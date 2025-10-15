"""用于批量刷新客户 Excel 表，提取邮箱中的最新邮件动态。

脚本会通过 Outlook/Exchange 接口连接邮箱，为每位客户分别检索拜访报告、
技术支持、询价三类邮件，并把最新的接收时间、主题及 AI 总结写入 Excel。
配置既可以通过 YAML 文件提供，也可以用环境变量注入。
"""

from __future__ import annotations

import argparse
import dataclasses
import datetime as dt
import logging
import os
import re
from typing import Dict, Iterable, Optional, Sequence, Tuple

import pandas as pd
import yaml
import requests

from exchangelib import (
    Account,
    Configuration,
    Credentials,
    DELEGATE,
    EWSTimeZone,
)
from exchangelib.errors import AutoDiscoverFailed, TransportError
from exchangelib.queryset import Q


# 默认的 DeepSeek 模型名称
DEFAULT_DEEPSEEK_MODEL = "deepseek-chat"
# 默认的 DeepSeek 接口地址
DEFAULT_DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"
# 直接内置的 DeepSeek 密钥（按照您的要求）
DEFAULT_DEEPSEEK_API_KEY = "sk-1ac41aa69c9646b4b1a7a60487ffe230"
# 直接内置的邮箱用户名和密码，方便立即调试（Outlook 登录）
DEFAULT_OUTLOOK_USERNAME = "eddy.wang@nxp.com"
DEFAULT_OUTLOOK_PASSWORD = "pP131313."


LOGGER = logging.getLogger(__name__)


def _to_bool(value) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        return value.strip().lower() in {"1", "true", "yes", "y", "on"}
    return bool(value)


@dataclasses.dataclass
class MailboxConfig:
    """封装连接 Outlook/Exchange 所需的基础参数。"""

    username: str
    password: str
    server: Optional[str] = None
    mailbox: str = "Inbox"
    autodiscover: bool = True

    @classmethod
    def from_mapping(cls, data: Dict[str, str]) -> "MailboxConfig":
        missing = {key for key in ("username", "password") if key not in data}
        if missing:
            raise ValueError(f"缺少邮箱配置字段: {sorted(missing)}")
        return cls(
            username=data["username"],
            password=data["password"],
            server=data.get("server"),
            mailbox=data.get("mailbox", "Inbox"),
            autodiscover=_to_bool(data.get("autodiscover", data.get("server") is None)),
        )


@dataclasses.dataclass
class SearchRule:
    """描述检索某类邮件时需要匹配的关键词。"""

    name: str
    required_keywords: Sequence[str]
    summary_type: str


@dataclasses.dataclass
class EmailHit:
    """保存命中的邮件主题、接收时间与正文。"""

    subject: str
    received_at: dt.datetime
    body: str

    def as_display(self) -> Tuple[str, str]:
        iso_time = self.received_at.isoformat(sep=" ", timespec="minutes")
        return iso_time, self.subject


CALL_REPORT_RULE = SearchRule(
    name="call_report",
    required_keywords=("call report",),
    summary_type="call",
)
TECH_SUPPORT_RULE = SearchRule(
    name="technical_support",
    required_keywords=("technical support",),
    summary_type="technical",
)
PRICE_SUPPORT_RULE = SearchRule(
    name="price_support",
    required_keywords=("price support",),
    summary_type="price",
)
CATEGORY_RULES = (CALL_REPORT_RULE, TECH_SUPPORT_RULE, PRICE_SUPPORT_RULE)


class MailClient:
    """对 exchangelib 的轻量封装，便于执行 Outlook 搜索。"""

    def __init__(self, config: MailboxConfig) -> None:
        self._config = config
        self._account: Optional[Account] = None
        self._folder = None
        self._timezone: EWSTimeZone = EWSTimeZone.timezone("UTC")

    def __enter__(self) -> "MailClient":
        LOGGER.debug("正在准备连接 Outlook 帐户 %s", self._config.username)
        credentials = Credentials(
            username=self._config.username,
            password=self._config.password,
        )
        try:
            if self._config.autodiscover:
                account = Account(
                    primary_smtp_address=self._config.username,
                    credentials=credentials,
                    autodiscover=True,
                    access_type=DELEGATE,
                )
            else:
                if not self._config.server:
                    raise ValueError("当 autodiscover 为 False 时必须提供 server 地址")
                configuration = Configuration(
                    server=self._config.server,
                    credentials=credentials,
                )
                account = Account(
                    primary_smtp_address=self._config.username,
                    credentials=credentials,
                    autodiscover=False,
                    config=configuration,
                    access_type=DELEGATE,
                )
        except AutoDiscoverFailed as exc:
            LOGGER.error("自动发现失败，请在配置中提供 server 或检查账号权限。")
            raise exc
        except TransportError as exc:
            LOGGER.error("连接 Outlook 服务失败，请确认网络或服务器设置。")
            raise exc
        self._account = account
        self._folder = self._resolve_folder(account, self._config.mailbox)
        if account.default_timezone:
            self._timezone = account.default_timezone
        return self

    def __exit__(self, exc_type, exc, tb) -> None:  # type: ignore[override]
        self._account = None
        self._folder = None

    def _resolve_folder(self, account: Account, path: str):
        """根据配置定位邮箱文件夹，默认使用收件箱。"""

        if not path or path.strip().lower() == "inbox":
            return account.inbox
        parts = [segment.strip() for segment in path.split("/") if segment.strip()]
        if not parts:
            return account.inbox
        candidates = [account.root, account.inbox.parent]
        for root in candidates:
            if root is None:
                continue
            current = root
            try:
                for part in parts:
                    current = current / part
                return current
            except Exception:
                continue
        LOGGER.warning("未能定位文件夹 %s，自动回退到收件箱", path)
        return account.inbox

    def search_latest(
        self, keywords: Sequence[str], since: dt.datetime
    ) -> Optional[EmailHit]:
        if self._account is None or self._folder is None:
            raise RuntimeError("邮件客户端尚未建立连接")

        timezone = self._timezone or EWSTimeZone.timezone("UTC")
        if since.tzinfo is None:
            since_ews = timezone.localize(since)
            cutoff_utc = since.replace(tzinfo=dt.timezone.utc)
        else:
            localized = since.astimezone(timezone).replace(tzinfo=None)
            since_ews = timezone.localize(localized)
            cutoff_utc = since.astimezone(dt.timezone.utc)

        query = Q(datetime_received__gte=since_ews)
        for keyword in keywords:
            query &= Q(subject__icontains=keyword)

        queryset = (
            self._folder.filter(query)
            .only("subject", "datetime_received", "text_body", "body")
            .order_by("-datetime_received")
        )
        message = queryset.first()
        if not message:
            LOGGER.debug("未找到满足条件 %s 的邮件", keywords)
            return None

        if message.datetime_received is None:
            received = dt.datetime.now(dt.timezone.utc)
        else:
            received = message.datetime_received.astimezone(dt.timezone.utc)

        body = ""
        text_body = getattr(message, "text_body", None)
        if text_body:
            body = str(text_body)
        else:
            raw_body = getattr(message, "body", None)
            if raw_body:
                body = str(raw_body)

        hit = EmailHit(
            subject=getattr(message, "subject", "") or "",
            received_at=received,
            body=body,
        )
        if hit.received_at < cutoff_utc:
            LOGGER.debug("最新邮件时间 %s 早于限定的 %s，忽略", hit.received_at, cutoff_utc)
            return None
        return hit


class DeepSeekClient:
    """与 DeepSeek API 交互的简单客户端。"""

    def __init__(
        self,
        api_key: str,
        api_url: str = DEFAULT_DEEPSEEK_URL,
        model: str = DEFAULT_DEEPSEEK_MODEL,
        timeout: int = 30,
    ) -> None:
        # 保存 DeepSeek 配置信息，供后续请求使用
        self.api_key = api_key
        self.api_url = api_url
        self.model = model
        self.timeout = timeout

    def summarize_call_report(self, row: pd.Series, body: str) -> str:
        """调用 DeepSeek 生成拜访报告的总结。"""

        prompt = build_call_prompt(row, body)
        return self._request_summary(prompt, fallback=legacy_call_summary(row, body))

    def summarize_support(self, prefix: str, body: str) -> str:
        """调用 DeepSeek 生成技术支持或价格支持总结。"""

        prompt = build_support_prompt(prefix, body)
        fallback = legacy_support_summary(prefix, body)
        return self._request_summary(prompt, fallback=fallback)

    def _request_summary(self, prompt: str, fallback: str) -> str:
        """向 DeepSeek 发送请求，如果失败则回退到本地总结。"""

        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        payload = {
            "model": self.model,
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "You are an English-writing assistant helping a sales manager "
                        "summarize customer emails into concise status updates."
                    ),
                },
                {"role": "user", "content": prompt},
            ],
            "temperature": 0.3,
            "max_tokens": 300,
        }

        try:
            # 发送 HTTP POST 请求到 DeepSeek
            response = requests.post(
                self.api_url,
                json=payload,
                headers=headers,
                timeout=self.timeout,
            )
            response.raise_for_status()
            data = response.json()
            # DeepSeek 接口兼容 OpenAI 格式，choices[0]["message"]["content"] 为模型回答
            content = (
                data.get("choices", [{}])[0]
                .get("message", {})
                .get("content", "")
                .strip()
            )
            if content:
                return content
        except Exception:  # pragma: no cover - 网络错误兜底
            LOGGER.exception("调用 DeepSeek API 失败，改用本地兜底总结")
        return fallback


def build_call_prompt(row: pd.Series, body: str) -> str:
    """构建拜访报告总结所需的提示词。"""

    customer = row.get("Customer", "Unknown Customer")
    project = row.get("Project Application", "N/A")
    chipset = row.get("Recommended NXP Chipset", "NXP solution")
    oem = row.get("Target OEM", "the OEM customer")
    sop = row.get("SOP Quarter", "2*Q*")
    lifecycle = row.get("Lifecycle Demand (K sets)", "**")
    lifecycle_display = lifecycle
    if isinstance(lifecycle, (int, float)):
        lifecycle_display = f"{lifecycle:g}"
    else:
        lifecycle_display = str(lifecycle)
    if lifecycle_display and not lifecycle_display.lower().endswith("ksets"):
        lifecycle_display = f"{lifecycle_display}Ksets"

    return (
        "Summarize the following call report email into one English paragraph that "
        "follows this template exactly:\n"
        f"**: {customer}/**: {project},{chipset} (LTR: $1M, Open) : project is for {oem}. "
        f"<SUMMARY> SOP date is about {sop}. Lifecycle demand is about {lifecycle_display}.\n"
        "Replace <SUMMARY> with 1-2 sentences covering project progress, competition, "
        "customer needs, and any action items. Expand important details when appropriate.\n"
        "If the email lacks details, craft a reasonable status based on the text.\n\n"
        "Email body:\n" + body.strip()
    )


def build_support_prompt(prefix: str, body: str) -> str:
    """构建技术/询价总结的提示词。"""

    return (
        "Summarize the following customer email into a concise English sentence starting with "
        f"'{prefix}:'. Mention the key problem, status, and next steps if available. "
        "Do not add extra introductions.\n\n"
        "Email body:\n" + body.strip()
    )


def legacy_call_summary(row: pd.Series, body: str) -> str:
    """当接口不可用时的本地拜访报告总结兜底逻辑。"""

    customer = row.get("Customer", "Unknown Customer")
    project = row.get("Project Application", "N/A")
    chipset = row.get("Recommended NXP Chipset", "NXP solution")
    oem = row.get("Target OEM", "the OEM customer")
    sop = row.get("SOP Quarter", "2*Q*")
    lifecycle = row.get("Lifecycle Demand (K sets)", "**")
    bullet = summarize_sentences(body, max_sentences=2)
    if not bullet:
        bullet = "No additional notes were detected in the latest report."
    if isinstance(lifecycle, (int, float)):
        lifecycle_display = f"{lifecycle:g}"
    else:
        lifecycle_display = str(lifecycle)
    if lifecycle_display and not lifecycle_display.lower().endswith("ksets"):
        lifecycle_display = f"{lifecycle_display}Ksets"
    return (
        f"**: {customer}/**: {project},{chipset} (LTR: $1M, Open) : project is for {oem}. "
        f"{bullet} SOP date is about {sop}. Lifecycle demand is about {lifecycle_display}."
    )


def legacy_support_summary(prefix: str, body: str) -> str:
    """当接口不可用时的本地支持总结兜底逻辑。"""

    summary = summarize_sentences(body, max_sentences=2)
    if summary:
        return f"{prefix}: {summary}"
    return f"{prefix}: Awaiting more details from the latest message."


def summarize_sentences(text: str, max_sentences: int = 3) -> str:
    """简单的句子提取函数，作为 DeepSeek 失败时的兜底方案。"""

    sentences = re.split(r"(?<=[.!?])\s+", text.strip())
    sentences = [sentence.strip() for sentence in sentences if sentence.strip()]
    if not sentences:
        return text.strip()[:240]
    return " ".join(sentences[:max_sentences])


def update_dataframe(
    df: pd.DataFrame,
    customer_column: str,
    client: MailClient,
    rules: Iterable[SearchRule],
    summarizer: DeepSeekClient,
    search_since: dt.datetime,
) -> pd.DataFrame:
    """遍历客户列表并调用 Outlook 与 DeepSeek，更新表格。"""

    df = df.copy()
    for rule in rules:
        time_column = f"Last {rule.name.replace('_', ' ').title()} Time"
        subject_column = f"Last {rule.name.replace('_', ' ').title()} Subject"
        summary_column = f"{rule.name.replace('_', ' ').title()} Summary"
        df[time_column] = "none"
        df[subject_column] = "none"
        df[summary_column] = "none"

        for index, row in df.iterrows():
            customer_name = str(row.get(customer_column, "")).strip()
            if not customer_name:
                continue
            keywords = list(rule.required_keywords) + [customer_name]
            hit = client.search_latest(keywords, since=search_since)
            if not hit:
                LOGGER.info("客户 %s 尚未匹配到 %s 邮件", customer_name, rule.name)
                continue
            received_at, subject = hit.as_display()
            df.at[index, time_column] = received_at
            df.at[index, subject_column] = subject

            if rule.summary_type == "call":
                df.at[index, summary_column] = summarizer.summarize_call_report(row, hit.body)
            elif rule.summary_type == "technical":
                df.at[index, summary_column] = summarizer.summarize_support(
                    "Technical status", hit.body
                )
            elif rule.summary_type == "price":
                df.at[index, summary_column] = summarizer.summarize_support(
                    "Pricing status", hit.body
                )
    # 根据拜访报告的时间输出 Warning 列
    warning_column = "Call Report Warning"
    call_time_column = "Last Call Report Time"
    now_utc = dt.datetime.now(dt.timezone.utc)
    threshold = now_utc - dt.timedelta(days=90)
    if warning_column not in df.columns:
        df[warning_column] = ""
    if call_time_column in df.columns:
        for index, value in df[call_time_column].items():
            warning_needed = False
            text_value = str(value).strip().lower()
            if not text_value or text_value == "none":
                warning_needed = True
            else:
                try:
                    parsed = dt.datetime.fromisoformat(str(value))
                    if parsed.tzinfo is None:
                        parsed = parsed.replace(tzinfo=dt.timezone.utc)
                    else:
                        parsed = parsed.astimezone(dt.timezone.utc)
                    if parsed < threshold:
                        warning_needed = True
                except Exception:
                    warning_needed = True
            df.at[index, warning_column] = "warning" if warning_needed else ""
    return df


def load_configuration(path: Optional[str]) -> Dict[str, str]:
    if path:
        with open(path, "r", encoding="utf-8") as handle:
            loaded = yaml.safe_load(handle) or {}
            if not isinstance(loaded, dict):
                raise ValueError("配置文件顶层必须是键值映射")
            return {str(k): v for k, v in loaded.items()}
    env_mapping = {
        "server": os.getenv("OUTLOOK_SERVER") or os.getenv("IMAP_SERVER"),
        "username": os.getenv("OUTLOOK_USERNAME") or os.getenv("IMAP_USERNAME"),
        "password": os.getenv("OUTLOOK_PASSWORD") or os.getenv("IMAP_PASSWORD"),
        "mailbox": os.getenv("OUTLOOK_MAILBOX") or os.getenv("IMAP_MAILBOX"),
        "autodiscover": os.getenv("OUTLOOK_AUTODISCOVER"),
        "deepseek_api_key": os.getenv("DEEPSEEK_API_KEY"),
        "deepseek_api_url": os.getenv("DEEPSEEK_API_URL"),
        "deepseek_model": os.getenv("DEEPSEEK_MODEL"),
    }
    mapping = {k: v for k, v in env_mapping.items() if v is not None}
    if not mapping.get("username"):
        mapping["username"] = DEFAULT_OUTLOOK_USERNAME
    if not mapping.get("password"):
        mapping["password"] = DEFAULT_OUTLOOK_PASSWORD
    return mapping


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", help="需要更新的 Excel 文件路径")
    parser.add_argument(
        "--config",
        help="包含邮箱与 DeepSeek 配置的 YAML 文件路径",
    )
    parser.add_argument(
        "--customer-column",
        default="Customer",
        help="Excel 中存放客户英文名称的列名",
    )
    parser.add_argument(
        "--sheet",
        default=0,
        help="需要更新的工作表名称或索引（默认第一个工作表）",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        help="日志级别（DEBUG、INFO、WARNING 等）",
    )
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> None:
    args = parse_args(argv)
    logging.basicConfig(level=getattr(logging, str(args.log_level).upper(), logging.INFO))

    config_mapping = load_configuration(args.config)
    mailbox_config = MailboxConfig.from_mapping(config_mapping)

    deepseek_api_key = config_mapping.get("deepseek_api_key", DEFAULT_DEEPSEEK_API_KEY)
    deepseek_api_url = config_mapping.get("deepseek_api_url", DEFAULT_DEEPSEEK_URL)
    deepseek_model = config_mapping.get("deepseek_model", DEFAULT_DEEPSEEK_MODEL)

    if deepseek_api_key == DEFAULT_DEEPSEEK_API_KEY:
        LOGGER.warning("当前正在使用代码内置的 DeepSeek 密钥，请注意妥善保密。")

    # 创建 DeepSeek 客户端，用于后续的内容理解与总结
    summarizer = DeepSeekClient(
        api_key=str(deepseek_api_key),
        api_url=str(deepseek_api_url),
        model=str(deepseek_model),
    )

    LOGGER.info("正在载入工作簿 %s", args.workbook)
    df = pd.read_excel(args.workbook, sheet_name=args.sheet)

    one_year_ago = dt.datetime.now(dt.timezone.utc) - dt.timedelta(days=365)

    with MailClient(mailbox_config) as client:
        LOGGER.info("开始刷新客户记录")
        updated = update_dataframe(
            df,
            args.customer_column,
            client,
            CATEGORY_RULES,
            summarizer=summarizer,
            search_since=one_year_ago,
        )

    LOGGER.info("写回更新后的数据到工作簿")
    writer_kwargs = {"engine": "openpyxl"}
    if os.path.exists(args.workbook):
        writer_kwargs.update({"mode": "a", "if_sheet_exists": "replace"})
    else:
        writer_kwargs.update({"mode": "w"})
    with pd.ExcelWriter(args.workbook, **writer_kwargs) as writer:
        updated.to_excel(writer, sheet_name=args.sheet, index=False)


if __name__ == "__main__":
    main()
