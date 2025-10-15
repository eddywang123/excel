# Excel 邮件追踪更新工具

本项目提供一个命令行脚本，定期读取您在 Excel 中维护的客户清单，并自动到邮箱
里寻找最新的拜访报告、技术支持和询价邮件。程序会把邮件的接收时间、主题以及
DeepSeek 生成的英文摘要写回到同一个工作簿，便于您一键刷新客户状态。

每次刷新时，脚本会：

- 针对 Excel 表格中的每一位客户，检索过去一年内的邮件。
- 若没有命中任何符合条件的邮件，会在时间、主题、摘要列写入 `none` 便于识别。
- 如果某位客户三个月内未收到最新的拜访报告，会在新增的 “Call Report Warning” 列标记 `warning`，帮助快速筛选。

## 文件说明

- `email_excel_updater.py`：主程序，负责连接 Outlook/Exchange 邮箱、调用 DeepSeek API，并把
  结果写入 Excel，同时根据拜访报告时间生成预警。
- `config_example.yaml`：示例配置文件，告诉您需要提供哪些邮箱参数以及（可选）覆
  盖 DeepSeek 设置的方式。
- `requirements.txt`：运行脚本所需的 Python 第三方库清单，便于一次性安装依赖。

## 安装步骤

1. 建议使用虚拟环境隔离依赖：
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # Windows 使用 .venv\Scripts\activate
   ```
2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```

## 本地运行指南（零基础版）

1. **安装 Python**：前往 [python.org](https://www.python.org/downloads/) 下载并安装最新的 Python 3 版本，勾选 “Add Python to PATH”。
2. **获取项目代码**：
   - 如果熟悉 Git，可以在命令行里执行 `git clone https://.../excel.git`（请替换成真实仓库地址）。
   - 如果不会使用 Git，可以在网页上点击 “Download ZIP”，解压到一个容易找到的文件夹。
3. **打开命令行**：
   - Windows：按 `Win + R` 输入 `cmd` 或使用 PowerShell。
   - macOS：打开 “终端 (Terminal)”。
   - Linux：打开您常用的终端程序。
4. **切换到项目目录**：使用 `cd` 命令进入解压后的项目文件夹，例如 `cd C:\Users\你的用户名\Downloads\excel`。
5. **（可选但推荐）创建虚拟环境**：
   ```bash
   python -m venv .venv
   # Windows 激活
   .venv\Scripts\activate
   # macOS/Linux 激活
   source .venv/bin/activate
   ```
6. **安装依赖包**：
   ```bash
   pip install -r requirements.txt
   ```
7. **把客户 Excel 放到项目目录**：将客户清单 Excel 文件复制到同一个文件夹内，比如命名为 `customers.xlsx`。
8. **准备配置（可选）**：如果需要覆盖默认的邮箱或 DeepSeek 设置，可以复制 `config_example.yaml` 为 `config.yaml` 并按需修改；若使用默认设置，可跳过此步。
9. **运行脚本**：
   ```bash
   python email_excel_updater.py customers.xlsx
   ```
   如果有自定义配置文件，改用：
   ```bash
   python email_excel_updater.py customers.xlsx --config config.yaml
   ```
10. **等待输出**：终端会显示进度日志，Excel 会在脚本结束后更新。请确保运行时 Excel 没有打开该文件。

完成上述步骤后，您就能在本地刷新客户清单。如果需要定期运行，可把第 9 步的命令加入计划任务工具里。

## 上传客户名单 Excel 的方式

1. 把您的客户清单保存成 `.xlsx` 文件。
2. 将该文件复制到本项目目录中（可通过 VS Code 拖拽、`scp`、FTP 或任意文件管理方
   式）。如果文件放在其他路径，也可以在运行脚本时写绝对路径。
3. 确认表格中至少包含一列客户英文名称（默认列名为 `Customer`，可在命令行参数中
   指定其他列名）。如果需要让拜访报告摘要更完整，可额外添加 `Project Application`、
   `Recommended NXP Chipset`、`Target OEM`、`SOP Quarter`、`Lifecycle Demand (K sets)`
   等列。

运行脚本时传入 Excel 路径，例如：
```bash
python email_excel_updater.py ./customers.xlsx --config config.yaml
```

## 邮箱连接与 DeepSeek 调用

- 邮箱通过 Outlook/Exchange 接口访问。脚本会默认使用自动发现（Autodiscover）功
  能，只需提供邮箱账号即可连接。如果贵司关闭了自动发现，可在配置文件中填入
  `server` 指向 EWS 地址（例如 `outlook.office365.com`），并将 `autodiscover` 设为
  `false`。脚本已经内置您的邮箱账号 `eddy.wang@nxp.com` 与密码 `pP131313.`，如需
  测试其他账号，可在配置里覆盖 `username` 与 `password`。若公司邮箱需要 VPN，
  请在运行脚本前先连接 VPN。
- 脚本内置了您提供的 DeepSeek 密钥 `sk-1ac41aa69c9646b4b1a7a60487ffe230`，无需额外
  填写即可调用接口。如需改用自己的密钥，可在配置文件或环境变量里覆盖
  `deepseek_api_key`。
- DeepSeek 生成的内容全部为英文，以满足拜访报告模板的要求。

示例配置（可选，如果您想覆盖默认值）：
```yaml
# 将此文件保存为 config.yaml，并用下方字段覆盖默认设置
# server: outlook.office365.com  # 若禁用自动发现请取消注释
# autodiscover: true             # 为 false 时必须提供 server
# username: your.name@example.com
# password: app-specific-password
# mailbox: Inbox/子文件夹        # 如需访问子文件夹，可用 / 分隔
# 如需覆盖默认的 DeepSeek 设定，可取消下列注释
# deepseek_api_key: sk-your-own-key
# deepseek_api_url: https://api.deepseek.com/chat/completions
# deepseek_model: deepseek-chat
```

## 使用方法

```bash
python email_excel_updater.py 路径/到/客户清单.xlsx --config config.yaml
```

常用参数：
- `--sheet`：需要更新的工作表名称或索引（默认第一个工作表）。
- `--customer-column`：客户英文名称所在的列名。
- `--log-level`：日志级别，调试时可以设为 `DEBUG` 观察详细过程。

## 自动化建议

如果希望定期执行，可在 Windows 任务计划程序或 Linux/macOS 的 cron 中定时运行
上述命令，使 Excel 始终保持最新状态。

## 安全提示

- 邮箱账号和密码请妥善保管，推荐使用应用专用密码。
- 由于密钥已经写入代码，请注意限制代码的访问范围，仅在可信环境运行。
- 在脚本执行期间请保持 Excel 处于关闭状态，避免写入冲突。
