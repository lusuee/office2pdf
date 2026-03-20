# office2pdf

调用本地 office 将文件转为 pdf 并返回

## 环境说明

- `uv` 负责管理 Python 版本、虚拟环境和 Python 包依赖。
- `uv` 不能把 Windows 专属模块变成跨平台可用，比如 `pywin32`、`win32event` 只能在 Windows Python 下使用。
- 本项目建议区分两种使用场景：
  - `main.py` 可在 WSL/Linux 环境下开发和检查代码结构。
  - `service.py`、`pywin32`、Windows 服务注册与 exe 打包需要在 Windows 环境执行。
- 运行时配置支持写入 `office2pdf.json`，放在源码目录或 `exe` 同目录即可，不必依赖环境变量。

## 开发运行

```bash
/home/linuxbrew/.linuxbrew/bin/uv sync
/home/linuxbrew/.linuxbrew/bin/uv run python main.py
```

说明：

- 在非 Windows 环境下，`uv sync` 不会安装 `pywin32`，这是预期行为。
- 如果编辑器使用的是 WSL/Linux 解释器，`service.py` 中的 Windows 专属导入可能无法被源码解析，这不影响 `main.py` 的开发。
- 如果需要让编辑器正确识别 `win32event` 等模块，请把解释器切换到 Windows 虚拟环境，例如 `D:\github\office2pdf\.venv\Scripts\python.exe`。

## 配置文件

程序会优先读取运行目录下的 `office2pdf.json`：

- 源码运行时：读取项目目录下的 `office2pdf.json`
- 打包 `exe` 运行时：读取 `exe` 同目录下的 `office2pdf.json`

可以参考 [office2pdf.json.example](/mnt/d/github/office2pdf/office2pdf.json.example)：

```json
{
  "host": "0.0.0.0",
  "port": 8081,
  "log_dir": "logs",
  "upload_dir": "uploads",
  "max_content_length": "200m"
}
```

说明：

- `log_dir` 和 `upload_dir` 支持相对路径；相对路径会按项目目录或 `exe` 目录计算。
- `max_content_length` 支持更直观的单位写法，例如 `200m`、`512k`、`1g`，也兼容直接写字节数。
- 如果同时配置了 `office2pdf.json` 和环境变量，优先使用配置文件。
- 不提供配置文件时，程序才会回退到环境变量和默认值。

## 生成 exe

在 Windows 环境执行：

```powershell
.\build.ps1
```

首次在 Windows 环境使用时，建议先执行：

```powershell
uv sync --group dev
```

如果 `.\build.ps1` 执行报错，可按下面排查：

1. 确认在 Windows PowerShell 中可以直接执行 `uv --version`。
2. 如果出现 `无法加载文件 .\build.ps1，因为在此系统上禁止运行脚本`，说明 PowerShell 执行策略阻止了 `.ps1` 脚本。
3. 可先在当前 PowerShell 会话临时放行脚本执行：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\build.ps1
```

4. 如果希望长期对当前用户生效，可执行：

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

然后重新打开 PowerShell，再执行：

```powershell
.\build.ps1
```

5. 如果出现脚本语法错误，或者错误信息里有乱码、`字符串缺少终止符`、`UnexpectedToken`，请先同步最新版 `build.ps1`。旧脚本可能存在 here-string 或文件编码问题。
6. 如果出现类似 `No Python at ...Python312\python.exe` 或 `Querying Python at .venv\Scripts\python.exe failed`，说明现有 `.venv` 仍然绑定旧 Python。
7. 删除旧虚拟环境并重新安装依赖：

```powershell
Remove-Item -Recurse -Force .venv
uv sync --group dev
```

8. 然后重新执行：

```powershell
.\build.ps1
```

常见原因：

- PowerShell 执行策略禁止运行 `.ps1`
- 使用了旧版本 `build.ps1`，其中包含 PowerShell here-string 语法问题
- 当前 PowerShell 里找不到 `uv`
- `.venv` 仍然指向已删除的旧 Python 安装
- 没有先在 Windows 环境完成 `uv sync --group dev`

输出文件：

- `dist/office2pdf.exe`
- `dist/office2pdf-service.exe`

打包后的运行目录说明：

- 单文件 `exe` 运行时，PyInstaller 会使用临时解包目录，但本项目的默认日志和上传目录会按 `exe` 所在目录计算。
- 如果未设置环境变量，默认目录如下：
  - 日志目录：`<exe所在目录>\logs`
  - 上传目录：`<exe所在目录>\uploads`
- 上传的原始文件和生成的 PDF 默认会保留在上传目录的日期子目录下，不会自动删除。

## 注册 Windows 服务

先构建 `dist/office2pdf-service.exe`，然后在管理员 PowerShell 中执行：

```powershell
.\dist\office2pdf-service.exe --startup auto install
.\dist\office2pdf-service.exe start
```

停止和卸载：

```powershell
.\dist\office2pdf-service.exe stop
.\dist\office2pdf-service.exe remove
```

## 环境变量

通常更推荐使用 `office2pdf.json`。下面这些环境变量仍然可以作为兜底配置：

- `OFFICE2PDF_PORT`: 服务端口，默认 `8081`
- `OFFICE2PDF_HOST`: 监听地址，默认 `0.0.0.0`
- `OFFICE2PDF_LOG_DIR`: 日志目录
- `OFFICE2PDF_UPLOAD_DIR`: 临时上传目录
- `OFFICE2PDF_MAX_CONTENT_LENGTH`: 上传大小限制，默认 `200m`（也兼容直接写字节数）
