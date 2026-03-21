# office2pdf

调用本地 office 将文件转为 pdf 并返回

## 环境说明

- `uv` 负责管理 Python 版本、虚拟环境和 Python 包依赖。
- 运行时配置支持写入 `office2pdf.json`，放在源码目录或 `exe` 同目录即可，不必依赖环境变量。
- 本项目当前锁定的 Python 和打包依赖组合已经验证可用，请不要随意升级 `pyinstaller`、`pyinstaller-hooks-contrib`、`pywin32` 或修改 `uv.lock`。
- 如需升级上述版本，请先在测试环境重新打包并验证后，再更新锁文件。

## 开发运行

```powershell
uv sync
uv run python main.py
```

说明：

- 建议在 Windows PowerShell 中执行以上命令。
- 如果使用编辑器，请将 Python 解释器切换到 `D:\github\office2pdf\.venv\Scripts\python.exe`。

## 配置文件

程序会优先读取运行目录下的 `office2pdf.json`：

- 直接运行源码时：读取项目目录下的 `office2pdf.json`
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

如果希望把 Visual C++ 运行库安装包一并放入 `dist`，可以执行：

```powershell
.\build.ps1 -VcRedistPath C:\path\to\vc_redist.x64.exe
```

输出文件：

- `dist/office2pdf.exe`
- `dist/office2pdf-dir\`
- `dist/office2pdf-service.exe`
- `dist\vc_redist.x64.exe`（仅在传入 `-VcRedistPath` 时生成）

打包后的运行目录说明：

- 单文件 `exe` 运行时，PyInstaller 会使用临时解包目录，但本项目的默认日志和上传目录会按 `exe` 所在目录计算。
- 目录版 `dist\office2pdf-dir\` 运行时不会使用 PyInstaller 的临时解包目录，更适合在老系统上排查运行库问题。
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
