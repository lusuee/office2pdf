param(
    [switch]$Clean
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$uvCommand = Get-Command uv -ErrorAction SilentlyContinue
if ($null -eq $uvCommand) {
    throw "uv was not found in PATH. Run 'uv --version' in PowerShell first."
}

$uv = $uvCommand.Source

if ($env:LOCALAPPDATA) {
    $env:UV_CACHE_DIR = Join-Path -Path $env:LOCALAPPDATA -ChildPath "uv\cache"
}

if ($Clean) {
    Remove-Item -Recurse -Force dist, build -ErrorAction SilentlyContinue
}

if ((Test-Path ".venv\pyvenv.cfg") -and (Test-Path ".python-version")) {
    $venvConfig = Get-Content ".venv\pyvenv.cfg" -Raw
    $pythonVersion = (Get-Content ".python-version" -Raw).Trim()
    if (($venvConfig -match "Python312") -and $pythonVersion.StartsWith("3.13")) {
        $message = @(
            "The current .venv is still linked to Python 3.12, but this project requires Python $pythonVersion."
            "Delete .venv and recreate it with:"
            "  Remove-Item -Recurse -Force .venv"
            "  uv sync --group dev"
            ""
            "Then run:"
            "  .\build.ps1"
        ) -join [Environment]::NewLine
        throw $message
    }
}

& $uv sync --group dev

& $uv run pyinstaller --noconfirm --clean --onefile --noconsole `
    --name office2pdf `
    main.py

& $uv run pyinstaller --noconfirm --clean --onefile `
    --name office2pdf-service `
    --hidden-import=win32timezone `
    --hidden-import=servicemanager `
    service.py

Write-Host "Build complete:"
Write-Host "  dist/office2pdf.exe"
Write-Host "  dist/office2pdf-service.exe"
