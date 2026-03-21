param(
    [switch]$Clean,
    [string]$VcRedistPath
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
    if (($venvConfig -match "Python313") -and $pythonVersion.StartsWith("3.12")) {
        $message = @(
            "The current .venv is still linked to Python 3.13, but this project requires Python $pythonVersion."
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

& $uv run pyinstaller --noconfirm --clean --onedir --noconsole `
    --name office2pdf-dir `
    main.py

& $uv run pyinstaller --noconfirm --clean --onefile `
    --name office2pdf-service `
    --hidden-import=win32timezone `
    --hidden-import=servicemanager `
    service.py

if ($VcRedistPath) {
    if (-not (Test-Path $VcRedistPath)) {
        throw "VC++ redistributable not found: $VcRedistPath"
    }
    Copy-Item -Force $VcRedistPath "dist\vc_redist.x64.exe"
    Write-Host "Copied VC++ redistributable to dist\vc_redist.x64.exe"
}

Write-Host "Build complete:"
Write-Host "  dist/office2pdf.exe"
Write-Host "  dist/office2pdf-dir/"
Write-Host "  dist/office2pdf-service.exe"
if ($VcRedistPath) {
    Write-Host "  dist/vc_redist.x64.exe"
}
