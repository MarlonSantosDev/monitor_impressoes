# Gera o .exe do Monitor de Impressões (PyInstaller)
# Uso: .\build.ps1
# Saída: dist\MonitorImpressoes.exe (copie para qualquer pasta e execute; não precisa de Python na máquina)

$ErrorActionPreference = "Stop"
$PastaProjeto = $PSScriptRoot

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Build: Monitor de Impressoes (.exe)" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

Set-Location $PastaProjeto

# Ambiente virtual (reutiliza o do projeto ou cria)
$venvPath = Join-Path $PastaProjeto "venv"
if (-not (Test-Path (Join-Path $venvPath "Scripts\python.exe"))) {
    Write-Host "[1/4] Criando ambiente virtual..." -ForegroundColor Yellow
    python -m venv venv
} else {
    Write-Host "[1/4] Usando ambiente virtual existente." -ForegroundColor Gray
}

$pip = Join-Path $venvPath "Scripts\pip.exe"
$py = Join-Path $venvPath "Scripts\python.exe"

Write-Host "[2/4] Instalando dependencias de build (PyInstaller)..." -ForegroundColor Yellow
& $pip install -r requirements-build.txt -q
if ($LASTEXITCODE -ne 0) { exit 1 }

Write-Host "[3/4] Gerando executavel com PyInstaller..." -ForegroundColor Yellow
$pyinstaller = Join-Path $venvPath "Scripts\pyinstaller.exe"
& $pyinstaller @(
    "--onefile",           # um unico .exe
    "--console",           # janela de console (para ver mensagens)
    "--name", "MonitorImpressoes",
    "--distpath", "dist",
    "--workpath", "build",
    "--specpath", ".",
    "--clean",
    "--noconfirm",
    "--hidden-import", "win32print",
    "--hidden-import", "win32api",
    "app.py"
)
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERRO: PyInstaller falhou." -ForegroundColor Red
    exit 1
}

Write-Host "[4/4] Limpando arquivos temporarios..." -ForegroundColor Yellow
if (Test-Path "build") { Remove-Item -Recurse -Force "build" }
$specFile = Join-Path $PastaProjeto "MonitorImpressoes.spec"
if (Test-Path $specFile) { Remove-Item $specFile }

$exePath = Join-Path $PastaProjeto "dist\MonitorImpressoes.exe"
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Build concluido com sucesso!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Executavel: $exePath" -ForegroundColor Cyan
Write-Host ""
Write-Host "Como usar:" -ForegroundColor White
Write-Host "  1. Copie dist\MonitorImpressoes.exe para a maquina destino (ou pasta desejada)." -ForegroundColor Gray
Write-Host "  2. Execute o .exe (duplo clique ou pelo prompt). Nao precisa instalar Python." -ForegroundColor Gray
Write-Host "  3. Log do dia na raiz: log_impressoes_DDMMYYYY.xlsx; pasta arquivos\ criada na raiz." -ForegroundColor Gray
Write-Host ""
