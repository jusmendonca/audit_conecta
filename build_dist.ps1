# Monta a versão distribuível do Auditoria Conecta+ para Windows.
#
# Gera dist\AuditoriaConecta\ com um Python embutido e todas as dependências,
# de modo que a aplicação rode em máquinas sem Python instalado. O usuário
# final descompacta a pasta e executa "Auditoria Conecta+.bat".
#
# Uso:  powershell -ExecutionPolicy Bypass -File build_dist.ps1

$ErrorActionPreference = "Stop"

$PyVersion = "3.11.9"
$Raiz      = $PSScriptRoot
$Build     = Join-Path $Raiz "build"
$Dist      = Join-Path $Raiz "dist\AuditoriaConecta"
$Runtime   = Join-Path $Dist "runtime"

Write-Host "== Auditoria Conecta+ · build da versão distribuível ==" -ForegroundColor Cyan

# 1. Limpa saídas anteriores
foreach ($d in @($Build, (Join-Path $Raiz "dist"))) {
    if (Test-Path $d) { Remove-Item -Recurse -Force $d }
}
New-Item -ItemType Directory -Force -Path $Build, $Runtime | Out-Null

# 2. Python embeddable
$ZipPy = Join-Path $Build "python-embed.zip"
$UrlPy = "https://www.python.org/ftp/python/$PyVersion/python-$PyVersion-embed-amd64.zip"
Write-Host "[1/5] Baixando Python $PyVersion (embeddable)..."
curl.exe -sSL -o $ZipPy $UrlPy
Expand-Archive -Path $ZipPy -DestinationPath $Runtime -Force

# 3. Habilita site-packages no runtime embutido (o ._pth vem com 'import site' comentado)
$Pth = Get-ChildItem -Path $Runtime -Filter "python*._pth" | Select-Object -First 1
$conteudo = Get-Content $Pth.FullName
$conteudo = $conteudo -replace '^#\s*import site', 'import site'
($conteudo + "Lib\site-packages") | Set-Content $Pth.FullName -Encoding ascii

# 4. pip + dependências
Write-Host "[2/5] Instalando pip..."
$GetPip = Join-Path $Build "get-pip.py"
curl.exe -sSL -o $GetPip "https://bootstrap.pypa.io/get-pip.py"
& "$Runtime\python.exe" $GetPip --no-warn-script-location -q

Write-Host "[3/5] Instalando dependências (pode levar alguns minutos)..."
& "$Runtime\python.exe" -m pip install --no-warn-script-location -q -r (Join-Path $Raiz "requirements.txt")

# 5. Código da aplicação
Write-Host "[4/5] Copiando a aplicação..."
$App = Join-Path $Dist "app"
New-Item -ItemType Directory -Force -Path $App | Out-Null
Copy-Item (Join-Path $Raiz "app.py") $App
Copy-Item (Join-Path $Raiz "modules") $App -Recurse
Get-ChildItem -Path $App -Include "__pycache__" -Recurse -Directory |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

# O SUPP_BASE_URL acompanha o pacote; o usuário final não precisa configurar nada.
# Gravamos em ascii: o -Encoding utf8 do PowerShell 5.1 escreve BOM, e tanto o
# parser TOML do Streamlit quanto o dotenv engasgam com ele.
"SUPP_BASE_URL=https://supersapiensbackend.agu.gov.br" |
    Set-Content (Join-Path $App ".env") -Encoding ascii

# Desativa a telemetria e a tela de boas-vindas do Streamlit no pacote.
$Cfg = Join-Path $Dist "app\.streamlit"
New-Item -ItemType Directory -Force -Path $Cfg | Out-Null
@"
[browser]
gatherUsageStats = false

[server]
headless = false

[global]
developmentMode = false
"@ | Set-Content (Join-Path $Cfg "config.toml") -Encoding ascii

# 6. Lançador
Write-Host "[5/5] Gerando o lançador..."
@"
@echo off
title Auditoria Conecta+
cd /d "%~dp0"
echo.
echo   Auditoria Conecta+ — iniciando...
echo   O navegador abrira em instantes. NAO FECHE esta janela: ela
echo   mantem a aplicacao no ar. Para encerrar, feche-a ou tecle Ctrl+C.
echo.
runtime\python.exe -m streamlit run app\app.py
if errorlevel 1 (
  echo.
  echo   A aplicacao terminou com erro. Copie a mensagem acima e envie ao suporte.
  pause
)
"@ | Set-Content (Join-Path $Dist "Auditoria Conecta+.bat") -Encoding ascii

Copy-Item (Join-Path $Raiz "LEIAME.txt") $Dist -ErrorAction SilentlyContinue

# 7. Zip final
$Zip = Join-Path $Raiz "dist\AuditoriaConecta.zip"
Compress-Archive -Path $Dist -DestinationPath $Zip -Force
Remove-Item -Recurse -Force $Build

$mb = [math]::Round((Get-Item $Zip).Length / 1MB, 1)
Write-Host ""
Write-Host "Pronto: dist\AuditoriaConecta.zip ($mb MB)" -ForegroundColor Green
Write-Host "Teste local: dist\AuditoriaConecta\'Auditoria Conecta+.bat'"
