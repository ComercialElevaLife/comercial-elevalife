@echo off
setlocal

cd /d "%~dp0"

set "BASE_FILE=%CD%\Base_Painel_KPI_.xlsx"
set "EXTRACTOR=%CD%\scripts\extract_kpi_data.py"
set "DASH_SERVER=%CD%\scripts\dashboard_server.py"
set "DASH_DIR=%CD%\dashboard"
set "DATA_FILE=%DASH_DIR%\data\kpi_data.json"
set "PORT=5500"
set "URL=http://localhost:%PORT%/"

where py >nul 2>&1
if not errorlevel 1 (
  set "PY_CMD=py -3"
) else (
  where python >nul 2>&1
  if errorlevel 1 (
    echo [ERRO] Python nao encontrado no PATH.
    echo Instale o Python e tente novamente.
    pause
    exit /b 1
  )
  set "PY_CMD=python"
)

echo Encerrando servidores antigos da porta %PORT%...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-CimInstance Win32_Process | ? { $_.Name -match '^python(\\.exe)?$' -and (($_.CommandLine -match 'http\\.server %PORT%') -or ($_.CommandLine -match 'dashboard_server\\.py')) } | % { Stop-Process -Id $_.ProcessId -Force }" >nul 2>&1

echo [1/3] Atualizando base JSON...
"%PY_CMD%" "%EXTRACTOR%" --input "%BASE_FILE%" --output "%DATA_FILE%"
if errorlevel 1 (
  echo [ERRO] Falha ao gerar o JSON.
  pause
  exit /b 1
)

echo [2/3] Iniciando servidor local na porta %PORT% ...
start "KPI Dashboard Server" cmd /k "cd /d \"%CD%\" && \"%PY_CMD%\" \"%DASH_SERVER%\" --port %PORT%"
if errorlevel 1 (
  echo [ERRO] Nao foi possivel iniciar o servidor local.
  pause
  exit /b 1
)

echo Aguardando servidor ficar pronto...
timeout /t 2 /nobreak >nul

echo [3/3] Abrindo navegador em %URL% ...
set "CHROME_EXE=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
set "EDGE_EXE=%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"
if not exist "%EDGE_EXE%" set "EDGE_EXE=%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"

if exist "%CHROME_EXE%" (
  start "" "%CHROME_EXE%" "%URL%"
) else if exist "%EDGE_EXE%" (
  start "" "%EDGE_EXE%" "%URL%"
) else (
  start "" "%URL%"
)

endlocal
exit /b 0
