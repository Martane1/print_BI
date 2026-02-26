@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%windows\install-windows.ps1"

if not exist "%PS_SCRIPT%" (
  echo Erro: arquivo nao encontrado: %PS_SCRIPT%
  pause
  exit /b 1
)

where powershell.exe >nul 2>nul
if errorlevel 1 (
  echo Erro: PowerShell nao encontrado neste Windows.
  pause
  exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
  echo.
  echo Falha na instalacao. Codigo: %EXIT_CODE%
  pause
  exit /b %EXIT_CODE%
)

echo.
echo Instalacao concluida.
echo Use o atalho "PRINT BI - Launcher" na Area de Trabalho.
pause
exit /b 0
