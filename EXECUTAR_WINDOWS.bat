@echo off
setlocal

set "ROOT_DIR=%~dp0"
set "INSTALL_SCRIPT=%ROOT_DIR%windows\install-windows.ps1"
set "RUN_SCRIPT=%ROOT_DIR%windows\run-print-bi.ps1"

if not exist "%INSTALL_SCRIPT%" (
  echo Erro: arquivo nao encontrado: %INSTALL_SCRIPT%
  pause
  exit /b 1
)

if not exist "%RUN_SCRIPT%" (
  echo Erro: arquivo nao encontrado: %RUN_SCRIPT%
  pause
  exit /b 1
)

where powershell.exe >nul 2>nul
if errorlevel 1 (
  echo Erro: PowerShell nao encontrado neste Windows.
  pause
  exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%INSTALL_SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
  echo.
  echo Falha na instalacao. Codigo: %EXIT_CODE%
  pause
  exit /b %EXIT_CODE%
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%RUN_SCRIPT%" -Action Menu
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" (
  echo.
  echo Falha na execucao. Codigo: %EXIT_CODE%
  pause
  exit /b %EXIT_CODE%
)

exit /b 0
