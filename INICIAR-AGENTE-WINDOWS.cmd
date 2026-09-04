@echo off
setlocal
cd /d "%~dp0"
set "LOCAL_OPERATOR_MODE=true"
echo Iniciando Teams Actas Agent desde:
echo %CD%
call pnpm install
if errorlevel 1 (
  echo No se pudieron instalar las dependencias.
  pause
  exit /b 1
)
call pnpm dev
pause
