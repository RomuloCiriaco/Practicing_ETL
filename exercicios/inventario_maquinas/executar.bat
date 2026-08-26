@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   SATO INVENTARIO - coleta nesta maquina
echo ============================================
echo.

if not exist "dados\scans" mkdir "dados\scans"
if not exist "dados\logs" mkdir "dados\logs"

for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "TS=%%I"
set "OUT=dados\scans\%COMPUTERNAME%_%TS%.json"
set "LOG=dados\logs\coleta_%COMPUTERNAME%_%TS%.log"

call :log "=== SATO coleta inicio ==="
call :log "Hostname=%COMPUTERNAME%"
call :log "Usuario=%USERNAME%"
call :log "Pasta=%CD%"
call :log "Saida=%OUT%"

if not exist "ferramentas\machine-scanner-windows.exe" (
  call :log "[ERRO] scanner nao encontrado: ferramentas\machine-scanner-windows.exe"
  echo [ERRO] Nao encontrei ferramentas\machine-scanner-windows.exe
  echo Baixe em: https://github.com/JorgeEd13/machine_scanner/releases/latest
  echo Log: %CD%\%LOG%
  pause
  exit /b 1
)

echo Hostname : %COMPUTERNAME%
echo Saida    : %OUT%
echo Log      : %LOG%
echo.
echo Coletando... (pode levar alguns segundos)
echo Dica: clique com o botao direito em executar.bat e
echo       "Executar como administrador" para ler o Serial completo.
echo.

call :log "Iniciando machine-scanner-windows.exe --json"
"ferramentas\machine-scanner-windows.exe" --json -o "%OUT%" >> "%LOG%" 2>&1
set "RC=%ERRORLEVEL%"
call :log "ExitCode=%RC%"

echo.
if not "%RC%"=="0" (
  call :log "[ERRO] scanner falhou com codigo %RC%"
  echo [ERRO] Scanner retornou codigo %RC%
  echo Veja o log: %CD%\%LOG%
  if exist "%OUT%" del "%OUT%" >nul 2>&1
  pause
  exit /b %RC%
)

if not exist "%OUT%" (
  call :log "[ERRO] JSON nao foi criado: %OUT%"
  echo [ERRO] Arquivo JSON nao foi criado.
  echo Veja o log: %CD%\%LOG%
  pause
  exit /b 1
)

for %%A in ("%OUT%") do call :log "JSON OK size=%%~zA bytes"
call :log "[OK] coleta concluida"
echo [OK] JSON salvo em:
echo      %CD%\%OUT%
echo Log:
echo      %CD%\%LOG%
echo.
echo Proximo passo: na sua mesa, rode consolidar_inventario.py
echo para gerar dados\inventario.csv com SATO-001 / MON-001...
echo.
pause
endlocal
exit /b 0

:log
>> "%LOG%" echo [%DATE% %TIME%] %~1
exit /b 0
