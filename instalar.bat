@echo off
chcp 1252 >nul
REM ==========================================
REM SCRIPT DE INSTALAÇÃO - DASHBOARD ESTOQUE CPFANI
REM ==========================================

setlocal enabledelayedexpansion

set "LOGFILE=%~dp0instalar.log"
echo. > "%LOGFILE%"

call :LOG "Iniciando processo de instalação em %DATE% %TIME%"

REM Verificar privilégios de administrador
net session >nul 2>&1
if %errorLevel% neq 0 (
    call :LOG "Privilégios de administrador não detectados. Solicitando elevação..."
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb runAs"
    exit /b
)

call :LOG "Privilégios de administrador confirmados."

REM Verificar se o Python já está instalado
call :LOG "Verificando se o Python está instalado..."
python --version >nul 2>&1
if %errorLevel% neq 0 (
    call :LOG "Python não encontrado. Iniciando download e instalação..."
    
    set "PYTHON_URL=https://www.python.org/ftp/python/3.11.8/python-3.11.8-amd64.exe"
    set "PYTHON_INSTALLER=%~dp0python-installer.exe"
    
    call :LOG "Baixando instalador do Python 3.11.8..."
    powershell -Command "try { Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%PYTHON_INSTALLER%' -UseBasicParsing; exit 0 } catch { exit 1 }"
    
    if exist "!PYTHON_INSTALLER!" (
        call :LOG "Instalando Python silenciosamente para todos os usuários..."
        "!PYTHON_INSTALLER!" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0 Include_launcher=1
        
        REM Atualizar PATH para a sessão atual
        set "PATH=C:\Program Files\Python311;C:\Program Files\Python311\Scripts;!PATH!"
        
        call :LOG "Instalação do Python concluída."
        del "!PYTHON_INSTALLER!"
    ) else (
        call :LOG "ERRO: Falha ao baixar o instalador do Python."
        pause
        exit /b 1
    )
) else (
    for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PYTHON_VER=%%i
    call :LOG "Python já está instalado: !PYTHON_VER!"
)

REM Verificar se o pip está instalado
call :LOG "Verificando se o pip está instalado..."
pip --version >nul 2>&1
if %errorLevel% neq 0 (
    call :LOG "ERRO: pip não encontrado. Tentando reparar..."
    python -m ensurepip --upgrade >> "%LOGFILE%" 2>&1
)

REM Instalar dependências
if not exist "%~dp0requirements.txt" (
    call :LOG "ERRO: Arquivo requirements.txt não encontrado em %~dp0"
    pause
    exit /b 1
)

call :LOG "Instalando/atualizando dependências do Python..."
pip install --upgrade pip >> "%LOGFILE%" 2>&1
pip install -r "%~dp0requirements.txt" >> "%LOGFILE%" 2>&1

if %errorLevel% neq 0 (
    call :LOG "ERRO: Falha ao instalar algumas dependências. Verifique o log."
) else (
    call :LOG "Dependências instaladas com sucesso."
)

call :LOG "Processo de instalação concluído em %DATE% %TIME%"
echo.
echo ==========================================
echo Instalação concluída!
echo Verifique o arquivo de log: %LOGFILE%
echo ==========================================
pause
exit /b

:LOG
echo [%TIME%] %~1 >> "%LOGFILE%"
echo [%TIME%] %~1
exit /b