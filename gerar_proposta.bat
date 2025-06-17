@echo off
setlocal enabledelayedexpansion

REM Obter o diretório do script
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Função para ativar o ambiente virtual
:activate_venv
if exist "%SCRIPT_DIR%venv\Scripts\activate.bat" (
    call "%SCRIPT_DIR%venv\Scripts\activate.bat"
    goto :eof
) else (
    echo ❌ Ambiente virtual não encontrado
    exit /b 1
)

REM Verificar se Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python não encontrado. Por favor, instale o Python 3.
    pause
    exit /b 1
)

REM Verificar se pip está instalado
pip --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Pip não encontrado. Por favor, instale o pip.
    pause
    exit /b 1
)

REM Remover ambiente virtual antigo se existir e estiver corrompido
if exist "%SCRIPT_DIR%venv" (
    call :activate_venv >nul 2>&1
    if errorlevel 1 (
        echo 🔧 Ambiente virtual corrompido, recriando...
        rmdir /s /q "%SCRIPT_DIR%venv"
    )
)

REM Criar ambiente virtual se não existir
if not exist "%SCRIPT_DIR%venv" (
    echo 🔧 Criando ambiente virtual...
    python -m venv "%SCRIPT_DIR%venv"
    call :activate_venv
    
    echo 📦 Instalando dependências...
    python -m pip install --upgrade pip
    pip install -e .
) else (
    call :activate_venv
)

REM Criar diretórios necessários
if not exist "%SCRIPT_DIR%input" mkdir "%SCRIPT_DIR%input"
if not exist "%SCRIPT_DIR%output" mkdir "%SCRIPT_DIR%output"
if not exist "%SCRIPT_DIR%templates" mkdir "%SCRIPT_DIR%templates"

REM Encontrar arquivo Excel mais recente
set "EXCEL_FILE="
for /f "delims=" %%i in ('dir /b /o-d "%SCRIPT_DIR%input\*.xlsx" 2^>nul') do (
    if not defined EXCEL_FILE set "EXCEL_FILE=%SCRIPT_DIR%input\%%i"
)

set "TEMPLATE_FILE=%SCRIPT_DIR%templates\modelo.pptx"

REM Obter data e hora atual para nome do arquivo
for /f "tokens=1-6 delims=/:. " %%a in ('echo %date% %time%') do (
    set "DATETIME=%%c%%a%%b_%%d%%e"
)

set "OUTPUT_FILE=%SCRIPT_DIR%output\proposta_%DATETIME%.pptx"

REM Verificar Excel
if not defined EXCEL_FILE (
    echo ❌ Nenhum arquivo Excel (.xlsx) encontrado na pasta input/
    echo Por favor, coloque seu arquivo Excel em: %SCRIPT_DIR%input/
    pause
    exit /b 1
)

echo 📊 Usando arquivo Excel: %EXCEL_FILE%

REM Verificar template
if not exist "%TEMPLATE_FILE%" (
    echo ❌ Template não encontrado: %TEMPLATE_FILE%
    echo Por favor, coloque o arquivo modelo.pptx em: %SCRIPT_DIR%templates/
    pause
    exit /b 1
)

REM Verificar se o comando está disponível
gerar-proposta --help >nul 2>&1
if errorlevel 1 (
    echo ❌ Comando não encontrado. Reinstalando pacote...
    pip install -e .
)

REM Executar script
echo 🚀 Gerando apresentação...
python -m proposta_solar.cli ^
    --excel "%EXCEL_FILE%" ^
    --template "%TEMPLATE_FILE%" ^
    --output "%OUTPUT_FILE%" ^
    --verbose

REM Verificar resultado
if %errorlevel% equ 0 (
    echo ✅ Apresentação gerada com sucesso!
    echo 📄 Arquivo salvo em: %OUTPUT_FILE%
    
    REM Abrir o arquivo no Windows
    start "" "%OUTPUT_FILE%"
) else (
    echo ❌ Erro ao gerar apresentação
    pause
    exit /b 1
)

pause 