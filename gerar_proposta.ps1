# Obter o diretório do script
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

# Função para ativar o ambiente virtual
function Activate-Venv {
    if (Test-Path "$ScriptDir\venv\Scripts\Activate.ps1") {
        & "$ScriptDir\venv\Scripts\Activate.ps1"
        return $true
    } else {
        Write-Host "❌ Ambiente virtual não encontrado" -ForegroundColor Red
        return $false
    }
}

# Verificar se Python está instalado
try {
    $pythonVersion = python --version 2>&1
    Write-Host "✅ Python encontrado: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "❌ Python não encontrado. Por favor, instale o Python 3." -ForegroundColor Red
    Read-Host "Pressione Enter para sair"
    exit 1
}

# Verificar se pip está instalado
try {
    $pipVersion = pip --version 2>&1
    Write-Host "✅ Pip encontrado: $pipVersion" -ForegroundColor Green
} catch {
    Write-Host "❌ Pip não encontrado. Por favor, instale o pip." -ForegroundColor Red
    Read-Host "Pressione Enter para sair"
    exit 1
}

# Remover ambiente virtual antigo se existir e estiver corrompido
if (Test-Path "$ScriptDir\venv") {
    try {
        Activate-Venv | Out-Null
    } catch {
        Write-Host "🔧 Ambiente virtual corrompido, recriando..." -ForegroundColor Yellow
        Remove-Item -Recurse -Force "$ScriptDir\venv"
    }
}

# Criar ambiente virtual se não existir
if (-not (Test-Path "$ScriptDir\venv")) {
    Write-Host "🔧 Criando ambiente virtual..." -ForegroundColor Yellow
    python -m venv "$ScriptDir\venv"
    Activate-Venv
    
    Write-Host "📦 Instalando dependências..." -ForegroundColor Yellow
    python -m pip install --upgrade pip
    pip install -e .
} else {
    Activate-Venv
}

# Criar diretórios necessários
@("input", "output", "templates") | ForEach-Object {
    if (-not (Test-Path "$ScriptDir\$_")) {
        New-Item -ItemType Directory -Path "$ScriptDir\$_" | Out-Null
    }
}

# Encontrar arquivo Excel mais recente
$ExcelFiles = Get-ChildItem "$ScriptDir\input\*.xlsx" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
$ExcelFile = if ($ExcelFiles) { $ExcelFiles[0].FullName } else { $null }

$TemplateFile = "$ScriptDir\templates\modelo.pptx"
$DateTime = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile = "$ScriptDir\output\proposta_$DateTime.pptx"

# Verificar Excel
if (-not $ExcelFile) {
    Write-Host "❌ Nenhum arquivo Excel (.xlsx) encontrado na pasta input/" -ForegroundColor Red
    Write-Host "Por favor, coloque seu arquivo Excel em: $ScriptDir\input\" -ForegroundColor Red
    Read-Host "Pressione Enter para sair"
    exit 1
}

Write-Host "📊 Usando arquivo Excel: $(Split-Path $ExcelFile -Leaf)" -ForegroundColor Green

# Verificar template
if (-not (Test-Path $TemplateFile)) {
    Write-Host "❌ Template não encontrado: $TemplateFile" -ForegroundColor Red
    Write-Host "Por favor, coloque o arquivo modelo.pptx em: $ScriptDir\templates\" -ForegroundColor Red
    Read-Host "Pressione Enter para sair"
    exit 1
}

# Verificar se o comando está disponível
try {
    gerar-proposta --help | Out-Null
} catch {
    Write-Host "❌ Comando não encontrado. Reinstalando pacote..." -ForegroundColor Yellow
    pip install -e .
}

# Executar script
Write-Host "🚀 Gerando apresentação (PPTX e PDF)..." -ForegroundColor Green
try {
    python -m proposta_solar.cli `
        --excel $ExcelFile `
        --template $TemplateFile `
        --output $OutputFile `
        --verbose
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ Apresentação gerada com sucesso!" -ForegroundColor Green
        Write-Host "📄 Arquivo PPTX salvo em: $OutputFile" -ForegroundColor Green
        Write-Host "📄 Arquivo PDF salvo em: $($OutputFile -replace '\.pptx$', '.pdf')" -ForegroundColor Green
        
        # Abrir o arquivo no Windows
        Start-Process $OutputFile
    } else {
        Write-Host "❌ Erro ao gerar apresentação" -ForegroundColor Red
        Read-Host "Pressione Enter para sair"
        exit 1
    }
} catch {
    Write-Host "❌ Erro ao executar o script: $_" -ForegroundColor Red
    Read-Host "Pressione Enter para sair"
    exit 1
}

Read-Host "Pressione Enter para sair" 