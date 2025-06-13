#!/bin/bash

# Obter o diretório do script
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Função para ativar o ambiente virtual
activate_venv() {
    if [[ "$OSTYPE" == "darwin"* ]]; then
        source "$SCRIPT_DIR/venv/bin/activate"
    else
        source "$SCRIPT_DIR/venv/Scripts/activate"
    fi
}

# Verificar se Python está instalado
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 não encontrado. Por favor, instale o Python 3."
    exit 1
fi

# Verificar se pip está instalado
if ! command -v pip3 &> /dev/null; then
    echo "❌ Pip não encontrado. Por favor, instale o pip."
    exit 1
fi

# Remover ambiente virtual antigo se existir e estiver corrompido
if [ -d "$SCRIPT_DIR/venv" ]; then
    if ! activate_venv 2>/dev/null; then
        echo "🔧 Ambiente virtual corrompido, recriando..."
        rm -rf "$SCRIPT_DIR/venv"
    fi
fi

# Criar ambiente virtual se não existir
if [ ! -d "$SCRIPT_DIR/venv" ]; then
    echo "🔧 Criando ambiente virtual..."
    python3 -m venv "$SCRIPT_DIR/venv"
    activate_venv
    
    echo "📦 Instalando dependências..."
    pip install --upgrade pip
    pip install -e .
else
    activate_venv
fi

# Criar diretórios necessários
mkdir -p "$SCRIPT_DIR/input"
mkdir -p "$SCRIPT_DIR/output"
mkdir -p "$SCRIPT_DIR/templates"

# Encontrar arquivo Excel mais recente
EXCEL_FILE=$(ls -t "$SCRIPT_DIR/input"/*.xlsx 2>/dev/null | head -n1)
TEMPLATE_FILE="$SCRIPT_DIR/templates/modelo.pptx"
OUTPUT_FILE="$SCRIPT_DIR/output/proposta_$(date +%Y%m%d_%H%M%S).pptx"

# Verificar Excel
if [ -z "$EXCEL_FILE" ]; then
    echo "❌ Nenhum arquivo Excel (.xlsx) encontrado na pasta input/"
    echo "Por favor, coloque seu arquivo Excel em: $SCRIPT_DIR/input/"
    exit 1
fi

echo "📊 Usando arquivo Excel: $(basename "$EXCEL_FILE")"

# Verificar template
if [ ! -f "$TEMPLATE_FILE" ]; then
    echo "❌ Template não encontrado: $TEMPLATE_FILE"
    echo "Por favor, coloque o arquivo modelo.pptx em: $SCRIPT_DIR/templates/"
    exit 1
fi

# Verificar se o comando está disponível
if ! command -v gerar-proposta &> /dev/null; then
    echo "❌ Comando não encontrado. Reinstalando pacote..."
    pip install -e .
fi

# Executar script
echo "🚀 Gerando apresentação..."
python -m proposta_solar.cli \
    --excel "$EXCEL_FILE" \
    --template "$TEMPLATE_FILE" \
    --output "$OUTPUT_FILE" \
    --verbose

# Verificar resultado
    if [ $? -eq 0 ]; then
    echo "✅ Apresentação gerada com sucesso!"
    echo "📄 Arquivo salvo em: $OUTPUT_FILE"
    
    # Abrir o arquivo no macOS
    if [[ "$OSTYPE" == "darwin"* ]]; then
        open "$OUTPUT_FILE"
    fi
else
    echo "❌ Erro ao gerar apresentação"
    exit 1
fi 