# Proposta Solar

Sistema para geração automática de propostas comerciais para projetos de energia solar.

## Funcionalidades

- Leitura de dados de planilha Excel
- Geração de gráficos personalizados
- Substituição automática de variáveis em template PowerPoint
- Geração de apresentação final em PowerPoint (PPTX)
- **Conversão automática para PDF**
- **Formatação automática de valores monetários em R$**
- Suporte multiplataforma (Windows, macOS, Linux)

## Requisitos

- Python 3.8+
- Dependências listadas em `requirements.txt`
- **Para conversão PDF**: PowerPoint (Windows) ou LibreOffice (todas as plataformas)

## Instalação

### Windows

#### Opção 1: Usando o arquivo .bat (Recomendado)
1. Clone o repositório:
```cmd
git clone [URL_DO_REPOSITÓRIO]
cd proposta_solar
```

2. Execute o arquivo `gerar_proposta.bat` clicando duas vezes nele ou via linha de comando:
```cmd
gerar_proposta.bat
```

#### Opção 2: Usando PowerShell
1. Clone o repositório:
```powershell
git clone [URL_DO_REPOSITÓRIO]
cd proposta_solar
```

2. Execute o script PowerShell:
```powershell
.\gerar_proposta.ps1
```

**Nota:** Se você receber um erro de política de execução no PowerShell, execute:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### Opção 3: Instalação manual
1. Clone o repositório:
```cmd
git clone [URL_DO_REPOSITÓRIO]
cd proposta_solar
```

2. Crie e ative um ambiente virtual:
```cmd
python -m venv venv
venv\Scripts\activate
```

3. Instale as dependências:
```cmd
pip install -r requirements.txt
pip install -e .
```

### Linux/Mac

1. Clone o repositório:
```bash
git clone [URL_DO_REPOSITÓRIO]
cd proposta_solar
```

2. Execute o arquivo .command:
```bash
./gerar_proposta.command
```

3. Ou instalação manual:
```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
pip install -e .
```

## Uso

### Windows
1. Coloque seu arquivo Excel na pasta `input/`
2. Coloque seu template PowerPoint na pasta `templates/`
3. Execute um dos arquivos:
   - `gerar_proposta.bat` (duplo clique)
   - `gerar_proposta.ps1` (PowerShell)
   - Ou via linha de comando: `python -m proposta_solar.cli`

### Linux/Mac
1. Coloque seu arquivo Excel na pasta `input/`
2. Coloque seu template PowerPoint na pasta `templates/`
3. Execute: `./gerar_proposta.command` ou `python -m proposta_solar.cli`

### Opções de Linha de Comando

```bash
# Gerar PPTX e PDF (padrão)
python -m proposta_solar.cli --excel input/dados.xlsx --template templates/modelo.pptx --output output/proposta.pptx

# Gerar apenas PPTX (sem PDF)
python -m proposta_solar.cli --excel input/dados.xlsx --template templates/modelo.pptx --output output/proposta.pptx --no-pdf

# Com logs detalhados
python -m proposta_solar.cli --excel input/dados.xlsx --template templates/modelo.pptx --output output/proposta.pptx --verbose
```

## Conversão para PDF

O sistema gera automaticamente tanto o arquivo PPTX quanto o PDF. A conversão usa:

### Windows
- **PowerPoint** (recomendado) - Requer Microsoft PowerPoint instalado
- **LibreOffice** (alternativa) - Se PowerPoint não estiver disponível

### Linux/Mac
- **LibreOffice** - Requer LibreOffice instalado

### Instalação do LibreOffice
- **Windows**: Baixe em https://www.libreoffice.org/download/download/
- **macOS**: `brew install --cask libreoffice`
- **Ubuntu/Debian**: `sudo apt install libreoffice`
- **CentOS/RHEL**: `sudo yum install libreoffice`

## Formatação Automática

O sistema formata automaticamente todos os valores monetários e datas para os formatos brasileiros padrão:

### **Formatação Monetária**
- **Antes**: `30178.57142857142`
- **Depois**: `R$ 30.178,57`

### **Formatação de Data**
- **Antes**: `2024-01-15`
- **Depois**: `15/01/2024`

### **Variáveis Formatadas**

#### **Valores Monetários**
- Valores totais: `valor_total`, `valor_total_1`, `valor_total_3`
- Valores à vista: `a_vista`, `a_vista_1`, `a_vista_3`
- Parcelas: `parcela1x`, `parcela2x`, `parcela3x`... até `parcela18x`
- Financiamentos: `fin12`, `fin24`, `fin36`, `fin48`, `fin60`, `fin72`, `fin84`, `fin96`
- Economias e gastos: `economia_5_anos`, `gasto_5_anos`, etc.

#### **Datas**
- Data principal: `data` (data da proposta)
- Outras datas: `date`, `dia`, `mes`, `ano`, `periodo`, `inicio`, `fim`, `validade`, `vencimento`, `prazo`, `duracao`

### **Formatos Suportados de Entrada**
- **Moeda**: Números, decimais, strings numéricas
- **Data**: dd/mm/aaaa, dd-mm-aaaa, aaaa-mm-dd, datas do Excel, múltiplos formatos

### **Características**
- **Moeda**: R$ com separador de milhares (.) e decimal (,)
- **Data**: dd/mm/aaaa com barras como separadores
- **Detecção**: Automática baseada no nome da variável
- **Flexibilidade**: Múltiplos formatos de entrada suportados

## Estrutura do Projeto

```
proposta_solar/
├── input/                    # Arquivos Excel de entrada
├── output/                   # Apresentações geradas (PPTX e PDF)
├── templates/                # Templates PowerPoint
├── src/                      # Código fonte Python
│   └── proposta_solar/
│       ├── __init__.py
│       ├── cli.py           # Interface de linha de comando
│       ├── presentation.py  # Lógica principal
│       └── variaveis.py     # Mapeamento de variáveis
├── venv/                     # Ambiente virtual Python
├── gerar_proposta.command   # Script para macOS/Linux
├── gerar_proposta.bat       # Script para Windows
├── gerar_proposta.ps1       # Script PowerShell
├── requirements.txt          # Dependências Python
├── setup.py                 # Configuração do pacote
├── FORMATACAO_MONETARIA.md  # Documentação da formatação automática
├── SOLUCAO_GRAFICO_CONSUMO.md # Solução para problema do gráfico
└── README.md                # Este arquivo
```

## Solução de Problemas

### **Erros Comuns**

#### **Erro no Slide 11**
- **Sintoma**: "Falha ao substituir todas as variáveis no Slide 11"
- **Causa**: Variáveis não encontradas ou mapeamento incorreto
- **Solução**: Verificar se todas as variáveis estão definidas em `variaveis.py`

#### **Gráfico de Consumo Não Exibe Barras Vermelhas**
- **Sintoma**: Gráfico "Produção Mensal x Consumo Médio" mostra apenas barras verdes
- **Causa**: Problema na leitura ou criação das barras de consumo
- **Solução**: Verificar se as células `I5:I16` contêm dados de consumo válidos
- **Documentação**: Ver `SOLUCAO_GRAFICO_CONSUMO.md`

#### **PDF não gerado**
- **Sintoma**: Apenas arquivo PPTX é criado
- **Causa**: PowerPoint ou LibreOffice não disponível
- **Solução**: Instalar Microsoft PowerPoint ou LibreOffice

#### **Permissão negado**
- **Linux/Mac**: Execute `chmod +x gerar_proposta.command`
- **Windows**: Execute como administrador se necessário

#### **Python não encontrado**
- **Linux/Mac**: Instale o Python 3.8+ via gerenciador de pacotes
- **Windows**: Instale Python 3.8+ e adicione ao PATH

#### **Dependências não encontradas**
- **Solução**: Execute `pip install -r requirements.txt`
- **Ambiente virtual**: Ative o ambiente virtual primeiro

### **Formatação Monetária**
- **Valor não formatado**: Verifique se o nome da variável contém padrões monetários
- **Formato incorreto**: Confirme se o valor na planilha é numérico
- **Verificar funcionamento**: Use `--verbose` e procure por logs de formatação

### **Formatação de Data**
- **Data não formatada**: Verifique se o nome da variável contém padrões de data
- **Formato incorreto**: Confirme se o valor na planilha é uma data válida
- **Verificar funcionamento**: Use `--verbose` e procure por logs de formatação

## Contribuição

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/nova-feature`)
3. Commit suas mudanças (`git commit -m 'Adiciona nova feature'`)
4. Push para a branch (`git push origin feature/nova-feature`)
5. Abra um Pull Request

## Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes. 