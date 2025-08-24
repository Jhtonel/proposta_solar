# Guia de Instalação - Windows

Este guia fornece instruções detalhadas para instalar e executar o sistema Proposta Solar no Windows.

## Pré-requisitos

### 1. Python 3.8 ou superior
- Baixe o Python em: https://www.python.org/downloads/
- **IMPORTANTE**: Durante a instalação, marque a opção "Add Python to PATH"
- Para verificar se foi instalado corretamente, abra o Prompt de Comando e digite:
  ```cmd
  python --version
  ```

### 2. Git (opcional, mas recomendado)
- Baixe o Git em: https://git-scm.com/download/win
- Permite clonar o repositório diretamente

### 3. Software para conversão PDF (opcional)
Para gerar arquivos PDF, você precisa de um dos seguintes:
- **Microsoft PowerPoint** (recomendado) - Já incluído no Office
- **LibreOffice** (alternativa gratuita) - Baixe em https://www.libreoffice.org/download/download/

## Métodos de Instalação

### Método 1: Execução Simples (Recomendado)

1. **Baixe o projeto**:
   - Se você tem Git: `git clone [URL_DO_REPOSITÓRIO]`
   - Ou baixe o ZIP do projeto e extraia

2. **Navegue até a pasta**:
   ```cmd
   cd proposta_solar
   ```

3. **Execute o arquivo .bat**:
   - Duplo clique em `gerar_proposta.bat`
   - Ou via linha de comando: `gerar_proposta.bat`

O script irá:
- Verificar se o Python está instalado
- Criar um ambiente virtual automaticamente
- Instalar todas as dependências
- Criar as pastas necessárias
- Executar o sistema
- **Gerar tanto PPTX quanto PDF**
- **Formatar automaticamente valores monetários em R$**

### Método 2: PowerShell (Alternativa Moderna)

1. **Abra o PowerShell como Administrador**

2. **Configure a política de execução** (se necessário):
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

3. **Execute o script PowerShell**:
   ```powershell
   .\gerar_proposta.ps1
   ```

### Método 3: Instalação Manual

1. **Abra o Prompt de Comando**

2. **Navegue até a pasta do projeto**:
   ```cmd
   cd caminho\para\proposta_solar
   ```

3. **Crie um ambiente virtual**:
   ```cmd
   python -m venv venv
   ```

4. **Ative o ambiente virtual**:
   ```cmd
   venv\Scripts\activate
   ```

5. **Instale as dependências**:
   ```cmd
   pip install -r requirements.txt
   pip install -e .
   ```

6. **Execute o sistema**:
   ```cmd
   python -m proposta_solar.cli
   ```

## Uso

### Preparação dos Arquivos

1. **Arquivo Excel**:
   - Coloque seu arquivo `.xlsx` na pasta `input/`
   - O sistema usará automaticamente o arquivo mais recente

2. **Template PowerPoint**:
   - Coloque seu arquivo `modelo.pptx` na pasta `templates/`

### Execução

1. **Execute um dos scripts**:
   - `gerar_proposta.bat` (duplo clique)
   - `gerar_proposta.ps1` (PowerShell)
   - Ou via linha de comando: `python -m proposta_solar.cli`

2. **Resultado**:
   - A apresentação PPTX será gerada na pasta `output/`
   - A apresentação PDF será gerada na pasta `output/`
   - O arquivo PPTX será aberto automaticamente
   - **Todos os valores monetários serão formatados em R$**

### Opções de Linha de Comando

```cmd
# Gerar PPTX e PDF (padrão)
python -m proposta_solar.cli --excel input/dados.xlsx --template templates/modelo.pptx --output output/proposta.pptx

# Gerar apenas PPTX (sem PDF)
python -m proposta_solar.cli --excel input/dados.xlsx --template templates/modelo.pptx --output output/proposta.pptx --no-pdf

# Com logs detalhados
python -m proposta_solar.cli --excel input/dados.xlsx --template templates/modelo.pptx --output output/proposta.pptx --verbose
```

## Conversão para PDF

O sistema gera automaticamente arquivos PDF usando:

### Método Principal (Windows)
- **Microsoft PowerPoint** - Se você tem o Office instalado
- O sistema detecta automaticamente e usa o PowerPoint

### Método Alternativo
- **LibreOffice** - Software gratuito e de código aberto
- Baixe em: https://www.libreoffice.org/download/download/
- Instale normalmente e o sistema detectará automaticamente

### Verificação
Para verificar se a conversão PDF está funcionando:
1. Execute o sistema normalmente
2. Verifique se dois arquivos foram criados na pasta `output/`:
   - `proposta_YYYYMMDD_HHMMSS.pptx`
   - `proposta_YYYYMMDD_HHMMSS.pdf`

## Formatação Automática

### **O que é Formatado**
O sistema detecta automaticamente e formata as seguintes variáveis:

#### **Valores Monetários**
- `valor_total` → R$ 30.178,57
- `valor_total_1` → R$ 30.178,57
- `valor_total_3` → R$ 30.178,57
- `a_vista` → R$ 30.178,57
- `a_vista_1` → R$ 30.178,57
- `a_vista_3` → R$ 30.178,57

#### **Parcelas**
- `parcela1x` → R$ 2.514,88
- `parcela2x` → R$ 1.257,44
- `parcela3x` → R$ 838,29
- ... até `parcela18x`

#### **Financiamentos**
- `fin12` → R$ 2.514,88
- `fin24` → R$ 1.257,44
- `fin36` → R$ 838,29
- ... até `fin96`

#### **Economias e Gastos**
- `economia_5_anos` → R$ 15.000,00
- `gasto_5_anos` → R$ 25.000,00
- `saldo_anual_rs` → R$ 3.600,00

#### **Datas**
- `data` → 15/01/2024
- `date` → 15/01/2024
- `dia` → 15/01/2024
- `mes` → 15/01/2024
- `ano` → 15/01/2024
- `periodo` → 15/01/2024
- `inicio` → 15/01/2024
- `fim` → 15/01/2024
- `validade` → 15/01/2024
- `vencimento` → 15/01/2024
- `prazo` → 15/01/2024
- `duracao` → 15/01/2024

### **Formatos Aplicados**

#### **Moeda (R$)**
- **Símbolo**: R$ (com espaço)
- **Separador de milhares**: Ponto (.)
- **Separador decimal**: Vírgula (,)
- **Casas decimais**: 2 dígitos
- **Valores negativos**: -R$ 500,00

#### **Data (dd/mm/aaaa)**
- **Dia**: 2 dígitos (01-31)
- **Mês**: 2 dígitos (01-12)
- **Ano**: 4 dígitos (2024)
- **Separadores**: Barra (/)
- **Formato**: dd/mm/aaaa

### **Formatos de Entrada Suportados**

#### **Para Moeda**
- Números: `30178.57`
- Decimais: `30178.57142857142`
- Strings numéricas: `"30178.57"`

#### **Para Data**
- `dd/mm/aaaa` → `dd/mm/aaaa` (mantém formato)
- `dd/mm/aa` → `dd/mm/aaaa` (expande ano)
- `aaaa-mm-dd` → `dd/mm/aaaa`
- `aa-mm-dd` → `dd/mm/aaaa`
- `dd-mm-aaaa` → `dd/mm/aaaa`
- `dd-mm-aa` → `dd/mm/aaaa`
- `dd.mm.aaaa` → `dd/mm/aaaa`
- `dd.mm.aa` → `dd/mm/aaaa`
- `aaaa/mm/dd` → `dd/mm/aaaa`
- `aa/mm/dd` → `dd/mm/aaaa`
- **Datas do Excel** (números) → `dd/mm/aaaa`

### **Exemplos de Transformação**

#### **Moeda**
```
Antes: 30178.57142857142
Depois: R$ 30.178,57

Antes: 2514.880952380952
Depois: R$ 2.514,88
```

#### **Data**
```
Antes: 2024-01-15
Depois: 15/01/2024

Antes: 15/01/24
Depois: 15/01/2024

Antes: 15.01.2024
Depois: 15/01/2024

Antes: 45295 (data do Excel)
Depois: 15/01/2024
```

## Solução de Problemas

### Erro: "Python não é reconhecido como comando"
- **Solução**: Reinstale o Python marcando "Add Python to PATH"
- Ou adicione manualmente o Python ao PATH do sistema

### Erro: "pip não é reconhecido como comando"
- **Solução**: Reinstale o Python ou execute:
  ```cmd
  python -m pip install --upgrade pip
  ```

### Erro de Política de Execução no PowerShell
- **Solução**: Execute como administrador:
  ```powershell
  Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
  ```

### Erro de Codificação (caracteres estranhos)
- **Solução**: Execute o script em um terminal com suporte a UTF-8
- Ou use o PowerShell em vez do Prompt de Comando

### Erro: "Módulo não encontrado"
- **Solução**: Certifique-se de que o ambiente virtual está ativado:
  ```cmd
  venv\Scripts\activate
  ```

### Erro: "Arquivo não encontrado"
- **Solução**: Verifique se os arquivos estão nas pastas corretas:
  - Excel: `input/arquivo.xlsx`
  - Template: `templates/modelo.pptx`

### PDF não é gerado
- **Solução 1**: Instale o Microsoft PowerPoint (Office)
- **Solução 2**: Instale o LibreOffice (gratuito)
- **Solução 3**: Use a opção `--no-pdf` para gerar apenas PPTX

### Erro: "Conversão para PDF não disponível"
- **Solução**: Instale a dependência comtypes:
  ```cmd
  pip install comtypes
  ```

### Formatação Monetária Não Funciona
- **Solução 1**: Verifique se o nome da variável contém padrões monetários
- **Solução 2**: Confirme se o valor na planilha é numérico
- **Solução 3**: Use `--verbose` para ver logs detalhados
- **Solução 4**: Verifique se a função `formatar_moeda()` está sendo chamada

### Formatação de Data Não Funciona
- **Solução 1**: Verifique se o nome da variável contém padrões de data
- **Solução 2**: Confirme se o valor na planilha é uma data válida
- **Solução 3**: Verifique se o formato de entrada é suportado
- **Solução 4**: Use `--verbose` para ver logs detalhados
- **Solução 5**: Verifique se a função `formatar_data()` está sendo chamada

### Problemas Gerais de Formatação
- **Solução 1**: Verifique se as funções de formatação estão sendo chamadas
- **Solução 2**: Confirme se o valor original é válido
- **Solução 3**: Verifique se há caracteres especiais no valor
- **Solução 4**: Use `--verbose` para identificar problemas
- **Solução 5**: Verifique se o nome da variável contém padrões reconhecidos

## Estrutura de Pastas

```
proposta_solar/
├── input/              # Coloque seus arquivos Excel aqui
├── output/             # Apresentações geradas (PPTX e PDF)
├── templates/          # Coloque modelo.pptx aqui
├── gerar_proposta.bat  # Script para Windows
├── gerar_proposta.ps1  # Script PowerShell
├── requirements.txt    # Dependências
├── FORMATACAO_MONETARIA.md # Documentação da formatação monetária
└── README.md
```

## Suporte

Se você encontrar problemas:
1. Verifique se o Python 3.8+ está instalado
2. Certifique-se de que os arquivos estão nas pastas corretas
3. Execute o script como administrador se necessário
4. Verifique se não há antivírus bloqueando a execução
5. Para problemas com PDF, instale PowerPoint ou LibreOffice
6. Para problemas de formatação monetária, verifique os logs com `--verbose` 