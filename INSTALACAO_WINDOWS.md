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
   - A apresentação será gerada na pasta `output/`
   - O arquivo será aberto automaticamente

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

## Estrutura de Pastas

```
proposta_solar/
├── input/              # Coloque seus arquivos Excel aqui
├── output/             # Apresentações geradas
├── templates/          # Coloque modelo.pptx aqui
├── gerar_proposta.bat  # Script para Windows
├── gerar_proposta.ps1  # Script PowerShell
└── requirements.txt    # Dependências
```

## Suporte

Se você encontrar problemas:
1. Verifique se o Python 3.8+ está instalado
2. Certifique-se de que os arquivos estão nas pastas corretas
3. Execute o script como administrador se necessário
4. Verifique se não há antivírus bloqueando a execução 