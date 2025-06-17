# Proposta Solar

Sistema para geração automática de propostas comerciais para projetos de energia solar.

## Funcionalidades

- Leitura de dados de planilha Excel
- Geração de gráficos personalizados
- Substituição automática de variáveis em template PowerPoint
- Geração de apresentação final em PowerPoint

## Requisitos

- Python 3.8+
- Dependências listadas em `requirements.txt`

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

## Estrutura do Projeto

```
proposta_solar/
├── input/              # Arquivos Excel de entrada
├── output/             # Apresentações geradas
├── templates/          # Templates PowerPoint
├── src/
│   └── proposta_solar/
│       ├── __init__.py
│       ├── cli.py
│       ├── presentation.py
│       └── variaveis.py
├── gerar_proposta.command  # Script para Linux/Mac
├── gerar_proposta.bat      # Script para Windows
├── gerar_proposta.ps1      # Script PowerShell para Windows
├── requirements.txt
├── setup.py
└── README.md
```

## Solução de Problemas

### Windows
- **Erro de política de execução no PowerShell**: Execute `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`
- **Python não encontrado**: Certifique-se de que o Python está instalado e adicionado ao PATH
- **Erro de codificação**: Execute o script em um terminal com suporte a UTF-8

### Linux/Mac
- **Permissão negada**: Execute `chmod +x gerar_proposta.command`
- **Python não encontrado**: Instale o Python 3.8+ via gerenciador de pacotes

## Contribuição

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/nova-feature`)
3. Commit suas mudanças (`git commit -m 'Adiciona nova feature'`)
4. Push para a branch (`git push origin feature/nova-feature`)
5. Abra um Pull Request

## Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes. 