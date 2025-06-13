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

1. Clone o repositório:
```bash
git clone [URL_DO_REPOSITÓRIO]
cd proposta_solar
```

2. Crie e ative um ambiente virtual:
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
# ou
.\venv\Scripts\activate  # Windows
```

3. Instale as dependências:
```bash
pip install -r requirements.txt
```

## Uso

1. Coloque seu arquivo Excel na pasta `input/`
2. Coloque seu template PowerPoint na pasta `templates/`
3. Execute o script:
```bash
python -m proposta_solar.cli
```

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
├── requirements.txt
└── README.md
```

## Contribuição

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/nova-feature`)
3. Commit suas mudanças (`git commit -m 'Adiciona nova feature'`)
4. Push para a branch (`git push origin feature/nova-feature`)
5. Abra um Pull Request

## Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes. 