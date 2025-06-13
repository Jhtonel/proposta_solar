# -*- coding: utf-8 -*-

"""
Este arquivo mapeia todas as variáveis usadas no sistema, servindo como referência
para qualquer alteração necessária na planilha ou no template.

Para adicionar uma nova variável:
1. Adicione a variável na planilha Excel (ex: {{nova_variavel}})
2. Adicione a variável no template PowerPoint (ex: {{nova_variavel}})
3. Adicione o mapeamento aqui em VARIAVEIS_PLANILHA
"""

# Mapeamento de variáveis da planilha para o template
# 'nome_variavel': 'célula_excel'
VARIAVEIS_PLANILHA = {
    'data': 'A2',                    # Data da proposta
    'representante_nome': 'B2',      # Nome do representante
    'cargo': 'C2',                   # Cargo do representante
    'cliente': 'D2',                 # Nome do cliente
    'projeto_tipo': 'E2',            # Tipo do projeto (Residencial, Comercial, etc)
    'modulos': 'F2',                 # Modelo dos módulos
    'modulos_quantidade': 'G2',      # Quantidade de módulos
    'inversor': 'H2',                # Modelo do inversor
    'inversor_quantidade': 'I2',     # Quantidade de inversores
    'potencia': 'J2',                # Potência total do sistema (kWp)
    'geracao_media': 'K2',           # Geração média mensal (kWh)
    'area': 'L2',                    # Área necessária (m²)
    'consumo_geracao': 'M2',         # Relação consumo/geração
    'fluxo_projetado': 'N2',         # Fluxo de caixa projetado
    'payback': 'O2',                 # Tempo de retorno do investimento
    'com_energia_solar': 'P2',       # Texto explicativo sobre geração solar
    'texto_analise_fin': 'Q2',       # Análise financeira
    'producao_mensal': 'R2',         # Texto sobre produção mensal
    'producao_x_consmed': 'S2',      # Relação produção/consumo médio
    'saldo_anual_rs': 'T2',         # Saldo anual em R$
    'valor_total': 'U2',            # Valor total do sistema
    'parcelas': 'V2',               # Condição de parcelamento
    'a_vista': 'W2',                # Valor à vista
    'cenario_atual': 'X2',          # Descrição do cenário atual
    'por_que_aumenta': 'Y2',        # Explicação sobre aumento da energia
    'economia_5_anos': 'Z2',        # Economia em 5 anos
    'economia_10_anos': 'AA2',      # Economia em 10 anos
    'economia_25_anos': 'AB2',      # Economia em 25 anos
    'gasto_5_anos': 'AC2',          # Gasto em 5 anos sem energia solar
    'gasto_10_anos': 'AD2',         # Gasto em 10 anos sem energia solar
    'gasto_25_anos': 'AE2'          # Gasto em 25 anos sem energia solar
}

# Mapeamento de variáveis por slide
VARIAVEIS_SLIDES = {
    0: ['data', 'representante_nome', 'cargo', 'cliente', 'projeto_tipo'],                                                         # Slide 1 (índice 0)
    1: ['cliente', 'cenario_atual', 'gasto_5_anos', 'gasto_10_anos', 'gasto_25_anos'],                                             # Slide 2
    2: ['cliente', 'por_que_aumenta'],                                                                                             # Slide 3
    3: ['cliente', 'producao_mensal'],                                                                                             # Slide 4
    4: ['cliente'],                                                                                                                # Slide 5
    5: ['cliente', 'economia_5_anos', 'economia_10_anos', 'economia_25_anos'],                                                     # Slide 6
    6: ['cliente'],                                                                                                                # Slide 7
    7: ['cliente', 'potencia', 'geracao_media', 'area', 'modulos', 'modulos_quantidade', 'inversor', 'inversor_quantidade'],       # Slide 8
    8: ['cliente'],                                                                                                                # Slide 9
    9: ['cliente', 'gasto_5_anos', 'gasto_10_anos', 'gasto_25_anos', 'economia_5_anos', 'economia_10_anos', 'economia_25_anos'],   # Slide 10
    10: ['cliente', 'valor_total', 'parcelas', 'a_vista'],                                                                         # Slide 11
    11: ['cliente'],                                                                                                               # Slide 12                                                                                                             # Slide 13
}

# Posições dos gráficos nos slides
POSICOES_GRAFICOS = {
    'graph1': {
        'slide': 1,
        'left': 18.54,    # cm
        'top': 8.83,      # cm
        'width': 29.4,    # cm
        'height': 17.02   # cm
    },
    'graph2': {
        'slide': 2,
        'left': 18.54,    # cm
        'top': 8.83,      # cm
        'width': 29.4,    # cm
        'height': 17.02   # cm
    },
    'graph3': {
        'slide': 3,
        'left': 18.54,    # cm
        'top': 8.83,      # cm
        'width': 29.4,    # cm
        'height': 17.02   # cm
    },
    'graph4': {
        'slide': 5,
        'left': 18.54,    # cm
        'top': 8.83,      # cm
        'width': 29.4,    # cm
        'height': 17.02   # cm
    },
    'graph5': {
        'slide': 9,
        'left': 12.66,    # cm
        'top': 11.87,     # cm
        'width': 24.97,    # cm
        'height': 14.46   # cm
    }
}

# Configuração dos gráficos
GRAFICOS = {
    # Gráfico 1: Custo Acumulado sem Energia Solar
    'graph1': {
        'titulo': 'Custo Acumulado S. Energia Solar',
        'ranges': {
            'ano': 'B5:B29',         # Anos (1 a 25)
            'valor': 'C5:C29'        # Valores acumulados por ano
        },
        'eixos': {
            'x': 'Anos',
            'y': 'Valor (R$)'
        },
        'tipo': 'barras',            # Gráfico de barras
        'cores': {
            'valor': 'vermelho'
        }
    },
    
    # Gráfico 2: Evolução da Conta de Energia
    'graph2': {
        'titulo': 'Evolução da Conta Média/Mês',
        'ranges': {
            'ano': 'B5:B29',         # Anos (1 a 25)
            'valor': 'E5:E29'        # Valores mensais por ano
        },
        'eixos': {
            'x': 'Anos',
            'y': 'Valor (R$)'
        },
        'tipo': 'linha',
        'cores': {
            'valor': 'vermelho'
        }
    },
    
    # Gráfico 3: Produção vs Consumo Mensal
    'graph3': {
        'titulo': 'Produção Mensal x Consumo Médio',
        'ranges': {
            'mes': 'G5:G16',         # Meses do ano
            'producao': 'H5:H16',    # Produção mensal
            'consumo': 'I5:I16'      # Consumo mensal
        },
        'eixos': {
            'x': 'Meses',
            'y': 'kWh'
        },
        'tipo': 'barras_duplas',
        'cores': {
            'producao': 'verde',
            'consumo': 'vermelho'
        }
    },
    
    # Gráfico 4: Fluxo de Caixa
    'graph4': {
        'titulo': 'Fluxo de Caixa',
        'ranges': {
            'ano': 'K5:K29',         # Anos
            'positivo': 'L5:L29',    # Valores positivos
            'negativo': 'M5:M29'     # Valores negativos
        },
        'eixos': {
            'x': 'Anos',
            'y': 'Valor (R$)'
        },
        'tipo': 'barras',           # Gráfico de barras
        'cores': {
            'positivo': 'verde',
            'negativo': 'vermelho'
        }
    },
    
    # Gráfico 5: Comparação Com vs Sem Energia Solar
    'graph5': {
        'titulo': 'Comparação C. Energia Solar x Sem Energia Solar',
        'ranges': {
            'ano': 'O5:O29',         # Anos
            'economia': 'P5:P29',    # Economia acumulada
            'custo': 'Q5:Q29'        # Custos acumulados
        },
        'eixos': {
            'x': 'Anos',
            'y': 'Valor (R$)'
        },
        'tipo': 'linhas',
        'cores': {
            'economia': 'verde',
            'custo': 'vermelho'
        }
    }
}

# Configurações de estilo dos gráficos
ESTILO_GRAFICOS = {
    # Dimensões do gráfico em centímetros
    'dimensoes': {
        'largura': 29.4,            # Largura em cm
        'altura': 17.02,            # Altura em cm
        'cm_para_polegadas': 0.393701  # Fator de conversão cm -> polegadas
    },
    
    # Cores padrão
    'cores': {
        'vermelho': '#FF0000',      # Cor para valores negativos/custos
        'verde': '#00FF00'          # Cor para valores positivos/economia
    },
    
    # Configurações gerais
    'geral': {
        'dpi': 300,                 # Resolução da imagem
        'tamanho_fonte': 10,        # Tamanho da fonte
        'espessura_linha': 2,       # Espessura das linhas
        'mostrar_grade': False,     # Se deve mostrar linhas de grade
        'mostrar_borda': {          # Quais bordas mostrar
            'topo': False,
            'direita': False,
            'inferior': True,
            'esquerda': True
        },
        'padding_titulo': 20        # Espaço entre título e gráfico
    }
} 