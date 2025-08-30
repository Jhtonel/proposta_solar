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
    'gasto_25_anos': 'AE2',          # Gasto em 25 anos sem energia solar
    'endereco_cliente': 'AF2',       # Endereço do cliente
    'telefone_cliente': 'AG2',       # Telefone do cliente
    'modulos1': 'AH2',              # Modelo do módulo 1
    'modulos_quantidade1': 'AI2',   # Quantidade de módulos 1
    'inversor1': 'AJ2',            # Modelo do inversor 1
    'inversor_quantidade1': 'AK2', # Quantidade de inversores 1
    'potencia1': 'AL2',           # Potência do sistema 1
    'geracao_media1': 'AM2',      # Geração média mensal 1
    'modulos3': 'AN2',              # Modelo do módulo 3
    'modulos_quantidade3': 'AO2',   # Quantidade de módulos 3
    'inversor3': 'AP2',            # Modelo do inversor 3
    'inversor_quantidade3': 'AQ2', # Quantidade de inversores 3
    'potencia3': 'AR2',           # Potência do sistema 3
    'geracao_media3': 'AS2',      # Geração média mensal 3
    'parcela1x': 'AT2',           # Parcela 1x
    'parcela2x': 'AU2',           # Parcela 2x 
    'parcela3x': 'AV2',           # Parcela 3x
    'parcela4x': 'AW2',           # Parcela 4x
    'parcela5x': 'AX2',           # Parcela 5x
    'parcela6x': 'AY2',           # Parcela 6x
    'parcela7x': 'AZ2',           # Parcela 7x
    'parcela8x': 'BA2',           # Parcela 8x
    'parcela9x': 'BB2',           # Parcela 9x
    'parcela10x': 'BC2',          # Parcela 10x
    'parcela11x': 'BD2',          # Parcela 11x
    'parcela12x': 'BE2',          # Parcela 12x
    'parcela13x': 'BF2',          # Parcela 13x
    'parcela14x': 'BG2',          # Parcela 14x
    'parcela15x': 'BH2',          # Parcela 15x
    'parcela16x': 'BI2',          # Parcela 16x
    'parcela17x': 'BJ2',          # Parcela 17x
    'parcela18x': 'BK2',          # Parcela 18x
    'parcela1x_1': 'BL2',         # Parcela 1x 1
    'parcela2x_1': 'BM2',         # Parcela 2x 1
    'parcela3x_1': 'BN2',         # Parcela 3x 1
    'parcela4x_1': 'BO2',         # Parcela 4x 1
    'parcela5x_1': 'BP2',         # Parcela 5x 1
    'parcela6x_1': 'BQ2',         # Parcela 6x 1
    'parcela7x_1': 'BR2',         # Parcela 7x 1
    'parcela8x_1': 'BS2',         # Parcela 8x 1
    'parcela9x_1': 'BT2',         # Parcela 9x 1
    'parcela10x_1': 'BU2',        # Parcela 10x 1
    'parcela11x_1': 'BV2',        # Parcela 11x 1
    'parcela12x_1': 'BW2',        # Parcela 12x 1
    'parcela13x_1': 'BX2',        # Parcela 13x 1
    'parcela14x_1': 'BY2',        # Parcela 14x 1
    'parcela15x_1': 'BZ2',        # Parcela 15x 1
    'parcela16x_1': 'CA2',        # Parcela 16x 1
    'parcela17x_1': 'CB2',        # Parcela 17x 1
    'parcela18x_1': 'CC2',        # Parcela 18x 1
    'parcela1x_3': 'CD2',         # Parcela 1x 3
    'parcela2x_3': 'CE2',         # Parcela 2x 3
    'parcela3x_3': 'CF2',         # Parcela 3x 3
    'parcela4x_3': 'CG2',         # Parcela 4x 3
    'parcela5x_3': 'CH2',         # Parcela 5x 3
    'parcela6x_3': 'CI2',         # Parcela 6x 3
    'parcela7x_3': 'CJ2',         # Parcela 7x 3
    'parcela8x_3': 'CK2',         # Parcela 8x 3
    'parcela9x_3': 'CL2',         # Parcela 9x 3
    'parcela10x_3': 'CM2',        # Parcela 10x 3
    'parcela11x_3': 'CN2',        # Parcela 11x 3
    'parcela12x_3': 'CO2',        # Parcela 12x 3
    'parcela13x_3': 'CP2',        # Parcela 13x 3
    'parcela14x_3': 'CQ2',        # Parcela 14x 3
    'parcela15x_3': 'CR2',        # Parcela 15x 3
    'parcela16x_3': 'CS2',        # Parcela 16x 3
    'parcela17x_3': 'CT2',        # Parcela 17x 3
    'parcela18x_3': 'CU2',        # Parcela 18x 3
    'fin12_1': 'CV2',             # Financiamento 12x 1
    'fin24_1': 'CW2',             # Financiamento 24x 1
    'fin36_1': 'CX2',             # Financiamento 36x 1
    'fin48_1': 'CY2',             # Financiamento 48x 1
    'fin60_1': 'CZ2',             # Financiamento 60x 1
    'fin72_1': 'DA2',             # Financiamento 12x 3
    'fin84_1': 'DB2',             # Financiamento 24x 3
    'fin96_1': 'DC2',             # Financiamento 36x 3
    'fin12': 'DD2',             # Financiamento 12x 1
    'fin24': 'DE2',             # Financiamento 24x 1
    'fin36': 'DF2',             # Financiamento 36x 1
    'fin48': 'DG2',             # Financiamento 48x 1
    'fin60': 'DH2',             # Financiamento 60x 1
    'fin72': 'DI2',             # Financiamento 72x 1
    'fin84': 'DJ2',             # Financiamento 84x 1
    'fin96': 'DK2',             # Financiamento 96x 1
    'fin12_3': 'DL2',             # Financiamento 12x 3
    'fin24_3': 'DM2',             # Financiamento 24x 3
    'fin36_3': 'DN2',             # Financiamento 36x 3
    'fin48_3': 'DO2',             # Financiamento 48x 3
    'fin60_3': 'DP2',             # Financiamento 60x 3
    'fin72_3': 'DQ2',             # Financiamento 72x 3
    'fin84_3': 'DR2',             # Financiamento 84x 3
    'fin96_3': 'DS2',             # Financiamento 96x 3
    'valor_total_1': 'DT2',             # Valor total 1
    'a_vista_1': 'DU2',             # A vista 1
    'valor_total_3': 'DV2',             # Valor total 3
    'a_vista_3': 'DW2',             # A vista 3
    'creditos': 'DY2',             # Creditos
    'creditos1': 'DX2',             # Creditos 1
    'creditos3': 'DZ2'             # Creditos 3
}

# Mapeamento de variáveis por slide
VARIAVEIS_SLIDES = {
    0: ['data', 'representante_nome', 'cargo', 'cliente', 'projeto_tipo', 'endereco_cliente', 'telefone_cliente'],  
    1: ['cliente'],
    2: ['cliente', 'cenario_atual', 'gasto_5_anos', 'gasto_10_anos', 'gasto_25_anos'],                                             
    3: ['cliente', 'por_que_aumenta'],                                                                                          
    4: ['cliente', 'com_energia_solar'],                                                                                             
    5: ['cliente', 'economia_5_anos', 'economia_10_anos', 'economia_25_anos'],                                                   
    6: ['cliente', 'modulos1', 'modulos_quantidade1', 'inversor1', 'inversor_quantidade1', 'potencia1', 'geracao_media1', 'modulos','modulos_quantidade', 'inversor', 'inversor_quantidade', 'potencia', 'geracao_media', 'modulos3', 'modulos_quantidade3', 'inversor3', 'inversor_quantidade3', 'potencia3', 'geracao_media3', 'creditos', 'creditos1', 'creditos3'],                                                                                                                # Slide 6
    7: ['cliente'],
    8: ['cliente'],
    9: ['cliente', 'gasto_5_anos', 'gasto_10_anos', 'gasto_25_anos', 'economia_5_anos', 'economia_10_anos', 'economia_25_anos'],   
    10: ['cliente', 'valor_total_1', 'a_vista_1', 'parcela1x_1', 'parcela2x_1', 'parcela3x_1', 'parcela4x_1', 'parcela5x_1', 'parcela6x_1', 'parcela7x_1', 'parcela8x_1', 'parcela9x_1', 'parcela10x_1', 'parcela11x_1', 'parcela12x_1', 'parcela13x_1', 'parcela14x_1', 'parcela15x_1', 'parcela16x_1', 'parcela17x_1', 'parcela18x_1', 'fin12_1', 'fin24_1', 'fin36_1', 'fin48_1', 'fin60_1', 'fin72_1', 'fin84_1', 'fin96_1'],                                                                         
    11: ['cliente', 'valor_total', 'a_vista', 'parcela1x', 'parcela2x', 'parcela3x', 'parcela4x', 'parcela5x', 'parcela6x', 'parcela7x', 'parcela8x', 'parcela9x', 'parcela10x', 'parcela11x', 'parcela12x', 'parcela13x', 'parcela14x', 'parcela15x', 'parcela16x', 'parcela17x', 'parcela18x', 'fin12', 'fin24', 'fin36', 'fin48', 'fin60', 'fin72', 'fin84', 'fin96'],                                                                                                               
    12: ['cliente', 'valor_total_3', 'a_vista_3', 'parcela1x_3', 'parcela2x_3', 'parcela3x_3', 'parcela4x_3', 'parcela5x_3', 'parcela6x_3', 'parcela7x_3', 'parcela8x_3', 'parcela9x_3', 'parcela10x_3', 'parcela11x_3', 'parcela12x_3', 'parcela13x_3', 'parcela14x_3', 'parcela15x_3', 'parcela16x_3', 'parcela17x_3', 'parcela18x_3', 'fin12_3', 'fin24_3', 'fin36_3', 'fin48_3', 'fin60_3', 'fin72_3', 'fin84_3', 'fin96_3'],                                                                                                               
    13: ['cliente'],                                                                                                               
}

# Posições dos gráficos nos slides
POSICOES_GRAFICOS = {
    'graph1': {
        'slide': 2,
        'left': 18.54,    # cm
        'top': 8.83,      # cm
        'width': 29.4,    # cm
        'height': 17.02   # cm
    },
    'graph2': {
        'slide': 3,
        'left': 18.54,    # cm
        'top': 8.83,      # cm
        'width': 29.4,    # cm
        'height': 17.02   # cm
    },
    'graph3': {
        'slide': 4,
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
        'vermelho': '#DC2626',      # Cor para valores negativos/custos
        'verde': '#059669'          # Cor para valores positivos/economia
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