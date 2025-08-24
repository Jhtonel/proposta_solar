# Formatação Automática - Proposta Solar

Este documento explica como o sistema formata automaticamente valores monetários e datas para os formatos brasileiros padrão.

## 🎯 Funcionalidades Implementadas

O sistema agora formata automaticamente:
- **Valores monetários** para o formato brasileiro: **R$ X.XXX,XX**
- **Datas** para o formato brasileiro: **dd/mm/aaaa**

## 💰 Formatação Monetária

### **Variáveis que Recebem Formatação Monetária**
- **Valores Principais**: `valor_total`, `valor_total_1`, `valor_total_3`, `a_vista`, `a_vista_1`, `a_vista_3`
- **Parcelas**: `parcela1x`, `parcela2x`, `parcela3x`... até `parcela18x`
- **Financiamentos**: `fin12`, `fin24`, `fin36`, `fin48`, `fin60`, `fin72`, `fin84`, `fin96`
- **Economias e Gastos**: `economia_5_anos`, `gasto_5_anos`, `saldo_anual_rs`, etc.

### **Formato Aplicado**
```
Antes: 30178.57142857142
Depois: R$ 30.178,57

Antes: 2514.880952380952
Depois: R$ 2.514,88
```

## 📅 Formatação de Data

### **Variáveis que Recebem Formatação de Data**
- **Data Principal**: `data` (data da proposta)
- **Outras Datas**: `date`, `dia`, `mes`, `ano`, `periodo`, `inicio`, `fim`, `validade`, `vencimento`, `prazo`, `duracao`

### **Formato Aplicado**
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

### **Formatos Suportados de Entrada**
O sistema reconhece e converte automaticamente:
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

## 🔧 Como Funciona

### **Detecção Automática**
O sistema identifica automaticamente:

#### **Variáveis Monetárias**
```python
padroes_monetarios = [
    'valor_total', 'a_vista', 'parcela', 'fin', 'economia', 'gasto',
    'saldo', 'fluxo', 'payback', 'custo'
]
```

#### **Variáveis de Data**
```python
padroes_data = [
    'data', 'date', 'dia', 'mes', 'ano', 'periodo', 'inicio', 'fim',
    'validade', 'vencimento', 'prazo', 'duracao'
]
```

### **Formatação Aplicada**
```python
# Para valores monetários
if is_variavel_monetaria(var_name):
    value = formatar_moeda(value)

# Para datas
if is_variavel_data(var_name):
    value = formatar_data(value)
```

## 📊 Exemplos Completos de Uso

### **Formatação Monetária**
```
valor_total: 30178.57142857142 → R$ 30.178,57
a_vista: 30178.57142857142 → R$ 30.178,57
parcela1x: 2514.880952380952 → R$ 2.514,88
fin12: 2514.880952380952 → R$ 2.514,88
economia_5_anos: 15000 → R$ 15.000,00
gasto_5_anos: 25000 → R$ 25.000,00
```

### **Formatação de Data**
```
data: 2024-01-15 → 15/01/2024
data: 15/01/24 → 15/01/2024
data: 15.01.2024 → 15/01/2024
data: 45295 → 15/01/2024 (data do Excel)
data: 15-01-2024 → 15/01/2024
```

## 🎨 Formatos Brasileiros Aplicados

### **Moeda (R$)**
- **Símbolo**: R$ (com espaço)
- **Separador de milhares**: Ponto (.)
- **Separador decimal**: Vírgula (,)
- **Casas decimais**: 2 dígitos
- **Valores negativos**: -R$ 500,00

### **Data (dd/mm/aaaa)**
- **Dia**: 2 dígitos (01-31)
- **Mês**: 2 dígitos (01-12)
- **Ano**: 4 dígitos (2024)
- **Separadores**: Barra (/)
- **Formato**: dd/mm/aaaa

## 🚀 Benefícios

### **1. Apresentação Profissional**
- Valores monetários sempre no formato correto
- Datas sempre no formato brasileiro padrão
- Consistência visual em toda a apresentação

### **2. Facilidade de Leitura**
- Formato familiar para usuários brasileiros
- Separação clara de milhares e datas
- Melhor compreensão dos valores

### **3. Automatização Total**
- Não é necessário formatar manualmente
- Reduz erros de formatação
- Mantém consistência em todas as propostas

### **4. Flexibilidade de Entrada**
- Aceita múltiplos formatos de data
- Converte automaticamente datas do Excel
- Tratamento robusto de erros

## ⚙️ Configuração

### **Padrões Detectados Automaticamente**
- `valor_total` → ✅ Formatação monetária
- `a_vista` → ✅ Formatação monetária
- `parcela1x` → ✅ Formatação monetária
- `data` → ✅ Formatação de data
- `cliente` → ❌ Sem formatação especial
- `representante_nome` → ❌ Sem formatação especial

### **Personalização**
Para adicionar novos padrões, edite as funções em `presentation.py`:

```python
# Para moeda
padroes_monetarios = [
    'valor_total', 'a_vista', 'parcela', 'fin', 'economia', 'gasto',
    'saldo', 'fluxo', 'payback', 'custo', 'novo_padrao'  # ← Adicione aqui
]

# Para data
padroes_data = [
    'data', 'date', 'dia', 'mes', 'ano', 'periodo', 'inicio', 'fim',
    'validade', 'vencimento', 'prazo', 'duracao', 'nova_data'  # ← Adicione aqui
]
```

## 🔍 Verificação

### **Como Verificar se Está Funcionando**
1. Execute o sistema com `--verbose`
2. Procure por logs como:
   ```
   Tentando substituir valor_total com valor: R$ 30.178,57
   Tentando substituir data com valor: 15/01/2024
   ```
3. Verifique a apresentação gerada

### **Logs de Formatação**
O sistema registra quando a formatação é aplicada:
```
INFO: Formatação monetária aplicada a: valor_total
INFO: Formatação de data aplicada a: data
```

## 🛠️ Solução de Problemas

### **Valor Monetário Não Formatado**
- Verifique se o nome da variável contém um dos padrões monetários
- Confirme se o valor na planilha é numérico
- Verifique os logs para identificar problemas

### **Data Não Formatada**
- Verifique se o nome da variável contém um dos padrões de data
- Confirme se o valor na planilha é uma data válida
- Verifique se o formato de entrada é suportado

### **Formato Incorreto**
- Verifique se as funções `formatar_moeda()` ou `formatar_data()` estão sendo chamadas
- Confirme se o valor original é válido
- Verifique se há caracteres especiais no valor

### **Performance**
- A formatação é aplicada apenas uma vez por variável
- Não afeta significativamente o tempo de processamento
- Cache interno para valores já formatados

## 📝 Notas Técnicas

### **Implementação**
- Função `formatar_moeda()` em `presentation.py`
- Função `formatar_data()` em `presentation.py`
- Detecção automática via `is_variavel_monetaria()` e `is_variavel_data()`
- Aplicada durante a substituição de variáveis

### **Compatibilidade**
- Funciona com todos os tipos de dados numéricos para moeda
- Suporta múltiplos formatos de data de entrada
- Converte automaticamente datas do Excel
- Tratamento robusto de erros

### **Extensibilidade**
- Fácil adicionar novos padrões de detecção
- Configurável para diferentes formatos
- Código modular e reutilizável
- Suporte para novos tipos de formatação

## 🎯 Resumo das Funcionalidades

### **✅ Implementado**
- Formatação automática de valores monetários em R$
- Formatação automática de datas em dd/mm/aaaa
- Detecção automática de variáveis
- Suporte a múltiplos formatos de entrada
- Tratamento robusto de erros
- Logs detalhados para debugging

### **🚀 Pronto para Uso**
O sistema agora formata automaticamente:
- **Moeda**: R$ X.XXX,XX
- **Data**: dd/mm/aaaa
- **Detecção**: Automática baseada no nome da variável
- **Flexibilidade**: Múltiplos formatos de entrada suportados
