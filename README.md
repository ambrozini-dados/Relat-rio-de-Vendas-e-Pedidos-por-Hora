# Relatório de Vendas e Pedidos por Hora (Python + Streamlit)

Dashboard interativo em **Streamlit** para análise de vendas, volume de pedidos e cancelamentos por faixa horária, replicando e expandindo a estrutura visual do painel analítico.

## Funcionalidades

- **Filtros com Seleção Múltipla (`multiselect`)**:
  - **Dia da Semana** (`segunda-feira` a `domingo`)
  - **Loja** (unidades físicas extraídas automaticamente de `NOME DA LOJA`: *Alípio de Melo, Barreiro, Estoril, Floresta, Mantiqueira, Planalto*)
  - **Marca** (nomes completos das lojas/operações no iFood, filtrados dinamicamente conforme as lojas marcadas)
  - **Período** e **Faixa Horária**
- **Leitura Inteligente do Excel iFood**: Corrige automaticamente metadados de dimensão (`reset_dimensions()`) das exportações `.xlsx` do portal para ler todos os registros sem perda de linhas.
- **Visualizações**:
  - Gráfico de barras agrupadas **Média de pedidos e cancelados por hora** (tema escuro com rótulos de dados no topo).
  - Gráfico combinado de **Média de Vendas por Hora (R$) e Ticket Médio**.
  - **Mapa de Calor (Dia da Semana × Hora)** para identificação rápida de picos operacionais.
  - Tabela consolidada com botão de exportação para `.xlsx`.

## Como Executar

1. Instale as dependências:
```bash
pip install -r requirements.txt
```

2. Coloque a planilha Excel na mesma pasta (ou faça upload diretamente pela barra lateral do app) e execute:
```bash
streamlit run app.py
```
