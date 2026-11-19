import io
import re
import unicodedata
import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA & ESTILO DARK (PADRÃO DASHBOARD EXECUTIVO)
# ==============================================================================
st.set_page_config(
    page_title="Relatório de Vendas e Pedidos por Hora",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .stApp {
            background-color: #121212;
            color: #F3F4F6;
        }
        [data-testid="stSidebar"] {
            background-color: #191919;
            border-right: 1px solid #2A2A2A;
        }
        div[data-testid="stMetric"] {
            background-color: #1E1E1E;
            border: 1px solid #2E2E2E;
            border-radius: 10px;
            padding: 14px 18px;
        }
        div[data-testid="stMetricLabel"] p {
            color: #A0AEC0 !important;
            font-size: 0.85rem !important;
            font-weight: 600 !important;
        }
        div[data-testid="stMetricValue"] {
            color: #FFFFFF !important;
            font-size: 1.55rem !important;
            font-weight: 700 !important;
        }
        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 2rem;
        }
        h1, h2, h3 {
            color: #F9FAFB !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==============================================================================
# FUNÇÕES DE TRATAMENTO E EXTRAÇÃO DE DADOS
# ==============================================================================
DIAS_SEMANA_MAP = {
    0: "segunda-feira",
    1: "terça-feira",
    2: "quarta-feira",
    3: "quinta-feira",
    4: "sexta-feira",
    5: "sábado",
    6: "domingo",
}
ORDEM_DIAS_SEMANA = [
    "segunda-feira",
    "terça-feira",
    "quarta-feira",
    "quinta-feira",
    "sexta-feira",
    "sábado",
    "domingo",
]

UNIDADES_PADRAO = {
    "alipio de melo": "Alípio de Melo",
    "barreiro": "Barreiro",
    "estoril": "Estoril",
    "floresta": "Floresta",
    "mantiqueira": "Mantiqueira",
    "planalto": "Planalto",
}


def remover_acentos(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn"
    )


def extrair_loja(nome_loja_bruto: str) -> str:
    """Extrai o nome padronizado da unidade (Loja) a partir da coluna NOME DA LOJA."""
    if not isinstance(nome_loja_bruto, str):
        return "Não Identificada"
    texto_norm = remover_acentos(nome_loja_bruto).lower().strip()
    for chave, nome_padrao in UNIDADES_PADRAO.items():
        if chave in texto_norm:
            return nome_padrao
    # Fallback: pega o trecho após o último hífen, se existir
    if "-" in nome_loja_bruto:
        return nome_loja_bruto.split("-")[-1].strip()
    return nome_loja_bruto.strip()


@st.cache_data(show_spinner="Carregando e processando base de pedidos...")
def carregar_dados(file_bytes: bytes) -> pd.DataFrame:
    """
    Lê a planilha Excel do relatório de pedidos (iFood/ERP) usando openpyxl
    com reset_dimensions() para evitar problemas de metadados de exportação.
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    ws.reset_dimensions()

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return pd.DataFrame()

    header = [str(col).strip() if col is not None else f"col_{i}" for i, col in enumerate(rows[0])]
    df = pd.DataFrame(rows[1:], columns=header)

    # Conversão de data e hora
    df["DATA_HORA"] = pd.to_datetime(
        df["DATA E HORA DO PEDIDO"], format="%d/%m/%Y %H:%M:%S", errors="coerce"
    )
    df = df.dropna(subset=["DATA_HORA"]).copy()

    df["Data"] = df["DATA_HORA"].dt.date
    df["Hora_Num"] = df["DATA_HORA"].dt.hour
    df["Hora"] = df["Hora_Num"].apply(lambda h: f"{h:02d}:00")
    df["Dia_Semana_Num"] = df["DATA_HORA"].dt.dayofweek
    df["Dia_Semana"] = df["Dia_Semana_Num"].map(DIAS_SEMANA_MAP)

    # Dimensões Loja (unidade física) e Marca (NOME DA LOJA no iFood)
    df["Marca"] = df["NOME DA LOJA"].astype(str).str.strip()
    df["Loja"] = df["Marca"].apply(extrair_loja)

    # Normalização de colunas financeiras
    colunas_valor = [
        "VALOR DOS ITENS (R$)",
        "TOTAL PAGO PELO CLIENTE (R$)",
        "VALOR LIQUIDO (R$)",
    ]
    for col in colunas_valor:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # Flags de status
    status_norm = df["STATUS FINAL DO PEDIDO"].astype(str).str.upper().str.strip()
    df["Is_Concluido"] = status_norm == "CONCLUIDO"
    df["Is_Cancelado"] = status_norm.isin(["CANCELADO", "CANCELAMENTO PARCIAL"])

    return df


def formatar_brl(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ==============================================================================
# INTERFACE PRINCIPAL & CARREGAMENTO DE ARQUIVO
# ==============================================================================
st.title("📊 Relatório de Vendas e Pedidos por Hora")
st.caption("Análise horária de volume de pedidos, cancelamentos e faturamento por Loja, Marca e Dia da Semana")

with st.sidebar:
    st.header("📁 Base de Dados")
    uploaded_file = st.file_uploader(
        "Carregar planilha Excel (.xlsx)",
        type=["xlsx"],
        help="Envie o arquivo relatorio-pedidos_*.xlsx exportado do portal.",
    )

    # Tenta carregar arquivo padrão na pasta local caso nenhum upload seja feito
    arquivo_padrao = "relatorio-pedidos_76c117942bf39c0113f8d1bec9d43fe58de8f71241a21490d0bb2343c11a18de_2026-10-15-2026-11-15.xlsx"
    file_bytes = None

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
    else:
        try:
            with open(arquivo_padrao, "rb") as f:
                file_bytes = f.read()
            st.success("Base padrão carregada automaticamente.")
        except FileNotFoundError:
            pass

if file_bytes is None:
    st.info("👈 Utilize o menu lateral para fazer upload da planilha Excel de pedidos (`.xlsx`).")
    st.stop()

df = carregar_dados(file_bytes)

if df.empty:
    st.error("Não foi possível ler registros válidos na planilha enviada.")
    st.stop()

# ==============================================================================
# FILTROS (MULTI-SELEÇÃO: LOJA, MARCA, DIAS DA SEMANA E PERÍODO)
# ==============================================================================
with st.sidebar:
    st.markdown("---")
    st.header("🎛️ Filtros")
    st.caption("Deixe em branco para considerar **Todos** ou selecione uma ou mais opções.")

    # 1. Filtro de Período
    data_min = df["Data"].min()
    data_max = df["Data"].max()
    periodo = st.date_input(
        "Período",
        value=(data_min, data_max),
        min_value=data_min,
        max_value=data_max,
    )

    # 2. Filtro de Dias da Semana (Multi-seleção)
    dias_selecionados = st.multiselect(
        "Dia da Semana",
        options=ORDEM_DIAS_SEMANA,
        default=[],
        placeholder="Todos os dias da semana",
    )

    # 3. Filtro de Loja (Multi-seleção)
    lojas_disponiveis = sorted(df["Loja"].unique().tolist())
    lojas_selecionadas = st.multiselect(
        "Loja",
        options=lojas_disponiveis,
        default=[],
        placeholder="Todas as lojas",
    )

    # 4. Filtro de Marca (Multi-seleção, dinâmico conforme Lojas escolhidas)
    if lojas_selecionadas:
        marcas_disponiveis = sorted(
            df[df["Loja"].isin(lojas_selecionadas)]["Marca"].unique().tolist()
        )
    else:
        marcas_disponiveis = sorted(df["Marca"].unique().tolist())

    marcas_selecionadas = st.multiselect(
        "Marca",
        options=marcas_disponiveis,
        default=[],
        placeholder="Todas as marcas",
    )

    st.markdown("---")
    faixa_horaria = st.slider(
        "Faixa Horária de Exibição",
        min_value=0,
        max_value=23,
        value=(7, 23),
        format="%02d:00",
    )

# Aplicação dos filtros
df_filtrado = df.copy()

if isinstance(periodo, tuple) and len(periodo) == 2:
    df_filtrado = df_filtrado[
        (df_filtrado["Data"] >= periodo[0]) & (df_filtrado["Data"] <= periodo[1])
    ]

if dias_selecionados:
    df_filtrado = df_filtrado[df_filtrado["Dia_Semana"].isin(dias_selecionados)]

if lojas_selecionadas:
    df_filtrado = df_filtrado[df_filtrado["Loja"].isin(lojas_selecionadas)]

if marcas_selecionadas:
    df_filtrado = df_filtrado[df_filtrado["Marca"].isin(marcas_selecionadas)]

if df_filtrado.empty:
    st.warning("Nenhum pedido encontrado para a combinação de filtros selecionada.")
    st.stop()

# ==============================================================================
# CÁLCULO DOS INDICADORES E MÉDIAS POR HORA
# ==============================================================================
# Quantidade de dias únicos no recorte filtrado (denominador da média diária por hora)
qtd_dias_unicos = max(df_filtrado["Data"].nunique(), 1)

df_concluidos = df_filtrado[df_filtrado["Is_Concluido"]]
df_cancelados = df_filtrado[df_filtrado["Is_Cancelado"]]

total_pedidos_concluidos = len(df_concluidos)
total_pedidos_cancelados = len(df_cancelados)
total_geral_pedidos = len(df_filtrado)

faturamento_itens = df_concluidos["VALOR DOS ITENS (R$)"].sum()
faturamento_liquido = df_concluidos["VALOR LIQUIDO (R$)"].sum()
ticket_medio = faturamento_itens / total_pedidos_concluidos if total_pedidos_concluidos > 0 else 0.0
taxa_cancelamento = (
    (total_pedidos_cancelados / total_geral_pedidos) * 100 if total_geral_pedidos > 0 else 0.0
)

# Cards de KPIs
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric(
    "Pedidos Concluídos",
    f"{total_pedidos_concluidos:,}".replace(",", "."),
    f"Média: {total_pedidos_concluidos / qtd_dias_unicos:.1f} / dia",
)
col2.metric(
    "Vendas Brutas (Itens)",
    formatar_brl(faturamento_itens),
    f"Média: {formatar_brl(faturamento_itens / qtd_dias_unicos)} / dia",
)
col3.metric(
    "Faturamento Líquido",
    formatar_brl(faturamento_liquido),
    f"Média: {formatar_brl(faturamento_liquido / qtd_dias_unicos)} / dia",
)
col4.metric(
    "Ticket Médio",
    formatar_brl(ticket_medio),
    f"{qtd_dias_unicos} dias analisados",
)
col5.metric(
    "Pedidos Cancelados",
    f"{total_pedidos_cancelados:,}".replace(",", "."),
    f"Taxa: {taxa_cancelamento:.2f}%",
    delta_color="inverse",
)

st.markdown("---")

# Agregação horária completa (00h a 23h)
horas_base = pd.DataFrame({"Hora_Num": range(faixa_horaria[0], faixa_horaria[1] + 1)})
horas_base["Hora"] = horas_base["Hora_Num"].apply(lambda h: f"{h:02d}:00")

agg_concluidos = (
    df_concluidos.groupby("Hora_Num")
    .agg(
        Total_Pedidos=("ID COMPLETO DO PEDIDO", "count"),
        Vendas_Brutas=("VALOR DOS ITENS (R$)", "sum"),
        Vendas_Liquidas=("VALOR LIQUIDO (R$)", "sum"),
    )
    .reset_index()
)

agg_cancelados = (
    df_cancelados.groupby("Hora_Num")
    .agg(Total_Cancelados=("ID COMPLETO DO PEDIDO", "count"))
    .reset_index()
)

df_hora = (
    horas_base.merge(agg_concluidos, on="Hora_Num", how="left")
    .merge(agg_cancelados, on="Hora_Num", how="left")
    .fillna(0)
)

df_hora["Media_Pedidos_Hora"] = df_hora["Total_Pedidos"] / qtd_dias_unicos
df_hora["Media_Cancelados_Hora"] = df_hora["Total_Cancelados"] / qtd_dias_unicos
df_hora["Media_Vendas_Brutas_Hora"] = df_hora["Vendas_Brutas"] / qtd_dias_unicos
df_hora["Media_Vendas_Liquidas_Hora"] = df_hora["Vendas_Liquidas"] / qtd_dias_unicos
df_hora["Ticket_Medio_Hora"] = df_hora.apply(
    lambda r: r["Vendas_Brutas"] / r["Total_Pedidos"] if r["Total_Pedidos"] > 0 else 0.0,
    axis=1,
)
df_hora["Taxa_Cancelamento_%"] = df_hora.apply(
    lambda r: (r["Total_Cancelados"] / (r["Total_Pedidos"] + r["Total_Cancelados"])) * 100
    if (r["Total_Pedidos"] + r["Total_Cancelados"]) > 0
    else 0.0,
    axis=1,
)

# ==============================================================================
# GRÁFICO 1: MÉDIA DE PEDIDOS E CANCELADOS POR HORA (ESTILO POWER BI DO PRINT)
# ==============================================================================
st.subheader("Média de pedidos e cancelados por hora")

fig_pedidos = go.Figure()

fig_pedidos.add_trace(
    go.Bar(
        x=df_hora["Hora"],
        y=df_hora["Media_Pedidos_Hora"],
        name="Média Pedidos por Hora",
        marker_color="#3399FF",
        text=df_hora["Media_Pedidos_Hora"].round(0).astype(int),
        textposition="outside",
        textfont=dict(color="#E5E7EB", size=12),
        hovertemplate=(
            "<b>Hora: %{x}</b><br>"
            "Média Pedidos/Dia: <b>%{y:.1f}</b><br>"
            "Total no Período: %{customdata[0]}<extra></extra>"
        ),
        customdata=df_hora[["Total_Pedidos"]].values,
    )
)

fig_pedidos.add_trace(
    go.Bar(
        x=df_hora["Hora"],
        y=df_hora["Media_Cancelados_Hora"],
        name="Média Cancelados por Hora",
        marker_color="#28C2C2",
        text=df_hora["Media_Cancelados_Hora"].round(0).astype(int),
        textposition="outside",
        textfont=dict(color="#E5E7EB", size=12),
        hovertemplate=(
            "<b>Hora: %{x}</b><br>"
            "Média Cancelados/Dia: <b>%{y:.1f}</b><br>"
            "Total Cancelados: %{customdata[0]} (%{customdata[1]:.1f}%)<extra></extra>"
        ),
        customdata=df_hora[["Total_Cancelados", "Taxa_Cancelamento_%"]].values,
    )
)

max_y_ped = max(df_hora["Media_Pedidos_Hora"].max() * 1.18, 10)

fig_pedidos.update_layout(
    barmode="group",
    bargap=0.22,
    bargroupgap=0.08,
    plot_bgcolor="#191919",
    paper_bgcolor="#191919",
    font=dict(color="#E5E7EB"),
    height=430,
    margin=dict(l=20, r=20, t=30, b=40),
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=-0.22,
        xanchor="left",
        x=0,
        font=dict(size=12, color="#D1D5DB"),
    ),
    xaxis=dict(
        title="",
        showgrid=False,
        tickfont=dict(color="#D1D5DB", size=12),
    ),
    yaxis=dict(
        title="",
        showgrid=False,
        showticklabels=False,
        range=[0, max_y_ped],
    ),
)

st.plotly_chart(fig_pedidos, use_container_width=True)

# ==============================================================================
# GRÁFICO 2: FATURAMENTO MÉDIO POR HORA & TICKET MÉDIO
# ==============================================================================
col_g1, col_g2 = st.columns([1.35, 1])

with col_g1:
    st.subheader("Média de Vendas por Hora (R$) e Ticket Médio")
    fig_vendas = go.Figure()

    fig_vendas.add_trace(
        go.Bar(
            x=df_hora["Hora"],
            y=df_hora["Media_Vendas_Brutas_Hora"],
            name="Média Vendas Brutas/Dia (R$)",
            marker_color="#10B981",
            text=df_hora["Media_Vendas_Brutas_Hora"].apply(lambda v: f"R$ {v:,.0f}".replace(",", ".")),
            textposition="outside",
            textfont=dict(color="#D1D5DB", size=10),
        )
    )

    fig_vendas.add_trace(
        go.Scatter(
            x=df_hora["Hora"],
            y=df_hora["Ticket_Medio_Hora"],
            name="Ticket Médio (R$)",
            mode="lines+markers",
            yaxis="y2",
            line=dict(color="#F59E0B", width=2.5),
            marker=dict(size=6, color="#F59E0B"),
        )
    )

    max_y_ven = max(df_hora["Media_Vendas_Brutas_Hora"].max() * 1.2, 100)
    max_y_tkt = max(df_hora["Ticket_Medio_Hora"].max() * 1.25, 20)

    fig_vendas.update_layout(
        plot_bgcolor="#191919",
        paper_bgcolor="#191919",
        font=dict(color="#E5E7EB"),
        height=390,
        margin=dict(l=20, r=20, t=30, b=40),
        legend=dict(orientation="h", y=-0.22, x=0),
        xaxis=dict(showgrid=False),
        yaxis=dict(showgrid=False, showticklabels=False, range=[0, max_y_ven]),
        yaxis2=dict(
            overlaying="y",
            side="right",
            showgrid=False,
            showticklabels=False,
            range=[0, max_y_tkt],
        ),
    )
    st.plotly_chart(fig_vendas, use_container_width=True)

with col_g2:
    st.subheader("Mapa de Calor: Pedidos por Dia × Hora")
    dias_por_dow = df_filtrado.groupby("Dia_Semana")["Data"].nunique().to_dict()
    pivot_heat = (
        df_concluidos[
            (df_concluidos["Hora_Num"] >= faixa_horaria[0])
            & (df_concluidos["Hora_Num"] <= faixa_horaria[1])
        ]
        .pivot_table(
            index="Hora",
            columns="Dia_Semana",
            values="ID COMPLETO DO PEDIDO",
            aggfunc="count",
            fill_value=0,
        )
    )
    # Reordena colunas segundo os dias da semana e normaliza pela média diária de cada dia
    cols_presentes = [d for d in ORDEM_DIAS_SEMANA if d in pivot_heat.columns]
    pivot_heat = pivot_heat[cols_presentes]
    for d in cols_presentes:
        pivot_heat[d] = (pivot_heat[d] / max(dias_por_dow.get(d, 1), 1)).round(1)

    fig_heat = px.imshow(
        pivot_heat,
        labels=dict(x="Dia da Semana", y="Hora", color="Média Pedidos"),
        color_continuous_scale="Blues",
        aspect="auto",
        text_auto=".0f",
    )
    fig_heat.update_layout(
        plot_bgcolor="#191919",
        paper_bgcolor="#191919",
        font=dict(color="#E5E7EB"),
        height=390,
        margin=dict(l=10, r=10, t=30, b=40),
        coloraxis_showscale=False,
    )
    st.plotly_chart(fig_heat, use_container_width=True)

# ==============================================================================
# TABELA ANALÍTICA DETALHADA E EXPORTAÇÃO
# ==============================================================================
st.subheader("📋 Detalhamento Analítico por Hora")

tabela_exibicao = df_hora[
    [
        "Hora",
        "Total_Pedidos",
        "Media_Pedidos_Hora",
        "Total_Cancelados",
        "Media_Cancelados_Hora",
        "Taxa_Cancelamento_%",
        "Vendas_Brutas",
        "Media_Vendas_Brutas_Hora",
        "Vendas_Liquidas",
        "Ticket_Medio_Hora",
    ]
].copy()

tabela_exibicao.columns = [
    "Hora",
    "Total Pedidos",
    "Média Pedidos/Dia",
    "Total Cancelados",
    "Média Cancelados/Dia",
    "Taxa Cancelamento (%)",
    "Vendas Brutas (R$)",
    "Média Vendas/Dia (R$)",
    "Faturamento Líquido (R$)",
    "Ticket Médio (R$)",
]

st.dataframe(
    tabela_exibicao.style.format(
        {
            "Total Pedidos": "{:,.0f}",
            "Média Pedidos/Dia": "{:,.1f}",
            "Total Cancelados": "{:,.0f}",
            "Média Cancelados/Dia": "{:,.1f}",
            "Taxa Cancelamento (%)": "{:.2f}%",
            "Vendas Brutas (R$)": "R$ {:,.2f}",
            "Média Vendas/Dia (R$)": "R$ {:,.2f}",
            "Faturamento Líquido (R$)": "R$ {:,.2f}",
            "Ticket Médio (R$)": "R$ {:,.2f}",
        }
    ),
    use_container_width=True,
    hide_index=True,
)

# Botão de download do relatório filtrado em Excel
buffer_excel = io.BytesIO()
with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
    tabela_exibicao.to_excel(writer, index=False, sheet_name="Vendas_Por_Hora")

st.download_button(
    label="⬇️ Baixar Tabela Consolidada em Excel (.xlsx)",
    data=buffer_excel.getvalue(),
    file_name="relatorio_vendas_por_hora.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
