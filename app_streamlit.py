"""
Dashboard Web para Análise de Extrato Bancário

Execute com:
    streamlit run app_streamlit.py

Funcionalidades:
- upload da planilha;
- leitura automática;
- filtros laterais;
- cards de resumo;
- gráficos;
- tabela filtrada;
- download do relatório Excel;
- análises de auditoria, ranking e comparação entre períodos.
"""

from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path
from typing import Any

import matplotlib.pyplot as plt
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from analisador_extrato import (
    COLUNAS_ESPERADAS,
    converter_data_excel,
    criar_resumos,
    formatar_planilha_excel,
    limpar_texto,
    normalizar_nome_coluna,
)


st.set_page_config(
    page_title="Análise de Extrato Bancário",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


APP_DIR = Path(__file__).resolve().parent
LOGO_CANDIDATOS = [
    APP_DIR / "assets" / "logo_ampla.jpg",
    APP_DIR / "assets" / "LOGO AMPLA_page-0001.jpg",
    APP_DIR / "LOGO AMPLA_page-0001.jpg",
    Path("/mnt/data/LOGO AMPLA_page-0001.jpg"),
]



# -----------------------------------------------------------------------------
# ESTILO VISUAL
# -----------------------------------------------------------------------------

st.markdown(
    """
    <style>
        :root {
            --azul-900: #0B1F3A;
            --azul-800: #123A63;
            --azul-700: #1D4ED8;
            --azul-600: #2563EB;
            --azul-500: #3B82F6;
            --azul-100: #DBEAFE;
            --azul-050: #EFF6FF;
            --cinza-900: #101828;
            --cinza-700: #344054;
            --cinza-600: #475467;
            --cinza-500: #667085;
            --cinza-300: #D0D5DD;
            --cinza-200: #EAECF0;
            --cinza-100: #F2F4F7;
            --cinza-050: #F8FAFC;
            --branco: #FFFFFF;
            --verde: #059669;
            --vermelho: #DC2626;
            --sombra-card: 0 12px 30px rgba(16, 24, 40, 0.08);
        }

        .stApp {
            background:
                radial-gradient(circle at top left, rgba(59, 130, 246, 0.16), transparent 30%),
                linear-gradient(180deg, #F8FAFC 0%, #EFF6FF 42%, #F8FAFC 100%);
            color: var(--cinza-900);
        }

        .main .block-container {
            padding-top: 1.4rem;
            padding-bottom: 2.6rem;
            max-width: 1480px;
        }

        header[data-testid="stHeader"] {
            background: rgba(248, 250, 252, 0.76);
            backdrop-filter: blur(12px);
            border-bottom: 1px solid rgba(208, 213, 221, 0.45);
        }

        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #FFFFFF 0%, #F8FAFC 100%);
            border-right: 1px solid var(--cinza-200);
            box-shadow: 6px 0 28px rgba(16, 24, 40, 0.04);
        }

        section[data-testid="stSidebar"] * {
            color: var(--cinza-900) !important;
        }

        section[data-testid="stSidebar"] h1,
        section[data-testid="stSidebar"] h2,
        section[data-testid="stSidebar"] h3,
        section[data-testid="stSidebar"] label {
            font-weight: 750 !important;
        }

        .app-title {
            font-size: 2.25rem;
            font-weight: 850;
            line-height: 1.1;
            color: var(--azul-900);
            margin-bottom: 0.25rem;
            letter-spacing: -0.04rem;
        }

        .app-subtitle {
            color: var(--cinza-600);
            font-size: 1.02rem;
            margin-bottom: 1.35rem;
        }

        .section-title {
            font-size: 1.28rem;
            font-weight: 800;
            color: var(--azul-900);
            margin-top: 1rem;
            margin-bottom: 0.7rem;
            letter-spacing: -0.015rem;
        }

        .painel {
            background: rgba(255, 255, 255, 0.92);
            border: 1px solid rgba(208, 213, 221, 0.72);
            border-radius: 1.05rem;
            padding: 1rem 1.1rem;
            box-shadow: var(--sombra-card);
        }

        .metric-card {
            background: rgba(255, 255, 255, 0.96);
            border: 1px solid rgba(208, 213, 221, 0.72);
            border-left: 5px solid var(--azul-600);
            border-radius: 1.05rem;
            padding: 1.05rem 1.1rem;
            min-height: 112px;
            box-shadow: var(--sombra-card);
        }

        .metric-label {
            color: var(--cinza-600);
            font-size: 0.82rem;
            font-weight: 760;
            letter-spacing: 0.03rem;
            margin-bottom: 0.5rem;
            text-transform: uppercase;
        }

        .metric-value {
            font-size: 1.72rem;
            font-weight: 850;
            line-height: 1.15;
            word-break: break-word;
            letter-spacing: -0.03rem;
        }

        .metric-help {
            margin-top: 0.45rem;
            color: var(--cinza-500);
            font-size: 0.82rem;
        }

        .valor-entrada { color: var(--verde); }
        .valor-saida { color: var(--vermelho); }
        .valor-info { color: var(--azul-600); }

        h1, h2, h3, h4, h5, h6, p, label, span, div {
            color: inherit;
        }

        div[data-testid="stAlert"] {
            background: #FFFFFF;
            color: var(--cinza-900);
            border: 1px solid var(--cinza-200);
            border-left: 5px solid var(--azul-600);
            border-radius: 0.85rem;
            box-shadow: 0 8px 22px rgba(16, 24, 40, 0.06);
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: .4rem;
            border-bottom: 1px solid var(--cinza-200);
        }

        .stTabs [data-baseweb="tab"] {
            background-color: #FFFFFF;
            border-radius: .8rem .8rem 0 0;
            color: var(--cinza-700) !important;
            border: 1px solid var(--cinza-200);
            border-bottom: none;
            padding: .55rem .85rem;
            font-weight: 700;
        }

        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, var(--azul-600), var(--azul-700)) !important;
            color: #FFFFFF !important;
            border-color: transparent;
        }

        .stTabs [aria-selected="true"] p,
        .stTabs [aria-selected="true"] span {
            color: #FFFFFF !important;
        }

        div[data-testid="stDataFrame"] {
            background: #FFFFFF;
            border: 1px solid var(--cinza-200);
            border-radius: .95rem;
            overflow: hidden;
            box-shadow: 0 10px 24px rgba(16, 24, 40, 0.06);
        }

        div[data-testid="stDataFrame"] * {
            color: var(--cinza-900);
        }

        .stDownloadButton button, .stButton button {
            background: linear-gradient(135deg, var(--azul-600), var(--azul-700));
            color: #FFFFFF !important;
            border: 1px solid rgba(37, 99, 235, 0.2);
            border-radius: .75rem;
            font-weight: 760;
            box-shadow: 0 8px 18px rgba(37, 99, 235, 0.22);
        }

        .stDownloadButton button:hover, .stButton button:hover {
            background: linear-gradient(135deg, #1D4ED8, #1E40AF);
            color: #FFFFFF !important;
            border-color: rgba(29, 78, 216, 0.32);
        }

        div[data-baseweb="select"] > div,
        div[data-baseweb="input"] > div,
        input,
        textarea {
            background-color: #FFFFFF !important;
            color: var(--cinza-900) !important;
            border-color: var(--cinza-200) !important;
            border-radius: .78rem !important;
        }

        div[data-baseweb="select"] span,
        div[data-baseweb="select"] svg,
        div[data-baseweb="input"] input,
        textarea {
            color: var(--cinza-900) !important;
            fill: var(--cinza-700) !important;
        }

        div[data-baseweb="select"] [role="option"] {
            color: var(--cinza-900) !important;
            background: #FFFFFF !important;
        }

        div[data-baseweb="popover"],
        div[data-baseweb="menu"] {
            background: #FFFFFF !important;
            color: var(--cinza-900) !important;
            border-radius: .8rem !important;
            box-shadow: 0 16px 36px rgba(16, 24, 40, 0.14) !important;
        }

        [data-testid="stDateInput"] input,
        [data-testid="stTextInput"] input {
            background: #FFFFFF !important;
            color: var(--cinza-900) !important;
        }

        [data-testid="stSlider"] * {
            color: var(--cinza-700) !important;
        }

        [data-testid="stFileUploader"] section {
            background: rgba(255, 255, 255, 0.92);
            border: 1px dashed var(--azul-500);
            border-radius: 1rem;
        }

        [data-testid="stFileUploader"] section * {
            color: var(--cinza-700) !important;
        }

        small, .stCaptionContainer, .stMarkdown p {
            color: var(--cinza-600);
        }

        hr {
            border-color: var(--cinza-200) !important;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------------------------------------------------------
# FUNÇÕES DE BASE
# -----------------------------------------------------------------------------

def formatar_moeda(valor: float | int | None) -> str:
    """Formata valores em real no padrão brasileiro."""
    valor = 0 if valor is None or pd.isna(valor) else float(valor)
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_percentual(valor: float | int | None) -> str:
    if valor is None or pd.isna(valor):
        return "0,00%"
    return f"{float(valor):,.2f}%".replace(",", "X").replace(".", ",").replace("X", ".")


def validar_colunas(df: pd.DataFrame) -> None:
    """Valida se a planilha possui as colunas mínimas esperadas."""
    colunas_faltantes = COLUNAS_ESPERADAS - set(df.columns)
    if colunas_faltantes:
        raise ValueError(
            "A planilha enviada não está no padrão esperado. "
            f"Colunas faltantes: {', '.join(sorted(colunas_faltantes))}. "
            "Verifique se o arquivo possui as colunas DATA, PAGAMENTO, ENTRADA, SAÍDA, SALDO, "
            "DESCRICAO, ONDE LANÇAR, DOCUMENTO e PLANO FINANCEIRO."
        )


def preparar_extrato(df_bruto: pd.DataFrame) -> pd.DataFrame:
    """Padroniza e enriquece os dados do extrato."""
    if df_bruto is None or df_bruto.empty:
        raise ValueError("A aba selecionada está vazia ou não possui dados para análise.")

    df = df_bruto.copy()
    df.columns = [normalizar_nome_coluna(c) for c in df.columns]
    validar_colunas(df)

    df = df.dropna(how="all").copy()
    if df.empty:
        raise ValueError("A aba selecionada não possui linhas preenchidas.")

    df["DATA"] = converter_data_excel(df["DATA"])

    linhas_sem_data = df["DATA"].isna().sum()
    df = df[df["DATA"].notna()].copy()
    if df.empty:
        raise ValueError("Não foi possível identificar datas válidas na coluna DATA.")

    for coluna in ["ENTRADA", "SAÍDA", "SALDO"]:
        df[coluna] = pd.to_numeric(df[coluna], errors="coerce").fillna(0)

    for coluna in ["PAGAMENTO", "DESCRICAO", "ONDE LANÇAR", "DOCUMENTO", "PLANO FINANCEIRO"]:
        df[coluna] = df[coluna].apply(limpar_texto)

    def classificar_tipo(row):
        entrada = float(row.get("ENTRADA", 0) or 0)
        saida = float(row.get("SAÍDA", 0) or 0)

        if saida != 0 and entrada == 0:
            return "SAÍDA"

        if entrada != 0 and saida == 0:
            return "ENTRADA"

        if entrada != 0 and saida != 0:
            return "MISTO"

        return "ZERADO"

    # GASTO sempre será positivo, mesmo que a coluna SAÍDA venha negativa.
    df["GASTO"] = df["SAÍDA"].abs()

    # VALOR representa o impacto financeiro real da linha:
    # entrada aumenta, saída diminui.
    df["VALOR"] = df["ENTRADA"] - df["GASTO"]

    # TIPO agora depende da coluna preenchida, não mais do sinal do valor.
    df["TIPO"] = df.apply(classificar_tipo, axis=1)

    df["ANO_MES"] = df["DATA"].dt.to_period("M").astype(str)
    df["DIA"] = df["DATA"].dt.date

    if linhas_sem_data:
        st.warning(
            f"{linhas_sem_data} linha(s) foram ignoradas porque não tinham data válida."
        )

    return df.sort_values("DATA")


@st.cache_data(show_spinner=False)
def listar_abas(arquivo_bytes: bytes) -> list[str]:
    """Lista as abas disponíveis no Excel enviado, lendo o arquivo direto da memória."""
    try:
        with pd.ExcelFile(BytesIO(arquivo_bytes), engine="openpyxl") as excel:
            return excel.sheet_names
    except Exception as erro:
        raise ValueError(
            "Não foi possível abrir o arquivo Excel. Confirme se o arquivo é .xlsx válido e não está corrompido."
        ) from erro


@st.cache_data(show_spinner=False)
def carregar_extrato_por_bytes(arquivo_bytes: bytes, aba: str | int | None) -> pd.DataFrame:
    """Carrega e prepara a aba selecionada, sem criar arquivo temporário."""
    try:
        df_bruto = pd.read_excel(BytesIO(arquivo_bytes), sheet_name=aba or 0, engine="openpyxl")
        return preparar_extrato(df_bruto)
    except ValueError:
        raise
    except Exception as erro:
        raise ValueError(
            "Não foi possível processar a planilha. Verifique se a aba selecionada contém uma tabela válida."
        ) from erro


def filtrar_dados(
    df: pd.DataFrame,
    documentos: list[str],
    planos: list[str],
    onde_lancar: list[str],
    tipo: str,
    texto: str,
    data_inicio,
    data_fim,
) -> pd.DataFrame:
    """Aplica filtros selecionados na lateral."""
    filtrado = df.copy()

    if data_inicio:
        filtrado = filtrado[filtrado["DATA"] >= pd.to_datetime(data_inicio)]

    if data_fim:
        filtrado = filtrado[filtrado["DATA"] <= pd.to_datetime(data_fim)]

    if documentos:
        filtrado = filtrado[filtrado["DOCUMENTO"].isin(documentos)]

    if planos:
        filtrado = filtrado[filtrado["PLANO FINANCEIRO"].isin(planos)]

    if onde_lancar:
        filtrado = filtrado[filtrado["ONDE LANÇAR"].isin(onde_lancar)]

    if tipo != "TODOS":
        filtrado = filtrado[filtrado["TIPO"] == tipo]

    if texto.strip():
        termo = texto.strip()
        colunas_busca = ["DESCRICAO", "PAGAMENTO", "DOCUMENTO", "PLANO FINANCEIRO", "ONDE LANÇAR"]
        mascara = pd.Series(False, index=filtrado.index)
        for coluna in colunas_busca:
            if coluna in filtrado.columns:
                mascara = mascara | filtrado[coluna].str.contains(termo, case=False, na=False)
        filtrado = filtrado[mascara]

    return filtrado.sort_values("DATA")


def periodo_anterior(data_inicio, data_fim) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Calcula o período anterior com a mesma quantidade de dias do período filtrado."""
    inicio = pd.to_datetime(data_inicio)
    fim = pd.to_datetime(data_fim)
    dias = max((fim - inicio).days + 1, 1)
    fim_anterior = inicio - pd.Timedelta(days=1)
    inicio_anterior = fim_anterior - pd.Timedelta(days=dias - 1)
    return inicio_anterior, fim_anterior


def calcular_indicadores(df: pd.DataFrame) -> dict[str, float | int]:
    """Calcula indicadores principais."""
    if df.empty:
        return {
            "entradas": 0.0,
            "saidas": 0.0,
            "resultado": 0.0,
            "lancamentos": 0,
            "qtd_saidas": 0,
            "ticket_medio_saida": 0.0,
        }

    total_entradas = float(df["ENTRADA"].sum())
    total_saidas = float(df.loc[df["TIPO"] == "SAÍDA", "GASTO"].sum())
    qtd_saidas = int((df["TIPO"] == "SAÍDA").sum())
    return {
        "entradas": total_entradas,
        "saidas": total_saidas,
        "resultado": float(df["VALOR"].sum()),
        "lancamentos": int(len(df)),
        "qtd_saidas": qtd_saidas,
        "ticket_medio_saida": total_saidas / max(qtd_saidas, 1),
    }


def gerar_excel_download(abas: dict[str, pd.DataFrame]) -> bytes:
    """Gera o relatório Excel em memória para download."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        for nome_aba, df_aba in abas.items():
            df_aba.to_excel(writer, sheet_name=nome_aba[:31], index=False)
        formatar_planilha_excel(writer, abas)
    buffer.seek(0)
    return buffer.getvalue()


# -----------------------------------------------------------------------------
# ANÁLISES AVANÇADAS
# -----------------------------------------------------------------------------

def criar_ranking_gastos(df: pd.DataFrame, grupo: str, top_n: int = 15) -> pd.DataFrame:
    """Cria ranking de gastos por uma dimensão."""
    if df.empty or grupo not in df.columns:
        return pd.DataFrame(columns=[grupo, "TOTAL_GASTO", "QTD_LANCAMENTOS", "PERCENTUAL_DO_TOTAL", "TICKET_MEDIO"])

    saidas = df[df["TIPO"] == "SAÍDA"].copy()
    total = saidas["GASTO"].sum()
    if saidas.empty or total == 0:
        return pd.DataFrame(columns=[grupo, "TOTAL_GASTO", "QTD_LANCAMENTOS", "PERCENTUAL_DO_TOTAL", "TICKET_MEDIO"])

    ranking = (
        saidas.groupby(grupo, as_index=False)
        .agg(
            TOTAL_GASTO=("GASTO", "sum"),
            QTD_LANCAMENTOS=("GASTO", "size"),
            TICKET_MEDIO=("GASTO", "mean"),
        )
        .sort_values("TOTAL_GASTO", ascending=False)
        .head(top_n)
    )
    ranking["PERCENTUAL_DO_TOTAL"] = ranking["TOTAL_GASTO"] / total * 100
    return ranking


def identificar_maiores_despesas(df: pd.DataFrame, top_n: int = 20) -> pd.DataFrame:
    """Lista as maiores despesas individuais."""
    colunas = [
        "DATA",
        "PAGAMENTO",
        "DESCRICAO",
        "DOCUMENTO",
        "PLANO FINANCEIRO",
        "ONDE LANÇAR",
        "GASTO",
        "SALDO",
    ]
    if df.empty:
        return pd.DataFrame(columns=colunas)
    return df[df["TIPO"] == "SAÍDA"].sort_values("GASTO", ascending=False).head(top_n)[colunas]


def comparar_periodos(df_original: pd.DataFrame, df_atual: pd.DataFrame, data_inicio, data_fim) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Compara o período filtrado com o período anterior equivalente."""
    inicio_ant, fim_ant = periodo_anterior(data_inicio, data_fim)
    df_anterior = df_original[(df_original["DATA"] >= inicio_ant) & (df_original["DATA"] <= fim_ant)].copy()

    atual = calcular_indicadores(df_atual)
    anterior = calcular_indicadores(df_anterior)

    linhas = []
    for chave, nome in [
        ("entradas", "Entradas"),
        ("saidas", "Saídas"),
        ("resultado", "Resultado líquido"),
        ("lancamentos", "Lançamentos"),
        ("ticket_medio_saida", "Média por saída"),
    ]:
        valor_atual = atual[chave]
        valor_anterior = anterior[chave]
        variacao_abs = float(valor_atual) - float(valor_anterior)
        variacao_pct = (variacao_abs / float(valor_anterior) * 100) if float(valor_anterior) != 0 else None
        linhas.append(
            {
                "INDICADOR": nome,
                "PERIODO_ATUAL": valor_atual,
                "PERIODO_ANTERIOR": valor_anterior,
                "VARIACAO_ABSOLUTA": variacao_abs,
                "VARIACAO_PERCENTUAL": variacao_pct,
            }
        )

    comparativo = pd.DataFrame(linhas)
    periodo_info = pd.DataFrame(
        [
            {
                "PERIODO": "Atual filtrado",
                "INICIO": pd.to_datetime(data_inicio).date(),
                "FIM": pd.to_datetime(data_fim).date(),
                "QTD_LANCAMENTOS": len(df_atual),
            },
            {
                "PERIODO": "Anterior equivalente",
                "INICIO": inicio_ant.date(),
                "FIM": fim_ant.date(),
                "QTD_LANCAMENTOS": len(df_anterior),
            },
        ]
    )
    return comparativo, periodo_info


def gerar_relatorio_analitico(
    df_filtrado: pd.DataFrame,
    df_original: pd.DataFrame,
    data_inicio,
    data_fim,
    top_n: int,
) -> dict[str, pd.DataFrame]:
    """Gera todas as abas do relatório de prestação de contas e auditoria."""
    abas = criar_resumos(df_filtrado)
    comparativo, periodos = comparar_periodos(df_original, df_filtrado, data_inicio, data_fim)

    abas["Maiores_Despesas"] = identificar_maiores_despesas(df_filtrado, top_n=top_n)
    abas["Ranking_Documento"] = criar_ranking_gastos(df_filtrado, "DOCUMENTO", top_n=top_n)
    abas["Ranking_Plano"] = criar_ranking_gastos(df_filtrado, "PLANO FINANCEIRO", top_n=top_n)
    abas["Ranking_Onde_Lancar"] = criar_ranking_gastos(df_filtrado, "ONDE LANÇAR", top_n=top_n)
    abas["Comparativo_Periodos"] = comparativo
    abas["Info_Periodos"] = periodos
    return abas


# -----------------------------------------------------------------------------
# COMPONENTES VISUAIS
# -----------------------------------------------------------------------------

def render_card(titulo: str, valor: str, classe_cor: str, ajuda: str = "") -> None:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{titulo}</div>
            <div class="metric-value {classe_cor}">{valor}</div>
            <div class="metric-help">{ajuda}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def exibir_cards(df: pd.DataFrame) -> None:
    """Exibe apenas os três cards solicitados."""
    indicadores = calcular_indicadores(df)
    col1, col2, col3 = st.columns(3)
    with col1:
        render_card("Entradas", formatar_moeda(indicadores["entradas"]), "valor-entrada", "Total recebido no período")
    with col2:
        render_card("Saídas", formatar_moeda(indicadores["saidas"]), "valor-saida", "Total gasto no período")
    with col3:
        render_card("Lançamentos", f"{indicadores['lancamentos']}", "valor-info", "Quantidade de movimentações")


def aplicar_layout_plotly(fig: go.Figure, altura: int = 430) -> go.Figure:
    """Padroniza gráficos em tema claro, limpo e profissional."""
    fig.update_layout(
        template="plotly_white",
        height=altura,
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(color="#101828", family="Arial"),
        title_font=dict(color="#0B1F3A", size=18),
        margin=dict(l=10, r=10, t=60, b=10),
        legend=dict(bgcolor="rgba(255,255,255,0)", font=dict(color="#344054")),
    )
    fig.update_xaxes(gridcolor="rgba(208, 213, 221, 0.55)", zerolinecolor="rgba(208, 213, 221, 0.85)")
    fig.update_yaxes(gridcolor="rgba(208, 213, 221, 0.55)", zerolinecolor="rgba(208, 213, 221, 0.85)")
    return fig


def grafico_barras(dados: pd.DataFrame, x: str, y: str, titulo: str, label_x: str, label_y: str, key: str):
    """Cria gráfico de barras horizontal com Plotly."""
    if dados.empty or x not in dados.columns or y not in dados.columns:
        st.info("Sem dados suficientes para este gráfico.")
        return

    dados_plot = dados.sort_values(y, ascending=True)
    fig = px.bar(
        dados_plot,
        x=y,
        y=x,
        orientation="h",
        title=titulo,
        labels={x: label_x, y: label_y},
        text=y,
    )
    fig.update_traces(texttemplate="R$ %{text:,.2f}", textposition="outside", cliponaxis=False)
    aplicar_layout_plotly(fig)
    st.plotly_chart(fig, use_container_width=True, key=key)


def grafico_linha_mensal(mensal: pd.DataFrame, key: str) -> None:
    """Exibe evolução mensal de entradas, saídas e resultado."""
    if mensal.empty:
        st.info("Sem dados suficientes para evolução mensal.")
        return

    colunas_obrigatorias = {"ANO_MES", "TOTAL_ENTRADAS", "TOTAL_SAIDAS_ABS", "RESULTADO_LIQUIDO"}
    if not colunas_obrigatorias.issubset(set(mensal.columns)):
        st.info("Resumo mensal indisponível para este conjunto de dados.")
        return

    dados = mensal[["ANO_MES", "TOTAL_ENTRADAS", "TOTAL_SAIDAS_ABS", "RESULTADO_LIQUIDO"]].copy()
    dados = dados.rename(
        columns={
            "TOTAL_ENTRADAS": "Entradas",
            "TOTAL_SAIDAS_ABS": "Saídas",
            "RESULTADO_LIQUIDO": "Resultado líquido",
        }
    )
    dados_longos = dados.melt(id_vars="ANO_MES", var_name="Indicador", value_name="Valor")

    fig = px.line(
        dados_longos,
        x="ANO_MES",
        y="Valor",
        color="Indicador",
        markers=True,
        title="Evolução mensal",
        labels={"ANO_MES": "Mês", "Valor": "Valor em R$"},
    )
    aplicar_layout_plotly(fig)
    st.plotly_chart(fig, use_container_width=True, key=key)


def grafico_pizza_plano(ranking: pd.DataFrame, key: str) -> None:
    """Mostra participação dos principais planos financeiros nas despesas."""
    if ranking.empty:
        st.info("Sem dados suficientes para participação por plano financeiro.")
        return
    fig = px.pie(
        ranking.head(8),
        names="PLANO FINANCEIRO",
        values="TOTAL_GASTO",
        title="Participação dos maiores planos financeiros",
        hole=0.45,
    )
    aplicar_layout_plotly(fig)
    st.plotly_chart(fig, use_container_width=True, key=key)


def grafico_comparativo(comparativo: pd.DataFrame, key: str) -> None:
    """Gráfico comparando período atual e período anterior."""
    if comparativo.empty:
        st.info("Sem dados suficientes para comparar períodos.")
        return
    dados = comparativo[comparativo["INDICADOR"].isin(["Entradas", "Saídas", "Resultado líquido"])].copy()
    dados = dados.melt(
        id_vars="INDICADOR",
        value_vars=["PERIODO_ATUAL", "PERIODO_ANTERIOR"],
        var_name="Período",
        value_name="Valor",
    )
    dados["Período"] = dados["Período"].replace(
        {"PERIODO_ATUAL": "Período atual", "PERIODO_ANTERIOR": "Período anterior"}
    )
    fig = px.bar(
        dados,
        x="INDICADOR",
        y="Valor",
        color="Período",
        barmode="group",
        title="Comparação entre períodos",
        labels={"INDICADOR": "Indicador", "Valor": "Valor em R$"},
    )
    aplicar_layout_plotly(fig)
    st.plotly_chart(fig, use_container_width=True, key=key)


def dataframe_monetario(df: pd.DataFrame) -> pd.DataFrame:
    """Retorna cópia com datas limpas para visualização amigável no st.dataframe."""
    exibicao = df.copy()
    for coluna in exibicao.columns:
        if pd.api.types.is_period_dtype(exibicao[coluna]):
            exibicao[coluna] = exibicao[coluna].astype(str)
    return exibicao




def caminho_logo() -> Path | None:
    """Retorna o caminho do logotipo para uso apenas nos relatórios exportados."""
    for candidato in LOGO_CANDIDATOS:
        if candidato.exists():
            return candidato
    return None


def maiores_despesas(df: pd.DataFrame, limite: int = 20) -> pd.DataFrame:
    """Retorna as maiores despesas individuais no formato usado pelos relatórios profissionais."""
    colunas = [
        "DATA",
        "PAGAMENTO",
        "DESCRICAO",
        "DOCUMENTO",
        "PLANO FINANCEIRO",
        "ONDE LANÇAR",
        "GASTO",
        "SAÍDA",
        "SALDO",
    ]
    if df.empty:
        return pd.DataFrame(columns=colunas)
    saidas = df[df["TIPO"] == "SAÍDA"].copy()
    return saidas.sort_values("GASTO", ascending=False)[colunas].head(limite)


def preparar_ranking(abas_relatorio: dict[str, pd.DataFrame], nome: str, limite: int = 15) -> pd.DataFrame:
    """Padroniza ranking para gráficos e relatórios profissionais."""
    ranking = abas_relatorio.get(nome, pd.DataFrame()).copy()
    if ranking.empty:
        return ranking
    return ranking.sort_values("TOTAL_GASTO", ascending=False).head(limite)

# -----------------------------------------------------------------------------
# Relatórios profissionais - Excel e PDF
# -----------------------------------------------------------------------------


def grafico_relatorio_png(dados: pd.DataFrame, categoria: str, valor: str, titulo: str, cor: str) -> BytesIO | None:
    """Cria gráfico em PNG para inserir no PDF."""
    if dados.empty or categoria not in dados.columns or valor not in dados.columns:
        return None

    plot_df = dados.sort_values(valor, ascending=False).head(10).sort_values(valor, ascending=True)
    if plot_df.empty:
        return None

    fig, ax = plt.subplots(figsize=(9.5, 5.2))

    barras = ax.barh(
        plot_df[categoria].astype(str),
        plot_df[valor].astype(float),
        color=cor
    )

    maior_valor = float(plot_df[valor].max()) if not plot_df.empty else 0

    for barra in barras:
        largura = barra.get_width()
        posicao_y = barra.get_y() + barra.get_height() / 2

        texto_valor = f"R$ {largura:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

        ax.text(
            largura + (maior_valor * 0.015),
            posicao_y,
            texto_valor,
            va="center",
            ha="left",
            fontsize=8.5,
            color="#344054",
            fontweight="bold",
        )

    ax.set_xlim(0, maior_valor * 1.22 if maior_valor > 0 else 1)
    ax.set_title(titulo, fontsize=15, fontweight="bold", color="#0B2342", pad=14)
    ax.set_xlabel("Valor em R$", color="#344054")
    ax.tick_params(axis="x", colors="#344054")
    ax.tick_params(axis="y", colors="#344054", labelsize=9)
    ax.grid(axis="x", linestyle="--", alpha=0.28)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_color("#D0D5DD")
    ax.spines["bottom"].set_color("#D0D5DD")
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    plt.tight_layout()

    buffer = BytesIO()
    fig.savefig(buffer, format="png", dpi=170, bbox_inches="tight")
    plt.close(fig)
    buffer.seek(0)
    return buffer


def preparar_dados_relatorio(df_filtrado: pd.DataFrame) -> dict[str, Any]:
    """Prepara somente os dados solicitados para Excel/PDF."""
    abas_relatorio = criar_resumos(df_filtrado)
    indicadores = calcular_indicadores(df_filtrado)
    ranking_documento = preparar_ranking(abas_relatorio, "Resumo_Documento", 15)
    ranking_plano = preparar_ranking(abas_relatorio, "Resumo_Plano", 15)
    top_despesas = maiores_despesas(df_filtrado, 20)
    return {
        "indicadores": indicadores,
        "ranking_documento": ranking_documento,
        "ranking_plano": ranking_plano,
        "maiores_despesas": top_despesas,
    }


def gerar_excel_profissional(df_filtrado: pd.DataFrame, filtros_aplicados: dict[str, Any]) -> bytes:
    """Gera um Excel visual e objetivo com os dados solicitados."""
    dados = preparar_dados_relatorio(df_filtrado)
    indicadores = dados["indicadores"]
    ranking_documento = dados["ranking_documento"]
    ranking_plano = dados["ranking_plano"]
    top_despesas = dados["maiores_despesas"]
    logo = caminho_logo()

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        workbook = writer.book
        ws = workbook.add_worksheet("Resumo")
        writer.sheets["Resumo"] = ws

        azul = "#0B2342"
        azul_claro = "#D9EAF7"
        verde = "#15803D"
        vermelho = "#B42318"
        cinza = "#667085"
        branco = "#FFFFFF"

        fmt_titulo = workbook.add_format({"bold": True, "font_size": 20, "font_color": azul})
        fmt_sub = workbook.add_format({"font_size": 10, "font_color": cinza})
        fmt_card_label = workbook.add_format({"bold": True, "font_size": 10, "font_color": branco, "align": "center"})
        fmt_card_entrada = workbook.add_format({"bold": True, "font_size": 15, "font_color": branco, "bg_color": verde, "align": "center", "valign": "vcenter", "border": 1, "border_color": "#FFFFFF", "num_format": 'R$ #,##0.00'})
        fmt_card_saida = workbook.add_format({"bold": True, "font_size": 15, "font_color": branco, "bg_color": vermelho, "align": "center", "valign": "vcenter", "border": 1, "border_color": "#FFFFFF", "num_format": 'R$ #,##0.00'})
        fmt_card_lanc = workbook.add_format({"bold": True, "font_size": 15, "font_color": branco, "bg_color": "#1570EF", "align": "center", "valign": "vcenter", "border": 1, "border_color": "#FFFFFF"})
        fmt_header = workbook.add_format({"bold": True, "font_color": branco, "bg_color": azul, "border": 1, "align": "center"})
        fmt_moeda = workbook.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00', "border": 1})
        fmt_int = workbook.add_format({"num_format": "0", "border": 1})
        fmt_texto = workbook.add_format({"border": 1})
        fmt_data = workbook.add_format({"num_format": "dd/mm/yyyy", "border": 1})
        fmt_secao = workbook.add_format({"bold": True, "font_size": 13, "font_color": azul, "bg_color": azul_claro, "border": 1})

        ws.hide_gridlines(2)
        ws.set_column("A:A", 3)
        ws.set_column("B:B", 18)
        ws.set_column("C:C", 18)
        ws.set_column("D:D", 18)
        ws.set_column("E:E", 18)
        ws.set_column("F:F", 18)
        ws.set_column("G:G", 18)
        ws.set_column("H:H", 18)
        ws.set_column("I:I", 18)
        ws.set_row(0, 26)

        if logo:
            ws.insert_image("B2", str(logo), {"x_scale": 0.16, "y_scale": 0.16, "object_position": 1})
        ws.write("E2", "Relatório Financeiro", fmt_titulo)
        ws.write("E3", "Resumo executivo gerado a partir dos filtros aplicados no dashboard.", fmt_sub)
        ws.write("E4", f"Período: {filtros_aplicados.get('periodo', 'Todos os dados')}", fmt_sub)
        ws.write("E5", f"Filtros: {filtros_aplicados.get('descricao', 'Sem filtros adicionais')}", fmt_sub)

        ws.merge_range("B8:C8", "ENTRADAS", fmt_card_label)
        ws.merge_range("D8:E8", "SAÍDAS", fmt_card_label)
        ws.merge_range("F8:G8", "LANÇAMENTOS", fmt_card_label)
        ws.merge_range("B9:C10", indicadores["entradas"], fmt_card_entrada)
        ws.merge_range("D9:E10", indicadores["saidas"], fmt_card_saida)
        ws.merge_range("F9:G10", indicadores["lancamentos"], fmt_card_lanc)
        ws.conditional_format("B9:G10", {"type": "no_errors", "format": workbook.add_format({"align": "center", "valign": "vcenter"})})

        # Abas de dados solicitadas
        ranking_documento.to_excel(writer, sheet_name="Ranking Documento", index=False, startrow=1)
        ranking_plano.to_excel(writer, sheet_name="Ranking Plano", index=False, startrow=1)
        top_despesas.to_excel(writer, sheet_name="Maiores Despesas", index=False, startrow=1)

        for sheet_name, df_sheet in [
            ("Ranking Documento", ranking_documento),
            ("Ranking Plano", ranking_plano),
            ("Maiores Despesas", top_despesas),
        ]:
            sheet = writer.sheets[sheet_name]
            sheet.hide_gridlines(2)
            sheet.freeze_panes(2, 0)
            sheet.write(0, 0, sheet_name, fmt_secao)
            if not df_sheet.empty:
                sheet.autofilter(1, 0, len(df_sheet) + 1, len(df_sheet.columns) - 1)
            for col_idx, col in enumerate(df_sheet.columns):
                sheet.write(1, col_idx, col, fmt_header)
                largura = min(max(len(str(col)) + 4, 14), 42)
                sheet.set_column(col_idx, col_idx, largura)
                if col in {"GASTO", "SAÍDA", "SALDO", "TOTAL_GASTO", "ENTRADA"}:
                    sheet.set_column(col_idx, col_idx, 18, fmt_moeda)
                elif col in {"QTD_LANCAMENTOS"}:
                    sheet.set_column(col_idx, col_idx, 16, fmt_int)
                elif col == "DATA":
                    sheet.set_column(col_idx, col_idx, 14, fmt_data)
                else:
                    sheet.set_column(col_idx, col_idx, largura, fmt_texto)

        # Gráfico 1 - Documento
        if not ranking_documento.empty:
            linha_ini = 14
            ws.write(linha_ini, 1, "Ranking de gasto por documento", fmt_secao)
            chart_doc = workbook.add_chart({"type": "bar"})
            max_row = len(ranking_documento) + 1
            chart_doc.add_series({
                "name": "Total gasto",
                "categories": "='Ranking Documento'!$A$3:$A$" + str(max_row + 1),
                "values": "='Ranking Documento'!$B$3:$B$" + str(max_row + 1),
                "fill": {"color": "#1570EF"},
            })
            chart_doc.set_title({"name": "Ranking de gasto por documento"})
            chart_doc.set_x_axis({"name": "Valor em R$", "num_format": 'R$ #,##0'})
            chart_doc.set_y_axis({"reverse": True})
            chart_doc.set_legend({"none": True})
            chart_doc.set_size({"width": 560, "height": 330})
            ws.insert_chart("B16", chart_doc)

        # Gráfico 2 - Plano
        if not ranking_plano.empty:
            ws.write(14, 6, "Ranking de gasto por plano financeiro", fmt_secao)
            chart_plano = workbook.add_chart({"type": "bar"})
            max_row_plano = len(ranking_plano) + 1
            chart_plano.add_series({
                "name": "Total gasto",
                "categories": "='Ranking Plano'!$A$3:$A$" + str(max_row_plano + 1),
                "values": "='Ranking Plano'!$B$3:$B$" + str(max_row_plano + 1),
                "fill": {"color": "#0E9384"},
            })
            chart_plano.set_title({"name": "Ranking de gasto por plano financeiro"})
            chart_plano.set_x_axis({"name": "Valor em R$", "num_format": 'R$ #,##0'})
            chart_plano.set_y_axis({"reverse": True})
            chart_plano.set_legend({"none": True})
            chart_plano.set_size({"width": 560, "height": 330})
            ws.insert_chart("G16", chart_plano)

        # Maiores despesas na aba resumo
        inicio_tabela = 37
        ws.merge_range(inicio_tabela, 1, inicio_tabela, 8, "Maiores despesas individuais", fmt_secao)
        colunas_top = ["DATA", "PAGAMENTO", "DESCRICAO", "DOCUMENTO", "PLANO FINANCEIRO", "ONDE LANÇAR", "GASTO"]
        top_resumo = top_despesas[colunas_top].head(10) if not top_despesas.empty else pd.DataFrame(columns=colunas_top)
        for j, col in enumerate(colunas_top, start=1):
            ws.write(inicio_tabela + 1, j, col, fmt_header)
        for i, (_, row) in enumerate(top_resumo.iterrows(), start=inicio_tabela + 2):
            for j, col in enumerate(colunas_top, start=1):
                val = row[col]
                if col == "DATA" and pd.notna(val):
                    ws.write_datetime(i, j, pd.to_datetime(val).to_pydatetime(), fmt_data)
                elif col == "GASTO":
                    ws.write_number(i, j, float(val), fmt_moeda)
                else:
                    ws.write(i, j, str(val), fmt_texto)

        ws.set_column("C:C", 24)
        ws.set_column("D:D", 34)
        ws.set_column("G:G", 24)
        ws.set_column("H:H", 24)

    buffer.seek(0)
    return buffer.getvalue()


def rodape_pdf(canvas, doc):
    """Adiciona rodapé simples no PDF."""
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#667085"))
    canvas.drawRightString(A4[0] - 1.5 * cm, 1.0 * cm, f"Página {doc.page}")
    canvas.drawString(1.5 * cm, 1.0 * cm, "Relatório financeiro gerado pelo dashboard de análise de extrato")
    canvas.restoreState()


def gerar_pdf_profissional(df_filtrado: pd.DataFrame, filtros_aplicados: dict[str, Any]) -> bytes:
    """Gera PDF profissional com logo, KPIs, gráficos e maiores despesas."""
    dados = preparar_dados_relatorio(df_filtrado)
    indicadores = dados["indicadores"]
    ranking_documento = dados["ranking_documento"]
    ranking_plano = dados["ranking_plano"]
    top_despesas = dados["maiores_despesas"]
    logo = caminho_logo()

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.35 * cm,
        leftMargin=1.35 * cm,
        topMargin=1.25 * cm,
        bottomMargin=1.5 * cm,
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TituloAzul", parent=styles["Title"], textColor=colors.HexColor("#0B2342"), fontSize=21, leading=26, alignment=TA_LEFT, spaceAfter=8))
    styles.add(ParagraphStyle(name="Subtitulo", parent=styles["Normal"], textColor=colors.HexColor("#475467"), fontSize=9.5, leading=13))
    styles.add(ParagraphStyle(name="Secao", parent=styles["Heading2"], textColor=colors.HexColor("#0B2342"), fontSize=14, leading=18, spaceBefore=12, spaceAfter=8))
    styles.add(ParagraphStyle(name="Centro", parent=styles["Normal"], alignment=TA_CENTER, textColor=colors.white, fontSize=9, leading=11))
    styles.add(ParagraphStyle(name="CardValor", parent=styles["Normal"], alignment=TA_CENTER, textColor=colors.white, fontSize=14, leading=18))

    elementos = []

    cabecalho = []
    if logo:
        cabecalho.append(Image(str(logo), width=4.1 * cm, height=2.5 * cm))
    else:
        cabecalho.append(Paragraph("AMPLA", styles["TituloAzul"]))
    cabecalho.append(
        [
            Paragraph("Relatório Financeiro", styles["TituloAzul"]),
            Paragraph("Prestação de contas com base nos filtros aplicados no dashboard.", styles["Subtitulo"]),
            Paragraph(f"<b>Período:</b> {filtros_aplicados.get('periodo', 'Todos os dados')}", styles["Subtitulo"]),
            Paragraph(f"<b>Filtros:</b> {filtros_aplicados.get('descricao', 'Sem filtros adicionais')}", styles["Subtitulo"]),
        ]
    )
    tabela_cabecalho = Table([cabecalho], colWidths=[5.0 * cm, 13.0 * cm])
    tabela_cabecalho.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    elementos.append(tabela_cabecalho)
    elementos.append(Spacer(1, 0.35 * cm))

    cards = [
        [Paragraph("ENTRADAS", styles["Centro"]), Paragraph("SAÍDAS", styles["Centro"]), Paragraph("LANÇAMENTOS", styles["Centro"])],
        [Paragraph(formatar_moeda(indicadores["entradas"]), styles["CardValor"]), Paragraph(formatar_moeda(indicadores["saidas"]), styles["CardValor"]), Paragraph(str(indicadores["lancamentos"]), styles["CardValor"])],
    ]
    tabela_cards = Table(cards, colWidths=[6.1 * cm, 6.1 * cm, 6.1 * cm], rowHeights=[0.72 * cm, 1.05 * cm])
    tabela_cards.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, 1), colors.HexColor("#15803D")),
        ("BACKGROUND", (1, 0), (1, 1), colors.HexColor("#B42318")),
        ("BACKGROUND", (2, 0), (2, 1), colors.HexColor("#1570EF")),
        ("BOX", (0, 0), (-1, -1), 0.7, colors.white),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.white),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    elementos.append(tabela_cards)
    elementos.append(Spacer(1, 0.4 * cm))

    elementos.append(Paragraph("Ranking de gastos", styles["Secao"]))
    grafico_doc = grafico_relatorio_png(ranking_documento, "DOCUMENTO", "TOTAL_GASTO", "Ranking de gasto por documento", "#1570EF")
    grafico_plano = grafico_relatorio_png(ranking_plano, "PLANO FINANCEIRO", "TOTAL_GASTO", "Ranking de gasto por plano financeiro", "#0E9384")

    imagens = []
    if grafico_doc:
        imagens.append(Image(grafico_doc, width=18.0 * cm, height=9.3 * cm))
    if grafico_plano:
        imagens.append(Image(grafico_plano, width=18.0 * cm, height=9.3 * cm))
    if imagens:
        for imagem in imagens:
            elementos.append(imagem)
            elementos.append(Spacer(1, 0.2 * cm))
    else:
        elementos.append(Paragraph("Não há despesas suficientes para gerar gráficos de ranking.", styles["Subtitulo"]))

    elementos.append(PageBreak())
    elementos.append(Paragraph("Maiores despesas individuais", styles["Secao"]))

    colunas = ["DATA", "PAGAMENTO", "DOCUMENTO", "PLANO FINANCEIRO", "GASTO"]
    top_pdf = top_despesas[colunas].head(15) if not top_despesas.empty else pd.DataFrame(columns=colunas)
    dados_tabela = [["Data", "Pagamento", "Documento", "Plano financeiro", "Valor"]]
    for _, row in top_pdf.iterrows():
        data_fmt = pd.to_datetime(row["DATA"]).strftime("%d/%m/%Y") if pd.notna(row["DATA"]) else "-"
        dados_tabela.append([
            data_fmt,
            (str(row["PAGAMENTO"])[:24] + "...") if len(str(row["PAGAMENTO"])) > 27 else str(row["PAGAMENTO"]),
            (str(row["DOCUMENTO"])[:16] + "...") if len(str(row["DOCUMENTO"])) > 19 else str(row["DOCUMENTO"]),
            (str(row["PLANO FINANCEIRO"])[:24] + "...") if len(str(row["PLANO FINANCEIRO"])) > 27 else str(row["PLANO FINANCEIRO"]),
            formatar_moeda(row["GASTO"]),
        ])

    if len(dados_tabela) == 1:
        dados_tabela.append(["-", "Sem despesas no filtro", "-", "-", "-"])

    tabela = Table(dados_tabela, colWidths=[2.3 * cm, 5.2 * cm, 3.2 * cm, 4.7 * cm, 3.0 * cm], repeatRows=1)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B2342")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 8.7),
        ("FONTSIZE", (0, 1), (-1, -1), 7.8),
        ("ALIGN", (4, 1), (4, -1), "RIGHT"),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D0D5DD")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    elementos.append(tabela)

    doc.build(elementos, onFirstPage=rodape_pdf, onLaterPages=rodape_pdf)
    buffer.seek(0)
    return buffer.getvalue()


def descrever_filtros(
    data_inicio: date | None,
    data_fim: date | None,
    documentos: list[str],
    planos: list[str],
    onde_lancar: list[str],
    tipo: str,
    texto: str,
) -> dict[str, str]:
    """Gera descrição textual dos filtros para os relatórios."""
    if data_inicio and data_fim:
        periodo = f"{data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}"
    else:
        periodo = "Todos os dados"

    partes = []
    if documentos:
        partes.append("Documento: " + ", ".join(documentos))
    if planos:
        partes.append("Plano financeiro: " + ", ".join(planos))
    if onde_lancar:
        partes.append("Onde lançar: " + ", ".join(onde_lancar))
    if tipo != "TODOS":
        partes.append("Tipo: " + tipo)
    if texto.strip():
        partes.append("Busca: " + texto.strip())

    return {"periodo": periodo, "descricao": "; ".join(partes) if partes else "Sem filtros adicionais"}



# -----------------------------------------------------------------------------
# INTERFACE PRINCIPAL
# -----------------------------------------------------------------------------

st.markdown('<div class="app-title">Análise de Extrato Bancário</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="app-subtitle">Dashboard para filtros, prestação de contas, auditoria e apoio à tomada de decisão.</div>',
    unsafe_allow_html=True,
)

arquivo = st.file_uploader("Envie sua planilha de extrato bancário (.xlsx)", type=["xlsx"])

if arquivo is None:
    st.info("Envie uma planilha Excel para iniciar a análise.")
    st.stop()

arquivo_bytes = arquivo.getvalue()

try:
    abas_disponiveis = listar_abas(arquivo_bytes)
except Exception as erro:
    st.error(str(erro))
    st.stop()

with st.sidebar:
    st.header("Filtros")
    aba_selecionada = st.selectbox("Aba da planilha", abas_disponiveis, index=0)

try:
    with st.spinner("Lendo e preparando os dados..."):
        df_original = carregar_extrato_por_bytes(arquivo_bytes, aba_selecionada)
except Exception as erro:
    st.error(str(erro))
    with st.expander("Ver detalhes técnicos"):
        st.exception(erro)
    st.stop()

if df_original.empty:
    st.warning("A planilha foi lida, mas não possui lançamentos para analisar.")
    st.stop()

min_data = df_original["DATA"].min()
max_data = df_original["DATA"].max()

with st.sidebar:
    st.divider()

    periodo = st.date_input(
        "Período",
        value=(min_data.date(), max_data.date()),
        min_value=min_data.date(),
        max_value=max_data.date(),
    )

    if isinstance(periodo, tuple) and len(periodo) == 2:
        data_inicio, data_fim = periodo
    else:
        data_inicio, data_fim = min_data.date(), max_data.date()
        st.warning("Selecione data inicial e data final para aplicar o período corretamente.")

    documentos = st.multiselect(
        "Documento",
        sorted(df_original["DOCUMENTO"].dropna().unique().tolist()),
    )

    planos = st.multiselect(
        "Plano financeiro",
        sorted(df_original["PLANO FINANCEIRO"].dropna().unique().tolist()),
    )

    onde_lancar = st.multiselect(
        "Onde lançar",
        sorted(df_original["ONDE LANÇAR"].dropna().unique().tolist()),
    )

    tipo = st.selectbox("Tipo de movimentação", ["TODOS", "ENTRADA", "SAÍDA", "ZERADO"])
    texto = st.text_input("Buscar texto", placeholder="Ex.: energia, asfalto, pix...")
    top_n = st.slider("Quantidade de itens nos rankings", min_value=5, max_value=30, value=15, step=5)

    st.caption("O comparativo usa automaticamente o período anterior com a mesma quantidade de dias.")

if pd.to_datetime(data_inicio) > pd.to_datetime(data_fim):
    st.error("A data inicial não pode ser maior que a data final.")
    st.stop()

df_filtrado = filtrar_dados(
    df_original,
    documentos=documentos,
    planos=planos,
    onde_lancar=onde_lancar,
    tipo=tipo,
    texto=texto,
    data_inicio=data_inicio,
    data_fim=data_fim,
)

if df_filtrado.empty:
    st.warning("Nenhum lançamento encontrado com os filtros selecionados.")
    st.stop()

try:
    abas_relatorio = gerar_relatorio_analitico(df_filtrado, df_original, data_inicio, data_fim, top_n=top_n)
    filtros_desc = descrever_filtros(data_inicio, data_fim, documentos, planos, onde_lancar, tipo, texto)
    excel_bytes = gerar_excel_profissional(df_filtrado, filtros_desc)
    pdf_bytes = gerar_pdf_profissional(df_filtrado, filtros_desc)
except Exception as erro:
    st.error("Não foi possível gerar os relatórios com os filtros atuais.")
    with st.expander("Ver detalhes técnicos"):
        st.exception(erro)
    st.stop()

st.markdown('<div class="section-title">Resumo do período filtrado</div>', unsafe_allow_html=True)
exibir_cards(df_filtrado)

st.caption(f"Exibindo {len(df_filtrado)} de {len(df_original)} lançamentos encontrados na planilha.")

col_down1, col_down2, _ = st.columns([1.25, 1.25, 4])
with col_down1:
    st.download_button(
        label="Baixar relatório Excel",
        data=excel_bytes,
        file_name="relatorio_financeiro_ampla.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with col_down2:
    st.download_button(
        label="Baixar relatório PDF",
        data=pdf_bytes,
        file_name="relatorio_financeiro_ampla.pdf",
        mime="application/pdf",
    )


tab_dashboard, tab_auditoria, tab_comparativo, tab_tabela, tab_relatorios = st.tabs(
    ["Dashboard", "Auditoria e rankings", "Comparação", "Tabela filtrada", "Relatórios"]
)

with tab_dashboard:
    st.markdown('<div class="section-title">Visão geral dos gastos</div>', unsafe_allow_html=True)
    col_esq, col_dir = st.columns(2)
    with col_esq:
        grafico_barras(
            abas_relatorio["Ranking_Documento"],
            x="DOCUMENTO",
            y="TOTAL_GASTO",
            titulo="Ranking de gastos por documento",
            label_x="Documento",
            label_y="Total gasto",
            key="dashboard_ranking_documento",
        )
    with col_dir:
        grafico_barras(
            abas_relatorio["Ranking_Plano"],
            x="PLANO FINANCEIRO",
            y="TOTAL_GASTO",
            titulo="Ranking de gastos por plano financeiro",
            label_x="Plano financeiro",
            label_y="Total gasto",
            key="dashboard_ranking_plano",
        )

    grafico_linha_mensal(abas_relatorio["Resumo_Mensal"], key="dashboard_evolucao_mensal")

with tab_auditoria:
    st.markdown('<div class="section-title">Maiores despesas individuais</div>', unsafe_allow_html=True)
    st.dataframe(
        dataframe_monetario(abas_relatorio["Maiores_Despesas"]),
        use_container_width=True,
        hide_index=True,
        column_config={
            "DATA": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
            "GASTO": st.column_config.NumberColumn("Gasto", format="R$ %.2f"),
            "SALDO": st.column_config.NumberColumn("Saldo", format="R$ %.2f"),
        },
    )

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Ranking por documento")
        st.dataframe(
            dataframe_monetario(abas_relatorio["Ranking_Documento"]),
            use_container_width=True,
            hide_index=True,
            column_config={
                "TOTAL_GASTO": st.column_config.NumberColumn("Total gasto", format="R$ %.2f"),
                "PERCENTUAL_DO_TOTAL": st.column_config.NumberColumn("% do total", format="%.2f%%"),
                "TICKET_MEDIO": st.column_config.NumberColumn("Ticket médio", format="R$ %.2f"),
            },
        )
    with col2:
        st.subheader("Ranking por plano financeiro")
        st.dataframe(
            dataframe_monetario(abas_relatorio["Ranking_Plano"]),
            use_container_width=True,
            hide_index=True,
            column_config={
                "TOTAL_GASTO": st.column_config.NumberColumn("Total gasto", format="R$ %.2f"),
                "PERCENTUAL_DO_TOTAL": st.column_config.NumberColumn("% do total", format="%.2f%%"),
                "TICKET_MEDIO": st.column_config.NumberColumn("Ticket médio", format="R$ %.2f"),
            },
        )

    col3, col4 = st.columns(2)
    with col3:
        grafico_barras(
            abas_relatorio["Ranking_Onde_Lancar"],
            x="ONDE LANÇAR",
            y="TOTAL_GASTO",
            titulo="Ranking de gastos por onde lançar",
            label_x="Onde lançar",
            label_y="Total gasto",
            key="auditoria_ranking_onde_lancar",
        )
    with col4:
        grafico_pizza_plano(abas_relatorio["Ranking_Plano"], key="auditoria_pizza_plano")

with tab_comparativo:
    st.markdown('<div class="section-title">Comparação entre períodos</div>', unsafe_allow_html=True)
    st.dataframe(abas_relatorio["Info_Periodos"], use_container_width=True, hide_index=True)
    grafico_comparativo(abas_relatorio["Comparativo_Periodos"], key="comparativo_periodos_grafico")

    comparativo_view = abas_relatorio["Comparativo_Periodos"].copy()
    st.dataframe(
        comparativo_view,
        use_container_width=True,
        hide_index=True,
        column_config={
            "PERIODO_ATUAL": st.column_config.NumberColumn("Período atual", format="%.2f"),
            "PERIODO_ANTERIOR": st.column_config.NumberColumn("Período anterior", format="%.2f"),
            "VARIACAO_ABSOLUTA": st.column_config.NumberColumn("Variação absoluta", format="%.2f"),
            "VARIACAO_PERCENTUAL": st.column_config.NumberColumn("Variação %", format="%.2f%%"),
        },
    )

with tab_tabela:
    st.markdown('<div class="section-title">Tabela final do filtro aplicado</div>', unsafe_allow_html=True)
    colunas_exibicao = [
        "DATA",
        "PAGAMENTO",
        "ENTRADA",
        "SAÍDA",
        "SALDO",
        "DESCRICAO",
        "ONDE LANÇAR",
        "DOCUMENTO",
        "PLANO FINANCEIRO",
        "TIPO",
    ]
    colunas_existentes = [c for c in colunas_exibicao if c in df_filtrado.columns]
    st.dataframe(
        dataframe_monetario(df_filtrado[colunas_existentes]),
        use_container_width=True,
        hide_index=True,
        column_config={
            "DATA": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
            "ENTRADA": st.column_config.NumberColumn("Entrada", format="R$ %.2f"),
            "SAÍDA": st.column_config.NumberColumn("Saída", format="R$ %.2f"),
            "SALDO": st.column_config.NumberColumn("Saldo", format="R$ %.2f"),
        },
    )

with tab_relatorios:
    st.markdown('<div class="section-title">Tabelas analíticas geradas</div>', unsafe_allow_html=True)

    st.subheader("Documento x Plano Financeiro")
    st.dataframe(
        dataframe_monetario(abas_relatorio["Documento_Plano"]),
        use_container_width=True,
        hide_index=True,
        column_config={"TOTAL_GASTO": st.column_config.NumberColumn("Total gasto", format="R$ %.2f")},
    )

    st.subheader("Resumo mensal")
    st.dataframe(
        dataframe_monetario(abas_relatorio["Resumo_Mensal"]),
        use_container_width=True,
        hide_index=True,
        column_config={
            "TOTAL_ENTRADAS": st.column_config.NumberColumn("Entradas", format="R$ %.2f"),
            "TOTAL_SAIDAS": st.column_config.NumberColumn("Saídas", format="R$ %.2f"),
            "TOTAL_SAIDAS_ABS": st.column_config.NumberColumn("Saídas abs.", format="R$ %.2f"),
            "RESULTADO_LIQUIDO": st.column_config.NumberColumn("Resultado", format="R$ %.2f"),
            "SALDO_FINAL": st.column_config.NumberColumn("Saldo final", format="R$ %.2f"),
        },
    )

    st.subheader("Ranking por onde lançar")
    st.dataframe(
        dataframe_monetario(abas_relatorio["Ranking_Onde_Lancar"]),
        use_container_width=True,
        hide_index=True,
        column_config={
            "TOTAL_GASTO": st.column_config.NumberColumn("Total gasto", format="R$ %.2f"),
            "PERCENTUAL_DO_TOTAL": st.column_config.NumberColumn("% do total", format="%.2f%%"),
            "TICKET_MEDIO": st.column_config.NumberColumn("Ticket médio", format="R$ %.2f"),
        },
    )
