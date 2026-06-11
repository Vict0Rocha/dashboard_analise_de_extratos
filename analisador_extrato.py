"""
Analisador de Extrato Bancário

Objetivo:
- Ler uma planilha de extrato bancário em Excel (.xlsx)
- Aplicar filtros simples
- Gerar relatórios em Excel
- Gerar gráficos em PNG para análise financeira

Exemplo de uso:
    python analisador_extrato.py --arquivo extrato.xlsx
    python analisador_extrato.py --arquivo extrato.xlsx --documento OBRA
    python analisador_extrato.py --arquivo extrato.xlsx --plano "COMBUSTIVEL" --inicio 2026-05-01 --fim 2026-05-31

Dependências:
    pip install pandas openpyxl matplotlib
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Optional

import matplotlib.pyplot as plt
import pandas as pd


COLUNAS_ESPERADAS = {
    "DATA",
    "PAGAMENTO",
    "ENTRADA",
    "SAÍDA",
    "SALDO",
    "DESCRICAO",
    "ONDE LANÇAR",
    "DOCUMENTO",
    "PLANO FINANCEIRO",
}


def normalizar_nome_coluna(nome: str) -> str:
    """Padroniza nomes de colunas removendo espaços duplicados."""
    return re.sub(r"\s+", " ", str(nome).strip()).upper()


def limpar_texto(valor) -> str:
    """Limpa campos de texto usados em filtros e agrupamentos."""
    if pd.isna(valor):
        return "NÃO INFORMADO"
    texto = str(valor).strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto.upper() if texto else "NÃO INFORMADO"


def converter_data_excel(serie: pd.Series) -> pd.Series:
    """
    Converte a coluna DATA.

    A planilha pode vir com datas reais do Excel ou números seriais, como 46161.
    O origin='1899-12-30' é o padrão usado para datas seriais do Excel no pandas.
    """
    if pd.api.types.is_datetime64_any_dtype(serie):
        return pd.to_datetime(serie, errors="coerce")

    numerica = pd.to_numeric(serie, errors="coerce")
    datas_serial = pd.to_datetime(numerica, unit="D", origin="1899-12-30", errors="coerce")
    datas_texto = pd.to_datetime(serie, dayfirst=True, errors="coerce")

    return datas_texto.fillna(datas_serial)


def carregar_extrato(caminho_arquivo: Path, aba: Optional[str] = None) -> pd.DataFrame:
    """Carrega e prepara a planilha de extrato."""
    if not caminho_arquivo.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")

    df = pd.read_excel(caminho_arquivo, sheet_name=aba or 0)
    df.columns = [normalizar_nome_coluna(c) for c in df.columns]

    colunas_faltantes = COLUNAS_ESPERADAS - set(df.columns)
    if colunas_faltantes:
        raise ValueError(
            "A planilha não possui todas as colunas esperadas. "
            f"Colunas faltantes: {', '.join(sorted(colunas_faltantes))}"
        )

    df = df.copy()
    df["DATA"] = converter_data_excel(df["DATA"])

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

    # GASTO sempre será positivo.
    # Assim, funciona tanto quando a coluna SAÍDA vier negativa quanto positiva.
    df["GASTO"] = df["SAÍDA"].abs()

    # VALOR representa o impacto financeiro real:
    # entrada aumenta o saldo, saída reduz o saldo.
    df["VALOR"] = df["ENTRADA"] - df["GASTO"]

    # TIPO passa a ser definido pelas colunas preenchidas,
    # não pelo sinal matemático do valor.
    df["TIPO"] = df.apply(classificar_tipo, axis=1)

    df["ANO_MES"] = df["DATA"].dt.to_period("M").astype(str)

    return df


def aplicar_filtros(
    df: pd.DataFrame,
    documento: Optional[str] = None,
    plano: Optional[str] = None,
    onde_lancar: Optional[str] = None,
    texto: Optional[str] = None,
    inicio: Optional[str] = None,
    fim: Optional[str] = None,
) -> pd.DataFrame:
    """Aplica filtros opcionais aos dados."""
    filtrado = df.copy()

    if documento:
        filtrado = filtrado[filtrado["DOCUMENTO"] == limpar_texto(documento)]

    if plano:
        filtrado = filtrado[filtrado["PLANO FINANCEIRO"] == limpar_texto(plano)]

    if onde_lancar:
        filtrado = filtrado[filtrado["ONDE LANÇAR"] == limpar_texto(onde_lancar)]

    if texto:
        termo = limpar_texto(texto)
        filtrado = filtrado[
            filtrado["DESCRICAO"].str.contains(termo, case=False, na=False)
            | filtrado["PAGAMENTO"].str.contains(termo, case=False, na=False)
        ]

    if inicio:
        data_inicio = pd.to_datetime(inicio, errors="raise")
        filtrado = filtrado[filtrado["DATA"] >= data_inicio]

    if fim:
        data_fim = pd.to_datetime(fim, errors="raise")
        filtrado = filtrado[filtrado["DATA"] <= data_fim]

    return filtrado.sort_values("DATA")


def criar_resumos(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Cria tabelas resumidas para análise."""
    saidas = df[df["TIPO"] == "SAÍDA"].copy()
    entradas = df[df["TIPO"] == "ENTRADA"].copy()

    resumo_documento = (
        saidas.groupby("DOCUMENTO", as_index=False)
        .agg(TOTAL_GASTO=("GASTO", "sum"), QTD_LANCAMENTOS=("GASTO", "size"))
        .sort_values("TOTAL_GASTO", ascending=False)
    )

    resumo_plano = (
        saidas.groupby("PLANO FINANCEIRO", as_index=False)
        .agg(TOTAL_GASTO=("GASTO", "sum"), QTD_LANCAMENTOS=("GASTO", "size"))
        .sort_values("TOTAL_GASTO", ascending=False)
    )

    documento_plano = (
        saidas.groupby(["DOCUMENTO", "PLANO FINANCEIRO"], as_index=False)
        .agg(TOTAL_GASTO=("GASTO", "sum"), QTD_LANCAMENTOS=("GASTO", "size"))
        .sort_values(["DOCUMENTO", "TOTAL_GASTO"], ascending=[True, False])
    )

    mensal = (
        df.groupby("ANO_MES", as_index=False)
        .agg(
            TOTAL_ENTRADAS=("ENTRADA", "sum"),
            TOTAL_SAIDAS=("SAÍDA", "sum"),
            SALDO_FINAL=("SALDO", "last"),
            QTD_LANCAMENTOS=("VALOR", "size"),
        )
        .sort_values("ANO_MES")
    )
    mensal["TOTAL_SAIDAS_ABS"] = mensal["TOTAL_SAIDAS"].abs()
    mensal["RESULTADO_LIQUIDO"] = mensal["TOTAL_ENTRADAS"] + mensal["TOTAL_SAIDAS"]

    top_despesas = saidas.sort_values("GASTO", ascending=False).head(20)
    top_entradas = entradas.sort_values("ENTRADA", ascending=False).head(20)

    indicadores = pd.DataFrame(
        [
            ["Total de entradas", df["ENTRADA"].sum()],
            ["Total de saídas", df["SAÍDA"].sum()],
            ["Total de saídas em módulo", saidas["GASTO"].sum()],
            ["Resultado líquido", df["VALOR"].sum()],
            ["Quantidade de lançamentos", len(df)],
            ["Quantidade de documentos", df["DOCUMENTO"].nunique()],
            ["Quantidade de planos financeiros", df["PLANO FINANCEIRO"].nunique()],
        ],
        columns=["INDICADOR", "VALOR"],
    )

    return {
        "Indicadores": indicadores,
        "Movimentacoes": df,
        "Resumo_Documento": resumo_documento,
        "Resumo_Plano": resumo_plano,
        "Documento_Plano": documento_plano,
        "Resumo_Mensal": mensal,
        "Top_Despesas": top_despesas,
        "Top_Entradas": top_entradas,
    }


def formatar_planilha_excel(writer: pd.ExcelWriter, abas: dict[str, pd.DataFrame]) -> None:
    """Aplica formatação simples ao arquivo Excel gerado."""
    workbook = writer.book

    formato_moeda = workbook.add_format({"num_format": 'R$ #,##0.00;[Red]-R$ #,##0.00'})
    formato_data = workbook.add_format({"num_format": "dd/mm/yyyy"})
    formato_cabecalho = workbook.add_format(
        {"bold": True, "bg_color": "#D9EAF7", "border": 1, "text_wrap": True}
    )

    for nome_aba, df in abas.items():
        worksheet = writer.sheets[nome_aba]
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, max(len(df), 1), max(len(df.columns) - 1, 0))

        for col_idx, coluna in enumerate(df.columns):
            worksheet.write(0, col_idx, coluna, formato_cabecalho)
            largura = min(max(len(str(coluna)) + 2, 12), 35)
            worksheet.set_column(col_idx, col_idx, largura)

            if coluna in {"ENTRADA", "SAÍDA", "SALDO", "VALOR", "GASTO", "TOTAL_GASTO", "TOTAL_ENTRADAS", "TOTAL_SAIDAS", "TOTAL_SAIDAS_ABS", "RESULTADO_LIQUIDO", "SALDO_FINAL"}:
                worksheet.set_column(col_idx, col_idx, 16, formato_moeda)
            elif coluna == "DATA":
                worksheet.set_column(col_idx, col_idx, 14, formato_data)


def salvar_relatorio_excel(abas: dict[str, pd.DataFrame], caminho_saida: Path) -> None:
    """Salva as tabelas de análise em um arquivo Excel."""
    with pd.ExcelWriter(caminho_saida, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        for nome_aba, df in abas.items():
            # Excel limita nome de aba a 31 caracteres.
            df.to_excel(writer, sheet_name=nome_aba[:31], index=False)
        formatar_planilha_excel(writer, abas)


def salvar_grafico_barras(df: pd.DataFrame, coluna_categoria: str, coluna_valor: str, titulo: str, caminho: Path, top_n: int = 10) -> None:
    """Gera gráfico de barras horizontal."""
    dados = df.sort_values(coluna_valor, ascending=False).head(top_n).sort_values(coluna_valor)
    if dados.empty:
        return

    plt.figure(figsize=(11, 6))
    plt.barh(dados[coluna_categoria], dados[coluna_valor])
    plt.title(titulo)
    plt.xlabel("Valor em R$")
    plt.tight_layout()
    plt.savefig(caminho, dpi=150)
    plt.close()


def gerar_graficos(abas: dict[str, pd.DataFrame], pasta_saida: Path) -> None:
    """Gera gráficos PNG com base nos resumos."""
    salvar_grafico_barras(
        abas["Resumo_Documento"],
        "DOCUMENTO",
        "TOTAL_GASTO",
        "Top gastos por documento",
        pasta_saida / "grafico_gastos_por_documento.png",
    )

    salvar_grafico_barras(
        abas["Resumo_Plano"],
        "PLANO FINANCEIRO",
        "TOTAL_GASTO",
        "Top gastos por plano financeiro",
        pasta_saida / "grafico_gastos_por_plano.png",
    )

    mensal = abas["Resumo_Mensal"]
    if not mensal.empty:
        plt.figure(figsize=(11, 6))
        plt.plot(mensal["ANO_MES"], mensal["TOTAL_ENTRADAS"], marker="o", label="Entradas")
        plt.plot(mensal["ANO_MES"], mensal["TOTAL_SAIDAS_ABS"], marker="o", label="Saídas")
        plt.title("Entradas x Saídas por mês")
        plt.xlabel("Mês")
        plt.ylabel("Valor em R$")
        plt.xticks(rotation=45)
        plt.legend()
        plt.tight_layout()
        plt.savefig(pasta_saida / "grafico_entradas_saidas_por_mes.png", dpi=150)
        plt.close()

    movimentacoes = abas["Movimentacoes"].sort_values("DATA")
    if not movimentacoes.empty:
        plt.figure(figsize=(11, 6))
        plt.plot(movimentacoes["DATA"], movimentacoes["SALDO"], marker="o")
        plt.title("Evolução do saldo")
        plt.xlabel("Data")
        plt.ylabel("Saldo em R$")
        plt.xticks(rotation=45)
        plt.tight_layout()
        plt.savefig(pasta_saida / "grafico_evolucao_saldo.png", dpi=150)
        plt.close()


def imprimir_resumo_console(abas: dict[str, pd.DataFrame], pasta_saida: Path) -> None:
    """Mostra um resumo simples no terminal."""
    indicadores = abas["Indicadores"]
    print("\nResumo geral")
    print("-" * 60)
    for _, linha in indicadores.iterrows():
        valor = linha["VALOR"]
        if isinstance(valor, (int, float)) and "Quantidade" not in linha["INDICADOR"]:
            print(f"{linha['INDICADOR']}: R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        else:
            print(f"{linha['INDICADOR']}: {valor}")

    print("\nArquivos gerados em:")
    print(pasta_saida.resolve())


def main() -> int:
    parser = argparse.ArgumentParser(description="Analisador simples de extrato bancário empresarial.")
    parser.add_argument("--arquivo", required=True, help="Caminho da planilha .xlsx do extrato.")
    parser.add_argument("--aba", default=None, help="Nome da aba da planilha. Se vazio, usa a primeira aba.")
    parser.add_argument("--saida", default="saida_analise_extrato", help="Pasta onde os relatórios serão salvos.")
    parser.add_argument("--documento", default=None, help="Filtrar por documento. Ex: OBRA, FABR, GERA.")
    parser.add_argument("--plano", default=None, help="Filtrar por plano financeiro. Ex: COMBUSTIVEL.")
    parser.add_argument("--onde-lancar", default=None, help="Filtrar por ONDE LANÇAR. Ex: CONTAS A PAGAR.")
    parser.add_argument("--texto", default=None, help="Buscar texto na descrição ou pagamento.")
    parser.add_argument("--inicio", default=None, help="Data inicial no formato AAAA-MM-DD.")
    parser.add_argument("--fim", default=None, help="Data final no formato AAAA-MM-DD.")

    args = parser.parse_args()

    try:
        arquivo = Path(args.arquivo)
        pasta_saida = Path(args.saida)
        pasta_saida.mkdir(parents=True, exist_ok=True)

        df = carregar_extrato(arquivo, aba=args.aba)
        df_filtrado = aplicar_filtros(
            df,
            documento=args.documento,
            plano=args.plano,
            onde_lancar=args.onde_lancar,
            texto=args.texto,
            inicio=args.inicio,
            fim=args.fim,
        )

        abas = criar_resumos(df_filtrado)
        caminho_excel = pasta_saida / "relatorio_analise_extrato.xlsx"
        salvar_relatorio_excel(abas, caminho_excel)
        gerar_graficos(abas, pasta_saida)
        imprimir_resumo_console(abas, pasta_saida)

        return 0
    except Exception as erro:
        print(f"Erro: {erro}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
