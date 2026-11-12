"""
Módulo de exportação e estilização avançada em Excel (openpyxl).
Gera um arquivo .xlsx executivo com 3 abas:
  1. Resumo Executivo (KPIs, tabelas agregadas e gráfico nativo)
  2. Base Consolidada (Dados higienizados, tipados e formatados)
  3. Auditoria de Erros (Quarentena de linhas rejeitadas com diagnóstico)
"""

from pathlib import Path
from typing import Any, Dict
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# Paleta de cores corporativa
COR_AZUL_ESCURO = "1F4E78"
COR_AZUL_MEDIO = "2F75B5"
COR_ZEBRA = "F2F6FA"
COR_VERDE_CLARO = "E2EFDA"
COR_VERMELHO_CLARO = "FCE4D6"
COR_VERMELHO_HEADER = "C00000"
COR_CINZA_BORDA = "D9D9D9"

FONTE_TITULO = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
FONTE_SUBTITULO = Font(name="Calibri", size=11, bold=True, color=COR_AZUL_ESCURO)
FONTE_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONTE_CORPO = Font(name="Calibri", size=10, color="111111")
FONTE_NEGRITO = Font(name="Calibri", size=10, bold=True, color="111111")
FONTE_ERRO = Font(name="Calibri", size=10, bold=True, color="9C0006")

FILL_TITULO = PatternFill(start_color=COR_AZUL_ESCURO, end_color=COR_AZUL_ESCURO, fill_type="solid")
FILL_HEADER = PatternFill(start_color=COR_AZUL_MEDIO, end_color=COR_AZUL_MEDIO, fill_type="solid")
FILL_HEADER_ERRO = PatternFill(start_color=COR_VERMELHO_HEADER, end_color=COR_VERMELHO_HEADER, fill_type="solid")
FILL_ZEBRA = PatternFill(start_color=COR_ZEBRA, end_color=COR_ZEBRA, fill_type="solid")
FILL_SUCESSO = PatternFill(start_color=COR_VERDE_CLARO, end_color=COR_VERDE_CLARO, fill_type="solid")
FILL_ALERTA = PatternFill(start_color=COR_VERMELHO_CLARO, end_color=COR_VERMELHO_CLARO, fill_type="solid")

BORDA_FINA = Border(
    left=Side(style="thin", color=COR_CINZA_BORDA),
    right=Side(style="thin", color=COR_CINZA_BORDA),
    top=Side(style="thin", color=COR_CINZA_BORDA),
    bottom=Side(style="thin", color=COR_CINZA_BORDA),
)

FORMATO_MOEDA = 'R$ #,##0.00'
FORMATO_INTEIRO = '#,##0'
FORMATO_PERCENTUAL = '0.0%'
FORMATO_DATA = 'DD/MM/YYYY'


def ajustar_largura_colunas(ws, largura_minima: int = 12, largura_maxima: int = 48) -> None:
    """Ajusta automaticamente a largura de todas as colunas da aba."""
    for col in ws.columns:
        letra_col = get_column_letter(col[0].column)
        maior_tamanho = 0
        for celula in col:
            # Ignora linha 1 se for título mesclado longo
            if celula.row == 1 and ws.merged_cells:
                continue
            valor_str = str(celula.value or "")
            maior_tamanho = max(maior_tamanho, len(valor_str))
        largura_ajustada = max(largura_minima, min(maior_tamanho + 4, largura_maxima))
        ws.column_dimensions[letra_col].width = largura_ajustada


def construir_aba_resumo(
    wb: Workbook,
    df_validos: pd.DataFrame,
    metricas: Dict[str, Any],
) -> None:
    """Cria a primeira aba com indicadores executivos (KPIs), tabelas por filial e gráfico."""
    ws = wb.active
    ws.title = "Resumo Executivo"
    ws.views.sheetView[0].showGridLines = True

    # Banner Principal
    ws.merge_cells("A1:F1")
    celula_titulo = ws["A1"]
    celula_titulo.value = "RELATÓRIO EXECUTIVO — CONSOLIDAÇÃO E AUDITORIA DE VENDAS"
    celula_titulo.font = FONTE_TITULO
    celula_titulo.fill = FILL_TITULO
    celula_titulo.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 34

    # Bloco 1: Indicadores de Qualidade & ETL
    ws["A3"] = "1. INDICADORES DO PROCESSAMENTO (PIPELINE ETL)"
    ws["A3"].font = FONTE_SUBTITULO

    cabecalhos_kpi = ["Indicador", "Resultado"]
    for col_idx, texto in enumerate(cabecalhos_kpi, start=1):
        c = ws.cell(row=4, column=col_idx, value=texto)
        c.font = FONTE_HEADER
        c.fill = FILL_HEADER
        c.alignment = Alignment(horizontal="left" if col_idx == 1 else "right", vertical="center")

    linhas_kpi = [
        ("Arquivos de Origem Processados", metricas["arquivos_processados"], FORMATO_INTEIRO),
        ("Total de Linhas Brutas Lidas", metricas["total_linhas_lidas"], FORMATO_INTEIRO),
        ("Registros Válidos e Higienizados", metricas["registros_validos"], FORMATO_INTEIRO),
        ("Registros Rejeitados (Aba Auditoria)", metricas["registros_rejeitados"], FORMATO_INTEIRO),
        ("Duplicidades Bloqueadas", metricas["duplicatas_removidas"], FORMATO_INTEIRO),
        ("Taxa de Qualidade / Aproveitamento", metricas["taxa_aproveitamento"] / 100.0, FORMATO_PERCENTUAL),
        ("Faturamento Total Consolidado", metricas["faturamento_total"], FORMATO_MOEDA),
        ("Ticket Médio por Venda", metricas["ticket_medio"], FORMATO_MOEDA),
    ]

    for idx, (nome_kpi, valor_kpi, formato) in enumerate(linhas_kpi, start=5):
        c_nome = ws.cell(row=idx, column=1, value=nome_kpi)
        c_val = ws.cell(row=idx, column=2, value=valor_kpi)

        c_nome.font = FONTE_NEGRITO if "Faturamento" in nome_kpi else FONTE_CORPO
        c_val.font = FONTE_NEGRITO
        c_val.number_format = formato
        c_val.alignment = Alignment(horizontal="right")

        c_nome.border = BORDA_FINA
        c_val.border = BORDA_FINA

        if "Rejeitados" in nome_kpi and valor_kpi > 0:
            c_val.fill = FILL_ALERTA
        elif "Faturamento" in nome_kpi or "Válidos" in nome_kpi:
            c_val.fill = FILL_SUCESSO

    # Bloco 2: Desempenho por Filial
    linha_inicio_filial = 15
    ws.cell(row=linha_inicio_filial, column=1, value="2. DESEMPENHO CONSOLIDADO POR FILIAL").font = FONTE_SUBTITULO

    headers_filial = [
        "Filial",
        "Qtd. Pedidos",
        "Itens Vendidos",
        "Faturamento Total",
        "Ticket Médio",
        "% Participação",
    ]
    linha_header_filial = linha_inicio_filial + 1
    for col_idx, h in enumerate(headers_filial, start=1):
        c = ws.cell(row=linha_header_filial, column=col_idx, value=h)
        c.font = FONTE_HEADER
        c.fill = FILL_HEADER
        c.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")

    if not df_validos.empty:
        resumo_filial = (
            df_validos.groupby("Filial", as_index=False)
            .agg(
                Qtd_Pedidos=("ID_Venda", "count"),
                Itens_Vendidos=("Quantidade", "sum"),
                Faturamento=("Valor_Total", "sum"),
                Ticket_Medio=("Valor_Total", "mean"),
            )
            .sort_values(by="Faturamento", ascending=False)
        )
        faturamento_geral = resumo_filial["Faturamento"].sum() or 1.0

        linha_atual = linha_header_filial + 1
        for _, row in resumo_filial.iterrows():
            participacao = float(row["Faturamento"]) / faturamento_geral
            valores = [
                row["Filial"],
                int(row["Qtd_Pedidos"]),
                int(row["Itens_Vendidos"]),
                float(row["Faturamento"]),
                float(row["Ticket_Medio"]),
                participacao,
            ]
            formatos = [None, FORMATO_INTEIRO, FORMATO_INTEIRO, FORMATO_MOEDA, FORMATO_MOEDA, FORMATO_PERCENTUAL]

            for col_idx, (val, fmt) in enumerate(zip(valores, formatos), start=1):
                c = ws.cell(row=linha_atual, column=col_idx, value=val)
                c.font = FONTE_CORPO
                c.border = BORDA_FINA
                if fmt:
                    c.number_format = fmt
                    c.alignment = Alignment(horizontal="right")
                if linha_atual % 2 == 1:
                    c.fill = FILL_ZEBRA
            linha_atual += 1

        # Gráfico de Barras Nativo no Excel
        grafico = BarChart()
        grafico.type = "col"
        grafico.style = 10
        grafico.title = "Faturamento por Filial (R$)"
        grafico.y_axis.title = "Faturamento (R$)"
        grafico.x_axis.title = "Filial"
        grafico.legend = None
        grafico.width = 16
        grafico.height = 10

        dados_ref = Reference(
            ws,
            min_col=4,
            min_row=linha_header_filial,
            max_row=linha_atual - 1,
        )
        categorias_ref = Reference(
            ws,
            min_col=1,
            min_row=linha_header_filial + 1,
            max_row=linha_atual - 1,
        )
        grafico.add_data(dados_ref, titles_from_data=True)
        grafico.set_categories(categorias_ref)
        ws.add_chart(grafico, "D3")

    ajustar_largura_colunas(ws)


def construir_aba_consolidada(wb: Workbook, df_validos: pd.DataFrame) -> None:
    """Cria a segunda aba com todos os registros higienizados e formatados."""
    ws = wb.create_sheet(title="Base Consolidada")
    ws.views.sheetView[0].showGridLines = True

    colunas = [
        "ID_Venda",
        "Data_Venda",
        "Filial",
        "Nome_Cliente",
        "CPF_Cliente",
        "Produto",
        "Categoria",
        "Quantidade",
        "Valor_Unitario",
        "Valor_Total",
        "Status_Pagamento",
        "Arquivo_Origem",
    ]

    headers_amigaveis = [
        "ID Venda",
        "Data da Venda",
        "Filial",
        "Cliente",
        "CPF",
        "Produto",
        "Categoria",
        "Qtd",
        "Valor Unitário",
        "Valor Total",
        "Status",
        "Arquivo de Origem",
    ]

    for col_idx, titulo in enumerate(headers_amigaveis, start=1):
        c = ws.cell(row=1, column=col_idx, value=titulo)
        c.font = FONTE_HEADER
        c.fill = FILL_HEADER
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    for row_idx, (_, linha) in enumerate(df_validos.iterrows(), start=2):
        for col_idx, nome_col in enumerate(colunas, start=1):
            valor = linha.get(nome_col)
            c = ws.cell(row=row_idx, column=col_idx, value=valor)
            c.font = FONTE_CORPO
            c.border = BORDA_FINA

            if row_idx % 2 == 1:
                c.fill = FILL_ZEBRA

            if nome_col == "Data_Venda":
                c.number_format = FORMATO_DATA
                c.alignment = Alignment(horizontal="center")
            elif nome_col in {"ID_Venda", "CPF_Cliente", "Status_Pagamento"}:
                c.alignment = Alignment(horizontal="center")
            elif nome_col == "Quantidade":
                c.number_format = FORMATO_INTEIRO
                c.alignment = Alignment(horizontal="right")
            elif nome_col in {"Valor_Unitario", "Valor_Total"}:
                c.number_format = FORMATO_MOEDA
                c.alignment = Alignment(horizontal="right")

    ws.freeze_panes = "A2"
    if not df_validos.empty:
        ultima_col = get_column_letter(len(colunas))
        ws.auto_filter.ref = f"A1:{ultima_col}{len(df_validos) + 1}"

    ajustar_largura_colunas(ws)


def construir_aba_auditoria(wb: Workbook, df_auditoria: pd.DataFrame) -> None:
    """Cria a terceira aba detalhando cada linha rejeitada e o motivo exato."""
    ws = wb.create_sheet(title="Auditoria de Erros")
    ws.views.sheetView[0].showGridLines = True

    colunas_auditoria = [
        "Arquivo_Origem",
        "Linha_Origem",
        "ID_Venda",
        "Data_Original",
        "Filial",
        "Cliente",
        "CPF_Original",
        "Produto",
        "Qtd_Original",
        "Valor_Original",
        "Motivo_Rejeicao",
    ]

    headers_auditoria = [
        "Arquivo Origem",
        "Linha Excel",
        "ID Venda",
        "Data Informada",
        "Filial",
        "Cliente",
        "CPF Informado",
        "Produto",
        "Qtd Informada",
        "Valor Informado",
        "Diagnóstico / Motivo da Rejeição",
    ]

    for col_idx, titulo in enumerate(headers_auditoria, start=1):
        c = ws.cell(row=1, column=col_idx, value=titulo)
        c.font = FONTE_HEADER
        c.fill = FILL_HEADER_ERRO
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    for row_idx, (_, linha) in enumerate(df_auditoria.iterrows(), start=2):
        for col_idx, nome_col in enumerate(colunas_auditoria, start=1):
            valor = linha.get(nome_col)
            if pd.isna(valor):
                valor = "[VAZIO]"
            c = ws.cell(row=row_idx, column=col_idx, value=valor)
            c.border = BORDA_FINA

            if nome_col == "Motivo_Rejeicao":
                c.font = FONTE_ERRO
                c.fill = FILL_ALERTA
            else:
                c.font = FONTE_CORPO
                if row_idx % 2 == 1:
                    c.fill = FILL_ZEBRA

            if nome_col in {"Linha_Origem", "ID_Venda", "CPF_Original"}:
                c.alignment = Alignment(horizontal="center")

    ws.freeze_panes = "A2"
    if not df_auditoria.empty:
        ultima_col = get_column_letter(len(colunas_auditoria))
        ws.auto_filter.ref = f"A1:{ultima_col}{len(df_auditoria) + 1}"

    ajustar_largura_colunas(ws, largura_maxima=60)


def exportar_relatorio_excel(
    caminho_saida: Path,
    df_validos: pd.DataFrame,
    df_auditoria: pd.DataFrame,
    metricas: Dict[str, Any],
) -> Path:
    """Gera e salva o arquivo Excel final com todas as abas formatadas."""
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    construir_aba_resumo(wb, df_validos, metricas)
    construir_aba_consolidada(wb, df_validos)
    construir_aba_auditoria(wb, df_auditoria)

    wb.save(caminho_saida)
    return caminho_saida
