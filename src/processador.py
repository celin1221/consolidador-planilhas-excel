"""
Módulo de processamento em lote (Batch ETL).
Varre o diretório de entrada ou lista de arquivos, carrega planilhas Excel (.xlsx) e CSV (.csv),
padroniza schemas, remove duplicidades e separa registros válidos dos rejeitados.
"""

from pathlib import Path
from typing import Any, Dict, Iterable, List, Set, Tuple
import pandas as pd

from src.validador import (
    COLUNAS_OBRIGATORIAS,
    padronizar_colunas,
    validar_registro,
)


def carregar_planilha(caminho: Path) -> pd.DataFrame:
    """Carrega um arquivo .xlsx ou .csv tratando codificações comuns."""
    sufixo = caminho.suffix.lower()
    if sufixo == ".xlsx":
        return pd.read_excel(caminho, dtype=str)
    if sufixo == ".csv":
        for encoding in ("utf-8-sig", "utf-8", "latin-1"):
            try:
                return pd.read_csv(caminho, dtype=str, sep=None, engine="python", encoding=encoding)
            except UnicodeDecodeError:
                continue
        raise ValueError(f"Não foi possível decodificar o arquivo CSV: {caminho.name}")
    raise ValueError(f"Formato de arquivo não suportado: {caminho.name}")


def processar_arquivos(
    arquivos: Iterable[Path],
) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    """
    Processa uma coleção de arquivos (.xlsx / .csv).
    Retorna:
      - df_validos: DataFrame pronto para análise
      - df_auditoria: DataFrame com as linhas rejeitadas e o motivo detalhado
      - metricas: Dicionário com estatísticas do processamento
    """
    lista_arquivos = sorted(
        [
            Path(p)
            for p in arquivos
            if Path(p).is_file()
            and Path(p).suffix.lower() in {".xlsx", ".csv"}
            and not Path(p).name.startswith("~$")
        ]
    )

    registros_validos: List[Dict[str, Any]] = []
    registros_auditoria: List[Dict[str, Any]] = []
    ids_ja_processados: Set[str] = set()

    total_linhas_lidas = 0
    total_duplicatas = 0
    arquivos_processados = 0

    for arquivo in lista_arquivos:
        df_bruto = carregar_planilha(arquivo)
        # Remove linhas 100% vazias
        df_bruto = df_bruto.dropna(how="all")
        df_padronizado = padronizar_colunas(df_bruto)
        arquivos_processados += 1

        # Verifica se falta alguma coluna obrigatória na planilha inteira
        colunas_faltantes = [
            col for col in COLUNAS_OBRIGATORIAS if col not in df_padronizado.columns
        ]

        for idx, serie_linha in df_padronizado.iterrows():
            total_linhas_lidas += 1
            numero_linha_excel = int(idx) + 2  # +2 considerando cabeçalho na linha 1
            dados_linha = serie_linha.to_dict()

            if colunas_faltantes:
                registros_auditoria.append(
                    {
                        "Arquivo_Origem": arquivo.name,
                        "Linha_Origem": numero_linha_excel,
                        "ID_Venda": dados_linha.get("ID_Venda", "N/A"),
                        "Data_Original": str(dados_linha.get("Data_Venda", "")),
                        "Filial": str(dados_linha.get("Filial", "")),
                        "Cliente": str(dados_linha.get("Nome_Cliente", "")),
                        "CPF_Original": str(dados_linha.get("CPF_Cliente", "")),
                        "Produto": str(dados_linha.get("Produto", "")),
                        "Qtd_Original": str(dados_linha.get("Quantidade", "")),
                        "Valor_Original": str(dados_linha.get("Valor_Unitario", "")),
                        "Motivo_Rejeicao": f"Colunas ausentes na planilha: {', '.join(colunas_faltantes)}",
                    }
                )
                continue

            registro_limpo, lista_erros = validar_registro(dados_linha)

            # Verificação de duplicidade pelo ID_Venda
            id_venda = registro_limpo.get("ID_Venda")
            if id_venda and not lista_erros:
                if id_venda in ids_ja_processados:
                    lista_erros.append(f"Registro duplicado (ID '{id_venda}' já contabilizado)")
                    total_duplicatas += 1
                else:
                    ids_ja_processados.add(id_venda)

            if lista_erros:
                registros_auditoria.append(
                    {
                        "Arquivo_Origem": arquivo.name,
                        "Linha_Origem": numero_linha_excel,
                        "ID_Venda": dados_linha.get("ID_Venda"),
                        "Data_Original": dados_linha.get("Data_Venda"),
                        "Filial": dados_linha.get("Filial"),
                        "Cliente": dados_linha.get("Nome_Cliente"),
                        "CPF_Original": dados_linha.get("CPF_Cliente"),
                        "Produto": dados_linha.get("Produto"),
                        "Qtd_Original": dados_linha.get("Quantidade"),
                        "Valor_Original": dados_linha.get("Valor_Unitario"),
                        "Motivo_Rejeicao": " | ".join(lista_erros),
                    }
                )
            else:
                registro_limpo["Arquivo_Origem"] = arquivo.name
                registros_validos.append(registro_limpo)

    df_validos = pd.DataFrame(registros_validos)
    if not df_validos.empty:
        df_validos = df_validos.sort_values(by=["Data_Venda", "ID_Venda"]).reset_index(drop=True)

    df_auditoria = pd.DataFrame(registros_auditoria)

    faturamento_total = (
        float(df_validos["Valor_Total"].sum()) if not df_validos.empty else 0.0
    )
    ticket_medio = (
        float(df_validos["Valor_Total"].mean()) if not df_validos.empty else 0.0
    )

    metricas = {
        "arquivos_processados": arquivos_processados,
        "total_linhas_lidas": total_linhas_lidas,
        "registros_validos": len(df_validos),
        "registros_rejeitados": len(df_auditoria),
        "duplicatas_removidas": total_duplicatas,
        "taxa_aproveitamento": (
            round((len(df_validos) / total_linhas_lidas) * 100, 1)
            if total_linhas_lidas > 0
            else 0.0
        ),
        "faturamento_total": round(faturamento_total, 2),
        "ticket_medio": round(ticket_medio, 2),
    }

    return df_validos, df_auditoria, metricas


def processar_diretorio(
    pasta_entrada: Path,
) -> Tuple[pd.DataFrame, pd.DataFrame, Dict[str, Any]]:
    """Processa todas as planilhas (.xlsx e .csv) contidas na pasta informada."""
    if not pasta_entrada.exists():
        return processar_arquivos([])
    arquivos = [p for p in pasta_entrada.iterdir() if p.is_file()]
    return processar_arquivos(arquivos)
