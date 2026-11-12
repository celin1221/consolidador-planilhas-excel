"""
Ponto de entrada principal do Pipeline ETL de Consolidação e Higienização de Planilhas.
Executa a leitura de múltiplos arquivos (.xlsx / .csv), valida regras de negócio
e gera o Relatório Executivo Consolidado em Excel.
"""

from pathlib import Path
from gerar_dados_exemplo import gerar_planilhas_exemplo
from src.exportador_excel import exportar_relatorio_excel
from src.processador import processar_diretorio


def executar_pipeline() -> None:
    pasta_entrada = Path("dados_entrada")
    caminho_saida = Path("dados_saida/Relatorio_Consolidado_Vendas.xlsx")

    print("=" * 68)
    print("  PIPELINE ETL — CONSOLIDADOR E HIGIENIZADOR DE PLANILHAS EXCEL")
    print("=" * 68)

    # Se não houver arquivos na pasta de entrada, gera automaticamente os dados de demonstração
    arquivos_existentes = (
        [
            p
            for p in pasta_entrada.iterdir()
            if p.suffix.lower() in {".xlsx", ".csv"} and not p.name.startswith("~$")
        ]
        if pasta_entrada.exists()
        else []
    )

    if not arquivos_existentes:
        print("\n[INFO] Pasta 'dados_entrada/' vazia. Gerando planilhas de demonstração...")
        gerar_planilhas_exemplo(pasta_entrada)

    print(f"\n[1/3] Lendo e padronizando planilhas de '{pasta_entrada}/'...")
    df_validos, df_auditoria, metricas = processar_diretorio(pasta_entrada)

    print("[2/3] Gerando relatório executivo formatado em Excel...")
    arquivo_gerado = exportar_relatorio_excel(
        caminho_saida=caminho_saida,
        df_validos=df_validos,
        df_auditoria=df_auditoria,
        metricas=metricas,
    )

    print("[3/3] Processamento concluído com sucesso!\n")
    print("-" * 68)
    print("  RESUMO DA EXECUÇÃO")
    print("-" * 68)
    print(f"  • Arquivos processados         : {metricas['arquivos_processados']}")
    print(f"  • Total de linhas lidas        : {metricas['total_linhas_lidas']}")
    print(f"  • Registros válidos (limpos)   : {metricas['registros_validos']}")
    print(f"  • Registros rejeitados (erros) : {metricas['registros_rejeitados']}")
    print(f"  • Duplicidades bloqueadas      : {metricas['duplicatas_removidas']}")
    print(f"  • Taxa de aproveitamento       : {metricas['taxa_aproveitamento']:.1f}%")
    print(f"  • Faturamento consolidado      : R$ {metricas['faturamento_total']:,.2f}")
    print("-" * 68)
    print(f"  Relatório salvo em: {arquivo_gerado.resolve()}")
    print("=" * 68)


if __name__ == "__main__":
    executar_pipeline()
