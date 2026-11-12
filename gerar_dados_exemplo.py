"""
Script gerador de massa de dados fictícia ("suja") para demonstrar o funcionamento do pipeline ETL.
Cria 3 arquivos em `dados_entrada/` (.xlsx e .csv) simulando planilhas enviadas por diferentes
filiais com colunas fora de padrão, CPFs inválidos, datas mal formatadas e duplicidades.
"""

from pathlib import Path
import pandas as pd


def gerar_planilhas_exemplo(pasta_destino: Path = Path("dados_entrada")) -> None:
    """Cria arquivos de exemplo com dados válidos e inconsistências propositais."""
    pasta_destino.mkdir(parents=True, exist_ok=True)

    # 1. Filial São Paulo (.xlsx) — Colunas tipo 1, valores com 'R$', espaços extras e erros de CPF
    dados_sp = [
        {
            "ID_Venda": "VND-1001",
            "Data_Venda": "10/03/2026",
            "Filial": "  são paulo - paulista ",
            "Nome_Cliente": "  JOÃO DA SILVA NETO ",
            "CPF_Cliente": "52998224725",
            "Produto": "notebook dell inspiron 15",
            "Categoria": "informática",
            "Quantidade": 2,
            "Valor_Unitario": "R$ 4.250,00",
            "Status_Pagamento": "pago",
        },
        {
            "ID_Venda": "VND-1002",
            "Data_Venda": "11/03/2026",
            "Filial": "São Paulo - Paulista",
            "Nome_Cliente": "maria fernanda de oliveira",
            "CPF_Cliente": "111.444.777-35",
            "Produto": "Monitor LG UltraWide 29",
            "Categoria": "Informática",
            "Quantidade": 2,
            "Valor_Unitario": "1.399,90",
            "Status_Pagamento": "Aprovado",
        },
        {
            "ID_Venda": "VND-1003",
            "Data_Venda": "12/03/2026",
            "Filial": "São Paulo - Paulista",
            "Nome_Cliente": "Carlos Eduardo Santos",
            "CPF_Cliente": "123.456.789-00",  # ERRO PROPOSITADO: Dígito verificador inválido
            "Produto": "Teclado Mecânico Sem Fio",
            "Categoria": "Periféricos",
            "Quantidade": 3,
            "Valor_Unitario": "450,00",
            "Status_Pagamento": "Pago",
        },
        {
            "ID_Venda": "VND-1004",
            "Data_Venda": "2026-03-14",
            "Filial": "São Paulo - Paulista",
            "Nome_Cliente": "Ana Beatriz Costa",
            "CPF_Cliente": "98765432100",
            "Produto": "Cadeira Ergonômica Office",
            "Categoria": "Escritório",
            "Quantidade": 2,
            "Valor_Unitario": 1290.50,
            "Status_Pagamento": "Concluído",
        },
        {
            "ID_Venda": "VND-1001",  # ERRO PROPOSITADO: Duplicata exata do primeiro pedido
            "Data_Venda": "10/03/2026",
            "Filial": "São Paulo - Paulista",
            "Nome_Cliente": "João da Silva Neto",
            "CPF_Cliente": "529.982.247-25",
            "Produto": "Notebook Dell Inspiron 15",
            "Categoria": "Informática",
            "Quantidade": 2,
            "Valor_Unitario": "4250.00",
            "Status_Pagamento": "Pago",
        },
    ]

    # 2. Filial Rio de Janeiro (.xlsx) — Nomes de colunas totalmente diferentes e erros operacionais
    dados_rj = [
        {
            "Num_Pedido": "VND-2001",
            "Dt_Venda": "12/03/2026",
            "Unidade": "rio de janeiro - centro",
            "Comprador": "pedro henrique almeida",
            "Documento": "705.484.450-52",
            "Mercadoria": "smartphone samsung galaxy s24",
            "Setor": "telefonia",
            "Qtde": 1,
            "Preco_Unitario": "R$ 3.899,00",
            "Situacao": "pago",
        },
        {
            "Num_Pedido": "VND-2002",
            "Dt_Venda": "13-03-2026",
            "Unidade": "Rio de Janeiro - Centro",
            "Comprador": "Juliana Ribeiro Dos Santos",
            "Documento": "12345678909",
            "Mercadoria": "Fone Bluetooth Noise Cancelling",
            "Setor": "Áudio",
            "Qtde": 4,
            "Preco_Unitario": "650,00",
            "Situacao": "Aguardando",
        },
        {
            "Num_Pedido": "VND-2003",
            "Dt_Venda": "31/02/2026",  # ERRO PROPOSITADO: Data inexistente
            "Unidade": "Rio de Janeiro - Centro",
            "Comprador": "Roberto Gomes",
            "Documento": "111.111.111-11",  # ERRO PROPOSITADO: CPF com dígitos repetidos
            "Mercadoria": "Mouse Gamer RGB",
            "Setor": "Periféricos",
            "Qtde": 1,
            "Preco_Unitario": "-150,00",  # ERRO PROPOSITADO: Valor negativo
            "Situacao": "Pago",
        },
        {
            "Num_Pedido": "VND-2004",
            "Dt_Venda": "15/03/2026",
            "Unidade": "Rio de Janeiro - Centro",
            "Comprador": "   ",  # ERRO PROPOSITADO: Nome do cliente vazio
            "Documento": "70548445052",
            "Mercadoria": "SSD NVMe 1TB Kingston",
            "Setor": "Hardware",
            "Qtde": 2,
            "Preco_Unitario": "499,90",
            "Situacao": "Pago",
        },
        {
            "Num_Pedido": "VND-2005",
            "Dt_Venda": "16/03/2026",
            "Unidade": "Rio de Janeiro - Centro",
            "Comprador": "camila mendes rodrigues",
            "Documento": "390.533.447-05",
            "Mercadoria": "tablet samsung tab s9",
            "Setor": "telefonia",
            "Qtde": 2,
            "Preco_Unitario": "R$ 2.750,00",
            "Situacao": "Pago",
        },
    ]

    # 3. Filial Curitiba (.csv) — Arquivo CSV com ponto-e-vírgula e outros nomes de colunas
    dados_cwb = [
        {
            "Codigo_Venda": "VND-3001",
            "Data_Emissao": "2026-03-11",
            "Loja": "curitiba - batel",
            "Cliente": "lucas martins da rocha",
            "CPF": "111.444.777-35",
            "Item": "macbook air m3 16gb",
            "Departamento": "informática",
            "Qtd": 1,
            "Vlr_Unitario": "R$ 8.999,00",
            "Status": "Pago",
        },
        {
            "Codigo_Venda": "VND-3002",
            "Data_Emissao": "16/03/2026",
            "Loja": "Curitiba - Batel",
            "Cliente": "Fernanda Lima de Souza",
            "CPF": "39053344705",
            "Item": "Dock Station USB-C Duplo HDMI",
            "Departamento": "Periféricos",
            "Qtd": 3,
            "Vlr_Unitario": "420,00",
            "Status": "Concluido",
        },
        {
            "Codigo_Venda": "VND-3003",
            "Data_Emissao": "17/03/2026",
            "Loja": "Curitiba - Batel",
            "Cliente": "Marcos Vinicius Pereira",
            "CPF": "52998224725",
            "Item": "Webcam Full HD 1080p",
            "Departamento": "Periféricos",
            "Qtd": 0,  # ERRO PROPOSITADO: Quantidade igual a zero
            "Vlr_Unitario": "310,00",
            "Status": "Pago",
        },
    ]

    pd.DataFrame(dados_sp).to_excel(
        pasta_destino / "vendas_filial_sao_paulo.xlsx", index=False
    )
    pd.DataFrame(dados_rj).to_excel(
        pasta_destino / "vendas_filial_rio_de_janeiro.xlsx", index=False
    )
    pd.DataFrame(dados_cwb).to_csv(
        pasta_destino / "vendas_filial_curitiba.csv", index=False, sep=";", encoding="utf-8-sig"
    )


if __name__ == "__main__":
    gerar_planilhas_exemplo()
    print("Planilhas de exemplo geradas com sucesso na pasta 'dados_entrada/'!")
