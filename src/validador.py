"""
Módulo de higienização, normalização e validação de dados.
Responsável por padronizar colunas e validar campos críticos (CPF, datas, valores monetários).
"""

from datetime import datetime, date
import re
from typing import Any, Dict, List, Optional, Tuple
import unicodedata
import pandas as pd


# Mapeamento de possíveis nomes de colunas nas planilhas das filiais para o padrão oficial
MAPA_COLUNAS: Dict[str, str] = {
    # ID da Venda
    "id_venda": "ID_Venda",
    "id": "ID_Venda",
    "codigo_venda": "ID_Venda",
    "cod_venda": "ID_Venda",
    "num_pedido": "ID_Venda",
    "pedido": "ID_Venda",
    # Data da Venda
    "data": "Data_Venda",
    "data_venda": "Data_Venda",
    "dt_venda": "Data_Venda",
    "data_emissao": "Data_Venda",
    "data_do_pedido": "Data_Venda",
    # Filial
    "filial": "Filial",
    "unidade": "Filial",
    "loja": "Filial",
    # Cliente
    "cliente": "Nome_Cliente",
    "nome_cliente": "Nome_Cliente",
    "nome_do_cliente": "Nome_Cliente",
    "comprador": "Nome_Cliente",
    # CPF
    "cpf": "CPF_Cliente",
    "cpf_cliente": "CPF_Cliente",
    "documento": "CPF_Cliente",
    "doc_cliente": "CPF_Cliente",
    # Produto
    "produto": "Produto",
    "item": "Produto",
    "descricao_produto": "Produto",
    "mercadoria": "Produto",
    # Categoria
    "categoria": "Categoria",
    "departamento": "Categoria",
    "setor": "Categoria",
    # Quantidade
    "quantidade": "Quantidade",
    "qtd": "Quantidade",
    "qtde": "Quantidade",
    # Valor Unitário
    "valor_unitario": "Valor_Unitario",
    "vlr_unitario": "Valor_Unitario",
    "preco_unitario": "Valor_Unitario",
    "preco": "Valor_Unitario",
    "unitario_r": "Valor_Unitario",
    # Status
    "status": "Status_Pagamento",
    "status_pagamento": "Status_Pagamento",
    "situacao": "Status_Pagamento",
}

COLUNAS_OBRIGATORIAS: List[str] = [
    "ID_Venda",
    "Data_Venda",
    "Filial",
    "Nome_Cliente",
    "CPF_Cliente",
    "Produto",
    "Categoria",
    "Quantidade",
    "Valor_Unitario",
]


def normalizar_chave_coluna(nome: Any) -> str:
    """Remove acentos, caracteres especiais e espaços para identificar o nome da coluna."""
    texto = str(nome).strip().lower()
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("utf-8")
    texto = re.sub(r"[^a-z0-9]+", "_", texto).strip("_")
    return texto


def padronizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """Renomeia as colunas do DataFrame para o schema padrão utilizando o dicionário de sinônimos."""
    novas_colunas = {}
    for col in df.columns:
        chave = normalizar_chave_coluna(col)
        if chave in MAPA_COLUNAS:
            novas_colunas[col] = MAPA_COLUNAS[chave]
    return df.rename(columns=novas_colunas)


def limpar_texto(valor: Any, title_case: bool = True) -> Optional[str]:
    """Remove espaços duplicados e padroniza capitalização de textos."""
    if pd.isna(valor) or valor is None:
        return None
    texto = re.sub(r"\s+", " ", str(valor)).strip()
    if not texto or texto.lower() in {"nan", "none", "null", "n/a", "-"}:
        return None
    if title_case:
        # Mantém preposições em minúsculo no meio de nomes
        preposicoes = {"de", "da", "do", "das", "dos", "e", "em", "para", "com"}
        palavras = texto.title().split(" ")
        palavras_formatadas = [
            p.lower() if (i > 0 and p.lower() in preposicoes) else p
            for i, p in enumerate(palavras)
        ]
        return " ".join(palavras_formatadas)
    return texto


def validar_e_formatar_cpf(valor: Any) -> Tuple[Optional[str], Optional[str]]:
    """
    Valida um CPF pelos dígitos verificadores oficiais e retorna formatado (XXX.XXX.XXX-XX).
    Retorna: (cpf_formatado, mensagem_erro)
    """
    if pd.isna(valor) or valor is None:
        return None, "CPF ausente"

    texto = str(valor).strip()
    # Caso o Excel tenha lido como float (ex: 12345678909.0)
    if texto.endswith(".0"):
        texto = texto[:-2]

    digitos = re.sub(r"\D", "", texto)

    if not digitos:
        return None, "CPF ausente"

    # Se perdeu zero à esquerda no Excel, só preenche se tiver entre 9 e 10 dígitos
    if 9 <= len(digitos) < 11:
        digitos = digitos.zfill(11)

    if len(digitos) != 11:
        return None, f"CPF com tamanho inválido ({len(digitos)} dígitos)"

    # Rejeita sequências repetidas como 111.111.111-11
    if digitos == digitos[0] * 11:
        return None, "CPF inválido (dígitos repetidos)"

    # Cálculo do 1º dígito verificador
    soma_1 = sum(int(digitos[i]) * (10 - i) for i in range(9))
    resto_1 = (soma_1 * 10) % 11
    dv_1 = 0 if resto_1 == 10 else resto_1

    # Cálculo do 2º dígito verificador
    soma_2 = sum(int(digitos[i]) * (11 - i) for i in range(10))
    resto_2 = (soma_2 * 10) % 11
    dv_2 = 0 if resto_2 == 10 else resto_2

    if int(digitos[9]) != dv_1 or int(digitos[10]) != dv_2:
        return None, "CPF inválido (dígito verificador incorreto)"

    cpf_formatado = f"{digitos[:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:]}"
    return cpf_formatado, None


def converter_valor_monetario(valor: Any) -> Tuple[Optional[float], Optional[str]]:
    """
    Converte valores em formato texto ('R$ 1.250,50', '1250.50') ou numérico para float.
    Retorna: (valor_float, mensagem_erro)
    """
    if pd.isna(valor) or valor is None:
        return None, "Valor unitário ausente"

    if isinstance(valor, (int, float)):
        val_float = float(valor)
    else:
        texto = str(valor).strip()
        if not texto:
            return None, "Valor unitário vazio"

        # Remove símbolo de moeda e espaços
        texto = re.sub(r"[R$\s]", "", texto, flags=re.IGNORECASE)

        # Formato brasileiro: 1.234,56 -> 1234.56
        if "," in texto and "." in texto:
            if texto.rfind(",") > texto.rfind("."):
                texto = texto.replace(".", "").replace(",", ".")
            else:
                texto = texto.replace(",", "")
        elif "," in texto:
            texto = texto.replace(",", ".")

        try:
            val_float = float(texto)
        except ValueError:
            return None, f"Valor monetário inválido ('{valor}')"

    if val_float <= 0:
        return None, f"Valor unitário deve ser maior que zero ({val_float:.2f})"

    return round(val_float, 2), None


def converter_quantidade(valor: Any) -> Tuple[Optional[int], Optional[str]]:
    """Valida e converte a quantidade vendida para inteiro positivo."""
    if pd.isna(valor) or valor is None:
        return None, "Quantidade ausente"

    try:
        texto = str(valor).strip().replace(",", ".")
        numero = float(texto)
        if not numero.is_integer():
            return None, f"Quantidade não inteira ('{valor}')"
        qtd = int(numero)
    except ValueError:
        return None, f"Quantidade inválida ('{valor}')"

    if qtd <= 0:
        return None, f"Quantidade deve ser maior que zero ({qtd})"

    return qtd, None


def converter_data(valor: Any) -> Tuple[Optional[date], Optional[str]]:
    """Converte diferentes representações de data para objeto datetime.date."""
    if pd.isna(valor) or valor is None:
        return None, "Data da venda ausente"

    if isinstance(valor, datetime):
        return valor.date(), None
    if isinstance(valor, date):
        return valor, None

    texto = str(valor).strip()
    if not texto:
        return None, "Data da venda vazia"

    formatos_aceitos = [
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%Y/%m/%d",
        "%d/%m/%y",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
    ]

    for fmt in formatos_aceitos:
        try:
            dt = datetime.strptime(texto, fmt).date()
            if dt.year < 2000 or dt.year > 2100:
                return None, f"Data fora do intervalo esperado ({dt.strftime('%d/%m/%Y')})"
            return dt, None
        except ValueError:
            continue

    return None, f"Formato de data inválido ('{valor}')"


def validar_registro(linha: Dict[str, Any]) -> Tuple[Dict[str, Any], List[str]]:
    """
    Aplica todas as regras de negócio e higienização em uma linha da planilha.
    Retorna o registro higienizado e uma lista de erros encontrados (vazia se válido).
    """
    erros: List[str] = []

    # 1. ID da Venda
    id_venda = limpar_texto(linha.get("ID_Venda"), title_case=False)
    if id_venda:
        id_venda = id_venda.upper()
    else:
        erros.append("ID da venda ausente")

    # 2. Data da Venda
    data_venda, erro_data = converter_data(linha.get("Data_Venda"))
    if erro_data:
        erros.append(erro_data)

    # 3. Filial
    filial = limpar_texto(linha.get("Filial"), title_case=True)
    if not filial:
        erros.append("Filial ausente")

    # 4. Nome do Cliente
    nome_cliente = limpar_texto(linha.get("Nome_Cliente"), title_case=True)
    if not nome_cliente:
        erros.append("Nome do cliente ausente")

    # 5. CPF do Cliente
    cpf_formatado, erro_cpf = validar_e_formatar_cpf(linha.get("CPF_Cliente"))
    if erro_cpf:
        erros.append(erro_cpf)

    # 6. Produto e Categoria
    produto = limpar_texto(linha.get("Produto"), title_case=True)
    if not produto:
        erros.append("Produto ausente")

    categoria = limpar_texto(linha.get("Categoria"), title_case=True) or "Sem Categoria"

    # 7. Quantidade e Valor Unitário
    quantidade, erro_qtd = converter_quantidade(linha.get("Quantidade"))
    if erro_qtd:
        erros.append(erro_qtd)

    valor_unitario, erro_valor = converter_valor_monetario(linha.get("Valor_Unitario"))
    if erro_valor:
        erros.append(erro_valor)

    # 8. Status do Pagamento
    status_bruto = limpar_texto(linha.get("Status_Pagamento"), title_case=True) or "Pago"
    mapa_status = {
        "Pago": "Aprovado",
        "Aprovado": "Aprovado",
        "Concluido": "Aprovado",
        "Concluído": "Aprovado",
        "Pendente": "Pendente",
        "Aguardando": "Pendente",
        "Cancelado": "Cancelado",
    }
    status_pagamento = mapa_status.get(status_bruto, status_bruto)

    # Cálculo do Valor Total
    valor_total = (
        round(quantidade * valor_unitario, 2)
        if (quantidade is not None and valor_unitario is not None)
        else None
    )

    registro_limpo = {
        "ID_Venda": id_venda,
        "Data_Venda": data_venda,
        "Filial": filial,
        "Nome_Cliente": nome_cliente,
        "CPF_Cliente": cpf_formatado,
        "Produto": produto,
        "Categoria": categoria,
        "Quantidade": quantidade,
        "Valor_Unitario": valor_unitario,
        "Valor_Total": valor_total,
        "Status_Pagamento": status_pagamento,
    }

    return registro_limpo, erros
