"""Testes automatizados para as funções de higienização e validação de dados."""

from datetime import date
import unittest
from src.validador import (
    converter_data,
    converter_quantidade,
    converter_valor_monetario,
    limpar_texto,
    validar_e_formatar_cpf,
)


class TestValidador(unittest.TestCase):
    def test_cpf_valido_formatado_e_sem_mascara(self):
        cpf, erro = validar_e_formatar_cpf("52998224725")
        self.assertIsNone(erro)
        self.assertEqual(cpf, "529.982.247-25")

    def test_cpf_invalido_digito_verificador(self):
        cpf, erro = validar_e_formatar_cpf("123.456.789-00")
        self.assertIsNone(cpf)
        self.assertIn("dígito verificador", erro)

    def test_cpf_invalido_sequencia_repetida(self):
        cpf, erro = validar_e_formatar_cpf("111.111.111-11")
        self.assertIsNone(cpf)
        self.assertIn("repetidos", erro)

    def test_conversao_moeda_brasileira(self):
        valor, erro = converter_valor_monetario("R$ 4.250,90")
        self.assertIsNone(erro)
        self.assertEqual(valor, 4250.90)

    def test_conversao_moeda_negativa_rejeitada(self):
        valor, erro = converter_valor_monetario("-150,00")
        self.assertIsNone(valor)
        self.assertIsNotNone(erro)

    def test_conversao_datas_multiplos_formatos(self):
        dt1, err1 = converter_data("15/03/2026")
        dt2, err2 = converter_data("2026-03-15")
        self.assertIsNone(err1)
        self.assertIsNone(err2)
        self.assertEqual(dt1, date(2026, 3, 15))
        self.assertEqual(dt2, date(2026, 3, 15))

    def test_limpeza_texto_com_preposicoes(self):
        nome = limpar_texto("  JOÃO   DA   SILVA  ")
        self.assertEqual(nome, "João da Silva")

    def test_quantidade_invalida(self):
        qtd, erro = converter_quantidade("0")
        self.assertIsNone(qtd)
        self.assertIsNotNone(erro)


if __name__ == "__main__":
    unittest.main()
