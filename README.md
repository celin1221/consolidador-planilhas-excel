# 📊 Pipeline ETL — Consolidador e Higienizador de Planilhas Excel

Aplicação em **Python** desenvolvida para automatizar a consolidação, higienização e auditoria de múltiplas planilhas de vendas (`.xlsx` e `.csv`) enviadas por diferentes filiais ou setores.

---

## 🎯 O Problema de Negócio

Em ambientes corporativos (Financeiro, Operações, Comercial e BI), é comum receber dezenas de planilhas descentralizadas com problemas recorrentes:
- **Colunas fora de padrão** (`Valor_Unitario`, `Preco_Unitario`, `Vlr_Unitario`, `CPF`, `Documento`).
- **Dados sujos ou mal formatados** (`"R$ 4.250,00"` como texto, datas em formatos mistos `DD/MM/YYYY` e `YYYY-MM-DD`, espaços duplicados e nomes em caixa alta/baixa).
- **Erros operacionais** (CPFs com dígito verificador inválido, quantidades zeradas, preços negativos ou campos obrigatórios vazios).
- **Pedidos duplicados** dentro da mesma planilha ou entre arquivos diferentes.

Este projeto resolve esse fluxo ponta a ponta através de um **Pipeline ETL (Extract, Transform, Load)** automatizado, gerando um **Relatório Executivo em Excel (`.xlsx`)** pronto para tomada de decisão.

---

## ✨ Funcionalidades Principais

1. **Ingestão Multi-formato (`.xlsx` e `.csv`)**: Varre automaticamente a pasta `dados_entrada/` e detecta o separador e codificação de arquivos CSV (`utf-8`, `latin-1`).
2. **Padronização Inteligente de Schema**: Mapeia sinônimos de colunas automaticamente para o modelo de dados oficial.
3. **Higienização e Validação Rigorosa**:
   - Algoritmo oficial da Receita Federal para validação dos dígitos verificadores de **CPF** e formatação padrão (`XXX.XXX.XXX-XX`).
   - Conversão de strings monetárias (`R$ 1.250,50`) para números reais (`float`).
   - Normalização de textos (`Title Case` respeitando preposições em português) e datas.
   - Bloqueio automático de **vendas duplicadas** (`ID_Venda`).
4. **Exportação Executiva com `openpyxl` (3 Abas)**:
   - 📈 **Aba 1 — `Resumo Executivo`**: Painel de KPIs (faturamento total, ticket médio, taxa de qualidade dos dados), tabela de desempenho por filial e **gráfico de barras nativo do Excel** gerado via código.
   - ✅ **Aba 2 — `Base Consolidada`**: Apenas os registros 100% válidos e padronizados, com formatação contábil (`R$ #,##0.00`), filtros automáticos (`AutoFilter`), congelamento de painel superior e efeito *zebra*.
   - 🚨 **Aba 3 — `Auditoria de Erros`**: Quarentena com todas as linhas rejeitadas, indicando o **arquivo de origem, número exato da linha no Excel e o diagnóstico detalhado do erro** para correção na ponta.

---

## 🗂️ Estrutura do Projeto

```text
ProjetoPhyton/
├── dados_entrada/               # Planilhas brutas (.xlsx e .csv) das filiais
├── dados_saida/                 # Relatório Excel consolidado gerado pelo pipeline
├── src/
│   ├── __init__.py
│   ├── validador.py             # Regras de negócio, limpeza e validação de CPF/datas/moeda
│   ├── processador.py           # Leitura em lote, mapeamento de colunas e deduplicação
│   └── exportador_excel.py      # Estilização avançada, KPIs e gráficos com openpyxl
├── tests/
│   └── test_validador.py        # Testes unitários automatizados (unittest)
├── gerar_dados_exemplo.py       # Gerador de planilhas "sujas" para demonstração
├── interface_gui.py             # Interface Gráfica Desktop (Tkinter/TTK com KPIs, filtros e gráfico)
├── main.py                      # Ponto de entrada via terminal (CLI)
├── requirements.txt             # Dependências do projeto
└── README.md
```

---

## 🚀 Como Executar

### 1. Instalar as dependências
```bash
pip install -r requirements.txt
```

### 2. Executar com Interface Gráfica (Recomendado)
```bash
python interface_gui.py
```
Abre um painel visual interativo com:
- **6 Cards de KPIs em tempo real** (arquivos lidos, linhas válidas, erros bloqueados, qualidade dos dados e faturamento).
- **Busca/Filtro instantâneo** na base consolidada.
- **Aba de Auditoria de Erros** destacando cada linha rejeitada e o motivo exato.
- **Gráfico de barras por filial** e botões para selecionar pastas/arquivos, exportar o `.xlsx` e abrir direto no Excel.

### 3. Executar via Terminal (CLI / Automação)
```bash
python main.py
```
> **Nota:** Caso a pasta `dados_entrada/` esteja vazia, o script gera automaticamente 3 planilhas de filiais (`.xlsx` e `.csv`) contendo dados válidos e erros propositais para você testar imediatamente!

### 4. Executar os testes automatizados
```bash
python -m unittest discover -s tests
```
