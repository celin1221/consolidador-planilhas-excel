"""
Interface Gráfica Desktop (GUI) para o Pipeline ETL de Consolidação e Higienização de Planilhas.
Construída com Tkinter / TTK (nativo do Python, sem necessidade de servidor externo).
"""

import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional
import pandas as pd

from gerar_dados_exemplo import gerar_planilhas_exemplo
from src.exportador_excel import exportar_relatorio_excel
from src.processador import processar_arquivos, processar_diretorio


class ConsolidadorExcelApp(tk.Tk):
    """Janela principal da aplicação de consolidação e auditoria de planilhas."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Consolidador e Higienizador de Planilhas Excel — Pipeline ETL")
        self.geometry("1240x760")
        self.minsize(1024, 640)
        self.configure(bg="#F3F6F9")

        # Estado da aplicação
        self.pasta_atual: Path = Path("dados_entrada")
        self.arquivos_selecionados: List[Path] = []
        self.ultimo_relatorio_salvo: Path = Path("dados_saida/Relatorio_Consolidado_Vendas.xlsx")
        self.df_validos: pd.DataFrame = pd.DataFrame()
        self.df_auditoria: pd.DataFrame = pd.DataFrame()
        self.metricas: dict = {}

        self._configurar_estilos()
        self._construir_interface()

        # Carrega automaticamente os dados ao iniciar
        self.after(150, self.acao_processar_padrao)

    def _configurar_estilos(self) -> None:
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        style.configure(
            "Treeview",
            background="#FFFFFF",
            foreground="#1E293B",
            rowheight=28,
            fieldbackground="#FFFFFF",
            font=("Segoe UI", 9),
        )
        style.configure(
            "Treeview.Heading",
            background="#1F4E78",
            foreground="#FFFFFF",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
        )
        style.map(
            "Treeview.Heading",
            background=[("active", "#2F75B5")],
        )
        style.map(
            "Treeview",
            background=[("selected", "#D0E4F7")],
            foreground=[("selected", "#0F172A")],
        )
        style.configure(
            "TNotebook",
            background="#F3F6F9",
            borderwidth=0,
        )
        style.configure(
            "TNotebook.Tab",
            font=("Segoe UI", 10, "bold"),
            padding=(16, 8),
            background="#E2E8F0",
            foreground="#334155",
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#1F4E78")],
            foreground=[("selected", "#FFFFFF")],
        )

    def _construir_interface(self) -> None:
        # 1. Cabeçalho Superior
        header = tk.Frame(self, bg="#1F4E78", height=74)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        header_left = tk.Frame(header, bg="#1F4E78")
        header_left.pack(side="left", padx=20, pady=10)

        tk.Label(
            header_left,
            text="📊 Consolidador e Higienizador de Planilhas Excel",
            font=("Segoe UI", 16, "bold"),
            fg="#FFFFFF",
            bg="#1F4E78",
        ).pack(anchor="w")

        tk.Label(
            header_left,
            text="Pipeline ETL automatizado: padronização de colunas, validação de CPF/datas/valores, deduplicação e relatório executivo.",
            font=("Segoe UI", 9),
            fg="#CBD5E1",
            bg="#1F4E78",
        ).pack(anchor="w")

        # 2. Barra de Ferramentas (Botões de Ação)
        toolbar = tk.Frame(self, bg="#FFFFFF", bd=1, relief="solid")
        toolbar.pack(fill="x", padx=18, pady=(14, 10))

        botoes = [
            ("📂 Selecionar Pasta", "#2F75B5", self.acao_selecionar_pasta),
            ("📄 Selecionar Arquivos", "#2F75B5", self.acao_selecionar_arquivos),
            ("🧪 Recriar Planilhas de Exemplo", "#475569", self.acao_gerar_exemplos),
            ("⚡ Processar e Atualizar", "#0D9488", self.acao_reprocessar),
            ("💾 Exportar Excel (.xlsx)", "#15803D", self.acao_exportar_como),
            ("📗 Abrir Relatório no Excel", "#B45309", self.acao_abrir_excel),
        ]

        for texto, cor, comando in botoes:
            btn = tk.Button(
                toolbar,
                text=texto,
                command=comando,
                bg=cor,
                fg="#FFFFFF",
                activebackground="#1E293B",
                activeforeground="#FFFFFF",
                font=("Segoe UI", 9, "bold"),
                relief="flat",
                padx=12,
                pady=7,
                cursor="hand2",
            )
            btn.pack(side="left", padx=6, pady=8)

        # 3. Painel de Cards de KPIs
        self.frame_kpis = tk.Frame(self, bg="#F3F6F9")
        self.frame_kpis.pack(fill="x", padx=14, pady=(0, 10))

        self.cards_kpi = {}
        definicao_cards = [
            ("arquivos", "Arquivos Lidos", "0", "#1F4E78"),
            ("linhas", "Linhas Brutas", "0", "#334155"),
            ("validos", "Registros Válidos", "0", "#15803D"),
            ("rejeitados", "Erros / Quarentena", "0", "#B91C1C"),
            ("aproveitamento", "Qualidade dos Dados", "0.0%", "#0369A1"),
            ("faturamento", "Faturamento Consolidado", "R$ 0,00", "#047857"),
        ]

        for i, (chave, titulo, valor_inicial, cor_borda) in enumerate(definicao_cards):
            self.frame_kpis.columnconfigure(i, weight=1)
            card = tk.Frame(
                self.frame_kpis,
                bg="#FFFFFF",
                highlightbackground=cor_borda,
                highlightthickness=2,
                padx=12,
                pady=8,
            )
            card.grid(row=0, column=i, sticky="nsew", padx=4)

            tk.Label(
                card,
                text=titulo.upper(),
                font=("Segoe UI", 8, "bold"),
                fg="#64748B",
                bg="#FFFFFF",
            ).pack(anchor="w")

            lbl_valor = tk.Label(
                card,
                text=valor_inicial,
                font=("Segoe UI", 15, "bold"),
                fg=cor_borda,
                bg="#FFFFFF",
            )
            lbl_valor.pack(anchor="w", pady=(2, 0))
            self.cards_kpi[chave] = lbl_valor

        # 4. Área Principal com Abas (Notebook)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=18, pady=(0, 8))

        self._construir_aba_validos()
        self._construir_aba_auditoria()
        self._construir_aba_resumo_filiais()

        # 5. Barra de Status Inferior
        status_frame = tk.Frame(self, bg="#E2E8F0", height=28)
        status_frame.pack(fill="x", side="bottom")
        self.var_status = tk.StringVar(value="Pronto.")
        tk.Label(
            status_frame,
            textvariable=self.var_status,
            font=("Segoe UI", 9),
            bg="#E2E8F0",
            fg="#334155",
             padx=12,
        ).pack(side="left", pady=3)

    def _construir_aba_validos(self) -> None:
        aba = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(aba, text="  ✅ Base Consolidada (Dados Limpos)  ")

        # Barra de busca rápida
        filtro_frame = tk.Frame(aba, bg="#F8FAFC", pady=6, padx=10)
        filtro_frame.pack(fill="x")

        tk.Label(
            filtro_frame,
            text="🔍 Filtrar registros (Cliente, Filial, Produto, CPF ou ID):",
            font=("Segoe UI", 9, "bold"),
            bg="#F8FAFC",
            fg="#334155",
        ).pack(side="left")

        self.var_busca = tk.StringVar()
        self.var_busca.trace_add("write", lambda *_: self._atualizar_tabela_validos())
        entry_busca = ttk.Entry(filtro_frame, textvariable=self.var_busca, width=42)
        entry_busca.pack(side="left", padx=8)

        tk.Button(
            filtro_frame,
            text="Limpar Filtro",
            command=lambda: self.var_busca.set(""),
            bg="#E2E8F0",
            fg="#1E293B",
            relief="flat",
            font=("Segoe UI", 8),
            padx=8,
        ).pack(side="left")

        colunas = (
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
        )
        headers = {
            "ID_Venda": ("ID Venda", 90, "center"),
            "Data_Venda": ("Data", 90, "center"),
            "Filial": ("Filial", 150, "w"),
            "Nome_Cliente": ("Cliente", 170, "w"),
            "CPF_Cliente": ("CPF", 115, "center"),
            "Produto": ("Produto", 190, "w"),
            "Categoria": ("Categoria", 105, "w"),
            "Quantidade": ("Qtd", 55, "e"),
            "Valor_Unitario": ("Vlr. Unitário", 105, "e"),
            "Valor_Total": ("Vlr. Total", 110, "e"),
            "Status_Pagamento": ("Status", 85, "center"),
            "Arquivo_Origem": ("Arquivo Origem", 170, "w"),
        }

        container = tk.Frame(aba, bg="#FFFFFF")
        container.pack(fill="both", expand=True, padx=6, pady=6)

        scroll_y = ttk.Scrollbar(container, orient="vertical")
        scroll_x = ttk.Scrollbar(container, orient="horizontal")

        self.tree_validos = ttk.Treeview(
            container,
            columns=colunas,
            show="headings",
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set,
        )
        scroll_y.config(command=self.tree_validos.yview)
        scroll_x.config(command=self.tree_validos.xview)

        for col, (titulo, largura, alinhamento) in headers.items():
            self.tree_validos.heading(col, text=titulo)
            self.tree_validos.column(col, width=largura, anchor=alinhamento)

        self.tree_validos.tag_configure("par", background="#FFFFFF")
        self.tree_validos.tag_configure("impar", background="#F2F6FA")

        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.tree_validos.pack(fill="both", expand=True)

    def _construir_aba_auditoria(self) -> None:
        aba = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(aba, text="  🚨 Auditoria de Erros (Quarentena)  ")

        banner_info = tk.Frame(aba, bg="#FEF2F2", pady=6, padx=10)
        banner_info.pack(fill="x")
        tk.Label(
            banner_info,
            text="⚠️ As linhas abaixo foram bloqueadas automaticamente pelo validador para não contaminar os indicadores executivos.",
            font=("Segoe UI", 9, "bold"),
            bg="#FEF2F2",
            fg="#991B1B",
        ).pack(anchor="w")

        colunas = (
            "Arquivo_Origem",
            "Linha_Origem",
            "ID_Venda",
            "Data_Original",
            "Filial",
            "Cliente",
            "CPF_Original",
            "Qtd_Original",
            "Valor_Original",
            "Motivo_Rejeicao",
        )
        headers = {
            "Arquivo_Origem": ("Arquivo Origem", 175, "w"),
            "Linha_Origem": ("Linha Excel", 80, "center"),
            "ID_Venda": ("ID Venda", 85, "center"),
            "Data_Original": ("Data Informada", 105, "center"),
            "Filial": ("Filial", 140, "w"),
            "Cliente": ("Cliente", 145, "w"),
            "CPF_Original": ("CPF Informado", 115, "center"),
            "Qtd_Original": ("Qtd", 55, "center"),
            "Valor_Original": ("Valor Informado", 105, "e"),
            "Motivo_Rejeicao": ("Diagnóstico / Motivo da Rejeição", 360, "w"),
        }

        container = tk.Frame(aba, bg="#FFFFFF")
        container.pack(fill="both", expand=True, padx=6, pady=6)

        scroll_y = ttk.Scrollbar(container, orient="vertical")
        scroll_x = ttk.Scrollbar(container, orient="horizontal")

        self.tree_auditoria = ttk.Treeview(
            container,
            columns=colunas,
            show="headings",
            yscrollcommand=scroll_y.set,
            xscrollcommand=scroll_x.set,
        )
        scroll_y.config(command=self.tree_auditoria.yview)
        scroll_x.config(command=self.tree_auditoria.xview)

        for col, (titulo, largura, alinhamento) in headers.items():
            self.tree_auditoria.heading(col, text=titulo)
            self.tree_auditoria.column(col, width=largura, anchor=alinhamento)

        self.tree_auditoria.tag_configure("erro", background="#FFF5F5", foreground="#7F1D1D")

        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        self.tree_auditoria.pack(fill="both", expand=True)

    def _construir_aba_resumo_filiais(self) -> None:
        aba = tk.Frame(self.notebook, bg="#FFFFFF")
        self.notebook.add(aba, text="  🏢 Desempenho por Filial & Gráfico  ")

        painel_superior = tk.Frame(aba, bg="#FFFFFF")
        painel_superior.pack(fill="x", padx=12, pady=10)

        tk.Label(
            painel_superior,
            text="Faturamento Consolidado e Participação por Filial",
            font=("Segoe UI", 11, "bold"),
            bg="#FFFFFF",
            fg="#1F4E78",
        ).pack(anchor="w", pady=(0, 6))

        colunas = ("Filial", "Pedidos", "Itens", "Faturamento", "Ticket_Medio", "Participacao")
        self.tree_filiais = ttk.Treeview(
            painel_superior,
            columns=colunas,
            show="headings",
            height=5,
        )
        definicoes = [
            ("Filial", "Filial", 220, "w"),
            ("Pedidos", "Qtd. Pedidos", 110, "center"),
            ("Itens", "Itens Vendidos", 120, "center"),
            ("Faturamento", "Faturamento Total", 160, "e"),
            ("Ticket_Medio", "Ticket Médio", 150, "e"),
            ("Participacao", "% Participação", 130, "e"),
        ]
        for col, texto, larg, anc in definicoes:
            self.tree_filiais.heading(col, text=texto)
            self.tree_filiais.column(col, width=larg, anchor=anc)
        self.tree_filiais.pack(fill="x")

        # Canvas para gráfico visual de barras horizontais
        tk.Label(
            aba,
            text="Visualização Gráfica — Faturamento por Filial (R$)",
            font=("Segoe UI", 11, "bold"),
            bg="#FFFFFF",
            fg="#1F4E78",
        ).pack(anchor="w", padx=12, pady=(12, 4))

        self.canvas_grafico = tk.Canvas(
            aba,
            bg="#F8FAFC",
            height=210,
            highlightthickness=1,
            highlightbackground="#CBD5E1",
        )
        self.canvas_grafico.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        self.canvas_grafico.bind("<Configure>", lambda _: self._desenhar_grafico_filiais())

    def _formatar_moeda_br(self, valor: Optional[float]) -> str:
        if valor is None or pd.isna(valor):
            return "R$ 0,00"
        texto = f"{float(valor):,.2f}"
        texto = texto.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {texto}"

    def _atualizar_interface_completa(self) -> None:
        m = self.metricas
        self.cards_kpi["arquivos"].config(text=str(m.get("arquivos_processados", 0)))
        self.cards_kpi["linhas"].config(text=str(m.get("total_linhas_lidas", 0)))
        self.cards_kpi["validos"].config(text=str(m.get("registros_validos", 0)))
        self.cards_kpi["rejeitados"].config(text=str(m.get("registros_rejeitados", 0)))
        self.cards_kpi["aproveitamento"].config(text=f"{m.get('taxa_aproveitamento', 0.0):.1f}%")
        self.cards_kpi["faturamento"].config(
            text=self._formatar_moeda_br(m.get("faturamento_total", 0.0))
        )

        self._atualizar_tabela_validos()
        self._atualizar_tabela_auditoria()
        self._atualizar_resumo_filiais()

    def _atualizar_tabela_validos(self) -> None:
        for item in self.tree_validos.get_children():
            self.tree_validos.delete(item)

        if self.df_validos.empty:
            return

        termo = self.var_busca.get().strip().lower()
        for idx, (_, row) in enumerate(self.df_validos.iterrows()):
            data_fmt = (
                row["Data_Venda"].strftime("%d/%m/%Y")
                if hasattr(row["Data_Venda"], "strftime")
                else str(row["Data_Venda"])
            )
            valores = (
                row["ID_Venda"],
                data_fmt,
                row["Filial"],
                row["Nome_Cliente"],
                row["CPF_Cliente"],
                row["Produto"],
                row["Categoria"],
                row["Quantidade"],
                self._formatar_moeda_br(row["Valor_Unitario"]),
                self._formatar_moeda_br(row["Valor_Total"]),
                row["Status_Pagamento"],
                row["Arquivo_Origem"],
            )
            if termo and not any(termo in str(v).lower() for v in valores):
                continue

            tag = "par" if idx % 2 == 0 else "impar"
            self.tree_validos.insert("", "end", values=valores, tags=(tag,))

    def _atualizar_tabela_auditoria(self) -> None:
        for item in self.tree_auditoria.get_children():
            self.tree_auditoria.delete(item)

        if self.df_auditoria.empty:
            return

        for _, row in self.df_auditoria.iterrows():
            valores = (
                row.get("Arquivo_Origem", ""),
                row.get("Linha_Origem", ""),
                row.get("ID_Venda", ""),
                row.get("Data_Original", ""),
                row.get("Filial", ""),
                row.get("Cliente", ""),
                row.get("CPF_Original", ""),
                row.get("Qtd_Original", ""),
                row.get("Valor_Original", ""),
                row.get("Motivo_Rejeicao", ""),
            )
            self.tree_auditoria.insert("", "end", values=valores, tags=("erro",))

    def _atualizar_resumo_filiais(self) -> None:
        for item in self.tree_filiais.get_children():
            self.tree_filiais.delete(item)

        if self.df_validos.empty:
            self._desenhar_grafico_filiais()
            return

        resumo = (
            self.df_validos.groupby("Filial", as_index=False)
            .agg(
                Pedidos=("ID_Venda", "count"),
                Itens=("Quantidade", "sum"),
                Faturamento=("Valor_Total", "sum"),
                Ticket_Medio=("Valor_Total", "mean"),
            )
            .sort_values(by="Faturamento", ascending=False)
        )
        total_geral = float(resumo["Faturamento"].sum()) or 1.0

        for _, row in resumo.iterrows():
            part = (float(row["Faturamento"]) / total_geral) * 100.0
            self.tree_filiais.insert(
                "",
                "end",
                values=(
                    row["Filial"],
                    int(row["Pedidos"]),
                    int(row["Itens"]),
                    self._formatar_moeda_br(row["Faturamento"]),
                    self._formatar_moeda_br(row["Ticket_Medio"]),
                    f"{part:.1f}%",
                ),
            )

        self._desenhar_grafico_filiais()

    def _desenhar_grafico_filiais(self) -> None:
        self.canvas_grafico.delete("all")
        if self.df_validos.empty:
            return

        resumo = (
            self.df_validos.groupby("Filial", as_index=False)
            .agg(Faturamento=("Valor_Total", "sum"))
            .sort_values(by="Faturamento", ascending=False)
        )
        maior_valor = float(resumo["Faturamento"].max()) or 1.0
        largura_canvas = max(self.canvas_grafico.winfo_width(), 650)

        margem_esq = 210
        margem_dir = 160
        largura_util = max(200, largura_canvas - margem_esq - margem_dir)
        cores = ["#1F4E78", "#2F75B5", "#0D9488", "#047857", "#6366F1"]

        y = 25
        altura_barra = 32
        espacamento = 18

        for idx, (_, row) in enumerate(resumo.iterrows()):
            filial = str(row["Filial"])
            fat = float(row["Faturamento"])
            comp = int((fat / maior_valor) * largura_util)
            cor = cores[idx % len(cores)]

            # Nome da filial
            self.canvas_grafico.create_text(
                margem_esq - 12,
                y + altura_barra // 2,
                text=filial,
                anchor="e",
                font=("Segoe UI", 10, "bold"),
                fill="#1E293B",
            )
            # Barra
            self.canvas_grafico.create_rectangle(
                margem_esq,
                y,
                margem_esq + comp,
                y + altura_barra,
                fill=cor,
                outline="",
            )
            # Rótulo de valor
            self.canvas_grafico.create_text(
                margem_esq + comp + 10,
                y + altura_barra // 2,
                text=self._formatar_moeda_br(fat),
                anchor="w",
                font=("Segoe UI", 10, "bold"),
                fill=cor,
            )
            y += altura_barra + espacamento

    # =========================================================================
    # Ações dos Botões
    # =========================================================================
    def acao_processar_padrao(self) -> None:
        if not self.pasta_atual.exists() or not any(self.pasta_atual.iterdir()):
            gerar_planilhas_exemplo(self.pasta_atual)
        self.acao_reprocessar()

    def acao_reprocessar(self) -> None:
        try:
            if self.arquivos_selecionados:
                self.df_validos, self.df_auditoria, self.metricas = processar_arquivos(
                    self.arquivos_selecionados
                )
                origem_str = f"{len(self.arquivos_selecionados)} arquivo(s) selecionado(s)"
            else:
                self.df_validos, self.df_auditoria, self.metricas = processar_diretorio(
                    self.pasta_atual
                )
                origem_str = f"pasta '{self.pasta_atual}'"

            exportar_relatorio_excel(
                caminho_saida=self.ultimo_relatorio_salvo,
                df_validos=self.df_validos,
                df_auditoria=self.df_auditoria,
                metricas=self.metricas,
            )
            self._atualizar_interface_completa()
            self.var_status.set(
                f"✅ Processado a partir de {origem_str} | Relatório atualizado em: {self.ultimo_relatorio_salvo}"
            )
        except Exception as exc:
            messagebox.showerror("Erro no Processamento", f"Ocorreu um erro:\n{exc}")

    def acao_selecionar_pasta(self) -> None:
        pasta = filedialog.askdirectory(
            title="Selecione a pasta com as planilhas (.xlsx e .csv)",
            initialdir=str(self.pasta_atual.resolve()),
        )
        if pasta:
            self.pasta_atual = Path(pasta)
            self.arquivos_selecionados = []
            self.acao_reprocessar()

    def acao_selecionar_arquivos(self) -> None:
        arquivos = filedialog.askopenfilenames(
            title="Selecione as planilhas (.xlsx ou .csv)",
            initialdir=str(self.pasta_atual.resolve()),
            filetypes=[("Planilhas Excel e CSV", "*.xlsx *.csv")],
        )
        if arquivos:
            self.arquivos_selecionados = [Path(a) for a in arquivos]
            self.acao_reprocessar()

    def acao_gerar_exemplos(self) -> None:
        self.pasta_atual = Path("dados_entrada")
        self.arquivos_selecionados = []
        gerar_planilhas_exemplo(self.pasta_atual)
        self.acao_reprocessar()
        messagebox.showinfo(
            "Planilhas de Exemplo Geradas",
            "As 3 planilhas de demonstração (São Paulo, Rio de Janeiro e Curitiba) "
            "foram recriadas na pasta 'dados_entrada/' e consolidadas com sucesso!",
        )

    def acao_exportar_como(self) -> None:
        if self.df_validos.empty and self.df_auditoria.empty:
            messagebox.showwarning("Aviso", "Não há dados carregados para exportar.")
            return

        caminho = filedialog.asksaveasfilename(
            title="Salvar Relatório Consolidado Excel",
            initialdir=str(Path("dados_saida").resolve()),
            initialfile="Relatorio_Consolidado_Vendas.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Planilha Excel", "*.xlsx")],
        )
        if caminho:
            self.ultimo_relatorio_salvo = Path(caminho)
            exportar_relatorio_excel(
                caminho_saida=self.ultimo_relatorio_salvo,
                df_validos=self.df_validos,
                df_auditoria=self.df_auditoria,
                metricas=self.metricas,
            )
            self.var_status.set(f"💾 Relatório exportado com sucesso em: {self.ultimo_relatorio_salvo}")
            messagebox.showinfo(
                "Exportação Concluída",
                f"Relatório Excel salvo com sucesso em:\n{self.ultimo_relatorio_salvo}",
            )

    def acao_abrir_excel(self) -> None:
        if not self.ultimo_relatorio_salvo.exists():
            self.acao_reprocessar()
        try:
            os.startfile(str(self.ultimo_relatorio_salvo.resolve()))
        except Exception as exc:
            messagebox.showerror(
                "Não foi possível abrir o arquivo",
                f"Arquivo salvo em:\n{self.ultimo_relatorio_salvo.resolve()}\n\nDetalhes: {exc}",
            )


if __name__ == "__main__":
    app = ConsolidadorExcelApp()
    app.mainloop()
