import sys
import os
import json
import traceback
import time
import tempfile
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import threading
import subprocess
import datetime
import config

# ==============================================================================
# 1. INTERCEPTADOR DE PROCESSOS
# ==============================================================================
if len(sys.argv) > 2 and sys.argv[1] == "--run-code":
    codigo = sys.argv[2]
    log_path = sys.argv[3] if len(sys.argv) > 3 else None

    if log_path:
        sys.stdout = open(log_path, "w", encoding="utf-8", buffering=1)
        sys.stderr = sys.stdout

    try:
        exec(codigo)
    except Exception:
        traceback.print_exc()
        sys.exit(1)

    sys.exit(0)


# ==============================================================================
# 2. FORÇAR PYINSTALLER A EMPACOTAR OS ROBÔS
# ==============================================================================
try:
    import robos.robo_rateio_malote
    import robos.robo_faturamento
    import robos.robo_juridico
    import robos.robo_incluir_encomendas
    import robos.robo_relatorio_correios
    import robos.produtividade
    import robos.robo_fechar_chamados
    import robos.robo_uber_relatorios
    import robos.robo_zmm180
    import robos.malote_web_scraper
    import robos.criar_rascunhos_uber
    import robos.robo_macro_contratos
    import robos.robo_cobrar_boleto
    import robos.robo_rateio_AGF
except ImportError:
    pass


# ==============================================================================
# 3. TEMA
# ==============================================================================
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("green")

CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


class PrintRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, text):
        try:
            self.text_widget.after(0, self._inserir_texto, text)
        except Exception:
            pass

    def _inserir_texto(self, text):
        try:
            self.text_widget.configure(state="normal")
            self.text_widget.insert(tk.END, text)
            self.text_widget.see(tk.END)
            self.text_widget.configure(state="disabled")
        except Exception:
            pass

    def flush(self):
        pass


class CentralAutomacaoMRV:
    def __init__(self, root):
        self.root = root

        self.processo_ativo = None
        self.foi_cancelado = False
        self.todos_botoes = []
        self.tela_atual = None
        self.logs_visiveis = True
        self.robos_filtrados = []
        self._resize_after_id = None

        self.PASTA_BASE = getattr(config, "PASTA_PROJETO", os.path.dirname(os.path.abspath(__file__)))
        self.ARQUIVO_HISTORICO = os.path.join(self.PASTA_BASE, "historico_execucoes.json")
        self.ARQUIVO_ACOES_RAPIDAS = os.path.join(self.PASTA_BASE, "acoes_rapidas.json")


        # Cores
        self.COR_BG = "#202020"
        self.COR_CARD = "#2A2A2A"
        self.COR_CARD_2 = "#303030"
        self.COR_TEXTO = "#F2F2F2"
        self.COR_TEXTO_FRACO = "#B8B8B8"
        self.COR_MRV = "#008542"
        self.COR_MRV_HOVER = "#006331"
        self.COR_CANCELAR = "#E74C3C"
        self.COR_CANCELAR_HOVER = "#C0392B"
        self.COR_SIDEBAR = "#171717"
        self.COR_SIDEBAR_HOVER = "#2A2A2A"
        self.COR_SIDEBAR_ACTIVE = "#008542"
        self.COR_AZUL = "#1F6FEB"
        self.COR_ROXO = "#7B61FF"
        self.COR_LARANJA = "#E67E22"
        self.COR_AMARELO = "#F1C40F"
        self.COR_VINHO = "#A93255"
        self.COR_CINZA = "#5D6D7E"

        self.cores_categoria = {
            "Correios & Faturamento": self.COR_AZUL,
            "Podio & Mensageria": self.COR_VINHO,
            "Agilis & Chamados": self.COR_MRV,
            "Outros (Uber / SAP / Contratos)": self.COR_LARANJA,
        }

        self.root.title("Hub Central de Automações - MRV")
        self.root.geometry("1250x780")
        self.root.minsize(1100, 700)

        try:
            self.root.after(0, lambda: self.root.state("zoomed"))
        except Exception:
            pass

        self.root.configure(fg_color=self.COR_BG)

        self._montar_lista_robos()

        # Layout principal
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

        self.sidebar_frame = ctk.CTkFrame(
            self.root,
            width=230,
            corner_radius=0,
            fg_color=self.COR_SIDEBAR
        )
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(8, weight=1)

        self.main_frame = ctk.CTkFrame(self.root, corner_radius=0, fg_color=self.COR_BG)
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        self.frames = {}

        self._construir_sidebar()
        self._construir_tela_inicio()
        self._construir_tela_robos()
        self._construir_tela_config()
        self._construir_tela_ajuda()

        self.root.bind("<Control-f>", self._atalho_busca)
        self.root.bind("<Control-F>", self._atalho_busca)
        
        # Vincula o evento de redimensionamento da janela
        self.root.bind("<Configure>", self._on_resize)

        self.selecionar_tela("inicio")
        self.root.after(900, self._verificar_credenciais_iniciais)

    # ==========================================================================
    # DADOS DOS ROBÔS
    # ==========================================================================
    def _montar_lista_robos(self):
        self.robos = [
            {
                "nome": "Rateio de Malote",
                "titulo": "Rateio de Malote (Centros de Custo)",
                "categoria": "Correios & Faturamento",
                "icone": "📦",
                "cor": self.COR_AZUL,
                "tempo": "2 a 10 min",
                "Prioridade": "Média",
                "requisitos": ["Excel", "Arquivos locais", "Correios"],
                "descricao": "Gera a planilha de rateio de malote por centro de custo.",
                "comando": "import robos.robo_rateio_malote as rrm; rrm.executar_rateio_malote()",
                "pasta": os.path.join(config.PASTA_ARQUIVOS, "rateio_malote"),
                "tipo": "pasta",
            },
            {
                "nome": "Rateio AGF",
                "requer_api_agilis": True,
                "titulo": "Rateio AGF",
                "categoria": "Correios & Faturamento",
                "icone": "📮",
                "cor": self.COR_AZUL,
                "tempo": "3 a 10 min",
                "Prioridade": "Média",
                "requisitos": ["Excel", "Arquivos locais","API Agilis"],
                "descricao": "Processa os arquivos de rateio AGF.",
                "comando": "import robos.robo_rateio_AGF as rra; rra.executar_rateio_AGF()",
                "pasta": os.path.join(config.PASTA_ARQUIVOS, "rateio_AGF"),
                "tipo": "pasta",
            },
            {
                "nome": "Faturamento 1",
                "titulo": "Faturamento 1: Gerar Rascunhos",
                "categoria": "Correios & Faturamento",
                "icone": "✉️",
                "cor": self.COR_ROXO,
                "tempo": "1 a 2 min",
                "Prioridade": "Baixo",
                "requisitos": ["Outlook"],
                "descricao": "Cria rascunhos de e-mail no Outlook.",
                "comando": "import robos.robo_faturamento as rf; rf.criar_rascunhos_correios()",
                "tipo": "direto",
            },
            {
                "nome": "Faturamento Completo",
                "titulo": "Faturamento 2: Processo Completo",
                "categoria": "Correios & Faturamento",
                "icone": "💰",
                "cor": self.COR_ROXO,
                "tempo": "2 a 5 min",
                "Prioridade": "Alto",
                "requisitos": ["Outlook", "MRV Pag", "Rede"],
                "descricao": "Executa o fluxo completo de e-mail até MRV Pag.",
                "comando": "import robos.robo_faturamento as rf; rf.executar_faturamento_completo()",
                "tipo": "direto",
            },
            {
                "nome": "Follow-up Boletos",
                "titulo": "Cobrança de boletos",
                "categoria": "Correios & Faturamento",
                "icone": "📨",
                "cor": self.COR_ROXO,
                "tempo": "2 a 5 min",
                "Prioridade": "Média",
                "requisitos": ["Outlook", "E-mails acessíveis"],
                "descricao": "Cria follow-up automático para boletos sem retorno.",
                "comando": "import robos.robo_cobrar_boleto as rcb; rcb.executar_cobranca_boletos()",
                "tipo": "direto",
            },
            {
                "nome": "Relatório Jurídico Montreal",
                "titulo": "Relatório Jurídico Montreal",
                "categoria": "Podio & Mensageria",
                "icone": "⚖️",
                "cor": self.COR_VINHO,
                "tempo": "1 a 3 min",
                "Prioridade": "Média",
                "requisitos": ["Podio", "Excel", "MFA"],
                "descricao": "Baixa ou formata relatório jurídico Montreal.",
                "handler": self._chamar_robo_juridico,
                "tipo": "especial",
            },
            {
                "nome": "Incluir Correspondências",
                "titulo": "Incluir Correspondências Rápidas",
                "categoria": "Podio & Mensageria",
                "icone": "📬",
                "cor": self.COR_VINHO,
                "tempo": "30 seg a 2 min",
                "Prioridade": "Média",
                "requisitos": ["Planilha encomendas", "Podio"],
                "descricao": "Inclui correspondências a partir da planilha encomendas.xlsx.",
                "handler": self._chamar_robo_incluir_encomendas,
                "tipo": "especial",
            },
            {
                "nome": "Relatório Correios",
                "requer_api_agilis": True,
                "titulo": "Gerar relatório de envio para Correios",
                "categoria": "Agilis & Chamados",
                "icone": "📊",
                "cor": self.COR_MRV,
                "tempo": "1 a 3 min",
                "Prioridade": "Média",
                "requisitos": ["Agilis", "Excel", "API Agilis"],
                "descricao": "Gera relatório de envios para os Correios.",
                "comando": "import robos.robo_relatorio_correios as rc; rc.executar_relatorio_completo()",
                "tipo": "direto",
            },
            {
                "nome": "Produtividade",
                "titulo": "Gerar Produtividade (Podio/Agilis/SAP)",
                "categoria": "Agilis & Chamados",
                "icone": "📈",
                "cor": self.COR_MRV,
                "tempo": "5 a 10 min",
                "Prioridade": "Alto",
                "requisitos": ["Podio", "Agilis", "SAP", "Bússola", "MFA"],
                "descricao": "Gera produtividade consolidando Podio, Agilis e SAP.",
                "handler": self._chamar_robo_produtividade,
                "tipo": "especial",
            },
            {
                "nome": "Fechar Chamados",
                "titulo": "Fechar Chamados a Vencer",
                "categoria": "Agilis & Chamados",
                "icone": "✅",
                "cor": self.COR_MRV,
                "tempo": "Variável",
                "Prioridade": "Alto",
                "requisitos": ["Agilis", "MFA"],
                "descricao": "Fecha chamados a vencer ou monitora chamados durante o dia.",
                "handler": self._chamar_robo_fechar_chamados,
                "tipo": "especial",
            },
            {
                "nome": "Macro Contratos",
                "titulo": "Atualizar Macro de Contratos",
                "categoria": "Outros (Uber / SAP / Contratos)",
                "icone": "📑",
                "cor": self.COR_LARANJA,
                "tempo": "2 a 4 min",
                "Prioridade": "Média",
                "requisitos": ["Excel fechado", "Rede"],
                "descricao": "Atualiza macros de contratos usando arquivos de rede.",
                "comando": "import robos.robo_macro_contratos as rmc; rmc.executar_macro_contratos()",
                "tipo": "direto",
            },
            {
                "nome": "Uber 1",
                "titulo": "Uber 1: Atualizar Responsáveis (SAP)",
                "categoria": "Outros (Uber / SAP / Contratos)",
                "icone": "🚗",
                "cor": self.COR_LARANJA,
                "tempo": "1 a 4 min",
                "Prioridade": "Média",
                "requisitos": ["SAP", "Excel", "Arquivos Uber"],
                "descricao": "Atualiza responsáveis com base na exportação do SAP.",
                "comando": "import robos.robo_uber_relatorios as ru; ru.etapa_1_atualizar_responsaveis()",
                "pasta": os.path.join(config.PASTA_ARQUIVOS, "uber"),
                "tipo": "pasta",
            },
            {
                "nome": "Uber 2",
                "titulo": "Uber 2: Gerar Relatórios e Pastas",
                "categoria": "Outros (Uber / SAP / Contratos)",
                "icone": "📁",
                "cor": self.COR_LARANJA,
                "tempo": "2 a 4 min",
                "Prioridade": "Média",
                "requisitos": ["Excel", "Arquivos Uber"],
                "descricao": "Gera relatórios, planilhas e pastas do fluxo Uber.",
                "comando": "import robos.robo_uber_relatorios as ru; ru.etapa_2_gerar_relatorios()",
                "pasta": os.path.join(config.PASTA_ARQUIVOS, "uber"),
                "tipo": "pasta",
            },
            {
                "nome": "Uber 3",
                "titulo": "Uber 3: Criar Rascunhos de E-mail",
                "categoria": "Outros (Uber / SAP / Contratos)",
                "icone": "📧",
                "cor": self.COR_LARANJA,
                "tempo": "1 a 4 min",
                "Prioridade": "Média",
                "requisitos": ["Outlook", "Arquivos Uber"],
                "descricao": "Cria rascunhos de e-mail a partir dos arquivos gerados.",
                "comando": "import robos.criar_rascunhos_uber as rr; rr.criar_rascunhos()",
                "pasta": os.path.join(config.PASTA_ARQUIVOS, "uber"),
                "tipo": "pasta",
            },
            {
                "nome": "ZMM180",
                "titulo": "Faturamento Transação ZMM180",
                "categoria": "Outros (Uber / SAP / Contratos)",
                "icone": "🧾",
                "cor": self.COR_AMARELO,
                "tempo": "20 a 40 min",
                "Prioridade": "Alto",
                "requisitos": ["SAP", "Edge", "OCR", "Não mexer no mouse"],
                "descricao": "Automação SAP/ZMM180 com PyAutoGUI e OCR.",
                "handler": self._chamar_robo_zmm180,
                "tipo": "especial",
            },
        ]

    # ==========================================================================
    # SIDEBAR
    # ==========================================================================
    def _construir_sidebar(self):
        self.logo_label = ctk.CTkLabel(
            self.sidebar_frame,
            text="MRV&CO\nAutomações",
            font=ctk.CTkFont(size=21, weight="bold"),
            text_color=self.COR_TEXTO
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=(30, 25))

        self.btn_nav_inicio = self._criar_botao_menu("🏠  Início", 1, "inicio")
        self.btn_nav_robos = self._criar_botao_menu("🤖  Robôs", 2, "robos")
        self.btn_nav_config = self._criar_botao_menu("⚙️  Configurações", 3, "config")
        self.btn_nav_ajuda = self._criar_botao_menu("❓  Ajuda", 4, "ajuda")

        self.sidebar_status = ctk.CTkFrame(self.sidebar_frame, fg_color="#111111", corner_radius=12)
        self.sidebar_status.grid(row=8, column=0, padx=14, pady=16, sticky="sew")

        ctk.CTkLabel(
            self.sidebar_status,
            text="Status do Hub",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=self.COR_TEXTO
        ).pack(anchor="w", padx=12, pady=(10, 4))

        self.lbl_status_credenciais = ctk.CTkLabel(
            self.sidebar_status,
            text=self._texto_status_credenciais(),
            font=ctk.CTkFont(size=12),
            text_color=self.COR_TEXTO_FRACO
        )
        self.lbl_status_credenciais.pack(anchor="w", padx=12, pady=(0, 10))

        ctk.CTkLabel(
            self.sidebar_status,
            text="Versão 3.0",
            font=ctk.CTkFont(size=11),
            text_color="#7D7D7D"
        ).pack(anchor="w", padx=12, pady=(0, 10))

    def _criar_botao_menu(self, texto, linha, nome_tela):
        btn = ctk.CTkButton(
            self.sidebar_frame,
            text=texto,
            corner_radius=0,
            height=52,
            border_spacing=12,
            text_color=self.COR_TEXTO,
            hover_color=self.COR_SIDEBAR_HOVER,
            fg_color="transparent",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w",
            command=lambda: self.selecionar_tela(nome_tela)
        )
        btn.grid(row=linha, column=0, sticky="ew")
        return btn

    def selecionar_tela(self, nome_tela):
        self.tela_atual = nome_tela

        self.btn_nav_inicio.configure(fg_color=self.COR_SIDEBAR_ACTIVE if nome_tela == "inicio" else "transparent")
        self.btn_nav_robos.configure(fg_color=self.COR_SIDEBAR_ACTIVE if nome_tela == "robos" else "transparent")
        self.btn_nav_config.configure(fg_color=self.COR_SIDEBAR_ACTIVE if nome_tela == "config" else "transparent")
        self.btn_nav_ajuda.configure(fg_color=self.COR_SIDEBAR_ACTIVE if nome_tela == "ajuda" else "transparent")

        for frame in self.frames.values():
            frame.grid_forget()

        self.frames[nome_tela].grid(row=0, column=0, sticky="nsew")

        if nome_tela == "inicio":
            self._atualizar_dashboard()

    # ==========================================================================
    # HELPERS VISUAIS
    # ==========================================================================
    def _criar_frame_base(self, nome):
        frame = ctk.CTkFrame(self.main_frame, fg_color=self.COR_BG)
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        self.frames[nome] = frame
        return frame

    def _card(self, parent, fg=None, radius=14):
        return ctk.CTkFrame(parent, fg_color=fg or self.COR_CARD, corner_radius=radius)

    def _label_titulo(self, parent, texto):
        return ctk.CTkLabel(
            parent,
            text=texto,
            font=ctk.CTkFont(size=26, weight="bold"),
            text_color=self.COR_TEXTO
        )

    def _label_sub(self, parent, texto):
        return ctk.CTkLabel(
            parent,
            text=texto,
            font=ctk.CTkFont(size=14),
            text_color=self.COR_TEXTO_FRACO
        )

    def _criar_chip(self, parent, texto, cor="#444444"):
        chip = ctk.CTkLabel(
            parent,
            text=texto,
            fg_color=cor,
            corner_radius=12,
            padx=8,
            pady=3,
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="white"
        )
        return chip

    # ==========================================================================
    # DASHBOARD
    # ==========================================================================
    def _construir_tela_inicio(self):
        frame = self._criar_frame_base("inicio")

        container = ctk.CTkScrollableFrame(frame, fg_color=self.COR_BG)
        container.pack(fill=tk.BOTH, expand=True, padx=28, pady=22)

        header = ctk.CTkFrame(container, fg_color="transparent")
        header.pack(fill=tk.X, pady=(0, 18))

        self.dashboard_titulo = self._label_titulo(header, "Bem-vindo ao Hub Central MRV")
        self.dashboard_titulo.pack(anchor="w")

        self.dashboard_sub = self._label_sub(
            header,
            "Acompanhe suas automações, execute ações rápidas e veja o histórico recente."
        )
        self.dashboard_sub.pack(anchor="w", pady=(4, 0))

        cards = ctk.CTkFrame(container, fg_color="transparent")
        cards.pack(fill=tk.X, pady=(0, 18))
        cards.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.card_total_robos = self._criar_card_metricas(cards, 0, "🤖", "Robôs disponíveis", "0")
        self.card_ultimo_robo = self._criar_card_metricas(cards, 1, "🕒", "Última execução", "Nenhuma")
        self.card_status = self._criar_card_metricas(cards, 2, "🔐", "Credenciais", "Verificando")

        self.card_versao = self._criar_card_metricas(cards, 3, "🚀", "Versão", "3.0")
        secao_rapida = self._card(container)
        secao_rapida.pack(fill=tk.X, pady=(0, 18))

        header_rapido = ctk.CTkFrame(secao_rapida, fg_color="transparent")
        header_rapido.pack(fill=tk.X, padx=18, pady=(16, 4))

        ctk.CTkLabel(
            header_rapido,
            text="Ações rápidas",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=self.COR_TEXTO
        ).pack(side="left")

        ctk.CTkButton(
            header_rapido,
            text="Configurar",
            width=120,
            height=30,
            fg_color=self.COR_CINZA,
            hover_color="#4A5560",
            command=self._abrir_config_acoes_rapidas
        ).pack(side="right")

        ctk.CTkLabel(
            secao_rapida,
            text="Escolha os robôs que devem aparecer como atalhos. Ao iniciar, o Hub abrirá a tela de execução automaticamente.",
            font=ctk.CTkFont(size=13),
            text_color=self.COR_TEXTO_FRACO
        ).pack(anchor="w", padx=18, pady=(0, 10))

        self.grid_rapido = ctk.CTkFrame(secao_rapida, fg_color="transparent")
        self.grid_rapido.pack(fill=tk.X, padx=18, pady=(0, 18))
        self.grid_rapido.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self._renderizar_acoes_rapidas()
        historico_card = self._card(container)
        historico_card.pack(fill=tk.BOTH, expand=True)

        ctk.CTkLabel(
            historico_card,
            text="Últimas execuções",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=self.COR_TEXTO
        ).pack(anchor="w", padx=18, pady=(16, 8))

        self.frame_historico = ctk.CTkFrame(historico_card, fg_color="transparent")
        self.frame_historico.pack(fill=tk.X, padx=18, pady=(0, 18))

    def _criar_card_metricas(self, parent, col, icone, titulo, valor):
        card = self._card(parent)
        card.grid(row=0, column=col, sticky="ew", padx=7)

        ctk.CTkLabel(
            card,
            text=icone,
            font=ctk.CTkFont(size=24)
        ).pack(anchor="w", padx=16, pady=(14, 4))

        lbl_valor = ctk.CTkLabel(
            card,
            text=valor,
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=self.COR_TEXTO
        )
        lbl_valor.pack(anchor="w", padx=16)

        ctk.CTkLabel(
            card,
            text=titulo,
            font=ctk.CTkFont(size=12),
            text_color=self.COR_TEXTO_FRACO
        ).pack(anchor="w", padx=16, pady=(0, 14))

        return lbl_valor

    def _criar_botao_acao_rapida(self, parent, robo, col):
        card = ctk.CTkButton(
            parent,
            text=f"{robo['icone']}  {robo['nome']}",
            height=52,
            fg_color=self.COR_MRV,
            hover_color=self.COR_MRV_HOVER,
            font=ctk.CTkFont(size=13, weight="bold"),
            command=lambda r=robo: self._executar_robo(r)
        )
        card.grid(row=0, column=col, sticky="ew", padx=6, pady=6)
        self.todos_botoes.append(card)

    def _atualizar_dashboard(self):
        if not hasattr(self, "card_total_robos"):
            return

        historico = self._carregar_historico()
        ultimo = historico[0] if historico else None

        self.card_total_robos.configure(text=str(len(self.robos)))
        self.card_ultimo_robo.configure(text=ultimo["robo"] if ultimo else "Nenhuma")
        self.card_status.configure(text="Configuradas" if self._credenciais_ok() else "Pendentes")

        for widget in self.frame_historico.winfo_children():
            widget.destroy()

        if not historico:
            ctk.CTkLabel(
                self.frame_historico,
                text="Nenhuma execução registrada ainda.",
                text_color=self.COR_TEXTO_FRACO,
                font=ctk.CTkFont(size=13)
            ).pack(anchor="w", pady=8)
            return

        for item in historico[:8]:
            self._linha_historico(self.frame_historico, item)

        self._renderizar_acoes_rapidas()

    def _linha_historico(self, parent, item):
        status = item.get("status", "")
        icone = "✅" if status == "sucesso" else "❌" if status == "erro" else "⚠️"

        linha = ctk.CTkFrame(parent, fg_color=self.COR_CARD_2, corner_radius=10)
        linha.pack(fill=tk.X, pady=4)

        ctk.CTkLabel(
            linha,
            text=f"{icone}  {item.get('robo', 'Processo')}",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=self.COR_TEXTO
        ).pack(side="left", padx=12, pady=9)

        ctk.CTkLabel(
            linha,
            text=f"{item.get('data', '')}  •  {status.upper()}",
            font=ctk.CTkFont(size=12),
            text_color=self.COR_TEXTO_FRACO
        ).pack(side="right", padx=12, pady=9)

    # ==========================================================================
    # TELA ROBÔS
    # ==========================================================================
    def _construir_tela_robos(self):
        frame = self._criar_frame_base("robos")

        # Layout principal da aba Robôs
        frame.grid_rowconfigure(0, weight=0)  # Header
        frame.grid_rowconfigure(1, weight=0)  # Busca
        frame.grid_rowconfigure(2, weight=1)  # Lista de robôs
        frame.grid_rowconfigure(3, weight=0)  # Execução/log
        frame.grid_columnconfigure(0, weight=1)

        # =========================
        # HEADER
        # =========================
        header = ctk.CTkFrame(frame, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=26, pady=(18, 8))
        header.grid_columnconfigure(0, weight=1)

        self._label_titulo(
            header,
            "🤖 Central de Robôs - Administrativo MRV"
        ).grid(row=0, column=0, sticky="w")

        self._label_sub(
            header,
            "Escolha um robô, confira os requisitos e acompanhe a execução em tempo real."
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        # =========================
        # BUSCA
        # =========================
        barra_busca = ctk.CTkFrame(frame, fg_color="transparent")
        barra_busca.grid(row=1, column=0, sticky="ew", padx=26, pady=(0, 8))
        barra_busca.grid_columnconfigure(0, weight=1)

        self.entry_busca_robos = ctk.CTkEntry(
            barra_busca,
            placeholder_text="Buscar robô por nome, categoria ou requisito...",
            height=38
        )
        self.entry_busca_robos.grid(row=0, column=0, sticky="ew")

        # Busca com pequeno debounce para não renderizar a cada tecla imediatamente
        self._busca_after_id = None
        self.entry_busca_robos.bind("<KeyRelease>", self._on_busca_robos)

        ctk.CTkButton(
            barra_busca,
            text="Limpar",
            width=90,
            height=38,
            fg_color=self.COR_CINZA,
            hover_color="#4A5560",
            command=self._limpar_busca_robos
        ).grid(row=0, column=1, padx=(10, 0))

        # =========================
        # LISTA DE ROBÔS
        # =========================
        self.frame_botoes_robos = ctk.CTkScrollableFrame(
            frame,
            fg_color="transparent"
        )
        self.frame_botoes_robos.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 8))

        # =========================
        # ÁREA DE EXECUÇÃO FIXA NO RODAPÉ
        # =========================
        self.frame_execucao = self._card(frame)
        self.frame_execucao.grid(row=3, column=0, sticky="ew", padx=26, pady=(0, 14))
        self.frame_execucao.grid_columnconfigure(0, weight=1)

        topo_exec = ctk.CTkFrame(self.frame_execucao, fg_color="transparent")
        topo_exec.grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 4))
        topo_exec.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            topo_exec,
            text="Execução",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=self.COR_TEXTO
        ).grid(row=0, column=0, sticky="w")

        self.btn_toggle_logs = ctk.CTkButton(
            topo_exec,
            text="Ocultar logs",
            width=110,
            height=28,
            fg_color=self.COR_CINZA,
            hover_color="#4A5560",
            command=self._alternar_logs
        )
        self.btn_toggle_logs.grid(row=0, column=1, sticky="e")

        self.btn_cancelar = ctk.CTkButton(
            self.frame_execucao,
            text="CANCELAR PROCESSO ATIVO",
            fg_color=self.COR_CANCELAR,
            hover_color=self.COR_CANCELAR_HOVER,
            font=ctk.CTkFont(size=13, weight="bold"),
            height=36,
            command=self.cancelar_processo,
            state="disabled"
        )
        self.btn_cancelar.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 6))

        self.progressbar = ctk.CTkProgressBar(
            self.frame_execucao,
            mode="determinate",
            height=8,
            progress_color=self.COR_MRV
        )
        self.progressbar.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 6))
        self.progressbar.set(0)

        self.console_container = ctk.CTkFrame(self.frame_execucao, fg_color="transparent")
        self.console_container.grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 12))
        self.console_container.grid_columnconfigure(0, weight=1)

        self.console = ctk.CTkTextbox(
            self.console_container,
            height=95,
            font=ctk.CTkFont(family="Consolas", size=12),
            text_color="#00FF06",
            fg_color="#141414",
            border_width=1,
            border_color="#3A3A3A"
        )
        self.console.grid(row=0, column=0, sticky="ew")

        self.console.configure(state="disabled")

        sys.stdout = PrintRedirector(self.console)
        sys.stderr = PrintRedirector(self.console)

        print("Sistema Central iniciado com sucesso!")
        print("Selecione o processo que deseja executar.")
        print("-" * 60)

        self._renderizar_robos()

    def _renderizar_robos(self):
        if not hasattr(self, "frame_botoes_robos"):
            return

        for widget in self.frame_botoes_robos.winfo_children():
            widget.destroy()

        self.todos_botoes = []

        termo = ""
        if hasattr(self, "entry_busca_robos"):
            termo = self.entry_busca_robos.get().strip().lower()

        categorias = {}
        for robo in self.robos:
            texto_busca = " ".join([
                robo.get("nome", ""),
                robo.get("titulo", ""),
                robo.get("categoria", ""),
                robo.get("descricao", ""),
                " ".join(robo.get("requisitos", []))
            ]).lower()

            if termo and termo not in texto_busca:
                continue

            categorias.setdefault(robo["categoria"], []).append(robo)

        if not categorias:
            ctk.CTkLabel(
                self.frame_botoes_robos,
                text="Nenhum robô encontrado para a busca informada.",
                text_color=self.COR_TEXTO_FRACO,
                font=ctk.CTkFont(size=14)
            ).pack(pady=30)
            return

        row = 0
        col = 0
        grid = ctk.CTkFrame(self.frame_botoes_robos, fg_color="transparent")
        grid.pack(fill=tk.BOTH, expand=True)
        grid.grid_columnconfigure((0, 1), weight=1)

        for categoria, lista in categorias.items():
            card = self._criar_card_categoria(grid, categoria, lista)
            card.grid(row=row, column=col, sticky="nsew", padx=8, pady=8)

            col += 1
            if col > 1:
                col = 0
                row += 1

    def _criar_card_categoria(self, parent, categoria, lista_robos):
        card = self._card(parent)
        cor = self.cores_categoria.get(categoria, self.COR_MRV)

        header = ctk.CTkFrame(card, fg_color="transparent")
        header.pack(fill=tk.X, padx=14, pady=(14, 8))

        ctk.CTkLabel(
            header,
            text=categoria,
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=self.COR_TEXTO
        ).pack(side="left")

        self._criar_chip(header, f"{len(lista_robos)} robôs", cor).pack(side="right")

        for robo in lista_robos:
            self._criar_linha_robo(card, robo)

        return card

    def _criar_linha_robo(self, parent, robo):
        linha = ctk.CTkFrame(parent, fg_color=self.COR_CARD_2, corner_radius=12)
        linha.pack(fill=tk.X, padx=12, pady=6)

        topo = ctk.CTkFrame(linha, fg_color="transparent")
        topo.pack(fill=tk.X, padx=12, pady=(10, 2))

        ctk.CTkLabel(
            topo,
            text=f"{robo['icone']}  {robo['titulo']}",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=self.COR_TEXTO
        ).pack(side="left", anchor="w")

        Prioridade = robo.get("Prioridade", "Média")
        cor_Prioridade = self.COR_MRV if Prioridade == "Baixo" else self.COR_AMARELO if Prioridade == "Médio" else self.COR_CANCELAR

        self._criar_chip(topo, f"Prioridade {Prioridade}", cor_Prioridade).pack(side="right")

        ctk.CTkLabel(
            linha,
            text=robo.get("descricao", ""),
            font=ctk.CTkFont(size=12),
            text_color=self.COR_TEXTO_FRACO,
            wraplength=520,
            justify="left"
        ).pack(anchor="w", padx=12, pady=(0, 4))

        info = ctk.CTkFrame(linha, fg_color="transparent")
        info.pack(fill=tk.X, padx=12, pady=(0, 8))

        self._criar_chip(info, f"⏱ {robo.get('tempo', 'Variável')}", "#444444").pack(side="left", padx=(0, 5))

        requisitos = robo.get("requisitos", [])
        if requisitos:
            self._criar_chip(info, " • ".join(requisitos[:3]), "#3D4852").pack(side="left", padx=(0, 5))

        btn = ctk.CTkButton(
            info,
            text="Executar",
            width=100,
            height=30,
            fg_color=self.COR_MRV,
            hover_color=self.COR_MRV_HOVER,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda r=robo: self._executar_robo(r)
        )
        btn.pack(side="right")
        self.todos_botoes.append(btn)

    def _limpar_busca_robos(self):
        self.entry_busca_robos.delete(0, tk.END)
        self._renderizar_robos()

    def _alternar_logs(self):
        self.logs_visiveis = not self.logs_visiveis

        if self.logs_visiveis:
            self.console_container.grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 12))
            self.btn_toggle_logs.configure(text="Ocultar logs")
        else:
            self.console_container.grid_remove()
            self.btn_toggle_logs.configure(text="Mostrar logs")

    def _on_busca_robos(self, event=None):
        if hasattr(self, "_busca_after_id") and self._busca_after_id:
            try:
                self.root.after_cancel(self._busca_after_id)
            except Exception:
                pass

        self._busca_after_id = self.root.after(120, self._renderizar_robos)
        
    # ==========================================================================
    # CONFIGURAÇÕES
    # ==========================================================================
    def _construir_tela_config(self):
        frame = self._criar_frame_base("config")

        container = ctk.CTkScrollableFrame(frame, fg_color=self.COR_BG)
        container.pack(fill=tk.BOTH, expand=True, padx=28, pady=22)

        self._label_titulo(container, "⚙️ Configurar Credenciais do Sistema").pack(anchor="w")
        self._label_sub(
            container,
            "Salve as credenciais usadas pelos robôs. Mantenha este arquivo seguro."
        ).pack(anchor="w", pady=(4, 16))

        status_card = self._card(container)
        status_card.pack(fill=tk.X, pady=(0, 16))

        ctk.CTkLabel(
            status_card,
            text="Status das credenciais",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=self.COR_TEXTO
        ).pack(anchor="w", padx=16, pady=(14, 4))

        self.lbl_config_status = ctk.CTkLabel(
            status_card,
            text=self._texto_status_credenciais(detalhado=True),
            font=ctk.CTkFont(size=13),
            text_color=self.COR_TEXTO_FRACO,
            justify="left"
        )
        self.lbl_config_status.pack(anchor="w", padx=16, pady=(0, 14))

        grid = ctk.CTkFrame(container, fg_color="transparent")
        grid.pack(fill=tk.BOTH, expand=True)
        grid.grid_columnconfigure((0, 1), weight=1)

        col_esq = self._card(grid)
        col_esq.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=0)

        col_dir = self._card(grid)
        col_dir.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=0)

        ctk.CTkLabel(
            col_esq,
            text="SISTEMAS MRV & AGILIS",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=self.COR_MRV
        ).pack(pady=(16, 12))

        self.entry_email = self._campo_config(col_esq, "E-mail MRV:", getattr(config, "EMAIL_USER", ""))
        self.entry_senha = self._campo_config(col_esq, "Senha MRV:", getattr(config, "SENHA_USER", ""), senha=True)
        self.entry_senha_malote = self._campo_config(col_esq, "Senha Malote Web:", getattr(config, "SENHA_MALOTE", ""), senha=True)
        self.entry_API_KEY_AGILIS = self._campo_config(col_esq, "Chave API Agilis:", getattr(config, "CHAVE_API_AGILIS", ""), senha=True)

        ctk.CTkButton(
            col_esq,
            text="Como gerar a chave API do Agilis?",
            height=32,
            fg_color=self.COR_CINZA,
            hover_color="#4A5560",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._abrir_tutorial_api_agilis
        ).pack(anchor="w", padx=28, pady=(0, 12))

        ctk.CTkLabel(
            col_dir,
            text="CORREIOS (SEDEX REVERSO)",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=self.COR_AMARELO
        ).pack(pady=(16, 12))

        self.entry_correios_cod = self._campo_config(col_dir, "Código Administrativo:", getattr(config, "CORREIOS_COD_ADM", ""))
        self.entry_correios_email = self._campo_config(col_dir, "E-mail Correios:", getattr(config, "CORREIOS_EMAIL", ""))
        self.entry_correios_senha = self._campo_config(col_dir, "Senha Correios:", getattr(config, "CORREIOS_SENHA", ""), senha=True)

        ctk.CTkLabel(
            col_dir,
            text="PODIO API (MENSAGERIA)",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=self.COR_VINHO
        ).pack(pady=(16, 12))

        self.entry_podio_client_id = self._campo_config(
            col_dir, 
            "Podio Client ID:", 
            getattr(config, "PODIO_CLIENT_ID", "")
        )
        self.entry_podio_client_secret = self._campo_config(
            col_dir, 
            "Podio Client Secret:", 
            getattr(config, "PODIO_CLIENT_SECRET", ""), 
            senha=True
        )

        self.entry_podio_app_id = self._campo_config(
            col_dir, 
            "Podio App ID:", 
            str(getattr(config, "PODIO_APP_ID", ""))
        )
        self.entry_podio_app_token = self._campo_config(
            col_dir, 
            "Podio App Token:", 
            getattr(config, "PODIO_APP_TOKEN", ""), 
            senha=True
        )

        botoes = ctk.CTkFrame(container, fg_color="transparent")
        botoes.pack(fill=tk.X, pady=20)

        ctk.CTkButton(
            botoes,
            text="Salvar Todas as Credenciais",
            command=self._salvar_credenciais,
            height=44,
            width=300,
            fg_color=self.COR_MRV,
            hover_color=self.COR_MRV_HOVER,
            font=ctk.CTkFont(weight="bold")
        ).pack(side="left")

        ctk.CTkButton(
            botoes,
            text="Voltar para Robôs",
            command=lambda: self.selecionar_tela("robos"),
            height=44,
            width=180,
            fg_color=self.COR_CINZA,
            hover_color="#4A5560",
            font=ctk.CTkFont(weight="bold")
        ).pack(side="left", padx=10)

    def _campo_config(self, parent, label, valor, senha=False):
        ctk.CTkLabel(
            parent,
            text=label,
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=self.COR_TEXTO
        ).pack(anchor="w", padx=28, pady=(8, 2))

        linha = ctk.CTkFrame(parent, fg_color="transparent")
        linha.pack(fill=tk.X, padx=28, pady=(0, 8))

        entry = ctk.CTkEntry(linha, height=32, show="*" if senha else "")
        entry.pack(side="left", fill=tk.X, expand=True)
        entry.insert(0, valor or "")

        if senha:
            btn = ctk.CTkButton(
                linha,
                text="👁",
                width=38,
                height=32,
                fg_color=self.COR_CINZA,
                hover_color="#4A5560",
                command=lambda e=entry: self._alternar_visao_senha(e)
            )
            btn.pack(side="right", padx=(6, 0))

        return entry

    def _alternar_visao_senha(self, entry):
        atual = entry.cget("show")
        entry.configure(show="" if atual == "*" else "*")

    def _salvar_credenciais(self):
        novo_email = self.entry_email.get().strip()
        novo_senha = self.entry_senha.get().strip()
        nova_senha_malote = self.entry_senha_malote.get().strip()
        nova_api_key_agilis = self.entry_API_KEY_AGILIS.get().strip()

        novo_correios_cod = self.entry_correios_cod.get().strip()
        novo_correios_email = self.entry_correios_email.get().strip()
        novo_correios_senha = self.entry_correios_senha.get().strip()
        
        # Captura os campos do Podio (incluindo os novos)
        novo_podio_id = self.entry_podio_client_id.get().strip()
        novo_podio_secret = self.entry_podio_client_secret.get().strip()
        novo_podio_app_id = self.entry_podio_app_id.get().strip()
        novo_podio_app_token = self.entry_podio_app_token.get().strip()

        # Salva no arquivo JSON config_mrv.json
        config.salvar_credenciais(
            novo_email,
            novo_senha,
            nova_senha_malote,
            nova_api_key_agilis,
            novo_correios_cod,
            novo_correios_email,
            novo_correios_senha,
            novo_podio_id,
            novo_podio_secret,
            novo_podio_app_id,
            novo_podio_app_token
        )

        # Atualiza as variáveis em memória para uso imediato
        config.EMAIL_USER = novo_email
        config.SENHA_USER = novo_senha
        config.SENHA_MALOTE = nova_senha_malote
        config.CHAVE_API_AGILIS = nova_api_key_agilis
        config.CORREIOS_COD_ADM = novo_correios_cod
        config.CORREIOS_EMAIL = novo_correios_email
        config.CORREIOS_SENHA = novo_correios_senha
        
        config.PODIO_CLIENT_ID = novo_podio_id
        config.PODIO_CLIENT_SECRET = novo_podio_secret
        config.PODIO_APP_ID = novo_podio_app_id
        config.PODIO_APP_TOKEN = novo_podio_app_token

        config.EMAIL_MRV = novo_email
        config.SENHA_MRV = novo_senha
        config.SENHA_MALOTE_MRV = nova_senha_malote
        config.API_KEY_AGILIS = nova_api_key_agilis

        self._salvar_metadata_config()
        self.lbl_status_credenciais.configure(text=self._texto_status_credenciais())
        self.lbl_config_status.configure(text=self._texto_status_credenciais(detalhado=True))

        messagebox.showinfo("Sucesso", "Todas as credenciais foram salvas com sucesso!")
        self.selecionar_tela("robos")

    # ==========================================================================
    # AJUDA
    # ==========================================================================
    def _construir_tela_ajuda(self):
        frame = self._criar_frame_base("ajuda")

        container = ctk.CTkFrame(frame, fg_color=self.COR_BG)
        container.pack(fill=tk.BOTH, expand=True, padx=28, pady=22)
        container.grid_columnconfigure(1, weight=1)
        container.grid_rowconfigure(1, weight=1)

        self._label_titulo(container, "❓ Guia de Uso e Tutorial").grid(row=0, column=0, columnspan=2, sticky="w")
        self._label_sub(container, "Consulte instruções por tema. Use Ctrl+F para buscar dentro do texto.").grid(
            row=0, column=1, sticky="e"
        )

        menu = self._card(container)
        menu.grid(row=1, column=0, sticky="ns", padx=(0, 12), pady=(16, 0))

        conteudo = self._card(container)
        conteudo.grid(row=1, column=1, sticky="nsew", pady=(16, 0))
        conteudo.grid_rowconfigure(0, weight=1)
        conteudo.grid_columnconfigure(0, weight=1)

        self.textbox_ajuda = ctk.CTkTextbox(
            conteudo,
            wrap="word",
            font=ctk.CTkFont(size=14),
            fg_color="#333333"
        )
        self.textbox_ajuda.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        self.secoes_ajuda = {
            "Primeiros passos": self._texto_ajuda_primeiros_passos(),
            "Credenciais": self._texto_ajuda_credenciais(),
            "Chave API Agilis": self._texto_ajuda_api_agilis(),
            "Correios e Faturamento": self._texto_ajuda_correios(),
            "Agilis e Chamados": self._texto_ajuda_agilis(),
            "Uber, SAP e Contratos": self._texto_ajuda_uber_sap(),
            "Cancelamento": self._texto_ajuda_cancelamento(),
            "FAQ": self._texto_ajuda_faq(),
        }

        for nome in self.secoes_ajuda:
            ctk.CTkButton(
                menu,
                text=nome,
                height=38,
                anchor="w",
                fg_color="transparent",
                hover_color=self.COR_SIDEBAR_HOVER,
                command=lambda n=nome: self._mostrar_secao_ajuda(n)
            ).pack(fill=tk.X, padx=10, pady=4)

        self._mostrar_secao_ajuda("Primeiros passos")

    def _texto_ajuda_api_agilis(self):
        return """CHAVE API AGILIS

    Alguns robôs precisam da Chave API do Agilis para consultar informações automaticamente.

    Robôs que dependem da chave:

    - Gerar relatório de envio para Correios
    - Rateio AGF

    Como gerar:

    1. Acesse o Agilis normalmente.
    2. Clique no ícone do perfil no canto superior direito.
    3. Clique em "Gerar chave API".
    4. Selecione "Nunca expira".
    5. Clique em "Gerar" ou "Regerar".
    6. Copie a chave gerada.
    7. Volte ao Hub Central MRV.
    8. Vá em Configurações.
    9. Cole a chave no campo "Chave API Agilis".
    10. Clique em "Salvar Todas as Credenciais".

    Observação:
    Na primeira geração, o botão pode aparecer com outro nome. Se a chave já existir, normalmente aparece como "Regerar". O processo é o mesmo.

    Importante:
    Não compartilhe sua chave API. Se esquecer a chave, gere uma nova.
    """

    def _mostrar_secao_ajuda(self, nome):
        self.textbox_ajuda.configure(state="normal")
        self.textbox_ajuda.delete("1.0", tk.END)
        self.textbox_ajuda.insert("1.0", self.secoes_ajuda.get(nome, ""))
        self.textbox_ajuda.configure(state="disabled")

    def _atalho_busca(self, event=None):
        if self.tela_atual == "ajuda":
            JanelaBusca(self.root, self.textbox_ajuda)
        elif self.tela_atual == "robos" and hasattr(self, "entry_busca_robos"):
            self.entry_busca_robos.focus()

    # ==========================================================================
    # EXECUÇÃO DOS ROBÔS
    # ==========================================================================
    def _executar_robo(self, robo):
        if robo.get("requer_api_agilis") and not self._validar_api_agilis():
            return

        if robo.get("tipo") == "especial":
            handler = robo.get("handler")
            if handler:
                handler()
            else:
                messagebox.showerror(
                    "Erro",
                    f"O robô '{robo.get('nome')}' não possui função de execução configurada."
                )
            return

        if robo.get("tipo") == "pasta":
            self._verificar_pasta_e_executar(
                robo["nome"],
                robo["comando"],
                robo.get("pasta")
            )
            return

        self.executar_processo_cancelavel(robo["nome"], comando_python=robo["comando"])

    def _verificar_credenciais_iniciais(self):
        if not self._credenciais_ok():
            messagebox.showwarning(
                "Atenção: Credenciais Ausentes",
                "Bem-vindo ao Hub Central!\n\n"
                "Notamos que suas credenciais ainda não foram configuradas.\n\n"
                "Você será redirecionado para a tela de configurações."
            )
            self.selecionar_tela("config")

    def _chamar_robo_juridico(self):
        resposta = messagebox.askyesnocancel(
            "Relatório Jurídico Montreal",
            "Você deseja que o robô baixe a planilha do Podio automaticamente?\n\n"
            "SIM: Baixar relatório de hoje.\n"
            "NÃO: Utilizar relatório já baixado.\n"
            "CANCELAR: Sair."
        )

        if resposta is True:
            self.executar_processo_cancelavel(
                "Relatório Jurídico Montreal",
                comando_python="import robos.robo_juridico as rj; rj.executar_juridico(pular_download=False)"
            )
        elif resposta is False:
            self.executar_processo_cancelavel(
                "Relatório Jurídico (Apenas Formatação)",
                comando_python="import robos.robo_juridico as rj; rj.executar_juridico(pular_download=True)"
            )

    def _verificar_pasta_e_executar(self, nome_processo, comando_python, caminho_pasta):
        msg = (
            f"Você já verificou/atualizou os arquivos para o processo '{nome_processo}'?\n\n"
            "OK para rodar o robô.\n"
            "Cancelar para abrir a pasta."
        )

        if messagebox.askokcancel(f"Lembrete - {nome_processo}", msg):
            self.executar_processo_cancelavel(nome_processo, comando_python=comando_python)
        else:
            if caminho_pasta and not os.path.exists(caminho_pasta):
                try:
                    os.makedirs(caminho_pasta)
                except Exception:
                    pass

            if caminho_pasta and os.path.exists(caminho_pasta):
                try:
                    os.startfile(caminho_pasta)
                except Exception as e:
                    messagebox.showwarning("Aviso", f"Não foi possível abrir a pasta:\n{e}")
            else:
                messagebox.showwarning("Aviso", f"A pasta não foi encontrada:\n{caminho_pasta}")

    def _chamar_robo_produtividade(self):
        resposta = messagebox.askyesnocancel(
            "Produtividade Setorial",
            "Você deseja que o robô baixe os relatórios automaticamente?\n\n"
            "SIM: O robô fará download e edição.\n"
            "NÃO: Pular download e apenas formatar planilhas existentes.\n"
            "CANCELAR: Abortar operação."
        )

        if resposta is True:
            if messagebox.askokcancel(
                "Aviso Importante",
                "ATENÇÃO\n\n"
                "NÃO MEXA no mouse ou teclado durante a extração do Bússola.\n\n"
                "Deseja continuar?"
            ):
                self.executar_processo_cancelavel(
                    "Produtividade (Completo)",
                    comando_python="import robos.produtividade as rp; rp.executar_robo_produtividade_setor(pular_extracao=False)"
                )
        elif resposta is False:
            self.executar_processo_cancelavel(
                "Produtividade (Apenas Edição)",
                comando_python="import robos.produtividade as rp; rp.executar_robo_produtividade_setor(pular_extracao=True)"
            )

    def _chamar_robo_fechar_chamados(self):
        resposta = messagebox.askyesnocancel(
            "Fechar Chamados",
            "Escolha o modo de execução:\n\n"
            "SIM: Monitorar o dia todo.\n"
            "NÃO: Fechar TODOS de hoje.\n"
            "CANCELAR: Abortar operação."
        )

        if resposta is True:
            self.executar_processo_cancelavel(
                "Monitorar Chamados",
                comando_python="import robos.robo_fechar_chamados as rfc; rfc.executar_fechamento(modo='monitorar')"
            )
        elif resposta is False:
            self.executar_processo_cancelavel(
                "Fechar Todos de Hoje",
                comando_python="import robos.robo_fechar_chamados as rfc; rfc.executar_fechamento(modo='todos_hoje')"
            )

    def _chamar_robo_incluir_encomendas(self):
        if messagebox.askokcancel(
            "Lembrete - Correspondências",
            "Você lembrou de preencher a planilha?\n\n"
            "OK para rodar o robô.\n"
            "Cancelar para abrir a planilha."
        ):
            self.executar_processo_cancelavel(
                "Incluir Correspondências",
                comando_python="import robos.robo_incluir_encomendas as rie; rie.executar_inclusao()"
            )
        else:
            caminho_planilha = os.path.join(config.PASTA_ARQUIVOS, "encomendas", "encomendas.xlsx")
            if os.path.exists(caminho_planilha):
                os.startfile(caminho_planilha)
            else:
                messagebox.showwarning(
                    "Aviso",
                    "A planilha 'encomendas.xlsx' ainda não foi encontrada na pasta:\n"
                    f"{os.path.join(config.PASTA_ARQUIVOS, 'encomendas')}"
                )

    def _chamar_robo_zmm180(self):
        if messagebox.askokcancel(
            "Aviso Importante - SAP e Edge",
            "ATENÇÃO\n\n"
            "1. Deixe o SAP aberto na SEGUNDA TELA.\n"
            "2. Deixe o documento aberto no Edge.\n"
            "3. NÃO MEXA no mouse.\n\n"
            "Deseja continuar?"
        ):
            self.executar_processo_cancelavel(
                "Faturamento ZMM180",
                comando_python="import robos.robo_zmm180 as rz; rz.executar_zmm180()"
            )

    def executar_processo_cancelavel(self, nome_processo, comando_python=None):
        if not comando_python:
            messagebox.showerror("Erro", "Comando do robô não informado.")
            return
        
        self._preparar_tela_execucao()

        for btn in self.todos_botoes:
            try:
                btn.configure(state="disabled")
            except Exception:
                pass

        self.progressbar.set(0)
        self.btn_cancelar.configure(state="normal")

        print(f">>> Iniciando: {nome_processo}...")
        threading.Thread(
            target=self._rodar_subprocesso,
            args=(nome_processo, comando_python),
            daemon=True
        ).start()

    def _rodar_subprocesso(self, nome_processo, comando_python):
        self.foi_cancelado = False

        fd, log_path = tempfile.mkstemp(suffix=".log", text=True)
        os.close(fd)

        inicio = time.time()

        try:
            if getattr(sys, "frozen", False):
                cmd = [sys.executable, "--run-code", comando_python, log_path]
            else:
                cmd = [sys.executable, sys.argv[0], "--run-code", comando_python, log_path]

            processo = subprocess.Popen(cmd, creationflags=CREATE_NO_WINDOW)
            self.processo_ativo = processo
            self.root.after(0, lambda: self.btn_cancelar.configure(state="normal"))

            linhas_log = []

            with open(log_path, "r", encoding="utf-8") as f:
                while processo.poll() is None:
                    linha = f.readline()

                    if linha:
                        self._processar_linha_log(linha, linhas_log)
                    else:
                        time.sleep(0.1)

                for linha in f.readlines():
                    self._processar_linha_log(linha, linhas_log)

            self.processo_ativo = None
            self.root.after(0, lambda: self.btn_cancelar.configure(state="disabled"))

            duracao = round(time.time() - inicio, 1)

            if self.foi_cancelado:
                print("\nO processo foi cancelado pelo usuário.")
                self._registrar_historico(nome_processo, "cancelado", duracao)
                self.root.after(0, lambda: messagebox.showwarning("Cancelado", "O processo foi cancelado pelo usuário."))

            elif processo.returncode == 0:
                self.progressbar.set(1)
                print("\nProcesso finalizado com sucesso!")
                self._registrar_historico(nome_processo, "sucesso", duracao)
                self.root.after(0, lambda: messagebox.showinfo("Sucesso", "A automação foi concluída com sucesso!"))

            elif processo.returncode == 1:
                print(f"\nO processo falhou. Código {processo.returncode}.")
                self._registrar_historico(nome_processo, "erro", duracao)

                texto_erro = self._extrair_erro(linhas_log)
                mensagem_popup = f"O processo foi interrompido pelo seguinte motivo:\n\n{texto_erro}"

                self.root.after(0, lambda: messagebox.showerror("Erro na Automação", mensagem_popup))

            else:
                print(f"\nO processo foi encerrado. Código {processo.returncode}.")
                self._registrar_historico(nome_processo, "encerrado", duracao)
                self.root.after(0, lambda: messagebox.showwarning("Encerrado", "O processo foi encerrado."))

        except Exception as e:
            print(f"\nErro ao iniciar o processo: {e}")
            self._registrar_historico(nome_processo, "erro", 0)
            self.root.after(0, lambda msg=str(e): messagebox.showerror("Erro Crítico", msg))

        finally:
            print("-" * 60)
            self.root.after(0, self._reativar_botoes)

            try:
                os.remove(log_path)
            except Exception:
                pass

    def _processar_linha_log(self, linha, linhas_log):
        if "[PROGRESSO:" in linha:
            try:
                valor_str = linha.split("[PROGRESSO:")[1].replace("]", "").strip()
                valor = float(valor_str)
                self.root.after(0, lambda v=valor: self.progressbar.set(v / 100.0))
            except Exception:
                pass
        else:
            print(linha, end="")
            linhas_log.append(linha.rstrip("\r\n"))

    def _extrair_erro(self, linhas_log):
        linhas_erro = [l for l in linhas_log if l.strip()]

        for i in range(len(linhas_erro)):
            if "Traceback (most recent call last):" in linhas_erro[i]:
                return "\n".join(linhas_erro[i:])

        return linhas_erro[-1] if linhas_erro else "Erro desconhecido."

    def cancelar_processo(self):
        if self.processo_ativo and self.processo_ativo.poll() is None:
            if messagebox.askyesno(
                "Atenção",
                "Tem certeza que deseja cancelar o robô?\n\n"
                "Isso tentará encerrar o processo ativo e os subprocessos abertos por ele."
            ):
                try:
                    self.foi_cancelado = True
                    subprocess.run(
                        ["taskkill", "/F", "/T", "/PID", str(self.processo_ativo.pid)],
                        creationflags=CREATE_NO_WINDOW
                    )

                    print("\n" + "=" * 50)
                    print("PROCESSO CANCELADO FORÇADAMENTE PELO USUÁRIO!")
                    print("=" * 50 + "\n")

                except Exception as e:
                    print(f"\nErro ao tentar cancelar: {e}")

    def _reativar_botoes(self):
        for btn in self.todos_botoes:
            try:
                btn.configure(state="normal")
            except Exception:
                pass

        self.btn_cancelar.configure(state="disabled")
        self.progressbar.set(0)

        if self.tela_atual == "inicio":
            self._atualizar_dashboard()

    def _preparar_tela_execucao(self):
        if self.tela_atual != "robos":
            self.selecionar_tela("robos")

        if hasattr(self, "console_container") and not self.logs_visiveis:
            self._alternar_logs()

        try:
            self.root.update_idletasks()
        except Exception:
            pass

    # ==========================================================================
    # HISTÓRICO
    # ==========================================================================
    def _carregar_historico(self):
        if not os.path.exists(self.ARQUIVO_HISTORICO):
            return []

        try:
            with open(self.ARQUIVO_HISTORICO, "r", encoding="utf-8") as f:
                dados = json.load(f)

            if isinstance(dados, list):
                return dados

            return []

        except Exception:
            return []

    def _registrar_historico(self, robo, status, duracao):
        historico = self._carregar_historico()

        historico.insert(0, {
            "robo": robo,
            "status": status,
            "duracao_segundos": duracao,
            "data": datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        })

        historico = historico[:50]

        try:
            with open(self.ARQUIVO_HISTORICO, "w", encoding="utf-8") as f:
                json.dump(historico, f, ensure_ascii=False, indent=2)
        except Exception:
            pass


    # ==========================================================================
    # AÇÕES RÁPIDAS
    # ==========================================================================
    def _acoes_rapidas_padrao(self):
        return ["Produtividade", "Rateio de Malote", "Uber 1", "Fechar Chamados"]


    def _carregar_acoes_rapidas(self):
        if not os.path.exists(self.ARQUIVO_ACOES_RAPIDAS):
            return self._acoes_rapidas_padrao()

        try:
            with open(self.ARQUIVO_ACOES_RAPIDAS, "r", encoding="utf-8") as f:
                dados = json.load(f)

            if isinstance(dados, list) and dados:
                return dados

            return self._acoes_rapidas_padrao()

        except Exception:
            return self._acoes_rapidas_padrao()


    def _salvar_acoes_rapidas(self, lista_robos):
        try:
            with open(self.ARQUIVO_ACOES_RAPIDAS, "w", encoding="utf-8") as f:
                json.dump(lista_robos, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível salvar as ações rápidas:\n\n{e}")


    def _renderizar_acoes_rapidas(self):
        if not hasattr(self, "grid_rapido"):
            return

        for widget in self.grid_rapido.winfo_children():
            widget.destroy()

        nomes_acoes = self._carregar_acoes_rapidas()
        col = 0

        for nome in nomes_acoes:
            robo = self._buscar_robo_por_nome(nome)

            if not robo:
                continue

            self._criar_botao_acao_rapida(self.grid_rapido, robo, col)
            col += 1

            if col >= 4:
                break

        if col == 0:
            ctk.CTkLabel(
                self.grid_rapido,
                text="Nenhuma ação rápida configurada.",
                text_color=self.COR_TEXTO_FRACO,
                font=ctk.CTkFont(size=13)
            ).grid(row=0, column=0, sticky="w", padx=6, pady=8)


    def _abrir_config_acoes_rapidas(self):
        janela = ctk.CTkToplevel(self.root)
        janela.title("Configurar Ações Rápidas")
        janela.geometry("520x580")
        janela.resizable(False, False)
        janela.attributes("-topmost", True)

        ctk.CTkLabel(
            janela,
            text="Configurar ações rápidas",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=self.COR_TEXTO
        ).pack(anchor="w", padx=22, pady=(20, 4))

        ctk.CTkLabel(
            janela,
            text="Escolha até 4 robôs para aparecerem na tela inicial.",
            font=ctk.CTkFont(size=13),
            text_color=self.COR_TEXTO_FRACO
        ).pack(anchor="w", padx=22, pady=(0, 14))

        frame_lista = ctk.CTkScrollableFrame(janela, fg_color=self.COR_CARD, corner_radius=12)
        frame_lista.pack(fill=tk.BOTH, expand=True, padx=22, pady=(0, 14))

        selecionados_atuais = set(self._carregar_acoes_rapidas())
        variaveis = {}

        for robo in self.robos:
            var = tk.BooleanVar(value=robo["nome"] in selecionados_atuais)
            variaveis[robo["nome"]] = var

            texto = f"{robo['icone']}  {robo['nome']}  •  {robo['categoria']}"

            chk = ctk.CTkCheckBox(
                frame_lista,
                text=texto,
                variable=var,
                font=ctk.CTkFont(size=13),
                checkbox_width=20,
                checkbox_height=20
            )
            chk.pack(anchor="w", padx=14, pady=7)

        frame_botoes = ctk.CTkFrame(janela, fg_color="transparent")
        frame_botoes.pack(fill=tk.X, padx=22, pady=(0, 20))

        def salvar():
            escolhidos = [
                nome for nome, var in variaveis.items()
                if var.get()
            ]

            if len(escolhidos) == 0:
                messagebox.showwarning(
                    "Ações rápidas",
                    "Selecione pelo menos 1 robô para aparecer nas ações rápidas."
                )
                return

            if len(escolhidos) > 4:
                messagebox.showwarning(
                    "Ações rápidas",
                    "Selecione no máximo 4 robôs."
                )
                return

            self._salvar_acoes_rapidas(escolhidos)
            self._renderizar_acoes_rapidas()
            janela.destroy()

        def restaurar_padrao():
            self._salvar_acoes_rapidas(self._acoes_rapidas_padrao())
            self._renderizar_acoes_rapidas()
            janela.destroy()

        ctk.CTkButton(
            frame_botoes,
            text="Salvar",
            height=40,
            fg_color=self.COR_MRV,
            hover_color=self.COR_MRV_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=salvar
        ).pack(side="left", fill=tk.X, expand=True, padx=(0, 6))

        ctk.CTkButton(
            frame_botoes,
            text="Restaurar padrão",
            height=40,
            fg_color=self.COR_CINZA,
            hover_color="#4A5560",
            font=ctk.CTkFont(weight="bold"),
            command=restaurar_padrao
        ).pack(side="left", fill=tk.X, expand=True, padx=(6, 0))

    # ==========================================================================
    # REDIMENSIONAMENTO DA JANELA (RESIZE)
    # ==========================================================================
    def _on_resize(self, event=None):
        """Gerencia o redimensionamento da janela com debounce para evitar travamentos."""
        if not hasattr(self, "frame_botoes_robos") or not hasattr(self, "console"):
            return

        if self._resize_after_id:
            try:
                self.root.after_cancel(self._resize_after_id)
            except Exception:
                pass

        self._resize_after_id = self.root.after(100, self._ajustar_alturas_resize)

    def _ajustar_alturas_resize(self):
        """Ajusta dinamicamente a altura da lista de robôs e do console com base na janela."""
        try:
            altura_janela = self.root.winfo_height()
        except Exception:
            return

        if altura_janela <= 720:
            altura_lista = 210
            altura_console = 75
        elif altura_janela <= 820:
            altura_lista = 260
            altura_console = 90
        elif altura_janela <= 920:
            altura_lista = 320
            altura_console = 110
        else:
            altura_lista = 380
            altura_console = 130

        try:
            self.frame_botoes_robos.configure(height=altura_lista)
            self.console.configure(height=altura_console)
        except Exception:
            pass

    # ==========================================================================
    # STATUS E UTILITÁRIOS
    # ==========================================================================
    def _validar_api_agilis(self):
        """Verifica se a chave API do Agilis está configurada."""
        api_key = getattr(config, "CHAVE_API_AGILIS", "") or getattr(config, "API_KEY_AGILIS", "")
        if not api_key or api_key.strip() == "":
            messagebox.showwarning(
                "Chave API Agilis Ausente",
                "Este robô requer a Chave API do Agilis para funcionar.\n\n"
                "Por favor, configure sua chave API na aba 'Configurações' antes de continuar."
            )
            self.selecionar_tela("config")
            return False
        return True

    def _credenciais_ok(self):
        email = getattr(config, "EMAIL_MRV", "") or getattr(config, "EMAIL_USER", "")
        senha = getattr(config, "SENHA_MRV", "") or getattr(config, "SENHA_USER", "")

        email_vazio = not email or email == "seu_email@mrv.com.br"
        senha_vazia = not senha or senha == "sua_senha"

        return not email_vazio and not senha_vazia

    def _texto_status_credenciais(self, detalhado=False):
        if self._credenciais_ok():
            base = "✅ Credenciais configuradas"
        else:
            base = "⚠️ Credenciais pendentes"

        if not detalhado:
            return base

        ultima = self._ler_metadata_config().get("ultima_atualizacao", "Não registrada")

        return (
            f"{base}\n"
            f"Última atualização: {ultima}\n\n"
            "Dica: não compartilhe o arquivo config_mrv.json com terceiros."
        )

    def _salvar_metadata_config(self):
        meta_path = os.path.join(self.PASTA_BASE, "config_status.json")

        dados = {
            "ultima_atualizacao": datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
        }

        try:
            with open(meta_path, "w", encoding="utf-8") as f:
                json.dump(dados, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _ler_metadata_config(self):
        meta_path = os.path.join(self.PASTA_BASE, "config_status.json")

        if not os.path.exists(meta_path):
            return {}

        try:
            with open(meta_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def _buscar_robo_por_nome(self, nome):
        for robo in self.robos:
            if robo.get("nome") == nome:
                return robo
        return None

    def _escurecer_cor(self, cor):
        mapa = {
            self.COR_MRV: self.COR_MRV_HOVER,
            self.COR_AZUL: "#174EA6",
            self.COR_ROXO: "#5A45D6",
            self.COR_LARANJA: "#BA5A12",
            self.COR_AMARELO: "#B7950B",
            self.COR_VINHO: "#7D2440",
            self.COR_CINZA: "#4A5560"
        }
        return mapa.get(cor, self.COR_MRV_HOVER)

    # ==========================================================================
    # TEXTOS DE AJUDA
    # ==========================================================================
    def _texto_ajuda_primeiros_passos(self):
        return """PRIMEIROS PASSOS

Bem-vindo à Central de Automações MRV.

1. Acesse a aba Configurações.
2. Preencha suas credenciais.
3. Salve as credenciais.
4. Acesse a aba Robôs.
5. Escolha o processo desejado.
6. Acompanhe o progresso e os logs pela área de execução.

Antes de rodar qualquer robô, confirme se os arquivos necessários estão nas pastas corretas.
"""

    def _texto_ajuda_credenciais(self):
        return """CREDENCIAIS

As credenciais são usadas pelos robôs para acessar sistemas internos, portais e integrações.

Campos principais:

- E-mail MRV
- Senha MRV
- Senha Malote Web
- Chave API Agilis
- Dados dos Correios

Importante:
- Não compartilhe o arquivo config_mrv.json.
- Atualize as senhas sempre que trocar sua senha corporativa.
- Alguns robôs ainda podem solicitar MFA no celular.
"""

    def _texto_ajuda_correios(self):
        return """CORREIOS E FATURAMENTO

Robôs disponíveis:

- Rateio de Malote
- Rateio AGF
- Faturamento 1
- Faturamento 2
- Cobrança de boletos de contratos

Arquivos comuns:
- Planilha dos Correios
- Relatório Agilis
- Base de centro de custo
- Acompanhamento VSC
- PDFs e planilhas conforme o processo

Atenção:
- Confira nomes e formatos antes de executar.
- O Faturamento pode depender do Outlook aberto e logado.
"""

    def _texto_ajuda_agilis(self):
        return """AGILIS E CHAMADOS

Robôs disponíveis:

- Relatório de envio para Correios
- Produtividade
- Fechar Chamados a Vencer

Produtividade:
- Pode baixar relatórios do Podio, Agilis e Bússola.
- Durante extrações com SAP/Bússola, não mexa no mouse ou teclado.

Fechar chamados:
- Monitorar o dia todo: fecha chamados próximos do vencimento.
- Fechar todos de hoje: executa uma única rodada nos chamados do dia.
"""

    def _texto_ajuda_uber_sap(self):
        return """UBER, SAP E CONTRATOS

Ordem recomendada dos robôs Uber:

1. Uber 1: Atualizar Responsáveis
2. Uber 2: Gerar Relatórios e Pastas
3. Uber 3: Criar Rascunhos de E-mail

Arquivos esperados:
- EXPORT_*.xlsx
- Responsaveis Por Centro de Custos.xlsx
- Responsaveis_Atualizado_SAP.xlsx
- Relatório do mês
- Pasta ano,mês gerada pelo Uber 2

SAP/ZMM180:
- Deixe o SAP aberto na segunda tela.
- Deixe o documento aberto no Edge, quando aplicável.
- Não mexa no mouse ou teclado durante a automação.
"""

    def _texto_ajuda_cancelamento(self):
        return """CANCELAMENTO DE EMERGÊNCIA

Use o botão vermelho CANCELAR PROCESSO ATIVO quando:

- O robô travar.
- Você precisar usar o computador imediatamente.
- Uma planilha estiver aberta e bloqueando o processo.
- Um navegador abrir em uma etapa errada.

O cancelamento tenta encerrar o processo principal e subprocessos relacionados.
"""

    def _texto_ajuda_faq(self):
        return """FAQ

1. O robô não abriu?
Verifique se as credenciais estão preenchidas e se o antivírus não bloqueou o executável.

2. O robô não encontrou a planilha?
Confirme a pasta do processo e o nome esperado do arquivo.

3. O progresso não saiu do zero?
Nem todos os robôs enviam mensagens [PROGRESSO:]. Mesmo assim, os logs continuam aparecendo.

4. Posso usar o computador enquanto roda?
Depende. Robôs de navegador geralmente permitem. Robôs SAP/PyAutoGUI não permitem.

5. O histórico fica salvo onde?
No arquivo historico_execucoes.json, dentro da pasta base do projeto.
"""

    def _abrir_tutorial_api_agilis(self):
        janela = ctk.CTkToplevel(self.root)
        janela.title("Como gerar a Chave API do Agilis")
        janela.geometry("720x560")
        janela.minsize(650, 500)
        janela.attributes("-topmost", True)

        janela.grid_columnconfigure(0, weight=1)
        janela.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(janela, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=22, pady=(20, 10))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="🔑 Como gerar a Chave API do Agilis",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=self.COR_TEXTO
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            header,
            text="Use este passo a passo para configurar a chave usada pelos robôs Relatório Correios e Rateio AGF.",
            font=ctk.CTkFont(size=13),
            text_color=self.COR_TEXTO_FRACO,
            wraplength=650,
            justify="left"
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        corpo = ctk.CTkScrollableFrame(janela, fg_color=self.COR_CARD, corner_radius=12)
        corpo.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 14))

        texto = """
    PASSO A PASSO

    1. Acesse o Agilis normalmente pelo navegador.

    2. Na tela inicial do Agilis, clique no ícone do seu perfil no canto superior direito.

    3. No menu lateral que abrir, clique na opção "Gerar chave API".

    4. Na tela de geração da chave, selecione a opção "Nunca expira".

    5. Clique no botão "Gerar" ou "Regerar".

    Observação:
    Se for a primeira vez que você gera a chave, o botão pode aparecer com outro nome, como "Gerar". 
    Se a chave já existir, o botão costuma aparecer como "Regerar". O processo é o mesmo.

    6. Copie a chave gerada.

    7. Volte para o Hub Central MRV.

    8. Acesse a aba "Configurações".

    9. Cole a chave no campo "Chave API Agilis".

    10. Clique em "Salvar Todas as Credenciais".

    ROBÔS QUE DEPENDEM DESSA CHAVE

    - Gerar relatório de envio para Correios
    - Rateio AGF

    IMPORTANTE

    - Não compartilhe sua chave API.
    - Se esquecer a chave, gere uma nova no Agilis.
    - Sempre prefira a opção "Nunca expira", para evitar que os robôs parem de funcionar futuramente.
    """

        ctk.CTkLabel(
            corpo,
            text=texto.strip(),
            font=ctk.CTkFont(size=14),
            text_color=self.COR_TEXTO,
            justify="left",
            wraplength=640
        ).pack(anchor="w", padx=18, pady=18)

        botoes = ctk.CTkFrame(janela, fg_color="transparent")
        botoes.grid(row=2, column=0, sticky="ew", padx=22, pady=(0, 18))
        botoes.grid_columnconfigure((0, 1), weight=1)

        def ir_configuracoes():
            janela.destroy()
            self.selecionar_tela("config")

        ctk.CTkButton(
            botoes,
            text="Ir para Configurações",
            height=40,
            fg_color=self.COR_MRV,
            hover_color=self.COR_MRV_HOVER,
            font=ctk.CTkFont(weight="bold"),
            command=ir_configuracoes
        ).grid(row=0, column=0, sticky="ew", padx=(0, 6))

        ctk.CTkButton(
            botoes,
            text="Fechar",
            height=40,
            fg_color=self.COR_CINZA,
            hover_color="#4A5560",
            font=ctk.CTkFont(weight="bold"),
            command=janela.destroy
        ).grid(row=0, column=1, sticky="ew", padx=(6, 0))

class JanelaBusca:
    def __init__(self, parent, textbox):
        self.top = ctk.CTkToplevel(parent)
        self.top.title("Buscar na Ajuda")
        self.top.geometry("380x86")
        self.top.attributes("-topmost", True)
        self.top.resizable(False, False)

        try:
            self.textbox = textbox._textbox
        except Exception:
            self.textbox = textbox

        self.last_pos = "1.0"

        self.entry = ctk.CTkEntry(self.top, placeholder_text="Digite para buscar...")
        self.entry.pack(side="left", padx=10, pady=12, expand=True, fill="x")
        self.entry.bind("<Return>", self.buscar)

        self.btn = ctk.CTkButton(self.top, text="Próximo", width=90, command=self.buscar)
        self.btn.pack(side="right", padx=10, pady=12)

        self.entry.focus()

    def buscar(self, event=None):
        query = self.entry.get()

        try:
            self.textbox.tag_remove("highlight", "1.0", "end")
        except Exception:
            pass

        if not query:
            return

        pos = self.textbox.search(query, self.last_pos, stopindex="end", nocase=True)

        if not pos:
            pos = self.textbox.search(query, "1.0", stopindex="end", nocase=True)

        if pos:
            end_pos = f"{pos}+{len(query)}c"
            self.textbox.tag_add("highlight", pos, end_pos)
            self.textbox.tag_config("highlight", background="#FFC000", foreground="black")
            self.textbox.see(pos)
            self.last_pos = end_pos
        else:
            self.last_pos = "1.0"
            messagebox.showinfo("Busca", "Nenhum resultado encontrado.")


if __name__ == "__main__":
    root = ctk.CTk()
    app = CentralAutomacaoMRV(root)
    root.mainloop()
