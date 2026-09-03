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
import re
import base64


# ==============================================================================
# 1. INTERCEPTADOR DE PROCESSOS
# ==============================================================================
if len(sys.argv) > 2 and sys.argv[1] == "--run-code":
    codigo = sys.argv[2]
    log_path = (
        sys.argv[3]
        if len(sys.argv) > 3
        else None
    )

    if log_path:
        sys.stdout = open(
            log_path,
            "w",
            encoding="utf-8",
            buffering=1
        )

        sys.stderr = sys.stdout

    try:
        print(
            "[SUBPROCESSO] Instância auxiliar iniciada.",
            flush=True
        )

        print(
            f"[SUBPROCESSO] PID: {os.getpid()}",
            flush=True
        )

        if os.environ.get("FATURAMENTO_PASTA"):
            print(
                "[SUBPROCESSO] Pasta recebida: "
                f"{os.environ['FATURAMENTO_PASTA']}",
                flush=True
            )

        print(
            "[SUBPROCESSO] Executando comando...",
            flush=True
        )

        exec(codigo)

        print(
            "[SUBPROCESSO] Comando concluído.",
            flush=True
        )

    except BaseException:
        traceback.print_exc()
        sys.stdout.flush()
        sys.exit(1)

    finally:
        try:
            sys.stdout.flush()
        except Exception:
            pass

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
    import robos.robo_rateio_uber_central
    import robos.robo_rateio_uber_tradicional
    import robos.atualizar_centro_custo_viagens

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
    def _tratar_erro_tkinter(
        self,
        tipo_erro,
        valor_erro,
        traceback_erro
    ):
        """
        Registra erros não tratados de callbacks do Tkinter.
        """

        texto_erro = "".join(
            traceback.format_exception(
                tipo_erro,
                valor_erro,
                traceback_erro
            )
        )

        try:
            caminho_log = os.path.join(
                self.PASTA_BASE,
                "erro_hub_tkinter.log"
            )

            with open(
                caminho_log,
                "a",
                encoding="utf-8"
            ) as arquivo:
                arquivo.write(
                    "\n" + "=" * 80 + "\n"
                )

                arquivo.write(
                    datetime.datetime.now().strftime(
                        "%d/%m/%Y %H:%M:%S"
                    )
                )

                arquivo.write("\n")
                arquivo.write(texto_erro)

        except Exception:
            pass

        print(texto_erro)

        try:
            messagebox.showerror(
                "Erro no Hub",
                "O Hub encontrou um erro na interface.\n\n"
                f"{valor_erro}\n\n"
                "O detalhe completo foi salvo em "
                "erro_hub_tkinter.log.",
                parent=self.root
            )
        except Exception:
            pass

    def __init__(self, root):
        self.root = root

        self.processo_ativo = None
        self.foi_cancelado = False
        self.janela_ajuda_robo = None
        self.todos_botoes = []
        self.tela_atual = None
        self.logs_visiveis = True
        self.robos_filtrados = []
        self._resize_after_id = None

        self.PASTA_BASE = getattr(config, "PASTA_PROJETO", os.path.dirname(os.path.abspath(__file__)))
        self.root.report_callback_exception = (
            self._tratar_erro_tkinter
        )
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
                "arquivo_ajuda": "robo_rateio_malote.md",
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
                "arquivo_ajuda": "robo_rateio_AGF.md"
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
                "arquivo_ajuda": "robo_faturamento.md",
                "secao_ajuda": "Faturamento 1: Gerar Rascunhos",
            },
            {
                "nome": "Faturamento Completo",
                "titulo": "Faturamento 2: Processo Completo",
                "categoria": "Correios & Faturamento",
                "icone": "💰",
                "cor": self.COR_ROXO,
                "tempo": "2 a 5 min",
                "Prioridade": "Alto",
                "requisitos": [
                    "Outlook ou pasta local",
                    "MRV Pag",
                    "Excel",
                    "PDF"
                ],
                "descricao": (
                    "Executa o faturamento pelo e-mail ou pelos arquivos "
                    "selecionados em uma pasta."
                ),
                "handler": self._chamar_robo_faturamento,
                "tipo": "especial",
                "arquivo_ajuda": "robo_faturamento.md",
                "secao_ajuda": "Faturamento 2: Processo Completo",
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
                "arquivo_ajuda": "robo_cobrar_boleto.md"
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
                "arquivo_ajuda": "robo_juridico.md"
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
                "arquivo_ajuda": "robo_incluir_encomendas.md"
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
                "arquivo_ajuda": "robo_relatorio_correios.md"
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
                "arquivo_ajuda": "produtividade.md"
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
                "arquivo_ajuda": "robo_fechar_chamados.md",
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
                "arquivo_ajuda": "robo_macro_contratos.md",
            },
            {
                "nome": "Rateio Uber Central",
                "requer_credenciais_uber": True,
                "titulo": "Rateio Uber Central",
                "categoria": "Outros (Uber / SAP / Contratos)",
                "icone": "🏢",
                "cor": self.COR_LARANJA,
                "tempo": "5 a 15 min",
                "requisitos": [
                    "Uber Business",
                    "Chrome",
                    "Excel",
                    "Rede"
                ],
                "descricao": (
                    "Baixa os relatórios da Uber Central, valida os centros "
                    "de custo e gera a planilha de rateio."
                ),
                "comando": (
                    "import robos.robo_rateio_uber_central as ruc; "
                    "ruc.executar_rateio_uber_central()"
                ),
                "tipo": "direto",
                "arquivo_ajuda": "robo_rateio_uber_central.md",
            },
            {
                "nome": "Rateio Uber Tradicional",
                "requer_credenciais_uber": True,
                "titulo": "Rateio Uber Tradicional",
                "categoria": "Outros (Uber / SAP / Contratos)",
                "icone": "🚕",
                "cor": self.COR_LARANJA,
                "tempo": "5 a 15 min",
                "requisitos": [
                    "Uber Business",
                    "Chrome",
                    "Excel",
                    "Rede"
                ],
                "descricao": (
                    "Baixa os relatórios da Uber Tradicional, valida os "
                    "centros de custo e gera a planilha de rateio."
                ),
                "comando": (
                    "import robos.robo_rateio_uber_tradicional as rut; "
                    "rut.executar_rateio_uber_tradicional()"
                ),
                "tipo": "direto",
                "arquivo_ajuda": "robo_rateio_uber_tradicional.md",
            },
            {
                "nome": "Responsáveis Viagens",
                "titulo": "Atualizar Responsáveis de Viagens",
                "categoria": "Outros (Uber / SAP / Contratos)",
                "icone": "🧳",
                "cor": self.COR_LARANJA,
                "tempo": "1 a 4 min",
                "Prioridade": "Média",
                "requisitos": [
                    "SAP",
                    "Excel",
                    "Base de Ativos",
                ],
                "descricao": (
                    "Atualiza o 1º aprovador e o superior do 1º aprovador "
                    "nos centros de custo de viagens."
                ),
                "comando": (
                    "import robos.atualizar_centro_custo_viagens as acv; "
                    "acv.atualizar_centro_custo_viagens()"
                ),
                "pasta": os.path.join(
                    config.PASTA_ARQUIVOS,
                    "uber",
                    "Responsaveis Viagens"
                ),
                "tipo": "pasta",
                "arquivo_ajuda": "atualizar_centro_custo_viagens.md",
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
                "arquivo_ajuda": "robo_uber_relatorios.md",
                "secao_ajuda": "Uber 1: Atualizar Responsáveis",
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
                "arquivo_ajuda": "robo_uber_relatorios.md",
                "secao_ajuda": "Uber 2: Gerar Relatórios e Pastas",
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
                "arquivo_ajuda": "criar_rascunhos_uber.md",
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
                "arquivo_ajuda": "robo_zmm180.md",
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

        btn_ajuda = ctk.CTkButton(
            topo,
            text="❓",
            width=32,
            height=28,
            fg_color=self.COR_CINZA,
            hover_color="#4A5560",
            command=lambda r=robo: self._abrir_ajuda_robo(r)
        )

        btn_ajuda.pack(side="right")

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

        chips = ctk.CTkFrame(info, fg_color="transparent")
        chips.pack(fill=tk.X, pady=(0, 6))

        self._criar_chip(
            chips,
            f"⏱ {robo.get('tempo', 'Variável')}",
            "#444444"
        ).pack(side="left", padx=(0, 5))

        requisitos = robo.get("requisitos", [])
        if requisitos:
            self._criar_chip(
                chips,
                " • ".join(requisitos[:3]),
                "#3D4852"
            ).pack(
                side="left",
                fill=tk.X,
                expand=True,
                padx=(0, 5)
            )

        btn = ctk.CTkButton(
            info,
            text="▶  Executar",
            height=32,
            fg_color=self.COR_MRV,
            hover_color=self.COR_MRV_HOVER,
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda r=robo: self._executar_robo(r)
        )
        btn.pack(fill=tk.X, expand=True)
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

        ctk.CTkButton(
            col_dir,
            text="Como gerar as chaves API do Podio?",
            height=32,
            fg_color=self.COR_CINZA,
            hover_color="#4A5560",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._abrir_tutorial_api_podio
        ).pack(anchor="w", padx=28, pady=(0, 12))

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
        ctk.CTkLabel(
            col_dir,
            text="UBER BUSINESS",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=self.COR_LARANJA
        ).pack(pady=(20, 12))

        self.entry_email_uber = self._campo_config(
            col_dir,
            "E-mail Uber:",
            getattr(config, "EMAIL_UBER", "")
        )

        self.entry_senha_uber = self._campo_config(
            col_dir,
            "Senha Uber:",
            getattr(config, "SENHA_UBER", ""),
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
        try:
            novo_email = self.entry_email.get().strip()
            novo_senha = self.entry_senha.get().strip()
            nova_senha_malote = self.entry_senha_malote.get().strip()
            nova_api_key_agilis = self.entry_API_KEY_AGILIS.get().strip()

            novo_correios_cod = self.entry_correios_cod.get().strip()
            novo_correios_email = self.entry_correios_email.get().strip()
            novo_correios_senha = self.entry_correios_senha.get().strip()

            novo_podio_id = self.entry_podio_client_id.get().strip()
            novo_podio_secret = self.entry_podio_client_secret.get().strip()
            novo_podio_app_id = self.entry_podio_app_id.get().strip()
            novo_podio_app_token = self.entry_podio_app_token.get().strip()

            novo_email_uber = self.entry_email_uber.get().strip()
            nova_senha_uber = self.entry_senha_uber.get().strip()

            print("Salvando credenciais...")

            config.salvar_credenciais(
                email=novo_email,
                senha=novo_senha,
                senha_malote=nova_senha_malote,
                chave_api_agilis=nova_api_key_agilis,
                correios_cod_adm=novo_correios_cod,
                correios_email=novo_correios_email,
                correios_senha=novo_correios_senha,
                podio_client_id=novo_podio_id,
                podio_client_secret=novo_podio_secret,
                podio_app_id=novo_podio_app_id,
                podio_app_token=novo_podio_app_token,
                email_uber=novo_email_uber,
                senha_uber=nova_senha_uber
            )

            # Atualiza as variáveis em memória.
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

            config.EMAIL_UBER = novo_email_uber
            config.SENHA_UBER = nova_senha_uber

            # Variáveis de compatibilidade.
            config.EMAIL_MRV = novo_email
            config.SENHA_MRV = novo_senha
            config.SENHA_MALOTE_MRV = nova_senha_malote
            config.API_KEY_AGILIS = nova_api_key_agilis

            self._salvar_metadata_config()

            self.lbl_status_credenciais.configure(
                text=self._texto_status_credenciais()
            )

            self.lbl_config_status.configure(
                text=self._texto_status_credenciais(detalhado=True)
            )

            print("Credenciais salvas com sucesso.")

            messagebox.showinfo(
                "Sucesso",
                "Todas as credenciais foram salvas com sucesso!",
                parent=self.root
            )

            self.selecionar_tela("robos")

        except TypeError as erro:
            traceback.print_exc()

            messagebox.showerror(
                "Erro de configuração",
                "A função config.salvar_credenciais não está compatível "
                "com os novos campos da Uber.\n\n"
                f"Detalhes:\n{erro}",
                parent=self.root
            )

        except Exception as erro:
            traceback.print_exc()

            messagebox.showerror(
                "Erro ao salvar",
                "Não foi possível salvar as credenciais.\n\n"
                f"Detalhes:\n{erro}",
                parent=self.root
            )

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
            "Chave API Podio": self._texto_ajuda_api_podio(),
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

    def _texto_ajuda_api_podio(self):
        return """CHAVE API PODIO

    Para que o robô de correspondências consiga inserir dados diretamente no Podio, é necessário configurar as chaves de API (Client ID e Client Secret).

    Como gerar as chaves no Podio:

    1. Acesse o link de configurações de API do Podio:
       https://podio.com/settings/api

    2. No campo "Nome do aplicativo (mostrado em atualizações)", preencha com o nome que desejar (ex: "Hub Central MRV").

    3. No campo "Domínio completo (sem protocolo) do seu URL de retorno (ex. mypodioapp.com)", você deve preencher com o domínio da sua empresa no Podio, sem o protocolo (https://) e sem barras adicionais.
       
       Exemplo prático:
       Se o link do seu aplicativo no Podio se parece com isto:
       https://podio.com/empresaaqui/processos/apps/teste
       
       Você deve pegar o "podio.com" e, sem espaçamento ou barras, adicionar o nome da sua empresa ("empresaaqui").
       O resultado final a ser preenchido deve ser exatamente:
       podio.comempresaaqui

    4. Clique em "Gerar chave de API" (Generate API Key).

    5. O Podio exibirá duas chaves:
       - Client ID
       - Client Secret

    6. Copie esses valores e cole-os nos respectivos campos na aba "Configurações" do Hub Central MRV.

    7. Lembre-se de preencher também o "Podio App ID" e o "Podio App Token" (que são obtidos diretamente nas configurações do próprio aplicativo/App no Podio) e clique em "Salvar Todas as Credenciais".
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

    def _obter_pasta_ajuda_robos(self):
        """
        Localiza a pasta ajuda_robos tanto no código-fonte
        quanto no executável gerado pelo PyInstaller.

        Estrutura esperada no desenvolvimento:

        AUTOMATIZAR_MENSAGERIA/
        ├── ajuda_robos/
        ├── robos/
        ├── dist/
        │   └── app_central.exe
        └── app_central.py
        """

        caminhos_possiveis = []

        if getattr(sys, "frozen", False):
            # Pasta em que está o executável:
            # AUTOMATIZAR_MENSAGERIA/dist/
            pasta_executavel = os.path.dirname(
                os.path.abspath(sys.executable)
            )

            # Raiz do projeto:
            # AUTOMATIZAR_MENSAGERIA/
            pasta_raiz_executavel = os.path.dirname(
                pasta_executavel
            )

            # Sua estrutura atual:
            # AUTOMATIZAR_MENSAGERIA/ajuda_robos/
            caminhos_possiveis.append(
                os.path.join(
                    pasta_raiz_executavel,
                    "ajuda_robos"
                )
            )

            # Estrutura alternativa para distribuição:
            # AUTOMATIZAR_MENSAGERIA/dist/ajuda_robos/
            caminhos_possiveis.append(
                os.path.join(
                    pasta_executavel,
                    "ajuda_robos"
                )
            )

            # Pasta temporária do PyInstaller, caso os arquivos
            # tenham sido incorporados ao executável.
            pasta_meipass = getattr(sys, "_MEIPASS", None)

            if pasta_meipass:
                caminhos_possiveis.append(
                    os.path.join(
                        pasta_meipass,
                        "ajuda_robos"
                    )
                )

        else:
            # Execução pelo código-fonte:
            # AUTOMATIZAR_MENSAGERIA/app_central.py
            pasta_codigo = os.path.dirname(
                os.path.abspath(__file__)
            )

            caminhos_possiveis.append(
                os.path.join(
                    pasta_codigo,
                    "ajuda_robos"
                )
            )

        # Também considera a pasta-base definida no config.py.
        pasta_base_config = getattr(
            config,
            "PASTA_PROJETO",
            None
        )

        if pasta_base_config:
            caminhos_possiveis.append(
                os.path.join(
                    pasta_base_config,
                    "ajuda_robos"
                )
            )

            # Caso PASTA_PROJETO aponte para dist.
            caminhos_possiveis.append(
                os.path.join(
                    os.path.dirname(pasta_base_config),
                    "ajuda_robos"
                )
            )

        # Remove caminhos duplicados sem alterar a ordem.
        caminhos_unicos = []

        for caminho in caminhos_possiveis:
            caminho_normalizado = os.path.normpath(caminho)

            if caminho_normalizado not in caminhos_unicos:
                caminhos_unicos.append(caminho_normalizado)

        # Retorna o primeiro caminho existente.
        for caminho in caminhos_unicos:
            if os.path.isdir(caminho):
                return caminho

        # Apresenta todos os locais verificados.
        caminhos_formatados = "\n".join(
            f"- {caminho}"
            for caminho in caminhos_unicos
        )

        raise FileNotFoundError(
            "A pasta de ajuda dos robôs não foi encontrada.\n\n"
            "Locais verificados:\n"
            f"{caminhos_formatados}"
        )

    def _carregar_arquivo_ajuda(self, nome_arquivo):
        """
        Carrega um arquivo Markdown da pasta ajuda_robos.
        """

        if not nome_arquivo:
            raise ValueError(
                "Nenhum arquivo de ajuda foi configurado para este robô."
            )

        pasta_ajuda = self._obter_pasta_ajuda_robos()

        caminho_arquivo = os.path.join(
            pasta_ajuda,
            nome_arquivo
        )

        if not os.path.isfile(caminho_arquivo):
            raise FileNotFoundError(
                "O arquivo de ajuda não foi encontrado:\n\n"
                f"{caminho_arquivo}"
            )

        with open(
            caminho_arquivo,
            "r",
            encoding="utf-8-sig"
        ) as arquivo:
            conteudo = arquivo.read()

        if not conteudo.strip():
            raise ValueError(
                "O arquivo de ajuda está vazio:\n\n"
                f"{caminho_arquivo}"
            )

        return conteudo

    def _extrair_secao_markdown(self, conteudo, nome_secao):
        """
        Extrai uma seção específica do Markdown.

        A seção começa em um cabeçalho que corresponda ao nome informado
        e termina no próximo cabeçalho de nível igual ou superior.
        """

        if not nome_secao:
            return conteudo

        linhas = conteudo.splitlines()

        indice_inicio = None
        nivel_inicio = None

        nome_procurado = nome_secao.strip().casefold()

        for indice, linha in enumerate(linhas):
            correspondencia = re.match(
                r"^\s*(#{1,6})\s+(.+?)\s*$",
                linha
            )

            if not correspondencia:
                continue

            marcadores = correspondencia.group(1)
            titulo_encontrado = correspondencia.group(2).strip()

            if titulo_encontrado.casefold() == nome_procurado:
                indice_inicio = indice
                nivel_inicio = len(marcadores)
                break

        if indice_inicio is None:
            raise ValueError(
                f"A seção de ajuda não foi encontrada:\n\n"
                f"{nome_secao}\n\n"
                f"Confira se o valor de 'secao_ajuda' é idêntico "
                f"ao título existente no arquivo Markdown."
            )

        indice_fim = len(linhas)

        for indice in range(indice_inicio + 1, len(linhas)):
            correspondencia = re.match(
                r"^\s*(#{1,6})\s+(.+?)\s*$",
                linhas[indice]
            )

            if not correspondencia:
                continue

            nivel_encontrado = len(correspondencia.group(1))

            if nivel_encontrado <= nivel_inicio:
                indice_fim = indice
                break

        secao = "\n".join(linhas[indice_inicio:indice_fim]).strip()

        return secao

    def _formatar_markdown_para_texto(self, conteudo):
        """
        Converte marcações básicas de Markdown em texto simples
        para exibição no CTkTextbox.
        """

        linhas_formatadas = []
        dentro_bloco_codigo = False

        for linha in conteudo.splitlines():
            linha_limpa = linha.rstrip()

            # Início ou fim de bloco de código Markdown.
            if linha_limpa.strip().startswith("```"):
                dentro_bloco_codigo = not dentro_bloco_codigo
                continue

            # Preserva a indentação de blocos de código.
            if dentro_bloco_codigo:
                linhas_formatadas.append(f"    {linha_limpa}")
                continue

            # Identifica títulos Markdown de nível 1 a 6.
            correspondencia_titulo = re.match(
                r"^\s*(#{1,6})\s+(.+?)\s*$",
                linha_limpa
            )

            if correspondencia_titulo:
                nivel = len(correspondencia_titulo.group(1))
                titulo = correspondencia_titulo.group(2).strip()

                # Remove marcações dentro dos títulos.
                titulo = re.sub(r"\*\*(.*?)\*\*", r"\1", titulo)
                titulo = re.sub(r"__(.*?)__", r"\1", titulo)
                titulo = re.sub(r"`(.*?)`", r"\1", titulo)

                if nivel == 1:
                    linhas_formatadas.append(titulo.upper())
                    linhas_formatadas.append("=" * len(titulo))
                else:
                    linhas_formatadas.append("")
                    linhas_formatadas.append(titulo)
                    linhas_formatadas.append("-" * len(titulo))

                continue

            # Remove negrito Markdown.
            linha_limpa = re.sub(r"\*\*(.*?)\*\*", r"\1", linha_limpa)
            linha_limpa = re.sub(r"__(.*?)__", r"\1", linha_limpa)

            # Remove itálico Markdown.
            linha_limpa = re.sub(
                r"(?<!\*)\*(?!\*)(.*?)\*(?!\*)",
                r"\1",
                linha_limpa
            )

            # Remove marcação de código na mesma linha.
            linha_limpa = re.sub(r"`(.*?)`", r"\1", linha_limpa)

            # Converte listas Markdown em marcadores visuais.
            linha_limpa = re.sub(
                r"^\s*[-*+]\s+",
                "• ",
                linha_limpa
            )

            linhas_formatadas.append(linha_limpa)

        texto = "\n".join(linhas_formatadas)

        # Impede quatro ou mais linhas vazias consecutivas.
        texto = re.sub(r"\n{4,}", "\n\n\n", texto)

        return texto.strip()
    
    def _abrir_ajuda_robo(self, robo):
        """
        Carrega e abre a ajuda contextual do robô selecionado.
        """

        try:
            nome_arquivo = robo.get("arquivo_ajuda")
            nome_secao = robo.get("secao_ajuda")

            conteudo = self._carregar_arquivo_ajuda(nome_arquivo)

            if nome_secao:
                conteudo = self._extrair_secao_markdown(
                    conteudo,
                    nome_secao
                )

            conteudo_formatado = self._formatar_markdown_para_texto(
                conteudo
            )

        except (FileNotFoundError, ValueError, OSError) as erro:
            messagebox.showerror(
                "Ajuda indisponível",
                str(erro),
                parent=self.root
            )
            return

        # Fecha a janela anterior, caso ainda esteja aberta.
        if self.janela_ajuda_robo is not None:
            try:
                if self.janela_ajuda_robo.winfo_exists():
                    self.janela_ajuda_robo.destroy()
            except Exception:
                pass

        janela = ctk.CTkToplevel(self.root)
        self.janela_ajuda_robo = janela

        titulo_robo = robo.get(
            "titulo",
            robo.get("nome", "Ajuda do robô")
        )

        janela.title(f"Ajuda - {titulo_robo}")
        janela.geometry("820x680")
        janela.minsize(650, 500)
        janela.configure(fg_color=self.COR_BG)

        # Mantém a janela vinculada ao Hub.
        janela.transient(self.root)

        # Coloca a janela na frente apenas no momento da abertura.
        janela.lift()
        janela.focus_force()

        janela.protocol(
            "WM_DELETE_WINDOW",
            self._fechar_ajuda_robo
        )

        # ==============================================================
        # CABEÇALHO
        # ==============================================================

        frame_cabecalho = ctk.CTkFrame(
            janela,
            fg_color=self.COR_CARD,
            corner_radius=0,
            height=90
        )
        frame_cabecalho.pack(
            fill=tk.X,
            padx=0,
            pady=0
        )
        frame_cabecalho.pack_propagate(False)

        icone = robo.get("icone", "❓")

        label_icone = ctk.CTkLabel(
            frame_cabecalho,
            text=icone,
            font=ctk.CTkFont(size=30)
        )
        label_icone.pack(
            side=tk.LEFT,
            padx=(24, 12),
            pady=20
        )

        frame_titulos = ctk.CTkFrame(
            frame_cabecalho,
            fg_color="transparent"
        )
        frame_titulos.pack(
            side=tk.LEFT,
            fill=tk.BOTH,
            expand=True,
            pady=14
        )

        label_titulo = ctk.CTkLabel(
            frame_titulos,
            text=titulo_robo,
            anchor="w",
            font=ctk.CTkFont(
                size=20,
                weight="bold"
            ),
            text_color=self.COR_TEXTO
        )
        label_titulo.pack(anchor="w")

        label_subtitulo = ctk.CTkLabel(
            frame_titulos,
            text="Ajuda específica desta automação",
            anchor="w",
            text_color=self.COR_TEXTO_FRACO,
            font=ctk.CTkFont(size=13)
        )
        label_subtitulo.pack(
            anchor="w",
            pady=(4, 0)
        )

        # ==============================================================
        # CONTEÚDO
        # ==============================================================

        frame_conteudo = ctk.CTkFrame(
            janela,
            fg_color="transparent"
        )
        frame_conteudo.pack(
            fill=tk.BOTH,
            expand=True,
            padx=20,
            pady=(20, 12)
        )

        caixa_texto = ctk.CTkTextbox(
            frame_conteudo,
            wrap="word",
            fg_color=self.COR_CARD,
            text_color=self.COR_TEXTO,
            corner_radius=12,
            border_width=1,
            border_color="#3A3A3A",
            font=ctk.CTkFont(
                family="Segoe UI",
                size=14
            )
        )
        caixa_texto.pack(
            fill=tk.BOTH,
            expand=True
        )

        caixa_texto.insert("1.0", conteudo_formatado)
        caixa_texto.configure(state="disabled")

        # ==============================================================
        # RODAPÉ
        # ==============================================================

        frame_rodape = ctk.CTkFrame(
            janela,
            fg_color="transparent"
        )
        frame_rodape.pack(
            fill=tk.X,
            padx=20,
            pady=(0, 18)
        )

        nome_exibido = nome_arquivo or "Ajuda não informada"

        label_arquivo = ctk.CTkLabel(
            frame_rodape,
            text=f"Documento: {nome_exibido}",
            text_color=self.COR_TEXTO_FRACO,
            font=ctk.CTkFont(size=11)
        )
        label_arquivo.pack(side=tk.LEFT)

        botao_fechar = ctk.CTkButton(
            frame_rodape,
            text="Fechar",
            width=110,
            height=36,
            fg_color=self.COR_MRV,
            hover_color=self.COR_MRV_HOVER,
            command=self._fechar_ajuda_robo
        )
        botao_fechar.pack(side=tk.RIGHT)

        # ==============================================================
        # CENTRALIZAÇÃO DA JANELA
        # ==============================================================

        janela.update_idletasks()

        largura = janela.winfo_width()
        altura = janela.winfo_height()

        posicao_x = (
            self.root.winfo_rootx()
            + (self.root.winfo_width() // 2)
            - (largura // 2)
        )

        posicao_y = (
            self.root.winfo_rooty()
            + (self.root.winfo_height() // 2)
            - (altura // 2)
        )

        # Evita posicionamento fora da tela.
        posicao_x = max(0, posicao_x)
        posicao_y = max(0, posicao_y)

        janela.geometry(
            f"{largura}x{altura}+{posicao_x}+{posicao_y}"
        )

    def _fechar_ajuda_robo(self):
        """
        Fecha a janela de ajuda contextual do robô.
        """

        if self.janela_ajuda_robo is not None:
            try:
                if self.janela_ajuda_robo.winfo_exists():
                    self.janela_ajuda_robo.destroy()
            except Exception:
                pass

        self.janela_ajuda_robo = None 
    
    # ==========================================================================
    # EXECUÇÃO DOS ROBÔS
    # ==========================================================================
    def _executar_robo(self, robo):
        if robo.get("requer_api_agilis") and not self._validar_api_agilis():
            return

        if (
            robo.get("requer_credenciais_uber")
            and not self._validar_credenciais_uber()
        ):
            return

        if robo.get("tipo") == "especial":
            handler = robo.get("handler")

            if not handler:
                messagebox.showerror(
                    "Erro",
                    f"O robô '{robo.get('nome')}' não possui "
                    "função de execução configurada.",
                    parent=self.root
                )
                return

            try:
                handler()

            except Exception as erro:
                traceback.print_exc()

                messagebox.showerror(
                    "Erro ao abrir o robô",
                    f"O robô '{robo.get('nome')}' encontrou "
                    "um erro antes de iniciar.\n\n"
                    f"Detalhes:\n{erro}",
                    parent=self.root
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

    def _selecionar_pasta_faturamento(self):
        """
        Abre o seletor do Windows em processo separado e envia
        a pasta escolhida ao robô de faturamento.
        """

        script_powershell = r"""
    Add-Type -AssemblyName System.Windows.Forms

    $dialogo = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialogo.Description = "Selecione a pasta com a planilha dos Correios, o Rateio Recebido e o boleto PDF"
    $dialogo.ShowNewFolderButton = $false

    $resultado = $dialogo.ShowDialog()

    if ($resultado -eq [System.Windows.Forms.DialogResult]::OK) {
        $caminho = $dialogo.SelectedPath
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($caminho)
        $base64 = [System.Convert]::ToBase64String($bytes)
        Write-Output $base64
    }
    """

        try:
            print(
                "Abrindo seletor de pasta do Windows..."
            )

            resultado = subprocess.run(
                [
                    "powershell.exe",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-STA",
                    "-Command",
                    script_powershell
                ],
                capture_output=True,
                creationflags=CREATE_NO_WINDOW
            )

        except Exception as erro:
            traceback.print_exc()

            messagebox.showerror(
                "Erro ao abrir seletor",
                "Não foi possível abrir o seletor de pasta.\n\n"
                f"Detalhes:\n{erro}",
                parent=self.root
            )
            return

        if resultado.returncode != 0:
            detalhe = resultado.stderr.decode(
                "utf-8",
                errors="replace"
            ).strip()

            messagebox.showerror(
                "Erro ao abrir seletor",
                "O seletor de pasta foi encerrado com erro.\n\n"
                f"Detalhes:\n{detalhe or 'Erro não identificado.'}",
                parent=self.root
            )
            return

        caminho_base64 = resultado.stdout.decode(
            "ascii",
            errors="ignore"
        ).strip()

        if not caminho_base64:
            print(
                "Seleção de pasta cancelada pelo usuário."
            )
            return

        try:
            pasta_selecionada = base64.b64decode(
                caminho_base64
            ).decode(
                "utf-8"
            )

        except Exception as erro:
            traceback.print_exc()

            messagebox.showerror(
                "Erro ao interpretar pasta",
                "O caminho selecionado não pôde ser interpretado.\n\n"
                f"Detalhes:\n{erro}",
                parent=self.root
            )
            return

        pasta_selecionada = os.path.normpath(
            pasta_selecionada.strip()
        )

        print(
            f"Pasta selecionada: {pasta_selecionada}"
        )

        if not os.path.isdir(pasta_selecionada):
            messagebox.showerror(
                "Pasta inválida",
                "A pasta selecionada não foi encontrada.\n\n"
                f"Caminho:\n{pasta_selecionada}",
                parent=self.root
            )
            return

        print(
            "Pasta recebida com sucesso."
        )

        comando_python = (
            "import os; "
            "import robos.robo_faturamento as rf; "
            "rf.executar_faturamento_por_pasta("
            "os.environ['FATURAMENTO_PASTA']"
            ")"
        )

        print(
            "Preparando execução do faturamento por pasta..."
        )

        self.executar_processo_cancelavel(
            "Faturamento 2 - Pasta selecionada",
            comando_python=comando_python,
            variaveis_ambiente={
                "FATURAMENTO_PASTA": pasta_selecionada
            }
        )

    def _chamar_robo_faturamento(self):
        """
        Permite escolher entre:
        - analisar os e-mails do Outlook;
        - selecionar uma pasta com os arquivos;
        - cancelar.
        """

        resposta = messagebox.askyesnocancel(
            "Faturamento 2",
            "Escolha como deseja executar o faturamento:\n\n"
            "SIM: analisar os e-mails do Outlook.\n\n"
            "NÃO: selecionar uma pasta com os arquivos.\n\n"
            "CANCELAR: não executar.",
            parent=self.root
        )

        if resposta is True:
            self.executar_processo_cancelavel(
                "Faturamento 2 - Analisar e-mail",
                comando_python=(
                    "import robos.robo_faturamento as rf; "
                    "rf.executar_faturamento_completo()"
                )
            )
            return

        if resposta is False:
            self.root.after_idle(
                self._selecionar_pasta_faturamento
            )
            return

        print(
            "Execução do Faturamento 2 cancelada."
        )

    def _validar_pasta_faturamento(
        self,
        pasta_selecionada
    ):
        """
        Verifica se a pasta contém exatamente:
        - uma planilha dos Correios com sete dígitos;
        - um arquivo Rateio Recebido em Excel;
        - um boleto em PDF.
        """

        erros = []

        if not pasta_selecionada:
            return [
                "Nenhuma pasta foi selecionada."
            ]

        if not os.path.exists(pasta_selecionada):
            return [
                "A pasta selecionada não existe."
            ]

        if not os.path.isdir(pasta_selecionada):
            return [
                "O caminho selecionado não é uma pasta."
            ]

        try:
            arquivos = [
                arquivo
                for arquivo in os.listdir(pasta_selecionada)
                if os.path.isfile(
                    os.path.join(
                        pasta_selecionada,
                        arquivo
                    )
                )
            ]

        except PermissionError:
            return [
                "O usuário não possui permissão para acessar a pasta."
            ]

        except OSError as erro:
            return [
                f"Não foi possível acessar a pasta: {erro}"
            ]

        planilhas_correios = [
            arquivo
            for arquivo in arquivos
            if re.fullmatch(
                r"\d{7}\.xlsx",
                arquivo,
                flags=re.IGNORECASE
            )
        ]

        rateios_recebidos = [
            arquivo
            for arquivo in arquivos
            if re.fullmatch(
                r"rateio[\s_-]*recebido\.xlsx",
                arquivo,
                flags=re.IGNORECASE
            )
        ]

        arquivos_pdf = [
            arquivo
            for arquivo in arquivos
            if arquivo.lower().endswith(".pdf")
        ]

        if not planilhas_correios:
            erros.append(
                "Planilha dos Correios no formato 1234567.xlsx"
            )

        elif len(planilhas_correios) > 1:
            erros.append(
                "Foi encontrada mais de uma planilha numérica "
                "dos Correios: "
                + ", ".join(planilhas_correios)
            )

        if not rateios_recebidos:
            erros.append(
                "Arquivo Rateio Recebido.xlsx"
            )

        elif len(rateios_recebidos) > 1:
            erros.append(
                "Foi encontrado mais de um Rateio Recebido: "
                + ", ".join(rateios_recebidos)
            )

        if not arquivos_pdf:
            erros.append(
                "Boleto em formato PDF"
            )

        elif len(arquivos_pdf) > 1:
            erros.append(
                "Foi encontrado mais de um PDF: "
                + ", ".join(arquivos_pdf)
            )

        return erros

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

    def executar_processo_cancelavel(self,nome_processo,comando_python=None,variaveis_ambiente=None):
        print(
            f"[HUB] Solicitação recebida: {nome_processo}"
        )

        if not comando_python:
            messagebox.showerror(
                "Erro",
                "Comando do robô não informado.",
                parent=self.root
            )
            return

        if (
            self.processo_ativo
            and self.processo_ativo.poll() is None
        ):
            messagebox.showwarning(
                "Processo em andamento",
                "Já existe uma automação em execução.",
                parent=self.root
            )
            return

        print(
            f"[HUB] Comando Python: {comando_python}"
        )

        self._preparar_tela_execucao()

        for btn in self.todos_botoes:
            try:
                btn.configure(state="disabled")
            except Exception:
                pass

        self.progressbar.set(0)

        self.btn_cancelar.configure(
            state="normal"
        )

        print(
            f">>> Iniciando: {nome_processo}..."
        )

        threading.Thread(
            target=self._rodar_subprocesso,
            args=(
                nome_processo,
                comando_python,
                variaveis_ambiente
            ),
            daemon=True
        ).start()

    def _rodar_subprocesso(self,nome_processo,comando_python,variaveis_ambiente=None):
        self.foi_cancelado = False

        fd, log_path = tempfile.mkstemp(
            suffix=".log",
            text=True
        )
        os.close(fd)

        inicio = time.time()

        print(
            f"Comando recebido: {comando_python}"
        )

        try:
            if getattr(sys, "frozen", False):
                cmd = [
                    sys.executable,
                    "--run-code",
                    comando_python,
                    log_path
                ]
            else:
                caminho_hub = os.path.abspath(
                    sys.argv[0]
                )

                cmd = [
                    sys.executable,
                    caminho_hub,
                    "--run-code",
                    comando_python,
                    log_path
                ]

            ambiente = os.environ.copy()

            if variaveis_ambiente:
                for chave, valor in variaveis_ambiente.items():
                    ambiente[str(chave)] = str(valor)

            print(
                f"Iniciando subprocesso: {cmd[0]}"
            )

            if "FATURAMENTO_PASTA" in ambiente:
                print(
                    "Pasta enviada ao subprocesso pela "
                    "variável FATURAMENTO_PASTA."
                )

            processo = subprocess.Popen(
                cmd,
                creationflags=CREATE_NO_WINDOW,
                env=ambiente
            )

            self.processo_ativo = processo

            self.root.after(
                0,
                lambda: self.btn_cancelar.configure(
                    state="normal"
                )
            )

            print(
                f"Subprocesso iniciado. PID: {processo.pid}"
            )

            linhas_log = []
            posicao_log = 0

            while processo.poll() is None:
                try:
                    with open(
                        log_path,
                        "r",
                        encoding="utf-8",
                        errors="replace"
                    ) as arquivo_log:
                        arquivo_log.seek(
                            posicao_log
                        )

                        novas_linhas = (
                            arquivo_log.readlines()
                        )

                        posicao_log = (
                            arquivo_log.tell()
                        )

                    for linha in novas_linhas:
                        self._processar_linha_log(
                            linha,
                            linhas_log
                        )

                except FileNotFoundError:
                    pass

                except Exception as erro_log:
                    print(
                        f"Erro ao acompanhar log: {erro_log}"
                    )

                time.sleep(0.15)

            # Lê as linhas restantes depois que o processo termina.
            try:
                with open(
                    log_path,
                    "r",
                    encoding="utf-8",
                    errors="replace"
                ) as arquivo_log:
                    arquivo_log.seek(
                        posicao_log
                    )

                    for linha in arquivo_log.readlines():
                        self._processar_linha_log(
                            linha,
                            linhas_log
                        )

            except Exception:
                pass

            codigo_retorno = processo.returncode

            self.processo_ativo = None

            self.root.after(
                0,
                lambda: self.btn_cancelar.configure(
                    state="disabled"
                )
            )

            duracao = round(
                time.time() - inicio,
                1
            )

            if self.foi_cancelado:
                print(
                    "\nO processo foi cancelado pelo usuário."
                )

                self._registrar_historico(
                    nome_processo,
                    "cancelado",
                    duracao
                )

                self.root.after(
                    0,
                    lambda: messagebox.showwarning(
                        "Cancelado",
                        "O processo foi cancelado pelo usuário.",
                        parent=self.root
                    )
                )

            elif codigo_retorno == 0:
                self.root.after(
                    0,
                    lambda: self.progressbar.set(1)
                )

                print(
                    "\nProcesso finalizado com sucesso!"
                )

                self._registrar_historico(
                    nome_processo,
                    "sucesso",
                    duracao
                )

                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Sucesso",
                        "A automação foi concluída com sucesso!",
                        parent=self.root
                    )
                )

            else:
                print(
                    f"\nO processo falhou. Código {codigo_retorno}."
                )

                self._registrar_historico(
                    nome_processo,
                    "erro",
                    duracao
                )

                texto_erro = self._extrair_erro(
                    linhas_log
                )

                mensagem_popup = (
                    "O processo foi interrompido pelo "
                    "seguinte motivo:\n\n"
                    f"{texto_erro}"
                )

                self.root.after(
                    0,
                    lambda mensagem=mensagem_popup: (
                        messagebox.showerror(
                            "Erro na Automação",
                            mensagem,
                            parent=self.root
                        )
                    )
                )

        except Exception as erro:
            self.processo_ativo = None

            print(
                f"\nErro ao iniciar o processo: {erro}"
            )

            traceback.print_exc()

            self._registrar_historico(
                nome_processo,
                "erro",
                0
            )

            self.root.after(
                0,
                lambda mensagem=str(erro): (
                    messagebox.showerror(
                        "Erro Crítico",
                        mensagem,
                        parent=self.root
                    )
                )
            )

        finally:
            print("-" * 60)

            self.root.after(
                0,
                self._reativar_botoes
            )

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
        processo = self.processo_ativo

        if not processo:
            messagebox.showinfo(
                "Cancelamento",
                "Não existe um processo ativo.",
                parent=self.root
            )
            return

        if processo.poll() is not None:
            self.processo_ativo = None

            messagebox.showinfo(
                "Cancelamento",
                "O processo já foi encerrado.",
                parent=self.root
            )
            return

        confirmar = messagebox.askyesno(
            "Atenção",
            "Tem certeza que deseja cancelar o robô?\n\n"
            "O processo ativo e os subprocessos relacionados "
            "serão encerrados.",
            parent=self.root
        )

        if not confirmar:
            return

        self.foi_cancelado = True

        print(
            f"\nCancelando processo PID {processo.pid}..."
        )

        try:
            resultado = subprocess.run(
                [
                    "taskkill",
                    "/F",
                    "/T",
                    "/PID",
                    str(processo.pid)
                ],
                capture_output=True,
                text=True,
                encoding="cp850",
                errors="replace",
                creationflags=CREATE_NO_WINDOW,
                timeout=15
            )

            print(
                resultado.stdout.strip()
                or resultado.stderr.strip()
                or "Comando de cancelamento enviado."
            )

        except subprocess.TimeoutExpired:
            try:
                processo.kill()
            except Exception:
                pass

            print(
                "O taskkill excedeu o tempo limite. "
                "O processo principal foi encerrado diretamente."
            )

        except Exception as erro:
            print(
                f"Erro ao executar taskkill: {erro}"
            )

            try:
                processo.kill()
            except Exception as erro_kill:
                print(
                    "Também não foi possível executar kill: "
                    f"{erro_kill}"
                )

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

    def _validar_credenciais_uber(self):
        email_uber = getattr(config, "EMAIL_UBER", "").strip()
        senha_uber = getattr(config, "SENHA_UBER", "").strip()

        campos_ausentes = []

        if not email_uber:
            campos_ausentes.append("E-mail Uber")

        if not senha_uber:
            campos_ausentes.append("Senha Uber")

        if campos_ausentes:
            messagebox.showwarning(
                "Credenciais Uber ausentes",
                "Os seguintes campos precisam ser configurados:\n\n"
                + "\n".join(f"• {campo}" for campo in campos_ausentes)
                + "\n\nAcesse a aba Configurações antes de continuar."
            )

            self.selecionar_tela("config")
            return False

        return True


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

    def _abrir_tutorial_api_podio(self):
        janela = ctk.CTkToplevel(self.root)
        janela.title("Como gerar as chaves API do Podio")
        janela.geometry("720x600")
        janela.minsize(650, 500)
        janela.attributes("-topmost", True)

        janela.grid_columnconfigure(0, weight=1)
        janela.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(janela, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=22, pady=(20, 10))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="🔑 Como gerar as chaves API do Podio",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=self.COR_TEXTO
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            header,
            text="Siga este passo a passo para obter o Client ID e o Client Secret necessários para a integração.",
            font=ctk.CTkFont(size=13),
            text_color=self.COR_TEXTO_FRACO,
            wraplength=650,
            justify="left"
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))

        corpo = ctk.CTkScrollableFrame(janela, fg_color=self.COR_CARD, corner_radius=12)
        corpo.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 14))

        texto = """
    PASSO A PASSO

    1. Acesse o link de configurações de API do Podio pelo seu navegador:
       https://podio.com/settings/api

    2. No campo "Nome do aplicativo (mostrado em atualizações)", digite o nome que preferir (ex: Hub Central MRV).

    3. No campo "Domínio completo (sem protocolo) do seu URL de retorno (ex. mypodioapp.com)", você deve preencher com o domínio da sua empresa no Podio, sem o protocolo (https://) e sem barras adicionais.

       Como estruturar o domínio:
       Se o link do seu aplicativo no Podio se parece com isto:
       https://podio.com/empresaaqui/processos/apps/teste
       
       Você deve pegar o "podio.com" e, sem espaçamento ou barras, adicionar o nome da sua empresa ("empresaaqui").
       O resultado final a ser preenchido deve ser exatamente:
       podio.comempresaaqui

    4. Clique em "Gerar chave de API" (Generate API Key).

    5. O Podio exibirá duas chaves na tela:
       - Client ID
       - Client Secret

    6. Copie esses valores.

    7. Volte para o Hub Central MRV.

    8. Acesse a aba "Configurações".

    9. Cole as chaves nos campos "Podio Client ID" e "Podio Client Secret".

    10. Clique em "Salvar Todas as Credenciais".
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
