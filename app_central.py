import ctypes

# ==============================================================================
# 🚀 FORÇAR O WINDOWS A RECONHECER TODOS OS MONITORES (ANTES DO TKINTER NASCER)
# ==============================================================================
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass
# ==============================================================================

import os
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import threading
import sys
import subprocess

# Configuração do Tema Moderno
ctk.set_appearance_mode("Dark")  # Pode ser "Light", "Dark" ou "System"
ctk.set_default_color_theme("green")

class PrintRedirector:
    """Redireciona os prints do terminal para a caixa de texto da interface de forma segura (Thread-Safe)."""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, text):
        self.text_widget.after(0, self._inserir_texto, text)

    def _inserir_texto(self, text):
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, text)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')

    def flush(self):
        pass

class CentralAutomacaoMRV:
    def __init__(self, root):
        self.processo_ativo = None 
        self.foi_cancelado = False 
        self.todos_botoes = []

        # Cores MRV
        self.COR_MRV = "#008542"          # Verde MRV
        self.COR_MRV_HOVER = "#006331"    # Verde MRV (Mais escuro para o hover)
        self.COR_CANCELAR = "#E74C3C"     # Vermelho
        self.COR_CANCELAR_HOVER = "#C0392B"

        self.root = root
        self.root.title("Hub Central de Automações - MRV")
        self.root.geometry("1050x850")
        
        # Título Principal (Emoji removido para evitar desalinhamento)
        lbl_titulo = ctk.CTkLabel(root, text="Central de Robôs - Administrativo MRV", 
                                  font=ctk.CTkFont(size=24, weight="bold"), text_color=self.COR_MRV)
        lbl_titulo.pack(pady=(20, 15))

        # Frame principal que vai segurar as 4 categorias
        frame_botoes = ctk.CTkFrame(root, fg_color="transparent")
        frame_botoes.pack(fill=tk.BOTH, expand=False, padx=20)
        
        frame_botoes.columnconfigure(0, weight=1)
        frame_botoes.columnconfigure(1, weight=1)

        # ======================================================================
        # FUNÇÕES AUXILIARES PARA CRIAR O DESIGN
        # ======================================================================
        def criar_quadro(parent, titulo, row, col):
            frm = ctk.CTkFrame(parent)
            frm.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
            # Emoji removido do título do quadro para evitar desalinhamento
            lbl = ctk.CTkLabel(frm, text=titulo, font=ctk.CTkFont(size=16, weight="bold"))
            lbl.pack(pady=(10, 10))
            return frm

        def criar_botao(parent, texto, comando, espaco_extra=False):
            pady_val = (15, 5) if espaco_extra else 5
            btn = ctk.CTkButton(parent, text=texto, command=comando, 
                                fg_color=self.COR_MRV, hover_color=self.COR_MRV_HOVER,
                                font=ctk.CTkFont(size=13, weight="bold"), height=35)
            btn.pack(fill=tk.X, padx=20, pady=pady_val)
            self.todos_botoes.append(btn)
            return btn

        cmd_placeholder = "import time; print('Executando processo simulado...'); time.sleep(2); print('Concluído!')"

        # --- CATEGORIA 1: CORREIOS & FATURAMENTO ---
        frame_correios = criar_quadro(frame_botoes, "Correios & Faturamento", 0, 0)
        
        criar_botao(frame_correios, "Relatório Encomendas do Dia", 
                    lambda: self.executar_processo_cancelavel("Relatório Encomendas do Dia", comando_python=cmd_placeholder))
        
        criar_botao(frame_correios, "Rateio de Malote (Centros de Custo)", 
                    lambda: self._verificar_planilhas_e_executar("Rateio de Malote", "import robos.robo_rateio_malote as rrm; rrm.executar_rateio_malote()"))
        
        # Espaço extra para agrupar o faturamento
        criar_botao(frame_correios, "Faturamento 1: Gerar Rascunhos", 
                    lambda: self.executar_processo_cancelavel("Faturamento 1", comando_python="import robos.robo_faturamento as rf; rf.criar_rascunhos_correios()"), espaco_extra=True)
        
        criar_botao(frame_correios, "Faturamento 2: Planilha Rateio Pag", 
                    lambda: self._verificar_planilhas_e_executar("Faturamento 2", "import robos.robo_faturamento as rf; rf.preparar_e_gerar_rateio()"))
        
        criar_botao(frame_correios, "Faturamento 3: Lançar NF (Portal)", 
                    lambda: self.executar_processo_cancelavel("Faturamento 3", comando_python="import robos.robo_faturamento as rf; rf.lancar_nota_fiscal()"))

        # --- CATEGORIA 2: PODIO & MENSAGERIA ---
        frame_podio = criar_quadro(frame_botoes, "Podio & Mensageria", 0, 1)
        
        criar_botao(frame_podio, "Relatório Jurídico Montreal", self._chamar_robo_juridico)
        
        criar_botao(frame_podio, "Incluir Correspondências Rápidas", self._chamar_robo_incluir_encomendas)

        # --- CATEGORIA 3: AGILIS & CHAMADOS ---
        frame_agilis = criar_quadro(frame_botoes, "Agilis & Chamados", 1, 0)
        
        criar_botao(frame_agilis, "Gerar relatório de envio para Correios", 
                    lambda: self.executar_processo_cancelavel("Relatório Correios", comando_python="import robos.robo_relatorio_correios as rc; rc.executar_relatorio_completo()"))
        
        criar_botao(frame_agilis, "Gerar Produtividade (Podio/Agilis/SAP)", self._chamar_robo_produtividade)
        
        criar_botao(frame_agilis, "Fechar Chamados a Vencer", 
                    lambda: self.executar_processo_cancelavel("Fechar Chamados", comando_python="import robos.robo_fechar_chamados as rfc; rfc.executar_fechamento()"))

        # --- CATEGORIA 4: OUTROS SISTEMAS ---
        frame_outros = criar_quadro(frame_botoes, "Outros (Uber / SAP)", 1, 1)
        
        criar_botao(frame_outros, "Uber 1: Atualizar Responsáveis (SAP)", 
                    lambda: self._verificar_planilhas_e_executar("Uber 1", "import robos.robo_uber_relatorios as ru; ru.etapa_1_atualizar_responsaveis()"))

        criar_botao(frame_outros, "Uber 2: Gerar Relatórios e Pastas", 
                    lambda: self._verificar_planilhas_e_executar("Uber 2", "import robos.robo_uber_relatorios as ru; ru.etapa_2_gerar_relatorios()"))

        criar_botao(frame_outros, "Uber 3: Criar Rascunhos de E-mail", 
                    lambda: self.executar_processo_cancelavel("Uber 3", comando_python="import robos.robo_uber_rascunhos as rr; rr.criar_rascunhos()"))
        
        criar_botao(frame_outros, "Faturamento Transação ZMM180", 
                    lambda: self.executar_processo_cancelavel("Faturamento ZMM180", comando_python=cmd_placeholder), espaco_extra=True)

        # ======================================================================
        # BOTÃO CANCELAR (MOVIDO PARA BAIXO)
        # ======================================================================
        self.btn_cancelar = ctk.CTkButton(root, text="CANCELAR PROCESSO ATIVO", 
                                          fg_color=self.COR_CANCELAR, hover_color=self.COR_CANCELAR_HOVER, 
                                          font=ctk.CTkFont(size=14, weight="bold"), height=40,
                                          command=self.cancelar_processo, state="disabled")
        self.btn_cancelar.pack(fill=tk.X, padx=30, pady=(20, 10))

        # ======================================================================
        # CONSOLE DE LOGS
        # ======================================================================
        lbl_console = ctk.CTkLabel(root, text="Console de Execução em Tempo Real:", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_console.pack(anchor=tk.W, padx=30, pady=(0, 5))

        self.console = ctk.CTkTextbox(root, height=200, font=ctk.CTkFont(family="Consolas", size=13), 
                                      text_color="#00FF00", fg_color="#1E1E1E", border_width=2, border_color="#333333")
        self.console.pack(fill=tk.BOTH, expand=True, padx=30, pady=(0, 20))
        self.console.configure(state='disabled')

        sys.stdout = PrintRedirector(self.console)
        sys.stderr = PrintRedirector(self.console)

        print("Sistema Central iniciado com sucesso, Pedro!")
        print("Selecione o processo que deseja executar.\n" + "-"*60)

    # ======================================================================
    # FUNÇÕES DE AVISO E VALIDAÇÃO (POKA-YOKE)
    # ======================================================================
    def _chamar_robo_juridico(self):
        resposta = messagebox.askyesnocancel(
            "Relatório Jurídico Montreal",
            "Você deseja que o robô baixe a planilha do Podio automaticamente?\n\n"
            "SIM: O robô fará tudo (Baixar, Formatar e Criar E-mail).\n"
            "NÃO: Eu já baixei manualmente (Apenas Formatar e Criar E-mail).\n"
            "CANCELAR: Abortar operação."
        )
        
        if resposta is True: 
            self.executar_processo_cancelavel("Relatório Jurídico Montreal", comando_python="import robos.robo_juridico as rj; rj.executar_juridico(pular_download=False)")
        elif resposta is False: 
            self.executar_processo_cancelavel("Relatório Jurídico (Apenas Formatação)", comando_python="import robos.robo_juridico as rj; rj.executar_juridico(pular_download=True)")

    def _verificar_planilhas_e_executar(self, nome_processo, comando_python):
        resposta = messagebox.askyesno(
            "Lembrete de Arquivos",
            f"LEMBRETE\n\nVocê já atualizou/trocou as planilhas na pasta para rodar o {nome_processo}?"
        )
        if resposta:
            self.executar_processo_cancelavel(nome_processo, comando_python=comando_python)

    def _chamar_robo_produtividade(self):
        resposta = messagebox.askokcancel(
            "Aviso Importante - SAP e Mouse",
            "ATENÇÃO\n\n"
            "1. Deixe o SAP aberto (após a tela de login) na SEGUNDA TELA.\n"
            "2. NÃO MEXA no mouse ou teclado durante o processo (principalmente na parte do Bússola).\n\n"
            "Deseja continuar?"
        )
        if resposta:
            self.executar_processo_cancelavel("Produtividade Setorial", comando_python="import robos.produtividade as rp; rp.executar_robo_produtividade_setor()")

    def _chamar_robo_incluir_encomendas(self):
        resposta = messagebox.askokcancel(
            "Lembrete - Correspondências",
            "LEMBRETE\n\n"
            "Você lembrou de apagar os dados antigos e preencher com os novos na planilha?\n\n"
            "Clique em OK para continuar ou Cancelar para verificar."
        )
        if resposta:
            self.executar_processo_cancelavel("Incluir Correspondências", comando_python="import robos.robo_incluir_encomendas as rie; rie.executar_inclusao()")

    # ======================================================================
    # MOTOR UNIVERSAL DE PROCESSOS ISOLADOS (CANCELÁVEIS)
    # ======================================================================
    def executar_processo_cancelavel(self, nome_processo, comando_python=None, script_path=None):
        for btn in self.todos_botoes:
            btn.configure(state="disabled")
            
        print(f">>> Iniciando: {nome_processo}...")
        threading.Thread(target=self._rodar_subprocesso, args=(comando_python, script_path), daemon=True).start()

    def _rodar_subprocesso(self, comando_python, script_path):
        self.foi_cancelado = False 
        
        try:
            if script_path:
                cmd = [sys.executable, "-X", "utf8", script_path]
            else:
                cmd = [sys.executable, "-X", "utf8", "-c", comando_python]

            processo = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                bufsize=1,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            self.processo_ativo = processo
            self.root.after(0, lambda: self.btn_cancelar.configure(state="normal"))
            
            linhas_log = []

            try:
                for linha in processo.stdout:
                    print(linha, end="")
                    linhas_log.append(linha.rstrip('\r\n')) 
            except ValueError:
                pass 
                
            processo.wait() 
            
            self.processo_ativo = None
            self.root.after(0, lambda: self.btn_cancelar.configure(state="disabled"))
            
            if self.foi_cancelado:
                print("\nO processo foi cancelado pelo usuário.")
                self.root.after(0, lambda: messagebox.showwarning("Cancelado", "O processo foi cancelado forçadamente pelo usuário."))
            
            elif processo.returncode == 0:
                print("\nProcesso finalizado com sucesso!")
                self.root.after(0, lambda: messagebox.showinfo("Sucesso", "A automação foi concluída com sucesso!"))
            
            elif processo.returncode == 1:
                print(f"\nO processo falhou (Código {processo.returncode}).")
                
                linhas_erro = [l for l in linhas_log if l.strip()]
                texto_erro = ""
                
                for i in range(len(linhas_erro)):
                    if "Traceback (most recent call last):" in linhas_erro[i]:
                        for j in range(i + 1, len(linhas_erro)):
                            if not linhas_erro[j].startswith(" "):
                                texto_erro = "\n".join(linhas_erro[j:])
                                linhas_texto = texto_erro.split("\n")
                                if ":" in linhas_texto[0]:
                                    msg_limpa = linhas_texto[0].split(":", 1)[1].strip()
                                    resto = "\n".join(linhas_texto[1:])
                                    texto_erro = msg_limpa + ("\n" + resto if resto else "")
                                break
                        break
                
                if not texto_erro:
                    texto_erro = linhas_erro[-1] if linhas_erro else "Erro desconhecido."
                
                mensagem_popup = f"O processo foi interrompido pelo seguinte motivo:\n\n{texto_erro}"
                self.root.after(0, lambda: messagebox.showerror("Erro na Automação", mensagem_popup))
            
            else:
                print(f"\nO processo foi encerrado (Código {processo.returncode}).")
                self.root.after(0, lambda: messagebox.showwarning("Encerrado", "O processo foi encerrado."))
                
        except Exception as e:
            print(f"\nErro ao iniciar o processo: {e}")
            self.root.after(0, lambda msg=str(e): messagebox.showerror("Erro Crítico", msg))
        finally:
            print("-" * 60)
            self.root.after(0, self._reativar_botoes)

    def cancelar_processo(self):
        if self.processo_ativo and self.processo_ativo.poll() is None:
            resposta = messagebox.askyesno("Atenção", "Tem certeza que deseja cancelar o robô?\nIsso fechará os navegadores abertos por ele.")
            if resposta:
                try:
                    self.foi_cancelado = True 
                    subprocess.run(['taskkill', '/F', '/T', '/PID', str(self.processo_ativo.pid)], creationflags=subprocess.CREATE_NO_WINDOW)
                    
                    if self.processo_ativo.stdout:
                        self.processo_ativo.stdout.close()
                        
                    print("\n" + "="*50)
                    print("PROCESSO CANCELADO FORÇADAMENTE PELO USUÁRIO!")
                    print("="*50 + "\n")
                except Exception as e:
                    print(f"\nErro ao tentar cancelar: {e}")

    def _reativar_botoes(self):
        for btn in self.todos_botoes:
            btn.configure(state="normal")

if __name__ == "__main__":
    root = ctk.CTk() 
    app = CentralAutomacaoMRV(root)
    root.mainloop()
