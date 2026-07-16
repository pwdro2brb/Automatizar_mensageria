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
from tkinter import ttk, messagebox
import threading
import sys
import subprocess

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
        
        # Botão de Cancelar (Vermelho e chamativo)
        self.btn_cancelar = tk.Button(root, text="🛑 CANCELAR PROCESSO ATIVO", bg="#ff4d4d", fg="white", font=("Arial", 10, "bold"), command=self.cancelar_processo, state="disabled")
        self.btn_cancelar.pack(fill=tk.X, pady=(10, 5))

        self.root = root
        self.root.title("Hub Central de Automações - MRV")
        self.root.geometry("950x750")
        self.root.configure(padx=15, pady=15)

        style = ttk.Style()
        style.configure("TButton", font=("Arial", 9, "bold"), padding=5)
        style.configure("TLabelframe.Label", font=("Arial", 10, "bold"), foreground="#003366")

        lbl_titulo = ttk.Label(root, text="🤖 Central de Robôs - Administrativo MRV", font=("Arial", 16, "bold"))
        lbl_titulo.pack(pady=(0, 15))

        frame_botoes = tk.Frame(root)
        frame_botoes.pack(fill=tk.X, pady=5)

        # ======================================================================
        # MAPEAMENTO DOS BOTÕES
        # ======================================================================
        
        cmd_placeholder = "import time; print('Executando processo simulado...'); time.sleep(2); print('✅ Concluído!')"

        # --- CATEGORIA 1: CORREIOS & FATURAMENTO ---
        frame_correios = ttk.LabelFrame(frame_botoes, text="📦 Correios & Faturamento")
        frame_correios.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.btn_enc_dia = ttk.Button(frame_correios, text="Relatório Encomendas do Dia", 
            command=lambda: self.executar_processo_cancelavel("Relatório Encomendas do Dia", comando_python=cmd_placeholder))
        self.btn_enc_dia.pack(fill=tk.X, padx=10, pady=5)

        self.btn_rateio_malote = ttk.Button(frame_correios, text="Rateio de Malote (Centros de Custo)", 
            command=lambda: self._verificar_planilhas_e_executar("Rateio de Malote", "import robos.robo_rateio_malote as rrm; rrm.executar_rateio_malote()"))
        self.btn_rateio_malote.pack(fill=tk.X, padx=10, pady=5)

        self.btn_fat_1 = ttk.Button(frame_correios, text="Faturamento 1: Gerar Rascunhos", 
            command=lambda: self.executar_processo_cancelavel("Faturamento 1", comando_python="import robos.robo_faturamento as rf; rf.criar_rascunhos_correios()"))
        self.btn_fat_1.pack(fill=tk.X, padx=10, pady=5)

        self.btn_fat_2 = ttk.Button(frame_correios, text="Faturamento 2: Planilha Rateio Pag", 
            command=lambda: self._verificar_planilhas_e_executar("Faturamento 2", "import robos.robo_faturamento as rf; rf.preparar_e_gerar_rateio()"))
        self.btn_fat_2.pack(fill=tk.X, padx=10, pady=5)

        self.btn_fat_3 = ttk.Button(frame_correios, text="Faturamento 3: Lançar NF (Portal)", 
            command=lambda: self.executar_processo_cancelavel("Faturamento 3", comando_python="import robos.robo_faturamento as rf; rf.lancar_nota_fiscal()"))
        self.btn_fat_3.pack(fill=tk.X, padx=10, pady=5)

        # --- CATEGORIA 2: PODIO & MENSAGERIA ---
        frame_podio = ttk.LabelFrame(frame_botoes, text="🏢 Podio & Mensageria")
        frame_podio.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        self.btn_juridico = ttk.Button(frame_podio, text="Relatório Jurídico Montreal", command=self._chamar_robo_juridico)
        self.btn_juridico.pack(fill=tk.X, padx=10, pady=5)

        self.btn_incluir_podio = ttk.Button(frame_podio, text="Incluir Correspondências Rápidas", 
            command=self._chamar_robo_incluir_encomendas)
        self.btn_incluir_podio.pack(fill=tk.X, padx=10, pady=5)

        # --- CATEGORIA 3: AGILIS & PRODUTIVIDADE ---
        frame_agilis = ttk.LabelFrame(frame_botoes, text="🎧 Agilis & Chamados")
        frame_agilis.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        self.btn_produtividade = ttk.Button(frame_agilis, text="Gerar relatório de envio para Correios", 
            command=lambda: self.executar_processo_cancelavel("Relatório Correios", comando_python="import robos.robo_relatorio_correios as rc; rc.executar_relatorio_completo()"))
        self.btn_produtividade.pack(fill=tk.X, padx=10, pady=5)

        self.btn_produtividade_setor = ttk.Button(frame_agilis, text="Gerar Produtividade (Podio/Agilis/SAP)", 
            command=self._chamar_robo_produtividade)
        self.btn_produtividade_setor.pack(fill=tk.X, padx=10, pady=5)

        self.btn_fechar_chamados = ttk.Button(frame_agilis, text="Fechar Chamados a Vencer", 
            command=lambda: self.executar_processo_cancelavel("Fechar Chamados", comando_python="import robos.robo_fechar_chamados as rfc; rfc.executar_fechamento()"))
        self.btn_fechar_chamados.pack(fill=tk.X, padx=10, pady=5)

        # --- CATEGORIA 4: OUTROS SISTEMAS ---
        frame_outros = ttk.LabelFrame(frame_botoes, text="🚗 Outros (Uber / SAP)")
        frame_outros.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")

        self.btn_uber_1 = ttk.Button(frame_outros, text="Uber 1: Atualizar Responsáveis (SAP)", 
            command=lambda: self._verificar_planilhas_e_executar("Uber 1", "import robos.robo_uber_relatorios as ru; ru.etapa_1_atualizar_responsaveis()"))
        self.btn_uber_1.pack(fill=tk.X, padx=10, pady=2)

        self.btn_uber_2 = ttk.Button(frame_outros, text="Uber 2: Gerar Relatórios e Pastas", 
            command=lambda: self._verificar_planilhas_e_executar("Uber 2", "import robos.robo_uber_relatorios as ru; ru.etapa_2_gerar_relatorios()"))
        self.btn_uber_2.pack(fill=tk.X, padx=10, pady=2)

        self.btn_uber_3 = ttk.Button(frame_outros, text="Uber 3: Criar Rascunhos de E-mail", 
            command=lambda: self.executar_processo_cancelavel("Uber 3", comando_python="import robos.robo_uber_rascunhos as rr; rr.criar_rascunhos()"))
        self.btn_uber_3.pack(fill=tk.X, padx=10, pady=2)

        self.btn_zmm180 = ttk.Button(frame_outros, text="Faturamento Transação ZMM180", 
            command=lambda: self.executar_processo_cancelavel("Faturamento ZMM180", comando_python=cmd_placeholder))
        self.btn_zmm180.pack(fill=tk.X, padx=10, pady=5)

        frame_botoes.columnconfigure(0, weight=1)
        frame_botoes.columnconfigure(1, weight=1)

        # ======================================================================
        # CONSOLE DE LOGS
        # ======================================================================
        lbl_console = ttk.Label(root, text="Console de Execução em Tempo Real:", font=("Arial", 10, "bold"))
        lbl_console.pack(anchor=tk.W, pady=(15, 5))

        self.console = tk.Text(root, height=15, bg="#1e1e1e", fg="#00ff00", font=("Consolas", 10))
        self.console.pack(fill=tk.BOTH, expand=True)
        self.console.configure(state='disabled')

        sys.stdout = PrintRedirector(self.console)
        sys.stderr = PrintRedirector(self.console)

        self.todos_botoes = [
            self.btn_enc_dia, self.btn_rateio_malote, self.btn_fat_1, self.btn_fat_2, self.btn_fat_3,
            self.btn_juridico, self.btn_incluir_podio, self.btn_produtividade, self.btn_fechar_chamados,self.btn_produtividade_setor,
            self.btn_uber_1, self.btn_uber_2, self.btn_uber_3, self.btn_zmm180
        ]

        print("✅ Sistema Central iniciado com sucesso, Pedro!")
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
            f"⚠️ LEMBRETE ⚠️\n\nVocê já atualizou/trocou as planilhas na pasta para rodar o {nome_processo}?"
        )
        if resposta:
            self.executar_processo_cancelavel(nome_processo, comando_python=comando_python)

    def _chamar_robo_produtividade(self):
        resposta = messagebox.askokcancel(
            "Aviso Importante - SAP e Mouse",
            "⚠️ ATENÇÃO ⚠️\n\n"
            "1. Deixe o SAP aberto (após a tela de login) na SEGUNDA TELA.\n"
            "2. NÃO MEXA no mouse ou teclado durante o processo (principalmente na parte do Bússola).\n\n"
            "Deseja continuar?"
        )
        if resposta:
            self.executar_processo_cancelavel("Produtividade Setorial", comando_python="import robos.produtividade as rp; rp.executar_robo_produtividade_setor()")

    def _chamar_robo_incluir_encomendas(self):
        resposta = messagebox.askokcancel(
            "Lembrete - Correspondências",
            "⚠️ LEMBRETE ⚠️\n\n"
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
            btn.state(['disabled'])
            
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
            self.root.after(0, lambda: self.btn_cancelar.config(state="normal"))
            
            linhas_log = []

            try:
                for linha in processo.stdout:
                    print(linha, end="")
                    linhas_log.append(linha.rstrip('\r\n')) 
            except ValueError:
                pass 
                
            processo.wait() 
            
            self.processo_ativo = None
            self.root.after(0, lambda: self.btn_cancelar.config(state="disabled"))
            
            if self.foi_cancelado:
                print("\n⚠️ O processo foi cancelado pelo usuário.")
                self.root.after(0, lambda: messagebox.showwarning("Cancelado", "O processo foi cancelado forçadamente pelo usuário."))
            
            elif processo.returncode == 0:
                print("\n✅ Processo finalizado com sucesso!")
                self.root.after(0, lambda: messagebox.showinfo("Sucesso", "A automação foi concluída com sucesso!"))
            
            elif processo.returncode == 1:
                print(f"\n⚠️ O processo falhou (Código {processo.returncode}).")
                
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
                print(f"\n⚠️ O processo foi encerrado (Código {processo.returncode}).")
                self.root.after(0, lambda: messagebox.showwarning("Encerrado", "O processo foi encerrado."))
                
        except Exception as e:
            print(f"\n❌ Erro ao iniciar o processo: {e}")
            self.root.after(0, lambda msg=str(e): messagebox.showerror("Erro Crítico", msg))
        finally:
            print("-" * 60)
            self.root.after(0, self._reativar_botoes)

    def cancelar_processo(self):
        """Mata o processo ativo e todos os navegadores/janelas que ele abriu."""
        if self.processo_ativo and self.processo_ativo.poll() is None:
            resposta = messagebox.askyesno("Atenção", "Tem certeza que deseja cancelar o robô?\nIsso fechará os navegadores abertos por ele.")
            if resposta:
                try:
                    self.foi_cancelado = True 
                    subprocess.run(['taskkill', '/F', '/T', '/PID', str(self.processo_ativo.pid)], creationflags=subprocess.CREATE_NO_WINDOW)
                    
                    if self.processo_ativo.stdout:
                        self.processo_ativo.stdout.close()
                        
                    print("\n" + "="*50)
                    print("🛑 PROCESSO CANCELADO FORÇADAMENTE PELO USUÁRIO!")
                    print("="*50 + "\n")
                except Exception as e:
                    print(f"\n❌ Erro ao tentar cancelar: {e}")

    def _reativar_botoes(self):
        for btn in self.todos_botoes:
            btn.state(['!disabled'])

if __name__ == "__main__":
    root = tk.Tk()
    app = CentralAutomacaoMRV(root)
    root.mainloop()
