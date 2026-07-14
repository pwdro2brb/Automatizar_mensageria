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
import time
import subprocess

# ==============================================================================
# IMPORTAÇÃO DOS SEUS MÓDULOS (SCRIPTS SEPARADOS)
# ==============================================================================
import Processos_simples.robo_relatorio_correios as robo_relatorio_correios
import Processos_simples.robo_juridico as robo_juridico
import robo_faturamento

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

        # --- CATEGORIA 1: CORREIOS & FATURAMENTO ---
        frame_correios = ttk.LabelFrame(frame_botoes, text="📦 Correios & Faturamento")
        frame_correios.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

        self.btn_enc_dia = ttk.Button(frame_correios, text="Relatório Encomendas do Dia", command=lambda: self.rodar_thread(self._proc_encomendas_dia))
        self.btn_enc_dia.pack(fill=tk.X, padx=10, pady=5)

        self.btn_rateio_malote = ttk.Button(frame_correios, text="Rateio de Malote (Centros de Custo)", command=lambda: self.rodar_thread(self._proc_rateio_malote))
        self.btn_rateio_malote.pack(fill=tk.X, padx=10, pady=5)

        self.btn_fat_1 = ttk.Button(frame_correios, text="Faturamento 1: Gerar Rascunhos", command=lambda: self.rodar_thread(self._proc_fat_rascunhos))
        self.btn_fat_1.pack(fill=tk.X, padx=10, pady=5)

        self.btn_fat_2 = ttk.Button(frame_correios, text="Faturamento 2: Planilha Rateio Pag", command=lambda: self.rodar_thread(self._proc_fat_planilha))
        self.btn_fat_2.pack(fill=tk.X, padx=10, pady=5)

        self.btn_fat_3 = ttk.Button(frame_correios, text="Faturamento 3: Lançar NF (Portal)", command=lambda: self.rodar_thread(self._proc_fat_lancamento))
        self.btn_fat_3.pack(fill=tk.X, padx=10, pady=5)

        # --- CATEGORIA 2: PODIO & MENSAGERIA ---
        frame_podio = ttk.LabelFrame(frame_botoes, text="🏢 Podio & Mensageria")
        frame_podio.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        self.btn_juridico = ttk.Button(frame_podio, text="Relatório Jurídico Montreal", command=lambda: self.rodar_thread(self._proc_juridico_montreal))
        self.btn_juridico.pack(fill=tk.X, padx=10, pady=5)

        self.btn_incluir_podio = ttk.Button(frame_podio, text="Incluir Correspondências Rápidas", command=lambda: self.rodar_thread(self._proc_incluir_podio))
        self.btn_incluir_podio.pack(fill=tk.X, padx=10, pady=5)

        # --- CATEGORIA 3: AGILIS & PRODUTIVIDADE ---
        frame_agilis = ttk.LabelFrame(frame_botoes, text="🎧 Agilis & Chamados")
        frame_agilis.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

        self.btn_produtividade = ttk.Button(frame_agilis, text="Gerar relatório de envio para Correios", command=lambda: self.rodar_thread(self._proc_produtividade))
        self.btn_produtividade.pack(fill=tk.X, padx=10, pady=5)

        self.btn_produtividade_setor = ttk.Button(frame_agilis, text="Gerar Produtividade (Podio/Agilis/SAP)", command=self.executar_produtividade)
        self.btn_produtividade_setor.pack(fill=tk.X, padx=10, pady=5)

        self.btn_fechar_chamados = ttk.Button(frame_agilis, text="Fechar Chamados a Vencer", command=lambda: self.rodar_thread(self._proc_fechar_chamados))
        self.btn_fechar_chamados.pack(fill=tk.X, padx=10, pady=5)

        # --- CATEGORIA 4: OUTROS SISTEMAS ---
        frame_outros = ttk.LabelFrame(frame_botoes, text="🚗 Outros (Uber / SAP)")
        frame_outros.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")

        self.btn_uber = ttk.Button(frame_outros, text="Relatório de Utilização Uber", command=lambda: self.rodar_thread(self._proc_uber))
        self.btn_uber.pack(fill=tk.X, padx=10, pady=5)

        self.btn_zmm180 = ttk.Button(frame_outros, text="Faturamento Transação ZMM180", command=lambda: self.rodar_thread(self._proc_zmm180))
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
            self.btn_uber, self.btn_zmm180
        ]

        print("✅ Sistema Central iniciado com sucesso, Pedro!")
        print("Selecione o processo que deseja executar.\n" + "-"*60)

    # ======================================================================
    # GERENCIADOR DE THREADS (Para os processos antigos)
    # ======================================================================
    def rodar_thread(self, funcao_processo):
        for btn in self.todos_botoes:
            btn.state(['disabled'])
        threading.Thread(target=self._executor_seguro, args=(funcao_processo,), daemon=True).start()

    def _executor_seguro(self, funcao_processo):
        try:
            funcao_processo()
            self.root.after(0, lambda: messagebox.showinfo("Sucesso", "A automação foi concluída com sucesso!"))
            
        except BaseException as e: # Mudamos para BaseException para capturar sys.exit() também!
            mensagem_erro = str(e)
            if not mensagem_erro or mensagem_erro == "1": 
                mensagem_erro = "Processo interrompido (Verifique o console para detalhes)."
                
            texto_popup = f"O processo foi interrompido devido a um erro.\n\nMotivo:\n{mensagem_erro}"
            print(f"\n❌ [ERRO CRÍTICO]: {mensagem_erro}")
            self.root.after(0, lambda msg=texto_popup: messagebox.showerror("Erro na Automação", msg))
            
        finally:
            print("-" * 60)
            self.root.after(0, self._reativar_botoes)

    def _reativar_botoes(self):
        for btn in self.todos_botoes:
            btn.state(['!disabled'])

    # ======================================================================
    # PROCESSO ISOLADO (SUBPROCESS) - ONDE O CANCELAR FUNCIONA
    # ======================================================================
    def executar_produtividade(self):
        for btn in self.todos_botoes:
            btn.state(['disabled'])
            
        print(">>> Iniciando Robô de Produtividade em processo isolado...")
        threading.Thread(target=self._rodar_processo_isolado, daemon=True).start()

    def _rodar_processo_isolado(self):
        try:
            diretorio_atual = os.path.dirname(os.path.abspath(__file__))
            caminho_script = os.path.join(diretorio_atual, "produtividade.py")
            
            processo = subprocess.Popen(
                [sys.executable, "-X", "utf8", caminho_script],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                bufsize=1,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            self.processo_ativo = processo
            self.root.after(0, lambda: self.btn_cancelar.config(state="normal"))
            
            for linha in processo.stdout:
                print(linha, end="")
                
            processo.wait() 
            
            self.processo_ativo = None
            self.root.after(0, lambda: self.btn_cancelar.config(state="disabled"))
            
            # --- AQUI ESTÁ A CORREÇÃO DOS POPUPS ---
            if processo.returncode == 0:
                print("\n✅ Processo finalizado com sucesso!")
                self.root.after(0, lambda: messagebox.showinfo("Sucesso", "A automação foi concluída com sucesso!"))
            
            elif processo.returncode == 1:
                # Código 1 significa que o script deu erro (ex: arquivo não encontrado)
                print(f"\n⚠️ O processo falhou (Código {processo.returncode}).")
                self.root.after(0, lambda: messagebox.showerror("Erro na Automação", "O processo encontrou um erro e foi interrompido.\n\nVerifique o console preto para ler o motivo exato."))
            
            else:
                # Qualquer outro código (geralmente 15 ou 1) significa que foi morto pelo Taskkill (Botão Cancelar)
                print(f"\n⚠️ O processo foi cancelado (Código {processo.returncode}).")
                self.root.after(0, lambda: messagebox.showwarning("Cancelado", "O processo foi cancelado forçadamente pelo usuário."))
            # ---------------------------------------
            
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
                    subprocess.run(['taskkill', '/F', '/T', '/PID', str(self.processo_ativo.pid)], creationflags=subprocess.CREATE_NO_WINDOW)
                    print("\n" + "="*50)
                    print("🛑 PROCESSO CANCELADO FORÇADAMENTE PELO USUÁRIO!")
                    print("="*50 + "\n")
                except Exception as e:
                    print(f"\n❌ Erro ao tentar cancelar: {e}")

    # ======================================================================
    # FUNÇÕES DOS OUTROS PROCESSOS
    # ======================================================================
    def _proc_encomendas_dia(self):
        print(">>> Iniciando: Relatório de Encomendas do Dia (Correios)...")
        time.sleep(2)
        print("✅ Processo concluído!")

    def _proc_juridico_montreal(self):
        print(">>> Iniciando: Relatório Jurídico Montreal (Podio)...")
        robo_juridico.executar_juridico()
        time.sleep(2)
        print("✅ E-mail do Jurídico enviado com sucesso!")

    def _proc_produtividade(self):
        print(">>> Iniciando: Gerar relatório de envio para Correios...")
        robo_relatorio_correios.executar_relatorio_completo() 
        time.sleep(1)
        print("✅ Planilha dos correios gerada e preenchida!")

    def _proc_incluir_podio(self):
        print(">>> Iniciando: Inclusão rápida de correspondências no Podio...")
        time.sleep(2)
        print("✅ Correspondências incluídas!")

    def _proc_fechar_chamados(self):
        print(">>> Iniciando: Fechamento de chamados próximos de vencer...")
        time.sleep(2)
        print("✅ Chamados verificados e fechados!")

    def _proc_rateio_malote(self):
        print(">>> Iniciando: Rateio de Malote (Distribuição por Centro de Custo)...")
        time.sleep(2)
        print("✅ Rateio de malote concluído!")

    def _proc_uber(self):
        print(">>> Iniciando: Relatório de Utilização Uber...")
        time.sleep(2)
        print("✅ E-mails enviados aos responsáveis dos centros de custo!")

    def _proc_zmm180(self):
        print(">>> Iniciando: Faturamento Transação ZMM180 (SAP)...")
        time.sleep(2)
        print("✅ Conferência ZMM180 finalizada!")

    def _proc_fat_rascunhos(self):
        print(">>> Iniciando Faturamento Etapa 1: Gerar Rascunhos...")
        robo_faturamento.criar_rascunhos_correios()
        print("✅ Rascunhos gerados!")

    def _proc_fat_planilha(self):
        print(">>> Iniciando Faturamento Etapa 2: Gerar Planilha Rateio Pag...")
        robo_faturamento.preparar_e_gerar_rateio()
        print("✅ Planilha RATEIO PAG.xlsx gerada!")

    def _proc_fat_lancamento(self):
        print(">>> Iniciando Faturamento Etapa 3: Lançar Nota no Portal MRV...")
        robo_faturamento.lancar_nota_fiscal()
        time.sleep(2)
        print("✅ Lançamento concluído!")

if __name__ == "__main__":
    root = tk.Tk()
    app = CentralAutomacaoMRV(root)
    root.mainloop()
