import tkinter as tk
from tkinter import ttk, messagebox
import threading
import sys
import time

# ==============================================================================
# IMPORTAÇÃO DOS SEUS MÓDULOS (SCRIPTS SEPARADOS)
# Exemplo de como você vai importar seus códigos:
import Processos_simples.robo_relatorio_correios as robo_relatorio_correios
import Processos_simples.robo_juridico as robo_juridico
import robo_faturamento
# ==============================================================================

class PrintRedirector:
    """Redireciona os prints do terminal para a caixa de texto da interface."""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, text):
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, text)
        self.text_widget.see(tk.END)
        self.text_widget.configure(state='disabled')

    def flush(self):
        pass

class CentralAutomacaoMRV:
    def __init__(self, root):
        self.root = root
        self.root.title("Hub Central de Automações - MRV")
        self.root.geometry("950x750")
        self.root.configure(padx=15, pady=15)

        # Estilos
        style = ttk.Style()
        style.configure("TButton", font=("Arial", 9, "bold"), padding=5)
        style.configure("TLabelframe.Label", font=("Arial", 10, "bold"), foreground="#003366")

        # Título Principal
        lbl_titulo = ttk.Label(root, text="🤖 Central de Robôs - Administrativo MRV", font=("Arial", 16, "bold"))
        lbl_titulo.pack(pady=(0, 15))

        # ======================================================================
        # ÁREA DE BOTÕES (Organizados por Categorias usando LabelFrames)
        # ======================================================================
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

        self.btn_fechar_chamados = ttk.Button(frame_agilis, text="Fechar Chamados a Vencer", command=lambda: self.rodar_thread(self._proc_fechar_chamados))
        self.btn_fechar_chamados.pack(fill=tk.X, padx=10, pady=5)

        # --- CATEGORIA 4: OUTROS SISTEMAS ---
        frame_outros = ttk.LabelFrame(frame_botoes, text="🚗 Outros (Uber / SAP)")
        frame_outros.grid(row=1, column=1, padx=10, pady=10, sticky="nsew")

        self.btn_uber = ttk.Button(frame_outros, text="Relatório de Utilização Uber", command=lambda: self.rodar_thread(self._proc_uber))
        self.btn_uber.pack(fill=tk.X, padx=10, pady=5)

        self.btn_zmm180 = ttk.Button(frame_outros, text="Faturamento Transação ZMM180", command=lambda: self.rodar_thread(self._proc_zmm180))
        self.btn_zmm180.pack(fill=tk.X, padx=10, pady=5)

        # Configurar grid para expandir igualmente
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

        # Lista de todos os botões para facilitar habilitar/desabilitar
        self.todos_botoes = [
            self.btn_enc_dia, self.btn_rateio_malote, self.btn_fat_1, self.btn_fat_2, self.btn_fat_3,
            self.btn_juridico, self.btn_incluir_podio, self.btn_produtividade, self.btn_fechar_chamados,
            self.btn_uber, self.btn_zmm180
        ]

        print("✅ Sistema Central iniciado com sucesso, Pedro!")
        print("Selecione o processo que deseja executar.\n" + "-"*60)

    # ======================================================================
    # GERENCIADOR DE THREADS E BOTÕES
    # ======================================================================
    def rodar_thread(self, funcao_processo):
        """Desabilita os botões e roda a função escolhida em segundo plano."""
        for btn in self.todos_botoes:
            btn.state(['disabled'])
        
        # Cria a thread passando a função alvo
        threading.Thread(target=self._executor_seguro, args=(funcao_processo,), daemon=True).start()

    def _executor_seguro(self, funcao_processo):
        """Executa a função e garante que os botões voltem ao normal no final."""
        try:
            funcao_processo()
            
            # Opcional: Mostra um popup de SUCESSO quando o robô termina sem erros
            self.root.after(0, lambda: messagebox.showinfo(
                "Sucesso", 
                "A automação foi concluída com sucesso!"
            ))
            
        except Exception as e:
            # 1. Salvamos o erro em formato de texto ANTES do bloco terminar
            mensagem_erro = str(e)
            texto_popup = f"O processo foi interrompido devido a um erro.\n\nMotivo:\n{mensagem_erro}"
            
            print(f"\n❌ [ERRO CRÍTICO]: {mensagem_erro}")
            
            # 2. Passamos o texto salvo para dentro do lambda (msg=texto_popup)
            self.root.after(0, lambda msg=texto_popup: messagebox.showerror(
                "Erro na Automação", 
                msg
            ))
            
        finally:
            print("-" * 60)
            # Reabilita os botões de forma segura na thread principal
            self.root.after(0, self._reativar_botoes)

    def _reativar_botoes(self):
        """Função auxiliar para reativar os botões na interface"""
        for btn in self.todos_botoes:
            btn.state(['!disabled'])
    # ======================================================================
    # FUNÇÕES DOS PROCESSOS (Aqui você chama seus scripts importados)
    # ======================================================================
    
    def _proc_encomendas_dia(self):
        print(">>> Iniciando: Relatório de Encomendas do Dia (Correios)...")
        # robo_encomendas.executar()
        time.sleep(2) # Simulação
        print("✅ Processo concluído!")

    def _proc_juridico_montreal(self):
        print(">>> Iniciando: Relatório Jurídico Montreal (Podio)...")
        # Aqui você chama a função principal do seu Código 2
        robo_juridico.executar_juridico()
        time.sleep(2)
        print("✅ E-mail do Jurídico enviado com sucesso!")

    def _proc_produtividade(self):
        print(">>> Iniciando: Gerar relatório de envio para Correios...")
        # Mude de processar_relatorio_email() para executar_relatorio_completo()
        robo_relatorio_correios.executar_relatorio_completo() 
        time.sleep(1)
        print("✅ Planilha de produtividade gerada e preenchida!")

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
        print("⚠️ Aviso: Este processo ainda está em desenvolvimento/incompleto.")
        time.sleep(2)
        print("✅ Conferência ZMM180 finalizada!")

    # --- Funções do Faturamento Correios (Seu Código 3) ---
    def _proc_fat_rascunhos(self):
        print(">>> Iniciando Faturamento Etapa 1: Gerar Rascunhos...")
        robo_faturamento.criar_rascunhos_correios() # Chama a função direto!
        print("✅ Rascunhos gerados!")

    def _proc_fat_planilha(self):
        print(">>> Iniciando Faturamento Etapa 2: Gerar Planilha Rateio Pag...")
        robo_faturamento.preparar_e_gerar_rateio() # Chama a função direto!
        print("✅ Planilha RATEIO PAG.xlsx gerada!")

    def _proc_fat_lancamento(self):
        print(">>> Iniciando Faturamento Etapa 3: Lançar Nota no Portal MRV...")
        robo_faturamento.lancar_nota_fiscal() # Chama a função direto!
        time.sleep(2)
        print("✅ Lançamento concluído!")

if __name__ == "__main__":
    root = tk.Tk()
    app = CentralAutomacaoMRV(root)
    root.mainloop()
