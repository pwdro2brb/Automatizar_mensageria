import sys
import os
import traceback
import time
import tempfile

# ==============================================================================
# 🚀 1. INTERCEPTADOR DE PROCESSOS (Bypass do PyInstaller --noconsole)
# ==============================================================================
if len(sys.argv) > 2 and sys.argv[1] == "--run-code":
    codigo = sys.argv[2]
    log_path = sys.argv[3] if len(sys.argv) > 3 else None
    
    if log_path:
        # Redireciona a "voz" do robô para um arquivo de texto temporário
        sys.stdout = open(log_path, "w", encoding="utf-8", buffering=1)
        sys.stderr = sys.stdout
        
    try:
        exec(codigo)
    except Exception as e:
        # Se der erro, escreve o erro completo no arquivo para o Hub ler
        traceback.print_exc()
        sys.exit(1)
    sys.exit(0)

# ==============================================================================
# 📦 2. FORÇAR O PYINSTALLER A EMPACOTAR OS ROBÔS
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
except ImportError:
    pass 

import ctypes

# ==============================================================================
# 🚀 FORÇAR O WINDOWS A RECONHECER TODOS OS MONITORES
# ==============================================================================
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import threading
import subprocess
import config

# Configuração do Tema Moderno
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("green")

class PrintRedirector:
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

        self.COR_MRV = "#008542"
        self.COR_MRV_HOVER = "#006331"
        self.COR_CANCELAR = "#E74C3C"
        self.COR_CANCELAR_HOVER = "#C0392B"

        self.root = root
        self.root.title("Hub Central de Automações - MRV")
        self.root.geometry("1050x700") # Reduzimos a altura padrão
        
        # Faz a janela abrir maximizada no Windows (respeitando a barra de tarefas)
        try:
            self.root.state('zoomed')
        except:
            pass
        
        lbl_titulo = ctk.CTkLabel(root, text="🤖 Central de Robôs - Administrativo MRV", font=ctk.CTkFont(size=16, weight="bold"))
        lbl_titulo.pack(pady=(0, 5))

        # Cria um espaço para colocar os botões do topo lado a lado
        frame_topo = ctk.CTkFrame(root, fg_color="transparent")
        frame_topo.pack(pady=(0, 15))

        btn_config = ctk.CTkButton(frame_topo, text="⚙️ Configurar Credenciais", command=self._abrir_popup_config, 
                                   fg_color="#3498DB", hover_color="#2980B9", font=ctk.CTkFont(weight="bold"))
        btn_config.pack(side=tk.LEFT, padx=10)

        btn_ajuda = ctk.CTkButton(frame_topo, text="❓ Ajuda / Tutorial", command=self._abrir_popup_ajuda, 
                                  fg_color="#F39C12", hover_color="#D68910", font=ctk.CTkFont(weight="bold"))
        btn_ajuda.pack(side=tk.LEFT, padx=10)


        # CTkScrollableFrame permite rolar a tela se os botões não couberem
        frame_botoes = ctk.CTkScrollableFrame(root, fg_color="transparent")
        frame_botoes.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        frame_botoes.columnconfigure(0, weight=1)
        frame_botoes.columnconfigure(1, weight=1)

        def criar_quadro(parent, titulo, row, col):
            frm = ctk.CTkFrame(parent)
            frm.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
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

        frame_correios = criar_quadro(frame_botoes, "Correios & Faturamento", 0, 0)
        criar_botao(frame_correios, "Rateio de Malote (Centros de Custo)", lambda: self._verificar_planilhas_e_executar("Rateio de Malote", "import robos.robo_rateio_malote as rrm; rrm.executar_rateio_malote()"))
        criar_botao(frame_correios, "Faturamento 1: Gerar Rascunhos", lambda: self.executar_processo_cancelavel("Faturamento 1", comando_python="import robos.robo_faturamento as rf; rf.criar_rascunhos_correios()"), espaco_extra=True)
        criar_botao(frame_correios, "Faturamento 2: Planilha Rateio Pag", lambda: self._verificar_planilhas_e_executar("Faturamento 2", "import robos.robo_faturamento as rf; rf.preparar_e_gerar_rateio()"))
        criar_botao(frame_correios, "Faturamento 3: Lançar NF (Portal)", lambda: self.executar_processo_cancelavel("Faturamento 3", comando_python="import robos.robo_faturamento as rf; rf.lancar_nota_fiscal()"))

        frame_podio = criar_quadro(frame_botoes, "Podio & Mensageria", 0, 1)
        criar_botao(frame_podio, "Relatório Jurídico Montreal", self._chamar_robo_juridico)
        criar_botao(frame_podio, "Incluir Correspondências Rápidas", self._chamar_robo_incluir_encomendas)

        frame_agilis = criar_quadro(frame_botoes, "Agilis & Chamados", 1, 0)
        criar_botao(frame_agilis, "Gerar relatório de envio para Correios", lambda: self.executar_processo_cancelavel("Relatório Correios", comando_python="import robos.robo_relatorio_correios as rc; rc.executar_relatorio_completo()"))
        criar_botao(frame_agilis, "Gerar Produtividade (Podio/Agilis/SAP)", self._chamar_robo_produtividade)
        criar_botao(frame_agilis, "Fechar Chamados a Vencer", lambda: self.executar_processo_cancelavel("Fechar Chamados", comando_python="import robos.robo_fechar_chamados as rfc; rfc.executar_fechamento()"))


        frame_outros = criar_quadro(frame_botoes, "Outros (Uber / SAP)", 1, 1)
        criar_botao(frame_outros, "Uber 1: Atualizar Responsáveis (SAP)", lambda: self._verificar_planilhas_e_executar("Uber 1", "import robos.robo_uber_relatorios as ru; ru.etapa_1_atualizar_responsaveis()"))
        criar_botao(frame_outros, "Uber 2: Gerar Relatórios e Pastas", lambda: self._verificar_planilhas_e_executar("Uber 2", "import robos.robo_uber_relatorios as ru; ru.etapa_2_gerar_relatorios()"))
        criar_botao(frame_outros, "Uber 3: Criar Rascunhos de E-mail", lambda: self.executar_processo_cancelavel("Uber 3", comando_python="import robos.criar_rascunhos_uber as rr; rr.criar_rascunhos()"))
        criar_botao(frame_outros, "Faturamento Transação ZMM180", self._chamar_robo_zmm180, espaco_extra=True)

        self.btn_cancelar = ctk.CTkButton(root, text="CANCELAR PROCESSO ATIVO", 
                                          fg_color=self.COR_CANCELAR, hover_color=self.COR_CANCELAR_HOVER, 
                                          font=ctk.CTkFont(size=14, weight="bold"), height=40,
                                          command=self.cancelar_processo, state="disabled")
        self.btn_cancelar.pack(fill=tk.X, padx=30, pady=(20, 10))

        lbl_console = ctk.CTkLabel(root, text="Console de Execução em Tempo Real:", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_console.pack(anchor=tk.W, padx=30, pady=(0, 5))

        self.console = ctk.CTkTextbox(root, height=150, font=ctk.CTkFont(family="Consolas", size=13), 
                                      text_color="#00FF00", fg_color="#1E1E1E", border_width=2, border_color="#333333")
        # expand=False garante que o console nunca mude de tamanho e seja esmagado
        self.console.pack(fill=tk.BOTH, expand=False, padx=30, pady=(0, 20))
        self.console.configure(state='disabled')

        sys.stdout = PrintRedirector(self.console)
        sys.stderr = PrintRedirector(self.console)

        print("Sistema Central iniciado com sucesso, Pedro!")
        print("Selecione o processo que deseja executar.\n" + "-"*60)

    def _chamar_robo_juridico(self):
        resposta = messagebox.askyesnocancel("Relatório Jurídico Montreal", "Você deseja que o robô baixe a planilha do Podio automaticamente?\n\nSIM: O robô fará tudo.\nNÃO: Eu já baixei manualmente.\nCANCELAR: Abortar operação.")
        if resposta is True: 
            self.executar_processo_cancelavel("Relatório Jurídico Montreal", comando_python="import robos.robo_juridico as rj; rj.executar_juridico(pular_download=False)")
        elif resposta is False: 
            self.executar_processo_cancelavel("Relatório Jurídico (Apenas Formatação)", comando_python="import robos.robo_juridico as rj; rj.executar_juridico(pular_download=True)")

    def _verificar_planilhas_e_executar(self, nome_processo, comando_python):
        if messagebox.askyesno("Lembrete de Arquivos", f"LEMBRETE\n\nVocê já atualizou/trocou as planilhas na pasta para rodar o {nome_processo}?"):
            self.executar_processo_cancelavel(nome_processo, comando_python=comando_python)

    def _chamar_robo_produtividade(self):
        if messagebox.askokcancel("Aviso Importante - SAP e Mouse", "ATENÇÃO\n\n1. Deixe o SAP aberto na SEGUNDA TELA.\n2. NÃO MEXA no mouse ou teclado.\n\nDeseja continuar?"):
            self.executar_processo_cancelavel("Produtividade Setorial", comando_python="import robos.produtividade as rp; rp.executar_robo_produtividade_setor()")

    def _chamar_robo_incluir_encomendas(self):
        if messagebox.askokcancel("Lembrete - Correspondências", "Você lembrou de preencher a planilha?\n\n• OK para rodar o robô.\n• Cancelar para ABRIR A PLANILHA."):
            self.executar_processo_cancelavel("Incluir Correspondências", comando_python="import robos.robo_incluir_encomendas as rie; rie.executar_inclusao()")
        else:
            caminho_planilha = os.path.join(config.PASTA_ARQUIVOS, "encomendas", "encomendas.xlsx")
            if os.path.exists(caminho_planilha):
                os.startfile(caminho_planilha)
            else:
                messagebox.showwarning("Aviso", f"A planilha 'encomendas.xlsx' ainda não foi encontrada na pasta:\n{os.path.join(config.PASTA_ARQUIVOS, 'encomendas')}")

    def _chamar_robo_zmm180(self):
        if messagebox.askokcancel("Aviso Importante - SAP e Edge", "ATENÇÃO\n\n1. Deixe o SAP aberto na SEGUNDA TELA.\n2. Deixe o documento aberto no Edge.\n3. NÃO MEXA no mouse.\n\nDeseja continuar?"):
            self.executar_processo_cancelavel("Faturamento ZMM180", comando_python="import robos.robo_zmm180 as rz; rz.executar_zmm180()")

    def executar_processo_cancelavel(self, nome_processo, comando_python=None):
        for btn in self.todos_botoes:
            btn.configure(state="disabled")
            
        print(f">>> Iniciando: {nome_processo}...")
        threading.Thread(target=self._rodar_subprocesso, args=(comando_python,), daemon=True).start()

    def _rodar_subprocesso(self, comando_python):
        self.foi_cancelado = False 
        
        # Cria um arquivo temporário para capturar os logs do robô
        fd, log_path = tempfile.mkstemp(suffix=".log", text=True)
        os.close(fd)
        
        try:
            if getattr(sys, 'frozen', False):
                cmd = [sys.executable, "--run-code", comando_python, log_path]
            else:
                cmd = [sys.executable, sys.argv[0], "--run-code", comando_python, log_path]

            processo = subprocess.Popen(cmd, creationflags=subprocess.CREATE_NO_WINDOW)
            
            self.processo_ativo = processo
            self.root.after(0, lambda: self.btn_cancelar.configure(state="normal"))
            
            linhas_log = []

            # Lê o arquivo temporário em tempo real enquanto o robô roda
            with open(log_path, "r", encoding="utf-8") as f:
                while processo.poll() is None:
                    linha = f.readline()
                    if linha:
                        print(linha, end="")
                        linhas_log.append(linha.rstrip('\r\n'))
                    else:
                        time.sleep(0.1)
                
                # Lê qualquer restinho de texto que ficou após o robô terminar
                for linha in f.readlines():
                    print(linha, end="")
                    linhas_log.append(linha.rstrip('\r\n'))
            
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
                
                # Agora o rastreador de erros vai achar o erro perfeitamente!
                linhas_erro = [l for l in linhas_log if l.strip()]
                texto_erro = ""
                
                for i in range(len(linhas_erro)):
                    if "Traceback (most recent call last):" in linhas_erro[i]:
                        texto_erro = "\n".join(linhas_erro[i:])
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
            # Apaga o arquivo temporário para não lotar o PC
            try:
                os.remove(log_path)
            except:
                pass

    def cancelar_processo(self):
        if self.processo_ativo and self.processo_ativo.poll() is None:
            if messagebox.askyesno("Atenção", "Tem certeza que deseja cancelar o robô?\nIsso fechará os navegadores abertos por ele."):
                try:
                    self.foi_cancelado = True 
                    subprocess.run(['taskkill', '/F', '/T', '/PID', str(self.processo_ativo.pid)], creationflags=subprocess.CREATE_NO_WINDOW)
                    print("\n" + "="*50)
                    print("PROCESSO CANCELADO FORÇADAMENTE PELO USUÁRIO!")
                    print("="*50 + "\n")
                except Exception as e:
                    print(f"\nErro ao tentar cancelar: {e}")

    def _reativar_botoes(self):
        for btn in self.todos_botoes:
            btn.configure(state="normal")

    def _abrir_popup_config(self):
        popup = ctk.CTkToplevel(self.root)
        popup.title("Configurar Credenciais")
        # Aumentei a altura da janela de 280 para 350 para caber o novo campo
        popup.geometry("400x350")
        popup.grab_set()
        popup.attributes("-topmost", True)

        ctk.CTkLabel(popup, text="E-mail MRV:", font=ctk.CTkFont(weight="bold")).pack(pady=(20, 0))
        entry_email = ctk.CTkEntry(popup, width=300)
        entry_email.pack(pady=5)
        entry_email.insert(0, config.EMAIL_USER)

        ctk.CTkLabel(popup, text="Senha MRV:", font=ctk.CTkFont(weight="bold")).pack(pady=(10, 0))
        entry_senha = ctk.CTkEntry(popup, width=300, show="*")
        entry_senha.pack(pady=5)
        entry_senha.insert(0, config.SENHA_USER)

        # Novo campo para a senha do Malote Web
        ctk.CTkLabel(popup, text="Senha Malote Web:", font=ctk.CTkFont(weight="bold")).pack(pady=(10, 0))
        entry_senha_malote = ctk.CTkEntry(popup, width=300, show="*")
        entry_senha_malote.pack(pady=5)
        entry_senha_malote.insert(0, getattr(config, "SENHA_MALOTE", ""))

        def salvar():
            novo_email = entry_email.get().strip()
            nova_senha = entry_senha.get().strip()
            nova_senha_malote = entry_senha_malote.get().strip()
            
            # Salva no JSON passando os 3 parâmetros
            config.salvar_credenciais(novo_email, nova_senha, nova_senha_malote)
            
            # Atualiza as variáveis em memória
            config.EMAIL_USER = novo_email
            config.SENHA_USER = nova_senha
            config.SENHA_MALOTE = nova_senha_malote
            
            messagebox.showinfo("Sucesso", "Credenciais salvas com sucesso!", parent=popup)
            popup.destroy()

        ctk.CTkButton(popup, text="Salvar", command=salvar, fg_color=self.COR_MRV, hover_color=self.COR_MRV_HOVER, font=ctk.CTkFont(weight="bold")).pack(pady=25)
    def _abrir_popup_ajuda(self):
        popup = ctk.CTkToplevel(self.root)
        popup.title("Ajuda e Tutorial")
        popup.geometry("650x550")
        popup.grab_set()
        popup.attributes("-topmost", True)

        lbl_titulo = ctk.CTkLabel(popup, text="Guia de Uso dos Robôs", font=ctk.CTkFont(size=18, weight="bold"))
        lbl_titulo.pack(pady=(15, 10))

        # Caixa de texto com barra de rolagem
        textbox = ctk.CTkTextbox(popup, width=600, height=450, wrap="word", font=ctk.CTkFont(size=14))
        textbox.pack(padx=20, pady=(0, 20), fill=tk.BOTH, expand=True)

        texto_ajuda = """Bem-vindo à Central de Automações MRV! Siga as instruções abaixo para garantir que tudo funcione perfeitamente.

PRIMEIRO PASSO: CREDENCIAIS

Antes de rodar qualquer robô, clique no botão "⚙️ Configurar Credenciais". 
Preencha seu E-mail MRV, sua Senha MRV e a Senha do Malote Web (se você tiver). O sistema salvará isso de forma segura para os robôs usarem automaticamente.

ROBÔS QUE ACESSAM SITES (Aviso Importante)

• Tempo de Carregamento: Os robôs que abrem o navegador (Chrome/Edge) podem demorar alguns segundos para iniciar. Aguarde a tela abrir sozinha, eles são os robôs de: Faturamento 3, Relatório Jurídico Montreal, incluir correspodências rápidas, gerar produtividade, gerar relatório de envio para correios, fechar chamados a vencer, rateio de malote.
• Código de Segurança (MFA): Na PRIMEIRA VEZ que o robô acessar os portais da MRV por processo (exceto o Malote Web), ele vai preencher seu e-mail e senha, mas o site pedirá a aprovação no seu celular (MFA/Token). Fique com o celular em mãos para aprovar o acesso!

REGRAS DOS PROCESSOS

• Correios & Faturamento: Lembre-se de colocar as planilhas e PDFs corretos dentro da pasta "arquivos", ela fica dentro da pasta "dist" antes de rodar os robôs. O robô vai ler os dados de lá.
• SAP (Produtividade, ZMM180): Quando o robô avisar, deixe o SAP aberto na SEGUNDA TELA e NÃO MEXA no mouse ou teclado enquanto ele trabalha.
• Podio & Mensageria: O robô fará os downloads e uploads automaticamente, apenas confirme as caixas de aviso que aparecerem na tela.

CANCELAMENTO DE EMERGÊNCIA

Se algo der errado, se você esquecer de fechar uma planilha, ou precisar usar o PC na mesma hora, clique no botão vermelho "CANCELAR PROCESSO ATIVO". Ele vai forçar a parada do robô imediatamente e fechar os navegadores abertos por ele.
"""
        # Insere o texto e bloqueia para o usuário não conseguir apagar
        textbox.insert("0.0", texto_ajuda)
        textbox.configure(state="disabled")




if __name__ == "__main__":
    root = ctk.CTk() 
    app = CentralAutomacaoMRV(root)
    root.mainloop()
