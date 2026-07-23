import sys
import os
import traceback
import time
import tempfile
import tkinter.messagebox as messagebox
import config

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
        self.root.geometry("1050x700")
        
        # Maximiza a janela de forma suave APÓS ela ser renderizada (evita o bug de piscar)
        self.root.after(0, lambda: self.root.state('zoomed'))
        
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
        criar_botao(frame_correios, "Rateio de Malote (Centros de Custo)", lambda: self._verificar_pasta_e_executar("Rateio de Malote", "import robos.robo_rateio_malote as rrm; rrm.executar_rateio_malote()", os.path.join(config.PASTA_ARQUIVOS, "rateio_malote")))
        criar_botao(frame_correios, "Faturamento 1: Gerar Rascunhos", lambda: self.executar_processo_cancelavel("Faturamento 1", comando_python="import robos.robo_faturamento as rf; rf.criar_rascunhos_correios()"), espaco_extra=True)
        criar_botao(frame_correios, "Faturamento 2: Planilha Rateio Pag", lambda: self._verificar_pasta_e_executar("Faturamento 2", "import robos.robo_faturamento as rf; rf.preparar_e_gerar_rateio()", os.path.join(config.PASTA_ARQUIVOS, "faturamento")))
        criar_botao(frame_correios, "Faturamento 3: Lançar NF (Portal)", lambda: self._verificar_pasta_e_executar("Faturamento 3", "import robos.robo_faturamento as rf; rf.lancar_nota_fiscal()", os.path.join(config.PASTA_ARQUIVOS, "faturamento")))


        frame_podio = criar_quadro(frame_botoes, "Podio & Mensageria", 0, 1)
        criar_botao(frame_podio, "Relatório Jurídico Montreal", self._chamar_robo_juridico)
        criar_botao(frame_podio, "Incluir Correspondências Rápidas", self._chamar_robo_incluir_encomendas)

        frame_agilis = criar_quadro(frame_botoes, "Agilis & Chamados", 1, 0)
        criar_botao(frame_agilis, "Gerar relatório de envio para Correios", lambda: self.executar_processo_cancelavel("Relatório Correios", comando_python="import robos.robo_relatorio_correios as rc; rc.executar_relatorio_completo()"))
        criar_botao(frame_agilis, "Gerar Produtividade (Podio/Agilis/SAP)", self._chamar_robo_produtividade)
        criar_botao(frame_agilis, "Fechar Chamados a Vencer", lambda: self.executar_processo_cancelavel("Fechar Chamados", comando_python="import robos.robo_fechar_chamados as rfc; rfc.executar_fechamento()"))


        frame_outros = criar_quadro(frame_botoes, "Outros (Uber / SAP)", 1, 1)
        criar_botao(frame_outros, "Uber 1: Atualizar Responsáveis (SAP)", lambda: self._verificar_pasta_e_executar("Uber 1", "import robos.robo_uber_relatorios as ru; ru.etapa_1_atualizar_responsaveis()", os.path.join(config.PASTA_ARQUIVOS, "uber")))
        criar_botao(frame_outros, "Uber 2: Gerar Relatórios e Pastas", lambda: self._verificar_pasta_e_executar("Uber 2", "import robos.robo_uber_relatorios as ru; ru.etapa_2_gerar_relatorios()", os.path.join(config.PASTA_ARQUIVOS, "uber")))
        criar_botao(frame_outros, "Uber 3: Criar Rascunhos de E-mail", lambda: self._verificar_pasta_e_executar("Uber 3", "import robos.criar_rascunhos_uber as rr; rr.criar_rascunhos()", os.path.join(config.PASTA_ARQUIVOS, "uber")))
        criar_botao(frame_outros, "Faturamento Transação ZMM180", self._chamar_robo_zmm180, espaco_extra=True)

        self.btn_cancelar = ctk.CTkButton(root, text="CANCELAR PROCESSO ATIVO", 
                                          fg_color=self.COR_CANCELAR, hover_color=self.COR_CANCELAR_HOVER, 
                                          font=ctk.CTkFont(size=14, weight="bold"), height=40,
                                          command=self.cancelar_processo, state="disabled")
        self.btn_cancelar.pack(fill=tk.X, padx=30, pady=(20, 10))

        lbl_console = ctk.CTkLabel(root, text="Console de Execução em Tempo Real:", font=ctk.CTkFont(size=14, weight="bold"))
        lbl_console.pack(anchor=tk.W, padx=30, pady=(0, 5))

        # --- NOVA BARRA DE PROGRESSO ---
        self.progressbar = ctk.CTkProgressBar(root, mode="determinate", height=8, progress_color=self.COR_MRV)
        self.progressbar.pack(fill=tk.X, padx=30, pady=(0, 5))
        self.progressbar.set(0) # Começa vazia
        # -------------------------------

        self.console = ctk.CTkTextbox(root, height=150, font=ctk.CTkFont(family="Consolas", size=13), 
                                      text_color="#00FF00", fg_color="#1E1E1E", border_width=2, border_color="#333333")
        # expand=False garante que o console nunca mude de tamanho e seja esmagado
        self.console.pack(fill=tk.BOTH, expand=False, padx=30, pady=(0, 20))
        self.console.configure(state='disabled')

        sys.stdout = PrintRedirector(self.console)
        sys.stderr = PrintRedirector(self.console)

        print("Sistema Central iniciado com sucesso!")
        print("Selecione o processo que deseja executar.\n" + "-"*60)

        # Chama a verificação de credenciais 1 segundo após abrir o app
        self.root.after(1000, self._verificar_credenciais_iniciais)

    def _verificar_credenciais_iniciais(self):
        # Verifica se as credenciais estão vazias ou com o texto padrão
        email_vazio = not config.EMAIL_MRV or config.EMAIL_MRV == "seu_email@mrv.com.br"
        senha_vazia = not config.SENHA_MRV or config.SENHA_MRV == "sua_senha"
        
        if email_vazio or senha_vazia:
            messagebox.showwarning(
                "Atenção: Credenciais Ausentes",
                "Bem-vindo ao Hub Central!\n\nNotamos que suas credenciais ainda não foram configuradas.\n\nPor favor, clique no botão 'Configurar Credenciais' e preencha seu e-mail e senha antes de executar os robôs."
            )

    def _chamar_robo_juridico(self):
        resposta = messagebox.askyesnocancel("Relatório Jurídico Montreal", "Você deseja que o robô baixe a planilha do Podio automaticamente?\n\nSIM: O robô fará tudo.\nNÃO: Eu já baixei manualmente.\nCANCELAR: Abortar operação.")
        if resposta is True: 
            self.executar_processo_cancelavel("Relatório Jurídico Montreal", comando_python="import robos.robo_juridico as rj; rj.executar_juridico(pular_download=False)")
        elif resposta is False: 
            self.executar_processo_cancelavel("Relatório Jurídico (Apenas Formatação)", comando_python="import robos.robo_juridico as rj; rj.executar_juridico(pular_download=True)")

    def _verificar_pasta_e_executar(self, nome_processo, comando_python, caminho_pasta):
        msg = f"Você já verificou/atualizou os arquivos para o processo '{nome_processo}'?\n\n• OK para rodar o robô.\n• Cancelar para ABRIR A PASTA."
        
        if messagebox.askokcancel(f"Lembrete - {nome_processo}", msg):
            self.executar_processo_cancelavel(nome_processo, comando_python=comando_python)
        else:
            # Se a pessoa clicou em Cancelar, tenta abrir a pasta
            if not os.path.exists(caminho_pasta):
                try:
                    os.makedirs(caminho_pasta) # Cria a pasta se ela não existir ainda
                except:
                    pass
            
            if os.path.exists(caminho_pasta):
                os.startfile(caminho_pasta)
            else:
                messagebox.showwarning("Aviso", f"A pasta não foi encontrada:\n{caminho_pasta}")

    def _chamar_robo_produtividade(self):
        resposta = messagebox.askyesnocancel(
            "Produtividade Setorial", 
            "Você deseja que o robô baixe os relatórios automaticamente (Podio, Agilis, Bússola)?\n\n"
            "SIM: O robô fará o download e a edição.\n"
            "NÃO: Pular download (apenas formatar planilhas já existentes na pasta).\n"
            "CANCELAR: Abortar operação."
        )
        
        if resposta is True:
            if messagebox.askokcancel("Aviso Importante", "ATENÇÃO\n\nNÃO MEXA no mouse ou teclado durante a extração do Bússola.\nDeseja continuar?"):
                self.executar_processo_cancelavel("Produtividade (Completo)", comando_python="import robos.produtividade as rp; rp.executar_robo_produtividade_setor(pular_extracao=False)")
        elif resposta is False:
            self.executar_processo_cancelavel("Produtividade (Apenas Edição)", comando_python="import robos.produtividade as rp; rp.executar_robo_produtividade_setor(pular_extracao=True)")
            
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
            
        # Zera a barra antes de começar
        self.progressbar.set(0)
            
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
                        # INTERCEPTADOR DE PROGRESSO
                        if "[PROGRESSO:" in linha:
                            try:
                                # Pega apenas o número. Ex: "[PROGRESSO: 50]" vira "50"
                                valor_str = linha.split("[PROGRESSO:")[1].replace("]", "").strip()
                                valor = float(valor_str)
                                # A barra vai de 0.0 a 1.0, então dividimos por 100
                                self.root.after(0, lambda v=valor: self.progressbar.set(v / 100.0))
                            except:
                                pass # Se der erro na conversão, apenas ignora
                        else:
                            # Se for texto normal, mostra no console
                            print(linha, end="")
                            linhas_log.append(linha.rstrip('\r\n'))
                    else:
                        time.sleep(0.1)
                
                # Lê qualquer restinho de texto que ficou após o robô terminar
                for linha in f.readlines():
                    if "[PROGRESSO:" in linha:
                        try:
                            valor_str = linha.split("[PROGRESSO:")[1].replace("]", "").strip()
                            valor = float(valor_str)
                            self.root.after(0, lambda v=valor: self.progressbar.set(v / 100.0))
                        except:
                            pass
                    else:
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
            
        # Zera a barra de progresso ao finalizar
        self.progressbar.set(0)

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

        # Adiciona o atalho CTRL+F na janela de ajuda
        popup.bind("<Control-f>", lambda e: JanelaBusca(popup, textbox))
        popup.bind("<Control-F>", lambda e: JanelaBusca(popup, textbox))

        texto_ajuda = """Bem-vindo à Central de Automações MRV! Siga as instruções abaixo para garantir que tudo funcione perfeitamente.

PRIMEIRO PASSO: CREDENCIAIS!!!!

Antes de rodar qualquer robô, clique no botão "⚙️ Configurar Credenciais". 
Preencha seu E-mail MRV, sua Senha MRV e a Senha do Malote Web (se você tiver). O sistema salvará isso de forma segura para os robôs usarem automaticamente.


ROBÔS QUE ACESSAM SITES (Aviso Importante) !!!!

• Tempo de Carregamento: Os robôs que abrem o navegador (Chrome/Edge) podem demorar alguns segundos para iniciar. Aguarde a tela abrir sozinha, eles são os robôs de: Faturamento 3, Relatório Jurídico Montreal, incluir correspodências rápidas, gerar produtividade, gerar relatório de envio para correios, fechar chamados a vencer, rateio de malote.
• Código de Segurança (MFA): Na PRIMEIRA VEZ que o robô acessar os portais da MRV por processo (exceto o Malote Web), ele vai preencher seu e-mail e senha, mas o site pedirá a aprovação no seu celular (MFA/Token). Fique com o celular em mãos para aprovar o acesso!


REGRAS DOS PROCESSOS!!!!

• Correios & Faturamento: Lembre-se de colocar as planilhas e PDFs corretos dentro da pasta "arquivos", ela fica dentro da pasta "dist" antes de rodar os robôs. O robô vai ler os dados de lá.
• SAP (Produtividade, ZMM180): Quando o robô avisar, deixe o SAP aberto na SEGUNDA TELA e NÃO MEXA no mouse ou teclado enquanto ele trabalha.
• Podio & Mensageria: O robô fará os downloads e uploads automaticamente, apenas confirme as caixas de aviso que aparecerem na tela.


ARQUIVOS OBRIGATÓRIOS PARA CADA ROBÔ!!!!

• Rateio de Malote: Planilha dos correios (ela se parece com isso 2554871.xlsx), 
planilha relatório agilis (ela se parece com isso Relatório Agilis -  01.05 a 23.06.xlsx), 
base de centro de custo, e acompanhamento de VSC mais recente. 
Todos estão dentro da pasta "dist/arquivos/rateio_malote".

• Faturamento 2: Planilha dos correios (ela se parece com isso 2554871.xlsx),
e planilha rateio recebido (ela se parece com isso Rateio Recebido.xlsx) 
todos eles ficam dentro da pasta "dist/arquivos/faturamento/testar_edicao".

• Faturamento 3: Planilha dados_puxados_preenchimento.xlsx ela fica na pasta "dist/arquivos/faturamento",
planilha RATEIO PAG.xlsx e o boleto (pdf), eles ficam na pasta "dist/arquivos/faturamento/exemplos".

• Produtividade: Planilha de produtividade (ela se parece com isso Produtividade 05 - 2026.xlsx) ela fica na pasta "dist/arquivos/produtividade".

• Incluir encomendas rápidass: Planilha encomendas.xlsx ela fica na pasta "dist/arquivos/encomendas".

• Uber 1: Planilha retirada do sap (ela se parece com isso EXPORT_20260721_132513.xlsx, ela sempre começa com "EXPORT_"),
Planilha Responsaveis Por Centro de Custos.xlsx as duas ficam na pasta "dist/arquivos/uber".

• Uber 2: A planilha Responsaveis_Atualizado_SAP.xlsx (ela aparece depois de rodar o Uber 1),
E a planilha de relatórios do mês (ela se parece com isso Relatório Junho - 2026 Atualizado.xlsx, ela sempre começa com "Relatório_") 
As duas ficam na pasta "dist/arquivos/uber".

• Uber 3: Os arquivos que são gerados dentro da  2026,06 (ou seja, ano,mes) dentro da pasta "dist/arquivos/uber".

CANCELAMENTO DE EMERGÊNCIA!!!!

• Se algo der errado, se você esquecer de fechar uma planilha, ou precisar usar o PC na mesma hora, clique no botão vermelho "CANCELAR PROCESSO ATIVO". Ele vai forçar a parada do robô imediatamente e fechar os navegadores abertos por ele.


OBSERVAÇÕES IMPORTANTES!!!!

• Os robôs de Uber, você deverá rodar na ordem: Uber 1, Uber 2 e depois Uber 3. Se você pular algum deles, o próximo não vai funcionar.

• Na planilha encomendas.xlsx, não altere os nomes das colunas, pois o robô depende delas para funcionar corretamente.

• O robô de faturamento 3 depende do robô de faturamento 2, então sempre rode o faturamento 2 antes do faturamento 3. 

• O robô Uber 2 vai gerar esses arquivos: PENDENCIAS_CARGO.xlsx, consolidado_para_envio_ATUALIZADO.xlsx, TESTE MACRO 2026,06 COM_ESTILO.xlsx, Enviar_e-mail original COM_ESTILO.xlsx e a pasta 2026,06 (ano,mes). Não altere os nomes deles, pois o robô Uber 3 depende desses nomes para funcionar corretamente.

• O robô Uber 1 vai gerar a planilha Responsaveis_Atualizado_SAP.xlsx, que é necessária para o Uber 2. É recomendável você analisar os nomes que o robô não conseguiu achar por não estar na base de ativos.

• O robô de Rateio de Malote vai gerar a planilha Rateio_Malote.xlsx.

"""
        # Insere o texto e bloqueia para o usuário não conseguir apagar
        textbox.insert("0.0", texto_ajuda)
        textbox.configure(state="disabled")


class JanelaBusca:
    def __init__(self, parent, textbox):
        self.top = ctk.CTkToplevel(parent)
        self.top.title("Buscar na Ajuda")
        self.top.geometry("350x80")
        self.top.attributes("-topmost", True) # Mantém a busca sempre na frente
        self.top.resizable(False, False)
        
        # O CTkTextbox usa um widget Text do Tkinter por baixo dos panos
        self.textbox = textbox._textbox 
        self.last_pos = "1.0"

        self.entry = ctk.CTkEntry(self.top, placeholder_text="Digite para buscar...")
        self.entry.pack(side="left", padx=10, pady=10, expand=True, fill="x")
        # Permite buscar apertando Enter
        self.entry.bind("<Return>", self.buscar)

        self.btn = ctk.CTkButton(self.top, text="Próximo", width=80, command=self.buscar)
        self.btn.pack(side="right", padx=10, pady=10)

        self.entry.focus()

    def buscar(self, event=None):
        query = self.entry.get()
        # Remove o grifo amarelo da busca anterior
        self.textbox.tag_remove("highlight", "1.0", "end")
        if not query:
            return

        # Procura a palavra a partir da última posição
        pos = self.textbox.search(query, self.last_pos, stopindex="end", nocase=True)
        
        # Se não achar, volta para o começo do texto e procura de novo
        if not pos:
            pos = self.textbox.search(query, "1.0", stopindex="end", nocase=True)

        if pos:
            # Calcula onde a palavra termina
            end_pos = f"{pos}+{len(query)}c"
            # Adiciona um grifo amarelo na palavra encontrada
            self.textbox.tag_add("highlight", pos, end_pos)
            self.textbox.tag_config("highlight", background="#FFC000", foreground="black")
            # Rola a tela até a palavra
            self.textbox.see(pos)
            # Salva a posição para a próxima busca
            self.last_pos = end_pos
        else:
            self.last_pos = "1.0"

if __name__ == "__main__":
    root = ctk.CTk() 
    app = CentralAutomacaoMRV(root)
    root.mainloop()
