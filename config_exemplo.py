import os
import json
import sys
import os

# ==============================================================================
# 1. CAMINHOS DINÂMICOS (Funciona em qualquer PC)
# ==============================================================================
# os.path.expanduser("~") pega a pasta raiz do usuário atual (Ex: C:\Users\joao.silva)
USER_HOME = os.path.expanduser("~")

PASTA_DOWNLOADS = os.path.join(USER_HOME, "Downloads")
PASTA_PRODUTIVIDADE = os.path.join(USER_HOME, "OneDrive - MRV", "Área de Trabalho", "produtividade")

# ==============================================================================
# 2. GERENCIAMENTO DE CREDENCIAIS (Salva em um arquivo config.json)
# ==============================================================================
ARQUIVO_CONFIG = "config_mrv.json"

def carregar_credenciais():
    """Lê o arquivo JSON se ele existir. Se não, retorna vazio."""
    if os.path.exists(ARQUIVO_CONFIG):
        try:
            with open(ARQUIVO_CONFIG, "r") as f:
                return json.load(f)
        except:
            pass
    return {"email": "", "senha": ""}

def salvar_credenciais(email, senha):
    """Salva o e-mail e a senha no arquivo JSON."""
    with open(ARQUIVO_CONFIG, "w") as f:
        json.dump({"email": email, "senha": senha}, f)

# Carrega as variáveis para serem usadas pelos robôs
credenciais = carregar_credenciais()
EMAIL_USER = credenciais.get("email", "")
SENHA_USER = credenciais.get("senha", "")


# ==============================================================================
# 3. MODO DE COMPATIBILIDADE (Para os robôs antigos continuarem funcionando)
# ==============================================================================
EMAIL_MRV = EMAIL_USER
SENHA_MRV = SENHA_USER


# ==============================================================================
# 4. RADAR DE PASTAS (Ignora a pasta temporária do PyInstaller)
# ==============================================================================
if getattr(sys, 'frozen', False):
    # Se estiver rodando como .exe, pega a pasta onde o .exe está salvo fisicamente
    PASTA_RAIZ = os.path.dirname(sys.executable)
else:
    # Se estiver rodando no VS Code, pega a pasta onde este config.py está
    PASTA_RAIZ = os.path.dirname(os.path.abspath(__file__))

# Caminho oficial da pasta de arquivos (Sempre vai procurar a pasta "arquivos" ao lado do programa)
PASTA_ARQUIVOS = os.path.join(PASTA_RAIZ, "arquivos")