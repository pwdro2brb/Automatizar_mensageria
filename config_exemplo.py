import os
import json
import sys

# ==============================================================================
# 1. CAMINHOS DINÂMICOS (Funciona em qualquer PC)
# ==============================================================================
USER_HOME = os.path.expanduser("~")

PASTA_DOWNLOADS = os.path.join(USER_HOME, "Downloads")
PASTA_PRODUTIVIDADE = os.path.join(USER_HOME, "OneDrive - MRV", "Área de Trabalho", "produtividade")

# ==============================================================================
# 2. GERENCIAMENTO DE CREDENCIAIS (Salva em um arquivo config_mrv.json)
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
    # Retorna estrutura padrão vazia
    return {
        "email": "", 
        "senha": "", 
        "senha_malote": "", 
        "chave_api_agilis": "",
        "correios_cod_adm": "",
        "correios_email": "",
        "correios_senha": ""
    }

def salvar_credenciais(email, senha, senha_malote, chave_api_agilis, correios_cod_adm, correios_email, correios_senha):
    """Salva todas as credenciais no arquivo JSON."""
    with open(ARQUIVO_CONFIG, "w") as f:
        json.dump({
            "email": email, 
            "senha": senha, 
            "senha_malote": senha_malote,
            "chave_api_agilis": chave_api_agilis,
            "correios_cod_adm": correios_cod_adm,
            "correios_email": correios_email,
            "correios_senha": correios_senha
        }, f)

# Carrega as variáveis para serem usadas pelos robôs
credenciais = carregar_credenciais()
EMAIL_USER = credenciais.get("email", "")
SENHA_USER = credenciais.get("senha", "")
SENHA_MALOTE = credenciais.get("senha_malote", "")
CHAVE_API_AGILIS = credenciais.get("chave_api_agilis", "")


CORREIOS_COD_ADM = credenciais.get("correios_cod_adm", "")
CORREIOS_EMAIL = credenciais.get("correios_email", "")
CORREIOS_SENHA = credenciais.get("correios_senha", "")

# ==============================================================================
# 3. MODO DE COMPATIBILIDADE
# ==============================================================================
EMAIL_MRV = EMAIL_USER
SENHA_MRV = SENHA_USER
SENHA_MALOTE_MRV = SENHA_MALOTE
API_KEY_AGILIS = CHAVE_API_AGILIS

# ==============================================================================
# 4. RADAR DE PASTAS
# ==============================================================================
if getattr(sys, 'frozen', False):
    PASTA_RAIZ = os.path.dirname(sys.executable)
else:
    PASTA_RAIZ = os.path.dirname(os.path.abspath(__file__))

PASTA_ARQUIVOS = os.path.join(PASTA_RAIZ, "arquivos")