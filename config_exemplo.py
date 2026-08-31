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
if getattr(sys, "frozen", False):
    PASTA_RAIZ = os.path.dirname(sys.executable)
else:
    PASTA_RAIZ = os.path.dirname(os.path.abspath(__file__))

ARQUIVO_CONFIG = os.path.join(
    PASTA_RAIZ,
    "config_mrv.json"
)

PASTA_ARQUIVOS = os.path.join(
    PASTA_RAIZ,
    "arquivos"
)


def carregar_credenciais():
    """Lê o arquivo JSON se ele existir. Se não, retorna vazio."""
    if os.path.exists(ARQUIVO_CONFIG):
        try:
            with open(ARQUIVO_CONFIG, "r") as f:
                return json.load(f)
        except:
            pass
    # Retorna estrutura padrão vazia (com os valores corretos do Podio como padrão)
    return {
        "email": "", 
        "senha": "", 
        "senha_malote": "", 
        "chave_api_agilis": "",
        "correios_cod_adm": " ",
        "correios_email": " ",
        "correios_senha": " ",
        "podio_client_id": "",       
        "podio_client_secret": "",
        "podio_app_id": " ",
        "podio_app_token": " ",
        "email_uber": "",
        "senha_uber": ""
    }

def salvar_credenciais(
    email,
    senha,
    senha_malote,
    chave_api_agilis,
    correios_cod_adm,
    correios_email,
    correios_senha,
    podio_client_id="",
    podio_client_secret="",
    podio_app_id="",
    podio_app_token="",
    email_uber="",
    senha_uber=""
):
    """Salva todas as credenciais no arquivo JSON."""

    dados = {
        "email": email,
        "senha": senha,
        "senha_malote": senha_malote,
        "chave_api_agilis": chave_api_agilis,
        "correios_cod_adm": correios_cod_adm,
        "correios_email": correios_email,
        "correios_senha": correios_senha,
        "podio_client_id": podio_client_id,
        "podio_client_secret": podio_client_secret,
        "podio_app_id": podio_app_id,
        "podio_app_token": podio_app_token,
        "email_uber": email_uber,
        "senha_uber": senha_uber
    }

    with open(
        ARQUIVO_CONFIG,
        "w",
        encoding="utf-8"
    ) as arquivo:
        json.dump(
            dados,
            arquivo,
            ensure_ascii=False,
            indent=2
        )

# Carrega as variáveis para serem usadas pelos robôs
credenciais = carregar_credenciais()
EMAIL_USER = credenciais.get("email", "")
SENHA_USER = credenciais.get("senha", "")
SENHA_MALOTE = credenciais.get("senha_malote", "")
CHAVE_API_AGILIS = credenciais.get("chave_api_agilis", "")

# Novas variáveis dos Correios (com valores padrão caso não existam no JSON)
CORREIOS_COD_ADM = credenciais.get("correios_cod_adm", " ")
CORREIOS_EMAIL = credenciais.get("correios_email", " ")
CORREIOS_SENHA = credenciais.get("correios_senha", " ")

# --- NOVAS VARIÁVEIS DA API DO PODIO ---
PODIO_CLIENT_ID = credenciais.get("podio_client_id", "")       
PODIO_CLIENT_SECRET = credenciais.get("podio_client_secret", "") 
PODIO_APP_ID = credenciais.get("podio_app_id", " ")
PODIO_APP_TOKEN = credenciais.get("podio_app_token", " ")

# --- NOVAS PARA ACESSO A PLATAFORMA UBER ---
EMAIL_UBER = credenciais.get("email_uber", "")
SENHA_UBER = credenciais.get("senha_uber", "")

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
