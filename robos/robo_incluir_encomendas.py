import pandas as pd
import requests
import os
import re
import sys
import unicodedata
import urllib3
from pathlib import Path
from datetime import datetime

# Desativa os avisos de SSL no console do Hub
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Garante que o script consiga importar o config e treinar_ia da pasta raiz do Hub
sys.path.append(str(Path(__file__).parent.parent))
import config
from treinar_ia import MAPA_PESSOAS, MAPA_ORIGEM_MALOTE

# ==============================================================================
# CONFIGURAÇÕES DO APLICATIVO (Lidas dinamicamente do Config do Hub)
# ==============================================================================

APP_ID = int(config.PODIO_APP_ID)
APP_TOKEN = str(config.PODIO_APP_TOKEN).strip()

# --- FUNÇÕES AUXILIARES ORIGINAIS ---
def normalizar_texto(texto):
    if not isinstance(texto, str): return str(texto).lower()
    return unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII').lower()

def descobrir_origem_malote(texto_remetente):
    if pd.isna(texto_remetente): return None
    texto_limpo = normalizar_texto(texto_remetente)
    for chave, origem in MAPA_ORIGEM_MALOTE.items():
        if chave in texto_limpo: return origem
    return None

def descobrir_destino(texto_remetente):
    if pd.isna(texto_remetente): return None
    texto_limpo = normalizar_texto(texto_remetente)
    for chave, valor in MAPA_PESSOAS.items():
        if chave in texto_limpo: return valor
    return None

# --- FUNÇÕES DE COMUNICAÇÃO COM A API DO PODIO ---
def obter_token_acesso(client_id, client_secret):
    """Autentica no Podio usando App Authentication (Bypassa o MFA da Microsoft)"""
    url = "https://api.podio.com/oauth/token/v2"
    
    # --- LINHAS DE DIAGNÓSTICO (DEBUG) ---
    id_mascarado = f"{client_id[:5]}...{client_id[-5:]}" if len(client_id) > 10 else "Muito curto ou vazio"
    secret_mascarado = f"{client_secret[:5]}...{client_secret[-5:]}" if len(client_secret) > 10 else "Muito curto ou vazio"
    app_token_mascarado = f"{APP_TOKEN[:5]}...{APP_TOKEN[-5:]}" if len(APP_TOKEN) > 10 else "Muito curto ou vazio"
    
    print(f" -> [DIAGNÓSTICO] Lendo Client ID do JSON: {id_mascarado}")
    print(f" -> [DIAGNÓSTICO] Lendo Client Secret do JSON: {secret_mascarado}")
    print(f" -> [DIAGNÓSTICO] Lendo APP_ID: {APP_ID}")
    print(f" -> [DIAGNÓSTICO] Lendo APP_TOKEN: {app_token_mascarado}")
    
    payload = {
        "grant_type": "app",
        "app_id": APP_ID,
        "app_token": APP_TOKEN,
        "client_id": client_id,
        "client_secret": client_secret
    }
    
    response = requests.post(url, json=payload, verify=False)
    
    if response.status_code != 200:
        print(f"\n❌ Erro de Autenticação no Podio (Código {response.status_code}):")
        print(f"Detalhes do erro: {response.text}\n")
        
    response.raise_for_status()
    return response.json()["access_token"]

def mapear_ids_categorias(token):
    """Busca a estrutura do App no Podio e mapeia os IDs das opções de categoria automaticamente"""
    url = f"https://api.podio.com/app/{APP_ID}"
    headers = {"Authorization": f"OAuth2 {token}"}
    response = requests.get(url, headers=headers, verify=False)
    response.raise_for_status()
    
    app_data = response.json()
    mapeamento = {}
    
    for field in app_data.get("fields", []):
        if field["type"] == "category":
            field_external_id = field["external_id"]
            mapeamento[field_external_id] = {}
            for option in field.get("config", {}).get("settings", {}).get("options", []):
                if option["status"] == "active":
                    mapeamento[field_external_id][option["text"].strip()] = option["id"]
    return mapeamento

def criar_item_podio(token, dados_item):
    """Cria um novo registro no aplicativo do Podio via API"""
    url = f"https://api.podio.com/item/app/{APP_ID}/"
    headers = {
        "Authorization": f"OAuth2 {token}",
        "Content-Type": "application/json"
    }
    
    payload = {
        "fields": {
            "numero-do-chamado-agilisrastreio": dados_item["codigo"],
            "remetente": dados_item["remetente"],
            "destinatario-2": dados_item["destinatario"],
            "rastreio": [dados_item["tipo_envio_id"]],
            "categoria": [dados_item["categoria_id"]],
            "data-de-recebimento": dados_item["data_recebimento"]
        }
    }
    
    if dados_item["is_malote"]:
        if dados_item["origem_malote"]:
            payload["fields"]["origem-do-malote"] = dados_item["origem_malote"]
        if dados_item["numero_malote"]:
            payload["fields"]["numero-do-malote"] = dados_item["numero_malote"]
        payload["fields"]["responsavel-pelo-envio-do-malote"] = "N/A"

    response = requests.post(url, json=payload, headers=headers, verify=False)
    if response.status_code not in [200, 201]:
        raise RuntimeError(f"Erro ao salvar no Podio: {response.text}")

# ==============================================================================
# FUNÇÃO PRINCIPAL (Chamada pelo Hub de Automações)
# ==============================================================================
def executar_inclusao():
    print("[PROGRESSO: 2]")
    
    # 1. Lógica de Pastas
    pasta_encomendas = os.path.join(config.PASTA_ARQUIVOS, "encomendas")
    
    if not os.path.exists(pasta_encomendas):
        os.makedirs(pasta_encomendas, exist_ok=True)
        print(f"📁 Pasta 'encomendas' criada automaticamente.")
        
    caminho_planilha = os.path.join(pasta_encomendas, "encomendas.xlsx")
    
    if not os.path.exists(caminho_planilha):
        raise FileNotFoundError(f"A planilha 'encomendas.xlsx' não foi encontrada!\nPor favor, coloque o arquivo dentro da pasta:\n{pasta_encomendas}")

    # ==========================================================================
    # VALIDAÇÕES DA PLANILHA (Vazia ou Aberta no Excel)
    # ==========================================================================
    try:
        tabela = pd.read_excel(caminho_planilha)
        tabela = tabela.dropna(how='all')
        
        if tabela.empty:
            os.startfile(caminho_planilha)
            raise RuntimeError("A planilha 'encomendas.xlsx' está vazia!\n\nO arquivo foi aberto automaticamente para você. Preencha os dados, salve, feche o Excel e tente novamente.")
            
        print(f"✅ Excel carregado com {len(tabela)} encomendas.")
        
    except PermissionError:
        raise RuntimeError("O arquivo 'encomendas.xlsx' está aberto no Excel!\n\nPor favor, feche o arquivo e rode o processo novamente.")
        
    except Exception as e:
        if isinstance(e, RuntimeError):
            raise e
        raise RuntimeError(f"Erro ao ler a planilha encomendas.xlsx: {e}")

    # ==========================================================================
    # VALIDAÇÃO E CARREGAMENTO DAS CREDENCIAIS DO CONFIG.PY
    # ==========================================================================
    client_id = getattr(config, "PODIO_CLIENT_ID", "")
    client_secret = getattr(config, "PODIO_CLIENT_SECRET", "")

    if not client_id or not client_secret:
        raise RuntimeError(
            "Chaves da API do Podio não configuradas!\n\n"
            "Por favor, acesse a aba 'Configurações' no Hub, preencha o 'Podio Client ID' "
            "e o 'Podio Client Secret', salve e tente novamente."
        )

    print("[PROGRESSO: 10]")
    print("Conectando à API do Podio...")
    try:
        token = obter_token_acesso(client_id, client_secret)
        print("🔑 Autenticado com sucesso!")
    except Exception as e:
        raise RuntimeError(f"Falha na autenticação com o Podio. Verifique suas chaves de API nas Configurações. Erro: {e}")

    print("Mapeando campos de categoria...")
    try:
        mapa_categorias = mapear_ids_categorias(token)
    except Exception as e:
        raise RuntimeError(f"Falha ao ler estrutura de categorias do Podio: {e}")

    total_linhas = len(tabela)
    print("[PROGRESSO: 15]")

    # ==========================================================================
    # LOOP PRINCIPAL DE PREENCHIMENTO VIA API
    # ==========================================================================
    for i, linha in tabela.iterrows():
        codigo_rastreio = str(linha.get('Codigo', '')).strip()
        codigo_upper = codigo_rastreio.upper()
        print(f"\nProcessando linha {i+2}: {codigo_rastreio}")
        
        agora = datetime.now()
        str_data_formatada = agora.strftime('%d/%m/%Y %H:%M')
        
        # Lógica de Categoria
        categoria_excel = str(linha.get('Categoria', '')).strip()
        if "Aditivo" in categoria_excel or "aditivo" in categoria_excel:
            categoria_final = "Aditivo"
            print(f" -> Categoria: ADITIVO (Lido do Excel)")
        else:
            categoria_final = "Remessa tarde" if agora.hour >= 12 else "Remessa manha"
            print(f" -> Categoria: {categoria_final} (Automático por Horário: {agora.hour}h)")

        # Lógica de Malote
        is_malote = False
        tem_origem = pd.notna(linha.get('Origem do malote')) and str(linha.get('Origem do malote')).strip() != ""
        tem_numero = pd.notna(linha.get('Número do malote')) and str(linha.get('Número do malote')).strip() != ""
        
        if tem_origem or tem_numero:
            is_malote = True
        elif codigo_upper != "SEM RASTREIO" and (re.search(r'\s', codigo_rastreio) or codigo_rastreio.isdigit()):
            is_malote = True

        if is_malote:
            tipo_envio_calculado = "Malote"
            origem_malote = descobrir_origem_malote(linha.get('Remetente'))
            if not origem_malote: 
                origem_malote = str(linha.get('Origem do malote', ''))
            print(f" -> MALOTE detectado. Origem: {origem_malote}")
        else:
            tipo_envio_calculado = "SEDEX/PAC"
            print(" -> SEDEX/PAC detectado")
            origem_malote = ""

        # Descoberta de Destino
        destino_calc = descobrir_destino(linha.get('Remetente'))
        if not destino_calc:
            destino_calc = str(linha.get('Destinatario', ''))

        # Busca os IDs numéricos das categorias no mapa dinâmico do Podio
        try:
            tipo_envio_id = mapa_categorias["rastreio"][tipo_envio_calculado]
            categoria_id = mapa_categorias["categoria"][categoria_final]
        except KeyError as e:
            print(f"⚠️ Opção de categoria '{e}' não encontrada no Podio. Pulando registro.")
            continue

        # Monta o dicionário de dados para a API
        dados_item = {
            "codigo": codigo_rastreio,
            "remetente": str(linha.get('Remetente', '')) if pd.notna(linha.get('Remetente')) else "",
            "destinatario": destino_calc,
            "tipo_envio_id": tipo_envio_id,
            "categoria_id": categoria_id,
            "data_recebimento": str_data_formatada,
            "is_malote": is_malote,
            "origem_malote": origem_malote,
            "numero_malote": str(linha.get('Número do malote', '')) if pd.notna(linha.get('Número do malote')) else ""
        }

        # Envia diretamente para o Podio
        try:
            criar_item_podio(token, dados_item)
            print(f" -> Cadastrado com sucesso!")
        except Exception as e:
            print(f" -> ❌ Erro ao cadastrar: {e}")
            raise RuntimeError(f"Erro ao cadastrar item {codigo_rastreio}: {e}")
        
        # Atualiza a barra de progresso do Hub (de 15% a 95%)
        if total_linhas > 0:
            progresso_atual = 15 + int(((i + 1) / total_linhas) * 80)
            print(f"[PROGRESSO: {progresso_atual}]")

    print("[PROGRESSO: 100]")
    print("\n✅ Fim do processamento! Todos os registros foram inseridos via API.")

if __name__ == "__main__":
    executar_inclusao()
