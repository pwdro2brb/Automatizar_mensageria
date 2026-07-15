import pandas as pd
import time
import os
import re
import sys
import unicodedata
from pathlib import Path
sys.path.append(str(Path(__file__).parent.parent)) # <--- ADICIONE ISSO
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

# Importações locais
from treinar_ia import MAPA_PESSOAS, MAPA_ORIGEM_MALOTE
from config import EMAIL_MRV, SENHA_MRV

# --- FUNÇÕES AUXILIARES ---
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

def preencher_remetente_iframe(driver, valor):
    if pd.isna(valor): return
    try:
        xpath = "//li[contains(@class, 'large-text-field')][.//div[contains(., 'Remetente')]]"
        linha = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", linha)
        try: linha.find_element(By.CLASS_NAME, "blank-placeholder").click()
        except: pass
        iframe = WebDriverWait(linha, 5).until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
        driver.switch_to.frame(iframe)
        corpo = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "tinymce")))
        corpo.click(); corpo.send_keys(str(valor))
        driver.switch_to.default_content(); time.sleep(0.5)
    except: driver.switch_to.default_content()

def preencher_destinatario_simples(driver, valor):
    if pd.isna(valor): return
    try:
        xpath = "//li[contains(@class, 'small-text-field')][.//div[contains(., 'Destinatário')]]"
        linha = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath)))
        try:
            placeholder = linha.find_element(By.CLASS_NAME, "blank-placeholder")
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", placeholder)
            placeholder.click()
        except: pass 
        inp = WebDriverWait(linha, 5).until(EC.visibility_of_element_located((By.TAG_NAME, "input")))
        inp.clear(); inp.send_keys(str(valor)); time.sleep(0.5); inp.send_keys(Keys.TAB)
    except Exception as e: 
        print(f"Erro ao preencher Destinatário: {e}")

def preencher_data_id(driver, valor):
    if pd.isna(valor): return
    try:
        linha = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "data-de-recebimento")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", linha)
        try: linha.find_element(By.CLASS_NAME, "blank-placeholder").click()
        except: linha.click()
        inp = WebDriverWait(linha, 5).until(EC.visibility_of_element_located((By.TAG_NAME, "input")))
        inp.clear(); time.sleep(0.5); inp.send_keys(str(valor)); time.sleep(0.5); inp.send_keys(Keys.TAB)
    except: pass

def preencher_codigo(driver, valor):
    if pd.isna(valor): return
    try:
        xpath = "//div[contains(@class, 'blank-placeholder')][contains(., 'Agilis')]"
        div = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", div)
        div.click()
        actions = ActionChains(driver); actions.send_keys(str(valor)); actions.send_keys(Keys.TAB); actions.perform()
    except: pass

def preencher_malote_iframe(driver, id_elemento, valor):
    if pd.isna(valor): return
    try:
        linha = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, id_elemento)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", linha)
        try: linha.find_element(By.CLASS_NAME, "blank-placeholder").click()
        except: pass
        iframe = WebDriverWait(linha, 5).until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
        driver.switch_to.frame(iframe)
        corpo = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "tinymce")))
        corpo.click(); corpo.send_keys(str(valor))
        driver.switch_to.default_content(); time.sleep(0.5)
    except: driver.switch_to.default_content()


# ==============================================================================
# FUNÇÃO PRINCIPAL
# ==============================================================================
def executar_inclusao():
    
    # 1. Lógica de Pastas
    pasta_base = Path(os.path.dirname(os.path.abspath(__file__)))
    pasta_encomendas = Path(__file__).parent.parent / "arquivos" / "encomendas"
    
    if not pasta_encomendas.exists():
        pasta_encomendas.mkdir(parents=True, exist_ok=True)
        print(f"📁 Pasta 'arquivos_encomendas' criada automaticamente.")
        
    caminho_planilha = pasta_encomendas / "encomendas.xlsx"
    
    if not caminho_planilha.exists():
        raise FileNotFoundError(f"A planilha 'encomendas.xlsx' não foi encontrada!\nPor favor, coloque o arquivo dentro da pasta:\n{pasta_encomendas}")

    # ==========================================================================
    # VALIDAÇÕES DA PLANILHA (Vazia ou Aberta no Excel)
    # ==========================================================================
    try:
        tabela = pd.read_excel(caminho_planilha)
        tabela = tabela.dropna(how='all')
        
        if tabela.empty:
            # Abre o arquivo no Excel automaticamente para o usuário
            os.startfile(caminho_planilha)
            raise RuntimeError("A planilha 'encomendas.xlsx' está vazia!\n\nO arquivo foi aberto automaticamente para você. Preencha os dados, salve, feche o Excel e tente novamente.")
            
        print(f"✅ Excel carregado com {len(tabela)} encomendas.")
        
    except PermissionError:
        # Captura o erro de "Access Denied" quando o arquivo está aberto
        raise RuntimeError("O arquivo 'encomendas.xlsx' está aberto no Excel!\n\nPor favor, feche o arquivo e rode o processo novamente.")
        
    except Exception as e:
        if isinstance(e, RuntimeError):
            raise e # Repassa o erro amigável para o popup
        raise RuntimeError(f"Erro ao ler a planilha encomendas.xlsx: {e}")
    
    print("\n--- INICIANDO ROBÔ DE ENCOMENDAS RÁPIDAS ---")
    
    # 2. Configura o Chrome para não fechar no final
    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)
    
    try:
        driver = webdriver.Chrome(options=chrome_options) 
        driver.get("https://podio.com/login")
        driver.maximize_window()
    except Exception as e:
        raise RuntimeError(f"Erro ao iniciar o Chrome. Verifique seu chromedriver. {e}")

    # --- Etapa 1: Aceitar Cookies E Clicar no Login da Microsoft ---
    try:
        print("Aguardando página de login...")
        try:
            print("Procurando o banner de cookies (OneTrust)...")
            accept_cookies_id = "onetrust-accept-btn-handler"
            cookie_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, accept_cookies_id)))
            print("Banner de cookies encontrado. Clicando em 'Aceitar'...")
            cookie_button.click()
            time.sleep(1) 
        except TimeoutException:
            print("Banner de cookies não encontrado ou já aceito. Continuando...")

        microsoft_login_xpath = "//a[@data-provider='live']"
        print("Procurando o botão de login da Microsoft...")
        microsoft_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, microsoft_login_xpath)))
        print("Clicando no botão da Microsoft...")
        microsoft_button.click()

    except Exception as e:
        driver.quit()
        raise RuntimeError(f"Erro na Etapa 1 (Cookies/Login): {e}")

    # --- Etapa 1.5: Lidar com o login da Microsoft ---
    try:
        print("Aguardando a nova janela/aba de login da Microsoft...")
        WebDriverWait(driver, 12).until(EC.number_of_windows_to_be(2))

        popup_window = None
        original_window = driver.current_window_handle 
        for window_handle in driver.window_handles:
            if window_handle != original_window:
                popup_window = window_handle
                break
                
        driver.switch_to.window(popup_window)
        print("Foco mudado para a janela de login da Microsoft.")
                
        print("Aguardando a tela de login da Microsoft...")
        
        email_field_microsoft = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "i0116")))
        print("Preenchendo e-mail da Microsoft...")
        email_field_microsoft.send_keys(EMAIL_MRV)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9"))).click()
        
        password_field_microsoft = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "i0118")))
        print("Preenchendo senha da Microsoft...")
        password_field_microsoft.send_keys(SENHA_MRV)
        
        print("Procurando o botão 'Entrar'...")
        tentativas = 0
        clicado_entrar = False
        while not clicado_entrar and tentativas < 5:
            try:
                entrar_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
                entrar_button.click()
                clicado_entrar = True
                print("Botão 'Entrar' clicado.")
            except StaleElementReferenceException:
                tentativas += 1; time.sleep(0.5)
        if not clicado_entrar: raise Exception("Falha ao clicar em Entrar")

        print("!!! AÇÃO MANUAL NECESSÁRIA !!!")
        print("Aguardando aprovação do MFA no seu celular (até 180s)...")
        
        tentativas = 0
        clicado_manter = False
        while not clicado_manter and tentativas < 5:
            try:
                keep_logged_in_button = WebDriverWait(driver, 180).until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
                keep_logged_in_button.click() 
                clicado_manter = True
                print("MFA Aprovado! Botão 'Manter conectado' clicado.")
            except StaleElementReferenceException:
                tentativas += 1; time.sleep(0.5)
            except TimeoutException:
                 print("Erro: Timeout após 180s. Você não aprovou o MFA a tempo?")
                 clicado_manter = False; break
        if not clicado_manter: raise Exception("Falha ao clicar em Manter Conectado")

        print("Login da Microsoft concluído na janela pop-up.")
        
        print("Aguardando janela pop-up fechar...")
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(1))
        
        nova_janela_principal = driver.window_handles[0]
        driver.switch_to.window(nova_janela_principal)
        print("Foco retornado para a janela principal do Podio.")
        print("Página principal carregada com sucesso.")
        
    except Exception as e:
        driver.quit()
        raise RuntimeError(f"Erro durante o login na Microsoft (Etapa 1.5): {e}")

    # --- LOOP PRINCIPAL DE PREENCHIMENTO ---
    url_app = "https://podio.com/mrvcombr/processos-e-informacoes-csc-teste/apps/mensageria/items/new"
    driver.get(url_app)

    for i, linha in tabela.iterrows():
        codigo_rastreio = str(linha.get('Codigo', '')).strip()
        codigo_upper = codigo_rastreio.upper()
        print(f"\nLinha {i+2}: {codigo_rastreio}")
        
        if "items/new" not in driver.current_url: driver.get(url_app)
        else: driver.refresh()
        
        try: WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "category-color-bg")))
        except: driver.get(url_app)

        agora = datetime.now()
        str_data_formatada = agora.strftime('%d/%m/%Y %H:%M')
        
        categoria_excel = str(linha.get('Categoria', '')).strip() 
        categoria_final = ""
        
        if "Aditivo" in categoria_excel or "aditivo" in categoria_excel:
            categoria_final = "Aditivo" 
            print(f" -> Categoria: ADITIVO (Lido do Excel)")
        else:
            if agora.hour >= 12:
                categoria_final = "Remessa tarde"
            else:
                categoria_final = "Remessa manha"
            print(f" -> Categoria: {categoria_final} (Automático por Horário: {agora.hour}h)")

        is_malote = False
        
        tem_origem = pd.notna(linha.get('Origem do malote')) and str(linha.get('Origem do malote')).strip() != ""
        tem_numero = pd.notna(linha.get('Número do malote')) and str(linha.get('Número do malote')).strip() != ""
        
        if tem_origem or tem_numero:
            is_malote = True
        elif codigo_upper != "SEM RASTREIO" and (re.search(r'\s', codigo_rastreio) or codigo_rastreio.isdigit()):
            is_malote = True

        if is_malote:
            tipo_envio_calculado = "Malote"
            responsavel_malote = "N/A"
            origem_malote = descobrir_origem_malote(linha.get('Remetente'))
            if not origem_malote: 
                origem_malote = linha.get('Origem do malote')
                
            print(f" -> MALOTE detectado. Origem: {origem_malote}")
        else:
            tipo_envio_calculado = "SEDEX/PAC"
            print(" -> SEDEX/PAC detectado")

        try: driver.find_element(By.XPATH, f"//li[contains(@class, 'category-color-bg')][contains(., '{tipo_envio_calculado}')]").click()
        except: pass
        
        try: driver.find_element(By.XPATH, f"//li[contains(@class, 'category-color-bg')][contains(., '{categoria_final}')]").click()
        except: print(f"Erro ao clicar na categoria: {categoria_final}")

        preencher_codigo(driver, codigo_rastreio)
        preencher_remetente_iframe(driver, linha.get('Remetente'))
        
        destino_calc = descobrir_destino(linha.get('Remetente'))
        if not destino_calc:
            destino_calc = linha.get('Destinatario') 
        
        if destino_calc and not pd.isna(destino_calc):
            print(f" -> Destino: {destino_calc}")
            preencher_destinatario_simples(driver, destino_calc)
        else:
            print(" -> Destino não encontrado (Vazio).")

        if is_malote:
            preencher_malote_iframe(driver, "origem-do-malote", origem_malote)
            preencher_malote_iframe(driver, "responsavel-pelo-envio-do-malote", responsavel_malote)
            
            num_malote = linha.get('Número do malote')
            preencher_malote_iframe(driver, "numero-do-malote", num_malote)

        preencher_data_id(driver, str_data_formatada)

        print("Salvando...")
        try:
            driver.find_element(By.XPATH, "//button[contains(., 'Salvar Rastreio')]").click()
            time.sleep(2)
            print("Sucesso!")
        except: print("Erro ao salvar")

    print("\n✅ Fim do processamento! O navegador permanecerá aberto.")

if __name__ == "__main__":
    executar_inclusao()
