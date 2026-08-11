import pandas as pd
import os
import re
import time
import requests
import urllib3
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException

# Importa as configurações globais do Hub
import config

# Desativa avisos de SSL corporativo se necessário
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ==========================================
# 0. CONFIGURAÇÕES GERAIS
# ==========================================
TECNICOS_VALIDOS = [
    "alfredo henrique goncalves pereira",
    "pedro henrique soares silva",
    "matheus silva de lemos",
    "joao vitor barbosa fernandes",
    "gabriel figueiredo emiliano"
]

# ==========================================
# FUNÇÃO DE DESTAQUE (BORDA VERMELHA)
# ==========================================
def destacar_elemento(driver, elemento):
    """Coloca uma borda vermelha ao redor do elemento para visualização."""
    try:
        driver.execute_script("arguments[0].style.border='3px solid red'", elemento)
        time.sleep(0.2)
    except:
        pass

# ==========================================
# 🚀 FUNÇÕES ULTRA-RÁPIDAS VIA API (AGILIS)
# ==========================================
def obter_dados_completos_chamado_api(chamado, api_key):
    """Busca os dados principais e todas as conversas do chamado em uma única chamada de API."""
    base_url = "https://agilis.mrv.com.br"
    headers = {
        "authtoken": api_key,
        "Accept": "application/vnd.manageengine.sdp.v3+json"
    }
    
    texto_acumulado = []
    tecnico_nome = ""
    coletor_custo = ""
    
    try:
        # 1. Buscar dados principais do chamado (Assunto, Descrição, Técnico)
        url_chamado = f"{base_url}/api/v3/requests/{chamado}"
        res_chamado = requests.get(url_chamado, headers=headers, verify=False)
        
        if res_chamado.status_code == 200:
            dados = res_chamado.json().get("request", {})
            
            tecnico_obj = dados.get("technician")
            if tecnico_obj and isinstance(tecnico_obj, dict):
                tecnico_nome = tecnico_obj.get("name", "")
            
            assunto = dados.get("subject", "")
            descricao_html = dados.get("description", "")
            
            texto_acumulado.append(assunto)
            if descricao_html:
                desc_limpa = re.sub(r'<[^>]+>', ' ', descricao_html)
                texto_acumulado.append(desc_limpa)
                
                match_cc = re.search(r'\*?\s*Coletor de Custo ADM:\s*([A-Za-z0-9]+)', desc_limpa, re.IGNORECASE)
                if match_cc:
                    coletor_custo = match_cc.group(1).strip().upper()
        else:
            return None, "", ""
            
        # 2. Buscar todas as conversas/interações internas
        url_conversas = f"{base_url}/api/v3/requests/{chamado}/conversations"
        res_conversas = requests.get(url_conversas, headers=headers, verify=False)
        
        if res_conversas.status_code == 200:
            conversas = res_conversas.json().get("conversations", [])
            for conv in conversas:
                content_url = conv.get("content_url")
                if content_url:
                    full_content_url = content_url if content_url.startswith("http") else f"{base_url}{content_url}"
                    res_content = requests.get(full_content_url, headers=headers, verify=False)
                    if res_content.status_code == 200:
                        conv_data = res_content.json()
                        notification = conv_data.get("notification", {})
                        conteudo_html = notification.get("description", "")
                        if conteudo_html:
                            conteudo_limpo = re.sub(r'<[^>]+>', ' ', conteudo_html)
                            texto_acumulado.append(conteudo_limpo)
                            
        texto_completo = " ".join(texto_acumulado)
        texto_completo = " ".join(texto_completo.split())
        
        return tecnico_nome, coletor_custo, texto_completo
        
    except Exception:
        return None, "", ""

def validar_chamado_no_agilis_api(chamado, etiqueta_planilha, cc_planilha, api_key):
    """Valida chamados normais procurando pelo CÓDIGO DE RASTREIO via API."""
    try:
        print(f"   [API] Pesquisando Chamado: {chamado} | Rastreio: {etiqueta_planilha}")
        tecnico_nome, cc_web, texto_completo = obter_dados_completos_chamado_api(chamado, api_key)
        
        if tecnico_nome is None:
            return False, cc_planilha, "Erro na validacao"
            
        tecnico_encontrado = any(t.lower() in tecnico_nome.lower() for t in TECNICOS_VALIDOS)
        if not tecnico_encontrado:
            print("      [DESCARTADO] Tecnico responsavel nao autorizado ou nao encontrado.")
            return False, cc_planilha, "Tecnico nao autorizado"
            
        novo_cc = cc_planilha
        if cc_web:
            cc_planilha_clean = str(cc_planilha).strip().upper()
            if cc_web != cc_planilha_clean:
                novo_cc = cc_web
                
        etiqueta_limpa = str(etiqueta_planilha).replace(" ", "").lower()
        texto_limpo = texto_completo.replace(" ", "").lower()
        
        if etiqueta_limpa in texto_limpo:
            print(f"      [SUCESSO] Tecnico e Rastreio ({etiqueta_planilha}) localizados!")
            return True, novo_cc, "Localizado"
        else:
            print(f"      [AVISO] Tecnico encontrado, mas o rastreio {etiqueta_planilha} NAO consta no chamado.")
            return False, cc_planilha, "Rastreio nao localizado"
            
    except Exception:
        return False, cc_planilha, "Erro na validacao"

def validar_chamado_por_nome_agilis_api(chamado, nome_procurado, cc_planilha, api_key):
    """Valida chamados procurando apenas pelo NOME DO DESTINATÁRIO via API."""
    try:
        print(f"   [API] -> Abrindo chamado candidato: {chamado} | Procurando nome: {nome_procurado}")
        tecnico_nome, cc_web, texto_completo = obter_dados_completos_chamado_api(chamado, api_key)
        
        if tecnico_nome is None:
            return False, cc_planilha, "Erro na validacao"
            
        tecnico_encontrado = any(t.lower() in tecnico_nome.lower() for t in TECNICOS_VALIDOS)
        if not tecnico_encontrado:
            print("      [DESCARTADO] Tecnico responsavel nao autorizado.")
            return False, cc_planilha, "Tecnico nao autorizado"
            
        novo_cc = cc_planilha
        if cc_web:
            cc_planilha_clean = str(cc_planilha).strip().upper()
            if cc_web != cc_planilha_clean:
                novo_cc = cc_web
                
        nome_limpo = str(nome_procurado).replace(" ", "").lower()
        texto_limpo = texto_completo.replace(" ", "").lower()
        
        if nome_limpo in texto_limpo:
            print(f"      [SUCESSO] Destinatario '{nome_procurado}' confirmado no chamado!")
            return True, novo_cc, "Localizado (Por Nome)"
        else:
            print(f"      [AVISO] Destinatario '{nome_procurado}' NAO consta neste chamado.")
            return False, cc_planilha, "Nome nao localizado"
            
    except Exception:
        return False, cc_planilha, "Erro na validacao"

def busca_profunda_agilis_api(chamado, etiquetas_pendentes, api_key):
    """Busca profunda de múltiplas etiquetas pendentes em um chamado via API de forma inteligente (Lazy)."""
    base_url = "https://agilis.mrv.com.br"
    headers = {
        "authtoken": api_key,
        "Accept": "application/vnd.manageengine.sdp.v3+json"
    }
    
    etiquetas_encontradas = []
    cc_web = ""
    
    try:
        url_chamado = f"{base_url}/api/v3/requests/{chamado}"
        res_chamado = requests.get(url_chamado, headers=headers, verify=False)
        if res_chamado.status_code != 200:
            return [], ""
            
        dados = res_chamado.json().get("request", {})
        
        tecnico_obj = dados.get("technician")
        tecnico_nome = tecnico_obj.get("name", "") if tecnico_obj and isinstance(tecnico_obj, dict) else ""
        tecnico_encontrado = any(t.lower() in tecnico_nome.lower() for t in TECNICOS_VALIDOS)
        if not tecnico_encontrado:
            return [], ""
        
        descricao_html = dados.get("description", "")
        assunto = dados.get("subject", "")
        desc_limpa = re.sub(r'<[^>]+>', ' ', descricao_html) if descricao_html else ""
        texto_acumulado = (assunto + " " + desc_limpa).replace(" ", "").lower()
        
        match_cc = re.search(r'\*?\s*Coletor de Custo ADM:\s*([A-Za-z0-9]+)', desc_limpa, re.IGNORECASE)
        if match_cc:
            cc_web = match_cc.group(1).strip().upper()
        
        for etiqueta in etiquetas_pendentes:
            etiqueta_limpa = str(etiqueta).replace(" ", "").lower()
            if etiqueta_limpa in texto_acumulado:
                etiquetas_encontradas.append(etiqueta)
        
        etiquetas_restantes = [e for e in etiquetas_pendentes if e not in etiquetas_encontradas]
        
        if etiquetas_restantes:
            url_conversas = f"{base_url}/api/v3/requests/{chamado}/conversations"
            res_conversas = requests.get(url_conversas, headers=headers, verify=False)
            if res_conversas.status_code == 200:
                conversas = res_conversas.json().get("conversations", [])
                for conv in conversas:
                    if not etiquetas_restantes:
                        break
                        
                    content_url = conv.get("content_url")
                    if content_url:
                        full_content_url = content_url if content_url.startswith("http") else f"{base_url}{content_url}"
                        res_content = requests.get(full_content_url, headers=headers, verify=False)
                        if res_content.status_code == 200:
                            conv_data = res_content.json()
                            notification = conv_data.get("notification", {})
                            conteudo_html = notification.get("description", "")
                            if conteudo_html:
                                conteudo_limpo = re.sub(r'<[^>]+>', ' ', conteudo_html).replace(" ", "").lower()
                                for etiqueta in list(etiquetas_restantes):
                                    etiqueta_limpa = str(etiqueta).replace(" ", "").lower()
                                    if etiqueta_limpa in conteudo_limpo:
                                        etiquetas_encontradas.append(etiqueta)
                                        etiquetas_restantes.remove(etiqueta)
                                        
        return etiquetas_encontradas, cc_web
    except Exception:
        return [], ""

# ==========================================
# 📅 AUXILIARES DINÂMICOS (DATA E PASTAS)
# ==========================================
def obter_caminho_pasta_dinamico():
    """
    Calcula o caminho da pasta na rede de forma dinâmica considerando o ciclo de faturamento dos Correios.
    Ciclo fecha dia 22/23. 
    - Do dia 23 em diante: os arquivos vão para a pasta do PRÓXIMO mês.
    - Do dia 01 ao dia 22: os arquivos estão na pasta do mês ATUAL.
    """
    base_rede = r"\\Bhz-fls-app1\mrvbh\Gerência Administrativa\Pública\NUCLEO DE CONTRATOS E APOIO A GESTÃO\CONTRATOS\Contratos Serviços\1. CORREIOS\2. Faturamento"
    
    meses_extenso = {
        1: "01 - Janeiro", 2: "02 - Fevereiro", 3: "03 - Março", 4: "04 - Abril",
        5: "05 - Maio", 6: "06 - Junho", 7: "07 - Julho", 8: "08 - Agosto",
        9: "09 - Setembro", 10: "10 - Outubro", 11: "11 - Novembro", 12: "12 - Dezembro"
    }
    
    hoje = datetime.now()
    dia = hoje.day
    
    # 1. Aplica a Regra do Dia 23 (Cutoff do faturamento)
    if dia >= 23:
        # Se for final de dezembro, o próximo mês é Janeiro do ano seguinte
        if hoje.month == 12:
            ano_alvo = hoje.year + 1
            mes_alvo = 1
        else:
            ano_alvo = hoje.year
            mes_alvo = hoje.month + 1
    else:
        ano_alvo = hoje.year
        mes_alvo = hoje.month
        
    caminho_alvo = os.path.join(base_rede, str(ano_alvo), meses_extenso[mes_alvo], "BH")
    
    # 2. Fallback de Segurança: Se a pasta calculada ainda não existir fisicamente na rede
    # (ex: se atrasarem a criação da pasta de Setembro no dia 23/24 de Agosto),
    # o robô retrocede um mês para garantir que encontre uma pasta válida e não quebre.
    if not os.path.exists(caminho_alvo):
        print(f"   ⚠️ Pasta calculada ({meses_extenso[mes_alvo]}) não existe na rede ainda.")
        
        if mes_alvo == 1:
            ano_fallback = ano_alvo - 1
            mes_fallback = 12
        else:
            ano_fallback = ano_alvo
            mes_fallback = mes_alvo - 1
            
        caminho_fallback = os.path.join(base_rede, str(ano_fallback), meses_extenso[mes_fallback], "BH")
        if os.path.exists(caminho_fallback):
            print(f"   🔄 Usando pasta anterior como fallback: {meses_extenso[mes_fallback]}")
            return caminho_fallback
            
    return caminho_alvo

def localizar_arquivos_inteligente(caminho_pasta):
    """Varre a pasta e localiza os arquivos corretos por padrão de nome, ignorando acentos e caixa alta/baixa."""
    if not os.path.exists(caminho_pasta):
        raise FileNotFoundError(f"A pasta de faturamento não foi encontrada: {caminho_pasta}")
        
    arquivos = os.listdir(caminho_pasta)
    
    arquivo_agf = None
    arquivo_consulta = None
    arquivo_agilis = None
    
    for arq in arquivos:
        arq_lower = arq.lower()
        
        # 1. Arquivo AGF (Exatamente 7 dígitos numéricos, ex: 2576410.xlsx)
        if re.match(r"^\d{7}\.xlsx?$", arq):
            arquivo_agf = os.path.join(caminho_pasta, arq)
            
        # 2. Consulta Postal (Contém "consulta" e "postal")
        elif "consulta" in arq_lower and "postal" in arq_lower and arq_lower.endswith(('.xlsx', '.xls')):
            arquivo_consulta = os.path.join(caminho_pasta, arq)
            
        # 3. Relatório Agilis (Contém "relatorio" ou "relatório" e "agilis")
        elif ("relatorio" in arq_lower or "relatório" in arq_lower) and "agilis" in arq_lower and arq_lower.endswith(('.xlsx', '.xls')):
            arquivo_agilis = os.path.join(caminho_pasta, arq)
            
    return arquivo_agf, arquivo_consulta, arquivo_agilis

def obter_ultimos_meses_busca(qtd=3):
    """Gera dinamicamente os nomes dos últimos meses em extenso para busca nos Correios."""
    nomes_meses = [
        "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
        "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
    ]
    hoje = datetime.now()
    resultado = []
    for i in range(qtd):
        mes_idx = (hoje.month - 1 - i) % 12
        resultado.append(nomes_meses[mes_idx])
    return resultado

# ==========================================
# FUNÇÃO: CORREIOS (SEDEX REVERSO)
# ==========================================
def processar_correios_reverso(aba_5, driver, wait):
    print("\n" + "="*50)
    print("INICIANDO VALIDACAO WEB NOS CORREIOS (SEDEX REVERSO)")
    print("="*50)
    
    wait_correios = WebDriverWait(driver, 60)
    
    try:
        driver.get("https://www2.correios.com.br/encomendas/servicosonline/default.cfm?s=true")
        time.sleep(5) 
        
        print("   - Realizando Login nos Correios...")
        
        cod_adm = wait_correios.until(EC.presence_of_element_located((By.ID, "tx_codigo")))
        destacar_elemento(driver, cod_adm)
        cod_adm.clear()
        cod_adm.send_keys(config.CORREIOS_COD_ADM)
        time.sleep(1)
        
        btn_pesquisar = wait_correios.until(EC.presence_of_element_located((By.XPATH, "//input[@value='Pesquisar' and @type='button']")))
        destacar_elemento(driver, btn_pesquisar)
        driver.execute_script("arguments[0].click();", btn_pesquisar)
        time.sleep(3) 
        
        try:
            btn_email = driver.find_element(By.XPATH, "//*[contains(text(), 'e-mail') or contains(text(), 'E-mail') or @value='email']")
            destacar_elemento(driver, btn_email)
            driver.execute_script("arguments[0].click();", btn_email)
            time.sleep(1)
        except:
            pass
            
        email_field = wait_correios.until(EC.presence_of_element_located((By.XPATH, "//input[@type='text' and contains(@name, 'mail')] | //input[@type='email']")))
        destacar_elemento(driver, email_field)
        email_field.clear()
        email_field.send_keys(config.CORREIOS_EMAIL)
        time.sleep(1)
        
        senha_field = driver.find_element(By.XPATH, "//input[@type='password']")
        destacar_elemento(driver, senha_field)
        senha_field.clear()
        senha_field.send_keys(config.CORREIOS_SENHA)
        time.sleep(1)
        
        btn_ok = wait_correios.until(EC.presence_of_element_located((By.XPATH, "//input[@value='Ok' and @type='button']")))
        destacar_elemento(driver, btn_ok)
        driver.execute_script("arguments[0].click();", btn_ok)
        
        print("   - Aguardando o login ser processado...")
        time.sleep(8) 
        
        print("   - Navegando para Logistica Reversa > Consultas Especificas...")
        log_reversa = wait_correios.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Logística Reversa') or contains(text(), 'Logistica Reversa')]")))
        destacar_elemento(driver, log_reversa)
        driver.execute_script("arguments[0].click();", log_reversa)
        time.sleep(3) 
        
        consultas = wait_correios.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@href, 'logisticaReversa/consultas')]")))
        destacar_elemento(driver, consultas)
        driver.execute_script("arguments[0].click();", consultas)
        
        combo_solicitacao = wait_correios.until(EC.presence_of_element_located((By.ID, "col_aut")))
        destacar_elemento(driver, combo_solicitacao)
        Select(combo_solicitacao).select_by_visible_text("AUTORIZAÇÃO DE POSTAGEM")
        
        print("   - Aguardando a pagina atualizar apos selecionar Autorizacao de Postagem...")
        time.sleep(3) 
        
        # 📅 OTIMIZAÇÃO: Meses de busca agora são calculados dinamicamente!
        meses_busca = obter_ultimos_meses_busca(3)
        
        for mes_alvo in meses_busca:
            pendentes = 0
            for idx, row in aba_5.iterrows():
                rastreio = str(row['Etiqueta']).strip()
                status = str(row.get('Validação', '')).strip()
                if status != 'Localizado' and rastreio and rastreio != 'nan':
                    pendentes += 1
                    
            if pendentes == 0:
                print(f"\n   [SUCESSO] Todos os rastreios foram localizados! Encerrando buscas.")
                break 
                
            print(f"\n   >>> Iniciando busca no mes de {mes_alvo} ({pendentes} rastreios pendentes) <<<")
            
            combo_periodo_element = wait_correios.until(EC.presence_of_element_located((By.NAME, "periodo")))
            destacar_elemento(driver, combo_periodo_element)
            combo_periodo = Select(combo_periodo_element)
            
            selecionou_mes = False
            for option in combo_periodo.options:
                if mes_alvo in option.text.upper():
                    option.click()
                    selecionou_mes = True
                    break
                    
            if not selecionou_mes:
                print(f"   [AVISO] Nao achou {mes_alvo} na lista. Pulando para o proximo...")
                continue
                
            print(f"   - Aguardando a pagina atualizar apos selecionar {mes_alvo}...")
            time.sleep(3) 
            
            xpath_checkboxes = "//td[b[contains(text(), 'Dados do Remetente') or contains(text(), 'Dados do Destinatário')]]//input[@type='Checkbox']"
            checkboxes = driver.find_elements(By.XPATH, xpath_checkboxes)
            
            for cb in checkboxes:
                if not cb.is_selected():
                    driver.execute_script("arguments[0].click();", cb)
            time.sleep(1)
                    
            btn_consulta = driver.find_element(By.XPATH, "//input[@value='Realizar Consulta']")
            destacar_elemento(driver, btn_consulta)
            driver.execute_script("arguments[0].click();", btn_consulta)
            
            print("   - Aguardando a NOVA GUIA de resultados abrir...")
            time.sleep(5) 
            
            driver.switch_to.window(driver.window_handles[-1])
            
            try:
                wait_correios.until(EC.presence_of_element_located((By.XPATH, "//tr[@class='cssLinhaTitulo']")))
                time.sleep(2)
                
                print(f"   - Iniciando a varredura dos rastreios na tabela de {mes_alvo}...")
                
                for idx, row in aba_5.iterrows():
                    rastreio = str(row['Etiqueta']).strip()
                    status_atual = str(row.get('Validação', '')).strip()
                    
                    if status_atual == 'Localizado' or not rastreio or rastreio == 'nan':
                        continue
                        
                    try:
                        xpath_linha = f"//tr[td[position()=5 and contains(text(), '{rastreio}')]]"
                        linha_rastreio = driver.find_element(By.XPATH, xpath_linha)
                        
                        coluna_autorizacao = linha_rastreio.find_element(By.XPATH, "./td[2]")
                        link_auth = coluna_autorizacao.find_element(By.TAG_NAME, "a")
                        num_auth = link_auth.text.strip()
                        
                        print(f"   [ACHOU] Rastreio {rastreio} -> Autorizacao: {num_auth}")
                        
                        aba_5.at[idx, 'N° da Autoriação'] = num_auth
                        aba_5.at[idx, 'Validação'] = "Localizado"
                        
                        destacar_elemento(driver, link_auth)
                        
                        janela_resultados = driver.current_window_handle
                        
                        driver.execute_script("arguments[0].click();", link_auth)
                        time.sleep(3) 
                        
                        driver.switch_to.window(driver.window_handles[-1])
                        
                        try:
                            xpath_solicitante = "//td[b[contains(text(), 'Usuário Solicitante')]]/following-sibling::td"
                            solicitante_elem = wait_correios.until(EC.presence_of_element_located((By.XPATH, xpath_solicitante)))
                            destacar_elemento(driver, solicitante_elem)
                            aba_5.at[idx, 'Responsavel'] = solicitante_elem.text.strip()
                        except:
                            aba_5.at[idx, 'Responsavel'] = "Não encontrado"
                            
                        try:
                            xpath_desc = "//td[b[contains(text(), 'Data de Entrega')]]"
                            desc_elem = driver.find_element(By.XPATH, xpath_desc)
                            destacar_elemento(driver, desc_elem)
                            
                            desc_text = desc_elem.text
                            linha_desc = desc_text.split('\n')[0].strip() 
                            
                            if '|' in linha_desc:
                                cc_bruto = linha_desc.split('|')[-1].strip().upper()
                            else:
                                cc_bruto = linha_desc.strip().upper()
                                
                            match_cc = re.search(r'\b[A-Z0-9]*\d+[A-Z0-9]*\b', cc_bruto)
                            if match_cc:
                                cc_limpo = match_cc.group(0)
                            else:
                                palavras = cc_bruto.split()
                                cc_limpo = palavras[-1] if palavras else ""
                                
                            aba_5.at[idx, 'Coletor de Custo'] = cc_limpo
                        except:
                            aba_5.at[idx, 'Coletor de Custo'] = "Não encontrado"
                        
                        driver.close()
                        driver.switch_to.window(janela_resultados)
                        time.sleep(1)
                        
                    except Exception:
                        pass
                        
            except TimeoutException:
                print(f"   [AVISO] A tabela de resultados nao carregou corretamente para {mes_alvo}.")
                
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
            time.sleep(2)
                
        for idx, row in aba_5.iterrows():
            if str(row.get('Validação', '')) != 'Localizado' and str(row['Etiqueta']).strip() != 'nan':
                aba_5.at[idx, 'Validação'] = 'Não Localizado'
                
    except Exception as e:
        print(f"   [ERRO FATAL NOS CORREIOS] {e}")

# ==========================================
# ⚙️ FUNÇÃO MESTRE: EXECUTAR RATEIO MALOTE
# ==========================================
def executar_rateio_AGF():
    print("[PROGRESSO: 5]")
    print("Iniciando processamento do Rateio AGF...")
    
    api_key = getattr(config, "CHAVE_API_AGILIS", "")
    if not api_key:
        print("❌ ERRO: Chave API do Agilis não configurada! Vá na aba Configurações e salve-a.")
        return False

    # 📅 1. CAMINHOS E ARQUIVOS DINÂMICOS
    try:
        caminho_pasta = obter_caminho_pasta_dinamico()
        print(f"📂 Pasta de trabalho identificada: {caminho_pasta}")
        
        caminho_arquivo_agf, caminho_consulta, caminho_agilis = localizar_arquivos_inteligente(caminho_pasta)
        caminho_novo_arquivo = os.path.join(caminho_pasta, "Rateio_AGF_Separado.xlsx")
        
        # Validações de existência dos arquivos
        if not caminho_arquivo_agf:
            print("❌ ERRO: Arquivo AGF (numérico de 7 dígitos) não encontrado na pasta!")
            return False
        if not caminho_consulta:
            print("❌ ERRO: Arquivo 'Consulta Postal' não encontrado na pasta!")
            return False
        if not caminho_agilis:
            print("❌ ERRO: Arquivo 'Relatório Agilis' não encontrado na pasta!")
            return False
            
        print(f"   ✅ Arquivo AGF localizado: {os.path.basename(caminho_arquivo_agf)}")
        print(f"   ✅ Consulta Postal localizada: {os.path.basename(caminho_consulta)}")
        print(f"   ✅ Relatório Agilis localizado: {os.path.basename(caminho_agilis)}")
        
    except Exception as e:
        print(f"❌ ERRO ao mapear arquivos na pasta: {e}")
        return False

    # 2. LER O ARQUIVO AGF ORIGINAL
    print("Lendo o arquivo AGF original...")
    try:
        df_base = pd.read_excel(caminho_arquivo_agf)
        df_base.columns = df_base.columns.str.strip()
        
        if 'Data de Postagem' not in df_base.columns:
            print(f"\n[ERRO FATAL] A coluna 'Data de Postagem' nao foi encontrada na planilha AGF!")
            print(f"-> Colunas lidas: {df_base.columns.tolist()}")
            return False
            
    except PermissionError:
        print(f"❌ ERRO: O arquivo '{os.path.basename(caminho_arquivo_agf)}' esta aberto. Feche-o e tente novamente.")
        return False
    except Exception as e:
        print(f"❌ ERRO ao ler planilha base: {e}")
        return False

    df_transacoes = df_base.dropna(subset=['Data de Postagem']).copy()

    # 3. SEPARAÇÃO DAS ABAS
    print("Separando as abas...")
    aba_1 = df_transacoes[df_transacoes['Titular do Cartao de Postagem'] == 'MRVHBH3022 TELEGRAMA']
    aba_2 = df_transacoes[df_transacoes['Titular do Cartao de Postagem'] == 'MRVHBH3016 TELEG GST COBR RENEG']
    aba_3 = df_transacoes[df_transacoes['Titular do Cartao de Postagem'] == 'MRVHBH3015']

    resp_4_lista = ['MRV ENGENHARIA E PARTICIPACOES S.A', 'MRV SEDE']
    df_resp_4_total = df_transacoes[df_transacoes['Titular do Cartao de Postagem'].isin(resp_4_lista)]

    aba_5 = df_resp_4_total[df_resp_4_total['Servico'].str.contains('REVERSO', na=False, case=False)].copy()
    aba_7 = df_resp_4_total[df_resp_4_total['Servico'].str.contains('MALOTE', na=False, case=False)]
    aba_4 = df_resp_4_total[~df_resp_4_total['Servico'].str.contains('REVERSO|MALOTE', na=False, case=False)].copy()

    titulares_mapeados = ['MRVHBH3022 TELEGRAMA', 'MRVHBH3016 TELEG GST COBR RENEG', 'MRVHBH3015'] + resp_4_lista
    filtro_titulares = ~df_transacoes['Titular do Cartao de Postagem'].isin(titulares_mapeados)
    filtro_servicos = ~df_transacoes['Servico'].str.contains('REVERSO|MALOTE', na=False, case=False)

    aba_6 = df_transacoes[filtro_titulares & filtro_servicos].copy()

    # PREPARAÇÃO DA ABA 5 (CRIANDO AS COLUNAS)
    if 'Peso Real' in aba_5.columns:
        idx_peso = aba_5.columns.get_loc('Peso Real') + 1
        novas_colunas = ['N° da Autoriação', 'Responsavel', 'Coletor de Custo', 'Validação', 'Valor']
        for i, col in enumerate(novas_colunas):
            if col not in aba_5.columns:
                aba_5.insert(idx_peso + i, col, "")

    # PREPARAÇÃO DA ABA 6 (DIVERSOS)
    print("Preparando a Aba 6 (DIVERSOS)...")
    if 'Peso Real' in aba_6.columns:
        idx_peso_6 = aba_6.columns.get_loc('Peso Real') + 1
        novas_colunas_6 = ['Cartão de Postagem', 'Centro de Custo/Diagrama', 'Validação', 'Valor']
        for i, col in enumerate(novas_colunas_6):
            if col not in aba_6.columns:
                aba_6.insert(idx_peso_6 + i, col, "")

        if 'Numero do Cartao de Postagem' in aba_6.columns:
            aba_6['Cartão de Postagem'] = aba_6['Numero do Cartao de Postagem']

        def extrair_cc_diagrama(texto):
            if pd.isna(texto): return ""
            texto_limpo = str(texto).strip().upper()
            match = re.search(r'\b[A-Z0-9]*\d+[A-Z0-9]*\b', texto_limpo)
            if match: return match.group(0)
            palavras = texto_limpo.split()
            return palavras[-1] if palavras else ""

        if 'Titular do Cartao de Postagem' in aba_6.columns:
            aba_6['Centro de Custo/Diagrama'] = aba_6['Titular do Cartao de Postagem'].apply(extrair_cc_diagrama)

        aba_6['Validação'] = aba_6['Centro de Custo/Diagrama'].apply(lambda cc: "Ok" if cc else "Não localizado")

    # 4. CRUZAMENTO ABA 4 x CONSULTA POSTAL
    print("Iniciando cruzamento da Aba 4 com a Consulta Postal...")
    try:
        df_consulta = pd.read_excel(caminho_consulta)
        df_consulta_reduzido = df_consulta[['Etiqueta', 'Destinatário']].drop_duplicates(subset=['Etiqueta'])
        aba_4 = aba_4.merge(df_consulta_reduzido, on='Etiqueta', how='left')
        aba_4['Nome_Destinatario'] = aba_4['Destinatário']
        
        def extrair_chamado(texto):
            if pd.isna(texto): return ""
            texto_str = str(texto).strip()
            match = re.search(r'(?<!\d)(\d{7})(?!\d)', texto_str)
            return match.group(1) if match else texto_str
                
        aba_4['Chamado'] = aba_4['Destinatário'].apply(extrair_chamado)
        aba_4 = aba_4.drop(columns=['Destinatário'])
    except Exception as e:
        print(f"⚠️ AVISO: Erro no cruzamento com Consulta Postal: {e}")

    for col in ['Centro de Custo/Diagrama', 'Solicitante', 'Validação', 'Valor']:
        if col not in aba_4.columns:
            aba_4[col] = ""

    # 4.2 CRUZAMENTO COM AGILIS (EXCEL)
    print("Iniciando cruzamento com a planilha Agilis...")
    df_agilis = None
    df_agilis_raw = None 

    try:
        df_agilis_raw = pd.read_excel(caminho_agilis)
        df_agilis = df_agilis_raw.copy()
        
        if 'Identificação da solicitação' in df_agilis.columns:
            df_agilis['Identificação da solicitação'] = df_agilis['Identificação da solicitação'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            df_agilis_raw['Identificação da solicitação'] = df_agilis_raw['Identificação da solicitação'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
            if 'Categoria' in df_agilis.columns:
                filtro_cat = df_agilis['Categoria'].astype(str).str.contains('SOLICITAÇÃO DE ENVIO DE CORRESPONDÊNCIA', case=False, na=False)
                df_agilis = df_agilis[filtro_cat].copy()
                
            def tem_tecnico_valido(row):
                linha_texto = " ".join(row.astype(str)).lower()
                return any(tecnico.lower() in linha_texto for tecnico in TECNICOS_VALIDOS)
            
            df_agilis = df_agilis[df_agilis.apply(tem_tecnico_valido, axis=1)].copy()
            
            def extrair_cc(descricao):
                if pd.isna(descricao): return ""
                match = re.search(r'\*\s*Coletor de Custo ADM:\s*([A-Za-z0-9]+)', str(descricao), re.IGNORECASE)
                return match.group(1) if match else ""
                
            if 'Descrição' in df_agilis.columns:
                for index, row in aba_4.iterrows():
                    if 'Localizado' in str(row.get('Validação', '')):
                        continue
                    valor_chamado = row['Chamado']
                    if pd.isna(valor_chamado) or str(valor_chamado).strip() == "": continue
                    valor_chamado = str(valor_chamado).strip()
                    
                    if re.fullmatch(r'\d{7}', valor_chamado):
                        match_agilis = df_agilis[df_agilis['Identificação da solicitação'] == valor_chamado]
                        if not match_agilis.empty:
                            aba_4.at[index, 'Centro de Custo/Diagrama'] = extrair_cc(match_agilis.iloc[0]['Descrição'])
    except Exception as e:
        print(f"⚠️ AVISO: Erro no cruzamento com Agilis: {e}")

    # ==========================================================================
    # ⚡ 4.3 VALIDAÇÃO ULTRA-RÁPIDA VIA API (AGILIS)
    # ==========================================================================
    print("\n" + "="*50)
    print("INICIANDO VALIDAÇÃO ULTRA-RÁPIDA VIA API (AGILIS)")
    print("="*50)
    
    chamados_ja_verificados = []
    total_linhas_aba4 = len(aba_4)

    print("\n--- ETAPA 1: Validando chamados ja preenchidos (Agilis API) ---")
    for index, row in aba_4.iterrows():
        progresso = 10 + int((index / total_linhas_aba4) * 30)
        print(f"[PROGRESSO: {progresso}]")

        chamado = row['Chamado']
        etiqueta = row['Etiqueta']
        cc_atual = row['Centro de Custo/Diagrama']
        status_atual = str(row.get('Validação', '')).strip()
        
        if 'Localizado' in status_atual:
            if pd.notna(chamado) and str(chamado).strip() != "" and re.fullmatch(r'\d{7}', str(chamado).strip()):
                chamados_ja_verificados.append(str(chamado).strip())
            continue
        
        if pd.isna(chamado) or str(chamado).strip() == "" or pd.isna(etiqueta):
            continue

        valor_chamado = str(chamado).strip()

        if re.fullmatch(r'\d{7}', valor_chamado):
            manter, novo_cc, status = validar_chamado_no_agilis_api(valor_chamado, etiqueta, cc_atual, api_key)
            aba_4.at[index, 'Validação'] = status
            if manter:
                aba_4.at[index, 'Centro de Custo/Diagrama'] = novo_cc
            
            if status in ["Tecnico nao autorizado", "Localizado", "Rastreio nao localizado"]:
                chamados_ja_verificados.append(valor_chamado)

    print("\n--- ETAPA 2: Busca Profunda Paralelizada (Agilis API) ---")
    print("[PROGRESSO: 45]")
    
    def obter_pendentes():
        mascara = (~aba_4['Validação'].astype(str).str.startswith('Localizado', na=False)) & (aba_4['Etiqueta'].notna()) & (aba_4['Etiqueta'].str.strip() != '')
        return aba_4.loc[mascara, 'Etiqueta'].tolist()

    etiquetas_pendentes = obter_pendentes()
    if etiquetas_pendentes and df_agilis is not None:
        todos_chamados_agilis = df_agilis['Identificação da solicitação'].dropna().unique()
        chamados_para_pesquisar = [c for c in todos_chamados_agilis if str(c).strip() not in chamados_ja_verificados]
        
        MAX_THREADS = 10
        print(f"   - Pesquisando {len(etiquetas_pendentes)} etiquetas pendentes em {len(chamados_para_pesquisar)} chamados...")
        print(f"   - Executando busca paralela com {MAX_THREADS} conexões simultâneas...")
        
        with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
            futuros = {
                executor.submit(busca_profunda_agilis_api, str(chamado).strip(), etiquetas_pendentes, api_key): str(chamado).strip()
                for chamado in chamados_para_pesquisar
                if re.fullmatch(r'\d{7}', str(chamado).strip())
            }
            
            for futuro in as_completed(futuros):
                chamado_str = futuros[futuro]
                if not etiquetas_pendentes:
                    break
                    
                try:
                    encontrados, cc_achado = futuro.result()
                    if encontrados:
                        chamados_ja_verificados.append(chamado_str)
                        for etiq in encontrados:
                            if etiq in etiquetas_pendentes:
                                idx = aba_4[aba_4['Etiqueta'] == etiq].index[0]
                                aba_4.at[idx, 'Chamado'] = chamado_str
                                aba_4.at[idx, 'Centro de Custo/Diagrama'] = cc_achado
                                aba_4.at[idx, 'Validação'] = "Localizado (Busca Profunda)"
                                print(f"      ✅ Etiqueta {etiq} localizada no chamado {chamado_str} (Busca Profunda)")
                                try:
                                    etiquetas_pendentes.remove(etiq)
                                except ValueError:
                                    pass
                except Exception as e:
                    print(f"      ⚠️ Erro ao processar chamado {chamado_str} na busca profunda: {e}")

    print("\n--- ETAPA 3: Busca Pesada por Nome na Descrição (Agilis API) ---")
    print("[PROGRESSO: 55]")
    for index, row in aba_4.iterrows():
        status_atual = str(row.get('Validação', '')).strip()
        if 'Localizado' in status_atual:
            continue
            
        etiqueta = str(row['Etiqueta']).strip().upper()
        cc_atual = row['Centro de Custo/Diagrama']
        nome_original = str(row.get('Nome_Destinatario', '')).strip()
        
        if not nome_original or nome_original == 'nan' or etiqueta == 'NAN' or etiqueta == '':
            continue
            
        nome_busca = re.sub(r'\d{7}', '', nome_original).replace('-', '').strip()
        
        if len(nome_busca) > 2:
            print(f"\n   [BUSCA PESADA POR NOME] Destinatario: {nome_busca} | Rastreio: {etiqueta}")
            if df_agilis is not None:
                nome_escapado = re.escape(nome_busca)
                match_agilis = df_agilis[df_agilis['Descrição'].astype(str).str.contains(nome_escapado, case=False, na=False)]
                
                encontrou_chamado_valido = False
                
                for _, linha_agilis in match_agilis.iterrows():
                    chamado_candidato = str(linha_agilis['Identificação da solicitação']).strip()
                    if not re.fullmatch(r'\d{7}', chamado_candidato): continue
                    
                    print(f"   -> Testando chamado candidato: {chamado_candidato}")
                    
                    if etiqueta.startswith('BN'):
                        manter, novo_cc, status = validar_chamado_por_nome_agilis_api(chamado_candidato, nome_busca, cc_atual, api_key)
                    else:
                        manter, novo_cc, status = validar_chamado_no_agilis_api(chamado_candidato, etiqueta, cc_atual, api_key)
                    
                    if manter:
                        aba_4.at[index, 'Chamado'] = chamado_candidato
                        aba_4.at[index, 'Centro de Custo/Diagrama'] = novo_cc
                        aba_4.at[index, 'Validação'] = status
                        chamados_ja_verificados.append(chamado_candidato)
                        encontrou_chamado_valido = True
                        break 
                
                if not encontrou_chamado_valido:
                    aba_4.at[index, 'Validação'] = "Não localizado"

    # ==========================================================================
    # 🌐 4.4 PARTE 2: CORREIOS (SEDEX REVERSO)
    # ==========================================================================
    print("\n" + "="*50)
    print("INICIANDO ETAPA DOS CORREIOS (SEDEX REVERSO)")
    print("="*50)
    print("[PROGRESSO: 70]")
    
    try:
        driver = webdriver.Chrome()
        wait_validacao = WebDriverWait(driver, 10)
        processar_correios_reverso(aba_5, driver, wait_validacao)
    except Exception as e:
        print(f"❌ Erro ao processar Correios via Selenium: {e}")
    finally:
        try:
            driver.quit()
            print("Navegador Selenium fechado.")
        except:
            pass

    # ==========================================================================
    # 4.5 ETAPA FINAL: PREENCHIMENTO (SOLICITANTE E NÃO LOCALIZADOS)
    # ==========================================================================
    print("\n--- ETAPA FINAL: Finalizando preenchimentos (Solicitante e Nao Localizados) ---")
    print("[PROGRESSO: 90]")
    col_solicitante = None
    if df_agilis_raw is not None:
        if 'Solicitante' in df_agilis_raw.columns:
            col_solicitante = 'Solicitante'
        elif 'Requerente' in df_agilis_raw.columns:
            col_solicitante = 'Requerente'

    if col_solicitante:
        mapa_solicitantes = df_agilis_raw.set_index('Identificação da solicitação')[col_solicitante].to_dict()
        for idx, row in aba_4.iterrows():
            chamado_atual = str(row['Chamado']).strip()
            if chamado_atual and chamado_atual in mapa_solicitantes:
                aba_4.at[idx, 'Solicitante'] = str(mapa_solicitantes[chamado_atual]).strip()

    for idx, row in aba_4.iterrows():
        val = str(row['Validação']).strip()
        etiq = str(row['Etiqueta']).strip()
        if etiq and etiq != 'nan' and not val.startswith('Localizado'):
            aba_4.at[idx, 'Validação'] = "Não localizado"

    # ==========================================================================
    # 5. SALVAR NO NOVO EXCEL COM AS 7 ABAS
    # ==========================================================================
    print("\nCriando o novo arquivo Excel final...")
    try:
        with pd.ExcelWriter(caminho_novo_arquivo, engine='openpyxl') as writer:
            aba_1.to_excel(writer, sheet_name='1-MRVHBH3022', index=False)
            aba_2.to_excel(writer, sheet_name='2-MRVHBH3016', index=False)
            aba_3.to_excel(writer, sheet_name='3-MRVHBH3015', index=False)
            
            if 'Nome_Destinatario' in aba_4.columns:
                aba_4 = aba_4.drop(columns=['Nome_Destinatario'])
                
            aba_4.to_excel(writer, sheet_name='4-MRV_ENGENHARIA_SEDE', index=False)
            aba_5.to_excel(writer, sheet_name='5-SEDEX_REVERSO', index=False)
            aba_6.to_excel(writer, sheet_name='6-DIVERSOS', index=False)
            aba_7.to_excel(writer, sheet_name='7-MALOTE', index=False)

        print(f"✅ Sucesso! O arquivo final foi salvo em: {caminho_novo_arquivo}")
        print("[PROGRESSO: 100]")
        return True
    except PermissionError:
        print(f"\n❌ [ERRO] O arquivo '{caminho_novo_arquivo}' esta aberto no Excel! Feche-o para salvar.")
        return False
    except Exception as e:
        print(f"❌ Erro ao salvar arquivo final: {e}")
        return False
