import time
import requests
import json
import subprocess
from datetime import datetime, timedelta, timezone
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.keys import Keys
import win32com.client
import os
import re
import traceback

# Importando as credenciais seguras
from config import EMAIL_MRV, SENHA_MRV

# =============================================
# FUNÇÕES AUXILIARES
# =============================================

def aguardar_proximo_ciclo(minutos=10):
    """Espera X minutos usando horário real, não sleep puro."""
    proximo = datetime.now() + timedelta(minutes=minutos)
    print(f"Próxima verificação às {proximo.strftime('%H:%M:%S')}")
    while datetime.now() < proximo:
        time.sleep(30)
    print("Hora de verificar novamente!")

def verificar_sessao_ativa(driver):
    """Verifica se o WebDriver ainda está funcional."""
    try:
        _ = driver.title
        return True
    except:
        return False

def fazer_login(driver, wait):
    """Faz o login completo no Agilis via Microsoft SSO."""
    URL_INICIAL = "https://agilis.mrv.com.br/HomePage.do?view_type=my_view"
    driver.get(URL_INICIAL)
    print(f"Página aberta: {URL_INICIAL}")
    print("Aguardando tela de login...")

    try:
        selector_login_integrado = (By.CSS_SELECTOR, "#saml-div a.sign-saml")
        wait.until(EC.element_to_be_clickable(selector_login_integrado)).click()
        print("Cliquei em 'Login Integrado Microsoft'.")
    except:
        print("Não encontrou pelo CSS. Tentando por XPATH...")
        selector_login_integrado = (By.XPATH, "//*[text()='Login Integrado Microsoft']")
        wait.until(EC.element_to_be_clickable(selector_login_integrado)).click()
        print("Cliquei em 'Login Integrado Microsoft'.")

    email_field = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "i0116"))
    )
    print("Preenchendo e-mail...")
    email_field.send_keys(EMAIL_MRV)
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "idSIButton9"))
    ).click()

    password_field = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "i0118"))
    )
    print("Preenchendo senha...")
    password_field.send_keys(SENHA_MRV)

    print("Clicando em 'Entrar'...")
    tentativas = 0
    clicado = False
    while not clicado and tentativas < 5:
        try:
            btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "idSIButton9"))
            )
            btn.click()
            clicado = True
            print("Botão 'Entrar' clicado.")
        except StaleElementReferenceException:
            tentativas += 1
            time.sleep(0.5)
    if not clicado:
        raise Exception("Falha ao clicar em Entrar")

    print("!!! AÇÃO MANUAL NECESSÁRIA !!!")
    print("Aguardando aprovação do MFA no seu celular (até 180s)...")

    tentativas = 0
    clicado_manter = False
    while not clicado_manter and tentativas < 5:
        try:
            btn_manter = WebDriverWait(driver, 180).until(
                EC.element_to_be_clickable((By.ID, "idSIButton9"))
            )
            btn_manter.click()
            clicado_manter = True
            print("MFA Aprovado! Botão 'Manter conectado' clicado.")
        except StaleElementReferenceException:
            tentativas += 1
            time.sleep(0.5)
        except TimeoutException:
            print("Erro: Timeout após 180s.")
            break
    if not clicado_manter:
        raise Exception("Falha ao clicar em Manter Conectado")

    print("Verificando se há tela intermediária de login...")
    try:
        # CORREÇÃO 2: Reduzido de 10 para 2 segundos. Se não aparecer rápido, segue a vida.
        btn_saml = WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "#saml-div a.sign-saml"))
        )
        btn_saml.click()
        print("Tela intermediária detectada. Cliquei em 'Login Integrado Microsoft' novamente.")
        time.sleep(3)
    except TimeoutException:
        print("Nenhuma tela intermediária. Seguindo normalmente.")

    print("Aguardando página principal do Agilis carregar...")
    try:
        # CORREÇÃO 2: Em vez de procurar um botão que pode não existir, apenas espera a URL mudar. É instantâneo!
        WebDriverWait(driver, 15).until(
            lambda d: "agilis.mrv.com.br" in d.current_url and "login" not in d.current_url.lower()
        )
        print("Página principal carregada!")
    except TimeoutException:
        print("Página principal demorou, mas continuando...")

    print("Login concluído com sucesso!")

def parsear_data_chamado(texto_data):
    """Converte 'May 22, 2026 05:49 PM' para um objeto datetime."""
    try:
        return datetime.strptime(texto_data.strip(), "%b %d, %Y %I:%M %p")
    except ValueError:
        print(f"  [AVISO] Não consegui parsear a data: '{texto_data}'")
        return None

def chamado_ainda_aberto(driver, chamado_id):
    """Verifica via API se o chamado ainda está aberto antes de processar."""
    session = get_session_from_selenium(driver)
    url = f"https://agilis.mrv.com.br/api/v3/requests/{chamado_id}"
    params = {
        "input_data": json.dumps({"fields_required": ["status"]}),
        "SUBREQUEST": "XMLHTTP"
    }
    try:
        data = session.get(url, params=params, timeout=10).json()
        status = data.get("request", {}).get("status", {}).get("name", "")
        return status in ("Open", "Aberto")
    except:
        return True  # Em caso de dúvida, tenta processar

def processar_chamado(driver, wait, max_tentativas=3):
    """Executa o fluxo completo de fechamento de um chamado."""
    for tentativa in range(1, max_tentativas + 1):
        try:
            if tentativa > 1:
                print(f"\n    🔄 Tentativa {tentativa}/{max_tentativas}...")
                print("    -> Atualizando a página do chamado...")
                driver.refresh()
                time.sleep(5)

            print("    -> Clicando em 'Responder a todos'...")
            btn_responder = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[data-button-role='reply']")
            ))
            btn_responder.click()
            time.sleep(2)

            print("    -> Abrindo dropdown de Modelo de Resposta...")
            dropdown_modelo = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "#s2id_rt_requests .select2-choice")
            ))
            dropdown_modelo.click()
            time.sleep(1.5)

            print("    -> Selecionando 'Encomenda não recebida na mensageria.'...")
            opcao_encomenda = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(@class,'select2-result-label')]//div[@title='Privado' and contains(.,'Encomenda não recebida na mensageria')]")
            ))
            opcao_encomenda.click()
            time.sleep(1.5)

            print("    -> Aceitando pop-up de confirmação (OK)...")
            try:
                alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
                alert.accept()
                print("    -> Pop-up aceito.")
            except TimeoutException:
                print("    -> Nenhum alert apareceu, tentando botão OK na página...")
                try:
                    btn_ok = WebDriverWait(driver, 5).until(EC.element_to_be_clickable(
                        (By.XPATH, "//button[text()='OK']")
                    ))
                    btn_ok.click ()
                except:
                    print("    -> Nenhum botão OK encontrado, continuando...")
            time.sleep(2)

            print("    -> Procurando botão de status...")
            try:
                status_span = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "span[name='viewWOStatus']")
                ))
                status_span.click()
                print("    -> Status clicado.")
                time.sleep(1.5)
            except TimeoutException:
                print(f"    ⚠️ Botão de status não encontrado! (Bug do Agilis)")
                if tentativa < max_tentativas:
                    print(f"    -> Vou atualizar a página e tentar novamente...")
                    continue
                else:
                    print(f"    ❌ Todas as {max_tentativas} tentativas falharam.")
                    return False

            print("    -> Selecionando 'Fechado'...")
            try:
                opcao_fechado = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                    (By.XPATH, "//div[contains(@class,'select2-result-label') and contains(.,'Fechado')]")
                ))
                opcao_fechado.click()
                print("    -> 'Fechado' selecionado.")
                time.sleep(1.5)
            except TimeoutException:
                print(f"    ⚠️ Opção 'Fechado' não apareceu! (Bug do Agilis)")
                if tentativa < max_tentativas:
                    print(f"    -> Vou atualizar a página e tentar novamente...")
                    continue
                else:
                    print(f"    ❌ Todas as {max_tentativas} tentativas falharam.")
                    return False

            print("    -> Clicando em 'Enviar'...")
            btn_enviar = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[data-cs-field='send_notification_btn']")
            ))
            btn_enviar.click()
            print("    -> Chamado fechado com sucesso! ✅")
            time.sleep(3)
            return True

        except Exception as e:
            print(f"    [ERRO] Falha na tentativa {tentativa}: {e}")
            if tentativa < max_tentativas:
                print(f"    -> Atualizando página e tentando novamente...")
            else:
                print(f"    ❌ Todas as {max_tentativas} tentativas falharam.")

    return False

def clicar_solicitacoes_vencem_hoje(driver, wait):
    try:
        link_vence_hoje = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "a[data-name='requests_due_today']")
        ))
        link_vence_hoje.click()
        print("   Clicou em 'Solicitações que vencem hoje'.")
        time.sleep(5)
        return True
    except TimeoutException:
        print("   Nenhum chamado vence hoje (ou link não encontrado).")
        return False

def voltar_para_lista(driver, wait):
    """Volta para a lista de 'Solicitações que vencem hoje' após processar um chamado."""
    URL_INICIAL = "https://agilis.mrv.com.br/HomePage.do?view_type=my_view"
    print("  -> Voltando para a lista de chamados...")
    driver.get(URL_INICIAL)
    time.sleep(5)
    if clicar_solicitacoes_vencem_hoje(driver, wait):
        print("  -> Lista de chamados recarregada com sucesso.")
        return True
    else:
        print("  -> Sem mais chamados que vencem hoje.")
        return False

# =============================================
# NOVAS FUNÇÕES — COLETA VIA API (substitui paginação DOM)
# =============================================

def get_session_from_selenium(driver):
    """Cria uma session requests autenticada com os cookies do Selenium."""
    session = requests.Session()
    for cookie in driver.get_cookies():
        session.cookies.set(cookie['name'], cookie['value'])
    session.headers.update({
        "User-Agent": driver.execute_script("return navigator.userAgent"),
        "Referer": driver.current_url,
        "X-Requested-With": "XMLHttpRequest"
    })
    return session

def buscar_chamados_vencem_hoje_api(driver, filter_id="3316"):
    """
    Busca TODOS os chamados que vencem hoje via API REST,
    percorrendo todas as páginas automaticamente.
    Retorna lista de dicts com id e due_by_time.
    """
    BASE_URL = "https://agilis.mrv.com.br/api/v3/requests"
    session = get_session_from_selenium(driver)

    todos = []
    start_index = 0
    row_count = 25
    has_more = True

    print(f"  🔍 Buscando chamados via API (filtro {filter_id})...")

    while has_more:
        input_data = {
            "list_info": {
                "start_index": start_index,
                "sort_field": "due_by_time",
                "sort_order": "desc",
                "row_count": str(row_count),
                "get_total_count": True,
                "filter_by": {"id": filter_id},
                "fields_required": [
                    "id", "subject", "status",
                    "due_by_time", "technician", "group"
                ]
            }
        }

        params = {
            "input_data": json.dumps(input_data),
            "SUBREQUEST": "XMLHTTP"
        }

        try:
            response = session.get(BASE_URL, params=params, timeout=30)
            data = response.json()
        except Exception as e:
            print(f"  [ERRO API] Falha na requisição: {e}")
            break

        chamados = data.get("requests", [])
        todos.extend(chamados)

        list_info = data.get("list_info", {})
        has_more = list_info.get("has_more_rows", False)
        total = list_info.get("total_count", len(todos))  # ✅ atualiza a cada página
        start_index += row_count

        print(f"  📄 {len(todos)}/{total} chamados coletados...")

    print(f"  ✅ Total coletado: {len(todos)} chamados.")
    return todos

def filtrar_proximos_vencer(chamados, minutos=60):
    """
    Filtra chamados que vencem em menos de X minutos a partir de agora.
    O campo due_by_time.value vem em milissegundos UTC da API.
    """
    agora_ms = datetime.now(timezone.utc).timestamp() * 1000
    limite_ms = agora_ms + (minutos * 60 * 1000)

    proximos = []
    for c in chamados:
        due = c.get("due_by_time", {})
        valor = due.get("value")
        if valor is None:
            continue
        due_ms = float(valor)
        # Inclui vencidos (due_ms < agora_ms) e próximos de vencer
        if due_ms <= limite_ms:
            minutos_restantes = (due_ms - agora_ms) / 60000
            proximos.append({
                "id": c["id"],
                "subject": c.get("subject", ""),
                "due_display": due.get("display_value", ""),
                "minutos_restantes": minutos_restantes,
                "status": c.get("status", {}).get("name", "")
            })

    # Ordena do mais urgente para o menos urgente
    proximos.sort(key=lambda x: x["minutos_restantes"])
    return proximos

# =============================================
# FUNÇÃO PRINCIPAL (Chamada pela Interface)
# =============================================
def executar_fechamento():
    print("[PROGRESSO: 5]")
    URL_INICIAL = "https://agilis.mrv.com.br/HomePage.do?view_type=my_view"
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 70)

    try:
        fazer_login(driver, wait)
        print("[PROGRESSO: 20]")

        # =============================================
        # LOOP PRINCIPAL
        # =============================================
        while True:
            if not verificar_sessao_ativa(driver):
                print("⚠️ Sessão perdida! Reiniciando o navegador...")
                try:
                    driver.quit()
                except:
                    pass
                driver = webdriver.Chrome()
                wait = WebDriverWait(driver, 70)
                fazer_login(driver, wait)

            print(f"\n{'='*60}")
            print(f"Verificação iniciada às {datetime.now().strftime('%H:%M:%S')}")
            print(f"{'='*60}")

            # =============================================
            # COLETA VIA API — substitui toda a paginação DOM
            # =============================================
            driver.get(URL_INICIAL)
            time.sleep(3)

            todos_chamados = buscar_chamados_vencem_hoje_api(driver, filter_id="3316")

            if not todos_chamados:
                print("[PROGRESSO: 100]")
                print("\n✅ Nenhum chamado para vencer hoje encontrado via API.")
                print("🛑 Encerrando o robô automaticamente para liberar a central.")
                break 

            proximos = filtrar_proximos_vencer(todos_chamados, minutos=60)

            if not proximos:
                print("[PROGRESSO: 100]")
                print(f"  ℹ️ {len(todos_chamados)} chamados encontrados, mas nenhum vence em menos de 1 hora.")
                print("🛑 Encerrando o robô automaticamente para liberar a central.")
                break

            print(f"\n  ⚠️ {len(proximos)} chamados para processar:")
            for c in proximos:
                sinal = "🔴" if c["minutos_restantes"] < 0 else "🟡"
                print(f"    {sinal} #{c['id']} | {c['due_display']} | {c['minutos_restantes']:.1f} min | {c['subject'][:50]}")

            # =============================================
            # LOOP DE PROCESSAMENTO DOS CHAMADOS
            # =============================================
            total_chamados = len(proximos)
            for i, chamado in enumerate(proximos):
                 # ✅ Verifica se ainda está aberto antes de abrir no navegador
                if not chamado_ainda_aberto(driver, chamado["id"]):
                    print(f"  ⏭️ Chamado #{chamado['id']} já foi fechado. Pulando...")
                    continue
                chamado_id = chamado["id"]
                print(f"\n{'='*60}")
                print(f"  🔓 Abrindo chamado #{chamado_id} ({chamado['due_display']})...")

                url_chamado = f"https://agilis.mrv.com.br/WorkOrder.do?woMode=viewWO&woID={chamado_id}"
                driver.get(url_chamado)
                time.sleep(4)

                sucesso = processar_chamado(driver, wait)

                if sucesso:
                    print(f"  ✅ Chamado #{chamado_id} fechado com sucesso!")
                else:
                    print(f"  ❌ Falha ao fechar chamado #{chamado_id}. Pulando...")

                # Volta para a home entre cada chamado
                driver.get(URL_INICIAL)
                time.sleep(3)
                
                # Progresso dinâmico de 20% a 95%
                progresso_atual = 20 + int(((i + 1) / total_chamados) * 75)
                print(f"[PROGRESSO: {progresso_atual}]")

            # =============================================
            # FIM DO CICLO
            # =============================================
            print("[PROGRESSO: 100]")
            print("\n" + "="*60)
            print("CICLO CONCLUÍDO!")
            print("="*60)
            break # Quebra o loop infinito para o robô poder finalizar e liberar a interface

    except Exception as e:
        print(f"\n[ERRO FATAL] {e}")
        traceback.print_exc()
        raise e 

    finally:
        print("\nFechando o navegador...")
        try:
            driver.quit()
        except:
            pass

if __name__ == "__main__":
    executar_fechamento()
