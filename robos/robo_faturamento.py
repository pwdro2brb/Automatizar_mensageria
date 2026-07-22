import os
import re
import time
import getpass
import unicodedata
import pandas as pd
from pathlib import Path
from PyPDF2 import PdfReader
from typing import Optional, Dict, List, Union, Tuple

# --- Selenium Imports ---
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    StaleElementReferenceException, TimeoutException, 
    ElementClickInterceptedException, ElementNotInteractableException, WebDriverException
)
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options 

# --- Pywinauto Imports ---
from pywinauto.application import Application
from pywinauto import Desktop
from pywinauto.timings import wait_until

import win32com.client as win32
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import numbers
import traceback
import sys

# ==============================================================================
# CONFIGURAÇÃO DE PASTAS DINÂMICAS
# ==============================================================================
# Garante que o robô ache o config.py na raiz do projeto
sys.path.append(str(Path(__file__).parent.parent))
import config
from config import EMAIL_MRV, SENHA_MRV

# Aponta dinamicamente para a nova pasta usando o Radar do config
PASTA_ARQUIVOS_RATEIO = Path(config.PASTA_ARQUIVOS) / "faturamento"

# Constantes Globais
CNPJ_CORREIOS_FIXO = "34028316001509"   
DATE_RE = r"([0-3]?\d/[01]?\d/\d{4})"   

# ==============================================================================
# 1. FUNÇÃO: GERAR RASCUNHOS
# ==============================================================================
def criar_rascunhos_correios():
    print("[PROGRESSO: 5]")
    print("Iniciando a criação de rascunhos no Outlook...")
    caminho_base = r"\\Bhz-fls-app1\mrvbh\Gerência Administrativa\Pública\NUCLEO DE CONTRATOS E APOIO A GESTÃO\CONTRATOS\Contratos Serviços\1. CORREIOS\2. Faturamento\2026"
    
    pastas_meses = [f for f in os.listdir(caminho_base) if os.path.isdir(os.path.join(caminho_base, f))]
    if not pastas_meses:
        print("Nenhuma pasta de mês encontrada no diretório.")
        return

    pastas_meses.sort()
    pasta_mes_recente = pastas_meses[-1]
    caminho_mes_recente = os.path.join(caminho_base, pasta_mes_recente)
    nome_mes = pasta_mes_recente.split("-")[-1].strip()

    print(f"Pasta mais recente encontrada: {pasta_mes_recente}")
    print("[PROGRESSO: 15]")

    contatos_para = {
        "Campinas": "flavia.pinho@mrv.com.br; ana.tilli@mrv.com.br",
        "Ribeirão Preto": "kaylana.alves@mrv.com.br",
        "Centro Oeste": "nicole.souza@mrv.com.br; maksuel.araujo@mrv.com.br; eunice.prudente@primeconstrucoes.com.br; maryanne.camargo@primeconstrucoes.com.br",
        "Nordeste": "langela.santos@mrv.com.br",
        "Sul": "victoria.gomes@mrv.com.br; filipe.avila@mrv.com.br; simone.csantos@mrv.com.br; monique.silva@mrv.com.br",
        "São Paulo": "telma.amattos@mrv.com.br; cristina.demetrio@parceiro.mrv.com.br; manoella.camargo@mrv.com.br; luciano.lsilva@mrv.com.br; nicoli.santos@mrv.com.br",
        "Triângulo": "kamilly.silva@mrv.com.br; kaylana.alves@mrv.com.br; maria.fernnanda@mrv.com.br"
    }

    agora = datetime.now()
    saudacao = "Bom dia" if agora.hour < 12 else "Boa tarde"
    prazo_rateio = agora + timedelta(hours=32)
    prazo_formatado = prazo_rateio.strftime("%d/%m/%Y às %H:%M")

    corpo_email = f"""
    <p style="font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000;">
        {saudacao}, Prezado(s)!<br><br>
        Segue em anexo o extrato dos Correios. O rateio deverá ser enviado até <b>{prazo_formatado}</b>.<br><br>
        Atenciosamente,
    </p>
    """

    outlook = win32.Dispatch('outlook.application')
    pastas_regionais = os.listdir(caminho_mes_recente)
    
    # Filtra apenas as pastas válidas para calcular o total
    regionais_validas = [r for r in pastas_regionais if os.path.isdir(os.path.join(caminho_mes_recente, r)) and r.upper() != "BH"]
    total_regionais = len(regionais_validas)
    
    print("[PROGRESSO: 25]")
    
    for i, regional in enumerate(regionais_validas):
        caminho_regional = os.path.join(caminho_mes_recente, regional)

        print(f"Gerando rascunho para: {regional}...")
        mail = outlook.CreateItem(0)
        mail.To = contatos_para.get(regional, "")
        
        cc_padrao = "vanessa.brodrigues@mrv.com.br; correiosbh@mrv.com.br"
        if regional in ["Triângulo", "Ribeirão Preto"]:
            mail.CC = f"{cc_padrao}; conceicao@mrv.com.br"
        else:
            mail.CC = cc_padrao

        mail.Subject = f"RES: Extrato Correios - {regional} ({nome_mes})"

        arquivos_na_pasta = os.listdir(caminho_regional)
        for arquivo in arquivos_na_pasta:
            caminho_arquivo = os.path.join(caminho_regional, arquivo)
            if os.path.isfile(caminho_arquivo):
                mail.Attachments.Add(caminho_arquivo)

        mail.Display() 
        assinatura_outlook = mail.HTMLBody
        mail.HTMLBody = f"<html><body>{corpo_email}{assinatura_outlook}</body></html>"
        mail.Save()
        mail.Close(0)
        
        # Calcula o progresso dinâmico (de 25% até 95%)
        progresso_atual = 25 + int(((i + 1) / total_regionais) * 70)
        print(f"[PROGRESSO: {progresso_atual}]")

    print("[PROGRESSO: 100]")
    print("\nProcesso concluído! Verifique a pasta 'Rascunhos' no seu Outlook.")

# ==============================================================================
# 2. FUNÇÃO: GERAR PLANILHA DE RATEIO
# ==============================================================================

def preparar_e_gerar_rateio():
    print("[PROGRESSO: 5]")
    print("Lendo planilhas na pasta testar_edicao...")
    pasta_trabalho = PASTA_ARQUIVOS_RATEIO / "testar_edicao"
    
    if not pasta_trabalho.exists():
        raise FileNotFoundError(f"A pasta 'testar_edicao' não foi encontrada dentro de:\n{PASTA_ARQUIVOS_RATEIO}")
    
    caminho_rr = None
    caminho_correios = None
    
    for ficheiro in pasta_trabalho.glob('*.xlsx'):
        nome_ficheiro = ficheiro.name.upper()
        if nome_ficheiro.startswith('~$') or nome_ficheiro == 'RATEIO PAG.XLSX': continue
        if 'RATEIO RECEBIDO' in nome_ficheiro: caminho_rr = ficheiro
        elif re.match(r'^\d{7}\.XLSX$', nome_ficheiro): caminho_correios = ficheiro

    if not caminho_rr and not caminho_correios:
        raise FileNotFoundError("Faltam as DUAS planilhas (Rateio Recebido e a dos Correios) na pasta 'testar_edicao'.")
    elif not caminho_rr:
        raise FileNotFoundError("Falta a planilha 'RATEIO RECEBIDO' na pasta 'testar_edicao'.")
    elif not caminho_correios:
        raise FileNotFoundError("Falta a planilha numérica dos Correios (ex: 1234567.xlsx) na pasta 'testar_edicao'.")

    print("[PROGRESSO: 20]")
    print(f">> Iniciando processamento...")
    
    pasta_exemplos = PASTA_ARQUIVOS_RATEIO / "exemplos"
    
    if not pasta_exemplos.exists():
        pasta_exemplos.mkdir(parents=True, exist_ok=True)
        print(f"📁 Pasta 'exemplos' criada automaticamente.")
        
    caminho_saida = pasta_exemplos / "RATEIO PAG.xlsx"
    
    print("[PROGRESSO: 30]")
    final = gerar_rateio_pag(caminho_correios=caminho_correios, caminho_rr=caminho_rr, saida=caminho_saida)
    
    print("[PROGRESSO: 100]")
    print(f"Total de linhas geradas no final: {len(final)}")
    print(f"✅ Arquivo RATEIO PAG.xlsx gerado e salvo direto na pasta 'exemplos'!")



# ==============================================================================
# 3. FUNÇÃO: LANÇAR NOTA FISCAL
# ==============================================================================
def lancar_nota_fiscal():
    print("[PROGRESSO: 2]")
    PASTA_BASE = PASTA_ARQUIVOS_RATEIO / "exemplos"
    ARQUIVO_REGRAS_XLSX = PASTA_ARQUIVOS_RATEIO / "dados_puxados_preenchimento.xlsx"

    if not ARQUIVO_REGRAS_XLSX.exists():
        raise FileNotFoundError(f"A planilha de regras 'dados_puxados_preenchimento.xlsx' não foi encontrada no caminho:\n{ARQUIVO_REGRAS_XLSX}")
    
    if not PASTA_BASE.exists():
        raise FileNotFoundError(f"A pasta 'exemplos' não foi encontrada dentro de:\n{PASTA_ARQUIVOS_RATEIO}")

    planilhas_encontradas = list(PASTA_BASE.glob("RATEIO PAG.xlsx"))
    if not planilhas_encontradas:
        raise FileNotFoundError("A planilha 'RATEIO PAG.xlsx' não foi encontrada na pasta 'exemplos'.\nVocê esqueceu de rodar a Etapa 2?")
    caminho_planilha_rateio = str(planilhas_encontradas[0].resolve())

    pdfs_encontrados = [arq for arq in PASTA_BASE.glob("*") if arq.suffix.lower() == ".pdf"]
    if not pdfs_encontrados:
        raise FileNotFoundError("Nenhum boleto em PDF foi encontrado na pasta 'exemplos'.")
    elif len(pdfs_encontrados) > 1:
        raise FileNotFoundError(f"Foram encontrados {len(pdfs_encontrados)} PDFs na pasta 'exemplos'.\nDeixe apenas UM boleto na pasta para o robô não se confundir!")
    caminho_boleto_pdf = str(pdfs_encontrados[0].resolve())

    print("Iniciando robô de lançamento...")
    print(f"✅ Planilha de Upload carregada: {caminho_planilha_rateio}")
    print(f"✅ Boleto carregado: {caminho_boleto_pdf}")

    print("[PROGRESSO: 10]")
    print("⏳ Extraindo dados do Boleto...")
    campos = extrair_campos_boleto(caminho_boleto_pdf)

    cnpj_correios = campos["cnpj_beneficiario"]
    num_doc       = campos["numero_documento"]
    vencimento    = campos["vencimento"]
    valor_boleto  = campos["valor_total_str"]
    cnpj_mrv      = campos["cnpj_pagador"]
    emissao_proc  = campos["data_processamento"]

    if not cnpj_mrv or not valor_boleto:
        raise ValueError("Não foi possível localizar o CNPJ da MRV ou o Valor no boleto PDF.")

    print(f"📌 Dados extraídos: CNPJ MRV: {cnpj_mrv} | Valor: R$ {valor_boleto} | Nº Doc: {num_doc}")

    df = pd.read_excel(ARQUIVO_REGRAS_XLSX, engine="openpyxl")

    texto_completo_pdf = norm_text(read_pdf_text(caminho_boleto_pdf)).upper()
    norm_limpo = texto_completo_pdf.replace(".", "").replace("/", "").replace("-", "")

    ID_REGIONAL = None
    candidatos = []

    for index, linha in df.iterrows():
        palavra_chave = str(linha.get("PALAVRA_CHAVE", "")).upper()
        if not palavra_chave or palavra_chave == "NAN": continue
        palavra_chave_limpa = palavra_chave.replace(".", "").replace("/", "").replace("-", "")
        if palavra_chave_limpa in norm_limpo:
            candidatos.append(linha)

    if len(candidatos) == 0:
        print("⚠️ Falha: Nenhuma 'PALAVRA_CHAVE' da planilha Excel foi encontrada no boleto.")
        return
    elif len(candidatos) == 1:
        linha_escolhida = candidatos[0]
        ID_REGIONAL  = linha_escolhida["ID"]
        descr        = linha_escolhida["DESCR"]
        material_cod = str(linha_escolhida["material_cod"])
    else:
        resultado = determinar_id_por_valor(valor_boleto, cnpj_mrv, df)
        if resultado is None:
            print("⚠️ Não foi possível determinar o ID regional pelo valor do boleto.")
            return
        ID_REGIONAL  = resultado["ID"]
        descr        = resultado["DESCR"]
        material_cod = resultado["material_cod"]

    print(f"📋 Regional: {ID_REGIONAL} | Descrição: {descr} | Material: {material_cod}")

    print("[PROGRESSO: 15]")
    try:
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True) 
        
        driver = webdriver.Chrome(options=chrome_options) 
        driver.get("https://mrvpag2.mrv.com.br/home")
        driver.maximize_window()
        wait_longo = WebDriverWait(driver, 180)
        wait = WebDriverWait(driver, 15)
        wait_rapido = WebDriverWait(driver, 2)
    except Exception as e:
        print(f"Erro ao iniciar o Chrome: {e}")
        return

    try:    
        print("Aguardando login...")
        wait.until(EC.presence_of_element_located((By.ID, "i0116"))).send_keys(EMAIL_MRV)
        click_anti_stale(wait, By.ID, "idSIButton9")
        wait.until(EC.presence_of_element_located((By.ID, "i0118"))).send_keys(SENHA_MRV)
        click_anti_stale(wait_longo, By.ID, "idSIButton9")
        print("!!! APROVE O MFA NO CELULAR !!!")
        click_anti_stale(wait_longo, By.ID, "idSIButton9") 
        
        print("[PROGRESSO: 25]")
        fechar_mensagem = WebDriverWait(driver, 100).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "mat-icon.btnCancelTest")))
        fechar_mensagem.click()
        print("Mensagem de boas-vindas fechada.")

        novo_protocolo = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.item-menu[routerlink='/protocolo'], a.item-menu[href='/protocolo']")))
        novo_protocolo.click()
        
        print("[PROGRESSO: 35]")
        tipo_nota = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "mat-card a.pointer")))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", tipo_nota)
        tipo_nota.click()
        
        tipo_de_documento = wait.until(EC.presence_of_element_located((By.XPATH, "//mat-expansion-panel-header[.//mat-panel-title[contains(normalize-space(),'Tipo de Documento')]]")))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", tipo_de_documento)
        if tipo_de_documento.get_attribute("aria-expanded") == "false":
            wait.until(EC.element_to_be_clickable(tipo_de_documento)).click()
        
        select_el = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "mat-select[formcontrolname='frmTipoDoc'], mat-select[aria-label='Qual o tipo do documento'], #mat-select-2")))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", select_el)
        select_el.click()

        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, ".cdk-overlay-pane .mat-select-panel")))
        opcao = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-option[.//span[contains(@class,'mat-option-text') and normalize-space()='NF somente de Serviços']]")))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", opcao)
        opcao.click()

        print("[PROGRESSO: 45]")
        input_enviar_arquivo = WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#frmFile")))
        input_enviar_arquivo.send_keys(caminho_boleto_pdf)

        btn_continuar = wait.until(EC.element_to_be_clickable((By.ID, "btnTipoDoc")))
        safe_click(driver, btn_continuar) 

        aprovador_select = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-select[@aria-label='APROVADOR' or @placeholder='APROVADOR']")))
        safe_click(driver, aprovador_select) 

        opcao_vanessa = wait.until(EC.element_to_be_clickable((By.XPATH, "//mat-option//span[normalize-space(.)='VANESSA DE BRITO RODRIGUES (VANESSA.BRODRIGUES)']/ancestor::mat-option")))
        safe_click(driver, opcao_vanessa) 

        print("[PROGRESSO: 55]")
        cnpj_input = driver.find_element(By.CSS_SELECTOR, "input[placeholder='CNPJ DA EMPRESA MRV'], input[aria-label='CNPJ DA EMPRESA MRV']")
        cnpj_input.clear()
        cnpj_input.send_keys(cnpj_mrv)
        cnpj_input.send_keys(Keys.TAB)

        xpath_linha_mrv = "(//tr[contains(@class,'mat-row')][.//td[contains(normalize-space(.),'MRV')]])[1]"
        linha_mrv = wait.until(EC.presence_of_element_located((By.XPATH, xpath_linha_mrv)))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", linha_mrv)

        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, xpath_linha_mrv)))
        except TimeoutException:
            pass 

        try:
            try:
                ActionChains(driver).move_to_element(linha_mrv).pause(0.1).click(linha_mrv).perform()
            except (ElementClickInterceptedException, StaleElementReferenceException):
                primeiro_td = linha_mrv.find_element(By.XPATH, ".//td[1]")
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", primeiro_td)
                try:
                    ActionChains(driver).move_to_element(primeiro_td).pause(0.1).click(primeiro_td).perform()
                except Exception:
                    driver.execute_script("arguments[0].click();", primeiro_td)
        except Exception as e:
            print("Erro ao clicar na linha MRV.")
            raise

        print("[PROGRESSO: 65]")
        wait_overlay_gone(driver, wait)
        wait_no_overlay(driver, wait)
        
        for _ in range(3):
            try:
                inp_num_doc = get_input_by_formcontrol(driver, wait, "frmNumDocumento")
                type_safely(driver, wait, inp_num_doc, num_doc)
                if (inp_num_doc.get_attribute("value") or "").strip() == num_doc: break
            except StaleElementReferenceException: continue

        wait_no_overlay(driver, wait)
        for _ in range(3):
            try:
                inp_cnpj_cor = get_input_by_formcontrol(driver, wait, "frmCnpjFornecedor")
                editable = ensure_enabled_and_editable(driver, inp_cnpj_cor, allow_force=True)
                if editable: type_safely(driver, wait, inp_cnpj_cor, cnpj_correios)
                else: js_set_value_and_dispatch(driver, inp_cnpj_cor, cnpj_correios)
                if (re.sub(r"\D", "", inp_cnpj_cor.get_attribute("value") or "")) == cnpj_correios: break
            except StaleElementReferenceException: continue

        wait_no_overlay(driver, wait)
        for _ in range(3):
            try:
                inp_emissao = get_input_by_formcontrol(driver, wait, "frmDtEmissao")
                try: click_with_fallback(driver, inp_emissao)
                except Exception: pass
                inp_emissao.send_keys(Keys.ESCAPE)
                type_safely(driver, wait, inp_emissao, emissao_proc)
                if (inp_emissao.get_attribute("value") or "").strip() == emissao_proc: break
            except StaleElementReferenceException: continue

        wait_no_overlay(driver, wait)
        for _ in range(3):
            try:
                inp_venc = get_input_by_formcontrol(driver, wait, "frmVencimento")
                ensure_enabled_and_editable(driver, inp_venc, allow_force=True)
                try: click_with_fallback(driver, inp_venc)
                except Exception: pass
                inp_venc.send_keys(Keys.ESCAPE)
                type_safely(driver, wait, inp_venc, vencimento)
                if (inp_venc.get_attribute("value") or "").strip() == vencimento: break
            except StaleElementReferenceException: continue

        wait_no_overlay(driver, wait)
        for _ in range(3):
            try:
                inp_valor = get_input_by_formcontrol(driver, wait, "frmValorTotalNf")
                click_with_fallback(driver, inp_valor)
                inp_valor.send_keys(Keys.CONTROL, 'a', Keys.DELETE)
                for ch in valor_boleto:
                    inp_valor.send_keys(ch)
                    time.sleep(0.01)
                if (inp_valor.get_attribute("value") or "").strip(): break
            except StaleElementReferenceException: continue

        qtd_cliques = click_ok_confirm(driver, wait_rapido, timeout=1, max_tentativas=1)
        if qtd_cliques > 0:
            print(f"✅ {qtd_cliques} diálogo(s) confirmado(s) com sucesso.")

        print("[PROGRESSO: 75]")
        campo_desc = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="frmDescNota"]')))
        driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', {bubbles: true}));
            arguments[0].dispatchEvent(new Event('change', {bubbles: true}));
        """, campo_desc, descr)
        campo_desc.send_keys(Keys.TAB) 
        time.sleep(0.2) 

        driver.execute_script("""
            let btns = Array.from(document.querySelectorAll("button"));
            let btn = btns.find(b => b.textContent.includes("CONTINUAR") && !b.disabled && !b.classList.contains("mat-button-disabled"));
            if(btn) btn.click();
        """)

        locator_link = (By.XPATH, "//a[.//span[normalize-space(.)='CONTINUAR']]")
        wait.until(EC.presence_of_element_located(locator_link)) 
        
        driver.execute_script("""
            let links = Array.from(document.querySelectorAll("a"));
            let link = links.find(a => a.textContent.includes("CONTINUAR") && !a.classList.contains("mat-button-disabled"));
            if(link) link.click();
        """)

        btn_adicionar = wait_rapido.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(.)='Adicionar']/ancestor::button[1]")))
        btn_adicionar.click()

        preencher_codigo_material_ultima_linha(driver, wait_rapido, material_cod, timeout=10)
        click_pesquisar(driver, wait_rapido)

        print("[PROGRESSO: 85]")
        print("Aguardando tabela carregar para selecionar o checkbox...")
        wait_longo = WebDriverWait(driver, 40)
        locator_checkbox = (By.XPATH, "(//td[contains(@class,'mat-column-select')]//mat-checkbox)[1]")
        wait_longo.until(EC.presence_of_element_located(locator_checkbox))
        
        driver.execute_script("""
            let matCheckbox = document.querySelector("td.mat-column-select mat-checkbox");
            if (matCheckbox) {
                let label = matCheckbox.querySelector("label");
                if (label) {
                    label.click();
                } else {
                    matCheckbox.click();
                }
            }
        """)
        print("✅ Primeira linha selecionada instantaneamente!")
        
        click_incluir_produtos(driver, wait_rapido)
        preencher_quantidade_e_valor(driver, wait_rapido, quantidade="1", valor_boleto=valor_boleto)
        abrir_select_justificativa(driver, wait_rapido)
        selecionar_opcao_justificativa_com_hover(driver, wait_rapido, texto_alvo="2 - Orientações do gestor/coordendor da área")
        click_continuar_proximo_ao_select(driver, wait_rapido)
        
        print("[PROGRESSO: 95]")
        print("Anexando planilha em segundo plano (sem abrir janela do Windows)...")
        try:
            inputs_file = driver.find_elements(By.XPATH, "//input[@type='file']")
            if not inputs_file:
                print("⚠️ Nenhum input de arquivo encontrado na página!")
            else:
                input_planilha = inputs_file[-1]
                driver.execute_script(
                    "arguments[0].style.display = 'block'; "
                    "arguments[0].style.visibility = 'visible'; "
                    "arguments[0].style.opacity = 1;", 
                    input_planilha
                )
                input_planilha.send_keys(caminho_planilha_rateio)
                print("✅ Planilha anexada com sucesso via HTML!")
                time.sleep(2) 
        except Exception as e:
            print(f"⚠️ Erro ao tentar anexar silenciosamente: {e}")
            
        total_ok = click_ok_confirm_repeatedly(driver, wait, max_clicks=3)
        print(f"[INFO] Botão OK clicado {total_ok} vez(es).")

        print("[PROGRESSO: 100]")
        print("✅ Fluxo concluído com sucesso! O navegador permanecerá aberto para conferência.")

    except Exception as e:
        print(f"❌ Erro Crítico durante a execução: {e}")
        traceback.print_exc()
        try:
            driver.save_screenshot("erro_final.png")
            debug_dump(driver, "erro_final")
        except Exception:
            pass

# ==============================================================================
# 4. FUNÇÕES AUXILIARES (Devem ficar na raiz do arquivo)
# ==============================================================================

def click_anti_stale(wait, by, seletor, tentativas=3):
    for _ in range(tentativas):
        try:
            elemento = wait.until(EC.element_to_be_clickable((by, seletor)))
            elemento.click()
            return True 
        except StaleElementReferenceException:
            time.sleep(0.5) 
    raise RuntimeError(f"O elemento {seletor} sumiu repetidas vezes.")

def scroll_center(driver, el):
    try: driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    except: pass

def wait_overlays_to_hide(wait):
    try: wait.until_not(lambda d: len(d.find_elements(By.CSS_SELECTOR, ".cdk-overlay-backdrop, .mat-progress-spinner")) > 0)
    except: pass

def safe_click(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    try:
        element.click()
    except Exception:
        driver.execute_script("arguments[0].click();", element)

def safe_click_diferenciado(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    try:
        element.click()
    except Exception:
        driver.execute_script("arguments[0].click();", element)

def fazer_upload_janela_windows(caminho_do_arquivo):
    from pywinauto import Desktop
    from pywinauto.timings import wait_until
    import time

    caminho_absoluto = str(Path(caminho_do_arquivo).resolve())
    titulos_possiveis = ["Abrir", "Open", "Abrir arquivo", "Escolher arquivo para carregar", "Escolher arquivo", "Select file", "File Upload", "Carregar"]
    classes_possiveis = ["#32770"]  

    janela = None
    backend_usado = None

    for backend in ["uia", "win32"]:
        if janela: break
        try:
            desktop = Desktop(backend=backend)
            for titulo in titulos_possiveis:
                try:
                    win = desktop.window(title=titulo)
                    if win.exists(timeout=1):
                        janela = win
                        backend_usado = backend
                        break
                except Exception: continue
        except Exception: continue

    if not janela:
        import re
        for backend in ["uia", "win32"]:
            if janela: break
            try:
                desktop = Desktop(backend=backend)
                for win in desktop.windows():
                    try:
                        titulo = win.window_text()
                        if any(t.lower() in titulo.lower() for t in ["abrir", "open", "upload", "file"]):
                            janela = win
                            backend_usado = backend
                            break
                    except Exception: continue
            except Exception: continue

    if not janela:
        for backend in ["uia", "win32"]:
            if janela: break
            try:
                desktop = Desktop(backend=backend)
                for classe in classes_possiveis:
                    try:
                        win = desktop.window(class_name=classe)
                        if win.exists(timeout=1):
                            janela = win
                            backend_usado = backend
                            break
                    except Exception: continue
            except Exception: continue

    if not janela:
        raise RuntimeError("⚠️ Janela de upload não encontrada!")

    try: janela.wait("ready", timeout=10)
    except Exception: time.sleep(1)  

    campo_preenchido = False
    if backend_usado == "uia":
        tentativas_campo = [
            lambda: janela.child_window(title="Nome do arquivo:", control_type="ComboBox").child_window(control_type="Edit"),
            lambda: janela.child_window(title="Nome do arquivo:", control_type="Edit"),
            lambda: janela.child_window(title="File name:", control_type="ComboBox").child_window(control_type="Edit"),
            lambda: janela.child_window(title="File name:", control_type="Edit"),
            lambda: janela.child_window(control_type="Edit", found_index=0),
        ]
    else:
        tentativas_campo = [
            lambda: janela.child_window(class_name="Edit", found_index=0),
            lambda: janela.child_window(title="Nome do arquivo:", class_name="ComboBoxEx32").child_window(class_name="Edit"),
            lambda: janela.child_window(class_name="ComboBoxEx32").child_window(class_name="Edit"),
        ]

    for get_campo in tentativas_campo:
        try:
            campo = get_campo()
            if campo.exists(timeout=2):
                campo.set_edit_text(caminho_absoluto)
                campo_preenchido = True
                break
        except Exception: continue

    if not campo_preenchido:
        try:
            from pywinauto.keyboard import send_keys
            janela.set_focus()
            time.sleep(0.3)
            send_keys("^a{DELETE}", pause=0.05)
            time.sleep(0.1)
            caminho_escaped = caminho_absoluto.replace("{", "{{").replace("}", "}}")
            send_keys(caminho_escaped, pause=0.02, with_spaces=True)
            campo_preenchido = True
        except Exception as e:
            raise RuntimeError(f"⚠️ Não consegui digitar o caminho do arquivo: {e}")

    time.sleep(0.5)
    botao_clicado = False
    nomes_botao = ["Abrir", "Open", "&Abrir", "&Open"]

    for nome_btn in nomes_botao:
        if botao_clicado: break
        try:
            if backend_usado == "uia": btn = janela.child_window(title=nome_btn, control_type="Button")
            else: btn = janela.child_window(title=nome_btn, class_name="Button")
            if btn.exists(timeout=2):
                btn.click()
                botao_clicado = True
        except Exception: continue

    if not botao_clicado:
        try:
            from pywinauto.keyboard import send_keys
            send_keys("{ENTER}")
            botao_clicado = True
        except Exception as e:
            raise RuntimeError(f"⚠️ Não consegui clicar no botão Abrir: {e}")

    for _ in range(20):
        try:
            if not janela.exists(timeout=0): return
        except Exception: return
        time.sleep(0.5)

def click_ok_confirm_repeatedly(driver, wait, max_clicks=5):    
    clicks = 0
    locators = [
        (By.ID, "btnTipoDoc"),
        (By.XPATH, "//button[contains(@class,'confirm') and normalize-space(.)='OK']")
    ]
    for _ in range(max_clicks):
        wait_overlays_to_hide(wait)
        btn = None
        for by, sel in locators:
            try:
                elem = driver.find_element(by, sel)
                if elem.is_displayed(): btn = elem; break
            except: pass
        if not btn: break
        
        safe_click(driver, btn)
        clicks += 1
        time.sleep(1)
    return clicks

def click_anexar_planilha(driver, wait, caminho_planilha=None):
    wait_overlays_to_hide(wait)
    locator_btn = (By.XPATH, "//button[.//span[normalize-space(.)='ANEXAR PLANILHA'] and not(@disabled) and not(contains(@class,'mat-button-disabled'))]")
    try: btn = wait.until(EC.presence_of_element_located(locator_btn))
    except TimeoutException: raise RuntimeError("Botão 'ANEXAR PLANILHA' não encontrado.")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    time.sleep(0.05)

    try: wait.until(EC.element_to_be_clickable(locator_btn))
    except TimeoutException: pass

    try:
        if not is_center_clickable_js(driver, btn):
            driver.execute_script("window.scrollBy(0, -80);")
            time.sleep(0.05)
        btn.click()
    except Exception:
        try: ActionChains(driver).move_to_element(btn).pause(0.05).click().perform()
        except Exception:
            try: driver.execute_script("arguments[0].click();", btn)
            except Exception as e3: raise RuntimeError(f"Falha ao clicar em 'ANEXAR PLANILHA': {repr(e3)}")

    if caminho_planilha:
        try: input_file = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='file']")))
        except TimeoutException: return
        caminho_planilha = str(Path(caminho_planilha).resolve())
        try: input_file.send_keys(caminho_planilha)
        except Exception:
            try:
                driver.execute_script("arguments[0].style.display='block'; arguments[0].style.visibility='visible';", input_file)
                time.sleep(0.05)
                input_file.send_keys(caminho_planilha)
            except Exception as e2: raise RuntimeError(f"Falha ao enviar arquivo para input[type=file]: {repr(e2)}")

def wait_overlay_gone(driver, wait, timeout=40):
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((
                By.CSS_SELECTOR,
                ".cdk-overlay-backdrop.cdk-overlay-backdrop-showing, .mat-progress-bar, .mat-spinner, .ngx-spinner-overlay"
            ))
        )
    except TimeoutException: pass

def get_visible_input(driver, formcontrol: str):
    candidates = driver.find_elements(By.CSS_SELECTOR, f"input[formcontrolname='{formcontrol}']")
    visibles = [el for el in candidates if el.is_displayed() and el.is_enabled()]
    if not visibles: raise TimeoutException(f"Input visível '{formcontrol}' não encontrado.")
    return visibles[0]

def focus_input(driver, wait, el):
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        wait.until(lambda d: el.is_displayed() and el.is_enabled())
        try: el.click()
        except ElementClickInterceptedException:
            container = el.find_element(By.XPATH, "./ancestor::mat-form-field//div[contains(@class,'mat-form-field-flex')]")
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", container)
            try: ActionChains(driver).move_to_element(container).pause(0.05).click(container).perform()
            except Exception: driver.execute_script("arguments[0].click();", container)
    except StaleElementReferenceException: pass

def clear_and_type(el, text: str):
    driver = el.parent  
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.15)
    except Exception: pass

    try: el.click()
    except Exception:
        try: driver.execute_script("arguments[0].click(); arguments[0].focus();", el)
        except Exception: pass

    try:
        el.send_keys(Keys.CONTROL, 'a', Keys.DELETE)
        time.sleep(0.1)
        if text: el.send_keys(str(text))
    except Exception:
        driver.execute_script("""
            const el = arguments[0], val = arguments[1] ?? '';
            el.value = '';
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.value = val;
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
        """, el, text or '')

def fill_input(driver, wait, formcontrol: str, value: str, numeric=False):
    wait_overlay_gone(driver, wait)
    el = get_visible_input(driver, formcontrol)
    focus_input(driver, wait, el)
    if numeric and value is not None: value = re.sub(r'\D', '', str(value))
    clear_and_type(el, value)
    try: WebDriverWait(driver, 5).until(lambda d: (el.get_attribute("value") or "") != "")
    except Exception: pass

def somente_digitos(s: str) -> str:
    return re.sub(r"\D", "", s or "")

def norm_text(s: str) -> str:
    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("ASCII")
    s = s.replace("\xa0", " ")
    return re.sub(r"[ \t]+", " ", s)

def read_pdf_text(pdf_path: str) -> str:
    reader = PdfReader(pdf_path)
    txt = [p.extract_text() or "" for p in reader.pages]
    texto = "\n".join(txt)
    if not texto.strip(): raise ValueError("PDF sem texto. Use OCR.")
    return texto

def extrair_cnpj_pagador(text_norm: str) -> Optional[str]:
    padrao_cnpj = r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b"
    cnpjs_encontrados = re.findall(padrao_cnpj, text_norm)
    cnpjs_correios = ["34.028.316/0015-09", "34.028.316/0001-03"]
    for cnpj in cnpjs_encontrados:
        if cnpj not in cnpjs_correios: return cnpj
    return None

def extrair_numero_documento_7d(text_norm: str) -> Optional[str]:
    mdoc = re.search(r"DOCUMENTO.{0,100}(\d{7})", text_norm, flags=re.I | re.S)
    return mdoc.group(1) if mdoc else None

def extrair_valor_total(text_norm: str) -> Optional[str]:
    anchor = r"(?i)VALOR\s*(?:DO\s*)?DOCUMENTO(?:\s*\(R\$\))?"
    money_re = r"(?:R\$\s*)?(\d{1,3}(?:\.\d{3})*,\d{2})"
    lines = text_norm.splitlines()
    for i, line in enumerate(lines):
        if re.search(anchor, line, flags=re.I):
            for j in (i, i+1, i+2): 
                if 0 <= j < len(lines):
                    m = re.search(money_re, lines[j], flags=re.I)
                    if m: return m.group(1)
    return None

def extrair_datas_correios(text_norm: str) -> dict:
    from datetime import datetime
    todas_datas_str = re.findall(r"\d{2}/\d{2}/\d{4}", text_norm)
    vistas = set()
    unicas_str = []
    for d in todas_datas_str:
        if d not in vistas:
            vistas.add(d)
            unicas_str.append(d)

    datas_parsed = []
    for d_str in unicas_str:
        try:
            dt = datetime.strptime(d_str, "%d/%m/%Y")
            datas_parsed.append((d_str, dt))
        except ValueError: continue  

    datas_parsed.sort(key=lambda x: x[1])
    if not datas_parsed: return {"emissao": "", "vencimento": ""}

    vencimento = datas_parsed[-1][0]  
    if len(datas_parsed) >= 3: emissao = datas_parsed[1][0]
    elif len(datas_parsed) == 2: emissao = datas_parsed[0][0]
    else: emissao = datas_parsed[0][0]

    dt_emissao = datetime.strptime(emissao, "%d/%m/%Y")
    dt_vencimento = datetime.strptime(vencimento, "%d/%m/%Y")
    if dt_vencimento < dt_emissao: emissao, vencimento = vencimento, emissao

    return {"emissao": emissao, "vencimento": vencimento}

def extrair_campos_boleto(pdf_path: str) -> Dict[str, Optional[str]]:
    texto = read_pdf_text(pdf_path)
    norm  = norm_text(texto)
    num_doc = extrair_numero_documento_7d(norm)
    datas = extrair_datas_correios(norm) 
    valor_total_str = extrair_valor_total(norm)
    cnpj_pagador_extraido = extrair_cnpj_pagador(norm)
    cnpj_pag = somente_digitos(cnpj_pagador_extraido) if cnpj_pagador_extraido else None
    
    return {
        "numero_documento": num_doc,
        "cnpj_pagador": cnpj_pag,
        "cnpj_beneficiario": CNPJ_CORREIOS_FIXO,
        "data_processamento": datas["emissao"],
        "vencimento": datas["vencimento"], 
        "valor_total_str": valor_total_str     
    }

def determinar_id_por_valor(valor_str: str, cnpj_pagador: str, df: pd.DataFrame) -> dict:
    valor_float = float(valor_str.replace(".", "").replace(",", "."))
    cnpj_limpo = cnpj_pagador.replace(".", "").replace("/", "").replace("-", "")
    df["PALAVRA_CHAVE_LIMPA"] = df["PALAVRA_CHAVE"].astype(str).str.replace(".", "", regex=False).str.replace("/", "", regex=False).str.replace("-", "", regex=False).str.upper()
    linhas_cnpj = df[df["PALAVRA_CHAVE_LIMPA"] == cnpj_limpo]

    if linhas_cnpj.empty: return None
    if len(linhas_cnpj) == 1:
        linha = linhas_cnpj.iloc[0]
        return {"ID": linha["ID"], "DESCR": linha["DESCR"], "material_cod": str(linha["material_cod"])}

    if valor_float < 400.00: id_alvo = 2
    elif valor_float < 2000.00: id_alvo = 7
    elif valor_float >= 40000.00: id_alvo = 8
    else: raise ValueError(f"⚠️ Valor R$ {valor_str} não se encaixa nas faixas.")

    linha_alvo = linhas_cnpj[linhas_cnpj["ID"] == id_alvo]
    if linha_alvo.empty: raise ValueError(f"⚠️ ID {id_alvo} não encontrado.")

    linha = linha_alvo.iloc[0]
    return {"ID": linha["ID"], "DESCR": linha["DESCR"], "material_cod": str(linha["material_cod"])}

def click_continuar_proximo_ao_select(driver, wait):
    container = driver.find_element(By.XPATH, "//mat-form-field[.//mat-select[@formcontrolname='justificativa']]")
    locator_local = (By.XPATH, ".//following::a[.//span[normalize-space(.)='CONTINUAR'] and (not(@aria-disabled) or @aria-disabled='false') and not(contains(@class,'mat-button-disabled'))][1]")
    try: link = container.find_element(*locator_local)
    except Exception: raise RuntimeError("Não achei o CONTINUAR referente a este passo.")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", link)
    time.sleep(0.05)
    try: wait.until(EC.element_to_be_clickable((By.XPATH, "//a[.//span[normalize-space(.)='CONTINUAR']]")))
    except TimeoutException: pass

    try: link.click()
    except Exception:
        try: ActionChains(driver).move_to_element(link).pause(0.05).click().perform()
        except Exception: driver.execute_script("arguments[0].click();", link)

def preencher_codigo_material_ultima_linha(driver, wait, material_cod, timeout=3):
    xpath_all = ("//input[(@formcontrolname='frmCodigoMaterial' or @name='codigoMaterial' or @placeholder='CÓDIGO DO MATERIAL') and not(@disabled)]")
    end = time.time() + timeout
    visiveis = []
    while time.time() < end:
        elems = driver.find_elements(By.XPATH, xpath_all)
        visiveis = [e for e in elems if e.is_displayed()]
        if visiveis: break
        time.sleep(0.2)

    if not visiveis: raise RuntimeError("Nenhum campo 'CÓDIGO DO MATERIAL' visível encontrado.")
    alvo = visiveis[-1]
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", alvo)
    time.sleep(0.05)
    try: wait.until(EC.element_to_be_clickable((By.XPATH, xpath_all)))
    except TimeoutException: pass

    try: alvo.click()
    except Exception: driver.execute_script("arguments[0].focus();", alvo)
    clear_and_type(alvo, material_cod)

def click_incluir_produtos(driver, wait):
    wait_overlays_to_hide(wait)
    locator_btn = (By.XPATH, "//mat-action-row//button[.//span[normalize-space(.)='INCLUIR PRODUTO(S)'] and not(@disabled) and not(contains(@class,'mat-button-disabled'))]")
    try: btn = wait.until(EC.presence_of_element_located(locator_btn))
    except TimeoutException: raise RuntimeError("Botão 'INCLUIR PRODUTO(S)' não encontrado.")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    time.sleep(0.05)
    try: wait.until(EC.element_to_be_clickable(locator_btn))
    except TimeoutException: pass

    try:
        if not is_center_clickable_js(driver, btn):
            driver.execute_script("window.scrollBy(0, -80);")
            time.sleep(0.05)
        btn.click()
        return
    except Exception: pass

    try:
        ActionChains(driver).move_to_element(btn).pause(0.05).click().perform()
        return
    except Exception: pass

    try:
        driver.execute_script("arguments[0].click();", btn)
        return
    except Exception as e3: raise RuntimeError(f"Falha ao clicar em 'INCLUIR PRODUTO(S)': {repr(e3)}")

def abrir_select_justificativa(driver, wait):
    wait_overlays_to_hide(wait)
    loc_select = (By.CSS_SELECTOR, "mat-select[formcontrolname='justificativa']")
    try: sel = wait.until(EC.presence_of_element_located(loc_select))
    except TimeoutException:
        try: sel = wait.until(EC.presence_of_element_located((By.XPATH, "//mat-select[@formcontrolname='justificativa' or @aria-label='Por quê o Pedido não foi criado antes da emissão da Nota Fiscal?' or @placeholder='Por quê o Pedido não foi criado antes da emissão da Nota Fiscal?']")))
        except TimeoutException: raise RuntimeError("Campo select 'Justificativa' não encontrado.")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", sel)
    time.sleep(0.05)
    try: sel.click()
    except Exception:
        try:
            trigger = sel.find_element(By.CSS_SELECTOR, ".mat-select-trigger")
            trigger.click()
        except Exception:
            try: driver.execute_script("arguments[0].click();", sel)
            except Exception as e3: raise RuntimeError(f"Falha ao abrir o select 'Justificativa': {repr(e3)}")

    try: wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".cdk-overlay-pane .mat-select-panel")))
    except TimeoutException: raise RuntimeError("Painel de opções do 'Justificativa' não apareceu.")

def verificar_textos_na_tabela(driver, wait, textos, timeout=3):
    locator_container = (By.CSS_SELECTOR, "div.table-container")
    try: container = wait.until(EC.presence_of_element_located(locator_container))
    except TimeoutException: raise RuntimeError("Não encontrei o container da tabela.")

    conteudo = " ".join(container.text.split())
    encontrados = [t for t in textos if t and t in conteudo]
    faltando = [t for t in textos if t and t not in conteudo]
    return encontrados, faltando, conteudo

def click_ok_confirm(driver, wait, timeout=3, max_tentativas=3):
    cliques = 0
    for tentativa in range(max_tentativas):
        locator_dialog = (By.CSS_SELECTOR, "mat-dialog-container")
        try:
            WebDriverWait(driver, timeout if tentativa == 0 else 5, poll_frequency=0.3, ignored_exceptions=[StaleElementReferenceException]).until(EC.presence_of_element_located(locator_dialog))
        except TimeoutException:
            if cliques == 0: return 0
            else: return cliques

        locators_btn = [
            (By.CSS_SELECTOR, "mat-dialog-container button#btnTipoDoc"),
            (By.CSS_SELECTOR, "mat-dialog-container button.confirm"),
            (By.XPATH, "//mat-dialog-container//button[normalize-space()='OK']"),
        ]

        btn = None
        for by, sel in locators_btn:
            try:
                elementos = driver.find_elements(by, sel)
                visiveis = [e for e in elementos if _is_displayed_safe(e)]
                if visiveis:
                    btn = visiveis[-1]  
                    break
            except Exception: continue

        if btn is None:
            try:
                btn = WebDriverWait(driver, 5, poll_frequency=0.2, ignored_exceptions=[StaleElementReferenceException]).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "mat-dialog-container button#btnTipoDoc")))
            except TimeoutException:
                if cliques > 0: return cliques
                continue

        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(0.1)
        except StaleElementReferenceException: continue  

        metodo = _clicar_botao_ok(driver, btn)
        if metodo: cliques += 1
        else: continue

        try: WebDriverWait(driver, 5, poll_frequency=0.2).until(_dialog_desapareceu())
        except TimeoutException: continue
        time.sleep(0.5)
    return cliques

def _is_displayed_safe(element):
    try: return element.is_displayed()
    except (StaleElementReferenceException, Exception): return False

def _clicar_botao_ok(driver, btn):
    try:
        btn.click()
        return "clique direto"
    except (ElementClickInterceptedException, ElementNotInteractableException): pass
    except StaleElementReferenceException: return None

    try:
        ActionChains(driver).move_to_element(btn).pause(0.1).click().perform()
        return "ActionChains"
    except Exception: pass

    try:
        driver.execute_script("arguments[0].click();", btn)
        return "JavaScript"
    except Exception: pass
    return None

class _dialog_desapareceu:
    def __call__(self, driver):
        dialogs = driver.find_elements(By.CSS_SELECTOR, "mat-dialog-container")
        visiveis = [d for d in dialogs if _is_displayed_safe(d)]
        return len(visiveis) == 0
    
def wait_no_overlay(driver, wait):
    try: wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, ".cdk-overlay-backdrop.cdk-overlay-backdrop-showing, .mat-progress-spinner, .mat-progress-bar")))
    except TimeoutException: pass

def js_set_value_and_dispatch(driver, el, value: str):
    driver.execute_script("""
        const el = arguments[0], v = arguments[1];
        el.focus();
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
        setter.call(el, v);
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        el.blur();
    """, el, value)

def get_input_by_formcontrol(driver, wait, formcontrol):
    sel = f"input[formcontrolname='{formcontrol}']"
    el = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, sel)))
    scroll_center(driver, el)
    return el    

def esperar_transicao_apos_primeiro(wait, btn_clique=None, timeout=40):
    if btn_clique is not None:
        try:
            wait.until(EC.staleness_of(btn_clique))
            return
        except TimeoutException: pass
        try:
            wait.until(EC.invisibility_of_element(btn_clique))
            return
        except TimeoutException: pass
    time.sleep(0.3)
    wait_overlays_to_hide(wait)

def click_with_fallback(driver, el):
    try: el.click()
    except (ElementClickInterceptedException, ElementNotInteractableException): driver.execute_script("arguments[0].click();", el)

def type_safely(driver, wait, el, value: str):
    try:
        click_with_fallback(driver, el)
        el.send_keys(Keys.CONTROL, 'a', Keys.DELETE)
        if value is not None: el.send_keys(value)
        time.sleep(0.05)
        v = el.get_attribute("value") or ""
        if v.strip() != (value or "").strip(): js_set_value_and_dispatch(driver, el, value or "")
    except StaleElementReferenceException: raise

def is_center_clickable_js(driver, el):
    try:
        return driver.execute_script("""
            const el = arguments[0];
            const r = el.getBoundingClientRect();
            if (r.width === 0 || r.height === 0) return false;
            const cx = r.left + r.width/2, cy = r.top + r.height/2;
            const e = document.elementFromPoint(cx, cy);
            return e && (e === el || el.contains(e));
        """, el)
    except Exception: return False

def debug_dump(driver, prefix):
    try: driver.save_screenshot(f"{prefix}.png")
    except: pass
    try:
        with open(f"{prefix}.html","w",encoding="utf-8") as f: f.write(driver.page_source)
    except: pass

def ensure_enabled_and_editable(driver, el, allow_force=False):
    readonly = el.get_attribute("readonly")
    disabled = el.get_attribute("disabled")
    if readonly is not None or disabled is not None:
        if not allow_force: return False
        driver.execute_script("arguments[0].removeAttribute('readonly'); arguments[0].removeAttribute('disabled');", el)
    return True

def click_primeiro_continuar(driver, wait, campo_desc_css='input[formcontrolname="frmDescNota"]'):
    try:
        campo_desc = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, campo_desc_css)))
        campo_desc.send_keys(Keys.TAB)
        time.sleep(0.1)
    except TimeoutException: pass

    wait_overlays_to_hide(wait)
    candidatos = driver.find_elements(By.XPATH, "//button[.//span[normalize-space(.)='CONTINUAR']]")
    if not candidatos: raise RuntimeError("Nenhum <button> CONTINUAR encontrado.")

    alvo = None
    for btn in candidatos:
        try:
            vis = btn.is_displayed()
            hab = btn.is_enabled() and btn.get_attribute("disabled") in (None, "", "false")
            aria = btn.get_attribute("aria-disabled")
            if vis and hab and (aria in (None, "", "false")) and "mat-button-disabled" not in (btn.get_attribute("class") or ""):
                alvo = btn
                break
        except Exception: continue

    if not alvo:
        for btn in candidatos:
            try:
                if btn.is_displayed():
                    alvo = btn
                    break
            except Exception: continue

    if not alvo: raise RuntimeError("Nenhum <button> CONTINUAR visível.")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", alvo)
    time.sleep(0.05)
    try: wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[normalize-space(.)='CONTINUAR']]")))
    except TimeoutException: pass

    try:
        if not is_center_clickable_js(driver, alvo):
            driver.execute_script("window.scrollBy(0, -80);")
            time.sleep(0.05)
        alvo.click()
        return alvo  
    except Exception: pass

    try:
        ActionChains(driver).move_to_element(alvo).pause(0.05).click().perform()
        return alvo
    except Exception: pass

    try:
        driver.execute_script("arguments[0].click();", alvo)
        return alvo
    except Exception as e3: raise RuntimeError(f"Falha ao clicar no primeiro CONTINUAR: {repr(e3)}")

def click_segundo_continuar(driver, wait):
    wait_overlays_to_hide(wait)
    locator_link = (By.XPATH, "//a[.//span[normalize-space(.)='CONTINUAR'] and (not(@aria-disabled) or @aria-disabled='false') and not(contains(@class,'mat-button-disabled'))]")
    try: wait.until(EC.presence_of_element_located(locator_link))
    except TimeoutException: raise RuntimeError("Nenhum <a> CONTINUAR presente após o primeiro clique.")

    links = driver.find_elements(*locator_link)
    if not links: raise RuntimeError("Nenhum <a> CONTINUAR encontrado (filtro).")

    alvo = None
    for a in links:
        try:
            vis = a.is_displayed()
            hab = a.is_enabled()
            aria = a.get_attribute("aria-disabled")
            cls = a.get_attribute("class") or ""
            if vis and hab and (aria in (None, "", "false")) and "mat-button-disabled" not in cls:
                alvo = a
                break
        except Exception: continue

    if not alvo:
        for a in links:
            try:
                if a.is_displayed():
                    alvo = a
                    break
            except Exception: continue

    if not alvo: raise RuntimeError("Nenhum <a> CONTINUAR visível.")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", alvo)
    time.sleep(0.05)
    try: wait.until(EC.element_to_be_clickable(locator_link))
    except TimeoutException: pass

    try:
        if not is_center_clickable_js(driver, alvo):
            driver.execute_script("window.scrollBy(0, -80);")
            time.sleep(0.05)
        alvo.click()
        return
    except Exception: pass

    try:
        ActionChains(driver).move_to_element(alvo).pause(0.05).click().perform()
        return
    except Exception: pass

    try:
        driver.execute_script("arguments[0].click();", alvo)
        return
    except Exception as e3: raise RuntimeError(f"Falha ao clicar no segundo CONTINUAR (<a>): {repr(e3)}")

def click_pesquisar(driver, wait):
    wait_overlays_to_hide(wait)
    locator_btn = (By.XPATH, "//button[.//span[normalize-space(.)='Pesquisar'] and not(@disabled) and not(contains(@class,'mat-button-disabled'))]")
    try: btn = wait.until(EC.presence_of_element_located(locator_btn))
    except TimeoutException: raise RuntimeError("Botão 'Pesquisar' não encontrado na tela.")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
    time.sleep(0.05)
    try: wait.until(EC.element_to_be_clickable(locator_btn))
    except TimeoutException: pass

    try:
        if not is_center_clickable_js(driver, btn):
            driver.execute_script("window.scrollBy(0, -80);")
            time.sleep(0.05)
        btn.click()
        return
    except Exception: pass

    try:
        ActionChains(driver).move_to_element(btn).pause(0.05).click().perform()
        return
    except Exception: pass

    try:
        driver.execute_script("arguments[0].click();", btn)
        return
    except Exception as e3: raise RuntimeError(f"Falha ao clicar no botão 'Pesquisar': {repr(e3)}")

def _aguardar_checkbox_interagivel(driver, timeout=40):
    locator_checkbox = (By.XPATH, "(//td[contains(@class,'mat-column-select')]//mat-checkbox//input[@type='checkbox'])[1]")
    wait = WebDriverWait(driver, timeout, poll_frequency=0.3, ignored_exceptions=[StaleElementReferenceException])
    try: cb_input = wait.until(EC.presence_of_element_located(locator_checkbox))
    except TimeoutException: raise RuntimeError(f"Checkbox não apareceu no DOM após {timeout}s.")

    try: cb_input = wait.until(EC.visibility_of_element_located(locator_checkbox))
    except TimeoutException:
        locator_mat_checkbox = (By.XPATH, "(//td[contains(@class,'mat-column-select')]//mat-checkbox)[1]")
        try: wait.until(EC.visibility_of_element_located(locator_mat_checkbox))
        except TimeoutException: raise RuntimeError(f"mat-checkbox não ficou visível após {timeout}s.")

    _aguardar_posicao_estavel(driver, cb_input, tentativas=5, intervalo=0.3)
    return cb_input

def _aguardar_posicao_estavel(driver, elemento, tentativas=5, intervalo=0.3):
    pos_anterior = None
    for _ in range(tentativas):
        try:
            rect = driver.execute_script(
                "var r = arguments[0].getBoundingClientRect();"
                "return {top: r.top, left: r.left, width: r.width, height: r.height};",
                elemento,
            )
        except StaleElementReferenceException:
            time.sleep(intervalo)
            continue

        if pos_anterior and rect == pos_anterior: return  
        pos_anterior = rect
        time.sleep(intervalo)

def _obter_label_do_checkbox(driver, cb_input):
    try:
        input_id = cb_input.get_attribute("id")
        if input_id: return driver.find_element(By.XPATH, f"//label[@for='{input_id}']")
    except Exception: pass

    try: return driver.find_element(By.XPATH, "(//td[contains(@class,'mat-column-select')]//mat-checkbox//label)[1]")
    except Exception: return None

def _clicar_com_fallback(driver, alvo, descricao="elemento"):
    try:
        alvo.click()
        return "Clique direto"
    except (ElementClickInterceptedException, Exception): pass

    try:
        ActionChains(driver).move_to_element(alvo).pause(0.1).click().perform()
        return "ActionChains"
    except Exception: pass

    try:
        driver.execute_script("arguments[0].click();", alvo)
        return "Clique via JS"
    except Exception as e: raise RuntimeError(f"Todas as estratégias de clique falharam para {descricao}: {repr(e)}")

def selecionar_primeira_linha_checkbox(driver, wait, timeout=40, textos_para_verificar=None, exigir_todos=False, clicar_mesmo_se_faltar=False):
    wait_overlays_to_hide(wait)
    resultado_textos = {"verificacao_feita": False, "encontrados": [], "faltando": []}

    if textos_para_verificar:
        encontrados, faltando, _conteudo = verificar_textos_na_tabela(driver, wait, textos_para_verificar, timeout=timeout)
        resultado_textos.update({"verificacao_feita": True, "encontrados": encontrados, "faltando": faltando})
        if faltando and exigir_todos and not clicar_mesmo_se_faltar: raise RuntimeError(f"Textos faltando na tabela: {faltando} | Encontrados: {encontrados}")
        if faltando and not clicar_mesmo_se_faltar: return {**resultado_textos, "clicou": False, "motivo": f"Faltando textos: {faltando}"}

    try: cb_input = _aguardar_checkbox_interagivel(driver, timeout=timeout)
    except RuntimeError: raise

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cb_input)
    time.sleep(0.15)
    label = _obter_label_do_checkbox(driver, cb_input)
    alvo = label if label else cb_input

    try:
        if not is_center_clickable_js(driver, alvo):
            driver.execute_script("window.scrollBy(0, -100);")
            time.sleep(0.15)
    except Exception: pass  

    metodo = _clicar_com_fallback(driver, alvo, descricao="checkbox primeira linha")
    return {**resultado_textos, "clicou": True, "motivo": metodo}

def preencher_quantidade_e_valor(driver, wait, quantidade="1", valor_boleto="123,45"):
    wait_overlays_to_hide(wait)
    locator_qtd = (By.XPATH, "//input[@id='quantidade' or @name='quantidade' or @formcontrolname='frmQuantidade']")
    try: qtd = wait.until(EC.presence_of_element_located(locator_qtd))
    except TimeoutException: raise RuntimeError("Campo 'Quantidade' não encontrado.")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", qtd)
    time.sleep(0.05)
    try: wait.until(EC.element_to_be_clickable(locator_qtd))
    except TimeoutException: pass

    try: clear_and_type(qtd, quantidade)
    except Exception: driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input',{bubbles:true}));", qtd, quantidade)

    locator_valor = (By.XPATH, "//input[@id='valorUnitario' or @name='valorUnitario' or @formcontrolname='frmValor']")
    try: valor = wait.until(EC.presence_of_element_located(locator_valor))
    except TimeoutException: raise RuntimeError("Campo 'Valor Unitário' não encontrado.")

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", valor)
    time.sleep(0.05)
    try: wait.until(EC.element_to_be_clickable(locator_valor))
    except TimeoutException: pass

    try: clear_and_type(valor, valor_boleto)
    except Exception:
        somente_digitos = re.sub(r"\D", "", str(valor_boleto))
        try: clear_and_type(valor, somente_digitos)
        except Exception: driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input',{bubbles:true}));", valor, somente_digitos)

def selecionar_opcao_justificativa_com_hover(driver, wait, texto_alvo="2 - Orientações do gestor/coordendor da área"):
    locator_option_exata = (By.XPATH, f"//mat-option//span[contains(normalize-space(.), '{texto_alvo.split(' - ')[-1].split()[0]}') and contains(normalize-space(.), 'Orientações do gestor')]")
    try: opt_span = wait.until(EC.presence_of_element_located(locator_option_exata))
    except TimeoutException:
        try: opt_span = wait.until(EC.presence_of_element_located((By.XPATH, "//mat-option//span[contains(normalize-space(.), 'Orientações do gestor')]")))
        except TimeoutException: raise RuntimeError("Não encontrei a opção de justificativa no painel.")

    try: mat_option = opt_span.find_element(By.XPATH, "./ancestor::mat-option[1]")
    except Exception: mat_option = opt_span

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", mat_option)
    time.sleep(0.05)

    try:
        ActionChains(driver).move_to_element(mat_option).pause(0.15).click().perform()
        return
    except Exception: pass

    try:
        mat_option.click()
        return
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", mat_option)
            return
        except Exception: pass

    try:
        from selenium.webdriver.common.keys import Keys
        body = driver.find_element(By.TAG_NAME, "body")
        for _ in range(10):  
            body.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.05)
        body.send_keys(Keys.ENTER)
    except Exception as e3: raise RuntimeError(f"Falha ao selecionar a justificativa (mesmo com hover): {repr(e3)}")

def _strip_accents(s: str) -> str:
    if pd.isna(s) or s is None: return ''
    s = unicodedata.normalize('NFKD', str(s))
    return ''.join(ch for ch in s if not unicodedata.combining(ch))

def _norm_colname(s: str) -> str:
    return _strip_accents(str(s)).lower().strip()

def _clean_str(x) -> str:
    if pd.isna(x): return ''
    if isinstance(x, float):
        if x.is_integer(): return str(int(x))
        return str(x)
    s = str(x).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

def _norm_coletor(x: str) -> str:
    up = _strip_accents(_clean_str(x)).upper()
    return re.sub(r'[^A-Z0-9]', '', up)

def _is_valid_coletor(coletor: str) -> bool:
    if pd.isna(coletor) or str(coletor).strip() == '' or str(coletor).strip().lower() == 'nan': return False
    c = _strip_accents(_clean_str(coletor)).upper()
    c_clean = re.sub(r'[^A-Z0-9]', '', c)
    if re.fullmatch(r'\d{6,}', c_clean): return True
    if 6 <= len(c_clean) <= 12 and any(x.isalpha() for x in c_clean) and any(x.isdigit() for x in c_clean): return True
    return False

def _tipo_de_coletor(coletor: str) -> str:
    c = _norm_coletor(coletor)
    if c == 'SEMCENTRODECUSTO': return '-'
    if re.fullmatch(r'\d+', c): return 'N'
    if any(x.isalpha() for x in c) and any(x.isdigit() for x in c): return 'K'
    return ''

def _clean_valor_series(s: pd.Series) -> pd.Series:
    def limpa_valor(val):
        if pd.isna(val) or str(val).strip() == '': return None
        if isinstance(val, (int, float)): return float(val)
        v = str(val).upper().replace('R$', '').replace('\xa0', '').replace(' ', '').strip()
        if '.' in v and ',' in v: v = v.replace('.', '').replace(',', '.')
        elif ',' in v: v = v.replace(',', '.')
        try: return float(v)
        except ValueError: return None
    return s.apply(limpa_valor)

def ler_rr_bruto(caminho_rr: Union[str, Path]) -> pd.DataFrame:
    print("[PROGRESSO: 40]")
    xls = pd.ExcelFile(caminho_rr, engine='openpyxl')
    frames = []
    
    for sh in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sh, header=None, engine='openpyxl')
        idx_ancora = -1
        max_knf = 0
        
        for j in range(df.shape[1]):
            col_str = df.iloc[:, j].astype(str).str.strip().str.upper()
            count_knf = col_str.isin(['K', 'N', 'F']).sum()
            if count_knf > max_knf and count_knf >= 2:
                max_knf = count_knf
                idx_ancora = j
                
        if idx_ancora == -1: continue
            
        df_regional = df.iloc[:, idx_ancora:].copy()
        df_regional.columns = range(df_regional.shape[1])
        
        col_coletor, col_subnum, col_valor = -1, -1, -1
        header_row_idx = -1
        
        for i, row in df_regional.head(20).iterrows():
            row_norm = [_norm_colname(str(x)) for x in row.values]
            if 'coletor' in row_norm or 'subnumero' in row_norm:
                for j in range(len(row_norm)):
                    val = row_norm[j]
                    if val == 'coletor' and col_coletor == -1: col_coletor = j
                    elif 'subnumero' in val and col_subnum == -1: col_subnum = j
                    elif 'valor' in val and 'servico' not in val and col_valor == -1: col_valor = j
                
                if col_valor == -1:
                    for j in range(len(row_norm)):
                        if 'valor' in row_norm[j] and col_valor == -1: col_valor = j
                        
                if col_coletor != -1 and col_valor != -1:
                    header_row_idx = i
                    break

        if header_row_idx != -1:
            tmp = df_regional.iloc[header_row_idx + 1:].copy()
            def safe_get(c): return tmp.iloc[:, c] if c != -1 and c < tmp.shape[1] else pd.Series(['']*len(tmp))
            tmp_clean = pd.DataFrame({
                'COLETOR_ORIG': safe_get(col_coletor),
                'SUBNUM_ORIG': safe_get(col_subnum),
                'VALOR': safe_get(col_valor)
            })
        else:
            col_coletor_tb = -1
            max_validos = 0
            for j in range(1, df_regional.shape[1]):
                mask = df_regional.iloc[:, j].astype(str).apply(_is_valid_coletor)
                qtd = mask.sum()
                if qtd > max_validos and qtd >= 2:
                    max_validos = qtd
                    col_coletor_tb = j
                    
            col_valor_tb = -1
            max_nums = 0
            for j in range(1, df_regional.shape[1]):
                if j == col_coletor_tb: continue
                nums = _clean_valor_series(df_regional.iloc[:, j]).notna().sum()
                if nums > max_nums:
                    max_nums = nums
                    col_valor_tb = j
                    
            if col_coletor_tb == -1 or col_valor_tb == -1: continue
                
            tmp_clean = pd.DataFrame({
                'COLETOR_ORIG': df_regional.iloc[:, col_coletor_tb],
                'SUBNUM_ORIG': pd.Series(['']*len(df_regional)), 
                'VALOR': df_regional.iloc[:, col_valor_tb]
            })

        tmp_clean['VALOR'] = _clean_valor_series(tmp_clean['VALOR'])
        tmp_clean = tmp_clean.dropna(subset=['VALOR'])
        
        def get_best_coletor(r):
            sub = _clean_str(r['SUBNUM_ORIG'])
            if _is_valid_coletor(sub): return sub
            col = _clean_str(r['COLETOR_ORIG'])
            if _is_valid_coletor(col): return col
            return None
            
        tmp_clean['COLETOR_FINAL'] = tmp_clean.apply(get_best_coletor, axis=1)
        tmp_clean = tmp_clean.dropna(subset=['COLETOR_FINAL'])
        
        if not tmp_clean.empty:
            tmp_clean['COLETOR'] = tmp_clean['COLETOR_FINAL'].apply(_norm_coletor)
            tmp_clean['TIPOCOLETOR'] = tmp_clean['COLETOR'].apply(_tipo_de_coletor)
            frames.append(tmp_clean[['TIPOCOLETOR', 'COLETOR', 'VALOR']])

    if frames: return pd.concat(frames, ignore_index=True)
    return pd.DataFrame(columns=['TIPOCOLETOR', 'COLETOR', 'VALOR'])

def _extrair_coletor_de_titular(texto: str) -> str:
    if pd.isna(texto) or str(texto).strip() == '': return "SEM CENTRO DE CUSTO"
    t = _strip_accents(str(texto)).upper()
    
    m = re.search(r'(?<!\d)(\d{6,})(?!\d)', t)
    if m: return _norm_coletor(m.group(1))
    
    palavras = t.split()
    for p in palavras:
        p_clean = re.sub(r'[^A-Z0-9]', '', p)
        if 6 <= len(p_clean) <= 12 and any(c.isalpha() for c in p_clean) and any(c.isdigit() for c in p_clean):
            return p_clean
            
    for i in range(len(palavras) - 1):
        p1 = re.sub(r'[^A-Z0-9]', '', palavras[i])
        p2 = re.sub(r'[^A-Z0-9]', '', palavras[i+1])
        comb = p1 + p2
        if 8 <= len(comb) <= 12 and any(c.isalpha() for c in comb) and any(c.isdigit() for c in comb):
            return comb
            
    return "SEM CENTRO DE CUSTO"

def ler_correios_bruto(caminho_correios: Union[str, Path]) -> Tuple[pd.DataFrame, float]:
    print("[PROGRESSO: 60]")
    df_raw = pd.read_excel(caminho_correios, header=None, engine='openpyxl')
    idx_header = -1
    col_titular, col_valor = -1, -1
    
    for i, row in df_raw.head(20).iterrows():
        row_norm = [_norm_colname(str(x)) for x in row.values]
        if any('titular do cartao' in c for c in row_norm) and any('valor do servico' in c for c in row_norm):
            idx_header = i
            for j, c in enumerate(row_norm):
                if 'titular do cartao' in c: col_titular = j
                if 'valor do servico' in c: col_valor = j
            break
            
    valor_liquido = 0.0
    idx_fim_tabela = len(df_raw)
    
    for i in range(len(df_raw) - 1, -1, -1):
        row_norm = [_norm_colname(str(x)) for x in df_raw.iloc[i].values]
        if any('valor liquido' in c for c in row_norm):
            idx_fim_tabela = i
            for j, c in enumerate(row_norm):
                if 'valor liquido' in c:
                    if i + 1 < len(df_raw):
                        val_raw = df_raw.iloc[i + 1, j]
                        valor_liquido = _clean_valor_series(pd.Series([val_raw])).iloc[0]
                    break
            break

    if idx_header == -1: 
        return pd.DataFrame(columns=['TIPOCOLETOR', 'COLETOR', 'VALOR']), valor_liquido
    
    df = df_raw.iloc[idx_header + 1 : idx_fim_tabela, [col_titular, col_valor]].copy()
    df.columns = ['TITULAR', 'VALOR']
    
    mask_ignorar = df['TITULAR'].astype(str).str.upper().str.contains('ENCARGO|DESCONTO|CREDITO')
    df = df[~mask_ignorar].copy()
    
    df['VALOR'] = _clean_valor_series(df['VALOR'])
    df = df.dropna(subset=['VALOR'])
    df['COLETOR'] = df['TITULAR'].apply(_extrair_coletor_de_titular)
    df['TIPOCOLETOR'] = df['COLETOR'].apply(_tipo_de_coletor)
    
    return df[['TIPOCOLETOR', 'COLETOR', 'VALOR']], valor_liquido

def _formatar_planilha_final(arquivo_xlsx: Union[str, Path], sheet='Planilha1'):
    wb = load_workbook(arquivo_xlsx)
    if sheet not in wb.sheetnames:
        wb.close(); return
    ws = wb[sheet]
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    try:
        idx_valor = header.index('VALOR') + 1
        idx_op = header.index('OPERACAO') + 1
    except ValueError:
        wb.close(); return
        
    for row in ws.iter_rows(min_row=2):
        cell_valor = row[idx_valor - 1]
        if isinstance(cell_valor.value, (int, float)):
            cell_valor.number_format = numbers.FORMAT_NUMBER_00
            
    wb.save(arquivo_xlsx)
    wb.close()

def gerar_rateio_pag(
    caminho_correios: Union[str, Path],
    caminho_rr: Union[str, Path],
    saida: Union[str, Path] = 'RATEIO PAG.xlsx',
    operacao_para_diagrama: int = 10,
    tolerancia_igual: float = 0.05, 
    debug: bool = True
) -> pd.DataFrame:

    df_rr_raw = ler_rr_bruto(caminho_rr)
    df_corr_raw, valor_liquido_correios = ler_correios_bruto(caminho_correios)

    print("[PROGRESSO: 80]")
    df_rr_raw['VALOR'] = pd.to_numeric(df_rr_raw['VALOR'], errors='coerce').fillna(0.0)
    df_corr_raw['VALOR'] = pd.to_numeric(df_corr_raw['VALOR'], errors='coerce').fillna(0.0)

    total_rr = float(df_rr_raw['VALOR'].sum()) if not df_rr_raw.empty else 0.0
    total_corr_soma = float(df_corr_raw['VALOR'].sum()) if not df_corr_raw.empty else 0.0
    total_corr = valor_liquido_correios if valor_liquido_correios > 0 else total_corr_soma

    if debug:
        print(f"[DEBUG] TOTAL RR               = R$ {total_rr:.2f}")
        print(f"[DEBUG] TOTAL CORREIOS (SOMA)  = R$ {total_corr_soma:.2f}")
        print(f"[DEBUG] TOTAL CORREIOS (LÍQ)   = R$ {valor_liquido_correios:.2f}")
        print(f"[DEBUG] DIFERENÇA (LÍQ - RR)   = R$ {round(total_corr - total_rr, 2):.2f}")

    linhas_finais = df_rr_raw.to_dict('records') if not df_rr_raw.empty else []

    if abs(total_corr - total_rr) > tolerancia_igual:
        if debug: print("[DEBUG] Valores divergem. Buscando pacotes exatos faltantes...")
        
        rr_valores_exatos = df_rr_raw['VALOR'].round(2).tolist()
        
        saldo_rr_cc = {}
        for _, row in df_rr_raw.iterrows():
            c = row['COLETOR']
            saldo_rr_cc[c] = saldo_rr_cc.get(c, 0.0) + row['VALOR']
            
        saldo_global_rr = total_rr
        pacotes_faltantes = []
        
        for _, row in df_corr_raw.iterrows():
            c = row['COLETOR']
            v = round(row['VALOR'], 2)
            tipo = row['TIPOCOLETOR']
            
            matched = False
            
            if c in saldo_rr_cc and saldo_rr_cc[c] >= v - 0.02:
                saldo_rr_cc[c] -= v
                saldo_global_rr -= v
                matched = True
                for i, rr_v in enumerate(rr_valores_exatos):
                    if abs(rr_v - v) <= 0.02:
                        rr_valores_exatos.pop(i)
                        break
            
            if not matched:
                for i, rr_v in enumerate(rr_valores_exatos):
                    if abs(rr_v - v) <= 0.02:
                        rr_valores_exatos.pop(i)
                        saldo_global_rr -= v
                        matched = True
                        break
                        
            if not matched:
                pacotes_faltantes.append({
                    'TIPOCOLETOR': tipo,
                    'COLETOR': c,
                    'VALOR': v
                })
                
        pacotes_adicionados = 0
        for pct in pacotes_faltantes:
            v = pct['VALOR']
            if saldo_global_rr >= v - 0.02:
                saldo_global_rr -= v
            else:
                linhas_finais.append(pct)
                pacotes_adicionados += 1
                
        if debug: print(f"[DEBUG] Adicionados {pacotes_adicionados} pacotes exatos dos Correios.")

    print("[PROGRESSO: 90]")
    final_base = pd.DataFrame(linhas_finais)
    if not final_base.empty:
        final_base = final_base.groupby(['TIPOCOLETOR', 'COLETOR'], as_index=False)['VALOR'].sum()

    soma_atual = final_base['VALOR'].sum() if not final_base.empty else 0.0
    diferenca_rateio = round(total_corr - soma_atual, 2)
    
    if abs(diferenca_rateio) > 0.02 and not final_base.empty:
        if debug: print(f"[DEBUG] Rateando R$ {diferenca_rateio:.2f} (Encargos/Descontos) proporcionalmente...")
        
        soma_validos = final_base['VALOR'].sum()
        if soma_validos > 0:
            final_base['VALOR_ADD'] = (final_base['VALOR'] / soma_validos) * diferenca_rateio
            final_base['VALOR_ADD'] = final_base['VALOR_ADD'].round(2)
            
            diff_centavos = round(diferenca_rateio - final_base['VALOR_ADD'].sum(), 2)
            if diff_centavos != 0:
                idx_max = final_base['VALOR'].idxmax() 
                final_base.loc[idx_max, 'VALOR_ADD'] += diff_centavos 
                
            final_base['VALOR'] += final_base['VALOR_ADD']
            final_base = final_base.drop(columns=['VALOR_ADD'])

    final = pd.DataFrame()
    if not final_base.empty:
        final['ITEM'] = [1] * len(final_base)
        final['TIPOCOLETOR'] = final_base['TIPOCOLETOR']
        final['COLETOR'] = final_base['COLETOR']
        final['OPERACAO'] = final_base['TIPOCOLETOR'].apply(lambda t: operacao_para_diagrama if t == 'N' else '')
        final['SUBNUMERO'] = ''
        final['VALOR'] = final_base['VALOR']
        final['DESCRICAO'] = ''

        final['__ord'] = final['TIPOCOLETOR'].map({'K': 0, 'N': 1}).fillna(2)
        final = final.sort_values(['__ord', 'COLETOR']).drop(columns='__ord').reset_index(drop=True)

    saida = Path(saida)
    with pd.ExcelWriter(saida, engine='openpyxl') as writer:
        final.to_excel(writer, sheet_name='Planilha1', index=False)

    _formatar_planilha_final(saida, 'Planilha1')

    if debug: print(f"[DEBUG] Arquivo gerado com sucesso: {saida.resolve()}")

    return final
