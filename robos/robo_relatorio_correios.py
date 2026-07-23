import time
import os
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.keys import Keys
import win32com.client

# Importando as credenciais do seu arquivo config.py
from config import EMAIL_MRV, SENHA_MRV

# --- CONFIGURAÇÃO ---
PASTA_DOWNLOAD = "C:/Users/pedro.henrsilva/Downloads"
NOME_ARQUIVO_FINAL = "Produtividade_EDITADO.xlsx"
PADRAO_RASTREIO = re.compile(r'\b[A-Z]{2}\d{9}[A-Z]{2}\b')
# --------------------

# ==============================================================================
# FUNÇÃO 1: VALIDAÇÃO DE CHAMADOS (Ferramenta auxiliar)
# ==============================================================================
def validar_chamado_no_agilis(chamado, driver, wait):
    try:
        print(f"   Validando chamado: {chamado}")

        # --- PASSO 1: Clicar na LUPA ---
        print("   - Clicando na lupa...")
        lupa_clicada = False
        estrategias_lupa = [
            (By.XPATH, "//*[@aria-label='Pesquisa']"),
            (By.XPATH, "//span[.//*[@aria-label='Pesquisa']]"),
            (By.XPATH, "//*[@viewBox='0 0 24 24']"),
            (By.XPATH, "//span[.//*[@viewBox='0 0 24 24']]"),
            (By.XPATH, "//*[contains(@class,'header-menu-icons')]"),
            (By.XPATH, "//span[.//*[contains(@class,'header-menu-icons')]]"),
            (By.XPATH, "//span[contains(@class,'search')]"),
            (By.XPATH, "//input[@id='subheader_search_box']/preceding-sibling::*[1]"),
            (By.XPATH, "//*[@role='img'][@aria-label='Pesquisa']"),
            (By.XPATH, "//*[contains(@style,'stroke-miterlimit')]/ancestor::span[1]"),
        ]

        for i, (by, seletor) in enumerate(estrategias_lupa, 1):
            try:
                elemento = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((by, seletor))
                )
                driver.execute_script("arguments[0].click();", elemento)
                lupa_clicada = True
                print(f"   ✅ Lupa clicada na tentativa {i}.")
                break
            except Exception:
                continue

        if not lupa_clicada:
            raise Exception("Não foi possível clicar na lupa após todas as tentativas.")

        time.sleep(0.8)

        # --- PASSO 2: Limpar e digitar o chamado ---
        print(f"   - Digitando chamado {chamado}...")
        campo_busca = wait.until(EC.element_to_be_clickable((By.ID, "subheader_search_box")))
        campo_busca.click()
        campo_busca.send_keys(Keys.CONTROL + "a")
        campo_busca.send_keys(Keys.DELETE)
        campo_busca.send_keys(str(chamado))
        campo_busca.send_keys(Keys.RETURN)
        time.sleep(2)

        # --- PASSO 3 e 4: Encontrar TODOS os painéis de resposta e iterar sobre eles ---
        print("   - Procurando painéis de conversa...")
        try:
            paineis = wait.until(EC.presence_of_all_elements_located(
                (By.CSS_SELECTOR, "z-collapsiblepanel[data-conv_type='reply']")
            ))
            print(f"   - Encontrados {len(paineis)} painéis de resposta. Verificando um por um...")

            for index, painel in enumerate(paineis):
                print(f"   - Analisando painel {index + 1} de {len(paineis)}...")
                
                aria_expanded = painel.get_attribute("aria-expanded")
                if aria_expanded != "true":
                    try:
                        header = painel.find_element(
                            By.CSS_SELECTOR,
                            "div.zcollapsiblepanel__header.zcollapsiblepanel--toggleableheader"
                        )
                        driver.execute_script("arguments[0].click();", header)
                        time.sleep(0.5) 
                    except Exception as e:
                        print(f"     ⚠️ Erro ao expandir painel {index + 1}: {e}")
                
                try:
                    painel_body = WebDriverWait(driver, 8).until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR, "z-collapsiblepanel[data-conv_type='reply'] div.panel-body.p0")
                        )
                    )

                    conteudo = painel_body.text
                    conteudo_lower = conteudo.lower()
                    
                    codigo_encontrado = PADRAO_RASTREIO.search(conteudo)
                    if codigo_encontrado:
                        print(f"   ✅ Chamado {chamado} MANTIDO (código de rastreio encontrado no painel {index + 1}: {codigo_encontrado.group()}).")
                        return True

                    if "encomenda enviada" in conteudo_lower or "encomenda recebida" in conteudo_lower:
                        print(f"   ✅ Chamado {chamado} MANTIDO (contém 'encomenda enviada' ou 'encomenda recebida' no painel {index + 1}).")
                        return True

                    if "segue o código de rastreio:" in conteudo_lower:
                        print(f"   ✅ Chamado {chamado} MANTIDO (contém 'segue o código de rastreio' no painel {index + 1}).")
                        return True

                except Exception as e:
                    print(f"     ⚠️ Não foi possível ler o corpo do painel {index + 1}.")
                    continue 

            print(f"   ❌ Nenhum critério de manutenção encontrado nos {len(paineis)} painéis.")

        except TimeoutException:
            print("   - Nenhum painel de resposta encontrado. Indo para verificação de status...")
            return False
        
    except TimeoutException:
        print(f"   ⚠️ Timeout ao validar chamado {chamado}. Removendo por falta de resposta.")
        return False
    except Exception as e:
        print(f"   ⚠️ Erro inesperado ao validar chamado {chamado}: {e}")
        return True

# ==============================================================================
# FUNÇÃO 2: BAIXAR RELATÓRIO NO AGILIS (SELENIUM)
# ==============================================================================
def baixar_relatorio_agilis():
    print("[PROGRESSO: 5]")
    print("Iniciando automação do relatório de envio dos correios...")
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 10)
    
    URL_INICIAL = "https://agilis.mrv.com.br/HomePage.do?view_type=my_view"
    try:
        driver.get(URL_INICIAL)
        print(f"Página aberta: {URL_INICIAL}")
        print("Aguardando tela de login...")

        try:
            selector_login_integrado = (By.LINK_TEXT, "Login Integrado Microsoft")
            wait.until(EC.element_to_be_clickable(selector_login_integrado)).click()
        except:
            print("Não encontrou por LINK_TEXT. Tentando por XPATH...")
            selector_login_integrado = (By.XPATH, "//*[text()='Login Integrado Microsoft']")
            wait.until(EC.element_to_be_clickable(selector_login_integrado)).click()

        print("0. Cliquei em 'Login Integrado Microsoft'.")
        
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

        print("[PROGRESSO: 15]")
        print("Login da Microsoft concluído na janela pop-up.")

        selector_relatorios = (By.LINK_TEXT, "Relatórios")
        wait.until(EC.element_to_be_clickable(selector_relatorios)).click()
        print("1. Cliquei em 'Relatórios'.")

        selector_contratos_adm = (By.LINK_TEXT, "Contratos - ADM")
        wait.until(EC.element_to_be_clickable(selector_contratos_adm)).click()
        print("   - Cliquei em 'Contratos - ADM'.")

        selector_produtividade = (By.LINK_TEXT, "Produtividade Contratos - ADM")
        wait.until(EC.element_to_be_clickable(selector_produtividade)).click()
        print("   - Cliquei em 'Produtividade Contratos - ADM'.")

        selector_editar = (By.CLASS_NAME, "linkborder") 
        wait.until(EC.element_to_be_clickable(selector_editar)).click()
        print("3. Cliquei em 'Editar'.")

        selector_coletor = (By.XPATH, "//option[text()='Coletor de custo ADM']")
        wait.until(EC.element_to_be_clickable(selector_coletor)).click()
        print("4. Selecionei 'Coletor de custo ADM'.")
        
        selector_seta_direita = (By.CLASS_NAME, "moverightButton") 
        driver.find_element(*selector_seta_direita).click()
        print("   - Cliquei na seta para mover.")

        print("4.5. Expandindo 'Passo 2: Opções de filtragem'...")
        try:
            elemento_clique = wait.until(EC.presence_of_element_located((By.ID, "rcstep2src")))
            driver.execute_script("arguments[0].click();", elemento_clique)
            print("   - SUCESSO: Cliquei em 'Opções de filtragem' (via JavaScript).")
            time.sleep(1) 
        except TimeoutException:
            print("    - FALHA: Não foi possível encontrar 'Passo 2: Opções de filtragem'.")
            raise 

        print("5. Selecionando o filtro 'Durante'...")
        try: 
            selector_radio_durante = (By.CSS_SELECTOR, "input[value='predefined']")
            wait.until(EC.element_to_be_clickable(selector_radio_durante)).click()
            print("    - SUCESSO: Filtro 'Durante' selecionado.")
        except TimeoutException:
            print("    - FALHA: Não foi possível encontrar o rádio 'Durante'.")
            raise 

        print("[PROGRESSO: 25]")
        selector_executar = (By.ID, "addnew223222")
        wait.until(EC.element_to_be_clickable(selector_executar)).click()
        print("6. Cliquei em 'Executar relatório'.")
        print("--- Relatório executado, aguardando 10s para carregar...")
        time.sleep(3) 

        print("7. Iniciando o download direto do relatório XLSX...")
        try:
            DOWNLOAD_XLSX_LINK = (By.ID, "exportxlsx")
            wait.until(EC.element_to_be_clickable(DOWNLOAD_XLSX_LINK)).click()
            print("   - Clique realizado no link 'Exportar arquivo como XLSX'.")
            print("   - Aguardando o download ser concluído...")
            time.sleep(5) 
            print("   - Relatório baixado com sucesso!")
        except Exception as e:
            print(f"ERRO ao tentar baixar o relatório XLS: {e}")
        
        print("--- Processo do site completo!!! ---")
        
        # IMPORTANTE: Retorna o driver aberto para a próxima função usar!
        return driver 

    except Exception as e:
        print(f"ERRO: A automação falhou. {e}")
        driver.quit()
        return None

# ==============================================================================
# FUNÇÃO 3: PROCESSAR EXCEL E VALIDAR CHAMADOS
# ==============================================================================
def processar_excel_e_validar(driver):
    print("[PROGRESSO: 35]")
    # Constantes de formatação do Excel
    xlUp = -4162
    xlCenter = -4108           
    xlTop = -4160              
    xlContinuous = 1           
    
    arquivo_salvo_path = None
    excel = None
    wb = None

    try:
        print(f"Procurando o relatório mais recente na pasta: {PASTA_DOWNLOAD}")
        arquivos_xls = [f for f in os.listdir(PASTA_DOWNLOAD) if f.lower().endswith('.xls') or f.lower().endswith('.xlsx')]

        if not arquivos_xls:
            raise FileNotFoundError(f"Nenhum arquivo encontrado na pasta {PASTA_DOWNLOAD}")

        caminhos_completos = [os.path.join(PASTA_DOWNLOAD, f) for f in arquivos_xls]
        arquivo_salvo_path = max(caminhos_completos, key=os.path.getmtime)
        print(f"Arquivo mais recente encontrado: {arquivo_salvo_path}")

    except Exception as e:
        print(f"ERRO ao tentar encontrar o arquivo de relatório: {e}")
        driver.quit()
        return

    try:
        print("\n--- Iniciando edição no Excel ---")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True  
        excel.DisplayAlerts = False 

        wb = excel.Workbooks.Open(arquivo_salvo_path)
        ws = wb.Worksheets(1) 

        print("Editando: Excluindo linhas 1-8...")
        ws.Rows("1:8").Delete()
        
        print("Editando: Removendo logos/imagens...")
        for shape in ws.Shapes:
            shape.Delete()
        
        print("Editando: Ajustando altura e largura...")
        ws.Rows.RowHeight = 12
        ws.Columns.ColumnWidth = 12
        
        print("Editando: Removendo linhas que não são de 'Solicitação de Envio de Correspondência'...")
        filtro_texto = "Solicitação de Envio de Correspondência"
        last_row = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
        
        for i in range(last_row, 1, -1): 
            cell_value = str(ws.Cells(i, 9).Value) 
            if filtro_texto not in cell_value:
                ws.Rows(i).Delete()
        print("Linhas indesejadas removidas.")

        print("\nIniciando verificação de horário para remoção adicional...")
        hora_atual = datetime.now().hour

        if hora_atual >= 12:
            print(f"A hora atual ({hora_atual}:00) é 12:00 ou mais. Removendo correspondências matutinas.")
            last_row = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
            for i in range(last_row, 1, -1):
                cell_text_k = str(ws.Cells(i, 11).Text)
                if re.search(r'(^|\s)(06|07|08|09|10|11):', cell_text_k):
                    ws.Rows(i).Delete()
            print("Remoção de correspondências matutinas concluída.")
        else:
            print(f"A hora atual ({hora_atual}:00) ainda é de manhã. Nenhuma remoção adicional será feita.")
        
        print("Editando: Adicionando fórmulas nas colunas M-R...")
        last_row = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        if last_row > 1:
            formula_N = (
                '=LET('
                'txt,F2,'
                'start,IFERROR(SEARCH("Código:",txt)+7,'
                'IFERROR(SEARCH("Coletor de Custo ADM",txt)+22,'
                'IFERROR(SEARCH("Centro de custo:",txt)+17,'
                'IFERROR(SEARCH(". Código:",txt)+8,"")))),'
                'raw,IF(start="","",MID(txt,start,999)),'
                'clean1,IF(raw="","",SUBSTITUTE(raw,CHAR(160)," ")),'
                'clean2,SUBSTITUTE(clean1,CHAR(10)," "),'
                'clean3,SUBSTITUTE(clean2,CHAR(13)," "),'
                'chunk,IF(raw="","",TRIM(clean3)),'
                'IF(chunk="","",'
                '   IF(ISNUMBER(--LEFT(chunk,1)),'
                '      LEFT(chunk,IFERROR(SEARCH(" ",chunk)-1,LEN(chunk))),'
                '      LEFT(chunk,10)'
                '   )'
                ')'
                ')'
            )
            ws.Range("N2").Formula = formula_N
            if last_row > 2:
                ws.Range("N2").Copy(ws.Range(f"N3:N{last_row}"))

            ws.Range(f"O2:O{last_row}").FormulaLocal = "=N2"
            ws.Range(f"P2:P{last_row}").FormulaLocal = "=TEXTODEPOIS(F2;\"Correspondência:\")"
            ws.Range(f"Q2:Q{last_row}").FormulaLocal = "=TEXTOANTES(P2;\"Cidade\")"
            ws.Range(f"R2:R{last_row}").FormulaLocal = "=TEXTODEPOIS(F2;\"Documentos:\")"
            ws.Range(f"S2:S{last_row}").FormulaLocal = "=TEXTOANTES(R2;\"*\")"
            print("Fórmulas aplicadas com sucesso.")

        print("Criando: Planilha de Resumo...")
        ws_summary = wb.Worksheets.Add(After=ws)
        ws_summary.Name = "Resumo"
        ws_summary.Activate()

        print("Criando: Layout do Resumo (Título e Cabeçalho)...")
        today_date = time.strftime("%d/%m/%Y")
        
        title_range = ws_summary.Range("A1:D1")
        title_range.Merge()
        title_range.Value = f"MRV - DATA - {today_date}"
        title_range.Font.Bold = True
        title_range.HorizontalAlignment = xlCenter 
        
        ws_summary.Range("A2").Value = "Centro de Custo"
        ws_summary.Range("B2").Value = "Chamado"
        ws_summary.Range("C2").Value = "Serviço"
        ws_summary.Range("D2").Value = "Quantidade"
        summary_header = ws_summary.Range("A2:D2")
        summary_header.Font.Bold = True
        summary_header.Font.Color = 16777215 
        summary_header.Interior.Color = 12611584 
        summary_header.AutoFilter()
        ws_summary.Columns("A:D").ColumnWidth = 22

        print("Copiando dados para o Resumo (somente valores)...")
        if last_row > 1:
            dest_last_row = 3 + (last_row - 2)
            ws_summary.Range(f"A3:A{dest_last_row}").NumberFormat = "0"
            ws_summary.Range(f"A3:A{dest_last_row}").Value = ws.Range(f"O2:O{last_row}").Value
            ws_summary.Range(f"B3:B{dest_last_row}").Value = ws.Range(f"B2:B{last_row}").Value
            ws_summary.Range(f"C3:C{dest_last_row}").Value = ws.Range(f"Q2:Q{last_row}").Value
            ws_summary.Range(f"D3:D{dest_last_row}").Value = ws.Range(f"S2:S{last_row}").Value
            
            print("Limpando asteriscos, quebras de linha e ajustando altura...")
            intervalo_dados = ws_summary.Range(f"A3:D{dest_last_row}")
            intervalo_dados.Replace("~*", "")
            intervalo_dados.Replace(chr(10), "")
            intervalo_dados.Replace(chr(13), "")
            intervalo_dados.WrapText = False
            
            for row_idx in range(3, dest_last_row + 1):
                for col_idx in range(1, 5): 
                    val_celula = ws_summary.Cells(row_idx, col_idx).Value
                    if val_celula is not None and isinstance(val_celula, str):
                        ws_summary.Cells(row_idx, col_idx).Value = val_celula.strip()

            ws_summary.Rows(f"3:{dest_last_row}").AutoFit()
            ws_summary.Range(f"A3:B{dest_last_row}").HorizontalAlignment = -4152 
            print("Dados copiados e formatados com sucesso.")
        else:
            print("AVISO: Nenhum dado filtrado para copiar.")
        
        excel.Application.CutCopyMode = False
        
        print("Criando: Rodapé do Resumo...")
        last_summary_row = ws_summary.Cells(ws_summary.Rows.Count, "A").End(xlUp).Row
        footer_row = max(last_summary_row + 2, 57)
        
        footer_cliente = ws_summary.Range(f"A{footer_row}:B{footer_row + 1}")
        footer_cliente.Merge()
        footer_cliente.Value = "CLIENTE:"
        footer_cliente.Font.Bold = True
        footer_cliente.VerticalAlignment = xlTop 

        footer_agf = ws_summary.Range(f"C{footer_row}:D{footer_row + 1}")
        footer_agf.Merge()
        footer_agf.Value = "AGF:"
        footer_agf.Font.Bold = True
        footer_agf.VerticalAlignment = xlTop 
        
        print("Criando: Bordas da planilha Resumo...")
        final_used_row = footer_row + 1
        full_range = ws_summary.Range(f"A1:D{final_used_row}")
        full_range.Borders.LineStyle = xlContinuous
        
        # --- VALIDAÇÃO DOS CHAMADOS NO AGILIS ---
        print("[PROGRESSO: 50]")
        print("\n--- Iniciando validação dos chamados no Agilis ---")
        wait_validacao = WebDriverWait(driver, 1)

        try:
            driver.get("https://agilis.mrv.com.br/HomePage.do?view_type=my_view")
            last_summary_row = ws_summary.Cells(ws_summary.Rows.Count, "B").End(xlUp).Row
            total_chamados = last_summary_row - 2
            print(f"Total de chamados para validar: {total_chamados}")

            for i in range(last_summary_row, 2, -1): 
                chamado = ws_summary.Cells(i, 2).Value 

                if chamado:
                    chamado_str = str(chamado).strip()
                    if chamado_str.endswith('.0'):
                        chamado_str = chamado_str[:-2] 

                    if chamado_str: 
                        manter = validar_chamado_no_agilis(chamado_str, driver, wait_validacao)
                        if not manter:
                            ws_summary.Rows(i).Delete()
                            print(f"   Linha {i} deletada do Resumo.")
                else:
                    print(f"   Linha {i} ignorada (chamado vazio).")
                    
                # Progresso dinâmico de 50% a 95%
                if total_chamados > 0:
                    chamados_processados = last_summary_row - i + 1
                    progresso_atual = 50 + int((chamados_processados / total_chamados) * 45)
                    print(f"[PROGRESSO: {progresso_atual}]")

        finally:
            driver.quit() # AQUI O NAVEGADOR É FECHADO DEFINITIVAMENTE
            print("--- Validação dos chamados concluída ---")

        # --- 3. SALVAR O ARQUIVO FINAL ---
        # Defina o caminho (ajuste se necessário)
        caminho_final = r"C:\Users\pedro.henrsilva\OneDrive - MRV\Área de Trabalho\Downloads"
        
        # 1. GARANTIA: Se a pasta não existir, o Python cria ela na hora!
        if not os.path.exists(caminho_final):
            os.makedirs(caminho_final)
            print(f"Pasta criada: {caminho_final}")

        caminho_completo = os.path.join(caminho_final, NOME_ARQUIVO_FINAL)
        
        # 2. GARANTIA: Força o Excel a não mostrar pop-ups de "Deseja substituir?"
        excel.DisplayAlerts = False 

        try:
            # Tenta salvar o arquivo
            wb.SaveAs(caminho_completo, FileFormat=51) # 51 = .xlsx
            print(f"\n--- SUCESSO! ---")
            print(f"Arquivo editado e com resumo salvo como: {caminho_completo}")
        except Exception as erro_salvar:
            print(f"\n❌ ERRO ESPECÍFICO AO SALVAR O ARQUIVO!")
            print("Verifique se o arquivo 'Produtividade_EDITADO.xlsx' já está ABERTO no seu computador e feche-o.")
            print(f"Detalhe técnico: {erro_salvar}")

        except Exception as e:
            print(f"ERRO durante a edição no Excel: {e}")
            
        finally:
            
            # Tenta fechar o navegador caso tenha dado erro antes de chegar no final
            try:
                driver.quit()
            except:
                pass

    except Exception as e:
        print(f"ERRO durante a edição no Excel: {e}")
        try:
            driver.quit()
        except:
            pass

# ==============================================================================
# FUNÇÃO MESTRE (A que o Hub Central vai chamar)
# ==============================================================================
def executar_relatorio_completo():
    print("Etapa 1/2: Baixando relatório no Agilis...")
    
    # A função de baixar retorna o navegador aberto
    navegador_aberto = baixar_relatorio_agilis()
    
    if navegador_aberto:
        print("Etapa 2/2: Processando a planilha no Excel e validando chamados...")
        # Passamos o navegador aberto para a função do Excel usar
        processar_excel_e_validar(navegador_aberto)
        print("[PROGRESSO: 100]")
        print("Automação do Relatório dos Correios finalizada!")
    else:
        print("A automação foi interrompida porque houve um erro no download.")
