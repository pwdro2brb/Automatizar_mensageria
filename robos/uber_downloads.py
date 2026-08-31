import os
import time
import traceback

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import config
from robos.uber_login import realizar_login_uber


def aguardar_download(
    pasta_downloads,
    extensao,
    criado_apos,
    timeout=180
):
    """
    Aguarda um novo arquivo ser concluido na pasta Downloads.

    Considera somente arquivos criados ou modificados depois
    do horario informado em criado_apos.
    """

    limite = time.time() + timeout
    extensao = extensao.lower()
    ultimo_log = 0

    while time.time() < limite:
        try:
            nomes = os.listdir(pasta_downloads)

        except OSError as erro:
            raise RuntimeError(
                "Nao foi possivel acessar a pasta Downloads: "
                f"{erro}"
            ) from erro

        arquivos_prontos = []
        arquivos_temporarios = []

        for nome in nomes:
            caminho = os.path.join(
                pasta_downloads,
                nome
            )

            if not os.path.isfile(caminho):
                continue

            try:
                data_modificacao = os.path.getmtime(caminho)

            except OSError:
                continue

            if data_modificacao < criado_apos:
                continue

            nome_minusculo = nome.lower()

            if nome_minusculo.endswith(
                (
                    ".crdownload",
                    ".part",
                    ".tmp"
                )
            ):
                arquivos_temporarios.append(caminho)
                continue

            if nome_minusculo.endswith(extensao):
                arquivos_prontos.append(caminho)

        if arquivos_prontos and not arquivos_temporarios:
            arquivo_mais_recente = max(
                arquivos_prontos,
                key=os.path.getmtime
            )

            try:
                tamanho_inicial = os.path.getsize(
                    arquivo_mais_recente
                )

                time.sleep(1)

                tamanho_final = os.path.getsize(
                    arquivo_mais_recente
                )

            except OSError:
                time.sleep(1)
                continue

            if (
                tamanho_inicial == tamanho_final
                and tamanho_final > 0
            ):
                return arquivo_mais_recente

        agora = time.time()

        if agora - ultimo_log >= 15:
            print(
                "Aguardando download. "
                f"Arquivos prontos: {len(arquivos_prontos)}. "
                f"Arquivos temporarios: {len(arquivos_temporarios)}."
            )

            ultimo_log = agora

        time.sleep(1)

    arquivos_diagnostico = []

    try:
        for nome in os.listdir(pasta_downloads):
            caminho = os.path.join(
                pasta_downloads,
                nome
            )

            if not os.path.isfile(caminho):
                continue

            try:
                data_modificacao = os.path.getmtime(caminho)

            except OSError:
                continue

            if data_modificacao >= criado_apos:
                arquivos_diagnostico.append(
                    (
                        data_modificacao,
                        nome
                    )
                )

    except OSError:
        pass

    arquivos_diagnostico.sort(
        reverse=True
    )

    nomes_recentes = [
        nome
        for _, nome in arquivos_diagnostico[:10]
    ]

    if nomes_recentes:
        texto_diagnostico = "\n".join(
            f"- {nome}"
            for nome in nomes_recentes
        )
    else:
        texto_diagnostico = (
            "- Nenhum arquivo novo apareceu na pasta Downloads."
        )

    raise TimeoutError(
        f"O download do arquivo {extensao} nao foi concluido "
        f"dentro de {timeout} segundos.\n\n"
        "Arquivos detectados depois do clique:\n"
        f"{texto_diagnostico}"
    )


def fechar_aviso_cookies(driver):
    """
    Fecha o aviso de cookies da Uber, caso esteja visivel.
    """

    seletores = [
        (
            By.XPATH,
            "//button[normalize-space()='Aceitar']"
        ),
        (
            By.XPATH,
            "//button[contains(normalize-space(), 'Aceitar')]"
        ),
        (
            By.XPATH,
            "//button[normalize-space()='Rejeitar']"
        ),
        (
            By.XPATH,
            "//button[contains(normalize-space(), 'Rejeitar')]"
        )
    ]

    for tipo, seletor in seletores:
        try:
            elementos = driver.find_elements(
                tipo,
                seletor
            )

            for elemento in elementos:
                if not elemento.is_displayed():
                    continue

                try:
                    elemento.click()

                except Exception:
                    driver.execute_script(
                        "arguments[0].click();",
                        elemento
                    )

                print("Aviso de cookies fechado.")
                time.sleep(1)

                return True

        except Exception:
            continue

    return False


def clicar_elemento(driver, elemento):
    """
    Tenta clicar normalmente e usa JavaScript como alternativa.
    """

    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});",
        elemento
    )

    time.sleep(1)

    try:
        elemento.click()

    except Exception:
        driver.execute_script(
            "arguments[0].click();",
            elemento
        )


def selecionar_organizacao(
    driver,
    nome_conta_uber,
    timeout=30
):
    """
    Localiza o texto da organizacao e clica no botao ancestral.
    """

    print(f"Selecionando a conta: {nome_conta_uber}")

    xpath_botao_conta = (
        "//div[contains(normalize-space(.), "
        f"'{nome_conta_uber}')]"
        "/ancestor::button[1]"
    )

    try:
        botao_conta = WebDriverWait(
            driver,
            timeout
        ).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    xpath_botao_conta
                )
            )
        )

        clicar_elemento(
            driver,
            botao_conta
        )

        print("Conta selecionada com sucesso.")

        WebDriverWait(driver, 40).until(
            lambda navegador: (
                "/dashboard/select"
                not in navegador.current_url.lower()
                or len(
                    navegador.find_elements(
                        By.CSS_SELECTOR,
                        'a[data-testid="billing-page-tab-side-nav"]'
                    )
                ) > 0
            )
        )

        return True

    except Exception as erro:
        raise RuntimeError(
            f"A organizacao '{nome_conta_uber}' "
            f"nao foi selecionada: {erro}"
        ) from erro


def acessar_aba_pagamentos(
    driver,
    tentativas=3
):
    """
    Acessa a aba de Pagamentos com novas tentativas
    caso a pagina seja atualizada durante o clique.
    """

    print("Acessando a aba de Pagamentos...")

    for tentativa in range(1, tentativas + 1):
        try:
            time.sleep(3)

            menu_pagamentos = WebDriverWait(
                driver,
                30
            ).until(
                EC.element_to_be_clickable(
                    (
                        By.CSS_SELECTOR,
                        'a[data-testid="billing-page-tab-side-nav"]'
                    )
                )
            )

            clicar_elemento(
                driver,
                menu_pagamentos
            )

            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        (
                            "//h2[contains("
                            "normalize-space(), "
                            "'Relatórios e faturas'"
                            ")]"
                        )
                    )
                )
            )

            print("Aba de Pagamentos acessada.")

            return True

        except Exception as erro:
            print(
                "A pagina foi atualizada durante o clique. "
                f"Tentativa {tentativa}/{tentativas}."
            )

            if tentativa == tentativas:
                print(
                    f"Detalhes da ultima tentativa: {erro}"
                )

            time.sleep(2)

    raise RuntimeError(
        "Nao foi possivel acessar a aba de Pagamentos "
        f"apos {tentativas} tentativas."
    )


def localizar_botao_download(driver):
    """
    Localiza o primeiro botao de download do demonstrativo.
    """

    return WebDriverWait(
        driver,
        30
    ).until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                (
                    "(//button["
                    "@data-uweb-guide-key='download-statement'"
                    "])[1]"
                )
            )
        )
    )


def localizar_opcao_download(
    driver,
    texto_opcao,
    timeout=20
):
    """
    Localiza o texto da opcao de download e retorna
    o elemento clicavel correspondente.
    """

    xpath_texto = (
        "//*[contains("
        "normalize-space(.), "
        f"'{texto_opcao}'"
        ")]"
    )

    texto_elemento = WebDriverWait(
        driver,
        timeout
    ).until(
        EC.visibility_of_element_located(
            (
                By.XPATH,
                xpath_texto
            )
        )
    )

    seletores_ancestrais = [
        "./ancestor::button[1]",
        "./ancestor::*[@role='menuitem'][1]",
        "./ancestor::*[@role='button'][1]",
        "./ancestor::*[@role='option'][1]"
    ]

    for seletor_ancestral in seletores_ancestrais:
        try:
            elemento_clicavel = texto_elemento.find_element(
                By.XPATH,
                seletor_ancestral
            )

            if (
                elemento_clicavel.is_displayed()
                and elemento_clicavel.is_enabled()
            ):
                return elemento_clicavel

        except Exception:
            continue

    return texto_elemento


def iniciar_download_csv(driver):
    """
    Abre o menu de download e seleciona Transacoes CSV.
    """

    print("Baixando Transacoes em formato CSV...")

    botao_download = localizar_botao_download(
        driver
    )

    clicar_elemento(
        driver,
        botao_download
    )

    print("Menu de download aberto.")

    opcao_csv = localizar_opcao_download(
        driver,
        "Transações (CSV)"
    )

    clicar_elemento(
        driver,
        opcao_csv
    )

    print("Opcao Transacoes CSV acionada.")


def iniciar_download_pdf(driver):
    """
    Abre o menu de download e seleciona Nota de debito.
    """

    print("Baixando Nota de debito em formato PDF...")

    botao_download = localizar_botao_download(
        driver
    )

    clicar_elemento(
        driver,
        botao_download
    )

    print("Menu de download aberto.")

    opcao_pdf = localizar_opcao_download(
        driver,
        "Nota de débito"
    )

    clicar_elemento(
        driver,
        opcao_pdf
    )

    print("Opcao Nota de debito acionada.")


def baixar_relatorios_uber(nome_conta_uber):
    """
    Acessa o Uber Business, realiza o login, seleciona
    a organizacao informada e baixa os arquivos CSV e PDF.

    Retorna uma tupla contendo:
    - caminho do arquivo CSV;
    - caminho do arquivo PDF.
    """

    print("[PROGRESSO: 12]")
    print("Abrindo o navegador para acessar a Uber...")

    caminho_downloads = getattr(
        config,
        "PASTA_DOWNLOADS",
        os.path.join(
            os.path.expanduser("~"),
            "Downloads"
        )
    )

    if not os.path.isdir(caminho_downloads):
        raise FileNotFoundError(
            "A pasta Downloads nao foi encontrada:\n"
            f"{caminho_downloads}"
        )

    chrome_options = webdriver.ChromeOptions()

    preferencias = {
        "download.default_directory": caminho_downloads,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }

    chrome_options.add_experimental_option(
        "prefs",
        preferencias
    )

    driver = None

    try:
        driver = webdriver.Chrome(
            options=chrome_options
        )

        driver.maximize_window()

        driver.get(
            "https://business.uber.com/"
        )

        print("[PROGRESSO: 15]")
        print("Iniciando autenticacao na Uber Business...")

        realizar_login_uber(
            driver
        )

        print("[PROGRESSO: 22]")
        print("Autenticacao concluida.")

        time.sleep(1)

        fechar_aviso_cookies(
            driver
        )

        selecionar_organizacao(
            driver,
            nome_conta_uber
        )

        print("[PROGRESSO: 28]")

        fechar_aviso_cookies(
            driver
        )

        acessar_aba_pagamentos(
            driver
        )

        time.sleep(3)

        print("[PROGRESSO: 35]")

        # =============================================================
        # DOWNLOAD DO CSV
        # =============================================================

        tempo_inicio_csv = time.time()

        iniciar_download_csv(
            driver
        )

        print("Download do CSV iniciado.")
        print("[PROGRESSO: 42]")

        arquivo_csv = aguardar_download(
            pasta_downloads=caminho_downloads,
            extensao=".csv",
            criado_apos=tempo_inicio_csv,
            timeout=180
        )

        print(
            "CSV baixado com sucesso: "
            f"{os.path.basename(arquivo_csv)}"
        )

        # =============================================================
        # ATUALIZAR A PAGINA
        # =============================================================

        print("Atualizando a pagina para limpar os menus...")

        driver.refresh()

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    (
                        "//h2[contains("
                        "normalize-space(), "
                        "'Relatórios e faturas'"
                        ")]"
                    )
                )
            )
        )

        time.sleep(3)

        fechar_aviso_cookies(
            driver
        )

        print("[PROGRESSO: 48]")

        # =============================================================
        # DOWNLOAD DO PDF
        # =============================================================

        tempo_inicio_pdf = time.time()

        iniciar_download_pdf(
            driver
        )

        print("Download do PDF iniciado.")

        arquivo_pdf = aguardar_download(
            pasta_downloads=caminho_downloads,
            extensao=".pdf",
            criado_apos=tempo_inicio_pdf,
            timeout=180
        )

        print(
            "PDF baixado com sucesso: "
            f"{os.path.basename(arquivo_pdf)}"
        )

        print("[PROGRESSO: 55]")
        print("Downloads concluidos com sucesso.")

        return arquivo_csv, arquivo_pdf

    except Exception as erro:
        print(
            "Erro durante a autenticacao ou download da Uber: "
            f"{erro}"
        )

        traceback.print_exc()

        raise RuntimeError(
            "Nao foi possivel concluir a autenticacao ou baixar "
            "os relatorios da Uber."
        ) from erro

    finally:
        if driver is not None:
            try:
                driver.quit()
                print("Navegador encerrado.")

            except Exception:
                pass