import time

import config

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


SELETORES_EMAIL = [
    (
        By.ID,
        "PHONE_NUMBER_or_EMAIL_ADDRESS"
    ),
    (
        By.CSS_SELECTOR,
        'input[data-testid="PHONE_NUMBER_or_EMAIL_ADDRESS"]'
    ),
    (
        By.CSS_SELECTOR,
        'input[type="email"]:not([hidden])'
    )
]

SELETORES_SENHA = [
    (
        By.ID,
        "PASSWORD"
    ),
    (
        By.NAME,
        "password"
    ),
    (
        By.CSS_SELECTOR,
        'input[autocomplete="current-password"]'
    ),
    (
        By.CSS_SELECTOR,
        'input[type="password"]'
    )
]

SELETORES_CODIGO_SMS = [
    (
        By.CSS_SELECTOR,
        'input[autocomplete="one-time-code"]'
    ),
    (
        By.CSS_SELECTOR,
        'input[inputmode="numeric"]'
    ),
    (
        By.CSS_SELECTOR,
        'input[data-testid*="OTP"]'
    ),
    (
        By.CSS_SELECTOR,
        'input[data-testid*="CODE"]'
    ),
    (
        By.CSS_SELECTOR,
        'input[name*="code"]'
    ),
    (
        By.CSS_SELECTOR,
        'input[id*="code"]'
    )
]

SELETORES_ENVIAR_SMS = [
    (
        By.ID,
        "alt-action-send-via-sms"
    ),
    (
        By.CSS_SELECTOR,
        'button[data-testid="Enviar código por SMS"]'
    ),
    (
        By.XPATH,
        "//button[contains(normalize-space(), 'Enviar código por SMS')]"
    )
]

SELETORES_USAR_SENHA = [
    (
        By.XPATH,
        "//button[contains(normalize-space(), 'Usar senha')]"
    ),
    (
        By.XPATH,
        "//button[contains(normalize-space(), 'Entrar com senha')]"
    ),
    (
        By.XPATH,
        "//button[contains(normalize-space(), 'senha')]"
    ),
    (
        By.CSS_SELECTOR,
        'button[data-testid*="password"]'
    ),
    (
        By.CSS_SELECTOR,
        'button[id*="password"]'
    )
]

SELETORES_AVANCAR = [
    (
        By.ID,
        "forward-button"
    ),
    (
        By.CSS_SELECTOR,
        'button[data-testid="forward-button"]'
    ),
    (
        By.CSS_SELECTOR,
        'button[type="submit"]'
    )
]

SELETORES_TELA_LOGADA = [
    (
        By.CSS_SELECTOR,
        'a[data-testid="billing-page-tab-side-nav"]'
    ),
    (
        By.XPATH,
        "//*[contains(normalize-space(), 'Escolha uma das organizações')]"
    ),
    (
        By.XPATH,
        "//*[contains(normalize-space(), 'Pesquisar por nome')]"
    ),
    (
        By.XPATH,
        "//*[contains(normalize-space(), 'Painéis')]"
    ),
    (
        By.XPATH,
        "//*[contains(normalize-space(), 'MRV Engenharia e Participações')]"
    ),
    (
        By.XPATH,
        "//*[contains(normalize-space(), 'Uber Central')]"
    )
]


def localizar_elemento_visivel(driver, seletores):
    """
    Procura imediatamente o primeiro elemento visível.

    Esta função não cria uma espera individual para cada seletor.
    Dessa forma, a detecção dos estados da página é rápida.
    """

    for tipo, seletor in seletores:
        try:
            elementos = driver.find_elements(tipo, seletor)

            for elemento in elementos:
                if elemento.is_displayed():
                    return elemento

        except Exception:
            continue

    return None


def clicar_elemento(driver, elemento):
    """
    Tenta clicar normalmente e utiliza JavaScript como alternativa.
    """

    try:
        elemento.click()
    except Exception:
        driver.execute_script(
            "arguments[0].click();",
            elemento
        )


def tela_uber_logada(driver):
    """
    Verifica se a autenticação terminou e a Uber já está
    na seleção de organização ou no dashboard.
    """

    url_atual = driver.current_url.lower()

    if "/dashboard" in url_atual:
        return True

    elemento_logado = localizar_elemento_visivel(
        driver,
        SELETORES_TELA_LOGADA
    )

    return elemento_logado is not None


def clicar_avancar(driver):
    """
    Localiza e clica no botão Avançar.
    """

    limite = time.time() + 15

    while time.time() < limite:
        botao = localizar_elemento_visivel(
            driver,
            SELETORES_AVANCAR
        )

        if botao is not None:
            try:
                if botao.is_enabled():
                    clicar_elemento(driver, botao)
                    return
            except Exception:
                pass

        time.sleep(0.25)

    raise TimeoutError(
        "O botão Avançar não ficou disponível."
    )


def aguardar_proximo_estado(driver, timeout=30):
    """
    Detecta o próximo estado apresentado pela autenticação da Uber.

    Possíveis retornos:
    - logado
    - enviar_sms
    - codigo_sms
    - senha
    - usar_senha
    - desconhecido
    """

    limite = time.time() + timeout

    while time.time() < limite:
        if tela_uber_logada(driver):
            return "logado", None

        campo_senha = localizar_elemento_visivel(
            driver,
            SELETORES_SENHA
        )

        if campo_senha is not None:
            return "senha", campo_senha

        campo_codigo = localizar_elemento_visivel(
            driver,
            SELETORES_CODIGO_SMS
        )

        if campo_codigo is not None:
            return "codigo_sms", campo_codigo

        botao_sms = localizar_elemento_visivel(
            driver,
            SELETORES_ENVIAR_SMS
        )

        if botao_sms is not None:
            return "enviar_sms", botao_sms

        botao_senha = localizar_elemento_visivel(
            driver,
            SELETORES_USAR_SENHA
        )

        if botao_senha is not None:
            return "usar_senha", botao_senha

        time.sleep(0.25)

    return "desconhecido", None


def aguardar_saida_da_tela_sms(driver, timeout=300):
    """
    Aguarda o preenchimento manual do SMS.

    A função termina imediatamente quando:
    - o campo de senha aparece;
    - uma opção alternativa aparece;
    - a tela de organização aparece;
    - o dashboard é aberto.
    """

    print("Digite o código recebido por SMS no navegador.")
    print("O código SMS não será armazenado pelo Hub.")

    limite = time.time() + timeout
    ultimo_aviso = 0

    while time.time() < limite:
        if tela_uber_logada(driver):
            print(
                "Código aceito. A Uber avançou diretamente "
                "para a tela inicial."
            )
            return "logado", None

        campo_senha = localizar_elemento_visivel(
            driver,
            SELETORES_SENHA
        )

        if campo_senha is not None:
            print("Código aceito. Campo de senha identificado.")
            return "senha", campo_senha

        botao_senha = localizar_elemento_visivel(
            driver,
            SELETORES_USAR_SENHA
        )

        if botao_senha is not None:
            print(
                "Código aceito. Opção de autenticação por senha "
                "identificada."
            )
            return "usar_senha", botao_senha

        agora = time.time()

        if agora - ultimo_aviso >= 30:
            restante = int(limite - agora)

            print(
                "Aguardando a conclusão da autenticação por SMS. "
                f"Tempo restante: {restante} segundos."
            )

            ultimo_aviso = agora

        time.sleep(0.25)

    raise TimeoutError(
        "O tempo para concluir a autenticação por SMS foi excedido."
    )


def preencher_senha(driver, senha_uber, campo_senha=None):
    """
    Preenche a senha e avança.
    """

    if campo_senha is None:
        estado, campo_senha = aguardar_proximo_estado(
            driver,
            timeout=20
        )

        if estado == "logado":
            return

        if estado != "senha" or campo_senha is None:
            raise RuntimeError(
                "O campo de senha da Uber não foi localizado."
            )

    print("Preenchendo a senha da Uber...")

    campo_senha.clear()
    campo_senha.send_keys(senha_uber)

    clicar_avancar(driver)


def aguardar_login_concluido(driver, timeout=90):
    """
    Aguarda a seleção de organização ou o dashboard.
    """

    limite = time.time() + timeout

    while time.time() < limite:
        if tela_uber_logada(driver):
            return True

        time.sleep(0.25)

    raise TimeoutError(
        "A Uber não apresentou a tela inicial após a autenticação."
    )


def realizar_login_uber(driver, timeout=30):
    """
    Realiza o login da Uber Business.

    Fluxos aceitos:
    1. E-mail, botão de SMS, código SMS e senha.
    2. E-mail, código SMS direto e senha.
    3. E-mail, opção para usar senha e campo de senha.
    4. E-mail e acesso direto à seleção de organização.
    5. Sessão previamente autenticada.
    """

    email_uber = getattr(
        config,
        "EMAIL_UBER",
        ""
    ).strip()

    senha_uber = getattr(
        config,
        "SENHA_UBER",
        ""
    ).strip()

    if not email_uber or not senha_uber:
        raise ValueError(
            "O e-mail e a senha da Uber não estão configurados no Hub."
        )

    if tela_uber_logada(driver):
        print("A sessão da Uber já está autenticada.")
        return

    print("Aguardando o campo de e-mail da Uber...")

    campo_email = None
    limite_email = time.time() + timeout

    while time.time() < limite_email:
        if tela_uber_logada(driver):
            print("A sessão da Uber já está autenticada.")
            return

        campo_email = localizar_elemento_visivel(
            driver,
            SELETORES_EMAIL
        )

        if campo_email is not None:
            break

        time.sleep(0.25)

    if campo_email is None:
        raise TimeoutError(
            "O campo de e-mail da Uber não foi localizado."
        )

    print("Preenchendo o e-mail da Uber...")

    campo_email.clear()
    campo_email.send_keys(email_uber)

    clicar_avancar(driver)

    print("E-mail confirmado. Verificando a próxima etapa...")

    limite_fluxo = time.time() + 360

    while time.time() < limite_fluxo:
        estado, elemento = aguardar_proximo_estado(
            driver,
            timeout=15
        )

        if estado == "logado":
            print(
                "Login concluído. A Uber abriu a seleção "
                "de organização."
            )
            return

        if estado == "enviar_sms":
            print("Selecionando Enviar código por SMS...")

            clicar_elemento(
                driver,
                elemento
            )

            estado_sms, elemento_sms = aguardar_saida_da_tela_sms(
                driver,
                timeout=300
            )

            if estado_sms == "logado":
                return

            if estado_sms == "senha":
                preencher_senha(
                    driver,
                    senha_uber,
                    elemento_sms
                )

                aguardar_login_concluido(driver)
                print("Login da Uber concluído.")
                return

            if estado_sms == "usar_senha":
                clicar_elemento(
                    driver,
                    elemento_sms
                )

                estado_senha, campo_senha = aguardar_proximo_estado(
                    driver,
                    timeout=20
                )

                if estado_senha == "logado":
                    return

                preencher_senha(
                    driver,
                    senha_uber,
                    campo_senha
                )

                aguardar_login_concluido(driver)
                print("Login da Uber concluído.")
                return

        if estado == "codigo_sms":
            print(
                "O campo do código SMS foi aberto diretamente."
            )

            estado_sms, elemento_sms = aguardar_saida_da_tela_sms(
                driver,
                timeout=300
            )

            if estado_sms == "logado":
                return

            if estado_sms == "senha":
                preencher_senha(
                    driver,
                    senha_uber,
                    elemento_sms
                )

                aguardar_login_concluido(driver)
                print("Login da Uber concluído.")
                return

            if estado_sms == "usar_senha":
                clicar_elemento(
                    driver,
                    elemento_sms
                )

                estado_senha, campo_senha = aguardar_proximo_estado(
                    driver,
                    timeout=20
                )

                if estado_senha == "logado":
                    return

                preencher_senha(
                    driver,
                    senha_uber,
                    campo_senha
                )

                aguardar_login_concluido(driver)
                print("Login da Uber concluído.")
                return

        if estado == "usar_senha":
            print("Selecionando a opção de autenticação por senha...")

            clicar_elemento(
                driver,
                elemento
            )

            estado_senha, campo_senha = aguardar_proximo_estado(
                driver,
                timeout=20
            )

            if estado_senha == "logado":
                return

            preencher_senha(
                driver,
                senha_uber,
                campo_senha
            )

            aguardar_login_concluido(driver)
            print("Login da Uber concluído.")
            return

        if estado == "senha":
            preencher_senha(
                driver,
                senha_uber,
                elemento
            )

            aguardar_login_concluido(driver)
            print("Login da Uber concluído.")
            return

        print(
            "A Uber apresentou uma etapa ainda não identificada. "
            "Aguardando alteração da página..."
        )

        time.sleep(1)

    raise TimeoutError(
        "Não foi possível concluir o fluxo de autenticação da Uber."
    )