import pandas as pd
import os
import win32com.client as win32
import datetime
import time
import shutil
import csv 
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from robos.uber_login import realizar_login_uber
from robos.uber_downloads import baixar_relatorios_uber
import config

def executar_rateio_uber_central():
    print("[PROGRESSO: 0]")
    print("Iniciando Rateio Uber Central...")

    # =====================================================================
    # 1. CRIAÇÃO DINÂMICA DA PASTA DO MÊS NA REDE (MÊS ANTERIOR)
    # =====================================================================
    
    print("Verificando a data atual e criando a pasta do mês anterior...")

    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    hoje = datetime.datetime.now()

    # Lógica para pegar o mês anterior
    primeiro_dia_mes_atual = hoje.replace(day=1)
    ultimo_dia_mes_anterior = primeiro_dia_mes_atual - datetime.timedelta(days=1)

    nome_mes = meses[ultimo_dia_mes_anterior.month - 1]
    numero_mes = str(ultimo_dia_mes_anterior.month).zfill(2)
    ano_pasta = str(ultimo_dia_mes_anterior.year)

    nome_pasta_mes = f"{numero_mes}. {nome_mes}"

    # Novo caminho para o Uber Central
    base_faturamento = rf"\\Bhz-fls-app1\mrvbh\Gerência Administrativa\Pública\NUCLEO DE CONTRATOS E APOIO A GESTÃO\CONTRATOS\Contratos Serviços\2. UBER\1. MRV\1. Faturamento\- Uber Central\{ano_pasta}"
    pasta_dados_uber = os.path.join(base_faturamento, nome_pasta_mes)

    os.makedirs(pasta_dados_uber, exist_ok=True)
    print(f"Pasta configurada com sucesso: {pasta_dados_uber}")

    # =====================================================================
    # 2. AUTOMAÇÃO WEB (SELENIUM) - UBER CENTRAL
    # =====================================================================

    arquivo_csv_recente, arquivo_pdf_recente = baixar_relatorios_uber(
    nome_conta_uber="Uber Central"
    )

    # =====================================================================
    # 3. MOVER ARQUIVOS E CONVERTER CSV PARA EXCEL
    # =====================================================================

    print("[PROGRESSO: 20]")
    print("Movendo os arquivos baixados para a pasta do faturamento...")

    if not os.path.isfile(arquivo_csv_recente):
        raise FileNotFoundError(
            f"O arquivo CSV baixado não foi encontrado:\n"
            f"{arquivo_csv_recente}"
        )

    if not os.path.isfile(arquivo_pdf_recente):
        raise FileNotFoundError(
            f"O arquivo PDF baixado não foi encontrado:\n"
            f"{arquivo_pdf_recente}"
        )

    nome_curto_csv = f"temp_{nome_mes}.csv"

    # O Uber Central preserva o nome original do PDF.
    nome_original_pdf = os.path.basename(
        arquivo_pdf_recente
    )

    destino_csv = os.path.join(
        pasta_dados_uber,
        nome_curto_csv
    )

    destino_pdf = os.path.join(
        pasta_dados_uber,
        nome_original_pdf
    )

    # Remove versões anteriores para evitar erro no shutil.move.
    if os.path.exists(destino_csv):
        os.remove(destino_csv)

    if os.path.exists(destino_pdf):
        os.remove(destino_pdf)

    shutil.move(
        arquivo_csv_recente,
        destino_csv
    )

    shutil.move(
        arquivo_pdf_recente,
        destino_pdf
    )

    print("Arquivos movidos para a pasta do Uber Central.")
    print("Convertendo o relatório CSV para Excel...")

    print("[PROGRESSO: 62]")

    # Primeiro tenta ler como UTF-8.
    try:
        with open(
            destino_csv,
            "r",
            encoding="utf-8-sig",
            newline=""
        ) as arquivo:
            conteudo_csv = arquivo.read()

    except UnicodeDecodeError:
        with open(
            destino_csv,
            "r",
            encoding="latin1",
            newline=""
        ) as arquivo:
            conteudo_csv = arquivo.read()

    # Identifica automaticamente o delimitador principal.
    primeiras_linhas = conteudo_csv.splitlines()[:10]
    amostra_csv = "\n".join(primeiras_linhas)

    try:
        dialeto = csv.Sniffer().sniff(
            amostra_csv,
            delimiters=",;"
        )
        delimitador = dialeto.delimiter
    except csv.Error:
        delimitador = ";"

    leitor = csv.reader(
        conteudo_csv.splitlines(),
        delimiter=delimitador
    )

    dados_csv = list(leitor)

    if not dados_csv:
        raise ValueError(
            "O arquivo CSV do Uber Central está vazio."
        )

    df_bruto = pd.DataFrame(dados_csv)

    nome_arquivo_bruto = f"Utilização {nome_mes}.xlsx"

    destino_excel_bruto = os.path.join(
        pasta_dados_uber,
        nome_arquivo_bruto
    )

    if os.path.exists(destino_excel_bruto):
        os.remove(destino_excel_bruto)

    df_bruto.to_excel(
        destino_excel_bruto,
        index=False,
        header=False,
        engine="openpyxl"
    )

    # Exclui o CSV temporário somente depois de criar o Excel.
    os.remove(destino_csv)

    print(
        "Arquivo bruto convertido e salvo como: "
        f"{nome_arquivo_bruto}"
    )

    caminho_origem = destino_excel_bruto

    arquivo_saida = f"Validacao_Utilizacao_{nome_mes}.xlsx"

    caminho_destino = os.path.join(
        pasta_dados_uber,
        arquivo_saida
    )

    pasta_base_cc = (
        r"\\Bhz-fls-app1\mrvbh"
        r"\Gerência Administrativa\Pública"
        r"\NUCLEO DE CONTRATOS E APOIO A GESTÃO"
        r"\CONTRATOS\Contratos Serviços"
        r"\1. CORREIOS"
    )

    print("Procurando a base de centro de custo mais recente...")

    if not os.path.isdir(pasta_base_cc):
        raise FileNotFoundError(
            "A pasta da base de centro de custo não foi encontrada:\n"
            f"{pasta_base_cc}"
        )

    arquivos_base_cc = []

    for nome_arquivo in os.listdir(pasta_base_cc):
        nome_maiusculo = nome_arquivo.upper()

        if "BASE CENTRO DE CUSTO" not in nome_maiusculo:
            continue

        if not nome_maiusculo.endswith((".XLSX", ".XLSM", ".XLS")):
            continue

        # Ignora arquivos temporários criados pelo Excel.
        if nome_arquivo.startswith("~$"):
            continue

        caminho_arquivo = os.path.join(
            pasta_base_cc,
            nome_arquivo
        )

        if os.path.isfile(caminho_arquivo):
            arquivos_base_cc.append(caminho_arquivo)

    if not arquivos_base_cc:
        raise FileNotFoundError(
            "Nenhum arquivo Excel contendo 'BASE CENTRO DE CUSTO' "
            "foi encontrado na pasta:\n"
            f"{pasta_base_cc}"
        )

    caminho_base_cc = max(
        arquivos_base_cc,
        key=os.path.getmtime
    )

    print(
        "Base de centro de custo selecionada: "
        f"{os.path.basename(caminho_base_cc)}"
    )

    print(
        "Última modificação: "
        f"{datetime.datetime.fromtimestamp(os.path.getmtime(caminho_base_cc)):%d/%m/%Y %H:%M}"
    )

    print("[PROGRESSO: 25]")


    # =====================================================================
    # 4. INÍCIO DO PROCESSAMENTO (Lógica de Negócio - Uber)
    # =====================================================================

    print("[PROGRESSO: 37]")

    print("Iniciando o processamento do Rateio Uber...")

    df_temp = pd.read_excel(caminho_origem, header=None)

    linha_cabecalho = 5 
    for i, row in df_temp.iterrows():
        row_str = ' '.join(str(val) for val in row.values)
        if 'ID da viagem' in row_str or 'Tipo de transa' in row_str:
            linha_cabecalho = i
            break

    df = pd.read_excel(caminho_origem, header=linha_cabecalho)
    df.columns = df.columns.str.strip()

    for col in df.columns:
        if 'ID da viagem' in str(col) and col != 'ID da viagem/Uber Eats':
            df.rename(columns={col: 'ID da viagem/Uber Eats'}, inplace=True)
        if 'Tipo de transa' in str(col) and col != 'Tipo de transação':
            df.rename(columns={col: 'Tipo de transação'}, inplace=True)

    if 'ID da viagem/Uber Eats' in df.columns:
        indice_id = df.columns.get_loc('ID da viagem/Uber Eats')
        df.insert(loc=indice_id, column='Conferência', value='')

    if 'Tipo de transação' in df.columns:
        df = df[df['Tipo de transação'] != 'Payment'].reset_index(drop=True)

    duplicados_mask = df.duplicated(subset=['ID da viagem/Uber Eats'], keep=False)
    df.loc[~duplicados_mask, 'Conferência'] = 'Ok'

    for id_viagem, grupo in df[duplicados_mask].groupby('ID da viagem/Uber Eats'):
        tipos_transacao = grupo['Tipo de transação'].values
        if 'Fare' in tipos_transacao and 'Adjustment' in tipos_transacao:
            df.loc[(df['ID da viagem/Uber Eats'] == id_viagem) & (df['Tipo de transação'] == 'Fare'), 'Conferência'] = 'Ok'
            df.loc[(df['ID da viagem/Uber Eats'] == id_viagem) & (df['Tipo de transação'] == 'Adjustment'), 'Conferência'] = 'Ok, Ajustado'

    df.loc[df['Conferência'] == '', 'Conferência'] = 'Verificar'

    # =====================================================================
    # 5. "CTRL+F" NA BASE DE CENTRO DE CUSTO
    # =====================================================================

    print("[PROGRESSO: 50]")

    print("Lendo a base de Centro de Custo para validação (CTRL+F)...")
    df_base_cc = pd.read_excel(caminho_base_cc)
    df_base_cc.columns = df_base_cc.columns.str.strip()

    coluna_busca_base = 'Centro cst/Ordem/Diagr.rede'

    if coluna_busca_base in df_base_cc.columns:
        valores_base = df_base_cc[coluna_busca_base].astype(str).str.strip().str.replace(r'\.0$', '', regex=True).unique()
    else:
        print(f"ERRO: Coluna '{coluna_busca_base}' não encontrada na base de Correios.")
        valores_base = []

    def tratar_codigo_uber(codigo):
        codigo = str(codigo).strip()
        if codigo.endswith('.0'):
            codigo = codigo[:-2]
        if ' ' in codigo:
            parte1 = codigo.split(' ')[0]
            if parte1.isdigit():
                codigo = parte1
        return codigo

    # Alterado para a coluna do Uber Central
    coluna_observacao = 'Observação de despesas do programa de vouchers'

    if coluna_observacao in df.columns:
        df['Código da despesa_limpo'] = df[coluna_observacao].apply(tratar_codigo_uber)
        indice_codigo = df.columns.get_loc(coluna_observacao)
        # Cria a coluna Validação logo à direita
        df.insert(loc=indice_codigo + 1, column='Validação', value='')
        df['Validação'] = df['Código da despesa_limpo'].apply(lambda x: x if x in valores_base else 'Não localizado')
        df = df.drop(columns=['Código da despesa_limpo'])

    # =====================================================================
    # 6. CRIAR O RATEIO (TABELA DINÂMICA)
    # =====================================================================

    print("[PROGRESSO: 62]")

    print("Criando a aba de Rateio (Tabela Dinâmica)...")
    coluna_valor = 'Valor da transação em BRL (com tributos)'

    if coluna_valor in df.columns and 'Validação' in df.columns:
        df_rateio = df.groupby('Validação', as_index=False)[coluna_valor].sum()
        df_rateio = df_rateio.sort_values(by=coluna_valor, ascending=False)
    else:
        df_rateio = pd.DataFrame()

    # =====================================================================
    # 7. APLICAR CORES E SALVAR EM MÚLTIPLAS ABAS
    # =====================================================================
    print("[PROGRESSO: 75]")

    def colorir_coluna_id(coluna):
        if coluna.name == 'ID da viagem/Uber Eats':
            return ['background-color: #FF9999' if duplicado else '' for duplicado in duplicados_mask]
        return [''] * len(coluna)

    print("Aplicando cores e salvando o arquivo...")
    df_estilizado = df.style.apply(colorir_coluna_id, axis=0)

    with pd.ExcelWriter(caminho_destino, engine='openpyxl') as writer:
        df_estilizado.to_excel(writer, index=False, sheet_name='Relatório')
        if not df_rateio.empty:
            df_rateio.to_excel(writer, index=False, sheet_name='Rateio')

    # =====================================================================
    # 8. CRIAR RASCUNHO DE E-MAIL NO OUTLOOK
    # =====================================================================

    print("[PROGRESSO: 87]")

    print("Verificando se há viagens para análise e criando rascunho de e-mail...")

    df_verificar = df[df['Conferência'] == 'Verificar']

    if not df_verificar.empty:
        # Atualizado para a nova coluna no e-mail
        colunas_email = ['ID da viagem/Uber Eats', 'Observação de despesas do programa de vouchers', 'Tipo de transação', 'Valor da transação em BRL (com tributos)']
        colunas_presentes = [col for col in colunas_email if col in df_verificar.columns]
        df_tabela_email = df_verificar[colunas_presentes]
        
        tabela_html = df_tabela_email.to_html(index=False)
        tabela_html = tabela_html.replace('<table border="1" class="dataframe">', '<table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 12px;" border="1" cellpadding="5">')
        tabela_html = tabela_html.replace('<th>', '<th style="background-color: #f2f2f2; text-align: left;">')
        
        try:
            outlook = win32.Dispatch('outlook.application')
            email = outlook.CreateItem(0) 
            
            email.To = "suporte-empresa@uber.com"
            email.CC = "correiosbh@mrv.com.br; vanessa.brodrigues@mrv.com.br"
            email.Subject = f"Cobrança Indevida - Uber Central {nome_mes}"
            
            email.HTMLBody = f"""
            <p style="font-family: Arial, sans-serif; font-size: 14px;">Bom dia, prezados!</p>
            <p style="font-family: Arial, sans-serif; font-size: 14px;">Tudo bem?</p>
            <p style="font-family: Arial, sans-serif; font-size: 14px;">Gentileza verificar os ID’s de viagem abaixo pois consta em mais de uma cobrança por viagem.</p>
            <br>
            {tabela_html}
            """
            email.Display()
        except Exception as e:
            print(f"Não foi possível abrir o Outlook automaticamente. Erro: {e}")

    # Abre a pasta pública para você ver os arquivos baixados e gerados
    os.startfile(pasta_dados_uber)
    print("[PROGRESSO: 100]")
    print("Rateio Uber Central concluído com sucesso!")


if __name__ == "__main__":
    executar_rateio_uber_central()

