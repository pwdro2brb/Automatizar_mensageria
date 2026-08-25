import win32com.client
import re
from datetime import datetime, timedelta
import pandas as pd
import os
import traceback

# Importa as configurações do Hub Central
import config

DATA_CORTE = datetime.now() - timedelta(days=30)

def executar_cobranca_boletos():
    print("[PROGRESSO: 5]")
    print("Iniciando automação de Follow-up de Boletos...")

    # ==========================
    # CONFIGURAÇÕES
    # ==========================
    DIAS_SEM_RETORNO = 5
    MARCADOR_ROBO = "[RECOBRANÇA BOLETO]"

    USUARIOS_MRV = [
    "Maria Eduarda Soares Rocha",
    "Matheus Silva De Lemos",
    "Carolina Pagnozzi Silva",
    ]

    DESTINATARIOS_INTERNOS = {
    "Maria Eduarda Soares Rocha",
    "Matheus Silva De Lemos",
    "Carolina Pagnozzi Silva",
    "faturamentoadm",
    "reajuste",
    }       

    EQUIPE_FATURAMENTO = {
    "Maria Eduarda Soares Rocha",
    "Matheus Silva De Lemos",
    "Carolina Pagnozzi Silva",
    }
    PALAVRAS_CANCELAMENTO = [
        "desconsiderar",
        "cancelar",
        "cancelado",
        "ignorar",
        "não cobrar",
        "nao cobrar",
        "cobrança indevida",
        "cobranca indevida",
        "encerrar",
        "já resolvido",
        "ja resolvido",
        "favor desconsiderar"
    ]

    PALAVRAS_BOLETO = [
    "boleto",
    "fatura",
    "pagamento",
    "cobranca",
    "cobrança",
    "pix"
    ]

    # Define a pasta onde o log será salvo (dist/arquivos/faturamento)
    PASTA_FATURAMENTO = os.path.join(config.PASTA_ARQUIVOS, "faturamento")
    os.makedirs(PASTA_FATURAMENTO, exist_ok=True)
    ARQUIVO_LOG = os.path.join(PASTA_FATURAMENTO, "Controle_Followups.xlsx")

    # ==========================
    # FUNÇÕES INTERNAS
    # ==========================

    def obter_data_email(email):
        """
        Retorna a data do email independente da pasta.
        """
        try:
            return email.ReceivedTime
        except:
            pass

        try:
            return email.SentOn
        except:
            pass

        return None


    def eh_email_valido(item):
        """
        Filtra somente MailItems do Outlook.
        """
        try:
            return item.Class == 43  # olMail
        except:
            return False           

    def extrair_dados(corpo):

        loja = ""
        vencimento = ""

        try:

            loja_match = re.search(
                r"referente a\s+(.*?)\s+com vencimento em",
                corpo,
                re.IGNORECASE | re.DOTALL
            )

            vencimento_match = re.search(
                r"com vencimento em\s+(\d{2}\.\d{2}\.\d{4})",
                corpo,
                re.IGNORECASE
            )

            if loja_match:
                loja = " ".join(
                    loja_match.group(1).split()
                )

            if vencimento_match:
                vencimento = vencimento_match.group(1)

        except:
            pass

        return loja, vencimento


    try:
        print("[PROGRESSO: 10]")
        print("Conectando ao Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        caixa_principal = namespace.Folders.Item(1)

        try:
            pasta_faturamento = caixa_principal.Folders["faturamentoadm"]
            print("Pasta 'faturamentoadm' encontrada.")

        except Exception:
            print(
                "Pasta 'faturamentoadm' não encontrada. "
                "Utilizando Caixa de Entrada."
            )

            pasta_faturamento = namespace.GetDefaultFolder(6)  # Inbox

        itens_enviados = namespace.GetDefaultFolder(5)

        rascunhos = namespace.GetDefaultFolder(16)

        # ==========================
        # LOG
        # ==========================
        print("[PROGRESSO: 15]")
        print("Carregando arquivo de controle...")
        if os.path.exists(ARQUIVO_LOG):
            df_log = pd.read_excel(ARQUIVO_LOG, engine="openpyxl")
        else:
            df_log = pd.DataFrame(columns=[
                "Data Criação", "Loja", "Vencimento", "Dias Sem Retorno", 
                "Assunto", "Status", "ConversationID", "Destinatários"
            ])

        # ==========================
        # VERIFICA RASCUNHOS JÁ GERADOS
        # ==========================
        print("[PROGRESSO: 20]")
        print("Verificando rascunhos já gerados...")
        assuntos_ja_criados = set()
        for item in rascunhos.Items:
            try:
                assunto = str(item.Subject)
                if MARCADOR_ROBO in assunto:
                    assuntos_ja_criados.add(assunto.replace(MARCADOR_ROBO, "").strip())
            except:
                pass

        # ==========================
        # AGRUPA CONVERSAS
        # ==========================
        print("[PROGRESSO: 30]")
        print("Lendo Caixa de Entrada e Itens Enviados...")
        conversas = {}
        pastas = [pasta_faturamento, itens_enviados]

        for pasta in pastas:
            for item in pasta.Items:
                try:

                    if not eh_email_valido(item):
                        continue
                    data_email = obter_data_email(item)

                    if data_email is None:
                        continue

                    data_email = data_email.replace(tzinfo=None)

                    if data_email < DATA_CORTE:
                        continue


                    assunto = str(item.Subject or "")

                    if "BOLETO" not in assunto.upper():
                        continue

                    conv_id = str(item.ConversationID)

                    if not conv_id:
                        continue

                    if conv_id not in conversas:
                        conversas[conv_id] = []

                    conversas[conv_id].append(item)

                except Exception as erro:
                    print(f"Erro ao ler item Outlook: {erro}")

        total_conversas = len(conversas)
        print(f"Conversas encontradas: {total_conversas}")

        # ==========================
        # PROCESSAMENTO
        # ==========================
        print("[PROGRESSO: 40]")
        print("Processando conversas e gerando rascunhos...")
        criados = 0

        for i, (conv_id, emails) in enumerate(conversas.items()):
            try:
                emails_validos = []

                for e in emails:
                    try:
                        _ = e.ReceivedTime
                        emails_validos.append(e)
                    except:
                        pass

                if not emails_validos:
                    continue

                emails_validos.sort(key=lambda x: x.ReceivedTime)

                primeira = emails_validos[0]
                ultima = emails_validos[-1]

                emails_validos = []

                for email in emails:
                    data_email = obter_data_email(email)

                    if data_email is not None:
                        emails_validos.append(email)

                if not emails_validos:
                    continue

                emails_validos.sort(
                    key=lambda x: obter_data_email(x)
                )

                primeira = emails_validos[0]
                ultima = emails_validos[-1]
                assunto_original = primeira.Subject

                if assunto_original.upper().startswith("ENC:"):
                    print("Ignorado -> Email encaminhado")
                    continue

                # evita criar duas vezes
                if assunto_original in assuntos_ja_criados:
                    continue

                nome_remetente = str(
                                    ultima.SenderName or ""
                )


                ultima_eh_mrv = any(
                    pessoa.lower() in nome_remetente.lower()
                    for pessoa in USUARIOS_MRV
                )

                if not ultima_eh_mrv:
                    continue

                # ==================================================
                # REGRA NOVA
                # VERIFICA SE HOUVE CANCELAMENTO
                # ==================================================

                conversa_cancelada = False

                for email in emails_validos:

                    try:
                        texto = (
                            str(email.Subject or "") + " " +
                            str(email.Body or "")
                        ).lower()

                        if any(
                            palavra in texto
                            for palavra in PALAVRAS_CANCELAMENTO
                        ):
                            conversa_cancelada = True
                            break

                    except:
                        pass

                if conversa_cancelada:
                    print("Ignorado -> Conversa cancelada/desconsiderada")
                    continue

                # ==================================================
                # REGRA 2
                # MAIS DE 5 DIAS
                # ==================================================

                data_final = obter_data_email(ultima)

                if data_final is None:
                    continue

                data_final = data_final.replace(tzinfo=None)

                dias_sem_retorno = (
                    datetime.now() - data_final
                ).days

                if dias_sem_retorno < DIAS_SEM_RETORNO:
                    print(f"Ignorado -> Apenas {dias_sem_retorno} dias")
                    continue


                # ==================================================
                # REGRA 3
                # NÃO PODE EXISTIR BOLETO/ANEXO DE PAGAMENTO
                # ==================================================

                tem_boleto = False

                for email in emails_validos:

                    try:

                        total_anexos = email.Attachments.Count

                        for i in range(1, total_anexos + 1):

                            nome_arquivo = (
                                email.Attachments.Item(i)
                                .FileName
                                .lower()
                            )

                            if (
                                nome_arquivo.endswith(".pdf")
                                or any(
                                    palavra in nome_arquivo
                                    for palavra in PALAVRAS_BOLETO
                                )
                            ):
                                tem_boleto = True
                                break

                        if tem_boleto:
                            break

                    except:
                        pass

                if tem_boleto:
                    print(
                        "Ignorado -> Encontrado possível boleto/anexo de pagamento"
                    )
                    continue


                # ==================================================
                # REGRA 4
                # NÃO PODE SER CONVERSA APENAS ENTRE
                # MATHEUS / CAROLINA / MARIA
                # ==================================================

                destinatarios = []

                try:

                    for r in ultima.Recipients:

                        try:

                            nome = str(r.Name).strip()

                            if nome:
                                destinatarios.append(nome)

                        except:
                            pass

                except:
                    pass

                destinatarios = list(set(destinatarios))

                print("\nDESTINATÁRIOS DA ÚLTIMA MENSAGEM:")

                for nome in destinatarios:
                    print(repr(nome))
                if destinatarios:

                    somente_faturamento = all(
                        nome in EQUIPE_FATURAMENTO
                        for nome in destinatarios
                    )

                somente_interno = False

                if destinatarios:

                    somente_interno = all(
                        nome in DESTINATARIOS_INTERNOS
                        for nome in destinatarios
                    )

                if somente_interno:

                    print(
                        f"Ignorado -> Conversa interna: "
                        f"{destinatarios}"
                    )

                    continue
                # NOVA PROTEÇÃO

                if len(emails_validos) > 15:

                    print(
                    f"Atenção -> Conversa longa"
                    f"({len(emails_validos)} emails)"
                    )

                loja, vencimento = extrair_dados(primeira.Body)

                print("=" * 60)
                print("ASSUNTO:", assunto_original)
                print("LOJA:", loja)
                print("VENCIMENTO:", vencimento)

                if not loja:

                    print(
                        f"Ignorado -> Loja não encontrada | {assunto_original}"
                    )

                    continue
                    
                ja_processado = df_log["ConversationID"].astype(str).eq(str(conv_id))
                if ja_processado.any():
                    print(f"  -> Conversa já registrada: {assunto_original}")
                    continue

                print("=" * 80)
                print("FOLLOW-UP VÁLIDO")
                print("ASSUNTO:", assunto_original)
                print("DIAS:", dias_sem_retorno)
                print("DESTINATÁRIOS:", destinatarios)
                print("=" * 80)    

                resposta = ultima.ReplyAll()
                resposta.Subject = f"{MARCADOR_ROBO} {resposta.Subject}"
                resposta.HTMLBody = f"""
                <p>Olá!</p>
                <p>Gostaríamos de saber se temos algum retorno referente ao Boleto da locação <b>{loja}</b> com vencimento no dia <b>{vencimento}</b>?</p>
                <br>
                <p>Atenciosamente,<br>Equipe de faturamento.</p>
                """ + resposta.HTMLBody

                resposta.Save()
                
                novo_registro = pd.DataFrame([{
                    "Data Criação": datetime.now(),
                    "Loja": loja,
                    "Vencimento": vencimento,
                    "Dias Sem Retorno": dias_sem_retorno,
                    "Assunto": assunto_original,
                    "Status": "Rascunho Criado",
                    "Destinatários": resposta.To,
                    "ConversationID": str(conv_id)
                }])

                df_log = pd.concat([df_log, novo_registro], ignore_index=True)
                criados += 1
                print(f"  ✅ Rascunho criado | {loja} | {dias_sem_retorno} dias")

            except Exception as erro:
                print(f"\n❌ Erro na conversa: {assunto_original if 'assunto_original' in locals() else 'Sem assunto'}")
                traceback.print_exc()

            # Atualiza a barra de progresso (de 40% a 90%)
            progresso_atual = 40 + int(((i + 1) / total_conversas) * 50)
            print(f"[PROGRESSO: {progresso_atual}]")

        print("[PROGRESSO: 95]")
        print("Salvando arquivo de controle...")
        df_log.to_excel(ARQUIVO_LOG, index=False, engine="openpyxl")
        
        print("[PROGRESSO: 100]")
        print(f"\nTotal de rascunhos criados: {criados}")
        print(f"Arquivo de controle salvo em: {ARQUIVO_LOG}")

    except Exception as e:
        print(f"\n[ERRO FATAL] {e}")
        traceback.print_exc()
        raise e

if __name__ == "__main__":
    executar_cobranca_boletos()
