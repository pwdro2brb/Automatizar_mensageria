import win32com.client
import re
from datetime import datetime
import pandas as pd
import os
import traceback

# Importa as configurações do Hub Central
import config

def executar_cobranca_boletos():
    print("[PROGRESSO: 5]")
    print("Iniciando automação de Follow-up de Boletos...")

    # ==========================
    # CONFIGURAÇÕES
    # ==========================
    DIAS_SEM_RETORNO = 5
    MARCADOR_ROBO = "[FOLLOW-UP BOLETO]"
    REMETENTES_FATURAMENTO = {
        "faturamentoadm@mrv.com.br",
        "matheus.lemos.silva@mrv.com.br",
        "pagnozzi.carolina@mrv.com.br",
        "maria.eduarocha@mrv.com.br"
    }

    # Define a pasta onde o log será salvo (dist/arquivos/faturamento)
    PASTA_FATURAMENTO = os.path.join(config.PASTA_ARQUIVOS, "faturamento")
    os.makedirs(PASTA_FATURAMENTO, exist_ok=True)
    ARQUIVO_LOG = os.path.join(PASTA_FATURAMENTO, "Controle_Followups.xlsx")

    # ==========================
    # FUNÇÕES INTERNAS
    # ==========================
    def email_remetente(mail):
        try:
            return mail.SenderEmailAddress.lower().strip()
        except:
            return ""

    def extrair_dados(corpo):
        loja = ""
        vencimento = ""
        loja_match = re.search(r"referente a\s+(.*?)\s+com vencimento em", corpo, re.IGNORECASE | re.DOTALL)
        vencimento_match = re.search(r"com vencimento em\s+(\d{2}\.\d{2}\.\d{4})", corpo, re.IGNORECASE)
        
        if loja_match:
            loja = " ".join(loja_match.group(1).split())
        if vencimento_match:
            vencimento = vencimento_match.group(1)
            
        return loja, vencimento

    try:
        print("[PROGRESSO: 10]")
        print("Conectando ao Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        caixa_entrada = namespace.GetDefaultFolder(6)
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
        pastas = [caixa_entrada, itens_enviados]

        for pasta in pastas:
            for email in pasta.Items:
                try:
                    assunto = str(email.Subject)
                    if "BOLETO" not in assunto.upper():
                        continue

                    conv_id = email.ConversationID
                    if conv_id not in conversas:
                        conversas[conv_id] = []
                    conversas[conv_id].append(email)
                except:
                    pass

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
                emails.sort(key=lambda x: x.ReceivedTime)
                primeira = emails[0]
                ultima = emails[-1]
                assunto_original = primeira.Subject

                # evita criar duas vezes
                if assunto_original in assuntos_ja_criados:
                    continue

                remetente_final = email_remetente(ultima)

                # se última mensagem não é do faturamento, existe retorno do fornecedor
                if remetente_final not in REMETENTES_FATURAMENTO:
                    continue

                data_final = ultima.ReceivedTime.replace(tzinfo=None)
                dias_sem_retorno = (datetime.now() - data_final).days

                if dias_sem_retorno < DIAS_SEM_RETORNO:
                    continue

                loja, vencimento = extrair_dados(primeira.Body)
                if not loja:
                    continue
                    
                ja_processado = df_log["ConversationID"].astype(str).eq(str(conv_id))
                if ja_processado.any():
                    print(f"  -> Conversa já registrada: {assunto_original}")
                    continue
                    
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
                print(f"  ❌ Erro ao processar conversa: {erro}")

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
