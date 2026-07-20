import re
import sys
from pathlib import Path
import win32com.client as win32
from datetime import datetime

import config
# ==============================================================================
# CONFIGURAÇÃO DE PASTAS DINÂMICAS
# ==============================================================================
PASTA_UBER = Path(config.PASTA_ARQUIVOS) / "uber"

# =========================
# CONFIGURAÇÕES
# =========================
# Calcula o mês passado automaticamente para o assunto
now = datetime.now()
prev_month = now.month - 1
prev_year = now.year
if prev_month == 0:
    prev_month = 12
    prev_year -= 1

meses_pt = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho", 
            7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}
nome_mes = meses_pt[prev_month]

ASSUNTO = f"Referente a utilização de uber no centro de custo no mês de {nome_mes} de {prev_year}"

CORPO = (
    f"Segue em anexo as utilizações do Uber coorporativo referente ao mês de {nome_mes} de {prev_year}, "
    "solicitados nos centros de custos sob sua responsabilidade. Caso estiver de acordo, não é necessário responder esse e-mail."
)

CC_OPCIONAL = "" 
EXTENSOES = {".xlsx", ".xlsm", ".xls"}

# =========================
# FUNÇÕES
# =========================
def achar_pasta_mes() -> Path:
    """Procura a pasta no formato YYYY,MM dentro de arquivos/uber"""
    if not PASTA_UBER.exists():
        raise RuntimeError(f"A pasta {PASTA_UBER} não existe.")
        
    pastas = [p for p in PASTA_UBER.iterdir() if p.is_dir() and re.match(r"^\d{4},\d{2}$", p.name)]
    
    if not pastas:
        raise RuntimeError(f"Nenhuma pasta de mês (ex: {prev_year},{prev_month:02d}) encontrada em {PASTA_UBER}")
        
    # Pega a mais recente baseada no nome (ano,mes)
    return max(pastas, key=lambda p: p.name)

def limpar_nome_arquivo(nome: str) -> str:
    nome = Path(nome).stem
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome

# =========================
# MAIN
# =========================
def criar_rascunhos():
    print("="*60)
    print(" UBER ETAPA 3: CRIAR RASCUNHOS DE E-MAIL ".center(60))
    print("="*60)
    
    try:
        pasta_mes = achar_pasta_mes()
    except Exception as e:
        print(f"❌ Erro: {e}")
        sys.exit(1)

    arquivos = sorted(
        [f for f in pasta_mes.iterdir() if f.is_file() and f.suffix.lower() in EXTENSOES]
    )

    if not arquivos:
        print(f"❌ Nenhum arquivo Excel encontrado dentro de: {pasta_mes}")
        sys.exit(1)

    print(f"📁 Lendo arquivos da pasta: {pasta_mes.name}")
    print("Iniciando comunicação com o Outlook...")
    
    try:
        outlook = win32.Dispatch("Outlook.Application")
    except Exception as e:
        print(f"❌ Erro ao abrir o Outlook: {e}")
        sys.exit(1)

    falhas = []

    for f in arquivos:
        destinatario = limpar_nome_arquivo(f.name)

        mail = outlook.CreateItem(0)  
        mail.To = destinatario
        mail.CC = CC_OPCIONAL
        mail.Subject = ASSUNTO

        mail.Attachments.Add(str(f))

        recip = mail.Recipients
        recip.ResolveAll()

        resolved = True
        for i in range(1, recip.Count + 1):
            if not recip.Item(i).Resolved:
                resolved = False

        if not resolved:
            falhas.append(destinatario)

        mail.BodyFormat = 2  
        mail.Display(False)  

        assinatura_html = mail.HTMLBody

        mail.HTMLBody = f"""
        <html>
        <body>
            <p>{CORPO}</p>
            <br>
            {assinatura_html}
        </body>
        </html>
        """

        mail.Save()
        mail.Close(0)  

        print(f"✅ Rascunho criado: {destinatario} | Anexo: {f.name}")

    if falhas:
        print("\n⚠️ ATENÇÃO: alguns destinatários NÃO foram resolvidos pelo Outlook (nome não encontrado):")
        for n in falhas:
            print(" -", n)
        print("\nDica: nesses casos você pode renomear o arquivo para o e-mail, ou usar um mapeamento Nome→E-mail.")
    else:
        print("\n✅ Todos os destinatários foram resolvidos com sucesso!")

    print(f"\n📁 Pasta processada: {pasta_mes.name}")
    print("Fim.")

if __name__ == "__main__":
    criar_rascunhos()
