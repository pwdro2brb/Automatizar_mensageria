# -*- coding: utf-8 -*-
"""
Versao otimizada da automacao de cobranca de boletos.
Adaptada para o Hub Central de Automações MRV.
"""

from __future__ import annotations

import logging
import re
import shutil
import sys
import time
import unicodedata
from collections import OrderedDict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client as win32

# =============================================================================
# CONFIGURACAO
# =============================================================================

ANO = datetime.now().year
STATUS_DESEJADO = "AG. BOLETO"

RAIZ = Path(r"\\Bhz-fls-app1\mrvbh\Gerência Administrativa\Pública\NUCLEO DE CONTRATOS E APOIO A GESTÃO\CONTRATOS\Contratos Comerciais")

PASTA_BASE = RAIZ / "3. BackUp Base Contratos Comerciais" / str(ANO)
PASTA_MACRO = RAIZ / "7. Macro cobrança boleto" / str(ANO)

MESES = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

LOG_FORMAT = "%(asctime)s | %(levelname)s | %(message)s"

# Constantes do Excel
XL_UP = -4162
XL_TO_LEFT = -4159
XL_CELL_TYPE_VISIBLE = 12
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
XL_OPENXML_WORKBOOK_MACRO_ENABLED = 52

@dataclass(frozen=True)
class Registro:
    loja: Any
    codigo: Any
    fornecedor: Any
    endereco: Any
    despesa: Any
    vencimento: Any

@dataclass(frozen=True)
class Consolidado:
    loja: Any
    vencimento: int
    despesas: str
    despesa_inicial: str
    endereco: Any

# =============================================================================
# LOCALIZACAO DOS ARQUIVOS
# =============================================================================

def normalizar(valor: Any) -> str:
    if valor is None:
        return ""
    texto = " ".join(str(valor).strip().split())
    texto = "".join(
        c for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )
    return texto.upper()

def mes_da_pasta(nome: str) -> int:
    match = re.match(r"\s*(\d{1,2})\s*[.\-_ ]", nome)
    return int(match.group(1)) if match else -1

def pasta_mensal_mais_recente() -> Path:
    candidatas = [
        pasta for pasta in PASTA_MACRO.iterdir()
        if pasta.is_dir() and 1 <= mes_da_pasta(pasta.name) <= 12
    ]
    if not candidatas:
        raise FileNotFoundError(f"Nenhuma pasta mensal encontrada em {PASTA_MACRO}")
    return max(candidatas, key=lambda p: mes_da_pasta(p.name))

def arquivo_mais_recente(pasta: Path, texto: str, extensoes: tuple[str, ...], recursivo=False) -> Path:
    itens = pasta.rglob("*") if recursivo else pasta.glob("*")
    candidatos = []
    alvo = normalizar(texto)

    for arquivo in itens:
        if not arquivo.is_file():
            continue
        if arquivo.name.startswith("~$"):
            continue
        if arquivo.suffix.lower() not in extensoes:
            continue
        if alvo not in normalizar(arquivo.name):
            continue
        if "ATUALIZADO" in normalizar(arquivo.stem):
            continue
        if arquivo.stat().st_size == 0:
            continue
        candidatos.append(arquivo)

    if not candidatos:
        raise FileNotFoundError(f'Arquivo contendo "{texto}" nao encontrado em {pasta}')
    return max(candidatos, key=lambda p: p.stat().st_mtime)

# =============================================================================
# EXCEL
# =============================================================================

def valor_matriz(valor):
    if valor is None:
        return []
    if isinstance(valor, tuple):
        resultado = []
        for item in valor:
            if isinstance(item, tuple):
                resultado.append(item[0])
            else:
                resultado.append(item)
        return resultado
    return [valor]

def encontrar_aba(workbook, nome: str):
    alvo = normalizar(nome)
    for planilha in workbook.Worksheets:
        if normalizar(planilha.Name) == alvo:
            return planilha
    raise KeyError(f'Aba "{nome}" nao encontrada em {workbook.Name}')

def encontrar_aba_base(workbook):
    candidatas = []
    for ws in workbook.Worksheets:
        match = re.search(r"Base Pagamento\s*(\d{4})", ws.Name, re.IGNORECASE)
        if match:
            candidatas.append((int(match.group(1)), ws))
    if not candidatas:
        raise KeyError('Aba "Base Pagamento AAAA" nao encontrada.')
    return max(candidatas, key=lambda item: item[0])[1]

def mapa_cabecalhos(ws) -> dict[str, int]:
    ultima_coluna = ws.Cells(1, ws.Columns.Count).End(XL_TO_LEFT).Column
    valores = ws.Range(ws.Cells(1, 1), ws.Cells(1, ultima_coluna)).Value
    primeira_linha = valores[0] if isinstance(valores, tuple) else (valores,)
    return {
        normalizar(valor): indice
        for indice, valor in enumerate(primeira_linha, start=1)
        if valor is not None
    }

def dia_vencimento(valor: Any) -> int:
    if isinstance(valor, (datetime, date)):
        return valor.day
    if valor is None or str(valor).strip() == "":
        raise ValueError("Vencimento vazio")
    return int(float(valor))

def ler_somente_ag_boleto(ws, mes: str) -> list[Registro]:
    mapa = mapa_cabecalhos(ws)
    campos = {
        "loja": "Loja",
        "codigo": "CÓDIGO FORNECEDOR",
        "fornecedor": "FORNECEDOR",
        "endereco": "Endereço",
        "despesa": f"Tipo de despesa {mes}",
        "vencimento": f"Vencimento {mes}",
        "status": f"Status {mes}",
    }

    colunas = {}
    for chave, cabecalho in campos.items():
        coluna = mapa.get(normalizar(cabecalho))
        if not coluna:
            raise KeyError(f'Coluna "{cabecalho}" nao encontrada na aba {ws.Name}')
        colunas[chave] = coluna

    col_status = colunas["status"]
    ultima_linha = ws.Cells(ws.Rows.Count, col_status).End(XL_UP).Row
    if ultima_linha < 2:
        return []

    logging.info("Lendo somente a coluna Status %s...", mes)
    status = valor_matriz(
        ws.Range(ws.Cells(2, col_status), ws.Cells(ultima_linha, col_status)).Value2
    )

    indices = [
        indice for indice, valor in enumerate(status)
        if normalizar(valor) == normalizar(STATUS_DESEJADO)
    ]
    logging.info("Linhas encontradas com AG. BOLETO: %d", len(indices))
    if not indices:
        return []

    dados = {}
    for chave in ("loja", "codigo", "fornecedor", "endereco", "despesa", "vencimento"):
        coluna = colunas[chave]
        dados[chave] = valor_matriz(
            ws.Range(ws.Cells(2, coluna), ws.Cells(ultima_linha, coluna)).Value2
        )

    registros = []
    for i in indices:
        registros.append(
            Registro(
                loja=dados["loja"][i],
                codigo=dados["codigo"][i],
                fornecedor=dados["fornecedor"][i],
                endereco=dados["endereco"][i],
                despesa=dados["despesa"][i],
                vencimento=dia_vencimento(dados["vencimento"][i]),
            )
        )
    return registros

def consolidar(registros: list[Registro]) -> list[Consolidado]:
    grupos = OrderedDict()
    for registro in registros:
        chave = (
            normalizar(registro.loja),
            normalizar(registro.endereco),
            registro.vencimento,
        )
        if chave not in grupos:
            grupos[chave] = {
                "loja": registro.loja,
                "vencimento": registro.vencimento,
                "endereco": registro.endereco,
                "despesas": [],
            }

        despesa = "" if registro.despesa is None else str(registro.despesa).strip()
        existentes = {normalizar(item) for item in grupos[chave]["despesas"]}
        if despesa and normalizar(despesa) not in existentes:
            grupos[chave]["despesas"].append(despesa)

    return [
        Consolidado(
            loja=g["loja"],
            vencimento=g["vencimento"],
            despesas=" + ".join(g["despesas"]),
            despesa_inicial=g["despesas"][0] if g["despesas"] else "",
            endereco=g["endereco"],
        )
        for g in grupos.values()
    ]

def ultima_linha_real(ws, colunas=(1,)) -> int:
    linhas = [ws.Cells(ws.Rows.Count, coluna).End(XL_UP).Row for coluna in colunas]
    return max(linhas + [1])

def limpar_intervalo_real(ws, col_inicio: int, col_fim: int, colunas_referencia=(1,)) -> None:
    ultima = ultima_linha_real(ws, colunas_referencia)
    if ultima >= 2:
        ws.Range(ws.Cells(2, col_inicio), ws.Cells(ultima, col_fim)).ClearContents()

def escrever_matriz(ws, linha: int, coluna: int, matriz: list[list[Any]]) -> None:
    if not matriz:
        return
    linhas = len(matriz)
    colunas = len(matriz[0])
    ws.Range(
        ws.Cells(linha, coluna),
        ws.Cells(linha + linhas - 1, coluna + colunas - 1),
    ).Value = tuple(tuple(item for item in registro) for registro in matriz)

def copiar_formatacao_linha(ws, linha_modelo: int, inicio: int, fim: int, col_fim: int) -> None:
    if fim < inicio:
        return
    origem = ws.Range(ws.Cells(linha_modelo, 1), ws.Cells(linha_modelo, col_fim))
    origem.Copy()
    destino = ws.Range(ws.Cells(inicio, 1), ws.Cells(fim, col_fim))
    destino.PasteSpecial(Paste=-4122)

def preencher_base_pagto(ws, registros, mes):
    logging.info("Atualizando Base Pagto...")
    limpar_intervalo_real(ws, 1, 6, (1, 2, 3, 4, 5, 6))
    ws.Range("A1:F1").Value = ((
        "Loja", "CÓDIGO FORNECEDOR", "FORNECEDOR", "Endereço",
        f"Tipo de despesa {mes}", f"Vencimento {mes}",
    ),)
    matriz = [[r.loja, r.codigo, r.fornecedor, r.endereco, r.despesa, r.vencimento] for r in registros]
    escrever_matriz(ws, 2, 1, matriz)
    copiar_formatacao_linha(ws, 2, 3, len(registros) + 1, 6)

def preencher_dinamica(ws, registros, mes):
    logging.info("Atualizando Dinamica BD PGTO...")
    limpar_intervalo_real(ws, 1, 4, (1, 2, 3, 4))
    ws.Range("A1:D1").Value = ((
        "Rótulos de Linha", f"Vencimento {mes}",
        f"Tipo de despesa {mes}", "Endereço",
    ),)

    matriz = []
    loja_anterior = None
    vencimento_anterior = None
    for r in registros:
        mudou_loja = normalizar(r.loja) != normalizar(loja_anterior)
        mudou_vencimento = r.vencimento != vencimento_anterior
        matriz.append([
            r.loja if mudou_loja else None,
            r.vencimento if mudou_loja or mudou_vencimento else None,
            r.despesa,
            r.endereco,
        ])
        loja_anterior = r.loja
        vencimento_anterior = r.vencimento

    escrever_matriz(ws, 2, 1, matriz)
    copiar_formatacao_linha(ws, 2, 3, len(matriz) + 1, 4)

def preencher_concatenado(ws, consolidados, mes_numero, ano):
    logging.info("Atualizando CONCATENADO...")
    limpar_intervalo_real(ws, 1, 6, (1, 2, 3, 4, 5, 6))
    ws.Range("A1:F1").Value = ((
        "Nome Loja", "Vencimento inicial", "Vencimento", "Despesas",
        "Despesas inicial", "Endereço Completo",
    ),)
    matriz = [[
        c.loja,
        c.vencimento,
        f"{c.vencimento:02d}.{mes_numero:02d}.{ano}",
        c.despesas,
        c.despesa_inicial,
        c.endereco,
    ] for c in consolidados]
    escrever_matriz(ws, 2, 1, matriz)
    copiar_formatacao_linha(ws, 2, 3, len(matriz) + 1, 6)

def preencher_macro(ws, consolidados, mes_numero, ano):
    logging.info("Atualizando Macro...")
    ultima_antiga = ultima_linha_real(ws, (7, 8, 10))
    ultima_nova = len(consolidados) + 1

    if consolidados:
        if ultima_nova >= 3:
            ws.Range("A2:M2").AutoFill(
                Destination=ws.Range(f"A2:M{ultima_nova}"),
                Type=0,
            )

        lojas = [[c.loja] for c in consolidados]
        vencimentos = [[f"{c.vencimento:02d}.{mes_numero:02d}.{ano}"] for c in consolidados]
        despesas = [[c.despesas] for c in consolidados]
        escrever_matriz(ws, 2, 7, lojas)
        escrever_matriz(ws, 2, 8, vencimentos)
        escrever_matriz(ws, 2, 10, despesas)

    if ultima_antiga > ultima_nova:
        ws.Range(
            ws.Cells(ultima_nova + 1, 1),
            ws.Cells(ultima_antiga, 11),
        ).ClearContents()

def processar() -> Path:
    inicio = time.perf_counter()
    print("[PROGRESSO: 10]")
    pasta_mes = pasta_mensal_mais_recente()
    mes_numero = mes_da_pasta(pasta_mes.name)
    mes = MESES[mes_numero]

    base = arquivo_mais_recente(
        PASTA_BASE, "Base de contratos comerciais", (".xlsx", ".xlsm")
    )
    macro = arquivo_mais_recente(
        pasta_mes, "MACRO BASE COBRANÇA DE BOLETOS", (".xlsm",), recursivo=True
    )
    saida = macro.with_name(f"{macro.stem} ATUALIZADO.xlsm")

    logging.info("Base: %s", base)
    logging.info("Macro: %s", macro)
    logging.info("Mes identificado: %s/%d", mes, ANO)

    if saida.exists():
        saida.unlink()
    shutil.copy2(macro, saida)

    pythoncom.CoInitialize()
    excel = None
    wb_base = None
    wb_macro = None

    try:
        print("[PROGRESSO: 20]")
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        try: excel.AskToUpdateLinks = False
        except Exception: pass

        try: excel.Calculation = XL_CALCULATION_MANUAL
        except Exception as e: logging.warning("Nao foi possivel alterar Calculation para Manual: %s", e)

        logging.info("Abrindo a Base de Contratos em modo somente leitura...")
        wb_base = excel.Workbooks.Open(str(base), UpdateLinks=0, ReadOnly=True, AddToMru=False)
        ws_origem = encontrar_aba_base(wb_base)
        
        print("[PROGRESSO: 40]")
        registros = ler_somente_ag_boleto(ws_origem, mes)
        wb_base.Close(SaveChanges=False)
        wb_base = None

        if not registros:
            raise RuntimeError(f'Nenhuma linha com Status {mes} = "{STATUS_DESEJADO}".')

        print("[PROGRESSO: 50]")
        consolidados = consolidar(registros)
        logging.info("Linhas consolidadas: %d", len(consolidados))

        print("[PROGRESSO: 60]")
        logging.info("Abrindo a copia da macro...")
        wb_macro = excel.Workbooks.Open(str(saida), UpdateLinks=0, ReadOnly=False, AddToMru=False)

        print("[PROGRESSO: 70]")
        preencher_base_pagto(encontrar_aba(wb_macro, "Base Pagto"), registros, mes)
        
        print("[PROGRESSO: 80]")
        preencher_dinamica(encontrar_aba(wb_macro, "Dinâmica BD PGTO"), registros, mes)
        
        print("[PROGRESSO: 90]")
        preencher_concatenado(encontrar_aba(wb_macro, "CONCATENADO"), consolidados, mes_numero, ANO)
        
        print("[PROGRESSO: 95]")
        preencher_macro(encontrar_aba(wb_macro, "Macro"), consolidados, mes_numero, ANO)

        excel.CutCopyMode = False

        try: excel.Calculation = XL_CALCULATION_AUTOMATIC
        except Exception: pass

        wb_macro.Save()
        wb_macro.Close(SaveChanges=True)
        wb_macro = None

    except Exception:
        if wb_base is not None: wb_base.Close(SaveChanges=False)
        if wb_macro is not None: wb_macro.Close(SaveChanges=False)
        if saida.exists():
            try: saida.unlink()
            except OSError: pass
        raise
    finally:
        if excel is not None:
            excel.DisplayAlerts = True
            excel.ScreenUpdating = True
            excel.EnableEvents = True
            excel.Quit()
        pythoncom.CoUninitialize()

    duracao = time.perf_counter() - inicio
    logging.info("Concluido em %.1f segundos.", duracao)
    logging.info("Arquivo criado: %s", saida)
    return saida

def executar_macro_contratos():
    """Função principal chamada pela interface central"""
    # Força o logging a sair no sys.stdout para aparecer no console da interface
    logging.basicConfig(level=logging.INFO, format=LOG_FORMAT, stream=sys.stdout, force=True)
    
    print("[PROGRESSO: 5]")
    print("=" * 60)
    print("  INICIANDO ATUALIZAÇÃO DA MACRO DE CONTRATOS")
    print("=" * 60)
    
    try:
        saida = processar()
        print("[PROGRESSO: 100]")
        print("\n✅ Processo concluído com sucesso!")
        print(f"📂 Arquivo salvo em: {saida}")
    except Exception as erro:
        logging.exception("Falha ao executar o processo: %s", erro)
        raise erro
