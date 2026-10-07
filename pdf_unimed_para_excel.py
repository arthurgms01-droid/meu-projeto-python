# -*- coding: utf-8 -*-
"""
Conversor de demonstrativos Unimed de PDF para Excel (.xlsx) já formatado.

Layouts suportados (detectados automaticamente pela primeira página):
    1. "Demonstrativo Mensalidade por Pagador (Portal Web)"
    2. "Demonstrativo coparticipação por pagador Pessoa Jurídica (Portal Web)"

Uso:
    Clique duplo no arquivo (ou: python pdf_unimed_para_excel.py)
        -> abre a tela do programa (tkinter): selecione PDFs ou uma pasta e
           clique em "CONVERTER PARA EXCEL".
    Arrastar PDFs/pasta sobre o .py
        -> abre a tela já com os arquivos na lista.
    python pdf_unimed_para_excel.py --console arquivo.pdf C:/pasta/com/pdfs
        -> converte direto no console, sem janela.

O Excel é gerado na mesma pasta do PDF, com o mesmo nome e extensão .xlsx.

Dependências:
    pip install pdfplumber openpyxl
"""

import os
import re
import sys
from collections import OrderedDict
from datetime import datetime
from decimal import Decimal

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ----------------------------------------------------------------------------
# Expressões regulares
# ----------------------------------------------------------------------------
VALOR = r"-?\d{1,3}(?:\.\d{3})*,\d{2}|-?\d+,\d{2}"
DATA = r"\d{2}/\d{2}/\d{4}"

RE_BENEFICIARIO = re.compile(
    r"^(?P<titular>.+?)\s+"
    r"(?P<carteira>\d{15,20})\s+"
    r"(?P<beneficiario>.+?)\s+"
    r"(?P<tipo>Tit|Dep|Agr|Agreg|Tit\.|Dep\.)\s+"
    r"(?:(?P<matricula>\S+)\s+)??"
    r"(?P<cpf_nasc>\d{11}|" + DATA + r")\s+"
    r"(?P<idade>\d{1,3})\s+"
    r"(?P<dt_contratacao>" + DATA + r")"
    r"(?:\s+(?P<dt_rescisao>" + DATA + r"))?\s+"
    r"(?P<produto>.+?)\s+"
    r"(?P<vl_mensalidade>" + VALOR + r")\s+"
    r"(?P<vl_outros>" + VALOR + r")\s*$"
)

# Linha sem CPF / Data de nascimento (fallback)
RE_BENEFICIARIO_SEM_CPF = re.compile(
    r"^(?P<titular>.+?)\s+"
    r"(?P<carteira>\d{15,20})\s+"
    r"(?P<beneficiario>.+?)\s+"
    r"(?P<tipo>Tit|Dep|Agr|Agreg|Tit\.|Dep\.)\s+"
    r"(?P<idade>\d{1,3})\s+"
    r"(?P<dt_contratacao>" + DATA + r")"
    r"(?:\s+(?P<dt_rescisao>" + DATA + r"))?\s+"
    r"(?P<produto>.+?)\s+"
    r"(?P<vl_mensalidade>" + VALOR + r")\s+"
    r"(?P<vl_outros>" + VALOR + r")\s*$"
)

RE_FAMILIA = re.compile(r"^Fam[ií]lia\s*:\s*(?P<nome>.+?)\s*$", re.IGNORECASE)
RE_TOTAL = re.compile(
    r"^Total\s*\(\s*(?P<qtd>\d+)\s*\)\s*(?P<vl_mensalidade>" + VALOR + r")\s+(?P<vl_outros>" + VALOR + r")\s*$",
    re.IGNORECASE,
)
RE_TOTAL_GERAL = re.compile(
    r"^Total\s+(?:Geral|do\s+Pagador|Pagador)\b.*?(?P<vl_mensalidade>" + VALOR + r")\s+(?P<vl_outros>" + VALOR + r")\s*$",
    re.IGNORECASE,
)
RE_RODAPE = re.compile(r"I+m+p+r+e+s+s+o+\s*e+m+\s*:|P+[aá]+g+i+n+a+\s+\d", re.IGNORECASE)
RE_PRODUTO = re.compile(r"^(?P<codigo>\d+)\s*-\s*(?P<nome>.+)$")

LINHAS_IGNORADAS_INICIO = (
    "Demonstrativo Mensalidade",
    "Titular Carteirinha",
    "Contratação Rescisão",
    "Contratacao Rescisao",
)


# ----------------------------------------------------------------------------
# Conversões
# ----------------------------------------------------------------------------
def valor_br_para_decimal(texto):
    if texto is None:
        return None
    texto = str(texto).strip()
    if not texto:
        return None
    texto = texto.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
    try:
        return Decimal(texto)
    except Exception:
        return None


def valor_para_float(texto):
    valor = valor_br_para_decimal(texto)
    return float(valor) if valor is not None else None


def texto_para_data(texto):
    if not texto:
        return None
    texto = texto.strip()
    for formato in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(texto, formato)
        except ValueError:
            continue
    return None


def moeda_br(valor):
    texto = f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}"


def formatar_cpf(cpf):
    digitos = re.sub(r"\D", "", cpf or "")
    if len(digitos) != 11:
        return cpf or ""
    return f"{digitos[0:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:11]}"


def normalizar_espacos(texto):
    return re.sub(r"\s+", " ", texto or "").strip()


# ----------------------------------------------------------------------------
# Leitura do cabeçalho (primeira página)
# ----------------------------------------------------------------------------
def extrair_cabecalho(linhas_primeira_pagina):
    cabecalho = OrderedDict(
        [
            ("Pagador (código)", ""),
            ("Pagador (nome)", ""),
            ("Título", ""),
            ("Contrato Interno", ""),
            ("Contrato Externo", ""),
            ("Estipulante", ""),
            ("Vl. Bruto", None),
            ("Vl. Líquido", None),
            ("Dt. Referência", ""),
            ("Parcela", ""),
        ]
    )
    texto = "\n".join(linhas_primeira_pagina)

    m = re.search(r"Pagador:\s*(\d+)\s*-\s*(.*?)\s*T[ií]tulo:", texto)
    if m:
        cabecalho["Pagador (código)"] = m.group(1)
        cabecalho["Pagador (nome)"] = normalizar_espacos(m.group(2))

    m = re.search(r"Vl\.\s*L[ií]quido:\s*(" + VALOR + r"|[\d.,]+)", texto)
    if m:
        cabecalho["Vl. Líquido"] = valor_para_float(m.group(1))

    m = re.search(r"Vl\.\s*Bruto:\s*(" + VALOR + r"|[\d.,]+)", texto)
    if m:
        cabecalho["Vl. Bruto"] = valor_para_float(m.group(1))

    m = re.search(r"Dt\.\s*refer[eê]ncia:\s*(\d{2}/\d{2}/\d{2,4})", texto, re.IGNORECASE)
    if m:
        cabecalho["Dt. Referência"] = m.group(1)

    m = re.search(r"Contrato Interno:\s*(\d+)", texto)
    if m:
        cabecalho["Contrato Interno"] = m.group(1)

    m = re.search(r"Contrato Externo:\s*(\d+)", texto)
    if m:
        cabecalho["Contrato Externo"] = m.group(1)

    m = re.search(r"Parcela:\s*(\d+)", texto)
    if m:
        cabecalho["Parcela"] = m.group(1)

    m = re.search(r"T[ií]tulo:\s*(\d+)", texto)
    if m:
        cabecalho["Título"] = m.group(1)

    # O nome do pagador e o número do título costumam quebrar para a linha
    # seguinte (ex.: "Beneficios da Saude Ltda 13395844").
    for indice, linha in enumerate(linhas_primeira_pagina):
        if linha.startswith("Pagador:") and indice + 1 < len(linhas_primeira_pagina):
            proxima = linhas_primeira_pagina[indice + 1].strip()
            if ":" not in proxima:
                m = re.match(r"^(.*?)\s*(\d{5,})\s*$", proxima)
                if m:
                    if m.group(1):
                        cabecalho["Pagador (nome)"] = normalizar_espacos(
                            cabecalho["Pagador (nome)"] + " " + m.group(1)
                        )
                    if not cabecalho["Título"]:
                        cabecalho["Título"] = m.group(2)
                else:
                    cabecalho["Pagador (nome)"] = normalizar_espacos(
                        cabecalho["Pagador (nome)"] + " " + proxima
                    )
            break

    # Estipulante: texto entre "Estipulante:" e "Vl. Bruto:" + continuação
    for indice, linha in enumerate(linhas_primeira_pagina):
        if "Estipulante:" in linha:
            m = re.search(r"Estipulante:\s*(.*?)\s*(?:Vl\.\s*Bruto:.*)?$", linha)
            estipulante = m.group(1) if m else ""
            if indice + 1 < len(linhas_primeira_pagina):
                proxima = linhas_primeira_pagina[indice + 1].strip()
                if (
                    proxima
                    and ":" not in proxima
                    and not proxima.startswith(LINHAS_IGNORADAS_INICIO)
                    and not RE_FAMILIA.match(proxima)
                ):
                    estipulante = estipulante + " " + proxima
            cabecalho["Estipulante"] = normalizar_espacos(estipulante)
            break

    return cabecalho


# ----------------------------------------------------------------------------
# Leitura das linhas de beneficiários
# ----------------------------------------------------------------------------
def linha_deve_ser_ignorada(linha):
    if not linha:
        return True
    if RE_RODAPE.search(linha):
        return True
    if linha.startswith(LINHAS_IGNORADAS_INICIO):
        return True
    if linha.startswith(("Pagador:", "Contrato Interno:")):
        return True
    return False


def montar_registro(match, familia, pagina):
    dados = match.groupdict()
    cpf_nasc = dados.get("cpf_nasc") or ""
    cpf = ""
    dt_nasc = None
    if re.fullmatch(r"\d{11}", cpf_nasc):
        cpf = formatar_cpf(cpf_nasc)
    elif cpf_nasc:
        dt_nasc = texto_para_data(cpf_nasc)

    produto_completo = normalizar_espacos(dados["produto"])
    m_prod = RE_PRODUTO.match(produto_completo)
    if m_prod:
        codigo_produto = m_prod.group("codigo")
        nome_produto = m_prod.group("nome").strip()
    else:
        codigo_produto = ""
        nome_produto = produto_completo

    tipo = dados["tipo"].replace(".", "")
    tipo_desc = {"Tit": "Titular", "Dep": "Dependente", "Agr": "Agregado", "Agreg": "Agregado"}.get(tipo, tipo)

    return {
        "familia": familia,
        "titular": normalizar_espacos(dados["titular"]),
        "carteirinha": dados["carteira"],
        "beneficiario": normalizar_espacos(dados["beneficiario"]),
        "tipo": tipo_desc,
        "matricula": (dados.get("matricula") or "").strip(),
        "cpf": cpf,
        "dt_nasc": dt_nasc,
        "idade": int(dados["idade"]),
        "dt_contratacao": texto_para_data(dados["dt_contratacao"]),
        "dt_rescisao": texto_para_data(dados.get("dt_rescisao")),
        "cod_produto": codigo_produto,
        "produto": nome_produto,
        "vl_mensalidade": valor_para_float(dados["vl_mensalidade"]),
        "vl_outros": valor_para_float(dados["vl_outros"]),
        "pagina": pagina,
    }


def casar_beneficiario(linha):
    m = RE_BENEFICIARIO.match(linha)
    if m:
        return m
    return RE_BENEFICIARIO_SEM_CPF.match(linha)


def extrair_dados_pdf(caminho_pdf):
    beneficiarios = []
    totais_familia = []
    linhas_nao_reconhecidas = []
    total_geral_pdf = None
    cabecalho = None

    familia_atual = ""
    pendente = None  # (texto, pagina) de linha quebrada aguardando continuação

    with pdfplumber.open(caminho_pdf) as pdf:
        total_paginas = len(pdf.pages)
        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            texto = pagina.extract_text() or ""
            linhas = [l.strip() for l in texto.splitlines()]
            linhas = [l for l in linhas if l]

            if cabecalho is None:
                cabecalho = extrair_cabecalho(linhas)

            for linha in linhas:
                if linha_deve_ser_ignorada(linha):
                    continue

                m_fam = RE_FAMILIA.match(linha)
                if m_fam:
                    if pendente:
                        linhas_nao_reconhecidas.append(pendente)
                        pendente = None
                    familia_atual = normalizar_espacos(m_fam.group("nome"))
                    continue

                m_tot = RE_TOTAL.match(linha)
                if m_tot:
                    if pendente:
                        linhas_nao_reconhecidas.append(pendente)
                        pendente = None
                    totais_familia.append(
                        {
                            "familia": familia_atual,
                            "qtd": int(m_tot.group("qtd")),
                            "vl_mensalidade": valor_para_float(m_tot.group("vl_mensalidade")),
                            "vl_outros": valor_para_float(m_tot.group("vl_outros")),
                        }
                    )
                    continue

                m_tg = RE_TOTAL_GERAL.match(linha)
                if m_tg:
                    total_geral_pdf = {
                        "vl_mensalidade": valor_para_float(m_tg.group("vl_mensalidade")),
                        "vl_outros": valor_para_float(m_tg.group("vl_outros")),
                    }
                    continue

                if not familia_atual:
                    # Ainda no cabeçalho do documento
                    continue

                # Tenta casar a linha sozinha ou juntando com a linha pendente
                candidatos = []
                if pendente:
                    candidatos.append((pendente[0] + " " + linha, pendente[1]))
                candidatos.append((linha, numero_pagina))

                reconhecida = False
                for indice, (texto_candidato, pagina_candidato) in enumerate(candidatos):
                    m_ben = casar_beneficiario(texto_candidato)
                    if m_ben:
                        if pendente and indice == 1:
                            linhas_nao_reconhecidas.append(pendente)
                        beneficiarios.append(montar_registro(m_ben, familia_atual, pagina_candidato))
                        pendente = None
                        reconhecida = True
                        break

                if not reconhecida:
                    if pendente:
                        # Permite quebra em até 3 linhas
                        juntado = pendente[0] + " " + linha
                        if len(juntado) < 400:
                            pendente = (juntado, pendente[1])
                        else:
                            linhas_nao_reconhecidas.append(pendente)
                            pendente = (linha, numero_pagina)
                    else:
                        pendente = (linha, numero_pagina)

        if pendente:
            linhas_nao_reconhecidas.append(pendente)

    if cabecalho is None:
        cabecalho = extrair_cabecalho([])

    return {
        "cabecalho": cabecalho,
        "beneficiarios": beneficiarios,
        "totais_familia": totais_familia,
        "total_geral_pdf": total_geral_pdf,
        "nao_reconhecidas": linhas_nao_reconhecidas,
        "total_paginas": total_paginas,
    }


# ----------------------------------------------------------------------------
# Geração do Excel formatado
# ----------------------------------------------------------------------------
COR_TITULO = "1F4E78"
COR_CABECALHO = "2E75B6"
COR_ZEBRA = "DDEBF7"
COR_TOTAL = "FFF2CC"
COR_OK = "C6EFCE"
COR_ERRO = "FFC7CE"

FONTE_TITULO = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
FONTE_CABECALHO = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONTE_NEGRITO = Font(name="Calibri", size=11, bold=True)
FONTE_NORMAL = Font(name="Calibri", size=11)

BORDA_FINA = Border(
    left=Side(style="thin", color="A6A6A6"),
    right=Side(style="thin", color="A6A6A6"),
    top=Side(style="thin", color="A6A6A6"),
    bottom=Side(style="thin", color="A6A6A6"),
)

FORMATO_MOEDA = '"R$" #,##0.00;[Red]-"R$" #,##0.00'
FORMATO_DATA = "DD/MM/YYYY"
FORMATO_INTEIRO = "0"


def preencher(cor):
    return PatternFill(start_color=cor, end_color=cor, fill_type="solid")


def escrever_titulo(ws, linha, texto, ultima_coluna):
    ws.merge_cells(start_row=linha, start_column=1, end_row=linha, end_column=ultima_coluna)
    celula = ws.cell(row=linha, column=1, value=texto)
    celula.font = FONTE_TITULO
    celula.fill = preencher(COR_TITULO)
    celula.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[linha].height = 24


def escrever_cabecalho_tabela(ws, linha, titulos):
    for coluna, titulo in enumerate(titulos, start=1):
        celula = ws.cell(row=linha, column=coluna, value=titulo)
        celula.font = FONTE_CABECALHO
        celula.fill = preencher(COR_CABECALHO)
        celula.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celula.border = BORDA_FINA
    ws.row_dimensions[linha].height = 30


def ajustar_larguras(ws, larguras):
    for coluna, largura in enumerate(larguras, start=1):
        ws.column_dimensions[get_column_letter(coluna)].width = largura


def criar_aba_beneficiarios(wb, dados, nome_arquivo):
    ws = wb.active
    ws.title = "Beneficiários"

    colunas = [
        ("Família", 34, None),
        ("Titular", 38, None),
        ("Carteirinha", 20, "@"),
        ("Beneficiário", 38, None),
        ("Tipo", 12, None),
        ("Matrícula", 12, "@"),
        ("CPF", 16, "@"),
        ("Dt. Nasc.", 12, FORMATO_DATA),
        ("Idade", 7, FORMATO_INTEIRO),
        ("Dt. Contratação", 14, FORMATO_DATA),
        ("Dt. Rescisão", 13, FORMATO_DATA),
        ("Cód. Produto", 11, "@"),
        ("Produto", 30, None),
        ("Vl. Mensalidade", 16, FORMATO_MOEDA),
        ("Vl. Outros", 14, FORMATO_MOEDA),
        ("Página PDF", 9, FORMATO_INTEIRO),
    ]
    total_colunas = len(colunas)
    cab = dados["cabecalho"]

    escrever_titulo(ws, 1, "Demonstrativo Mensalidade por Pagador - Unimed", total_colunas)

    info = [
        ("Pagador:", f'{cab["Pagador (código)"]} - {cab["Pagador (nome)"]}'.strip(" -"),
         "Contrato Interno:", cab["Contrato Interno"]),
        ("Estipulante:", cab["Estipulante"], "Contrato Externo:", cab["Contrato Externo"]),
        ("Título:", cab["Título"], "Parcela:", cab["Parcela"]),
        ("Dt. Referência:", cab["Dt. Referência"], "Vl. Bruto:", cab["Vl. Bruto"]),
        ("Arquivo:", nome_arquivo, "Vl. Líquido:", cab["Vl. Líquido"]),
    ]
    linha = 2
    for rotulo1, valor1, rotulo2, valor2 in info:
        c = ws.cell(row=linha, column=1, value=rotulo1)
        c.font = FONTE_NEGRITO
        ws.merge_cells(start_row=linha, start_column=2, end_row=linha, end_column=8)
        c = ws.cell(row=linha, column=2, value=valor1)
        c.font = FONTE_NORMAL
        c.alignment = Alignment(horizontal="left")
        ws.merge_cells(start_row=linha, start_column=12, end_row=linha, end_column=13)
        c = ws.cell(row=linha, column=12, value=rotulo2)
        c.font = FONTE_NEGRITO
        c.alignment = Alignment(horizontal="right")
        ws.merge_cells(start_row=linha, start_column=14, end_row=linha, end_column=15)
        c = ws.cell(row=linha, column=14, value=valor2)
        c.font = FONTE_NORMAL
        if isinstance(valor2, float):
            c.number_format = FORMATO_MOEDA
        c.alignment = Alignment(horizontal="right")
        linha += 1

    linha_cabecalho = linha + 1
    escrever_cabecalho_tabela(ws, linha_cabecalho, [c[0] for c in colunas])

    chaves = [
        "familia", "titular", "carteirinha", "beneficiario", "tipo", "matricula", "cpf",
        "dt_nasc", "idade", "dt_contratacao", "dt_rescisao", "cod_produto", "produto",
        "vl_mensalidade", "vl_outros", "pagina",
    ]

    linha = linha_cabecalho + 1
    primeira_linha_dados = linha
    familia_anterior = None
    zebra = False
    for registro in dados["beneficiarios"]:
        if registro["familia"] != familia_anterior:
            zebra = not zebra
            familia_anterior = registro["familia"]
        for coluna, chave in enumerate(chaves, start=1):
            valor = registro[chave]
            if valor == "":
                valor = None
            celula = ws.cell(row=linha, column=coluna, value=valor)
            celula.font = FONTE_NORMAL
            celula.border = BORDA_FINA
            formato = colunas[coluna - 1][2]
            if formato:
                celula.number_format = formato
            if zebra:
                celula.fill = preencher(COR_ZEBRA)
            if chave in ("tipo", "idade", "dt_nasc", "dt_contratacao", "dt_rescisao",
                         "cod_produto", "pagina", "carteirinha", "cpf", "matricula"):
                celula.alignment = Alignment(horizontal="center")
        linha += 1
    ultima_linha_dados = linha - 1

    # Linha de totais
    linha_total = linha
    for coluna in range(1, total_colunas + 1):
        celula = ws.cell(row=linha_total, column=coluna)
        celula.fill = preencher(COR_TOTAL)
        celula.border = BORDA_FINA
        celula.font = FONTE_NEGRITO
    ws.cell(row=linha_total, column=1, value="TOTAL GERAL")
    letra_benef = get_column_letter(4)
    letra_mens = get_column_letter(14)
    letra_outros = get_column_letter(15)
    if ultima_linha_dados >= primeira_linha_dados:
        ws.cell(row=linha_total, column=4,
                value=f"=COUNTA({letra_benef}{primeira_linha_dados}:{letra_benef}{ultima_linha_dados})")
        ws.cell(row=linha_total, column=14,
                value=f"=SUM({letra_mens}{primeira_linha_dados}:{letra_mens}{ultima_linha_dados})")
        ws.cell(row=linha_total, column=15,
                value=f"=SUM({letra_outros}{primeira_linha_dados}:{letra_outros}{ultima_linha_dados})")
    else:
        ws.cell(row=linha_total, column=4, value=0)
        ws.cell(row=linha_total, column=14, value=0)
        ws.cell(row=linha_total, column=15, value=0)
    ws.cell(row=linha_total, column=3, value="Qtd. vidas:").alignment = Alignment(horizontal="right")
    ws.cell(row=linha_total, column=4).alignment = Alignment(horizontal="left")
    ws.cell(row=linha_total, column=14).number_format = FORMATO_MOEDA
    ws.cell(row=linha_total, column=15).number_format = FORMATO_MOEDA

    ajustar_larguras(ws, [c[1] for c in colunas])
    ws.freeze_panes = ws.cell(row=linha_cabecalho + 1, column=1)
    if ultima_linha_dados >= primeira_linha_dados:
        ws.auto_filter.ref = (
            f"A{linha_cabecalho}:{get_column_letter(total_colunas)}{ultima_linha_dados}"
        )
    ws.sheet_view.zoomScale = 90
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{linha_cabecalho}:{linha_cabecalho}"

    return ws, primeira_linha_dados, ultima_linha_dados, linha_total


def criar_aba_familias(wb, dados, ref_beneficiarios):
    ws = wb.create_sheet("Resumo por Família")
    _, primeira, ultima, _ = ref_beneficiarios

    titulos = [
        "Família", "Qtd. Vidas (PDF)", "Vl. Mensalidade (PDF)", "Vl. Outros (PDF)",
        "Qtd. Vidas (Calculado)", "Vl. Mensalidade (Calculado)", "Diferença", "Status",
    ]
    escrever_titulo(ws, 1, "Resumo por Família - Conferência", len(titulos))
    escrever_cabecalho_tabela(ws, 2, titulos)

    # Totais calculados a partir das linhas extraídas
    calculado = OrderedDict()
    for registro in dados["beneficiarios"]:
        item = calculado.setdefault(registro["familia"], {"qtd": 0, "vl": Decimal("0")})
        item["qtd"] += 1
        item["vl"] += Decimal(str(registro["vl_mensalidade"] or 0))

    totais_pdf = OrderedDict()
    for total in dados["totais_familia"]:
        totais_pdf[total["familia"]] = total

    familias = list(OrderedDict.fromkeys(list(calculado.keys()) + list(totais_pdf.keys())))

    linha = 3
    for familia in familias:
        pdf_tot = totais_pdf.get(familia)
        calc = calculado.get(familia, {"qtd": 0, "vl": Decimal("0")})
        ws.cell(row=linha, column=1, value=familia)
        ws.cell(row=linha, column=2, value=pdf_tot["qtd"] if pdf_tot else None)
        ws.cell(row=linha, column=3, value=pdf_tot["vl_mensalidade"] if pdf_tot else None)
        ws.cell(row=linha, column=4, value=pdf_tot["vl_outros"] if pdf_tot else None)
        nome_escapado = familia.replace('"', '""')
        ws.cell(row=linha, column=5,
                value=f'=COUNTIF(\'Beneficiários\'!$A${primeira}:$A${ultima},"{nome_escapado}")'
                if ultima >= primeira else 0)
        ws.cell(row=linha, column=6,
                value=f'=SUMIF(\'Beneficiários\'!$A${primeira}:$A${ultima},"{nome_escapado}",'
                      f'\'Beneficiários\'!$N${primeira}:$N${ultima})'
                if ultima >= primeira else 0)
        ws.cell(row=linha, column=7, value=f"=ROUND(F{linha}-C{linha},2)")

        ok = (
            pdf_tot is not None
            and pdf_tot["qtd"] == calc["qtd"]
            and abs(Decimal(str(pdf_tot["vl_mensalidade"])) - calc["vl"]) < Decimal("0.01")
        )
        status = "OK" if ok else ("SEM TOTAL NO PDF" if pdf_tot is None else "DIVERGENTE")
        ws.cell(row=linha, column=8, value=status)

        for coluna in range(1, len(titulos) + 1):
            celula = ws.cell(row=linha, column=coluna)
            celula.border = BORDA_FINA
            celula.font = FONTE_NORMAL
            if coluna in (3, 4, 6, 7):
                celula.number_format = FORMATO_MOEDA
            if coluna in (2, 5, 8):
                celula.alignment = Alignment(horizontal="center")
        ws.cell(row=linha, column=8).fill = preencher(COR_OK if ok else COR_ERRO)
        ws.cell(row=linha, column=8).font = FONTE_NEGRITO
        linha += 1

    ultima_fam = linha - 1
    for coluna in range(1, len(titulos) + 1):
        celula = ws.cell(row=linha, column=coluna)
        celula.fill = preencher(COR_TOTAL)
        celula.border = BORDA_FINA
        celula.font = FONTE_NEGRITO
    ws.cell(row=linha, column=1, value="TOTAL")
    if ultima_fam >= 3:
        for coluna in (2, 3, 4, 5, 6, 7):
            letra = get_column_letter(coluna)
            ws.cell(row=linha, column=coluna, value=f"=SUM({letra}3:{letra}{ultima_fam})")
    for coluna in (3, 4, 6, 7):
        ws.cell(row=linha, column=coluna).number_format = FORMATO_MOEDA
    for coluna in (2, 5):
        ws.cell(row=linha, column=coluna).alignment = Alignment(horizontal="center")

    ajustar_larguras(ws, [40, 14, 20, 16, 16, 22, 14, 20])
    ws.freeze_panes = "A3"
    if ultima_fam >= 3:
        ws.auto_filter.ref = f"A2:{get_column_letter(len(titulos))}{ultima_fam}"
    return ws


def criar_aba_conferencia(wb, dados, ref_beneficiarios):
    ws = wb.create_sheet("Conferência")
    _, _, _, linha_total = ref_beneficiarios
    cab = dados["cabecalho"]

    escrever_titulo(ws, 1, "Conferência Geral", 4)
    escrever_cabecalho_tabela(ws, 2, ["Item", "Valor PDF", "Valor Calculado", "Diferença"])

    soma_mens = sum(Decimal(str(r["vl_mensalidade"] or 0)) for r in dados["beneficiarios"])
    soma_outros = sum(Decimal(str(r["vl_outros"] or 0)) for r in dados["beneficiarios"])
    total_calculado = soma_mens + soma_outros

    itens = [
        ("Vl. Bruto (cabeçalho)", cab["Vl. Bruto"],
         f"='Beneficiários'!N{linha_total}+'Beneficiários'!O{linha_total}", True),
        ("Vl. Líquido (cabeçalho)", cab["Vl. Líquido"],
         f"='Beneficiários'!N{linha_total}+'Beneficiários'!O{linha_total}", True),
        ("Soma dos totais por família (PDF)",
         float(sum(Decimal(str(t["vl_mensalidade"])) for t in dados["totais_familia"])),
         f"='Beneficiários'!N{linha_total}", True),
        ("Qtd. vidas (totais por família no PDF)",
         sum(t["qtd"] for t in dados["totais_familia"]),
         f"='Beneficiários'!D{linha_total}", False),
    ]
    if dados["total_geral_pdf"]:
        itens.append(("Total geral impresso no PDF", dados["total_geral_pdf"]["vl_mensalidade"],
                      f"='Beneficiários'!N{linha_total}", True))

    linha = 3
    for rotulo, valor_pdf, formula, moeda in itens:
        ws.cell(row=linha, column=1, value=rotulo).font = FONTE_NEGRITO
        ws.cell(row=linha, column=2, value=valor_pdf)
        ws.cell(row=linha, column=3, value=formula)
        ws.cell(row=linha, column=4, value=f'=IF(B{linha}="","",ROUND(C{linha}-B{linha},2))')
        for coluna in range(1, 5):
            celula = ws.cell(row=linha, column=coluna)
            celula.border = BORDA_FINA
            if coluna > 1:
                celula.number_format = FORMATO_MOEDA if moeda else FORMATO_INTEIRO
                celula.alignment = Alignment(horizontal="right")
        linha += 1

    linha += 1
    ws.cell(row=linha, column=1, value="Páginas lidas no PDF:").font = FONTE_NEGRITO
    ws.cell(row=linha, column=2, value=dados["total_paginas"])
    linha += 1
    ws.cell(row=linha, column=1, value="Beneficiários extraídos:").font = FONTE_NEGRITO
    ws.cell(row=linha, column=2, value=len(dados["beneficiarios"]))
    linha += 1
    ws.cell(row=linha, column=1, value="Total extraído (Mens. + Outros):").font = FONTE_NEGRITO
    c = ws.cell(row=linha, column=2, value=float(total_calculado))
    c.number_format = FORMATO_MOEDA
    linha += 1
    ws.cell(row=linha, column=1, value="Linhas não reconhecidas:").font = FONTE_NEGRITO
    c = ws.cell(row=linha, column=2, value=len(dados["nao_reconhecidas"]))
    c.fill = preencher(COR_OK if not dados["nao_reconhecidas"] else COR_ERRO)
    c.font = FONTE_NEGRITO

    if dados["nao_reconhecidas"]:
        linha += 2
        escrever_cabecalho_tabela(ws, linha, ["Página", "Texto não reconhecido (verificar manualmente)", "", ""])
        ws.merge_cells(start_row=linha, start_column=2, end_row=linha, end_column=4)
        linha += 1
        for texto, pagina in dados["nao_reconhecidas"]:
            ws.cell(row=linha, column=1, value=pagina).alignment = Alignment(horizontal="center")
            ws.merge_cells(start_row=linha, start_column=2, end_row=linha, end_column=4)
            ws.cell(row=linha, column=2, value=texto).alignment = Alignment(wrap_text=True)
            linha += 1

    ajustar_larguras(ws, [42, 22, 22, 18])
    return ws


def gerar_excel(dados, caminho_saida, nome_arquivo_pdf):
    wb = Workbook()
    ref = criar_aba_beneficiarios(wb, dados, nome_arquivo_pdf)
    criar_aba_familias(wb, dados, ref)
    criar_aba_conferencia(wb, dados, ref)
    wb.save(caminho_saida)


# ============================================================================
# LAYOUT 2 - DEMONSTRATIVO DE COPARTICIPAÇÃO
# (funções separadas; o layout de mensalidade acima não é alterado)
# ============================================================================
LAYOUT_MENSALIDADE = "mensalidade"
LAYOUT_COPARTICIPACAO = "coparticipacao"

DATA_CURTA_OU_LONGA = r"\d{2}/\d{2}/(?:\d{4}|\d{2})"

RE_COPART_LANCAMENTO = re.compile(
    r"^(?P<carteira>\d{15,20})\s+"
    r"(?P<beneficiario>.+?)\s+"
    r"(?P<tipo>Tit|Dep|Agr|Agreg|Tit\.|Dep\.)\s+"
    r"(?:(?P<matricula>\S+)\s+)??"
    r"(?P<cpf_nasc>\d{11}|" + DATA_CURTA_OU_LONGA + r")\s+"
    r"(?P<dt_atendimento>" + DATA_CURTA_OU_LONGA + r")\s+"
    r"(?P<conta>\d+)\s+"
    r"(?P<procedimento>\S+)\s+"
    r"(?P<tipo_guia>.+?)\s+"
    r"(?P<qtd>\d+(?:,\d+)?)\s+"
    r"(?P<vl_unitario>" + VALOR + r")\s+"
    r"(?P<vl_total>" + VALOR + r")\s*$"
)

# Lançamento sem CPF / Data de nascimento (fallback)
RE_COPART_LANCAMENTO_SEM_CPF = re.compile(
    r"^(?P<carteira>\d{15,20})\s+"
    r"(?P<beneficiario>.+?)\s+"
    r"(?P<tipo>Tit|Dep|Agr|Agreg|Tit\.|Dep\.)\s+"
    r"(?P<dt_atendimento>" + DATA_CURTA_OU_LONGA + r")\s+"
    r"(?P<conta>\d+)\s+"
    r"(?P<procedimento>\S+)\s+"
    r"(?P<tipo_guia>.+?)\s+"
    r"(?P<qtd>\d+(?:,\d+)?)\s+"
    r"(?P<vl_unitario>" + VALOR + r")\s+"
    r"(?P<vl_total>" + VALOR + r")\s*$"
)

RE_COPART_TOTAL = re.compile(
    r"^Total\s*\(\s*(?P<qtd>\d+)\s*\)\s*(?P<vl_total>" + VALOR + r")\s*$",
    re.IGNORECASE,
)
RE_COPART_TOTAL_GERAL = re.compile(
    r"^Total\s+(?:Geral|do\s+Pagador|Pagador)\b.*?(?P<vl_total>" + VALOR + r")\s*$",
    re.IGNORECASE,
)
RE_COPART_IDENTIFICACAO = re.compile(r"co-?\s*participa[cç][aã]o|Vl\.\s*co-?\s*partic", re.IGNORECASE)


def detectar_layout(caminho_pdf):
    with pdfplumber.open(caminho_pdf) as pdf:
        if not pdf.pages:
            return LAYOUT_MENSALIDADE
        texto = pdf.pages[0].extract_text() or ""
    if RE_COPART_IDENTIFICACAO.search(texto):
        return LAYOUT_COPARTICIPACAO
    return LAYOUT_MENSALIDADE


def extrair_cabecalho_copart(linhas_primeira_pagina):
    base = extrair_cabecalho(linhas_primeira_pagina)
    texto = "\n".join(linhas_primeira_pagina)

    estipulante = re.sub(r"\s*Vl\.\s*co-?\s*partic\.?:\s*[\d.,]+\s*(?:Parcela:\s*\d+)?", "",
                         base["Estipulante"], flags=re.IGNORECASE)
    estipulante = normalizar_espacos(estipulante)

    vl_copart = None
    m = re.search(r"Vl\.\s*co-?\s*partic\.?:\s*(" + VALOR + r"|[\d.,]+)", texto, re.IGNORECASE)
    if m:
        vl_copart = valor_para_float(m.group(1))

    return OrderedDict(
        [
            ("Pagador (código)", base["Pagador (código)"]),
            ("Pagador (nome)", base["Pagador (nome)"]),
            ("Título", base["Título"]),
            ("Contrato Interno", base["Contrato Interno"]),
            ("Contrato Externo", base["Contrato Externo"]),
            ("Estipulante", estipulante),
            ("Vl. Coparticipação", vl_copart),
            ("Dt. Referência", base["Dt. Referência"]),
            ("Parcela", base["Parcela"]),
        ]
    )


def montar_lancamento_copart(match, familia, pagina):
    dados = match.groupdict()
    cpf_nasc = dados.get("cpf_nasc") or ""
    cpf = ""
    dt_nasc = None
    if re.fullmatch(r"\d{11}", cpf_nasc):
        cpf = formatar_cpf(cpf_nasc)
    elif cpf_nasc:
        dt_nasc = texto_para_data(cpf_nasc)

    tipo = dados["tipo"].replace(".", "")
    tipo_desc = {"Tit": "Titular", "Dep": "Dependente", "Agr": "Agregado", "Agreg": "Agregado"}.get(tipo, tipo)

    qtd_texto = dados["qtd"]
    qtd = valor_para_float(qtd_texto) if "," in qtd_texto else int(qtd_texto)

    return {
        "familia": familia,
        "carteirinha": dados["carteira"],
        "beneficiario": normalizar_espacos(dados["beneficiario"]),
        "tipo": tipo_desc,
        "matricula": (dados.get("matricula") or "").strip(),
        "cpf": cpf,
        "dt_nasc": dt_nasc,
        "dt_atendimento": texto_para_data(dados["dt_atendimento"]),
        "conta": dados["conta"],
        "procedimento": dados["procedimento"],
        "tipo_guia": normalizar_espacos(dados["tipo_guia"]),
        "qtd": qtd,
        "vl_unitario": valor_para_float(dados["vl_unitario"]),
        "vl_total": valor_para_float(dados["vl_total"]),
        "pagina": pagina,
    }


def casar_lancamento_copart(linha):
    m = RE_COPART_LANCAMENTO.match(linha)
    if m:
        return m
    return RE_COPART_LANCAMENTO_SEM_CPF.match(linha)


def extrair_dados_copart(caminho_pdf):
    lancamentos = []
    totais_familia = []
    linhas_nao_reconhecidas = []
    total_geral_pdf = None
    cabecalho = None

    familia_atual = ""
    pendente = None  # (texto, pagina) de linha quebrada aguardando continuação
    em_cabecalho = False

    with pdfplumber.open(caminho_pdf) as pdf:
        total_paginas = len(pdf.pages)
        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            texto = pagina.extract_text() or ""
            linhas = [l.strip() for l in texto.splitlines()]
            linhas = [l for l in linhas if l]

            if cabecalho is None:
                cabecalho = extrair_cabecalho_copart(linhas)

            for linha in linhas:
                if RE_RODAPE.search(linha):
                    continue

                # Cabeçalho do documento (repete no topo de cada página)
                if linha.startswith("Demonstrativo"):
                    em_cabecalho = True
                    continue
                if linha.startswith("Carteirinha"):
                    em_cabecalho = False
                    continue

                m_fam = RE_FAMILIA.match(linha)
                m_tot = RE_COPART_TOTAL.match(linha)
                m_tg = RE_COPART_TOTAL_GERAL.match(linha)
                m_lanc = casar_lancamento_copart(linha)
                if em_cabecalho:
                    if m_fam or m_tot or m_tg or m_lanc:
                        em_cabecalho = False
                    else:
                        continue

                if m_fam:
                    if pendente:
                        linhas_nao_reconhecidas.append(pendente)
                        pendente = None
                    familia_atual = normalizar_espacos(m_fam.group("nome"))
                    continue

                if m_tot:
                    if pendente:
                        linhas_nao_reconhecidas.append(pendente)
                        pendente = None
                    totais_familia.append(
                        {
                            "familia": familia_atual,
                            "qtd": int(m_tot.group("qtd")),
                            "vl_total": valor_para_float(m_tot.group("vl_total")),
                        }
                    )
                    continue

                if m_tg:
                    if pendente:
                        linhas_nao_reconhecidas.append(pendente)
                        pendente = None
                    total_geral_pdf = {"vl_total": valor_para_float(m_tg.group("vl_total"))}
                    continue

                if not familia_atual:
                    continue

                # Tenta casar a linha sozinha ou juntando com a linha pendente
                candidatos = []
                if pendente:
                    candidatos.append((pendente[0] + " " + linha, pendente[1]))
                candidatos.append((linha, numero_pagina))

                reconhecida = False
                for indice, (texto_candidato, pagina_candidato) in enumerate(candidatos):
                    m_cand = casar_lancamento_copart(texto_candidato)
                    if m_cand:
                        if pendente and indice == 1:
                            linhas_nao_reconhecidas.append(pendente)
                        lancamentos.append(montar_lancamento_copart(m_cand, familia_atual, pagina_candidato))
                        pendente = None
                        reconhecida = True
                        break

                if not reconhecida:
                    if pendente:
                        juntado = pendente[0] + " " + linha
                        if len(juntado) < 400:
                            pendente = (juntado, pendente[1])
                        else:
                            linhas_nao_reconhecidas.append(pendente)
                            pendente = (linha, numero_pagina)
                    else:
                        pendente = (linha, numero_pagina)

        if pendente:
            linhas_nao_reconhecidas.append(pendente)

    if cabecalho is None:
        cabecalho = extrair_cabecalho_copart([])

    return {
        "cabecalho": cabecalho,
        "lancamentos": lancamentos,
        "totais_familia": totais_familia,
        "total_geral_pdf": total_geral_pdf,
        "nao_reconhecidas": linhas_nao_reconhecidas,
        "total_paginas": total_paginas,
    }


def criar_aba_lancamentos_copart(wb, dados, nome_arquivo):
    ws = wb.active
    ws.title = "Coparticipação"

    colunas = [
        ("Família", 38, None),
        ("Carteirinha", 20, "@"),
        ("Beneficiário", 38, None),
        ("Tipo", 12, None),
        ("Matrícula", 12, "@"),
        ("CPF", 16, "@"),
        ("Dt. Nasc.", 12, FORMATO_DATA),
        ("Dt. Atendimento", 14, FORMATO_DATA),
        ("Conta", 12, "@"),
        ("Procedimento", 14, "@"),
        ("Tipo Guia", 14, None),
        ("Qtd", 7, "0.##"),
        ("Vl. Unitário", 14, FORMATO_MOEDA),
        ("Vl. Total", 15, FORMATO_MOEDA),
        ("Página PDF", 9, FORMATO_INTEIRO),
    ]
    total_colunas = len(colunas)
    cab = dados["cabecalho"]

    escrever_titulo(ws, 1, "Demonstrativo Coparticipação por Pagador - Unimed", total_colunas)

    info = [
        ("Pagador:", f'{cab["Pagador (código)"]} - {cab["Pagador (nome)"]}'.strip(" -"),
         "Contrato Interno:", cab["Contrato Interno"]),
        ("Estipulante:", cab["Estipulante"], "Contrato Externo:", cab["Contrato Externo"]),
        ("Título:", cab["Título"], "Parcela:", cab["Parcela"]),
        ("Dt. Referência:", cab["Dt. Referência"], "Vl. Coparticipação:", cab["Vl. Coparticipação"]),
        ("Arquivo:", nome_arquivo, "", None),
    ]
    linha = 2
    for rotulo1, valor1, rotulo2, valor2 in info:
        c = ws.cell(row=linha, column=1, value=rotulo1)
        c.font = FONTE_NEGRITO
        ws.merge_cells(start_row=linha, start_column=2, end_row=linha, end_column=7)
        c = ws.cell(row=linha, column=2, value=valor1)
        c.font = FONTE_NORMAL
        c.alignment = Alignment(horizontal="left")
        ws.merge_cells(start_row=linha, start_column=11, end_row=linha, end_column=12)
        c = ws.cell(row=linha, column=11, value=rotulo2 or None)
        c.font = FONTE_NEGRITO
        c.alignment = Alignment(horizontal="right")
        ws.merge_cells(start_row=linha, start_column=13, end_row=linha, end_column=14)
        c = ws.cell(row=linha, column=13, value=valor2)
        c.font = FONTE_NORMAL
        if isinstance(valor2, float):
            c.number_format = FORMATO_MOEDA
        c.alignment = Alignment(horizontal="right")
        linha += 1

    linha_cabecalho = linha + 1
    escrever_cabecalho_tabela(ws, linha_cabecalho, [c[0] for c in colunas])

    chaves = [
        "familia", "carteirinha", "beneficiario", "tipo", "matricula", "cpf", "dt_nasc",
        "dt_atendimento", "conta", "procedimento", "tipo_guia", "qtd", "vl_unitario",
        "vl_total", "pagina",
    ]
    centralizadas = ("carteirinha", "tipo", "matricula", "cpf", "dt_nasc", "dt_atendimento",
                     "conta", "procedimento", "tipo_guia", "qtd", "pagina")

    linha = linha_cabecalho + 1
    primeira_linha_dados = linha
    familia_anterior = None
    zebra = False
    for registro in dados["lancamentos"]:
        if registro["familia"] != familia_anterior:
            zebra = not zebra
            familia_anterior = registro["familia"]
        for coluna, chave in enumerate(chaves, start=1):
            valor = registro[chave]
            if valor == "":
                valor = None
            celula = ws.cell(row=linha, column=coluna, value=valor)
            celula.font = FONTE_NORMAL
            celula.border = BORDA_FINA
            formato = colunas[coluna - 1][2]
            if formato:
                celula.number_format = formato
            if zebra:
                celula.fill = preencher(COR_ZEBRA)
            if chave in centralizadas:
                celula.alignment = Alignment(horizontal="center")
        linha += 1
    ultima_linha_dados = linha - 1

    # Linha de totais
    linha_total = linha
    for coluna in range(1, total_colunas + 1):
        celula = ws.cell(row=linha_total, column=coluna)
        celula.fill = preencher(COR_TOTAL)
        celula.border = BORDA_FINA
        celula.font = FONTE_NEGRITO
    ws.cell(row=linha_total, column=1, value="TOTAL GERAL")
    ws.cell(row=linha_total, column=2, value="Lançamentos:").alignment = Alignment(horizontal="right")
    if ultima_linha_dados >= primeira_linha_dados:
        ws.cell(row=linha_total, column=3,
                value=f"=COUNTA(C{primeira_linha_dados}:C{ultima_linha_dados})")
        ws.cell(row=linha_total, column=12,
                value=f"=SUM(L{primeira_linha_dados}:L{ultima_linha_dados})")
        ws.cell(row=linha_total, column=14,
                value=f"=SUM(N{primeira_linha_dados}:N{ultima_linha_dados})")
    else:
        ws.cell(row=linha_total, column=3, value=0)
        ws.cell(row=linha_total, column=12, value=0)
        ws.cell(row=linha_total, column=14, value=0)
    ws.cell(row=linha_total, column=3).alignment = Alignment(horizontal="left")
    ws.cell(row=linha_total, column=12).alignment = Alignment(horizontal="center")
    ws.cell(row=linha_total, column=14).number_format = FORMATO_MOEDA

    ajustar_larguras(ws, [c[1] for c in colunas])
    ws.freeze_panes = ws.cell(row=linha_cabecalho + 1, column=1)
    if ultima_linha_dados >= primeira_linha_dados:
        ws.auto_filter.ref = (
            f"A{linha_cabecalho}:{get_column_letter(total_colunas)}{ultima_linha_dados}"
        )
    ws.sheet_view.zoomScale = 90
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{linha_cabecalho}:{linha_cabecalho}"

    return ws, primeira_linha_dados, ultima_linha_dados, linha_total


def criar_aba_beneficiarios_copart(wb, dados, ref_lancamentos):
    ws = wb.create_sheet("Resumo por Beneficiário")
    _, primeira, ultima, _ = ref_lancamentos

    titulos = ["Família", "Carteirinha", "Beneficiário", "Tipo", "CPF", "Qtd. Lançamentos", "Vl. Total"]
    escrever_titulo(ws, 1, "Coparticipação por Beneficiário", len(titulos))
    escrever_cabecalho_tabela(ws, 2, titulos)

    beneficiarios = OrderedDict()
    for registro in dados["lancamentos"]:
        chave = registro["carteirinha"]
        if chave not in beneficiarios:
            beneficiarios[chave] = registro

    linha = 3
    for carteirinha, registro in beneficiarios.items():
        ws.cell(row=linha, column=1, value=registro["familia"])
        ws.cell(row=linha, column=2, value=carteirinha).number_format = "@"
        ws.cell(row=linha, column=3, value=registro["beneficiario"])
        ws.cell(row=linha, column=4, value=registro["tipo"])
        ws.cell(row=linha, column=5, value=registro["cpf"] or None)
        if ultima >= primeira:
            ws.cell(row=linha, column=6,
                    value=f"=COUNTIF('Coparticipação'!$B${primeira}:$B${ultima},B{linha})")
            ws.cell(row=linha, column=7,
                    value=f"=SUMIF('Coparticipação'!$B${primeira}:$B${ultima},B{linha},"
                          f"'Coparticipação'!$N${primeira}:$N${ultima})")
        for coluna in range(1, len(titulos) + 1):
            celula = ws.cell(row=linha, column=coluna)
            celula.border = BORDA_FINA
            celula.font = FONTE_NORMAL
            if coluna in (2, 4, 5, 6):
                celula.alignment = Alignment(horizontal="center")
        ws.cell(row=linha, column=7).number_format = FORMATO_MOEDA
        linha += 1

    ultima_ben = linha - 1
    for coluna in range(1, len(titulos) + 1):
        celula = ws.cell(row=linha, column=coluna)
        celula.fill = preencher(COR_TOTAL)
        celula.border = BORDA_FINA
        celula.font = FONTE_NEGRITO
    ws.cell(row=linha, column=1, value="TOTAL")
    if ultima_ben >= 3:
        ws.cell(row=linha, column=6, value=f"=SUM(F3:F{ultima_ben})")
        ws.cell(row=linha, column=7, value=f"=SUM(G3:G{ultima_ben})")
    ws.cell(row=linha, column=6).alignment = Alignment(horizontal="center")
    ws.cell(row=linha, column=7).number_format = FORMATO_MOEDA

    ajustar_larguras(ws, [40, 20, 40, 12, 16, 16, 16])
    ws.freeze_panes = "A3"
    if ultima_ben >= 3:
        ws.auto_filter.ref = f"A2:{get_column_letter(len(titulos))}{ultima_ben}"
    return ws


def criar_aba_familias_copart(wb, dados, ref_lancamentos):
    ws = wb.create_sheet("Resumo por Família")
    _, primeira, ultima, _ = ref_lancamentos

    titulos = [
        "Família", "Qtd. Lançamentos (PDF)", "Vl. Total (PDF)",
        "Qtd. Lançamentos (Calculado)", "Vl. Total (Calculado)", "Diferença", "Status",
    ]
    escrever_titulo(ws, 1, "Resumo por Família - Conferência", len(titulos))
    escrever_cabecalho_tabela(ws, 2, titulos)

    calculado = OrderedDict()
    for registro in dados["lancamentos"]:
        item = calculado.setdefault(registro["familia"], {"qtd": 0, "vl": Decimal("0")})
        item["qtd"] += 1
        item["vl"] += Decimal(str(registro["vl_total"] or 0))

    totais_pdf = OrderedDict()
    for total in dados["totais_familia"]:
        totais_pdf[total["familia"]] = total

    familias = list(OrderedDict.fromkeys(list(calculado.keys()) + list(totais_pdf.keys())))

    linha = 3
    for familia in familias:
        pdf_tot = totais_pdf.get(familia)
        calc = calculado.get(familia, {"qtd": 0, "vl": Decimal("0")})
        ws.cell(row=linha, column=1, value=familia)
        ws.cell(row=linha, column=2, value=pdf_tot["qtd"] if pdf_tot else None)
        ws.cell(row=linha, column=3, value=pdf_tot["vl_total"] if pdf_tot else None)
        nome_escapado = familia.replace('"', '""')
        ws.cell(row=linha, column=4,
                value=f'=COUNTIF(\'Coparticipação\'!$A${primeira}:$A${ultima},"{nome_escapado}")'
                if ultima >= primeira else 0)
        ws.cell(row=linha, column=5,
                value=f'=SUMIF(\'Coparticipação\'!$A${primeira}:$A${ultima},"{nome_escapado}",'
                      f'\'Coparticipação\'!$N${primeira}:$N${ultima})'
                if ultima >= primeira else 0)
        ws.cell(row=linha, column=6, value=f'=IF(C{linha}="","",ROUND(E{linha}-C{linha},2))')

        ok = (
            pdf_tot is not None
            and pdf_tot["qtd"] == calc["qtd"]
            and abs(Decimal(str(pdf_tot["vl_total"])) - calc["vl"]) < Decimal("0.01")
        )
        status = "OK" if ok else ("SEM TOTAL NO PDF" if pdf_tot is None else "DIVERGENTE")
        ws.cell(row=linha, column=7, value=status)

        for coluna in range(1, len(titulos) + 1):
            celula = ws.cell(row=linha, column=coluna)
            celula.border = BORDA_FINA
            celula.font = FONTE_NORMAL
            if coluna in (3, 5, 6):
                celula.number_format = FORMATO_MOEDA
            if coluna in (2, 4, 7):
                celula.alignment = Alignment(horizontal="center")
        ws.cell(row=linha, column=7).fill = preencher(COR_OK if ok else COR_ERRO)
        ws.cell(row=linha, column=7).font = FONTE_NEGRITO
        linha += 1

    ultima_fam = linha - 1
    for coluna in range(1, len(titulos) + 1):
        celula = ws.cell(row=linha, column=coluna)
        celula.fill = preencher(COR_TOTAL)
        celula.border = BORDA_FINA
        celula.font = FONTE_NEGRITO
    ws.cell(row=linha, column=1, value="TOTAL")
    if ultima_fam >= 3:
        for coluna in (2, 3, 4, 5, 6):
            letra = get_column_letter(coluna)
            ws.cell(row=linha, column=coluna, value=f"=SUM({letra}3:{letra}{ultima_fam})")
    for coluna in (3, 5, 6):
        ws.cell(row=linha, column=coluna).number_format = FORMATO_MOEDA
    for coluna in (2, 4):
        ws.cell(row=linha, column=coluna).alignment = Alignment(horizontal="center")

    ajustar_larguras(ws, [44, 16, 18, 18, 20, 14, 20])
    ws.freeze_panes = "A3"
    if ultima_fam >= 3:
        ws.auto_filter.ref = f"A2:{get_column_letter(len(titulos))}{ultima_fam}"
    return ws


def criar_aba_conferencia_copart(wb, dados, ref_lancamentos):
    ws = wb.create_sheet("Conferência")
    _, _, _, linha_total = ref_lancamentos
    cab = dados["cabecalho"]

    escrever_titulo(ws, 1, "Conferência Geral - Coparticipação", 4)
    escrever_cabecalho_tabela(ws, 2, ["Item", "Valor PDF", "Valor Calculado", "Diferença"])

    total_calculado = sum(Decimal(str(r["vl_total"] or 0)) for r in dados["lancamentos"])

    itens = [
        ("Vl. Coparticipação (cabeçalho)", cab["Vl. Coparticipação"],
         f"='Coparticipação'!N{linha_total}", True),
        ("Soma dos totais por família (PDF)",
         float(sum(Decimal(str(t["vl_total"])) for t in dados["totais_familia"])),
         f"='Coparticipação'!N{linha_total}", True),
        ("Qtd. lançamentos (totais por família no PDF)",
         sum(t["qtd"] for t in dados["totais_familia"]),
         f"='Coparticipação'!C{linha_total}", False),
    ]
    if dados["total_geral_pdf"]:
        itens.append(("Total geral impresso no PDF", dados["total_geral_pdf"]["vl_total"],
                      f"='Coparticipação'!N{linha_total}", True))

    linha = 3
    for rotulo, valor_pdf, formula, moeda in itens:
        ws.cell(row=linha, column=1, value=rotulo).font = FONTE_NEGRITO
        ws.cell(row=linha, column=2, value=valor_pdf)
        ws.cell(row=linha, column=3, value=formula)
        ws.cell(row=linha, column=4, value=f'=IF(B{linha}="","",ROUND(C{linha}-B{linha},2))')
        for coluna in range(1, 5):
            celula = ws.cell(row=linha, column=coluna)
            celula.border = BORDA_FINA
            if coluna > 1:
                celula.number_format = FORMATO_MOEDA if moeda else FORMATO_INTEIRO
                celula.alignment = Alignment(horizontal="right")
        linha += 1

    linha += 1
    ws.cell(row=linha, column=1, value="Páginas lidas no PDF:").font = FONTE_NEGRITO
    ws.cell(row=linha, column=2, value=dados["total_paginas"])
    linha += 1
    ws.cell(row=linha, column=1, value="Lançamentos extraídos:").font = FONTE_NEGRITO
    ws.cell(row=linha, column=2, value=len(dados["lancamentos"]))
    linha += 1
    ws.cell(row=linha, column=1, value="Beneficiários distintos:").font = FONTE_NEGRITO
    ws.cell(row=linha, column=2, value=len({r["carteirinha"] for r in dados["lancamentos"]}))
    linha += 1
    ws.cell(row=linha, column=1, value="Total extraído:").font = FONTE_NEGRITO
    c = ws.cell(row=linha, column=2, value=float(total_calculado))
    c.number_format = FORMATO_MOEDA
    linha += 1
    ws.cell(row=linha, column=1, value="Linhas não reconhecidas:").font = FONTE_NEGRITO
    c = ws.cell(row=linha, column=2, value=len(dados["nao_reconhecidas"]))
    c.fill = preencher(COR_OK if not dados["nao_reconhecidas"] else COR_ERRO)
    c.font = FONTE_NEGRITO

    if dados["nao_reconhecidas"]:
        linha += 2
        escrever_cabecalho_tabela(ws, linha, ["Página", "Texto não reconhecido (verificar manualmente)", "", ""])
        ws.merge_cells(start_row=linha, start_column=2, end_row=linha, end_column=4)
        linha += 1
        for texto, pagina in dados["nao_reconhecidas"]:
            ws.cell(row=linha, column=1, value=pagina).alignment = Alignment(horizontal="center")
            ws.merge_cells(start_row=linha, start_column=2, end_row=linha, end_column=4)
            ws.cell(row=linha, column=2, value=texto).alignment = Alignment(wrap_text=True)
            linha += 1

    ajustar_larguras(ws, [46, 22, 22, 18])
    return ws


def gerar_excel_copart(dados, caminho_saida, nome_arquivo_pdf):
    wb = Workbook()
    ref = criar_aba_lancamentos_copart(wb, dados, nome_arquivo_pdf)
    criar_aba_beneficiarios_copart(wb, dados, ref)
    criar_aba_familias_copart(wb, dados, ref)
    criar_aba_conferencia_copart(wb, dados, ref)
    wb.save(caminho_saida)


def converter_pdf_copart(caminho_pdf, log=print):
    log("=" * 60)
    log(f"LENDO O ARQUIVO (COPARTICIPAÇÃO): {caminho_pdf}")
    dados = extrair_dados_copart(caminho_pdf)

    caminho_saida = os.path.splitext(caminho_pdf)[0] + ".xlsx"
    try:
        gerar_excel_copart(dados, caminho_saida, os.path.basename(caminho_pdf))
    except PermissionError:
        base = os.path.splitext(caminho_pdf)[0]
        caminho_saida = f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        log("  Aviso: o Excel original está aberto. Salvando com outro nome.")
        gerar_excel_copart(dados, caminho_saida, os.path.basename(caminho_pdf))

    soma = sum(Decimal(str(r["vl_total"] or 0)) for r in dados["lancamentos"])
    vl_copart = dados["cabecalho"]["Vl. Coparticipação"]
    confere = None
    log(f"  Páginas lidas.............: {dados['total_paginas']}")
    log(f"  Lançamentos extraídos.....: {len(dados['lancamentos'])}")
    log(f"  Beneficiários distintos...: {len({r['carteirinha'] for r in dados['lancamentos']})}")
    log(f"  Famílias (totais no PDF)..: {len(dados['totais_familia'])}")
    log(f"  Soma extraída.............: {moeda_br(soma)}")
    if vl_copart is not None:
        diferenca = soma - Decimal(str(vl_copart))
        confere = abs(diferenca) < Decimal("0.01")
        status = "OK" if confere else f"DIFERENÇA DE {moeda_br(diferenca)}"
        log(f"  Vl. co-partic. cabeçalho..: {moeda_br(vl_copart)}  -> {status}")
    if dados["nao_reconhecidas"]:
        log(f"  ATENÇÃO: {len(dados['nao_reconhecidas'])} linha(s) não reconhecida(s) "
            f"- veja a aba 'Conferência'.")
    log(f"  EXCEL GERADO: {caminho_saida}")
    return {
        "saida": caminho_saida,
        "confere": confere,
        "nao_reconhecidas": len(dados["nao_reconhecidas"]),
    }


# ----------------------------------------------------------------------------
# Fluxo principal
# ----------------------------------------------------------------------------
def converter_pdf(caminho_pdf, log=print):
    if detectar_layout(caminho_pdf) == LAYOUT_COPARTICIPACAO:
        return converter_pdf_copart(caminho_pdf, log)
    log("=" * 60)
    log(f"LENDO O ARQUIVO: {caminho_pdf}")
    dados = extrair_dados_pdf(caminho_pdf)

    caminho_saida = os.path.splitext(caminho_pdf)[0] + ".xlsx"
    try:
        gerar_excel(dados, caminho_saida, os.path.basename(caminho_pdf))
    except PermissionError:
        base = os.path.splitext(caminho_pdf)[0]
        caminho_saida = f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        log("  Aviso: o Excel original está aberto. Salvando com outro nome.")
        gerar_excel(dados, caminho_saida, os.path.basename(caminho_pdf))

    soma = sum(Decimal(str(r["vl_mensalidade"] or 0)) + Decimal(str(r["vl_outros"] or 0))
               for r in dados["beneficiarios"])
    bruto = dados["cabecalho"]["Vl. Bruto"]
    confere = None
    log(f"  Páginas lidas.............: {dados['total_paginas']}")
    log(f"  Beneficiários extraídos...: {len(dados['beneficiarios'])}")
    log(f"  Famílias (totais no PDF)..: {len(dados['totais_familia'])}")
    log(f"  Soma extraída.............: {moeda_br(soma)}")
    if bruto is not None:
        diferenca = soma - Decimal(str(bruto))
        confere = abs(diferenca) < Decimal("0.01")
        status = "OK" if confere else f"DIFERENÇA DE {moeda_br(diferenca)}"
        log(f"  Vl. Bruto do cabeçalho....: {moeda_br(bruto)}  -> {status}")
    if dados["nao_reconhecidas"]:
        log(f"  ATENÇÃO: {len(dados['nao_reconhecidas'])} linha(s) não reconhecida(s) "
            f"- veja a aba 'Conferência'.")
    log(f"  EXCEL GERADO: {caminho_saida}")
    return {
        "saida": caminho_saida,
        "confere": confere,
        "nao_reconhecidas": len(dados["nao_reconhecidas"]),
    }


def listar_pdfs(argumentos, log=print):
    arquivos = []
    for item in argumentos:
        item = item.strip().strip('"')
        if os.path.isdir(item):
            for nome in sorted(os.listdir(item)):
                if nome.lower().endswith(".pdf"):
                    arquivos.append(os.path.join(item, nome))
        elif os.path.isfile(item) and item.lower().endswith(".pdf"):
            arquivos.append(item)
        else:
            log(f"Ignorado (não é PDF/pasta válida): {item}")
    return arquivos


def converter_lista(arquivos, log=print, ao_avancar=None):
    gerados = 0
    alertas = 0
    for indice, caminho_pdf in enumerate(arquivos, start=1):
        try:
            resultado = converter_pdf(caminho_pdf, log)
            gerados += 1
            if resultado["confere"] is False or resultado["nao_reconhecidas"]:
                alertas += 1
        except Exception as erro:
            log(f"  ERRO ao converter {caminho_pdf}: {erro}")
        if ao_avancar:
            ao_avancar(indice, len(arquivos))
    log("=" * 60)
    log(f"Concluído: {gerados} de {len(arquivos)} arquivo(s) convertido(s).")
    if alertas:
        log(f"Atenção: {alertas} arquivo(s) com diferença ou linhas não reconhecidas.")
    return gerados, alertas


# ----------------------------------------------------------------------------
# Interface gráfica (tkinter)
# ----------------------------------------------------------------------------
def abrir_no_sistema(caminho):
    if sys.platform.startswith("win"):
        os.startfile(caminho)
    elif sys.platform == "darwin":
        import subprocess
        subprocess.Popen(["open", caminho])
    else:
        import subprocess
        subprocess.Popen(["xdg-open", caminho])


def iniciar_interface():
    import queue
    import threading
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter.scrolledtext import ScrolledText

    janela = tk.Tk()
    janela.title("Unimed - Conversor de Demonstrativo PDF para Excel")
    janela.geometry("900x620")
    janela.minsize(720, 480)

    fila = queue.Queue()
    arquivos = []
    estado = {"rodando": False, "ultima_pasta": ""}

    # --- Topo: botões de seleção ---
    quadro_botoes = ttk.Frame(janela, padding=(10, 10, 10, 0))
    quadro_botoes.pack(fill="x")

    # --- Lista de arquivos ---
    quadro_lista = ttk.LabelFrame(janela, text="PDFs selecionados", padding=8)
    quadro_lista.pack(fill="both", expand=False, padx=10, pady=8)
    lista = tk.Listbox(quadro_lista, height=8, selectmode="extended")
    barra_lista = ttk.Scrollbar(quadro_lista, orient="vertical", command=lista.yview)
    lista.configure(yscrollcommand=barra_lista.set)
    lista.pack(side="left", fill="both", expand=True)
    barra_lista.pack(side="right", fill="y")

    rotulo_qtd = ttk.Label(janela, text="0 arquivo(s) selecionado(s)")
    rotulo_qtd.pack(anchor="w", padx=12)

    # --- Progresso ---
    quadro_prog = ttk.Frame(janela, padding=(10, 6))
    quadro_prog.pack(fill="x")
    progresso = ttk.Progressbar(quadro_prog, mode="determinate")
    progresso.pack(side="left", fill="x", expand=True)
    rotulo_prog = ttk.Label(quadro_prog, text="", width=12, anchor="e")
    rotulo_prog.pack(side="right", padx=(8, 0))

    # --- Log ---
    quadro_log = ttk.LabelFrame(janela, text="Resultado", padding=8)
    quadro_log.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    texto_log = ScrolledText(quadro_log, height=12, font=("Consolas", 9), state="disabled")
    texto_log.pack(fill="both", expand=True)
    texto_log.tag_configure("ok", foreground="#1E7B34")
    texto_log.tag_configure("erro", foreground="#C00000")
    texto_log.tag_configure("aviso", foreground="#B36B00")

    def atualizar_lista():
        lista.delete(0, "end")
        for caminho in arquivos:
            lista.insert("end", caminho)
        rotulo_qtd.config(text=f"{len(arquivos)} arquivo(s) selecionado(s)")

    def adicionar(caminhos):
        for caminho in caminhos:
            caminho = os.path.normpath(caminho)
            if caminho not in arquivos:
                arquivos.append(caminho)
        if caminhos:
            estado["ultima_pasta"] = os.path.dirname(os.path.normpath(caminhos[-1]))
        atualizar_lista()

    def selecionar_pdfs():
        selecionados = filedialog.askopenfilenames(
            title="Selecione o(s) PDF(s) do Demonstrativo Unimed",
            initialdir=estado["ultima_pasta"] or None,
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
        )
        adicionar(list(selecionados))

    def selecionar_pasta():
        pasta = filedialog.askdirectory(
            title="Selecione a pasta com os PDFs",
            initialdir=estado["ultima_pasta"] or None,
        )
        if not pasta:
            return
        encontrados = listar_pdfs([pasta], escrever_log)
        if not encontrados:
            messagebox.showinfo("Nenhum PDF", "Não há arquivos PDF nessa pasta.")
            return
        adicionar(encontrados)
        estado["ultima_pasta"] = os.path.normpath(pasta)

    def remover_selecionados():
        for indice in reversed(lista.curselection()):
            del arquivos[indice]
        atualizar_lista()

    def limpar():
        arquivos.clear()
        atualizar_lista()

    def escrever_log(mensagem):
        fila.put(("log", mensagem))

    def inserir_no_log(mensagem):
        tag = None
        if "-> OK" in mensagem or mensagem.startswith("Concluído"):
            tag = "ok"
        elif "ERRO" in mensagem or "DIFERENÇA" in mensagem:
            tag = "erro"
        elif "ATENÇÃO" in mensagem or "Atenção" in mensagem or "Aviso" in mensagem:
            tag = "aviso"
        texto_log.configure(state="normal")
        texto_log.insert("end", mensagem + "\n", tag)
        texto_log.see("end")
        texto_log.configure(state="disabled")

    def ao_avancar(atual, total):
        fila.put(("progresso", (atual, total)))

    def trabalho(lista_arquivos):
        try:
            gerados, alertas = converter_lista(lista_arquivos, escrever_log, ao_avancar)
            fila.put(("fim", (gerados, len(lista_arquivos), alertas)))
        except Exception as erro:
            fila.put(("log", f"ERRO inesperado: {erro}"))
            fila.put(("fim", (0, len(lista_arquivos), 0)))

    def converter():
        if estado["rodando"]:
            return
        if not arquivos:
            messagebox.showwarning("Atenção", "Selecione pelo menos um PDF ou uma pasta.")
            return
        estado["rodando"] = True
        for botao in botoes_bloqueaveis:
            botao.configure(state="disabled")
        progresso.configure(maximum=len(arquivos), value=0)
        rotulo_prog.config(text=f"0 / {len(arquivos)}")
        threading.Thread(target=trabalho, args=(list(arquivos),), daemon=True).start()

    def abrir_pasta():
        pasta = estado["ultima_pasta"]
        if not pasta and arquivos:
            pasta = os.path.dirname(arquivos[0])
        if pasta and os.path.isdir(pasta):
            try:
                abrir_no_sistema(pasta)
            except Exception as erro:
                messagebox.showerror("Erro", f"Não foi possível abrir a pasta:\n{erro}")
        else:
            messagebox.showinfo("Pasta", "Nenhuma pasta selecionada ainda.")

    def processar_fila():
        try:
            while True:
                tipo, conteudo = fila.get_nowait()
                if tipo == "log":
                    inserir_no_log(conteudo)
                elif tipo == "progresso":
                    atual, total = conteudo
                    progresso.configure(value=atual)
                    rotulo_prog.config(text=f"{atual} / {total}")
                elif tipo == "fim":
                    gerados, total, alertas = conteudo
                    estado["rodando"] = False
                    for botao in botoes_bloqueaveis:
                        botao.configure(state="normal")
                    mensagem = f"{gerados} de {total} arquivo(s) convertido(s)."
                    if alertas:
                        mensagem += (f"\n\n{alertas} arquivo(s) com diferença ou linhas não "
                                     f"reconhecidas. Confira a aba 'Conferência' desses Excel.")
                        messagebox.showwarning("Concluído com avisos", mensagem)
                    else:
                        messagebox.showinfo("Concluído", mensagem)
        except queue.Empty:
            pass
        janela.after(100, processar_fila)

    def ao_fechar():
        if estado["rodando"] and not messagebox.askyesno(
            "Sair", "A conversão ainda está em andamento. Deseja sair mesmo assim?"
        ):
            return
        janela.destroy()

    btn_pdfs = ttk.Button(quadro_botoes, text="Selecionar PDFs...", command=selecionar_pdfs)
    btn_pasta = ttk.Button(quadro_botoes, text="Selecionar pasta...", command=selecionar_pasta)
    btn_remover = ttk.Button(quadro_botoes, text="Remover selecionados", command=remover_selecionados)
    btn_limpar = ttk.Button(quadro_botoes, text="Limpar lista", command=limpar)
    btn_converter = ttk.Button(quadro_botoes, text="CONVERTER PARA EXCEL", command=converter)
    btn_abrir = ttk.Button(quadro_botoes, text="Abrir pasta", command=abrir_pasta)
    for botao in (btn_pdfs, btn_pasta, btn_remover, btn_limpar):
        botao.pack(side="left", padx=(0, 6))
    btn_abrir.pack(side="right")
    btn_converter.pack(side="right", padx=(0, 6))
    botoes_bloqueaveis = [btn_pdfs, btn_pasta, btn_remover, btn_limpar, btn_converter]

    lista.bind("<Delete>", lambda evento: remover_selecionados())
    janela.protocol("WM_DELETE_WINDOW", ao_fechar)

    # PDFs/pastas passados ao abrir (ex.: arrastados sobre o .py) já entram na lista
    if sys.argv[1:]:
        adicionar(listar_pdfs(sys.argv[1:], escrever_log))

    inserir_no_log("Selecione os PDFs ou uma pasta e clique em 'CONVERTER PARA EXCEL'.")
    janela.after(100, processar_fila)
    janela.mainloop()


# ----------------------------------------------------------------------------
# Modo linha de comando (sem janela): python pdf_unimed_para_excel.py --console <pdfs/pastas>
# ----------------------------------------------------------------------------
def main_console(argumentos):
    if not argumentos:
        caminho = input("Informe o caminho do PDF ou da pasta com PDFs: ").strip().strip('"')
        argumentos = [caminho] if caminho else []
    arquivos = listar_pdfs(argumentos)
    if not arquivos:
        print("Nenhum arquivo PDF selecionado.")
        return
    converter_lista(arquivos)


def main():
    argumentos = sys.argv[1:]
    if "--console" in argumentos:
        argumentos = [a for a in argumentos if a != "--console"]
        main_console(argumentos)
        if sys.stdin.isatty():
            try:
                input("\nPressione ENTER para fechar...")
            except EOFError:
                pass
        return
    try:
        iniciar_interface()
    except ImportError:
        print("tkinter não está disponível. Usando o modo console.")
        main_console(argumentos)
    except Exception as erro:
        if "display" in str(erro).lower():
            print("Não foi possível abrir a janela. Usando o modo console.")
            main_console(argumentos)
        else:
            raise


if __name__ == "__main__":
    main()
