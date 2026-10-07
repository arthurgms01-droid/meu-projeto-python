# -*- coding: utf-8 -*-
"""
Conversor do "Demonstrativo Mensalidade por Pagador (Portal Web)" - Unimed
de PDF para Excel (.xlsx) já formatado.

Uso:
    python pdf_unimed_para_excel.py                    -> abre janela para escolher PDF(s)
    python pdf_unimed_para_excel.py arquivo.pdf        -> converte um arquivo
    python pdf_unimed_para_excel.py C:/pasta/com/pdfs  -> converte todos os PDFs da pasta

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


# ----------------------------------------------------------------------------
# Fluxo principal
# ----------------------------------------------------------------------------
def converter_pdf(caminho_pdf):
    print("=" * 60)
    print(f"LENDO O ARQUIVO: {caminho_pdf}")
    dados = extrair_dados_pdf(caminho_pdf)

    caminho_saida = os.path.splitext(caminho_pdf)[0] + ".xlsx"
    try:
        gerar_excel(dados, caminho_saida, os.path.basename(caminho_pdf))
    except PermissionError:
        base = os.path.splitext(caminho_pdf)[0]
        caminho_saida = f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        print("  Aviso: o Excel original está aberto. Salvando com outro nome.")
        gerar_excel(dados, caminho_saida, os.path.basename(caminho_pdf))

    soma = sum(Decimal(str(r["vl_mensalidade"] or 0)) + Decimal(str(r["vl_outros"] or 0))
               for r in dados["beneficiarios"])
    bruto = dados["cabecalho"]["Vl. Bruto"]
    print(f"  Páginas lidas.............: {dados['total_paginas']}")
    print(f"  Beneficiários extraídos...: {len(dados['beneficiarios'])}")
    print(f"  Famílias (totais no PDF)..: {len(dados['totais_familia'])}")
    print(f"  Soma extraída.............: {moeda_br(soma)}")
    if bruto is not None:
        diferenca = soma - Decimal(str(bruto))
        status = "OK" if abs(diferenca) < Decimal("0.01") else f"DIFERENÇA DE {moeda_br(diferenca)}"
        print(f"  Vl. Bruto do cabeçalho....: {moeda_br(bruto)}  -> {status}")
    if dados["nao_reconhecidas"]:
        print(f"  ATENÇÃO: {len(dados['nao_reconhecidas'])} linha(s) não reconhecida(s) "
              f"- veja a aba 'Conferência'.")
    print(f"  EXCEL GERADO: {caminho_saida}")
    return caminho_saida


def listar_pdfs(argumentos):
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
            print(f"Ignorado (não é PDF/pasta válida): {item}")
    return arquivos


def selecionar_pdfs_interativo():
    try:
        import tkinter as tk
        from tkinter import filedialog

        raiz = tk.Tk()
        raiz.withdraw()
        raiz.attributes("-topmost", True)
        selecionados = filedialog.askopenfilenames(
            title="Selecione o(s) PDF(s) do Demonstrativo Unimed",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
        )
        raiz.destroy()
        return list(selecionados)
    except Exception:
        caminho = input("Informe o caminho do PDF ou da pasta com PDFs: ").strip().strip('"')
        return [caminho] if caminho else []


def main():
    argumentos = sys.argv[1:]
    if not argumentos:
        argumentos = selecionar_pdfs_interativo()

    arquivos = listar_pdfs(argumentos)
    if not arquivos:
        print("Nenhum arquivo PDF selecionado.")
        return

    gerados = 0
    for caminho_pdf in arquivos:
        try:
            converter_pdf(caminho_pdf)
            gerados += 1
        except Exception as erro:
            print(f"  ERRO ao converter {caminho_pdf}: {erro}")

    print("=" * 60)
    print(f"Concluído: {gerados} de {len(arquivos)} arquivo(s) convertido(s).")


if __name__ == "__main__":
    try:
        main()
    finally:
        if not sys.argv[1:] or sys.stdin.isatty():
            try:
                input("\nPressione ENTER para fechar...")
            except EOFError:
                pass
