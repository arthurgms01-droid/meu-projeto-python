# -*- coding: utf-8 -*-
"""
ROBO DE DOWNLOADS NDI - HAPVIDA - RELATÓRIOS ANALÍTICOS - PLAYWRIGHT

Objetivo:
- Ler o arquivo NDI_CONTRATOS_DOWNLOAD.xlsx.
- Acessar o portal Hapvida/NDI.
- Localizar contrato e vencimento (mesmo fluxo do robô de Boletos e NFs).
- Na tela "Detalhes da fatura", baixar os relatórios analíticos:
  - Mensalidade     : Relatório.csv, Relatório.txt e Relatório.pdf
                      (linhas "Relatório analítico (CSV/TXT/PDF)").
  - Coparticipação  : CSV, TXT e PDF da seção "Relatórios"
                      (botões "Baixar relatório", entregues em .zip).
- Opcionalmente (BAIXAR_BOLETO_NF=1) baixar também Boleto.pdf e Nota Fiscal.

Mapeamento dos botões:
    Obtido da gravação do Gravador_Acoes_Portal_Generico (sessão de 05/10/2026).
    Comportamentos observados e tratados:
    - CSV da mensalidade e os três .zip da coparticipação disparam download
      por meio de uma aba auxiliar que abre e fecha sozinha.
    - PDF e TXT da mensalidade abrem uma nova aba com a URL direta do arquivo
      (visualizador de PDF / texto puro), sem evento de download.

Diferenciais mantidos da versão Playwright original:
- Captura de download por evento, nova aba, URL direta e Blob.
- Salva screenshots em erros e diagnóstico estruturado de falhas (CSV + JSON).
- Não altera o arquivo NDI_CONTRATOS_DOWNLOAD.xlsx por padrão.
- Controle TXT de downloads, duplicidade (retomada), evidência por contrato e resumo final.
- Retentativa individual por documento e por contrato.
- Cronometragem detalhada em CSV.

Novidades desta versão:
- Rastreamento de cada arquivo salvo em Logs/rastreamento_downloads_analiticos_ndi.csv
  (método de captura, bytes, SHA-256, validação e membros do ZIP).
- Validação do conteúdo: CRC do ZIP, assinatura do PDF e texto plausível.
- Extração automática dos .zip (EXTRAIR_ZIP_ANALITICO=0 desativa).

Organização dos arquivos:
- Pasta única por tipo, sem subpastas por mês ou contrato:
      Downloads/MENSALIDADE/
      Downloads/COPARTICIPACAO/
- Os arquivos mantêm o nome original entregue pelo portal
  (ex.: EMPRESA_FM_0T9W1_REMESSA_6365924.csv). O nome padrão do robô só é
  usado quando o portal não informa nome (captura por Blob).
- Coparticipação: o .zip é extraído e removido; ficam só os arquivos.
- Evidências por contrato: Logs/Evidencias/.

Instalação:
    py -m pip install pandas openpyxl playwright
    py -m playwright install chromium

Credenciais:
    A senha não fica no script. Crie um arquivo .env na mesma pasta do robô:
        HAPVIDA_EMAIL=usuario_da_nova_empresa
        HAPVIDA_SENHA=senha_da_nova_empresa

    O .env tem prioridade sobre variáveis de ambiente do Windows com o mesmo
    nome. Se nenhuma credencial for encontrada, o script pergunta no terminal.

Configurações opcionais (.env ou variáveis de ambiente):
    FORMATOS_ANALITICOS=CSV,TXT,PDF      formatos a baixar
    BAIXAR_BOLETO_NF=0                   1 = baixa também Boleto e Nota Fiscal
    EXTRAIR_ZIP_ANALITICO=1              extrai e remove o .zip; 0 = mantém apenas o .zip
    TIMEOUT_DOWNLOAD_ANALITICO_MS=60000  espera máxima por arquivo
    WORKERS_PARALELOS=1                  janelas do navegador trabalhando ao mesmo tempo
    INTERVALO_INICIO_WORKERS_SEG=5       intervalo entre a abertura de cada janela
    PULAR_CONTRATO_CONCLUIDO=1           0 = sempre reabre contratos já concluídos
"""

import csv
import ctypes
import getpass
import builtins
import hashlib
import json
import os
import queue
import re
import shutil
import threading
import time
import zipfile
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

import pandas as pd
# --- INÍCIO INSERÇÃO PLAYWRIGHT (IMPORTS) ---
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import expect, sync_playwright
# --- FIM INSERÇÃO PLAYWRIGHT (IMPORTS) ---


# --- INÍCIO INSERÇÃO CREDENCIAIS E CONFIGURAÇÕES VIA .ENV ---
def carregar_arquivo_env(caminho_env=None, sobrescrever=True):
    """
    Carrega pares CHAVE=VALOR do arquivo .env para os.environ.

    Não depende de bibliotecas externas. Aceita linhas em branco, comentários
    iniciados por #, prefixo opcional "export" e valores entre aspas.
    Com sobrescrever=True o .env prevalece sobre variáveis já existentes no
    Windows, evitando que credenciais antigas gravadas com setx sejam usadas.
    """
    caminho_env = Path(caminho_env) if caminho_env else Path(__file__).resolve().parent / ".env"
    if not caminho_env.exists():
        return []

    carregadas = []
    try:
        with open(caminho_env, "r", encoding="utf-8-sig") as arquivo:
            for linha in arquivo:
                linha = linha.strip()
                if not linha or linha.startswith("#") or "=" not in linha:
                    continue
                if linha.lower().startswith("export "):
                    linha = linha[7:].strip()
                chave, valor = linha.split("=", 1)
                chave = chave.strip()
                valor = valor.strip()
                if not chave:
                    continue
                if len(valor) >= 2 and valor[0] == valor[-1] and valor[0] in {'"', "'"}:
                    valor = valor[1:-1]
                if sobrescrever or chave not in os.environ:
                    os.environ[chave] = valor
                    carregadas.append(chave)
    except Exception as e:
        print(f"[AVISO] Não foi possível ler o arquivo .env: {e}")
    return carregadas


def preparar_caminho_longo_windows(caminho):
    """
    No Windows, caminhos com mais de 259 caracteres falham ao gravar arquivos.
    A pasta do robô dentro do OneDrive já é longa, e somada às subpastas e ao
    nome padronizado dos arquivos ultrapassa esse limite.

    O prefixo de caminho estendido faz o Windows aceitar caminhos longos sem depender de
    configuração do registro. CAMINHO_LONGO_WINDOWS=0 desativa.
    """
    caminho = Path(caminho)
    if os.name != "nt":
        return caminho
    if os.getenv("CAMINHO_LONGO_WINDOWS", "1").strip().lower() not in {"1", "true", "sim", "s"}:
        return caminho
    texto = str(caminho)
    if texto.startswith("\\\\"):
        # Já possui prefixo ou é caminho de rede (UNC): mantém como está.
        return caminho
    if re.match(r"^[A-Za-z]:\\", texto):
        return Path("\\\\?\\" + texto)
    return caminho


# O .env é lido antes da criação da classe, pois as configurações do robô
# também são obtidas por os.getenv() no __init__.
VARIAVEIS_ENV_CARREGADAS = carregar_arquivo_env()

# As credenciais da empresa ficam somente no .env:
#   HAPVIDA_EMAIL
#   HAPVIDA_SENHA
LOGIN_EMAIL_PADRAO = ""
LOGIN_SENHA_PADRAO = ""
# --- FIM INSERÇÃO CREDENCIAIS E CONFIGURAÇÕES VIA .ENV ---


# --- INÍCIO INSERÇÃO EXECUÇÃO PARALELA (INFRAESTRUTURA) ---
# Trava única para os arquivos de log compartilhados (controle TXT, CSVs de
# tempos/rastreamento/diagnóstico, logs de erro). Reentrante porque algumas
# gravações inicializam o arquivo dentro da própria gravação.
TRAVA_ARQUIVOS = threading.RLock()
TRAVA_CONSOLE = threading.Lock()
_contexto_thread = threading.local()


def print(*args, **kwargs):
    """
    Substitui o print deste módulo: com mais de uma janela em paralelo, cada
    linha do console sai inteira e com o prefixo da janela ([W1], [W2]...).
    Com uma única janela não há prefixo e a saída fica igual à anterior.
    """
    prefixo = getattr(_contexto_thread, "prefixo", "")
    with TRAVA_CONSOLE:
        if prefixo:
            builtins.print(prefixo, *args, **kwargs)
        else:
            builtins.print(*args, **kwargs)
# --- FIM INSERÇÃO EXECUÇÃO PARALELA (INFRAESTRUTURA) ---


# --- INÍCIO INSERÇÃO PLAYWRIGHT (CLASSE MANTENDO FLUXO ORIGINAL) ---
class HapvidaNDIAnaliticosAutomation:
    TIPO_CARNET = "CARNET COMPLEMENTAR-COPARTICIPACAO"
    TIPO_MENSALIDADE = "MENSALIDADE PLANO SAUDE"
    PRIORIDADE_TIPOS = [TIPO_CARNET, TIPO_MENSALIDADE]

    # --- INÍCIO INSERÇÃO ESCOPO DE DOWNLOAD: RELATÓRIOS ANALÍTICOS ---
    FORMATOS_ANALITICOS_SUPORTADOS = ("CSV", "TXT", "PDF")
    DOCUMENTOS_ANALITICOS = ("ANALITICO_CSV", "ANALITICO_TXT", "ANALITICO_PDF")
    # Coparticipação por teclado: 1 clique no título "Relatórios" e TAB até o botão.
    TABS_RELATORIO_COPARTICIPACAO = {"CSV": 1, "TXT": 2, "PDF": 3}
    # Valor padrão; o __init__ redefine conforme FORMATOS_ANALITICOS e BAIXAR_BOLETO_NF.
    DOCUMENTOS_PERMITIDOS = DOCUMENTOS_ANALITICOS
    # Ordem fixa das colunas de documentos no controle TXT. Documentos fora do
    # escopo da execução são gravados com "-", mantendo o layout estável.
    COLUNAS_DOCUMENTOS_CONTROLE = (
        ("BOLETO", "Boleto"),
        ("NF", "Nota Fiscal"),
        ("ANALITICO_CSV", "Analítico CSV"),
        ("ANALITICO_TXT", "Analítico TXT"),
        ("ANALITICO_PDF", "Analítico PDF"),
    )
    # --- FIM INSERÇÃO ESCOPO DE DOWNLOAD: RELATÓRIOS ANALÍTICOS ---

    # --- INÍCIO INSERÇÃO CONFIRMAÇÃO PRECISA DAS TELAS ---
    # Os seletores antigos (//*[contains(., 'texto')]) casavam com o próprio
    # <html>, cujo texto inclui elementos ocultos e a tela anterior; a espera
    # terminava antes de a tela mudar. Estes miram o botão/aba que só existe
    # na tela de destino, e "visible=true" faz o Playwright usar o primeiro
    # elemento visível (wait_for_selector avalia só o primeiro que casar).
    SELETOR_TELA_CONTRATO = (
        "xpath=//*[self::button or @role='button' or @role='tab']"
        "[contains(normalize-space(.), 'Extrato')] >> visible=true"
    )
    SELETOR_TELA_EXTRATO = (
        "xpath=//*[self::button or @role='button' or @role='tab']"
        "[contains(normalize-space(.), 'Em Aberto') or contains(normalize-space(.), 'Histórico')"
        " or contains(normalize-space(.), 'Historico')] >> visible=true"
    )
    # --- FIM INSERÇÃO CONFIRMAÇÃO PRECISA DAS TELAS ---

    TIMEOUT_CURTO = 5_000
    TIMEOUT_PADRAO = 20_000
    TIMEOUT_LONGO = 60_000
    TIMEOUT_DOWNLOAD = 120_000

    def __init__(self, file_path):
        self.file_path = Path(file_path).resolve()
        self.url_login = (
            "https://hapvidaadb2bprd.b2clogin.com/"
            "hapvidaadb2bprd.onmicrosoft.com/oauth2/v2.0/authorize"
            "?p=B2C_1_portal-empresa-login"
            "&client_id=a873a0fd-ed81-4cf2-bb25-0f6dca61bc46"
            "&nonce=defaultNonce"
            "&redirect_uri=https%3A%2F%2Fportal-empresa.hapvidagndi.com.br%2Fautenticando"
            "&scope=openid%20offline_access"
            "&response_type=code+id_token"
            "&prompt=login"
        )
        self.url_painel = None

        # --- INÍCIO INSERÇÃO CONFIGURAÇÃO DOS RELATÓRIOS ANALÍTICOS ---
        valores_verdadeiros = {"1", "true", "sim", "s"}
        formatos_env = os.getenv("FORMATOS_ANALITICOS", "CSV,TXT,PDF")
        self.formatos_analiticos = []
        for formato in re.split(r"[,;\s]+", formatos_env.upper()):
            if formato in self.FORMATOS_ANALITICOS_SUPORTADOS and formato not in self.formatos_analiticos:
                self.formatos_analiticos.append(formato)
        if not self.formatos_analiticos:
            self.formatos_analiticos = list(self.FORMATOS_ANALITICOS_SUPORTADOS)

        self.baixar_boleto_nf = os.getenv("BAIXAR_BOLETO_NF", "0").strip().lower() in valores_verdadeiros
        documentos = ["BOLETO", "NF"] if self.baixar_boleto_nf else []
        documentos.extend(f"ANALITICO_{formato}" for formato in self.formatos_analiticos)
        self.DOCUMENTOS_PERMITIDOS = tuple(documentos)

        self.extrair_zip_ativo = os.getenv("EXTRAIR_ZIP_ANALITICO", "1").strip().lower() in valores_verdadeiros
        self.timeout_download_analitico = int(os.getenv("TIMEOUT_DOWNLOAD_ANALITICO_MS", "60000"))
        self.timeout_botao_analitico_total = int(os.getenv("TIMEOUT_BOTAO_ANALITICO_TOTAL_MS", "8000"))
        self.intervalo_monitoramento_analitico = max(
            100,
            int(os.getenv("INTERVALO_MONITORAMENTO_ANALITICO_MS", "250")),
        )
        # Quando o clique abre uma aba com URL direta, o robô aguarda esta
        # carência por um evento de download antes de buscar a URL por request.
        self.carencia_nova_aba_analitico = max(
            0.0,
            float(os.getenv("CARENCIA_NOVA_ABA_ANALITICO_SEG", "1.5").replace(",", ".")),
        )
        # Coparticipação: método principal por teclado (clique em "Relatórios" + TAB + ENTER).
        # COPARTICIPACAO_POR_TECLADO=0 volta a usar somente os seletores dos botões.
        self.coparticipacao_por_teclado = os.getenv(
            "COPARTICIPACAO_POR_TECLADO", "1"
        ).strip().lower() in valores_verdadeiros
        self.intervalo_tab_analitico = max(
            0.0,
            float(os.getenv("INTERVALO_TAB_ANALITICO_SEG", "0.20").replace(",", ".")),
        )
        # --- FIM INSERÇÃO CONFIGURAÇÃO DOS RELATÓRIOS ANALÍTICOS ---

        # --- INÍCIO INSERÇÃO ESPERA DAS LINHAS DO EXTRATO ---
        # Tempo máximo para as linhas da tabela do extrato aparecerem depois de
        # abrir uma aba/página, e tempo sem linhas e sem indicador de carregamento
        # para considerar a tabela realmente vazia.
        self.timeout_linhas_extrato = int(os.getenv("TIMEOUT_LINHAS_EXTRATO_MS", "15000"))
        self.carencia_extrato_vazio = max(
            0.5,
            float(os.getenv("CARENCIA_EXTRATO_VAZIO_SEG", "3").replace(",", ".")),
        )
        self.linhas_extrato_vistas = []
        # --- FIM INSERÇÃO ESPERA DAS LINHAS DO EXTRATO ---

        # URL da lista de faturas do contrato atual, usada para reabrir a fatura.
        self.url_extrato_atual = None
        self.ultimo_erro_captura_analitico = ""

        self.base_dir = preparar_caminho_longo_windows(Path(__file__).resolve().parent)
        self.downloads_dir = self.base_dir / "Downloads"
        self.logs_dir = self.base_dir / "Logs"
        self.screenshots_dir = self.logs_dir / "Screenshots"
        # Evidências por contrato ficam nos Logs, fora das pastas de download.
        self.evidencias_dir = self.logs_dir / "Evidencias"
        self.downloads_dir.mkdir(parents=True, exist_ok=True)
        self.logs_dir.mkdir(parents=True, exist_ok=True)
        self.screenshots_dir.mkdir(parents=True, exist_ok=True)
        self.evidencias_dir.mkdir(parents=True, exist_ok=True)

        self.log_file = self.logs_dir / "downloads_concluidos.log"
        self.erros_log_file = self.logs_dir / "erros_execucao.log"
        # Registro, na raiz de Logs, do que havia na tela quando um botão de
        # relatório analítico não foi encontrado.
        self.arquivo_tela_sem_botao = self.logs_dir / "tela_sem_botao_analitico_ndi.txt"

        # Para não pular downloads antigos, deixe 0.
        # Para reativar o log e pular arquivos já baixados, defina USAR_LOG_DOWNLOADS=1.
        self.usar_log_downloads = os.getenv("USAR_LOG_DOWNLOADS", "0").strip().lower() in {"1", "true", "sim", "s"}
        self.downloads_concluidos = self.ler_log()

        self.contrato_atual = None
        self.vencimento_atual = None
        self.tipo_fatura_atual = None
        self.docs_linha_atual = []
        # Guarda a origem da fatura atual para permitir reabertura completa
        # durante a retentativa do boleto.
        self.ocorrencia_fatura_atual = None

        # --- INÍCIO INSERÇÃO CONTROLE TXT / NÃO ALTERAR EXCEL ORIGINAL ---
        # Regra solicitada: o arquivo NDI_CONTRATOS_DOWNLOAD.xlsx será usado somente para leitura.
        # O controle da execução, duplicidade, evidências e resumo serão gravados em arquivos TXT.
        self.atualizar_excel_original = os.getenv("ATUALIZAR_EXCEL_ORIGINAL", "0").strip().lower() in {"1", "true", "sim", "s"}
        # Arquivo próprio deste robô, para não misturar com o controle do robô
        # de Boletos e NFs caso os dois sejam executados na mesma pasta.
        self.controle_txt_file = self.logs_dir / "controle_downloads_analiticos_ndi.txt"
        self.controle_txt_cabecalho = (
            ["Contrato", "Vencimento", "Tipo da fatura"]
            + [rotulo for _, rotulo in self.COLUNAS_DOCUMENTOS_CONTROLE]
            + ["Status", "Mensagem de erro", "Data/hora"]
        )
        self.status_docs_fatura_atual = {}
        self.caminhos_docs_fatura_atual = {}
        self.mensagens_fatura_atual = []
        self.controle_fatura_registrado = False

        # --- INÍCIO INSERÇÃO DIAGNÓSTICO ESTRUTURADO DE FALHAS ---
        self.diagnosticos_dir = self.logs_dir / "Diagnosticos"
        self.diagnosticos_dir.mkdir(parents=True, exist_ok=True)
        self.arquivo_diagnosticos_csv = self.logs_dir / "diagnostico_falhas_ndi.csv"
        self.arquivo_resumo_diagnosticos = self.logs_dir / "resumo_diagnostico_falhas_ndi.txt"
        self.diagnosticos_execucao = []
        self.falha_atual_documento = {}
        self.contagem_codigos_falha = {}
        self.chaves_diagnostico_registradas = set()
        self.documento_download_atual = ""
        self.tentativa_download_atual = 0
        self.salvar_html_diagnostico = os.getenv(
            "SALVAR_HTML_DIAGNOSTICO", "0"
        ).strip().lower() in {"1", "true", "sim", "s"}
        self.inicializar_arquivo_diagnosticos()
        # --- FIM INSERÇÃO DIAGNÓSTICO ESTRUTURADO DE FALHAS ---

        # --- INÍCIO INSERÇÃO RASTREAMENTO DE DOWNLOADS (MODELO DO GRAVADOR) ---
        self.arquivo_rastreamento_csv = self.logs_dir / "rastreamento_downloads_analiticos_ndi.csv"
        self.inicializar_arquivo_rastreamento()
        # --- FIM INSERÇÃO RASTREAMENTO DE DOWNLOADS (MODELO DO GRAVADOR) ---

        self.pasta_download_atual = self.downloads_dir
        self.pasta_contrato_atual = self.downloads_dir
        self.evidencias_contrato_atual = []
        self.index_linha_atual = None
        self.resumo_execucao = {
            "contratos_processados": 0,
            "contratos_sucesso": 0,
            "contratos_pendencia": 0,
            "contratos_erro": 0,
            "faturas_processadas": 0,
            "arquivos_baixados": 0,
            "arquivos_pulados": 0,
            "arquivos_com_erro": 0,
            "retentativas_download": 0,
            "arquivos_recuperados_retentativa": 0,
            "arquivos_falha_apos_retentativas": 0,
            "falhas_diagnosticadas": 0,
            "falhas_recuperadas_apos_diagnostico": 0,
            "BOLETO": 0,
            "NF": 0,
            "ANALITICO_CSV": 0,
            "ANALITICO_TXT": 0,
            "ANALITICO_PDF": 0,
            "arquivos_extraidos_zip": 0,
            "contratos_pulados_concluidos": 0,
        }
        self.inicializar_controle_txt()
        self.documentos_ok_controle = self.carregar_documentos_ok_controle()

        # --- INÍCIO INSERÇÃO PULAR CONTRATO JÁ CONCLUÍDO ---
        # Marcador por contrato + vencimento + escopo de documentos, gravado só
        # quando o contrato termina com todos os documentos OK. Com ele o robô
        # nem pesquisa o contrato numa nova execução da mesma planilha.
        self.pular_contrato_concluido = os.getenv(
            "PULAR_CONTRATO_CONCLUIDO", "1"
        ).strip().lower() in {"1", "true", "sim", "s"}
        self.arquivo_contratos_concluidos = self.logs_dir / "contratos_concluidos_analiticos_ndi.txt"
        self.contratos_concluidos = self.carregar_contratos_concluidos()
        # --- FIM INSERÇÃO PULAR CONTRATO JÁ CONCLUÍDO ---
        # --- FIM INSERÇÃO CONTROLE TXT / NÃO ALTERAR EXCEL ORIGINAL ---

        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None

        # --- INÍCIO INSERÇÃO MODO RÁPIDO COM SESSÃO REUTILIZÁVEL ---
        # Ativado por padrão. Para voltar ao fluxo antigo, que reinicia o navegador
        # em cada contrato, configure a variável de ambiente MODO_RAPIDO=0.
        self.modo_rapido = os.getenv("MODO_RAPIDO", "1").strip().lower() in {"1", "true", "sim", "s"}
        self.tempos_contratos = []
        # --- FIM INSERÇÃO MODO RÁPIDO COM SESSÃO REUTILIZÁVEL ---

        # --- INÍCIO INSERÇÃO EXECUÇÃO PARALELA ---
        # WORKERS_PARALELOS=1 mantém o comportamento sequencial. Com 2 ou mais,
        # cada janela tem navegador e login próprios e retira contratos de uma
        # fila comum, de modo que nenhuma linha da planilha é processada duas vezes.
        self.quantidade_workers = max(1, int(os.getenv("WORKERS_PARALELOS", "1")))
        self.intervalo_inicio_workers = max(
            0.0,
            float(os.getenv("INTERVALO_INICIO_WORKERS_SEG", "5").replace(",", ".")),
        )
        self.modo_worker = False
        self.numero_worker = 0
        self.fila_linhas = None
        self.df_contratos_preparado = None
        self.progresso_paralelo = None
        self.evento_parada = None
        self.contratos_finalizados = 0
        self.indices_processados = []
        # --- FIM INSERÇÃO EXECUÇÃO PARALELA ---

        # --- INÍCIO INSERÇÃO CRONOMETRAGEM DETALHADA ---
        self.tempos_etapas = []
        self.inicio_execucao_geral = None
        self.total_contratos_planejados = 0
        self.arquivo_tempos_csv = self.logs_dir / "tempos_execucao_ndi.csv"
        self.arquivo_tempos_txt = self.logs_dir / "resumo_tempos_execucao_ndi.txt"
        self.inicializar_arquivo_tempos()
        # --- FIM INSERÇÃO CRONOMETRAGEM DETALHADA ---

        # --- INÍCIO INSERÇÃO OTIMIZAÇÃO DE TEMPO / CONFIGURAÇÕES ---
        # A Nota Fiscal continua com a espera conservadora anterior.
        self.timeout_evento_pdf = int(os.getenv("TIMEOUT_EVENTO_PDF_MS", "7000"))
        # O boleto do portal pode permanecer com o círculo de carregamento enquanto
        # o servidor gera o PDF. Durante esse período não devemos recarregar nem
        # clicar novamente. O robô monitora download, popup, Blob, URL e pasta local
        # durante até 120 segundos antes de considerar a tentativa malsucedida.
        self.timeout_geracao_boleto = int(
            os.getenv("TIMEOUT_GERACAO_BOLETO_MS", "120000")
        )
        self.intervalo_monitoramento_boleto = max(
            200,
            int(os.getenv("INTERVALO_MONITORAMENTO_BOLETO_MS", "500")),
        )
        self.intervalo_log_boleto = max(
            1.0,
            float(os.getenv("INTERVALO_LOG_BOLETO_SEG", "10").replace(",", ".")),
        )
        # Mantido como alias para compatibilidade com configurações e trechos antigos.
        self.timeout_evento_boleto = self.timeout_geracao_boleto
        self.timeout_botao_boleto_total = int(
            os.getenv("TIMEOUT_BOTAO_BOLETO_TOTAL_MS", "10000")
        )
        self.timeout_recarregar_fatura_boleto = int(
            os.getenv("TIMEOUT_RECARREGAR_FATURA_BOLETO_MS", "15000")
        )
        # A prefeitura às vezes abre uma página incompleta. A procura pelo botão
        # da NFS-e agora tem um limite total curto, em vez de vários timeouts somados.
        self.timeout_popup_nf = int(os.getenv("TIMEOUT_POPUP_NF_MS", "8000"))
        self.timeout_carregamento_nf = int(os.getenv("TIMEOUT_CARREGAMENTO_NF_MS", "8000"))
        self.timeout_botao_nf = int(os.getenv("TIMEOUT_BOTAO_NF_MS", "10000"))
        # A página da Prefeitura precisa terminar completamente antes do fluxo
        # confirmado manualmente: clique neutro + TAB 1x + ENTER.
        self.timeout_carregamento_completo_nf = int(
            os.getenv("TIMEOUT_CARREGAMENTO_COMPLETO_NF_MS", "20000")
        )
        self.timeout_download_nf_teclado = int(
            os.getenv("TIMEOUT_DOWNLOAD_NF_TECLADO_MS", "8000")
        )
        self.clique_neutro_nf_x = int(os.getenv("CLIQUE_NEUTRO_NF_X", "20"))
        self.clique_neutro_nf_y = int(os.getenv("CLIQUE_NEUTRO_NF_Y", "20"))
        self.intervalo_tab_nf = max(
            0.0,
            float(os.getenv("INTERVALO_TAB_NF_SEG", "0.20").replace(",", ".")),
        )
        self.timeout_respiro_rapido = int(os.getenv("TIMEOUT_RESPIRO_RAPIDO_MS", "3500"))
        self.max_paginas_historico = int(os.getenv("MAX_PAGINAS_HISTORICO", "20"))
        self.tentativas_por_contrato = max(1, int(os.getenv("TENTATIVAS_POR_CONTRATO", "2")))

        # --- INÍCIO INSERÇÃO RETENTATIVA INDIVIDUAL DE DOWNLOAD ---
        # Quantidade total de tentativas para cada documento. O padrão 2 significa:
        # 1 tentativa normal + 1 nova tentativa automática em caso de falha.
        self.tentativas_download_arquivo = max(
            1,
            int(os.getenv("TENTATIVAS_DOWNLOAD_ARQUIVO", "2")),
        )
        # Boleto: 1ª tentativa aguarda a geração por até 120 segundos. Somente
        # após o tempo limite a 2ª e última tentativa recarrega a mesma fatura.
        self.tentativas_download_boleto = max(
            1,
            int(os.getenv("TENTATIVAS_DOWNLOAD_BOLETO", "2")),
        )
        # Nota Fiscal mantém uma nova tentativa automática se necessário.
        self.tentativas_download_nf = max(
            1,
            int(os.getenv("TENTATIVAS_DOWNLOAD_NF", "2")),
        )
        self.intervalo_retentativa_download = max(
            0.0,
            float(os.getenv("INTERVALO_RETENTATIVA_DOWNLOAD_SEG", "1.5").replace(",", ".")),
        )
        self.recarregar_tela_retentativa_download = os.getenv(
            "RECARREGAR_TELA_RETENTATIVA_DOWNLOAD",
            "1",
        ).strip().lower() in {"1", "true", "sim", "s"}
        # --- FIM INSERÇÃO RETENTATIVA INDIVIDUAL DE DOWNLOAD ---
        # --- FIM INSERÇÃO OTIMIZAÇÃO DE TEMPO / CONFIGURAÇÕES ---


    # ----------------------------------------------------------------------
    # Log, status e utilitários
    # ----------------------------------------------------------------------
    # --- INÍCIO INSERÇÃO CRONOMETRAGEM DETALHADA ---
    def formatar_duracao(self, segundos):
        segundos = max(0.0, float(segundos or 0.0))
        horas = int(segundos // 3600)
        minutos = int((segundos % 3600) // 60)
        segundos_restantes = segundos % 60
        if horas:
            return f"{horas:02d}:{minutos:02d}:{segundos_restantes:05.2f}"
        return f"{minutos:02d}:{segundos_restantes:05.2f}"

    def inicializar_arquivo_tempos(self):
        try:
            if not self.arquivo_tempos_csv.exists():
                with TRAVA_ARQUIVOS, open(self.arquivo_tempos_csv, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f, delimiter=";")
                    writer.writerow([
                        "DATA_HORA", "CONTRATO", "VENCIMENTO", "TIPO_FATURA",
                        "ETAPA", "DURACAO_SEGUNDOS", "DURACAO_FORMATADA",
                        "STATUS", "DETALHES"
                    ])
        except Exception as e:
            print(f"[AVISO] Não foi possível inicializar o CSV de tempos: {e}")

    def registrar_tempo_etapa(self, etapa, segundos, status="OK", detalhes=""):
        registro = {
            "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "contrato": str(self.contrato_atual or ""),
            "vencimento": str(self.vencimento_atual or ""),
            "tipo_fatura": str(self.tipo_fatura_atual or ""),
            "etapa": str(etapa),
            "segundos": float(segundos),
            "duracao": self.formatar_duracao(segundos),
            "status": str(status),
            "detalhes": self.limpar_campo_txt(detalhes),
        }
        self.tempos_etapas.append(registro)
        try:
            self.inicializar_arquivo_tempos()
            with TRAVA_ARQUIVOS, open(self.arquivo_tempos_csv, "a", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow([
                    registro["data_hora"], registro["contrato"], registro["vencimento"],
                    registro["tipo_fatura"], registro["etapa"], f"{registro['segundos']:.4f}".replace(".", ","),
                    registro["duracao"], registro["status"], registro["detalhes"]
                ])
        except Exception as e:
            print(f"[AVISO] Não foi possível gravar tempo da etapa: {e}")

    @contextmanager
    def medir_etapa(self, etapa, detalhes=""):
        inicio = time.perf_counter()
        status = "OK"
        erro = ""
        print(f"[CRONÔMETRO] Início: {etapa}")
        try:
            yield
        except Exception as e:
            status = "ERRO"
            erro = self.resumir_erro(e)
            raise
        finally:
            duracao = time.perf_counter() - inicio
            detalhes_finais = detalhes
            if erro:
                detalhes_finais = f"{detalhes} | {erro}" if detalhes else erro
            self.registrar_tempo_etapa(etapa, duracao, status, detalhes_finais)
            print(f"[CRONÔMETRO] Fim: {etapa} | {self.formatar_duracao(duracao)} | {status}")

    def imprimir_projecao_execucao(self, contratos_concluidos):
        if not self.inicio_execucao_geral or contratos_concluidos <= 0:
            return
        decorrido = time.perf_counter() - self.inicio_execucao_geral
        media = decorrido / contratos_concluidos
        restantes = max(0, self.total_contratos_planejados - contratos_concluidos)
        estimado_restante = media * restantes
        estimado_total = decorrido + estimado_restante
        previsao_fim = datetime.fromtimestamp(time.time() + estimado_restante)
        print("\n[PROJEÇÃO DE TEMPO]")
        print(f"Concluídos              : {contratos_concluidos}/{self.total_contratos_planejados}")
        print(f"Tempo decorrido         : {self.formatar_duracao(decorrido)}")
        print(f"Média real por contrato : {self.formatar_duracao(media)}")
        print(f"Tempo restante estimado : {self.formatar_duracao(estimado_restante)}")
        print(f"Tempo total estimado    : {self.formatar_duracao(estimado_total)}")
        print(f"Previsão de término     : {previsao_fim.strftime('%d/%m/%Y %H:%M:%S')}")

    def gerar_resumo_tempos(self):
        if not self.tempos_etapas:
            return
        totais = {}
        contagens = {}
        for item in self.tempos_etapas:
            etapa = item["etapa"]
            totais[etapa] = totais.get(etapa, 0.0) + item["segundos"]
            contagens[etapa] = contagens.get(etapa, 0) + 1
        ranking = sorted(totais.items(), key=lambda x: x[1], reverse=True)
        linhas = [
            "=" * 100,
            "RESUMO DETALHADO DOS TEMPOS",
            "=" * 100,
        ]
        for etapa, total in ranking:
            media = total / contagens[etapa]
            linhas.append(
                f"{etapa:<45} | ocorrências: {contagens[etapa]:>4} | "
                f"total: {self.formatar_duracao(total):>11} | média: {self.formatar_duracao(media):>11}"
            )
        linhas.extend([
            "=" * 100,
            f"CSV detalhado: {self.arquivo_tempos_csv}",
            "=" * 100,
        ])
        resumo = "\n".join(linhas)
        print("\n" + resumo)
        try:
            with open(self.arquivo_tempos_txt, "w", encoding="utf-8") as f:
                f.write(resumo + "\n")
        except Exception as e:
            print(f"[AVISO] Não foi possível salvar o resumo TXT de tempos: {e}")
    # --- FIM INSERÇÃO CRONOMETRAGEM DETALHADA ---

    def ler_log(self):
        if not self.usar_log_downloads:
            return set()
        if not self.log_file.exists():
            return set()
        try:
            with open(self.log_file, "r", encoding="utf-8") as f:
                return {linha.strip() for linha in f if linha.strip()}
        except Exception:
            return set()

    def registrar_log(self, registro):
        if not self.usar_log_downloads:
            return
        if registro in self.downloads_concluidos:
            return
        self.downloads_concluidos.add(registro)
        with TRAVA_ARQUIVOS, open(self.log_file, "a", encoding="utf-8") as f:
            f.write(registro + "\n")

    def montar_chave_log_doc(self, contrato, vencimento, tipo_fatura, documento):
        return f"CONTRATO={contrato} | VENCIMENTO={vencimento} | TIPO_FATURA={tipo_fatura} | DOCUMENTO={documento}"

    def verificar_documento_baixado(self, contrato, vencimento, tipo_fatura, documento):
        # --- INÍCIO INSERÇÃO CONTROLE TXT / NÃO BAIXAR DUPLICADO ---
        chave_controle = self.montar_chave_controle_doc(contrato, vencimento, tipo_fatura, documento)
        if chave_controle in self.documentos_ok_controle:
            print(f"{documento} já consta como OK no controle TXT. Pulando download duplicado...")
            self.marcar_documento_fatura(documento, "JA_BAIXADO", "Documento já baixado anteriormente conforme controle TXT.")
            # --- INÍCIO INSERÇÃO CONTROLE TXT / DETALHE DOCUMENTO PULADO ---
            doc_normalizado = self.normalizar_documento_controle(documento)
            detalhe_doc = f"{doc_normalizado}_JA_BAIXADO"
            if detalhe_doc not in self.docs_linha_atual:
                self.docs_linha_atual.append(detalhe_doc)
            # --- FIM INSERÇÃO CONTROLE TXT / DETALHE DOCUMENTO PULADO ---
            return True
        # --- FIM INSERÇÃO CONTROLE TXT / NÃO BAIXAR DUPLICADO ---

        if not self.usar_log_downloads:
            return False
        chave = self.montar_chave_log_doc(contrato, vencimento, tipo_fatura, documento)
        if chave in self.downloads_concluidos:
            # --- INÍCIO INSERÇÃO CONTROLE TXT / NÃO BAIXAR DUPLICADO ---
            self.marcar_documento_fatura(documento, "JA_BAIXADO", "Documento já baixado anteriormente conforme log antigo.")
            doc_normalizado = self.normalizar_documento_controle(documento)
            detalhe_doc = f"{doc_normalizado}_JA_BAIXADO"
            if detalhe_doc not in self.docs_linha_atual:
                self.docs_linha_atual.append(detalhe_doc)
            # --- FIM INSERÇÃO CONTROLE TXT / NÃO BAIXAR DUPLICADO ---
            return True
        return False

    def registrar_documento_log(self, contrato, vencimento, tipo_fatura, documento):
        chave = self.montar_chave_log_doc(contrato, vencimento, tipo_fatura, documento)
        self.registrar_log(chave)
        # --- INÍCIO INSERÇÃO CONTROLE TXT / REGISTRO POR DOCUMENTO ---
        self.documentos_ok_controle.add(self.montar_chave_controle_doc(contrato, vencimento, tipo_fatura, documento))
        self.marcar_documento_fatura(documento, "OK")
        # --- FIM INSERÇÃO CONTROLE TXT / REGISTRO POR DOCUMENTO ---
        if documento not in self.docs_linha_atual:
            self.docs_linha_atual.append(documento)

    # --- INÍCIO INSERÇÃO CONTROLE TXT / FUNÇÕES AUXILIARES ---
    def normalizar_documento_controle(self, documento):
        doc = str(documento or "").strip().upper()
        if doc in {"NOTA FISCAL", "NOTA_FISCAL", "NFE", "NFS", "NFS-E"}:
            return "NF"
        if doc in {"BOLETO.PDF", "BOLETO PDF"}:
            return "BOLETO"
        return doc

    def montar_chave_controle_doc(self, contrato, vencimento, tipo_fatura, documento):
        doc = self.normalizar_documento_controle(documento)
        return (str(contrato).strip(), str(vencimento).strip(), str(tipo_fatura).strip(), doc)

    def inicializar_controle_txt(self):
        try:
            if not self.controle_txt_file.exists():
                with TRAVA_ARQUIVOS, open(self.controle_txt_file, "w", encoding="utf-8") as f:
                    f.write("|".join(self.controle_txt_cabecalho) + "\n")
                print(f"Controle TXT criado: {self.controle_txt_file}")
        except Exception as e:
            print(f"[AVISO] Não foi possível inicializar o controle TXT: {e}")

    def carregar_documentos_ok_controle(self):
        documentos_ok = set()
        try:
            if not self.controle_txt_file.exists():
                return documentos_ok

            with open(self.controle_txt_file, "r", encoding="utf-8") as f:
                linhas = [linha.rstrip("\n") for linha in f if linha.strip()]

            if not linhas:
                return documentos_ok

            inicio = 1 if linhas[0].startswith("Contrato|") else 0

            status_concluidos = {"OK", "JA_BAIXADO", "JÁ_BAIXADO", "JA BAIXADO", "JÁ BAIXADO"}
            total_colunas = len(self.controle_txt_cabecalho)

            for linha in linhas[inicio:]:
                partes = linha.split("|")
                if len(partes) < total_colunas:
                    continue

                contrato = partes[0].strip()
                vencimento = partes[1].strip()
                tipo_fatura = partes[2].strip()

                for posicao, (doc, _rotulo) in enumerate(self.COLUNAS_DOCUMENTOS_CONTROLE, start=3):
                    if partes[posicao].strip().upper() in status_concluidos:
                        documentos_ok.add(self.montar_chave_controle_doc(contrato, vencimento, tipo_fatura, doc))

        except Exception as e:
            print(f"[AVISO] Não foi possível carregar controle TXT para duplicidade: {e}")

        print(f"Documentos já baixados carregados do controle TXT: {len(documentos_ok)}")
        return documentos_ok

    # --- INÍCIO INSERÇÃO PULAR CONTRATO JÁ CONCLUÍDO ---
    def montar_chave_contrato_concluido(self, contrato, vencimento):
        """
        Contrato + vencimento + documentos do escopo atual. Se o escopo mudar
        (ex.: FORMATOS_ANALITICOS ou BAIXAR_BOLETO_NF), a chave muda e o
        contrato volta a ser processado.
        """
        escopo = ",".join(sorted(self.DOCUMENTOS_PERMITIDOS))
        return (
            self.limpar_campo_txt(contrato).upper(),
            self.limpar_campo_txt(vencimento),
            escopo,
        )

    def carregar_contratos_concluidos(self):
        concluidos = set()
        try:
            if self.arquivo_contratos_concluidos.exists():
                with open(self.arquivo_contratos_concluidos, "r", encoding="utf-8") as f:
                    for linha in f:
                        partes = linha.rstrip("\n").split("|")
                        if len(partes) < 3 or partes[0] == "Contrato":
                            continue
                        concluidos.add((partes[0].strip().upper(), partes[1].strip(), partes[2].strip()))
        except Exception as e:
            print(f"[AVISO] Não foi possível carregar os contratos concluídos: {e}")
        print(f"Contratos já concluídos carregados: {len(concluidos)}")
        return concluidos

    def contrato_ja_concluido(self, contrato, vencimento):
        if not self.pular_contrato_concluido:
            return False
        return self.montar_chave_contrato_concluido(contrato, vencimento) in self.contratos_concluidos

    def registrar_contrato_concluido(self, contrato, vencimento):
        chave = self.montar_chave_contrato_concluido(contrato, vencimento)
        if chave in self.contratos_concluidos:
            return
        self.contratos_concluidos.add(chave)
        try:
            with TRAVA_ARQUIVOS:
                novo = not self.arquivo_contratos_concluidos.exists()
                with open(self.arquivo_contratos_concluidos, "a", encoding="utf-8") as f:
                    if novo:
                        f.write("Contrato|Vencimento|Documentos|Data/hora\n")
                    f.write(
                        "|".join(list(chave) + [datetime.now().strftime("%d/%m/%Y %H:%M:%S")]) + "\n"
                    )
        except Exception as e:
            print(f"[AVISO] Não foi possível registrar o contrato concluído: {e}")
    # --- FIM INSERÇÃO PULAR CONTRATO JÁ CONCLUÍDO ---

    def limpar_campo_txt(self, valor):
        texto = str(valor or "").replace("\r", " ").replace("\n", " ").replace("|", "/")
        texto = re.sub(r"\s+", " ", texto).strip()
        return texto

    def obter_competencia_por_vencimento(self, vencimento):
        try:
            return datetime.strptime(str(vencimento), "%d/%m/%Y").strftime("%Y-%m")
        except Exception:
            try:
                return pd.to_datetime(vencimento, dayfirst=True).strftime("%Y-%m")
            except Exception:
                return "COMPETENCIA_NAO_IDENTIFICADA"

    def obter_pasta_contrato(self, contrato=None, vencimento=None):
        # Pasta única: não há mais subpastas por competência nem por contrato.
        # As evidências por contrato são gravadas em Logs/Evidencias.
        self.evidencias_dir.mkdir(parents=True, exist_ok=True)
        return self.evidencias_dir

    def definir_pasta_download_atual(self, contrato, vencimento, tipo_fatura):
        # Pasta única por tipo: Downloads/MENSALIDADE ou Downloads/COPARTICIPACAO.
        tipo_limpo = self.limpar_nome_arquivo(self.obter_sufixo_tipo_fatura(tipo_fatura))
        self.pasta_download_atual = self.downloads_dir / tipo_limpo
        self.pasta_download_atual.mkdir(parents=True, exist_ok=True)
        return self.pasta_download_atual

    def obter_pasta_download_atual(self):
        try:
            self.pasta_download_atual.mkdir(parents=True, exist_ok=True)
            return self.pasta_download_atual
        except Exception:
            self.downloads_dir.mkdir(parents=True, exist_ok=True)
            return self.downloads_dir

    def iniciar_estado_fatura(self, contrato, vencimento, tipo_fatura):
        self.definir_pasta_download_atual(contrato, vencimento, tipo_fatura)
        self.status_docs_fatura_atual = {
            doc: "PENDENTE" for doc in self.DOCUMENTOS_PERMITIDOS
        }
        # Não reaproveita diagnóstico de uma fatura anterior.
        self.falha_atual_documento = {}
        self.documento_download_atual = ""
        self.tentativa_download_atual = 0
        self.caminhos_docs_fatura_atual = {}
        self.mensagens_fatura_atual = []
        self.controle_fatura_registrado = False

    def marcar_documento_fatura(self, documento, status, mensagem="", caminho=None):
        doc = self.normalizar_documento_controle(documento)
        if doc not in self.DOCUMENTOS_PERMITIDOS:
            return

        status = str(status or "").strip().upper() or "PENDENTE"
        anterior = self.status_docs_fatura_atual.get(doc, "PENDENTE")

        if anterior in {"OK", "JA_BAIXADO"} and status not in {"OK"}:
            return

        self.status_docs_fatura_atual[doc] = status

        if caminho:
            self.caminhos_docs_fatura_atual[doc] = str(caminho)

        if mensagem:
            mensagem_limpa = self.limpar_campo_txt(f"{doc}: {mensagem}")
            if mensagem_limpa and mensagem_limpa not in self.mensagens_fatura_atual:
                self.mensagens_fatura_atual.append(mensagem_limpa)

    def obter_status_geral_fatura(self):
        valores = [
            self.status_docs_fatura_atual.get(doc, "PENDENTE")
            for doc in self.DOCUMENTOS_PERMITIDOS
        ]
        concluidos = {"OK", "JA_BAIXADO"}

        if valores and all(status in concluidos for status in valores):
            if all(status == "JA_BAIXADO" for status in valores):
                return "JA BAIXADO / SEM DUPLICAR"
            return "SUCESSO COMPLETO"

        if any(status in concluidos for status in valores):
            return "SUCESSO PARCIAL / PENDENCIA"

        if any(status == "ERRO" for status in valores):
            return "COM ERRO/PENDENCIA"
        if any(status == "NAO_DISPONIVEL" for status in valores):
            return "SEM DOCUMENTOS DISPONIVEIS"
        return "SEM DOWNLOAD"

    def classificar_resultado_contrato_atual(self):
        """Classifica o contrato considerando todos os documentos de todas as faturas."""
        status_documentos = []
        for evidencia in self.evidencias_contrato_atual:
            status_docs = evidencia.get("status_docs", {}) or {}
            for doc in self.DOCUMENTOS_PERMITIDOS:
                if doc in status_docs:
                    status_documentos.append(status_docs.get(doc, "PENDENTE"))

        concluidos = {"OK", "JA_BAIXADO"}
        if status_documentos and all(status in concluidos for status in status_documentos):
            return "COMPLETO"
        if any(status in concluidos for status in status_documentos):
            return "PARCIAL"
        return "SEM_DOWNLOAD"

    def montar_detalhes_resultado_contrato(self):
        detalhes = list(self.docs_linha_atual)
        pendencias = []
        for evidencia in self.evidencias_contrato_atual:
            status_docs = evidencia.get("status_docs", {}) or {}
            for doc in self.DOCUMENTOS_PERMITIDOS:
                status = status_docs.get(doc, "PENDENTE")
                if status not in {"OK", "JA_BAIXADO"}:
                    item = f"{doc}_{status}"
                    if item not in pendencias:
                        pendencias.append(item)

        for item in pendencias:
            if item not in detalhes:
                detalhes.append(item)
        return ", ".join(detalhes) if detalhes else "Sem documentos baixados"

    def registrar_linha_controle_fatura(self):
        if self.controle_fatura_registrado:
            return

        data_hora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        status_geral = self.obter_status_geral_fatura()
        mensagem = "; ".join(self.mensagens_fatura_atual) if self.mensagens_fatura_atual else ""

        linha = [
            self.limpar_campo_txt(self.contrato_atual),
            self.limpar_campo_txt(self.vencimento_atual),
            self.limpar_campo_txt(self.tipo_fatura_atual),
        ]
        for doc, _rotulo in self.COLUNAS_DOCUMENTOS_CONTROLE:
            if doc in self.DOCUMENTOS_PERMITIDOS:
                linha.append(self.status_docs_fatura_atual.get(doc, "PENDENTE"))
            else:
                # Documento fora do escopo desta execução.
                linha.append("-")
        linha.extend([
            status_geral,
            self.limpar_campo_txt(mensagem),
            data_hora,
        ])

        try:
            self.inicializar_controle_txt()
            with TRAVA_ARQUIVOS, open(self.controle_txt_file, "a", encoding="utf-8") as f:
                f.write("|".join(linha) + "\n")
            print(f"Linha adicionada ao controle TXT: {self.controle_txt_file.name}")
        except Exception as e:
            print(f"[AVISO] Não foi possível gravar linha no controle TXT: {e}")

        self.evidencias_contrato_atual.append({
            "contrato": str(self.contrato_atual),
            "vencimento": str(self.vencimento_atual),
            "tipo_fatura": str(self.tipo_fatura_atual),
            "status_docs": dict(self.status_docs_fatura_atual),
            "status_geral": status_geral,
            "mensagem": mensagem,
            "caminhos": dict(self.caminhos_docs_fatura_atual),
            "data_hora": data_hora,
        })

        self.resumo_execucao["faturas_processadas"] += 1
        for doc, status in self.status_docs_fatura_atual.items():
            if status == "OK":
                self.resumo_execucao["arquivos_baixados"] += 1
                if doc in self.resumo_execucao:
                    self.resumo_execucao[doc] += 1
            elif status == "JA_BAIXADO":
                self.resumo_execucao["arquivos_pulados"] += 1
            elif status == "ERRO":
                self.resumo_execucao["arquivos_com_erro"] += 1

        self.controle_fatura_registrado = True

    def adicionar_evidencia_erro_contrato(self, mensagem):
        self.evidencias_contrato_atual.append({
            "contrato": str(self.contrato_atual or ""),
            "vencimento": str(self.vencimento_atual or ""),
            "tipo_fatura": str(self.tipo_fatura_atual or "ERRO_CONTRATO"),
            "status_docs": dict(self.status_docs_fatura_atual or {}),
            "status_geral": "ERRO",
            "mensagem": self.limpar_campo_txt(mensagem),
            "caminhos": dict(self.caminhos_docs_fatura_atual or {}),
            "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        })

    def gerar_evidencia_contrato(self, status_contrato, detalhes_contrato=""):
        try:
            pasta_contrato = self.obter_pasta_contrato(self.contrato_atual, self.vencimento_atual)
            pasta_contrato.mkdir(parents=True, exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            contrato_limpo = self.limpar_nome_arquivo(str(self.contrato_atual or "SEM_CONTRATO"))
            caminho = pasta_contrato / f"EVIDENCIA_CONTRATO_{contrato_limpo}_{timestamp}.txt"

            linhas = []
            linhas.append("=======================================================")
            linhas.append("          EVIDENCIA AUTOMATICA DO CONTRATO")
            linhas.append("=======================================================")
            linhas.append(f"Contrato: {self.contrato_atual}")
            linhas.append(f"Vencimento: {self.vencimento_atual}")
            linhas.append(f"Status do contrato: {status_contrato}")
            linhas.append(f"Detalhes: {detalhes_contrato}")
            linhas.append(f"Data/hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
            linhas.append(f"Pasta da evidência: {pasta_contrato}")
            linhas.append("")

            if not self.evidencias_contrato_atual:
                linhas.append("Nenhuma fatura/documento foi registrado para este contrato.")
            else:
                for idx, ev in enumerate(self.evidencias_contrato_atual, 1):
                    linhas.append(f"--- Fatura {idx} ---")
                    linhas.append(f"Tipo da fatura: {ev.get('tipo_fatura', '')}")
                    linhas.append(f"Status geral: {ev.get('status_geral', '')}")
                    status_docs = ev.get("status_docs", {})
                    for doc, rotulo in self.COLUNAS_DOCUMENTOS_CONTROLE:
                        if doc in self.DOCUMENTOS_PERMITIDOS or doc in status_docs:
                            linhas.append(f"{rotulo}: {status_docs.get(doc, '-')}")
                    if ev.get("mensagem"):
                        linhas.append(f"Mensagem: {ev.get('mensagem')}")
                    caminhos = ev.get("caminhos", {})
                    if caminhos:
                        linhas.append("Arquivos salvos:")
                        for doc, caminho_doc in caminhos.items():
                            linhas.append(f"- {doc}: {caminho_doc}")
                    linhas.append(f"Data/hora da fatura: {ev.get('data_hora', '')}")
                    linhas.append("")

            with open(caminho, "w", encoding="utf-8") as f:
                f.write("\n".join(linhas))

            print(f"Evidência do contrato salva: {caminho}")
            return caminho
        except Exception as e:
            print(f"[AVISO] Não foi possível gerar evidência do contrato: {e}")
            return None

    def imprimir_resumo_final(self):
        escopo = "RELATÓRIOS ANALÍTICOS (" + ", ".join(self.formatos_analiticos) + ")"
        if self.baixar_boleto_nf:
            escopo += " + BOLETOS E NOTAS FISCAIS"
        print(f"\nMODO DE DOWNLOAD: {escopo}")
        print("\n=======================================================")
        print("                  RESUMO FINAL")
        print("=======================================================")
        print(f"Contratos processados: {self.resumo_execucao['contratos_processados']}")
        print(f"Contratos com sucesso completo: {self.resumo_execucao['contratos_sucesso']}")
        print(f"Contratos com sucesso parcial/pendência: {self.resumo_execucao['contratos_pendencia']}")
        print(f"Contratos com erro: {self.resumo_execucao['contratos_erro']}")
        print(f"Contratos pulados (já concluídos antes): {self.resumo_execucao.get('contratos_pulados_concluidos', 0)}")
        print(f"Faturas processadas: {self.resumo_execucao['faturas_processadas']}")
        print(f"Arquivos baixados nesta execução: {self.resumo_execucao['arquivos_baixados']}")
        print(f"Arquivos pulados por duplicidade: {self.resumo_execucao['arquivos_pulados']}")
        print(f"Arquivos com erro: {self.resumo_execucao['arquivos_com_erro']}")
        print(f"Novas tentativas de download: {self.resumo_execucao['retentativas_download']}")
        print(f"Arquivos recuperados em nova tentativa: {self.resumo_execucao['arquivos_recuperados_retentativa']}")
        print(f"Arquivos que falharam após todas as tentativas: {self.resumo_execucao['arquivos_falha_apos_retentativas']}")
        print(f"Falhas diagnosticadas: {self.resumo_execucao['falhas_diagnosticadas']}")
        print(f"Falhas recuperadas após diagnóstico: {self.resumo_execucao['falhas_recuperadas_apos_diagnostico']}")
        for doc, rotulo in self.COLUNAS_DOCUMENTOS_CONTROLE:
            if doc in self.DOCUMENTOS_PERMITIDOS:
                print(f"{rotulo} - arquivos baixados: {self.resumo_execucao.get(doc, 0)}")
        print(f"Arquivos extraídos de ZIP: {self.resumo_execucao['arquivos_extraidos_zip']}")
        print(f"Controle TXT: {self.controle_txt_file}")
        print(f"Rastreamento CSV: {self.arquivo_rastreamento_csv}")
        print(f"Diagnóstico CSV: {self.arquivo_diagnosticos_csv}")
        print(f"Evidências de diagnóstico: {self.diagnosticos_dir}")
        print(f"Pasta Downloads: {self.downloads_dir}")
        print("=======================================================")
    # --- FIM INSERÇÃO CONTROLE TXT / FUNÇÕES AUXILIARES ---

    # --- INÍCIO INSERÇÃO DIAGNÓSTICO ESTRUTURADO DE FALHAS ---
    def inicializar_arquivo_diagnosticos(self):
        try:
            if not self.arquivo_diagnosticos_csv.exists():
                with TRAVA_ARQUIVOS, open(self.arquivo_diagnosticos_csv, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f, delimiter=";")
                    writer.writerow([
                        "DATA_HORA", "CONTRATO", "VENCIMENTO", "TIPO_FATURA",
                        "DOCUMENTO", "TENTATIVA", "ETAPA", "CODIGO_FALHA",
                        "ORIGEM_PROVAVEL", "PONTO_CRUCIAL", "CAUSA_PROVAVEL",
                        "SINTOMA_PORTAL", "ACAO_SUGERIDA", "MENSAGEM_TECNICA",
                        "URL", "TITULO", "READY_STATE", "ABAS_ABERTAS",
                        "ELEMENTO_ATIVO", "SCREENSHOT", "ARQUIVO_JSON", "ARQUIVO_HTML"
                    ])
        except Exception as e:
            print(f"[AVISO] Não foi possível inicializar o CSV de diagnósticos: {e}")

    def capturar_estado_pagina_diagnostico(self, page=None):
        page = page or self.page
        estado = {
            "url": "",
            "titulo": "",
            "ready_state": "",
            "visibility_state": "",
            "elemento_ativo": {},
            "elementos_interativos": [],
            "texto_visivel": "",
            "abas_abertas": [],
        }
        if not page:
            return estado

        try:
            estado["url"] = str(page.url or "")
        except Exception:
            pass
        try:
            estado["titulo"] = str(page.title() or "")
        except Exception:
            pass
        try:
            dados_dom = page.evaluate(
                """
                () => {
                    const visivel = (el) => {
                        if (!el) return false;
                        const s = window.getComputedStyle(el);
                        const r = el.getBoundingClientRect();
                        return s && s.display !== 'none' && s.visibility !== 'hidden' &&
                               Number(s.opacity || 1) > 0 && r.width > 0 && r.height > 0;
                    };
                    const ativo = document.activeElement;
                    const descrever = (el) => {
                        if (!el) return {};
                        const texto = (el.innerText || el.textContent || el.value || '').trim();
                        return {
                            tag: (el.tagName || '').toLowerCase(),
                            id: el.id || '',
                            name: el.getAttribute ? (el.getAttribute('name') || '') : '',
                            role: el.getAttribute ? (el.getAttribute('role') || '') : '',
                            type: el.getAttribute ? (el.getAttribute('type') || '') : '',
                            placeholder: el.getAttribute ? (el.getAttribute('placeholder') || '') : '',
                            aria_label: el.getAttribute ? (el.getAttribute('aria-label') || '') : '',
                            texto: texto.slice(0, 180),
                            disabled: !!el.disabled
                        };
                    };
                    const interativos = Array.from(document.querySelectorAll(
                        'button, a, input, [role="button"]'
                    )).filter(visivel).slice(0, 30).map(descrever);
                    const corpo = document.body ? (document.body.innerText || '') : '';
                    return {
                        ready_state: document.readyState || '',
                        visibility_state: document.visibilityState || '',
                        elemento_ativo: descrever(ativo),
                        elementos_interativos: interativos,
                        texto_visivel: corpo.replace(/\\s+/g, ' ').trim().slice(0, 2500)
                    };
                }
                """
            )
            if isinstance(dados_dom, dict):
                estado.update(dados_dom)
        except Exception as e:
            estado["erro_captura_dom"] = self.resumir_erro(e)

        try:
            if self.context:
                for aba in list(self.context.pages):
                    try:
                        estado["abas_abertas"].append({
                            "url": str(aba.url or ""),
                            "titulo": str(aba.title() or ""),
                            "fechada": bool(aba.is_closed()),
                        })
                    except Exception:
                        pass
        except Exception:
            pass
        return estado

    def detectar_sintoma_portal(self, estado):
        texto = " ".join([
            str(estado.get("titulo", "")),
            str(estado.get("texto_visivel", "")),
        ]).lower()
        sintomas = [
            ("sessão expirada", "O portal informou que a sessão expirou."),
            ("sessao expirada", "O portal informou que a sessão expirou."),
            ("acesso negado", "O portal informou acesso negado."),
            ("unauthorized", "A página retornou indicação de acesso não autorizado."),
            ("forbidden", "A página retornou indicação de acesso proibido."),
            ("internal server error", "A página apresentou erro interno do servidor."),
            ("service unavailable", "O serviço externo está indisponível."),
            ("temporariamente indisponível", "O serviço informou indisponibilidade temporária."),
            ("temporariamente indisponivel", "O serviço informou indisponibilidade temporária."),
            ("tente novamente", "A página solicitou uma nova tentativa."),
            ("erro inesperado", "O portal exibiu uma mensagem de erro inesperado."),
            ("não foi possível", "A página informou que não foi possível concluir a operação."),
            ("nao foi possivel", "A página informou que não foi possível concluir a operação."),
            ("timeout", "A página apresentou indicação de timeout."),
            ("gateway", "A página apresentou erro de gateway/proxy."),
        ]
        for termo, descricao in sintomas:
            if termo in texto:
                return descricao
        return "Nenhuma mensagem explícita de erro foi encontrada no conteúdo visível da página."

    def classificar_falha_download(self, documento, etapa, mensagem, estado=None):
        doc = self.normalizar_documento_controle(documento)
        etapa_normalizada = str(etapa or "").upper()
        mensagem_normalizada = str(mensagem or "").lower()
        estado = estado or {}
        texto_pagina = str(estado.get("texto_visivel", "")).lower()

        if "session" in mensagem_normalizada or "sessão" in mensagem_normalizada or "sessao" in mensagem_normalizada or "sessão expirada" in texto_pagina or "sessao expirada" in texto_pagina:
            return {
                "codigo": "SESSAO_PORTAL_EXPIRADA",
                "origem": "PORTAL_HAPVIDA",
                "ponto_crucial": "Sessão autenticada antes da tentativa de download.",
                "causa": "A sessão do portal provavelmente expirou ou perdeu a autorização necessária.",
                "acao": "Reiniciar a sessão, refazer o login e reabrir o contrato antes de repetir o documento.",
            }

        if "GRAVA" in etapa_normalizada or "SALVAR" in etapa_normalizada or "permission" in mensagem_normalizada or "acesso negado" in mensagem_normalizada:
            return {
                "codigo": "ERRO_GRAVACAO_LOCAL",
                "origem": "SISTEMA_DE_ARQUIVOS",
                "ponto_crucial": "Gravação do PDF na pasta local.",
                "causa": "O conteúdo pode ter sido recebido, mas não foi possível salvar o arquivo no disco.",
                "acao": "Verificar permissão da pasta, arquivo aberto, caminho longo, espaço em disco e antivírus.",
            }

        if "TIMEOUT" in etapa_normalizada or "timeout" in mensagem_normalizada:
            return {
                "codigo": "TIMEOUT_REDE_OU_PORTAL",
                "origem": "REDE_OU_PORTAL",
                "ponto_crucial": "Espera pela resposta do portal após uma ação.",
                "causa": "O portal ou a conexão não respondeu dentro do limite configurado.",
                "acao": "Conferir conectividade, indisponibilidade do portal e aumentar apenas o timeout desta etapa se a página estiver carregando lentamente.",
            }

        if doc == "BOLETO":
            if "GERACAO_EXCEDEU_TEMPO_LIMITE" in etapa_normalizada or "GERAÇÃO_EXCEDEU_TEMPO_LIMITE" in etapa_normalizada:
                return {
                    "codigo": "BOLETO_GERACAO_EXCEDEU_TEMPO_LIMITE",
                    "origem": "PORTAL_HAPVIDA",
                    "ponto_crucial": "Processamento do boleto após o clique em Boleto.pdf.",
                    "causa": "O portal permaneceu processando o boleto além do limite configurado, sem disponibilizar download, popup, Blob, URL de PDF ou arquivo local válido.",
                    "acao": "Recarregar a tela da fatura somente após o tempo limite e executar uma última tentativa, evitando cliques repetidos enquanto o indicador estiver carregando.",
                }
            if "LOCALIZAR" in etapa_normalizada or "BOTAO" in etapa_normalizada or "BOTÃO" in etapa_normalizada:
                return {
                    "codigo": "BOLETO_BOTAO_NAO_RENDERIZADO",
                    "origem": "PORTAL_HAPVIDA_OU_LAYOUT",
                    "ponto_crucial": "Tela Detalhes da fatura, antes do clique em Boleto.pdf.",
                    "causa": "O botão do boleto não foi renderizado, mudou de posição/texto ou o documento não está disponível para a fatura.",
                    "acao": "Validar visualmente a tela, revisar o seletor e confirmar se o boleto aparece manualmente para o mesmo contrato e vencimento.",
                }
            if "RETENTATIVA" in etapa_normalizada or "REABRIR" in etapa_normalizada:
                return {
                    "codigo": "BOLETO_FATURA_NAO_REABERTA",
                    "origem": "PORTAL_HAPVIDA_OU_NAVEGACAO",
                    "ponto_crucial": "Retorno ao extrato e reentrada na fatura antes da nova tentativa.",
                    "causa": "A automação não conseguiu reconstruir o estado da fatura para repetir o boleto.",
                    "acao": "Conferir a aba/página/vencimento salvos e os marcadores usados para voltar ao extrato e reabrir a fatura.",
                }
            return {
                "codigo": "BOLETO_CLIQUE_SEM_GERAR_PDF",
                "origem": "PORTAL_HAPVIDA",
                "ponto_crucial": "Geração do PDF imediatamente após o clique em Boleto.pdf.",
                "causa": "O botão foi localizado e clicado, mas o portal não gerou download, popup, Blob válido nem mudança de URL.",
                "acao": "Tratar como falha do gerador de boleto do portal; reabrir a fatura, repetir o clique e guardar screenshot/estado da página para suporte do portal.",
            }

        if doc == "NF":
            if (
                "TECLADO" in etapa_normalizada
                or "TAB_1X" in etapa_normalizada
                or "CLIQUE_TAB" in etapa_normalizada
            ):
                elementos = estado.get("elementos_interativos", []) or []
                existe_download_visivel = any(
                    "download" in str(item.get("texto", "")).lower()
                    for item in elementos
                    if isinstance(item, dict)
                )
                if existe_download_visivel:
                    return {
                        "codigo": "NF_BOTAO_VISIVEL_MAS_CLIQUE_TAB_NAO_DISPAROU_DOWNLOAD",
                        "origem": "AUTOMACAO_DE_FOCO_OU_PORTAL_DA_PREFEITURA",
                        "ponto_crucial": "Foco após clique neutro + TAB 1x e acionamento por ENTER na página da NFS-e.",
                        "causa": "O botão Download NFS-e estava visível, mas clique neutro + TAB 1x + ENTER não gerou evento de download. O clique pode não ter estabelecido o foco esperado ou a página pode ter ficado temporariamente sem resposta.",
                        "acao": "Usar o elemento ativo registrado no JSON, repetir clique neutro + TAB 1x + ENTER e, como contingência, clicar diretamente no botão Download NFS-e visível.",
                    }
                return {
                    "codigo": "NF_CLIQUE_TAB_1X_ENTER_SEM_DISPARAR_DOWNLOAD",
                    "origem": "PORTAL_DA_PREFEITURA_OU_ORDEM_DE_FOCO",
                    "ponto_crucial": "Página da Prefeitura completamente carregada, imediatamente após clique neutro + TAB 1x + ENTER.",
                    "causa": "A sequência confirmada manualmente não disparou o download nesta tentativa; o ponto clicado, a ordem de foco ou a resposta da página podem ter variado.",
                    "acao": "Reabrir a NF, repetir clique neutro + TAB 1x + ENTER, conferir o elemento focado e usar o botão visível como contingência.",
                }
            if "LOCALIZAR_ACESSO" in etapa_normalizada or "ACESSAR_NOTA" in etapa_normalizada:
                return {
                    "codigo": "NF_BOTAO_ACESSO_NAO_RENDERIZADO",
                    "origem": "PORTAL_HAPVIDA_OU_LAYOUT",
                    "ponto_crucial": "Tela Detalhes da fatura, antes de abrir a nota fiscal.",
                    "causa": "O botão Acessar Nota Fiscal não apareceu, mudou de seletor ou a NF não está disponível.",
                    "acao": "Confirmar manualmente a disponibilidade da NF e revisar o seletor do botão na tela da fatura.",
                }
            if "ABRIR_PORTAL_NF" in etapa_normalizada or "NAO_ABRIU" in etapa_normalizada:
                return {
                    "codigo": "NF_PORTAL_EXTERNO_NAO_ABRIU",
                    "origem": "PORTAL_HAPVIDA_OU_PREFEITURA",
                    "ponto_crucial": "Redirecionamento do Hapvida para o portal externo da nota fiscal.",
                    "causa": "O clique não abriu popup e não alterou a URL, indicando falha no redirecionamento ou bloqueio da página externa.",
                    "acao": "Repetir a abertura, verificar bloqueio de popup e registrar a URL/estado para identificar se o Hapvida forneceu o link externo.",
                }
            if "LOCALIZAR_DOWNLOAD_NF" in etapa_normalizada or "BOTAO_DOWNLOAD" in etapa_normalizada:
                return {
                    "codigo": "NF_PREFEITURA_PAGINA_INCOMPLETA",
                    "origem": "PORTAL_DA_PREFEITURA",
                    "ponto_crucial": "Página externa da NFS-e, após o redirecionamento e antes do download.",
                    "causa": "A página externa abriu, mas não exibiu o botão de download/impressão dentro do tempo limite.",
                    "acao": "Fechar a página incompleta e abrir a NF novamente; usar título, URL, texto e screenshot para identificar erro da prefeitura.",
                }
            return {
                "codigo": "NF_CLIQUE_SEM_GERAR_PDF",
                "origem": "PORTAL_DA_PREFEITURA",
                "ponto_crucial": "Geração do PDF após o clique em download/impressão da NFS-e.",
                "causa": "O botão da NFS-e foi encontrado, mas nenhum arquivo PDF válido foi capturado ou salvo.",
                "acao": "Reabrir a NF e repetir o download; conferir se o portal mudou o mecanismo para impressão, popup ou Blob.",
            }

        # --- INÍCIO INSERÇÃO DIAGNÓSTICO DOS RELATÓRIOS ANALÍTICOS ---
        if doc.startswith("ANALITICO_"):
            formato = doc.replace("ANALITICO_", "")
            if "LOCALIZAR" in etapa_normalizada or "BOTAO" in etapa_normalizada or "BOTÃO" in etapa_normalizada:
                return {
                    "codigo": "ANALITICO_BOTAO_NAO_RENDERIZADO",
                    "origem": "PORTAL_HAPVIDA_OU_LAYOUT",
                    "ponto_crucial": f"Tela Detalhes da fatura, antes do clique no relatório analítico {formato}.",
                    "causa": (
                        f"O botão do relatório analítico {formato} (Relatório.{formato.lower()} ou "
                        "Baixar relatório) não foi renderizado, mudou de texto/posição ou o relatório "
                        "não está disponível para a fatura."
                    ),
                    "acao": "Conferir manualmente se o relatório aparece para o mesmo contrato e vencimento e revisar os seletores em obter_seletores_analitico.",
                }
            if "VALIDAR" in etapa_normalizada:
                return {
                    "codigo": "ANALITICO_ARQUIVO_INVALIDO",
                    "origem": "PORTAL_HAPVIDA_OU_ARMAZENAMENTO",
                    "ponto_crucial": f"Validação do conteúdo do relatório analítico {formato} depois de salvo.",
                    "causa": "O portal entregou um arquivo vazio, corrompido ou uma página de erro no lugar do relatório.",
                    "acao": "Repetir o download; se persistir, abrir o arquivo marcado como .invalido e acionar o suporte do portal.",
                }
            if "RETENTATIVA" in etapa_normalizada:
                return {
                    "codigo": "ANALITICO_FATURA_NAO_RECARREGADA",
                    "origem": "PORTAL_HAPVIDA_OU_NAVEGACAO",
                    "ponto_crucial": "Recarga da tela Detalhes da fatura antes da nova tentativa.",
                    "causa": "A tela da fatura não pôde ser confirmada depois do recarregamento.",
                    "acao": "Conferir a disponibilidade do portal e o marcador usado para confirmar a tela da fatura.",
                }
            return {
                "codigo": "ANALITICO_CLIQUE_SEM_GERAR_ARQUIVO",
                "origem": "PORTAL_HAPVIDA",
                "ponto_crucial": f"Entrega do arquivo logo após o clique no relatório analítico {formato}.",
                "causa": "O botão foi localizado e clicado, mas o portal não gerou download, nova aba com URL do arquivo nem mudança de URL.",
                "acao": "Repetir o clique após recarregar a fatura; guardar screenshot e estado da página para o suporte do portal.",
            }
        # --- FIM INSERÇÃO DIAGNÓSTICO DOS RELATÓRIOS ANALÍTICOS ---

        return {
            "codigo": "FALHA_NAO_CLASSIFICADA",
            "origem": "INDETERMINADA",
            "ponto_crucial": str(etapa or "Etapa não informada"),
            "causa": "A falha não correspondeu a uma regra conhecida de diagnóstico.",
            "acao": "Usar o screenshot, JSON, URL, elementos visíveis e mensagem técnica para criar uma regra específica.",
        }

    def salvar_screenshot_pagina_diagnostico(self, page, nome_base):
        try:
            if not page or page.is_closed():
                return None
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            contrato = self.limpar_nome_arquivo(str(self.contrato_atual or "SEM_CONTRATO"))
            nome = self.limpar_nome_arquivo(f"{timestamp}_{contrato}_{nome_base}") + ".png"
            caminho = self.diagnosticos_dir / nome
            page.screenshot(path=str(caminho), full_page=True)
            return caminho
        except Exception as e:
            print(f"[AVISO] Não foi possível salvar screenshot diagnóstico: {self.resumir_erro(e)}")
            return None

    def salvar_html_pagina_diagnostico(self, page, nome_base):
        if not self.salvar_html_diagnostico:
            return None
        try:
            if not page or page.is_closed():
                return None
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            contrato = self.limpar_nome_arquivo(str(self.contrato_atual or "SEM_CONTRATO"))
            nome = self.limpar_nome_arquivo(f"{timestamp}_{contrato}_{nome_base}") + ".html"
            caminho = self.diagnosticos_dir / nome
            caminho.write_text(page.content(), encoding="utf-8")
            return caminho
        except Exception as e:
            print(f"[AVISO] Não foi possível salvar HTML diagnóstico: {self.resumir_erro(e)}")
            return None

    def registrar_diagnostico_falha(self, documento, etapa, mensagem, page=None, excecao=None):
        doc = self.normalizar_documento_controle(documento or self.documento_download_atual or "DESCONHECIDO")
        tentativa = int(self.tentativa_download_atual or 0)
        pagina_alvo = page or self.page
        estado = self.capturar_estado_pagina_diagnostico(pagina_alvo)
        classificacao = self.classificar_falha_download(doc, etapa, mensagem, estado)
        sintoma = self.detectar_sintoma_portal(estado)
        mensagem_tecnica = self.resumir_erro(excecao if excecao is not None else mensagem)

        chave = (
            str(self.contrato_atual or ""), str(self.vencimento_atual or ""),
            str(self.tipo_fatura_atual or ""), doc, tentativa,
            str(etapa), classificacao["codigo"]
        )
        if chave in self.chaves_diagnostico_registradas:
            return self.falha_atual_documento.get(doc)
        self.chaves_diagnostico_registradas.add(chave)

        base_evidencia = self.limpar_nome_arquivo(
            f"{doc}_T{tentativa}_{classificacao['codigo']}"
        )
        screenshot = self.salvar_screenshot_pagina_diagnostico(pagina_alvo, base_evidencia)
        html = self.salvar_html_pagina_diagnostico(pagina_alvo, base_evidencia)

        diagnostico = {
            "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "contrato": str(self.contrato_atual or ""),
            "vencimento": str(self.vencimento_atual or ""),
            "tipo_fatura": str(self.tipo_fatura_atual or ""),
            "documento": doc,
            "tentativa": tentativa,
            "etapa": str(etapa),
            "codigo": classificacao["codigo"],
            "origem_provavel": classificacao["origem"],
            "ponto_crucial": classificacao["ponto_crucial"],
            "causa_provavel": classificacao["causa"],
            "sintoma_portal": sintoma,
            "acao_sugerida": classificacao["acao"],
            "mensagem_tecnica": mensagem_tecnica,
            "estado_pagina": estado,
            "screenshot": str(screenshot or ""),
            "html": str(html or ""),
            "recuperado": False,
        }

        timestamp_arquivo = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        contrato_limpo = self.limpar_nome_arquivo(str(self.contrato_atual or "SEM_CONTRATO"))
        nome_json = self.limpar_nome_arquivo(
            f"{timestamp_arquivo}_{contrato_limpo}_{base_evidencia}"
        ) + ".json"
        caminho_json = self.diagnosticos_dir / nome_json
        diagnostico["arquivo_json"] = str(caminho_json)
        try:
            caminho_json.write_text(
                json.dumps(diagnostico, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception as e:
            print(f"[AVISO] Não foi possível salvar JSON diagnóstico: {self.resumir_erro(e)}")
            diagnostico["arquivo_json"] = ""

        self.diagnosticos_execucao.append(diagnostico)
        self.falha_atual_documento[doc] = diagnostico
        self.contagem_codigos_falha[diagnostico["codigo"]] = (
            self.contagem_codigos_falha.get(diagnostico["codigo"], 0) + 1
        )
        self.resumo_execucao["falhas_diagnosticadas"] += 1

        try:
            self.inicializar_arquivo_diagnosticos()
            with TRAVA_ARQUIVOS, open(self.arquivo_diagnosticos_csv, "a", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";")
                ativo = estado.get("elemento_ativo", {}) or {}
                writer.writerow([
                    diagnostico["data_hora"], diagnostico["contrato"],
                    diagnostico["vencimento"], diagnostico["tipo_fatura"],
                    diagnostico["documento"], diagnostico["tentativa"],
                    diagnostico["etapa"], diagnostico["codigo"],
                    diagnostico["origem_provavel"], diagnostico["ponto_crucial"],
                    diagnostico["causa_provavel"], diagnostico["sintoma_portal"],
                    diagnostico["acao_sugerida"], diagnostico["mensagem_tecnica"],
                    estado.get("url", ""), estado.get("titulo", ""),
                    estado.get("ready_state", ""), len(estado.get("abas_abertas", [])),
                    json.dumps(ativo, ensure_ascii=False), diagnostico["screenshot"],
                    diagnostico["arquivo_json"], diagnostico["html"],
                ])
        except Exception as e:
            print(f"[AVISO] Não foi possível gravar CSV diagnóstico: {self.resumir_erro(e)}")

        print("\n[DIAGNÓSTICO DA FALHA]")
        print(f"Código             : {diagnostico['codigo']}")
        print(f"Origem provável    : {diagnostico['origem_provavel']}")
        print(f"Ponto crucial      : {diagnostico['ponto_crucial']}")
        print(f"Causa provável     : {diagnostico['causa_provavel']}")
        print(f"Sintoma identificado: {diagnostico['sintoma_portal']}")
        print(f"Ação sugerida      : {diagnostico['acao_sugerida']}")
        print(f"Mensagem técnica   : {diagnostico['mensagem_tecnica']}")
        if diagnostico["screenshot"]:
            print(f"Screenshot         : {diagnostico['screenshot']}")
        if diagnostico["arquivo_json"]:
            print(f"Detalhes técnicos  : {diagnostico['arquivo_json']}")
        return diagnostico

    def registrar_diagnostico_se_ausente(self, documento, etapa, mensagem, page=None, excecao=None):
        doc = self.normalizar_documento_controle(documento)
        atual = self.falha_atual_documento.get(doc)
        if atual and int(atual.get("tentativa", 0) or 0) == int(self.tentativa_download_atual or 0):
            return atual
        return self.registrar_diagnostico_falha(doc, etapa, mensagem, page=page, excecao=excecao)

    def registrar_recuperacao_apos_diagnostico(self, documento, tentativa):
        doc = self.normalizar_documento_controle(documento)
        diagnostico = self.falha_atual_documento.get(doc)
        if not diagnostico:
            return
        diagnostico["recuperado"] = True
        diagnostico["tentativa_recuperacao"] = int(tentativa)
        self.resumo_execucao["falhas_recuperadas_apos_diagnostico"] += 1
        try:
            caminho_json = diagnostico.get("arquivo_json")
            if caminho_json:
                Path(caminho_json).write_text(
                    json.dumps(diagnostico, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
        except Exception:
            pass
        print(
            f"[DIAGNÓSTICO] A falha {diagnostico.get('codigo')} foi recuperada "
            f"na tentativa {tentativa}."
        )

    def gerar_resumo_diagnosticos(self):
        linhas = [
            "=" * 100,
            "RESUMO DOS DIAGNÓSTICOS DE FALHA",
            "=" * 100,
            f"Eventos diagnosticados: {len(self.diagnosticos_execucao)}",
            f"Falhas recuperadas depois do diagnóstico: {self.resumo_execucao.get('falhas_recuperadas_apos_diagnostico', 0)}",
            "",
        ]
        if not self.diagnosticos_execucao:
            linhas.append("Nenhuma falha de download foi diagnosticada nesta execução.")
        else:
            linhas.append("Ocorrências por código:")
            for codigo, quantidade in sorted(
                self.contagem_codigos_falha.items(),
                key=lambda item: item[1],
                reverse=True,
            ):
                linhas.append(f"- {codigo}: {quantidade}")
            linhas.append("")
            linhas.append("Eventos:")
            for item in self.diagnosticos_execucao:
                recuperado = "SIM" if item.get("recuperado") else "NÃO"
                linhas.append(
                    f"- {item.get('contrato')} | {item.get('documento')} | "
                    f"tentativa {item.get('tentativa')} | {item.get('codigo')} | "
                    f"origem: {item.get('origem_provavel')} | recuperado: {recuperado}"
                )
                linhas.append(f"  Ponto crucial: {item.get('ponto_crucial')}")
                linhas.append(f"  Causa provável: {item.get('causa_provavel')}")
                if item.get("screenshot"):
                    linhas.append(f"  Screenshot: {item.get('screenshot')}")
                if item.get("arquivo_json"):
                    linhas.append(f"  JSON: {item.get('arquivo_json')}")
        linhas.extend([
            "",
            f"CSV detalhado: {self.arquivo_diagnosticos_csv}",
            f"Pasta de evidências: {self.diagnosticos_dir}",
            "=" * 100,
        ])
        resumo = "\n".join(linhas)
        try:
            self.arquivo_resumo_diagnosticos.write_text(resumo + "\n", encoding="utf-8")
        except Exception as e:
            print(f"[AVISO] Não foi possível salvar resumo diagnóstico: {self.resumir_erro(e)}")
        print("\n" + resumo)
        return resumo
    # --- FIM INSERÇÃO DIAGNÓSTICO ESTRUTURADO DE FALHAS ---

    # --- INÍCIO INSERÇÃO RASTREAMENTO DE DOWNLOADS (MODELO DO GRAVADOR) ---
    def inicializar_arquivo_rastreamento(self):
        try:
            if not self.arquivo_rastreamento_csv.exists():
                with TRAVA_ARQUIVOS, open(self.arquivo_rastreamento_csv, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f, delimiter=";")
                    writer.writerow([
                        "DATA_HORA", "CONTRATO", "VENCIMENTO", "TIPO_FATURA",
                        "DOCUMENTO", "TENTATIVA", "STATUS", "METODO", "ARQUIVO",
                        "BYTES", "SHA256", "VALIDACAO", "DURACAO_SEGUNDOS",
                        "URL_ORIGEM", "MEMBROS_ZIP", "ARQUIVOS_EXTRAIDOS",
                    ])
        except Exception as e:
            print(f"[AVISO] Não foi possível inicializar o CSV de rastreamento: {e}")

    def calcular_sha256_arquivo(self, caminho):
        try:
            sha = hashlib.sha256()
            with open(caminho, "rb") as arquivo:
                for bloco in iter(lambda: arquivo.read(1024 * 1024), b""):
                    sha.update(bloco)
            return sha.hexdigest()
        except Exception:
            return ""

    def limpar_url_para_log(self, url):
        """Remove query string e fragmento: URLs de armazenamento podem conter token de acesso."""
        try:
            partes = urlparse(str(url or ""))
            if not partes.scheme:
                return ""
            if partes.scheme == "blob":
                return "blob:"
            return f"{partes.scheme}://{partes.netloc}{partes.path}"
        except Exception:
            return ""

    def registrar_rastreamento_download(
        self,
        documento,
        status,
        metodo="",
        caminho=None,
        validacao="",
        duracao=0.0,
        url="",
        membros_zip=None,
        extraidos=None,
    ):
        """Grava uma linha por arquivo capturado, no mesmo espírito do downloads.csv do gravador."""
        tamanho = ""
        sha256 = ""
        nome_arquivo = ""
        try:
            if caminho and Path(caminho).exists():
                nome_arquivo = Path(caminho).name
                tamanho = Path(caminho).stat().st_size
                sha256 = self.calcular_sha256_arquivo(caminho)
        except Exception:
            pass

        try:
            self.inicializar_arquivo_rastreamento()
            with TRAVA_ARQUIVOS, open(self.arquivo_rastreamento_csv, "a", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f, delimiter=";")
                writer.writerow([
                    datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                    str(self.contrato_atual or ""),
                    str(self.vencimento_atual or ""),
                    str(self.tipo_fatura_atual or ""),
                    self.normalizar_documento_controle(documento),
                    int(self.tentativa_download_atual or 0),
                    str(status),
                    str(metodo),
                    nome_arquivo,
                    tamanho,
                    sha256,
                    str(validacao),
                    f"{float(duracao or 0.0):.3f}".replace(".", ","),
                    self.limpar_url_para_log(url),
                    json.dumps(membros_zip or [], ensure_ascii=False),
                    json.dumps([Path(p).name for p in (extraidos or [])], ensure_ascii=False),
                ])
        except Exception as e:
            print(f"[AVISO] Não foi possível gravar o rastreamento do download: {self.resumir_erro(e)}")
    # --- FIM INSERÇÃO RASTREAMENTO DE DOWNLOADS (MODELO DO GRAVADOR) ---

    def registrar_erro(self, contexto, erro_msg):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        msg = self.resumir_erro(erro_msg)
        with TRAVA_ARQUIVOS, open(self.erros_log_file, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] CONTRATO: {self.contrato_atual} | {contexto} | ERRO: {msg}\n")

    def resumir_erro(self, erro):
        texto = str(erro).strip()
        if not texto:
            return erro.__class__.__name__ if hasattr(erro, "__class__") else "Erro"
        linhas = [linha.strip() for linha in texto.splitlines() if linha.strip()]
        if not linhas:
            return texto[:250]
        primeira = linhas[0]
        if primeira.lower().startswith("message:") and len(linhas) > 1:
            primeira = linhas[1]
        return primeira[:350]

    def salvar_screenshot(self, nome_base, page=None):
        """Salva screenshot da página informada; por padrão usa a página principal."""
        try:
            pagina_alvo = page or self.page
            if not pagina_alvo or pagina_alvo.is_closed():
                return None
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            contrato = self.limpar_nome_arquivo(str(self.contrato_atual or "SEM_CONTRATO"))
            nome = self.limpar_nome_arquivo(f"{timestamp}_{contrato}_{nome_base}") + ".png"
            caminho = self.screenshots_dir / nome
            pagina_alvo.screenshot(path=str(caminho), full_page=True)
            print(f"Screenshot salvo: {caminho}")
            return caminho
        except Exception as e:
            print(f"[AVISO] Falha ao salvar screenshot: {e}")
            return None

    def limpar_nome_arquivo(self, nome):
        nome = str(nome).strip()
        nome = re.sub(r"[\\/:*?\"<>|]+", "_", nome)
        nome = re.sub(r"\s+", "_", nome)
        nome = nome.replace("__", "_")
        return nome.strip("._ ") or "arquivo"

    # --- INÍCIO INSERÇÃO NOME ORIGINAL DO ARQUIVO ---
    def limpar_nome_original(self, nome):
        """
        Mantém o nome entregue pelo portal. Remove apenas o caminho e os
        caracteres que o Windows não aceita em nomes de arquivo.
        Retorna "" quando não há nome aproveitável.
        """
        nome = str(nome or "").replace("\\", "/").split("/")[-1].strip()
        nome = re.sub(r"[:*?\"<>|\x00-\x1f]+", "_", nome)
        nome = nome.strip(". ")
        if not Path(nome).stem:
            return ""
        return nome

    def obter_nome_original_resposta(self, url="", headers=None):
        """
        Nome original de um arquivo obtido por URL: primeiro o cabeçalho
        Content-Disposition da resposta; depois o nome no fim da URL.
        """
        headers = {str(k).lower(): str(v) for k, v in (headers or {}).items()}
        disposicao = headers.get("content-disposition", "")
        if disposicao:
            # filename*=UTF-8''nome%20codificado.csv
            m = re.search(r"filename\*\s*=\s*(?:[\w-]+)?'[^']*'([^;]+)", disposicao, re.IGNORECASE)
            if m:
                nome = self.limpar_nome_original(unquote(m.group(1).strip().strip('"')))
                if nome:
                    return nome
            m = re.search(r"filename\s*=\s*\"?([^\";]+)\"?", disposicao, re.IGNORECASE)
            if m:
                nome = self.limpar_nome_original(unquote(m.group(1).strip()))
                if nome:
                    return nome

        try:
            caminho_url = unquote(urlparse(str(url or "")).path)
        except Exception:
            caminho_url = ""
        nome = self.limpar_nome_original(caminho_url)
        # Só aceita o fim da URL quando é realmente um arquivo (evita .aspx, .ashx etc.).
        if nome and Path(nome).suffix.lower() in {".csv", ".txt", ".pdf", ".zip"}:
            return nome
        return ""

    def montar_caminho_destino(self, pasta_destino, nome_original, novo_nome_base, ext):
        """
        Usa o nome original quando existir; o nome padrão do robô fica apenas
        como contingência (ex.: PDF capturado por Blob, que não tem nome).
        O contador _1, _2... só é usado para nunca sobrescrever um arquivo.
        """
        nome_original = self.limpar_nome_original(nome_original)
        if nome_original:
            base = Path(nome_original).stem
            ext = Path(nome_original).suffix or ext
        else:
            base = self.limpar_nome_arquivo(novo_nome_base)

        caminho = pasta_destino / f"{base}{ext}"
        contador = 1
        while caminho.exists():
            caminho = pasta_destino / f"{base}_{contador}{ext}"
            contador += 1
        return caminho
    # --- FIM INSERÇÃO NOME ORIGINAL DO ARQUIVO ---

    def formatar_vencimento_nome_arquivo(self, vencimento_alvo):
        try:
            return datetime.strptime(str(vencimento_alvo), "%d/%m/%Y").strftime("%d-%m-%y")
        except Exception:
            try:
                return pd.to_datetime(vencimento_alvo, dayfirst=True).strftime("%d-%m-%y")
            except Exception:
                return str(vencimento_alvo).replace("/", "-")

    def obter_sufixo_tipo_fatura(self, tipo_fatura):
        if tipo_fatura == self.TIPO_MENSALIDADE:
            return "MENSALIDADE"
        if tipo_fatura == self.TIPO_CARNET:
            return "COPARTICIPACAO"
        return "OUTRO"

    def salvar_download(self, download, novo_nome_base, extensao_fallback=""):
        suggested = download.suggested_filename or ""
        ext = Path(suggested).suffix
        if not ext and extensao_fallback:
            ext = extensao_fallback if extensao_fallback.startswith(".") else f".{extensao_fallback}"
        if not ext:
            ext = ".bin"

        # --- INÍCIO INSERÇÃO PASTA ÚNICA POR TIPO / NOME ORIGINAL ---
        pasta_destino = self.obter_pasta_download_atual()
        caminho = self.montar_caminho_destino(pasta_destino, suggested, novo_nome_base, ext)
        # --- FIM INSERÇÃO PASTA ÚNICA POR TIPO / NOME ORIGINAL ---

        download.save_as(str(caminho))
        print(f"Arquivo salvo: {caminho.name}")
        return caminho

    def obter_numero_nf_url(self, page):
        try:
            parsed = urlparse(page.url)
            params = parse_qs(parsed.query)
            numero_nf = params.get("nf", [None])[0]
            if numero_nf:
                print(f"Número da NF identificado pela URL: {numero_nf}")
                return str(numero_nf).strip()
        except Exception:
            pass

        try:
            texto = page.locator("body").inner_text(timeout=self.TIMEOUT_CURTO)
            numeros = re.findall(r"\b\d{6,}\b", texto)
            if numeros:
                print(f"Número da NF identificado por fallback no texto: {numeros[0]}")
                return numeros[0]
        except Exception:
            pass

        return "NUMERO_NF_NAO_IDENTIFICADO"

    # ----------------------------------------------------------------------
    # Navegador e esperas Playwright
    # ----------------------------------------------------------------------
    # --- INÍCIO INSERÇÃO JANELA MAXIMIZADA / SUPORTE A DOIS MONITORES ---
    def obter_monitores_windows(self):
        """
        Retorna as áreas úteis dos monitores do Windows em pixels reais.

        A ordem é:
        1. monitor principal do Windows;
        2. demais monitores, ordenados pela posição na área de trabalho.

        Em outro sistema operacional, ou se a enumeração falhar, retorna uma
        área de fallback.
        """
        fallback = [{
            "left": 0,
            "top": 0,
            "width": 1920,
            "height": 1080,
            "primary": True,
            "device": "MONITOR_FALLBACK",
        }]

        if os.name != "nt":
            return fallback

        try:
            from ctypes import wintypes

            try:
                # Evita que o Windows devolva medidas reduzidas quando a escala
                # da tela estiver em 125%, 150% etc.
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

            class RECT(ctypes.Structure):
                _fields_ = [
                    ("left", ctypes.c_long),
                    ("top", ctypes.c_long),
                    ("right", ctypes.c_long),
                    ("bottom", ctypes.c_long),
                ]

            class MONITORINFOEXW(ctypes.Structure):
                _fields_ = [
                    ("cbSize", wintypes.DWORD),
                    ("rcMonitor", RECT),
                    ("rcWork", RECT),
                    ("dwFlags", wintypes.DWORD),
                    ("szDevice", wintypes.WCHAR * 32),
                ]

            monitores = []
            callback_type = ctypes.WINFUNCTYPE(
                ctypes.c_int,
                ctypes.c_void_p,
                ctypes.c_void_p,
                ctypes.POINTER(RECT),
                ctypes.c_void_p,
            )

            def callback_monitor(hmonitor, hdc, lprect, lparam):
                info = MONITORINFOEXW()
                info.cbSize = ctypes.sizeof(MONITORINFOEXW)
                if ctypes.windll.user32.GetMonitorInfoW(hmonitor, ctypes.byref(info)):
                    area = info.rcWork
                    monitores.append({
                        "left": int(area.left),
                        "top": int(area.top),
                        "width": int(area.right - area.left),
                        "height": int(area.bottom - area.top),
                        "primary": bool(info.dwFlags & 1),
                        "device": str(info.szDevice),
                    })
                return 1

            callback_ref = callback_type(callback_monitor)
            ctypes.windll.user32.EnumDisplayMonitors(
                None,
                None,
                callback_ref,
                None,
            )

            if not monitores:
                return fallback

            principais = [m for m in monitores if m["primary"]]
            secundarios = sorted(
                [m for m in monitores if not m["primary"]],
                key=lambda m: (m["left"], m["top"]),
            )
            return principais + secundarios
        except Exception as e:
            print(f"[AVISO] Não foi possível identificar os monitores do Windows: {e}")
            return fallback

    def selecionar_monitor_navegador(self):
        monitores = self.obter_monitores_windows()
        try:
            numero_tela = int(os.getenv("NAVEGADOR_TELA", "1").strip())
        except Exception:
            numero_tela = 1

        numero_tela = max(1, numero_tela)
        indice = min(numero_tela - 1, len(monitores) - 1)
        monitor = dict(monitores[indice])
        monitor["numero"] = indice + 1
        monitor["total"] = len(monitores)
        return monitor

    def maximizar_janela_navegador(self, page=None, monitor=None):
        """
        Posiciona a janela no monitor escolhido e maximiza a janela real do
        Chromium. O viewport da página acompanha a janela porque o contexto é
        criado com no_viewport=True.
        """
        page = page or self.page
        if not page or not self.context:
            return False

        headless = os.getenv("PLAYWRIGHT_HEADLESS", "0").strip().lower() in {
            "1", "true", "sim", "s"
        }
        if headless:
            return False

        maximizar = os.getenv("NAVEGADOR_MAXIMIZADO", "1").strip().lower() in {
            "1", "true", "sim", "s"
        }
        monitor = monitor or self.selecionar_monitor_navegador()

        ultimo_erro = None
        for tentativa in range(1, 4):
            sessao_cdp = None
            try:
                sessao_cdp = self.context.new_cdp_session(page)
                resposta = sessao_cdp.send("Browser.getWindowForTarget")
                window_id = resposta["windowId"]

                # Primeiro normaliza e move a janela para dentro do monitor
                # selecionado. Isso é necessário para que o comando maximizar
                # seja aplicado na tela correta.
                margem = 20
                largura = max(800, int(monitor["width"]) - (margem * 2))
                altura = max(600, int(monitor["height"]) - (margem * 2))
                sessao_cdp.send(
                    "Browser.setWindowBounds",
                    {
                        "windowId": window_id,
                        "bounds": {
                            "windowState": "normal",
                            "left": int(monitor["left"]) + margem,
                            "top": int(monitor["top"]) + margem,
                            "width": largura,
                            "height": altura,
                        },
                    },
                )
                time.sleep(0.20)

                if maximizar:
                    sessao_cdp.send(
                        "Browser.setWindowBounds",
                        {
                            "windowId": window_id,
                            "bounds": {"windowState": "maximized"},
                        },
                    )

                print(
                    f"Navegador aberto na tela {monitor['numero']}/{monitor['total']} "
                    f"({monitor['device']}) e "
                    f"{'maximizado' if maximizar else 'redimensionado'}."
                )
                return True
            except Exception as e:
                ultimo_erro = e
                if tentativa < 3:
                    time.sleep(0.50)
            finally:
                if sessao_cdp:
                    try:
                        sessao_cdp.detach()
                    except Exception:
                        pass

        print(f"[AVISO] Não foi possível maximizar a janela do navegador: {ultimo_erro}")
        return False
    # --- FIM INSERÇÃO JANELA MAXIMIZADA / SUPORTE A DOIS MONITORES ---

    def iniciar_navegador(self):
        print("Iniciando navegador Playwright...")
        headless = os.getenv("PLAYWRIGHT_HEADLESS", "0").strip().lower() in {"1", "true", "sim", "s"}
        usar_chrome_instalado = os.getenv("PLAYWRIGHT_USAR_CHROME", "0").strip().lower() in {"1", "true", "sim", "s"}

        self.playwright = sync_playwright().start()
        monitor = self.selecionar_monitor_navegador()

        argumentos_chromium = [
            "--start-maximized",
            "--incognito",
            "--disable-gpu",
            "--disable-pdf-extension",
            "--disable-features=PdfOopif",
        ]

        # Já solicita ao Chromium que nasça no monitor escolhido. Depois da
        # criação da página, o CDP confirma a posição e maximiza de verdade.
        if not headless:
            argumentos_chromium.extend([
                f"--window-position={int(monitor['left'])},{int(monitor['top'])}",
                f"--window-size={int(monitor['width'])},{int(monitor['height'])}",
            ])

        launch_kwargs = {
            "headless": headless,
            "args": argumentos_chromium,
        }
        if usar_chrome_instalado:
            launch_kwargs["channel"] = "chrome"

        self.browser = self.playwright.chromium.launch(**launch_kwargs)
        self.context = self.browser.new_context(
            accept_downloads=True,
            # IMPORTANTE: não usar viewport fixo. A página acompanha o tamanho
            # real da janela maximizada do Chromium.
            no_viewport=True,
            locale="pt-BR",
        )
        self.context.set_default_timeout(self.TIMEOUT_PADRAO)
        self.context.set_default_navigation_timeout(self.TIMEOUT_LONGO)
        self.page = self.context.new_page()
        self.page.set_default_timeout(self.TIMEOUT_PADRAO)
        self.page.set_default_navigation_timeout(self.TIMEOUT_LONGO)

        self.maximizar_janela_navegador(self.page, monitor)

    def fechar_navegador(self):
        try:
            if self.context:
                self.context.close()
        except Exception:
            pass
        try:
            if self.browser:
                self.browser.close()
        except Exception:
            pass
        try:
            if self.playwright:
                self.playwright.stop()
        except Exception:
            pass

        # --- INÍCIO INSERÇÃO REINÍCIO DO NAVEGADOR POR CONTRATO ---
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        self.url_painel = None
        # --- FIM INSERÇÃO REINÍCIO DO NAVEGADOR POR CONTRATO ---

    def respirar_sistema(self, page=None, timeout=12_000, segundos_fallback=1.5):
        """
        Respiro padrão depois de telas pesadas.
        Usa networkidle, mas sem travar o robô caso o portal mantenha requests ativos.
        """
        page = page or self.page
        # --- INÍCIO INSERÇÃO OTIMIZAÇÃO DE TEMPO / NETWORKIDLE ---
        if self.modo_rapido:
            timeout = min(timeout, self.timeout_respiro_rapido)
            segundos_fallback = min(segundos_fallback, 0.35)
        # --- FIM INSERÇÃO OTIMIZAÇÃO DE TEMPO / NETWORKIDLE ---
        try:
            page.wait_for_load_state("domcontentloaded", timeout=timeout)
        except Exception:
            pass
        try:
            page.wait_for_load_state("networkidle", timeout=timeout)
        except Exception:
            time.sleep(segundos_fallback)

    def aguardar_selector(self, selector, descricao="elemento", page=None, timeout=None, state="visible"):
        page = page or self.page
        timeout = timeout or self.TIMEOUT_PADRAO
        print(f"Aguardando {descricao}...")
        return page.wait_for_selector(selector, state=state, timeout=timeout)

    def localizar_primeiro(self, descricao, seletores, page=None, timeout_por_selector=None, somente_visivel=True):
        page = page or self.page
        timeout_por_selector = timeout_por_selector or self.TIMEOUT_CURTO
        estado = "visible" if somente_visivel else "attached"

        # --- INÍCIO INSERÇÃO ESPERA SIMULTÂNEA DOS SELETORES ---
        # Antes cada seletor era esperado em sequência (até timeout_por_selector
        # cada um): se o primeiro não existisse, o robô ficava parado até o fim
        # do prazo antes de tentar o próximo. Agora todos são esperados ao mesmo
        # tempo e, assim que qualquer um aparece, vence o de maior prioridade
        # (ordem da lista) que estiver disponível naquele momento. O prazo total
        # é o mesmo de antes no pior caso (soma dos prazos individuais).
        timeout_total = timeout_por_selector * max(1, len(seletores))
        combinado = None
        try:
            for seletor in seletores:
                locator = page.locator(seletor)
                combinado = locator if combinado is None else combinado.or_(locator)
        except Exception:
            combinado = None

        if combinado is not None:
            try:
                combinado.first.wait_for(state=estado, timeout=timeout_total)
            except PlaywrightTimeoutError:
                print(f"[AVISO] {descricao} não localizado pelos seletores disponíveis.")
                return None
            except Exception:
                combinado = None

        if combinado is not None:
            for seletor in seletores:
                try:
                    locator = page.locator(seletor).first
                    disponivel = locator.is_visible() if somente_visivel else locator.count() > 0
                except Exception:
                    continue
                if disponivel:
                    print(f"{descricao} localizado pelo seletor: {seletor}")
                    return locator
            print(f"{descricao} localizado (seletor combinado).")
            return combinado.first
        # --- FIM INSERÇÃO ESPERA SIMULTÂNEA DOS SELETORES ---

        # Contingência: espera sequencial original (ex.: seletor inválido no combinado).
        for seletor in seletores:
            try:
                locator = page.locator(seletor).first
                locator.wait_for(
                    state=estado,
                    timeout=timeout_por_selector,
                )
                print(f"{descricao} localizado pelo seletor: {seletor}")
                return locator
            except PlaywrightTimeoutError:
                continue
            except Exception:
                continue

        print(f"[AVISO] {descricao} não localizado pelos seletores disponíveis.")
        return None

    def clicar_e_aguardar(
        self,
        locator,
        descricao,
        page=None,
        proximo_selector=None,
        usar_networkidle=True,
        timeout=None,
    ):
        """
        Clique robusto Playwright.
        - Se proximo_selector for informado, usa wait_for_selector logo após o clique.
        - Se usar_networkidle=True, aplica page.wait_for_load_state('networkidle') depois do clique.
        """
        page = page or self.page
        timeout = timeout or self.TIMEOUT_PADRAO

        print(f"Clicando em: {descricao}")
        try:
            locator.scroll_into_view_if_needed(timeout=self.TIMEOUT_CURTO)
        except Exception:
            pass

        try:
            locator.click(timeout=timeout)
        except Exception as e_click:
            print(f"[AVISO] Clique normal falhou em {descricao}. Tentando clique via JavaScript...")
            try:
                handle = locator.element_handle(timeout=self.TIMEOUT_CURTO)
                if handle:
                    page.evaluate("el => el.click()", handle)
                else:
                    raise e_click
            except Exception:
                raise e_click

        if proximo_selector:
            self.aguardar_selector(
                proximo_selector,
                descricao=f"próxima tela após {descricao}",
                page=page,
                timeout=timeout,
                state="visible",
            )

        # --- INÍCIO INSERÇÃO MODO RÁPIDO / EVITAR ESPERA NETWORKIDLE DUPLICADA ---
        # Quando a próxima tela já foi confirmada por um seletor visível, não é
        # necessário aguardar também o estado networkidle. O portal pode manter
        # conexões em segundo plano e provocar esperas desnecessárias.
        if self.modo_rapido and usar_networkidle and proximo_selector:
            return
        # --- FIM INSERÇÃO MODO RÁPIDO / EVITAR ESPERA NETWORKIDLE DUPLICADA ---

        if usar_networkidle:
            self.respirar_sistema(page=page)

    def baixar_por_click(self, locator, descricao, novo_nome_base, page=None, extensao_fallback="", registrar_documento=None):
        page = page or self.page
        try:
            print(f"Preparando captura de download: {descricao}")
            # --- INÍCIO INSERÇÃO RELATÓRIOS NOVA ABA / URL DIRETA (PREPARAR FALLBACK) ---
            paginas_antes = []
            url_antes = ""
            try:
                paginas_antes = list(self.context.pages)
                url_antes = page.url
            except Exception:
                pass
            # --- FIM INSERÇÃO RELATÓRIOS NOVA ABA / URL DIRETA (PREPARAR FALLBACK) ---
            with page.expect_download(timeout=self.TIMEOUT_DOWNLOAD) as download_info:
                self.clicar_e_aguardar(locator, descricao, page=page, usar_networkidle=False)
            download = download_info.value
            caminho = self.salvar_download(download, novo_nome_base, extensao_fallback=extensao_fallback)
            # --- INÍCIO INSERÇÃO CONTROLE TXT / CAMINHO DO ARQUIVO ---
            if registrar_documento and caminho:
                self.marcar_documento_fatura(registrar_documento, "OK", caminho=caminho)
            # --- FIM INSERÇÃO CONTROLE TXT / CAMINHO DO ARQUIVO ---
            if registrar_documento:
                self.registrar_documento_log(
                    self.contrato_atual,
                    self.vencimento_atual,
                    self.tipo_fatura_atual,
                    registrar_documento,
                )
            return caminho
        except PlaywrightTimeoutError:
            print(f"[AVISO] Download não iniciou/concluiu no tempo esperado: {descricao}. Verificando se abriu nova aba/URL direta...")
            # --- INÍCIO INSERÇÃO RELATÓRIOS NOVA ABA / URL DIRETA (SALVAR FALLBACK) ---
            caminho = self.baixar_arquivo_aberto_em_nova_aba_ou_url(
                page_origem=page,
                paginas_antes=paginas_antes,
                url_antes=url_antes,
                descricao=descricao,
                novo_nome_base=novo_nome_base,
                extensao_fallback=extensao_fallback,
            )
            if caminho:
                if registrar_documento:
                    self.marcar_documento_fatura(registrar_documento, "OK", caminho=caminho)
                    self.registrar_documento_log(
                        self.contrato_atual,
                        self.vencimento_atual,
                        self.tipo_fatura_atual,
                        registrar_documento,
                    )
                return caminho
            # --- FIM INSERÇÃO RELATÓRIOS NOVA ABA / URL DIRETA (SALVAR FALLBACK) ---
            # --- INÍCIO INSERÇÃO CONTROLE TXT / STATUS DOCUMENTO ---
            if registrar_documento:
                self.marcar_documento_fatura(registrar_documento, "ERRO", f"Download não iniciou/concluiu no tempo esperado e nenhuma nova aba/URL salvável foi encontrada: {descricao}")
            # --- FIM INSERÇÃO CONTROLE TXT / STATUS DOCUMENTO ---
            return None
        except Exception as e:
            print(f"[AVISO] Falha no download por clique ({descricao}): {self.resumir_erro(e)}")
            self.registrar_erro(f"Falha no download por clique: {descricao}", self.resumir_erro(e))
            # --- INÍCIO INSERÇÃO CONTROLE TXT / STATUS DOCUMENTO ---
            if registrar_documento:
                self.marcar_documento_fatura(registrar_documento, "ERRO", self.resumir_erro(e))
            # --- FIM INSERÇÃO CONTROLE TXT / STATUS DOCUMENTO ---
            return None

    # --- INÍCIO INSERÇÃO RELATÓRIOS NOVA ABA / URL DIRETA (FUNÇÕES AUXILIARES) ---
    def baixar_arquivo_aberto_em_nova_aba_ou_url(
        self,
        page_origem,
        paginas_antes,
        url_antes,
        descricao,
        novo_nome_base,
        extensao_fallback="",
        fechar_nova_aba=True,
    ):
        """
        Fallback para relatórios que não disparam download automático.

        Em alguns casos do portal Hapvida, principalmente TXT e PDF de relatórios,
        o clique/ENTER abre uma nova aba com URL direta, por exemplo:
        https://...blob.core.windows.net/.../mensalidade/txt/ARQUIVO.TXT?sv=...

        Quando isso acontece, page.expect_download() não captura o arquivo.
        Esta função detecta nova aba ou mudança de URL e salva o conteúdo via request.
        """
        page_origem = page_origem or self.page
        caminho = None

        try:
            self.respirar_sistema(page=page_origem, timeout=15_000, segundos_fallback=2)
        except Exception:
            pass

        try:
            paginas_depois = list(self.context.pages)
        except Exception:
            paginas_depois = []

        novas_paginas = [p for p in paginas_depois if p not in (paginas_antes or [])]

        for pagina_aberta in reversed(novas_paginas):
            try:
                pagina_aberta.set_default_timeout(self.TIMEOUT_PADRAO)
                pagina_aberta.set_default_navigation_timeout(self.TIMEOUT_LONGO)
                try:
                    pagina_aberta.wait_for_load_state("domcontentloaded", timeout=20_000)
                except Exception:
                    pass

                print(f"Nova aba detectada para {descricao}: {pagina_aberta.url}")
                caminho = self.baixar_url_da_pagina_por_request(
                    pagina_aberta,
                    novo_nome_base,
                    extensao_fallback=extensao_fallback,
                )

                if caminho:
                    print(f"Arquivo salvo a partir da nova aba: {caminho.name}")
                    return caminho

            except Exception as e:
                print(f"[AVISO] Falha ao tratar nova aba de {descricao}: {self.resumir_erro(e)}")
            finally:
                try:
                    if fechar_nova_aba and not pagina_aberta.is_closed():
                        pagina_aberta.close()
                except Exception:
                    pass
                try:
                    page_origem.bring_to_front()
                except Exception:
                    pass

        try:
            url_depois = page_origem.url
        except Exception:
            url_depois = ""

        if url_depois and url_antes and url_depois != url_antes:
            try:
                print(f"A própria aba mudou de URL para {descricao}: {url_depois}")
                caminho = self.baixar_url_da_pagina_por_request(
                    page_origem,
                    novo_nome_base,
                    extensao_fallback=extensao_fallback,
                )
                if caminho:
                    print(f"Arquivo salvo a partir da URL atual: {caminho.name}")
                    return caminho
            finally:
                try:
                    page_origem.go_back(wait_until="domcontentloaded", timeout=self.TIMEOUT_LONGO)
                    self.respirar_sistema(page=page_origem, timeout=15_000, segundos_fallback=2)
                except Exception:
                    pass

        # Última tentativa para PDFs gerados como Blob no contexto original.
        try:
            if str(extensao_fallback).lower().endswith("pdf"):
                caminho = self.salvar_blob_pdf_capturado(page_origem, novo_nome_base, extensao_fallback)
                if caminho:
                    print(f"Arquivo PDF salvo por Blob capturado: {caminho.name}")
                    return caminho
        except Exception:
            pass

        print(f"[AVISO] Nenhuma nova aba/URL direta salvável foi encontrada para {descricao}.")
        return None
    # --- FIM INSERÇÃO RELATÓRIOS NOVA ABA / URL DIRETA (FUNÇÕES AUXILIARES) ---


    # --- INÍCIO INSERÇÃO CORREÇÃO BOLETO/NF (DOWNLOAD ROBUSTO IGUAL AO SELENIUM ORIGINAL) ---
    def salvar_bytes_em_arquivo(self, conteudo, novo_nome_base, extensao_fallback=".pdf", nome_original=""):
        try:
            ext = extensao_fallback if str(extensao_fallback).startswith(".") else f".{extensao_fallback}"
            # --- INÍCIO INSERÇÃO PASTA ÚNICA POR TIPO / NOME ORIGINAL ---
            pasta_destino = self.obter_pasta_download_atual()
            caminho = self.montar_caminho_destino(pasta_destino, nome_original, novo_nome_base, ext)
            # --- FIM INSERÇÃO PASTA ÚNICA POR TIPO / NOME ORIGINAL ---

            with open(caminho, "wb") as f:
                f.write(conteudo)

            print(f"Arquivo salvo por fallback de URL: {caminho.name}")
            return caminho
        except Exception as e:
            print(f"[AVISO] Falha ao salvar bytes do arquivo: {self.resumir_erro(e)}")
            self.registrar_erro("Falha ao salvar bytes do arquivo", self.resumir_erro(e))
            return None

    # --- INÍCIO INSERÇÃO BOLETO BLOB PDF (CAPTURA DO BLOB GERADO PELO PORTAL) ---
    def preparar_captura_blob_pdf(self, page=None):
        """
        O boleto do Hapvida pode ser criado pelo front-end como Blob e aberto em uma nova aba
        com URL no formato blob:https://portal-empresa...

        Quando isso acontece, page.expect_download() não dispara e o fetch(blob:) na nova aba
        falha, porque o Blob foi criado no contexto da página original.

        Esta rotina instala um interceptador em URL.createObjectURL ANTES do clique no Boleto.pdf.
        Assim, quando o portal criar o Blob do PDF, o robô guarda o objeto Blob na página original
        e consegue salvar os bytes no disco depois do clique.
        """
        page = page or self.page
        try:
            page.evaluate(
                """
                () => {
                    window.__hapvidaCapturedBlobs = [];

                    if (!window.__hapvidaOriginalCreateObjectURL) {
                        window.__hapvidaOriginalCreateObjectURL = URL.createObjectURL.bind(URL);

                        URL.createObjectURL = function(blob) {
                            const blobUrl = window.__hapvidaOriginalCreateObjectURL(blob);

                            try {
                                window.__hapvidaCapturedBlobs = window.__hapvidaCapturedBlobs || [];
                                window.__hapvidaCapturedBlobs.push({
                                    url: blobUrl,
                                    type: blob && blob.type ? String(blob.type) : '',
                                    size: blob && blob.size ? Number(blob.size) : 0,
                                    createdAt: Date.now(),
                                    blob: blob
                                });
                            } catch (e) {
                                // Não interrompe o portal caso a captura falhe.
                            }

                            return blobUrl;
                        };
                    }
                }
                """
            )
            return True
        except Exception as e:
            print(f"[AVISO] Não foi possível preparar captura de Blob PDF: {self.resumir_erro(e)}")
            return False

    def salvar_blob_pdf_capturado(self, page=None, novo_nome_base="ARQUIVO_PDF", extensao_fallback=".pdf"):
        """
        Salva o último Blob PDF capturado na página original.
        Este é o fallback principal para boleto aberto como blob: no visualizador do Chrome.
        """
        page = page or self.page
        try:
            resultado = page.evaluate(
                """
                async () => {
                    const lista = window.__hapvidaCapturedBlobs || [];

                    for (let i = lista.length - 1; i >= 0; i--) {
                        const item = lista[i];
                        const blob = item && item.blob ? item.blob : null;

                        if (!blob) {
                            continue;
                        }

                        const tipo = (item.type || blob.type || '').toLowerCase();
                        const tamanho = item.size || blob.size || 0;

                        // Boleto normalmente vem como application/pdf.
                        // Mantemos tamanho > 1000 como fallback porque alguns retornos vêm sem content-type.
                        if (tipo.includes('pdf') || tamanho > 1000) {
                            const buffer = await blob.arrayBuffer();
                            return {
                                url: item.url || '',
                                type: tipo,
                                size: tamanho,
                                bytes: Array.from(new Uint8Array(buffer))
                            };
                        }
                    }

                    return null;
                }
                """
            )

            if not resultado or not resultado.get("bytes"):
                return None

            conteudo = bytes(resultado["bytes"])
            if not conteudo:
                return None

            if not conteudo.startswith(b"%PDF"):
                print("[AVISO] Blob capturado não começa com %PDF, mas será salvo mesmo assim para conferência.")

            caminho = self.salvar_bytes_em_arquivo(conteudo, novo_nome_base, extensao_fallback)
            if caminho:
                print(f"Blob PDF capturado e salvo com sucesso: {caminho.name}")
            return caminho

        except Exception as e:
            print(f"[AVISO] Não foi possível salvar Blob PDF capturado: {self.resumir_erro(e)}")
            return None
    # --- FIM INSERÇÃO BOLETO BLOB PDF (CAPTURA DO BLOB GERADO PELO PORTAL) ---

    def baixar_url_da_pagina_por_request(self, page_alvo, novo_nome_base, extensao_fallback=".pdf"):
        """
        Fallback necessário porque, no Playwright, alguns PDFs do Hapvida/Prefeitura
        não disparam evento de download: eles abrem em nova aba ou no visualizador PDF.
        O Selenium antigo baixava porque usava a preferência do Chrome para sempre abrir PDF externamente.
        Aqui tentamos capturar a URL já autenticada e salvar o conteúdo manualmente.
        """
        try:
            page_alvo.wait_for_load_state("domcontentloaded", timeout=15_000)
        except Exception:
            pass

        try:
            url = page_alvo.url
        except Exception:
            url = ""

        if not url or url.startswith("about:"):
            print("[AVISO] Fallback de URL não executado: URL vazia/about:blank.")
            return None

        if url.startswith("blob:"):
            try:
                conteudo_lista = page_alvo.evaluate(
                    """
                    async () => {
                        const resp = await fetch(window.location.href);
                        const buffer = await resp.arrayBuffer();
                        return Array.from(new Uint8Array(buffer));
                    }
                    """
                )
                conteudo = bytes(conteudo_lista)
                if conteudo:
                    return self.salvar_bytes_em_arquivo(conteudo, novo_nome_base, extensao_fallback)
            except Exception as e:
                print(f"[AVISO] Não foi possível baixar blob via JavaScript: {self.resumir_erro(e)}")
                return None

        try:
            resposta = self.context.request.get(url, timeout=self.TIMEOUT_DOWNLOAD)
            if not resposta.ok:
                print(f"[AVISO] Fallback de URL retornou HTTP {resposta.status}: {url}")
                return None
            conteudo = resposta.body()
            if not conteudo:
                print("[AVISO] Fallback de URL retornou conteúdo vazio.")
                return None
            nome_original = self.obter_nome_original_resposta(url, resposta.headers)
            return self.salvar_bytes_em_arquivo(
                conteudo,
                novo_nome_base,
                extensao_fallback,
                nome_original=nome_original,
            )
        except Exception as e:
            print(f"[AVISO] Falha no fallback de URL autenticada: {self.resumir_erro(e)}")
            self.registrar_erro("Falha no fallback de URL autenticada", self.resumir_erro(e))
            return None

    # --- INÍCIO INSERÇÃO ESPERA INTELIGENTE DA GERAÇÃO DO BOLETO ---
    def indicador_carregamento_boleto_visivel(self, page=None):
        """Detecta indicadores visíveis de processamento próximos ao boleto ou globais."""
        page = page or self.page
        if not page:
            return False
        try:
            return bool(
                page.evaluate(
                    r"""
                    () => {
                        const visivel = (el) => {
                            if (!el) return false;
                            const estilo = window.getComputedStyle(el);
                            const ret = el.getBoundingClientRect();
                            return estilo &&
                                   estilo.display !== 'none' &&
                                   estilo.visibility !== 'hidden' &&
                                   Number(estilo.opacity || 1) > 0 &&
                                   ret.width > 0 && ret.height > 0;
                        };

                        const seletores = [
                            '[role="progressbar"]',
                            '[aria-busy="true"]',
                            '[class*="spinner" i]',
                            '[class*="loading" i]',
                            '[class*="loader" i]',
                            '[class*="progress" i]',
                            '.MuiCircularProgress-root'
                        ];

                        for (const seletor of seletores) {
                            try {
                                const elementos = Array.from(document.querySelectorAll(seletor));
                                if (elementos.some(visivel)) return true;
                            } catch (e) {
                                // Ignora seletor não suportado pelo navegador.
                            }
                        }

                        // Procura um SVG/indicador dentro da região textual do Boleto.
                        const candidatos = Array.from(document.querySelectorAll('div, section, article, td'));
                        for (const el of candidatos) {
                            const texto = (el.innerText || '').replace(/\s+/g, ' ').trim();
                            if (!texto || !/^Boleto(?:\.pdf)?$/i.test(texto)) continue;
                            let atual = el;
                            for (let nivel = 0; nivel < 4 && atual; nivel++, atual = atual.parentElement) {
                                const indicadores = atual.querySelectorAll(
                                    'svg, [role="progressbar"], [aria-busy="true"], ' +
                                    '[class*="spinner" i], [class*="loading" i], [class*="loader" i]'
                                );
                                if (Array.from(indicadores).some(visivel)) return true;
                            }
                        }
                        return false;
                    }
                    """
                )
            )
        except Exception:
            return False

    def fechar_paginas_abertas_apos_clique(self, paginas_antes, page_origem=None):
        """Fecha somente abas criadas depois do clique e devolve o foco à fatura."""
        page_origem = page_origem or self.page
        try:
            paginas_atuais = list(self.context.pages) if self.context else []
        except Exception:
            paginas_atuais = []
        for pagina_aberta in paginas_atuais:
            if pagina_aberta in (paginas_antes or []):
                continue
            try:
                if not pagina_aberta.is_closed():
                    pagina_aberta.close()
            except Exception:
                pass
        try:
            if page_origem and not page_origem.is_closed():
                page_origem.bring_to_front()
        except Exception:
            pass

    def aguardar_geracao_boleto_inteligente(
        self,
        locator,
        descricao,
        novo_nome_base,
        page=None,
        extensao_fallback=".pdf",
    ):
        """
        Clica uma única vez e aguarda o servidor terminar a geração do boleto.

        Durante a espera monitora, sem recarregar e sem repetir o clique:
        - evento de download do Playwright;
        - Blob PDF criado na página original;
        - nova aba com PDF/blob;
        - mudança da URL da aba original;
        - PDF válido que tenha aparecido na pasta de destino;
        - indicador visual de carregamento do portal.
        """
        page = page or self.page
        timeout_ms = max(1_000, int(self.timeout_geracao_boleto))
        intervalo_ms = max(200, int(self.intervalo_monitoramento_boleto))
        inicio_monotonic = time.monotonic()
        inicio_arquivo = time.time()
        limite = inicio_monotonic + (timeout_ms / 1000.0)
        proximo_log = inicio_monotonic
        downloads_detectados = []
        urls_processadas = set()
        indicador_ja_informado = False
        indicador_estava_visivel = False

        try:
            paginas_antes = list(self.context.pages) if self.context else []
        except Exception:
            paginas_antes = []
        try:
            url_antes = str(page.url or "")
        except Exception:
            url_antes = ""

        def ao_download(download):
            downloads_detectados.append(download)

        listener_instalado = False
        try:
            page.on("download", ao_download)
            listener_instalado = True
        except Exception as e:
            print(f"[AVISO] Não foi possível instalar listener do boleto: {self.resumir_erro(e)}")

        try:
            self.preparar_captura_blob_pdf(page)
            print(f"Clicando uma única vez em: {descricao}")
            self.clicar_e_aguardar(
                locator,
                descricao,
                page=page,
                usar_networkidle=False,
            )
            print(
                "Boleto acionado. Aguardando o portal concluir a geração por até "
                f"{timeout_ms / 1000:.0f} segundo(s), sem recarregar a página..."
            )

            while time.monotonic() < limite:
                agora = time.monotonic()
                decorrido = agora - inicio_monotonic

                # 1. Evento de download tradicional.
                if downloads_detectados:
                    download = downloads_detectados.pop(0)
                    caminho = self.salvar_download(
                        download,
                        novo_nome_base,
                        extensao_fallback=extensao_fallback,
                    )
                    if caminho:
                        print(
                            "Boleto concluído por evento de download após "
                            f"{decorrido:.1f} segundo(s)."
                        )
                        self.fechar_paginas_abertas_apos_clique(paginas_antes, page)
                        return caminho

                # 2. Blob criado na página original.
                caminho = self.salvar_blob_pdf_capturado(
                    page,
                    novo_nome_base,
                    extensao_fallback,
                )
                if caminho:
                    print(
                        "Boleto concluído por Blob após "
                        f"{decorrido:.1f} segundo(s)."
                    )
                    self.fechar_paginas_abertas_apos_clique(paginas_antes, page)
                    return caminho

                # 3. PDF que apareceu na pasta local.
                caminho = self.localizar_pdf_boleto_gerado_recentemente(
                    novo_nome_base,
                    inicio_arquivo,
                )
                if caminho:
                    print(
                        "Boleto identificado na pasta após "
                        f"{decorrido:.1f} segundo(s)."
                    )
                    self.fechar_paginas_abertas_apos_clique(paginas_antes, page)
                    return caminho

                # 4. Nova aba criada pelo portal.
                try:
                    paginas_atuais = list(self.context.pages) if self.context else []
                except Exception:
                    paginas_atuais = []
                novas_paginas = [p for p in paginas_atuais if p not in paginas_antes]
                for pagina_pdf in novas_paginas:
                    try:
                        if pagina_pdf.is_closed():
                            continue
                        url_pdf = str(pagina_pdf.url or "")
                    except Exception:
                        continue
                    if not url_pdf or url_pdf.startswith("about:"):
                        continue
                    chave_url = f"popup::{url_pdf}"
                    if chave_url in urls_processadas:
                        continue
                    urls_processadas.add(chave_url)
                    caminho = self.baixar_url_da_pagina_por_request(
                        pagina_pdf,
                        novo_nome_base,
                        extensao_fallback,
                    )
                    if caminho:
                        print(
                            "Boleto concluído pela nova aba após "
                            f"{decorrido:.1f} segundo(s)."
                        )
                        self.fechar_paginas_abertas_apos_clique(paginas_antes, page)
                        return caminho

                # 5. A própria aba mudou para uma URL de PDF/blob.
                try:
                    url_atual = str(page.url or "")
                except Exception:
                    url_atual = ""
                if url_atual and url_atual != url_antes:
                    chave_url = f"origem::{url_atual}"
                    if chave_url not in urls_processadas:
                        urls_processadas.add(chave_url)
                        caminho = self.baixar_url_da_pagina_por_request(
                            page,
                            novo_nome_base,
                            extensao_fallback,
                        )
                        if caminho:
                            print(
                                "Boleto concluído pela URL atual após "
                                f"{decorrido:.1f} segundo(s)."
                            )
                            try:
                                page.go_back(
                                    wait_until="domcontentloaded",
                                    timeout=self.TIMEOUT_LONGO,
                                )
                            except Exception:
                                pass
                            return caminho

                # 6. Estado visual do indicador de carregamento.
                indicador_visivel = self.indicador_carregamento_boleto_visivel(page)
                if indicador_visivel and not indicador_ja_informado:
                    print(
                        "Indicador de carregamento do boleto detectado. "
                        "O portal ainda está gerando o arquivo; o robô continuará aguardando."
                    )
                    indicador_ja_informado = True
                if indicador_estava_visivel and not indicador_visivel:
                    print(
                        "O indicador de carregamento desapareceu. "
                        "Continuando a verificar download, Blob e nova aba."
                    )
                indicador_estava_visivel = indicador_visivel

                if agora >= proximo_log:
                    restante = max(0.0, (limite - agora))
                    estado = "carregando" if indicador_visivel else "aguardando resposta"
                    print(
                        f"[BOLETO EM GERAÇÃO] {decorrido:.0f}s decorridos | "
                        f"{restante:.0f}s restantes | estado: {estado}."
                    )
                    proximo_log = agora + self.intervalo_log_boleto

                # wait_for_timeout mantém o Playwright processando eventos.
                try:
                    page.wait_for_timeout(intervalo_ms)
                except Exception:
                    time.sleep(intervalo_ms / 1000.0)

            # Verificação final imediatamente antes de declarar timeout.
            if downloads_detectados:
                caminho = self.salvar_download(
                    downloads_detectados.pop(0),
                    novo_nome_base,
                    extensao_fallback=extensao_fallback,
                )
                if caminho:
                    self.fechar_paginas_abertas_apos_clique(paginas_antes, page)
                    return caminho

            caminho = self.salvar_blob_pdf_capturado(
                page,
                novo_nome_base,
                extensao_fallback,
            )
            if not caminho:
                caminho = self.localizar_pdf_boleto_gerado_recentemente(
                    novo_nome_base,
                    inicio_arquivo,
                )
            if caminho:
                self.fechar_paginas_abertas_apos_clique(paginas_antes, page)
                return caminho

            print(
                "[AVISO] O portal permaneceu processando o boleto por "
                f"{timeout_ms / 1000:.0f} segundo(s), mas nenhum PDF foi disponibilizado."
            )
            self.fechar_paginas_abertas_apos_clique(paginas_antes, page)
            return None

        finally:
            if listener_instalado:
                try:
                    page.remove_listener("download", ao_download)
                except Exception:
                    pass
    # --- FIM INSERÇÃO ESPERA INTELIGENTE DA GERAÇÃO DO BOLETO ---

    def baixar_por_click_com_fallback_pdf(self, locator, descricao, novo_nome_base, page=None, extensao_fallback=".pdf", registrar_documento=None):
        """
        Mantém o mecanismo Playwright expect_download, mas adiciona fallback para o comportamento
        que funcionava no Selenium original: Boleto/NF podem abrir PDF em nova aba ou navegar
        na própria aba, sem gerar evento de download.
        """
        page = page or self.page
        # --- INÍCIO INSERÇÃO BOLETO BLOB PDF (PREPARAR CAPTURA ANTES DO CLIQUE) ---
        # Precisa ocorrer antes do clique, porque o portal cria o Blob no momento em que o Boleto.pdf é acionado.
        self.preparar_captura_blob_pdf(page)
        # --- FIM INSERÇÃO BOLETO BLOB PDF (PREPARAR CAPTURA ANTES DO CLIQUE) ---
        paginas_antes = []
        url_antes = ""
        try:
            paginas_antes = list(self.context.pages)
            url_antes = page.url
        except Exception:
            pass

        caminho = None

        # --- INÍCIO INSERÇÃO TIMEOUT ESPECÍFICO DO BOLETO ---
        documento_normalizado = self.normalizar_documento_controle(registrar_documento)
        eh_boleto = documento_normalizado == "BOLETO" or "BOLETO" in str(descricao or "").upper()
        timeout_evento_atual = self.timeout_evento_boleto if eh_boleto else self.timeout_evento_pdf
        # --- FIM INSERÇÃO TIMEOUT ESPECÍFICO DO BOLETO ---

        # O boleto não usa mais uma espera curta seguida de recarga imediata.
        # Clicamos uma única vez e aguardamos inteligentemente a geração completa.
        if eh_boleto:
            try:
                caminho = self.aguardar_geracao_boleto_inteligente(
                    locator,
                    descricao,
                    novo_nome_base,
                    page=page,
                    extensao_fallback=extensao_fallback,
                )
                if caminho and registrar_documento:
                    self.marcar_documento_fatura(
                        registrar_documento,
                        "OK",
                        caminho=caminho,
                    )
                    self.registrar_documento_log(
                        self.contrato_atual,
                        self.vencimento_atual,
                        self.tipo_fatura_atual,
                        registrar_documento,
                    )
                return caminho
            except Exception as e:
                print(
                    f"[AVISO] Falha durante a espera inteligente do boleto: "
                    f"{self.resumir_erro(e)}"
                )
                if registrar_documento:
                    self.registrar_diagnostico_se_ausente(
                        registrar_documento,
                        "ESPERA_INTELIGENTE_BOLETO",
                        self.resumir_erro(e),
                        page=page,
                        excecao=e,
                    )
                return None

        try:
            print(f"Preparando captura robusta de download/PDF: {descricao}")
            if eh_boleto:
                print(f"Timeout do evento tradicional do boleto: {timeout_evento_atual / 1000:.1f} segundo(s).")
            try:
                with page.expect_download(timeout=timeout_evento_atual) as download_info:
                    self.clicar_e_aguardar(locator, descricao, page=page, usar_networkidle=False)
                download = download_info.value
                caminho = self.salvar_download(download, novo_nome_base, extensao_fallback=extensao_fallback)
            except PlaywrightTimeoutError:
                print(f"[AVISO] {descricao} não disparou evento de download. Verificando nova aba/URL...")
                self.respirar_sistema(page=page, timeout=self.timeout_respiro_rapido, segundos_fallback=0.35)

                # --- INÍCIO INSERÇÃO BOLETO BLOB PDF (SALVAR BLOB CAPTURADO APÓS CLIQUE) ---
                # Se o portal abriu o boleto como blob:, o conteúdo real fica preso na página original.
                # Tentamos salvar o Blob antes de depender da URL da nova aba.
                caminho = self.salvar_blob_pdf_capturado(page, novo_nome_base, extensao_fallback)
                # --- FIM INSERÇÃO BOLETO BLOB PDF (SALVAR BLOB CAPTURADO APÓS CLIQUE) ---

                paginas_depois = []
                try:
                    paginas_depois = list(self.context.pages)
                except Exception:
                    paginas_depois = []

                novas_paginas = [p for p in paginas_depois if p not in paginas_antes]

                # --- INÍCIO INSERÇÃO ENCERRAMENTO IMEDIATO APÓS BLOB SALVO ---
                # Quando o Blob já foi salvo, o arquivo está concluído. Não aguardamos
                # o carregamento da aba blob: e não executamos os fallbacks de URL.
                # Apenas fechamos rapidamente as abas que o clique possa ter aberto.
                if caminho:
                    print(f"Blob/PDF já salvo com sucesso para {descricao}. Encerrando fallbacks adicionais.")
                    for pagina_aberta in novas_paginas:
                        try:
                            if not pagina_aberta.is_closed():
                                pagina_aberta.close()
                        except Exception:
                            pass
                    try:
                        page.bring_to_front()
                    except Exception:
                        pass
                # --- FIM INSERÇÃO ENCERRAMENTO IMEDIATO APÓS BLOB SALVO ---
                elif novas_paginas:
                    pagina_pdf = novas_paginas[-1]
                    try:
                        pagina_pdf.set_default_timeout(self.TIMEOUT_PADRAO)
                        pagina_pdf.set_default_navigation_timeout(self.TIMEOUT_LONGO)
                        pagina_pdf.wait_for_load_state("domcontentloaded", timeout=20_000)
                    except Exception:
                        pass
                    print(f"Nova aba detectada para {descricao}: {pagina_pdf.url}")
                    caminho = self.baixar_url_da_pagina_por_request(pagina_pdf, novo_nome_base, extensao_fallback)
                    try:
                        if not pagina_pdf.is_closed():
                            pagina_pdf.close()
                    except Exception:
                        pass
                    try:
                        page.bring_to_front()
                    except Exception:
                        pass
                else:
                    try:
                        url_depois = page.url
                    except Exception:
                        url_depois = ""
                    if url_depois and url_depois != url_antes:
                        print(f"A própria aba mudou de URL para {descricao}: {url_depois}")
                        caminho = self.baixar_url_da_pagina_por_request(page, novo_nome_base, extensao_fallback)
                        try:
                            page.go_back(wait_until="domcontentloaded", timeout=self.TIMEOUT_LONGO)
                            self.respirar_sistema(page=page, timeout=15_000, segundos_fallback=2)
                        except Exception:
                            pass
                    else:
                        # --- INÍCIO INSERÇÃO BOLETO BLOB PDF (ÚLTIMA TENTATIVA) ---
                        caminho = self.salvar_blob_pdf_capturado(page, novo_nome_base, extensao_fallback)
                        # --- FIM INSERÇÃO BOLETO BLOB PDF (ÚLTIMA TENTATIVA) ---
                        if not caminho:
                            mensagem_sem_resposta = (
                                f"Após o clique em {descricao}, não houve evento de download, "
                                "popup, Blob válido nem mudança de URL."
                            )
                            print(f"[AVISO] Nenhum download, popup, Blob ou mudança de URL detectado para {descricao}.")
                            if registrar_documento:
                                self.registrar_diagnostico_se_ausente(
                                    registrar_documento,
                                    "GERAR_OU_CAPTURAR_PDF",
                                    mensagem_sem_resposta,
                                    page=page,
                                )

            # Se todos os caminhos de captura terminaram sem arquivo, registra o
            # ponto exato mesmo quando houve popup ou mudança de URL sem PDF válido.
            if not caminho and registrar_documento:
                self.registrar_diagnostico_se_ausente(
                    registrar_documento,
                    "GERAR_OU_CAPTURAR_PDF",
                    f"{descricao} foi acionado, mas nenhum PDF válido foi salvo.",
                    page=page,
                )

            # --- INÍCIO INSERÇÃO CONTROLE TXT / CAMINHO DO ARQUIVO ---
            if caminho and registrar_documento:
                self.marcar_documento_fatura(registrar_documento, "OK", caminho=caminho)
            # --- FIM INSERÇÃO CONTROLE TXT / CAMINHO DO ARQUIVO ---
            if caminho and registrar_documento:
                self.registrar_documento_log(
                    self.contrato_atual,
                    self.vencimento_atual,
                    self.tipo_fatura_atual,
                    registrar_documento,
                )
            return caminho

        except Exception as e:
            print(f"[AVISO] Falha na captura robusta ({descricao}): {self.resumir_erro(e)}")
            self.registrar_erro(f"Falha na captura robusta: {descricao}", self.resumir_erro(e))
            if registrar_documento:
                self.registrar_diagnostico_se_ausente(
                    registrar_documento,
                    "CAPTURA_ROBUSTA_PDF",
                    self.resumir_erro(e),
                    page=page,
                    excecao=e,
                )
            return None

    # ----------------------------------------------------------------------
    # Login e navegação principal
    # ----------------------------------------------------------------------
    def obter_credenciais(self):
        # --- INÍCIO INSERÇÃO LOGIN AUTOMÁTICO (OBTER CREDENCIAIS) ---
        email_padrao = str(LOGIN_EMAIL_PADRAO).strip()
        senha_padrao = str(LOGIN_SENHA_PADRAO).strip()

        email = os.getenv("HAPVIDA_EMAIL", email_padrao).strip()
        senha = os.getenv("HAPVIDA_SENHA", senha_padrao).strip()

        if email and senha:
            origem = ".env" if "HAPVIDA_SENHA" in VARIAVEIS_ENV_CARREGADAS else "variáveis de ambiente"
            print(f"Credenciais carregadas automaticamente ({origem}).")
            return email, senha

        print("[AVISO] HAPVIDA_EMAIL/HAPVIDA_SENHA não encontrados no .env. As credenciais serão solicitadas.")
        # --- FIM INSERÇÃO LOGIN AUTOMÁTICO (OBTER CREDENCIAIS) ---

        email = os.getenv("HAPVIDA_EMAIL", "").strip()
        senha = os.getenv("HAPVIDA_SENHA", "").strip()

        if not email:
            email = input("E-mail Hapvida/NDI: ").strip()
        if not senha:
            senha = getpass.getpass("Senha Hapvida/NDI: ").strip()

        if not email or not senha:
            raise ValueError("E-mail e senha são obrigatórios para login.")

        return email, senha

    def fazer_login(self):
        print("Acessando o portal...")
        self.page.goto(self.url_login, wait_until="domcontentloaded", timeout=self.TIMEOUT_LONGO)
        self.respirar_sistema(timeout=20_000, segundos_fallback=2)

        email, senha = self.obter_credenciais()

        print("Preenchendo credenciais...")
        email_input = self.localizar_primeiro(
            "campo de e-mail",
            [
                "xpath=/html/body/main/section[1]/div[1]/form/div[3]/div[1]/input",
                "xpath=//input[@type='email']",
                "xpath=//input[contains(@name, 'email') or contains(@id, 'email')]",
                "xpath=//input[contains(@placeholder, 'E-mail') or contains(@placeholder, 'email')]",
            ],
            timeout_por_selector=10_000,
        )
        if not email_input:
            raise RuntimeError("Campo de e-mail não localizado.")
        email_input.fill(email)

        senha_input = self.localizar_primeiro(
            "campo de senha",
            [
                "xpath=/html/body/main/section[1]/div[1]/form/div[3]/div[2]/input",
                "xpath=//input[@type='password']",
                "xpath=//input[contains(@name, 'password') or contains(@id, 'password')]",
            ],
            timeout_por_selector=10_000,
        )
        if not senha_input:
            raise RuntimeError("Campo de senha não localizado.")
        senha_input.fill(senha)

        btn_login = self.localizar_primeiro(
            "botão de login",
            [
                "xpath=/html/body/main/section[1]/div[1]/form/div[3]/div[4]/button",
                "xpath=//button[contains(normalize-space(.), 'Entrar')]",
                "xpath=//button[@type='submit']",
            ],
            timeout_por_selector=10_000,
        )
        if not btn_login:
            raise RuntimeError("Botão de login não localizado.")

        self.clicar_e_aguardar(
            btn_login,
            "botão de login",
            proximo_selector="xpath=//*[contains(., 'Perfil de acesso') or contains(., 'CNPJ') or contains(., 'Contrato') or contains(., 'Empresa')]",
            usar_networkidle=True,
            timeout=self.TIMEOUT_LONGO,
        )

        self.url_painel = self.page.url
        print("Login efetuado. Painel carregado.")

    def voltar_para_painel(self):
        print("\nRetornando para o painel principal...")
        if not self.url_painel:
            raise RuntimeError("URL do painel não foi definida.")
        self.page.goto(self.url_painel, wait_until="domcontentloaded", timeout=self.TIMEOUT_LONGO)
        self.respirar_sistema(timeout=20_000, segundos_fallback=2)

    def selecionar_empresa_matriz(self, modulo):
        # --- INÍCIO INSERÇÃO ACESSO DIRETO À PESQUISA DE CONTRATOS ---
        # O portal atualizado já abre diretamente na tela de pesquisa de contratos
        # após o login. A seleção anterior de ADESÃO/PME por CNPJ não é mais executada.
        print("Portal atualizado: seleção de ADESÃO/PME dispensada. Prosseguindo para a pesquisa do contrato...")
        return
        # --- FIM INSERÇÃO ACESSO DIRETO À PESQUISA DE CONTRATOS ---

        modulo_txt = str(modulo).strip().upper()
        if modulo_txt == "ADESÃO":
            cnpj_numeros = "17670901000193"
            cnpj_formatado = "17.670.901/0001-93"
            print(f"Módulo ADESÃO detectado. Buscando CNPJ: {cnpj_numeros} ou {cnpj_formatado}...")
        elif modulo_txt == "PME":
            cnpj_numeros = "42700473000141"
            cnpj_formatado = "42.700.473/0001-41"
            print(f"Módulo PME detectado. Buscando CNPJ: {cnpj_numeros} ou {cnpj_formatado}...")
        else:
            raise ValueError(f"Módulo desconhecido na planilha: {modulo}")

        elemento_cnpj = self.localizar_primeiro(
            "empresa matriz/CNPJ",
            [
                f"xpath=//*[contains(normalize-space(.), '{cnpj_numeros}') or contains(normalize-space(.), '{cnpj_formatado}')]",
            ],
            timeout_por_selector=20_000,
        )
        if not elemento_cnpj:
            raise RuntimeError("CNPJ da empresa matriz não localizado.")

        self.clicar_e_aguardar(elemento_cnpj, "empresa matriz/CNPJ", usar_networkidle=True)

        print("Confirmando seleção da empresa com TAB + ENTER...")
        self.page.keyboard.press("Tab")
        time.sleep(0.4)
        self.page.keyboard.press("Enter")
        self.respirar_sistema(timeout=20_000, segundos_fallback=2)

        self.aguardar_selector(
            "xpath=/html/body/div/div/div/main/div/div/main/div[1]/div/div[1]/div/div/div/input | //input",
            descricao="campo de busca de contrato",
            timeout=self.TIMEOUT_LONGO,
        )

    def buscar_contrato(self, contrato):
        print(f"Buscando contrato: {contrato}...")

        # --- INÍCIO INSERÇÃO SELEÇÃO DA OPERADORA NDI ---
        # Regra do portal atualizado: antes de editar/pesquisar o número do contrato,
        # é obrigatório selecionar a operadora NDI na lista suspensa.
        # Gravação de 28/09/2026: a lista apresenta CCG e NDI. A seleção deve
        # usar o nome exato, pois o primeiro item não é necessariamente NDI.
        try:
            print("Selecionando a operadora NDI antes da pesquisa do contrato...")

            seletor_operadora = self.localizar_primeiro(
                'lista suspensa "Selecione a operadora"',
                [
                    "xpath=//*[@id='subHeaderPageElementId']//*[@role='combobox']",
                    "xpath=//*[@role='combobox' and (normalize-space(.)='Selecione a operadora' or normalize-space(.)='NDI' or normalize-space(.)='CCG')]",
                ],
                timeout_por_selector=self.TIMEOUT_CURTO,
            )
            if not seletor_operadora:
                raise RuntimeError('Lista suspensa "Selecione a operadora" não localizada.')

            operadora_atual = (
                seletor_operadora.inner_text(timeout=self.TIMEOUT_CURTO)
                .replace("\u200b", "")
                .strip()
            )
            if operadora_atual != "NDI":
                seletor_operadora.click(timeout=self.TIMEOUT_PADRAO)

                opcao_ndi = self.localizar_primeiro(
                    'opção "NDI" da operadora',
                    [
                        "xpath=//*[@role='listbox']//*[@role='option' and normalize-space(.)='NDI']",
                        "xpath=//li[normalize-space(.)='NDI']",
                    ],
                    timeout_por_selector=self.TIMEOUT_CURTO,
                )
                if not opcao_ndi:
                    raise RuntimeError('Opção "NDI" não localizada na lista de operadoras.')

                opcao_ndi.click(timeout=self.TIMEOUT_PADRAO)
            else:
                print("Operadora NDI já está selecionada nesta tela.")

            # A confirmação usa o combobox, não o texto da opção ainda aberta.
            # O locator é resolvido novamente se o portal recriar o componente.
            expect(seletor_operadora).to_have_text(
                re.compile(r"^[\s\u200b]*NDI[\s\u200b]*$"),
                timeout=self.TIMEOUT_PADRAO,
            )

            # Confirma que a seleção liberou o campo mostrado na gravação.
            # O ID React do input é dinâmico; o placeholder identifica o campo.
            campo_contrato = self.localizar_primeiro(
                "campo de contrato após selecionar NDI",
                [
                    "xpath=//input[@placeholder='Buscar por contrato, nome fantasia, razão social e CNPJ']",
                    "xpath=//*[@id='subHeaderPageElementId']//input[contains(@placeholder, 'contrato') or contains(@placeholder, 'Contrato')]",
                ],
                timeout_por_selector=self.TIMEOUT_CURTO,
            )
            if not campo_contrato:
                raise RuntimeError("Campo de pesquisa do contrato não localizado após selecionar NDI.")

            expect(campo_contrato).to_be_editable(timeout=self.TIMEOUT_PADRAO)
            print("Operadora NDI confirmada. Campo de pesquisa do contrato liberado.")

        except Exception as erro_operadora:
            raise RuntimeError(
                f"Não foi possível selecionar a operadora NDI antes da pesquisa do contrato: "
                f"{self.resumir_erro(erro_operadora)}"
            ) from erro_operadora
        # --- FIM INSERÇÃO SELEÇÃO DA OPERADORA NDI ---

        # --- INÍCIO INSERÇÃO OTIMIZAÇÃO BUSCA DO CONTRATO POR TECLADO ---
        # Fluxo observado diretamente no portal após a página estar carregada:
        #   clicar no texto "CNPJ"
        #   TAB 3x -> barra de pesquisa
        #   digitar contrato
        #   TAB 1x -> botão Buscar
        #   ENTER -> executar a pesquisa
        #
        # Este passa a ser o caminho principal porque elimina os timeouts causados
        # pelos XPaths absolutos antigos. Os seletores continuam abaixo apenas como
        # contingência, com timeout curto, caso o foco da página esteja diferente.
        try:
            print('Pesquisa rápida: clicando no texto "CNPJ"...')

            texto_cnpj = self.localizar_primeiro(
                'texto "CNPJ"',
                [
                    "xpath=//*[normalize-space(text())='CNPJ']",
                    "xpath=//*[contains(normalize-space(text()), 'CNPJ')]",
                ],
                timeout_por_selector=3_000,
            )
            if not texto_cnpj:
                raise RuntimeError('Texto "CNPJ" não localizado para iniciar a pesquisa.')

            texto_cnpj.click()
            time.sleep(0.15)

            print("TAB 3x para acessar a barra de pesquisa...")
            for _ in range(3):
                self.page.keyboard.press("Tab")
                time.sleep(0.15)

            print(f"Digitando o contrato pela navegação rápida: {contrato}")
            self.page.keyboard.press("Control+A")
            self.page.keyboard.type(str(contrato), delay=20)

            print("TAB 1x para acessar o botão Buscar...")
            self.page.keyboard.press("Tab")
            time.sleep(0.15)

            print("ENTER para executar a pesquisa do contrato...")
            self.page.keyboard.press("Enter")

            self.aguardar_selector(
                "xpath=//*[contains(normalize-space(.), 'Perfil de acesso')]",
                descricao=f"resultado da pesquisa do contrato {contrato}",
                timeout=10_000,
            )

            print("Contrato localizado com sucesso pela navegação rápida por teclado.")
            return

        except Exception as erro_teclado:
            print(
                f"[AVISO] Pesquisa rápida pelo teclado não foi confirmada: "
                f"{self.resumir_erro(erro_teclado)}"
            )
            print("Executando contingência pelos seletores atualizados...")

            try:
                self.page.keyboard.press("Escape")
            except Exception:
                pass
        # --- FIM INSERÇÃO OTIMIZAÇÃO BUSCA DO CONTRATO POR TECLADO ---

        # --- INÍCIO CONTINGÊNCIA BUSCA DO CONTRATO POR SELETORES ---
        # Os seletores que funcionaram nos testes ficam primeiro. Os XPaths
        # absolutos antigos foram removidos para não consumir 20 + 10 segundos.
        barra_busca = self.localizar_primeiro(
            "barra de busca de contrato",
            [
                "xpath=//input[contains(@placeholder, 'Contrato') or contains(@placeholder, 'contrato') or contains(@placeholder, 'Buscar') or contains(@placeholder, 'buscar')]",
                "xpath=//main//input",
            ],
            timeout_por_selector=3_000,
        )
        if not barra_busca:
            raise RuntimeError("Barra de busca do contrato não localizada.")

        barra_busca.fill(str(contrato))

        btn_buscar = self.localizar_primeiro(
            "botão buscar contrato",
            [
                "xpath=//button[contains(normalize-space(.), 'Buscar')]",
                "xpath=//main//button[1]",
            ],
            timeout_por_selector=3_000,
        )
        if not btn_buscar:
            raise RuntimeError("Botão Buscar não localizado.")

        self.clicar_e_aguardar(
            btn_buscar,
            "botão buscar contrato - contingência",
            proximo_selector="xpath=//*[contains(normalize-space(.), 'Perfil de acesso')]",
            usar_networkidle=False,
            timeout=10_000,
        )
        print("Contrato localizado pela contingência com seletores.")
        # --- FIM CONTINGÊNCIA BUSCA DO CONTRATO POR SELETORES ---

    def entrar_no_contrato(self):
        print("Entrando nas opções do contrato selecionado...")
        elemento_perfil = self.localizar_primeiro(
            "Perfil de acesso",
            [
                "xpath=//*[contains(normalize-space(.), 'Perfil de acesso')]",
            ],
            timeout_por_selector=20_000,
        )
        if not elemento_perfil:
            raise RuntimeError("Elemento 'Perfil de acesso' não localizado.")

        self.clicar_e_aguardar(elemento_perfil, "Perfil de acesso", usar_networkidle=True)

        print("Confirmando entrada no contrato com TAB 2x + ENTER...")
        self.page.keyboard.press("Tab")
        time.sleep(0.2)
        self.page.keyboard.press("Tab")
        time.sleep(0.2)
        self.page.keyboard.press("Enter")

        self.aguardar_selector(
            self.SELETOR_TELA_CONTRATO,
            descricao="tela do contrato (botão Extrato Financeiro)",
            timeout=self.TIMEOUT_LONGO,
        )
        self.respirar_sistema(timeout=20_000, segundos_fallback=2)

    def acessar_extrato_financeiro(self):
        print("Acessando a aba de Extrato Financeiro...")
        # --- INÍCIO INSERÇÃO CRONOMETRAGEM DO EXTRATO EM SUB-ETAPAS ---
        # Separa "achar o botão" de "esperar a tela do extrato" no CSV de tempos,
        # para saber onde está o tempo gasto nesta etapa.
        with self.medir_etapa("EXTRATO_LOCALIZAR_BOTAO"):
            # Seletores por texto primeiro; o XPath absoluto fica por último,
            # pois quebra com qualquer mudança de layout do portal.
            btn_extrato = self.localizar_primeiro(
                "aba/botão Extrato Financeiro",
                [
                    "xpath=//*[self::button or @role='button'][contains(normalize-space(.), 'Extrato Financeiro')]",
                    "xpath=//*[self::button or @role='button'][contains(normalize-space(.), 'Extrato')]",
                    "xpath=/html/body/div/div/div/main/div/div/main/div/div[3]/div/button[3]",
                ],
                timeout_por_selector=20_000,
            )
        if not btn_extrato:
            raise RuntimeError("Botão/aba de Extrato Financeiro não localizado.")

        with self.medir_etapa("EXTRATO_CLICAR_E_AGUARDAR_TELA"):
            self.clicar_e_aguardar(
                btn_extrato,
                "Extrato Financeiro",
                proximo_selector=self.SELETOR_TELA_EXTRATO,
                usar_networkidle=True,
                timeout=self.TIMEOUT_LONGO,
            )
        # --- FIM INSERÇÃO CRONOMETRAGEM DO EXTRATO EM SUB-ETAPAS ---

    def voltar_para_extrato(self):
        print("\nVoltando para o Extrato Financeiro...")
        try:
            self.page.go_back(wait_until="domcontentloaded", timeout=self.TIMEOUT_LONGO)
        except Exception:
            pass
        self.respirar_sistema(timeout=20_000, segundos_fallback=2)
        self.aguardar_selector(
            self.SELETOR_TELA_EXTRATO,
            descricao="tabela do extrato financeiro",
            timeout=self.TIMEOUT_LONGO,
        )

    # ----------------------------------------------------------------------
    # Busca de faturas, abas e paginação
    # ----------------------------------------------------------------------
    def navegar_para_aba_e_pagina(self, tab_nome, pagina):
        # --- INÍCIO INSERÇÃO OTIMIZAÇÃO DE TEMPO / PAGINAÇÃO DIRETA ---
        if tab_nome == "ABERTO":
            print("Abrindo aba Em Aberto...")
            btn = self.localizar_primeiro(
                "aba Em Aberto",
                [
                    "xpath=//*[@id='root']/div/div/main/div/div/main/div[3]/div/div[3]/div[1]/div/div/div/button[1]",
                    "xpath=//*[self::button or @role='button'][contains(normalize-space(.), 'Em Aberto')]",
                ],
                timeout_por_selector=4_000 if self.modo_rapido else 10_000,
            )
            if btn:
                self.clicar_e_aguardar(btn, "aba Em Aberto", usar_networkidle=False)
                return True
            return False

        if tab_nome == "HISTORICO":
            print(f"Abrindo aba Histórico, página {pagina}...")
            btn = self.localizar_primeiro(
                "aba Histórico",
                [
                    "xpath=//*[@id='root']/div/div/main/div/div/main/div[3]/div/div[3]/div[1]/div/div/div/button[2]",
                    "xpath=//*[self::button or @role='button'][contains(normalize-space(.), 'Histórico') or contains(normalize-space(.), 'Historico')]",
                ],
                timeout_por_selector=4_000 if self.modo_rapido else 10_000,
            )
            if not btn:
                return False
            self.clicar_e_aguardar(btn, "aba Histórico", usar_networkidle=False)

            if pagina <= 1:
                return True

            seletores_pagina = [
                f"xpath=//*[self::button or self::a or @role='button'][normalize-space(.)='{pagina}' and not(@disabled)]",
                f"xpath=//*[contains(@class, 'pagination')]//*[self::button or self::a][normalize-space(.)='{pagina}']",
                f"xpath=//*[contains(@aria-label, 'página {pagina}') or contains(@aria-label, 'Página {pagina}')]",
            ]
            botao_pagina = self.localizar_primeiro(
                f"botão da página {pagina}",
                seletores_pagina,
                timeout_por_selector=1_500 if self.modo_rapido else 4_000,
            )
            if botao_pagina:
                self.clicar_e_aguardar(botao_pagina, f"página {pagina} do Histórico", usar_networkidle=False)
                return True

            # Fallback compatível com o portal antigo. Se a página não existir, retorna False
            # para interromper a varredura, evitando percorrer páginas inexistentes.
            proximo = self.localizar_primeiro(
                "botão Próxima página",
                [
                    "xpath=//*[self::button or self::a or @role='button'][not(@disabled) and (contains(@aria-label, 'Próxima') or contains(@aria-label, 'proxima') or normalize-space(.)='>')]",
                    "xpath=//*[self::button or self::a or @role='button'][not(@disabled) and contains(normalize-space(.), 'Próxima')]",
                ],
                timeout_por_selector=1_500 if self.modo_rapido else 4_000,
            )
            if proximo:
                self.clicar_e_aguardar(proximo, "Próxima página do Histórico", usar_networkidle=False)
                return True

            print(f"[INFO] Página {pagina} não disponível. Encerrando paginação do Histórico.")
            return False

        return False
        # --- FIM INSERÇÃO OTIMIZAÇÃO DE TEMPO / PAGINAÇÃO DIRETA ---

    # --- INÍCIO INSERÇÃO ESPERA DAS LINHAS DO EXTRATO ---
    def aguardar_linhas_extrato(self, tab_nome="", pagina=1):
        """
        Aguarda a tabela do extrato terminar de renderizar antes da leitura.

        A leitura das linhas logo após o clique na aba pode ocorrer antes de o
        portal preencher a tabela. Esta espera termina quando:
        - há linhas com data e a quantidade ficou estável entre duas leituras; ou
        - não há linhas nem indicador de carregamento durante a carência
          configurada (tabela realmente vazia); ou
        - o tempo máximo foi atingido.

        Retorna a quantidade de linhas com data encontradas.
        """
        timeout_s = max(1.0, self.timeout_linhas_extrato / 1000.0)
        inicio = time.monotonic()
        limite = inicio + timeout_s
        quantidade_anterior = -1
        quantidade = 0
        inicio_sem_linhas = None

        while time.monotonic() < limite:
            try:
                estado = self.page.evaluate(
                    r"""
                    () => {
                        const visivel = (el) => {
                            if (!el) return false;
                            const s = window.getComputedStyle(el);
                            const r = el.getBoundingClientRect();
                            return s && s.display !== 'none' && s.visibility !== 'hidden' &&
                                   r.width > 0 && r.height > 0;
                        };
                        const linhas = Array.from(document.querySelectorAll('tr')).filter(
                            (tr) => /\d{2}\/\d{2}\/\d{4}/.test(tr.innerText || '')
                        ).length;
                        const seletores = [
                            '[role="progressbar"]', '[aria-busy="true"]',
                            '.MuiCircularProgress-root', '.MuiSkeleton-root',
                            '[class*="spinner" i]', '[class*="loading" i]',
                            '[class*="loader" i]', '[class*="skeleton" i]'
                        ];
                        let carregando = false;
                        for (const seletor of seletores) {
                            try {
                                if (Array.from(document.querySelectorAll(seletor)).some(visivel)) {
                                    carregando = true;
                                    break;
                                }
                            } catch (e) {}
                        }
                        return { linhas, carregando };
                    }
                    """
                )
                quantidade = int(estado.get("linhas", 0) or 0)
                carregando = bool(estado.get("carregando"))
            except Exception:
                quantidade = 0
                carregando = True

            agora = time.monotonic()
            if quantidade > 0:
                inicio_sem_linhas = None
                if quantidade == quantidade_anterior and not carregando:
                    return quantidade
            elif carregando:
                inicio_sem_linhas = None
            else:
                if inicio_sem_linhas is None:
                    inicio_sem_linhas = agora
                elif (agora - inicio_sem_linhas) >= self.carencia_extrato_vazio:
                    print(f"Tabela do extrato sem linhas em {tab_nome} (página {pagina}).")
                    return 0

            quantidade_anterior = quantidade
            try:
                self.page.wait_for_timeout(300)
            except Exception:
                time.sleep(0.30)

        if quantidade <= 0:
            print(
                f"[AVISO] As linhas do extrato não apareceram em {timeout_s:.0f}s "
                f"em {tab_nome} (página {pagina})."
            )
        return quantidade

    def registrar_falta_de_fatura_no_extrato(self, vencimento_alvo):
        """
        Quando o vencimento não é encontrado, deixa registrado o que havia na
        tela: screenshot, datas e linhas lidas em cada aba/página.
        Retorna um resumo curto para a evidência do contrato.
        """
        datas = []
        for item in self.linhas_extrato_vistas:
            for data in re.findall(r"\b\d{2}/\d{2}/\d{4}\b", item["texto"]):
                if data not in datas:
                    datas.append(data)

        print(f"Linhas lidas no extrato: {len(self.linhas_extrato_vistas)}")
        if not self.linhas_extrato_vistas:
            print("Nenhuma linha com data foi lida em nenhuma aba (tabela vazia ou não carregada).")
        for item in self.linhas_extrato_vistas[:40]:
            print(f"  [{item['tab']} pág. {item['pagina']}] {item['texto'][:200]}")
        if len(self.linhas_extrato_vistas) > 40:
            print(f"  ... e mais {len(self.linhas_extrato_vistas) - 40} linha(s).")

        if datas:
            resumo = (
                f"Datas existentes no extrato ({len(self.linhas_extrato_vistas)} linha(s) lidas): "
                + ", ".join(datas[:40])
            )
        else:
            resumo = "Nenhuma linha com data foi lida no extrato (tabela vazia ou não carregada)."
        print(resumo)

        self.registrar_erro(
            f"Fatura não encontrada para o vencimento {vencimento_alvo}",
            resumo,
        )
        try:
            with TRAVA_ARQUIVOS, open(self.erros_log_file, "a", encoding="utf-8") as f:
                for item in self.linhas_extrato_vistas[:200]:
                    f.write(
                        f"    LINHA DO EXTRATO [{item['tab']} pág. {item['pagina']}]: "
                        f"{item['texto'][:300]}\n"
                    )
        except Exception:
            pass

        self.salvar_screenshot("fatura_nao_encontrada_no_extrato")
        return resumo
    # --- FIM INSERÇÃO ESPERA DAS LINHAS DO EXTRATO ---

    def pesquisar_ocorrencias_na_tela(self, vencimento_alvo, tab_nome, pagina):
        ocorrencias = []
        try:
            # --- INÍCIO INSERÇÃO ESPERA DAS LINHAS DO EXTRATO ---
            self.aguardar_linhas_extrato(tab_nome, pagina)
            # --- FIM INSERÇÃO ESPERA DAS LINHAS DO EXTRATO ---
            linhas = self.page.locator("xpath=//tr")
            total_linhas = linhas.count()

            for idx in range(total_linhas):
                linha = linhas.nth(idx)
                try:
                    texto_linha = linha.inner_text(timeout=3_000).upper()
                except Exception:
                    continue

                # --- INÍCIO INSERÇÃO REGISTRO DAS LINHAS VISTAS NO EXTRATO ---
                texto_resumido = re.sub(r"\s+", " ", texto_linha).strip()
                if re.search(r"\d{2}/\d{2}/\d{4}", texto_resumido):
                    self.linhas_extrato_vistas.append({
                        "tab": tab_nome,
                        "pagina": pagina,
                        "texto": texto_resumido,
                    })
                # --- FIM INSERÇÃO REGISTRO DAS LINHAS VISTAS NO EXTRATO ---

                if str(vencimento_alvo) not in texto_linha:
                    continue

                if self.TIPO_CARNET in texto_linha:
                    tipo_fatura = self.TIPO_CARNET
                elif self.TIPO_MENSALIDADE in texto_linha:
                    tipo_fatura = self.TIPO_MENSALIDADE
                else:
                    tipo_fatura = "OUTRO"

                ocorrencias.append(
                    {
                        "tipo_fatura": tipo_fatura,
                        "linha_index": idx,
                        "tab": tab_nome,
                        "pagina": pagina,
                    }
                )
        except Exception as e:
            print(f"[AVISO] Falha ao pesquisar ocorrências na tela: {self.resumir_erro(e)}")

        return ocorrencias

    def abrir_fatura_por_ocorrencia(self, ocorrencia, vencimento_alvo, navegar=True):
        tab_nome = ocorrencia["tab"]
        pagina = ocorrencia["pagina"]
        linha_index = ocorrencia["linha_index"]
        tipo_fatura = ocorrencia["tipo_fatura"]

        if navegar:
            self.navegar_para_aba_e_pagina(tab_nome, pagina)
            # --- INÍCIO INSERÇÃO RECONFERÊNCIA DA LINHA DA FATURA ---
            # O índice da linha foi obtido em uma leitura anterior da tabela. Depois
            # de voltar de outra fatura a ordem pode ter mudado (faturas com o mesmo
            # vencimento), então a linha é localizada de novo pelo tipo da fatura.
            try:
                ocorrencias_atuais = self.pesquisar_ocorrencias_na_tela(vencimento_alvo, tab_nome, pagina)
                indices_do_tipo = [
                    oc["linha_index"] for oc in ocorrencias_atuais
                    if oc.get("tipo_fatura") == tipo_fatura
                ]
                if indices_do_tipo and linha_index not in indices_do_tipo:
                    print(
                        f"[AVISO] A fatura {tipo_fatura} mudou da linha {linha_index} "
                        f"para a linha {indices_do_tipo[0]}. Usando a linha atual."
                    )
                    linha_index = indices_do_tipo[0]
                    ocorrencia["linha_index"] = linha_index
            except Exception as e:
                print(f"[AVISO] Não foi possível reconferir a linha da fatura: {self.resumir_erro(e)}")
            # --- FIM INSERÇÃO RECONFERÊNCIA DA LINHA DA FATURA ---

        print(f"Abrindo fatura {tipo_fatura} na linha {linha_index}...")
        linhas = self.page.locator("xpath=//tr")
        linha = linhas.nth(linha_index)

        celula_vencimento = linha.locator(f"xpath=.//*[contains(normalize-space(.), '{vencimento_alvo}')]" ).first
        celula_vencimento.wait_for(state="visible", timeout=self.TIMEOUT_PADRAO)

        self.clicar_e_aguardar(celula_vencimento, "célula do vencimento", usar_networkidle=False)

        print("Executando TAB 1x + ENTER para entrar na fatura...")
        self.page.keyboard.press("Tab")
        time.sleep(0.4)
        self.page.keyboard.press("Enter")

        # Clique seguido de espera explícita pela próxima tela.
        self.aguardar_selector(
            "xpath=//*[contains(normalize-space(.), 'Detalhes da fatura') or contains(normalize-space(.), 'Relatórios') or contains(normalize-space(.), 'Relatorios')]",
            descricao="tela Detalhes da fatura/Relatórios",
            timeout=self.TIMEOUT_LONGO,
        )
        self.respirar_sistema(timeout=20_000, segundos_fallback=2)

    def aguardar_tela_fatura(self):
        try:
            self.aguardar_selector(
                "xpath=//*[contains(normalize-space(.), 'Detalhes da fatura') or contains(normalize-space(.), 'Relatórios') or contains(normalize-space(.), 'Relatorios')]",
                descricao="tela da fatura",
                timeout=self.TIMEOUT_LONGO,
            )
            self.respirar_sistema(timeout=20_000, segundos_fallback=2)
            return True
        except Exception:
            print("[AVISO] Tela da fatura não carregou no tempo esperado.")
            return False

    # --- INÍCIO INSERÇÃO RETENTATIVA INDIVIDUAL DE DOWNLOAD ---
    def limpar_mensagens_documento_retentativa(self, documento):
        """Remove a mensagem transitória da tentativa anterior antes de repetir o documento."""
        doc = self.normalizar_documento_controle(documento)
        prefixo = f"{doc}:"
        self.mensagens_fatura_atual = [
            mensagem
            for mensagem in self.mensagens_fatura_atual
            if not str(mensagem).strip().upper().startswith(prefixo)
        ]

    # --- INÍCIO INSERÇÃO FLUXO 7 ETAPAS DO BOLETO ---
    def obter_seletores_boleto(self):
        """Seletores conhecidos do botão/atalho Boleto.pdf, em ordem de prioridade."""
        return [
            "xpath=//*[normalize-space()='Boleto.pdf']/ancestor-or-self::*[self::button or self::a or @role='button'][1]",
            "xpath=//*[self::button or self::a or self::p or self::span][normalize-space()='Boleto.pdf']",
            "xpath=//*[self::button or self::a or self::p or self::span][contains(normalize-space(.), 'Boleto.pdf')]",
            "xpath=//*[contains(normalize-space(.), 'Boleto.pdf')]/ancestor-or-self::*[self::button or self::a or @role='button'][1]",
            "xpath=//*[self::button or self::a or @role='button'][contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'boleto')]",
            "xpath=//*[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'documento de cobrança')]/ancestor-or-self::*[self::button or self::a or @role='button'][1]",
            "xpath=//*[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'documento de cobranca')]/ancestor-or-self::*[self::button or self::a or @role='button'][1]",
            "xpath=/html/body/div/div/div/main/div/div/main/div/div[1]/div/div[10]/div/div/p",
            "xpath=/html/body/div/div/div/main/div/div/main/div/div[1]/div/div[11]/div/div/p",
        ]

    def localizar_botao_boleto_rapido(self, timeout_total_ms=None, page=None):
        """
        Procura todos os seletores sem somar um timeout para cada XPath.
        O limite informado vale para a busca inteira.
        """
        page = page or self.page
        timeout_total_ms = (
            self.timeout_botao_boleto_total
            if timeout_total_ms is None
            else max(0, int(timeout_total_ms))
        )
        limite = time.perf_counter() + (timeout_total_ms / 1000.0)
        seletores = self.obter_seletores_boleto()

        while time.perf_counter() <= limite:
            for seletor in seletores:
                try:
                    candidatos = page.locator(seletor)
                    quantidade = min(candidatos.count(), 10)
                    for indice in range(quantidade):
                        candidato = candidatos.nth(indice)
                        if candidato.is_visible() and candidato.is_enabled():
                            print(f"Boleto.pdf localizado pelo seletor rápido: {seletor}")
                            return candidato
                except Exception:
                    continue
            if timeout_total_ms <= 0:
                break
            time.sleep(0.20)
        return None

    def localizar_pdf_boleto_gerado_recentemente(self, novo_nome_base, inicio_tentativa):
        """Verifica se o PDF apareceu na pasta mesmo sem o evento Playwright ser capturado."""
        try:
            pasta = self.obter_pasta_download_atual()
            # Com o nome original do portal, o arquivo não segue o nome padrão do
            # robô: considera qualquer PDF da pasta gravado após o clique.
            candidatos = sorted(
                pasta.glob("*.[pP][dD][fF]"),
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
            for caminho in candidatos:
                try:
                    if caminho.stat().st_mtime < (inicio_tentativa - 1.0):
                        continue
                    if caminho.stat().st_size < 1000:
                        continue
                    with open(caminho, "rb") as arquivo:
                        if arquivo.read(5) != b"%PDF-":
                            continue
                    print(
                        "PDF do boleto identificado na pasta após o clique, "
                        f"mesmo sem evento de download: {caminho.name}"
                    )
                    return caminho
                except Exception:
                    continue
        except Exception:
            pass
        return None

    def aguardar_tela_fatura_rapido(self, timeout_ms=None):
        """Confirma a tela de detalhes sem aplicar os respiros longos do fluxo normal."""
        timeout_ms = timeout_ms or self.timeout_recarregar_fatura_boleto
        seletor = (
            "xpath=//*[contains(normalize-space(.), 'Detalhes da fatura') "
            "or contains(normalize-space(.), 'Relatórios') "
            "or contains(normalize-space(.), 'Relatorios')]"
        )
        try:
            self.page.locator(seletor).first.wait_for(
                state="visible",
                timeout=timeout_ms,
            )
            return True
        except Exception:
            return False

    def recarregar_mesma_fatura_para_retentativa_boleto(self):
        """2º nível: recarrega a mesma tela de detalhes e aguarda Boleto.pdf novamente."""
        print("[RETENTATIVA BOLETO - NÍVEL 2] Recarregando a mesma página da fatura...")
        self.fechar_paginas_auxiliares_download()
        pagina = self.page
        pagina.bring_to_front()
        pagina.reload(
            wait_until="domcontentloaded",
            timeout=self.TIMEOUT_LONGO,
        )
        try:
            pagina.wait_for_function(
                "() => document.readyState === 'complete'",
                timeout=self.timeout_recarregar_fatura_boleto,
            )
        except Exception:
            pass

        if not self.aguardar_tela_fatura_rapido():
            raise RuntimeError("A tela de detalhes da fatura não reapareceu após o recarregamento.")

        boleto = self.localizar_botao_boleto_rapido(
            timeout_total_ms=self.timeout_botao_boleto_total,
            page=pagina,
        )
        if boleto is None:
            raise RuntimeError("Boleto.pdf não reapareceu após recarregar a mesma fatura.")

        print("[RETENTATIVA BOLETO - NÍVEL 2] Página recarregada e Boleto.pdf disponível.")
        return True
    # --- FIM INSERÇÃO FLUXO 7 ETAPAS DO BOLETO ---

    def fechar_paginas_auxiliares_download(self):
        pagina_principal = self.page
        try:
            if self.context:
                for pagina_aberta in list(self.context.pages):
                    if pagina_principal and pagina_aberta == pagina_principal:
                        continue
                    try:
                        if not pagina_aberta.is_closed():
                            pagina_aberta.close()
                    except Exception:
                        pass
        except Exception:
            pass

        try:
            if pagina_principal and not pagina_principal.is_closed():
                pagina_principal.bring_to_front()
        except Exception:
            pass

    def reabrir_fatura_atual_para_retentativa_boleto(self):
        """Volta ao extrato, relocaliza a fatura e entra novamente antes do novo boleto."""
        ocorrencia_original = dict(self.ocorrencia_fatura_atual or {})
        if not ocorrencia_original:
            raise RuntimeError("Ocorrência da fatura atual não está disponível para reabertura.")

        print("[RETENTATIVA BOLETO] Voltando completamente ao Extrato Financeiro...")
        self.fechar_paginas_auxiliares_download()

        try:
            self.voltar_para_extrato()
        except Exception as erro_voltar:
            print(
                f"[AVISO] Retorno normal ao extrato falhou: "
                f"{self.resumir_erro(erro_voltar)}. Tentando voltar pela navegação."
            )
            try:
                self.page.go_back(wait_until="domcontentloaded", timeout=self.TIMEOUT_LONGO)
            except Exception:
                pass
            self.aguardar_selector(
                self.SELETOR_TELA_EXTRATO,
                descricao="Extrato Financeiro para retentativa do boleto",
                timeout=self.TIMEOUT_LONGO,
            )

        tab_nome = ocorrencia_original.get("tab", "ABERTO")
        pagina = int(ocorrencia_original.get("pagina", 1) or 1)
        tipo_fatura = ocorrencia_original.get("tipo_fatura", self.tipo_fatura_atual)

        print(
            f"[RETENTATIVA BOLETO] Relocalizando fatura {tipo_fatura} "
            f"na aba {tab_nome}, página {pagina}..."
        )
        if not self.navegar_para_aba_e_pagina(tab_nome, pagina):
            raise RuntimeError("Não foi possível reabrir a aba/página da fatura para a retentativa do boleto.")

        ocorrencias_novas = self.pesquisar_ocorrencias_na_tela(
            self.vencimento_atual,
            tab_nome,
            pagina,
        )
        ocorrencia_atualizada = next(
            (oc for oc in ocorrencias_novas if oc.get("tipo_fatura") == tipo_fatura),
            None,
        )
        if not ocorrencia_atualizada:
            ocorrencia_atualizada = next(iter(ocorrencias_novas), None)
        if not ocorrencia_atualizada:
            raise RuntimeError("A fatura não foi localizada novamente no extrato.")

        self.ocorrencia_fatura_atual = dict(ocorrencia_atualizada)
        # A aba e a página já foram posicionadas acima; não navegamos duas vezes.
        self.abrir_fatura_por_ocorrencia(
            ocorrencia_atualizada,
            self.vencimento_atual,
            navegar=False,
        )
        if self.localizar_botao_boleto_rapido(
            timeout_total_ms=self.timeout_botao_boleto_total,
            page=self.page,
        ) is None:
            raise RuntimeError(
                "A fatura foi reaberta pelo Extrato Financeiro, mas Boleto.pdf não ficou disponível."
            )
        print("[RETENTATIVA BOLETO - NÍVEL 3] Fatura reaberta e Boleto.pdf disponível.")
        return True

    def preparar_retentativa_download(self, documento, tentativa_atual, total_tentativas):
        """
        Prepara a nova tentativa conforme o documento.

        BOLETO:
        - tentativa 2: recarrega a mesma página da fatura;
        - tentativa 3: volta ao Extrato Financeiro e reabre a fatura.

        NF:
        - recarrega a tela da fatura antes da nova tentativa.
        """
        doc = self.normalizar_documento_controle(documento)
        proxima_tentativa = tentativa_atual + 1
        print(
            f"[RETENTATIVA] Preparando nova tentativa de {doc} "
            f"({proxima_tentativa}/{total_tentativas})..."
        )

        self.fechar_paginas_auxiliares_download()
        if self.intervalo_retentativa_download > 0:
            time.sleep(self.intervalo_retentativa_download)

        preparado = False
        try:
            if doc == "BOLETO" and proxima_tentativa == 2:
                preparado = self.recarregar_mesma_fatura_para_retentativa_boleto()

            elif doc == "BOLETO" and proxima_tentativa >= 3:
                print(
                    "[RETENTATIVA BOLETO - NÍVEL 3] "
                    "Voltando ao Extrato Financeiro e reabrindo a fatura..."
                )
                preparado = self.reabrir_fatura_atual_para_retentativa_boleto()

            elif doc.startswith("ANALITICO_"):
                # No portal, recarregar (F5) a tela de detalhes devolve o usuário à
                # lista de faturas. Por isso a fatura é reaberta pelo extrato.
                print(f"[RETENTATIVA] Reabrindo a fatura pelo extrato antes de repetir {doc}...")
                preparado = self.reabrir_fatura_atual_para_analitico()

            elif self.recarregar_tela_retentativa_download:
                pagina_principal = self.page
                print(f"[RETENTATIVA] Recarregando a tela da fatura antes de repetir {doc}...")
                pagina_principal.reload(
                    wait_until="domcontentloaded",
                    timeout=self.TIMEOUT_LONGO,
                )
                preparado = self.aguardar_tela_fatura_rapido(
                    timeout_ms=self.timeout_recarregar_fatura_boleto,
                )

        except Exception as e:
            preparado = False
            mensagem = self.resumir_erro(e)
            print(f"[AVISO] Preparação da tentativa {proxima_tentativa} de {doc} falhou: {mensagem}")
            self.registrar_diagnostico_falha(
                doc,
                "PREPARAR_RETENTATIVA_DOWNLOAD",
                mensagem,
                excecao=e,
            )

        if not preparado:
            print(
                f"[AVISO] A tela não pôde ser confirmada para a tentativa "
                f"{proxima_tentativa}/{total_tentativas} de {doc}."
            )

        self.limpar_mensagens_documento_retentativa(doc)
        self.status_docs_fatura_atual[doc] = "PENDENTE"

    def executar_download_com_retentativas(self, documento, funcao_download, *args, **kwargs):
        """
        Executa o download de um documento de forma independente.

        O arquivo que já foi concluído não é repetido. Somente o documento que
        retornou False é tentado novamente. A função chamada deve retornar True
        quando o arquivo foi salvo ou já constava no controle.
        """
        doc = self.normalizar_documento_controle(documento)
        if doc == "BOLETO":
            total_tentativas = self.tentativas_download_boleto
        elif doc == "NF":
            total_tentativas = self.tentativas_download_nf
        else:
            total_tentativas = self.tentativas_download_arquivo
        ultimo_erro = ""

        for tentativa in range(1, total_tentativas + 1):
            self.documento_download_atual = doc
            self.tentativa_download_atual = tentativa
            # Mantém o diagnóstico da tentativa anterior até sabermos se a nova
            # tentativa recuperou a falha. Se falhar novamente, ele será substituído.

            if tentativa == 1:
                print(f"[{doc}] Tentativa {tentativa}/{total_tentativas}...")
            else:
                self.resumo_execucao["retentativas_download"] += 1
                self.preparar_retentativa_download(doc, tentativa - 1, total_tentativas)
                print(f"[{doc}] Nova tentativa {tentativa}/{total_tentativas}...")

            inicio_tentativa = time.perf_counter()
            sucesso = False
            try:
                sucesso = bool(funcao_download(*args, **kwargs))
            except Exception as e:
                ultimo_erro = self.resumir_erro(e)
                print(f"[AVISO] Exceção na tentativa {tentativa} de {doc}: {ultimo_erro}")
                self.registrar_erro(
                    f"Tentativa {tentativa}/{total_tentativas} do download {doc}",
                    ultimo_erro,
                )
                self.registrar_diagnostico_se_ausente(
                    doc,
                    "EXCECAO_DOWNLOAD",
                    ultimo_erro,
                    excecao=e,
                )
                sucesso = False
            finally:
                duracao_tentativa = time.perf_counter() - inicio_tentativa
                diagnostico_tentativa = self.falha_atual_documento.get(doc) or {}
                detalhes_tentativa = f"Tentativa {tentativa}/{total_tentativas}"
                if diagnostico_tentativa:
                    detalhes_tentativa += (
                        f" | {diagnostico_tentativa.get('codigo', '')} | "
                        f"{diagnostico_tentativa.get('causa_provavel', '')}"
                    )
                self.registrar_tempo_etapa(
                    f"DOWNLOAD_{doc}_TENTATIVA_{tentativa}",
                    duracao_tentativa,
                    "OK" if sucesso else "FALHA",
                    detalhes_tentativa,
                )

            if not sucesso:
                self.registrar_diagnostico_se_ausente(
                    doc,
                    "RESULTADO_SEM_ARQUIVO",
                    f"A tentativa {tentativa}/{total_tentativas} terminou sem produzir arquivo.",
                )

            if sucesso:
                if tentativa > 1:
                    self.resumo_execucao["arquivos_recuperados_retentativa"] += 1
                    self.limpar_mensagens_documento_retentativa(doc)
                    self.marcar_documento_fatura(
                        doc,
                        "OK",
                        f"Arquivo recuperado com sucesso na tentativa {tentativa}/{total_tentativas}.",
                    )
                    print(
                        f"[RETENTATIVA] {doc} recuperado com sucesso na "
                        f"tentativa {tentativa}/{total_tentativas}."
                    )
                    self.registrar_recuperacao_apos_diagnostico(doc, tentativa)
                return True

            if tentativa < total_tentativas:
                print(
                    f"[AVISO] {doc} não foi concluído na tentativa "
                    f"{tentativa}/{total_tentativas}. Uma nova tentativa será executada."
                )

        self.resumo_execucao["arquivos_falha_apos_retentativas"] += 1
        diagnostico_final = self.falha_atual_documento.get(doc) or {}
        mensagem_final = f"Arquivo não concluído após {total_tentativas} tentativa(s)."
        if diagnostico_final:
            mensagem_final += (
                f" Código: {diagnostico_final.get('codigo', '')}. "
                f"Origem provável: {diagnostico_final.get('origem_provavel', '')}. "
                f"Ponto crucial: {diagnostico_final.get('ponto_crucial', '')}. "
                f"Causa provável: {diagnostico_final.get('causa_provavel', '')}."
            )
        elif ultimo_erro:
            mensagem_final += f" Último erro: {ultimo_erro}"
        self.marcar_documento_fatura(doc, "ERRO", mensagem_final)
        print(f"[ERRO] {doc} não foi concluído após {total_tentativas} tentativa(s).")
        return False
    # --- FIM INSERÇÃO RETENTATIVA INDIVIDUAL DE DOWNLOAD ---

    # ----------------------------------------------------------------------
    # Downloads: Boleto e Nota Fiscal
    # ----------------------------------------------------------------------
    def baixar_boleto(self, contrato, vencimento_alvo, tipo_fatura):
        documento = "BOLETO"
        if self.verificar_documento_baixado(
            contrato,
            vencimento_alvo,
            tipo_fatura,
            documento,
        ):
            print("Boleto já consta no log. Pulando...")
            return True

        print("Iniciando download do boleto...")
        sufixo_tipo = self.obter_sufixo_tipo_fatura(tipo_fatura)
        vencimento_fmt = self.formatar_vencimento_nome_arquivo(vencimento_alvo)
        novo_nome_base = f"{contrato}_BOLETO_VENC_{vencimento_fmt}_{sufixo_tipo}"

        boleto = self.localizar_botao_boleto_rapido(
            timeout_total_ms=self.timeout_botao_boleto_total,
            page=self.page,
        )
        if boleto is None:
            print("[AVISO] Boleto.pdf não encontrado para esta fatura.")
            self.registrar_diagnostico_se_ausente(
                documento,
                "LOCALIZAR_BOTAO_BOLETO",
                (
                    "Boleto.pdf não ficou visível e habilitado dentro do limite "
                    f"total de {self.timeout_botao_boleto_total / 1000:.1f}s."
                ),
                page=self.page,
            )
            self.marcar_documento_fatura(
                documento,
                "NAO_DISPONIVEL",
                "Boleto.pdf não encontrado ou não renderizado na tela.",
            )
            return False

        # Etapas 1 e 2: clicar e aguardar download/popup/Blob entre 5 e 8 segundos.
        inicio_tentativa = time.time()
        caminho = self.baixar_por_click_com_fallback_pdf(
            boleto,
            "Boleto.pdf",
            novo_nome_base,
            extensao_fallback=".pdf",
            registrar_documento=documento,
        )

        # Etapa 3: antes de declarar falha, verificar se o PDF apareceu na pasta.
        if caminho is None:
            caminho = self.localizar_pdf_boleto_gerado_recentemente(
                novo_nome_base,
                inicio_tentativa,
            )
            if caminho is not None:
                self.marcar_documento_fatura(documento, "OK", caminho=caminho)
                self.registrar_documento_log(
                    contrato,
                    vencimento_alvo,
                    tipo_fatura,
                    documento,
                )

        if caminho is None:
            self.registrar_diagnostico_se_ausente(
                documento,
                "GERACAO_EXCEDEU_TEMPO_LIMITE",
                (
                    "Boleto.pdf foi acionado e o robô aguardou a geração durante "
                    f"{self.timeout_geracao_boleto / 1000:.0f} segundos, mas o portal "
                    "não disponibilizou download, popup, Blob válido, URL de PDF "
                    "nem arquivo recente na pasta."
                ),
                page=self.page,
            )
            self.marcar_documento_fatura(
                documento,
                "ERRO",
                (
                    "O portal permaneceu processando o boleto além do limite de "
                    f"{self.timeout_geracao_boleto / 1000:.0f} segundos e não gerou o PDF."
                ),
            )
            return False

        return True

    def baixar_nota_fiscal(self, contrato, vencimento_alvo, tipo_fatura):
        documento = "NF"
        if self.verificar_documento_baixado(contrato, vencimento_alvo, tipo_fatura, documento):
            print("Nota fiscal já consta no log. Pulando...")
            return True

        print("Iniciando download da nota fiscal...")
        sufixo_tipo = self.obter_sufixo_tipo_fatura(tipo_fatura)
        vencimento_fmt = self.formatar_vencimento_nome_arquivo(vencimento_alvo)

        btn_nf = self.localizar_primeiro(
            "Acessar Nota Fiscal",
            [
                "xpath=//*[normalize-space()='Acessar Nota Fiscal']/ancestor-or-self::*[self::button or self::a or @role='button'][1]",
                "xpath=//*[self::button or self::a or self::span][normalize-space()='Acessar Nota Fiscal']",
                "xpath=//*[self::button or self::a or self::span][contains(normalize-space(.), 'Acessar Nota Fiscal')]",
                "xpath=//*[contains(normalize-space(.), 'Acessar Nota Fiscal')]/ancestor-or-self::*[self::button or self::a or @role='button'][1]",
                "xpath=/html/body/div/div/div/main/div/div/main/div/div[1]/div/div[11]/div/button/span",
                "xpath=/html/body/div/div/div/main/div/div/main/div/div[1]/div/div[12]/div/button/span",
                "xpath=//*[self::button or self::a or @role='button'][contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'nota fiscal')]",
                "xpath=//*[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'nota fiscal')]/ancestor-or-self::*[self::button or self::a or @role='button'][1]",
                "xpath=//*[self::button or self::a or @role='button'][contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'nfs-e')]",
                "xpath=//*[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'visualizar nota')]/ancestor-or-self::*[self::button or self::a or @role='button'][1]",
                "xpath=//*[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'baixar nota')]/ancestor-or-self::*[self::button or self::a or @role='button'][1]",
            ],
            timeout_por_selector=8_000,
        )
        if not btn_nf:
            print("[AVISO] Botão 'Acessar Nota Fiscal' não encontrado para esta fatura.")
            self.registrar_diagnostico_se_ausente(
                documento,
                "LOCALIZAR_ACESSO_NOTA_FISCAL",
                "Botão Acessar Nota Fiscal não encontrado na tela da fatura.",
            )
            self.marcar_documento_fatura(
                documento,
                "NAO_DISPONIVEL",
                "Botão Acessar Nota Fiscal não encontrado na tela.",
            )
            return False

        pagina_fatura = self.page
        url_fatura = pagina_fatura.url
        pagina_nf = None
        abriu_popup = False

        def descrever_foco(pagina):
            try:
                return pagina.evaluate(
                    """
                    () => {
                        const el = document.activeElement;
                        if (!el) return {};
                        return {
                            tag: (el.tagName || '').toLowerCase(),
                            id: el.id || '',
                            name: el.getAttribute ? (el.getAttribute('name') || '') : '',
                            type: el.getAttribute ? (el.getAttribute('type') || '') : '',
                            value: ('value' in el ? (el.value || '') : ''),
                            texto: ((el.innerText || el.textContent || el.value || '') + '')
                                .replace(/\\s+/g, ' ').trim().slice(0, 180),
                            disabled: !!el.disabled
                        };
                    }
                    """
                )
            except Exception:
                return {}

        def texto_foco(dados):
            if not dados:
                return "não identificado"
            partes = [
                str(dados.get("tag", "")),
                f"id={dados.get('id', '')}" if dados.get("id") else "",
                f"name={dados.get('name', '')}" if dados.get("name") else "",
                f"texto={dados.get('texto', '')}" if dados.get("texto") else "",
            ]
            return " | ".join(parte for parte in partes if parte) or "não identificado"

        def registrar_nf_concluida(caminho):
            if not caminho:
                return False
            self.marcar_documento_fatura(documento, "OK", caminho=caminho)
            self.registrar_documento_log(
                self.contrato_atual,
                self.vencimento_atual,
                self.tipo_fatura_atual,
                documento,
            )
            return True

        try:
            print("Clicando em Acessar Nota Fiscal e aguardando popup ou navegação...")
            try:
                with pagina_fatura.expect_popup(timeout=self.timeout_popup_nf) as popup_info:
                    self.clicar_e_aguardar(
                        btn_nf,
                        "Acessar Nota Fiscal",
                        page=pagina_fatura,
                        usar_networkidle=False,
                    )
                pagina_nf = popup_info.value
                abriu_popup = True
                pagina_nf.set_default_timeout(self.TIMEOUT_PADRAO)
                pagina_nf.set_default_navigation_timeout(self.TIMEOUT_LONGO)
            except PlaywrightTimeoutError:
                pagina_nf = pagina_fatura
                abriu_popup = False

            if pagina_nf == pagina_fatura and pagina_fatura.url == url_fatura:
                print("[AVISO] Nota fiscal não abriu nova aba nem mudou a URL.")
                self.registrar_diagnostico_se_ausente(
                    documento,
                    "ABRIR_PORTAL_NF_NAO_ABRIU",
                    "O clique em Acessar Nota Fiscal não abriu popup e não alterou a URL.",
                    page=pagina_fatura,
                )
                self.marcar_documento_fatura(
                    documento,
                    "ERRO",
                    "Nota fiscal não abriu nova aba nem mudou a URL.",
                )
                return False

            pagina_nf.bring_to_front()

            try:
                pagina_nf.wait_for_load_state(
                    "domcontentloaded",
                    timeout=self.timeout_carregamento_completo_nf,
                )
            except Exception:
                pass

            try:
                pagina_nf.wait_for_function(
                    "() => document.readyState === 'complete'",
                    timeout=self.timeout_carregamento_completo_nf,
                )
                print("Página da Prefeitura carregada completamente.")
            except Exception as erro_ready:
                print(
                    "[AVISO] A página da Prefeitura não confirmou readyState=complete "
                    f"no limite configurado: {self.resumir_erro(erro_ready)}"
                )
                # A tentativa continuará, pois alguns portais mantêm recursos pendentes
                # mesmo quando os botões já estão utilizáveis.

            time.sleep(0.25)

            numero_nf = self.obter_numero_nf_url(pagina_nf)
            novo_nome_base = (
                f"{contrato}_NF_{numero_nf}_VENC_{vencimento_fmt}_{sufixo_tipo}"
            )

            # ==============================================================
            # MÉTODO PRINCIPAL CONFIRMADO MANUALMENTE:
            # CLIQUE EM ÁREA NEUTRA + TAB 1x + ENTER
            # ==============================================================
            print("Método principal da NF: clique neutro + TAB 1x + ENTER...")
            foco_antes = descrever_foco(pagina_nf)
            print(f"Foco antes do clique neutro: {texto_foco(foco_antes)}")

            caminho = None
            erro_teclado = ""
            try:
                # O clique estabelece um ponto inicial previsível para a navegação
                # por teclado sem acionar nenhum dos botões da NFS-e.
                pagina_nf.locator("body").click(
                    position={
                        "x": self.clique_neutro_nf_x,
                        "y": self.clique_neutro_nf_y,
                    },
                    force=True,
                    timeout=3_000,
                )
                print(
                    "Clique neutro realizado na página da Prefeitura em "
                    f"({self.clique_neutro_nf_x}, {self.clique_neutro_nf_y})."
                )
                if self.intervalo_tab_nf:
                    time.sleep(self.intervalo_tab_nf)

                pagina_nf.keyboard.press("Tab")
                if self.intervalo_tab_nf:
                    time.sleep(self.intervalo_tab_nf)

                foco_depois = descrever_foco(pagina_nf)
                print(f"Foco após TAB 1x: {texto_foco(foco_depois)}")
                print("ENTER para disparar o download da NFS-e...")

                with pagina_nf.expect_download(
                    timeout=self.timeout_download_nf_teclado
                ) as download_info:
                    pagina_nf.keyboard.press("Enter")

                download = download_info.value
                caminho = self.salvar_download(
                    download,
                    novo_nome_base,
                    extensao_fallback=".pdf",
                )
                if registrar_nf_concluida(caminho):
                    print(
                        "Nota fiscal baixada com sucesso por "
                        "clique neutro + TAB 1x + ENTER."
                    )
                    return True

            except PlaywrightTimeoutError as e:
                erro_teclado = self.resumir_erro(e)
                print(
                    "[AVISO] Clique neutro + TAB 1x + ENTER não disparou "
                    "download no limite. Executando contingência pelo botão visível..."
                )
            except Exception as e:
                erro_teclado = self.resumir_erro(e)
                print(
                    "[AVISO] Falha no método clique neutro + TAB 1x + ENTER: "
                    f"{erro_teclado}. Executando contingência pelo botão visível..."
                )

            # ==============================================================
            # CONTINGÊNCIA: localizar individualmente o botão DOWNLOAD NFS-e
            # Sem XPath combinado com .first, que poderia escolher item oculto.
            # ==============================================================
            seletores_download_nf = [
                "xpath=/html/body/form/div[3]/div/div[2]/div/input[1]",
                "xpath=//input[contains(translate(@value, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download nfs-e')]",
                "xpath=//input[contains(translate(@value, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download nfs')]",
                "xpath=//input[contains(translate(@value, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download')]",
                "xpath=//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download nfs-e')]",
                "xpath=//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download')]",
                "xpath=//a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download nfs-e')]",
                "xpath=//a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download')]",
            ]

            btn_download_nf = None
            seletor_encontrado = ""
            limite_busca = time.perf_counter() + max(1.0, self.timeout_botao_nf / 1000.0)

            while time.perf_counter() < limite_busca and btn_download_nf is None:
                for seletor in seletores_download_nf:
                    try:
                        candidatos = pagina_nf.locator(seletor)
                        quantidade = min(candidatos.count(), 10)
                        for indice in range(quantidade):
                            candidato = candidatos.nth(indice)
                            if candidato.is_visible() and candidato.is_enabled():
                                btn_download_nf = candidato
                                seletor_encontrado = seletor
                                break
                        if btn_download_nf is not None:
                            break
                    except Exception:
                        continue
                if btn_download_nf is None:
                    time.sleep(0.25)

            if btn_download_nf is not None:
                print(
                    "Botão Download NFS-e localizado pela contingência: "
                    f"{seletor_encontrado}"
                )
                caminho = self.baixar_por_click_com_fallback_pdf(
                    btn_download_nf,
                    "Download NFS-e - contingência após clique neutro + TAB 1x + ENTER",
                    novo_nome_base,
                    page=pagina_nf,
                    extensao_fallback=".pdf",
                    registrar_documento=documento,
                )
                if caminho is not None:
                    print("Nota fiscal baixada pela contingência do botão visível.")
                    return True

            # Falha definitiva desta tentativa: teclado e contingência não concluíram.
            foco_final = descrever_foco(pagina_nf)
            mensagem_falha = (
                "A página da Prefeitura abriu, mas clique neutro + TAB 1x + ENTER não gerou "
                "download e a contingência pelo botão Download NFS-e também "
                "não conseguiu salvar o PDF. "
                f"Foco final: {texto_foco(foco_final)}. "
                f"Erro do teclado: {erro_teclado or 'não informado'}."
            )
            print(f"[AVISO] {mensagem_falha}")
            self.registrar_diagnostico_se_ausente(
                documento,
                "DOWNLOAD_NF_CLIQUE_TAB_1X_E_CONTINGENCIA",
                mensagem_falha,
                page=pagina_nf,
            )
            self.marcar_documento_fatura(documento, "ERRO", mensagem_falha)
            self.salvar_screenshot(
                "nf_clique_tab_1x_enter_e_contingencia_sem_download",
                page=pagina_nf,
            )
            return False

        except Exception as e:
            print(f"[AVISO] Erro ao baixar nota fiscal: {self.resumir_erro(e)}")
            self.registrar_diagnostico_se_ausente(
                documento,
                "EXCECAO_DOWNLOAD_NF",
                self.resumir_erro(e),
                page=pagina_nf or pagina_fatura,
                excecao=e,
            )
            self.marcar_documento_fatura(documento, "ERRO", self.resumir_erro(e))
            self.registrar_erro(
                "Falha no download da nota fiscal",
                self.resumir_erro(e),
            )
            self.salvar_screenshot(
                "erro_nota_fiscal",
                page=pagina_nf or pagina_fatura,
            )
            return False
        finally:
            try:
                if abriu_popup and pagina_nf and not pagina_nf.is_closed():
                    pagina_nf.close()
            except Exception:
                pass
            try:
                self.page = pagina_fatura
                if not abriu_popup and pagina_fatura.url != url_fatura:
                    pagina_fatura.go_back(
                        wait_until="domcontentloaded",
                        timeout=self.TIMEOUT_LONGO,
                    )
                    self.respirar_sistema(
                        page=pagina_fatura,
                        timeout=20_000,
                        segundos_fallback=2,
                    )
                pagina_fatura.bring_to_front()
            except Exception:
                pass

    def executar_etapa_download_cronometrada(self, etapa, funcao, *args, **kwargs):
        """Cronometra o download e marca OK, FALHA ou ERRO conforme o retorno real."""
        inicio = time.perf_counter()
        sucesso = False
        status = "FALHA"
        detalhes = ""
        print(f"[CRONÔMETRO] Início: {etapa}")
        try:
            sucesso = bool(funcao(*args, **kwargs))
            status = "OK" if sucesso else "FALHA"
            if not sucesso:
                doc = etapa.replace("DOWNLOAD_", "").strip().upper()
                diagnostico = self.falha_atual_documento.get(doc) or {}
                detalhes = "Arquivo não concluído após as tentativas configuradas."
                if diagnostico:
                    detalhes += (
                        f" Código: {diagnostico.get('codigo', '')}. "
                        f"Ponto crucial: {diagnostico.get('ponto_crucial', '')}."
                    )
            return sucesso
        except Exception as e:
            status = "ERRO"
            detalhes = self.resumir_erro(e)
            raise
        finally:
            duracao = time.perf_counter() - inicio
            self.registrar_tempo_etapa(etapa, duracao, status, detalhes)
            print(f"[CRONÔMETRO] Fim: {etapa} | {self.formatar_duracao(duracao)} | {status}")

    def baixar_documentos_fatura(self, contrato, vencimento_alvo, tipo_fatura):
        print(
            "Baixando documentos da fatura: relatórios analíticos "
            f"({', '.join(self.formatos_analiticos)})"
            f"{' + boleto e nota fiscal' if self.baixar_boleto_nf else ''}..."
        )
        resultados = []

        # --- INÍCIO BLOCO OPCIONAL BOLETO E NOTA FISCAL (BAIXAR_BOLETO_NF=1) ---
        if self.baixar_boleto_nf:
            print(
                "Retentativas automáticas configuradas: "
                f"BOLETO até {self.tentativas_download_boleto} tentativa(s) "
                f"(aguarda até {self.timeout_geracao_boleto / 1000:.0f}s por tentativa; "
                "a última ocorre após recarregar a mesma fatura); "
                f"NF até {self.tentativas_download_nf} tentativa(s)."
            )

            sucesso_boleto = self.executar_etapa_download_cronometrada(
                "DOWNLOAD_BOLETO",
                self.executar_download_com_retentativas,
                "BOLETO",
                self.baixar_boleto,
                contrato,
                vencimento_alvo,
                tipo_fatura,
            )
            resultados.append(sucesso_boleto)

            sucesso_nf = self.executar_etapa_download_cronometrada(
                "DOWNLOAD_NOTA_FISCAL",
                self.executar_download_com_retentativas,
                "NF",
                self.baixar_nota_fiscal,
                contrato,
                vencimento_alvo,
                tipo_fatura,
            )
            resultados.append(sucesso_nf)
        # --- FIM BLOCO OPCIONAL BOLETO E NOTA FISCAL ---

        # --- INÍCIO INSERÇÃO DOWNLOAD DOS RELATÓRIOS ANALÍTICOS ---
        print(
            "Relatórios analíticos: até "
            f"{self.tentativas_download_arquivo} tentativa(s) por arquivo, "
            f"aguardando até {self.timeout_download_analitico / 1000:.0f}s em cada uma."
        )
        for formato in self.formatos_analiticos:
            documento = f"ANALITICO_{formato}"
            # Cada formato é independente: a falha de um não impede os demais.
            sucesso = self.executar_etapa_download_cronometrada(
                f"DOWNLOAD_{documento}",
                self.executar_download_com_retentativas,
                documento,
                self.baixar_relatorio_analitico,
                contrato,
                vencimento_alvo,
                tipo_fatura,
                formato,
            )
            resultados.append(sucesso)
        # --- FIM INSERÇÃO DOWNLOAD DOS RELATÓRIOS ANALÍTICOS ---

        # Mantém o fluxo da fatura ativo quando pelo menos um documento concluiu.
        # A classificação completa/parcial é feita separadamente no contrato.
        return any(resultados)

    # ----------------------------------------------------------------------
    # Downloads: Relatórios analíticos (Mensalidade e Coparticipação)
    # ----------------------------------------------------------------------
    # --- INÍCIO INSERÇÃO RELATÓRIOS ANALÍTICOS ---
    def obter_tipo_fatura_da_tela(self, page=None):
        """Lê o campo "Tipo de fatura" da tela Detalhes da fatura. Retorna None se não identificar."""
        page = page or self.page
        try:
            texto = page.locator("body").inner_text(timeout=self.TIMEOUT_CURTO).upper()
        except Exception:
            return None
        trecho = re.search(r"TIPO DE FATURA\s*([^\n]*(?:\n[^\n]*)?)", texto)
        if not trecho:
            return None
        valor = re.sub(r"\s+", " ", trecho.group(1))
        for tipo in (self.TIPO_CARNET, self.TIPO_MENSALIDADE):
            if tipo in valor:
                return tipo
        return None

    def conferir_tipo_da_fatura_aberta(self, tipo_esperado):
        """
        Garante que a fatura aberta é do tipo que o robô pretendia abrir.
        Se a tela mostrar outro tipo, reabre a fatura pela lista; se continuar
        errada, interrompe esta fatura com mensagem clara em vez de baixar
        (ou procurar) relatórios na fatura errada.
        """
        try:
            url_fatura = str(self.page.url or "")
        except Exception:
            url_fatura = ""
        tipo_tela = self.obter_tipo_fatura_da_tela()
        print(f"Fatura aberta: {url_fatura} | tipo na tela: {tipo_tela or 'não identificado'}")

        if not tipo_tela or tipo_esperado not in self.PRIORIDADE_TIPOS or tipo_tela == tipo_esperado:
            return True

        print(
            f"[AVISO] A fatura aberta é '{tipo_tela}', mas a esperada era '{tipo_esperado}'. "
            "Reabrindo pela lista de faturas..."
        )
        self.reabrir_fatura_atual_para_analitico()
        tipo_tela = self.obter_tipo_fatura_da_tela()
        if tipo_tela and tipo_tela != tipo_esperado:
            raise RuntimeError(
                f"A fatura aberta é '{tipo_tela}', mas a esperada era '{tipo_esperado}'."
            )
        return True

    def rolar_tela_fatura_ate_o_fim(self, page=None):
        """
        Rola a página e os painéis internos até o fim. A seção "Relatórios" da
        coparticipação fica abaixo dos detalhes e pode só ser montada pelo portal
        quando entra na área visível.
        """
        page = page or self.page
        try:
            page.evaluate(
                """
                () => {
                    const alvos = [document.scrollingElement, document.documentElement, document.body]
                        .concat(Array.from(document.querySelectorAll('main, div')));
                    for (const el of alvos) {
                        if (!el) continue;
                        if (el.scrollHeight > el.clientHeight + 20) {
                            el.scrollTop = el.scrollHeight;
                        }
                    }
                    window.scrollTo(0, document.body ? document.body.scrollHeight : 0);
                }
                """
            )
        except Exception:
            pass

    def registrar_tela_sem_botao_analitico(self, documento, formato, page=None):
        """
        Grava em Logs/tela_sem_botao_analitico_ndi.txt o que havia na tela quando
        o botão não foi encontrado: URL, tipo da fatura, todos os botões/links
        (texto, se está visível e se está desabilitado) e o texto da página.
        Retorna um resumo curto da situação para a mensagem do controle.
        """
        page = page or self.page
        fmt = str(formato or "").upper()
        dados = {}
        try:
            dados = page.evaluate(
                """
                () => {
                    const visivel = (el) => {
                        const s = window.getComputedStyle(el);
                        const r = el.getBoundingClientRect();
                        return s.display !== 'none' && s.visibility !== 'hidden' && r.width > 0 && r.height > 0;
                    };
                    const limpar = (t) => (t || '').replace(/\\s+/g, ' ').trim();
                    const botoes = Array.from(document.querySelectorAll('button, a, [role="button"]'))
                        .map((el) => ({
                            tag: el.tagName.toLowerCase(),
                            texto: limpar(el.innerText || el.textContent).slice(0, 80),
                            visivel: visivel(el),
                            desabilitado: !!el.disabled || el.getAttribute('aria-disabled') === 'true',
                            linha: limpar(el.parentElement && el.parentElement.parentElement
                                ? el.parentElement.parentElement.innerText : '').slice(0, 120)
                        }))
                        .filter((b) => b.texto);
                    const principal = document.querySelector('main main') || document.querySelector('main') || document.body;
                    return {
                        url: location.href,
                        botoes: botoes.slice(0, 120),
                        texto: (principal.innerText || '').slice(0, 6000)
                    };
                }
                """
            ) or {}
        except Exception as e:
            dados = {"erro": self.resumir_erro(e)}

        botoes = dados.get("botoes", []) or []
        texto_pagina = str(dados.get("texto", "") or "")
        texto_maiusculo = texto_pagina.upper()

        de_relatorio = [b for b in botoes if "RELAT" in str(b.get("texto", "")).upper()]
        if re.search(rf"{fmt}\s*\(\s*0\s+ARQUIVO", texto_maiusculo):
            resumo = f"A tela informa {fmt} (0 arquivos): o portal não tem relatório {fmt} para esta fatura."
        elif de_relatorio and all(b.get("desabilitado") for b in de_relatorio):
            resumo = "Os botões de relatório existem, mas estão desabilitados pelo portal."
        elif de_relatorio and not any(b.get("visivel") for b in de_relatorio):
            resumo = "Os botões de relatório existem no HTML, mas não estão visíveis."
        elif de_relatorio:
            resumo = (
                f"Há {len(de_relatorio)} botão(ões) de relatório na tela, mas nenhum foi "
                f"associado ao formato {fmt}."
            )
        elif "RELATÓRIOS" in texto_maiusculo or "RELATORIOS" in texto_maiusculo:
            resumo = "A seção Relatórios aparece na tela, mas sem botões de download."
        else:
            resumo = "A tela desta fatura não tem seção Relatórios nem botões de relatório analítico."

        try:
            with TRAVA_ARQUIVOS, open(self.arquivo_tela_sem_botao, "a", encoding="utf-8") as f:
                f.write("=" * 100 + "\n")
                f.write(
                    f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} | CONTRATO: {self.contrato_atual} | "
                    f"VENCIMENTO: {self.vencimento_atual} | TIPO ESPERADO: {self.tipo_fatura_atual} | "
                    f"DOCUMENTO: {documento} | TENTATIVA: {self.tentativa_download_atual}\n"
                )
                f.write(f"URL: {dados.get('url', '')}\n")
                f.write(f"TIPO DE FATURA NA TELA: {self.obter_tipo_fatura_da_tela(page) or 'não identificado'}\n")
                f.write(f"RESUMO: {resumo}\n")
                if dados.get("erro"):
                    f.write(f"ERRO AO LER A TELA: {dados.get('erro')}\n")
                f.write(f"BOTOES E LINKS ({len(botoes)}):\n")
                for b in botoes:
                    f.write(
                        f"  <{b.get('tag')}> '{b.get('texto')}' | visível={b.get('visivel')} | "
                        f"desabilitado={b.get('desabilitado')} | linha: {b.get('linha')}\n"
                    )
                f.write("TEXTO DA TELA:\n")
                f.write(texto_pagina + "\n")
        except Exception as e:
            print(f"[AVISO] Não foi possível gravar o registro da tela: {self.resumir_erro(e)}")

        print(f"Situação da tela: {resumo}")
        return resumo

    def tela_detalhes_fatura_aberta(self, page=None):
        """Confirma que a aba está na tela "Detalhes da fatura" (e não na lista de faturas)."""
        page = page or self.page
        try:
            if "/detalhes/" in str(page.url or "").lower():
                return True
        except Exception:
            pass
        try:
            titulo = page.locator("xpath=//*[normalize-space(.)='Detalhes da fatura']").first
            return bool(titulo.is_visible())
        except Exception:
            return False

    def reabrir_fatura_atual_para_analitico(self):
        """
        Volta à lista de faturas, relocaliza a fatura atual e entra nela de novo.
        Usado na retentativa e sempre que a aba não estiver na tela de detalhes.
        """
        ocorrencia_original = dict(self.ocorrencia_fatura_atual or {})
        if not ocorrencia_original:
            raise RuntimeError("Ocorrência da fatura atual não está disponível para reabertura.")

        self.fechar_paginas_auxiliares_download()
        seletor_extrato = self.SELETOR_TELA_EXTRATO

        if self.tela_detalhes_fatura_aberta():
            try:
                self.page.go_back(wait_until="domcontentloaded", timeout=self.TIMEOUT_LONGO)
            except Exception:
                pass

        # Se ainda não estiver na lista de faturas, abre a URL guardada do extrato.
        if self.tela_detalhes_fatura_aberta() and self.url_extrato_atual:
            self.page.goto(
                self.url_extrato_atual,
                wait_until="domcontentloaded",
                timeout=self.TIMEOUT_LONGO,
            )

        self.aguardar_selector(
            seletor_extrato,
            descricao="lista de faturas para reabrir a fatura",
            timeout=self.TIMEOUT_LONGO,
        )

        tab_nome = ocorrencia_original.get("tab", "ABERTO")
        pagina = int(ocorrencia_original.get("pagina", 1) or 1)
        tipo_fatura = ocorrencia_original.get("tipo_fatura", self.tipo_fatura_atual)

        print(f"Relocalizando fatura {tipo_fatura} na aba {tab_nome}, página {pagina}...")
        if not self.navegar_para_aba_e_pagina(tab_nome, pagina):
            raise RuntimeError("Não foi possível reabrir a aba/página da fatura.")

        ocorrencias_novas = self.pesquisar_ocorrencias_na_tela(self.vencimento_atual, tab_nome, pagina)
        ocorrencia_atualizada = next(
            (oc for oc in ocorrencias_novas if oc.get("tipo_fatura") == tipo_fatura),
            None,
        )
        if not ocorrencia_atualizada:
            raise RuntimeError("A fatura não foi localizada novamente na lista de faturas.")

        self.ocorrencia_fatura_atual = dict(ocorrencia_atualizada)
        self.abrir_fatura_por_ocorrencia(ocorrencia_atualizada, self.vencimento_atual, navegar=False)
        if not self.tela_detalhes_fatura_aberta():
            raise RuntimeError("A tela Detalhes da fatura não abriu após a reabertura.")
        print("Fatura reaberta pela lista de faturas.")
        return True

    def focar_relatorio_coparticipacao_por_teclado(self, formato, page=None):
        """
        Coparticipação: 1 clique no título "Relatórios" e TAB até o botão do formato.
            CSV = TAB 1x | TXT = TAB 2x | PDF = TAB 3x

        Deixa o botão em foco e NÃO pressiona ENTER. Antes de liberar o ENTER,
        confere o elemento em foco: precisa ser um botão "Baixar relatório"
        habilitado e, quando a linha indicar o formato, tem de ser o formato
        pedido. Isso impede que o ENTER acione outro elemento (por exemplo o
        botão "Voltar") se a ordem de foco da tela for diferente da esperada.

        Retorna (True, descrição do foco) ou (False, motivo).
        """
        page = page or self.page
        fmt = str(formato or "").strip().upper()
        quantidade_tabs = self.TABS_RELATORIO_COPARTICIPACAO.get(fmt)
        if not quantidade_tabs:
            return False, f"Formato {fmt} sem quantidade de TAB definida."

        titulo = None
        limite = time.perf_counter() + (self.timeout_botao_analitico_total / 1000.0)
        seletores_titulo = [
            "xpath=//*[normalize-space(.)='Relatórios']",
            "xpath=//*[normalize-space(.)='Relatorios']",
        ]
        primeira_passagem = True
        while titulo is None and time.perf_counter() <= limite:
            if not primeira_passagem:
                self.rolar_tela_fatura_ate_o_fim(page)
                time.sleep(0.20)
            primeira_passagem = False
            for seletor in seletores_titulo:
                try:
                    candidatos = page.locator(seletor)
                    # O último da lista é o elemento mais interno com esse texto.
                    for indice in range(min(candidatos.count(), 10) - 1, -1, -1):
                        candidato = candidatos.nth(indice)
                        if candidato.is_visible():
                            titulo = candidato
                            break
                except Exception:
                    continue
                if titulo is not None:
                    break
        if titulo is None:
            return False, 'Título "Relatórios" não encontrado na tela da fatura.'

        try:
            titulo.scroll_into_view_if_needed(timeout=self.TIMEOUT_CURTO)
        except Exception:
            pass
        print('Coparticipação: 1 clique no título "Relatórios"...')
        titulo.click(timeout=self.TIMEOUT_PADRAO)
        if self.intervalo_tab_analitico:
            time.sleep(self.intervalo_tab_analitico)

        # --- INÍCIO INSERÇÃO TAB ADAPTATIVO NA COPARTICIPAÇÃO ---
        # Antes: TAB fixo (CSV 1x, TXT 2x, PDF 3x) e, se o foco não estivesse no
        # botão certo, desistia e ia para o clique. Na execução de 06/10/2026 o
        # CSV caiu no clique em todos os contratos. Agora o robô avança um TAB por
        # vez (até TAB esperado + 3) e para no botão "Baixar relatório" correto,
        # identificado pelo formato da linha ou pela posição do botão abaixo do
        # título "Relatórios" (1º = CSV, 2º = TXT, 3º = PDF). Se nenhum TAB chegar
        # lá, foca o botão certo diretamente e o acionamento continua sendo ENTER.
        ordem_esperada = quantidade_tabs - 1
        maximo_tabs = quantidade_tabs + 3
        script_foco = """
            (titulo) => {
                const el = document.activeElement;
                if (!el || el === document.body) return null;
                const limpar = (t) => (t || '').replace(/\\s+/g, ' ').trim();
                const seletor = 'button, a, [role="button"]';
                let bloco = el;
                let pai = el.parentElement;
                while (pai && pai.querySelectorAll(seletor).length <= 1) {
                    bloco = pai;
                    pai = pai.parentElement;
                }
                const visivel = (b) => {
                    const r = b.getBoundingClientRect();
                    const s = window.getComputedStyle(b);
                    return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
                };
                // Botões "Baixar relatório" que vêm depois do título "Relatórios".
                const botoes = Array.from(document.querySelectorAll(seletor)).filter((b) =>
                    /baixar relat/i.test(limpar(b.innerText || b.textContent)) && visivel(b) &&
                    (titulo.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING)
                );
                return {
                    tag: (el.tagName || '').toLowerCase(),
                    role: el.getAttribute('role') || '',
                    texto: limpar(el.innerText || el.textContent).slice(0, 80),
                    linha: limpar(bloco.innerText || bloco.textContent).slice(0, 120),
                    desabilitado: !!el.disabled || el.getAttribute('aria-disabled') === 'true',
                    ordem: botoes.indexOf(el),
                    total_botoes: botoes.length
                };
            }
        """

        def avaliar_foco(foco):
            """Retorna ('OK' | 'CONTINUAR' | 'PAROU', descrição)."""
            if not foco:
                return "CONTINUAR", "nenhum elemento em foco"
            descricao = (
                f"<{foco.get('tag')}> '{foco.get('texto')}' | linha: '{foco.get('linha')}' | "
                f"posição: {foco.get('ordem')}/{foco.get('total_botoes')}"
            )
            texto_foco = str(foco.get("texto", "")).upper()
            linha_foco = str(foco.get("linha", "")).upper()
            clicavel = foco.get("tag") in {"button", "a"} or foco.get("role") == "button"
            if not clicavel or "BAIXAR RELAT" not in texto_foco:
                return "CONTINUAR", descricao
            if foco.get("desabilitado"):
                return "PAROU", f"O botão em foco está desabilitado: {descricao}"

            formato_linha = ""
            for candidato in self.FORMATOS_ANALITICOS_SUPORTADOS:
                if re.match(rf"^{candidato}(\W|$)", linha_foco):
                    formato_linha = candidato
                    break
            if formato_linha == fmt:
                return "OK", descricao
            if formato_linha:
                ordem_linha = self.TABS_RELATORIO_COPARTICIPACAO.get(formato_linha, 0) - 1
                if ordem_linha > ordem_esperada:
                    return "PAROU", f"O foco passou do relatório {fmt} (está no {formato_linha}): {descricao}"
                return "CONTINUAR", descricao

            ordem = foco.get("ordem", -1)
            if ordem == ordem_esperada:
                return "OK", descricao
            if isinstance(ordem, int) and ordem > ordem_esperada:
                return "PAROU", f"O foco passou do relatório {fmt}: {descricao}"
            return "CONTINUAR", descricao

        print(f"TAB até o relatório {fmt} (esperado {quantidade_tabs}x, máximo {maximo_tabs}x)...")
        ultima_descricao = ""
        for numero_tab in range(1, maximo_tabs + 1):
            page.keyboard.press("Tab")
            if self.intervalo_tab_analitico:
                time.sleep(self.intervalo_tab_analitico)
            try:
                foco = titulo.evaluate(script_foco)
            except Exception as e:
                return False, f"Não foi possível ler o elemento em foco: {self.resumir_erro(e)}"

            situacao, ultima_descricao = avaliar_foco(foco)
            print(f"Foco após TAB {numero_tab}x: {ultima_descricao}")
            if situacao == "OK":
                if numero_tab != quantidade_tabs:
                    print(
                        f"[INFO] Relatório {fmt} alcançado com TAB {numero_tab}x "
                        f"(esperado {quantidade_tabs}x)."
                    )
                return True, f"TAB {numero_tab}x | {ultima_descricao}"
            if situacao == "PAROU":
                break

        # Último recurso do método por teclado: foca diretamente o botão na
        # posição esperada abaixo de "Relatórios" e confere antes do ENTER.
        try:
            focado = titulo.evaluate(
                """
                (titulo, ordem) => {
                    const limpar = (t) => (t || '').replace(/\\s+/g, ' ').trim();
                    const visivel = (b) => {
                        const r = b.getBoundingClientRect();
                        const s = window.getComputedStyle(b);
                        return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
                    };
                    const botoes = Array.from(document.querySelectorAll('button, a, [role="button"]')).filter((b) =>
                        /baixar relat/i.test(limpar(b.innerText || b.textContent)) && visivel(b) &&
                        (titulo.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING)
                    );
                    const alvo = botoes[ordem];
                    if (!alvo) return false;
                    alvo.scrollIntoView({block: 'center'});
                    alvo.focus();
                    return document.activeElement === alvo;
                }
                """,
                ordem_esperada,
            )
            if focado:
                situacao, descricao_direta = avaliar_foco(titulo.evaluate(script_foco))
                print(f"Foco direto no botão {ordem_esperada + 1} de Relatórios: {descricao_direta}")
                if situacao == "OK":
                    return True, f"foco direto | {descricao_direta}"
        except Exception as e:
            print(f"[AVISO] Foco direto no botão do relatório {fmt} falhou: {self.resumir_erro(e)}")

        return False, f"Botão do relatório {fmt} não alcançado pelo teclado. Último foco: {ultima_descricao}"
        # --- FIM INSERÇÃO TAB ADAPTATIVO NA COPARTICIPAÇÃO ---

    def obter_seletores_analitico(self, formato):
        """
        Seletores do botão do relatório analítico, em ordem de prioridade.

        Mapeados a partir da gravação do portal (05/10/2026):

        Layout 1 - MENSALIDADE PLANO SAUDE (dentro de "Detalhes da fatura"):
            Relatório analítico (PDF)   [Relatório.pdf]
            Relatório analítico (CSV)   [Relatório.csv]
            Relatório analítico (TXT)   [Relatório.txt]

        Layout 2 - CARNET COMPLEMENTAR-COPARTICIPACAO (seção "Relatórios"):
            CSV (1 arquivos)   [Baixar relatório]
            TXT (1 arquivos)   [Baixar relatório]
            PDF (1 arquivos)   [Baixar relatório]

        No layout 2 os três botões têm o mesmo texto; por isso o botão é
        localizado pela linha que começa com o formato e contém um único botão.
        Os caminhos CSS posicionais da gravação (div:nth-of-type(14) etc.) não
        são usados porque mudam conforme os campos exibidos na fatura.
        """
        fmt = str(formato or "").strip().upper()
        ext = fmt.lower()
        maiusculas = "ABCDEFGHIJKLMNOPQRSTUVWXYZÓ"
        minusculas = "abcdefghijklmnopqrstuvwxyzó"
        clicavel = "self::button or self::a or @role='button'"
        return [
            # Layout 1: botão identificado pelo próprio texto.
            f"xpath=//*[{clicavel}][normalize-space(.)='Relatório.{ext}']",
            f"xpath=//*[{clicavel}][translate(normalize-space(.), '{maiusculas}', '{minusculas}')='relatório.{ext}']",
            f"xpath=//*[{clicavel}][translate(normalize-space(.), '{maiusculas}', '{minusculas}')='relatorio.{ext}']",
            # Layout 1: botão único da linha "Relatório analítico (FMT)".
            f"xpath=//div[starts-with(normalize-space(.), 'Relatório analítico ({fmt})') and count(.//button)=1]//button",
            f"xpath=//div[starts-with(normalize-space(.), 'Relatorio analitico ({fmt})') and count(.//button)=1]//button",
            # Layout 2 sem depender de espaços: no HTML, "CSV" e "(1 arquivos)" podem
            # estar em elementos separados, e o XPath os enxerga colados ("CSV(1 arquivos)").
            f"xpath=//div[starts-with(translate(normalize-space(.), ' \u00a0', ''), '{fmt}(') and count(.//button)=1]//button",
            # Layout 2: botão único da linha "FMT (N arquivos)".
            f"xpath=//div[starts-with(normalize-space(.), '{fmt} (') and count(.//button)=1]//button",
            f"xpath=//div[starts-with(normalize-space(.), '{fmt} (') and count(.//*[{clicavel}])=1]//*[{clicavel}]",
        ]

    def localizar_botao_relatorio_por_linha(self, formato, page=None):
        """
        Localiza o botão "Baixar relatório" pelo texto VISÍVEL da linha em que ele está.

        Os três botões da seção "Relatórios" têm o mesmo texto. Para cada um, sobe
        na árvore enquanto o bloco contiver um único botão e lê o texto visível
        desse bloco (innerText, que respeita quebras entre elementos), por exemplo
        "CSV (1 arquivos) Baixar relatório". O botão escolhido é o da linha que
        começa com o formato pedido.
        """
        page = page or self.page
        fmt = str(formato or "").strip().upper()
        try:
            botoes = page.locator(
                "xpath=//*[self::button or self::a or @role='button']"
                "[contains(normalize-space(.), 'Baixar relat') or contains(normalize-space(.), 'Relat')]"
            )
            quantidade = min(botoes.count(), 40)
        except Exception:
            return None

        for indice in range(quantidade):
            try:
                botao = botoes.nth(indice)
                if not (botao.is_visible() and botao.is_enabled()):
                    continue
                rotulo = botao.evaluate(
                    """
                    (el) => {
                        const seletor = 'button, a, [role="button"]';
                        let bloco = el;
                        let pai = el.parentElement;
                        while (pai && pai.querySelectorAll(seletor).length <= 1) {
                            bloco = pai;
                            pai = pai.parentElement;
                        }
                        return (bloco.innerText || bloco.textContent || '')
                            .replace(/\\s+/g, ' ').trim();
                    }
                    """
                )
            except Exception:
                continue

            rotulo_maiusculo = str(rotulo or "").upper()
            if re.match(rf"^{fmt}(\W|$)", rotulo_maiusculo) or f"({fmt})" in rotulo_maiusculo:
                print(f"Relatório analítico {fmt} localizado pela linha do botão: '{str(rotulo)[:80]}'")
                return botao
        return None

    def localizar_botao_analitico_rapido(self, formato, timeout_total_ms=None, page=None):
        """
        Procura todos os seletores do formato sem somar um timeout para cada XPath.
        O limite informado vale para a busca inteira.
        """
        page = page or self.page
        timeout_total_ms = (
            self.timeout_botao_analitico_total
            if timeout_total_ms is None
            else max(0, int(timeout_total_ms))
        )
        limite = time.perf_counter() + (timeout_total_ms / 1000.0)
        seletores = self.obter_seletores_analitico(formato)
        primeira_passagem = True

        while time.perf_counter() <= limite:
            if not primeira_passagem:
                # Não achou de primeira: rola até o fim para forçar o portal a
                # montar a seção "Relatórios" e tenta de novo.
                self.rolar_tela_fatura_ate_o_fim(page)
            primeira_passagem = False
            for seletor in seletores:
                try:
                    candidatos = page.locator(seletor)
                    quantidade = min(candidatos.count(), 10)
                    for indice in range(quantidade):
                        candidato = candidatos.nth(indice)
                        if candidato.is_visible() and candidato.is_enabled():
                            print(f"Relatório analítico {formato} localizado pelo seletor: {seletor}")
                            return candidato
                except Exception:
                    continue
            # Contingência independente da estrutura do HTML.
            candidato = self.localizar_botao_relatorio_por_linha(formato, page=page)
            if candidato is not None:
                return candidato
            if timeout_total_ms <= 0:
                break
            time.sleep(0.20)
        return None

    def obter_extensao_por_url(self, url, extensao_fallback):
        """Usa a extensão do arquivo indicado na URL quando ela é conhecida."""
        try:
            ext = Path(urlparse(str(url or "")).path).suffix.lower()
        except Exception:
            ext = ""
        if ext in {".csv", ".txt", ".pdf", ".zip"}:
            return ext
        return extensao_fallback

    def baixar_url_analitico_por_request(self, page_alvo, url, novo_nome_base, extensao_fallback):
        """
        Salva o arquivo que o portal abriu em uma aba (PDF no visualizador ou
        TXT como texto puro), sem depender do botão de salvar do navegador.

        1º: request autenticado do contexto, com a mesma URL (inclui o token).
        2º: fetch executado dentro da própria aba, que é da mesma origem.
        """
        conteudo = None
        headers_resposta = {}

        try:
            resposta = self.context.request.get(url, timeout=self.TIMEOUT_DOWNLOAD)
            if resposta.ok:
                conteudo = resposta.body()
                headers_resposta = resposta.headers
            else:
                print(
                    f"[AVISO] URL do relatório retornou HTTP {resposta.status}: "
                    f"{self.limpar_url_para_log(url)}"
                )
        except Exception as e:
            print(f"[AVISO] Falha no request da URL do relatório: {self.resumir_erro(e)}")

        if not conteudo:
            try:
                if page_alvo and not page_alvo.is_closed():
                    conteudo_lista = page_alvo.evaluate(
                        """
                        async () => {
                            const resp = await fetch(window.location.href);
                            if (!resp.ok) return null;
                            const buffer = await resp.arrayBuffer();
                            return Array.from(new Uint8Array(buffer));
                        }
                        """
                    )
                    if conteudo_lista:
                        conteudo = bytes(conteudo_lista)
            except Exception as e:
                print(f"[AVISO] Falha no fetch dentro da aba do relatório: {self.resumir_erro(e)}")

        if not conteudo:
            return None

        extensao = self.obter_extensao_por_url(url, extensao_fallback)
        nome_original = self.obter_nome_original_resposta(url, headers_resposta)
        return self.salvar_bytes_em_arquivo(conteudo, novo_nome_base, extensao, nome_original=nome_original)

    def capturar_arquivo_analitico_por_clique(self, locator, descricao, novo_nome_base, formato, page=None, acionar=None):
        """
        Clica uma única vez no relatório e rastreia o que o portal fizer.

        Segue a lógica de rastreamento do gravador de ações: observa a aba de
        origem e todas as abas criadas depois do clique, associando o arquivo ao
        clique que o originou. Caminhos monitorados:

        1. Evento de download na aba de origem ou em qualquer aba nova
           (CSV da mensalidade e .zip da coparticipação abrem uma aba auxiliar
           que fecha sozinha).
        2. Nova aba que permanece aberta com a URL direta do arquivo
           (PDF e TXT da mensalidade).
        3. A própria aba navegar para a URL do arquivo.
        4. Blob PDF criado na página de origem (último recurso, só para PDF).

        Quando "acionar" é informado (função sem argumentos), ele é executado no
        lugar do clique no locator; é usado pelo método de teclado, em que o
        botão já está em foco e o acionamento é a tecla ENTER.

        Retorna {"caminho", "metodo", "url"} ou None.
        """
        page = page or self.page
        fmt = str(formato or "").strip().upper()
        extensao_fallback = f".{fmt.lower()}"
        timeout_ms = max(1_000, int(self.timeout_download_analitico))
        intervalo_ms = int(self.intervalo_monitoramento_analitico)

        downloads_detectados = []
        paginas_ouvidas = []
        primeira_vez_pagina = {}
        urls_processadas = set()

        def ao_download(download):
            downloads_detectados.append(download)

        def ouvir_pagina(pagina):
            try:
                pagina.on("download", ao_download)
                paginas_ouvidas.append(pagina)
            except Exception:
                pass

        def ao_nova_pagina(pagina):
            # O download pode ser emitido pela aba auxiliar, não pela aba da fatura.
            ouvir_pagina(pagina)

        # URLs de arquivo solicitadas por qualquer aba depois do clique. Servem de
        # contingência quando o evento de download não chega ou não pode ser salvo.
        urls_arquivo_rede = []
        self.ultimo_erro_captura_analitico = ""

        def ao_request(request):
            try:
                url_request = str(request.url or "")
                extensao_request = Path(urlparse(url_request).path).suffix.lower()
                if extensao_request in {".csv", ".txt", ".pdf", ".zip"} and url_request not in urls_arquivo_rede:
                    urls_arquivo_rede.append(url_request)
            except Exception:
                pass

        try:
            paginas_antes = list(self.context.pages) if self.context else []
        except Exception:
            paginas_antes = []
        try:
            url_antes = str(page.url or "")
        except Exception:
            url_antes = ""

        def salvar_download_detectado():
            while downloads_detectados:
                download = downloads_detectados.pop(0)
                try:
                    url_download = str(download.url or "")
                except Exception:
                    url_download = ""
                try:
                    caminho_download = self.salvar_download(
                        download,
                        novo_nome_base,
                        extensao_fallback=extensao_fallback,
                    )
                except Exception as e:
                    self.ultimo_erro_captura_analitico = (
                        f"O download foi iniciado, mas não pôde ser salvo: {self.resumir_erro(e)}"
                    )
                    print(f"[AVISO] {descricao}: {self.ultimo_erro_captura_analitico}")
                    self.registrar_erro(f"Falha ao salvar o download de {descricao}", self.resumir_erro(e))
                    # Tenta novamente pela URL do próprio download.
                    if url_download and url_download not in urls_arquivo_rede:
                        urls_arquivo_rede.insert(0, url_download)
                    continue
                if caminho_download:
                    # O nome sugerido pelo portal é mantido exatamente como veio
                    # (inclusive a extensão em maiúsculas, ex.: .ZIP, .CSV).
                    return {
                        "caminho": Path(caminho_download),
                        "metodo": "download",
                        "url": url_download,
                    }
            return None

        ouvir_pagina(page)
        listener_contexto = False
        try:
            self.context.on("page", ao_nova_pagina)
            listener_contexto = True
        except Exception as e:
            print(f"[AVISO] Não foi possível instalar o rastreador de novas abas: {self.resumir_erro(e)}")
        listener_rede = False
        try:
            self.context.on("request", ao_request)
            listener_rede = True
        except Exception as e:
            print(f"[AVISO] Não foi possível instalar o rastreador de rede: {self.resumir_erro(e)}")

        try:
            if fmt == "PDF":
                self.preparar_captura_blob_pdf(page)

            if acionar is not None:
                print(f"Acionando uma única vez: {descricao}")
                acionar()
            else:
                print(f"Clicando uma única vez em: {descricao}")
                self.clicar_e_aguardar(
                    locator,
                    descricao,
                    page=page,
                    usar_networkidle=False,
                )

            inicio = time.monotonic()
            limite = inicio + (timeout_ms / 1000.0)

            while time.monotonic() < limite:
                agora = time.monotonic()

                # 1. Evento de download (aba de origem ou aba auxiliar).
                resultado = salvar_download_detectado()
                if resultado:
                    print(
                        f"{descricao} concluído por evento de download após "
                        f"{agora - inicio:.1f} segundo(s)."
                    )
                    return resultado

                # 2. Nova aba que ficou aberta exibindo o arquivo.
                try:
                    paginas_atuais = list(self.context.pages) if self.context else []
                except Exception:
                    paginas_atuais = []
                for pagina_nova in paginas_atuais:
                    if pagina_nova in paginas_antes:
                        continue
                    try:
                        if pagina_nova.is_closed():
                            continue
                        url_nova = str(pagina_nova.url or "")
                    except Exception:
                        continue
                    # Somente URLs que podem ser buscadas (ignora about:, chrome-error: etc.).
                    if not url_nova.lower().startswith(("http://", "https://", "blob:")):
                        continue

                    # Dá preferência ao evento de download: se a aba for apenas
                    # o gatilho de um download, o arquivo chega pelo passo 1.
                    chave_pagina = id(pagina_nova)
                    if chave_pagina not in primeira_vez_pagina:
                        primeira_vez_pagina[chave_pagina] = agora
                    if (agora - primeira_vez_pagina[chave_pagina]) < self.carencia_nova_aba_analitico:
                        continue

                    chave_url = f"popup::{url_nova}"
                    if chave_url in urls_processadas:
                        continue
                    urls_processadas.add(chave_url)

                    print(
                        f"Nova aba detectada para {descricao}: "
                        f"{self.limpar_url_para_log(url_nova)}"
                    )
                    if url_nova.startswith("blob:"):
                        caminho = self.baixar_url_da_pagina_por_request(
                            pagina_nova,
                            novo_nome_base,
                            extensao_fallback,
                        )
                    else:
                        caminho = self.baixar_url_analitico_por_request(
                            pagina_nova,
                            url_nova,
                            novo_nome_base,
                            extensao_fallback,
                        )
                    if caminho:
                        print(
                            f"{descricao} concluído pela URL da nova aba após "
                            f"{agora - inicio:.1f} segundo(s)."
                        )
                        return {
                            "caminho": caminho,
                            "metodo": "nova_aba_url",
                            "url": url_nova,
                        }

                # 2b. Arquivo solicitado na rede, sem download salvo nem aba aberta.
                if urls_arquivo_rede and (agora - inicio) >= (self.carencia_nova_aba_analitico * 2):
                    for url_rede in list(urls_arquivo_rede):
                        chave_url = f"rede::{url_rede}"
                        if chave_url in urls_processadas or f"popup::{url_rede}" in urls_processadas:
                            continue
                        urls_processadas.add(chave_url)
                        print(
                            f"Arquivo de {descricao} identificado na rede: "
                            f"{self.limpar_url_para_log(url_rede)}"
                        )
                        caminho = self.baixar_url_analitico_por_request(
                            None,
                            url_rede,
                            novo_nome_base,
                            extensao_fallback,
                        )
                        if caminho:
                            return {
                                "caminho": caminho,
                                "metodo": "url_rede",
                                "url": url_rede,
                            }

                # 3. A própria aba navegou para o arquivo.
                try:
                    url_atual = str(page.url or "")
                except Exception:
                    url_atual = ""
                if url_atual and url_antes and url_atual != url_antes:
                    chave_url = f"origem::{url_atual}"
                    if chave_url not in urls_processadas:
                        urls_processadas.add(chave_url)
                        print(
                            f"A própria aba mudou de URL para {descricao}: "
                            f"{self.limpar_url_para_log(url_atual)}"
                        )
                        caminho = self.baixar_url_analitico_por_request(
                            page,
                            url_atual,
                            novo_nome_base,
                            extensao_fallback,
                        )
                        if caminho:
                            return {
                                "caminho": caminho,
                                "metodo": "url_mesma_aba",
                                "url": url_atual,
                            }

                # wait_for_timeout mantém o Playwright processando os eventos.
                try:
                    page.wait_for_timeout(intervalo_ms)
                except Exception:
                    time.sleep(intervalo_ms / 1000.0)

            # Verificação final imediatamente antes de declarar que nada chegou.
            resultado = salvar_download_detectado()
            if resultado:
                return resultado

            if fmt == "PDF":
                caminho = self.salvar_blob_pdf_capturado(page, novo_nome_base, extensao_fallback)
                if caminho:
                    return {
                        "caminho": caminho,
                        "metodo": "blob",
                        "url": "blob:",
                    }

            print(
                f"[AVISO] {descricao} foi acionado, mas o portal não entregou arquivo "
                f"em {timeout_ms / 1000:.0f} segundo(s)."
            )
            return None

        finally:
            if listener_contexto:
                try:
                    self.context.remove_listener("page", ao_nova_pagina)
                except Exception:
                    pass
            if listener_rede:
                try:
                    self.context.remove_listener("request", ao_request)
                except Exception:
                    pass
            for pagina_ouvida in paginas_ouvidas:
                try:
                    pagina_ouvida.remove_listener("download", ao_download)
                except Exception:
                    pass

            # Fecha somente as abas criadas por este clique e devolve o foco à fatura.
            self.fechar_paginas_abertas_apos_clique(paginas_antes, page)

            # Se a aba da fatura saiu da tela de detalhes, volta para ela.
            try:
                if url_antes and not page.is_closed() and str(page.url or "") != url_antes:
                    page.go_back(wait_until="domcontentloaded", timeout=self.TIMEOUT_LONGO)
                    self.respirar_sistema(page=page, timeout=15_000, segundos_fallback=2)
            except Exception:
                pass

    def validar_arquivo_analitico(self, caminho, formato):
        """
        Confere se o arquivo salvo é realmente o relatório.

        Retorna (valido, codigo_validacao, membros_zip). Os códigos de sucesso
        são os mesmos usados pelo gravador:
        ZIP_CRC_OK, ASSINATURA_PDF_OK e TEXTO_PLAUSIVEL.
        """
        fmt = str(formato or "").strip().upper()
        try:
            caminho = Path(caminho)
            if not caminho.exists():
                return False, "ARQUIVO_NAO_ENCONTRADO", []
            if caminho.stat().st_size <= 0:
                return False, "ARQUIVO_VAZIO", []

            with open(caminho, "rb") as arquivo:
                cabecalho = arquivo.read(4096)

            extensao = caminho.suffix.lower()

            if extensao == ".zip" or cabecalho.startswith(b"PK\x03\x04"):
                try:
                    with zipfile.ZipFile(caminho) as pacote:
                        membros = [
                            {"name": item.filename, "bytes": item.file_size}
                            for item in pacote.infolist()
                            if not item.is_dir()
                        ]
                        if not membros:
                            return False, "ZIP_SEM_ARQUIVOS", []
                        if pacote.testzip() is not None:
                            return False, "ZIP_CRC_INVALIDO", membros
                    return True, "ZIP_CRC_OK", membros
                except zipfile.BadZipFile:
                    return False, "ZIP_INVALIDO", []

            if fmt == "PDF" or extensao == ".pdf":
                if cabecalho.lstrip()[:5] == b"%PDF-":
                    return True, "ASSINATURA_PDF_OK", []
                return False, "PDF_SEM_ASSINATURA", []

            # CSV e TXT: rejeita páginas/respostas de erro salvas no lugar do relatório.
            inicio = cabecalho.lstrip(b"\xef\xbb\xbf \t\r\n").lower()
            if inicio.startswith((b"<?xml", b"<!doctype", b"<html", b"<error")):
                return False, "RESPOSTA_DE_ERRO_NO_LUGAR_DO_RELATORIO", []
            if b"\x00" in cabecalho:
                return False, "CONTEUDO_BINARIO_INESPERADO", []
            return True, "TEXTO_PLAUSIVEL", []

        except Exception as e:
            return False, f"FALHA_NA_VALIDACAO: {self.resumir_erro(e)}", []

    def extrair_zip_analitico(self, caminho_zip, novo_nome_base):
        """
        Extrai os arquivos do .zip para a mesma pasta, mantendo o nome original
        de cada arquivo dentro do .zip. Retorna a lista de caminhos extraídos.
        A remoção do .zip é feita por baixar_relatorio_analitico após o registro.
        """
        extraidos = []
        caminho_zip = Path(caminho_zip)
        pasta_destino = caminho_zip.parent

        with zipfile.ZipFile(caminho_zip) as pacote:
            membros = [item for item in pacote.infolist() if not item.is_dir()]
            for numero, item in enumerate(membros, 1):
                # Usa apenas o nome final do membro: nunca grava fora da pasta de destino.
                nome_membro = Path(str(item.filename).replace("\\", "/")).name
                sufixo = "" if len(membros) == 1 else f"_{numero}"
                nome_padrao = f"{novo_nome_base}{sufixo}"
                extensao = Path(nome_membro).suffix or ".bin"

                destino = self.montar_caminho_destino(pasta_destino, nome_membro, nome_padrao, extensao)

                with pacote.open(item) as origem, open(destino, "wb") as saida:
                    shutil.copyfileobj(origem, saida)
                print(f"Arquivo extraído do ZIP: {destino.name}")
                extraidos.append(destino)

        return extraidos

    def baixar_relatorio_analitico(self, contrato, vencimento_alvo, tipo_fatura, formato):
        fmt = str(formato or "").strip().upper()
        documento = f"ANALITICO_{fmt}"

        if self.verificar_documento_baixado(contrato, vencimento_alvo, tipo_fatura, documento):
            print(f"Relatório analítico {fmt} já consta no controle. Pulando...")
            return True

        print(f"Iniciando download do relatório analítico {fmt}...")
        sufixo_tipo = self.obter_sufixo_tipo_fatura(tipo_fatura)
        vencimento_fmt = self.formatar_vencimento_nome_arquivo(vencimento_alvo)
        novo_nome_base = f"{contrato}_ANALITICO_{fmt}_VENC_{vencimento_fmt}_{sufixo_tipo}"
        descricao = f"Relatório analítico {fmt}"

        # Garante que a aba está na tela de detalhes antes de procurar o botão.
        if not self.tela_detalhes_fatura_aberta():
            print("[AVISO] A aba não está na tela Detalhes da fatura. Reabrindo a fatura...")
            try:
                self.reabrir_fatura_atual_para_analitico()
            except Exception as e:
                mensagem = f"Não foi possível reabrir a fatura antes do {descricao}: {self.resumir_erro(e)}"
                print(f"[AVISO] {mensagem}")
                self.registrar_diagnostico_se_ausente(
                    documento,
                    "REABRIR_FATURA_RETENTATIVA_ANALITICO",
                    mensagem,
                    page=self.page,
                    excecao=e,
                )
                self.marcar_documento_fatura(documento, "ERRO", mensagem)
                return False

        inicio_captura = time.perf_counter()
        resultado = None

        # --- INÍCIO INSERÇÃO COPARTICIPAÇÃO POR TECLADO ---
        # Método principal da coparticipação:
        #   1 clique em "Relatórios" -> TAB (CSV 1x, TXT 2x, PDF 3x) -> ENTER.
        if tipo_fatura == self.TIPO_CARNET and self.coparticipacao_por_teclado:
            try:
                foco_ok, detalhe_foco = self.focar_relatorio_coparticipacao_por_teclado(fmt, page=self.page)
            except Exception as e:
                foco_ok, detalhe_foco = False, self.resumir_erro(e)

            if foco_ok:
                resultado = self.capturar_arquivo_analitico_por_clique(
                    None,
                    f"{descricao} (Relatórios + TAB {self.TABS_RELATORIO_COPARTICIPACAO[fmt]}x + ENTER)",
                    novo_nome_base,
                    fmt,
                    page=self.page,
                    acionar=lambda: self.page.keyboard.press("Enter"),
                )
                if resultado and resultado.get("caminho"):
                    resultado["metodo"] = f"teclado_{resultado.get('metodo', '')}"
                else:
                    resultado = None
                    print(
                        "[AVISO] O método por teclado não entregou arquivo. "
                        "Executando contingência pelo botão da tela..."
                    )
            else:
                print(
                    f"[AVISO] Método por teclado não aplicado: {detalhe_foco} "
                    "Executando contingência pelo botão da tela..."
                )
        # --- FIM INSERÇÃO COPARTICIPAÇÃO POR TECLADO ---

        if resultado is None:
            botao = self.localizar_botao_analitico_rapido(
                fmt,
                timeout_total_ms=self.timeout_botao_analitico_total,
                page=self.page,
            )
            if botao is None:
                print(f"[AVISO] Botão do relatório analítico {fmt} não encontrado para esta fatura.")
                situacao_tela = self.registrar_tela_sem_botao_analitico(documento, fmt, page=self.page)
                self.registrar_diagnostico_se_ausente(
                    documento,
                    "LOCALIZAR_BOTAO_ANALITICO",
                    (
                        f"{situacao_tela} Botão do relatório analítico {fmt} não ficou visível e "
                        f"habilitado em {self.timeout_botao_analitico_total / 1000:.1f}s."
                    ),
                    page=self.page,
                )
                self.marcar_documento_fatura(
                    documento,
                    "NAO_DISPONIVEL",
                    f"Botão do relatório analítico {fmt} não encontrado. {situacao_tela}",
                )
                return False


            resultado = self.capturar_arquivo_analitico_por_clique(
                botao,
                descricao,
                novo_nome_base,
                fmt,
                page=self.page,
            )
        duracao_captura = time.perf_counter() - inicio_captura

        if not resultado or not resultado.get("caminho"):
            mensagem = (
                f"{descricao} foi acionado, mas o portal não disponibilizou download, "
                "nova aba com o arquivo nem mudança de URL dentro de "
                f"{self.timeout_download_analitico / 1000:.0f} segundos."
            )
            if self.ultimo_erro_captura_analitico:
                mensagem += f" Detalhe: {self.ultimo_erro_captura_analitico}"
            self.registrar_rastreamento_download(
                documento,
                "FALHA",
                validacao="ERRO_AO_SALVAR" if self.ultimo_erro_captura_analitico else "SEM_ARQUIVO",
                duracao=duracao_captura,
            )
            self.registrar_diagnostico_se_ausente(
                documento,
                "CAPTURAR_ARQUIVO_ANALITICO",
                mensagem,
                page=self.page,
            )
            self.marcar_documento_fatura(documento, "ERRO", mensagem)
            return False

        caminho = Path(resultado["caminho"])
        metodo = resultado.get("metodo", "")
        url_origem = resultado.get("url", "")

        valido, validacao, membros_zip = self.validar_arquivo_analitico(caminho, fmt)
        # Garante que o arquivo entregue é do formato pedido (protege contra o
        # acionamento do botão de outro formato).
        if valido and membros_zip:
            extensoes = {Path(str(m.get("name", ""))).suffix.lower().lstrip(".") for m in membros_zip}
            extensoes_conhecidas = extensoes & {f.lower() for f in self.FORMATOS_ANALITICOS_SUPORTADOS}
            if extensoes_conhecidas and fmt.lower() not in extensoes_conhecidas:
                valido = False
                validacao = "FORMATO_DIFERENTE_DO_SOLICITADO"
        if not valido:
            # Mantém o arquivo para conferência, mas fora do padrão de nomes válidos.
            caminho_invalido = caminho
            try:
                caminho_invalido = caminho.with_name(caminho.name + ".invalido")
                contador = 1
                while caminho_invalido.exists():
                    caminho_invalido = caminho.with_name(f"{caminho.name}_{contador}.invalido")
                    contador += 1
                caminho.rename(caminho_invalido)
            except Exception:
                caminho_invalido = caminho

            mensagem = (
                f"{descricao} foi salvo, mas o conteúdo foi reprovado ({validacao}). "
                f"Arquivo mantido para conferência: {caminho_invalido.name}."
            )
            print(f"[AVISO] {mensagem}")
            self.registrar_rastreamento_download(
                documento,
                "INVALIDO",
                metodo=metodo,
                caminho=caminho_invalido,
                validacao=validacao,
                duracao=duracao_captura,
                url=url_origem,
                membros_zip=membros_zip,
            )
            self.registrar_diagnostico_se_ausente(
                documento,
                "VALIDAR_ARQUIVO_ANALITICO",
                mensagem,
                page=self.page,
            )
            self.marcar_documento_fatura(documento, "ERRO", mensagem)
            return False

        extraidos = []
        if self.extrair_zip_ativo and validacao == "ZIP_CRC_OK":
            try:
                extraidos = self.extrair_zip_analitico(caminho, novo_nome_base)
                self.resumo_execucao["arquivos_extraidos_zip"] += len(extraidos)
            except Exception as e:
                # O .zip válido já está salvo; a falha na extração não invalida o download.
                print(f"[AVISO] ZIP salvo, mas a extração falhou: {self.resumir_erro(e)}")
                self.registrar_erro(f"Falha ao extrair ZIP do {descricao}", self.resumir_erro(e))

        self.registrar_rastreamento_download(
            documento,
            "SALVO",
            metodo=metodo,
            caminho=caminho,
            validacao=validacao,
            duracao=duracao_captura,
            url=url_origem,
            membros_zip=membros_zip,
            extraidos=extraidos,
        )

        # --- INÍCIO INSERÇÃO MANTER SOMENTE ARQUIVOS DESCOMPACTADOS ---
        # O rastreamento acima já gravou bytes/SHA-256 do .zip e os nomes extraídos.
        # Com a extração concluída, o .zip é removido e ficam só os arquivos.
        arquivos_finais = [caminho]
        if extraidos:
            arquivos_finais = list(extraidos)
            try:
                caminho.unlink()
                print(f"ZIP removido após extração: {caminho.name}")
            except Exception as e:
                arquivos_finais = [caminho] + list(extraidos)
                print(f"[AVISO] Arquivos extraídos, mas o ZIP não pôde ser removido: {self.resumir_erro(e)}")
                self.registrar_erro(f"Falha ao remover ZIP do {descricao}", self.resumir_erro(e))
        # --- FIM INSERÇÃO MANTER SOMENTE ARQUIVOS DESCOMPACTADOS ---

        caminhos_evidencia = " ; ".join(str(item) for item in arquivos_finais)
        self.marcar_documento_fatura(documento, "OK", caminho=caminhos_evidencia)
        self.registrar_documento_log(contrato, vencimento_alvo, tipo_fatura, documento)
        nomes_finais = ", ".join(Path(item).name for item in arquivos_finais)
        print(f"{descricao} concluído: {nomes_finais} | {validacao} | método: {metodo}")
        return True
    # --- FIM INSERÇÃO RELATÓRIOS ANALÍTICOS ---

    # ----------------------------------------------------------------------
    # Processamento das faturas
    # ----------------------------------------------------------------------
    def processar_faturas_do_vencimento(self, vencimento_alvo):
        print(f"\nProcurando faturas para o vencimento: {vencimento_alvo}...")
        ocorrencias = []
        self.linhas_extrato_vistas = []
        try:
            self.url_extrato_atual = str(self.page.url or "")
        except Exception:
            self.url_extrato_atual = None

        print("-> Verificando aba 'Em Aberto'...")
        self.navegar_para_aba_e_pagina("ABERTO", 1)
        ocorrencias.extend(self.pesquisar_ocorrencias_na_tela(vencimento_alvo, "ABERTO", 1))

        if not ocorrencias:
            print("-> Fatura não encontrada em 'Em Aberto'. Verificando aba 'Histórico'...")
            # --- INÍCIO INSERÇÃO OTIMIZAÇÃO DE TEMPO / PARADA ANTECIPADA DO HISTÓRICO ---
            for pagina in range(1, self.max_paginas_historico + 1):
                pagina_disponivel = self.navegar_para_aba_e_pagina("HISTORICO", pagina)
                if not pagina_disponivel:
                    break
                encontradas = self.pesquisar_ocorrencias_na_tela(vencimento_alvo, "HISTORICO", pagina)
                if encontradas:
                    ocorrencias.extend(encontradas)
                    break
            # --- FIM INSERÇÃO OTIMIZAÇÃO DE TEMPO / PARADA ANTECIPADA DO HISTÓRICO ---

        if not ocorrencias:
            print("Nenhuma fatura encontrada para esta data em nenhuma aba.")
            # --- INÍCIO INSERÇÃO DIAGNÓSTICO DE FATURA NÃO ENCONTRADA ---
            resumo_extrato = self.registrar_falta_de_fatura_no_extrato(vencimento_alvo)
            # --- FIM INSERÇÃO DIAGNÓSTICO DE FATURA NÃO ENCONTRADA ---
            # --- INÍCIO INSERÇÃO CONTROLE TXT / EVIDÊNCIA SEM FATURA ---
            self.adicionar_evidencia_erro_contrato(
                f"Nenhuma fatura encontrada para o vencimento {vencimento_alvo}. {resumo_extrato}"
            )
            # --- FIM INSERÇÃO CONTROLE TXT / EVIDÊNCIA SEM FATURA ---
            return False

        ocorrencias.sort(
            key=lambda oc: self.PRIORIDADE_TIPOS.index(oc["tipo_fatura"])
            if oc["tipo_fatura"] in self.PRIORIDADE_TIPOS
            else 99
        )

        print(f"Faturas mapeadas para processamento: {[oc['tipo_fatura'] for oc in ocorrencias]}")
        houve_download = False
        total = len(ocorrencias)

        for i, ocorrencia in enumerate(ocorrencias, 1):
            tipo_fatura = ocorrencia["tipo_fatura"]
            self.vencimento_atual = vencimento_alvo
            self.tipo_fatura_atual = tipo_fatura
            self.ocorrencia_fatura_atual = dict(ocorrencia)
            # --- INÍCIO INSERÇÃO CONTROLE TXT / PASTA POR FATURA ---
            self.iniciar_estado_fatura(self.contrato_atual, vencimento_alvo, tipo_fatura)
            # --- FIM INSERÇÃO CONTROLE TXT / PASTA POR FATURA ---

            print(
                f"\n-> Iniciando extração do tipo: {tipo_fatura} "
                f"(Aba: {ocorrencia['tab']}, Pág: {ocorrencia['pagina']})"
            )

            try:
                self.abrir_fatura_por_ocorrencia(ocorrencia, vencimento_alvo)
                self.ocorrencia_fatura_atual = dict(ocorrencia)
                self.conferir_tipo_da_fatura_aberta(tipo_fatura)

                # --- INÍCIO INSERÇÃO FLUXO DOS RELATÓRIOS ANALÍTICOS ---
                # A fatura é aberta e o processamento segue para os relatórios analíticos
                # (e, se BAIXAR_BOLETO_NF=1, também para boleto e nota fiscal).
                sucesso_download = False

                try:
                    sucesso_docs = self.baixar_documentos_fatura(
                        self.contrato_atual,
                        vencimento_alvo,
                        tipo_fatura,
                    )
                    if sucesso_docs:
                        sucesso_download = True
                except Exception as e_docs:
                    print(f"[AVISO] Falha na rotina de documentos da fatura: {self.resumir_erro(e_docs)}")
                    self.registrar_erro("Falha na rotina de documentos da fatura", self.resumir_erro(e_docs))
                # --- FIM INSERÇÃO FLUXO DOS RELATÓRIOS ANALÍTICOS ---

                if sucesso_download:
                    houve_download = True

                if i < total:
                    self.voltar_para_extrato()

            except Exception as e:
                print(f"[ERRO] Falha ao acessar/processar fatura {tipo_fatura}: {self.resumir_erro(e)}")
                self.registrar_erro(
                    f"Falha na fatura {tipo_fatura} (Aba: {ocorrencia['tab']}, Pág: {ocorrencia['pagina']})",
                    self.resumir_erro(e),
                )
                # --- INÍCIO INSERÇÃO CONTROLE TXT / STATUS DA FATURA ---
                self.mensagens_fatura_atual.append(self.limpar_campo_txt(f"Falha na fatura {tipo_fatura}: {self.resumir_erro(e)}"))
                for doc in self.DOCUMENTOS_PERMITIDOS:
                    if self.status_docs_fatura_atual.get(doc, "PENDENTE") == "PENDENTE":
                        self.status_docs_fatura_atual[doc] = "ERRO"
                # --- FIM INSERÇÃO CONTROLE TXT / STATUS DA FATURA ---
                self.salvar_screenshot(f"erro_fatura_{tipo_fatura}")

                try:
                    self.voltar_para_extrato()
                except Exception:
                    pass

            # --- INÍCIO INSERÇÃO CONTROLE TXT / REGISTRAR FATURA ---
            self.registrar_linha_controle_fatura()
            # --- FIM INSERÇÃO CONTROLE TXT / REGISTRAR FATURA ---

        return houve_download

    # ----------------------------------------------------------------------
    # Execução principal
    # ----------------------------------------------------------------------
    def validar_colunas_excel(self, df):
        colunas_obrigatorias = ["MÓDULO", "CONTRATO", "VENCIMENTO"]
        faltantes = [col for col in colunas_obrigatorias if col not in df.columns]
        if faltantes:
            raise ValueError(f"Colunas obrigatórias ausentes no Excel: {faltantes}")

    # --- INÍCIO INSERÇÃO MODO RÁPIDO COM SESSÃO REUTILIZÁVEL ---
    def iniciar_sessao_rapida(self):
        """Abre uma sessão nova e realiza o login sem esperas fixas."""
        print("Iniciando/reiniciando sessão rápida do portal...")
        self.fechar_navegador()
        self.iniciar_navegador()
        self.fazer_login()

        self.aguardar_selector(
            "xpath=/html/body/div/div/div/main/div/div/main/div[1]/div/div[1]/div/div/div/input | //input[contains(@placeholder, 'Contrato') or contains(@placeholder, 'contrato') or contains(@placeholder, 'Buscar') or contains(@placeholder, 'buscar')] | //main//input",
            descricao="campo de pesquisa de contratos após o login",
            timeout=self.TIMEOUT_LONGO,
        )

    def preparar_pesquisa_proximo_contrato(self):
        """Retorna à pesquisa usando a sessão atual; reinicia apenas se necessário."""
        print("\nPreparando pesquisa do próximo contrato na mesma sessão...")
        try:
            if not self.page or not self.url_painel:
                raise RuntimeError("Sessão ou URL do painel indisponível.")

            self.page.goto(
                self.url_painel,
                wait_until="domcontentloaded",
                timeout=self.TIMEOUT_LONGO,
            )
            self.aguardar_selector(
                "xpath=/html/body/div/div/div/main/div/div/main/div[1]/div/div[1]/div/div/div/input | //input[contains(@placeholder, 'Contrato') or contains(@placeholder, 'contrato') or contains(@placeholder, 'Buscar') or contains(@placeholder, 'buscar')] | //main//input",
                descricao="campo de pesquisa do próximo contrato",
                timeout=self.TIMEOUT_LONGO,
            )
            print("Sessão reaproveitada com sucesso.")
        except Exception as e:
            print(f"[AVISO] Não foi possível reaproveitar a sessão: {self.resumir_erro(e)}")
            print("Reiniciando navegador e refazendo o login automaticamente...")
            self.iniciar_sessao_rapida()

    def imprimir_resumo_desempenho(self):
        if not self.tempos_contratos:
            return

        total_segundos = sum(item[1] for item in self.tempos_contratos)
        media_segundos = total_segundos / len(self.tempos_contratos)
        mais_rapido = min(self.tempos_contratos, key=lambda item: item[1])
        mais_lento = max(self.tempos_contratos, key=lambda item: item[1])

        print("\n" + "=" * 90)
        print("RESUMO DE DESEMPENHO - MODO RÁPIDO")
        print("=" * 90)
        print(f"Contratos cronometrados : {len(self.tempos_contratos)}")
        print(f"Tempo total             : {total_segundos:.2f} segundos")
        print(f"Tempo médio por contrato: {media_segundos:.2f} segundos")
        print(f"Contrato mais rápido    : {mais_rapido[0]} | {mais_rapido[1]:.2f} segundos")
        print(f"Contrato mais lento     : {mais_lento[0]} | {mais_lento[1]:.2f} segundos")
        print("=" * 90)

    # --- INÍCIO INSERÇÃO EXECUÇÃO PARALELA (PLANILHA E FILA) ---
    def preparar_planilha_contratos(self):
        """Lê e normaliza a planilha (mesmas regras de antes, agora reutilizáveis)."""
        print(f"Lendo base de dados Excel: {self.file_path}")
        df_contratos = pd.read_excel(self.file_path)
        self.validar_colunas_excel(df_contratos)

        df_contratos.dropna(subset=["CONTRATO", "VENCIMENTO"], inplace=True)
        df_contratos["VENCIMENTO"] = pd.to_datetime(
            df_contratos["VENCIMENTO"],
            dayfirst=True,
            errors="coerce",
        ).dt.strftime("%d/%m/%Y")
        df_contratos.dropna(subset=["VENCIMENTO"], inplace=True)

        if "STATUS" not in df_contratos.columns:
            df_contratos["STATUS"] = ""
        if "DETALHES_DOWNLOAD" not in df_contratos.columns:
            df_contratos["DETALHES_DOWNLOAD"] = ""
        if "ULTIMA_EXECUCAO" not in df_contratos.columns:
            df_contratos["ULTIMA_EXECUCAO"] = ""
        return df_contratos

    def iterar_linhas_contratos(self, df_contratos):
        """
        Sequencial: percorre a planilha na ordem.
        Paralelo: retira a próxima linha da fila comum a todas as janelas.
        """
        if self.fila_linhas is None:
            yield from df_contratos.iterrows()
            return
        while True:
            if self.evento_parada is not None and self.evento_parada.is_set():
                return
            try:
                index = self.fila_linhas.get_nowait()
            except queue.Empty:
                return
            self.indices_processados.append(index)
            yield index, df_contratos.loc[index]

    def contar_contrato_finalizado(self):
        """Quantidade de contratos finalizados, somando todas as janelas."""
        self.contratos_finalizados += 1
        progresso = self.progresso_paralelo
        if not progresso:
            return self.contratos_finalizados
        with progresso["trava"]:
            progresso["concluidos"] += 1
            return progresso["concluidos"]
    # --- FIM INSERÇÃO EXECUÇÃO PARALELA (PLANILHA E FILA) ---

    def executar_modo_rapido(self):
        df_contratos = None
        index_atual = None
        contratos_iniciados = 0
        self.tempos_contratos = []
        self.tempos_etapas = []
        if not self.modo_worker:
            self.inicio_execucao_geral = time.perf_counter()

        try:
            if self.df_contratos_preparado is not None:
                # Janela paralela: a planilha já foi lida pelo coordenador.
                df_contratos = self.df_contratos_preparado
            else:
                with self.medir_etapa("LEITURA_E_PREPARACAO_EXCEL"):
                    df_contratos = self.preparar_planilha_contratos()
                self.total_contratos_planejados = len(df_contratos)

            print("Modo rápido ATIVADO: o navegador e o login serão reutilizados entre os contratos.")

            for index, row in self.iterar_linhas_contratos(df_contratos):
                inicio_contrato = time.perf_counter()
                index_atual = index
                modulo = row["MÓDULO"]
                contrato = row["CONTRATO"]
                vencimento = row["VENCIMENTO"]

                self.contrato_atual = contrato
                self.vencimento_atual = vencimento
                self.tipo_fatura_atual = None
                self.docs_linha_atual = []
                self.ocorrencia_fatura_atual = None
                self.status_docs_fatura_atual = {}
                self.caminhos_docs_fatura_atual = {}
                self.mensagens_fatura_atual = []
                self.controle_fatura_registrado = False
                self.index_linha_atual = index
                self.evidencias_contrato_atual = []
                self.pasta_contrato_atual = self.obter_pasta_contrato(contrato, vencimento)

                # --- INÍCIO INSERÇÃO PULAR CONTRATO JÁ CONCLUÍDO ---
                if self.contrato_ja_concluido(contrato, vencimento):
                    print(
                        f"Linha {index + 1} | Contrato {contrato} | Venc {vencimento}: "
                        "já concluído em execução anterior. Pulando sem abrir o portal."
                    )
                    self.resumo_execucao["contratos_pulados_concluidos"] += 1
                    df_contratos.at[index, "STATUS"] = "Já concluído anteriormente"
                    df_contratos.at[index, "ULTIMA_EXECUCAO"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    self.contar_contrato_finalizado()
                    continue
                # --- FIM INSERÇÃO PULAR CONTRATO JÁ CONCLUÍDO ---

                self.resumo_execucao["contratos_processados"] += 1

                print("\n=======================================================")
                print(f"   Processando Linha {index + 1} | Contrato: {contrato} | Venc: {vencimento}")
                print("=======================================================")

                try:
                    # --- INÍCIO INSERÇÃO OTIMIZAÇÃO DE TEMPO / RECUPERAÇÃO DO CONTRATO ATUAL ---
                    ultimo_erro_tentativa = None
                    houve_download = False
                    for tentativa in range(1, self.tentativas_por_contrato + 1):
                        try:
                            if contratos_iniciados == 0 or not self.page or tentativa > 1:
                                with self.medir_etapa("ABRIR_NAVEGADOR_E_LOGIN", f"Tentativa {tentativa}"):
                                    self.iniciar_sessao_rapida()
                            else:
                                with self.medir_etapa("REAPROVEITAR_SESSAO", f"Tentativa {tentativa}"):
                                    self.preparar_pesquisa_proximo_contrato()
                            contratos_iniciados += 1

                            with self.medir_etapa("SELECIONAR_EMPRESA_MATRIZ"):
                                self.selecionar_empresa_matriz(modulo)
                            with self.medir_etapa("BUSCAR_CONTRATO"):
                                self.buscar_contrato(contrato)
                            with self.medir_etapa("ENTRAR_NO_CONTRATO"):
                                self.entrar_no_contrato()
                            with self.medir_etapa("ACESSAR_EXTRATO_FINANCEIRO"):
                                self.acessar_extrato_financeiro()
                            with self.medir_etapa("LOCALIZAR_E_PROCESSAR_FATURAS"):
                                houve_download = self.processar_faturas_do_vencimento(vencimento)
                            ultimo_erro_tentativa = None
                            break
                        except Exception as erro_tentativa:
                            ultimo_erro_tentativa = erro_tentativa
                            if tentativa >= self.tentativas_por_contrato:
                                raise
                            print(
                                f"[AVISO] Tentativa {tentativa} falhou para o contrato {contrato}: "
                                f"{self.resumir_erro(erro_tentativa)}"
                            )
                            print("Reiniciando a sessão e repetindo o contrato atual uma vez...")
                            self.fechar_navegador()

                    if ultimo_erro_tentativa is not None:
                        raise ultimo_erro_tentativa
                    # --- FIM INSERÇÃO OTIMIZAÇÃO DE TEMPO / RECUPERAÇÃO DO CONTRATO ATUAL ---

                    resultado_contrato = self.classificar_resultado_contrato_atual()
                    if resultado_contrato == "COMPLETO":
                        status = "Sucesso completo - Todos os documentos concluídos"
                        self.resumo_execucao["contratos_sucesso"] += 1
                        # --- INÍCIO INSERÇÃO PULAR CONTRATO JÁ CONCLUÍDO ---
                        self.registrar_contrato_concluido(contrato, vencimento)
                        # --- FIM INSERÇÃO PULAR CONTRATO JÁ CONCLUÍDO ---
                    elif resultado_contrato == "PARCIAL":
                        status = "Pendência - Download parcial"
                        self.resumo_execucao["contratos_pendencia"] += 1
                    else:
                        status = "Pendência - Nenhum documento concluído"
                        self.resumo_execucao["contratos_pendencia"] += 1

                    detalhes = self.montar_detalhes_resultado_contrato()
                    df_contratos.at[index, "STATUS"] = status
                    df_contratos.at[index, "DETALHES_DOWNLOAD"] = detalhes
                    df_contratos.at[index, "ULTIMA_EXECUCAO"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

                    print(f"Contrato {contrato} finalizado: {status} | {detalhes}")

                except Exception as e_linha:
                    erro = self.resumir_erro(e_linha)
                    print(f"[ERRO] Falha na linha {index + 1}: {erro}")
                    self.registrar_erro(f"Erro na linha {index + 1}", erro)
                    self.salvar_screenshot(f"erro_linha_{index + 1}")

                    self.resumo_execucao["contratos_erro"] += 1
                    self.adicionar_evidencia_erro_contrato(erro)

                    df_contratos.at[index, "STATUS"] = "Erro - Verificar Logs/Screenshots"
                    df_contratos.at[index, "DETALHES_DOWNLOAD"] = erro[:320]
                    df_contratos.at[index, "ULTIMA_EXECUCAO"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

                    print("A sessão será reiniciada automaticamente antes do próximo contrato.")
                    self.fechar_navegador()

                finally:
                    tempo_contrato = time.perf_counter() - inicio_contrato
                    self.tempos_contratos.append((str(contrato), tempo_contrato))
                    self.registrar_tempo_etapa("TOTAL_CONTRATO", tempo_contrato)
                    print(f"Tempo do contrato {contrato}: {self.formatar_duracao(tempo_contrato)} ({tempo_contrato:.2f} segundos)")
                    self.imprimir_projecao_execucao(self.contar_contrato_finalizado())

                    try:
                        status_evidencia = df_contratos.at[index, "STATUS"] if "STATUS" in df_contratos.columns else ""
                        detalhes_evidencia = df_contratos.at[index, "DETALHES_DOWNLOAD"] if "DETALHES_DOWNLOAD" in df_contratos.columns else ""
                        with self.medir_etapa("GERAR_EVIDENCIA_CONTRATO"):
                            self.gerar_evidencia_contrato(status_evidencia, detalhes_evidencia)
                    except Exception as e_evidencia:
                        print(f"[AVISO] Falha ao gerar evidência no final do contrato: {e_evidencia}")

                    if self.atualizar_excel_original:
                        try:
                            df_contratos.to_excel(self.file_path, index=False)
                            print("Arquivo Excel atualizado após o contrato.")
                        except Exception as e_excel_linha:
                            print(f"[ERRO] Não foi possível salvar o Excel após o contrato. Feche a planilha se estiver aberta. Erro: {e_excel_linha}")
                    else:
                        print("Arquivo Excel original NÃO foi alterado. Controle salvo em TXT.")

        except Exception as e:
            erro = self.resumir_erro(e)
            print(f"Ocorreu um erro crítico durante a execução: {erro}")
            self.registrar_erro("Erro Crítico na Execução", erro)
            if df_contratos is not None and index_atual is not None:
                try:
                    df_contratos.at[index_atual, "STATUS"] = "Erro Crítico - Verificar Log de Erros"
                    df_contratos.at[index_atual, "DETALHES_DOWNLOAD"] = erro[:320]
                    df_contratos.at[index_atual, "ULTIMA_EXECUCAO"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                except Exception:
                    pass

        finally:
            # --- INÍCIO INSERÇÃO EXECUÇÃO PARALELA (FIM DA JANELA) ---
            # Janela paralela: só fecha o navegador; os resumos são consolidados
            # pelo coordenador depois que todas as janelas terminam.
            if self.modo_worker:
                self.fechar_navegador()
                print("Janela finalizada: não há mais contratos na fila.")
                return
            # --- FIM INSERÇÃO EXECUÇÃO PARALELA (FIM DA JANELA) ---

            if self.atualizar_excel_original:
                print("\nSalvando status no arquivo Excel...")
                try:
                    if df_contratos is not None:
                        df_contratos.to_excel(self.file_path, index=False)
                        print("Arquivo Excel atualizado com sucesso!")
                except Exception as e:
                    print(f"[ERRO] Não foi possível salvar o Excel. Feche a planilha se estiver aberta. Erro: {e}")
            else:
                print("\nArquivo Excel original mantido sem alteração. Controle da execução foi gravado em TXT.")

            self.imprimir_resumo_final()
            self.imprimir_resumo_desempenho()
            if self.inicio_execucao_geral:
                tempo_total_execucao = time.perf_counter() - self.inicio_execucao_geral
                self.registrar_tempo_etapa("TOTAL_EXECUCAO", tempo_total_execucao)
                print(f"Tempo total real da execução: {self.formatar_duracao(tempo_total_execucao)}")
            self.gerar_resumo_tempos()
            self.gerar_resumo_diagnosticos()
            self.fechar_navegador()
    # --- FIM INSERÇÃO MODO RÁPIDO COM SESSÃO REUTILIZÁVEL ---

    # --- INÍCIO INSERÇÃO EXECUÇÃO PARALELA (COORDENADOR) ---
    def executar_paralelo(self):
        """
        Abre N janelas independentes (navegador e login próprios, uma thread
        cada) que retiram contratos de uma fila comum. Esta instância apenas
        coordena: lê a planilha, distribui as linhas e consolida os resumos.
        """
        self.tempos_contratos = []
        self.tempos_etapas = []
        self.inicio_execucao_geral = time.perf_counter()

        with self.medir_etapa("LEITURA_E_PREPARACAO_EXCEL"):
            df_contratos = self.preparar_planilha_contratos()
        self.total_contratos_planejados = len(df_contratos)

        quantidade = max(1, min(self.quantidade_workers, len(df_contratos)))
        print(
            f"Execução paralela: {quantidade} janela(s) para {len(df_contratos)} contrato(s). "
            f"Intervalo entre aberturas: {self.intervalo_inicio_workers:.1f}s."
        )

        # Credenciais resolvidas uma única vez aqui; as janelas não perguntam no terminal.
        email, senha = self.obter_credenciais()
        os.environ["HAPVIDA_EMAIL"] = email
        os.environ["HAPVIDA_SENHA"] = senha

        fila = queue.Queue()
        for index in df_contratos.index:
            fila.put(index)
        progresso = {"concluidos": 0, "trava": threading.Lock()}
        evento_parada = threading.Event()
        workers = []
        trava_workers = threading.Lock()

        def rodar_worker(numero):
            _contexto_thread.prefixo = f"[W{numero}]"
            try:
                worker = HapvidaNDIAnaliticosAutomation(self.file_path)
                worker.modo_worker = True
                worker.numero_worker = numero
                worker.fila_linhas = fila
                worker.df_contratos_preparado = df_contratos.copy()
                worker.progresso_paralelo = progresso
                worker.evento_parada = evento_parada
                worker.inicio_execucao_geral = self.inicio_execucao_geral
                worker.total_contratos_planejados = self.total_contratos_planejados
                # Cada janela tem só uma parte das linhas: o Excel é salvo pelo coordenador.
                worker.atualizar_excel_original = False
                with trava_workers:
                    workers.append(worker)
                worker.executar_modo_rapido()
            except Exception as e:
                print(f"[ERRO] Janela {numero} encerrada por falha: {self.resumir_erro(e)}")
                self.registrar_erro(f"Falha na janela paralela {numero}", self.resumir_erro(e))

        threads = []
        try:
            for numero in range(1, quantidade + 1):
                if fila.empty():
                    break
                thread = threading.Thread(
                    target=rodar_worker,
                    args=(numero,),
                    name=f"Worker-{numero}",
                    daemon=True,
                )
                thread.start()
                threads.append(thread)
                # Escalona as aberturas para não fazer todos os logins no mesmo instante.
                if numero < quantidade and self.intervalo_inicio_workers:
                    limite = time.monotonic() + self.intervalo_inicio_workers
                    while time.monotonic() < limite and not fila.empty():
                        time.sleep(0.2)

            while any(thread.is_alive() for thread in threads):
                for thread in threads:
                    thread.join(timeout=1.0)
        except KeyboardInterrupt:
            print("\n[AVISO] Interrupção solicitada: as janelas terminam o contrato atual e param.")
            evento_parada.set()
            for thread in threads:
                thread.join(timeout=120)

        # --- Consolidação dos resultados de todas as janelas ---
        for worker in workers:
            for chave, valor in worker.resumo_execucao.items():
                if isinstance(valor, (int, float)):
                    self.resumo_execucao[chave] = self.resumo_execucao.get(chave, 0) + valor
            self.tempos_etapas.extend(worker.tempos_etapas)
            self.tempos_contratos.extend(worker.tempos_contratos)
            self.diagnosticos_execucao.extend(worker.diagnosticos_execucao)
            for codigo, quantidade_codigo in worker.contagem_codigos_falha.items():
                self.contagem_codigos_falha[codigo] = self.contagem_codigos_falha.get(codigo, 0) + quantidade_codigo
            for index in worker.indices_processados:
                for coluna in ("STATUS", "DETALHES_DOWNLOAD", "ULTIMA_EXECUCAO"):
                    try:
                        df_contratos.at[index, coluna] = worker.df_contratos_preparado.at[index, coluna]
                    except Exception:
                        pass

        if self.atualizar_excel_original:
            print("\nSalvando status no arquivo Excel...")
            try:
                df_contratos.to_excel(self.file_path, index=False)
                print("Arquivo Excel atualizado com sucesso!")
            except Exception as e:
                print(f"[ERRO] Não foi possível salvar o Excel. Feche a planilha se estiver aberta. Erro: {e}")
        else:
            print("\nArquivo Excel original mantido sem alteração. Controle da execução foi gravado em TXT.")

        self.imprimir_resumo_final()
        self.imprimir_resumo_desempenho()
        tempo_total_execucao = time.perf_counter() - self.inicio_execucao_geral
        self.registrar_tempo_etapa("TOTAL_EXECUCAO", tempo_total_execucao)
        print(f"Tempo total real da execução: {self.formatar_duracao(tempo_total_execucao)}")
        self.gerar_resumo_tempos()
        self.gerar_resumo_diagnosticos()
    # --- FIM INSERÇÃO EXECUÇÃO PARALELA (COORDENADOR) ---

    def executar(self):
        # --- INÍCIO INSERÇÃO ATIVAÇÃO DO MODO RÁPIDO ---
        if self.modo_rapido:
            # --- INÍCIO INSERÇÃO EXECUÇÃO PARALELA (ATIVAÇÃO) ---
            if self.quantidade_workers > 1:
                return self.executar_paralelo()
            # --- FIM INSERÇÃO EXECUÇÃO PARALELA (ATIVAÇÃO) ---
            return self.executar_modo_rapido()
        # --- FIM INSERÇÃO ATIVAÇÃO DO MODO RÁPIDO ---

        df_contratos = None
        index_atual = None

        try:
            print(f"Lendo base de dados Excel: {self.file_path}")
            df_contratos = pd.read_excel(self.file_path)
            self.validar_colunas_excel(df_contratos)

            df_contratos.dropna(subset=["CONTRATO", "VENCIMENTO"], inplace=True)
            df_contratos["VENCIMENTO"] = pd.to_datetime(
                df_contratos["VENCIMENTO"],
                dayfirst=True,
                errors="coerce",
            ).dt.strftime("%d/%m/%Y")
            df_contratos.dropna(subset=["VENCIMENTO"], inplace=True)

            if "STATUS" not in df_contratos.columns:
                df_contratos["STATUS"] = ""
            if "DETALHES_DOWNLOAD" not in df_contratos.columns:
                df_contratos["DETALHES_DOWNLOAD"] = ""
            if "ULTIMA_EXECUCAO" not in df_contratos.columns:
                df_contratos["ULTIMA_EXECUCAO"] = ""

            for index, row in df_contratos.iterrows():
                index_atual = index
                modulo = row["MÓDULO"]
                contrato = row["CONTRATO"]
                vencimento = row["VENCIMENTO"]

                self.contrato_atual = contrato
                self.vencimento_atual = vencimento
                self.tipo_fatura_atual = None
                self.docs_linha_atual = []
                self.ocorrencia_fatura_atual = None
                # --- INÍCIO INSERÇÃO CONTROLE TXT / LIMPAR ESTADO DO CONTRATO ---
                self.status_docs_fatura_atual = {}
                self.caminhos_docs_fatura_atual = {}
                self.mensagens_fatura_atual = []
                self.controle_fatura_registrado = False
                # --- FIM INSERÇÃO CONTROLE TXT / LIMPAR ESTADO DO CONTRATO ---
                # --- INÍCIO INSERÇÃO CONTROLE TXT / EVIDÊNCIA POR CONTRATO ---
                self.index_linha_atual = index
                self.evidencias_contrato_atual = []
                self.pasta_contrato_atual = self.obter_pasta_contrato(contrato, vencimento)
                self.resumo_execucao["contratos_processados"] += 1
                # --- FIM INSERÇÃO CONTROLE TXT / EVIDÊNCIA POR CONTRATO ---

                print("\n=======================================================")
                print(f"   Processando Linha {index + 1} | Contrato: {contrato} | Venc: {vencimento}")
                print("=======================================================")

                # --- INÍCIO INSERÇÃO REINÍCIO DO NAVEGADOR POR CONTRATO ---
                # Nova regra solicitada: depois de concluir todos os downloads de um contrato,
                # fechar o navegador e recomeçar do zero no próximo contrato.
                self.iniciar_navegador()
                # --- FIM INSERÇÃO REINÍCIO DO NAVEGADOR POR CONTRATO ---

                try:
                    # --- INÍCIO INSERÇÃO REINÍCIO DO NAVEGADOR POR CONTRATO ---
                    self.fazer_login()
                    time.sleep(2)
                    # --- FIM INSERÇÃO REINÍCIO DO NAVEGADOR POR CONTRATO ---

                    self.selecionar_empresa_matriz(modulo)
                    self.buscar_contrato(contrato)
                    self.entrar_no_contrato()
                    self.acessar_extrato_financeiro()

                    houve_download = self.processar_faturas_do_vencimento(vencimento)

                    resultado_contrato = self.classificar_resultado_contrato_atual()
                    if resultado_contrato == "COMPLETO":
                        status = "Sucesso completo - Todos os documentos concluídos"
                        self.resumo_execucao["contratos_sucesso"] += 1
                    elif resultado_contrato == "PARCIAL":
                        status = "Pendência - Download parcial"
                        self.resumo_execucao["contratos_pendencia"] += 1
                    else:
                        status = "Pendência - Nenhum documento concluído"
                        self.resumo_execucao["contratos_pendencia"] += 1

                    detalhes = self.montar_detalhes_resultado_contrato()
                    df_contratos.at[index, "STATUS"] = status
                    df_contratos.at[index, "DETALHES_DOWNLOAD"] = detalhes
                    df_contratos.at[index, "ULTIMA_EXECUCAO"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    print(f"Contrato {contrato} finalizado: {status} | {detalhes}")

                except Exception as e_linha:
                    erro = self.resumir_erro(e_linha)
                    print(f"[ERRO] Falha na linha {index + 1}: {erro}")
                    self.registrar_erro(f"Erro na linha {index + 1}", erro)
                    self.salvar_screenshot(f"erro_linha_{index + 1}")

                    # --- INÍCIO INSERÇÃO CONTROLE TXT / RESUMO POR CONTRATO ---
                    self.resumo_execucao["contratos_erro"] += 1
                    self.adicionar_evidencia_erro_contrato(erro)
                    # --- FIM INSERÇÃO CONTROLE TXT / RESUMO POR CONTRATO ---

                    df_contratos.at[index, "STATUS"] = "Erro - Verificar Logs/Screenshots"
                    df_contratos.at[index, "DETALHES_DOWNLOAD"] = erro[:320]
                    df_contratos.at[index, "ULTIMA_EXECUCAO"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

                finally:
                    # --- INÍCIO INSERÇÃO REINÍCIO DO NAVEGADOR POR CONTRATO ---
                    print("Fechando navegador deste contrato para reiniciar no próximo...")
                    self.fechar_navegador()
                    # --- INÍCIO INSERÇÃO CONTROLE TXT / NÃO ALTERAR EXCEL ORIGINAL ---
                    try:
                        status_evidencia = df_contratos.at[index, "STATUS"] if "STATUS" in df_contratos.columns else ""
                        detalhes_evidencia = df_contratos.at[index, "DETALHES_DOWNLOAD"] if "DETALHES_DOWNLOAD" in df_contratos.columns else ""
                        self.gerar_evidencia_contrato(status_evidencia, detalhes_evidencia)
                    except Exception as e_evidencia:
                        print(f"[AVISO] Falha ao gerar evidência no final do contrato: {e_evidencia}")

                    if self.atualizar_excel_original:
                        try:
                            df_contratos.to_excel(self.file_path, index=False)
                            print("Arquivo Excel atualizado após o contrato.")
                        except Exception as e_excel_linha:
                            print(f"[ERRO] Não foi possível salvar o Excel após o contrato. Feche a planilha se estiver aberta. Erro: {e_excel_linha}")
                    else:
                        print("Arquivo Excel original NÃO foi alterado. Controle salvo em TXT.")
                    # --- FIM INSERÇÃO CONTROLE TXT / NÃO ALTERAR EXCEL ORIGINAL ---
                    time.sleep(3)
                    # --- FIM INSERÇÃO REINÍCIO DO NAVEGADOR POR CONTRATO ---

        except Exception as e:
            erro = self.resumir_erro(e)
            print(f"Ocorreu um erro crítico durante a execução: {erro}")
            self.registrar_erro("Erro Crítico na Execução", erro)
            if df_contratos is not None and index_atual is not None:
                try:
                    df_contratos.at[index_atual, "STATUS"] = "Erro Crítico - Verificar Log de Erros"
                    df_contratos.at[index_atual, "DETALHES_DOWNLOAD"] = erro[:320]
                    df_contratos.at[index_atual, "ULTIMA_EXECUCAO"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                except Exception:
                    pass

        finally:
            # --- INÍCIO INSERÇÃO CONTROLE TXT / NÃO ALTERAR EXCEL ORIGINAL ---
            if self.atualizar_excel_original:
                print("\nSalvando status no arquivo Excel...")
                try:
                    if df_contratos is not None:
                        df_contratos.to_excel(self.file_path, index=False)
                        print("Arquivo Excel atualizado com sucesso!")
                except Exception as e:
                    print(f"[ERRO] Não foi possível salvar o Excel. Feche a planilha se estiver aberta. Erro: {e}")
            else:
                print("\nArquivo Excel original mantido sem alteração. Controle da execução foi gravado em TXT.")
            # --- FIM INSERÇÃO CONTROLE TXT / NÃO ALTERAR EXCEL ORIGINAL ---

            # --- INÍCIO INSERÇÃO RESUMO FINAL DETALHADO ---
            self.imprimir_resumo_final()
            self.gerar_resumo_diagnosticos()
            # --- FIM INSERÇÃO RESUMO FINAL DETALHADO ---

            self.fechar_navegador()

# --- FIM INSERÇÃO PLAYWRIGHT (CLASSE MANTENDO FLUXO ORIGINAL) ---

if __name__ == "__main__":
    print("\n=======================================================")
    print("         ROBO DE DOWNLOADS NDI - HAPVIDA")
    print("     RELATÓRIOS ANALÍTICOS - VERSÃO PLAYWRIGHT")
    print("=======================================================\n")

    diretorio_atual = Path(__file__).resolve().parent
    nome_arquivo = "NDI_CONTRATOS_DOWNLOAD.xlsx"
    arquivo_excel = diretorio_atual / nome_arquivo

    automacao = HapvidaNDIAnaliticosAutomation(arquivo_excel)
    automacao.executar()

    print("\n=======================================================")
    print("         AUTOMACAO FINALIZADA!")
    print("=======================================================")
    os.system("pause")
