#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Gravador de Portais Playwright — v1.2 — Python 3.10+ / Windows.

INSTALAÇÃO (Prompt de Comando):
    py -m pip install playwright
    py -m playwright install chromium
    py Gravador_Portais_Playwright.py

USO:
1. Informe a URL e abra o navegador. A gravação começa PAUSADA.
2. Faça login manualmente; clique em Iniciar/retomar e demonstre o fluxo.
3. Use Marcar etapa para informar contrato, competência e tipo de documento.
4. Finalize pela janela do gravador; aguarde os downloads e a exportação.

SAÍDAS: eventos.jsonl (incremental), gravacao.json, relatorio.md,
         downloads.csv, execucao.log, downloads/ e capturas/ (somente por solicitação).
O JSONL preserva os eventos já escritos caso o processo seja interrompido.
Para recuperar os relatórios: py Gravador_Portais_Playwright.py --recuperar PASTA

Login automático opcional: configure os seletores CSS dos três controles
pela interface e PORTAL_USUARIO / PORTAL_SENHA no ambiente. A confirmação
de que o login terminou é manual: clique Iniciar/retomar após entrar.
O programa não grava cookies, cabeçalhos de autenticação ou estado de sessão.
Com Capturar PDFs habilitado, salva respostas PDF recebidas durante a gravação.
Não refaz requisições HTTP: lê a resposta existente; para blob, tenta ler o
objeto local nas abas abertas. PDF limitado a 50 MiB (a leitura HTTP pode
alocar o corpo inteiro antes de conferir o tamanho quando este é desconhecido).
Valores de campos comuns são gravados; senha, OTP e campos reconhecidos como
sensíveis são mascarados. A classificação é heurística: pause para segredos
em campos não convencionais. Capturas solicitadas podem conter dados pessoais.

LIMITES: não é um executor automático da gravação. Os seletores são candidatos
para revisão; páginas canvas, Shadow DOM fechado, interfaces do navegador,
janelas de impressão/Salvar como e CAPTCHA não são mapeados como DOM comum.
Downloads são salvos quando o navegador emite o evento download; PDFs abertos
no visualizador também têm tentativa de captura pela resposta HTTP.
ZIPs são verificados por CRC sem extração, com limite de 256 MiB descompactados,
10.000 entradas e 30 segundos. ZIP criptografado exige verificação manual.
Arquivos idênticos na sessão compartilham a mesma cópia por SHA-256.
A captura alternativa depende da disponibilidade da resposta/blob no navegador.
Eventos de teclado/clique relacionados permanecem no histórico bruto e recebem
indicação para não serem reproduzidos duas vezes. O roteiro exige revisão.
Sem interceptação de rede, alteração do portal ou aceitação automática de diálogos.
"""
from __future__ import annotations

import argparse
import asyncio
import csv
import base64
import zipfile
import hashlib
import json
import os
from pathlib import Path
import queue
import re
import sys
import threading
import time
from datetime import datetime, timezone
from urllib.parse import urlsplit, urlunsplit
import uuid

VERSION = '1.2'
PDF_LIMIT = 50 * 1024 * 1024
ZIP_MAX_BYTES = 256 * 1024 * 1024
ZIP_MAX_ENTRIES = 10000
ZIP_MAX_SECONDS = 30
# Windows: MAX_PATH = 260 incluindo o terminador; margem para o prefixo uuid.
WINDOWS_PATH_LIMIT = 250
# Substrings de content-type tratadas como documento (XML de NF-e/CT-e, planilhas, ZIP etc.).
DOCUMENT_MIME_PARTS = ('application/pdf', 'text/csv', 'text/plain', 'octet-stream', 'zip',
                       'application/xml', 'text/xml', 'ms-excel', 'spreadsheetml',
                       'wordprocessingml', 'msword')

# Executado em cada documento e iframe, inclusive os criados posteriormente.
JS = r"""(() => {
if (window.__portalRecorder) return;
window.__portalRecorder = true;
let enabled = false, serial = 0, lastFocus = null, pendingTab = null;
let lastPointer = null, activation = null, labelClick = null;
const documentId = Math.random().toString(36).slice(2);
const dirty = new Set(); let values = new WeakMap();
const clip = x => String(x || '').replace(/\s+/g, ' ').trim().slice(0, 240);
const sensitive = el => !!el && (/password/i.test(el.type || '') ||
 /pass|senha|secret|token|otp|verification|one.time|security.code|cvv|credit.card/i.test(
 [el.name, el.id, el.autocomplete, el.getAttribute('aria-label'), el.placeholder].join(' ')));
function path(el) {
 const parts = [];
 while (el && el.nodeType === 1 && parts.length < 12) {
  let s = el.localName;
  if (el.id && !dynamicId(el.id)) { parts.unshift('#' + CSS.escape(el.id)); break; }
  const parent = el.parentElement;
  if (parent) {
   const same = [...parent.children].filter(x => x.localName === el.localName);
   if (same.length > 1) s += ':nth-of-type(' + (same.indexOf(el) + 1) + ')';
  }
  parts.unshift(s); el = parent;
 }
 return parts.join(' > ');
}
function dynamicId(id) { return /^:r|^radix-|^headlessui-|^[a-f0-9]{16,}$/i.test(id || ''); }
function contextOf(el) {
 const ancestors = []; let node = el;
 for (let i = 0; node && i < 5; i++) {
  const parent = node.parentElement || node.getRootNode()?.host;
  if (!parent) break;
  if (!['BODY','HTML'].includes(parent.tagName)) {
   const text = sensitive(parent) ? '' : clip(parent.innerText || parent.textContent);
   ancestors.push({tag:parent.localName,text,css:path(parent),
    headings:[...parent.querySelectorAll('h1,h2,h3,h4,legend,[role="heading"]')].slice(0,4).map(x=>clip(x.textContent))});
  }
  node = parent;
 }
 const row = el.closest('tr,[role="row"]');
 return {row:row ? clip(row.innerText) : '', ancestors};
}
function control(e) {
 return e.composedPath().find(x => x?.nodeType === 1 &&
  x.matches('button,a,input,select,textarea,[role="button"],[role="link"],[role="combobox"],[role="option"],label,cr-icon-button')) || target(e);
}
function describe(el) {
 if (!el || el.nodeType !== 1) return null;
 const secret = sensitive(el), candidates = [], root = el.getRootNode();
 const add = (kind, value, extra = {}) => {
  if (!value) return;
  const item = {kind, value, ...extra};
  if (kind === 'css') item.matches = count(value);
  candidates.push(item);
 };
 const count = selector => {
  try { return root.querySelectorAll(selector).length; }
  catch (_) { return -1; }
 };
 const testId = el.getAttribute('data-testid');
 if (testId) add('test_id', testId, {matches:count('[data-testid="' + CSS.escape(testId) + '"]')});
 if (el.id) add('css', '#' + CSS.escape(el.id), {dynamic:dynamicId(el.id), fragile:dynamicId(el.id)});
 const label = clip(el.labels && [...el.labels].map(x => x.textContent).join(' '));
 if (label) add('label', label);
 const aria = clip(el.getAttribute('aria-label'));
 const role = el.getAttribute('role') || ({button:'button',a:'link',select:'combobox',textarea:'textbox'}[el.localName]) || '';
 const text = secret ? '' : clip(el.innerText || (['button','submit','reset'].includes(el.type) ? el.value : ''));
 if (role && (aria || text)) add('role', role, {name: aria || text, heuristic: true});
 if (el.getAttribute('name')) add('css', el.localName + '[name=' + JSON.stringify(el.getAttribute('name')) + ']');
 if (el.placeholder && !secret) add('placeholder', el.placeholder, {matches:count('[placeholder="' + CSS.escape(el.placeholder) + '"]')});
 if (text && text.length < 150) add('text', text);
 add('css', path(el), {fragile: true});
 // Candidatos que correspondem a mais de um elemento descem abaixo dos únicos.
 const rank = c => c.fragile ? 90 : ({test_id:0,label:1,role:2,placeholder:3,css:4,text:5}[c.kind] ?? 50) + (c.matches > 1 ? 10 : 0);
 candidates.sort((a,b) => rank(a)-rank(b));
 const hosts = []; let r = root;
 while (r && r.host) { hosts.unshift(path(r.host)); r = r.host.getRootNode(); }
 return {tag: el.localName, type: el.type || '', label, text, sensitive: secret,
  candidates, shadow_hosts: hosts, context:contextOf(el)};
}
function emit(kind, data = {}) {
 if (!enabled) return;
 const payload = {kind, document_id: documentId, local_seq: ++serial,
  browser_time: new Date().toISOString(), ...data};
 window.__recordPortal(payload).catch(() => {});
 return documentId + ":" + serial;
}
function target(e) { return e.composedPath().find(x => x && x.nodeType === 1); }
function flushOne(el) {
 if (!dirty.has(el)) return;
 dirty.delete(el);
 let value;
 if (sensitive(el)) value = '[PROTEGIDO]';
 else if (el.type === 'file') value = '[ARQUIVO LOCAL]';
 else if (el.isContentEditable) value = (el.textContent || '').slice(0, 4000);
 else value = String(el.value ?? '').slice(0, 4000);
 if (values.get(el) === value) return;
 values.set(el, value);
 emit('fill', {element: describe(el), value});
}
window.__flushPortal = () => { for (const el of dirty) flushOne(el); };
window.__setPortalRecording = state => {
 if (!state) window.__flushPortal();
 enabled = !!state; dirty.clear(); valuesReset(); pendingTab = null; lastPointer = null; activation = null; labelClick = null;
};
function valuesReset() { lastFocus = null; values = new WeakMap(); }
window.__recordPortal({kind:'recorder_ready'}).then(state => {enabled = !!state;}).catch(() => {});
document.addEventListener('input', e => {
 if (!enabled) return;
 const el = target(e); if (el) dirty.add(el);
}, true);
document.addEventListener('focusout', e => flushOne(target(e)), true);
document.addEventListener('change', e => {
 const el = target(e); if (!el || !enabled) return;
 if (el.tagName === 'SELECT') {
  dirty.delete(el);
  emit('select', {element: describe(el), values: sensitive(el) ? ['[PROTEGIDO]'] : [...el.selectedOptions].map(x => x.value)});
 } else if (['checkbox','radio'].includes(el.type)) {
  dirty.delete(el); emit('check', {element:describe(el), checked:el.checked});
 } else { dirty.add(el); flushOne(el); }
}, true);
document.addEventListener('pointerdown', e => {
 if (!enabled) return;
 window.__flushPortal();
 const el = control(e), info = describe(el);
 const id = emit('pointer_down', {element:info,actual_target:describe(target(e)),
  is_trusted:e.isTrusted,button:e.button,x:e.clientX,y:e.clientY,replay_hint:'context_only'});
 lastPointer = {id,at:performance.now(),element:info};
 activation = null;
}, true);
document.addEventListener('click', e => {
 if (!enabled) return;
 window.__flushPortal();
 const el = control(e);
 const linkedKey = e.detail === 0 && activation && activation.el === el &&
  performance.now()-activation.at < 1500 ? activation.id : null;
 // Clique no <label> é repassado pelo navegador ao controle associado.
 const linkedLabel = !linkedKey && labelClick && labelClick.control === el &&
  performance.now()-labelClick.at < 1000 ? labelClick.id : null;
 const recent = lastPointer && performance.now()-lastPointer.at < 2000;
 const id = emit('click', {element:describe(el), actual_target:describe(target(e)),
  is_trusted:e.isTrusted, generated_by_key:linkedKey, generated_by_label:linkedLabel,
  pointer_id:recent ? lastPointer.id : null,
  pointer_element:recent ? lastPointer.element : null,
  replay_hint:linkedKey ? 'effect_of_key_do_not_repeat' : (linkedLabel ? 'effect_of_label_do_not_repeat' :
   (!e.isTrusted ? 'script_generated_review' : 'action')),
  button:e.button,detail:e.detail,x:e.clientX,y:e.clientY,
  modifiers:{ctrl:e.ctrlKey,alt:e.altKey,shift:e.shiftKey,meta:e.metaKey}});
 if (linkedKey) activation = null;
 labelClick = el?.localName === 'label' && el.control ? {id,control:el.control,at:performance.now()} : null;
}, true);
document.addEventListener('dblclick', e => { if (enabled) emit('double_click', {element:describe(target(e))}); }, true);
document.addEventListener('contextmenu', e => { if (enabled) emit('context_menu', {element:describe(target(e))}); }, true);
document.addEventListener('keydown', e => {
 if (!enabled) return;
 // Preenchimento automático do Chrome dispara keydown sem e.key.
 const el = target(e), key = typeof e.key === 'string' ? e.key : '';
 if (!key) return;
 // Não captura caracteres digitados nem atalhos dentro de campos protegidos.
 if (sensitive(el) && !['Tab','Enter','Escape'].includes(key)) return;
 if (key.length === 1 && !(key === ' ' && el?.matches('button,[role="button"],input[type="checkbox"]')) && !e.ctrlKey && !e.altKey && !e.metaKey) return;
 if (['Shift','Control','Alt','Meta'].includes(key)) return;
 window.__flushPortal();
 const keyId = documentId + ':' + (serial + 1);
 emit('key', {key_id:keyId,is_trusted:e.isTrusted,key,code:e.code,repeat:e.repeat,element:describe(el),
  modifiers:{ctrl:e.ctrlKey,alt:e.altKey,shift:e.shiftKey,meta:e.metaKey}});
 if (['Enter',' '].includes(key)) activation = {id:keyId,el:control(e),at:performance.now()};
 if (key === 'Tab') pendingTab = {key_id:keyId, from:describe(el), at:performance.now()};
}, true);
document.addEventListener('focusin', e => {
 if (!enabled) return;
 const current = describe(target(e));
 emit('focus', {from:lastFocus,to:current}); lastFocus = current;
 if (pendingTab && performance.now() - pendingTab.at < 2000) {
  emit('tab_destination', {key_id:pendingTab.key_id,from:pendingTab.from,to:current});
 }
 pendingTab = null;
}, true);
document.addEventListener('submit', e => {if (!enabled) return; window.__flushPortal(); emit('submit',{element:describe(target(e))});}, true);
// Um temporizador por contêiner: rolagens simultâneas não se sobrescrevem.
const scrollTimers = new WeakMap();
document.addEventListener('scroll', e => {
 if (!enabled) return;
 const el = target(e), key = el || document;
 clearTimeout(scrollTimers.get(key));
 scrollTimers.set(key, setTimeout(() => {
  scrollTimers.delete(key);
  emit('scroll',{element:describe(el), x:el ? el.scrollLeft : scrollX,y:el ? el.scrollTop : scrollY});
 }, 250));
}, true);
window.addEventListener('beforeunload', () => window.__flushPortal());
window.addEventListener('beforeprint', () => emit('print_requested'));
window.addEventListener('afterprint', () => emit('print_closed'));
})();"""


def safe_url(url):
    """Preserva caminho; remove usuário, senha, query e fragmento."""
    try:
        p = urlsplit(url)
        if p.scheme not in ('http', 'https'):
            return p.scheme + ':' if p.scheme else ''
        host = p.hostname or ''
        if ':' in host:
            host = '[' + host + ']'
        if p.port:
            host += ':' + str(p.port)
        return urlunsplit((p.scheme, host, p.path, '', ''))
    except ValueError:
        return '[URL inválida]'


def safe_name(name, max_len=160):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).strip(' .')
    if len(name) > max_len:
        # Trunca o nome preservando a extensão.
        stem, dot, ext = name.rpartition('.')
        if dot and stem and len(ext) <= 10:
            name = stem[:max_len - len(ext) - 1].rstrip(' .') + '.' + ext
        else:
            name = name[:max_len].rstrip(' .')
    if name.split('.')[0].upper() in {'CON','PRN','AUX','NUL', *('COM'+str(i) for i in range(1,10)), *('LPT'+str(i) for i in range(1,10))}:
        name = '_' + name
    return name or 'documento'


def clean_error(error):
    """Remove valores de login, URLs assinadas e pares sensíveis do diagnóstico."""
    text = str(error)
    for key in ('PORTAL_USUARIO', 'PORTAL_SENHA'):
        value = os.getenv(key)
        if value:
            text = text.replace(value, '[PROTEGIDO]')
    text = re.sub(r'https?://[^\s<>"\']+', lambda m: safe_url(m.group()), text)
    text = re.sub(r'(?i)(authorization|cookie|password|senha|access_token|refresh_token|id_token|token|secret|sig)(\s*[:=]\s*)([^\s,;]+)', r'\1\2[PROTEGIDO]', text)
    text = re.sub(r'(?i)Bearer\s+\S+', 'Bearer [PROTEGIDO]', text)
    return text[:2000]


def inspect_zip(path):
    result = {'zip_members': [], 'zip_members_truncated': False}
    try:
        with zipfile.ZipFile(path) as z:
            items = z.infolist()
            result['zip_members'] = [{'name':i.filename, 'bytes':i.file_size, 'compressed_bytes':i.compress_size} for i in items[:1000]]
            result['zip_members_truncated'] = len(items) > 1000
            result['zip_total_members'] = len(items)
            if len(items) > ZIP_MAX_ENTRIES or sum(i.file_size for i in items) > ZIP_MAX_BYTES:
                result['validation'] = 'ZIP_LIMITE_VERIFICACAO'
            elif any(i.flag_bits & 1 for i in items):
                result['validation'] = 'ZIP_CRIPTOGRAFADO'
            elif not any(not i.is_dir() for i in items):
                result['validation'] = 'ZIP_SEM_ARQUIVOS'
            else:
                started = time.monotonic()
                total = 0
                for item in items:
                    if item.is_dir():
                        continue
                    with z.open(item) as member:
                        while True:
                            block = member.read(1024*1024)
                            if not block:
                                break
                            total += len(block)
                            if total > ZIP_MAX_BYTES or time.monotonic()-started > ZIP_MAX_SECONDS:
                                result['validation'] = 'ZIP_LIMITE_VERIFICACAO'
                                return result
                result['validation'] = 'ZIP_CRC_OK'
    except Exception as exc:
        result.update(validation='ZIP_FALHA_VERIFICACAO', validation_error=clean_error(exc))
    return result


def inspect_file(path):
    size = path.stat().st_size
    digest = hashlib.sha256()
    with path.open('rb') as f:
        head = f.read(8192)
        digest.update(head)
        for block in iter(lambda: f.read(1024 * 1024), b''):
            digest.update(block)
    prefix = head.lstrip().lower()
    if not size:
        status = 'VAZIO'
    elif prefix.startswith((b'<!doctype html', b'<html')):
        status = 'ATENCAO_HTML'
    elif path.suffix.lower() == '.pdf' or b'%PDF-' in head[:1024]:
        status = 'ASSINATURA_PDF_OK' if b'%PDF-' in head[:1024] else 'PDF_SEM_ASSINATURA'
    elif path.suffix.lower() in ('.csv', '.txt'):
        status = 'TEXTO_PLAUSIVEL' if b'\x00' not in head else 'VERIFICAR_CODIFICACAO'
    else:
        status = 'NAO_VALIDADO'
    result = {'bytes': size, 'sha256': digest.hexdigest(), 'validation': status}
    if size and (path.suffix.lower() == '.zip' or head.startswith(b'PK\x03\x04')):
        result.update(inspect_zip(path))
    return result


def export_session(folder):
    events, invalid = [], 0
    with (folder / 'eventos.jsonl').open(encoding='utf-8') as f:
        for line in f:
            if not line.strip():
                continue
            try:
                item = json.loads(line)
            except json.JSONDecodeError:
                invalid += 1
                continue
            # Linha JSON válida, mas sem estrutura de evento, também é descartada.
            if isinstance(item, dict) and isinstance(item.get('kind'), str):
                events.append(item)
            else:
                invalid += 1
    recorded_version = next((e.get('version', VERSION) for e in events if e.get('kind') == 'session_started'), VERSION)
    result = {'version': recorded_version, 'exporter_version': VERSION, 'invalid_lines': invalid, 'events': events}
    temporary = folder / 'gravacao.json.tmp'
    temporary.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    temporary.replace(folder / 'gravacao.json')
    columns = ['id', 'time', 'kind', 'status', 'page', 'stage', 'attempt_id', 'method', 'file',
               'bytes', 'sha256', 'validation', 'duration_s', 'trigger_candidate', 'duplicate_of',
               'error', 'browser_failure', 'zip_members']
    terminal = ('download_finished', 'download_failed', 'capture_finished', 'capture_failed')
    finished_ids = {e.get('attempt_id') for e in events if e['kind'] in terminal}
    rows = []
    for e in events:
        if e['kind'] in terminal or (e['kind'] in ('download_started','capture_started') and e.get('attempt_id') not in finished_ids):
            row = dict(e)
            row['status'] = 'SALVO' if e['kind'].endswith('_finished') else ('FALHOU' if e['kind'].endswith('_failed') else 'INCOMPLETO')
            row['error'] = e.get('error', e.get('note','')) if row['status'] != 'SALVO' else ''
            row['zip_members'] = json.dumps(e.get('zip_members', []), ensure_ascii=False)
            rows.append(row)
    def csv_cell(value):
        if isinstance(value,str) and (value.startswith(('\t','\r')) or value.lstrip().startswith(('=','+','-','@'))):
            return "'" + value
        return value
    with (folder / 'downloads.csv').open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=columns, delimiter=';', extrasaction='ignore')
        writer.writeheader()
        writer.writerows({k:csv_cell(v) for k,v in row.items()} for row in rows)
    def md(value):
        return str(value).replace('|', r'\|').replace('\n', ' ')[:900]
    def element_name(el):
        if not el:
            return '(sem elemento)'
        return el.get('label') or el.get('text') or next((c['value'] for c in el.get('candidates',[]) if c['kind']=='css'), el.get('tag',''))
    lines = ['# Gravação de portal', '', f'Eventos: {len(events)}. Linhas incompletas ignoradas: {invalid}.', '',
             'Arquivos salvos: %s; tentativas com falha: %s; incompletas: %s.' % tuple(sum(r['status']==v for r in rows) for v in ('SALVO','FALHOU','INCOMPLETO')),
             'Contagens representam operações; arquivos idênticos podem compartilhar uma cópia.',
             'Seletores/contextos e relações entre ações são candidatos heurísticos para revisão.',
             'CRC verifica o ZIP; assinatura PDF não confirma dados fiscais nem integridade completa do PDF.', '',
             '| Nº | Hora | Aba | Etapa | Evento | Detalhe |', '|---|---|---|---|---|---|']
    for e in events:
        kind = e['kind']
        detail = e.get('file') or e.get('url') or e.get('note') or ''
        if kind == 'key':
            mods = e.get('modifiers',{})
            detail = '+'.join([name for key,name in [('ctrl','Ctrl'),('alt','Alt'),('shift','Shift'),('meta','Meta')] if mods.get(key)] + [str(e.get('key',''))]) + ' → ' + element_name(e.get('element'))
        elif kind == 'tab_destination':
            detail = element_name(e.get('from')) + ' → ' + element_name(e.get('to'))
        elif kind == 'focus':
            detail = element_name(e.get('to'))
        elif e.get('element'):
            detail = element_name(e['element'])
            ancestors = e['element'].get('context',{}).get('ancestors',[])
            if ancestors:
                detail += ' | contexto: ' + ancestors[min(1,len(ancestors)-1)].get('text','')
        if e.get('replay_hint'):
            detail += ' [' + e['replay_hint'] + ']'
        if e.get('error'):
            detail += ' | ' + e['error']
        if e.get('validation'):
            detail += ' | ' + e['validation']
        if 'duration_s' in e:
            detail += ' | ' + str(e['duration_s']) + ' s'
        lines.append('| ' + ' | '.join(md(e.get(k, '')) for k in ['id', 'time', 'page', 'stage', 'kind']) + ' | ' + md(detail) + ' |')
    (folder / 'relatorio.md').write_text('\n'.join(lines), encoding='utf-8')
    return len(events)


class Recorder:
    def __init__(self, config, commands, updates):
        self.config, self.commands, self.updates = config, commands, updates
        self.active = False
        self.stage = ''
        self.seq = 0
        self.pages, self.frames = {}, {}
        self.tasks = set()
        self.last_action = {}
        self.event_ids = {}
        self.requests = {}
        self.parents = {}
        self.hash_files = {}
        self.file_lock = asyncio.Lock()
        self.pdf_semaphore = asyncio.Semaphore(2)
        self.folder = Path(config['output']) / ('sessao_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '_' + uuid.uuid4().hex[:6])
        self.folder.mkdir(parents=True)
        (self.folder / 'downloads').mkdir()
        (self.folder / 'capturas').mkdir()
        self.stream = (self.folder / 'eventos.jsonl').open('a', encoding='utf-8', buffering=1)
        self.context = None

    def inform(self, message):
        with (self.folder / 'execucao.log').open('a', encoding='utf-8') as log:
            log.write(datetime.now().isoformat(timespec='seconds') + ' | ' + message + '\n')
        self.updates.put(('log', message))

    def event(self, kind, **data):
        self.seq += 1
        event = {'id': self.seq, 'time': datetime.now(timezone.utc).isoformat(),
                 'kind': kind, 'stage': self.stage, **data}
        self.stream.write(json.dumps(event, ensure_ascii=False) + '\n')
        self.stream.flush()
        self.updates.put(('count', self.seq))
        return self.seq

    def page_id(self, page):
        if page not in self.pages:
            self.pages[page] = 'aba_' + str(len(self.pages) + 1)
        return self.pages[page]

    def frame_info(self, frame):
        chain = []
        while frame:
            if frame not in self.frames:
                self.frames[frame] = 'frame_' + str(len(self.frames) + 1)
            chain.insert(0, {'id': self.frames[frame], 'name': frame.name, 'url': safe_url(frame.url)})
            frame = frame.parent_frame
        return chain

    def spawn(self, coro):
        task = asyncio.create_task(coro)
        self.tasks.add(task)
        task.add_done_callback(self.task_done)

    def task_done(self, task):
        self.tasks.discard(task)
        if not task.cancelled() and task.exception():
            self.event('task_error', error=clean_error(task.exception()), error_type=type(task.exception()).__name__)
            self.inform('Falha em tarefa auxiliar; consulte os eventos.')

    async def binding(self, source, payload):
        if payload.get('kind') == 'recorder_ready':
            return self.active
        if not self.active:
            return False
        page = source['page']
        kind = payload.pop('kind', 'unknown')
        number = self.event(kind, page=self.page_id(page), frames=self.frame_info(source['frame']), **payload)
        local_id = str(payload.get('document_id','')) + ':' + str(payload.get('local_seq',''))
        self.event_ids[local_id] = number
        if len(self.event_ids) > 10000:
            self.event_ids.pop(next(iter(self.event_ids)))
        generated_key = payload.get('generated_by_key')
        generated_label = payload.get('generated_by_label')
        if generated_key:
            self.event('action_relation', page=self.page_id(page), source_event=self.event_ids.get(generated_key),
                       effect_event=number, note='Clique resultante do teclado; não reproduzir ambos.')
        elif generated_label:
            self.event('action_relation', page=self.page_id(page), source_event=self.event_ids.get(generated_label),
                       effect_event=number, note='Clique repassado pelo label; não reproduzir ambos.')
        if kind in ('click', 'key', 'select') and payload.get('is_trusted',True) and not generated_key and not generated_label:
            self.last_action[page] = (number, time.monotonic())
        return True

    async def set_active(self, active):
        # Flush antes de desligar; Python continua aceitando o último valor.
        if not active:
            await self.flush()
        self.active = active
        await self.evaluate_frames('(v) => window.__setPortalRecording && window.__setPortalRecording(v)', active)
        self.event('recording_started' if active else 'recording_paused')
        self.updates.put(('state', 'GRAVANDO' if active else 'PAUSADO'))

    async def evaluate_frames(self, script, arg=None):
        async def one(frame):
            try:
                await frame.evaluate(script, arg)
            except Exception:
                pass  # frame pode desaparecer durante navegação
        frames = [frame for page in list(self.context.pages) for frame in page.frames]
        await asyncio.gather(*(one(frame) for frame in frames))

    async def flush(self):
        await self.evaluate_frames('() => window.__flushPortal && window.__flushPortal()')
        await asyncio.sleep(0.1)

    def attach(self, page):
        pid = self.page_id(page)
        self.event('page_opened', page=pid, url=safe_url(page.url))
        page.on('download', lambda download: self.spawn(self.download(page, download)) if self.active else None)
        page.on('framenavigated', lambda frame: self.event('navigation', page=pid, frames=self.frame_info(frame), url=safe_url(frame.url)) if self.active else None)
        page.on('close', lambda: self.event('page_closed', page=pid))
        page.on('pageerror', lambda error: self.event('page_error', page=pid, note='Erro JavaScript no portal; mensagem omitida para evitar dados sensíveis.') if self.active else None)
        page.on('request', lambda request: self.request_started(page, request))
        page.on('requestfailed', lambda request: self.request_failed(page, request))
        page.on('response', lambda response: self.response(page, response))
        # Não registra dialog listener: preserva o tratamento padrão do Playwright.
        self.spawn(self.opener(page))

    async def opener(self, page):
        parent = await page.opener()
        if parent:
            self.parents[page] = parent
            self.event('popup_parent', page=self.page_id(page), parent=self.page_id(parent))
            if parent in self.last_action:
                self.last_action[page] = self.last_action[parent]

    def trigger(self, page):
        candidate = self.last_action.get(page)
        if candidate and time.monotonic()-candidate[1] <= 180:
            return candidate[0]
        return None

    def request_started(self, page, request):
        if self.active:
            self.requests[request] = {'page':self.page_id(page), 'stage':self.stage,
                                      'trigger_candidate':self.trigger(page), 'started':time.monotonic()}
            if len(self.requests) > 5000:
                self.requests.pop(next(iter(self.requests)))

    def request_failed(self, page, request):
        meta = self.requests.pop(request, None)
        if meta:
            self.event('request_failed', page=self.page_id(page), stage=meta['stage'],
                       url=safe_url(request.url), error=clean_error(request.failure or 'Falha de rede'),
                       trigger_candidate=meta['trigger_candidate'])

    def response(self, page, response):
        meta = self.requests.pop(response.request, None)
        if not self.active and not meta:
            return
        meta = meta or {'stage':self.stage, 'trigger_candidate':self.trigger(page), 'started':time.monotonic()}
        mime = response.headers.get('content-type', '').split(';')[0].strip().lower()
        attachment = response.headers.get('content-disposition', '').strip().lower().startswith('attachment')
        if attachment or any(x in mime for x in DOCUMENT_MIME_PARTS):
            self.event('document_response', page=self.page_id(page), stage=meta['stage'], url=safe_url(response.url),
                       status=response.status, mime=mime, attachment=attachment, trigger_candidate=meta['trigger_candidate'],
                       response_wait_s=round(time.monotonic()-meta['started'],3),
                       note='Resposta observada; consulte capture_finished/download_finished para arquivo salvo.')
        elif response.status >= 400 and response.request.resource_type in ('document', 'xhr', 'fetch'):
            # Ignora 4xx/5xx de imagens, fontes e scripts de terceiros.
            self.event('http_error', page=self.page_id(page), url=safe_url(response.url), status=response.status)
        if self.config.get('capture_pdf', True) and mime == 'application/pdf' and response.status == 200 and response.url.startswith(('http://','https://')):
            self.spawn(self.capture_response(page, response, meta))

    async def finish_file(self, path, meta, kind, started):
        info = await asyncio.to_thread(inspect_file, path)
        duplicate = None
        async with self.file_lock:
            duplicate = self.hash_files.get(info['sha256'])
            if duplicate and duplicate != path.name:
                path.unlink(missing_ok=True)
                name = duplicate
            else:
                name = path.name
                self.hash_files[info['sha256']] = name
        self.event(kind, **meta, file=name, duplicate_of=duplicate,
                   duration_s=round(time.monotonic()-started,3), **info)
        self.inform(('Arquivo já salvo: ' if duplicate else 'Arquivo salvo: ') + name + ' — ' + info['validation'])

    async def capture_response(self, page, response, request_meta):
        started = time.monotonic()
        name = 'pdf_resposta_' + uuid.uuid4().hex[:10] + '.pdf'
        path = self.folder / 'downloads' / name
        meta = {'attempt_id':uuid.uuid4().hex, 'page':self.page_id(page), 'stage':request_meta['stage'],
                'method':'response_pdf', 'trigger_candidate':request_meta['trigger_candidate'], 'url':safe_url(response.url)}
        self.event('capture_started', **meta, file=name)
        try:
            limit = PDF_LIMIT
            length = response.headers.get('content-length','')
            if length.isdigit() and int(length)>limit:
                raise ValueError('PDF excede limite de captura de 50 MiB')
            async with self.pdf_semaphore:
                data = await asyncio.wait_for(response.body(), timeout=180)
                if len(data)>limit:
                    raise ValueError('PDF excede limite de captura de 50 MiB')
                if b'%PDF-' not in data[:1024]:
                    raise ValueError('Resposta declarada PDF sem assinatura PDF')
                await asyncio.to_thread(path.write_bytes, data)
                del data
            await self.finish_file(path, meta, 'capture_finished', started)
        except Exception as exc:
            self.event('capture_failed', **meta, file=name, error=clean_error(exc),
                       error_type=type(exc).__name__, duration_s=round(time.monotonic()-started,3))
            self.inform('Captura PDF não concluída: ' + clean_error(exc))

    async def capture_blob(self, page, url, parent_attempt, stage):
        """Lê somente o blob local; não envia uma segunda requisição HTTP."""
        started = time.monotonic()
        meta = {'attempt_id':uuid.uuid4().hex, 'page':self.page_id(page), 'stage':stage,
                'method':'blob_pdf', 'parent_attempt':parent_attempt, 'trigger_candidate':self.trigger(page)}
        name = 'pdf_blob_' + uuid.uuid4().hex[:10] + '.pdf'
        self.event('capture_started', **meta, file=name)
        errors = []
        candidates = list(dict.fromkeys([page, self.parents.get(page)] + list(self.context.pages)))
        for candidate in candidates:
            if candidate is None or candidate.is_closed():
                continue
            try:
                encoded = await asyncio.wait_for(candidate.evaluate("""async ({url, limit}) => {
                    if (!url.startsWith('blob:')) throw new Error('Somente blob local');
                    const response = await fetch(url);
                    if (!response.ok) throw new Error('Falha ao ler blob');
                    const blob = await response.blob();
                    if (blob.size > limit) throw new Error('Blob excede 50 MiB');
                    return await new Promise((resolve,reject) => {
                        const reader = new FileReader();
                        reader.onload = () => resolve(reader.result.split(',')[1]);
                        reader.onerror = () => reject(new Error('Falha FileReader'));
                        reader.readAsDataURL(blob);
                    });
                }""", {'url': url, 'limit': PDF_LIMIT}), timeout=15)
                data = base64.b64decode(encoded, validate=True)
                if b'%PDF-' not in data[:1024]:
                    raise ValueError('Blob não contém assinatura PDF')
                path = self.folder / 'downloads' / name
                await asyncio.to_thread(path.write_bytes, data)
                await self.finish_file(path, meta, 'capture_finished', started)
                return
            except Exception as exc:
                errors.append(type(exc).__name__ + ': ' + clean_error(exc))
        self.event('capture_failed', **meta, file=name, duration_s=round(time.monotonic()-started,3),
                   error=' | '.join(errors)[-2000:] or 'Nenhuma aba disponível para ler o blob')

    async def download(self, page, download):
        started = time.monotonic()
        folder = self.folder / 'downloads'
        budget = max(40, min(160, WINDOWS_PATH_LIMIT - len(str(folder.resolve())) - 12))
        name = uuid.uuid4().hex[:10] + '_' + safe_name(download.suggested_filename, budget)
        path = folder / name
        meta = {'attempt_id':uuid.uuid4().hex, 'page':self.page_id(page), 'stage':self.stage,
                'method':'download', 'trigger_candidate':self.trigger(page), 'url':safe_url(download.url)}
        self.event('download_started', **meta, file=name)
        try:
            await asyncio.wait_for(download.save_as(str(path)), timeout=300)
            await self.finish_file(path, meta, 'download_finished', started)
        except Exception as exc:
            failure = ''
            try:
                failure = await asyncio.wait_for(download.failure(),timeout=3) or ''
            except Exception:
                pass
            self.event('download_failed', **meta, file=name, error=clean_error(exc),
                       browser_failure=clean_error(failure), error_type=type(exc).__name__,
                       duration_s=round(time.monotonic()-started,3))
            self.inform('Download falhou: ' + name + ' | ' + clean_error(failure or exc))
            if self.config.get('capture_pdf',True) and download.url.startswith('blob:'):
                await self.capture_blob(page, download.url, meta['attempt_id'], meta['stage'])

    async def screenshot(self):
        for page in list(self.context.pages):
            if page.is_closed():
                continue
            path = self.folder / 'capturas' / (self.page_id(page) + '_' + uuid.uuid4().hex[:8] + '.png')
            try:
                await page.screenshot(path=str(path), full_page=False, timeout=10000,
                                      mask=[page.locator('input[type="password"]')])
                self.event('screenshot', page=self.page_id(page), file=str(path.relative_to(self.folder)))
            except Exception as exc:
                self.event('screenshot_failed', page=self.page_id(page), error=clean_error(exc), error_type=type(exc).__name__)
        self.inform('Capturas solicitadas processadas.')

    async def auto_login(self, page):
        c = self.config
        username, password = os.getenv('PORTAL_USUARIO'), os.getenv('PORTAL_SENHA')
        if not username or not password:
            self.inform('Login automático não executado: configure PORTAL_USUARIO e PORTAL_SENHA. Faça login manual.')
            return
        if not all(c[k] for k in ('user_selector', 'password_selector', 'submit_selector')):
            self.inform('Login automático não executado: faltam seletores. Faça login manual.')
            return
        try:
            await page.locator(c['user_selector']).fill(username)
            await page.locator(c['password_selector']).fill(password)
            await page.locator(c['submit_selector']).click()
            self.inform('Credenciais enviadas. Conclua MFA/CAPTCHA e confirme o acesso antes de iniciar a gravação.')
        except Exception as exc:
            self.inform('Login automático não concluído (' + type(exc).__name__ + '). Faça login manual.')

    async def run(self):
        from playwright.async_api import async_playwright
        try:
            async with async_playwright() as playwright:
                launch = {'headless': False, 'args': ['--start-maximized']}
                if self.config['browser'] != 'chromium':
                    launch['channel'] = self.config['browser']
                browser = await playwright.chromium.launch(**launch)
                self.context = await browser.new_context(accept_downloads=True, no_viewport=True)
                self.context.set_default_timeout(20000)
                await self.context.expose_binding('__recordPortal', self.binding)
                await self.context.add_init_script(JS)
                self.context.on('page', self.attach)
                page = await self.context.new_page()
                self.event('session_started', version=VERSION)
                self.inform('Saída: ' + str(self.folder))
                try:
                    await page.goto(self.config['url'], wait_until='domcontentloaded', timeout=60000)
                except Exception as exc:
                    self.inform('Navegação inicial: ' + type(exc).__name__ + ' | ' + clean_error(exc))
                if self.config['auto_login'] and not page.is_closed():
                    await self.auto_login(page)
                self.updates.put(('state', 'PAUSADO — faça login e clique em Iniciar/retomar'))
                while browser.is_connected() and self.context.pages:
                    try:
                        command, value = self.commands.get_nowait()
                    except queue.Empty:
                        await asyncio.sleep(0.05)
                        continue
                    if command == 'stop':
                        break
                    if command == 'start':
                        await self.set_active(True)
                    elif command == 'pause':
                        await self.set_active(False)
                    elif command == 'stage':
                        await self.flush()
                        self.stage = value
                        self.event('stage', note=value)
                        self.inform('Etapa: ' + value)
                    elif command == 'screenshot':
                        await self.screenshot()
                if browser.is_connected():
                    await self.set_active(False)
                    self.inform('Finalizando; aguardando downloads pendentes (até 5 minutos)...')
                    if self.tasks:
                        done, pending = await asyncio.wait(list(self.tasks), timeout=305)
                        for task in pending:
                            task.cancel()
                        if pending:
                            await asyncio.gather(*pending, return_exceptions=True)
                    await browser.close()
                self.event('session_finished')
        finally:
            if self.tasks:
                for task in list(self.tasks):
                    task.cancel()
                await asyncio.gather(*list(self.tasks), return_exceptions=True)
            self.stream.close()
            count = export_session(self.folder)
            self.inform(f'Relatórios salvos: {count} eventos em {self.folder}')


def gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    from tkinter.scrolledtext import ScrolledText

    root = tk.Tk()
    root.title('Gravador de Portais • Playwright ' + VERSION)
    root.geometry('1050x770')
    commands, updates = queue.Queue(), queue.Queue()
    worker = None
    closing = False
    fields = {}
    panel = ttk.Frame(root, padding=12)
    panel.pack(fill='both', expand=True)
    panel.columnconfigure(1, weight=1)
    def field(row, key, title, initial=''):
        ttk.Label(panel, text=title).grid(row=row, column=0, sticky='w', pady=4)
        variable = tk.StringVar(value=initial)
        ttk.Entry(panel, textvariable=variable).grid(row=row, column=1, sticky='ew', pady=4)
        fields[key] = variable
    field(0, 'url', 'URL do portal')
    field(1, 'output', 'Pasta de saída', str(Path.home() / 'Documents' / 'Gravacoes_Portais'))
    ttk.Button(panel, text='Escolher', command=lambda: choose()).grid(row=1, column=2, padx=5)
    def choose():
        path = filedialog.askdirectory()
        if path:
            fields['output'].set(path)
    browser = tk.StringVar(value='chromium')
    ttk.Label(panel, text='Navegador').grid(row=2, column=0, sticky='w')
    ttk.Combobox(panel, textvariable=browser, values=['chromium','msedge','chrome'], state='readonly').grid(row=2, column=1, sticky='w')
    automatic = tk.BooleanVar(value=False)
    ttk.Checkbutton(panel, text='Login automático opcional (variáveis PORTAL_USUARIO e PORTAL_SENHA)', variable=automatic).grid(row=3, column=0, columnspan=3, sticky='w', pady=6)
    field(4, 'user_selector', 'CSS do usuário')
    field(5, 'password_selector', 'CSS da senha')
    field(6, 'submit_selector', 'CSS do botão Entrar')
    capture_pdf = tk.BooleanVar(value=True)
    ttk.Checkbutton(panel, text=f'Capturar PDFs recebidos pela página e tentar recuperar PDFs blob (até {PDF_LIMIT // (1024*1024)} MiB)', variable=capture_pdf).grid(row=7, column=0, columnspan=3, sticky='w', pady=6)
    buttons = ttk.Frame(panel)
    buttons.grid(row=8, column=0, columnspan=3, sticky='w', pady=8)
    controls = []
    def send(command, value=None):
        commands.put((command, value))
    def start_worker(config):
        try:
            recorder = Recorder(config, commands, updates)
            asyncio.run(recorder.run())
        except Exception as exc:
            updates.put(('log', 'Falha: ' + type(exc).__name__ + ' | ' + clean_error(exc)))
            # Mensagens Playwright podem incluir valores de campos: não exportar traceback.
        finally:
            updates.put(('done', None))
    def open_browser():
        nonlocal worker
        if worker and worker.is_alive():
            return
        config = {key: value.get().strip() for key, value in fields.items()}
        if urlsplit(config['url']).scheme not in ('http','https') or not urlsplit(config['url']).netloc:
            messagebox.showerror('URL', 'Informe uma URL completa iniciando com https:// ou http://.')
            return
        if not config['output']:
            messagebox.showerror('Pasta', 'Escolha a pasta de saída.')
            return
        config.update(browser=browser.get(), auto_login=automatic.get(), capture_pdf=capture_pdf.get())
        while not commands.empty():
            commands.get_nowait()
        open_button.config(state='disabled')
        for button in controls:
            button.config(state='normal')
        state.set('ABRINDO — aguarde')
        worker = threading.Thread(target=start_worker, args=(config,), daemon=False)
        worker.start()
    open_button = ttk.Button(buttons, text='Abrir navegador', command=open_browser)
    open_button.pack(side='left', padx=3)
    for title, cmd in [('Iniciar/retomar','start'),('Pausar','pause'),('Capturar telas','screenshot'),('Finalizar e salvar','stop')]:
        button = ttk.Button(buttons, text=title, command=lambda c=cmd: send(c), state='disabled')
        button.pack(side='left', padx=3)
        controls.append(button)
    field(9, 'stage', 'Etapa / contrato / competência')
    # stage é metadado opcional, não parâmetro de execução.
    mark = ttk.Button(panel, text='Marcar etapa', command=lambda: send('stage', fields['stage'].get()), state='disabled')
    mark.grid(row=9, column=2, padx=4)
    controls.append(mark)
    state, count = tk.StringVar(value='AGUARDANDO'), tk.StringVar(value='0 eventos')
    ttk.Label(panel, textvariable=state).grid(row=10, column=0, columnspan=2, sticky='w', pady=8)
    ttk.Label(panel, textvariable=count).grid(row=10, column=2)
    log = ScrolledText(panel, height=16, wrap='word', state='disabled')
    log.grid(row=11, column=0, columnspan=3, sticky='nsew')
    panel.rowconfigure(11, weight=1)
    ttk.Label(panel, text='Finalize após os documentos aparecerem. Janelas nativas não são gravadas; capturas PDF podem exigir alternativa.').grid(row=12, column=0, columnspan=3, sticky='w', pady=8)
    def poll():
        nonlocal worker
        for _ in range(500):
            try:
                kind, value = updates.get_nowait()
            except queue.Empty:
                break
            if kind == 'log':
                log.config(state='normal')
                log.insert('end', datetime.now().strftime('%H:%M:%S') + ' | ' + value + '\n')
                log.see('end')
                log.config(state='disabled')
            elif kind == 'state':
                state.set(value)
            elif kind == 'count':
                count.set(str(value) + ' eventos')
            elif kind == 'done':
                state.set('ENCERRADO — consulte o resultado no painel')
                open_button.config(state='normal')
                for button in controls:
                    button.config(state='disabled')
                if closing:
                    root.destroy()
                    return
        root.after(100, poll)
    def close():
        nonlocal closing
        if worker and worker.is_alive():
            closing = True
            state.set('FINALIZANDO — aguarde os arquivos serem salvos')
            send('stop')
        else:
            root.destroy()
    root.protocol('WM_DELETE_WINDOW', close)
    poll()
    root.mainloop()


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('--recuperar', type=Path, help='Regenera relatórios a partir de eventos.jsonl de uma sessão encerrada')
    args = parser.parse_args()
    if args.recuperar:
        if not (args.recuperar / 'eventos.jsonl').is_file():
            print(f'Arquivo eventos.jsonl não encontrado em: {args.recuperar}')
            sys.exit(1)
        print(f'{export_session(args.recuperar)} eventos recuperados.')
        return
    try:
        import playwright.async_api  # noqa: F401
    except ImportError:
        text = 'Instale: py -m pip install playwright\nDepois: py -m playwright install chromium'
        print(text)
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror('Dependência ausente', text)
            root.destroy()
        except Exception:
            pass
        return
    gui()


if __name__ == '__main__':
    main()
