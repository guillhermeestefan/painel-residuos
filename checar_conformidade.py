#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Verificador semanal do Painel de Residuos.

Envia um e-mail quando ha:
  A) Nao-conformidade programa x etapa
     - etapa atual da obra = fase vigente hoje (Firestore painel/caracteristicas)
     - "em uso" = programa ativo AND status preenchido (nao vazio / "-" / "Nao")
     - nao-conforme = programa se aplica a etapa atual mas nao esta em uso
  B) Semaforo de custo amarelo/vermelho
     - realizado = volume construtivo da obra (m3)
     - orcado = fator(0,20 alvenaria / 0,22 concreto) x area (m2)
     - vermelho = realizado > orcado
     - amarelo  = realizado >= limite da fase atual x orcado
                  (Estrutura 60%, Estrutura e Acabamentos 70%, Acabamentos 90%)

Replica exatamente a logica do index.html. Le dados ao vivo do Firestore
(leitura publica). Sem Firestore, cai nos defaults commitados.

Config por variaveis de ambiente (GitHub secrets):
  FIREBASE_API_KEY (opcional)
  MAIL_SERVER, MAIL_PORT, MAIL_USERNAME, MAIL_PASSWORD, MAIL_FROM, MAIL_TO
"""

import os, re, json, sys, ssl, smtplib, datetime, unicodedata
import urllib.request, urllib.error
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ID = "painel-residuos"
HOJE = datetime.date.today().isoformat()

# ---- constantes (identicas ao index.html) ----
PROGRAMA_ETAPAS = {
    'Caçambas Padrão (4m³)': 'todas',
    'Caçambas de Gesso (5m³)': ['Estrutura e Acabamentos', 'Acabamentos'],
    'Big Bags': 'todas',
    'Logística Reversa de Bloco': ['Estrutura', 'Estrutura e Acabamentos'],
    'Ensacados': ['Estrutura', 'Estrutura e Acabamentos', 'Acabamentos'],
    'Painel Eletrônico Tapume': 'todas',
    'Baldes': ['Estrutura e Acabamentos', 'Acabamentos'],
    'Pró-Lata': ['Estrutura e Acabamentos', 'Acabamentos'],
    'Madeira - Rafa Entulhos': ['Estrutura', 'Estrutura e Acabamentos'],
    'Sucata - Rafa Entulhos': ['Fundações', 'Estrutura', 'Estrutura e Acabamentos'],
    'Agregado Reciclado (Classe A)': ['Fundações', 'Estrutura'],
}
ETAPAS_CONSTR = ['Limpeza do Terreno', 'Fundações', 'Estrutura',
                 'Estrutura e Acabamentos', 'Acabamentos']
FATOR_ORC = {'Alvenaria estrutural': 0.2, 'Concreto armado': 0.22}
LIM_AMARELO = {'Estrutura': 0.6, 'Estrutura e Acabamentos': 0.7, 'Acabamentos': 0.9}
NAO_CONSTRUTIVAS = {'Terraplenagem', 'Limpeza do Terreno',
                    'Stand de Vendas - Construção', 'Stand de Vendas - Demolição'}


def strip(s):
    s = s or ''
    return ''.join(c for c in unicodedata.normalize('NFD', s)
                   if unicodedata.category(c) != 'Mn').lower()


# ---------- defaults commitados ----------
def _read_js_assignment(path, var):
    try:
        txt = open(path, encoding='utf-8').read()
    except OSError:
        return None
    m = re.search(r'window\.' + re.escape(var) + r'\s*=\s*(\{.*?\}|\[.*?\])\s*;',
                  txt, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(1))
    except json.JSONDecodeError:
        return None


def load_defaults():
    prog = _read_js_assignment(os.path.join(BASE_DIR, 'programas-init.js'),
                               'PROGRAMAS_INIT') or {}
    lista = _read_js_assignment(os.path.join(BASE_DIR, 'programas-init.js'),
                                'PROGRAMAS_LISTA') or list(PROGRAMA_ETAPAS.keys())
    return prog, lista


def load_raw():
    path = os.path.join(BASE_DIR, 'base_consolidada.json')
    try:
        data = json.load(open(path, encoding='utf-8'))
    except OSError:
        return []
    for r in data:
        r['quantidade'] = float(r.get('quantidade') or 0)
        r['etapaOrig'] = r.get('etapa') or ''
    return data


# ---------- Firestore REST (leitura publica) ----------
def _decode(v):
    if 'nullValue' in v: return None
    if 'booleanValue' in v: return v['booleanValue']
    if 'integerValue' in v: return int(v['integerValue'])
    if 'doubleValue' in v: return float(v['doubleValue'])
    if 'stringValue' in v: return v['stringValue']
    if 'timestampValue' in v: return v['timestampValue']
    if 'mapValue' in v:
        return {k: _decode(x) for k, x in v['mapValue'].get('fields', {}).items()}
    if 'arrayValue' in v:
        return [_decode(x) for x in v['arrayValue'].get('values', [])]
    return None


def _api_key_from_config():
    """Le apiKey do firebase-config.js (ja publico no repo) como fallback."""
    try:
        txt = open(os.path.join(BASE_DIR, 'firebase-config.js'), encoding='utf-8').read()
    except OSError:
        return ''
    m = re.search(r'apiKey:\s*"([^"]+)"', txt)
    return m.group(1) if m else ''


def firestore_get(doc):
    api_key = os.environ.get('FIREBASE_API_KEY', '') or _api_key_from_config()
    url = (f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}"
           f"/databases/(default)/documents/painel/{doc}")
    if api_key:
        url += f"?key={api_key}"
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'conx-checker'})
        with urllib.request.urlopen(req, timeout=30) as resp:
            payload = json.loads(resp.read().decode('utf-8'))
    except Exception as e:  # noqa: BLE001 (falha silenciosa proposital)
        print(f"[aviso] Firestore '{doc}' indisponivel: {e}", file=sys.stderr)
        return None
    fields = payload.get('fields', {})
    if 'dados' not in fields:
        return None
    return _decode(fields['dados'])


# ---------- regra: etapa vigente ----------
def fase_na_data(attrs, obra, data_iso):
    fs = ((attrs.get(obra) or {}).get('fases')) or []
    ordf = sorted([f for f in fs if f.get('etapa') and f.get('data')],
                  key=lambda f: f['data'])
    cur = None
    for f in ordf:
        if f['data'] <= data_iso and (not f.get('dataFim') or data_iso <= f['dataFim']):
            cur = f['etapa']
    return cur


# ---------- A) conformidade programa x etapa ----------
def prog_aplicavel(p, et):
    if not et:
        return False
    r = PROGRAMA_ETAPAS.get(p)
    if not r:
        return False
    return (et in ETAPAS_CONSTR) if r == 'todas' else (et in r)


def prog_em_uso(cell):
    if not cell or not cell.get('ativo'):
        return False
    s = str(cell.get('status') or '').strip().lower()
    return s not in ('', '-', 'não', 'nao')


def checar_conformidade(obras, prog, attrs, lista):
    por_obra, por_programa = {}, {}
    for o in obras:
        et = fase_na_data(attrs, o, HOJE)
        if not et:
            continue
        for p in lista:
            if not prog_aplicavel(p, et):
                continue
            cell = (prog.get(o) or {}).get(p) or {}
            if not prog_em_uso(cell):
                por_obra.setdefault(o, {'etapa': et, 'progs': []})['progs'].append(p)
                por_programa.setdefault(p, []).append(o)
    return por_obra, por_programa


# ---------- B) semaforo de custo ----------
def etapa_especial(residuo):
    r = strip(residuo)
    if re.search(r'solo|terrapl|escava|inerte', r): return 'Terraplenagem'
    if 'stand' in r: return 'Stand de Vendas - Demolição'
    if 'demoli' in r: return 'Limpeza do Terreno'
    return None


def record_etapa(r, attrs):
    return (etapa_especial(r.get('residuo', '')) or
            fase_na_data(attrs, r.get('obra', ''), r.get('data', '')) or
            r.get('etapaOrig') or 'Não informada')


def volume_construtivo(raw, attrs):
    vol = {}
    for r in raw:
        if record_etapa(r, attrs) in NAO_CONSTRUTIVAS:
            continue
        vol[r['obra']] = vol.get(r['obra'], 0.0) + r['quantidade']
    return vol


def status_custo(attrs, o, real):
    a = attrs.get(o) or {}
    area = float(a.get('area') or 0)
    fat = FATOR_ORC.get(a.get('estrutura'))
    if not area or not fat:
        return 'cinza'
    orc = fat * area
    if real > orc:
        return 'vermelho'
    lim = LIM_AMARELO.get(fase_na_data(attrs, o, HOJE))
    if lim is not None and real >= lim * orc:
        return 'amarelo'
    return 'verde'


def checar_custo(obras, attrs, vol):
    alertas = []  # (obra, status, real, orcado, pct, etapa)
    for o in obras:
        real = vol.get(o, 0.0)
        st = status_custo(attrs, o, real)
        if st not in ('amarelo', 'vermelho'):
            continue
        a = attrs.get(o) or {}
        orc = FATOR_ORC.get(a.get('estrutura'), 0) * float(a.get('area') or 0)
        pct = (real / orc * 100) if orc else 0
        alertas.append({'obra': o, 'status': st, 'real': real, 'orc': orc,
                        'pct': pct, 'etapa': fase_na_data(attrs, o, HOJE) or '—'})
    ordem = {'vermelho': 0, 'amarelo': 1}
    alertas.sort(key=lambda x: (ordem[x['status']], -x['pct']))
    return alertas


# ---------- e-mail ----------
def montar_html(por_obra, por_programa, custo):
    total_nc = sum(len(v['progs']) for v in por_obra.values())
    partes = [f"""<div style="font-family:Arial,Helvetica,sans-serif;color:#222;max-width:760px">
  <h2 style="color:#1F6F43;margin:0 0 4px">Painel de Resíduos — Alertas da semana</h2>
  <p style="color:#666;margin:0 0 18px">Verificação de {HOJE}.</p>"""]

    # Secao custo
    if custo:
        vermelhos = sum(1 for c in custo if c['status'] == 'vermelho')
        amarelos = sum(1 for c in custo if c['status'] == 'amarelo')
        linhas = []
        for c in custo:
            cor = '#C0463B' if c['status'] == 'vermelho' else '#E0A423'
            tag = 'ESTOUROU' if c['status'] == 'vermelho' else 'ATENÇÃO'
            linhas.append(
                f"<tr><td style='padding:8px;border-bottom:1px solid #eee;font-weight:600'>{c['obra']}</td>"
                f"<td style='padding:8px;border-bottom:1px solid #eee'><span style='background:{cor};color:#fff;padding:2px 8px;border-radius:4px;font-size:12px'>{tag}</span></td>"
                f"<td style='padding:8px;border-bottom:1px solid #eee;color:#555'>{c['etapa']}</td>"
                f"<td style='padding:8px;border-bottom:1px solid #eee;text-align:right'>{c['real']:.0f} / {c['orc']:.0f} m³</td>"
                f"<td style='padding:8px;border-bottom:1px solid #eee;text-align:right;font-weight:600;color:{cor}'>{c['pct']:.0f}%</td></tr>")
        partes.append(f"""
  <h3 style="margin:18px 0 6px;color:#C0463B">Semáforo de custo — {vermelhos} vermelho(s), {amarelos} amarelo(s)</h3>
  <table style="border-collapse:collapse;width:100%;font-size:14px">
    <thead><tr style="background:#FAFBFC;text-align:left">
      <th style="padding:8px;border-bottom:2px solid #ddd">Obra</th>
      <th style="padding:8px;border-bottom:2px solid #ddd">Status</th>
      <th style="padding:8px;border-bottom:2px solid #ddd">Etapa</th>
      <th style="padding:8px;border-bottom:2px solid #ddd;text-align:right">Realizado / Orçado</th>
      <th style="padding:8px;border-bottom:2px solid #ddd;text-align:right">% do orçado</th>
    </tr></thead><tbody>{''.join(linhas)}</tbody></table>""")

    # Secao conformidade
    if por_obra:
        linhas = []
        for o in sorted(por_obra):
            progs = ', '.join(por_obra[o]['progs'])
            linhas.append(
                f"<tr><td style='padding:8px;border-bottom:1px solid #eee;font-weight:600'>{o}</td>"
                f"<td style='padding:8px;border-bottom:1px solid #eee;color:#555'>{por_obra[o]['etapa']}</td>"
                f"<td style='padding:8px;border-bottom:1px solid #eee;color:#C0463B'>{progs}</td></tr>")
        resumo = ''.join(
            f"<li><b>{p}</b>: {len(obs)} obra(s) — {', '.join(sorted(obs))}</li>"
            for p, obs in sorted(por_programa.items(), key=lambda kv: -len(kv[1])))
        partes.append(f"""
  <h3 style="margin:22px 0 6px;color:#C0463B">Não-conformidade programa × etapa — {total_nc} item(ns) em {len(por_obra)} obra(s)</h3>
  <table style="border-collapse:collapse;width:100%;font-size:14px">
    <thead><tr style="background:#FAFBFC;text-align:left">
      <th style="padding:8px;border-bottom:2px solid #ddd">Obra</th>
      <th style="padding:8px;border-bottom:2px solid #ddd">Etapa atual</th>
      <th style="padding:8px;border-bottom:2px solid #ddd">Programas a cobrar</th>
    </tr></thead><tbody>{''.join(linhas)}</tbody></table>
  <p style="font-size:13px;color:#444;margin:8px 0 0"><b>Por programa:</b></p>
  <ul style="font-size:13px;line-height:1.6;color:#333">{resumo}</ul>""")

    partes.append("""
  <p style="color:#999;font-size:12px;margin-top:22px">Mensagem automática semanal.
  "Em uso" = programa ativo + status preenchido. Abra o painel para a matriz completa.</p>
</div>""")
    return ''.join(partes)


def enviar_email(html, assunto):
    server = os.environ.get('MAIL_SERVER', '')
    port = int(os.environ.get('MAIL_PORT', '587'))
    user = os.environ.get('MAIL_USERNAME', '')
    pwd = os.environ.get('MAIL_PASSWORD', '')
    mail_from = os.environ.get('MAIL_FROM', user)
    mail_to = os.environ.get('MAIL_TO', 'guilherme.estefan@conx.com.br')
    if not (server and user and pwd):
        print("[aviso] SMTP nao configurado. E-mail nao enviado. Previa:\n")
        print(html)
        return False
    msg = MIMEMultipart('alternative')
    msg['Subject'] = assunto
    msg['From'] = formataddr(("Painel de Resíduos Conx", mail_from))
    msg['To'] = mail_to
    msg.attach(MIMEText("Seu cliente de e-mail não suporta HTML.", 'plain', 'utf-8'))
    msg.attach(MIMEText(html, 'html', 'utf-8'))
    dest = [a.strip() for a in mail_to.split(',') if a.strip()]
    ctx = ssl.create_default_context()
    if port == 465:
        with smtplib.SMTP_SSL(server, port, context=ctx, timeout=30) as s:
            s.login(user, pwd); s.sendmail(mail_from, dest, msg.as_string())
    else:
        with smtplib.SMTP(server, port, timeout=30) as s:
            s.ehlo(); s.starttls(context=ctx); s.login(user, pwd)
            s.sendmail(mail_from, dest, msg.as_string())
    print(f"[ok] E-mail enviado para {mail_to}.")
    return True


def main():
    defaults_prog, lista = load_defaults()
    raw = load_raw()
    obras = []
    seen = set()
    for r in raw:
        o = (r.get('obra') or '').strip()
        if o and o not in seen:
            seen.add(o); obras.append(o)

    prog_fs = firestore_get('programas')
    attrs = firestore_get('caracteristicas') or {}
    prog = prog_fs if prog_fs else defaults_prog
    fonte = 'Firestore' if prog_fs else 'defaults(programas-init.js)'

    por_obra, por_programa = checar_conformidade(obras, prog, attrs, lista)
    vol = volume_construtivo(raw, attrs)
    custo = checar_custo(obras, attrs, vol)

    total_nc = sum(len(v['progs']) for v in por_obra.values())
    n_etapa = sum(1 for o in obras if fase_na_data(attrs, o, HOJE))
    print(f"Obras: {len(obras)} | com etapa: {n_etapa} | fonte programas: {fonte}")
    print(f"Nao-conformidades: {total_nc} em {len(por_obra)} obra(s)")
    print(f"Custo: {sum(1 for c in custo if c['status']=='vermelho')} vermelho, "
          f"{sum(1 for c in custo if c['status']=='amarelo')} amarelo")

    if total_nc == 0 and not custo:
        print("Nada a reportar. Nenhum e-mail enviado.")
        return
    partes = []
    if custo:
        partes.append(f"{sum(1 for c in custo if c['status']=='vermelho')}🔴 "
                      f"{sum(1 for c in custo if c['status']=='amarelo')}🟡 custo")
    if total_nc:
        partes.append(f"{total_nc} não-conformidade(s)")
    assunto = f"[Painel Resíduos] {' + '.join(partes)} — {HOJE}"
    html = montar_html(por_obra, por_programa, custo)
    enviar_email(html, assunto)


if __name__ == '__main__':
    main()
