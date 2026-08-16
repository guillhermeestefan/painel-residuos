#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Notificador mensal do Painel de Residuos (fluxo em 2 etapas).

Roda toda 1a segunda-feira do mes (e a cada segunda para escalar).
So considera obras NAO concluidas. Tres gatilhos por obra:
  - Sinal amarelo de geracao de residuos
  - Sinal vermelho de geracao de residuos
  - Atraso na utilizacao de algum programa (aplicavel na etapa, mas nao em uso)

Etapa 1 (na 1a segunda do mes): e-mail para Engenheiro Responsavel,
  Administrativo da obra, Analista de Qualidade e Gestor de Meio Ambiente.
Etapa 2 (apos o engenheiro registrar justificativa + plano de acao no
  painel): e-mail para Gerente de Qualidade e Coordenador de Qualidade,
  ja com a justificativa e o plano.

Fontes ao vivo (Firestore, leitura publica):
  painel/caracteristicas  -> etapas, area, estrutura, e-mails por obra, concluida
  painel/programas        -> status/ativo dos programas
  painel/config           -> 3 e-mails fixos (gestorMA, gerenteQ, coordQ)
  painel/justificativas   -> justificativa + plano por obra e ciclo (YYYY-MM)
Estado anti-duplicidade: estado_notificacoes.json (commitado no repo).

Config por variaveis de ambiente (GitHub secrets):
  MAIL_SERVER, MAIL_PORT, MAIL_USERNAME, MAIL_PASSWORD, MAIL_FROM
  MAIL_TO (opcional, recebe copia de tudo para monitoramento)
  FIREBASE_API_KEY (opcional; senao le do firebase-config.js)
"""

import os, re, json, sys, ssl, smtplib, datetime, calendar, unicodedata
import urllib.request, urllib.error
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ID = "painel-residuos"
PAINEL_URL = "https://guillhermeestefan.github.io/painel-residuos/"
ESTADO_PATH = os.path.join(BASE_DIR, "estado_notificacoes.json")

HOJE = datetime.date.today()
CICLO = HOJE.strftime("%Y-%m")

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


def valido(email):
    return bool(email) and '@' in email and '.' in email.split('@')[-1]


def primeira_segunda(ano, mes):
    c = calendar.Calendar(firstweekday=0)
    for d in c.itermonthdates(ano, mes):
        if d.month == mes and d.weekday() == 0:
            return d
    return datetime.date(ano, mes, 1)


# ---------- defaults / dados ----------
def _read_js_assignment(path, var):
    try:
        txt = open(path, encoding='utf-8').read()
    except OSError:
        return None
    m = re.search(r'window\.' + re.escape(var) + r'\s*=\s*(\{.*?\}|\[.*?\])\s*;', txt, re.S)
    if not m:
        return None
    try:
        return json.loads(m.group(1))
    except json.JSONDecodeError:
        return None


def load_defaults():
    prog = _read_js_assignment(os.path.join(BASE_DIR, 'programas-init.js'), 'PROGRAMAS_INIT') or {}
    lista = _read_js_assignment(os.path.join(BASE_DIR, 'programas-init.js'), 'PROGRAMAS_LISTA') or list(PROGRAMA_ETAPAS.keys())
    return prog, lista


def load_raw():
    try:
        data = json.load(open(os.path.join(BASE_DIR, 'base_consolidada.json'), encoding='utf-8'))
    except OSError:
        return []
    for r in data:
        r['quantidade'] = float(r.get('quantidade') or 0)
        r['etapaOrig'] = r.get('etapa') or ''
    return data


def load_estado():
    try:
        return json.load(open(ESTADO_PATH, encoding='utf-8'))
    except (OSError, json.JSONDecodeError):
        return {}


def save_estado(estado):
    with open(ESTADO_PATH, 'w', encoding='utf-8') as f:
        json.dump(estado, f, ensure_ascii=False, indent=1)


# ---------- Firestore REST ----------
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
    except Exception as e:  # noqa: BLE001
        print(f"[aviso] Firestore '{doc}' indisponivel: {e}", file=sys.stderr)
        return None
    fields = payload.get('fields', {})
    if 'dados' not in fields:
        return None
    return _decode(fields['dados'])


# ---------- regra ----------
def fase_na_data(attrs, obra, data_iso):
    fs = ((attrs.get(obra) or {}).get('fases')) or []
    ordf = sorted([f for f in fs if f.get('etapa') and f.get('data')], key=lambda f: f['data'])
    cur = None
    for f in ordf:
        if f['data'] <= data_iso and (not f.get('dataFim') or data_iso <= f['dataFim']):
            cur = f['etapa']
    return cur


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
    lim = LIM_AMARELO.get(fase_na_data(attrs, o, HOJE.isoformat()))
    if lim is not None and real >= lim * orc:
        return 'amarelo'
    return 'verde'


TIPOS = [('atraso', 'Atraso de programa'), ('custo', 'Sinal de custo (amarelo/vermelho)')]


def alertas_tipos(o, attrs, prog, lista, vol):
    et = fase_na_data(attrs, o, HOJE.isoformat())
    atraso = []
    if et:
        for p in lista:
            if prog_aplicavel(p, et) and not prog_em_uso((prog.get(o) or {}).get(p) or {}):
                atraso.append('Atraso na utilização do programa: ' + p)
    custo = []
    st = status_custo(attrs, o, vol.get(o, 0.0))
    if st == 'vermelho':
        custo.append('Sinal vermelho de geração de resíduos (realizado passou do orçado)')
    elif st == 'amarelo':
        custo.append('Sinal amarelo de geração de resíduos (atingiu o limite da fase atual)')
    return et, {'atraso': atraso, 'custo': custo}


# ---------- e-mail ----------
def _wrap(inner):
    return (f'<div style="font-family:Arial,Helvetica,sans-serif;color:#222;max-width:720px">{inner}'
            f'<p style="color:#999;font-size:12px;margin-top:22px">Mensagem automática do Painel de Resíduos. '
            f'Abra o painel: <a href="{PAINEL_URL}">{PAINEL_URL}</a></p></div>')


def html_etapa1(obra, etapa, tipos):
    secs = ''
    for k, nome in TIPOS:
        it = tipos.get(k) or []
        if not it:
            continue
        li = ''.join(f'<li>{i}</li>' for i in it)
        secs += (f'<p style="font-size:14px;margin:12px 0 2px"><b>{nome}:</b></p>'
                 f'<ul style="font-size:14px;line-height:1.6;color:#333">{li}</ul>')
    return _wrap(f"""
  <h2 style="color:#C0463B;margin:0 0 4px">{obra} — alertas do mês</h2>
  <p style="color:#666;margin:0 0 8px">Etapa atual: <b>{etapa or '—'}</b> · verificação de {HOJE.isoformat()}.</p>
  {secs}
  <p style="font-size:14px;background:#FFF6E6;border:1px solid #E0A423;border-radius:6px;padding:10px 12px">
  <b>Engenheiro Responsável:</b> faça login no painel e registre a <b>justificativa</b> e o <b>plano de ação</b>
  de cada tipo na seção “Justificativas e plano de ação”. A Gerência de Qualidade só é notificada após esse registro.</p>""")


def html_etapa2(obra, etapa, nome_tipo, itens, just, plano, data_ts):
    li = ''.join(f'<li>{i}</li>' for i in itens)
    quando = ''
    if data_ts:
        try:
            quando = datetime.datetime.fromtimestamp(int(data_ts) / 1000).strftime('%d/%m/%Y')
        except (ValueError, OSError):
            quando = ''
    return _wrap(f"""
  <h2 style="color:#1F6F43;margin:0 0 4px">{obra} — justificativa registrada ({nome_tipo})</h2>
  <p style="color:#666;margin:0 0 14px">Etapa atual: <b>{etapa or '—'}</b>{' · registrado em ' + quando if quando else ''}.</p>
  <p style="font-size:14px;margin:0 0 6px"><b>Alertas ({nome_tipo}):</b></p>
  <ul style="font-size:14px;line-height:1.6;color:#333">{li}</ul>
  <p style="font-size:14px;margin:12px 0 4px"><b>Justificativa do Engenheiro Responsável:</b></p>
  <p style="font-size:14px;background:#F7F9FA;border:1px solid #e3e6ea;border-radius:6px;padding:10px 12px;white-space:pre-wrap">{just}</p>
  <p style="font-size:14px;margin:12px 0 4px"><b>Plano de ação:</b></p>
  <p style="font-size:14px;background:#F7F9FA;border:1px solid #e3e6ea;border-radius:6px;padding:10px 12px;white-space:pre-wrap">{plano}</p>""")


def enviar(assunto, html, destinatarios):
    dest = sorted(set(d for d in destinatarios if valido(d)))
    if not dest:
        print(f"   (sem destinatários válidos — não enviado: {assunto})")
        return False
    server = os.environ.get('MAIL_SERVER', '')
    port = int(os.environ.get('MAIL_PORT', '587'))
    user = os.environ.get('MAIL_USERNAME', '')
    pwd = os.environ.get('MAIL_PASSWORD', '')
    mail_from = os.environ.get('MAIL_FROM', user)
    if not (server and user and pwd):
        print(f"   [PRÉVIA — SMTP não configurado] Para: {', '.join(dest)} | {assunto}")
        return False
    msg = MIMEMultipart('alternative')
    msg['Subject'] = assunto
    msg['From'] = formataddr(("Painel de Resíduos Conx", mail_from))
    msg['To'] = ', '.join(dest)
    msg.attach(MIMEText("Seu cliente de e-mail não suporta HTML.", 'plain', 'utf-8'))
    msg.attach(MIMEText(html, 'html', 'utf-8'))
    ctx = ssl.create_default_context()
    if port == 465:
        with smtplib.SMTP_SSL(server, port, context=ctx, timeout=30) as sv:
            sv.login(user, pwd); sv.sendmail(mail_from, dest, msg.as_string())
    else:
        with smtplib.SMTP(server, port, timeout=30) as sv:
            sv.ehlo(); sv.starttls(context=ctx); sv.login(user, pwd)
            sv.sendmail(mail_from, dest, msg.as_string())
    print(f"   [ENVIADO] {assunto} -> {', '.join(dest)}")
    return True


def main():
    raw = load_raw()
    obras, seen = [], set()
    for r in raw:
        o = (r.get('obra') or '').strip()
        if o and o not in seen:
            seen.add(o); obras.append(o)

    defaults_prog, lista = load_defaults()
    prog = firestore_get('programas') or defaults_prog
    attrs = firestore_get('caracteristicas') or {}
    config = firestore_get('config') or {}
    justs = firestore_get('justificativas') or {}
    estado = load_estado()
    estado.setdefault(CICLO, {})

    vol = volume_construtivo(raw, attrs)
    pmonday = primeira_segunda(HOJE.year, HOJE.month)
    monitor = [e.strip() for e in os.environ.get('MAIL_TO', '').split(',') if e.strip()]

    fixos_ma = config.get('gestorMA', '')
    fixos_ger = config.get('gerenteQ', '')
    fixos_coord = config.get('coordQ', '')

    print(f"Ciclo {CICLO} | hoje {HOJE} | 1a segunda {pmonday} | obras {len(obras)}")
    env1 = env2 = 0

    for o in obras:
        if (attrs.get(o) or {}).get('concluida'):
            continue
        etapa, tipos = alertas_tipos(o, attrs, prog, lista, vol)
        if not (tipos['atraso'] or tipos['custo']):
            continue
        st = estado[CICLO].setdefault(o, {})
        st.setdefault('stage2', {})
        emails = (attrs.get(o) or {}).get('emails') or {}

        # Etapa 1 (uma vez por obra, na 1a segunda) — todos os tipos juntos
        if HOJE >= pmonday and not st.get('stage1'):
            dest = [emails.get('eng'), emails.get('adm'), emails.get('est'), fixos_ma] + monitor
            if enviar(f"[Painel Resíduos] {o} — alertas do mês ({CICLO})",
                      html_etapa1(o, etapa, tipos), dest):
                st['stage1'] = datetime.datetime.now().isoformat(timespec='seconds')
                env1 += 1

        # Etapa 2 por tipo (apos justificativa + plano daquele tipo)
        jcyc = (justs.get(o) or {}).get(CICLO) or {}
        for k, nome in TIPOS:
            it = tipos.get(k) or []
            if not it:
                continue
            jt = jcyc.get(k) or {}
            if st.get('stage1') and jt.get('justificativa') and jt.get('plano') and not st['stage2'].get(k):
                dest = [fixos_ger, fixos_coord] + monitor
                if enviar(f"[Painel Resíduos] {o} — justificativa registrada: {nome} ({CICLO})",
                          html_etapa2(o, etapa, nome, it, jt['justificativa'], jt['plano'], jt.get('data')), dest):
                    st['stage2'][k] = datetime.datetime.now().isoformat(timespec='seconds')
                    env2 += 1

    save_estado(estado)
    print(f"Etapa 1 enviadas: {env1} | Etapa 2 enviadas: {env2}")


if __name__ == '__main__':
    main()
