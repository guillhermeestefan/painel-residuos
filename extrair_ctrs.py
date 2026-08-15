#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extração de CTRs do SP Regula -> dados_ctr.js (para o painel no GitHub Pages).

Roda na nuvem (GitHub Actions), sem depender do seu computador:
  - obras: lidas de obras.csv (Obra, CNPJ, Etapa, Ativo) -- SEM senha (repo público)
  - senha: vem da variável de ambiente SENHA_SPREGULA (GitHub Secret)
  - 1ª execução (sem base ainda): puxa histórico completo desde 01/01/2020
  - execuções seguintes: só o período recente (--dias) e acumula
"""

import argparse, csv, json, os, sys, unicodedata
from datetime import date, timedelta
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent
OBRAS_CSV = BASE / "obras.csv"
NOVAS_CSV = BASE / "novas_obras.csv"
DOWNLOADS = BASE / "downloads_ctr"
CONSOLIDADO = BASE / "base_consolidada.xlsx"
BASE_JSON = BASE / "base_consolidada.json"
DADOS_JS = BASE / "dados_ctr.js"
SITE = "https://rcc-spregula.coletas.online/"
SENHA_ENV = os.environ.get("SENHA_SPREGULA", "conx2016")

CLASSE_MAP = {
    "concreto":"A","argamassa":"A","alvenaria":"A","ceramic":"A","cerâmic":"A",
    "solo":"A","terra":"A","escava":"A","entulho":"A","bloco":"A","telha":"A","tijolo":"A",
    "madeira":"B","metal":"B","ferro":"B","aço":"B","papel":"B","papelão":"B",
    "plastic":"B","plástic":"B","gesso":"B","sucata":"B",
    "misturad":"C","lã de vidro":"C","lã de rocha":"C","isopor":"C",
    "tinta":"D","solvente":"D","óleo":"D","oleo":"D","amianto":"D",
    "lâmpada":"D","lampada":"D","bateria":"D","pilha":"D","perigos":"D",
}

def strip_acc(s):
    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c)).lower()

def classe_de(residuo):
    r = strip_acc(residuo)
    for k, v in CLASSE_MAP.items():
        if strip_acc(k) in r: return v
    return "C"

def etapa_do_residuo(residuo, etapa_cadastro):
    r = strip_acc(residuo)
    if any(k in r for k in ("solo","terrapl","escava")): return "Terraplenagem"
    if "demoli" in r: return "Limpeza do Terreno - Demolição"
    if "stand" in r: return "Stand de Vendas - Demolição"
    return etapa_cadastro or "Não informada"

def normaliza_status(s):
    x = strip_acc(s)
    if "recusad" in x: return "Recusada no Destino Final"
    if "recebida" in x or "baixad" in x: return "Baixada"
    if "transito" in x or "retirada" in x: return "Em Trânsito"
    if "obra" in x: return "Em Obra"
    if "recebimento" in x: return "Pendente Recebimento"
    if "envio" in x: return "Pendente Envio"
    return s or "Baixada"

def _inativa(v):
    return str(v).strip().lower() in ("não","nao","n","false","0")

def _ler_csv(path):
    out = []
    if not path.exists():
        return out
    with open(path, encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            r = { (k or "").strip().lower(): (v or "").strip() for k, v in r.items() }
            obra = r.get("obra","")
            cnpj = "".join(ch for ch in r.get("login (cnpj)", r.get("cnpj","")) if ch.isdigit())
            senha = r.get("senha","") or SENHA_ENV
            etapa = r.get("etapa atual", r.get("etapa",""))
            ativo = r.get("ativo","Sim")
            if obra and cnpj:
                out.append({"obra":obra,"cnpj":cnpj,"senha":senha,"etapa":etapa,"ativo":ativo})
    return out

def ler_cadastro():
    linhas = _ler_csv(OBRAS_CSV) + _ler_csv(NOVAS_CSV)
    if not linhas:
        sys.exit("[ERRO] obras.csv não encontrado/vazio.")
    obras, vistos, nomes = [], set(), set()
    for o in linhas:
        nomes.add(o["obra"])
        if _inativa(o["ativo"]) or o["obra"] in vistos:
            continue
        if not o["senha"]:
            continue
        obras.append({"obra":o["obra"],"cnpj":o["cnpj"],"senha":o["senha"],"etapa":o["etapa"]})
        vistos.add(o["obra"])
    if not obras:
        sys.exit("[ERRO] Nenhuma obra ativa em obras.csv.")
    return obras, {o["obra"] for o in obras}

def clicar(page, seletores, timeout=8000, obrig=True):
    for sel in seletores:
        try:
            page.locator(sel).first.click(timeout=timeout); return True
        except Exception:
            continue
    if obrig:
        raise RuntimeError(f"Nenhum seletor clicável: {seletores[:2]}")
    return False

def extrair_obra(page, obra, desde, ate, downloads_dir):
    page.goto(SITE, wait_until="domcontentloaded", timeout=60000)
    page.fill("input[type='text']", obra["cnpj"])
    page.fill("input[type='password']", obra["senha"])
    if not clicar(page, ["input[type='submit'][value*='LOGIN' i]",
                         "input[type='button'][value*='LOGIN' i]",
                         "button:has-text('Login')"], obrig=False):
        page.locator("input[type='password']").press("Enter")
    page.wait_for_load_state("networkidle", timeout=60000)

    achou = False
    for nome in ("Gerador","Obra"):
        if clicar(page, [f"text={nome}", f"input[value='{nome}']",
                         f"input[value*='{nome}' i]",
                         f":is(a,button,div,span):has-text('{nome}')"], obrig=False):
            achou = True; break
    if not achou:
        raise RuntimeError("Módulo (Gerador/Obra) não encontrado (login inválido?)")
    page.wait_for_load_state("networkidle", timeout=60000)

    clicar(page, ["text=FECHAR VISUALIZAÇÃO DE AVISOS","input[value*='FECHAR' i]",
                  ":is(a,button):has-text('FECHAR')"], timeout=6000, obrig=False)
    page.wait_for_timeout(600)

    clicar(page, ["a:has-text('Relatórios')","text=Relatórios"], timeout=15000)
    page.wait_for_load_state("networkidle", timeout=60000)
    clicar(page, ["text=registradas",":is(a,div,span):has-text('registradas')"], timeout=15000)
    page.wait_for_timeout(1500)

    d1 = desde.strftime("%d/%m/%Y"); d2 = ate.strftime("%d/%m/%Y")
    page.evaluate(
        """([d1, d2]) => {
            const q = (suf) => document.querySelector("[id$='"+suf+"']");
            const set = (suf, v) => { const el = q(suf); if(el){ el.value = v;
                ['input','change','blur'].forEach(e=>el.dispatchEvent(new Event(e,{bubbles:true}))); } };
            set('ed_DataInicio', d1); set('ed_DataFim', d2);
            const tr = q('ddl_Transportador');
            if (tr) { tr.value = '0'; tr.dispatchEvent(new Event('change',{bubbles:true})); }
        }""", [d1, d2])
    page.wait_for_timeout(500)

    with page.expect_download(timeout=120000) as dl_info:
        if not clicar(page, ["input[id$='bt_Exportar']","input[value*='Exportar' i]"], obrig=False):
            page.locator("text=EXPORTAR").first.click()
    download = dl_info.value
    safe = strip_acc(obra["obra"]).replace(" ","_")[:40] or "obra"
    sufixo = Path(download.suggested_filename).suffix or ".xls"
    dest = downloads_dir / f"ctr_{safe}_{ate:%Y%m%d}{sufixo}"
    download.save_as(str(dest))
    return dest

def _rows_from_df(df):
    return [tuple(df.columns)] + [tuple(r) for r in df.itertuples(index=False, name=None)]

def _ler_html(raw):
    import io, pandas as pd
    try:
        txt = raw.decode("utf-8")
    except UnicodeDecodeError:
        txt = raw.decode("latin-1")
    dfs = pd.read_html(io.StringIO(txt), thousands='.', decimal=',')
    if not dfs:
        raise ValueError("nenhuma tabela HTML encontrada")
    df = max(dfs, key=lambda d: d.shape[0])
    return _rows_from_df(df)

def _ler_spreadsheetml(raw):
    # Excel 2003 XML (SpreadsheetML) - export comum do SP Regula
    from xml.etree import ElementTree as ET
    txt = raw.decode("utf-8", "ignore")
    root = ET.fromstring(txt)
    SS = "{urn:schemas-microsoft-com:office:spreadsheet}"
    rows = []
    for row in root.iter(SS + "Row"):
        cells, col = [], 0
        for cell in row.findall(SS + "Cell"):
            idx = cell.get(SS + "Index")
            if idx:
                col = int(idx) - 1
                while len(cells) < col:
                    cells.append(None)
            data = cell.find(SS + "Data")
            cells.append(data.text if data is not None else None)
            col += 1
        rows.append(tuple(cells))
    if not rows:
        raise ValueError("SpreadsheetML sem linhas")
    return rows

def _ler_xlsx(raw):
    import io
    wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True, read_only=True)
    return list(wb.active.iter_rows(values_only=True))

def _ler_xls_bin(raw):
    import io, pandas as pd
    df = pd.read_excel(io.BytesIO(raw), engine="xlrd")
    return _rows_from_df(df)

def _ler_csv_txt(raw):
    import csv as _csv, io as _io
    try:
        txt = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        txt = raw.decode("latin-1")
    amostra = txt[:5000]
    sep = max((";", ","), key=lambda s: amostra.count(s))
    if amostra.count("\t") > amostra.count(sep):
        sep = "\t"
    rows = [tuple(r) for r in _csv.reader(_io.StringIO(txt), delimiter=sep)]
    if len(rows) < 2:
        raise ValueError("CSV com menos de 2 linhas")
    return rows

def carregar_tabela(path):
    """Le o export do SP Regula em qualquer formato entregue: HTML disfarcado
    de .xls, Excel 2003 XML (SpreadsheetML), .xlsx real, .xls binario antigo
    (BIFF) ou CSV. Detecta por conteudo e, se falhar, faz cascata de parsers."""
    path = Path(path)
    raw = open(path, "rb").read()
    if not raw:
        raise ValueError("arquivo vazio (0 bytes) - export nao gerou dados")
    low = raw.lower()
    cabeca = low[:20000]

    ordem = []
    def add(fn):
        if fn not in ordem:
            ordem.append(fn)

    # 1) deteccao por conteudo (nao so pela extensao ou primeiros 4000 bytes)
    if b"mso-application" in cabeca or (b"spreadsheet" in cabeca and b"<workbook" in cabeca):
        add(_ler_spreadsheetml)
    if b"<table" in low or b"<html" in low:
        add(_ler_html)
    if raw[:2] == b"PK":
        add(_ler_xlsx)
    if raw[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        add(_ler_xls_bin)
    # 2) fallback: cobre qualquer coisa que a deteccao nao pegou
    for fn in (_ler_html, _ler_spreadsheetml, _ler_xlsx, _ler_xls_bin, _ler_csv_txt):
        add(fn)

    erros = []
    for fn in ordem:
        try:
            rows = fn(raw)
            if rows:
                return rows
        except Exception as e:
            erros.append(fn.__name__ + ": " + str(e).splitlines()[0][:60])
    raise ValueError("formato nao reconhecido (magic=" + raw[:48].hex() +
                     ") | tentativas -> " + " ; ".join(erros))

def ler_xls_ctr(path, obra):
    rows = carregar_tabela(path)
    if not rows: return []
    header = [strip_acc(c) for c in rows[0]]
    def col(*names):
        for n in names:
            for i,h in enumerate(header):
                if strip_acc(n)==h: return i
        for n in names:
            for i,h in enumerate(header):
                if strip_acc(n) in h: return i
        return None
    c_num=col("Numero"); c_data=col("DtGuia","data","emiss")
    c_transp=col("tbTransportador_Fantasia","transportador"); c_res=col("TipoResiduo","residuo","material")
    c_qtd=col("VolumeRecebido","Capacidade (M3)","volume","quantidade","m3")
    c_dest=col("Destino","destinac"); c_status=col("Status","situacao")
    c_ender=col("Endereco"); c_pgrcc=col("NumeroPGRCC")
    def cel(r,i): return "" if i is None or i>=len(r) or r[i] is None else str(r[i]).strip()
    regs=[]
    for r in rows[1:]:
        if r is None or all(c is None for c in r): continue
        residuo=cel(r,c_res) or "Não informado"
        try:
            qtd=float(str(r[c_qtd]).replace(",",".")) if c_qtd is not None and r[c_qtd] not in (None,"") else 0.0
        except (ValueError,TypeError): qtd=0.0
        raw=cel(r,c_data); data=raw
        if "/" in raw:
            p=raw.split()[0].split("/")
            if len(p)==3: data=f"{p[2]}-{p[1].zfill(2)}-{p[0].zfill(2)}"
        st=cel(r,c_status)
        regs.append({"obra":obra["obra"],"numero":cel(r,c_num),
            "etapa":etapa_do_residuo(residuo,obra.get("etapa","")),
            "residuo":residuo,"classe":classe_de(residuo),
            "transportador":cel(r,c_transp) or "Não informado",
            "quantidade":round(qtd,2),"destino":cel(r,c_dest) or "Não informado",
            "data":data,"status":normaliza_status(st),
            "divergencia":"divergencia" in strip_acc(st) and "sem" not in strip_acc(st),
            "endereco":cel(r,c_ender),"pgrcc":cel(r,c_pgrcc)})
    return regs

def salvar_consolidado(regs_novos, ativas=None):
    acervo={}
    if BASE_JSON.exists():
        try:
            for r in json.loads(BASE_JSON.read_text(encoding="utf-8")):
                acervo[(r.get("obra"),r.get("numero"))]=r
        except Exception: pass
    for r in regs_novos:
        acervo[(r.get("obra"),r.get("numero"))]=r
    regs=list(acervo.values())
    if ativas is not None:
        regs=[r for r in regs if r.get("obra") in ativas]
    regs=sorted(regs, key=lambda r:(r.get("data",""),r.get("obra","")))
    BASE_JSON.write_text(json.dumps(regs,ensure_ascii=False),encoding="utf-8")
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="CTRs"
    cols=["numero","data","obra","etapa","residuo","classe","transportador","quantidade","destino","status","divergencia","pgrcc"]
    ws.append([c.capitalize() for c in cols])
    for r in regs: ws.append([r.get(c,"") for c in cols])
    wb.save(CONSOLIDADO)
    meta={"gerado_em":date.today().strftime("%d/%m/%Y"),"total":len(regs),"novos_nesta_rodada":len(regs_novos)}
    DADOS_JS.write_text("window.DADOS_CTR = "+json.dumps(regs,ensure_ascii=False)+";\n"+
                        "window.DADOS_META = "+json.dumps(meta,ensure_ascii=False)+";\n",encoding="utf-8")
    return len(regs)

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--dias", type=int, default=30)
    ap.add_argument("--desde", help="dd/mm/aaaa")
    ap.add_argument("--headful", action="store_true")
    args=ap.parse_args()

    from playwright.sync_api import sync_playwright
    DOWNLOADS.mkdir(exist_ok=True)
    ate=date.today()
    # 1ª execução (sem base): histórico completo. Depois: janela recente.
    if args.desde:
        d,m,a=args.desde.split("/"); desde=date(int(a),int(m),int(d))
    elif not BASE_JSON.exists():
        desde=date(2020,1,1); print("[INFO] 1ª execução: puxando histórico completo desde 01/01/2020")
    else:
        desde=ate-timedelta(days=args.dias)

    obras, ativas = ler_cadastro()
    print(f"[INFO] {len(obras)} obra(s) ativa(s) | período {desde:%d/%m/%Y} a {ate:%d/%m/%Y}")

    todos=[]
    with sync_playwright() as p:
        browser=p.chromium.launch(headless=not args.headful)
        for o in obras:
            ctx=browser.new_context(accept_downloads=True); page=ctx.new_page()
            try:
                print(f"  -> {o['obra']} (CNPJ {o['cnpj']}) ...", end=" ", flush=True)
                arq=extrair_obra(page,o,desde,ate,DOWNLOADS)
                regs=ler_xls_ctr(arq,o); todos.extend(regs)
                print(f"OK ({len(regs)} CTRs)")
            except Exception as e:
                print(f"FALHOU: {str(e).splitlines()[0][:90]}")
            finally:
                ctx.close()
        browser.close()

    total=salvar_consolidado(todos, ativas)
    print(f"[OK] +{len(todos)} CTRs nesta rodada | base agora com {total} CTRs")

if __name__ == "__main__":
    main()
