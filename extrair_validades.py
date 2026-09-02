#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extrai a data de validade (Fim da Vigencia) do cadastro de cada obra no portal
cadastros-spregula.coletas.online e gera validades.json / validades.js.

- Reaproveita o CNPJ + senha por obra do cadastro (mesma senha do painel/CTRs),
  lidos por ler_cadastro() do extrair_ctrs.py (Firestore + CSV).
- Avisa por e-mail as obras que vencem em ate DIAS_ALERTA_VALIDADE dias (ou ja vencidas),
  usando os mesmos secrets MAIL_* do fluxo.
- Nunca apaga o que ja existe: obras que falharem mantem a validade anterior.

Fluxo no portal (mapeado a partir do passo a passo):
  /auth/MeuCadastro -> digita CNPJ -> "Consultar" -> "Acessar" -> senha -> "Confirmar"
  -> a pagina do cadastro mostra "Fim da Vigencia: DD/MM/AAAA".
"""

import argparse, json, os, re, sys
from datetime import date, datetime
from pathlib import Path

from extrair_ctrs import ler_cadastro, strip_acc, _enviar_alerta

BASE = Path(__file__).resolve().parent
VAL_JSON = BASE / "validades.json"
VAL_JS = BASE / "validades.js"
PORTAL = "https://cadastros-spregula.coletas.online/auth/MeuCadastro"
DIAS_ALERTA_VALIDADE = 60


def _load_validades():
    if VAL_JSON.exists():
        try:
            return json.loads(VAL_JSON.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _clic(page, textos, timeout=10000):
    for t in textos:
        try:
            el = page.locator(f"button:has-text('{t}')").first
            el.wait_for(state="visible", timeout=timeout)
            el.click()
            return True
        except Exception:
            continue
    return False


def _iso(br):
    d, m, a = br.split("/")
    return f"{int(a):04d}-{int(m):02d}-{int(d):02d}"


def extrair_validade_obra(page, cnpj, senha):
    """Retorna (validade_br, numero_cadastro) ou levanta excecao."""
    page.goto(PORTAL, wait_until="domcontentloaded", timeout=45000)
    inp = page.locator("input[type=text]").first
    inp.wait_for(state="visible", timeout=20000)
    inp.fill("".join(ch for ch in str(cnpj) if ch.isdigit()))
    if not _clic(page, ["Consultar"]):
        raise RuntimeError("botao Consultar nao encontrado")
    if not _clic(page, ["Acessar"]):
        raise RuntimeError("botao Acessar nao encontrado (CNPJ sem cadastro?)")
    pw = page.locator("input[type=password]").first
    pw.wait_for(state="visible", timeout=20000)
    pw.fill(str(senha))
    if not _clic(page, ["Confirmar"]):
        raise RuntimeError("botao Confirmar nao encontrado")
    # aguarda a pagina do cadastro carregar o texto da vigencia
    page.wait_for_function(
        "() => /Fim da Vig[eê\\u00ea]ncia/i.test(document.body.innerText)", timeout=30000
    )
    txt = page.inner_text("body")
    m = re.search(r"Fim da Vig[eê]ncia\s*:?\s*([0-3]?\d/[01]?\d/\d{4})", txt, re.I)
    if not m:
        raise RuntimeError("data de vigencia nao localizada na pagina")
    mc = re.search(r"(?:N[uú]mero Cadastro|Cadastro)\s*:?\s*(\d{3,})", txt)
    return m.group(1), (mc.group(1) if mc else "")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--obras", default="", help="rodar apenas estas obras (nomes por virgula)")
    ap.add_argument("--headful", action="store_true")
    args = ap.parse_args()

    from playwright.sync_api import sync_playwright

    obras, ativas = ler_cadastro()
    if args.obras.strip():
        alvo = {strip_acc(x) for x in args.obras.split(",") if x.strip()}
        obras = [o for o in obras if strip_acc(o["obra"]) in alvo]
    print(f"[INFO] Validade: {len(obras)} obra(s) a consultar")

    estado = _load_validades()
    hoje = date.today()
    resultados = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not args.headful)
        import time
        for o in obras:
            nome, cnpj, senha = o["obra"], o.get("cnpj", ""), o.get("senha", "")
            ok = False
            for tent in range(1, 3):
                ctx = browser.new_context(); page = ctx.new_page()
                try:
                    print(f"  -> {nome} (validade) tent {tent}/2 ...", end=" ", flush=True)
                    val_br, numero = extrair_validade_obra(page, cnpj, senha)
                    reg = estado.get(nome, {})
                    reg.update({
                        "obra": nome, "cnpj": "".join(c for c in str(cnpj) if c.isdigit()),
                        "numero_cadastro": numero or reg.get("numero_cadastro", ""),
                        "validade": _iso(val_br), "validade_br": val_br,
                        "atualizado_em": hoje.strftime("%Y-%m-%d"), "erro": "",
                    })
                    estado[nome] = reg
                    resultados[nome] = reg
                    print(f"OK ({val_br})")
                    ok = True
                except Exception as e:
                    msg = str(e).splitlines()[0][:90]
                    print(f"falhou: {msg}")
                    if tent == 2:
                        reg = estado.get(nome, {})
                        reg["erro"] = msg  # preserva validade anterior, so registra o erro
                        estado[nome] = reg
                finally:
                    ctx.close()
                if ok:
                    break
                time.sleep(4)
            time.sleep(1)
        browser.close()

    # grava estado (json) e o arquivo do painel (js)
    VAL_JSON.write_text(json.dumps(estado, ensure_ascii=False, indent=1), encoding="utf-8")
    VAL_JS.write_text(
        "window.VALIDADES = " + json.dumps(estado, ensure_ascii=False) + ";\n"
        "window.VALIDADES_META = " + json.dumps({"gerado_em": hoje.strftime("%d/%m/%Y")}, ensure_ascii=False) + ";\n",
        encoding="utf-8")
    print(f"[OK] {len(resultados)} validade(s) atualizada(s) nesta rodada | {len(estado)} no total")

    # alerta de vencimento (<= DIAS_ALERTA_VALIDADE dias ou vencido)
    venc = []
    for nome, reg in estado.items():
        if nome not in ativas or not reg.get("validade"):
            continue
        try:
            dias = (date.fromisoformat(reg["validade"]) - hoje).days
        except Exception:
            continue
        if dias <= DIAS_ALERTA_VALIDADE:
            venc.append((nome, reg["validade_br"], dias))
    if venc:
        venc.sort(key=lambda x: x[2])
        linhas = "".join(
            f"<li><b>{n}</b> — vence em <b>{v}</b> ({'VENCIDO ha '+str(-d)+' dia(s)' if d < 0 else 'faltam '+str(d)+' dia(s)'})</li>"
            for n, v, d in venc)
        html = (f"<p>Os cadastros abaixo no SP Regula vencem em ate {DIAS_ALERTA_VALIDADE} dias "
                f"(ou ja venceram). Providencie a renovacao:</p><ul>{linhas}</ul>")
        _enviar_alerta(f"[Painel de Residuos] {len(venc)} cadastro(s) SP Regula a vencer (<= {DIAS_ALERTA_VALIDADE} dias)", html)
        print(f"[ALERTA] {len(venc)} cadastro(s) a vencer: " + ", ".join(n for n, _, _ in venc))
    else:
        print(f"[OK] Nenhum cadastro vencendo nos proximos {DIAS_ALERTA_VALIDADE} dias.")


if __name__ == "__main__":
    main()
