#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Processa os pedidos de "Esqueci a senha" do painel.

- Le painel/reset (campo `pedidos`: lista de {email, ts}) no Firestore.
- Para cada e-mail que existe em painel/usuarios (e esta ativo), envia a senha
  para o proprio e-mail usando os secrets MAIL_* (mesmo servidor dos alertas).
- Limpa a fila (pedidos = []) ao final, para nao reenviar.
- Nunca revela se um e-mail existe ou nao: apenas nao envia para quem nao esta cadastrado.

Roda no GitHub Actions a cada ~10 min (workflow reset_senha.yml) ou manualmente.
"""

import os, re, ssl, json, smtplib, urllib.request
from email.mime.text import MIMEText
from email.utils import formataddr
from pathlib import Path

BASE = Path(__file__).resolve().parent
PROJECT_ID = "painel-residuos"


def _api_key():
    k = os.environ.get("FIREBASE_API_KEY", "")
    if k:
        return k
    try:
        t = open(BASE / "firebase-config.js", encoding="utf-8").read()
        m = re.search(r'apiKey:\s*"([^"]+)"', t)
        return m.group(1) if m else ""
    except OSError:
        return ""


def _dec(v):
    if not isinstance(v, dict):
        return v
    if "stringValue" in v:
        return v["stringValue"]
    if "integerValue" in v:
        return int(v["integerValue"])
    if "doubleValue" in v:
        return v["doubleValue"]
    if "booleanValue" in v:
        return v["booleanValue"]
    if "nullValue" in v:
        return None
    if "timestampValue" in v:
        return v["timestampValue"]
    if "mapValue" in v:
        return {k: _dec(x) for k, x in v["mapValue"].get("fields", {}).items()}
    if "arrayValue" in v:
        return [_dec(x) for x in v["arrayValue"].get("values", [])]
    return None


def _get_doc(doc):
    url = (f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}"
           f"/databases/(default)/documents/painel/{doc}")
    k = _api_key()
    if k:
        url += f"?key={k}"
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "conx-reset"})
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read().decode("utf-8")).get("fields", {})
    except Exception as e:  # noqa: BLE001
        print(f"[aviso] leitura de painel/{doc} falhou: {e}")
        return {}


def _clear_pedidos():
    url = (f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}"
           f"/databases/(default)/documents/painel/reset?updateMask.fieldPaths=pedidos")
    k = _api_key()
    if k:
        url += f"&key={k}"
    body = json.dumps({"fields": {"pedidos": {"arrayValue": {"values": []}}}}).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="PATCH",
                                headers={"Content-Type": "application/json", "User-Agent": "conx-reset"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            r.read()
        return True
    except Exception as e:  # noqa: BLE001
        print(f"[erro] nao consegui limpar a fila: {e}")
        return False


def _enviar(to, login, senha):
    host = os.environ.get("MAIL_SERVER")
    port = int(os.environ.get("MAIL_PORT") or 587)
    user = os.environ.get("MAIL_USERNAME")
    pw = os.environ.get("MAIL_PASSWORD")
    frm = os.environ.get("MAIL_FROM") or user
    if not (host and user and pw):
        print("[erro] MAIL_* nao configurado nos secrets")
        return False
    html = (
        "<p>Olá,</p>"
        "<p>Você (ou alguém) solicitou a senha de acesso ao <b>Painel de Gestão de Resíduos</b>.</p>"
        f"<p><b>Login:</b> {login}<br><b>Senha:</b> {senha}</p>"
        "<p>Por segurança, recomendamos trocar a senha após entrar, no botão "
        "<b>“Trocar minha senha”</b> no topo do painel.</p>"
        "<p style='color:#888;font-size:12px'>Se você não fez este pedido, pode ignorar este e-mail.</p>"
    )
    msg = MIMEText(html, "html", "utf-8")
    msg["Subject"] = "[Painel de Resíduos] Sua senha de acesso"
    msg["From"] = formataddr(("Painel de Resíduos CONX", frm))
    msg["To"] = to
    ctx = ssl.create_default_context()
    try:
        if port == 465:
            with smtplib.SMTP_SSL(host, port, context=ctx, timeout=30) as s:
                s.login(user, pw)
                s.sendmail(frm, [to], msg.as_string())
        else:
            with smtplib.SMTP(host, port, timeout=30) as s:
                s.starttls(context=ctx)
                s.login(user, pw)
                s.sendmail(frm, [to], msg.as_string())
        return True
    except Exception as e:  # noqa: BLE001
        print(f"[erro] envio para {to} falhou: {e}")
        return False


def main():
    resetf = _get_doc("reset")
    pedidos = _dec(resetf.get("pedidos")) if "pedidos" in resetf else []
    pedidos = pedidos or []
    if not pedidos:
        print("[ok] Nenhum pedido de senha pendente.")
        return

    usf = _get_doc("usuarios")
    usuarios = _dec(usf.get("dados")) if "dados" in usf else {}
    umap = {(k or "").strip().lower(): v for k, v in (usuarios or {}).items()}

    enviados, vistos = 0, set()
    for p in pedidos:
        em = ""
        if isinstance(p, dict):
            em = (p.get("email") or "").strip().lower()
        if not em or em in vistos:
            continue
        vistos.add(em)
        rec = umap.get(em)
        if not isinstance(rec, dict):
            print(f"[skip] e-mail nao cadastrado: {em}")
            continue
        if rec.get("ativo") is False:
            print(f"[skip] usuario inativo: {em}")
            continue
        senha = rec.get("senha")
        if not senha:
            print(f"[skip] usuario sem senha: {em}")
            continue
        if _enviar(em, em, senha):
            enviados += 1
            print(f"[ok] senha enviada para {em}")

    _clear_pedidos()
    print(f"[fim] {enviados} e-mail(s) enviado(s); fila de pedidos limpa.")


if __name__ == "__main__":
    main()
