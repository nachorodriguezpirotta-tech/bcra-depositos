#!/usr/bin/env python3
"""Genera el dashboard y lo manda por mail (cuerpo HTML + gráfico inline + .html adjunto).

Credenciales Gmail (cuenta asistente.revolv@gmail.com), por orden de preferencia:
  1) env  MAIL_OAUTH_REFRESH_TOKEN + MAIL_OAUTH_CLIENT_ID + MAIL_OAUTH_CLIENT_SECRET   (GitHub Actions)
  2) archivo token_mail.json local (de asistente-revolv)                               (local)

Destinatarios: env MAIL_TO (coma) ; default nacho.rodriguezpirotta@gmail.com
"""
import os, sys, base64, datetime as dt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication

from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build as gbuild

import generate
import dashboard
import render_png

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
LOCAL_TOKEN = "/Users/ignaciorodriguezpirotta/Documents/Claude/asistente-revolv/token_mail.json"
HERE = os.path.dirname(os.path.abspath(__file__))
HTML_OUT = os.path.join(HERE, "dashboard.html")
PNG_OUT = os.path.join(HERE, "dashboard.png")
LOGO_OUT = os.path.join(HERE, "win_logo.png")


def _data_uri(path):
    with open(path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode("ascii")


def get_creds():
    rt = os.environ.get("MAIL_OAUTH_REFRESH_TOKEN")
    cid = os.environ.get("MAIL_OAUTH_CLIENT_ID")
    csec = os.environ.get("MAIL_OAUTH_CLIENT_SECRET")
    if rt and cid and csec:
        creds = Credentials(token=None, refresh_token=rt,
                            token_uri="https://oauth2.googleapis.com/token",
                            client_id=cid, client_secret=csec, scopes=SCOPES)
        creds.refresh(Request())
        return creds
    if os.path.exists(LOCAL_TOKEN):
        import json
        creds = Credentials.from_authorized_user_info(json.load(open(LOCAL_TOKEN)), SCOPES)
        if not creds.valid and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        return creds
    raise RuntimeError("Sin credenciales de mail (ni env OAuth ni token_mail.json).")


def plain_body(last_date, last_val, table, mm):
    L = ["Depósitos en USD del Sector Privado — BCRA (millones de u$s).",
         f"Último dato: {last_date:%d/%m/%Y} = {last_val:,.0f}", ""]
    for label, d, v, delta, pct in table:
        if label == "Último":
            L.append(f"  {label:14} {d:%d/%m/%Y}  {v:>9,.0f}")
        else:
            L.append(f"  {label:14} {d:%d/%m/%Y}  {v:>9,.0f}   Δ {delta:>+8,.0f}  ({pct:+.1%})")
    if mm:
        L += ["", f"  Money Market USD   {mm[1]:>9,.0f}  (al {mm[0]})",
              f"  Liquidez USD total {last_val + mm[1]:>9,.0f}"]
    L += ["", "Ver dashboard (HTML) adjunto."]
    return "\n".join(L)


def main():
    if "--no-download" not in sys.argv or not os.path.exists(generate.SRC):
        generate.download()
    data = generate.load_series()
    last_date, last_val, table = generate.build_table(data)
    mm = generate.fetch_mm_usd()
    png = generate.render_chart(data)
    # logo Win (SVG -> PNG) para usar como imagen
    try:
        render_png.render_svg(dashboard.LOGO_SVG, LOGO_OUT)
        logo_cid, logo_b64 = "cid:logo", _data_uri(LOGO_OUT)
    except Exception as e:
        print(f"  [aviso] no se pudo renderizar el logo: {e}")
        logo_cid = logo_b64 = None

    # cuerpo del mail: gráfico + logo inline vía CID
    html_mail = dashboard.build_html(last_date, last_val, table, mm,
                                     chart_src="cid:chart", logo_src=logo_cid)
    # adjunto standalone: gráfico + logo embebidos en base64 (autocontenido)
    html_file = dashboard.build_html(last_date, last_val, table, mm,
                                     chart_src=dashboard.chart_data_uri(png), logo_src=logo_b64)
    with open(HTML_OUT, "w", encoding="utf-8") as f:
        f.write(html_file)
    # foto del dashboard (PNG) para adjuntar
    try:
        render_png.render(HTML_OUT, PNG_OUT)
        png_ok = True
    except Exception as e:
        print(f"  [aviso] no se pudo renderizar PNG: {e}")
        png_ok = False

    to = [x.strip() for x in os.environ.get("MAIL_TO", "nacho.rodriguezpirotta@gmail.com").split(",") if x.strip()]
    is_test = "--test" in sys.argv
    subject = f"{'[PRUEBA] ' if is_test else ''}Depósitos USD Sector Privado (BCRA) — {last_date:%d/%m/%Y}"

    root = MIMEMultipart("mixed")
    root["To"] = ", ".join(to)
    root["From"] = "Win Securities <asistente.revolv@gmail.com>"
    root["Subject"] = subject

    related = MIMEMultipart("related")
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(plain_body(last_date, last_val, table, mm), "plain", "utf-8"))
    alt.attach(MIMEText(html_mail, "html", "utf-8"))
    related.attach(alt)
    with open(png, "rb") as f:
        img = MIMEImage(f.read(), _subtype="png")
    img.add_header("Content-ID", "<chart>")
    img.add_header("Content-Disposition", "inline", filename="grafico.png")
    related.attach(img)
    if logo_cid and os.path.exists(LOGO_OUT):
        with open(LOGO_OUT, "rb") as f:
            limg = MIMEImage(f.read(), _subtype="png")
        limg.add_header("Content-ID", "<logo>")
        limg.add_header("Content-Disposition", "inline", filename="win_logo.png")
        related.attach(limg)
    root.attach(related)

    # adjunto: PNG (foto) si se pudo renderizar; si no, el HTML standalone
    if png_ok:
        with open(PNG_OUT, "rb") as f:
            att = MIMEImage(f.read(), _subtype="png")
        att.add_header("Content-Disposition", "attachment",
                       filename=f"depositos_usd_{last_date:%Y%m%d}.png")
    else:
        with open(HTML_OUT, "rb") as f:
            att = MIMEApplication(f.read(), _subtype="html")
        att.add_header("Content-Disposition", "attachment", filename="dashboard_depositos_usd.html")
    root.attach(att)

    service = gbuild("gmail", "v1", credentials=get_creds(), cache_discovery=False)
    raw = base64.urlsafe_b64encode(root.as_bytes()).decode("ascii")
    sent = service.users().messages().send(userId="me", body={"raw": raw}).execute()
    print(f"✅ Mail enviado a {to}  msg_id={sent['id']}")


if __name__ == "__main__":
    main()
