#!/usr/bin/env python3
"""Envía por mail el Excel de depósitos USD (BCRA), con un resumen en el cuerpo.

Variables de entorno (secrets en GHA):
  MAIL_USER       -> casilla que envía  (ej: asistente.revolv@gmail.com)
  MAIL_PASSWORD   -> app password de esa casilla (16 caracteres, sin espacios)
  MAIL_TO         -> destinatarios separados por coma
  MAIL_CC         -> (opcional)

Uso:
  python3 send_mail.py            # genera el Excel (bajando data) y lo envía
  python3 send_mail.py --no-download
"""
import os, sys, smtplib, ssl, datetime as dt
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

import generate  # reutiliza la lógica de descarga + build


def build_body(last_date, table, mm=None):
    L = [f"Depósitos en USD del Sector Privado — BCRA (serie diaria, millones de u$s).",
         f"Último dato disponible: {last_date:%d/%m/%Y}.", ""]
    L.append("Nivel y variación vs. último dato:")
    for label, d, v, delta, pct in table:
        if label == "Último":
            L.append(f"  {label:14} {d:%d/%m/%Y}   {v:>9,.0f}")
        else:
            L.append(f"  {label:14} {d:%d/%m/%Y}   {v:>9,.0f}   Δ {delta:>+8,.0f}  ({pct:+.1%})")
    if mm:
        L += ["", "Money Market USD (FCI, patrimonio en u$s):",
              f"  Depósitos USD          {table[0][2]:>9,.0f}",
              f"  Money Market USD       {mm[1]:>9,.0f}   (al {mm[0]})",
              f"  Liquidez USD total     {table[0][2] + mm[1]:>9,.0f}"]
    L += ["", "Detalle completo + gráfico en el Excel adjunto.",
          "Fuente: BCRA, Principales pasivos de las entidades financieras (diar_dep), hoja Sector_privado, columna CB.",
          "", "(Automático — generado semanalmente.)"]
    return "\n".join(L)


def send(subject, body, attachment):
    user = os.environ["MAIL_USER"]
    pwd = os.environ["MAIL_PASSWORD"]
    to = [x.strip() for x in os.environ["MAIL_TO"].split(",") if x.strip()]
    cc = [x.strip() for x in os.environ.get("MAIL_CC", "").split(",") if x.strip()]

    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = user
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg.attach(MIMEText(body, "plain", "utf-8"))
    with open(attachment, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="vnd.ms-excel")
    part.add_header("Content-Disposition", "attachment",
                    filename=os.path.basename(attachment))
    msg.attach(part)

    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ctx) as s:
        s.login(user, pwd)
        s.sendmail(user, to + cc, msg.as_string())
    print(f"[mail enviado desde {user} a {to + cc}]")


def main():
    if "--no-download" not in sys.argv or not os.path.exists(generate.SRC):
        generate.download()
    data = generate.load_series()
    last_date, last_val, table = generate.build_table(data)
    mm = generate.fetch_mm_usd()
    generate.write_xlsx(data, last_date, last_val, table, mm)
    subject = f"Depósitos USD Sector Privado (BCRA) — {last_date:%d/%m/%Y}"
    send(subject, build_body(last_date, table, mm), generate.OUT)


if __name__ == "__main__":
    main()
