#!/usr/bin/env python3
"""Envío de prueba YA, vía Gmail API (token_mail.json de asistente.revolv@gmail.com).
Reutiliza la auth de asistente-revolv. No necesita app password."""
import os, sys, base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

REVOLV = "/Users/ignaciorodriguezpirotta/Documents/Claude/asistente-revolv"
sys.path.insert(0, REVOLV)
os.chdir(REVOLV)  # auth_mail/config resuelven paths relativos al repo

from auth_mail import get_mail_credentials
from googleapiclient.discovery import build

sys.path.insert(0, "/Users/ignaciorodriguezpirotta/Documents/Claude/bcra-depositos")
import generate

TO = "nacho.rodriguezpirotta@gmail.com"


def main():
    data = generate.load_series()
    last_date, last_val, table = generate.build_table(data)
    mm = generate.fetch_mm_usd()
    generate.write_xlsx(data, last_date, last_val, table, mm)

    body = ["Depósitos en USD del Sector Privado — BCRA (millones de u$s).",
            f"Último dato: {last_date:%d/%m/%Y}.", "", "Nivel y variación vs. último:"]
    for label, d, v, delta, pct in table:
        if label == "Último":
            body.append(f"  {label:14} {d:%d/%m/%Y}   {v:>9,.0f}")
        else:
            body.append(f"  {label:14} {d:%d/%m/%Y}   {v:>9,.0f}   Δ {delta:>+8,.0f}  ({pct:+.1%})")
    if mm:
        body += ["", "Money Market USD (FCI, patrimonio en u$s):",
                 f"  Depósitos USD          {last_val:>9,.0f}",
                 f"  Money Market USD       {mm[1]:>9,.0f}   (al {mm[0]})",
                 f"  Liquidez USD total     {last_val + mm[1]:>9,.0f}"]
    body += ["", "Gráfico (desde Milei) + cuadro completo en el Excel adjunto.",
             "Fuente: BCRA diar_dep, hoja Sector_privado, columna CB · MM: argentinadatos/CAFCI.",
             "", "** PRUEBA — se enviará automático todos los lunes 13:00. **"]

    msg = MIMEMultipart()
    msg["To"] = TO
    msg["From"] = "Win Securities <asistente.revolv@gmail.com>"
    msg["Subject"] = f"[PRUEBA] Depósitos USD Sector Privado (BCRA) — {last_date:%d/%m/%Y}"
    msg.attach(MIMEText("\n".join(body), "plain", "utf-8"))
    with open(generate.OUT, "rb") as f:
        part = MIMEApplication(f.read(), _subtype="vnd.ms-excel")
    part.add_header("Content-Disposition", "attachment", filename="depositos_usd_sector_privado.xlsx")
    msg.attach(part)

    creds = get_mail_credentials()
    service = build("gmail", "v1", credentials=creds, cache_discovery=False)
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("ascii")
    sent = service.users().messages().send(userId="me", body={"raw": raw}).execute()
    print(f"✅ Mail enviado a {TO}. message_id={sent['id']}")


if __name__ == "__main__":
    main()
