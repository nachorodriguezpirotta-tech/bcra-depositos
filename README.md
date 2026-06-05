# BCRA — Depósitos en USD del Sector Privado

Reporte automático semanal (dashboard HTML por mail) de los depósitos en dólares
del sector privado, con gráfico desde la asunción de Milei + cuadro comparativo +
patrimonio de Money Market en USD.

## Qué hace
1. Baja `diar_dep.xls` del BCRA (Principales pasivos de las entidades financieras).
2. Lee hoja `Sector_privado`, **columna CB** (depósitos USD del sector privado, millones u$s).
3. Suma el patrimonio de los FCI Money Market en dólares (argentinadatos/CAFCI).
4. Genera un dashboard HTML y lo manda por mail (gráfico inline + `.html` adjunto).

## Archivos
- `generate.py` — descarga, parseo, gráfico (matplotlib) y build del Excel.
- `dashboard.py` — template del dashboard HTML.
- `send_dashboard.py` — genera y envía el mail (Gmail API, OAuth). **Lo que corre el cron.**
- `send_mail.py` — variante por SMTP (alternativa, no usada por el cron).
- `.github/workflows/bcra-depositos.yml` — cron **lunes 13:00 ART** + `workflow_dispatch`.

## Correr local
```bash
pip install -r requirements.txt
python3 send_dashboard.py --no-download --test   # usa el .xls ya bajado, asunto [PRUEBA]
python3 generate.py                              # solo genera el Excel
```

## Secrets (GitHub Actions)
Manda desde `asistente.revolv@gmail.com` vía Gmail API (OAuth, scope gmail.send):
- `MAIL_OAUTH_REFRESH_TOKEN`, `MAIL_OAUTH_CLIENT_ID`, `MAIL_OAUTH_CLIENT_SECRET`
- `MAIL_TO` — destinatarios (coma)

Fuente: BCRA, https://www.bcra.gob.ar/depositos-y-otros-pasivos-de-las-entidades-financieras/
