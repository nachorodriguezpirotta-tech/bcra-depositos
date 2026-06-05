#!/usr/bin/env python3
"""Genera el dashboard HTML (mismo contenido que el Excel, presentación linda).

build_html(...) devuelve un HTML con estilos inline (apto como cuerpo de mail y
como archivo standalone). El gráfico se pasa como `chart_src`:
  - "cid:chart"        -> para el cuerpo del mail (imagen adjunta inline)
  - "data:image/png;base64,..."  -> para el .html standalone autocontenido
"""
import base64

BG = "#0B1220"; CARD = "#151E32"; CARD2 = "#1C2740"
BLUE = "#5B9BD5"; GREEN = "#2ECC71"; RED = "#E74C3C"
TXT = "#E8EDF5"; MUT = "#8A97AD"; LINE = "#26324A"
FONT = "'Montserrat','Segoe UI',Arial,sans-serif"


def _money(v):
    return f"{v:,.0f}".replace(",", ".")


def _pct(p):
    s = f"{p*100:+.1f}".replace(".", ",")
    return f"{s}%"


def _delta(v):
    return f"{v:+,.0f}".replace(",", "X").replace("-", "−").replace("X", ".")


def chart_data_uri(png_path):
    with open(png_path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode("ascii")


def build_html(last_date, last_val, table, mm, chart_src):
    milei = next((r for r in table if r[0] == "Inicio Milei"), None)
    milei_pct = milei[4] if milei else 0

    # filas del cuadro comparativo
    rows = []
    for label, d, v, delta, pct in table:
        if label == "Último":
            rows.append(f"""
            <tr>
              <td style="padding:11px 14px;border-bottom:1px solid {LINE};color:{TXT};font-weight:700">{label}</td>
              <td style="padding:11px 14px;border-bottom:1px solid {LINE};color:{MUT};font-size:13px">{d:%d/%m/%Y}</td>
              <td style="padding:11px 14px;border-bottom:1px solid {LINE};color:{TXT};font-weight:700;text-align:right">{_money(v)}</td>
              <td style="padding:11px 14px;border-bottom:1px solid {LINE};color:{MUT};text-align:right">—</td>
              <td style="padding:11px 14px;border-bottom:1px solid {LINE};color:{MUT};text-align:right">—</td>
            </tr>""")
        else:
            col = GREEN if delta >= 0 else RED
            rows.append(f"""
            <tr>
              <td style="padding:11px 14px;border-bottom:1px solid {LINE};color:{TXT}">{label}</td>
              <td style="padding:11px 14px;border-bottom:1px solid {LINE};color:{MUT};font-size:13px">{d:%d/%m/%Y}</td>
              <td style="padding:11px 14px;border-bottom:1px solid {LINE};color:{TXT};text-align:right">{_money(v)}</td>
              <td style="padding:11px 14px;border-bottom:1px solid {LINE};color:{col};text-align:right">{_delta(delta)}</td>
              <td style="padding:11px 14px;border-bottom:1px solid {LINE};color:{col};text-align:right;font-weight:700">{_pct(pct)}</td>
            </tr>""")
    rows_html = "".join(rows)

    # bloque money market
    if mm:
        mm_fecha, mm_total = mm
        total_liq = last_val + mm_total
        mm_html = f"""
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:8px">
          <tr>
            <td width="33%" style="padding:6px">
              <div style="background:{CARD2};border:1px solid {LINE};border-radius:12px;padding:16px">
                <div style="color:{MUT};font-size:12px;text-transform:uppercase;letter-spacing:.5px">Depósitos USD</div>
                <div style="color:{TXT};font-size:22px;font-weight:700;margin-top:4px">{_money(last_val)}</div>
              </div>
            </td>
            <td width="33%" style="padding:6px">
              <div style="background:{CARD2};border:1px solid {LINE};border-radius:12px;padding:16px">
                <div style="color:{MUT};font-size:12px;text-transform:uppercase;letter-spacing:.5px">Money Market USD</div>
                <div style="color:{TXT};font-size:22px;font-weight:700;margin-top:4px">{_money(mm_total)}</div>
              </div>
            </td>
            <td width="33%" style="padding:6px">
              <div style="background:linear-gradient(135deg,#1B3A5C,#15233B);border:1px solid {BLUE};border-radius:12px;padding:16px">
                <div style="color:{BLUE};font-size:12px;text-transform:uppercase;letter-spacing:.5px">Liquidez USD total</div>
                <div style="color:#fff;font-size:22px;font-weight:700;margin-top:4px">{_money(total_liq)}</div>
              </div>
            </td>
          </tr>
        </table>
        <div style="color:{MUT};font-size:11px;margin-top:6px">Money Market USD al {mm_fecha} · Fuente: argentinadatos/CAFCI (FCI Mercado de Dinero en dólares).</div>
        """
    else:
        mm_html = f'<div style="color:{MUT};font-size:12px">Money Market USD: no disponible en esta corrida.</div>'

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:{BG};font-family:{FONT}">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:{BG};padding:24px 0">
<tr><td align="center">
<table role="presentation" width="680" cellpadding="0" cellspacing="0" style="max-width:680px;width:100%">

  <!-- Header -->
  <tr><td style="padding:0 8px 18px">
    <div style="color:{MUT};font-size:12px;letter-spacing:2px;text-transform:uppercase">BCRA · Sector Privado</div>
    <div style="color:{TXT};font-size:26px;font-weight:800;margin-top:4px">Depósitos en USD</div>
    <div style="color:{MUT};font-size:13px;margin-top:2px">Serie diaria · millones de u$s · último dato {last_date:%d/%m/%Y}</div>
  </td></tr>

  <!-- Hero -->
  <tr><td style="padding:8px">
    <div style="background:{CARD};border:1px solid {LINE};border-radius:16px;padding:22px 24px">
      <table role="presentation" width="100%"><tr>
        <td>
          <div style="color:{MUT};font-size:12px;text-transform:uppercase;letter-spacing:.5px">Nivel actual</div>
          <div style="color:#fff;font-size:42px;font-weight:800;line-height:1.1;margin-top:4px">{_money(last_val)}<span style="font-size:16px;color:{MUT};font-weight:600"> M u$s</span></div>
        </td>
        <td align="right" valign="top">
          <span style="display:inline-block;background:rgba(46,204,113,.15);color:{GREEN};font-weight:700;font-size:14px;padding:8px 14px;border-radius:999px">▲ {_pct(milei_pct)} desde Milei</span>
        </td>
      </tr></table>
    </div>
  </td></tr>

  <!-- Chart -->
  <tr><td style="padding:8px">
    <div style="background:{CARD};border:1px solid {LINE};border-radius:16px;padding:16px">
      <img src="{chart_src}" width="100%" style="display:block;border-radius:8px" alt="Evolución depósitos USD">
    </div>
  </td></tr>

  <!-- Tabla comparativa -->
  <tr><td style="padding:8px">
    <div style="background:{CARD};border:1px solid {LINE};border-radius:16px;padding:6px 8px 10px">
      <div style="color:{TXT};font-size:15px;font-weight:700;padding:12px 14px 8px">Comparativo</div>
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">
        <tr>
          <th style="text-align:left;padding:8px 14px;color:{MUT};font-size:11px;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid {LINE}">Período</th>
          <th style="text-align:left;padding:8px 14px;color:{MUT};font-size:11px;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid {LINE}">Fecha</th>
          <th style="text-align:right;padding:8px 14px;color:{MUT};font-size:11px;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid {LINE}">Nivel</th>
          <th style="text-align:right;padding:8px 14px;color:{MUT};font-size:11px;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid {LINE}">Δ u$s</th>
          <th style="text-align:right;padding:8px 14px;color:{MUT};font-size:11px;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid {LINE}">Δ %</th>
        </tr>
        {rows_html}
      </table>
    </div>
  </td></tr>

  <!-- Money Market -->
  <tr><td style="padding:8px">
    <div style="background:{CARD};border:1px solid {LINE};border-radius:16px;padding:16px">
      <div style="color:{TXT};font-size:15px;font-weight:700;margin-bottom:4px">Liquidez en dólares</div>
      {mm_html}
    </div>
  </td></tr>

  <!-- Footer -->
  <tr><td style="padding:14px 16px 4px">
    <div style="color:{MUT};font-size:11px;line-height:1.6">
      Fuente: BCRA — Principales pasivos de las entidades financieras (diar_dep), hoja Sector_privado, columna CB.<br>
      Reporte automático semanal · generado {last_date:%d/%m/%Y}.
    </div>
  </td></tr>

</table>
</td></tr>
</table>
</body></html>"""
