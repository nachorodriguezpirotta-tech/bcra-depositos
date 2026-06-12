#!/usr/bin/env python3
"""Genera el dashboard HTML (mismo contenido que el Excel, presentación linda).

build_html(...) devuelve un HTML con estilos inline (apto como cuerpo de mail y
como archivo standalone). El gráfico se pasa como `chart_src`:
  - "cid:chart"        -> para el cuerpo del mail (imagen adjunta inline)
  - "data:image/png;base64,..."  -> para el .html standalone autocontenido
"""
import base64

# Paleta oficial Win Securities (Manual de Marca ene-2022):
# AZUL #2E5C9D · GRIS #404041 · CELESTE #4489EB · OCRE #976828 · MOSTAZA #E69C3C
# Tipografía oficial: Scania Sans (fallback web a sans similares).
AZUL = "#2E5C9D"; CELESTE = "#4489EB"; MOSTAZA = "#E69C3C"
BG = "#0B1220"; CARD = "#151E32"; CARD2 = "#1C2740"
BLUE = CELESTE; GREEN = "#2ECC71"; RED = "#E74C3C"
TXT = "#E8EDF5"; MUT = "#8A97AD"; LINE = "#26324A"
FONT = "'Scania Sans','Scania Sans CY','Montserrat','Segoe UI',Arial,sans-serif"


# Logo Win Securities recreado en vectorial (W de doble línea azul + wordmark),
# sobre fondo dark del dashboard para que blende en el header.
LOGO_SVG = """<svg xmlns="http://www.w3.org/2000/svg" width="408" height="100" viewBox="0 0 408 100">
  <rect width="408" height="100" fill="#0B1220"/>
  <polyline points="18,20 50,84 82,34 114,84 146,20" fill="none" stroke="#2E5C9D" stroke-width="15" stroke-linejoin="miter"/>
  <polyline points="18,20 50,84 82,34 114,84 146,20" fill="none" stroke="#0B1220" stroke-width="5.5" stroke-linejoin="miter"/>
  <text x="176" y="56" font-family="'Scania Sans',Montserrat,Arial,sans-serif" font-size="52" font-weight="800" fill="#FFFFFF">WIN</text>
  <text x="178" y="85" font-family="'Scania Sans',Montserrat,Arial,sans-serif" font-size="21" font-weight="600" letter-spacing="6.5" fill="#8A97AD">SECURITIES</text>
</svg>"""


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


def build_html(last_date, last_val, table, mm, chart_src, logo_src=None):
    logo_img = (f'<img src="{logo_src}" height="44" style="display:block" alt="Win Securities">'
                if logo_src else "")
    milei = next((r for r in table if r[0] == "Inicio Milei"), None)
    milei_pct = milei[4] if milei else 0

    # filas del cuadro comparativo
    rows = []
    for label, d, v, delta, pct in table:
        if label == "Último":
            rows.append(f"""
            <tr>
              <td style="padding:14px 16px;border-bottom:1px solid {LINE};color:{TXT};font-weight:700;font-size:22px">{label}</td>
              <td style="padding:14px 16px;border-bottom:1px solid {LINE};color:{MUT};font-size:18px">{d:%d/%m/%Y}</td>
              <td style="padding:14px 16px;border-bottom:1px solid {LINE};color:{TXT};font-weight:700;text-align:right;font-size:22px">{_money(v)}</td>
              <td style="padding:14px 16px;border-bottom:1px solid {LINE};color:{MUT};text-align:right;font-size:20px">—</td>
              <td style="padding:14px 16px;border-bottom:1px solid {LINE};color:{MUT};text-align:right;font-size:20px">—</td>
            </tr>""")
        else:
            col = GREEN if delta >= 0 else RED
            rows.append(f"""
            <tr>
              <td style="padding:14px 16px;border-bottom:1px solid {LINE};color:{TXT};font-size:22px">{label}</td>
              <td style="padding:14px 16px;border-bottom:1px solid {LINE};color:{MUT};font-size:18px">{d:%d/%m/%Y}</td>
              <td style="padding:14px 16px;border-bottom:1px solid {LINE};color:{TXT};text-align:right;font-size:22px">{_money(v)}</td>
              <td style="padding:14px 16px;border-bottom:1px solid {LINE};color:{col};text-align:right;font-size:20px">{_delta(delta)}</td>
              <td style="padding:14px 16px;border-bottom:1px solid {LINE};color:{col};text-align:right;font-weight:700;font-size:22px">{_pct(pct)}</td>
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
              <div style="background:{CARD2};border:1px solid {LINE};border-radius:12px;padding:18px">
                <div style="color:{MUT};font-size:16px;text-transform:uppercase;letter-spacing:.5px">Depósitos USD</div>
                <div style="color:{TXT};font-size:34px;font-weight:700;margin-top:6px">{_money(last_val)}</div>
              </div>
            </td>
            <td width="33%" style="padding:6px">
              <div style="background:{CARD2};border:1px solid {LINE};border-radius:12px;padding:18px">
                <div style="color:{MUT};font-size:16px;text-transform:uppercase;letter-spacing:.5px">Money Market USD</div>
                <div style="color:{TXT};font-size:34px;font-weight:700;margin-top:6px">{_money(mm_total)}</div>
              </div>
            </td>
            <td width="33%" style="padding:6px">
              <div style="background:linear-gradient(135deg,{AZUL},#15233B);border:1px solid {CELESTE};border-radius:12px;padding:18px">
                <div style="color:#BBD4F5;font-size:16px;text-transform:uppercase;letter-spacing:.5px">Liquidez USD total</div>
                <div style="color:#fff;font-size:34px;font-weight:700;margin-top:6px">{_money(total_liq)}</div>
              </div>
            </td>
          </tr>
        </table>
        <div style="color:{MUT};font-size:15px;margin-top:8px">Money Market USD al {mm_fecha} · Fuente: CAFCI (FCI Mercado de Dinero en dólares).</div>
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
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>
      <td valign="middle">
        <div style="color:{CELESTE};font-size:17px;letter-spacing:2px;text-transform:uppercase;font-weight:600">BCRA · Sector Privado</div>
        <div style="color:{TXT};font-size:42px;font-weight:800;margin-top:5px">Depósitos en USD</div>
        <div style="color:{MUT};font-size:19px;margin-top:4px">Serie diaria · millones de u$s · último dato {last_date:%d/%m/%Y}</div>
      </td>
      <td valign="middle" align="right">{logo_img}</td>
    </tr></table>
  </td></tr>

  <!-- Hero -->
  <tr><td style="padding:8px">
    <div style="background:{CARD};border:1px solid {LINE};border-radius:16px;padding:22px 24px">
      <table role="presentation" width="100%"><tr>
        <td>
          <div style="color:{MUT};font-size:17px;text-transform:uppercase;letter-spacing:.5px">Nivel actual</div>
          <div style="color:#fff;font-size:70px;font-weight:800;line-height:1.1;margin-top:6px">{_money(last_val)}<span style="font-size:25px;color:{MUT};font-weight:600"> M u$s</span></div>
        </td>
        <td align="right" valign="top">
          <span style="display:inline-block;background:rgba(46,204,113,.15);color:{GREEN};font-weight:700;font-size:21px;padding:11px 20px;border-radius:999px">▲ {_pct(milei_pct)} desde Milei</span>
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
      <div style="color:{TXT};font-size:23px;font-weight:700;padding:14px 16px 10px">Comparativo</div>
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">
        <tr>
          <th style="text-align:left;padding:10px 16px;color:{MUT};font-size:16px;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid {LINE}">Período</th>
          <th style="text-align:left;padding:10px 16px;color:{MUT};font-size:16px;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid {LINE}">Fecha</th>
          <th style="text-align:right;padding:10px 16px;color:{MUT};font-size:16px;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid {LINE}">Nivel</th>
          <th style="text-align:right;padding:10px 16px;color:{MUT};font-size:16px;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid {LINE}">Δ u$s</th>
          <th style="text-align:right;padding:10px 16px;color:{MUT};font-size:16px;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid {LINE}">Δ %</th>
        </tr>
        {rows_html}
      </table>
    </div>
  </td></tr>

  <!-- Money Market -->
  <tr><td style="padding:8px">
    <div style="background:{CARD};border:1px solid {LINE};border-radius:16px;padding:16px">
      <div style="color:{TXT};font-size:23px;font-weight:700;margin-bottom:6px">Liquidez en dólares</div>
      {mm_html}
    </div>
  </td></tr>

  <!-- Footer -->
  <tr><td style="padding:16px 16px 4px">
    <div style="color:{MUT};font-size:15px;line-height:1.6">
      Fuente: BCRA y CAFCI.<br>
      Reporte automático semanal · generado {last_date:%d/%m/%Y}.
    </div>
  </td></tr>

</table>
</td></tr>
</table>
</body></html>"""
