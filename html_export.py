"""
Export du rapport de visualisation en HTML interactif autonome.

Genere un fichier HTML self-contained avec :
- Page de garde + KPIs
- Sommaire sticky cliquable (memes ancres que le rapport en ligne)
- Graphiques Plotly interactifs (zoom, hover, legende cliquable)
- Tableaux annuels
- Plotly charge via CDN

Le fichier est ouvrable hors-ligne (sauf Plotly CDN). Partageable par email.
"""

import base64
from pathlib import Path

import plotly.graph_objects as go
import plotly.io as pio

from calculs import projection_complete, calc_tableau_pret


def _fig_html(fig, height=400):
    """Convertit une figure Plotly en HTML standalone (sans Plotly JS lib)."""
    if fig is None:
        return ""
    fig.update_layout(height=height, margin=dict(l=50, r=30, t=50, b=40),
                      autosize=True, paper_bgcolor="white", plot_bgcolor="#fafafa")
    return pio.to_html(
        fig, full_html=False, include_plotlyjs=False,
        config={"displayModeBar": True, "displaylogo": False,
                "modeBarButtonsToRemove": ["lasso2d", "select2d"]},
    )


def _table_html(headers, rows, classes="data-table"):
    h = f'<table class="{classes}"><thead><tr>'
    h += "".join(f"<th>{c}</th>" for c in headers)
    h += "</tr></thead><tbody>"
    for r in rows:
        h += "<tr>" + "".join(f"<td>{c}</td>" for c in r) + "</tr>"
    h += "</tbody></table>"
    return h


def _img_data_uri(path):
    p = Path(path)
    if not p.exists():
        return None
    return "data:image/jpeg;base64," + base64.b64encode(p.read_bytes()).decode()


def _kpi_card(value, label, color):
    return (
        f'<div style="background:linear-gradient(135deg,{color}15,{color}30); '
        f'padding:18px; border-radius:12px; text-align:center; '
        f'border:1px solid {color}50; flex:1; min-width:160px;">'
        f'<div style="font-size:1.6em; font-weight:700; color:{color};">{value}</div>'
        f'<div style="font-size:0.85em; color:#555;">{label}</div></div>'
    )


def _build_charts(p, ann, x_lbl):
    """Construit toutes les figures principales du rapport."""
    K = lambda v: [x / 1000 for x in v]
    leg = dict(orientation="h", y=-0.18, xanchor="center", x=0.5)
    figs = {}

    # 1. Ventes par service
    fig = go.Figure()
    for col, lbl, clr in [
        ("ca_hebergement", "Hebergement", "#667eea"),
        ("ca_brasserie", "Brasserie", "#f5576c"),
        ("ca_bar", "Bar", "#ffcc00"),
        ("ca_spa", "Spa", "#11998e"),
        ("ca_salles", "Salles", "#a0522d"),
        ("ca_loyer_restaurant", "Location resto.", "#ff8c00"),
    ]:
        if col in ann.columns:
            fig.add_trace(go.Bar(x=x_lbl, y=K(ann[col].values), name=lbl, marker_color=clr))
    fig.update_layout(title="Ventes par service (K€)", barmode="stack",
                      xaxis=dict(type="category"), yaxis=dict(tickformat=",.0f"), legend=leg)
    figs["ventes"] = fig

    # 2. Marge par service
    ann["marge_heberg"] = ann["ca_hebergement"] - ann["cv_hebergement"] - ann["cf_directs_hebergement"]
    ann["marge_brass"] = ann["ca_brasserie"] - ann["cv_brasserie"] - ann["cf_directs_brasserie"]
    ann["marge_bar"] = ann["ca_bar"] - ann["cv_bar"] - ann["cf_directs_bar"]
    ann["marge_spa"] = ann["ca_spa"] - ann["cv_spa"] - ann["cf_directs_spa"]
    ann["marge_salles"] = ann["ca_salles"] - ann["cf_directs_evenements"]
    ann["marge_resto"] = ann.get("ca_loyer_restaurant", 0)
    fig = go.Figure()
    for col, lbl, clr in [
        ("marge_heberg", "Hebergement", "#667eea"),
        ("marge_brass", "Brasserie", "#f5576c"),
        ("marge_bar", "Bar", "#ffcc00"),
        ("marge_spa", "Spa", "#11998e"),
        ("marge_salles", "Salles", "#a0522d"),
        ("marge_resto", "Location resto.", "#ff8c00"),
    ]:
        fig.add_trace(go.Bar(x=x_lbl, y=K(ann[col].values), name=lbl, marker_color=clr))
    fig.update_layout(title="Marge par service (K€)", barmode="stack",
                      xaxis=dict(type="category"), yaxis=dict(tickformat=",.0f"), legend=leg)
    figs["marge"] = fig

    # 3. Du EBITDA au Resultat Net (avec Subside RW)
    ebitda_k = K(ann["ebitda"].values)
    amort_k = K(ann["amortissement"].values)
    int_k = K(ann["dette_interets"].values)
    sub_k = K(ann["subside_rw"].values) if "subside_rw" in ann.columns else [0] * len(x_lbl)
    rn_k = K(ann["resultat_net"].values)
    fig = go.Figure()
    fig.add_trace(go.Bar(x=x_lbl, y=ebitda_k, name="EBITDA", marker_color="#11998e"))
    fig.add_trace(go.Bar(x=x_lbl, y=[-v for v in amort_k], name="Amortissement", marker_color="#764ba2"))
    fig.add_trace(go.Bar(x=x_lbl, y=[-v for v in int_k], name="Interets", marker_color="#ffcc00"))
    fig.add_trace(go.Bar(x=x_lbl, y=sub_k, name="Subsides RW", marker_color="#f093fb"))
    fig.add_trace(go.Scatter(x=x_lbl, y=rn_k, name="Resultat Net", mode="lines+markers+text",
                             line=dict(color="#f5576c", width=3),
                             text=[f"<b>{v:,.0f}</b>" for v in rn_k],
                             textposition="bottom center", textfont=dict(size=11, color="#b71c2e")))
    fig.add_hline(y=0, line_dash="dot", line_color="gray")
    fig.update_layout(title="Du EBITDA au Resultat Net (K€)", barmode="group",
                      xaxis=dict(type="category"), yaxis=dict(tickformat=",.0f"), legend=leg)
    figs["ebitda_rn"] = fig

    # 4. Cash flow + tresorerie
    surplus = (p.get("fonds_propres_initial", 0)
               + sum(pr["montant"] for pr in p.get("prets", []))
               - sum(inv["montant"] for inv in p.get("investissements", [])))
    cf_k = K(ann["cash_flow"].values)
    treso_k = [v / 1000 + surplus / 1000 for v in ann["cash_flow_cumul"].values]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=x_lbl, y=cf_k, name="Cash flow", marker_color="#4facfe"))
    fig.add_trace(go.Scatter(x=x_lbl, y=treso_k, name="Tresorerie cumulee",
                             mode="lines+markers", line=dict(color="#f5576c", width=3)))
    fig.add_hline(y=0, line_dash="dot", line_color="gray")
    fig.update_layout(title="Cash flow & Tresorerie (K€)", xaxis=dict(type="category"),
                      yaxis=dict(tickformat=",.0f"), legend=leg)
    figs["cashflow"] = fig

    # 5. Endettement par pret
    enc_by_pret = {}
    from datetime import date as _date
    for pr in p.get("prets", []):
        nm = pr.get("nom", "Pret")
        if pr.get("subside_rw"):
            nm += " (RW)"
        dfp = calc_tableau_pret(pr, p["date_ouverture"], p["nb_mois_projection"])
        vals = []
        for a in ann["annee"]:
            d_fin = _date(int(a), 12, 31)
            r = dfp[dfp["date"] <= d_fin]
            vals.append(r.iloc[-1]["capital_restant"] / 1000 if not r.empty else 0)
        enc_by_pret[nm] = vals
    fig = go.Figure()
    colors = ["#f5576c", "#667eea", "#11998e", "#ffcc00", "#764ba2", "#a0522d", "#f093fb"]
    for ci, (nm, vals) in enumerate(enc_by_pret.items()):
        fig.add_trace(go.Bar(x=x_lbl, y=vals, name=nm, marker_color=colors[ci % len(colors)]))
    fig.update_layout(title="Evolution endettement (K€)", barmode="stack",
                      xaxis=dict(type="category"), yaxis=dict(tickformat=",.0f"), legend=leg)
    figs["endettement"] = fig

    return figs


def build_rapport_html(plan_nom, params):
    """
    Genere le rapport complet en HTML interactif autonome.
    Retourne une string HTML pretente a etre telechargee.
    """
    p = params
    df = projection_complete(p)
    ann = df.groupby("annee").agg({
        "ca_total": "sum", "ca_hebergement": "sum", "ca_brasserie": "sum",
        "ca_bar": "sum", "ca_spa": "sum", "ca_salles": "sum",
        "ca_loyer_restaurant": "sum",
        "cv_hebergement": "sum", "cv_brasserie": "sum", "cv_bar": "sum", "cv_spa": "sum",
        "cf_directs_hebergement": "sum", "cf_directs_brasserie": "sum",
        "cf_directs_bar": "sum", "cf_directs_spa": "sum", "cf_directs_evenements": "sum",
        "marge": "sum", "subside_rw": "sum",
        "ebitda": "sum", "amortissement": "sum", "dette_interets": "sum",
        "resultat_net": "sum", "cash_flow": "sum", "cash_flow_cumul": "last",
    }).reset_index()
    x_lbl = [str(int(a)) for a in ann["annee"]]

    figs = _build_charts(p, ann, x_lbl)

    # ── KPIs ──
    nb_ch = p["nb_chambres"]
    total_inv = sum(i["montant"] for i in p.get("investissements", []))
    total_inv += sum(i["montant"] for i in p.get("rocher_data", {}).get("investissements", []))
    total_fp = (p.get("fonds_propres_initial", 0)
                + p.get("rocher_data", {}).get("fonds_propres_initial", 0))
    total_dette = sum(pr["montant"] for pr in p.get("prets", []))
    total_dette += sum(pr["montant"] for pr in p.get("rocher_data", {}).get("prets", []))

    cover = _img_data_uri(Path(__file__).parent / "assets" / "chateau_1.jpg")
    cover_html = ""
    if cover:
        cover_html = (
            f'<div style="position:relative; border-radius:14px; overflow:hidden; margin-bottom:24px;">'
            f'<img src="{cover}" style="width:100%; height:280px; object-fit:cover; filter:brightness(0.45);">'
            f'<div style="position:absolute; top:50%; left:50%; transform:translate(-50%,-50%); text-align:center; width:90%;">'
            f'<p style="color:rgba(255,255,255,0.7); font-size:0.9em; letter-spacing:3px; '
            f'text-transform:uppercase; margin:0 0 8px 0;">Plan Financier</p>'
            f'<h1 style="color:white; margin:0; font-size:2.4em; '
            f'text-shadow:0 2px 8px rgba(0,0,0,0.5); font-weight:700;">'
            f'{p.get("nom_hotel", plan_nom)}</h1>'
            f'</div></div>'
        )

    # ── TOC ──
    sections = [
        ("sec-1", "Synthese", "#4f46e5"),
        ("sec-2", "Chiffre d'affaires", "#11998e"),
        ("sec-3", "Marges & Resultat", "#f5576c"),
        ("sec-4", "Cash flow", "#4facfe"),
        ("sec-5", "Endettement", "#764ba2"),
        ("sec-6", "Tableau annuel", "#0b6e66"),
    ]
    toc = "".join(
        f'<a href="#{aid}" style="display:inline-block; padding:6px 12px; margin:3px; '
        f'background:{clr}15; color:{clr}; border:1px solid {clr}50; border-radius:6px; '
        f'text-decoration:none; font-size:0.85em; font-weight:600;">{lbl}</a>'
        for aid, lbl, clr in sections
    )

    # ── Tableau annuel ──
    fmt = lambda v: f"{v/1000:,.0f}".replace(",", " ")
    tbl_rows = []
    for _, row in ann.iterrows():
        tbl_rows.append([
            int(row["annee"]),
            fmt(row["ca_total"]),
            fmt(row["marge"]),
            fmt(row["ebitda"]),
            fmt(row["resultat_net"]),
            fmt(row["cash_flow"]),
            fmt(row["cash_flow_cumul"]),
        ])
    tbl_html = _table_html(
        ["Annee", "CA total (K€)", "Marge (K€)", "EBITDA (K€)", "RN (K€)",
         "Cash flow (K€)", "CF cumule (K€)"],
        tbl_rows,
    )

    # ── Assemblage HTML ──
    style = """
    <style>
        * { box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
               max-width: 1200px; margin: 0 auto; padding: 16px; color: #1f2937; line-height: 1.5; }
        h1, h2, h3 { color: #1f2937; }
        h2 { border-bottom: 3px solid #4facfe; padding-bottom: 8px; margin-top: 36px;
             scroll-margin-top: 80px; }
        .toc { position: sticky; top: 0; z-index: 100; background: white;
               padding: 12px 14px; margin: 0 0 16px 0; border-radius: 10px;
               box-shadow: 0 2px 12px rgba(0,0,0,0.08); border: 1px solid #e5e7eb; }
        .toc-label { font-size: 0.75em; font-weight: 700; color: #6b7280;
                     text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px; }
        .data-table { width: 100%; border-collapse: collapse; margin: 16px 0;
                      font-size: 0.9em; }
        .data-table th { background: #f3f4f6; padding: 10px; text-align: left;
                         border-bottom: 2px solid #d1d5db; }
        .data-table td { padding: 8px 10px; border-bottom: 1px solid #e5e7eb;
                         text-align: right; }
        .data-table td:first-child { text-align: left; font-weight: 600; }
        .kpi-row { display: flex; gap: 12px; margin: 16px 0; flex-wrap: wrap; }
        .footer { margin-top: 40px; padding-top: 16px; border-top: 2px solid #dee2e6;
                  text-align: center; color: #888; font-size: 0.9em; }
    </style>
    """

    plotly_cdn = '<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>'

    html = f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="utf-8">
    <title>Plan Financier — {p.get('nom_hotel', plan_nom)}</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    {plotly_cdn}
    {style}
</head>
<body>
    {cover_html}

    <div class="toc">
        <div class="toc-label">Sommaire</div>
        <div>{toc}</div>
    </div>

    <h2 id="sec-1">Synthese</h2>
    <div class="kpi-row">
        {_kpi_card(f"{nb_ch}", "Chambres", "#4f46e5")}
        {_kpi_card(f"{total_inv/1e6:,.1f} M€", "Investissement total", "#16a34a")}
        {_kpi_card(f"{total_fp/1e6:,.1f} M€", "Fonds propres", "#2563eb")}
        {_kpi_card(f"{total_dette/1e6:,.1f} M€", "Endettement total", "#dc2626")}
    </div>
    <p>Debut d'activite : <b>{p['date_ouverture'].strftime('%B %Y')}</b> &mdash;
       Projection : <b>{p['nb_mois_projection']//12} ans</b></p>

    <h2 id="sec-2" style="border-bottom-color:#11998e;">Chiffre d'affaires</h2>
    {_fig_html(figs['ventes'], 450)}

    <h2 id="sec-3" style="border-bottom-color:#f5576c;">Marges & Resultat</h2>
    {_fig_html(figs['marge'], 450)}
    {_fig_html(figs['ebitda_rn'], 450)}

    <h2 id="sec-4" style="border-bottom-color:#4facfe;">Cash flow</h2>
    {_fig_html(figs['cashflow'], 400)}

    <h2 id="sec-5" style="border-bottom-color:#764ba2;">Endettement</h2>
    {_fig_html(figs['endettement'], 400)}

    <h2 id="sec-6" style="border-bottom-color:#0b6e66;">Tableau annuel</h2>
    {tbl_html}

    <div class="footer">
        Rapport genere pour le plan <b>{plan_nom}</b>
    </div>
</body>
</html>
"""
    return html
