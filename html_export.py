"""
Export du rapport de visualisation en HTML interactif autonome.

Genere un fichier HTML self-contained avec :
- Page de garde + KPIs
- Sommaire sticky cliquable
- 8 sections couvrant l'integralite du rapport (Montage, Capital, Injection,
  Investissements, Moyens & Besoins, Plan Rocher, Plan Chateau, Simulation)
- Graphiques Plotly interactifs (zoom, hover, legende cliquable)
- Plotly charge via CDN

Le fichier est ouvrable hors-ligne (sauf Plotly CDN). Partageable par email.
"""

import base64
from datetime import date as _date
from pathlib import Path

import plotly.graph_objects as go
import plotly.io as pio

from calculs import projection_complete, calc_tableau_pret


# ────────────────────────────────────────────────────────────────────────────
# Helpers
# ────────────────────────────────────────────────────────────────────────────

def _fig_html(fig, height=400):
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


def _fmt_eur(v):
    return f"{v:,.0f} €".replace(",", " ")


def _fmt_k(v):
    return f"{v/1000:,.0f}".replace(",", " ")


# ────────────────────────────────────────────────────────────────────────────
# Construction des figures principales
# ────────────────────────────────────────────────────────────────────────────

def _build_charts(p, ann, x_lbl):
    K = lambda v: [x / 1000 for x in v]
    leg = dict(orientation="h", y=-0.18, xanchor="center", x=0.5)
    figs = {}

    # Ventes par service
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

    # Marges intermediaires
    ann["marge_heberg"] = ann["ca_hebergement"] - ann["cv_hebergement"] - ann["cf_directs_hebergement"]
    ann["marge_brass"] = ann["ca_brasserie"] - ann["cv_brasserie"] - ann["cf_directs_brasserie"]
    ann["marge_bar"] = ann["ca_bar"] - ann["cv_bar"] - ann["cf_directs_bar"]
    ann["marge_spa"] = ann["ca_spa"] - ann["cv_spa"] - ann["cf_directs_spa"]
    ann["marge_salles"] = ann["ca_salles"] - ann["cf_directs_evenements"]
    ann["marge_resto"] = ann.get("ca_loyer_restaurant", 0)

    # Marge par service (sans subside)
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
    marge_tot = K(ann["marge"].values)
    max_mh = max(marge_tot) if marge_tot else 1
    fig.add_trace(go.Scatter(x=x_lbl, y=[v + max_mh * 0.04 for v in marge_tot], mode="text",
                             text=[f"<b>{v:,.0f}</b>" for v in marge_tot], textposition="top center",
                             textfont=dict(size=11, color="#1a1a6e"), showlegend=False, hoverinfo="skip"))
    fig.update_layout(title="Marge par service (K€)", barmode="stack",
                      xaxis=dict(type="category"),
                      yaxis=dict(tickformat=",.0f", range=[0, max_mh * 1.5]), legend=leg)
    figs["marge"] = fig

    # Repartition des marges (camembert cumul)
    marge_cumul = {
        "Hebergement": ann["marge_heberg"].sum(),
        "Brasserie": ann["marge_brass"].sum(),
        "Bar": ann["marge_bar"].sum(),
        "Spa": ann["marge_spa"].sum(),
        "Salles": ann["marge_salles"].sum(),
        "Location resto.": ann["marge_resto"].sum() if isinstance(ann["marge_resto"], type(ann["marge_heberg"])) else 0,
    }
    marge_cumul = {k: v for k, v in marge_cumul.items() if v > 0}
    if marge_cumul:
        clr_map = {"Hebergement": "#667eea", "Brasserie": "#f5576c", "Bar": "#ffcc00",
                   "Spa": "#11998e", "Salles": "#a0522d", "Location resto.": "#ff8c00"}
        fig = go.Figure(go.Pie(labels=list(marge_cumul.keys()), values=list(marge_cumul.values()),
                               marker=dict(colors=[clr_map[k] for k in marge_cumul]),
                               textinfo="label+percent", textfont=dict(size=13)))
        fig.update_layout(title="Repartition des marges (cumul sur la projection)")
        figs["marge_pie"] = fig

    # Frais fixes indirects
    if "cf_indirects_total" in ann.columns:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=x_lbl, y=K(ann["cf_indirects_total"].values),
                             name="CF indirects", marker_color="#a78bfa"))
        fig.update_layout(title="Frais fixes indirects (K€)", xaxis=dict(type="category"),
                          yaxis=dict(tickformat=",.0f"), legend=leg)
        figs["cf_indirects"] = fig

    # Du EBITDA au Resultat Net (avec Subside RW)
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

    # Cash flow + tresorerie
    surplus = (p.get("fonds_propres_initial", 0)
               + sum(pr["montant"] for pr in p.get("prets", []))
               - sum(inv["montant"] for inv in p.get("investissements", [])))
    cf_k = K(ann["cash_flow"].values)
    treso_k = [v / 1000 + surplus / 1000 for v in ann["cash_flow_cumul"].values]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=x_lbl, y=cf_k, name="Cash flow", marker_color="#4facfe",
                         text=[f"<b>{v:,.0f}</b>" for v in cf_k], textposition="outside",
                         textfont=dict(size=10, color="#1a5e96")))
    fig.add_trace(go.Scatter(x=x_lbl, y=treso_k, name="Tresorerie cumulee",
                             mode="lines+markers+text", line=dict(color="#f5576c", width=3),
                             text=[f"<b>{v:,.0f}</b>" if v < 0 else "" for v in treso_k],
                             textposition="bottom center", textfont=dict(size=10, color="#b71c2e")))
    fig.add_hline(y=0, line_dash="dot", line_color="gray")
    fig.update_layout(title="Cash flow & Tresorerie (K€)", xaxis=dict(type="category"),
                      yaxis=dict(tickformat=",.0f"), legend=leg)
    figs["cashflow"] = fig

    # Bilan : evolution FP / FP+Quasi-FP
    fp_init = p.get("fonds_propres_initial", 0)
    pret_rocher_mt = next((pr["montant"] for pr in p.get("prets", []) if "rocher" in pr.get("nom", "").lower()), 0)
    res_cum = ann["resultat_net"].cumsum()
    fp_k = [(fp_init + v) / 1000 for v in res_cum]
    fp_quasi_k = [(fp_init + pret_rocher_mt + v) / 1000 for v in res_cum]
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=x_lbl, y=fp_k, name="Fonds propres", mode="lines+markers+text",
                             line=dict(color="#11998e", width=3),
                             text=[f"<b>{v:,.0f}</b>" for v in fp_k], textposition="bottom center",
                             textfont=dict(size=10, color="#0b6e66")))
    fig.add_trace(go.Scatter(x=x_lbl, y=fp_quasi_k, name="FP + Quasi-FP", mode="lines+markers",
                             line=dict(color="#667eea", width=2, dash="dash")))
    fig.add_hline(y=0, line_dash="dot", line_color="gray")
    fig.update_layout(title="Evolution des fonds propres (K€)", xaxis=dict(type="category"),
                      yaxis=dict(tickformat=",.0f"), legend=leg)
    figs["bilan_fp"] = fig

    # Endettement par pret
    enc_by_pret = {}
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
    enc_total_k = [sum(v[i] for v in enc_by_pret.values()) for i in range(len(x_lbl))]
    fig.add_trace(go.Scatter(x=x_lbl, y=enc_total_k, mode="markers+text",
                             text=[f"<b>{v:,.0f}</b>" for v in enc_total_k], textposition="top center",
                             textfont=dict(size=11, color="#1a1a6e"),
                             marker=dict(size=1, color="rgba(0,0,0,0)"),
                             showlegend=False, hoverinfo="skip"))
    fig.update_layout(title="Evolution endettement (K€)", barmode="stack",
                      xaxis=dict(type="category"), yaxis=dict(tickformat=",.0f"), legend=leg)
    figs["endettement"] = fig

    # Ratio solvabilite
    ratios = []
    for i in range(len(x_lbl)):
        enc_i_eur = sum(v[i] for v in enc_by_pret.values()) * 1000 if enc_by_pret else 0
        res_cum_i = res_cum.iloc[i] if i < len(res_cum) else 0
        fp_quasi = fp_init + res_cum_i + pret_rocher_mt
        total_passif = enc_i_eur + fp_init + res_cum_i + pret_rocher_mt
        ratios.append(fp_quasi / total_passif * 100 if total_passif > 0 else 0)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=x_lbl, y=ratios, name="Ratio solvabilite",
                             mode="lines+markers+text", line=dict(color="#667eea", width=3),
                             fill="tozeroy", fillcolor="rgba(102,126,234,0.1)",
                             text=[f"<b>{v:.0f}%</b>" for v in ratios], textposition="top center",
                             textfont=dict(size=11, color="#23408f")))
    fig.add_hline(y=30, line_dash="dash", line_color="orange",
                  annotation_text="Seuil 30%")
    fig.update_layout(title="Ratio (FP+Quasi-FP) / Total passif (%)",
                      xaxis=dict(type="category"), yaxis=dict(range=[0, 105], tickformat=".0f"))
    figs["solvabilite"] = fig

    # Camembert capital social (Rocher + Chateau)
    rd = p.get("rocher_data", {})
    cap_inv = {}
    for inv in rd.get("fonds_propres_investisseurs", []):
        cap_inv[inv["nom"]] = cap_inv.get(inv["nom"], 0) + inv["montant"]
    for inv in p.get("fonds_propres_investisseurs", []):
        cap_inv[inv["nom"]] = cap_inv.get(inv["nom"], 0) + inv["montant"]
    if cap_inv:
        pie_clrs = ["#4f46e5", "#11998e", "#f5576c", "#ffcc00", "#764ba2", "#a0522d"]
        fig = go.Figure(go.Pie(labels=list(cap_inv.keys()), values=list(cap_inv.values()),
                               marker=dict(colors=pie_clrs[:len(cap_inv)]),
                               textinfo="label+value+percent",
                               texttemplate="<b>%{label}</b><br>%{value:,.0f} €<br>(%{percent})",
                               textfont=dict(size=12), hole=0.3))
        fig.update_layout(title="Repartition du capital social (consolide)")
        figs["capital_pie"] = fig

    return figs


# ────────────────────────────────────────────────────────────────────────────
# Sections du rapport
# ────────────────────────────────────────────────────────────────────────────

def _section_montage(p):
    """Section 1 : Montage financier (Rocher + Chateau)."""
    rd = p.get("rocher_data", {})
    fp_ch = p.get("fonds_propres_initial", 0)
    fp_ro = rd.get("fonds_propres_initial", 0)
    prets_ch = p.get("prets", [])
    prets_ro = rd.get("prets", [])

    def _fp_block(fp_total, investisseurs):
        rows = "".join(
            f'<li>{inv["nom"]} : <b>{_fmt_eur(inv["montant"])}</b></li>'
            for inv in investisseurs
        )
        return (
            f'<p><b>Fonds propres : {_fmt_eur(fp_total)}</b></p>'
            f'<ul style="margin:0 0 8px 0;">{rows}</ul>'
        )

    def _prets_block(prets):
        if not prets:
            return ""
        std = [pr for pr in prets if not pr.get("subside_rw")]
        rw = [pr for pr in prets if pr.get("subside_rw")]
        html = ""
        if std:
            html += "<p><b>Dettes bancaires :</b></p><ul>"
            for pr in std:
                creancier = pr.get("creancier", "")
                creancier_lbl = f" ({creancier})" if creancier else ""
                html += (f'<li>{pr["nom"]}{creancier_lbl} : <b>{_fmt_eur(pr["montant"])}</b> '
                         f'@ {pr["taux_annuel"]:.1%} / {pr["duree_ans"]} ans</li>')
            html += "</ul>"
        if rw:
            html += "<p><b>Dettes garanties Region Wallonne :</b></p><ul>"
            for pr in rw:
                creancier = pr.get("creancier", "")
                creancier_lbl = f" ({creancier})" if creancier else ""
                html += (f'<li>🟢 {pr["nom"]}{creancier_lbl} : <b>{_fmt_eur(pr["montant"])}</b> '
                         f'@ {pr["taux_annuel"]:.1%} / {pr["duree_ans"]} ans &mdash; <i>Subside RW</i></li>')
            html += "</ul>"
        return html

    return (
        '<div class="two-col">'
        '<div class="col"><h3>Immobilière Rocher</h3>'
        f'{_fp_block(fp_ro, rd.get("fonds_propres_investisseurs", []))}'
        f'{_prets_block(prets_ro)}'
        '</div>'
        '<div class="col"><h3>Château d\'Argenteau</h3>'
        f'{_fp_block(fp_ch, p.get("fonds_propres_investisseurs", []))}'
        f'{_prets_block(prets_ch)}'
        '</div></div>'
    )


def _section_investissements(p):
    """Section 4 : Tableaux d'investissements pour Rocher et Chateau."""
    rd = p.get("rocher_data", {})

    def _inv_table(invs):
        rows = []
        total = 0
        for inv in invs:
            if inv.get("montant", 0) > 0:
                rows.append([
                    inv.get("categorie", "-"),
                    _fmt_eur(inv["montant"]),
                    f"{inv.get('duree_amort', 0)} ans" if inv.get("duree_amort", 0) > 0 else "-",
                ])
                total += inv["montant"]
        rows.append(["<b>Total</b>", f"<b>{_fmt_eur(total)}</b>", ""])
        return _table_html(["Categorie", "Montant", "Amortissement"], rows)

    return (
        '<div class="two-col">'
        '<div class="col"><h3>Immobilière Rocher</h3>'
        f'{_inv_table(rd.get("investissements", []))}'
        '</div>'
        '<div class="col"><h3>Château d\'Argenteau</h3>'
        f'{_inv_table(p.get("investissements", []))}'
        '</div></div>'
    )


def _section_personnel(p):
    """Tableau des effectifs (ETP)."""
    categories = [
        ("personnel_hebergement", "Hebergement"),
        ("personnel_brasserie", "Brasserie"),
        ("personnel_bar", "Bar"),
        ("personnel_spa", "Spa"),
        ("personnel_indirect", "Indirect"),
    ]
    rows = []
    total_etp = 0
    total_masse = 0
    for key, lbl in categories:
        for poste in p.get(key, []):
            etp = poste.get("etp", 0)
            cout = poste.get("cout_brut", 0)
            cp = p.get("charges_patronales_pct", 0.35)
            masse = etp * cout * (1 + cp)
            rows.append([lbl, poste.get("poste", "-"), f"{etp:.1f}", _fmt_eur(masse)])
            total_etp += etp
            total_masse += masse
    rows.append(["", "<b>Total</b>", f"<b>{total_etp:.1f}</b>", f"<b>{_fmt_eur(total_masse)}</b>"])
    return _table_html(["Categorie", "Poste", "ETP", "Masse salariale annuelle"], rows)


def _section_pl(ann):
    """Tableau P&L annuel detaille."""
    rows = []
    for _, row in ann.iterrows():
        rows.append([
            int(row["annee"]),
            _fmt_k(row["ca_total"]),
            _fmt_k(row["marge"]),
            _fmt_k(row.get("ebitda", 0)),
            _fmt_k(row["resultat_net"]),
            _fmt_k(row["cash_flow"]),
            _fmt_k(row["cash_flow_cumul"]),
        ])
    return _table_html(
        ["Annee", "CA total (K€)", "Marge (K€)", "EBITDA (K€)", "RN (K€)",
         "Cash flow (K€)", "CF cumule (K€)"],
        rows,
    )


# ────────────────────────────────────────────────────────────────────────────
# Calculs et graphes Rocher (replique de app.py _module_rocher partiel)
# ────────────────────────────────────────────────────────────────────────────

def _compute_rocher(p):
    """Calcule les flux annuels de l'Immobiliere Rocher.
    Retourne (ann_ro, prets_ro, fp_ro, x_lbl)."""
    import pandas as pd
    from dateutil.relativedelta import relativedelta

    rd = p.get("rocher_data", {})
    fp_ro = rd.get("fonds_propres_initial", 0)
    prets_ro = rd.get("prets", [])
    loyer_m = p.get("loyer_mensuel", 0)
    nb_mois = p["nb_mois_projection"]

    # Pret du Chateau au Rocher : Rocher RECOIT les remboursements
    pret_rocher_ch = next(
        (pr for pr in p.get("prets", []) if "rocher" in pr.get("nom", "").lower()), None
    )
    df_pret_ch_ro = (
        calc_tableau_pret(pret_rocher_ch, p["date_ouverture"], nb_mois)
        if pret_rocher_ch else pd.DataFrame()
    )

    # Prets bancaires Rocher (charges)
    dfs_ro = {
        pr["nom"]: calc_tableau_pret(pr, p["date_ouverture"], nb_mois)
        for pr in prets_ro if pr["montant"] > 0
    }

    # Amortissement annuel Rocher
    amort_an_ro = sum(
        inv["montant"] / inv["duree_amort"]
        for inv in rd.get("investissements", [])
        if inv.get("duree_amort", 0) > 0
    )
    amort_m_ro = amort_an_ro / 12

    rows_ro = []
    for m in range(nb_mois):
        d = p["date_ouverture"] + relativedelta(months=m)
        rev = loyer_m
        int_rec = 0; cap_rec = 0
        if not df_pret_ch_ro.empty:
            r = df_pret_ch_ro[df_pret_ch_ro["date"] == d]
            if not r.empty:
                int_rec = r.iloc[0]["interets"]
                cap_rec = r.iloc[0]["capital"]
        int_pay = 0; cap_pay = 0
        for _df in dfs_ro.values():
            r = _df[_df["date"] == d]
            if not r.empty:
                int_pay += r.iloc[0]["interets"]
                cap_pay += r.iloc[0]["capital"]
        # Subside RW Rocher (15 ans, a partir de An 2)
        annee_idx = m // 12
        sub_rw = 0
        if annee_idx >= 1:
            for _pr in prets_ro:
                if _pr.get("subside_rw"):
                    duree_sub = 15
                    if (annee_idx - 1) < duree_sub:
                        sub_rw += _pr["montant"] / duree_sub / 12
        resultat = rev + int_rec - int_pay - amort_m_ro + sub_rw
        cash = rev + int_rec + cap_rec - int_pay - cap_pay
        rows_ro.append({
            "date": d, "annee": d.year,
            "loyer": rev, "interets_recus": int_rec, "capital_recu": cap_rec,
            "interets_payes": int_pay, "capital_paye": cap_pay,
            "amortissement": amort_m_ro, "subside_rw": sub_rw,
            "produits": rev + int_rec + sub_rw,
            "charges": int_pay + amort_m_ro,
            "resultat": resultat, "cash": cash,
        })
    df_ro = pd.DataFrame(rows_ro)
    df_ro["cash_cumul"] = df_ro["cash"].cumsum()
    df_ro["res_cumul"] = df_ro["resultat"].cumsum()
    ann_ro = df_ro.groupby("annee").agg({
        "loyer": "sum", "interets_recus": "sum", "capital_recu": "sum",
        "interets_payes": "sum", "capital_paye": "sum",
        "amortissement": "sum", "subside_rw": "sum",
        "produits": "sum", "charges": "sum",
        "resultat": "sum", "cash": "sum",
        "cash_cumul": "last", "res_cumul": "last",
    }).reset_index()
    return ann_ro, prets_ro, fp_ro


def _build_rocher_charts(p, ann_ro, prets_ro, fp_ro):
    K = lambda v: [x / 1000 for x in v]
    leg = dict(orientation="h", y=-0.18, xanchor="center", x=0.5)
    figs = {}
    xr = [str(int(a)) for a in ann_ro["annee"]]

    # P&L Rocher
    fig = go.Figure()
    fig.add_trace(go.Bar(x=xr, y=K(ann_ro["produits"].values), name="Produits", marker_color="#38ef7d"))
    fig.add_trace(go.Bar(x=xr, y=[-v for v in K(ann_ro["charges"].values)], name="Charges", marker_color="#f5576c"))
    res_k = K(ann_ro["resultat"].values)
    fig.add_trace(go.Scatter(x=xr, y=res_k, name="Resultat",
                             mode="lines+markers+text", line=dict(color="#3b3b98", width=3),
                             text=[f"<b>{v:,.0f}</b>" for v in res_k],
                             textposition="top center", textfont=dict(size=11, color="#23408f")))
    fig.add_hline(y=0, line_dash="dot", line_color="gray")
    fig.update_layout(title="Rocher - P&L (K€)", barmode="group",
                      xaxis=dict(type="category"), yaxis=dict(tickformat=",.0f"), legend=leg)
    figs["ro_pnl"] = fig

    # Cash flow Rocher
    fig = go.Figure()
    cash_k = K(ann_ro["cash"].values)
    cumul_k = K(ann_ro["cash_cumul"].values)
    fig.add_trace(go.Bar(x=xr, y=cash_k, name="Cash periode", marker_color="#4facfe",
                         text=[f"<b>{v:,.0f}</b>" for v in cash_k], textposition="outside",
                         textfont=dict(size=10, color="#1a5e96")))
    fig.add_trace(go.Scatter(x=xr, y=cumul_k, name="Cash cumule",
                             mode="lines+markers+text", line=dict(color="#f5576c", width=3),
                             text=[f"<b>{v:,.0f}</b>" for v in cumul_k],
                             textposition="top center", textfont=dict(size=11, color="#b71c2e")))
    fig.add_hline(y=0, line_dash="dot", line_color="gray")
    fig.update_layout(title="Rocher - Cash flow (K€)", xaxis=dict(type="category"),
                      yaxis=dict(tickformat=",.0f"), legend=leg)
    figs["ro_cashflow"] = fig

    # Fonds propres Rocher
    fig = go.Figure()
    cp_k = [(fp_ro + v) / 1000 for v in ann_ro["res_cumul"].values]
    fig.add_trace(go.Scatter(x=xr, y=cp_k, name="Capitaux propres",
                             mode="lines+markers+text", line=dict(color="#667eea", width=3),
                             fill="tozeroy", fillcolor="rgba(102,126,234,0.15)",
                             text=[f"<b>{v:,.0f}</b>" for v in cp_k],
                             textposition="top center", textfont=dict(size=11, color="#3b3b98")))
    fig.update_layout(title="Rocher - Fonds propres (K€)", xaxis=dict(type="category"),
                      yaxis=dict(tickformat=",.0f"), legend=leg)
    figs["ro_fp"] = fig

    # Endettement Rocher
    enc_by_pret = {}
    from datetime import date as _date
    for pr in prets_ro:
        if pr["montant"] > 0:
            nm = pr["nom"]
            if pr.get("subside_rw"):
                nm += " (RW)"
            dfp = calc_tableau_pret(pr, p["date_ouverture"], p["nb_mois_projection"])
            vals = []
            for a in ann_ro["annee"]:
                d_fin = _date(int(a), 12, 31)
                r = dfp[dfp["date"] <= d_fin]
                vals.append(r.iloc[-1]["capital_restant"] / 1000 if not r.empty else 0)
            enc_by_pret[nm] = vals
    if enc_by_pret:
        fig = go.Figure()
        colors = ["#f5576c", "#667eea", "#11998e", "#ffcc00", "#764ba2", "#a0522d", "#f093fb"]
        for ci, (nm, vals) in enumerate(enc_by_pret.items()):
            fig.add_trace(go.Bar(x=xr, y=vals, name=nm, marker_color=colors[ci % len(colors)]))
        total_k = [sum(v[i] for v in enc_by_pret.values()) for i in range(len(xr))]
        fig.add_trace(go.Scatter(x=xr, y=total_k, mode="markers+text",
                                 text=[f"<b>{v:,.0f}</b>" for v in total_k],
                                 textposition="top center", textfont=dict(size=11, color="#1a1a6e"),
                                 marker=dict(size=1, color="rgba(0,0,0,0)"),
                                 showlegend=False, hoverinfo="skip", cliponaxis=False))
        fig.update_layout(title="Rocher - Evolution endettement (K€)", barmode="stack",
                          xaxis=dict(type="category"),
                          yaxis=dict(tickformat=",.0f",
                                     range=[0, max(total_k) * 1.18] if total_k else None),
                          legend=leg)
        figs["ro_endettement"] = fig

    return figs


def _section_rocher_table(ann_ro):
    """Tableau Details des chiffres Rocher."""
    fmt = lambda v: f"{v:,.0f} €".replace(",", " ")
    rows = []
    for _, row in ann_ro.iterrows():
        rows.append([
            int(row["annee"]),
            fmt(row["loyer"]),
            fmt(row["interets_recus"]),
            fmt(row["subside_rw"]),
            fmt(row["capital_recu"]),
            fmt(row["interets_payes"]),
            fmt(row["amortissement"]),
            fmt(row["resultat"]),
            fmt(row["res_cumul"]),
            fmt(row["capital_paye"]),
            fmt(row["cash"]),
            fmt(row["cash_cumul"]),
        ])
    return _table_html(
        ["Annee", "Loyer recu", "Interets recus", "Subside RW", "Capital recu (Ch.)",
         "Interets payes", "Amortissement", "Resultat", "Resultat cumule",
         "Remb. capital paye", "Cash flow", "Cash flow cumul"],
        rows,
    )


# ────────────────────────────────────────────────────────────────────────────
# Assemblage final
# ────────────────────────────────────────────────────────────────────────────

def build_rapport_html(plan_nom, params):
    """Genere le rapport complet en HTML interactif autonome."""
    p = params
    df = projection_complete(p)
    ann = df.groupby("annee").agg({
        "ca_total": "sum", "ca_hebergement": "sum", "ca_brasserie": "sum",
        "ca_bar": "sum", "ca_spa": "sum", "ca_salles": "sum",
        "ca_loyer_restaurant": "sum",
        "cv_hebergement": "sum", "cv_brasserie": "sum", "cv_bar": "sum", "cv_spa": "sum",
        "cf_directs_hebergement": "sum", "cf_directs_brasserie": "sum",
        "cf_directs_bar": "sum", "cf_directs_spa": "sum", "cf_directs_evenements": "sum",
        "cf_indirects_total": "sum",
        "marge": "sum", "subside_rw": "sum",
        "ebitda": "sum", "amortissement": "sum", "dette_interets": "sum",
        "resultat_net": "sum", "cash_flow": "sum", "cash_flow_cumul": "last",
    }).reset_index()
    x_lbl = [str(int(a)) for a in ann["annee"]]

    figs = _build_charts(p, ann, x_lbl)

    # Donnees et figures Rocher
    ann_ro, prets_ro_data, fp_ro_init = _compute_rocher(p)
    figs_ro = _build_rocher_charts(p, ann_ro, prets_ro_data, fp_ro_init)
    rocher_table_html = _section_rocher_table(ann_ro)

    # ── KPIs ──
    rd = p.get("rocher_data", {})
    nb_ch = p["nb_chambres"]
    total_inv = sum(i["montant"] for i in p.get("investissements", []))
    total_inv += sum(i["montant"] for i in rd.get("investissements", []))
    total_fp = (p.get("fonds_propres_initial", 0) + rd.get("fonds_propres_initial", 0))
    total_dette = sum(pr["montant"] for pr in p.get("prets", []))
    total_dette += sum(pr["montant"] for pr in rd.get("prets", []))

    # Page de garde
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

    # Sommaire
    sections = [
        ("sec-1", "1. Montage financier", "#4facfe"),
        ("sec-2", "2. Capital social", "#4facfe"),
        ("sec-3", "3. Investissements", "#4facfe"),
        ("sec-4", "4. Plan Rocher", "#11998e"),
        ("sec-5", "5. Plan Chateau", "#f5576c"),
        ("sec-6", "6. Bilan & Endettement", "#764ba2"),
        ("sec-7", "7. Ratios", "#0b6e66"),
        ("sec-8", "8. Simulation", "#a78bfa"),
    ]
    toc = "".join(
        f'<a href="#{aid}" style="display:inline-block; padding:6px 12px; margin:3px; '
        f'background:{clr}15; color:{clr}; border:1px solid {clr}50; border-radius:6px; '
        f'text-decoration:none; font-size:0.85em; font-weight:600;">{lbl}</a>'
        for aid, lbl, clr in sections
    )

    # Commentaires
    comment_ro = p.get("commentaire_rocher", "")
    comment_ch = p.get("commentaire_chateau", "")

    # ── CSS ──
    style = """
    <style>
        * { box-sizing: border-box; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
               max-width: 1200px; margin: 0 auto; padding: 16px; color: #1f2937; line-height: 1.5; }
        h1, h2, h3 { color: #1f2937; }
        h2 { padding-bottom: 8px; margin-top: 36px; scroll-margin-top: 80px; border-bottom: 3px solid #4facfe; }
        h3 { margin-top: 22px; color: #374151; }
        .toc { position: sticky; top: 0; z-index: 100; background: white;
               padding: 12px 14px; margin: 0 0 16px 0; border-radius: 10px;
               box-shadow: 0 2px 12px rgba(0,0,0,0.08); border: 1px solid #e5e7eb; }
        .toc-label { font-size: 0.75em; font-weight: 700; color: #6b7280;
                     text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px; }
        .data-table { width: 100%; border-collapse: collapse; margin: 16px 0; font-size: 0.9em; }
        .data-table th { background: #f3f4f6; padding: 10px; text-align: left;
                         border-bottom: 2px solid #d1d5db; }
        .data-table td { padding: 8px 10px; border-bottom: 1px solid #e5e7eb; text-align: right; }
        .data-table td:first-child { text-align: left; font-weight: 600; }
        .kpi-row { display: flex; gap: 12px; margin: 16px 0; flex-wrap: wrap; }
        .two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin: 16px 0; }
        .col { background: #fafafa; padding: 16px; border-radius: 10px; border: 1px solid #e5e7eb; }
        .comment-box { background: #f0fdf4; border-left: 4px solid #11998e;
                       padding: 14px 18px; border-radius: 6px; margin: 16px 0;
                       font-size: 0.95em; line-height: 1.6; color: #1a3a2a; }
        .comment-box.ch { background: #fef2f2; border-left-color: #f5576c; color: #3a1a1a; }
        .info-box { background: #eff6ff; border-left: 4px solid #2563eb;
                    padding: 12px 16px; border-radius: 6px; margin: 16px 0;
                    font-size: 0.9em; color: #1e3a5f; }
        .footer { margin-top: 40px; padding-top: 16px; border-top: 2px solid #dee2e6;
                  text-align: center; color: #888; font-size: 0.9em; }
        @media (max-width: 768px) { .two-col { grid-template-columns: 1fr; } }
    </style>
    """

    plotly_cdn = '<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>'

    # ── HTML ──
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

    <div class="kpi-row">
        {_kpi_card(f"{nb_ch}", "Chambres", "#4f46e5")}
        {_kpi_card(f"{total_inv/1e6:,.1f} M€", "Investissement total", "#16a34a")}
        {_kpi_card(f"{total_fp/1e6:,.1f} M€", "Fonds propres", "#2563eb")}
        {_kpi_card(f"{total_dette/1e6:,.1f} M€", "Endettement total", "#dc2626")}
    </div>
    <p>Debut d'activite : <b>{p['date_ouverture'].strftime('%B %Y')}</b> &mdash;
       Projection : <b>{p['nb_mois_projection']//12} ans</b></p>

    <div class="toc">
        <div class="toc-label">Sommaire</div>
        <div>{toc}</div>
    </div>

    <h2 id="sec-1">1. Montage financier</h2>
    {_section_montage(p)}

    <h2 id="sec-2" style="border-bottom-color:#4facfe;">2. Capital social</h2>
    {_fig_html(figs['capital_pie'], 450) if 'capital_pie' in figs else '<p>Aucun investisseur renseigne.</p>'}

    <h2 id="sec-3" style="border-bottom-color:#4facfe;">3. Investissements initiaux</h2>
    {_section_investissements(p)}

    <h2 id="sec-4" style="border-bottom-color:#11998e;">4. Plan financier — Immobilière Rocher</h2>
    {f'<div class="comment-box">{comment_ro}</div>' if comment_ro else ''}
    <p><i>L'Immobilière Rocher porte les investissements immobiliers du projet et perçoit un loyer
    du Château d'Argenteau ainsi que les intérêts du prêt intra-groupe.</i></p>

    <h3>P&L</h3>
    {_fig_html(figs_ro['ro_pnl'], 450)}

    <h3>Cash flow</h3>
    {_fig_html(figs_ro['ro_cashflow'], 450)}

    <h3>Details des chiffres</h3>
    {rocher_table_html}

    <h3>Fonds propres</h3>
    {_fig_html(figs_ro['ro_fp'], 380)}

    <h3>Endettement</h3>
    {_fig_html(figs_ro.get('ro_endettement'), 400) if 'ro_endettement' in figs_ro else ''}

    <h2 id="sec-5" style="border-bottom-color:#f5576c;">5. Plan financier — Château d'Argenteau</h2>
    {f'<div class="comment-box ch">{comment_ch}</div>' if comment_ch else ''}

    <h3>Chiffre d'affaires</h3>
    {_fig_html(figs['ventes'], 450)}

    <h3>Marge par service</h3>
    {_fig_html(figs['marge'], 480)}
    {_fig_html(figs.get('marge_pie'), 400) if 'marge_pie' in figs else ''}

    <h3>Frais fixes indirects</h3>
    {_fig_html(figs.get('cf_indirects'), 350) if 'cf_indirects' in figs else ''}

    <h3>Personnel</h3>
    {_section_personnel(p)}

    <h3>Resultats — Du EBITDA au Resultat Net</h3>
    {_fig_html(figs['ebitda_rn'], 450)}

    <h3>Cash flow & Tresorerie</h3>
    {_fig_html(figs['cashflow'], 420)}

    <h3>Tableau annuel</h3>
    {_section_pl(ann)}

    <h2 id="sec-6" style="border-bottom-color:#764ba2;">6. Bilan & Endettement</h2>
    <h3>Evolution des fonds propres</h3>
    {_fig_html(figs['bilan_fp'], 380)}

    <h3>Encours de la dette</h3>
    {_fig_html(figs['endettement'], 400)}

    <h2 id="sec-7" style="border-bottom-color:#0b6e66;">7. Ratios</h2>
    <h3>Solvabilité — (FP + Quasi-FP) / Total passif</h3>
    {_fig_html(figs['solvabilite'], 380)}

    <h2 id="sec-8" style="border-bottom-color:#a78bfa;">8. Simulation</h2>
    <div class="info-box">
        <b>Note :</b> la simulation interactive (curseurs pour faire varier hypothèses
        et observer l'impact en temps réel sur EBITDA, Résultat Net et trésorerie)
        est disponible uniquement dans l'application Streamlit.<br>
        Les graphiques ci-dessous représentent l'état de référence (curseurs à zéro,
        sans variation appliquée) pour les indicateurs clés.
    </div>
    <h3>EBITDA — Référence</h3>
    {_fig_html(figs['ebitda_rn'], 380)}
    <h3>Cash flow & Trésorerie — Référence</h3>
    {_fig_html(figs['cashflow'], 380)}

    <div class="footer">
        Rapport généré pour le plan <b>{plan_nom}</b>
    </div>
</body>
</html>
"""
    return html
