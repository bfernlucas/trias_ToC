"""
Render PNG estático do mapa do ecossistema Trias Brasil.

Gera duas versões:
  - ecossistema_completo.png — todas as 167 organizações (com camada institucional)
  - ecossistema_brasileiro.png — apenas 118 nós (4 parceiros + 114 do ecossistema BR)

Usa a mesma estrutura de grafo, paleta e classificação do mapa interativo.
"""
from openpyxl import load_workbook
import networkx as nx
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.lines import Line2D
import re
from collections import Counter, defaultdict
from pathlib import Path
import importlib.util

# Reusa logica do build_network.py (carrega o modulo)
ROOT = Path("/home/user/trias_ToC")
SCRIPT = ROOT / "outputs" / "network" / "build_network.py"
OUT_NET = ROOT / "outputs" / "network"

# Como o build_network.py gera arquivos no import, vamos ler dados crus de
# forma independente.

DOCS_SRC = ROOT / "docs-referencia"
wb = load_workbook(DOCS_SRC / "Stakeholder Ecosystem Mapping.xlsx", data_only=True)
ws = wb["Stakeholder list"]
rows = list(ws.iter_rows(values_only=True))
records = []
for r in rows[6:]:
    if not r[1]:
        continue
    records.append({
        "id": str(r[1]).strip(),
        "name": (str(r[2]) if r[2] else "").strip(),
        "sector": (str(r[3]) if r[3] else "").strip(),
        "type": (str(r[4]) if r[4] else "").strip(),
        "description": (str(r[5]) if r[5] else "").strip(),
        "territory": (str(r[6]) if r[6] else "").strip(),
        "impact_area": (str(r[7]) if r[7] else "").strip(),
        "biome_primary": (str(r[8]) if r[8] else "").strip(),
        "biome_secondary": (str(r[9]) if r[9] else "").strip(),
        "role": (str(r[11]) if r[11] else "").strip(),
        "notes": (str(r[18]) if r[18] else "").strip(),
    })
_IKF_PAT = re.compile(r"\bIKEA\s+Foundation\b|\bIKF\b", re.IGNORECASE)
for rec in records:
    rec["notes"] = _IKF_PAT.sub("o financiador analisado", rec["notes"])
    rec["description"] = _IKF_PAT.sub("o financiador analisado", rec["description"])

# Importa as definições de PARTNERS, TRIAS_INSTITUTIONAL, scoring do build_network.py
spec = importlib.util.spec_from_file_location("build_net", SCRIPT)
# Não executamos o build_net inteiro — apenas extraímos as constantes via exec
# em namespace isolado.
import importlib, sys
src = SCRIPT.read_text(encoding="utf-8")
# Vamos truncar antes do `# ========== 7. JSON do grafo para D3 ==========`
# para evitar gerar HTMLs. Pegar tudo até o `Grafo:` print.
marker = "# ========== 7. JSON do grafo para D3 =========="
truncated = src.split(marker)[0]
ns = {"__name__": "build_net_partial"}
exec(truncated, ns)

PARTNERS = ns["PARTNERS"]
TRIAS_INSTITUTIONAL = ns["TRIAS_INSTITUTIONAL"]
SECTOR_COLOR = ns["SECTOR_COLOR"]
PARTNER_COLOR = ns["PARTNER_COLOR"]
INSTITUTIONAL_COLOR = ns["INSTITUTIONAL_COLOR"]
G = ns["G"]
results = ns["results"]

print(f"Grafo carregado: {G.number_of_nodes()} nós, {G.number_of_edges()} arestas")

# ========== Posicionamento ==========
plt.rcParams.update({
    "font.family": "DejaVu Sans",  # Roboto similar fallback
    "font.size": 8,
    "axes.facecolor": "white",
    "figure.facecolor": "white",
})

def compute_layout(G, partners_xy, institutional_zones=None, seed=42,
                   iterations=600, k=1.2):
    fixed = dict(partners_xy)
    initial = dict(partners_xy)
    if institutional_zones:
        for nid in G.nodes:
            if G.nodes[nid].get("kind") == "trias_institutional":
                subcat = G.nodes[nid].get("subcat", "")
                zone = institutional_zones.get(subcat, {"x": 0, "y": 0})
                initial[nid] = (zone["x"], zone["y"])
                fixed[nid] = initial[nid]
    pos = nx.spring_layout(G, k=k, iterations=iterations, pos=initial,
                          fixed=list(fixed.keys()), seed=seed, weight="weight")
    return pos

PARTNERS_XY = {
    "UNICAFES_PA":  (-3.0, 2.0),
    "UNICAFES_RO":  (3.0, 2.0),
    "CSA_BRASIL":   (-3.0, -2.0),
    "UNICATADORES": (3.0, -2.0),
}

INSTITUTIONAL_ZONES = {
    'Doador institucional':         { "x":  0.0, "y":  5.5 },
    'Rede peer (AgriCord)':         { "x":  5.8, "y":  3.0 },
    'Rede belga/europeia':          { "x": -5.8, "y":  1.5 },
    'Diplomacia/mercado':           { "x":  5.8, "y": -3.0 },
    'Setor privado parceiro':       { "x": -5.8, "y": -3.0 },
    'Setor privado parceiro · SAM': { "x":  0.0, "y": -5.5 },
    'Academia':                     { "x": -3.5, "y":  4.6 },
}

# ========== Função genérica de render ==========
# Localized strings
LANG = {
    "pt": {
        "title": "Ecossistema de stakeholders",
        "eyebrow": "TRIAS · BRASIL · DGD 2027–2031",
        "version": "Versão 4.0 · maio 2026",
        "mbo_partner": "Parceiro MBO",
        "inst_layer": "Camada institucional Trias",
        "connections": "Conexões",
        "edge_alignment": "Alinhamento temático/territorial",
        "edge_peer": "Peer MBO (rede UNICOPAS/Trias)",
        "edge_inst": "Vínculo institucional Trias",
        "title_brasileiro": "Rede ancorada nos 4 parceiros MBO · 114 stakeholders do ecossistema brasileiro",
        "title_completo": "Mapa completo · ecossistema brasileiro + 49 atores institucionais Trias",
    },
    "en": {
        "title": "Stakeholder Ecosystem",
        "eyebrow": "TRIAS · BRAZIL · DGD 2027–2031",
        "version": "Version 4.0 · May 2026",
        "mbo_partner": "MBO Partner",
        "inst_layer": "Trias institutional layer",
        "connections": "Connections",
        "edge_alignment": "Thematic / territorial alignment",
        "edge_peer": "MBO peer (UNICOPAS / Trias network)",
        "edge_inst": "Trias institutional tie",
        "title_brasileiro": "Network anchored on the 4 MBO partners · 114 stakeholders in the Brazilian ecosystem",
        "title_completo": "Full map · Brazilian ecosystem + 49 Trias institutional actors",
    },
}

def render(G, pos, out_path, title, show_institutional=True,
           figsize=(24, 16), label_threshold=0.0, lang="pt"):
    L = LANG[lang]
    fig, ax = plt.subplots(figsize=figsize, dpi=180)
    # auto bounds based on positions
    xs = [pos[n][0] for n in pos]
    ys = [pos[n][1] for n in pos]
    pad = 1.0
    ax.set_xlim(min(xs) - pad, max(xs) + pad)
    ax.set_ylim(min(ys) - pad, max(ys) + pad)
    ax.set_aspect("equal")
    ax.axis("off")

    # === Edges ===
    for u, v, data in G.edges(data=True):
        if u not in pos or v not in pos:
            continue
        if data.get("kind") == "institutional" and not show_institutional:
            continue
        x1, y1 = pos[u]; x2, y2 = pos[v]
        kind = data.get("kind", "alignment")
        if kind == "peer":
            ax.plot([x1, x2], [y1, y2], color="#B91C1C", alpha=0.30,
                    linestyle="--", linewidth=1.2, zorder=1)
        elif kind == "institutional":
            ax.plot([x1, x2], [y1, y2], color=INSTITUTIONAL_COLOR, alpha=0.18,
                    linewidth=0.6, zorder=1)
        else:
            w = data.get("weight", 4)
            ax.plot([x1, x2], [y1, y2], color="#0F172A", alpha=0.10,
                    linewidth=0.4 + w * 0.06, zorder=1)

    # === Nodes ===
    for nid, attrs in G.nodes(data=True):
        if nid not in pos: continue
        if attrs.get("kind") == "trias_institutional" and not show_institutional:
            continue
        x, y = pos[nid]
        kind = attrs.get("kind", "stakeholder")
        if kind == "partner":
            size = 1500
            ax.scatter(x, y, s=size, c=PARTNER_COLOR, marker="D",
                       edgecolors="white", linewidths=2.0, zorder=4)
        elif kind == "trias_institutional":
            size = 130
            ax.scatter(x, y, s=size, c=INSTITUTIONAL_COLOR, marker="s",
                       edgecolors="white", linewidths=1.0, zorder=3)
        else:
            n_p = attrs.get("n_partners", 1)
            size = 80 + n_p * 80
            ax.scatter(x, y, s=size, c=attrs.get("color", "#999"), marker="o",
                       edgecolors="white", linewidths=1.0, zorder=3)

    # === Labels === (com leve offset abaixo do nó, bbox branco para legibilidade)
    for nid, attrs in G.nodes(data=True):
        if nid not in pos: continue
        if attrs.get("kind") == "trias_institutional" and not show_institutional:
            continue
        x, y = pos[nid]
        kind = attrs.get("kind", "stakeholder")
        name = attrs.get("name", nid)
        if kind == "partner":
            ax.text(x, y - 0.35, name, ha="center", va="top", fontsize=13,
                    fontweight="bold", color="#0F172A", zorder=5)
        elif kind == "trias_institutional":
            short = name if len(name) <= 26 else name[:23] + "…"
            ax.text(x, y - 0.17, short, ha="center", va="top", fontsize=7.0,
                    color="#475569", zorder=5,
                    bbox=dict(facecolor="white", edgecolor="none",
                              boxstyle="round,pad=0.10", alpha=0.78))
        else:
            n_p = attrs.get("n_partners", 1)
            bw = attrs.get("betweenness", 0)
            importance = n_p >= 3 or bw > 0.015
            if importance:
                short = name if len(name) <= 30 else name[:27] + "…"
                ax.text(x, y - 0.17, short, ha="center", va="top", fontsize=7.6,
                        color="#0F172A", zorder=5, fontweight="500",
                        bbox=dict(facecolor="white", edgecolor="none",
                                  boxstyle="round,pad=0.10", alpha=0.85))
            else:
                short = name if len(name) <= 26 else name[:23] + "…"
                ax.text(x, y - 0.15, short, ha="center", va="top", fontsize=6.2,
                        color="#64748B", zorder=4,
                        bbox=dict(facecolor="white", edgecolor="none",
                                  boxstyle="round,pad=0.06", alpha=0.65))

    # === Header (titulo + meta) ===
    fig.text(0.04, 0.96, L["title"],
             fontsize=26, fontweight="bold", color="#0F172A", ha="left", va="top")
    fig.text(0.04, 0.925, title, fontsize=12, color="#334155",
             style="italic", ha="left", va="top")
    fig.text(0.04, 0.905, L["eyebrow"],
             fontsize=8.5, color="#64748B", ha="left", va="top",
             family="monospace")
    fig.text(0.96, 0.96, L["version"],
             fontsize=8.5, color="#64748B", ha="right", va="top",
             family="monospace")

    # === Node category legend ===
    legend_handles = [
        Line2D([0], [0], marker="D", color="w", markerfacecolor=PARTNER_COLOR,
               markeredgecolor="white", markeredgewidth=1.2, markersize=14,
               label=f"{L['mbo_partner']} ({sum(1 for n in G.nodes if G.nodes[n].get('kind')=='partner')})"),
    ]
    sector_counts = Counter()
    for n in G.nodes:
        if G.nodes[n].get("kind") == "stakeholder":
            sector_counts[G.nodes[n].get("sector", "Others")] += 1
    for sector, color in SECTOR_COLOR.items():
        legend_handles.append(
            Line2D([0], [0], marker="o", color="w", markerfacecolor=color,
                   markeredgecolor="white", markeredgewidth=1.0, markersize=11,
                   label=f"{sector} ({sector_counts.get(sector, 0)})")
        )
    if show_institutional:
        n_inst = sum(1 for n in G.nodes if G.nodes[n].get("kind") == "trias_institutional")
        legend_handles.append(
            Line2D([0], [0], marker="s", color="w", markerfacecolor=INSTITUTIONAL_COLOR,
                   markeredgecolor="white", markeredgewidth=1.0, markersize=11,
                   label=f"{L['inst_layer']} ({n_inst})")
        )

    leg = ax.legend(handles=legend_handles, loc="lower left",
                    bbox_to_anchor=(0.0, -0.02),
                    frameon=True, fontsize=9, labelspacing=0.7,
                    handletextpad=0.8, borderpad=1.0,
                    facecolor="white", edgecolor="#E1E5EB", ncol=1)
    leg.get_frame().set_linewidth(0.8)

    # === Edge legend ===
    edge_handles = [
        Line2D([0], [0], color="#0F172A", alpha=0.30, linewidth=1.4,
               label=L["edge_alignment"]),
        Line2D([0], [0], color="#B91C1C", alpha=0.5, linewidth=1.4,
               linestyle="--", label=L["edge_peer"]),
    ]
    if show_institutional:
        edge_handles.append(
            Line2D([0], [0], color=INSTITUTIONAL_COLOR, alpha=0.4, linewidth=1.0,
                   label=L["edge_inst"])
        )
    leg2 = ax.legend(handles=edge_handles, loc="lower right",
                     bbox_to_anchor=(1.0, -0.02),
                     frameon=True, fontsize=9, labelspacing=0.6,
                     handletextpad=0.6, borderpad=0.8,
                     facecolor="white", edgecolor="#E1E5EB", title=L["connections"],
                     title_fontsize=9)
    leg2.get_title().set_fontweight("600")
    leg2.get_title().set_color("#0F172A")
    leg2.get_frame().set_linewidth(0.8)
    ax.add_artist(leg)  # garante que ambas legendas fiquem

    plt.subplots_adjust(left=0.04, right=0.96, top=0.88, bottom=0.10)
    plt.savefig(out_path, dpi=180, bbox_inches="tight", facecolor="white",
                pad_inches=0.4)
    plt.close()
    print(f"PNG salvo: {out_path}")

# ========== Renders PT-BR ==========
G_br = G.copy()
inst_nodes = [n for n in G_br.nodes if G_br.nodes[n].get("kind") == "trias_institutional"]
G_br.remove_nodes_from(inst_nodes)
pos_br = compute_layout(G_br, PARTNERS_XY, seed=42, iterations=800, k=1.5)
render(G_br, pos_br,
       OUT_NET / "ecossistema_brasileiro.png",
       title=LANG["pt"]["title_brasileiro"],
       show_institutional=False,
       figsize=(26, 17), lang="pt")

pos_full = compute_layout(G, PARTNERS_XY, INSTITUTIONAL_ZONES, seed=42,
                          iterations=1000, k=1.2)
render(G, pos_full,
       OUT_NET / "ecossistema_completo.png",
       title=LANG["pt"]["title_completo"],
       show_institutional=True,
       figsize=(30, 20), lang="pt")

# ========== Renders English ==========
render(G_br, pos_br,
       OUT_NET / "ecosystem_brazil.png",
       title=LANG["en"]["title_brasileiro"],
       show_institutional=False,
       figsize=(26, 17), lang="en")
render(G, pos_full,
       OUT_NET / "ecosystem_full.png",
       title=LANG["en"]["title_completo"],
       show_institutional=True,
       figsize=(30, 20), lang="en")

print("Done.")
