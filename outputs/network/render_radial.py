"""
Render PNG do ecossistema Trias Brasil em layout radial/constelação.

A visualização representa o ecossistema como um sistema orbital:
- Centro: a Trias como hub articulador
- Anel interno (4 cardeais): os 4 parceiros MBO como âncoras
- Centro do diagrama: tri-bridges (atores que conectam 3 MBOs simultaneamente)
- Faixas internas entre MBOs: dual-bridges (atores que conectam 2 MBOs)
- Arcos externos por MBO: stakeholders alinhados a apenas 1 parceiro

Cores codificam setor; tamanho codifica número de MBOs alcançados.
Áreas de fundo coloridas (muito sutis) marcam o "território de influência"
de cada MBO.

Saídas: outputs/network/ecossistema_radial.png e ecosystem_radial_en.png
(versão inglesa).
"""
from openpyxl import load_workbook
import networkx as nx
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.lines import Line2D
from matplotlib.patches import Wedge, Circle, FancyArrowPatch
from matplotlib.path import Path as MplPath
from matplotlib.patches import PathPatch
import numpy as np
import re
from collections import Counter, defaultdict
from pathlib import Path
import importlib.util

ROOT = Path("/home/user/trias_ToC")
SCRIPT = ROOT / "outputs" / "network" / "build_network.py"
OUT_NET = ROOT / "outputs" / "network"
OUT_DOCS = ROOT / "docs"

# Reutiliza definições do build_network.py truncando antes da geração de HTML
src = SCRIPT.read_text(encoding="utf-8")
marker = "# ========== 7. JSON do grafo para D3 =========="
truncated = src.split(marker)[0]
ns = {"__name__": "build_net_partial"}
exec(truncated, ns)

PARTNERS = ns["PARTNERS"]
SECTOR_COLOR = ns["SECTOR_COLOR"]
PARTNER_COLOR = ns["PARTNER_COLOR"]
G = ns["G"]
results = ns["results"]
records = ns["records"]

print(f"Grafo carregado: {G.number_of_nodes()} nós, {G.number_of_edges()} arestas")

# ========== Configuração visual ==========
plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 9,
    "axes.facecolor": "#FAFBFC",
    "figure.facecolor": "#FAFBFC",
})

# Partner positions: 4 cardinal points
PARTNER_ANGLES = {
    "UNICAFES_PA":  np.pi/2,       # North
    "UNICAFES_RO":  0,             # East
    "UNICATADORES": -np.pi/2,      # South
    "CSA_BRASIL":   np.pi,         # West
}
PARTNER_RADIUS = 6.5

# Anel "Trias hub" no centro
HUB_RADIUS = 0.7

# Anel de tri-bridges
TRIBRIDGE_RADIUS = 2.3

# Anel de dual-bridges
DUALBRIDGE_RADIUS = 4.3

# Stakeholders externos (1 MBO) — projetados em arcos OUTWARD da MBO
OUTER_RADIUS_MIN = 8.2
OUTER_RADIUS_MAX = 11.0

# Cores de "território" de cada MBO (background sector tint)
PARTNER_TINT = {
    "UNICAFES_PA":  "#E8EFE3",  # Amazônia rural - verde sálvia muito claro
    "UNICAFES_RO":  "#EDE7DD",  # Amazônia rural - terra clara
    "CSA_BRASIL":   "#E5E9F0",  # Rural-urbano - azul gelo
    "UNICATADORES": "#EFE5E0",  # Urbano circular - bege rosado
}

def polar_to_xy(r, theta):
    return r * np.cos(theta), r * np.sin(theta)


# ========== Classificar stakeholders por papel posicional ==========
# Nodes a posicionar: parceiros + stakeholders conectados
# (sem camada institucional Trias — fica para o mapa "completo")
all_stk = [n for n in G.nodes if G.nodes[n].get("kind") == "stakeholder"]

# 1) tri-bridges (3 MBOs)
tri = [n for n in all_stk if G.nodes[n].get("n_partners", 0) >= 3]
# 2) dual-bridges (2 MBOs)
dual = [n for n in all_stk if G.nodes[n].get("n_partners", 0) == 2]
# 3) single (1 MBO)
single = {}  # partner_id -> list of nodes
for n in all_stk:
    if G.nodes[n].get("n_partners", 0) == 1:
        conns = G.nodes[n].get("connections", {})
        pid = list(conns.keys())[0]
        single.setdefault(pid, []).append(n)

print(f"  tri-bridges: {len(tri)}, dual-bridges: {len(dual)}, "
      f"single-partner: {sum(len(v) for v in single.values())}")

# ========== Posicionamento radial ==========
pos = {}

# Parceiros MBO nos 4 cardeais
for pid, ang in PARTNER_ANGLES.items():
    pos[pid] = polar_to_xy(PARTNER_RADIUS, ang)

# Tri-bridges no anel central, distribuídos uniformemente em círculo
n_tri = len(tri)
# Ordena por relevância (soma dos scores) para colocar os mais alinhados no topo
tri_sorted = sorted(tri, key=lambda n: -sum(G.nodes[n].get("connections", {}).values()))
for i, nid in enumerate(tri_sorted):
    # Distribuição uniforme em 360° começando no topo, horário
    ang = np.pi/2 - (i / max(n_tri, 1)) * 2 * np.pi
    # Alterna leve raio para reduzir sobreposição de labels
    r_jitter = (i % 2) * 0.25
    pos[nid] = polar_to_xy(TRIBRIDGE_RADIUS + r_jitter, ang)

# Dual-bridges: distribuídos ao longo de arcos entre os pares de MBOs
# Para evitar sobreposição, cada par recebe um leque mais amplo e 2 raios alternados
pair_to_bridges = defaultdict(list)
for nid in dual:
    pids = tuple(sorted(G.nodes[nid].get("connections", {}).keys()))
    pair_to_bridges[pids].append(nid)

for pair_key, items in pair_to_bridges.items():
    p1, p2 = pair_key
    a1 = PARTNER_ANGLES[p1]; a2 = PARTNER_ANGLES[p2]
    # ângulo médio com cuidado para wraparound
    da = (a2 - a1 + np.pi) % (2*np.pi) - np.pi
    mid_ang = a1 + da/2
    items_sorted = sorted(items,
                          key=lambda n: -sum(G.nodes[n].get("connections", {}).values()))
    n_items = len(items_sorted)
    is_diagonal = abs(abs(da) - np.pi) < 0.1
    # leque mais amplo para muitos itens; diagonais usam leque ainda maior
    if is_diagonal:
        spread = min(0.9, 0.20 * n_items)
    else:
        spread = min(0.75, 0.18 * n_items)
    for j, nid in enumerate(items_sorted):
        if n_items == 1:
            offset = 0
        else:
            offset = (j - (n_items - 1) / 2) * (spread / max(n_items - 1, 1))
        # 3 camadas de raio para reduzir overlap em pares densos
        layer = j % 3
        r_offset = (layer - 1) * 0.55
        pos[nid] = polar_to_xy(DUALBRIDGE_RADIUS + r_offset, mid_ang + offset)

# Single-partner stakeholders: em arco OUTWARD da sua MBO
# Arco aumenta proporcionalmente ao número de partners
for pid, nodes in single.items():
    if not nodes: continue
    base_ang = PARTNER_ANGLES[pid]
    # Ordena por score (mais alinhado mais perto da MBO)
    nodes_sorted = sorted(nodes,
                          key=lambda n: -G.nodes[n].get("connections", {}).get(pid, 0))
    n = len(nodes_sorted)
    # Arco span dinâmico: 50° para 5 nós, até 80° para 30+ nós
    arc_span = np.pi * min(0.50, 0.18 + n * 0.012)
    # 4 camadas radiais para reduzir overlap em arcos densos
    n_layers = 4
    for i, nid in enumerate(nodes_sorted):
        if n == 1:
            ang = base_ang
        else:
            # Distribui por camadas: camada interna pega os mais alinhados
            layer = i % n_layers
            # dentro da camada, distribui pelo arco
            in_layer_idx = i // n_layers
            n_in_layer = (n + n_layers - 1 - layer) // n_layers
            if n_in_layer <= 1:
                ang = base_ang
            else:
                ang = (base_ang - arc_span/2
                       + (in_layer_idx / (n_in_layer - 1)) * arc_span)
            # leve jitter angular para evitar grade muito rígida
            ang += ((i * 37) % 17 - 8) * 0.003
        layer = i % n_layers
        r = OUTER_RADIUS_MIN + layer * (OUTER_RADIUS_MAX - OUTER_RADIUS_MIN) / (n_layers - 1)
        pos[nid] = polar_to_xy(r, ang)


# ========== Renderização ==========
LANG = {
    "pt": {
        "title": "O ecossistema em que a Trias opera",
        "subtitle": ("Brasil · DGD 2027–2031 · 4 parceiros MBO ancorando "
                     f"{len(all_stk)} stakeholders do ecossistema"),
        "eyebrow": "TRIAS · BRASIL · DGD 2027–2031",
        "hub": "Trias",
        "ring_tri":  "Tri-bridges · conectam 3 MBOs",
        "ring_dual": "Dual-bridges · conectam 2 MBOs",
        "ring_outer": "Atores alinhados a 1 MBO",
        "legend_title": "Setor do stakeholder",
        "edge_alignment": "Alinhamento temático/territorial",
        "edge_peer": "Peer MBO (rede UNICOPAS)",
        "footer": ("Posição radial = quantos MBOs o ator alcança. "
                   "Centro = articulação compartilhada · Periferia = nicho específico"),
        "n_tri_label": f"{len(tri)} tri-bridges",
        "n_dual_label": f"{len(dual)} dual-bridges",
        "n_single_label": f"{sum(len(v) for v in single.values())} stakeholders de 1 MBO",
    },
    "en": {
        "title": "The ecosystem in which Trias operates",
        "subtitle": ("Brazil · DGD 2027–2031 · 4 MBO partners anchoring "
                     f"{len(all_stk)} ecosystem stakeholders"),
        "eyebrow": "TRIAS · BRAZIL · DGD 2027–2031",
        "hub": "Trias",
        "ring_tri":  "Tri-bridges · connect 3 MBOs",
        "ring_dual": "Dual-bridges · connect 2 MBOs",
        "ring_outer": "Actors aligned with 1 MBO",
        "legend_title": "Stakeholder sector",
        "edge_alignment": "Thematic / territorial alignment",
        "edge_peer": "MBO peer (UNICOPAS network)",
        "footer": ("Radial position = how many MBOs the actor reaches. "
                   "Center = shared articulation · Periphery = specific niche"),
        "n_tri_label": f"{len(tri)} tri-bridges",
        "n_dual_label": f"{len(dual)} dual-bridges",
        "n_single_label": f"{sum(len(v) for v in single.values())} single-MBO stakeholders",
    },
}

SECTOR_EN = {
    "Organizações sociais":     "Civil society",
    "Cooperação internacional": "International coop.",
    "Setor privado":            "Private sector",
    "Academia":                 "Academia",
    "Governo":                  "Government",
    "Outros":                   "Other",
}

def render_radial(out_path, lang="pt"):
    L = LANG[lang]
    fig, ax = plt.subplots(figsize=(22, 22), dpi=180)
    LIM = 13.0
    ax.set_xlim(-LIM, LIM)
    ax.set_ylim(-LIM, LIM)
    ax.set_aspect("equal")
    ax.axis("off")

    # === Background MBO "territories" as colored wedges ===
    # cada MBO domina um quadrante (90° em torno de seu ângulo)
    for pid, ang in PARTNER_ANGLES.items():
        wedge = Wedge((0, 0), LIM, np.degrees(ang) - 45, np.degrees(ang) + 45,
                      facecolor=PARTNER_TINT[pid], edgecolor="none",
                      alpha=0.55, zorder=0)
        ax.add_patch(wedge)

    # === Concentric guide rings (subtle) ===
    for r, lbl_key in [(HUB_RADIUS, None),
                       (TRIBRIDGE_RADIUS, "ring_tri"),
                       (DUALBRIDGE_RADIUS, "ring_dual"),
                       (PARTNER_RADIUS, None),
                       (OUTER_RADIUS_MAX + 0.5, "ring_outer")]:
        c = Circle((0, 0), r, fill=False, edgecolor="#94A3B8", alpha=0.25,
                   linewidth=0.6, linestyle="--", zorder=0.5)
        ax.add_patch(c)
    # Labels dos anéis (à direita, em pequeno) — somente os 3 informativos
    ring_label_positions = [
        (TRIBRIDGE_RADIUS, L["ring_tri"]),
        (DUALBRIDGE_RADIUS, L["ring_dual"]),
        (OUTER_RADIUS_MAX + 0.5, L["ring_outer"]),
    ]
    for r, lbl in ring_label_positions:
        ax.text(r * np.cos(np.pi/4 + 0.05), r * np.sin(np.pi/4 + 0.05),
                lbl, fontsize=7.5, color="#64748B", style="italic",
                ha="left", va="bottom", zorder=2,
                bbox=dict(facecolor="#FAFBFC", edgecolor="none", pad=1.5, alpha=0.9))

    # === Edges (alignment) ===
    for u, v, data in G.edges(data=True):
        if u not in pos or v not in pos:
            continue
        kind = data.get("kind", "alignment")
        x1, y1 = pos[u]; x2, y2 = pos[v]
        if kind == "peer":
            ax.plot([x1, x2], [y1, y2], color="#B91C1C", alpha=0.30,
                    linestyle="--", linewidth=1.4, zorder=1.5)
        elif kind == "institutional":
            continue  # camada institucional fora deste mapa
        else:
            w = data.get("weight", 4)
            # arestas curvas leves: usa FancyArrowPatch para slight curve
            ax.plot([x1, x2], [y1, y2], color="#475569", alpha=0.12,
                    linewidth=0.4 + w * 0.05, zorder=1, solid_capstyle="round")

    # === Trias central hub ===
    hub_circle = Circle((0, 0), HUB_RADIUS, facecolor="#0F172A",
                        edgecolor="white", linewidth=2.5, zorder=5)
    ax.add_patch(hub_circle)
    ax.text(0, 0, L["hub"], ha="center", va="center", fontsize=13,
            fontweight="bold", color="white", zorder=6)

    # Linhas finas do Trias para cada MBO (apoio visual)
    for pid in PARTNERS:
        px, py = pos[pid]
        ax.plot([0, px], [0, py], color="#0F172A", alpha=0.18,
                linewidth=1.2, zorder=1.2, linestyle=":")

    # === MBO nodes (4 anchors) ===
    for pid in PARTNERS:
        x, y = pos[pid]
        ax.scatter(x, y, s=2800, c=PARTNER_COLOR, marker="D",
                   edgecolors="white", linewidths=3.5, zorder=4)
        # Label da MBO bem afastado (radialmente para fora do centro)
        ang = PARTNER_ANGLES[pid]
        lx = x + 0.95 * np.cos(ang)
        ly = y + 0.95 * np.sin(ang)
        ha = "center"; va = "center"
        if abs(np.cos(ang)) > 0.5:
            ha = "left" if np.cos(ang) > 0 else "right"
        if abs(np.sin(ang)) > 0.5:
            va = "bottom" if np.sin(ang) > 0 else "top"
        ax.text(lx, ly, PARTNERS[pid]["name"], ha=ha, va=va,
                fontsize=15, fontweight="bold", color="#0F172A", zorder=6,
                bbox=dict(facecolor="white", edgecolor="#B91C1C",
                          boxstyle="round,pad=0.42", alpha=0.98, linewidth=1.5))

    # === Stakeholder nodes ===
    for nid, attrs in G.nodes(data=True):
        if nid not in pos: continue
        if attrs.get("kind") != "stakeholder":
            continue
        x, y = pos[nid]
        n_p = attrs.get("n_partners", 1)
        # tamanho cresce com n_partners
        size = 90 + n_p * 110
        color = attrs.get("color", "#999")
        ax.scatter(x, y, s=size, c=color, marker="o",
                   edgecolors="white", linewidths=1.2, zorder=3)

    # === Labels seletivos ===
    # Etiqueta apenas os atores estruturais: tri-bridges, top 1 dual por par,
    # top 2 single por MBO. Resto fica visual sem texto para reduzir ruído.
    label_set = set(tri)
    for pair_key, items in pair_to_bridges.items():
        items_sorted = sorted(items,
                              key=lambda n: -sum(G.nodes[n].get("connections", {}).values()))
        for nid in items_sorted[:1]:
            label_set.add(nid)
    for pid, nodes in single.items():
        nodes_sorted = sorted(nodes,
                              key=lambda n: -G.nodes[n].get("connections", {}).get(pid, 0))
        for nid in nodes_sorted[:2]:
            label_set.add(nid)

    # Helper: label offset radial (afastar do centro) — labels ficam FORA do nó
    def label_offset(x, y, dist=0.5):
        r = np.sqrt(x**2 + y**2)
        if r < 0.1:
            return 0, -dist
        return (x / r) * dist, (y / r) * dist

    for nid in label_set:
        if nid not in pos: continue
        x, y = pos[nid]
        name = G.nodes[nid].get("name", nid)
        # Tira parênteses para encurtar
        short = re.sub(r"\s*\([^)]+\)\s*$", "", name)
        if len(short) > 28:
            short = short[:25] + "…"
        n_p = G.nodes[nid].get("n_partners", 1)
        # Offset radial maior para tri-bridges (saem do anel central)
        dist = 0.45 if n_p >= 3 else 0.40
        dx, dy = label_offset(x, y, dist)
        # ha/va inteligentes baseados no quadrante
        ha = "left" if dx > 0.1 else "right" if dx < -0.1 else "center"
        va = "bottom" if dy > 0.1 else "top" if dy < -0.1 else "center"
        fontsize = 9.0 if n_p >= 3 else 8.5 if n_p == 2 else 7.6
        fontweight = "700" if n_p >= 3 else "600" if n_p == 2 else "500"
        color = "#0F172A" if n_p >= 2 else "#334155"
        ax.text(x + dx, y + dy, short, ha=ha, va=va,
                fontsize=fontsize, color=color, fontweight=fontweight,
                zorder=5,
                bbox=dict(facecolor="white", edgecolor="none",
                          boxstyle="round,pad=0.18", alpha=0.92))

    # === Header ===
    fig.text(0.06, 0.965, L["title"], fontsize=28, fontweight="bold",
             color="#0F172A", ha="left", va="top")
    fig.text(0.06, 0.935, L["subtitle"], fontsize=12.5,
             color="#334155", ha="left", va="top", style="italic")
    fig.text(0.06, 0.915, L["eyebrow"], fontsize=9,
             color="#64748B", ha="left", va="top", family="monospace")

    # === Legenda de setor (canto inferior esquerdo) ===
    sector_counts = Counter()
    for n in all_stk:
        sector_counts[G.nodes[n].get("sector", "Outros")] += 1
    handles = []
    for sector, color in SECTOR_COLOR.items():
        label = SECTOR_EN.get(sector, sector) if lang == "en" else sector
        n = sector_counts.get(sector, 0)
        if n == 0: continue
        handles.append(Line2D([0], [0], marker="o", color="w",
                              markerfacecolor=color, markeredgecolor="white",
                              markeredgewidth=1.0, markersize=11,
                              label=f"{label} ({n})"))
    handles.append(Line2D([0], [0], marker="D", color="w",
                          markerfacecolor=PARTNER_COLOR, markeredgecolor="white",
                          markeredgewidth=1.2, markersize=14,
                          label=("Parceiro MBO (4)" if lang == "pt"
                                 else "MBO partner (4)")))
    handles.append(Line2D([0], [0], marker="o", color="w",
                          markerfacecolor="#0F172A", markeredgecolor="white",
                          markeredgewidth=1.2, markersize=12,
                          label=("Trias (hub articulador)" if lang == "pt"
                                 else "Trias (articulating hub)")))
    leg = ax.legend(handles=handles, loc="lower left",
                    bbox_to_anchor=(-0.05, -0.08),
                    frameon=True, fontsize=9.5, labelspacing=0.8,
                    handletextpad=0.9, borderpad=1.0,
                    facecolor="white", edgecolor="#CBD5E1",
                    title=L["legend_title"], title_fontsize=10.5)
    leg.get_title().set_fontweight("700")
    leg.get_title().set_color("#0F172A")
    leg.get_frame().set_linewidth(0.9)

    # === Legenda de arestas (canto inferior direito) ===
    edge_handles = [
        Line2D([0], [0], color="#475569", alpha=0.35, linewidth=1.4,
               label=L["edge_alignment"]),
        Line2D([0], [0], color="#B91C1C", alpha=0.5, linewidth=1.6,
               linestyle="--", label=L["edge_peer"]),
        Line2D([0], [0], color="#0F172A", alpha=0.5, linewidth=1.2,
               linestyle=":", label=("Eixo Trias → MBO" if lang == "pt"
                                    else "Trias → MBO axis")),
    ]
    leg2 = ax.legend(handles=edge_handles, loc="lower right",
                     bbox_to_anchor=(1.05, -0.08),
                     frameon=True, fontsize=9.5, labelspacing=0.7,
                     handletextpad=0.8, borderpad=1.0,
                     facecolor="white", edgecolor="#CBD5E1",
                     title=("Tipos de conexão" if lang == "pt"
                            else "Connection types"),
                     title_fontsize=10.5)
    leg2.get_title().set_fontweight("700")
    leg2.get_title().set_color("#0F172A")
    leg2.get_frame().set_linewidth(0.9)
    ax.add_artist(leg)

    # === Counts annotation (top right) ===
    counts_txt = (f"{L['n_tri_label']}\n"
                  f"{L['n_dual_label']}\n"
                  f"{L['n_single_label']}")
    fig.text(0.94, 0.93, counts_txt, fontsize=10.5, color="#0F172A",
             ha="right", va="top", family="monospace",
             bbox=dict(facecolor="white", edgecolor="#CBD5E1",
                       boxstyle="round,pad=0.55", linewidth=0.9))

    # === Footer ===
    fig.text(0.5, 0.035, L["footer"], fontsize=9.5,
             color="#64748B", ha="center", va="bottom", style="italic")

    plt.subplots_adjust(left=0.02, right=0.98, top=0.90, bottom=0.07)
    plt.savefig(out_path, dpi=180, bbox_inches="tight", facecolor="#FAFBFC",
                pad_inches=0.5)
    plt.close()
    print(f"PNG salvo: {out_path}")


# ========== Gerar PNGs ==========
render_radial(OUT_NET / "ecossistema_radial.png", lang="pt")
render_radial(OUT_NET / "ecosystem_radial_en.png", lang="en")

# Cópia para docs/ (GitHub Pages)
import shutil
shutil.copy(OUT_NET / "ecossistema_radial.png", OUT_DOCS / "ecossistema_radial.png")
print(f"Copiado para {OUT_DOCS / 'ecossistema_radial.png'}")
print("Done.")
