"""
Render lado a lado: alluvial (setor → papel Trias → parceiro MBO) +
chord/arc (bridges compartilhados entre os 4 parceiros).
"""
from openpyxl import load_workbook
import networkx as nx
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch
from matplotlib.path import Path as MplPath
import re
import numpy as np
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path("/home/user/trias_ToC")
SCRIPT = ROOT / "outputs" / "network" / "build_network.py"
OUT_NET = ROOT / "outputs" / "network"

# Reusa logica do build_network.py
src = SCRIPT.read_text(encoding="utf-8")
marker = "# ========== 7. JSON do grafo para D3 =========="
truncated = src.split(marker)[0]
ns = {"__name__": "build_net_partial"}
exec(truncated, ns)
PARTNERS = ns["PARTNERS"]
SECTOR_COLOR = ns["SECTOR_COLOR"]
PARTNER_COLOR = ns["PARTNER_COLOR"]
INSTITUTIONAL_COLOR = ns["INSTITUTIONAL_COLOR"]
G = ns["G"]
results = ns["results"]
TRIAS_MODE_RULES = ns["TRIAS_MODE_RULES"]
print(f"Grafo: {G.number_of_nodes()} nós, {G.number_of_edges()} arestas")

# Mapa modo → papel Trias
MODE_TO_ROLE = {}
for _, mode, role, _ in TRIAS_MODE_RULES:
    if mode not in MODE_TO_ROLE:
        MODE_TO_ROLE[mode] = role
ROLES_ORDER = ["Process Facilitator", "Thematic Advisor",
               "Peer-to-Peer Facilitator", "Bridge Builder", "Financer"]
# Normalizar role names (alguns têm "+")
def normalize_role(r):
    # "Financer + Bridge Builder" → "Financer + Bridge Builder", manter combinado
    return r
# Vamos usar o ROLE primário (primeiro mencionado) p/ não inflar
def primary_role(role_str):
    for canonical in ROLES_ORDER:
        if canonical in role_str:
            return canonical
    return role_str

# ===========================================================
# DADOS PARA ALLUVIAL
# ===========================================================
# Para cada (stakeholder, partner), conta fluxo: setor → papel_primario → partner
flows_s_r = Counter()   # (setor, role)
flows_r_p = Counter()   # (role, partner_name)
sectors_set = list(SECTOR_COLOR.keys())
partners_set = [PARTNERS[pid]["name"] for pid in PARTNERS]

for r in results:
    stk = r["stk"]
    sector = stk["sector"] or "Others"
    modes = r["trias_modes"]
    roles = list({primary_role(m["trias_role"]) for m in modes if primary_role(m["trias_role"]) in ROLES_ORDER})
    if not roles:
        continue
    for pid in r["scores"]:
        pname = PARTNERS[pid]["name"]
        # cada conexão MBO contribui 1/N por role (para que a soma de fluxos = total conexões)
        for role in roles:
            w = 1.0 / len(roles)
            flows_s_r[(sector, role)] += w
            flows_r_p[(role, pname)] += w

# Totais por categoria
sector_totals = Counter()
role_totals_left = Counter()  # entrada (de setores)
role_totals_right = Counter() # saída (para parceiros)
partner_totals = Counter()
for (s, r), v in flows_s_r.items():
    sector_totals[s] += v
    role_totals_left[r] += v
for (r, p), v in flows_r_p.items():
    role_totals_right[r] += v
    partner_totals[p] += v

# ===========================================================
# DADOS PARA CHORD (bridges entre pares de parceiros)
# ===========================================================
bridge_pairs = Counter()
for r in results:
    pids = list(r["scores"].keys())
    if len(pids) >= 2:
        for i in range(len(pids)):
            for j in range(i+1, len(pids)):
                a, b = sorted([pids[i], pids[j]])
                bridge_pairs[(a, b)] += 1

# Auto-loops = conexões exclusivas a 1 parceiro (não bridge)
exclusive = Counter()
for r in results:
    if len(r["scores"]) == 1:
        pid = list(r["scores"].keys())[0]
        exclusive[pid] += 1

# ===========================================================
# RENDER
# ===========================================================
plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "font.size": 9,
})

fig = plt.figure(figsize=(28, 13), dpi=180, facecolor="white")
# Layout: alluvial à esquerda (mais largo), chord à direita
gs = fig.add_gridspec(1, 2, width_ratios=[1.6, 1.0], left=0.04, right=0.97,
                       top=0.84, bottom=0.10, wspace=0.10)
ax_a = fig.add_subplot(gs[0, 0])
ax_c = fig.add_subplot(gs[0, 1])

# ------------------------------ ALLUVIAL ------------------------------
ax_a.set_xlim(0, 10)
ax_a.set_ylim(0, 100)
ax_a.axis("off")
ax_a.set_title("Como o capital institucional flui até as MBOs",
               fontsize=14, fontweight="bold", color="#0F172A", loc="left", pad=12)
ax_a.text(0, 96, "Setor do stakeholder → papel Trias mobilizado → parceiro MBO",
          fontsize=9, color="#64748B", style="italic")

# Posições x das 3 colunas
COL_X = {"sector": (0.5, 1.6), "role": (4.0, 5.1), "partner": (8.6, 9.7)}
PAD = 1.8  # padding vertical entre blocos
# Escala total: soma fluxos = altura total
total_flow = sum(flows_s_r.values())
USABLE_H = 88.0  # altura útil (0..88, com header em cima)
H_SCALE = USABLE_H / total_flow

# Função para alocar bins verticalmente
def alloc_bins(items, total_flow):
    """items: lista de (key, value). Retorna {key: (y_bottom, y_top)}."""
    y = 92  # de cima pra baixo
    result = {}
    n_items = len([k for k, v in items if v > 0])
    n_gaps = max(0, n_items - 1)
    available = USABLE_H - n_gaps * PAD
    h_unit = available / total_flow if total_flow > 0 else 0
    for k, v in items:
        if v <= 0: continue
        h = v * h_unit
        result[k] = (y - h, y)
        y -= h + PAD
    return result

# Alocação left (setores)
sector_items = [(s, sector_totals.get(s, 0)) for s in sectors_set]
sector_items.sort(key=lambda x: -x[1])
sector_bins = alloc_bins(sector_items, total_flow)

# Alocação middle (roles) — usar role_totals_left
role_items = [(r, role_totals_left.get(r, 0)) for r in ROLES_ORDER]
role_items.sort(key=lambda x: -x[1])
role_bins = alloc_bins(role_items, total_flow)

# Alocação right (partners)
partner_items = [(p, partner_totals.get(p, 0)) for p in partners_set]
partner_items.sort(key=lambda x: -x[1])
partner_bins = alloc_bins(partner_items, total_flow)

# Cores para roles e partners
ROLE_COLOR = {
    "Process Facilitator":   "#7A6E5C",  # marrom claro
    "Thematic Advisor":      "#4D7A6A",  # verde
    "Peer-to-Peer Facilitator": "#6B5882",  # roxo
    "Bridge Builder":        "#3B5380",  # azul
    "Financer":              "#806237",  # bronze
}
PARTNER_COLORS = {p: PARTNER_COLOR for p in partners_set}

# Desenha blocos
def draw_block(x0, x1, y0, y1, color, label, count, side="left"):
    rect = plt.Rectangle((x0, y0), x1 - x0, y1 - y0,
                         facecolor=color, edgecolor="white", linewidth=1.5, zorder=3)
    ax_a.add_patch(rect)
    ycen = (y0 + y1) / 2
    if side == "left":
        ax_a.text(x0 - 0.15, ycen, f"{label}", ha="right", va="center",
                  fontsize=9.5, color="#0F172A", fontweight="500", zorder=4)
        ax_a.text(x0 - 0.15, ycen - 1.6, f"n={count:.0f}", ha="right", va="top",
                  fontsize=7.5, color="#64748B", zorder=4)
    elif side == "right":
        ax_a.text(x1 + 0.15, ycen, f"{label}", ha="left", va="center",
                  fontsize=9.5, color="#0F172A", fontweight="500", zorder=4)
        ax_a.text(x1 + 0.15, ycen - 1.6, f"n={count:.0f}", ha="left", va="top",
                  fontsize=7.5, color="#64748B", zorder=4)
    else:
        # middle: label vertical à direita
        ax_a.text(x1 + 0.15, ycen, f"{label}", ha="left", va="center",
                  fontsize=9.5, color="#0F172A", fontweight="500", zorder=4)
        ax_a.text(x1 + 0.15, ycen - 1.6, f"n={count:.0f}", ha="left", va="top",
                  fontsize=7.5, color="#64748B", zorder=4)

# Desenha setores (left)
for s, (y0, y1) in sector_bins.items():
    draw_block(COL_X["sector"][0], COL_X["sector"][1], y0, y1,
               SECTOR_COLOR[s], s, sector_totals[s], side="left")

# Desenha roles (middle)
for r, (y0, y1) in role_bins.items():
    draw_block(COL_X["role"][0], COL_X["role"][1], y0, y1,
               ROLE_COLOR[r], r, role_totals_left[r], side="middle")

# Desenha partners (right)
for p, (y0, y1) in partner_bins.items():
    draw_block(COL_X["partner"][0], COL_X["partner"][1], y0, y1,
               PARTNER_COLOR, p, partner_totals[p], side="right")

# Desenha bandas de fluxo (bezier)
def draw_flow(x0, x1, y0_top, y0_bot, y1_top, y1_bot, color, alpha=0.4):
    # ribbon: lados curvos
    xm = (x0 + x1) / 2
    verts_top = [(x0, y0_top), (xm, y0_top), (xm, y1_top), (x1, y1_top)]
    verts_bot = [(x1, y1_bot), (xm, y1_bot), (xm, y0_bot), (x0, y0_bot)]
    # path bezier cubic
    codes_top = [MplPath.MOVETO, MplPath.CURVE4, MplPath.CURVE4, MplPath.CURVE4]
    codes_bot = [MplPath.LINETO, MplPath.CURVE4, MplPath.CURVE4, MplPath.CURVE4]
    codes = codes_top + codes_bot + [MplPath.CLOSEPOLY]
    verts = verts_top + verts_bot + [(x0, y0_top)]
    path = MplPath(verts, codes)
    patch = mpatches.PathPatch(path, facecolor=color, edgecolor="none",
                                alpha=alpha, zorder=2)
    ax_a.add_patch(patch)

# Alocar fluxos dentro de cada bloco (cumulative)
# Esquerda: para cada setor, fluxos saindo (para cada role)
left_cursors = {s: sector_bins[s][1] for s in sector_bins}  # do topo do bloco
mid_cursors_in = {r: role_bins[r][1] for r in role_bins}
mid_cursors_out = {r: role_bins[r][1] for r in role_bins}
right_cursors = {p: partner_bins[p][1] for p in partner_bins}

# Ordenar fluxos por destino para reduzir crossings
flows_sr_sorted = sorted(flows_s_r.items(), key=lambda x: (
    list(sector_bins.keys()).index(x[0][0]) if x[0][0] in sector_bins else 99,
    list(role_bins.keys()).index(x[0][1]) if x[0][1] in role_bins else 99
))
for (s, r), v in flows_sr_sorted:
    if v <= 0 or s not in sector_bins or r not in role_bins: continue
    h = v * H_SCALE
    y0_top = left_cursors[s]
    y0_bot = y0_top - h
    y1_top = mid_cursors_in[r]
    y1_bot = y1_top - h
    draw_flow(COL_X["sector"][1], COL_X["role"][0],
              y0_top, y0_bot, y1_top, y1_bot,
              SECTOR_COLOR[s], alpha=0.42)
    left_cursors[s] = y0_bot
    mid_cursors_in[r] = y1_bot

# Direita: fluxos role → partner (cor pela role)
flows_rp_sorted = sorted(flows_r_p.items(), key=lambda x: (
    list(role_bins.keys()).index(x[0][0]) if x[0][0] in role_bins else 99,
    list(partner_bins.keys()).index(x[0][1]) if x[0][1] in partner_bins else 99
))
for (r, p), v in flows_rp_sorted:
    if v <= 0 or r not in role_bins or p not in partner_bins: continue
    h = v * H_SCALE
    y0_top = mid_cursors_out[r]
    y0_bot = y0_top - h
    y1_top = right_cursors[p]
    y1_bot = y1_top - h
    draw_flow(COL_X["role"][1], COL_X["partner"][0],
              y0_top, y0_bot, y1_top, y1_bot,
              ROLE_COLOR[r], alpha=0.42)
    mid_cursors_out[r] = y0_bot
    right_cursors[p] = y1_bot

# Header das colunas
ax_a.text(COL_X["sector"][0] - 0.15, 95, "SETOR", ha="right", va="bottom",
          fontsize=8, color="#94A3B8", fontweight="bold")
ax_a.text((COL_X["role"][0] + COL_X["role"][1])/2, 95,
          "PAPEL TRIAS MOBILIZADO", ha="center", va="bottom",
          fontsize=8, color="#94A3B8", fontweight="bold")
ax_a.text(COL_X["partner"][1] + 0.15, 95, "PARCEIRO MBO", ha="left", va="bottom",
          fontsize=8, color="#94A3B8", fontweight="bold")

# ------------------------------ CHORD/ARC ------------------------------
ax_c.set_xlim(-1.6, 1.6)
ax_c.set_ylim(-1.6, 1.6)
ax_c.set_aspect("equal")
ax_c.axis("off")
ax_c.set_title("Bridges compartilhados entre os 4 parceiros MBO",
               fontsize=14, fontweight="bold", color="#0F172A", loc="left", pad=12)
ax_c.text(-1.6, 1.50, "Espessura do arco = nº de stakeholders bridge entre o par",
          fontsize=9, color="#64748B", style="italic")

# 4 nós nas extremidades de um círculo
partners_order = ["UNICAFES_PA", "UNICAFES_RO", "CSA_BRASIL", "UNICATADORES"]
# angles: PA top-left, RO top-right, CSA bottom-left, UNICAT bottom-right
angles = {
    "UNICAFES_PA":   135,
    "UNICAFES_RO":   45,
    "CSA_BRASIL":    225,
    "UNICATADORES":  315,
}
R = 1.0
node_pos = {pid: (R * np.cos(np.radians(angles[pid])),
                  R * np.sin(np.radians(angles[pid])))
            for pid in partners_order}

# Desenha arcos (chord-style com bezier passando pelo centro)
max_bridges = max(bridge_pairs.values()) if bridge_pairs else 1
for (a, b), n in bridge_pairs.items():
    if n == 0: continue
    p1 = node_pos[a]; p2 = node_pos[b]
    # bezier passando perto do centro
    cp = (0, 0)
    verts = [p1, cp, p2]
    codes = [MplPath.MOVETO, MplPath.CURVE3, MplPath.CURVE3]
    path = MplPath(verts, codes)
    # espessura proporcional ao número de bridges
    width = 0.5 + (n / max_bridges) * 18
    alpha = 0.25 + 0.55 * (n / max_bridges)
    patch = mpatches.PathPatch(path, facecolor="none", edgecolor="#B91C1C",
                                linewidth=width, alpha=alpha, zorder=2,
                                capstyle="round")
    ax_c.add_patch(patch)
    # rótulo no meio
    mid_x = (p1[0] + p2[0]) / 4  # mais perto do centro
    mid_y = (p1[1] + p2[1]) / 4
    ax_c.text(mid_x, mid_y, str(n), fontsize=10.5, color="#B91C1C",
              fontweight="bold", ha="center", va="center", zorder=4,
              bbox=dict(facecolor="white", edgecolor="none",
                        boxstyle="circle,pad=0.25", alpha=0.95))

# Desenha nós (losangos vermelhos como no mapa)
for pid in partners_order:
    x, y = node_pos[pid]
    ax_c.scatter(x, y, s=1800, c=PARTNER_COLOR, marker="D",
                 edgecolors="white", linewidths=2.0, zorder=5)
    # label fora do círculo
    label_offset = 0.25
    dx = label_offset * np.cos(np.radians(angles[pid]))
    dy = label_offset * np.sin(np.radians(angles[pid]))
    # Halign por quadrante
    ha = "right" if x < 0 else "left"
    va = "top" if y < 0 else "bottom"
    name = PARTNERS[pid]["name"]
    n_excl = exclusive.get(pid, 0)
    ax_c.text(x + dx, y + dy, name, ha=ha, va=va, fontsize=11,
              fontweight="bold", color="#0F172A", zorder=6)
    ax_c.text(x + dx, y + dy - 0.10 if va == "top" else y + dy + 0.10,
              f"{n_excl} exclusivos · {sum(1 for r in results if pid in r['scores'])} total",
              ha=ha, va=va, fontsize=8, color="#64748B",
              fontweight="normal", zorder=6)

# Legenda explicativa abaixo do chord
ax_c.text(0, -1.45, "Leitura · arcos espessos = parceiros com ecossistemas que se sobrepõem",
          ha="center", va="top", fontsize=9, color="#475569", style="italic")
ax_c.text(0, -1.55, "Ausência de arcos = ecossistemas separados (caso UNICATADORES vs Amazônia)",
          ha="center", va="top", fontsize=8.5, color="#64748B")

# ===========================================================
# HEADER GERAL
# ===========================================================
fig.text(0.04, 0.96, "Ecossistema de stakeholders",
         fontsize=26, fontweight="bold", color="#0F172A", ha="left", va="top")
fig.text(0.04, 0.925, "Fluxos institucionais e sobreposição entre parceiros MBO",
         fontsize=12, color="#334155", style="italic", ha="left", va="top")
fig.text(0.04, 0.905, "TRIAS · BRASIL · DGD 2027–2031",
         fontsize=8.5, color="#64748B", ha="left", va="top", family="monospace")
fig.text(0.97, 0.96, "Versão 4.0 · maio 2026",
         fontsize=8.5, color="#64748B", ha="right", va="top", family="monospace")

# Salva
out_path = OUT_NET / "alluvial_chord.png"
plt.savefig(out_path, dpi=180, facecolor="white", bbox_inches="tight",
            pad_inches=0.3)
plt.close()
print(f"PNG salvo: {out_path}")
print(f"  Alluvial: {len(flows_s_r)} fluxos S→R, {len(flows_r_p)} fluxos R→P")
print(f"  Chord: {len(bridge_pairs)} pares com bridges")
for (a,b), n in sorted(bridge_pairs.items(), key=lambda x: -x[1]):
    print(f"    {PARTNERS[a]['name']} ↔ {PARTNERS[b]['name']}: {n}")
