"""
Renders standalone (English):
  - alluvial.png  — Stakeholder sector → Trias role → MBO partner
  - bridges.png   — Shared bridges between MBO partners (chord/arc)

Distinct palettes for sectors (cool/neutral) and Trias roles (warm/violet).
"""
from openpyxl import load_workbook
import networkx as nx
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.path import Path as MplPath
import numpy as np
from collections import Counter
from pathlib import Path

ROOT = Path("/home/user/trias_ToC")
SCRIPT = ROOT / "outputs" / "network" / "build_network.py"
OUT_NET = ROOT / "outputs" / "network"

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
TRIAS_MODE_RULES = ns["TRIAS_MODE_RULES"]
print(f"Graph: {G.number_of_nodes()} nodes, {G.number_of_edges()} edges")

# PT → EN translation for sector labels in the chart
SECTOR_LABEL_EN = {
    "Organizações sociais":     "Civil society organizations",
    "Cooperação internacional": "International cooperation",
    "Setor privado":            "Private sector",
    "Academia":                 "Academia",
    "Governo":                  "Government",
    "Outros":                   "Others",
}

# Distinct palette for Trias roles (warm + violet, no overlap with sectors)
ROLE_COLOR = {
    "Process Facilitator":      "#F2A93B",  # amber
    "Thematic Advisor":         "#D97757",  # terracotta
    "Peer-to-Peer Facilitator": "#C25E91",  # rose-mauve
    "Bridge Builder":           "#6B4FB8",  # deep violet
    "Financer":                 "#4A1F5E",  # dark plum
}
ROLES_ORDER = list(ROLE_COLOR.keys())

def primary_role(role_str):
    """Mapeia role string (ex.: 'Financer + Bridge Builder') ao papel canônico
    primário. Financer tem prioridade — é o papel mais distintivo do programa."""
    if "Financer" in role_str:
        return "Financer"
    for canonical in ROLES_ORDER:
        if canonical in role_str:
            return canonical
    return role_str

# ===========================================================
# DATA: alluvial flows
# ===========================================================
flows_s_r = Counter()
flows_r_p = Counter()
sectors_set = list(SECTOR_COLOR.keys())
partners_set = [PARTNERS[pid]["name"] for pid in PARTNERS]

for r in results:
    stk = r["stk"]
    sector = stk["sector"] or "Others"
    modes = r["trias_modes"]
    roles = list({primary_role(m["trias_role"]) for m in modes
                  if primary_role(m["trias_role"]) in ROLES_ORDER})
    if not roles: continue
    for pid in r["scores"]:
        pname = PARTNERS[pid]["name"]
        for role in roles:
            w = 1.0 / len(roles)
            flows_s_r[(sector, role)] += w
            flows_r_p[(role, pname)] += w

sector_totals = Counter()
role_totals_left = Counter()
role_totals_right = Counter()
partner_totals = Counter()
for (s, r), v in flows_s_r.items():
    sector_totals[s] += v
    role_totals_left[r] += v
for (r, p), v in flows_r_p.items():
    role_totals_right[r] += v
    partner_totals[p] += v

# ===========================================================
# DATA: bridges chord
# ===========================================================
bridge_pairs = Counter()
for r in results:
    pids = list(r["scores"].keys())
    if len(pids) >= 2:
        for i in range(len(pids)):
            for j in range(i+1, len(pids)):
                a, b = sorted([pids[i], pids[j]])
                bridge_pairs[(a, b)] += 1
exclusive = Counter()
for r in results:
    if len(r["scores"]) == 1:
        exclusive[list(r["scores"].keys())[0]] += 1

plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9})

# ===========================================================
# CHART 1 — ALLUVIAL
# ===========================================================
def render_alluvial(out_path):
    fig = plt.figure(figsize=(17, 11), dpi=180, facecolor="white")
    ax = fig.add_subplot(111)
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 100)
    ax.axis("off")

    # Header
    fig.text(0.05, 0.96, "How institutional capital flows to MBOs",
             fontsize=22, fontweight="bold", color="#0F172A", ha="left", va="top")
    fig.text(0.05, 0.925,
             "Stakeholder sector → Trias role mobilized → MBO partner",
             fontsize=11, color="#334155", style="italic", ha="left", va="top")
    fig.text(0.05, 0.905, "TRIAS · BRAZIL · DGD 2027–2031",
             fontsize=8.5, color="#64748B", ha="left", va="top", family="monospace")
    fig.text(0.97, 0.96, "Version 4.0 · May 2026",
             fontsize=8.5, color="#64748B", ha="right", va="top", family="monospace")

    COL_X = {"sector": (0.5, 1.6), "role": (4.0, 5.1), "partner": (8.6, 9.7)}
    PAD = 1.8
    total_flow = sum(flows_s_r.values())
    USABLE_H = 86.0
    H_SCALE = USABLE_H / total_flow

    def alloc_bins(items, total_flow):
        y = 90
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

    sector_items = sorted([(s, sector_totals.get(s, 0)) for s in sectors_set], key=lambda x: -x[1])
    sector_bins = alloc_bins(sector_items, total_flow)
    role_items = sorted([(r, role_totals_left.get(r, 0)) for r in ROLES_ORDER], key=lambda x: -x[1])
    role_bins = alloc_bins(role_items, total_flow)
    partner_items = sorted([(p, partner_totals.get(p, 0)) for p in partners_set], key=lambda x: -x[1])
    partner_bins = alloc_bins(partner_items, total_flow)

    def draw_block(x0, x1, y0, y1, color, label, count, side="left"):
        rect = plt.Rectangle((x0, y0), x1 - x0, y1 - y0,
                             facecolor=color, edgecolor="white", linewidth=1.5, zorder=3)
        ax.add_patch(rect)
        ycen = (y0 + y1) / 2
        if side == "left":
            ax.text(x0 - 0.15, ycen, f"{label}", ha="right", va="center",
                    fontsize=10.5, color="#0F172A", fontweight="500", zorder=4)
            ax.text(x0 - 0.15, ycen - 1.8, f"n={count:.0f}", ha="right", va="top",
                    fontsize=8, color="#64748B", zorder=4)
        elif side == "right":
            ax.text(x1 + 0.15, ycen, f"{label}", ha="left", va="center",
                    fontsize=10.5, color="#0F172A", fontweight="500", zorder=4)
            ax.text(x1 + 0.15, ycen - 1.8, f"n={count:.0f}", ha="left", va="top",
                    fontsize=8, color="#64748B", zorder=4)
        else:
            ax.text(x1 + 0.15, ycen, f"{label}", ha="left", va="center",
                    fontsize=10.5, color="#0F172A", fontweight="500", zorder=4)
            ax.text(x1 + 0.15, ycen - 1.8, f"n={count:.0f}", ha="left", va="top",
                    fontsize=8, color="#64748B", zorder=4)

    for s, (y0, y1) in sector_bins.items():
        draw_block(COL_X["sector"][0], COL_X["sector"][1], y0, y1,
                   SECTOR_COLOR[s], SECTOR_LABEL_EN.get(s, s), sector_totals[s], side="left")
    for r, (y0, y1) in role_bins.items():
        draw_block(COL_X["role"][0], COL_X["role"][1], y0, y1,
                   ROLE_COLOR[r], r, role_totals_left[r], side="middle")
    for p, (y0, y1) in partner_bins.items():
        draw_block(COL_X["partner"][0], COL_X["partner"][1], y0, y1,
                   PARTNER_COLOR, p, partner_totals[p], side="right")

    def draw_flow(x0, x1, y0_top, y0_bot, y1_top, y1_bot, color, alpha=0.4):
        xm = (x0 + x1) / 2
        verts_top = [(x0, y0_top), (xm, y0_top), (xm, y1_top), (x1, y1_top)]
        verts_bot = [(x1, y1_bot), (xm, y1_bot), (xm, y0_bot), (x0, y0_bot)]
        codes = ([MplPath.MOVETO, MplPath.CURVE4, MplPath.CURVE4, MplPath.CURVE4]
                 + [MplPath.LINETO, MplPath.CURVE4, MplPath.CURVE4, MplPath.CURVE4]
                 + [MplPath.CLOSEPOLY])
        verts = verts_top + verts_bot + [(x0, y0_top)]
        path = MplPath(verts, codes)
        patch = mpatches.PathPatch(path, facecolor=color, edgecolor="none",
                                    alpha=alpha, zorder=2)
        ax.add_patch(patch)

    left_cursors = {s: sector_bins[s][1] for s in sector_bins}
    mid_cursors_in = {r: role_bins[r][1] for r in role_bins}
    mid_cursors_out = {r: role_bins[r][1] for r in role_bins}
    right_cursors = {p: partner_bins[p][1] for p in partner_bins}

    flows_sr_sorted = sorted(flows_s_r.items(), key=lambda x: (
        list(sector_bins.keys()).index(x[0][0]) if x[0][0] in sector_bins else 99,
        list(role_bins.keys()).index(x[0][1]) if x[0][1] in role_bins else 99
    ))
    for (s, r), v in flows_sr_sorted:
        if v <= 0 or s not in sector_bins or r not in role_bins: continue
        h = v * H_SCALE
        y0_top = left_cursors[s]; y0_bot = y0_top - h
        y1_top = mid_cursors_in[r]; y1_bot = y1_top - h
        draw_flow(COL_X["sector"][1], COL_X["role"][0],
                  y0_top, y0_bot, y1_top, y1_bot,
                  SECTOR_COLOR[s], alpha=0.42)
        left_cursors[s] = y0_bot
        mid_cursors_in[r] = y1_bot

    flows_rp_sorted = sorted(flows_r_p.items(), key=lambda x: (
        list(role_bins.keys()).index(x[0][0]) if x[0][0] in role_bins else 99,
        list(partner_bins.keys()).index(x[0][1]) if x[0][1] in partner_bins else 99
    ))
    for (r, p), v in flows_rp_sorted:
        if v <= 0 or r not in role_bins or p not in partner_bins: continue
        h = v * H_SCALE
        y0_top = mid_cursors_out[r]; y0_bot = y0_top - h
        y1_top = right_cursors[p]; y1_bot = y1_top - h
        draw_flow(COL_X["role"][1], COL_X["partner"][0],
                  y0_top, y0_bot, y1_top, y1_bot,
                  ROLE_COLOR[r], alpha=0.42)
        mid_cursors_out[r] = y0_bot
        right_cursors[p] = y1_bot

    # Column headers
    ax.text(COL_X["sector"][0] - 0.15, 94, "SECTOR", ha="right", va="bottom",
            fontsize=8.5, color="#94A3B8", fontweight="bold")
    ax.text((COL_X["role"][0] + COL_X["role"][1])/2, 94,
            "TRIAS ROLE MOBILIZED", ha="center", va="bottom",
            fontsize=8.5, color="#94A3B8", fontweight="bold")
    ax.text(COL_X["partner"][1] + 0.15, 94, "MBO PARTNER", ha="left", va="bottom",
            fontsize=8.5, color="#94A3B8", fontweight="bold")

    plt.savefig(out_path, dpi=180, facecolor="white", bbox_inches="tight", pad_inches=0.3)
    plt.close()
    print(f"PNG saved: {out_path}")

# ===========================================================
# CHART 2 — BRIDGES (chord/arc)
# ===========================================================
def render_bridges(out_path):
    fig = plt.figure(figsize=(13, 11), dpi=180, facecolor="white")
    ax = fig.add_subplot(111)
    ax.set_xlim(-1.7, 1.7)
    ax.set_ylim(-1.7, 1.7)
    ax.set_aspect("equal")
    ax.axis("off")

    # Header
    fig.text(0.06, 0.96, "Shared bridges across the 4 MBO partners",
             fontsize=22, fontweight="bold", color="#0F172A", ha="left", va="top")
    fig.text(0.06, 0.925,
             "Arc thickness = number of stakeholders that connect the pair",
             fontsize=11, color="#334155", style="italic", ha="left", va="top")
    fig.text(0.06, 0.905, "TRIAS · BRAZIL · DGD 2027–2031",
             fontsize=8.5, color="#64748B", ha="left", va="top", family="monospace")
    fig.text(0.94, 0.96, "Version 4.0 · May 2026",
             fontsize=8.5, color="#64748B", ha="right", va="top", family="monospace")

    partners_order = ["UNICAFES_PA", "UNICAFES_RO", "CSA_BRASIL", "UNICATADORES"]
    angles = {"UNICAFES_PA": 135, "UNICAFES_RO": 45,
              "CSA_BRASIL": 225, "UNICATADORES": 315}
    R = 1.0
    node_pos = {pid: (R * np.cos(np.radians(angles[pid])),
                      R * np.sin(np.radians(angles[pid])))
                for pid in partners_order}

    max_bridges = max(bridge_pairs.values()) if bridge_pairs else 1
    for (a, b), n in bridge_pairs.items():
        if n == 0: continue
        p1 = node_pos[a]; p2 = node_pos[b]
        cp = (0, 0)
        verts = [p1, cp, p2]
        codes = [MplPath.MOVETO, MplPath.CURVE3, MplPath.CURVE3]
        path = MplPath(verts, codes)
        width = 0.5 + (n / max_bridges) * 22
        alpha = 0.30 + 0.55 * (n / max_bridges)
        patch = mpatches.PathPatch(path, facecolor="none", edgecolor="#B91C1C",
                                    linewidth=width, alpha=alpha, zorder=2,
                                    capstyle="round")
        ax.add_patch(patch)
        # label at midpoint near center
        mid_x = (p1[0] + p2[0]) / 4
        mid_y = (p1[1] + p2[1]) / 4
        ax.text(mid_x, mid_y, str(n), fontsize=12, color="#B91C1C",
                fontweight="bold", ha="center", va="center", zorder=4,
                bbox=dict(facecolor="white", edgecolor="#B91C1C",
                          boxstyle="circle,pad=0.30", linewidth=0.8, alpha=0.96))

    # Partners as red diamonds with labels
    for pid in partners_order:
        x, y = node_pos[pid]
        ax.scatter(x, y, s=2200, c=PARTNER_COLOR, marker="D",
                   edgecolors="white", linewidths=2.0, zorder=5)
        # Position label outside the circle
        label_offset_r = 1.40
        lx = label_offset_r * np.cos(np.radians(angles[pid]))
        ly = label_offset_r * np.sin(np.radians(angles[pid]))
        ha = "right" if lx < 0 else "left"
        va = "top" if ly < 0 else "bottom"
        name = PARTNERS[pid]["name"]
        n_excl = exclusive.get(pid, 0)
        n_total = sum(1 for r in results if pid in r["scores"])
        ax.text(lx, ly, name, ha=ha, va=va, fontsize=12.5,
                fontweight="bold", color="#0F172A", zorder=6)
        dy_meta = -0.13 if va == "top" else 0.13
        ax.text(lx, ly + dy_meta, f"{n_excl} exclusive · {n_total} total",
                ha=ha, va=va, fontsize=8.5, color="#64748B", zorder=6)

    # Reading guide at bottom
    ax.text(0, -1.50, "How to read · thick arcs = partners with overlapping stakeholder ecosystems",
            ha="center", va="top", fontsize=9.5, color="#475569", style="italic")
    ax.text(0, -1.62,
            "Absence of arcs = separated ecosystems · UNICATADORES has zero bridges with the Amazonian partners",
            ha="center", va="top", fontsize=9, color="#64748B")

    plt.savefig(out_path, dpi=180, facecolor="white", bbox_inches="tight", pad_inches=0.3)
    plt.close()
    print(f"PNG saved: {out_path}")

# ===========================================================
# RUN
# ===========================================================
render_alluvial(OUT_NET / "alluvial.png")
render_bridges(OUT_NET / "bridges.png")

# remove arquivo combinado antigo se existir
old = OUT_NET / "alluvial_chord.png"
if old.exists():
    old.unlink()
    print(f"Removed: {old}")

print("\nBridge pairs:")
for (a, b), n in sorted(bridge_pairs.items(), key=lambda x: -x[1]):
    print(f"  {PARTNERS[a]['name']} ↔ {PARTNERS[b]['name']}: {n}")
