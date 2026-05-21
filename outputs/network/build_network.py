"""
Constrói a rede do ecossistema Trias Brasil DGD 2027-2031 e renderiza
visualização interativa estilo Kumu (D3.js) em docs/index.html para
publicação via GitHub Pages.

Lê Stakeholder Ecosystem Mapping.xlsx, ancora os 4 parceiros MBO
(UNICAFES PA, UNICAFES RO, CSA Brasil, UNICATADORES) e calcula score de
alinhamento por biome, impact area, role e keywords específicas.
"""
from openpyxl import load_workbook
import networkx as nx
import json
import re
from collections import Counter
from pathlib import Path

ROOT = Path("/home/user/trias_ToC")
DOCS_SRC = ROOT / "docs-referencia"
OUT_NET = ROOT / "outputs" / "network"
OUT_PAGES = ROOT / "docs"
OUT_NET.mkdir(parents=True, exist_ok=True)
OUT_PAGES.mkdir(parents=True, exist_ok=True)

# ---------- 1. Carregar base ----------
wb = load_workbook(DOCS_SRC / "Stakeholder Ecosystem Mapping.xlsx", data_only=True)
ws = wb["Stakeholder list"]
rows = list(ws.iter_rows(values_only=True))
HEADER_ROW = 5
records = []
for r in rows[HEADER_ROW + 1:]:
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
        "value_chain": (str(r[12]) if r[12] else "").strip(),
        "potential_partnership": (str(r[13]) if r[13] else "").strip(),
        "hq": (str(r[14]) if r[14] else "").strip(),
        "funding_role": (str(r[15]) if r[15] else "").strip(),
        "url": (str(r[17]) if r[17] else "").strip(),
        "notes": (str(r[18]) if r[18] else "").strip(),
    })
print(f"Loaded {len(records)} stakeholders from database.")

# ---------- 2. Perfis-âncora dos 4 parceiros MBO ----------
PARTNERS = {
    "UNICAFES_PA": {
        "name": "UNICAFES Pará",
        "long_name": "União das Cooperativas da Agricultura Familiar e Economia Solidária — Pará",
        "sector": "MBO partner",
        "type": "Trias partner — Amazonia (rural)",
        "biomes": {"Amazonia"},
        "impacts": {"Land, Food and Forest"},
        "roles": {"Implementation capacity (delivery)",
                  "Community legitimacy / proximate leadership",
                  "Territorial anchor"},
        "keywords": ["agricultura familiar", "cooperativa", "amazon", "bioeconomia",
                     "açaí", "cacau", "cocoa", "coffee", "café", "agroflorest",
                     "sociobiodiversidade", "family farming", "agroecolog",
                     "northern brazil", "pará"],
        "territory": "Pará state · Amazonia",
        "description": ("2º nível MBO, Amazônia rural. Strategic anchor para a "
                        "bioeconomia pan-amazônica, cooperativismo e cadeias da "
                        "sociobiodiversidade (açaí, castanha, mel, cacau, café). "
                        "Opera 13 cooperativas de primeiro nível."),
    },
    "UNICAFES_RO": {
        "name": "UNICAFES Rondônia",
        "long_name": "União das Cooperativas da Agricultura Familiar e Economia Solidária — Rondônia",
        "sector": "MBO partner",
        "type": "Trias partner — Amazonia (exit 2028)",
        "biomes": {"Amazonia"},
        "impacts": {"Land, Food and Forest"},
        "roles": {"Implementation capacity (delivery)",
                  "Community legitimacy / proximate leadership",
                  "Territorial anchor"},
        "keywords": ["agricultura familiar", "cooperativa", "amazon", "bioeconomia",
                     "café", "coffee", "cocoa", "cacau", "agroflorest",
                     "blended finance", "pes", "carbon", "rondônia", "rondonia"],
        "territory": "Rondônia state · Amazonia",
        "description": ("2º nível MBO, Amazônia rural. Exit strategy em 2028. Foco "
                        "em cooperativismo, agroflorestal e finanças inovadoras "
                        "(blended finance, PES, mercados de carbono). Coordena 15 "
                        "cooperativas de primeiro nível."),
    },
    "CSA_BRASIL": {
        "name": "CSA Brasil",
        "long_name": "Comunidade que Sustenta a Agricultura — Brasil",
        "sector": "MBO partner",
        "type": "Trias partner — National (rural-urban)",
        "biomes": {"Mata Atlântica", "Cerrado", "Amazonia"},
        "impacts": {"Land, Food and Forest"},
        "roles": {"Implementation capacity (delivery)",
                  "Community legitimacy / proximate leadership",
                  "Convenor/enablers"},
        "keywords": ["community supported", "agroecolog", "food system",
                     "sistema alimentar", "saudáve", "food and nutrition",
                     "agricultura urbana", "peri-urb", "agricultor familiar",
                     "organic", "orgânico", "consumer", "rural-urban", "solidari"],
        "territory": "Nacional · rural-urbano",
        "description": ("3º nível MBO. Rede nacional de 200 unidades CSA em 19 "
                        "estados. Foco em educação alimentar, economia circular "
                        "rural-urbana e agricultura apoiada pela comunidade."),
    },
    "UNICATADORES": {
        "name": "UNICATADORES",
        "long_name": "União Nacional de Catadoras e Catadores",
        "sector": "MBO partner",
        "type": "Trias partner — National (urban)",
        "biomes": set(),
        "impacts": {"Land, Food and Forest", "Buildings & Transport"},
        "roles": {"Implementation capacity (delivery)",
                  "Policy influence / regulation",
                  "Community legitimacy / proximate leadership"},
        "keywords": ["catador", "waste pick", "recicl", "circular econom",
                     "resíduo sólido", "residuos solidos", "waste management",
                     "pnrs", "logística reversa", "extended producer", "epr",
                     "lixo", "urban"],
        "territory": "Nacional · urbano",
        "description": ("3º nível MBO maduro. Federação nacional representando 230 "
                        "cooperativas de catadores em 26 estados (~50 mil catadores, "
                        "60% mulheres). Sede em SP, ativa em PNRS e no Comitê "
                        "Interministerial CIISC."),
    },
}

# ---------- 3. Score ----------
def tokenize(s):
    return set(t.strip() for t in re.split(r"[,;]", s) if t.strip())

def score(stk, p):
    breakdown = []
    sc = 0.0
    biome_hit = False
    b_primary = tokenize(stk["biome_primary"])
    b_secondary = tokenize(stk["biome_secondary"])
    for biome in p["biomes"]:
        matched = False
        for tk in b_primary:
            if biome.lower() in tk.lower():
                sc += 3
                breakdown.append(f"biome primário: {tk}")
                biome_hit = True
                matched = True
                break
        if not matched:
            for tk in b_secondary:
                if biome.lower() in tk.lower():
                    sc += 1
                    breakdown.append(f"biome secundário: {tk}")
                    biome_hit = True
                    break
    impacts = tokenize(stk["impact_area"])
    for imp in p["impacts"]:
        for tk in impacts:
            if imp.lower() in tk.lower() or tk.lower() in imp.lower():
                sc += 1
                breakdown.append(f"impact: {tk}")
                break
    roles = tokenize(stk["role"])
    role_hits = 0
    matched_roles = []
    for r in p["roles"]:
        for tk in roles:
            if r.lower() == tk.lower():
                role_hits += 1
                matched_roles.append(tk)
                break
    if role_hits:
        sc += min(role_hits, 2) * 0.5
        breakdown.append(f"role: {', '.join(matched_roles)}")
    blob = (stk["description"] + " " + stk["notes"] + " " + stk["type"] +
            " " + stk["territory"] + " " + stk["name"]).lower()
    kw_hits = [kw for kw in p["keywords"] if kw.lower() in blob]
    if kw_hits:
        bonus = min(len(kw_hits), 4) * 2
        sc += bonus
        breakdown.append(f"keywords: {', '.join(kw_hits[:4])}")
    if "national" in stk["territory"].lower():
        sc += 0.5
        breakdown.append("alcance nacional")
    # Gates
    if p["biomes"] == {"Amazonia"} and not biome_hit and not kw_hits:
        return 0, []
    if p["name"] == "UNICATADORES" and not kw_hits:
        return 0, []
    return sc, breakdown

THRESHOLD = 4.5
results = []
for stk in records:
    scores = {}
    for pid, p in PARTNERS.items():
        sc, br = score(stk, p)
        if sc >= THRESHOLD:
            scores[pid] = {"score": round(sc, 1), "breakdown": br}
    if scores:
        results.append({"stk": stk, "scores": scores})
print(f"Conectados (score >= {THRESHOLD}): {len(results)}")

# ---------- 4. Estatísticas ----------
bridges = [r for r in results if len(r["scores"]) >= 2]
n_b3 = sum(1 for r in results if len(r["scores"]) >= 3)
n_b4 = sum(1 for r in results if len(r["scores"]) == 4)

# Cor por setor (paleta clara para tema escuro)
SECTOR_COLOR = {
    "Civil Society Organization (CSO)": "#5B9BD5",
    "Funders": "#F4A261",
    "Private sector": "#2A9D8F",
    "Public sector": "#8AB17D",
    "Others": "#C28BD7",
}
PARTNER_COLOR = "#E76F51"

# ---------- 5. Construir grafo (networkx) ----------
G = nx.Graph()

for pid, p in PARTNERS.items():
    G.add_node(pid,
               kind="partner",
               name=p["name"],
               long_name=p["long_name"],
               sector="MBO partner (Trias)",
               type=p["type"],
               territory=p["territory"],
               description=p["description"],
               color=PARTNER_COLOR)

for r in results:
    stk = r["stk"]
    node_id = f"S{stk['id']}"
    sector = stk["sector"] or "Others"
    G.add_node(node_id,
               kind="stakeholder",
               name=stk["name"],
               sector=sector,
               type=stk["type"],
               territory=stk["territory"],
               biome=stk["biome_primary"],
               impact=stk["impact_area"],
               role=stk["role"],
               funding_role=stk["funding_role"],
               hq=stk["hq"],
               url=stk["url"],
               description=stk["description"],
               notes=stk["notes"],
               n_partners=len(r["scores"]),
               color=SECTOR_COLOR.get(sector, "#999999"),
               connections={pid: r["scores"][pid]["score"] for pid in r["scores"]})
    for pid, s in r["scores"].items():
        G.add_edge(node_id, pid,
                   weight=s["score"],
                   breakdown=" · ".join(s["breakdown"]))

# Peer links entre os 4 parceiros
peer_pairs = [("UNICAFES_PA", "UNICAFES_RO"),
              ("UNICAFES_PA", "CSA_BRASIL"),
              ("UNICAFES_RO", "CSA_BRASIL"),
              ("UNICAFES_PA", "UNICATADORES"),
              ("UNICAFES_RO", "UNICATADORES"),
              ("CSA_BRASIL", "UNICATADORES")]
for a, b in peer_pairs:
    G.add_edge(a, b, weight=8, kind="peer",
               breakdown="rede de parceiros MBO da Trias Brasil")

print(f"Grafo: {G.number_of_nodes()} nós, {G.number_of_edges()} arestas")

# ---------- 6. Métricas ----------
betweenness = nx.betweenness_centrality(G, weight="weight")
degree = dict(G.degree())
for nid in G.nodes:
    G.nodes[nid]["betweenness"] = round(betweenness[nid], 4)
    G.nodes[nid]["degree"] = degree[nid]

# ---------- 7. Exportar JSON do grafo para o D3 ----------
nodes_json = []
for nid in G.nodes:
    n = dict(G.nodes[nid])
    n["id"] = nid
    # truncar campos longos no payload visível
    n["description_full"] = n.get("description", "")
    nodes_json.append(n)

links_json = []
for u, v, d in G.edges(data=True):
    links_json.append({
        "source": u,
        "target": v,
        "weight": d.get("weight", 1),
        "kind": d.get("kind", "alignment"),
        "breakdown": d.get("breakdown", ""),
    })

# Top stakeholders por parceiro
top_by_partner = {}
for pid, p in PARTNERS.items():
    lst = sorted(
        [(f"S{r['stk']['id']}", G.nodes[f"S{r['stk']['id']}"], r["scores"][pid]["score"])
         for r in results if pid in r["scores"]],
        key=lambda x: -x[2])
    top_by_partner[pid] = [{"name": x[1]["name"], "type": x[1]["type"],
                            "id": x[0], "score": x[2]}
                           for x in lst[:12]]

# Bridges ordenados
bridges_data = []
for r in sorted(bridges,
                key=lambda x: (-len(x["scores"]),
                               -sum(s["score"] for s in x["scores"].values()))):
    stk = r["stk"]
    bridges_data.append({
        "id": f"S{stk['id']}",
        "name": stk["name"],
        "type": stk["type"],
        "n": len(r["scores"]),
        "partners": [PARTNERS[pid]["name"] for pid in r["scores"]],
    })

graph_data = {
    "nodes": nodes_json,
    "links": links_json,
    "partners": [{"id": pid, **{k: v for k, v in p.items() if k != "biomes"}}
                 for pid, p in PARTNERS.items()],
    "sectors": SECTOR_COLOR,
    "stats": {
        "total": len(records),
        "connected": len(results),
        "bridges": len(bridges),
        "tri": n_b3,
        "quad": n_b4,
    },
    "top_by_partner": top_by_partner,
    "bridges": bridges_data,
}

# Sets não serializáveis
def jdefault(o):
    if isinstance(o, set):
        return list(o)
    raise TypeError
graph_json_str = json.dumps(graph_data, default=jdefault, ensure_ascii=False)

# ---------- 8. Renderizar HTML estilo Kumu (D3) ----------
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Ecossistema Trias Brasil DGD 2027–2031</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="Rede de stakeholders ancorada nos 4 parceiros MBO da Trias Brasil — UNICAFES Pará, UNICAFES Rondônia, CSA Brasil, UNICATADORES.">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@500&display=swap" rel="stylesheet">
<script src="https://d3js.org/d3.v7.min.js"></script>
<style>
  :root {
    --bg: #0e1218;
    --panel: #161c25;
    --panel-2: #1c2330;
    --border: #28313f;
    --text: #e9edf2;
    --muted: #8595a8;
    --muted-2: #5e6b7c;
    --accent: #e76f51;
    --accent-soft: rgba(231,111,81,0.15);
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; height: 100%; font-family: 'Inter', sans-serif;
                background: var(--bg); color: var(--text); overflow: hidden; }
  #app { display: grid; grid-template-columns: 290px 1fr 340px; height: 100vh; }
  aside { background: var(--panel); border-right: 1px solid var(--border);
          overflow-y: auto; padding: 22px 18px; }
  aside.right { border-right: none; border-left: 1px solid var(--border); }
  #stage { position: relative; background:
      radial-gradient(ellipse at center, #141a23 0%, var(--bg) 70%); overflow: hidden; }
  svg.network { width: 100%; height: 100%; cursor: grab; display: block; }
  svg.network:active { cursor: grabbing; }
  h1 { font-size: 16.5px; margin: 0 0 4px 0; font-weight: 700; letter-spacing: -0.2px; color: #fff; }
  h2 { font-size: 10.5px; text-transform: uppercase; letter-spacing: 1.4px;
       color: var(--muted-2); margin: 24px 0 10px 0; font-weight: 600; }
  .subtitle { font-size: 11px; color: var(--muted); margin-bottom: 16px; }
  .stats { display: grid; grid-template-columns: 1fr 1fr; gap: 7px; }
  .stat { background: var(--panel-2); border: 1px solid var(--border);
          border-radius: 10px; padding: 11px 10px; }
  .stat .n { font-size: 22px; font-weight: 700; color: #fff; line-height: 1; }
  .stat .l { font-size: 9.5px; color: var(--muted); text-transform: uppercase;
             letter-spacing: 0.5px; margin-top: 5px; font-weight: 600; }
  .legend-item { display: flex; align-items: center; gap: 9px; font-size: 11.5px;
                  padding: 4px 0; color: var(--text); cursor: pointer; user-select: none;
                  border-radius: 5px; padding: 4px 6px; margin: 0 -6px; transition: background .15s; }
  .legend-item:hover { background: var(--panel-2); }
  .legend-item.dim { opacity: 0.4; }
  .legend-item .dot { width: 11px; height: 11px; border-radius: 50%; flex-shrink: 0; }
  .legend-item .count { margin-left: auto; color: var(--muted); font-size: 10.5px;
                         font-family: 'JetBrains Mono', monospace; }
  .filter-btn { display: inline-block; background: var(--panel-2);
                border: 1px solid var(--border); color: var(--muted); padding: 5px 11px;
                border-radius: 999px; font-size: 11px; cursor: pointer;
                margin: 0 4px 6px 0; user-select: none; transition: all 0.15s; }
  .filter-btn:hover { background: var(--border); color: var(--text); }
  .filter-btn.active { background: var(--accent); color: #fff; border-color: var(--accent); }
  #search { width: 100%; padding: 9px 12px; border-radius: 8px;
            border: 1px solid var(--border); background: var(--bg);
            color: var(--text); font-size: 12.5px; font-family: inherit; }
  #search:focus { outline: none; border-color: var(--accent); }
  .top-block { background: var(--panel-2); border: 1px solid var(--border);
                border-radius: 10px; padding: 11px 13px; margin-bottom: 8px; }
  .top-block summary { cursor: pointer; font-size: 12px; outline: none;
                        user-select: none; color: var(--text); }
  .top-block summary b { color: #fff; }
  .top-list { padding: 10px 0 4px 0; margin: 0; list-style: none; font-size: 11.5px;
              line-height: 1.45; }
  .top-list li { padding: 6px 0; border-bottom: 1px solid var(--border);
                  cursor: pointer; transition: padding 0.15s; }
  .top-list li:last-child { border-bottom: none; }
  .top-list li:hover { padding-left: 4px; color: var(--accent); }
  .badge { background: var(--bg); color: var(--muted); font-size: 10px;
            padding: 2px 7px; border-radius: 999px; font-family: 'JetBrains Mono', monospace;
            border: 1px solid var(--border); }
  .muted { color: var(--muted); font-size: 10.5px; }
  .bridges-list { list-style: none; padding: 0; margin: 0; font-size: 11.5px;
                   line-height: 1.4; }
  .bridges-list li { background: var(--panel-2); border: 1px solid var(--border);
                     border-radius: 8px; padding: 9px 12px; margin-bottom: 6px;
                     cursor: pointer; transition: all 0.15s; }
  .bridges-list li:hover { border-color: var(--accent); background:
                            rgba(231,111,81,0.06); }
  details > summary { list-style: none; }
  details > summary::-webkit-details-marker { display: none; }
  details > summary::before { content: '▸  '; color: var(--muted-2); }
  details[open] > summary::before { content: '▾  '; }
  ::-webkit-scrollbar { width: 6px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }
  ::-webkit-scrollbar-thumb:hover { background: var(--muted-2); }

  /* SVG styling */
  .link { stroke: rgba(180,190,210,0.18); fill: none; stroke-linecap: round;
           transition: stroke 0.25s, stroke-width 0.25s, opacity 0.25s; }
  .link.peer { stroke: rgba(231,111,81,0.4); stroke-dasharray: 4 3; }
  .link.dim { opacity: 0.05; }
  .link.focus { stroke: rgba(231,111,81,0.65); stroke-width: 1.6px; }
  .node { cursor: pointer; }
  .node-halo { fill: var(--accent); opacity: 0; transition: opacity 0.25s; }
  .node-halo.show { opacity: 0.2; }
  .node-circle { stroke: var(--bg); stroke-width: 1.5px;
                 transition: stroke 0.25s, stroke-width 0.25s; filter: drop-shadow(0 1px 3px rgba(0,0,0,0.5)); }
  .node.partner .node-circle { stroke: #fff; stroke-width: 2.5px; }
  .node.dim .node-circle { opacity: 0.18; }
  .node.dim text { opacity: 0.12; }
  .node.focused .node-circle { stroke: #fff; stroke-width: 2.5px; }
  .node text { font-size: 10.5px; fill: var(--text); pointer-events: none;
                text-anchor: middle; font-weight: 500;
                paint-order: stroke; stroke: var(--bg); stroke-width: 3px;
                stroke-linejoin: round; transition: opacity 0.25s; }
  .node.partner text { font-weight: 700; font-size: 12.5px; fill: #fff; }
  .partner-anchor { pointer-events: none; }

  /* Inspector card */
  .inspector { position: absolute; right: 18px; top: 18px; width: 320px;
                background: var(--panel); border: 1px solid var(--border);
                border-radius: 12px; padding: 16px 18px; box-shadow: 0 10px 40px rgba(0,0,0,0.5);
                z-index: 10; display: none; max-height: calc(100vh - 40px); overflow-y: auto; }
  .inspector.show { display: block; }
  .inspector .close { position: absolute; right: 12px; top: 10px; cursor: pointer;
                       color: var(--muted); font-size: 18px; background: none;
                       border: none; padding: 4px 8px; }
  .inspector .close:hover { color: var(--text); }
  .inspector h3 { margin: 0 0 4px 0; font-size: 15px; padding-right: 24px; color: #fff; }
  .inspector .ins-type { font-size: 11px; color: var(--muted); margin-bottom: 12px; }
  .inspector .row { margin: 8px 0; font-size: 12px; line-height: 1.45; }
  .inspector .row .k { color: var(--muted-2); font-size: 10.5px; text-transform: uppercase;
                       letter-spacing: 0.6px; font-weight: 600; margin-bottom: 2px; }
  .inspector .row .v { color: var(--text); }
  .inspector .conn { display: inline-block; background: var(--accent-soft);
                      color: var(--accent); border: 1px solid rgba(231,111,81,0.3);
                      padding: 3px 8px; border-radius: 6px; font-size: 11px; margin: 2px 4px 2px 0;
                      font-weight: 500; }

  /* Floating controls */
  .controls { position: absolute; left: 18px; bottom: 18px; display: flex;
              gap: 6px; z-index: 5; }
  .ctrl-btn { background: var(--panel); border: 1px solid var(--border);
              color: var(--text); width: 34px; height: 34px; border-radius: 8px;
              cursor: pointer; font-size: 14px; transition: all 0.15s; font-family: inherit;
              display: flex; align-items: center; justify-content: center; }
  .ctrl-btn:hover { background: var(--panel-2); border-color: var(--accent); color: var(--accent); }
  .focus-hint { position: absolute; bottom: 20px; right: 18px; font-size: 11px;
                color: var(--muted-2); pointer-events: none; }
</style>
</head>
<body>
<div id="app">

  <aside class="left">
    <h1>Ecossistema Trias Brasil</h1>
    <div class="subtitle">DGD 2027–2031 · 4 parceiros MBO ancoram a rede</div>

    <div class="stats">
      <div class="stat"><div class="n" id="s-total"></div><div class="l">stakeholders</div></div>
      <div class="stat"><div class="n" id="s-connected"></div><div class="l">conectados</div></div>
      <div class="stat"><div class="n" id="s-bridges"></div><div class="l">bridges (2+)</div></div>
      <div class="stat"><div class="n" id="s-tri"></div><div class="l">tri-bridges</div></div>
    </div>

    <h2>Busca</h2>
    <input id="search" placeholder="Filtrar por nome…" autocomplete="off">

    <h2>Filtro por parceiro</h2>
    <div id="partner-filters">
      <span class="filter-btn active" data-partner="ALL">Todos</span>
      <span class="filter-btn" data-partner="UNICAFES_PA">UNICAFES PA</span>
      <span class="filter-btn" data-partner="UNICAFES_RO">UNICAFES RO</span>
      <span class="filter-btn" data-partner="CSA_BRASIL">CSA Brasil</span>
      <span class="filter-btn" data-partner="UNICATADORES">UNICATADORES</span>
    </div>

    <h2>Setor (clique para filtrar)</h2>
    <div id="legend"></div>

    <h2>Como interagir</h2>
    <div class="muted" style="line-height: 1.5;">
      <b>Hover</b> exibe halo no nó.<br>
      <b>Clique</b> num nó ativa o modo foco — apenas a vizinhança permanece destacada.
      <br><b>Clique no vazio</b> para sair do foco.
      <br><b>Roda do mouse</b> dá zoom, arraste move a câmera.
    </div>
  </aside>

  <main id="stage">
    <svg class="network" id="net"></svg>

    <div class="controls">
      <button class="ctrl-btn" id="btn-reset" title="Resetar visão">⌂</button>
      <button class="ctrl-btn" id="btn-zoom-in" title="Zoom in">+</button>
      <button class="ctrl-btn" id="btn-zoom-out" title="Zoom out">−</button>
      <button class="ctrl-btn" id="btn-physics" title="Pausar/retomar física">⏸</button>
    </div>

    <div class="inspector" id="inspector">
      <button class="close" id="inspector-close">×</button>
      <div id="inspector-body"></div>
    </div>
  </main>

  <aside class="right">
    <h2>Bridges — atores que conectam múltiplos parceiros</h2>
    <ul class="bridges-list" id="bridges-list"></ul>

    <h2>Top atores por parceiro</h2>
    <div id="top-blocks"></div>
  </aside>
</div>

<script>
const DATA = __GRAPH_JSON__;

// ---------- DOM helpers ----------
const $ = sel => document.querySelector(sel);
const $$ = sel => Array.from(document.querySelectorAll(sel));

// ---------- Fill stats and lists ----------
$('#s-total').textContent = DATA.stats.total;
$('#s-connected').textContent = DATA.stats.connected;
$('#s-bridges').textContent = DATA.stats.bridges;
$('#s-tri').textContent = DATA.stats.tri;

// Sector counts
const sectorCounts = {};
DATA.nodes.forEach(n => {
  if (n.kind === 'partner') return;
  const s = n.sector || 'Others';
  sectorCounts[s] = (sectorCounts[s] || 0) + 1;
});
const legendEl = $('#legend');
const activeSectors = new Set(Object.keys(DATA.sectors));
Object.entries(DATA.sectors).forEach(([sector, col]) => {
  const div = document.createElement('div');
  div.className = 'legend-item';
  div.dataset.sector = sector;
  div.innerHTML = `<span class="dot" style="background:${col}"></span>${sector}<span class="count">${sectorCounts[sector]||0}</span>`;
  div.addEventListener('click', () => {
    if (activeSectors.has(sector)) {
      activeSectors.delete(sector);
      div.classList.add('dim');
    } else {
      activeSectors.add(sector);
      div.classList.remove('dim');
    }
    applySectorFilter();
  });
  legendEl.appendChild(div);
});

// Bridges list
const bridgesEl = $('#bridges-list');
DATA.bridges.slice(0, 25).forEach(b => {
  const li = document.createElement('li');
  li.innerHTML = `<b>${b.name}</b> <span class="badge">${b.n}×</span><br><span class="muted">→ ${b.partners.join(' · ')}</span>`;
  li.addEventListener('click', () => selectNode(b.id, true));
  bridgesEl.appendChild(li);
});

// Top by partner
const tbEl = $('#top-blocks');
DATA.partners.forEach(p => {
  const lst = DATA.top_by_partner[p.id] || [];
  const det = document.createElement('details');
  det.className = 'top-block';
  det.innerHTML = `<summary><b>${p.name}</b> — ${lst.length} no top</summary>`;
  const ol = document.createElement('ol');
  ol.className = 'top-list';
  lst.forEach(item => {
    const li = document.createElement('li');
    li.innerHTML = `<b>${item.name}</b> <span class="badge">${item.score}</span><br><span class="muted">${item.type || ''}</span>`;
    li.addEventListener('click', () => selectNode(item.id, true));
    ol.appendChild(li);
  });
  det.appendChild(ol);
  tbEl.appendChild(det);
});

// ---------- D3 setup ----------
const svg = d3.select('#net');
const W = () => svg.node().clientWidth;
const H = () => svg.node().clientHeight;

const root = svg.append('g').attr('class', 'root');
const linkLayer = root.append('g').attr('class', 'links');
const nodeLayer = root.append('g').attr('class', 'nodes');

const zoom = d3.zoom().scaleExtent([0.2, 5]).on('zoom', ev => {
  root.attr('transform', ev.transform);
});
svg.call(zoom);

// Initial partner positions — anchored in a square layout to give structure
function partnerAnchors() {
  const cx = W() / 2, cy = H() / 2;
  const r = Math.min(W(), H()) * 0.22;
  return {
    UNICAFES_PA:  {x: cx - r,     y: cy - r * 0.9},
    UNICAFES_RO:  {x: cx + r,     y: cy - r * 0.9},
    CSA_BRASIL:   {x: cx - r,     y: cy + r * 0.9},
    UNICATADORES: {x: cx + r,     y: cy + r * 0.9},
  };
}

// Prepare nodes & links (note: d3 mutates objects)
const nodes = DATA.nodes.map(n => Object.assign({}, n));
const links = DATA.links.map(l => Object.assign({}, l));

// Place partners at initial anchors and fix them with weaker pull
const anchors = partnerAnchors();
nodes.forEach(n => {
  if (n.kind === 'partner') {
    n.fx = anchors[n.id].x;
    n.fy = anchors[n.id].y;
  }
});

// Force simulation
const sim = d3.forceSimulation(nodes)
  .force('link', d3.forceLink(links).id(d => d.id)
                  .distance(l => l.kind === 'peer' ? 220 : 120 + (8 - l.weight) * 12)
                  .strength(l => l.kind === 'peer' ? 0.05 : 0.4))
  .force('charge', d3.forceManyBody().strength(d => d.kind === 'partner' ? -1400 : -260))
  .force('center', d3.forceCenter(0, 0).strength(0.02))
  .force('collide', d3.forceCollide().radius(d => nodeRadius(d) + 4))
  .alphaDecay(0.025);

function nodeRadius(d) {
  if (d.kind === 'partner') return 22;
  return 6 + (d.n_partners || 1) * 4;
}

// Render links (curved)
const linkSel = linkLayer.selectAll('path.link')
  .data(links).join('path')
  .attr('class', d => 'link ' + (d.kind === 'peer' ? 'peer' : ''))
  .attr('stroke-width', d => d.kind === 'peer' ? 1.6 : 0.5 + d.weight * 0.18);

// Render nodes
const nodeSel = nodeLayer.selectAll('g.node')
  .data(nodes).join('g')
  .attr('class', d => 'node ' + (d.kind === 'partner' ? 'partner' : ''))
  .attr('data-id', d => d.id)
  .call(d3.drag()
    .on('start', dragstart)
    .on('drag', dragmove)
    .on('end', dragend));

// halo (behind circle)
nodeSel.append('circle')
  .attr('class', 'node-halo')
  .attr('r', d => nodeRadius(d) + 8);
nodeSel.append('circle')
  .attr('class', 'node-circle')
  .attr('r', nodeRadius)
  .attr('fill', d => d.color);

// labels: only partners always; others on hover or when zoomed
nodeSel.append('text')
  .attr('dy', d => nodeRadius(d) + 13)
  .text(d => {
    if (d.kind === 'partner') return d.name;
    if (d.n_partners >= 2 || d.betweenness > 0.01) {
      return d.name.length > 30 ? d.name.slice(0, 27) + '…' : d.name;
    }
    return '';
  });

// Hover behavior
nodeSel.on('mouseenter', function(ev, d) {
  d3.select(this).select('.node-halo').classed('show', true);
})
.on('mouseleave', function(ev, d) {
  d3.select(this).select('.node-halo').classed('show', false);
})
.on('click', function(ev, d) {
  ev.stopPropagation();
  selectNode(d.id, false);
});

// Click empty: clear focus
svg.on('click', () => clearFocus());

sim.on('tick', () => {
  linkSel.attr('d', d => {
    const dx = d.target.x - d.source.x;
    const dy = d.target.y - d.source.y;
    const dr = Math.sqrt(dx*dx + dy*dy) * 2.2;
    return `M${d.source.x},${d.source.y}A${dr},${dr} 0 0,1 ${d.target.x},${d.target.y}`;
  });
  nodeSel.attr('transform', d => `translate(${d.x},${d.y})`);
});

function dragstart(ev, d) {
  if (!ev.active) sim.alphaTarget(0.3).restart();
  d.fx = d.x; d.fy = d.y;
}
function dragmove(ev, d) {
  d.fx = ev.x; d.fy = ev.y;
}
function dragend(ev, d) {
  if (!ev.active) sim.alphaTarget(0);
  if (d.kind === 'partner') return;  // partners stay fixed if not partner? keep free
  d.fx = null; d.fy = null;
}

// ---------- Focus mode ----------
let focusId = null;
function selectNode(id, fromList) {
  const target = nodes.find(n => n.id === id);
  if (!target) return;
  focusId = id;
  const neighbors = new Set([id]);
  links.forEach(l => {
    const s = l.source.id || l.source;
    const t = l.target.id || l.target;
    if (s === id) neighbors.add(t);
    if (t === id) neighbors.add(s);
  });
  nodeSel.classed('dim', n => !neighbors.has(n.id));
  nodeSel.classed('focused', n => n.id === id);
  linkSel.classed('dim', l => {
    const s = l.source.id || l.source;
    const t = l.target.id || l.target;
    return !(s === id || t === id);
  }).classed('focus', l => {
    const s = l.source.id || l.source;
    const t = l.target.id || l.target;
    return s === id || t === id;
  });
  showInspector(target);
  if (fromList) {
    // center the view roughly on the node
    const tr = d3.zoomTransform(svg.node());
    const tx = W()/2 - target.x * tr.k;
    const ty = H()/2 - target.y * tr.k;
    svg.transition().duration(600)
      .call(zoom.transform, d3.zoomIdentity.translate(tx, ty).scale(tr.k));
  }
}

function clearFocus() {
  focusId = null;
  nodeSel.classed('dim', false).classed('focused', false);
  linkSel.classed('dim', false).classed('focus', false);
  applySectorFilter();
  applyPartnerFilter();
  $('#inspector').classList.remove('show');
}

function showInspector(n) {
  const el = $('#inspector-body');
  if (n.kind === 'partner') {
    el.innerHTML = `
      <h3>${n.name}</h3>
      <div class="ins-type">${n.type}</div>
      <div class="row"><div class="k">Nome completo</div><div class="v">${n.long_name||''}</div></div>
      <div class="row"><div class="k">Território</div><div class="v">${n.territory||''}</div></div>
      <div class="row"><div class="k">Descrição</div><div class="v">${n.description||''}</div></div>
    `;
  } else {
    const conns = Object.entries(n.connections||{}).map(([pid, sc]) => {
      const pname = DATA.partners.find(p => p.id === pid).name;
      return `<span class="conn">${pname} · ${sc}</span>`;
    }).join('');
    el.innerHTML = `
      <h3>${n.name}</h3>
      <div class="ins-type">${n.type || ''} · ${n.sector || ''}</div>
      <div class="row"><div class="k">Território</div><div class="v">${n.territory||'—'}</div></div>
      <div class="row"><div class="k">Biome primário</div><div class="v">${n.biome||'—'}</div></div>
      <div class="row"><div class="k">Impact area</div><div class="v">${n.impact||'—'}</div></div>
      <div class="row"><div class="k">Role no ecossistema</div><div class="v">${n.role||'—'}</div></div>
      <div class="row"><div class="k">Funding role</div><div class="v">${n.funding_role||'—'}</div></div>
      <div class="row"><div class="k">HQ</div><div class="v">${n.hq||'—'}</div></div>
      ${n.url ? `<div class="row"><div class="k">Site</div><div class="v"><a href="${n.url}" target="_blank" style="color:var(--accent)">${n.url}</a></div></div>` : ''}
      <div class="row"><div class="k">Conexões com parceiros</div><div class="v">${conns}</div></div>
      ${n.description ? `<div class="row"><div class="k">Descrição</div><div class="v" style="font-size:11.5px;color:var(--muted)">${n.description}</div></div>` : ''}
    `;
  }
  $('#inspector').classList.add('show');
}
$('#inspector-close').addEventListener('click', () => clearFocus());

// ---------- Search ----------
$('#search').addEventListener('input', e => {
  const q = e.target.value.toLowerCase().trim();
  if (!q) {
    clearFocus();
    return;
  }
  nodeSel.classed('dim', n => !n.name.toLowerCase().includes(q));
  linkSel.classed('dim', l => {
    const sname = (l.source.name || nodes.find(n => n.id === l.source).name).toLowerCase();
    const tname = (l.target.name || nodes.find(n => n.id === l.target).name).toLowerCase();
    return !sname.includes(q) && !tname.includes(q);
  });
});

// ---------- Sector filter ----------
function applySectorFilter() {
  nodeSel.classed('dim', n => {
    if (n.kind === 'partner') return false;
    return !activeSectors.has(n.sector || 'Others');
  });
  linkSel.classed('dim', l => {
    const s = nodes.find(n => n.id === (l.source.id || l.source));
    const t = nodes.find(n => n.id === (l.target.id || l.target));
    return (s.kind !== 'partner' && !activeSectors.has(s.sector || 'Others')) ||
           (t.kind !== 'partner' && !activeSectors.has(t.sector || 'Others'));
  });
}

// ---------- Partner filter ----------
let activePartner = 'ALL';
$$('.filter-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    $$('.filter-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    activePartner = btn.dataset.partner;
    applyPartnerFilter();
  });
});
function applyPartnerFilter() {
  if (activePartner === 'ALL') {
    nodeSel.classed('dim', false);
    linkSel.classed('dim', false);
    return;
  }
  const connected = new Set([activePartner]);
  links.forEach(l => {
    const s = l.source.id || l.source;
    const t = l.target.id || l.target;
    if (s === activePartner) connected.add(t);
    if (t === activePartner) connected.add(s);
  });
  nodeSel.classed('dim', n => !connected.has(n.id));
  linkSel.classed('dim', l => {
    const s = l.source.id || l.source;
    const t = l.target.id || l.target;
    return !(connected.has(s) && connected.has(t));
  });
}

// ---------- Controls ----------
$('#btn-zoom-in').addEventListener('click', () =>
  svg.transition().call(zoom.scaleBy, 1.4));
$('#btn-zoom-out').addEventListener('click', () =>
  svg.transition().call(zoom.scaleBy, 0.7));
$('#btn-reset').addEventListener('click', () => {
  svg.transition().duration(600).call(zoom.transform, d3.zoomIdentity);
});
let physOn = true;
$('#btn-physics').addEventListener('click', () => {
  physOn = !physOn;
  $('#btn-physics').textContent = physOn ? '⏸' : '▶';
  if (physOn) sim.alphaTarget(0.05).restart();
  else sim.alphaTarget(0).stop();
});

// Initial zoom centered around origin
svg.call(zoom.transform, d3.zoomIdentity.translate(W()/2, H()/2));

// Adjust on resize
window.addEventListener('resize', () => {
  const a = partnerAnchors();
  // Re-anchor relative to new center
  nodes.forEach(n => {
    if (n.kind === 'partner' && a[n.id]) {
      n.fx = a[n.id].x - W()/2;
      n.fy = a[n.id].y - H()/2;
    }
  });
  sim.alpha(0.3).restart();
});

// Translate partner fixed coords to origin-centered space initially
const a0 = partnerAnchors();
nodes.forEach(n => {
  if (n.kind === 'partner') {
    n.fx = a0[n.id].x - W()/2;
    n.fy = a0[n.id].y - H()/2;
    n.x = n.fx; n.y = n.fy;
  }
});
sim.alpha(1).restart();
</script>
</body>
</html>
"""

html = HTML_TEMPLATE.replace("__GRAPH_JSON__", graph_json_str)

# Write to both places
(OUT_NET / "ecossistema_trias_brasil.html").write_text(html, encoding="utf-8")
(OUT_PAGES / "index.html").write_text(html, encoding="utf-8")
print(f"HTML salvo em:\n  {OUT_NET / 'ecossistema_trias_brasil.html'}\n  {OUT_PAGES / 'index.html'}")

# ---------- 9. Markdown report ----------
md = [f"# Análise de rede — Ecossistema Trias Brasil DGD 2027-2031\n"]
md.append("Rede de stakeholders ancorada nos 4 parceiros MBO da Trias Brasil "
          "(UNICAFES Pará, UNICAFES Rondônia, CSA Brasil e UNICATADORES). "
          "Fonte: `Stakeholder Ecosystem Mapping.xlsx`.\n")
md.append(f"\n**Visualização interativa**: ver `docs/index.html` "
          f"(publicado via GitHub Pages quando habilitado no repo).\n")
md.append("## Resumo\n")
md.append(f"- **{len(records)}** stakeholders na base de origem\n"
          f"- **{len(results)}** conectados a pelo menos 1 parceiro (score ≥ {THRESHOLD})\n"
          f"- **{len(bridges)}** bridges (conectados a 2 ou mais parceiros)\n"
          f"- **{n_b3}** conectam 3+ parceiros\n"
          f"- **{n_b4}** conectam os 4 parceiros\n")

md.append("\n## Metodologia\n")
md.append("Edges ponderadas por alinhamento temático/territorial:\n"
          "- biome primário (+3) / secundário (+1)\n"
          "- impact area (+1) — peso reduzido porque 'Land, Food and Forest' "
          "aparece em 244/360 stakeholders\n"
          "- role no ecossistema (+0.5/role, cap 1)\n"
          "- keywords específicas por parceiro (+2/keyword, cap 8) — "
          "principal diferenciador\n"
          "- alcance nacional (+0.5)\n"
          "- **gate amazônico**: para UNICAFES PA/RO, sem biome amazônica nem "
          "keyword amazônica o score zera\n"
          "- **gate de catadores**: para UNICATADORES, exige keyword de "
          "resíduo/catador/circular\n"
          f"\nThreshold para criar edge: score ≥ {THRESHOLD}.\n")

md.append("\n## Bridges (conectam múltiplos parceiros)\n")
for b in bridges_data[:30]:
    md.append(f"- **{b['name']}** ({b['type']}) — {b['n']}× · "
              f"{', '.join(b['partners'])}")

md.append("\n\n## Top 10 stakeholders por parceiro\n")
for p in PARTNERS:
    md.append(f"\n### {PARTNERS[p]['name']}\n")
    for item in top_by_partner[p][:10]:
        md.append(f"- **{item['name']}** — score {item['score']} · {item['type']}")

md.append("\n\n## Top 15 por betweenness centrality (potenciais brokers)\n")
top_bw = sorted(betweenness.items(), key=lambda x: -x[1])[:15]
for nid, b in top_bw:
    label = G.nodes[nid].get("name", nid)
    md.append(f"- {label} — {b:.4f}")

(OUT_NET / "analise_rede.md").write_text("\n".join(md), encoding="utf-8")
print(f"Relatório: {OUT_NET / 'analise_rede.md'}")
