"""
Construção da rede do ecossistema Trias Brasil DGD 2027-2031.

Lê a base Stakeholder Ecosystem Mapping.xlsx, ancora os 4 parceiros estratégicos
(UNICAFES PA, UNICAFES RO, CSA Brasil, UNICATADORES) e calcula um score de
alinhamento de cada stakeholder com cada parceiro com base em:
 - biome (territorial)
 - impact area (temático)
 - role no ecossistema
 - palavras-chave da descrição

Gera HTML interativo (pyvis) e um relatório markdown com os top stakeholders
por parceiro e os bridges (que conectam múltiplos parceiros).
"""
from openpyxl import load_workbook
from pyvis.network import Network
import networkx as nx
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path("/home/user/trias_ToC")
DOCS = ROOT / "docs-referencia"
OUT = ROOT / "outputs" / "network"
OUT.mkdir(parents=True, exist_ok=True)

# ---------- 1. Carregar dados ----------
wb = load_workbook(DOCS / "Stakeholder Ecosystem Mapping.xlsx", data_only=True)
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

# ---------- 2. Perfis-âncora dos 4 parceiros ----------
PARTNERS = {
    "UNICAFES_PA": {
        "name": "UNICAFES Pará",
        "long_name": "União das Cooperativas da Agricultura Familiar e Economia Solidária — Pará",
        "sector": "Civil Society Organization (CSO)",
        "type": "MBO partner (Trias)",
        "biomes": {"Amazonia"},
        "impacts": {"Land, Food and Forest"},
        "roles": {"Implementation capacity (delivery)", "Community legitimacy / proximate leadership",
                  "Territorial anchor"},
        "keywords": ["agricultura familiar", "cooperativa", "amazon", "bioeconomia",
                     "açaí", "cacau", "cocoa", "coffee", "café", "agroflorest", "sociobiodiversidade",
                     "family farming", "agroecolog", "northern brazil", "pará"],
        "territory": "Amazonia / Pará state",
        "description": ("2º nível MBO, Amazônia rural. Strategic anchor para bioeconomia "
                        "pan-amazônica, cooperativismo e cadeias da sociobiodiversidade (açaí, "
                        "castanha, mel, cacau, café). Atua em 13 cooperativas de primeiro nível."),
    },
    "UNICAFES_RO": {
        "name": "UNICAFES Rondônia",
        "long_name": "União das Cooperativas da Agricultura Familiar e Economia Solidária — Rondônia",
        "sector": "Civil Society Organization (CSO)",
        "type": "MBO partner (Trias, exit 2028)",
        "biomes": {"Amazonia"},
        "impacts": {"Land, Food and Forest"},
        "roles": {"Implementation capacity (delivery)", "Community legitimacy / proximate leadership",
                  "Territorial anchor"},
        "keywords": ["agricultura familiar", "cooperativa", "amazon", "bioeconomia",
                     "café", "coffee", "cocoa", "cacau", "agroflorest", "blended finance",
                     "pes", "carbon", "rondônia", "rondonia"],
        "territory": "Amazonia / Rondônia state",
        "description": ("2º nível MBO, Amazônia rural, exit strategy em 2028. Foco em "
                        "cooperativismo, agroflorestal e finanças inovadoras (blended finance, "
                        "PES, mercados de carbono). Coordena 15 cooperativas de primeiro nível."),
    },
    "CSA_BRASIL": {
        "name": "CSA Brasil",
        "long_name": "Comunidade que Sustenta a Agricultura — Brasil",
        "sector": "Civil Society Organization (CSO)",
        "type": "MBO partner (Trias)",
        "biomes": {"Mata Atlântica", "Cerrado", "Amazonia"},  # nacional
        "impacts": {"Land, Food and Forest"},
        "roles": {"Implementation capacity (delivery)", "Community legitimacy / proximate leadership",
                  "Convenor/enablers"},
        "keywords": ["community supported", "agroecolog", "food system", "sistema alimentar",
                     "saudáve", "food and nutrition", "agricultura urbana", "peri-urb",
                     "agricultor familiar", "organic", "orgânico", "consumer", "rural-urban",
                     "solidari"],
        "territory": "National (rural-urban)",
        "description": ("3º nível MBO, rural-urbano. Rede nacional de 200 unidades CSA em "
                        "19 estados. Foco em educação alimentar, economia circular rural-"
                        "urbana e agricultura apoiada pela comunidade."),
    },
    "UNICATADORES": {
        "name": "UNICATADORES",
        "long_name": "União Nacional de Catadoras e Catadores",
        "sector": "Others",
        "type": "MBO partner (Trias)",
        "biomes": set(),  # urban, not biome-bound
        "impacts": {"Land, Food and Forest", "Buildings & Transport"},
        "roles": {"Implementation capacity (delivery)", "Policy influence / regulation",
                  "Community legitimacy / proximate leadership"},
        "keywords": ["catador", "waste pick", "recicl", "circular econom", "resíduo sólido",
                     "residuos solidos", "waste management", "pnrs", "logística reversa",
                     "extended producer", "epr", "lixo", "urban"],
        "territory": "National (urban)",
        "description": ("3º nível MBO maduro, urbano. Federação nacional representando 230 "
                        "cooperativas de catadores em 26 estados (~50 mil catadores, 60% mulheres). "
                        "Sede em SP, ativa em diálogos do PNRS e do Comitê Interministerial CIISC."),
    },
}

# ---------- 3. Função de score (stakeholder → parceiro) ----------
def tokenize(s):
    return set(t.strip() for t in re.split(r"[,;]", s) if t.strip())

def has_token(s, target):
    return any(target.lower() in t.lower() for t in tokenize(s))

def score(stk, p):
    """Retorna (score, breakdown).

    Calibração: o impact area 'Land, Food and Forest' aparece em 244/360 dos
    stakeholders, então tem peso baixo. Biome (Amazônia) e keywords específicas
    são os principais diferenciadores. Exige sinal específico para conectar.
    """
    breakdown = []
    sc = 0
    # Biome match (peso forte — diferencia parceiros amazônicos dos demais)
    b_primary = tokenize(stk["biome_primary"])
    b_secondary = tokenize(stk["biome_secondary"])
    biome_hit = False
    for biome in p["biomes"]:
        for tk in b_primary:
            if biome.lower() in tk.lower():
                sc += 3
                breakdown.append(f"biome_primary={tk} (+3)")
                biome_hit = True
                break
        if not biome_hit:
            for tk in b_secondary:
                if biome.lower() in tk.lower():
                    sc += 1
                    breakdown.append(f"biome_secondary={tk} (+1)")
                    biome_hit = True
                    break
    # Impact area — peso reduzido porque é largo demais
    impacts = tokenize(stk["impact_area"])
    for imp in p["impacts"]:
        for tk in impacts:
            if imp.lower() in tk.lower() or tk.lower() in imp.lower():
                sc += 1
                breakdown.append(f"impact={tk} (+1)")
                break
    # Role — peso baixo
    roles = tokenize(stk["role"])
    role_hits = 0
    for r in p["roles"]:
        for tk in roles:
            if r.lower() == tk.lower():
                role_hits += 1
                breakdown.append(f"role={tk} (+0.5)")
                break
    sc += min(role_hits, 2) * 0.5
    # Keywords específicas (sinal forte e diagnóstico)
    blob = (stk["description"] + " " + stk["notes"] + " " +
            stk["type"] + " " + stk["territory"] + " " + stk["name"]).lower()
    kw_hits = []
    for kw in p["keywords"]:
        if kw.lower() in blob:
            kw_hits.append(kw)
    if kw_hits:
        bonus = min(len(kw_hits), 4) * 2  # até 8 pts
        sc += bonus
        breakdown.append(f"keywords={kw_hits[:5]} (+{bonus})")
    # Territorial National — neutro/baixo
    if "national" in stk["territory"].lower():
        sc += 0.5
        breakdown.append("territory=National (+0.5)")
    # GATE: para parceiros amazônicos, sem biome amazônica nem keyword amazônica,
    # zera o score (evita falsos positivos de empresas nacionais broad-spectrum)
    if p["biomes"] == {"Amazonia"} and not biome_hit and not kw_hits:
        return 0, ["GATE: no Amazon biome nor amazon keyword"]
    # GATE: para UNICATADORES, exige keyword (catador/waste/circular/PNRS)
    if p["name"] == "UNICATADORES" and not kw_hits:
        return 0, ["GATE: no waste/recycling/circular keyword"]
    return sc, breakdown

# ---------- 4. Calcular score de cada stakeholder vs cada parceiro ----------
THRESHOLD = 4.5  # mínimo para criar edge

results = []
for stk in records:
    scores = {}
    for pid, p in PARTNERS.items():
        sc, br = score(stk, p)
        if sc >= THRESHOLD:
            scores[pid] = {"score": sc, "breakdown": br}
    if scores:
        results.append({"stk": stk, "scores": scores})

print(f"Stakeholders com pelo menos uma conexão (score >= {THRESHOLD}): {len(results)}")

# Bridges: stakeholders conectados a 2+ parceiros
bridges = [r for r in results if len(r["scores"]) >= 2]
print(f"Bridges (conectados a 2+ parceiros): {len(bridges)}")
print(f"Bridges (3+): {sum(1 for r in results if len(r['scores']) >= 3)}")
print(f"Bridges (4): {sum(1 for r in results if len(r['scores']) == 4)}")

# ---------- 5. Construir grafo networkx ----------
G = nx.Graph()

SECTOR_COLOR = {
    "Civil Society Organization (CSO)": "#4E79A7",
    "Funders": "#F28E2B",
    "Private sector": "#76B7B2",
    "Public sector": "#59A14F",
    "Others": "#B07AA1",
}
PARTNER_COLOR = "#E15759"

# nós-âncora
for pid, p in PARTNERS.items():
    G.add_node(
        pid,
        label=p["name"],
        title=f"<b>{p['name']}</b><br>{p['long_name']}<br><br>{p['description']}",
        group="partner",
        color=PARTNER_COLOR,
        size=55,
        shape="star",
        partner=True,
    )

# nós-stakeholders e arestas
for r in results:
    stk = r["stk"]
    node_id = f"S{stk['id']}"
    n_partners = len(r["scores"])
    total_score = sum(s["score"] for s in r["scores"].values())
    # tooltip rico
    role_short = stk["role"][:120] + ("…" if len(stk["role"]) > 120 else "")
    desc_short = stk["description"][:280] + ("…" if len(stk["description"]) > 280 else "")
    title = (
        f"<b>{stk['name']}</b><br>"
        f"<i>{stk['type']}</i> · {stk['sector']}<br>"
        f"<b>Território:</b> {stk['territory']} · "
        f"<b>Biome:</b> {stk['biome_primary']}<br>"
        f"<b>Impact:</b> {stk['impact_area']}<br>"
        f"<b>Role:</b> {role_short}<br>"
        f"<b>Funding role:</b> {stk['funding_role']}<br><br>"
        f"{desc_short}<br><br>"
        f"<b>Conectado a:</b> {', '.join(PARTNERS[pid]['name'] for pid in r['scores'])}"
    )
    G.add_node(
        node_id,
        label=stk["name"][:42] + ("…" if len(stk["name"]) > 42 else ""),
        title=title,
        group=stk["sector"] or "Others",
        color=SECTOR_COLOR.get(stk["sector"], "#999999"),
        size=12 + n_partners * 8,  # bridges aparecem maiores
        n_partners=n_partners,
        score_total=total_score,
        partner=False,
        url=stk["url"],
    )
    for pid, s in r["scores"].items():
        G.add_edge(node_id, pid, weight=s["score"], title=f"score {s['score']}")

# arestas peer-to-peer entre os 4 parceiros (todos MBOs do UNICOPAS / mesma rede)
peer_partners = [("UNICAFES_PA", "UNICAFES_RO"),
                 ("UNICAFES_PA", "CSA_BRASIL"),
                 ("UNICAFES_RO", "CSA_BRASIL"),
                 ("UNICAFES_PA", "UNICATADORES"),
                 ("UNICAFES_RO", "UNICATADORES"),
                 ("CSA_BRASIL", "UNICATADORES")]
for a, b in peer_partners:
    G.add_edge(a, b, weight=10, dashes=True, color="#E15759",
               title="Parceiros MBO — rede UNICOPAS / Trias Brasil")

print(f"\nGrafo: {G.number_of_nodes()} nós, {G.number_of_edges()} arestas")

# ---------- 6. Métricas de rede ----------
degree_cent = nx.degree_centrality(G)
betweenness = nx.betweenness_centrality(G, weight="weight")
print("\nTop 10 por betweenness (potenciais brokers):")
top_bw = sorted(betweenness.items(), key=lambda x: -x[1])[:15]
for nid, b in top_bw:
    label = G.nodes[nid].get("label", nid)
    print(f"  {b:.4f}  {label}")

# ---------- 7. Exportar HTML interativo (pyvis) ----------
net = Network(
    height="100vh",
    width="100%",
    bgcolor="#0f1419",
    font_color="#e6e6e6",
    notebook=False,
    directed=False,
    cdn_resources="remote",
)
net.from_nx(G)
# Opções de layout e física
net.set_options("""
{
  "physics": {
    "enabled": true,
    "barnesHut": {
      "gravitationalConstant": -8000,
      "centralGravity": 0.25,
      "springLength": 130,
      "springConstant": 0.04,
      "damping": 0.4,
      "avoidOverlap": 0.3
    },
    "minVelocity": 0.5,
    "stabilization": {"iterations": 250, "fit": true}
  },
  "nodes": {
    "borderWidth": 1,
    "borderWidthSelected": 4,
    "font": {"color": "#e6e6e6", "size": 12, "face": "Inter, sans-serif", "strokeWidth": 0},
    "shadow": {"enabled": true, "size": 12, "color": "rgba(0,0,0,0.5)"}
  },
  "edges": {
    "color": {"color": "rgba(180,180,200,0.35)", "highlight": "#fff", "inherit": false},
    "smooth": {"type": "continuous"},
    "scaling": {"min": 0.5, "max": 5},
    "selectionWidth": 2
  },
  "interaction": {
    "hover": true,
    "tooltipDelay": 80,
    "navigationButtons": true,
    "keyboard": true,
    "multiselect": true
  }
}
""")

net_html = OUT / "_pyvis_raw.html"
net.write_html(str(net_html), notebook=False, open_browser=False)

# ---------- 8. Envolver com header customizado (legenda, busca, filtros) ----------
raw = net_html.read_text(encoding="utf-8")

# extrair o trecho do <body> do pyvis
import re as _re
body_match = _re.search(r"<body[^>]*>(.*)</body>", raw, _re.DOTALL)
body_inner = body_match.group(1) if body_match else raw

# stats
n_total = len(records)
n_connected = len(results)
n_bridges = len(bridges)
n_b3 = sum(1 for r in results if len(r["scores"]) >= 3)
n_b4 = sum(1 for r in results if len(r["scores"]) == 4)

sector_counts = Counter()
for r in results:
    sector_counts[r["stk"]["sector"] or "Others"] += 1

# Top 5 por parceiro
top_by_partner = {}
for pid, p in PARTNERS.items():
    lst = [(r["stk"], r["scores"][pid]["score"])
           for r in results if pid in r["scores"]]
    lst.sort(key=lambda x: -x[1])
    top_by_partner[pid] = lst[:10]

legend_html = ""
for sector, col in SECTOR_COLOR.items():
    legend_html += f'<div class="legend-item"><span class="dot" style="background:{col}"></span>{sector} ({sector_counts.get(sector, 0)})</div>'
legend_html += f'<div class="legend-item"><span class="dot star" style="background:{PARTNER_COLOR}"></span>Parceiro MBO (Trias)</div>'

# Top stakeholders panel
top_html = ""
for pid, lst in top_by_partner.items():
    pname = PARTNERS[pid]["name"]
    top_html += f'<details class="top-block"><summary><b>{pname}</b> — {len([r for r in results if pid in r["scores"]])} conexões</summary>'
    top_html += '<ol class="top-list">'
    for stk, sc in lst[:10]:
        top_html += f'<li><b>{stk["name"]}</b> <span class="badge">score {sc}</span><br><span class="muted">{stk["type"]} · {stk["territory"]}</span></li>'
    top_html += '</ol></details>'

# Bridges panel (conectados a 2+)
bridges_html = ""
bridges_sorted = sorted(bridges, key=lambda r: (-len(r["scores"]), -sum(s["score"] for s in r["scores"].values())))
for r in bridges_sorted[:25]:
    stk = r["stk"]
    connections = " · ".join(PARTNERS[p]["name"] for p in r["scores"])
    bridges_html += f'<li><b>{stk["name"]}</b> <span class="badge">{len(r["scores"])}× </span><br><span class="muted">→ {connections}</span></li>'

# HTML final
final_html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Ecossistema Trias Brasil DGD 2027-2031 — Rede de Stakeholders</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  * {{ box-sizing: border-box; }}
  html, body {{ margin: 0; padding: 0; height: 100%; font-family: 'Inter', sans-serif; background: #0f1419; color: #e6e6e6; overflow: hidden; }}
  #app {{ display: grid; grid-template-columns: 320px 1fr 340px; height: 100vh; }}
  #left, #right {{ background: #161d26; border-right: 1px solid #243040; overflow-y: auto; padding: 18px 16px; }}
  #right {{ border-right: none; border-left: 1px solid #243040; }}
  #center {{ position: relative; background: #0f1419; }}
  #mynetwork {{ height: 100% !important; width: 100% !important; }}
  h1 {{ font-size: 17px; margin: 0 0 4px 0; font-weight: 600; letter-spacing: -0.2px; }}
  h2 {{ font-size: 12px; text-transform: uppercase; letter-spacing: 1px; color: #7a8a9a; margin: 22px 0 8px 0; font-weight: 600; }}
  .subtitle {{ font-size: 11px; color: #7a8a9a; margin-bottom: 12px; }}
  .stats {{ display: grid; grid-template-columns: 1fr 1fr; gap: 8px; margin-bottom: 4px; }}
  .stat {{ background: #1c2530; border: 1px solid #243040; border-radius: 8px; padding: 10px; }}
  .stat .n {{ font-size: 22px; font-weight: 700; color: #fff; line-height: 1; }}
  .stat .l {{ font-size: 10px; color: #7a8a9a; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 4px; }}
  .legend-item {{ display: flex; align-items: center; gap: 8px; font-size: 12px; padding: 4px 0; }}
  .dot {{ width: 12px; height: 12px; border-radius: 50%; flex-shrink: 0; }}
  .dot.star {{ width: 14px; height: 14px; border-radius: 2px; transform: rotate(45deg); }}
  .top-block {{ background: #1c2530; border: 1px solid #243040; border-radius: 8px; padding: 10px 12px; margin-bottom: 8px; }}
  .top-block summary {{ cursor: pointer; font-size: 12px; outline: none; user-select: none; }}
  .top-block summary b {{ color: #fff; }}
  .top-list {{ padding-left: 18px; margin: 10px 0 4px 0; font-size: 11.5px; line-height: 1.5; }}
  .top-list li {{ margin-bottom: 8px; }}
  .badge {{ background: #2a3a4d; color: #b9c6d2; font-size: 10px; padding: 1px 6px; border-radius: 6px; font-weight: 600; }}
  .muted {{ color: #7a8a9a; font-size: 11px; }}
  #search {{ width: 100%; padding: 8px 10px; border-radius: 6px; border: 1px solid #243040; background: #0f1419; color: #e6e6e6; font-size: 12px; font-family: inherit; margin-bottom: 12px; }}
  #search:focus {{ outline: none; border-color: #4E79A7; }}
  .filter-btn {{ display: inline-block; background: #1c2530; border: 1px solid #243040; color: #b9c6d2; padding: 5px 10px; border-radius: 6px; font-size: 11px; cursor: pointer; margin: 0 4px 4px 0; user-select: none; }}
  .filter-btn:hover {{ background: #243040; }}
  .filter-btn.active {{ background: #4E79A7; color: #fff; border-color: #4E79A7; }}
  .bridges-list {{ list-style: none; padding: 0; margin: 0; font-size: 11.5px; line-height: 1.5; }}
  .bridges-list li {{ background: #1c2530; border: 1px solid #243040; border-radius: 6px; padding: 8px 10px; margin-bottom: 6px; }}
  .header-bar {{ padding: 14px 20px; background: #161d26; border-bottom: 1px solid #243040; display: flex; align-items: center; justify-content: space-between; }}
  .header-bar h1 {{ font-size: 15px; }}
  .header-bar .meta {{ font-size: 11px; color: #7a8a9a; }}
  .vis-network {{ outline: none !important; }}
  .vis-network:focus {{ outline: none !important; }}
  details > summary {{ list-style: none; }}
  details > summary::-webkit-details-marker {{ display: none; }}
  details > summary::before {{ content: '▸ '; color: #7a8a9a; }}
  details[open] > summary::before {{ content: '▾ '; }}
  ::-webkit-scrollbar {{ width: 8px; }}
  ::-webkit-scrollbar-track {{ background: #0f1419; }}
  ::-webkit-scrollbar-thumb {{ background: #243040; border-radius: 4px; }}
</style>
</head>
<body>
<div id="app">
  <aside id="left">
    <h1>Ecossistema Trias Brasil</h1>
    <div class="subtitle">DGD 2027-2031 · Rede de stakeholders ancorada nos 4 parceiros MBO</div>

    <div class="stats">
      <div class="stat"><div class="n">{n_total}</div><div class="l">stakeholders na base</div></div>
      <div class="stat"><div class="n">{n_connected}</div><div class="l">conectados</div></div>
      <div class="stat"><div class="n">{n_bridges}</div><div class="l">bridges (2+)</div></div>
      <div class="stat"><div class="n">{n_b4}</div><div class="l">conectam os 4</div></div>
    </div>

    <h2>Busca</h2>
    <input id="search" placeholder="Filtrar por nome…" />

    <h2>Filtro por parceiro</h2>
    <div id="partner-filters">
      <span class="filter-btn active" data-partner="ALL">Todos</span>
      <span class="filter-btn" data-partner="UNICAFES_PA">UNICAFES PA</span>
      <span class="filter-btn" data-partner="UNICAFES_RO">UNICAFES RO</span>
      <span class="filter-btn" data-partner="CSA_BRASIL">CSA Brasil</span>
      <span class="filter-btn" data-partner="UNICATADORES">UNICATADORES</span>
    </div>

    <h2>Legenda</h2>
    {legend_html}

    <h2>Como ler</h2>
    <div style="font-size: 11.5px; line-height: 1.5; color: #b9c6d2;">
      As estrelas vermelhas são os 4 parceiros MBO. Os demais nós (círculos) são stakeholders do ecossistema brasileiro, classificados por setor.
      <br><br>
      O <b>tamanho</b> do nó indica quantos parceiros ele conecta (bridges aparecem maiores). A <b>espessura</b> da aresta reflete o score de alinhamento temático/territorial.
      <br><br>
      <i>Passe o mouse sobre um nó para ver detalhes. Clique para destacar a vizinhança.</i>
    </div>
  </aside>

  <main id="center">
    {body_inner}
  </main>

  <aside id="right">
    <h2>Bridges — atores que conectam múltiplos parceiros</h2>
    <ul class="bridges-list">
      {bridges_html}
    </ul>

    <h2>Top atores por parceiro</h2>
    {top_html}
  </aside>
</div>

<script>
// Search filter
document.getElementById('search').addEventListener('input', (e) => {{
  const q = e.target.value.toLowerCase();
  if (!window.network) return;
  const allIds = window.network.body.data.nodes.getIds();
  if (!q) {{
    window.network.body.data.nodes.update(allIds.map(id => ({{id, hidden: false}})));
    return;
  }}
  const updates = allIds.map(id => {{
    const node = window.network.body.data.nodes.get(id);
    const label = (node.label || '').toLowerCase();
    const title = (node.title || '').toLowerCase();
    const match = label.includes(q) || title.includes(q);
    return {{id, hidden: !match}};
  }});
  window.network.body.data.nodes.update(updates);
}});

// Partner filter
document.querySelectorAll('.filter-btn').forEach(btn => {{
  btn.addEventListener('click', () => {{
    document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    const partner = btn.dataset.partner;
    if (!window.network) return;
    const allNodes = window.network.body.data.nodes.get();
    const allEdges = window.network.body.data.edges.get();
    if (partner === 'ALL') {{
      window.network.body.data.nodes.update(allNodes.map(n => ({{id: n.id, hidden: false}})));
      return;
    }}
    const connected = new Set([partner]);
    allEdges.forEach(e => {{
      if (e.from === partner) connected.add(e.to);
      if (e.to === partner) connected.add(e.from);
    }});
    window.network.body.data.nodes.update(
      allNodes.map(n => ({{id: n.id, hidden: !connected.has(n.id)}}))
    );
  }});
}});
</script>
</body>
</html>
"""

out_path = OUT / "ecossistema_trias_brasil.html"
out_path.write_text(final_html, encoding="utf-8")
print(f"\nHTML gerado: {out_path}")

# ---------- 9. Markdown report ----------
md = [f"# Análise de rede — Ecossistema Trias Brasil DGD 2027-2031\n"]
md.append("Rede de stakeholders ancorada nos 4 parceiros MBO da Trias Brasil "
          "(UNICAFES Pará, UNICAFES Rondônia, CSA Brasil e UNICATADORES). "
          "Fonte: `Stakeholder Ecosystem Mapping.xlsx`.\n")
md.append("## Resumo\n")
md.append(f"- **{n_total}** stakeholders na base de origem\n"
          f"- **{n_connected}** conectados a pelo menos 1 parceiro (score ≥ {THRESHOLD})\n"
          f"- **{n_bridges}** bridges (conectados a 2 ou mais parceiros)\n"
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
          "keyword amazônica o score zera (evita falsos positivos de grandes "
          "empresas de cobertura nacional)\n"
          "- **gate de catadores**: para UNICATADORES, exige pelo menos uma "
          "keyword de resíduos/catador/circular\n"
          f"\nThreshold para criar edge: score ≥ {THRESHOLD}.\n")

md.append("\n## Bridges (conectam múltiplos parceiros)\n")
for r in bridges_sorted[:30]:
    stk = r["stk"]
    cons = ", ".join(PARTNERS[p]["name"] for p in r["scores"])
    md.append(f"- **{stk['name']}** ({stk['type']}) — {len(r['scores'])}× · {cons}")

md.append("\n\n## Top 10 stakeholders por parceiro\n")
for pid, lst in top_by_partner.items():
    md.append(f"\n### {PARTNERS[pid]['name']}\n")
    for stk, sc in lst[:10]:
        md.append(f"- **{stk['name']}** — score {sc} · {stk['type']} · {stk['territory']}")

md.append("\n\n## Top 15 por betweenness centrality (potenciais brokers)\n")
for nid, b in top_bw[:15]:
    label = G.nodes[nid].get("label", nid)
    md.append(f"- {label} — {b:.4f}")

(OUT / "analise_rede.md").write_text("\n".join(md), encoding="utf-8")
print(f"Relatório gerado: {OUT / 'analise_rede.md'}")
