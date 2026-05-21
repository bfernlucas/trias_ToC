"""
Constrói a rede do ecossistema Trias Brasil DGD 2027-2031 e renderiza
visualização interativa em tema claro, com painéis colapsáveis e estatísticas
agregadas. Saída em docs/index.html para servir via GitHub Pages.

Pipeline:
 1. Carrega Stakeholder Ecosystem Mapping.xlsx
 2. Ancora os 4 parceiros MBO com perfis temáticos/territoriais
 3. Calcula score de alinhamento (biome + impact area + role + keywords)
 4. Aplica gates para reduzir falsos positivos (amazônico, catadores)
 5. Computa métricas de rede e distribuições agregadas
 6. Gera HTML autocontido (D3.js) + relatório markdown analítico
"""
from openpyxl import load_workbook
import networkx as nx
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path("/home/user/trias_ToC")
DOCS_SRC = ROOT / "docs-referencia"
OUT_NET = ROOT / "outputs" / "network"
OUT_PAGES = ROOT / "docs"
OUT_NET.mkdir(parents=True, exist_ok=True)
OUT_PAGES.mkdir(parents=True, exist_ok=True)

# ========== 1. Carregar base ==========
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
# Sanitiza menções a IKF/IKEA nos textos livres (notes, description) — a base
# de origem foi construída em outro contexto e referências ao financiador
# original poluem a leitura do mapa Trias.
import re as _re_sanit
_IKF_PAT = _re_sanit.compile(r"\bIKEA\s+Foundation\b|\bIKF\b", _re_sanit.IGNORECASE)
for rec in records:
    rec["notes"] = _IKF_PAT.sub("o financiador analisado", rec["notes"])
    rec["description"] = _IKF_PAT.sub("o financiador analisado", rec["description"])

print(f"Loaded {len(records)} stakeholders.")

# ========== 2. Perfis-âncora dos parceiros ==========
PARTNERS = {
    "UNICAFES_PA": {
        "name": "UNICAFES Pará",
        "long_name": "União das Cooperativas da Agricultura Familiar e Economia Solidária — Pará",
        "maturity": "2º nível · em consolidação",
        "type": "MBO de 2º nível · Amazônia rural",
        "biomes": {"Amazonia"},
        "impacts": {"Land, Food and Forest"},
        "roles": {"Implementation capacity (delivery)",
                  "Community legitimacy / proximate leadership",
                  "Territorial anchor"},
        "keywords": ["agricultura familiar", "cooperativa", "amazon", "bioeconomia",
                     "açaí", "cacau", "cocoa", "coffee", "café", "agroflorest",
                     "sociobiodiversidade", "family farming", "agroecolog",
                     "northern brazil", "pará"],
        "territory": "Pará · Amazônia Legal",
        "units": "13 cooperativas de primeiro nível",
        "members": "1.241 agricultores familiares (FFEs)",
        "women_pct": 29, "youth_pct": 15, "staff": 2,
        "value_chains": "açaí, castanha, mel, cacau, café",
        "strategic_role": ("Âncora estratégica para a bioeconomia pan-amazônica e expansão "
                           "do cooperativismo familiar no Pará. Cacau e café conectam às "
                           "cadeias da Trias SAM em Peru e Equador."),
        "description": ("Articula 13 cooperativas de primeiro nível no estado do Pará. "
                        "Foco em cadeias da sociobiodiversidade (açaí, castanha, mel, "
                        "cacau, café) e bioeconomia amazônica. 1.241 agricultores "
                        "familiares afiliados, 29% mulheres, 15% jovens."),
        "trias_priorities": [
            "Governança cooperativa e fortalecimento de serviços aos membros",
            "Acesso a mercados diferenciados (orgânico, EUDR-compliant, fair trade)",
            "Práticas agroflorestais e agroecológicas",
            "Inclusão de mulheres e jovens nos espaços de decisão"
        ],
    },
    "UNICAFES_RO": {
        "name": "UNICAFES Rondônia",
        "long_name": "União das Cooperativas da Agricultura Familiar e Economia Solidária — Rondônia",
        "maturity": "2º nível · em consolidação · exit 2028",
        "type": "MBO de 2º nível · Amazônia rural · exit 2028",
        "biomes": {"Amazonia"},
        "impacts": {"Land, Food and Forest"},
        "roles": {"Implementation capacity (delivery)",
                  "Community legitimacy / proximate leadership",
                  "Territorial anchor"},
        "keywords": ["agricultura familiar", "cooperativa", "amazon", "bioeconomia",
                     "café", "coffee", "cocoa", "cacau", "agroflorest",
                     "blended finance", "pes", "carbon", "rondônia", "rondonia"],
        "territory": "Rondônia · Amazônia Legal",
        "units": "15-16 cooperativas de primeiro nível",
        "members": "1.265 agricultores familiares (FFEs)",
        "women_pct": 56, "youth_pct": 30, "staff": 2,
        "value_chains": "açaí, castanha, mel, cacau, café",
        "strategic_role": ("Pilotagem de instrumentos de finança climática (blended finance, "
                           "PES, mercados de carbono) na Amazônia. Transição planejada para "
                           "papel de parceria estratégica a partir de 2028, com complementos "
                           "via IKI, GIZ e MDA."),
        "description": ("Coordena 15-16 cooperativas de primeiro nível em Rondônia. Exit "
                        "strategy programado para 2028 (transição gradual 2027-2028). "
                        "1.265 agricultores afiliados, 56% mulheres, 30% jovens — perfil "
                        "de gênero/juventude significativamente mais inclusivo do que PA."),
        "trias_priorities": [
            "Cooperativism expansion (governança e serviços aos cooperados)",
            "Agricultura climate-resilient (agroecológica e agroflorestal)",
            "Finanças inovadoras (blended finance, PES, carbono)",
            "Articulação com IKI, GIZ, MDA para complementaridade pós-2028"
        ],
    },
    "CSA_BRASIL": {
        "name": "CSA Brasil",
        "long_name": "Comunidade que Sustenta a Agricultura — Brasil",
        "maturity": "3º nível · em consolidação",
        "type": "MBO de 3º nível · rural-urbano nacional",
        "biomes": {"Mata Atlântica", "Cerrado", "Amazonia"},
        "impacts": {"Land, Food and Forest"},
        "roles": {"Implementation capacity (delivery)",
                  "Community legitimacy / proximate leadership",
                  "Convenor/enablers"},
        "keywords": ["community supported", "agroecolog", "food system",
                     "sistema alimentar", "saudáve", "food and nutrition",
                     "agricultura urbana", "peri-urb", "agricultor familiar",
                     "organic", "orgânico", "consumer", "rural-urban", "solidari"],
        "territory": "Nacional · 19 estados · 5 regiões · rural-urbano",
        "units": "200 unidades CSA",
        "members": "540 FFEs + 3.600 NFEs (co-agricultores)",
        "women_pct": 70, "youth_pct": 30, "staff": 3,
        "value_chains": "agroecologia, cestas saudáveis, educação alimentar",
        "strategic_role": ("Modelo nacional para conectar produção agroecológica a consumidores "
                           "em centros urbanos. Membro do Urgenci (rede internacional CSA). "
                           "Foco de expansão da Trias: levar o modelo CSA para a região Norte."),
        "description": ("Rede nacional de 200 unidades CSA em 19 estados (5 regiões). 540 "
                        "agricultores familiares + 3.600 co-agricultores urbanos. 70% "
                        "mulheres — o perfil mais feminilizado dos 4 parceiros. Membro ativo "
                        "do Urgenci, conectando experiências brasileiras a movimentos globais."),
        "trias_priorities": [
            "Expansão do modelo CSA para a região Norte (gap territorial)",
            "Educação alimentar e cadeias curtas peri-urbanas",
            "Conexão de produtores agroecológicos a co-agricultores urbanos",
            "Liderança feminina (70% das beneficiárias)"
        ],
    },
    "UNICATADORES": {
        "name": "UNICATADORES",
        "long_name": "União Nacional de Catadoras e Catadores de Materiais Recicláveis",
        "maturity": "3º nível · maduro",
        "type": "MBO de 3º nível · urbano nacional",
        "biomes": set(),
        "impacts": {"Land, Food and Forest", "Buildings & Transport"},
        "roles": {"Implementation capacity (delivery)",
                  "Policy influence / regulation",
                  "Community legitimacy / proximate leadership"},
        "keywords": ["catador", "waste pick", "recicl", "circular econom",
                     "resíduo sólido", "residuos solidos", "waste management",
                     "pnrs", "logística reversa", "extended producer", "epr",
                     "lixo", "urban"],
        "territory": "Nacional · 26 estados · urbano",
        "units": "230 cooperativas de catadores",
        "members": "~50.000 catadores organizados (NFEs)",
        "women_pct": 60, "youth_pct": 30, "staff": 28,
        "value_chains": "papel, plástico, metais, eletrônicos, vidro",
        "strategic_role": ("Única MBO madura entre os 4 parceiros. Operação policy-driven, "
                           "ancorada no PNRS (Política Nacional de Resíduos Sólidos) e no "
                           "Comitê Interministerial CIISC. Equipe de 28 funcionários e "
                           "capacidade própria significativa — Trias atua como parceiro "
                           "estratégico, não como organizational developer."),
        "description": ("Federação nacional de 230 cooperativas em 26 estados. ~50.000 "
                        "catadores organizados, 60% mulheres. Sede em São Paulo, 28 "
                        "funcionários. Participa do CIISC e dos fóruns da PNRS. Foco em "
                        "advocacy, capacitação de liderança jovem e feminina, conectividade "
                        "digital e resiliência climática urbana."),
        "trias_priorities": [
            "Pilotos de PES e mercados de carbono para serviços ecossistêmicos urbanos",
            "Logística reversa, EPR (Extended Producer Responsibility), contratação municipal",
            "Capacitação de liderança jovem e feminina",
            "Conectividade digital e ferramentas de traceabilidade"
        ],
    },
}

# ========== 3. Score de alinhamento ==========
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
                breakdown.append({"k": "biome (primário)", "v": tk, "pts": 3})
                biome_hit = True; matched = True; break
        if not matched:
            for tk in b_secondary:
                if biome.lower() in tk.lower():
                    sc += 1
                    breakdown.append({"k": "biome (secundário)", "v": tk, "pts": 1})
                    biome_hit = True; break
    impacts = tokenize(stk["impact_area"])
    for imp in p["impacts"]:
        for tk in impacts:
            if imp.lower() in tk.lower() or tk.lower() in imp.lower():
                sc += 1
                breakdown.append({"k": "impact area", "v": tk, "pts": 1})
                break
    roles = tokenize(stk["role"])
    role_hits = []
    for r in p["roles"]:
        for tk in roles:
            if r.lower() == tk.lower():
                role_hits.append(tk); break
    if role_hits:
        pts = min(len(role_hits), 2) * 0.5
        sc += pts
        breakdown.append({"k": "role", "v": ", ".join(role_hits), "pts": pts})
    blob = (stk["description"] + " " + stk["notes"] + " " + stk["type"] +
            " " + stk["territory"] + " " + stk["name"]).lower()
    kw_hits = [kw for kw in p["keywords"] if kw.lower() in blob]
    if kw_hits:
        pts = min(len(kw_hits), 4) * 2
        sc += pts
        breakdown.append({"k": "keywords", "v": ", ".join(kw_hits[:4]), "pts": pts})
    if "national" in stk["territory"].lower():
        sc += 0.5
        breakdown.append({"k": "alcance", "v": "nacional", "pts": 0.5})
    if p["biomes"] == {"Amazonia"} and not biome_hit and not kw_hits:
        return 0, []
    if p["name"] == "UNICATADORES" and not kw_hits:
        return 0, []
    return round(sc, 1), breakdown

THRESHOLD = 4.5
results = []
for stk in records:
    scores = {}
    for pid, p in PARTNERS.items():
        sc, br = score(stk, p)
        if sc >= THRESHOLD:
            scores[pid] = {"score": sc, "breakdown": br}
    if scores:
        results.append({"stk": stk, "scores": scores})
print(f"Conectados (score >= {THRESHOLD}): {len(results)}")

# ========== 3a. Camada institucional Trias ==========
# Atores que compõem a base institucional, programática e operacional da Trias
# globalmente — financiadores históricos, redes belgas/europeias, agri-agências
# irmãs da AgriCord, parceiros diplomáticos e corporativos. Fontes: trias.ngo
# (Worldwide, South America, East Africa), Annual Report 2024, AgriCord.org,
# Annex 1 e Narrative DGD. Estes são vínculos *existentes*, não potenciais —
# por isso entram como camada separada, sem entrar no cálculo de adicionalidade.
ALL_MBOS = ["UNICAFES_PA", "UNICAFES_RO", "CSA_BRASIL", "UNICATADORES"]
AMAZON_MBOS = ["UNICAFES_PA", "UNICAFES_RO"]

TRIAS_INSTITUTIONAL = {
    # ---- Doadores institucionais ----
    "T_DGD": {"name": "DGD", "long_name": "Direction-Générale de la Coopération au Développement (Bélgica)",
              "subcat": "Doador institucional", "type": "Agência de cooperação bilateral",
              "country": "Bélgica",
              "role": "Principal financiador do programa Trias DGD 2027-2031",
              "mbo_relations": {
                  "UNICAFES_PA": "Financia o programa DGD 2027-2031, do qual a MBO é parceira central",
                  "UNICAFES_RO": "Financia o programa DGD 2027-2031 incluindo o exit pathway 2028",
                  "CSA_BRASIL": "Financia o programa DGD 2027-2031, pilar rural-urbano",
                  "UNICATADORES": "Financia o programa DGD 2027-2031, pilar economia circular",
              }},
    "T_EU":  {"name": "Comissão Europeia", "long_name": "European Commission · DG INTPA",
              "subcat": "Doador institucional", "type": "Multilateral",
              "country": "União Europeia",
              "role": "Acordo EU-Mercosur, EUDR (deforestation regulation), DG INTPA",
              "mbo_relations": {
                  "UNICAFES_PA": "EUDR cria barreira/oportunidade direta para cacau e café amazônicos",
                  "UNICAFES_RO": "EUDR aplicada às cadeias de cacau e café de Rondônia",
                  "CSA_BRASIL": "EU-Mercosur Agreement afeta sistemas alimentares e exportações",
              }},
    "T_ENABEL": {"name": "Enabel", "long_name": "Agência Belga de Cooperação para o Desenvolvimento",
                 "subcat": "Doador institucional", "type": "Agência de cooperação bilateral",
                 "country": "Bélgica",
                 "role": "Cooperação técnica complementar à DGD, foco em programas bilaterais",
                 "mbo_relations": {
                     "UNICAFES_PA": "Complementaridade técnica para programas amazônicos belgas",
                     "UNICAFES_RO": "Complementaridade técnica para programas amazônicos belgas",
                 }},
    "T_WB":  {"name": "World Bank", "long_name": "Banco Mundial",
              "subcat": "Doador institucional", "type": "Multilateral · banco de desenvolvimento",
              "country": "EUA",
              "role": "Sustainable Landscapes Amazonia, financiamento climático",
              "mbo_relations": {
                  "UNICAFES_PA": "Sustainable Landscapes Amazonia (PA é estado-foco)",
                  "UNICAFES_RO": "Sustainable Landscapes Amazonia (RO é estado-foco)",
              }},
    "T_AFD": {"name": "AFD", "long_name": "Agence Française de Développement",
              "subcat": "Doador institucional", "type": "Agência de cooperação bilateral",
              "country": "França",
              "role": "Co-funding em SAM, voluntários AFD nos parceiros brasileiros",
              "mbo_relations": {
                  "UNICAFES_PA": "Voluntários AFD em SAM (mencionado no Narrative)",
                  "UNICAFES_RO": "Voluntários AFD em SAM, perfil cooperativista",
                  "CSA_BRASIL": "Voluntários AFD em CSA (modelo educacional francês)",
              }},
    "T_IKI": {"name": "IKI", "long_name": "Internationale Klimaschutzinitiative (BMUV, Alemanha)",
              "subcat": "Doador institucional", "type": "Fundo climático bilateral",
              "country": "Alemanha",
              "role": "Iniciativa Internacional de Proteção Climática alemã",
              "mbo_relations": {
                  "UNICAFES_RO": "Mencionado no Annex 1 como complementaridade pós-2028 (instrumentos climáticos)",
                  "UNICAFES_PA": "Potencial expansão de finanças climáticas para PA",
              }},
    "T_GSTIC": {"name": "GSTIC", "long_name": "Global Sustainable Tech & Innovation Centre",
                "subcat": "Doador institucional", "type": "Multilateral · tecnologia",
                "country": "Bélgica",
                "role": "Co-funding em tecnologia e inovação para sustentabilidade",
                "mbo_relations": {
                    "CSA_BRASIL": "Conectividade digital e tecnologia para sistemas alimentares",
                    "UNICATADORES": "Ferramentas digitais para traceabilidade reciclagem (AI/MRV)",
                }},
    "T_MDA": {"name": "MDA", "long_name": "Ministério do Desenvolvimento Agrário e Agricultura Familiar",
              "subcat": "Doador institucional", "type": "Agência pública federal",
              "country": "Brasil",
              "role": "PRONAF, PAA, política nacional de agricultura familiar",
              "mbo_relations": {
                  "UNICAFES_PA": "PRONAF Amazônia, crédito para agricultura familiar PA",
                  "UNICAFES_RO": "PRONAF Amazônia, complementaridade pós-2028 (Annex 1)",
                  "CSA_BRASIL": "PNAE, PAA — políticas de compras públicas alimentares",
              }},

    # ---- Agri-agências da AgriCord ----
    "T_AGRICORD": {"name": "AgriCord", "long_name": "AgriCord — Aliança global de agri-agências",
                   "subcat": "Rede peer (AgriCord)", "type": "Aliança global",
                   "country": "Bélgica (secretariado)",
                   "role": "Aliança da qual Trias é membro-fundador via Boerenbond",
                   "mbo_relations": {
                       "UNICAFES_PA": "Knowledge management e peer-learning AgriCord",
                       "UNICAFES_RO": "Knowledge management e peer-learning AgriCord",
                       "CSA_BRASIL": "Knowledge management e peer-learning AgriCord",
                       "UNICATADORES": "Knowledge management AgriCord (caso atípico, urbano)",
                   }},
    "T_AGRITERRA": {"name": "Agriterra", "long_name": "Agriterra — Cooperative Development Agency",
                    "subcat": "Rede peer (AgriCord)", "type": "Agri-agência",
                    "country": "Holanda",
                    "role": "Company Assessment é referência metodológica do Diagnóstico Técnico ToC",
                    "mbo_relations": {
                        "UNICAFES_PA": "Company Assessment aplicável a MBO 2º nível em consolidação",
                        "UNICAFES_RO": "Company Assessment aplicável a MBO 2º nível em consolidação",
                        "CSA_BRASIL": "Toolkit de governança aplicável a MBO 3º nível",
                    }},
    "T_CRESOL_AA": {"name": "Cresol Agri-Agency", "long_name": "Cresol Agri-Agency (Sistema Cresol)",
                    "subcat": "Rede peer (AgriCord)", "type": "Agri-agência",
                    "country": "Brasil",
                    "role": "Única agri-agência brasileira da AgriCord; parceria histórica Trias",
                    "mbo_relations": {
                        "UNICAFES_PA": "Crédito rural especializado, sistema irmão UNICAFES",
                        "UNICAFES_RO": "Crédito rural especializado, sistema irmão UNICAFES",
                        "CSA_BRASIL": "Cresol financia agricultores fornecedores de unidades CSA",
                        "UNICATADORES": "Possível parceria em microcrédito para cooperativas urbanas",
                    }},
    "T_FFD": {"name": "FFD", "long_name": "Food and Forest Development Finland",
              "subcat": "Rede peer (AgriCord)", "type": "Agri-agência",
              "country": "Finlândia",
              "role": "Mandato florestal; ToC inspirada pelo Forest and Farm Facility",
              "mbo_relations": {
                  "UNICAFES_PA": "Forest/farm — sociobiodiversidade amazônica direta",
                  "UNICAFES_RO": "Forest/farm — agroflorestal e bioeconomia",
              }},
    "T_SOLIDARIDAD": {"name": "Solidaridad", "long_name": "Solidaridad Network",
                      "subcat": "Rede peer (AgriCord)", "type": "ONG internacional",
                      "country": "Holanda",
                      "role": "Cadeias sustentáveis; certificação RTRS (Cresol/Colruyt 2017)",
                      "mbo_relations": {
                          "UNICAFES_PA": "Certificação de cadeias de cacau e café",
                          "UNICAFES_RO": "Certificação de cadeias de cacau e café",
                          "CSA_BRASIL": "Padrões de sustentabilidade em sistemas alimentares",
                      }},
    "T_CSA_BE": {"name": "CSA (Bélgica)", "long_name": "Collectif Stratégies Alimentaires",
                 "subcat": "Rede peer (AgriCord)", "type": "Agri-agência",
                 "country": "Bélgica",
                 "role": "Sistemas alimentares; peer temático direto com CSA Brasil",
                 "mbo_relations": {
                     "CSA_BRASIL": "Peer temático direto — sistemas alimentares e cooperação",
                 }},
    "T_FERT": {"name": "Fert", "long_name": "Fert — Coopération internationale agricole",
               "subcat": "Rede peer (AgriCord)", "type": "Agri-agência",
               "country": "França",
               "role": "Cooperativismo agrícola francês",
               "mbo_relations": {
                   "UNICAFES_PA": "Peer-learning entre cooperativismo francês e brasileiro",
                   "UNICAFES_RO": "Peer-learning entre cooperativismo francês e brasileiro",
               }},
    "T_UPA_DI": {"name": "UPA DI", "long_name": "UPA Développement International",
                 "subcat": "Rede peer (AgriCord)", "type": "Agri-agência",
                 "country": "Quebec, Canadá",
                 "role": "Cooperativismo agrícola das Américas",
                 "mbo_relations": {
                     "UNICAFES_PA": "Intercâmbio cooperativo americano-canadense",
                     "UNICAFES_RO": "Intercâmbio cooperativo americano-canadense",
                 }},
    "T_WE_EFFECT": {"name": "We Effect", "long_name": "We Effect (former Swedish Cooperative Centre)",
                    "subcat": "Rede peer (AgriCord)", "type": "Agri-agência",
                    "country": "Suécia",
                    "role": "Direitos rurais, gênero, cooperativismo",
                    "mbo_relations": {
                        "UNICAFES_PA": "Agenda de gênero em cooperativismo rural (29% mulheres em PA é prioridade)",
                        "UNICAFES_RO": "Agenda de gênero em cooperativismo rural",
                        "CSA_BRASIL": "Gênero como driver (70% mulheres na rede)",
                        "UNICATADORES": "Gênero em cooperativas urbanas (60% mulheres catadoras)",
                    }},
    "T_ASPRODEB": {"name": "Asprodeb", "long_name": "Asprodeb — Senegal",
                   "subcat": "Rede peer (AgriCord)", "type": "Agri-agência",
                   "country": "Senegal",
                   "role": "Peer south-south em agricultura familiar",
                   "mbo_relations": {
                       "UNICAFES_PA": "Peer south-south agricultura familiar",
                       "UNICAFES_RO": "Peer south-south agricultura familiar",
                   }},
    "T_AHA":      {"name": "AHA", "long_name": "AHA — Agri-agency",
                   "subcat": "Rede peer (AgriCord)", "type": "Agri-agência",
                   "country": "Multipaís",
                   "role": "Peer-learning AgriCord",
                   "mbo_relations": {
                       "CSA_BRASIL": "Peer-learning multipaís de food systems",
                   }},
    "T_ASIADHRRA": {"name": "AsiaDHRRA", "long_name": "AsiaDHRRA — Asian Partnership",
                    "subcat": "Rede peer (AgriCord)", "type": "Agri-agência",
                    "country": "Filipinas (regional)",
                    "role": "Rede asiática de cooperativismo e capacitação rural",
                    "mbo_relations": {
                        "CSA_BRASIL": "Peer asiática em sistemas alimentares",
                    }},

    # ---- Redes belgas/europeias ----
    "T_11_11":  {"name": "11.11.11", "long_name": "11.11.11 · Umbrella belga de solidariedade internacional",
                 "subcat": "Rede belga/europeia", "type": "Federação de ONGs",
                 "country": "Bélgica",
                 "role": "Articulação política das ONGs belgas; Trias é membro",
                 "mbo_relations": {
                     "UNICAFES_PA": "Voz coletiva belga inclui a operação PA",
                     "UNICAFES_RO": "Voz coletiva belga inclui a operação RO",
                     "CSA_BRASIL": "Voz coletiva belga inclui o pilar rural-urbano",
                     "UNICATADORES": "Voz coletiva belga inclui o pilar economia circular",
                 }},
    "T_NGO_FED": {"name": "NGO Federation BE", "long_name": "Federação Belga de ONGs (ngo-federatie)",
                  "subcat": "Rede belga/europeia", "type": "Federação de ONGs",
                  "country": "Bélgica",
                  "role": "Diálogo com DGD; harmonização de procedimentos",
                  "mbo_relations": {
                      "UNICAFES_PA": "Procedimentos administrativos DGD",
                      "UNICAFES_RO": "Procedimentos administrativos DGD",
                      "CSA_BRASIL": "Procedimentos administrativos DGD",
                      "UNICATADORES": "Procedimentos administrativos DGD",
                  }},
    "T_CCH":    {"name": "Coalition Against Hunger", "long_name": "Coalition Against Hunger (CCH) Bélgica",
                 "subcat": "Rede belga/europeia", "type": "Coalizão temática",
                 "country": "Bélgica",
                 "role": "Sistemas alimentares sustentáveis; Trias planeja rejoin 2026",
                 "mbo_relations": {
                     "UNICAFES_PA": "Advocacy belga para sistemas alimentares amazônicos",
                     "UNICAFES_RO": "Advocacy belga para sistemas alimentares amazônicos",
                     "CSA_BRASIL": "CSA é diretamente sistemas alimentares — alinhamento natural",
                 }},
    "T_BEYOND_CHOC": {"name": "Beyond Chocolate", "long_name": "Beyond Chocolate · Plataforma belga para cacau sustentável",
                      "subcat": "Rede belga/europeia", "type": "Plataforma setorial",
                      "country": "Bélgica",
                      "role": "Plataforma belga para cacau sustentável",
                      "mbo_relations": {
                          "UNICAFES_PA": "Cacau é cadeia central em PA",
                          "UNICAFES_RO": "Cacau é cadeia central em RO",
                      }},
    "T_BASCOF":  {"name": "BASCOF", "long_name": "Belgian Sustainable Coffee Initiative",
                  "subcat": "Rede belga/europeia", "type": "Plataforma setorial",
                  "country": "Bélgica",
                  "role": "Iniciativa belga em formação para café sustentável",
                  "mbo_relations": {
                      "UNICAFES_PA": "Café é cadeia central em PA",
                      "UNICAFES_RO": "Café é cadeia central em RO",
                  }},
    "T_URGENCI": {"name": "Urgenci", "long_name": "Urgenci · International CSA Network",
                  "subcat": "Rede belga/europeia", "type": "Rede temática internacional",
                  "country": "Internacional",
                  "role": "Rede internacional de Community-Supported Agriculture",
                  "mbo_relations": {
                      "CSA_BRASIL": "CSA Brasil é membro ativo da rede internacional",
                  }},
    "T_UN_GC":   {"name": "UN Global Compact", "long_name": "UN Global Compact",
                  "subcat": "Rede belga/europeia", "type": "Pacto global",
                  "country": "ONU",
                  "role": "Trias é signatária; due diligence ética de parceiros",
                  "mbo_relations": {
                      "UNICAFES_PA": "Due diligence ética aplicada a parceiros do programa",
                      "UNICAFES_RO": "Due diligence ética aplicada a parceiros do programa",
                      "CSA_BRASIL": "Due diligence ética aplicada a parceiros do programa",
                      "UNICATADORES": "Due diligence ética aplicada a parceiros do programa",
                  }},

    # ---- Diplomacia e mercado Bélgica-Brasil ----
    "T_EMB_BE":  {"name": "Embaixada da Bélgica", "long_name": "Embaixada da Bélgica em Brasília",
                  "subcat": "Diplomacia/mercado", "type": "Diplomático",
                  "country": "Bélgica/Brasil",
                  "role": "Coordena agenda bilateral e missões econômicas",
                  "mbo_relations": {
                      "UNICAFES_PA": "Apoio diplomático em iniciativas amazônicas",
                      "UNICAFES_RO": "Apoio diplomático em iniciativas amazônicas",
                      "CSA_BRASIL": "Apoio diplomático em iniciativas alimentares",
                      "UNICATADORES": "Apoio diplomático em iniciativas urbanas",
                  }},
    "T_AWEX":    {"name": "AWEX", "long_name": "Wallonia Export and Investment Agency",
                  "subcat": "Diplomacia/mercado", "type": "Agência de comércio",
                  "country": "Bélgica (Valônia)",
                  "role": "Parceira do Gastronomy Lab Santarém-PA (com BID e Liège)",
                  "mbo_relations": {
                      "UNICAFES_PA": "Gastronomy Lab Santarém-PA é parceria direta AWEX-BID-Liège",
                  }},
    "T_HUB_BR":  {"name": "Hub Brussels", "long_name": "Hub Brussels · Brussels Export",
                  "subcat": "Diplomacia/mercado", "type": "Agência de comércio",
                  "country": "Bélgica (Bruxelas)",
                  "role": "Comércio bilateral Bruxelas-Brasil",
                  "mbo_relations": {
                      "CSA_BRASIL": "Comércio bilateral de produtos agroecológicos",
                  }},
    "T_FIT":     {"name": "FIT", "long_name": "Flanders Investment & Trade",
                  "subcat": "Diplomacia/mercado", "type": "Agência de comércio",
                  "country": "Bélgica (Flandres)",
                  "role": "Colruyt é empresa flamenga, foco em cadeias amazônicas",
                  "mbo_relations": {
                      "UNICAFES_PA": "Comércio Flandres-Amazônia (Colruyt mel, cacau)",
                      "UNICAFES_RO": "Comércio Flandres-Amazônia (cacau, café)",
                  }},
    "T_BELGALUX": {"name": "Belgalux", "long_name": "Belgalux · Câmara comercial Bélgica-Brasil",
                   "subcat": "Diplomacia/mercado", "type": "Câmara de comércio",
                   "country": "Bélgica/Brasil",
                   "role": "Câmara comercial bilateral",
                   "mbo_relations": {
                       "UNICAFES_PA": "Câmara bilateral suporta cadeias exportadoras",
                       "UNICAFES_RO": "Câmara bilateral suporta cadeias exportadoras",
                       "CSA_BRASIL": "Câmara bilateral suporta produtos agroecológicos",
                   }},
    "T_BEM_2024": {"name": "Belgian Economic Mission 2024", "long_name": "Belgian Economic Mission to Brazil (Princesa Astrid, nov/2024)",
                   "subcat": "Diplomacia/mercado", "type": "Missão diplomática (histórica)",
                   "country": "Bélgica/Brasil",
                   "role": "405 participantes, 173 empresas; marco histórico Trias-Brasil",
                   "mbo_relations": {
                       "UNICAFES_PA": "Trias participou da missão e MoU Suzano referencia PA",
                       "UNICAFES_RO": "Trias participou da missão",
                       "CSA_BRASIL": "Trias participou da missão",
                       "UNICATADORES": "Trias participou da missão",
                   }},
    "T_ABC":     {"name": "ABC", "long_name": "Agência Brasileira de Cooperação · MRE",
                  "subcat": "Diplomacia/mercado", "type": "Agência pública federal",
                  "country": "Brasil",
                  "role": "Cooperação Sul-Sul",
                  "mbo_relations": {
                      "UNICAFES_PA": "Sul-Sul no corredor amazônico trinacional",
                      "UNICAFES_RO": "Sul-Sul no corredor amazônico trinacional",
                  }},
    "T_ITAMARATY": {"name": "Itamaraty", "long_name": "Ministério das Relações Exteriores",
                    "subcat": "Diplomacia/mercado", "type": "Agência pública federal",
                    "country": "Brasil",
                    "role": "Diplomacia e cooperação internacional",
                    "mbo_relations": {
                        "UNICAFES_PA": "Diplomacia amazônica (COP30, NDC, sociobiodiversidade)",
                        "UNICAFES_RO": "Diplomacia amazônica e mercados de carbono",
                        "CSA_BRASIL": "Soberania alimentar nos fóruns internacionais",
                        "UNICATADORES": "Política internacional de resíduos e economia circular",
                    }},
    "T_SAF":     {"name": "SAF/MDA", "long_name": "Secretaria de Agricultura Familiar · MDA",
                  "subcat": "Diplomacia/mercado", "type": "Agência pública federal",
                  "country": "Brasil",
                  "role": "Política nacional de agricultura familiar (PRONAF, PAA, PNAE)",
                  "mbo_relations": {
                      "UNICAFES_PA": "PRONAF e acesso a políticas de AF",
                      "UNICAFES_RO": "PRONAF e acesso a políticas de AF",
                      "CSA_BRASIL": "PNAE, PAA — compras públicas de agricultura familiar",
                  }},

    # ---- Setor privado parceiro + MBOs SAM ----
    "T_COLRUYT": {"name": "Colruyt Group", "long_name": "Colruyt Group · varejista belga",
                  "subcat": "Setor privado parceiro", "type": "Varejo",
                  "country": "Bélgica",
                  "role": "Mel orgânico Coopemapi 2024-2027; soja Cresol 2017; parceria histórica",
                  "mbo_relations": {
                      "UNICAFES_PA": "Off-take de mel orgânico (cooperativa Coopemapi na Amazônia)",
                      "UNICAFES_RO": "Off-take potencial de cacau e café orgânicos",
                      "CSA_BRASIL": "Modelo de varejo orgânico belga, peer para CSA",
                  }},
    "T_BOERENBOND": {"name": "Boerenbond", "long_name": "Boerenbond · Belgian Farmers' Association",
                     "subcat": "Setor privado parceiro", "type": "Associação de produtores",
                     "country": "Bélgica",
                     "role": "Co-fundadora da Trias; mandante AgriCord",
                     "mbo_relations": {
                         "UNICAFES_PA": "Origem institucional da Trias enquanto agri-agência",
                         "UNICAFES_RO": "Origem institucional da Trias",
                         "CSA_BRASIL": "Origem institucional da Trias",
                         "UNICATADORES": "Origem institucional da Trias",
                     }},
    "T_COOPEMAPI": {"name": "Coopemapi", "long_name": "Coopemapi · Cooperativa de Apicultores Amazônicos",
                    "subcat": "Setor privado parceiro", "type": "Cooperativa de produtores",
                    "country": "Brasil (Amazônia)",
                    "role": "Fornecedora do mel orgânico para Colruyt 2024-2027",
                    "mbo_relations": {
                        "UNICAFES_PA": "Cooperativa amazônica, peer direto de UNICAFES PA na cadeia do mel",
                        "UNICAFES_RO": "Modelo replicável para cooperativas RO",
                    }},
    "T_KALLARI": {"name": "Kallari", "long_name": "Asociación Kallari · Cooperativa de cacao",
                  "subcat": "Setor privado parceiro · SAM", "type": "Cooperativa de produtores",
                  "country": "Equador (Amazônia)",
                  "role": "Cacau orgânico; peer regional no corredor amazônico trinacional",
                  "mbo_relations": {
                      "UNICAFES_PA": "Peer trinacional, cadeia do cacau amazônico",
                      "UNICAFES_RO": "Peer trinacional, cadeia do cacau amazônico",
                  }},
    "T_UNOCACE": {"name": "Unocace", "long_name": "Unión de Organizaciones Campesinas Cacaoteras del Ecuador",
                  "subcat": "Setor privado parceiro · SAM", "type": "Cooperativa de produtores",
                  "country": "Equador",
                  "role": "Cacau; peer regional UNICAFES PA/RO",
                  "mbo_relations": {
                      "UNICAFES_PA": "Peer cacau equatoriano-brasileiro",
                      "UNICAFES_RO": "Peer cacau equatoriano-brasileiro",
                  }},
    "T_APROCAM": {"name": "Aprocam", "long_name": "Aprocam · Cooperativa peruana de cacao",
                  "subcat": "Setor privado parceiro · SAM", "type": "Cooperativa de produtores",
                  "country": "Peru",
                  "role": "Cacau peruano; peer regional UNICAFES PA/RO",
                  "mbo_relations": {
                      "UNICAFES_PA": "Peer trinacional Peru-Brasil cacau",
                      "UNICAFES_RO": "Peer trinacional Peru-Brasil cacau",
                  }},
    "T_AGROPAPA": {"name": "AGROPAPA Tungurahua", "long_name": "AGROPAPA · Cooperativa equatoriana de batata",
                   "subcat": "Setor privado parceiro · SAM", "type": "Cooperativa de produtores",
                   "country": "Equador (Tungurahua)",
                   "role": "Modelo de business partner management replicável",
                   "mbo_relations": {
                       "UNICAFES_PA": "Modelo cooperativista equatoriano aplicável a PA",
                       "UNICAFES_RO": "Modelo cooperativista equatoriano aplicável a RO",
                   }},
    "T_CONPAPA": {"name": "CONPAPA Chimborazo", "long_name": "CONPAPA · Consorcio equatoriano de papas",
                  "subcat": "Setor privado parceiro · SAM", "type": "Cooperativa de produtores",
                  "country": "Equador (Chimborazo)",
                  "role": "Modelo cooperativista replicável",
                  "mbo_relations": {
                      "UNICAFES_PA": "Modelo cooperativista equatoriano aplicável a PA",
                      "UNICAFES_RO": "Modelo cooperativista equatoriano aplicável a RO",
                  }},
    "T_COOPAGROS": {"name": "COOPAGROS", "long_name": "COOPAGROS · Cooperativa peruana",
                    "subcat": "Setor privado parceiro · SAM", "type": "Cooperativa de produtores",
                    "country": "Peru",
                    "role": "Business partner management; peer regional",
                    "mbo_relations": {
                        "UNICAFES_PA": "Peer cooperativista Peru-Brasil",
                        "UNICAFES_RO": "Peer cooperativista Peru-Brasil",
                    }},

    # ---- Academia ----
    "T_LIEGE": {"name": "Université de Liège", "long_name": "Université de Liège (ULiège)",
                "subcat": "Academia", "type": "Universidade",
                "country": "Bélgica",
                "role": "Gastronomy Lab em Santarém-PA (com BID, AWEX, Suzano)",
                "mbo_relations": {
                    "UNICAFES_PA": "Gastronomy Lab Santarém-PA — parceria ativa direta",
                }},
    "T_VLERICK": {"name": "Vlerick Business School", "long_name": "Vlerick Business School",
                  "subcat": "Academia", "type": "Business school",
                  "country": "Bélgica",
                  "role": "Capacitação executiva para lideranças MBO (Narrative)",
                  "mbo_relations": {
                      "UNICAFES_PA": "Capacitação executiva para staff e lideranças (2 staff)",
                      "UNICAFES_RO": "Capacitação executiva para staff e lideranças (2 staff)",
                      "CSA_BRASIL": "Capacitação executiva para liderança (3 staff)",
                      "UNICATADORES": "Capacitação executiva para liderança (28 staff)",
                  }},
    "T_BRS":    {"name": "BRS", "long_name": "Belgian Raiffeisen Society",
                 "subcat": "Academia", "type": "Cooperativa financeira / academia",
                 "country": "Bélgica",
                 "role": "Expertise em microfinanças cooperativas",
                 "mbo_relations": {
                     "UNICAFES_PA": "Microfinanças cooperativas para agricultores familiares",
                     "UNICAFES_RO": "Microfinanças e blended finance",
                     "CSA_BRASIL": "Modelo de finança solidária para CSA",
                 }},
    "T_A4D":    {"name": "Academics for Development", "long_name": "Academics for Development",
                 "subcat": "Academia", "type": "Rede de voluntários acadêmicos",
                 "country": "Bélgica",
                 "role": "Apoio voluntário acadêmico a OSCs",
                 "mbo_relations": {
                     "UNICAFES_PA": "Voluntariado acadêmico em projetos pontuais",
                     "UNICAFES_RO": "Voluntariado acadêmico em projetos pontuais",
                     "CSA_BRASIL": "Voluntariado acadêmico em projetos pontuais",
                     "UNICATADORES": "Voluntariado acadêmico em projetos pontuais",
                 }},
}

# Paleta para a camada institucional — cor única, distinta dos setores
INSTITUTIONAL_COLOR = "#3A5A87"  # azul-ardósia profundo

# ========== 3b. Modo de articulação Trias ==========
# Mapeia o papel declarado do stakeholder ao modo de articulação que a Trias
# pode mobilizar com ele, segundo os 5 papéis declarados no Annex 1 (Theory of
# Change) — Process Facilitator, Thematic Advisor, Peer-to-Peer Facilitator,
# Bridge Builder e Financer. Cada modo descreve OPERACIONALMENTE como o
# stakeholder pode contribuir para o programa.
TRIAS_MODE_RULES = [
    # (role substring, modo, papel Trias mobilizado, descrição operacional)
    ("Funder",                "Co-financiamento",
     "Financer + Bridge Builder",
     "captação conjunta, edital co-financiado, fundos paralelos"),
    ("Capital provider",      "Finanças mistas/inovadoras",
     "Bridge Builder + Financer",
     "blended finance, garantias, crédito, instrumentos climáticos"),
    ("Capital provider",      "Co-financiamento",
     "Financer + Bridge Builder",
     "captação conjunta, edital co-financiado, fundos paralelos"),
    ("Strategic partner",     "Parceria estratégica",
     "Bridge Builder",
     "co-criação de agenda, alinhamento programático plurianual"),
    ("Technical assistance",  "Advisory técnico-temático",
     "Thematic Advisor",
     "ATER, formação técnica, mentoria temática para MBOs"),
    ("Research",              "Pesquisa & evidência",
     "Thematic Advisor",
     "estudos, baselines, MRV, avaliação"),
    ("Data / MRV",            "MRV / dados",
     "Thematic Advisor",
     "plataformas, monitoramento ambiental, traceabilidade"),
    ("Monitoring",            "MRV / dados",
     "Thematic Advisor",
     "sistemas MEAL, avaliação, learning"),
    ("Project developer",     "Co-implementação",
     "Peer-to-Peer Facilitator",
     "execução conjunta de projetos com as MBOs"),
    ("Implementer",           "Co-implementação",
     "Peer-to-Peer Facilitator",
     "execução conjunta de projetos com as MBOs"),
    ("Implementation capacity", "Co-implementação",
     "Peer-to-Peer Facilitator",
     "execução conjunta de projetos com as MBOs"),
    ("Policy influence",      "Advocacy compartilhado",
     "Bridge Builder",
     "incidência conjunta em PNRS, PNAE, PLANAPO, MROSC"),
    ("Convenor",              "Articulação multistakeholder",
     "Bridge Builder",
     "convocatória, plataforma multi-ator, redes temáticas"),
    ("enablers",              "Articulação multistakeholder",
     "Bridge Builder",
     "convocatória, plataforma multi-ator, redes temáticas"),
    ("Territorial anchor",    "Ancoragem territorial",
     "Process Facilitator",
     "presença local, legitimidade junto aos membros"),
    ("Community legitimacy",  "Ancoragem territorial",
     "Process Facilitator",
     "legitimidade comunitária, voz de membros"),
    ("proximate leadership",  "Ancoragem territorial",
     "Process Facilitator",
     "liderança próxima dos beneficiários"),
]

def derive_trias_modes(role_str):
    """Retorna lista de modos de articulação Trias com base no papel do stakeholder."""
    if not role_str:
        return []
    seen = set()
    out = []
    role_lower = role_str.lower()
    for rule, mode, trias_role, desc in TRIAS_MODE_RULES:
        if rule.lower() in role_lower and mode not in seen:
            seen.add(mode)
            out.append({"mode": mode, "trias_role": trias_role, "desc": desc})
    return out

# ========== 3c. Adicionalidade Trias ==========
# Para cada par (stakeholder, parceiro MBO), estima quanto a mediação da Trias
# agrega valor além do que aconteceria espontaneamente. Heurística baseada em:
# (a) gap territorial entre stakeholder e MBO, (b) maturidade do MBO (mais
# imaturo = mais bridging necessário), (c) escala/formalidade do stakeholder
# (mais institucional = mais necessidade de tradução), e (d) penalização se a
# relação já parece existir (mencionada em notas ou descrição).
PARTNER_MATURITY_SCORE = {
    "UNICAFES_PA": 1.0,    # 2º nível em consolidação → mais bridging necessário
    "UNICAFES_RO": 1.0,    # idem
    "CSA_BRASIL": 0.5,     # 3º nível em consolidação
    "UNICATADORES": 0.0,   # 3º nível maduro, Trias é parceiro estratégico
}

def compute_additionality(stk, partner_id):
    """Retorna (score 0-3, classe, fatores) — quanto maior, mais adicionalidade."""
    factors = []
    sc = 0.0
    territory = (stk["territory"] or "").lower()
    type_ = (stk["type"] or "").lower()
    notes_blob = (stk["notes"] + " " + stk["description"]).lower()
    # (a) gap territorial — parceiros amazônicos vs stakeholder não-amazônico
    p = PARTNERS[partner_id]
    if p["biomes"] == {"Amazonia"}:
        primary_biome = (stk["biome_primary"] or "").lower()
        if "amazon" not in primary_biome and "national" not in territory:
            sc += 1.0
            factors.append("gap territorial (parceiro amazônico, stakeholder não-amazônico)")
        elif "national" in territory:
            sc += 0.5
            factors.append("alcance nacional (precisa aterrissagem amazônica)")
    if "cross-border" in territory or "international" in territory:
        sc += 1.0
        factors.append("ator internacional — bridging Brasil-mundo")
    # (b) maturidade do MBO
    mat = PARTNER_MATURITY_SCORE[partner_id]
    if mat > 0:
        sc += mat
        factors.append(f"MBO em desenvolvimento ({p['maturity']})")
    # (c) escala/formalidade do stakeholder
    big_keywords = ["development bank", "multilateral", "bilateral",
                    "philanthropy foundation", "large", "industry platform"]
    if any(k in type_ for k in big_keywords):
        sc += 0.7
        factors.append("ator institucional de grande porte")
    # Impact investor pequeno também precisa ponte
    if "impact invest" in type_ or "fund" in type_.lower():
        sc += 0.4
        factors.append("intermediário financeiro — exige tradução de pipeline")
    # (d) penalização: relação aparentemente já existente
    mbo_terms = ["unicafes", "unicatadores", "csa brasil", "catador cooperat",
                 "agricultura familiar trias", "trias parc"]
    if any(t in notes_blob for t in mbo_terms):
        sc -= 1.5
        factors.append("relação aparentemente preexistente — penalização")
    sc = max(0, min(3, sc))
    if sc >= 2.0: cls = "Alta"
    elif sc >= 1.0: cls = "Média"
    else: cls = "Baixa"
    return round(sc, 1), cls, factors

for r in results:
    r["trias_modes"] = derive_trias_modes(r["stk"]["role"])
    r["additionality"] = {}
    for pid in r["scores"]:
        sc_add, cls, factors = compute_additionality(r["stk"], pid)
        r["additionality"][pid] = {"score": sc_add, "class": cls, "factors": factors}

# ========== 4. Estatísticas agregadas ==========
def split_tokens(s):
    if not s: return []
    return [t.strip() for t in re.split(r"[,;]", s) if t.strip()]

stats_all = defaultdict(Counter)
stats_connected = defaultdict(Counter)
for stk in records:
    for tok in split_tokens(stk["sector"]):
        stats_all["sector"][tok] += 1
    for tok in split_tokens(stk["biome_primary"]):
        stats_all["biome"][tok] += 1
    for tok in split_tokens(stk["role"]):
        stats_all["role"][tok] += 1
    for tok in split_tokens(stk["funding_role"]):
        stats_all["funding_role"][tok] += 1
    for tok in split_tokens(stk["value_chain"]):
        stats_all["value_chain"][tok] += 1
    for tok in split_tokens(stk["territory"]):
        stats_all["territory"][tok] += 1

for r in results:
    stk = r["stk"]
    for tok in split_tokens(stk["sector"]):
        stats_connected["sector"][tok] += 1
    for tok in split_tokens(stk["biome_primary"]):
        stats_connected["biome"][tok] += 1
    for tok in split_tokens(stk["role"]):
        stats_connected["role"][tok] += 1
    for tok in split_tokens(stk["funding_role"]):
        stats_connected["funding_role"][tok] += 1
    for tok in split_tokens(stk["value_chain"]):
        stats_connected["value_chain"][tok] += 1

# Conexões por parceiro (counts e top tipos)
per_partner_stats = {}
for pid, p in PARTNERS.items():
    matches = [r for r in results if pid in r["scores"]]
    sector_c = Counter()
    type_c = Counter()
    for r in matches:
        sector_c[r["stk"]["sector"] or "Others"] += 1
        type_c[r["stk"]["type"] or "—"] += 1
    per_partner_stats[pid] = {
        "n": len(matches),
        "sectors": dict(sector_c.most_common()),
        "types": dict(type_c.most_common(8)),
    }

# Bridges
bridges = [r for r in results if len(r["scores"]) >= 2]
n_b3 = sum(1 for r in results if len(r["scores"]) >= 3)
n_b4 = sum(1 for r in results if len(r["scores"]) == 4)
n_b2 = sum(1 for r in results if len(r["scores"]) == 2)
n_b1 = sum(1 for r in results if len(r["scores"]) == 1)

# ========== 5. Paleta setorial sutil ==========
SECTOR_COLOR = {
    "Civil Society Organization (CSO)": "#4A6FA5",
    "Funders":                          "#C68A5D",
    "Private sector":                   "#6BA08C",
    "Public sector":                    "#8DA570",
    "Others":                           "#9B7AB0",
}
PARTNER_COLOR = "#B14545"

# ========== 6. Grafo ==========
G = nx.Graph()
for pid, p in PARTNERS.items():
    G.add_node(pid, kind="partner", **{k: (list(v) if isinstance(v, set) else v)
                                       for k, v in p.items()}, color=PARTNER_COLOR)

for r in results:
    stk = r["stk"]
    node_id = f"S{stk['id']}"
    sector = stk["sector"] or "Others"
    modes = r["trias_modes"]
    # adicionalidade agregada do stakeholder = média das conexões com parceiros
    addit_scores = [r["additionality"][pid]["score"] for pid in r["scores"]]
    addit_max = max(addit_scores) if addit_scores else 0
    addit_avg = round(sum(addit_scores) / len(addit_scores), 1) if addit_scores else 0
    if addit_avg >= 2.0: addit_class = "Alta"
    elif addit_avg >= 1.0: addit_class = "Média"
    else: addit_class = "Baixa"
    G.add_node(node_id, kind="stakeholder",
               name=stk["name"], sector=sector, type=stk["type"],
               territory=stk["territory"], biome=stk["biome_primary"],
               biome_secondary=stk["biome_secondary"],
               impact=stk["impact_area"], role=stk["role"],
               value_chain=stk["value_chain"],
               funding_role=stk["funding_role"],
               hq=stk["hq"], url=stk["url"],
               description=stk["description"], notes=stk["notes"],
               n_partners=len(r["scores"]),
               trias_modes=[m["mode"] for m in modes],
               trias_mode_details=modes,
               additionality_per_partner={pid: r["additionality"][pid]
                                          for pid in r["scores"]},
               additionality_avg=addit_avg,
               additionality_max=addit_max,
               additionality_class=addit_class,
               color=SECTOR_COLOR.get(sector, "#999999"),
               connections={pid: r["scores"][pid]["score"] for pid in r["scores"]},
               score_breakdowns={pid: r["scores"][pid]["breakdown"]
                                 for pid in r["scores"]})
    for pid, s in r["scores"].items():
        G.add_edge(node_id, pid, weight=s["score"],
                   trias_modes=[m["mode"] for m in modes],
                   additionality_class=r["additionality"][pid]["class"])

# Peer links MBO
peer_pairs = [("UNICAFES_PA", "UNICAFES_RO"),
              ("UNICAFES_PA", "CSA_BRASIL"),
              ("UNICAFES_RO", "CSA_BRASIL"),
              ("UNICAFES_PA", "UNICATADORES"),
              ("UNICAFES_RO", "UNICATADORES"),
              ("CSA_BRASIL", "UNICATADORES")]
for a, b in peer_pairs:
    G.add_edge(a, b, weight=8, kind="peer")

# Camada institucional Trias
for inst_id, inst in TRIAS_INSTITUTIONAL.items():
    mbo_rels = inst.get("mbo_relations", {})
    G.add_node(inst_id, kind="trias_institutional",
               name=inst["name"], long_name=inst["long_name"],
               subcat=inst["subcat"], type=inst["type"],
               country=inst["country"], role=inst["role"],
               color=INSTITUTIONAL_COLOR,
               connects_to=list(mbo_rels.keys()),
               mbo_relations=mbo_rels)
    for mbo, reason in mbo_rels.items():
        G.add_edge(inst_id, mbo, weight=6, kind="institutional", reason=reason)

betweenness = nx.betweenness_centrality(G, weight="weight")
for nid in G.nodes:
    G.nodes[nid]["betweenness"] = round(betweenness[nid], 4)
    G.nodes[nid]["degree"] = G.degree(nid)

print(f"Grafo: {G.number_of_nodes()} nós, {G.number_of_edges()} arestas")

# Top brokers
top_brokers = sorted(
    [(nid, G.nodes[nid]) for nid in G.nodes if G.nodes[nid]["kind"] == "stakeholder"],
    key=lambda x: -x[1]["betweenness"])[:15]

# Top stakeholders por parceiro (com score)
top_by_partner = {}
for pid in PARTNERS:
    lst = sorted(
        [(f"S{r['stk']['id']}", G.nodes[f"S{r['stk']['id']}"], r["scores"][pid]["score"])
         for r in results if pid in r["scores"]],
        key=lambda x: -x[2])
    top_by_partner[pid] = [{"id": x[0], "name": x[1]["name"], "type": x[1]["type"],
                            "sector": x[1]["sector"], "score": x[2]}
                           for x in lst[:15]]

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
        "sector": stk["sector"],
        "n": len(r["scores"]),
        "partners": [PARTNERS[pid]["name"] for pid in r["scores"]],
        "total_score": sum(r["scores"][pid]["score"] for pid in r["scores"]),
    })

# ========== 7. JSON do grafo para D3 ==========
nodes_json = []
for nid in G.nodes:
    nodes_json.append({**dict(G.nodes[nid]), "id": nid})
links_json = []
for u, v, d in G.edges(data=True):
    links_json.append({
        "source": u, "target": v,
        "weight": d.get("weight", 1),
        "kind": d.get("kind", "alignment"),
    })

# Listas de valores únicos para filtros
def unique_tokens(field):
    s = set()
    for r in nodes_json:
        if r.get("kind") == "stakeholder":
            for tok in split_tokens(r.get(field, "")):
                s.add(tok)
    return sorted(s)

# Distribuições derivadas: modos Trias, adicionalidade
mode_counter = Counter()
addit_class_counter = Counter()
addit_class_per_partner = defaultdict(Counter)
trias_role_counter = Counter()
for r in results:
    for m in r["trias_modes"]:
        mode_counter[m["mode"]] += 1
        trias_role_counter[m["trias_role"]] += 1
    for pid, ad in r["additionality"].items():
        addit_class_counter[ad["class"]] += 1
        addit_class_per_partner[pid][ad["class"]] += 1

mode_options = sorted(mode_counter.keys())
addit_options = ["Alta", "Média", "Baixa"]

graph_data = {
    "nodes": nodes_json,
    "links": links_json,
    "partners": [{"id": pid, **{k: (list(v) if isinstance(v, set) else v)
                                for k, v in p.items()}}
                 for pid, p in PARTNERS.items()],
    "sectors": SECTOR_COLOR,
    "stats": {
        "total_db": len(records),
        "connected": len(results),
        "bridges": len(bridges),
        "tri": n_b3,
        "quad": n_b4,
        "by_partner": {pid: per_partner_stats[pid]["n"] for pid in PARTNERS},
        "dist_sector_all": dict(stats_all["sector"]),
        "dist_sector_connected": dict(stats_connected["sector"]),
        "dist_biome_all": dict(stats_all["biome"]),
        "dist_biome_connected": dict(stats_connected["biome"]),
        "dist_role_all": dict(stats_all["role"].most_common(10)),
        "dist_role_connected": dict(stats_connected["role"].most_common(10)),
        "dist_funding_all": dict(stats_all["funding_role"]),
        "dist_funding_connected": dict(stats_connected["funding_role"]),
        "dist_chain_all": dict(stats_all["value_chain"]),
        "dist_chain_connected": dict(stats_connected["value_chain"]),
    },
    "top_by_partner": top_by_partner,
    "bridges": bridges_data,
    "top_brokers": [{"id": x[0], "name": x[1]["name"], "betweenness": x[1]["betweenness"],
                     "sector": x[1]["sector"], "type": x[1]["type"]}
                    for x in top_brokers],
    "filter_options": {
        "sector": unique_tokens("sector"),
        "biome": sorted({tok for r in nodes_json if r.get("kind") == "stakeholder"
                          for tok in split_tokens(r.get("biome", ""))}),
        "role": sorted({tok for r in nodes_json if r.get("kind") == "stakeholder"
                          for tok in split_tokens(r.get("role", ""))}),
        "funding_role": sorted({r.get("funding_role", "") for r in nodes_json
                                if r.get("kind") == "stakeholder" and r.get("funding_role")}),
        "trias_mode": mode_options,
        "additionality": addit_options,
    },
    "dist_trias_mode": dict(mode_counter.most_common()),
    "dist_trias_role": dict(trias_role_counter.most_common()),
    "dist_additionality": dict(addit_class_counter),
    "addit_per_partner": {pid: dict(addit_class_per_partner[pid]) for pid in PARTNERS},
    "institutional_color": INSTITUTIONAL_COLOR,
    "dist_institutional_subcat": dict(Counter(inst["subcat"] for inst in TRIAS_INSTITUTIONAL.values())),
    "dist_institutional_country": dict(Counter(inst["country"] for inst in TRIAS_INSTITUTIONAL.values()).most_common()),
    "n_institutional": len(TRIAS_INSTITUTIONAL),
    "config": {
        "threshold": THRESHOLD,
    },
}

def jdefault(o):
    if isinstance(o, set): return list(o)
    raise TypeError
graph_json_str = json.dumps(graph_data, default=jdefault, ensure_ascii=False)

print("Estatísticas-chave:")
print(f"  Setor (na base): {dict(stats_all['sector'])}")
print(f"  Biome (na base): {dict(stats_all['biome'])}")
print(f"  Funding role (na base): {dict(stats_all['funding_role'])}")
print(f"  Conexões por parceiro: {{ {', '.join(f'{p}: {n}' for p, n in graph_data['stats']['by_partner'].items())} }}")

# ========== 8. HTML template (tema claro, sidebars colapsáveis) ==========
HTML = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>Ecossistema Trias Brasil — DGD 2027–2031</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Source+Serif+4:wght@400;600&family=JetBrains+Mono:wght@500&display=swap" rel="stylesheet">
<script src="https://d3js.org/d3.v7.min.js"></script>
<style>
  :root {
    --bg: #FCFCFD;
    --panel: #FFFFFF;
    --panel-2: #F5F6F8;
    --border: #E2E5EA;
    --border-2: #CFD4DC;
    --text: #1A2230;
    --text-2: #3C4757;
    --muted: #6B7888;
    --muted-2: #97A0AE;
    --accent: #B14545;
    --accent-soft: rgba(177,69,69,0.08);
    --link: rgba(70,80,100,0.18);
    --link-strong: rgba(40,50,70,0.50);
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; height: 100%; font-family: 'Inter', sans-serif;
                background: var(--bg); color: var(--text); overflow: hidden; }
  body { font-size: 13px; }
  #app { display: grid; height: 100vh; grid-template-rows: 44px 1fr;
         grid-template-columns: var(--lw, 280px) 1fr var(--rw, 320px);
         grid-template-areas: "header header header" "left stage right";
         transition: grid-template-columns 0.25s ease; }
  body.left-hidden { --lw: 0px; }
  body.right-hidden { --rw: 0px; }
  /* camada institucional escondida por padrao */
  body .node.trias-institutional,
  body .link.institutional { display: none; }
  body.institutional-on .node.trias-institutional,
  body.institutional-on .link.institutional { display: block; }
  header.bar { grid-area: header; background: var(--panel);
               border-bottom: 1px solid var(--border); display: flex;
               align-items: center; padding: 0 14px; gap: 10px; z-index: 20; }
  header.bar .title { font-family: 'Source Serif 4', serif; font-weight: 600;
                       font-size: 15px; color: var(--text); letter-spacing: -0.2px; }
  header.bar .subtitle { font-size: 11px; color: var(--muted);
                          padding-left: 10px; border-left: 1px solid var(--border);
                          margin-left: 6px; }
  header.bar .spacer { flex: 1; }
  header.bar .icon-btn { width: 32px; height: 30px; border: 1px solid var(--border);
                          background: var(--panel); color: var(--text-2);
                          border-radius: 6px; cursor: pointer; display: flex;
                          align-items: center; justify-content: center; font-size: 14px;
                          font-family: inherit; transition: all 0.12s; }
  header.bar .icon-btn:hover { border-color: var(--border-2); background: var(--panel-2); }
  header.bar .icon-btn.active { background: var(--text); color: #fff; border-color: var(--text); }
  header.bar .text-btn { padding: 6px 12px; border: 1px solid var(--border);
                          background: var(--panel); color: var(--text-2);
                          border-radius: 6px; cursor: pointer; font-size: 12px;
                          font-family: inherit; transition: all 0.12s; }
  header.bar .text-btn:hover { border-color: var(--border-2); background: var(--panel-2); }

  aside.left { grid-area: left; background: var(--panel);
                border-right: 1px solid var(--border); overflow-y: auto;
                overflow-x: hidden; padding: 18px 16px; }
  aside.right { grid-area: right; background: var(--panel);
                 border-left: 1px solid var(--border); overflow-y: auto;
                 padding: 18px 16px; }
  body.left-hidden aside.left { padding: 0; overflow: hidden; }
  body.right-hidden aside.right { padding: 0; overflow: hidden; }
  main#stage { grid-area: stage; position: relative; background: var(--bg); overflow: hidden; }
  svg.network { width: 100%; height: 100%; display: block; cursor: grab; }
  svg.network:active { cursor: grabbing; }

  h2 { font-size: 10px; text-transform: uppercase; letter-spacing: 1.5px;
        color: var(--muted-2); margin: 22px 0 8px 0; font-weight: 600; }
  h2:first-child { margin-top: 0; }
  .small { font-size: 11px; color: var(--muted); }
  .mono { font-family: 'JetBrains Mono', monospace; font-size: 11px; }

  /* Filters */
  #search { width: 100%; padding: 8px 11px; border-radius: 6px;
            border: 1px solid var(--border); background: var(--panel);
            color: var(--text); font-size: 12.5px; font-family: inherit;
            transition: border-color 0.12s; }
  #search:focus { outline: none; border-color: var(--text); }
  .filter-group { margin-bottom: 4px; }
  .check { display: flex; align-items: center; gap: 8px; padding: 4px 6px;
            border-radius: 4px; cursor: pointer; user-select: none; font-size: 12px;
            transition: background 0.1s; }
  .check:hover { background: var(--panel-2); }
  .check input { margin: 0; cursor: pointer; accent-color: var(--text); }
  .check .swatch { width: 10px; height: 10px; border-radius: 2px; flex-shrink: 0; }
  .check .count { margin-left: auto; color: var(--muted); font-family: 'JetBrains Mono', monospace; font-size: 10.5px; }
  .check .label { flex: 1; color: var(--text-2); }
  .filter-actions { display: flex; gap: 6px; margin: 6px 0 4px 0; }
  .filter-actions button { flex: 1; padding: 5px; border: 1px solid var(--border);
                            background: var(--panel); color: var(--muted);
                            border-radius: 4px; cursor: pointer; font-size: 10.5px;
                            font-family: inherit; }
  .filter-actions button:hover { color: var(--text); border-color: var(--border-2); }

  /* Partner chips */
  .partner-chips { display: flex; flex-wrap: wrap; gap: 5px; }
  .partner-chip { padding: 5px 10px; border-radius: 999px; background: var(--panel);
                   border: 1px solid var(--border); color: var(--text-2);
                   font-size: 11px; cursor: pointer; user-select: none;
                   transition: all 0.12s; }
  .partner-chip:hover { border-color: var(--border-2); }
  .partner-chip.active { background: var(--accent); color: #fff; border-color: var(--accent); }

  /* Right panel: stats vs inspector */
  .tabs { display: flex; gap: 0; margin-bottom: 16px; border-bottom: 1px solid var(--border); }
  .tab { padding: 8px 0; flex: 1; text-align: center; cursor: pointer;
          font-size: 11.5px; color: var(--muted); border-bottom: 2px solid transparent;
          font-weight: 500; transition: all 0.12s; }
  .tab:hover { color: var(--text); }
  .tab.active { color: var(--text); border-bottom-color: var(--text); }
  .tab-pane { display: none; }
  .tab-pane.active { display: block; }

  /* Big stat cards */
  .stat-row { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; margin-bottom: 6px; }
  .stat-card { padding: 10px 12px; background: var(--panel-2); border-radius: 6px;
                border: 1px solid var(--border); }
  .stat-card .v { font-size: 22px; font-weight: 700; color: var(--text); line-height: 1;
                   font-feature-settings: "tnum"; }
  .stat-card .k { font-size: 10px; color: var(--muted); text-transform: uppercase;
                   letter-spacing: 0.5px; margin-top: 4px; }

  /* Bar charts */
  .bars { margin: 4px 0; }
  .bar-row { display: grid; grid-template-columns: 1fr 30px; gap: 6px; align-items: center;
              font-size: 11px; padding: 3px 0; cursor: default; }
  .bar-row .lbl { color: var(--text-2); }
  .bar-row .num { text-align: right; font-family: 'JetBrains Mono', monospace; color: var(--muted); font-size: 10.5px; }
  .bar-row .track { grid-column: 1 / -1; height: 4px; background: var(--panel-2);
                     border-radius: 2px; overflow: hidden; }
  .bar-row .fill { height: 100%; background: var(--text); border-radius: 2px;
                    transition: width 0.4s ease; }
  .bar-row.with-sub .lbl { display: flex; align-items: center; gap: 6px; }
  .bar-row.with-sub .swatch { width: 9px; height: 9px; border-radius: 2px; flex-shrink: 0; }

  /* Inspector */
  .ins-empty { color: var(--muted); font-size: 12px; padding: 12px 0; line-height: 1.5; }
  .ins-header { padding-bottom: 12px; border-bottom: 1px solid var(--border); margin-bottom: 12px; }
  .ins-header .name { font-family: 'Source Serif 4', serif; font-size: 16px; font-weight: 600;
                       color: var(--text); line-height: 1.25; margin-bottom: 4px; }
  .ins-header .type { font-size: 11px; color: var(--muted); }
  .ins-row { margin: 10px 0; font-size: 12px; line-height: 1.45; }
  .ins-row .k { color: var(--muted-2); font-size: 9.5px; text-transform: uppercase;
                 letter-spacing: 0.7px; font-weight: 600; margin-bottom: 3px; }
  .ins-row .v { color: var(--text-2); }
  .ins-row a { color: var(--accent); text-decoration: none; word-break: break-all; }
  .ins-row a:hover { text-decoration: underline; }
  .conn-row { display: flex; justify-content: space-between; align-items: center;
               padding: 8px 10px; background: var(--panel-2); border-radius: 5px;
               margin-bottom: 4px; font-size: 11.5px; }
  .conn-row .pname { color: var(--text); font-weight: 500; }
  .conn-row .pscore { font-family: 'JetBrains Mono', monospace; color: var(--accent); font-size: 11px; }
  .breakdown-line { font-size: 11px; color: var(--muted); padding: 2px 0;
                     font-family: 'JetBrains Mono', monospace; }
  .breakdown-line .pts { color: var(--accent); margin-left: 6px; }

  /* Lists */
  .item-list { list-style: none; padding: 0; margin: 0; }
  .item-list li { padding: 8px 10px; border: 1px solid var(--border); border-radius: 5px;
                   margin-bottom: 5px; cursor: pointer; font-size: 11.5px; line-height: 1.4;
                   transition: all 0.12s; background: var(--panel); }
  .item-list li:hover { border-color: var(--accent); background: var(--accent-soft); }
  .item-list .pill { display: inline-block; background: var(--panel-2);
                      padding: 1px 6px; border-radius: 3px; font-family: 'JetBrains Mono', monospace;
                      font-size: 10px; color: var(--muted); }
  .item-list .meta { color: var(--muted); font-size: 10.5px; margin-top: 3px; }

  /* SVG */
  .link { stroke: var(--link); fill: none; stroke-linecap: round;
           transition: stroke 0.25s, stroke-width 0.25s, opacity 0.25s; }
  .link.peer { stroke: rgba(177,69,69,0.25); stroke-dasharray: 4 3; }
  .link.institutional { stroke: rgba(58,90,135,0.18); stroke-width: 0.8px; }
  .link.dim { opacity: 0.06; }
  .link.focus { stroke: rgba(177,69,69,0.65); }
  .link.institutional.focus { stroke: rgba(58,90,135,0.85); stroke-width: 1.8px; }
  .btn-institutional-on { background: var(--text); color: #fff; border-color: var(--text); }
  .node { cursor: pointer; }
  .node-halo { fill: var(--accent); opacity: 0; transition: opacity 0.25s; }
  .node-halo.show { opacity: 0.14; }
  .node-circle { stroke: #fff; stroke-width: 1.5px;
                  transition: stroke 0.25s, stroke-width 0.25s; }
  .node.partner .node-circle { stroke: var(--text); stroke-width: 2.5px; }
  .node.dim .node-circle { opacity: 0.16; }
  .node.dim text { opacity: 0.12; }
  .node.focused .node-circle { stroke: var(--text); stroke-width: 2.5px; }
  .node text { font-size: 10.5px; fill: var(--text); pointer-events: none;
                text-anchor: middle; font-weight: 500;
                paint-order: stroke; stroke: var(--bg); stroke-width: 3px;
                stroke-linejoin: round; transition: opacity 0.25s; }
  .node.partner text { font-weight: 700; font-size: 13px; fill: var(--text); }

  /* Controls overlay */
  .stage-controls { position: absolute; left: 14px; bottom: 14px; display: flex;
                     flex-direction: column; gap: 4px; z-index: 5; }
  .stage-controls .icon-btn { width: 30px; height: 30px; border: 1px solid var(--border);
                               background: var(--panel); border-radius: 5px; cursor: pointer;
                               font-size: 14px; color: var(--text-2); }
  .stage-controls .icon-btn:hover { border-color: var(--border-2); }
  .stage-hint { position: absolute; right: 14px; bottom: 14px; font-size: 10.5px;
                color: var(--muted-2); user-select: none; pointer-events: none; }

  /* Modal */
  .modal-bg { position: fixed; inset: 0; background: rgba(20,28,40,0.45);
               z-index: 100; display: none; }
  .modal-bg.show { display: flex; align-items: center; justify-content: center; }
  .modal { background: var(--panel); border-radius: 10px; width: 720px; max-width: 95vw;
            max-height: 88vh; overflow-y: auto; padding: 24px 28px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.2); }
  .modal h3 { font-family: 'Source Serif 4', serif; font-size: 19px; margin: 0 0 16px 0;
              font-weight: 600; color: var(--text); }
  .modal h4 { font-size: 12px; text-transform: uppercase; letter-spacing: 1.3px;
              color: var(--muted-2); margin: 20px 0 8px 0; font-weight: 600; }
  .modal p { line-height: 1.55; color: var(--text-2); font-size: 13px; }
  .modal ul { padding-left: 20px; }
  .modal li { font-size: 12.5px; line-height: 1.6; color: var(--text-2); margin-bottom: 4px; }
  .modal .close-btn { float: right; background: none; border: none; font-size: 22px;
                       color: var(--muted); cursor: pointer; padding: 0 4px; }
  .modal .close-btn:hover { color: var(--text); }
  .modal .formula-table { width: 100%; border-collapse: collapse; margin: 8px 0; }
  .modal .formula-table th, .modal .formula-table td { padding: 6px 8px; text-align: left;
                                                       border-bottom: 1px solid var(--border);
                                                       font-size: 12px; }
  .modal .formula-table th { color: var(--muted); font-weight: 600;
                              text-transform: uppercase; font-size: 10.5px; letter-spacing: 0.7px; }
  .modal code { background: var(--panel-2); padding: 1px 6px; border-radius: 3px;
                 font-size: 11.5px; font-family: 'JetBrains Mono', monospace; }

  /* Scrollbar */
  ::-webkit-scrollbar { width: 8px; }
  ::-webkit-scrollbar-track { background: transparent; }
  ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 4px; }
  ::-webkit-scrollbar-thumb:hover { background: var(--border-2); }

  /* Tooltip */
  .tip { position: absolute; pointer-events: none; background: var(--panel);
          border: 1px solid var(--border-2); box-shadow: 0 4px 16px rgba(0,0,0,0.08);
          padding: 8px 11px; border-radius: 6px; font-size: 11.5px;
          max-width: 260px; z-index: 50; display: none; line-height: 1.4; }
  .tip strong { color: var(--text); }
  .tip .tip-meta { color: var(--muted); font-size: 10.5px; margin-top: 4px; }
</style>
</head>
<body>
<div id="app">
  <header class="bar">
    <button class="icon-btn" id="toggle-left" title="Mostrar/ocultar filtros (F)" aria-label="Toggle filters">☰</button>
    <div class="title">Ecossistema Trias Brasil</div>
    <div class="subtitle">DGD 2027–2031 · rede ancorada nos 4 parceiros MBO</div>
    <div class="spacer"></div>
    <button class="text-btn" id="btn-institutional" title="Mostrar/ocultar camada institucional Trias (I)">Camada institucional Trias</button>
    <button class="text-btn" id="btn-methodology">Metodologia</button>
    <button class="icon-btn" id="toggle-right" title="Mostrar/ocultar painel direito (D)" aria-label="Toggle right panel">▣</button>
  </header>

  <aside class="left">
    <h2>Busca</h2>
    <input id="search" placeholder="Nome ou trecho da descrição" autocomplete="off">

    <h2>Parceiro MBO</h2>
    <div class="partner-chips" id="partner-filters">
      <span class="partner-chip active" data-partner="ALL">Todos</span>
    </div>

    <h2>Setor <span class="filter-actions" style="float:right;display:inline-flex;margin:0;width:auto;"><button data-group="sector" data-action="all">tudo</button><button data-group="sector" data-action="none">limpar</button></span></h2>
    <div class="filter-group" id="filter-sector"></div>

    <h2>Bioma primário <span class="filter-actions" style="float:right;display:inline-flex;margin:0;width:auto;"><button data-group="biome" data-action="all">tudo</button><button data-group="biome" data-action="none">limpar</button></span></h2>
    <div class="filter-group" id="filter-biome"></div>

    <h2>Papel no ecossistema <span class="filter-actions" style="float:right;display:inline-flex;margin:0;width:auto;"><button data-group="role" data-action="all">tudo</button><button data-group="role" data-action="none">limpar</button></span></h2>
    <div class="filter-group" id="filter-role"></div>

    <h2>Funding role <span class="filter-actions" style="float:right;display:inline-flex;margin:0;width:auto;"><button data-group="funding_role" data-action="all">tudo</button><button data-group="funding_role" data-action="none">limpar</button></span></h2>
    <div class="filter-group" id="filter-funding"></div>

    <h2>Modo de articulação Trias <span class="filter-actions" style="float:right;display:inline-flex;margin:0;width:auto;"><button data-group="trias_mode" data-action="all">tudo</button><button data-group="trias_mode" data-action="none">limpar</button></span></h2>
    <div class="small" style="margin-bottom:6px;">derivado do papel do stakeholder, mapeado aos 5 papéis Trias do Annex 1</div>
    <div class="filter-group" id="filter-mode"></div>

    <h2>Adicionalidade Trias <span class="filter-actions" style="float:right;display:inline-flex;margin:0;width:auto;"><button data-group="additionality" data-action="all">tudo</button><button data-group="additionality" data-action="none">limpar</button></span></h2>
    <div class="small" style="margin-bottom:6px;">quanto a mediação Trias adiciona vs. relação espontânea</div>
    <div class="filter-group" id="filter-addit"></div>

    <h2>Camadas no mapa</h2>
    <div style="font-size:11.5px;line-height:1.6;color:var(--text-2);">
      <div style="display:flex;align-items:center;gap:8px;padding:3px 0;"><span style="width:14px;height:14px;background:#B14545;display:inline-block;transform:rotate(45deg);margin-left:3px;"></span> Parceiro MBO (4)</div>
      <div style="display:flex;align-items:center;gap:8px;padding:3px 0;"><span style="width:11px;height:11px;border-radius:50%;background:#4A6FA5;display:inline-block;"></span> Ecossistema brasileiro</div>
      <div style="display:flex;align-items:center;gap:8px;padding:3px 0;"><span style="width:11px;height:11px;border-radius:2px;background:#3A5A87;display:inline-block;"></span> Camada institucional Trias</div>
    </div>

    <h2>Atalhos de teclado</h2>
    <div class="muted" style="line-height: 1.6;">
      <b>F</b> esconde filtros · <b>D</b> esconde painel direito · <b>I</b> esconde camada institucional · <b>Esc</b> limpa seleção.
    </div>
  </aside>

  <main id="stage">
    <svg class="network" id="net"></svg>
    <div class="stage-controls">
      <button class="icon-btn" id="btn-reset" title="Resetar visão">⌂</button>
      <button class="icon-btn" id="btn-zoom-in" title="Zoom in">+</button>
      <button class="icon-btn" id="btn-zoom-out" title="Zoom out">−</button>
      <button class="icon-btn" id="btn-physics" title="Pausar/retomar física">⏸</button>
    </div>
    <div class="stage-hint">clique no vazio para limpar a seleção</div>
    <div class="tip" id="tip"></div>
  </main>

  <aside class="right">
    <div class="tabs">
      <div class="tab active" data-tab="stats">Estatísticas</div>
      <div class="tab" data-tab="inspector">Inspecionar</div>
      <div class="tab" data-tab="rank">Rankings</div>
    </div>

    <div class="tab-pane active" id="pane-stats">
      <h2>Cobertura</h2>
      <div class="stat-row">
        <div class="stat-card"><div class="v" id="kpi-db"></div><div class="k">na base</div></div>
        <div class="stat-card"><div class="v" id="kpi-conn"></div><div class="k">conectados</div></div>
      </div>
      <div class="stat-row">
        <div class="stat-card"><div class="v" id="kpi-bridges"></div><div class="k">bridges (2+)</div></div>
        <div class="stat-card"><div class="v" id="kpi-tri"></div><div class="k">tri-bridges</div></div>
      </div>

      <h2>Conexões por parceiro</h2>
      <div class="bars" id="bars-partner"></div>

      <h2>Composição por setor</h2>
      <div class="bars" id="bars-sector"></div>

      <h2>Composição por bioma</h2>
      <div class="bars" id="bars-biome"></div>

      <h2>Top papéis</h2>
      <div class="bars" id="bars-role"></div>

      <h2>Funding role</h2>
      <div class="bars" id="bars-funding"></div>

      <h2>Posição na cadeia</h2>
      <div class="bars" id="bars-chain"></div>

      <h2>Modo de articulação Trias</h2>
      <div class="small" style="margin-bottom:6px;">entre os 114 stakeholders conectados</div>
      <div class="bars" id="bars-mode"></div>

      <h2>Papel Trias mobilizado</h2>
      <div class="small" style="margin-bottom:6px;">os 5 papéis declarados no Annex 1</div>
      <div class="bars" id="bars-trias-role"></div>

      <h2>Adicionalidade Trias</h2>
      <div class="small" style="margin-bottom:6px;">distribuição das conexões potenciais</div>
      <div class="bars" id="bars-addit"></div>

      <h2>Camada institucional Trias</h2>
      <div class="small" style="margin-bottom:6px;"><span id="n-inst"></span> atores institucionais vinculados aos parceiros MBO. Tecla <b>I</b> para esconder/mostrar.</div>
      <div class="bars" id="bars-inst-subcat"></div>

      <h2>Origem / país (camada institucional)</h2>
      <div class="bars" id="bars-inst-country"></div>
    </div>

    <div class="tab-pane" id="pane-inspector">
      <div class="ins-empty" id="ins-empty">Clique num nó da rede para inspecionar.</div>
      <div id="ins-content" style="display:none;"></div>
    </div>

    <div class="tab-pane" id="pane-rank">
      <h2>Bridges — conectam múltiplos parceiros</h2>
      <ul class="item-list" id="bridges-list"></ul>
      <h2>Brokers — alta intermediação (betweenness)</h2>
      <ul class="item-list" id="brokers-list"></ul>
      <h2>Top por parceiro</h2>
      <div id="top-blocks"></div>
    </div>
  </aside>
</div>

<!-- Methodology modal -->
<div class="modal-bg" id="modal-bg">
  <div class="modal">
    <button class="close-btn" id="modal-close">×</button>
    <h3>Como a rede foi construída</h3>
    <p>A rede é uma <em>projeção bipartite</em>: cada stakeholder externo recebe uma aresta ponderada para cada um dos 4 parceiros MBO da Trias se o alinhamento temático/territorial passar de um limiar. Não há ligações diretas entre stakeholders externos — a leitura visual privilegia o papel de cada um <em>em relação aos parceiros</em>.</p>

    <h4>Base de origem</h4>
    <p>360 organizações registradas em <code>Stakeholder Ecosystem Mapping.xlsx</code> (última atualização: 11/02/2026). Cada registro traz: setor, tipo, descrição, território, bioma primário e secundário, área de impacto, papel no ecossistema, posição na cadeia de valor e funding role (donor / grantee / re-granter).</p>

    <h4>Nós-âncora (parceiros MBO)</h4>
    <p>Os 4 parceiros não estão na base original (apenas UNICAFES Nacional e UNICATADORES aparecem como entradas). Foram adicionados manualmente com perfis derivados do <em>Annex 1 — Theory of Change</em> e do <em>BRAZIL_DGD Narrative_DRAFT</em>:</p>
    <ul>
      <li><strong>UNICAFES Pará / Rondônia</strong> — biome: Amazônia; impact area: Land, Food and Forest; keywords: agricultura familiar, cooperativa, bioeconomia, sociobiodiversidade, café/cacau/açaí.</li>
      <li><strong>CSA Brasil</strong> — biomes: Mata Atlântica, Cerrado, Amazônia; impact area: Land, Food and Forest; keywords: community supported, agroecologia, food system, peri-urb, rural-urban.</li>
      <li><strong>UNICATADORES</strong> — sem biome (urbano); impact areas: Land/Food/Forest + Buildings/Transport; keywords: catador, recicl, circular econom, PNRS, EPR.</li>
    </ul>

    <h4>Score de alinhamento</h4>
    <p>Para cada par (stakeholder, parceiro) somam-se pontos quando há sobreposição:</p>
    <table class="formula-table">
      <thead><tr><th>Critério</th><th>Pontos</th><th>Observação</th></tr></thead>
      <tbody>
        <tr><td>Bioma primário coincide</td><td>+3</td><td>diferenciador territorial mais forte</td></tr>
        <tr><td>Bioma secundário coincide</td><td>+1</td><td>—</td></tr>
        <tr><td>Impact area coincide</td><td>+1</td><td>peso baixo: 244/360 têm "Land, Food and Forest"</td></tr>
        <tr><td>Role coincide (até 2)</td><td>+0,5 por match</td><td>cap em 1,0</td></tr>
        <tr><td>Keywords específicas (até 4)</td><td>+2 por keyword</td><td>cap em 8 — principal diferenciador temático</td></tr>
        <tr><td>Alcance nacional</td><td>+0,5</td><td>—</td></tr>
      </tbody>
    </table>
    <p>Keywords são buscadas no campo combinado de descrição + notas + tipo + território + nome (case-insensitive).</p>

    <h4>Gates (regras de corte)</h4>
    <ul>
      <li><strong>Gate amazônico</strong> — para UNICAFES PA/RO, sem bioma amazônica <em>nem</em> keyword amazônica o score é zerado. Evita falsos positivos com empresas/instituições nacionais de cobertura genérica.</li>
      <li><strong>Gate de catadores</strong> — para UNICATADORES, exige pelo menos uma keyword de resíduo, catador, circular ou PNRS. Sem isso o score é zerado.</li>
    </ul>

    <h4>Limiar de conexão</h4>
    <p>Edge é criada quando o score ≥ <code>4,5</code>. Esse valor foi calibrado iterativamente: limiares mais baixos (3,0) deixavam empresas nacionais broad-spectrum (grandes varejistas, montadoras) aparecerem como bridges; limiares mais altos (6,0) descartavam parceiros legítimos. O número final de stakeholders conectados (114 de 360) representa o ecossistema imediatamente relevante para a operação dos 4 parceiros, segundo os atributos disponíveis na base.</p>

    <h4>Origem da base de dados</h4>
    <p>A planilha <code>Stakeholder Ecosystem Mapping.xlsx</code> é um mapeamento do ecossistema brasileiro de filantropia climática, cooperativismo e bioeconomia, com 360 atores categorizados por setor, território, bioma, papel e potencial de articulação. A análise apresentada aqui aplica essa base a partir da lente da Trias e dos seus 4 parceiros MBO — toda a lógica de articulação foi derivada algoritmicamente dos 5 papéis Trias declarados no Annex 1 (Theory of Change), e não do schema original da planilha.</p>

    <h4>Camada institucional Trias (overlay)</h4>
    <p>Sobrepostos à base brasileira, foram adicionados ~50 atores institucionais que compõem a rede global da Trias e que não estavam na planilha de origem: doadores históricos (DGD, EU, Enabel, IFAD, AFD, IKI, GIZ), agri-agências da AgriCord (Agriterra, Solidaridad, FFD, Cresol Agri-Agency, Fert, We Effect, Asprodeb, UPA DI), redes belgas (11.11.11, NGO Federation, Coalition Against Hunger, Beyond Chocolate, BASCOF, Urgenci, UN Global Compact), diplomacia e mercado Bélgica-Brasil (Embaixada belga, AWEX, Hub Brussels, FIT, Belgalux, ABC, Itamaraty, SAF), parceiros corporativos (Colruyt, Boerenbond, Coopemapi) e MBOs irmãs do corredor amazônico SAM (Kallari, Unocace, Aprocam, AGROPAPA, CONPAPA, COOPAGROS), além de academia (Université de Liège, Vlerick, BRS). Fontes: <code>trias.ngo</code>, <code>agricord.org</code>, Annual Report Trias 2024, Annex 1 ToC, Narrative DGD.</p>
    <p>Esses nós representam <strong>vínculos institucionais existentes ou em construção formal</strong>, não potenciais — por isso são renderizados em <strong>quadrados azul-ardósia</strong> (vs. círculos coloridos do ecossistema brasileiro) e <strong>não entram no cálculo de adicionalidade</strong>. Podem ser ocultados pelo botão "Camada institucional Trias" no header (tecla I).</p>

    <h4>Modos de articulação Trias (derivados do papel)</h4>
    <p>Cada stakeholder conectado recebe um ou mais modos de articulação, mapeados aos 5 papéis declarados no Annex 1 (Theory of Change, seção sobre "complementary roles"):</p>
    <table class="formula-table">
      <thead><tr><th>Papel Trias</th><th>Modo operacional</th></tr></thead>
      <tbody>
        <tr><td>Process Facilitator</td><td>Ancoragem territorial · facilitação de OS/ID</td></tr>
        <tr><td>Thematic Advisor</td><td>Advisory técnico · pesquisa & evidência · MRV/dados</td></tr>
        <tr><td>Peer-to-Peer Facilitator</td><td>Co-implementação · mediação entre MBOs</td></tr>
        <tr><td>Bridge Builder</td><td>Co-financiamento · parceria estratégica · advocacy compartilhado · articulação multistakeholder</td></tr>
        <tr><td>Financer</td><td>Finanças mistas/inovadoras · mobilização direta de recursos</td></tr>
      </tbody>
    </table>

    <h4>Adicionalidade Trias</h4>
    <p>Para cada conexão (stakeholder, parceiro MBO), uma heurística estima quanto a mediação da Trias <em>adiciona valor</em> além do que aconteceria espontaneamente. O conceito vem da literatura de impact investing e cooperação para o desenvolvimento — Trias só justifica seu papel de hub onde a relação não emerge sem mediação.</p>
    <table class="formula-table">
      <thead><tr><th>Fator</th><th>Pontos</th></tr></thead>
      <tbody>
        <tr><td>Gap territorial — parceiro amazônico, stakeholder não-amazônico</td><td>+1,0</td></tr>
        <tr><td>Stakeholder nacional precisa aterrissagem amazônica</td><td>+0,5</td></tr>
        <tr><td>Stakeholder internacional / cross-border</td><td>+1,0</td></tr>
        <tr><td>MBO em desenvolvimento (UNICAFES PA/RO: +1,0; CSA: +0,5; UNICATADORES: 0)</td><td>varia</td></tr>
        <tr><td>Ator institucional de grande porte (multilateral, foundation, dev bank)</td><td>+0,7</td></tr>
        <tr><td>Intermediário financeiro (impact investor, fundo)</td><td>+0,4</td></tr>
        <tr><td>Relação aparentemente preexistente (menção em notas)</td><td>−1,5</td></tr>
      </tbody>
    </table>
    <p>Classificação final: <strong>Alta</strong> (≥ 2,0) · <strong>Média</strong> (1,0–2,0) · <strong>Baixa</strong> (&lt; 1,0). Score cap em 0–3.</p>

    <h4>O que a rede <em>não</em> captura</h4>
    <ul>
      <li>Relacionamentos efetivos (contratos, parcerias formais) — a base é descritiva, não relacional.</li>
      <li>Histórico de colaboração com a Trias ou entre os parceiros — só captura preexistência se for textualmente declarada nas notas.</li>
      <li>Alinhamento ideológico/político — ex.: posição sobre agronegócio, marco regulatório socioambiental.</li>
      <li>Capacidade técnica ou solidez institucional — não há indicadores quantitativos na base original.</li>
      <li>Os scores de adicionalidade são <em>proxies heurísticos</em>, não medidas validadas.</li>
    </ul>
    <p>Por isso, a rede deve ser lida como <strong>hipótese inicial de proximidade e prioridade</strong>, validada com a equipe Brasil/SAM e com os parceiros MBO antes de qualquer ação de articulação.</p>
  </div>
</div>

<script>
const DATA = __GRAPH_JSON__;
const $ = s => document.querySelector(s);
const $$ = s => Array.from(document.querySelectorAll(s));

// ========== Sidebar toggles ==========
$('#toggle-left').onclick = () => {
  document.body.classList.toggle('left-hidden');
  $('#toggle-left').classList.toggle('active', document.body.classList.contains('left-hidden'));
  resize();
};
$('#toggle-right').onclick = () => {
  document.body.classList.toggle('right-hidden');
  $('#toggle-right').classList.toggle('active', document.body.classList.contains('right-hidden'));
  resize();
};
document.addEventListener('keydown', ev => {
  if (ev.target.tagName === 'INPUT') return;
  if (ev.key === 'f' || ev.key === 'F') $('#toggle-left').click();
  if (ev.key === 'd' || ev.key === 'D') $('#toggle-right').click();
  if (ev.key === 'i' || ev.key === 'I') $('#btn-institutional').click();
  if (ev.key === 'Escape') clearFocus();
});

// Toggle camada institucional Trias — ESCONDIDA por padrão
$('#btn-institutional').onclick = () => {
  document.body.classList.toggle('institutional-on');
  const on = document.body.classList.contains('institutional-on');
  $('#btn-institutional').classList.toggle('btn-institutional-on', on);
  $('#btn-institutional').textContent = on
    ? 'Ocultar camada institucional'
    : 'Mostrar camada institucional Trias';
  if (on) { sim.alpha(0.3).restart(); }
};
// Texto inicial do botão
$('#btn-institutional').textContent = 'Mostrar camada institucional Trias';

// ========== Methodology modal ==========
$('#btn-methodology').onclick = () => $('#modal-bg').classList.add('show');
$('#modal-close').onclick = () => $('#modal-bg').classList.remove('show');
$('#modal-bg').onclick = (ev) => { if (ev.target === $('#modal-bg')) $('#modal-bg').classList.remove('show'); };

// ========== Tabs ==========
$$('.tab').forEach(t => t.addEventListener('click', () => {
  $$('.tab').forEach(x => x.classList.remove('active'));
  $$('.tab-pane').forEach(x => x.classList.remove('active'));
  t.classList.add('active');
  $('#pane-' + t.dataset.tab).classList.add('active');
}));

// ========== Stats panel ==========
$('#kpi-db').textContent = DATA.stats.total_db;
$('#kpi-conn').textContent = DATA.stats.connected;
$('#kpi-bridges').textContent = DATA.stats.bridges;
$('#kpi-tri').textContent = DATA.stats.tri;

function renderBars(containerSel, data, opts = {}) {
  const el = $(containerSel);
  el.innerHTML = '';
  const entries = Object.entries(data).filter(([k,v]) => v > 0).sort((a,b) => b[1]-a[1]);
  const max = Math.max(...entries.map(e => e[1]), 1);
  entries.forEach(([k, v]) => {
    const row = document.createElement('div');
    row.className = 'bar-row' + (opts.swatch ? ' with-sub' : '');
    const swatch = opts.swatch ? `<span class="swatch" style="background:${opts.swatch[k] || '#999'}"></span>` : '';
    const pct = (v / max * 100).toFixed(0);
    row.innerHTML = `
      <div class="lbl">${swatch}${k}</div>
      <div class="num">${v}</div>
      <div class="track"><div class="fill" style="width:${pct}%; ${opts.swatch && opts.swatch[k] ? 'background:'+opts.swatch[k] : ''}"></div></div>
    `;
    el.appendChild(row);
  });
}

// Per-partner connection counts
const partnerBars = {};
DATA.partners.forEach(p => { partnerBars[p.name] = DATA.stats.by_partner[p.id]; });
const partnerColors = {};
DATA.partners.forEach(p => { partnerColors[p.name] = 'var(--accent)'; });
renderBars('#bars-partner', partnerBars);

renderBars('#bars-sector', DATA.stats.dist_sector_all, {swatch: DATA.sectors});
renderBars('#bars-biome', DATA.stats.dist_biome_all);
renderBars('#bars-role', DATA.stats.dist_role_all);
renderBars('#bars-funding', DATA.stats.dist_funding_all);
renderBars('#bars-chain', DATA.stats.dist_chain_all);
renderBars('#bars-mode', DATA.dist_trias_mode);
renderBars('#bars-trias-role', DATA.dist_trias_role);
renderBars('#bars-addit', DATA.dist_additionality);
renderBars('#bars-inst-subcat', DATA.dist_institutional_subcat);
renderBars('#bars-inst-country', DATA.dist_institutional_country);
$('#n-inst').textContent = DATA.n_institutional;

// ========== Rankings panel ==========
const bridgesEl = $('#bridges-list');
DATA.bridges.slice(0, 30).forEach(b => {
  const li = document.createElement('li');
  li.innerHTML = `<div><strong>${b.name}</strong> <span class="pill">${b.n}×</span></div>
                  <div class="meta">${b.type || '—'} · ${b.partners.join(' · ')}</div>`;
  li.onclick = () => selectNode(b.id, true);
  bridgesEl.appendChild(li);
});

const brokersEl = $('#brokers-list');
DATA.top_brokers.forEach(b => {
  const li = document.createElement('li');
  li.innerHTML = `<div><strong>${b.name}</strong> <span class="pill">${b.betweenness.toFixed(3)}</span></div>
                  <div class="meta">${b.sector || '—'}${b.type ? ' · ' + b.type : ''}</div>`;
  li.onclick = () => selectNode(b.id, true);
  brokersEl.appendChild(li);
});

const topBlocksEl = $('#top-blocks');
DATA.partners.forEach(p => {
  const lst = DATA.top_by_partner[p.id] || [];
  const det = document.createElement('details');
  det.style.cssText = 'background:var(--panel);border:1px solid var(--border);border-radius:6px;padding:9px 12px;margin-bottom:6px;';
  det.innerHTML = `<summary style="cursor:pointer;font-size:11.5px;color:var(--text-2);font-weight:600;">${p.name} — ${lst.length} no top</summary>`;
  const ol = document.createElement('ul');
  ol.style.cssText = 'list-style:none;padding:0;margin:8px 0 0 0;';
  lst.forEach(item => {
    const li = document.createElement('li');
    li.style.cssText = 'padding:5px 0;border-bottom:1px solid var(--border);font-size:11px;cursor:pointer;';
    li.innerHTML = `<strong>${item.name}</strong> <span class="pill">${item.score}</span><br><span style="color:var(--muted);font-size:10.5px;">${item.type || ''}</span>`;
    li.onclick = () => selectNode(item.id, true);
    ol.appendChild(li);
  });
  det.appendChild(ol);
  topBlocksEl.appendChild(det);
});

// ========== Filters ==========
const filterState = {
  sector: new Set(DATA.filter_options.sector),
  biome: new Set(DATA.filter_options.biome),
  role: new Set(DATA.filter_options.role),
  funding_role: new Set(DATA.filter_options.funding_role),
  trias_mode: new Set(DATA.filter_options.trias_mode),
  additionality: new Set(DATA.filter_options.additionality),
  partner: 'ALL',
  search: '',
};

const FILTER_CONTAINERS = {
  sector: '#filter-sector', biome: '#filter-biome',
  role: '#filter-role', funding_role: '#filter-funding',
  trias_mode: '#filter-mode', additionality: '#filter-addit',
};

function getNodeTokens(n, group) {
  if (group === 'trias_mode') return n.trias_modes || [];
  if (group === 'additionality') return n.additionality_class ? [n.additionality_class] : [];
  const field = group === 'biome' ? 'biome' : group;
  const v = n[field] || '';
  return String(v).split(/[,;]/).map(s => s.trim()).filter(Boolean);
}

function buildFilter(group, options, colorMap = null) {
  const c = $(FILTER_CONTAINERS[group]); c.innerHTML = '';
  const counts = {};
  DATA.nodes.forEach(n => {
    if (n.kind !== 'stakeholder') return;
    getNodeTokens(n, group).forEach(tok => {
      if (options.includes(tok)) counts[tok] = (counts[tok] || 0) + 1;
    });
  });
  options.forEach(opt => {
    const lab = document.createElement('label'); lab.className = 'check';
    const sw = colorMap && colorMap[opt] ? `<span class="swatch" style="background:${colorMap[opt]}"></span>` : '';
    lab.innerHTML = `<input type="checkbox" checked>${sw}<span class="label">${opt || '—'}</span><span class="count">${counts[opt] || 0}</span>`;
    c.appendChild(lab);
    lab.querySelector('input').addEventListener('change', e => {
      if (e.target.checked) filterState[group].add(opt);
      else filterState[group].delete(opt);
      applyFilters();
    });
  });
}
buildFilter('sector', DATA.filter_options.sector, DATA.sectors);
buildFilter('biome', DATA.filter_options.biome);
buildFilter('role', DATA.filter_options.role);
buildFilter('funding_role', DATA.filter_options.funding_role);
buildFilter('trias_mode', DATA.filter_options.trias_mode);
buildFilter('additionality', DATA.filter_options.additionality);

$$('.filter-actions button').forEach(btn => btn.addEventListener('click', () => {
  const g = btn.dataset.group, a = btn.dataset.action;
  if (a === 'all') {
    filterState[g] = new Set(DATA.filter_options[g] || []);
  } else {
    filterState[g] = new Set();
  }
  $$(FILTER_CONTAINERS[g] + ' input').forEach((inp, i) => {
    inp.checked = filterState[g].has(DATA.filter_options[g][i]);
  });
  applyFilters();
}));

// Partner chips
const partnerFiltersEl = $('#partner-filters');
DATA.partners.forEach(p => {
  const span = document.createElement('span');
  span.className = 'partner-chip';
  span.dataset.partner = p.id;
  span.textContent = p.name;
  partnerFiltersEl.appendChild(span);
});
$$('.partner-chip').forEach(chip => chip.addEventListener('click', () => {
  $$('.partner-chip').forEach(c => c.classList.remove('active'));
  chip.classList.add('active');
  filterState.partner = chip.dataset.partner;
  applyFilters();
}));

$('#search').addEventListener('input', e => {
  filterState.search = e.target.value.toLowerCase().trim();
  applyFilters();
});

// ========== D3 setup ==========
const svg = d3.select('#net');
const stage = $('#stage');
const W = () => stage.clientWidth;
const H = () => stage.clientHeight;

const root = svg.append('g').attr('class', 'root');
const linkLayer = root.append('g').attr('class', 'links');
const nodeLayer = root.append('g').attr('class', 'nodes');

const zoom = d3.zoom().scaleExtent([0.2, 6]).on('zoom', ev => {
  root.attr('transform', ev.transform);
});
svg.call(zoom);

function partnerAnchors() {
  const cx = 0, cy = 0;  // origin-centered space
  const r = Math.min(W(), H()) * 0.24;
  return {
    UNICAFES_PA:  {x: -r,        y: -r * 0.9},
    UNICAFES_RO:  {x:  r,        y: -r * 0.9},
    CSA_BRASIL:   {x: -r,        y:  r * 0.9},
    UNICATADORES: {x:  r,        y:  r * 0.9},
  };
}

const nodes = DATA.nodes.map(n => ({...n}));
const links = DATA.links.map(l => ({...l}));

const anchors = partnerAnchors();
nodes.forEach(n => {
  if (n.kind === 'partner') {
    n.fx = anchors[n.id].x; n.fy = anchors[n.id].y;
    n.x = n.fx; n.y = n.fy;
  }
});

// Layout-alvo dos nós institucionais por subcategoria (em coordenadas relativas
// ao centro). Pelo menos 1.5x o raio do cluster MBO para ficarem na periferia.
const INSTITUTIONAL_LAYOUT = {
  'Doador institucional':         { x:  0.0, y: -1.6 },  // topo
  'Rede peer (AgriCord)':         { x:  1.6, y: -0.8 },  // topo-direita
  'Rede belga/europeia':          { x: -1.6, y: -0.4 },  // esquerda
  'Diplomacia/mercado':           { x:  1.6, y:  0.8 },  // direita-baixo
  'Setor privado parceiro':       { x: -1.6, y:  0.8 },  // esquerda-baixo
  'Setor privado parceiro · SAM': { x:  0.0, y:  1.6 },  // base (perto da Amazônia)
  'Academia':                     { x: -0.9, y: -1.4 },  // topo-esquerda
};

function instTarget(d, axis) {
  if (d.kind !== 'trias_institutional') return 0;
  const t = INSTITUTIONAL_LAYOUT[d.subcat];
  if (!t) return 0;
  const scale = Math.min(W(), H()) * 0.32;
  return t[axis] * scale;
}

const sim = d3.forceSimulation(nodes)
  .force('link', d3.forceLink(links).id(d => d.id)
                   .distance(l => {
                     if (l.kind === 'peer') return 220;
                     if (l.kind === 'institutional') return 160;
                     return 110 + (8 - l.weight) * 12;
                   })
                   .strength(l => {
                     if (l.kind === 'peer') return 0.04;
                     if (l.kind === 'institutional') return 0.08;
                     return 0.45;
                   }))
  .force('charge', d3.forceManyBody().strength(d => {
    if (d.kind === 'partner') return -1500;
    if (d.kind === 'trias_institutional') return -120;
    return -240;
  }))
  .force('center', d3.forceCenter(0, 0).strength(0.015))
  .force('collide', d3.forceCollide().radius(d => nodeRadius(d) + 4))
  .force('inst-x', d3.forceX(d => instTarget(d, 'x')).strength(d => d.kind === 'trias_institutional' ? 0.18 : 0))
  .force('inst-y', d3.forceY(d => instTarget(d, 'y')).strength(d => d.kind === 'trias_institutional' ? 0.18 : 0))
  .alphaDecay(0.025);

function nodeRadius(d) {
  if (d.kind === 'partner') return 20;
  if (d.kind === 'trias_institutional') return 7;
  return 5 + (d.n_partners || 1) * 3.5;
}

const linkSel = linkLayer.selectAll('path.link')
  .data(links).join('path')
  .attr('class', d => 'link ' + (d.kind === 'peer' ? 'peer ' : '')
                              + (d.kind === 'institutional' ? 'institutional ' : ''))
  .attr('stroke-width', d => {
    if (d.kind === 'peer') return 1.5;
    if (d.kind === 'institutional') return 1.2;
    return 0.5 + d.weight * 0.16;
  });

const nodeSel = nodeLayer.selectAll('g.node')
  .data(nodes).join('g')
  .attr('class', d => 'node '
                       + (d.kind === 'partner' ? 'partner ' : '')
                       + (d.kind === 'trias_institutional' ? 'trias-institutional ' : ''))
  .attr('data-id', d => d.id)
  .call(d3.drag().on('start', dragstart).on('drag', dragmove).on('end', dragend));

// halo
nodeSel.append('circle').attr('class', 'node-halo').attr('r', d => nodeRadius(d) + 8);
// shape: circle for stakeholders/partners; rounded square for institutional
nodeSel.each(function(d) {
  const r = nodeRadius(d);
  const sel = d3.select(this);
  if (d.kind === 'trias_institutional') {
    sel.append('rect').attr('class', 'node-circle')
       .attr('x', -r).attr('y', -r).attr('width', r*2).attr('height', r*2)
       .attr('rx', 3).attr('ry', 3)
       .attr('fill', d.color);
  } else {
    sel.append('circle').attr('class', 'node-circle')
       .attr('r', r).attr('fill', d.color);
  }
});
nodeSel.append('text')
  .attr('dy', d => nodeRadius(d) + 12)
  .text(d => {
    if (d.kind === 'partner') return d.name;
    if (d.kind === 'trias_institutional') {
      // só mostra label se nome curto, evita poluição
      return d.name.length > 20 ? d.name.slice(0, 17) + '…' : d.name;
    }
    // ecossistema brasileiro: só labels para tri-bridges e altos brokers
    if ((d.n_partners || 0) >= 3 || d.betweenness > 0.015) {
      return d.name.length > 32 ? d.name.slice(0, 29) + '…' : d.name;
    }
    return '';
  })
  .style('font-size', d => d.kind === 'trias_institutional' ? '9.5px' : null)
  .style('fill', d => d.kind === 'trias_institutional' ? 'var(--muted)' : null);

const tip = $('#tip');
nodeSel.on('mouseenter', function(ev, d) {
  d3.select(this).select('.node-halo').classed('show', true);
  tip.style.display = 'block';
  const parts = [];
  if (d.kind === 'partner') {
    parts.push(`<strong>${d.name}</strong>`);
    parts.push(`<div class="tip-meta">${d.type}<br>${d.territory}</div>`);
  } else {
    parts.push(`<strong>${d.name}</strong>`);
    const subs = [];
    if (d.type) subs.push(d.type);
    if (d.sector) subs.push(d.sector);
    if (subs.length) parts.push(`<div class="tip-meta">${subs.join(' · ')}</div>`);
    if (d.biome) parts.push(`<div class="tip-meta">${d.biome}</div>`);
  }
  tip.innerHTML = parts.join('');
})
.on('mousemove', function(ev) {
  const rect = stage.getBoundingClientRect();
  tip.style.left = (ev.clientX - rect.left + 12) + 'px';
  tip.style.top = (ev.clientY - rect.top + 12) + 'px';
})
.on('mouseleave', function() {
  d3.select(this).select('.node-halo').classed('show', false);
  tip.style.display = 'none';
})
.on('click', function(ev, d) {
  ev.stopPropagation();
  selectNode(d.id, false);
});
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

function dragstart(ev, d) { if (!ev.active) sim.alphaTarget(0.3).restart(); d.fx = d.x; d.fy = d.y; }
function dragmove(ev, d) { d.fx = ev.x; d.fy = ev.y; }
function dragend(ev, d) {
  if (!ev.active) sim.alphaTarget(0);
  if (d.kind !== 'partner') { d.fx = null; d.fy = null; }
}

// ========== Focus mode + inspector ==========
function selectNode(id, fromList) {
  const target = nodes.find(n => n.id === id);
  if (!target) return;
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
  // switch right tab to inspector
  $$('.tab').forEach(x => x.classList.remove('active'));
  $$('.tab-pane').forEach(x => x.classList.remove('active'));
  $('.tab[data-tab="inspector"]').classList.add('active');
  $('#pane-inspector').classList.add('active');
  if (fromList) {
    const tr = d3.zoomTransform(svg.node());
    const tx = W()/2 - target.x * tr.k;
    const ty = H()/2 - target.y * tr.k;
    svg.transition().duration(600)
      .call(zoom.transform, d3.zoomIdentity.translate(tx - W()/2, ty - H()/2).scale(tr.k));
  }
}

function clearFocus() {
  applyFilters();  // restore dim state from filters
  nodeSel.classed('focused', false);
  linkSel.classed('focus', false);
  $('#ins-content').style.display = 'none';
  $('#ins-empty').style.display = 'block';
}

function showInspector(n) {
  const ec = $('#ins-content'); const ee = $('#ins-empty');
  ee.style.display = 'none'; ec.style.display = 'block';
  if (n.kind === 'trias_institutional') {
    const rels = n.mbo_relations || {};
    let relsHtml = '';
    Object.entries(rels).forEach(([pid, reason]) => {
      const pname = DATA.partners.find(p => p.id === pid).name;
      relsHtml += `<div class="conn-row" style="display:block;">
        <div class="pname" style="font-weight:600;color:var(--text);">${pname}</div>
        <div style="color:var(--muted);font-size:11px;margin-top:3px;line-height:1.4;">${reason}</div>
      </div>`;
    });
    ec.innerHTML = `
      <div class="ins-header">
        <div class="name">${n.name}</div>
        <div class="type">Camada institucional Trias · ${n.subcat}</div>
      </div>
      <div class="ins-row"><div class="k">Nome completo</div><div class="v">${n.long_name || ''}</div></div>
      <div class="ins-row"><div class="k">Tipo</div><div class="v">${n.type || ''}</div></div>
      <div class="ins-row"><div class="k">País / origem</div><div class="v">${n.country || ''}</div></div>
      <div class="ins-row"><div class="k">Papel na operação Trias</div><div class="v">${n.role || ''}</div></div>
      <div class="ins-row"><div class="k">Interesse específico de cada MBO neste ator</div><div class="v">${relsHtml || '—'}</div></div>
      <div class="ins-row"><div class="k">Nota metodológica</div><div class="v" style="font-size:11px;color:var(--muted);">Vínculos institucionais existentes ou em construção formal — não potenciais. As conexões foram qualificadas caso-a-caso pelo interesse e perfil temático/territorial de cada MBO.</div></div>
    `;
    return;
  }
  if (n.kind === 'partner') {
    const priorities = (n.trias_priorities || []).map(p => `<li>${p}</li>`).join('');
    ec.innerHTML = `
      <div class="ins-header">
        <div class="name">${n.name}</div>
        <div class="type">${n.maturity || n.type || ''}</div>
      </div>
      <div class="ins-row"><div class="k">Nome completo</div><div class="v">${n.long_name || ''}</div></div>
      <div class="ins-row"><div class="k">Território</div><div class="v">${n.territory || ''}</div></div>
      <div class="ins-row"><div class="k">Unidades organizativas</div><div class="v">${n.units || '—'}</div></div>
      <div class="ins-row"><div class="k">Beneficiários</div><div class="v">${n.members || '—'}</div></div>
      <div class="ins-row"><div class="k">Perfil de gênero/juventude</div><div class="v">${n.women_pct || 0}% mulheres · ${n.youth_pct || 0}% jovens · ${n.staff || 0} staff</div></div>
      <div class="ins-row"><div class="k">Cadeias / produtos</div><div class="v">${n.value_chains || '—'}</div></div>
      <div class="ins-row"><div class="k">Papel estratégico no programa</div><div class="v">${n.strategic_role || ''}</div></div>
      <div class="ins-row"><div class="k">Prioridades de intervenção Trias</div><div class="v"><ul style="margin:4px 0 0 0;padding-left:18px;">${priorities}</ul></div></div>
      <div class="ins-row"><div class="k">Descrição</div><div class="v" style="font-size:11.5px;color:var(--muted);">${n.description || ''}</div></div>
    `;
    return;
  }
  // stakeholder
  let connsHtml = '';
  Object.entries(n.connections || {}).forEach(([pid, sc]) => {
    const pname = DATA.partners.find(p => p.id === pid).name;
    const br = (n.score_breakdowns && n.score_breakdowns[pid]) || [];
    let brHtml = '';
    br.forEach(b => {
      brHtml += `<div class="breakdown-line">${b.k}: ${b.v}<span class="pts">+${b.pts}</span></div>`;
    });
    connsHtml += `
      <div class="conn-row"><span class="pname">${pname}</span><span class="pscore">${sc}</span></div>
      ${brHtml ? `<div style="padding:4px 0 8px 10px;">${brHtml}</div>` : ''}
    `;
  });
  // Modos de articulação Trias
  let modesHtml = '';
  (n.trias_mode_details || []).forEach(m => {
    modesHtml += `<div class="conn-row" style="background:var(--accent-soft);display:block;">
      <div style="display:flex;justify-content:space-between;"><span class="pname">${m.mode}</span><span class="pscore" style="color:var(--muted);font-size:10px;">${m.trias_role}</span></div>
      <div style="color:var(--muted);font-size:10.5px;margin-top:3px;">${m.desc}</div>
    </div>`;
  });
  // Conexões com breakdown + adicionalidade
  connsHtml = '';
  Object.entries(n.connections || {}).forEach(([pid, sc]) => {
    const pname = DATA.partners.find(p => p.id === pid).name;
    const br = (n.score_breakdowns && n.score_breakdowns[pid]) || [];
    const ad = (n.additionality_per_partner && n.additionality_per_partner[pid]) || {};
    let brHtml = '';
    br.forEach(b => {
      brHtml += `<div class="breakdown-line">${b.k}: ${b.v}<span class="pts">+${b.pts}</span></div>`;
    });
    const adClass = ad.class || '—';
    const adColor = adClass === 'Alta' ? 'var(--accent)' : adClass === 'Média' ? '#9b7bb0' : 'var(--muted)';
    const adFactors = (ad.factors || []).map(f => `<li>${f}</li>`).join('');
    connsHtml += `
      <div class="conn-row"><span class="pname">${pname}</span><span class="pscore">score ${sc} · adic. <b style="color:${adColor}">${adClass} (${ad.score || 0})</b></span></div>
      <div style="padding:4px 0 10px 10px;font-size:10.5px;">
        ${brHtml ? `<div style="margin-bottom:5px;color:var(--muted);">${brHtml}</div>` : ''}
        ${adFactors ? `<div style="color:var(--muted-2);"><b>Fatores de adicionalidade:</b><ul style="margin:2px 0 0 14px;padding:0;">${adFactors}</ul></div>` : ''}
      </div>
    `;
  });
  ec.innerHTML = `
    <div class="ins-header">
      <div class="name">${n.name}</div>
      <div class="type">${n.type || '—'} · ${n.sector || '—'}</div>
    </div>
    <div class="ins-row"><div class="k">Território</div><div class="v">${n.territory || '—'}</div></div>
    <div class="ins-row"><div class="k">Bioma primário</div><div class="v">${n.biome || '—'}</div></div>
    ${n.biome_secondary ? `<div class="ins-row"><div class="k">Bioma secundário</div><div class="v">${n.biome_secondary}</div></div>` : ''}
    <div class="ins-row"><div class="k">Impact area</div><div class="v">${n.impact || '—'}</div></div>
    <div class="ins-row"><div class="k">Papel no ecossistema</div><div class="v">${n.role || '—'}</div></div>
    <div class="ins-row"><div class="k">Posição na cadeia</div><div class="v">${n.value_chain || '—'}</div></div>
    <div class="ins-row"><div class="k">Funding role</div><div class="v">${n.funding_role || '—'}</div></div>
    <div class="ins-row"><div class="k">HQ</div><div class="v">${n.hq || '—'}</div></div>
    ${n.url ? `<div class="ins-row"><div class="k">Site</div><div class="v"><a href="${n.url}" target="_blank">${n.url}</a></div></div>` : ''}
    ${n.description ? `<div class="ins-row"><div class="k">Descrição</div><div class="v" style="font-size:11.5px;color:var(--muted);">${n.description}</div></div>` : ''}
    <div class="ins-row"><div class="k">Adicionalidade Trias (média)</div><div class="v"><b>${n.additionality_class || '—'}</b> · score ${n.additionality_avg || 0}/3</div></div>
    ${modesHtml ? `<div class="ins-row"><div class="k">Modos de articulação que a Trias pode mobilizar</div><div class="v">${modesHtml}</div></div>` : ''}
    <div class="ins-row"><div class="k">Conexões com parceiros · breakdown do score · adicionalidade</div><div class="v">${connsHtml}</div></div>
    ${n.notes ? `<div class="ins-row"><div class="k">Notas estratégicas (do mapeamento)</div><div class="v" style="font-size:11.5px;color:var(--muted);">${n.notes}</div></div>` : ''}
  `;
}

// ========== Apply filters ==========
function applyFilters() {
  const partner = filterState.partner;
  const partnerConnected = new Set([partner]);
  if (partner !== 'ALL') {
    links.forEach(l => {
      const s = l.source.id || l.source;
      const t = l.target.id || l.target;
      if (s === partner) partnerConnected.add(t);
      if (t === partner) partnerConnected.add(s);
    });
  }
  const q = filterState.search;
  nodeSel.classed('dim', n => {
    if (n.kind === 'partner') return false;
    if (n.kind === 'trias_institutional') {
      // institucionais só são afetados por filtro de parceiro e busca
      if (partner !== 'ALL' && !partnerConnected.has(n.id)) return true;
      if (q && !n.name.toLowerCase().includes(q) &&
              !(n.role || '').toLowerCase().includes(q)) return true;
      return false;
    }
    if (partner !== 'ALL' && !partnerConnected.has(n.id)) return true;
    const sects = (n.sector || '').split(/[,;]/).map(s => s.trim());
    if (!sects.some(s => filterState.sector.has(s))) return true;
    const biomes = (n.biome || '').split(/[,;]/).map(s => s.trim()).filter(Boolean);
    if (biomes.length > 0 && !biomes.some(b => filterState.biome.has(b))) return true;
    const roles = (n.role || '').split(/[,;]/).map(s => s.trim()).filter(Boolean);
    if (roles.length > 0 && !roles.some(r => filterState.role.has(r))) return true;
    if (n.funding_role && !filterState.funding_role.has(n.funding_role)) return true;
    const modes = n.trias_modes || [];
    if (modes.length > 0 && !modes.some(m => filterState.trias_mode.has(m))) return true;
    if (n.additionality_class && !filterState.additionality.has(n.additionality_class)) return true;
    if (q && !n.name.toLowerCase().includes(q) &&
            !(n.description || '').toLowerCase().includes(q)) return true;
    return false;
  });
  linkSel.classed('dim', l => {
    const s = l.source.id || l.source;
    const t = l.target.id || l.target;
    const sNode = nodes.find(n => n.id === s);
    const tNode = nodes.find(n => n.id === t);
    const sDim = d3.select(`g.node[data-id="${s}"]`).classed('dim');
    const tDim = d3.select(`g.node[data-id="${t}"]`).classed('dim');
    return sDim || tDim;
  });
}

// ========== Controls ==========
$('#btn-zoom-in').onclick = () => svg.transition().call(zoom.scaleBy, 1.4);
$('#btn-zoom-out').onclick = () => svg.transition().call(zoom.scaleBy, 0.7);
$('#btn-reset').onclick = () => svg.transition().duration(600).call(zoom.transform, d3.zoomIdentity);
let physOn = true;
$('#btn-physics').onclick = () => {
  physOn = !physOn;
  $('#btn-physics').textContent = physOn ? '⏸' : '▶';
  if (physOn) sim.alphaTarget(0.05).restart();
  else sim.alphaTarget(0).stop();
};

function resize() {
  setTimeout(() => {
    const a = partnerAnchors();
    nodes.forEach(n => { if (n.kind === 'partner' && a[n.id]) { n.fx = a[n.id].x; n.fy = a[n.id].y; } });
    sim.alpha(0.3).restart();
  }, 300);
}
window.addEventListener('resize', resize);

// Initial centering
svg.call(zoom.transform, d3.zoomIdentity.translate(W()/2, H()/2));
sim.alpha(1).restart();
</script>
</body>
</html>
"""

html = HTML.replace("__GRAPH_JSON__", graph_json_str)
(OUT_NET / "ecossistema_trias_brasil.html").write_text(html, encoding="utf-8")
(OUT_PAGES / "index.html").write_text(html, encoding="utf-8")
print(f"HTML: {OUT_PAGES / 'index.html'}")

# ========== 9. Relatório analítico ==========
def fmt_pct(n, total):
    if total == 0: return "0,0%"
    return f"{n/total*100:.1f}%".replace(".", ",")

bridges_list = sorted(bridges_data, key=lambda x: -x["n"])

md = []
md.append("# Ecossistema Trias Brasil — DGD 2027–2031\n")
md.append("Descrição analítica do ecossistema em que a Trias atuará no próximo "
          "ciclo programático, a partir do mapeamento de 360 organizações "
          "(`Stakeholder Ecosystem Mapping.xlsx`, última atualização 11/02/2026) "
          "e da rede ancorada nos 4 parceiros MBO (UNICAFES Pará, UNICAFES "
          "Rondônia, CSA Brasil, UNICATADORES).\n")
md.append("> **Como ler este documento.** A primeira seção descreve a composição "
          "geral do ecossistema, sem filtro. A segunda mostra o subconjunto "
          "diretamente relevante para os parceiros MBO (114 organizações). As "
          "seções seguintes detalham bridges, brokers e a assinatura "
          "característica de cada parceiro.\n\n")

# Metodologia
md.append("## Metodologia\n")
md.append("A rede é uma **projeção bipartite**: cada stakeholder externo recebe "
          "uma aresta ponderada para cada um dos 4 parceiros MBO se o "
          "alinhamento temático/territorial passar de um limiar. Não há "
          "ligações diretas entre stakeholders externos. A leitura visual "
          "privilegia o papel de cada um *em relação aos parceiros*.\n")
md.append("\n**Critérios de score**\n")
md.append("| Critério | Pontos | Observação |")
md.append("|---|---|---|")
md.append("| Bioma primário coincide | +3 | diferenciador territorial mais forte |")
md.append("| Bioma secundário coincide | +1 | — |")
md.append("| Impact area coincide | +1 | peso baixo: 244/360 têm \"Land, Food and Forest\" |")
md.append("| Role coincide (até 2) | +0,5/match | cap 1,0 |")
md.append("| Keywords específicas (até 4) | +2/keyword | cap 8 — principal diferenciador |")
md.append("| Alcance nacional | +0,5 | — |")
md.append(f"\n**Limiar para conexão**: score ≥ {THRESHOLD}.\n")
md.append("\n**Gates**\n")
md.append("- *Amazônico* — para UNICAFES PA/RO, sem bioma amazônica nem keyword "
          "amazônica o score zera. Evita falsos positivos com organizações "
          "nacionais broad-spectrum.")
md.append("- *Catadores* — para UNICATADORES, exige pelo menos uma keyword de "
          "resíduo, catador, circular ou PNRS.\n")
md.append("\n**O que a rede não captura**: relacionamentos formais (contratos, "
          "parcerias), histórico de colaboração, alinhamento ideológico, capacidade "
          "institucional. A rede é uma hipótese inicial de proximidade baseada em "
          "atributos declarados, a ser validada com a equipe Brasil/SAM.\n\n")

# Composição geral
md.append("## 1. Composição geral do ecossistema (n = 360)\n")
md.append("### Setor\n")
for k, v in stats_all["sector"].most_common():
    md.append(f"- **{k}** — {v} ({fmt_pct(v, 360)})")
md.append("\n### Tipo de organização (top 10)\n")
tipo_c = Counter()
for stk in records:
    if stk["type"]:
        tipo_c[stk["type"]] += 1
for k, v in tipo_c.most_common(10):
    md.append(f"- {k} — {v}")
md.append("\n### Bioma primário\n")
for k, v in stats_all["biome"].most_common():
    md.append(f"- {k} — {v} ({fmt_pct(v, 360)})")
md.append("\n### Funding role\n")
for k, v in stats_all["funding_role"].most_common():
    md.append(f"- {k} — {v} ({fmt_pct(v, 360)})")
md.append("\n### Papel no ecossistema (top 10)\n")
for k, v in stats_all["role"].most_common(10):
    md.append(f"- {k} — {v}")
md.append("\n### Posição na cadeia\n")
for k, v in stats_all["value_chain"].most_common():
    md.append(f"- {k} — {v} ({fmt_pct(v, 360)})")

# Leitura concreta
md.append("\n### Leitura concreta\n")
md.append("- **CSOs dominam** (128 / 360, 35,6%), seguidas por categoria "
          "\"Others\" (73) — que inclui organizações com classificação dupla ou "
          "atípica como UNICATADORES e movimentos sociais.")
md.append("- A **filantropia institucional** é o segundo maior pool (49 "
          "fundações + 30 alianças/iniciativas multistakeholder).")
md.append("- O setor privado mapeado tem **43 grandes empresas** vs. apenas 11 "
          "investidores de impacto e 4-5 SMEs — o ecossistema retratado "
          "privilegia o corporate, não o tecido empreendedor de pequeno e médio porte.")
md.append("- **Amazônia concentra o maior número de atores temáticos** (105 / 360, "
          "29%), mas Cerrado (78) e Caatinga (54) também têm presença "
          "significativa. Mata Atlântica, Pantanal, Pampa e Costeiro/marinho aparecem "
          "predominantemente como biomas secundários.")
md.append("- A **arquitetura de financiamento** é fortemente assimétrica: 198 "
          "grantees, 116 doadores, 37 regranters. A relação 1:5 entre regranter "
          "e grantees indica que a maior parte da captação se dá por relação "
          "direta doador→grantee, com pouco papel intermediário formal.")
md.append("- **Policy influence** é o papel mais declarado (105), confirmando "
          "que muitas organizações se posicionam como atores de incidência "
          "política — uma característica relevante para um ecossistema que "
          "discute marcos como PNRS, PNAE, MROSC, Plano Safra Familiar e Plano "
          "Nacional de Bioeconomia.\n")

# Subset conectado
md.append(f"\n## 2. Subconjunto conectado aos parceiros MBO (n = {len(results)})\n")
md.append(f"Do total de 360 organizações, **{len(results)} ({fmt_pct(len(results), 360)})** "
          f"têm score de alinhamento ≥ {THRESHOLD} com pelo menos um dos 4 parceiros. "
          f"Esse é o ecossistema imediatamente acionável pela Trias para mediação "
          f"de relacionamentos.\n")
md.append("\n### Distribuição da relevância\n")
md.append(f"- **{n_b1}** conectam exatamente 1 parceiro")
md.append(f"- **{n_b2}** conectam 2 parceiros (bridges duplas)")
md.append(f"- **{n_b3}** conectam 3 parceiros (tri-bridges)")
md.append(f"- **{n_b4}** conectam os 4 parceiros\n")

md.append("\n### Conexões por parceiro\n")
for pid, p in PARTNERS.items():
    n = per_partner_stats[pid]["n"]
    md.append(f"- **{p['name']}** — {n} organizações relevantes")
md.append("\nNote a assimetria: UNICAFES Pará e Rondônia concentram a maior "
          "densidade de potenciais aliados (compartilham território e tema), "
          "enquanto UNICATADORES tem um ecossistema próprio e relativamente "
          "isolado das demais — refletindo a clivagem entre as agendas de "
          "agricultura familiar/bioeconomia e a agenda de catadores/economia "
          "circular urbana.\n")

# Por parceiro — perfil enriquecido
md.append("\n## 3. As 4 MBOs parceiras — perfis e assinatura de ecossistema\n")
md.append("Os 4 parceiros têm maturidade, território e tema distintos. Compreender "
          "essas diferenças é pré-condição para a análise de adicionalidade. Os "
          "dados abaixo vêm do *Annex 1 — Theory of Change* e do *BRAZIL_DGD "
          "Narrative_DRAFT*.\n")
md.append("\n### Quadro comparativo\n")
md.append("| MBO | Maturidade | Território | Membros | Mulheres | Jovens | Staff | Conectados |")
md.append("|---|---|---|---|---|---|---|---|")
for pid, p in PARTNERS.items():
    md.append(f"| {p['name']} | {p['maturity']} | {p['territory']} | "
              f"{p['members']} | {p['women_pct']}% | {p['youth_pct']}% | "
              f"{p['staff']} | {per_partner_stats[pid]['n']} |")
md.append("\n**Leitura comparativa**\n")
md.append("- **UNICAFES Rondônia e CSA Brasil têm o perfil de gênero mais "
          "inclusivo** (56% e 70% mulheres respectivamente). UNICAFES Pará está "
          "abaixo (29%), o que justifica a prioridade de inclusão de mulheres "
          "como agenda da intervenção Trias.")
md.append("- **UNICATADORES é a única madura** e a única com staff substancial "
          "(28 funcionários vs. 2-3 dos outros). Trias não atua como "
          "organisational developer ali — atua como parceiro estratégico "
          "em pilotos de PES/carbono, EPR e finanças.")
md.append("- **CSA Brasil tem alcance nacional mais amplo** (19 estados) mas "
          "membership menor (540 + 3.600). É o canal preferencial para o pilar "
          "rural-urbano e a internacionalização (Urgenci).")
md.append("- **UNICAFES PA e RO compartilham território (Amazônia), cadeias "
          "(açaí, castanha, mel, cacau, café) e maturidade** — peer learning "
          "natural entre as duas é uma das alavancas estratégicas mais óbvias.\n")

for pid, p in PARTNERS.items():
    md.append(f"\n### {p['name']}\n")
    md.append(f"**{p['maturity']} · {p['territory']}**\n")
    md.append(f"- *Unidades*: {p['units']}")
    md.append(f"- *Membros*: {p['members']} · {p['women_pct']}% mulheres · {p['youth_pct']}% jovens")
    md.append(f"- *Staff*: {p['staff']}")
    md.append(f"- *Cadeias/produtos*: {p['value_chains']}")
    md.append(f"\n*Papel estratégico no programa*: {p['strategic_role']}\n")
    md.append("\n**Prioridades de intervenção Trias**")
    for pr in p["trias_priorities"]:
        md.append(f"- {pr}")
    md.append(f"\n**Composição do ecossistema conectado** ({per_partner_stats[pid]['n']} stakeholders):")
    for s, n in per_partner_stats[pid]["sectors"].items():
        md.append(f"- {s} — {n}")
    md.append("\n**Distribuição de adicionalidade** das conexões:")
    for cls, n in addit_class_per_partner[pid].most_common():
        md.append(f"- {cls} — {n}")
    md.append("\n**Top 5 por alinhamento**:")
    for item in top_by_partner[pid][:5]:
        md.append(f"- **{item['name']}** (score {item['score']}) — {item['type']}")
    md.append("")

# Bridges
md.append("\n## 4. Bridges — atores que conectam múltiplos parceiros\n")
md.append("Bridges são organizações com alinhamento simultâneo a dois ou mais "
          "parceiros MBO. Têm valor estratégico desproporcional para o "
          "reposicionamento da Trias como hub: uma ação de articulação que "
          "envolva um bridge propaga efeitos por múltiplos parceiros ao mesmo tempo.\n")
md.append(f"\n**Total**: {len(bridges_data)} bridges identificadas.\n")
md.append("\n### Tri-bridges (3 parceiros) — núcleo da articulação\n")
for b in bridges_data:
    if b["n"] >= 3:
        md.append(f"- **{b['name']}** ({b['type']}) — {' · '.join(b['partners'])}")
md.append("\n### Bridges duplas — destaques\n")
for b in bridges_data[:25]:
    if b["n"] == 2:
        md.append(f"- **{b['name']}** ({b['type']}) — {' · '.join(b['partners'])}")

# Brokers
md.append("\n\n## 5. Brokers — alta intermediação\n")
md.append("Betweenness centrality mede o quanto um nó está em caminhos curtos "
          "entre outros nós. Em redes ancoradas como esta, os 4 parceiros "
          "dominam por construção; entre os stakeholders externos, os "
          "brokers são os que mais agregam fluxo de informação/recursos potenciais.\n")
for b in graph_data["top_brokers"][:15]:
    md.append(f"- **{b['name']}** — betweenness {b['betweenness']:.4f} · {b['sector']}")

# Gaps
## Seção 6 — Lógica de intervenção Trias
md.append("\n\n## 6. Lógica de intervenção Trias e modos de articulação\n")
md.append("A Trias declara no *Annex 1 — Theory of Change* cinco papéis "
          "complementares no programa Brasil 2027–2031:\n")
md.append("- **Process Facilitator** — facilita processos de OS/ID "
          "(organisational strengthening / institutional development) junto às MBOs")
md.append("- **Thematic Advisor** — provê advisory técnico-temático em clima, "
          "inclusão, gestão financeira e desenvolvimento de negócios")
md.append("- **Peer-to-Peer Facilitator** — media intercâmbios e aprendizado "
          "entre MBOs e entre cooperativas afiliadas")
md.append("- **Bridge Builder** — conecta MBOs ao ecossistema relevante "
          "(setor privado, autoridades, academia, multilaterais)")
md.append("- **Financer** — financiamento direto e mobilização de recursos "
          "via co-funding e finanças mistas\n")
md.append("Cada stakeholder conectado tem um ou mais **modos de articulação** "
          "que correspondem a esses papéis. Os modos são derivados do papel "
          "declarado da organização no mapeamento e descrevem operacionalmente "
          "como a Trias pode mobilizar a relação.\n")
md.append("\n### Distribuição dos modos de articulação (114 stakeholders conectados)\n")
md.append("| Modo de articulação | Contagem | Papel Trias | Operacionalização |")
md.append("|---|---|---|---|")
mode_desc_map = {}
for _, mode, role, desc in TRIAS_MODE_RULES:
    if mode not in mode_desc_map:
        mode_desc_map[mode] = (role, desc)
for mode, n in mode_counter.most_common():
    role, desc = mode_desc_map.get(mode, ("—", "—"))
    md.append(f"| {mode} | {n} | {role} | {desc} |")

md.append("\n### Distribuição por papel Trias mobilizado\n")
for role, n in trias_role_counter.most_common():
    md.append(f"- **{role}** — {n}")

md.append("\n### Stakeholders por modo de articulação (top 5)\n")
by_mode = defaultdict(list)
for r in results:
    for m in r["trias_modes"]:
        by_mode[m["mode"]].append(r["stk"])
for mode, lst in sorted(by_mode.items(), key=lambda x: -len(x[1])):
    md.append(f"\n**{mode}** ({len(lst)})")
    for stk in lst[:5]:
        md.append(f"- {stk['name']} ({stk['type'] or '—'})")

## Seção 7 — Adicionalidade
md.append("\n\n## 7. Análise de adicionalidade Trias\n")
md.append("Adicionalidade é a medida de quanto a presença da Trias agrega valor "
          "*além* do que aconteceria sem ela. Em programas de cooperação para o "
          "desenvolvimento, é critério crescente de legitimidade — particularmente "
          "para um reposicionamento como hub/facilitador, em que a justificativa "
          "do papel deve ser explícita.\n")
md.append("\n### Definição operacional usada\n")
md.append("A adicionalidade Trias por conexão (stakeholder × parceiro MBO) é "
          "estimada por uma heurística que pondera:\n")
md.append("- **Gap territorial** — quanto mais distante o stakeholder do "
          "território do parceiro (ex.: amazônico), mais ponte é necessária")
md.append("- **Maturidade do MBO** — UNICAFES PA/RO (2º nível em consolidação) "
          "exigem mais bridging do que UNICATADORES (3º nível maduro)")
md.append("- **Escala/formalidade do stakeholder** — multilaterais, fundações "
          "grandes e dev banks precisam de tradução para chegar a uma MBO de "
          "primeiro ou segundo nível")
md.append("- **Penalização por preexistência** — se o stakeholder já é "
          "mencionado nas notas do parceiro, a relação existe e a adicionalidade "
          "Trias é baixa (papel limitado a coordenação)\n")
md.append(f"\n### Distribuição agregada (182 conexões)\n")
for cls, n in [("Alta", addit_class_counter.get("Alta", 0)),
               ("Média", addit_class_counter.get("Média", 0)),
               ("Baixa", addit_class_counter.get("Baixa", 0))]:
    pct = (n / 182 * 100) if 182 > 0 else 0
    md.append(f"- **{cls}** — {n} conexões ({pct:.0f}%)")

md.append("\n### Adicionalidade por parceiro\n")
md.append("| MBO | Alta | Média | Baixa | Total |")
md.append("|---|---|---|---|---|")
for pid, p in PARTNERS.items():
    counts = addit_class_per_partner[pid]
    total = sum(counts.values())
    md.append(f"| {p['name']} | {counts.get('Alta',0)} | {counts.get('Média',0)} "
              f"| {counts.get('Baixa',0)} | {total} |")

md.append("\n**Leitura da adicionalidade por parceiro**\n")
md.append("- **UNICAFES Pará e Rondônia** têm a maior proporção de conexões de "
          "alta adicionalidade. Lógico: são MBOs em consolidação, em território "
          "amazônico, com pouca capacidade própria de bridging para fora da "
          "região. O papel da Trias é mais transformador ali.")
md.append("- **CSA Brasil** tem adicionalidade média predominante. Sua escala "
          "nacional e maturidade relativa diminuem a necessidade de mediação "
          "Trias para alguns atores, mas a expansão para o Norte abre janelas "
          "de alta adicionalidade.")
md.append("- **UNICATADORES** tem adicionalidade média a baixa. A organização "
          "já é atora política reconhecida (CIISC, PNRS) e tem 28 staff — não "
          "precisa de Trias para acessar muitas das suas conexões. O papel Trias "
          "ali é mais de parceria estratégica em pilotos específicos (PES, "
          "carbono, EPR) do que de bridging.\n")

md.append("\n### Conexões de alta adicionalidade — onde Trias adiciona mais valor\n")
high_addit = []
for r in results:
    for pid, ad in r["additionality"].items():
        if ad["class"] == "Alta":
            high_addit.append({
                "stk": r["stk"], "partner": PARTNERS[pid]["name"],
                "score": ad["score"], "factors": ad["factors"]
            })
high_addit.sort(key=lambda x: -x["score"])
md.append("\nTop 15 conexões com maior score de adicionalidade:\n")
for h in high_addit[:15]:
    md.append(f"- **{h['stk']['name']} → {h['partner']}** (adic. {h['score']}/3) — "
              f"{h['stk']['type'] or '—'}. Fatores: {'; '.join(h['factors'])}")

md.append("\n\n## 8. Lacunas e observações estratégicas\n")
md.append("**(a) UNICATADORES como ilha temática.** Nenhum stakeholder do "
          "recorte amazônico (bioeconomia, agricultura familiar) se conecta a "
          "UNICATADORES com score ≥ 4,5. Isso confirma a hipótese de "
          "fragmentação do ecossistema apontada na árvore de causas-raiz "
          "(Annex 1, seção 1.3). Para a Trias atuar como hub entre as agendas "
          "rural-amazônica e urbano-circular, será preciso construir as pontes "
          "que hoje não existem — possivelmente via temas transversais como "
          "clima, gênero ou políticas públicas (PNRS + PNAE).\n")
md.append("**(b) Concentração da filantropia no eixo Rio-São Paulo.** A "
          "maioria das fundações e regranters mapeados tem sede em RJ ou SP. "
          "Para uma operação centrada na Amazônia rural, a captação local "
          "(BASA, Banco do Brasil, governos estaduais do Pará/Rondônia, FAPESP "
          "amazônica) aparece sub-representada no mapeamento e merece "
          "complementação.\n")
md.append("**(c) Setor privado: grandes empresas vs. negócios de impacto.** Há "
          "43 grandes empresas vs. 11 impact investors e poucos SMEs. Para o "
          "objetivo de fortalecer cadeias de valor da sociobiodiversidade e "
          "economia circular, o ecossistema retratado privilegia parceiros B2B "
          "de grande porte (Natura, Suzano, Vale, Colruyt) — útil para off-take "
          "agreements, mas insuficiente para construir um tecido de "
          "fornecedores intermediários. AMAZ, Yunus, ICE e iniciativas como o "
          "Sebrae merecem reforço.\n")
md.append("**(d) Academia e dados.** WRI Brasil, IPAM, ICV, CPI, Earth "
          "Innovation, Agroicone e INPE aparecem como bridges para os "
          "parceiros amazônicos — uma força do ecossistema. A integração entre "
          "esses centros e os parceiros MBO é, na prática, uma das maiores "
          "alavancas para evidência, MRV e advocacy baseada em dados.\n")
md.append("**(e) UNICAFES Nacional como hub interno.** A entidade nacional "
          "aparece como bridge tri-partite (PA, RO, CSA Brasil) e funciona "
          "como ponte natural dentro da rede UNICOPAS. Estratégia de "
          "fortalecimento da UNICAFES Nacional gera spillover para os três "
          "parceiros do recorte rural/peri-urbano.\n")

# Reading guide
md.append("\n## 9. Como usar este mapa\n")
md.append("**Para descrever o ecossistema na narrativa do programa.** As "
          "distribuições da seção 1 (sector, biome, role) fornecem os números "
          "agregados para a seção C.3 do BRAZIL_DGD Narrative_DRAFT. A "
          "assinatura por parceiro (seção 3) alimenta a seção D do mesmo "
          "documento.\n")
md.append("**Para identificar parceiros prioritários de articulação.** Os "
          "bridges (seção 4) são os candidatos naturais para o papel de hub "
          "que a Trias quer ocupar — articulação com 1 deles propaga para 2-3 "
          "parceiros MBO. Tri-bridges merecem entrevista qualitativa antes do "
          "início do ciclo.\n")
md.append("**Para sustentabilidade pós-2031.** A seção 6 indica onde o "
          "ecossistema é frágil (UNICATADORES isolado, captação local "
          "sub-representada, baixa densidade de impact investors) — informação "
          "que alimenta a estratégia de saída (Lacuna 4 do Diagnóstico "
          "Técnico).\n")

# Fix walrus issue
try:
    md_text = "\n".join(md)
    # remove walrus operator if accidentally embedded
    md_text = md_text.replace("DATA_TOP_BROKERS := graph_data[\"top_brokers\"]",
                              "graph_data['top_brokers']")
except Exception:
    md_text = "\n".join(md)

(OUT_NET / "analise_rede.md").write_text(md_text, encoding="utf-8")
print(f"Relatório: {OUT_NET / 'analise_rede.md'}")
