"""
Adiciona 15 instituições ao Stakeholder Ecosystem Mapping.xlsx (IDs 361-375)
para ampliar o ecossistema conectado à UNICATADORES. Ver memo:
outputs/network/ampliacao_unicatadores.md.

Esta rotina é one-shot: roda uma vez para gravar as linhas, depois fica
arquivada como documentação da operação.
"""
from openpyxl import load_workbook
from pathlib import Path

ROOT = Path("/home/user/trias_ToC")
XLSX = ROOT / "docs-referencia" / "Stakeholder Ecosystem Mapping.xlsx"

NEW = [
    {
        "id": 361, "name": "CEMPRE — Compromisso Empresarial para a Reciclagem",
        "sector": "Others", "type": "Industry platforms and sectoral initiatives",
        "description": (
            "Brazilian business coalition that promotes recycling and the circular "
            "economy through technical research, advocacy and capacity building. "
            "Founded in 1992, it brings together major consumer-goods companies "
            "committed to extended producer responsibility (EPR), supports waste "
            "picker cooperatives integration into reverse logistics chains and "
            "publishes the Ciclosoft survey on selective collection (coleta seletiva)."
        ),
        "territory": "National", "impact_area": "Buildings & Transport",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — operates in urban waste systems across municipalities "
            "rather than biome-delimited territories."
        ),
        "role": "Convenor/enablers, Technical assistance / research, Policy influence / regulation",
        "value_chain": "Midstream", "partnership": "Strategic partner",
        "hq": "São Paulo, SP, Brazil", "funding_role": "Re-granter",
        "url": "https://cempre.org.br/",
        "notes": (
            "Referência histórica do setor brasileiro de reciclagem e implementação "
            "da PNRS; conecta brand owners a cooperativas de catadores via acordos "
            "de logística reversa e EPR. Bridge natural para advocacy compartilhado "
            "com a UNICATADORES."
        ),
    },
    {
        "id": 362, "name": "WIEGO — Women in Informal Employment: Globalizing and Organizing",
        "sector": "Civil Society Organization (CSO)", "type": "Grassroots networks",
        "description": (
            "Global network focused on improving the status of the working poor, "
            "especially women, in the informal economy. Through research, organizing "
            "and policy advocacy it supports waste picker cooperatives worldwide, "
            "including significant work with Brazilian catador networks on social "
            "protection, gender equity and integration into municipal solid waste "
            "and recycling systems."
        ),
        "territory": "Cross-border / international", "impact_area": "Buildings & Transport",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — operates in urban informality and waste management "
            "systems globally."
        ),
        "role": "Policy influence / regulation, Community legitimacy / proximate leadership, Technical assistance / research",
        "value_chain": "Midstream", "partnership": "Strategic partner",
        "hq": "Manchester, UK", "funding_role": "Re-granter",
        "url": "https://www.wiego.org/",
        "notes": (
            "Referência internacional para a interseção catador + gênero + "
            "informalidade — alinhamento direto com a UNICATADORES (60% mulheres "
            "catadoras). Bridge para fóruns globais (OIT Recomendação 204, ILO "
            "Recommendation on Transition from the Informal to the Formal Economy)."
        ),
    },
    {
        "id": 363, "name": "ANCAT — Associação Nacional dos Catadores e Catadoras de Materiais Recicláveis",
        "sector": "Others", "type": "Grassroots networks",
        "description": (
            "National association of waste picker cooperatives and individual "
            "catadores, working alongside MNCR and Unicatadores in the National "
            "Committee for the Inclusion of Waste Pickers (CIISC). Coordinates "
            "technical training, cooperative strengthening and policy advocacy "
            "related to the PNRS, reverse logistics (logística reversa) and "
            "credit access for catador organizations."
        ),
        "territory": "National", "impact_area": "Buildings & Transport",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — urban waste systems across Brazilian municipalities."
        ),
        "role": "Community legitimacy / proximate leadership, Implementation capacity (delivery), Policy influence / regulation",
        "value_chain": "Downstream", "partnership": "Strategic partner",
        "hq": "São Paulo, SP, Brazil", "funding_role": "Grantee",
        "url": "https://ancat.org.br/",
        "notes": (
            "Co-coordenadora da CIISC ao lado de Unicatadores e MNCR — peer "
            "político natural da UNICATADORES em advocacy de PNRS e em "
            "capacitação de cooperativas."
        ),
    },
    {
        "id": 364, "name": "Ambipar Group",
        "sector": "Private sector", "type": "Private companies (large / leading)",
        "description": (
            "Brazilian-listed environmental services company operating across "
            "emergency response, waste valorization and circular economy. Through "
            "its Boomera division and reverse logistics (logística reversa) "
            "operations, it manages industrial and post-consumer waste streams, "
            "operates recycling plants and contracts with waste picker "
            "cooperatives across the country."
        ),
        "territory": "Cross-border / international", "impact_area": "Buildings & Transport",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — focus on urban waste streams nationwide."
        ),
        "role": "Strategic partner, Implementation capacity (delivery), Capital provider / Financial intermediary",
        "value_chain": "Downstream", "partnership": "Co-funding partner",
        "hq": "São Paulo, SP, Brazil", "funding_role": "Donor",
        "url": "https://ambipar.com/",
        "notes": (
            "Mencionada no diagnóstico Trias como um dos principais operadores "
            "privados de logística reversa em escala industrial. Potencial parceiro "
            "de off-take/contrato para cooperativas Unicatadores e pilotos de "
            "EPR e finanças circulares."
        ),
    },
    {
        "id": 365, "name": "Eureciclo",
        "sector": "Private sector", "type": "Industry platforms and sectoral initiatives",
        "description": (
            "Brazilian company that operates a recycling credit certification "
            "platform connecting brand owners with EPR (extended producer "
            "responsibility) obligations to waste picker cooperatives. Certifies "
            "packaging recovery volumes against the PNRS, transfers revenue to "
            "cooperatives and supports their operational and technological "
            "upgrading."
        ),
        "territory": "National", "impact_area": "Buildings & Transport",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — operates across Brazilian urban waste systems."
        ),
        "role": "Capital provider / Financial intermediary, Convenor/enablers, Implementation capacity (delivery)",
        "value_chain": "Midstream", "partnership": "Strategic partner",
        "hq": "São Paulo, SP, Brazil", "funding_role": "Re-granter",
        "url": "https://eureciclo.com.br/",
        "notes": (
            "Eureciclo conecta diretamente a obrigação de logística reversa de "
            "marcas a cooperativas de catadores — bridge financeiro-operacional "
            "para a rede UNICATADORES no cumprimento da PNRS e EPR."
        ),
    },
    {
        "id": 366, "name": "ABRELPE — Associação Brasileira de Empresas de Limpeza Pública e Resíduos Especiais",
        "sector": "Private sector", "type": "Industry platforms and sectoral initiatives",
        "description": (
            "Industry association representing private urban cleaning, solid waste "
            "(resíduo sólido) collection and special waste treatment companies in "
            "Brazil. Publishes the annual Solid Waste Outlook (Panorama dos "
            "Resíduos Sólidos), advises on PNRS implementation and represents the "
            "sector in regulatory and policy dialogues, often as counterpart to "
            "waste picker cooperatives in municipal contracting debates."
        ),
        "territory": "National", "impact_area": "Buildings & Transport",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — urban services across all municipalities."
        ),
        "role": "Policy influence / regulation, Convenor/enablers",
        "value_chain": "Downstream", "partnership": "Strategic partner",
        "hq": "São Paulo, SP, Brazil", "funding_role": "Grantee",
        "url": "https://abrelpe.org.br/",
        "notes": (
            "Contraparte setorial das cooperativas de catadores nos debates de "
            "contratação pública municipal de coleta seletiva e implementação "
            "da PNRS (logística reversa, EPR). Referência setorial obrigatória "
            "para a UNICATADORES em advocacy regulatório."
        ),
    },
    {
        "id": 367, "name": "Ambev",
        "sector": "Private sector", "type": "Private companies (large / leading)",
        "description": (
            "Largest brewery in Latin America with major operations across Brazil. "
            "Key brand owner with extended producer responsibility (EPR) "
            "obligations on aluminum cans, glass bottles and PET packaging. "
            "Operates the Reciclo platform and supports recycling cooperative "
            "networks through procurement of post-consumer material."
        ),
        "territory": "Cross-border / international", "impact_area": "Buildings & Transport",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — operates across all Brazilian urban packaging streams."
        ),
        "role": "Strategic partner, Implementation capacity (delivery)",
        "value_chain": "Downstream", "partnership": "Co-funding partner",
        "hq": "São Paulo, SP, Brazil", "funding_role": "Donor",
        "url": "https://www.ambev.com.br/",
        "notes": (
            "Maior cliente brasileiro de logística reversa de vidro e alumínio; "
            "programa Reciclo conecta a cooperativas de catadores — interlocutor "
            "natural para a UNICATADORES em pilotos de off-take e contratualização "
            "EPR."
        ),
    },
    {
        "id": 368, "name": "Coca-Cola Brasil",
        "sector": "Private sector", "type": "Private companies (large / leading)",
        "description": (
            "Brazilian operations of The Coca-Cola Company, including bottlers; "
            "leading brand owner with EPR obligations on PET, glass and aluminum "
            "packaging. Through the Coletivo Reciclagem program (operated with "
            "Instituto Coca-Cola Brasil) it supports waste picker cooperatives "
            "technical upgrading, income generation and integration into reverse "
            "logistics chains."
        ),
        "territory": "Cross-border / international", "impact_area": "Buildings & Transport",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — operates across all Brazilian urban packaging streams."
        ),
        "role": "Strategic partner, Implementation capacity (delivery)",
        "value_chain": "Downstream", "partnership": "Co-funding partner",
        "hq": "Rio de Janeiro, RJ, Brazil", "funding_role": "Donor",
        "url": "https://www.cocacolabrasil.com.br/",
        "notes": (
            "Programa Coletivo Reciclagem é histórico parceiro de cooperativas de "
            "catadores no Brasil; bridge natural para a UNICATADORES em pilotos "
            "de EPR, capacitação e contratualização da logística reversa."
        ),
    },
    {
        "id": 369, "name": "Braskem",
        "sector": "Private sector", "type": "Private companies (large / leading)",
        "description": (
            "Brazilian petrochemical company and one of the world's largest "
            "producers of polyolefins; signatory of the Operation Clean Sweep "
            "program and operator of post-consumer plastic recycling initiatives. "
            "Supports waste picker cooperative networks through technical training "
            "and procurement of recycled plastic streams under the PNRS."
        ),
        "territory": "Cross-border / international", "impact_area": "Buildings & Transport",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — operates across all Brazilian plastic packaging streams."
        ),
        "role": "Strategic partner, Implementation capacity (delivery)",
        "value_chain": "Midstream", "partnership": "Co-funding partner",
        "hq": "São Paulo, SP, Brazil", "funding_role": "Donor",
        "url": "https://www.braskem.com.br/",
        "notes": (
            "Maior produtor de plásticos do Brasil — diretamente impactada pelas "
            "obrigações de logística reversa de embalagens da PNRS. Potencial "
            "parceiro para pilotos de economia circular do plástico envolvendo "
            "cooperativas da UNICATADORES."
        ),
    },
    {
        "id": 370, "name": "Instituto Coca-Cola Brasil",
        "sector": "Funders", "type": "Philanthropy foundations",
        "description": (
            "Private institute backed by the Coca-Cola Brazil System with a "
            "long-standing portfolio dedicated to recycling and inclusive circular "
            "economy. Through the Coletivo Reciclagem program it directly funds "
            "and capacitates waste picker cooperatives across Brazil, focused on "
            "income generation, gender equity and digital tools for catador "
            "organizations."
        ),
        "territory": "National", "impact_area": "Buildings & Transport",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — funding deployed across Brazilian urban waste systems."
        ),
        "role": "Funder, Implementation capacity (delivery)",
        "value_chain": "Midstream", "partnership": "Capacity building partner",
        "hq": "Rio de Janeiro, RJ, Brazil", "funding_role": "Donor",
        "url": "https://www.institutococacola.org.br/",
        "notes": (
            "Um dos principais financiadores domésticos específicos de cooperativas "
            "de catadores via Coletivo Reciclagem — alta complementaridade direta "
            "com a agenda da UNICATADORES em renda, gênero e digitalização."
        ),
    },
    {
        "id": 371, "name": "Fundação Banco do Brasil (FBB)",
        "sector": "Funders", "type": "Philanthropy foundations",
        "description": (
            "Corporate foundation of Banco do Brasil; flagship Brazilian funder of "
            "solidarity economy initiatives, with decades of support to waste "
            "picker cooperatives (PROCAT) and family farming cooperatives. "
            "Operates the Banco de Tecnologias Sociais and finances networks of "
            "recycling cooperatives nationwide. Bridges urban catador agendas and "
            "rural family farming cooperativism."
        ),
        "territory": "National",
        "impact_area": "Buildings & Transport, \"Land, Food and Forest\"",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — national-level funder across multiple territories "
            "and themes."
        ),
        "role": "Funder, Convenor/enablers, Implementation capacity (delivery)",
        "value_chain": "Midstream", "partnership": "Co-funding partner",
        "hq": "Brasília, DF, Brazil", "funding_role": "Donor",
        "url": "https://www.fbb.org.br/",
        "notes": (
            "Histórico financiador da economia solidária e de cooperativas de "
            "catadores via PROCAT — um dos canais mais sólidos para captação "
            "doméstica em economia circular. Bridge potencial entre as agendas "
            "UNICATADORES, CSA Brasil e UNICAFES (todos cooperativismo)."
        ),
    },
    {
        "id": 372, "name": "BVRio — Bolsa Verde do Rio de Janeiro",
        "sector": "Civil Society Organization (CSO)", "type": "Multi-stakeholder initiatives",
        "description": (
            "Non-profit institution that designs and operates market-based "
            "environmental solutions. Created the Reverse Logistics Credit (CLR — "
            "Crédito de Logística Reversa) platform which monetizes EPR "
            "obligations, channeling revenue to waste picker cooperatives, and "
            "develops impact finance instruments for the circular economy and "
            "forest products."
        ),
        "territory": "Cross-border / international",
        "impact_area": "Buildings & Transport, \"Land, Food and Forest\"",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — operates market-based instruments across urban "
            "waste systems and forest products."
        ),
        "role": "Capital provider / Financial intermediary, Convenor/enablers, Policy influence / regulation",
        "value_chain": "Midstream", "partnership": "Strategic partner",
        "hq": "Rio de Janeiro, RJ, Brazil", "funding_role": "Re-granter",
        "url": "https://www.bvrio.org/",
        "notes": (
            "Operadora do Crédito de Logística Reversa (CLR) — mecanismo "
            "financeiro inovador que diretamente capitaliza cooperativas de "
            "catadores. Bridge essencial entre o pilar Trias de finanças "
            "mistas/inovadoras e o ecossistema UNICATADORES."
        ),
    },
    {
        "id": 373, "name": "Tetra Pak Brasil",
        "sector": "Private sector", "type": "Private companies (large / leading)",
        "description": (
            "Brazilian operations of the global packaging multinational; pioneer "
            "in long-life carton recycling (reciclagem) with a national network "
            "of cooperativas de catadores that recover its packaging, and major "
            "investor in domestic recycling infrastructure for multi-material "
            "streams."
        ),
        "territory": "Cross-border / international", "impact_area": "Buildings & Transport",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — packaging recovery operations across Brazilian "
            "urban systems."
        ),
        "role": "Strategic partner, Implementation capacity (delivery)",
        "value_chain": "Downstream", "partnership": "Co-funding partner",
        "hq": "São Paulo, SP, Brazil", "funding_role": "Donor",
        "url": "https://www.tetrapak.com/pt-br",
        "notes": (
            "Histórica articuladora da reciclagem de longa vida no Brasil, com "
            "programa próprio de apoio a cooperativas de catadores. Presença em "
            "Norte/Nordeste onde a infraestrutura de coleta é mais frágil — "
            "bridge potencial para a expansão regional da UNICATADORES."
        ),
    },
    {
        "id": 374, "name": "Plastic Bank",
        "sector": "Private sector", "type": "Multi-stakeholder initiatives",
        "description": (
            "Canadian social enterprise that monetizes plastic recycling through "
            "digital deposit-refund systems. Provides bonus payments and benefits "
            "to waste pickers (catadores) in exchange for plastic recovery in "
            "urban areas; pioneer model of payment-for-ecosystem-services (PES) "
            "applied to municipal solid waste streams."
        ),
        "territory": "Cross-border / international", "impact_area": "Buildings & Transport",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — urban plastic recovery operations in coastal cities "
            "globally."
        ),
        "role": "Capital provider / Financial intermediary, Implementation capacity (delivery)",
        "value_chain": "Midstream", "partnership": "Strategic partner",
        "hq": "Vancouver, Canada", "funding_role": "Re-granter",
        "url": "https://plasticbank.com/",
        "notes": (
            "Modelo internacional de PES para resíduos urbanos diretamente "
            "relevante para a prioridade Trias de pilotos de PES e mercados de "
            "carbono para serviços ecossistêmicos urbanos — referência "
            "metodológica para a UNICATADORES."
        ),
    },
    {
        "id": 375, "name": "INSEA — Instituto Nenuca de Desenvolvimento Sustentável",
        "sector": "Civil Society Organization (CSO)", "type": "Local implementers",
        "description": (
            "Brazilian NGO based in Belo Horizonte that has supported waste "
            "picker cooperatives since the 1980s. Co-architect of the Minas "
            "Gerais Bolsa Reciclagem state policy and historical mentor of the "
            "MNCR. Offers technical assistance, cooperative organizing and policy "
            "advocacy specialized in recycling (reciclagem) cooperative networks "
            "and circular economy."
        ),
        "territory": "National", "impact_area": "Buildings & Transport",
        "biome_primary": None, "biome_secondary": None,
        "biome_rationale": (
            "No biome lens — operates with urban catador cooperatives across "
            "Brazilian states, headquartered in Minas Gerais."
        ),
        "role": "Implementation capacity (delivery), Technical assistance / research, Community legitimacy / proximate leadership",
        "value_chain": "Downstream", "partnership": "Strategic partner",
        "hq": "Belo Horizonte, MG, Brazil", "funding_role": "Grantee",
        "url": "https://www.insea.org.br/",
        "notes": (
            "Mentora histórica do MNCR e referência nacional em desenvolvimento "
            "institucional de cooperativas de catadores — natural peer técnico "
            "para a UNICATADORES em capacitação e governança cooperativa."
        ),
    },
]

# Mapping: dict key -> sheet column index (1-based)
COL = {
    "id": 2, "name": 3, "sector": 4, "type": 5, "description": 6,
    "territory": 7, "impact_area": 8, "biome_primary": 9, "biome_secondary": 10,
    "biome_rationale": 11, "role": 12, "value_chain": 13, "partnership": 14,
    "hq": 15, "funding_role": 16, "url": 18, "notes": 19,
}

# --- Step 1: read computed IDs from the data-only view ---
wb_data = load_workbook(XLSX, data_only=True)
ws_data = wb_data["Stakeholder list"]
computed_ids = {}
for r in range(7, ws_data.max_row + 1):
    v = ws_data.cell(row=r, column=2).value
    if v is not None:
        computed_ids[r] = int(v)
wb_data.close()
print(f"Read {len(computed_ids)} computed IDs from data_only view")

# --- Step 2: load with formulas, replace ID column with static ints ---
wb = load_workbook(XLSX)
ws = wb["Stakeholder list"]
# Replace formulas with static integer IDs to avoid openpyxl wiping
# cached values on subsequent saves
for r, val in computed_ids.items():
    ws.cell(row=r, column=2, value=val)

start_row = ws.max_row + 1
# Find first truly empty row (max_row may include trailing blanks)
for r in range(7, ws.max_row + 2):
    if ws.cell(row=r, column=2).value is None and ws.cell(row=r, column=3).value is None:
        start_row = r
        break

print(f"Inserting {len(NEW)} rows starting at row {start_row}")
for i, rec in enumerate(NEW):
    row = start_row + i
    for key, col in COL.items():
        val = rec.get(key)
        if val is not None:
            ws.cell(row=row, column=col, value=val)
    # Always set MEL type column (col 20) to em-dash like others
    ws.cell(row=row, column=20, value="—")

wb.save(XLSX)
print(f"Saved {XLSX}")
print(f"New max_row: {ws.max_row}")
