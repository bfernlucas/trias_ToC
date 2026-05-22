# Ampliação do ecossistema — UNICATADORES

Memo de incorporação de **15 instituições** ao `Stakeholder Ecosystem Mapping.xlsx`
(IDs 361–375), endereçando a lacuna estrutural identificada na seção 8 do
`analise_rede.md`: das 360 organizações mapeadas, apenas 4 alcançavam o
limiar de alinhamento (score ≥ 4,5) com a UNICATADORES, configurando uma
**ilha temática** desconectada das demais agendas do programa.

## Por que essa expansão

O diagnóstico técnico aponta três frentes em que a fragmentação dói:

1. **UNICATADORES como ilha** — sem bridges com a agenda amazônica, embora as
   duas partilhem políticas federais (PNRS, PNAE, Plano Safra).
2. **Sub-representação da economia circular** — apenas 4 stakeholders
   conectados, contra 41–73 dos demais parceiros MBO. A rede atual
   subestima a densidade real do ecossistema brasileiro de resíduos
   sólidos urbanos.
3. **Lacuna nos pilares operacionais Trias para UNICATADORES** — pilotos
   de PES/carbono, EPR, contratação municipal e finanças circulares
   exigem interlocutores que hoje não aparecem no mapa (Eureciclo, BVRio,
   Ambipar, Plastic Bank).

Os 15 nós foram selecionados sob três critérios cumulativos: **(a)** são
organizações (não políticas ou programas — fora do escopo desta rodada);
**(b)** UNICATADORES já articula ou potencialmente articularia diretamente
com elas no ciclo 2027-2031; e **(c)** correspondem a pelo menos uma das
quatro prioridades de intervenção Trias declaradas no *Annex 1 — Theory of
Change*.

## Lista das 15 instituições adicionadas

Agrupadas por função estratégica:

### Infraestrutura técnica, advocacy e organização do setor (3)

| ID | Organização | Função para UNICATADORES |
|---|---|---|
| 361 | **CEMPRE** — Compromisso Empresarial para a Reciclagem | Coalizão empresarial histórica (1992), referência técnica em PNRS/EPR; conecta brand owners a cooperativas |
| 362 | **WIEGO** — Women in Informal Employment Globalizing and Organizing | Rede global de informalidade + gênero; direta para a agenda de 60% de mulheres catadoras |
| 363 | **ANCAT** — Associação Nacional dos Catadores e Catadoras de Materiais Recicláveis | Co-coordenadora da CIISC com Unicatadores e MNCR; peer político natural |

### EPR e logística reversa — operadores e plataformas (3)

| ID | Organização | Função para UNICATADORES |
|---|---|---|
| 364 | **Ambipar Group** (inclui Boomera) | Maior operador integrado de logística reversa no Brasil; contratante potencial para cooperativas |
| 365 | **Eureciclo** | Plataforma de créditos de reciclagem que monetiza obrigação EPR diretamente em receita para cooperativas |
| 366 | **ABRELPE** — Associação Brasileira de Empresas de Limpeza Pública | Contraparte setorial nos debates de contratação pública municipal e PNRS |

### Brand owners com obrigação de logística reversa (3)

| ID | Organização | Função para UNICATADORES |
|---|---|---|
| 367 | **Ambev** | Maior cliente brasileiro de vidro/alumínio; programa Reciclo articula cooperativas |
| 368 | **Coca-Cola Brasil** | Sistema Coca-Cola; programa Coletivo Reciclagem histórico parceiro de cooperativas |
| 369 | **Braskem** | Maior produtor de plásticos do país; obrigação direta de logística reversa de embalagens |

### Financiadores específicos da economia circular (3)

| ID | Organização | Função para UNICATADORES |
|---|---|---|
| 370 | **Instituto Coca-Cola Brasil** | Financiador doméstico mais consistente de cooperativas via Coletivo Reciclagem |
| 371 | **Fundação Banco do Brasil (FBB)** | Apoio histórico via PROCAT; bridge potencial entre UNICATADORES, CSA Brasil e UNICAFES (todos cooperativismo) |
| 372 | **BVRio** — Bolsa Verde do Rio de Janeiro | Crédito de Logística Reversa (CLR) — instrumento financeiro que capitaliza diretamente cooperativas |

### Bridges adicionais (3)

| ID | Organização | Função para UNICATADORES |
|---|---|---|
| 373 | **Tetra Pak Brasil** | Pioneira na reciclagem de longa vida; presença em Norte/Nordeste, bridge para expansão regional da UNICATADORES |
| 374 | **Plastic Bank** | Referência internacional de PES para resíduos urbanos — alinhamento direto com a prioridade Trias de pilotos PES/carbono urbanos |
| 375 | **INSEA** — Instituto Nenuca de Desenvolvimento Sustentável | Mentora histórica do MNCR; referência nacional em desenvolvimento institucional de cooperativas de catadores |

## Como entram no cálculo da rede

Os nós novos passam pelo mesmo pipeline da `build_network.py`: tokenização
do *blob* (description + notes + name + territory), gate de keyword
catador-específico (`waste pick`, `recicl`, `catador`, `pnrs`, `epr`,
`circular econom`, `logística reversa`, `urban`, `extended producer`), e
score combinado com bioma, impact area, role e alcance territorial. As 15
fichas foram redigidas com 3–5 keywords cada para garantir gatilho no
gate da UNICATADORES (limiar score ≥ 4,5).

A camada institucional Trias (`TRIAS_INSTITUTIONAL`) permanece separada e
**não** foi modificada — as 15 entram como stakeholders do ecossistema
brasileiro genérico, candidatas naturais a bridges para os demais
parceiros MBO via mecânica do score.

## Critérios deliberadamente fora do escopo

A pedido, esta rodada **não inclui**:

- **Políticas ou programas** (PNRS, CIISC, Bolsa Reciclagem MG, Coletivo
  Reciclagem) — esses ficam representados via as instituições que os
  operam.
- **Atores municipais e estaduais individuais** (SEMAD-MG, prefeituras).
- **Outros financiadores fora do recorte circular específico**
  (Itaú Social, Fundação Itaú) — relevantes mas mais difusos.
- **Operadores de menor escala** (Recicleiros, Polen, Triciclo) —
  candidatos a rodada futura se o recorte for ampliado.

## Lacunas residuais a validar

Com a equipe Trias Brasil:

- **Verificar relacionamentos existentes** das 15 organizações com a
  UNICATADORES — algumas podem ter contratos formais que mudam a
  classificação de adicionalidade (de "Alta" para "Baixa").
- **Confirmar localização do bridge UNICATADORES–UNICAFES** — Fundação
  Banco do Brasil aparece como candidato natural, mas pode haver outros.
- **EPR setoriais (ABIPLA, ABIVIDRO, ABIPET, ABIHPEC)** — propositadamente
  agregados sob a categoria de brand owners + Eureciclo; se a UNICATADORES
  vê valor em relação direta com PROs setoriais, adicionar em rodada
  separada.
