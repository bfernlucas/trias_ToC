# docs/ — GitHub Pages

Esta pasta hospeda a visualização interativa do **ecossistema Trias Brasil DGD 2027–2031**.

## Como publicar

1. Acesse **Settings → Pages** no repositório no GitHub.
2. Em **Source**, selecione `Deploy from a branch`.
3. Em **Branch**, escolha `claude/brazil-philanthropy-ecosystem-8WW8r` (ou faça merge para `main`) e o folder `/docs`.
4. Salve. Em poucos minutos a URL `https://bfernlucas.github.io/trias_toc/` ficará disponível.

## Conteúdo

- `index.html` — rede de stakeholders ancorada nos 4 parceiros MBO (UNICAFES Pará, UNICAFES Rondônia, CSA Brasil, UNICATADORES). Build em D3.js, autocontido.

## Iteração

O HTML é regenerado pelo script `outputs/network/build_network.py` a partir da
base `docs-referencia/Stakeholder Ecosystem Mapping.xlsx`. Para refinar
keywords/scoring, edite o script e rode:

```
python3 outputs/network/build_network.py
```

A saída é gravada simultaneamente em `outputs/network/ecossistema_trias_brasil.html`
e em `docs/index.html`.
