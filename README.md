# Metabase Contabilidade

Documentos de clientes e propostas comerciais.

## Propostas

- `propostas/marma-ltda/` – MARMA LTDA (confecção, Santa Cruz do Capibaribe/PE)
  - `01-proposta-planejamento-trabalhista.html` / `.pdf` – planejamento da regularização trabalhista
  - `02-proposta-acompanhamento-mensal.html` / `.pdf` – contabilidade, DP e jurídico mensal
  - `notas-tecnicas-internas.md` – fundamentos e premissas (uso interno, não enviar)

Logomarca: salvar em `assets/logo-metabase.png` e regerar os PDFs:

```sh
cd propostas/marma-ltda
for f in 0*.html; do chromium --headless --no-sandbox --no-pdf-header-footer --print-to-pdf="${f%.html}.pdf" "$f"; done
```
