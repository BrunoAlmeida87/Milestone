# Contexto do projeto — Acompanhamento de itens J06 (SBR4)

> **Para a IA:** leia este arquivo antes de qualquer alteração. Ele explica o que é o
> projeto, como os dados fluem e como fazer as mudanças mais comuns sem quebrar nada.
> O idioma de trabalho com o usuário (Bruno) é **português**.

## 1. O que é

Painel de acompanhamento dos **218 itens do milestone J06** do submarino **SBR4**
(itens cujo `ActualJx = J06`). Mostra, por item, o **status anterior** e o **status
atual**, a análise, o responsável e a categoria. Existe em dois formatos:

- **`J06_Items_Analysis.xlsx`** — planilha (cabeçalho congelado + filtros).
- **`J06_Dashboard.html`** — dashboard interativo, publicado no **GitHub Pages**
  (`https://brunoalmeida87.github.io/Milestone/`, servido de `main` via `index.html`).

Ambos são **gerados** por `build_j06.py`. **Nunca edite os dois arquivos de saída à
mão** — edite a fonte/o template e rode o script.

## 2. Fluxo de dados

```
data/GTO__LIST_OF_ITEMS.xlsx  ─┐
data/status_history.json      ─┼─► build_j06.py ─► J06_Items_Analysis.xlsx
data/enrichment.json          ─┤                └► J06_Dashboard.html (via templates/dashboard_template.html)
templates/dashboard_template.html ┘
```

- **`data/GTO__LIST_OF_ITEMS.xlsx`** — **fonte primária** (o usuário substitui este
  arquivo, mantendo o nome, quando há atualização). Colunas relevantes na linha 2:
  `Item`, `Updated Status`, `Original Jx`, `Actual Jx`, `Description`, `Certificate`,
  `Context`. A data de atualização vem da célula *"Gerado em ..."* na linha 1.
- **`data/status_history.json`** — lista de *snapshots* de status. O último snapshot é
  o **Status Atual**; o penúltimo é o **Status Anterior**. O script acrescenta um novo
  snapshot **somente quando os status mudam** (idempotente).
- **`data/enrichment.json`** — dados curados à mão que **não existem no GTO**:
  `analysis` (EN), `analysis_pt` (PT), `resp` (responsável), `category`
  (`General J06` ou `B05`), `note`. Chave = número do item (string).

## 3. Lógica de status (regra central)

| Coluna | Origem | Regra |
|---|---|---|
| **Status Atual** | `Updated Status` do GTO | Classificação vigente na versão mais recente. |
| **Status Anterior** | Snapshot imediatamente anterior em `status_history.json` | Classificação antes da última atualização, casada por `Item`. |

- Item **encerrado** = Status Atual em `{"1 - Validated by ICN", "2 - Not Blocking"}`
  (`CLOSED_SET`). KPIs e resumo por responsável usam sempre o **Status Atual**.
- Item **alterado** = tem status anterior e ele difere do atual (`changed(d)` no HTML;
  aba **Status Changes** no Excel).
- Item **novo** (sem histórico anterior) → Status Anterior aparece como `— (novo item)`.

## 4. Saída Excel (`J06_Items_Analysis.xlsx`)

Gerada em `write_excel()`. Três abas:
- **Merged Analysis** — 12 colunas (`HEADERS`): Item, Context, OriginalJx, ActualJx,
  Category, Description, Certificate / Evidence, **Status Anterior**, **Status Atual**,
  Analysis (EN), Responsible, Analysis (PT). Cabeçalho **congelado** (`freeze_panes="A4"`),
  **autofilter** em `A3:L{fim}`, células de status coloridas por `STATUS_FILL`, "Status
  Atual" em **negrito** quando mudou. Data da última atualização no **canto superior
  direito** (célula `H2`).
- **Summary by Responsible** — pendências abertas agrupadas por responsável.
- **Status Changes** — apenas os itens que mudaram de status na última atualização.

Não há fórmulas na saída → **não precisa** rodar `recalc.py`.

## 5. Saída HTML (`J06_Dashboard.html`)

Gerada em `write_html()` a partir de **`templates/dashboard_template.html`**, que tem 4
placeholders: `__DATA_JSON__`, `__LAST_UPDATED__`, `__ITEM_COUNT__`, `__SOURCE_LABEL__`.
Os dados vão embutidos como `const DATA = [...]` (um objeto por item; campos: `item, ojx,
ajx, desc, cert, category, status, status_prev, analysis, analysis_pt, resp, note, closed`).

Recursos do dashboard (tudo client-side, sem backend):
- Tabela **responsiva** com colunas **Status Anterior** e **Status Atual** + selo
  `alterado`; agrupável por Responsible / Status / Category / None; visão **List** e **Board**.
- **Filtros** multi-seleção (Status, Responsible, Category) + busca textual.
- Checkbox **"Somente alterados"** (`#onlyChanged`) → mostra só os itens alterados
  (ignora o filtro de "encerrados").
- Botão **"◆ Ressaltar alterados"** (`#hlChanged`) → alterna a classe `hl-mode` em
  `#groups`, destacando em dourado as linhas/post-its com classe `changed-row`.
- **Última atualização** no canto superior direito (`.updbadge`).
- Gráficos usam **Chart.js via CDN**; o código está protegido por
  `if(typeof Chart!=='undefined'){...}` — se o CDN falhar (ex.: offline), o resto do
  dashboard ainda renderiza.

## 6. Como fazer as alterações mais comuns

**Atualizar os dados (fluxo normal):**
1. Substituir `data/GTO__LIST_OF_ITEMS.xlsx` pelo arquivo novo (mesmo nome).
2. `python3 build_j06.py`
3. Commit + push na branch de trabalho → PR → merge na `main` (o Pages republica sozinho).

**Adicionar um status novo** (ex.: `9 - Xxx`): inclua a cor em `STATUS_FILL`
(`build_j06.py`) **e** em `STATUS_CLASS`/`STATUS_COLOR` no template; se precisar de nova
classe de badge, adicione o CSS `.b-xxx`.

**Preencher análise/responsável de um item novo:** edite `data/enrichment.json`
(chave = número do item) e rode o script. Itens sem entrada saem com análise em branco
(`— pending —`).

**Mudar colunas/formatação do Excel:** ajuste `HEADERS`/`COL_WIDTHS`/`write_excel()`.
**Mudar layout/recursos do HTML:** edite `templates/dashboard_template.html` (NÃO mexa
nos dados lá — eles vêm dos placeholders) e rode o script.

## 7. Ambiente / dependências

Python 3 com `openpyxl` e `pandas` (instale com `pip install openpyxl pandas` se faltar).
Para tirar screenshot do HTML offline, use o Chromium pré-instalado em
`/opt/pw-browsers/chromium-1194/chrome-linux/chrome` via Playwright.

## 8. Git / publicação

- Branch de trabalho padrão desta automação:
  `claude/data-analysis-status-tracking-sdbj1a`. Publicação no Pages é a partir da `main`.
- Após atualizar, o deploy do Pages ("pages build and deployment") roda automático no
  push da `main`. Verifique a conclusão em Actions.
- Documento complementar voltado ao usuário: **`README_ATUALIZACAO.md`**.
