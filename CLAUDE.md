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

- **`data/GTO__LIST_OF_ITEMS.xlsx`** — **fonte primária**. O usuário pode largar o
  Excel novo em `data/` com **qualquer nome** (`GTO*.xlsx`); o script pega o mais
  recente (`resolve_source`) e o renomeia para o nome canônico (`consolidate_source`),
  removendo cópias antigas. Colunas relevantes na linha 2:
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
| **Status Anterior** | Snapshot **baseline fixo** em `status_history.json` | Referência antiga **fixa** (não desloca o "atual" para "anterior" a cada update). |

- **Status Anterior é FIXO** no snapshot marcado com `"baseline": true` (por padrão o
  primeiro, o estado pré-atualização do repositório). Para avançar a referência no
  futuro, marque outro snapshot com `"baseline": true` (e remova a flag do antigo).
- **Status 8 não existe:** `"8 - Awaiting full B05 completion"` é normalizado para
  `"7 - Missing Vacuum Test or Sign"` em toda parte (`STATUS_REMAP` / `norm_status`).
- Item **encerrado** = Status Atual em `{"1 - Validated by ICN", "2 - Not Blocking"}`
  (`CLOSED_SET`). KPIs e resumo por responsável usam sempre o **Status Atual**.
- O script grava cada atualização com mudança como um snapshot em `status_history.json`;
  a **linha do tempo por item** (`build_item_histories`) alimenta o fluxograma do HTML.
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

Gerada em `write_html()` a partir de **`templates/dashboard_template.html`**, que tem 5
placeholders: `__DATA_JSON__`, `__PROGRESS_JSON__`, `__LAST_UPDATED__`, `__ITEM_COUNT__`,
`__SOURCE_LABEL__`.
Os dados vão embutidos como `const DATA = [...]` (um objeto por item; campos: `item, ojx,
ajx, desc, cert, category, status, status_prev, analysis, analysis_pt, resp, note, closed,
hist`). `hist` = linha do tempo `[{t: rótulo, s: status}, ...]` para o fluxograma.

Recursos do dashboard (tudo client-side, sem backend):
- Tabela **responsiva** com colunas **Status Anterior** e **Status Atual** + selo
  `alterado`; agrupável por Responsible / Status / Category / None; visão **List** e **Board**.
- **Filtros** multi-seleção (Status, Responsible, Category) + busca textual.
- Checkbox **"Somente alterados"** (`#onlyChanged`) → mostra só os itens alterados
  (ignora o filtro de "encerrados").
- Botão **"◆ Ressaltar alterados"** (`#hlChanged`) → alterna a classe `hl-mode` em
  `#groups`, destacando em dourado as linhas/post-its com classe `changed-row`.
- **Fluxograma de status por item:** clicar no **nº do item** (na tabela ou nos
  post-its) abre um **modal** (`#modal`) com o fluxograma das mudanças de status do
  item ao longo do tempo (campo `hist` de cada objeto em `DATA`; nós coloridos por
  `STATUS_COLOR`, do mais antigo ao mais recente). Fecha no ✕, no fundo ou com Esc.
- **Abas** (`.tabs`): **📊 Dashboard** (tabela/filtros/gráficos) e **📈 Progresso**.
- **Aba Progresso** (`#tab-progress`, dados em `const PROGRESS = {...}` via
  `build_progress`): compara a **primeira fotografia (Inicial)** com a **atual** — anel de
  % concluído, cartões Antes × Hoje com delta colorido (melhor/pior), seção **"O que
  aconteceu"** (destaques + fluxo *de → para* dos itens que mudaram, calculado no
  cliente a partir de `status_prev`→`status` em `DATA`), gráfico de status Antes×Hoje,
  gráfico de evolução no tempo e tabela-resumo por snapshot.
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
