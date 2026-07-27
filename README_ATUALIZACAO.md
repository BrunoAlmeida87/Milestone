# J06 — Acompanhamento de Status (SBR4)

Pipeline que lê o arquivo **GTO – List of Items** e regenera o Excel e o
dashboard HTML com duas colunas de status: **Status Anterior** e **Status Atual**.

## Como atualizar (fluxo mensal / sob demanda)

1. Coloque o arquivo novo na pasta [`data/`](data/) com o **mesmo nome**:
   `data/GTO__LIST_OF_ITEMS.xlsx` (substitua o existente).
2. Rode o pipeline:

   ```bash
   python3 build_j06.py
   ```

3. Serão regenerados automaticamente:
   - `J06_Items_Analysis.xlsx` — planilha com cabeçalho congelado e filtros.
   - `J06_Dashboard.html` — dashboard responsivo.

A cada execução com dados diferentes, o **Status Atual** vira **Status Anterior**
na próxima atualização, sem intervenção manual (o histórico fica em
`data/status_history.json`).

## Lógica de status (como foi calculado)

| Coluna | Origem | Regra |
|---|---|---|
| **Status Atual** | Coluna `Updated Status` do arquivo **GTO** (fonte primária) | Classificação vigente na versão mais recente do GTO. |
| **Status Anterior** | Snapshot imediatamente anterior (`data/status_history.json`) | Classificação vigente **antes** desta atualização, casada por número do `Item`. |

- **Snapshot base (1ª execução):** o Status Anterior foi extraído do estado
  anterior já existente no repositório (`J06_Items_Analysis.xlsx` original) —
  ou seja, é histórico real, não uma dedução. Nesta primeira atualização,
  **34 dos 218 itens** mudaram de status.
- **Item novo** (não existia no snapshot anterior): Status Anterior aparece como
  `— (novo item)`.
- **Encerrado / aberto:** um item é considerado *encerrado* (fora das pendências
  abertas) quando o Status Atual é `1 - Validated by ICN` ou `2 - Not Blocking`.
  Os KPIs e o resumo por responsável usam sempre o **Status Atual**.

## Estrutura dos arquivos

```
data/
  GTO__LIST_OF_ITEMS.xlsx   Fonte primária (você atualiza este arquivo).
  status_history.json       Histórico de snapshots de status (base do "Anterior").
  enrichment.json           Análise EN/PT, responsável e categoria (curados à mão).
templates/
  dashboard_template.html   Molde do dashboard (não editar dados aqui).
build_j06.py                Script único de geração.
J06_Items_Analysis.xlsx     Saída — Excel.
J06_Dashboard.html          Saída — dashboard.
```

## Observações

- **Análise / Responsável** não existem no GTO; são mantidos em `enrichment.json`.
  Itens novos que aparecerem em um GTO futuro entram com análise em branco
  (`— pending —`) até serem preenchidos nesse arquivo.
- A data da última atualização é lida do campo *"Gerado em"* do próprio GTO e
  exibida no canto superior direito do dashboard e no cabeçalho do Excel.
