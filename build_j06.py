#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_j06.py — Pipeline de atualizacao do acompanhamento de itens J06 (SBR4).

O QUE FAZ
---------
Le o arquivo-fonte GTO (List of Items) que fica na pasta ./data/ e regenera:
  * J06_Items_Analysis.xlsx  (planilha com cabecalho congelado + filtros)
  * J06_Dashboard.html       (dashboard responsivo)
Ambos tem duas colunas de status:
  * Status Anterior  -> snapshot BASELINE fixo (classificacao antiga de referencia)
  * Status Atual     -> classificacao vigente no GTO mais recente

REGRAS ESPECIAIS
----------------
* Status Anterior e FIXO no snapshot marcado como "baseline" em status_history.json
  (nao deslocamos o "atual" para "anterior" a cada atualizacao). Para avancar a
  referencia, marque outro snapshot com "baseline": true (e remova a flag do atual).
* O status "8 - Awaiting full B05 completion" NAO existe -> e normalizado para
  "7 - Missing Vacuum Test or Sign" em toda parte (STATUS_REMAP).
* Cada atualizacao com mudancas e gravada como um snapshot no historico; a linha do
  tempo por item alimenta o fluxograma que abre ao clicar num item no dashboard.

COMO ATUALIZAR (fluxo do usuario)
---------------------------------
1. Coloque o Excel novo na pasta ./data/ (qualquer nome tipo "GTO ... .xlsx";
   o script pega o mais recente e o renomeia para GTO__LIST_OF_ITEMS.xlsx).
2. Rode:  python3 build_j06.py

FONTES DE DADOS (pasta ./data/)
-------------------------------
  GTO__LIST_OF_ITEMS.xlsx  Fonte primaria (status atual, descricao, certificado...).
  status_history.json      Historico de snapshots de status (baseline + atualizacoes).
  enrichment.json          Analise EN/PT, responsavel, categoria (curados a mao).
"""

import json
import re
import shutil
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
CANONICAL_SRC = DATA / "GTO__LIST_OF_ITEMS.xlsx"
HISTORY_JSON = DATA / "status_history.json"
ENRICH_JSON = DATA / "enrichment.json"
TEMPLATE_HTML = ROOT / "templates" / "dashboard_template.html"
OUT_XLSX = ROOT / "J06_Items_Analysis.xlsx"
OUT_HTML = ROOT / "J06_Dashboard.html"

# Status considerados "encerrados" (nao entram nas pendencias abertas).
CLOSED_SET = {"1 - Validated by ICN", "2 - Not Blocking"}

# Status 8 nao existe -> tratar como 7.
STATUS_REMAP = {"8 - Awaiting full B05 completion": "7 - Missing Vacuum Test or Sign"}


def norm_status(s):
    s = (s or "").strip()
    return STATUS_REMAP.get(s, s)


# Cor de preenchimento por status (Excel).
STATUS_FILL = {
    "1 - Validated by ICN": "C6EFCE",
    "2 - Not Blocking": "E2EFDA",
    "3 - Blocking": "FFC7CE",
    "4 - Under Analysis": "FCE4D6",
    "4 - Under Analysis To Not Blocking": "FCE4D6",
    "5 - Waiting Proof": "FFEB9C",
    "6 - Waiting B05": "DDEBF7",
    "7 - Missing Vacuum Test or Sign": "FFF2CC",
}

CLOSED_MARKERS = ("Status closed", "Status encerrado", "no pending action")


# --------------------------------------------------------------------------- #
# Fonte GTO
# --------------------------------------------------------------------------- #
def resolve_source() -> Path:
    """Pega o Excel GTO mais recente na pasta data/ (tolerante ao nome)."""
    candidates = sorted(DATA.glob("GTO*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not candidates:
        raise FileNotFoundError("Nenhum arquivo GTO*.xlsx encontrado em ./data/.")
    return candidates[0]


def consolidate_source(used: Path):
    """Renomeia o arquivo usado para o nome canonico e remove GTO*.xlsx extras."""
    if used.resolve() != CANONICAL_SRC.resolve():
        shutil.copy2(used, CANONICAL_SRC)
    for p in DATA.glob("GTO*.xlsx"):
        if p.resolve() != CANONICAL_SRC.resolve():
            p.unlink()


def read_gto(path: Path):
    """Le a planilha GTO. Retorna (registros_por_item, meta). Status normalizados."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    header = {ws.cell(2, c).value: c for c in range(1, ws.max_column + 1)}

    def col(name):
        if name not in header:
            raise KeyError(f"Coluna '{name}' nao encontrada na fonte GTO.")
        return header[name]

    c_item, c_upd = col("Item"), col("Updated Status")
    c_ojx, c_ajx = col("Original Jx"), col("Actual Jx")
    c_desc, c_cert, c_ctx = col("Description"), col("Certificate"), col("Context")

    items = {}
    for r in range(3, ws.max_row + 1):
        it = ws.cell(r, c_item).value
        if it is None:
            continue
        items[str(it)] = {
            "item": it,
            "context": ws.cell(r, c_ctx).value or "",
            "ojx": ws.cell(r, c_ojx).value or "",
            "ajx": ws.cell(r, c_ajx).value or "",
            "desc": ws.cell(r, c_desc).value or "",
            "cert": ws.cell(r, c_cert).value or "",
            "status": norm_status(ws.cell(r, c_upd).value),
        }

    meta = {"contract": "", "generated": ""}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(1, c).value
        if isinstance(v, str):
            if v.startswith("Contrato"):
                meta["contract"] = v.split(":", 1)[-1].strip()
            elif v.startswith("Gerado"):
                meta["generated"] = v.replace("Gerado em", "").replace("às", "").strip()
    return items, meta


def norm_generated(raw: str):
    """'27/7/2026 11:52' -> ('27/07/2026 11:52', '27/07 11:52')."""
    if not raw:
        d = date.today()
        return d.strftime("%d/%m/%Y"), d.strftime("%d/%m")
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})(?:\D+(\d{1,2}):(\d{2}))?", raw)
    if not m:
        return raw, raw
    d, mo, y = int(m.group(1)), int(m.group(2)), m.group(3)
    full = f"{d:02d}/{mo:02d}/{y}"
    short = f"{d:02d}/{mo:02d}"
    if m.group(4):
        full += f" {int(m.group(4)):02d}:{m.group(5)}"
        short += f" {int(m.group(4)):02d}:{m.group(5)}"
    return full, short


def _short_from_label(s):
    m = re.search(r"(\d{1,2})/(\d{1,2})/\d{4}\D+(\d{1,2}):(\d{2})", s.get("label", ""))
    if m:
        return f"{int(m.group(1)):02d}/{int(m.group(2)):02d} {int(m.group(3)):02d}:{m.group(4)}"
    return s.get("label", "snapshot")[:16]


# --------------------------------------------------------------------------- #
# Historico de status (baseline fixo + atualizacoes) e linha do tempo por item
# --------------------------------------------------------------------------- #
def sync_history(current: dict, source_label: str, short_label: str):
    """Normaliza o historico, garante baseline, anexa snapshot novo se mudou.
    Retorna (baseline_statuses, snapshots)."""
    h = json.loads(HISTORY_JSON.read_text(encoding="utf-8"))
    snaps = h["snapshots"]

    # Normaliza (8->7) e garante rotulo curto.
    for s in snaps:
        s["statuses"] = {k: norm_status(v) for k, v in s["statuses"].items()}
        if not s.get("short"):
            s["short"] = "Inicial" if s is snaps[0] else _short_from_label(s)

    # Garante um snapshot baseline (o primeiro, por padrao).
    if not any(s.get("baseline") for s in snaps):
        snaps[0]["baseline"] = True
        snaps[0]["short"] = "Inicial"

    # Anexa a atualizacao atual apenas se os status mudaram em relacao ao ultimo.
    if current != snaps[-1]["statuses"]:
        snaps.append({
            "label": source_label,
            "short": short_label,
            "date": date.today().isoformat(),
            "statuses": current,
        })

    baseline = next(s for s in snaps if s.get("baseline"))["statuses"]
    HISTORY_JSON.write_text(json.dumps(h, ensure_ascii=False, indent=1), encoding="utf-8")
    return baseline, snaps


def build_item_histories(snapshots):
    """Linha do tempo por item: [{t: rotulo, s: status}, ...] colapsando repeticoes."""
    hist = {}
    for s in snapshots:
        lbl, statuses = s["short"], s["statuses"]
        for item, stat in statuses.items():
            seq = hist.setdefault(item, [])
            if not seq or seq[-1]["s"] != stat:
                seq.append({"t": lbl, "s": stat})
    return hist


# --------------------------------------------------------------------------- #
# Montagem dos registros finais
# --------------------------------------------------------------------------- #
def build_records(gto, baseline, histories, enrich):
    records = []
    for key, g in gto.items():
        e = enrich.get(key, {})
        status = g["status"]
        closed = status in CLOSED_SET

        analysis = (e.get("analysis") or "").strip()
        analysis_pt = (e.get("analysis_pt") or "").strip()
        resp = (e.get("resp") or "").strip()
        category = e.get("category") or ("B05" if _looks_b05(g) else "General J06")
        note = e.get("note") or ""

        if not closed and any(mk in analysis for mk in CLOSED_MARKERS):
            analysis, analysis_pt = "", ""
        if closed and not analysis:
            analysis = "Status closed - no pending action."
            analysis_pt = "Status encerrado - sem acao pendente."
        if not closed and (not resp or resp.startswith("—") or resp.startswith("-")):
            resp = ""
        if closed:
            resp = resp or "— (Closed – no pending)"

        records.append({
            "item": g["item"],
            "context": g["context"],
            "ojx": g["ojx"],
            "ajx": g["ajx"],
            "category": category,
            "desc": g["desc"],
            "cert": g["cert"],
            "status_prev": baseline.get(key),   # Status Anterior (baseline fixo)
            "status": status,                    # Status Atual
            "analysis": analysis,
            "analysis_pt": analysis_pt,
            "resp": resp,
            "note": note,
            "closed": closed,
            "hist": histories.get(key, [{"t": "atual", "s": status}]),
        })
    records.sort(key=lambda d: int(d["item"]))
    return records


def _looks_b05(g):
    blob = f"{g.get('desc','')} {g.get('cert','')}".upper()
    return "B05" in blob


# --------------------------------------------------------------------------- #
# Saida Excel
# --------------------------------------------------------------------------- #
HEADERS = [
    "Item", "Context", "OriginalJx", "ActualJx", "Category", "Description",
    "Certificate / Evidence", "Status Anterior", "Status Atual",
    "Analysis (EN)", "Responsible", "Analysis (PT)",
]
COL_WIDTHS = [7, 10, 11, 10, 12, 44, 24, 20, 20, 40, 22, 34]
NCOL = len(HEADERS)


def write_excel(records, last_updated, source_label):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Merged Analysis"

    thin = Side(style="thin", color="D0D7E5")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    title_fill = PatternFill("solid", fgColor="1F3864")
    hdr_fill = PatternFill("solid", fgColor="2E5496")
    last_col = get_column_letter(NCOL)

    ws.merge_cells(f"A1:{last_col}1")
    t = ws["A1"]
    t.value = "J06 Items: Analysis, Status & Responsibility"
    t.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    t.fill = title_fill
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:G2")
    s = ws["A2"]
    n_open = sum(1 for r in records if not r["closed"])
    s.value = (f"SBR4 · {source_label} · {len(records)} itens (ActualJx = J06) · "
               f"{n_open} pendencias abertas")
    s.font = Font(name="Arial", size=10, italic=True, color="44546A")
    s.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(f"H2:{last_col}2")
    u = ws["H2"]
    u.value = f"Última atualização: {last_updated}"
    u.font = Font(name="Arial", size=10, bold=True, color="2E5496")
    u.alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[2].height = 18

    for j, h in enumerate(HEADERS, start=1):
        c = ws.cell(3, j, h)
        c.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border
    ws.row_dimensions[3].height = 30

    center = Alignment(horizontal="center", vertical="top", wrap_text=True)
    left = Alignment(horizontal="left", vertical="top", wrap_text=True)
    for i, d in enumerate(records):
        r = 4 + i
        changed = d["status_prev"] not in (None, "") and d["status_prev"] != d["status"]
        row = [
            d["item"], d["context"], d["ojx"], d["ajx"], d["category"], d["desc"],
            d["cert"],
            d["status_prev"] if d["status_prev"] else "— (novo item)",
            d["status"], d["analysis"], d["resp"], d["analysis_pt"],
        ]
        for j, val in enumerate(row, start=1):
            c = ws.cell(r, j, val)
            c.font = Font(name="Arial", size=10)
            c.border = border
            c.alignment = center if j in (1, 2, 3, 4, 5, 8, 9) else left
        if STATUS_FILL.get(d["status_prev"]):
            ws.cell(r, 8).fill = PatternFill("solid", fgColor=STATUS_FILL[d["status_prev"]])
        if STATUS_FILL.get(d["status"]):
            ws.cell(r, 9).fill = PatternFill("solid", fgColor=STATUS_FILL[d["status"]])
        if changed:
            ws.cell(r, 9).font = Font(name="Arial", size=10, bold=True)

    for j, w in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{last_col}{3 + len(records)}"

    _write_summary(wb, records)
    _write_changelog(wb, records, last_updated)
    _write_history_sheet(wb, records)
    wb.save(OUT_XLSX)


def _write_summary(wb, records):
    ws = wb.create_sheet("Summary by Responsible")
    fmt = lambda r: r if (r and r.strip() and not r.startswith("—")) else "(To be assigned)"
    groups = {}
    for d in records:
        if d["closed"]:
            continue
        groups.setdefault(fmt(d["resp"]), []).append(d)

    ws.merge_cells("A1:E1")
    ws["A1"] = "Open Items by Responsible (excludes closed: Validated / Not Blocking)"
    ws["A1"].font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F3864")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    for j, h in enumerate(["Responsible", "Total", "Blocking", "Waiting/Other", "Items"], 1):
        c = ws.cell(2, j, h)
        c.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2E5496")
        c.alignment = Alignment(horizontal="center", vertical="center")
    for i, g in enumerate(sorted(groups, key=lambda g: -len(groups[g]))):
        rows = groups[g]
        blk = sum(1 for d in rows if d["status"] == "3 - Blocking")
        items = ", ".join(str(d["item"]) for d in sorted(rows, key=lambda d: int(d["item"])))
        for j, val in enumerate([g, len(rows), blk, len(rows) - blk, items], 1):
            c = ws.cell(3 + i, j, val)
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(horizontal="left" if j in (1, 5) else "center",
                                    vertical="top", wrap_text=(j == 5))
    for col, w in zip("ABCDE", [26, 9, 11, 14, 60]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A3"


def _write_changelog(wb, records, last_updated):
    changed = [d for d in records
               if d["status_prev"] not in (None, "") and d["status_prev"] != d["status"]]
    ws = wb.create_sheet("Status Changes")
    ws.merge_cells("A1:D1")
    ws["A1"] = f"Mudanças de status (Anterior → Atual) — {len(changed)} item(ns)"
    ws["A1"].font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F3864")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    for j, h in enumerate(["Item", "Status Anterior", "Status Atual", "Description"], 1):
        c = ws.cell(2, j, h)
        c.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2E5496")
        c.alignment = Alignment(horizontal="center", vertical="center")
    for i, d in enumerate(sorted(changed, key=lambda d: int(d["item"]))):
        for j, val in enumerate([d["item"], d["status_prev"], d["status"], d["desc"]], 1):
            c = ws.cell(3 + i, j, val)
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(horizontal="center" if j in (1, 2, 3) else "left",
                                    vertical="top", wrap_text=(j == 4))
        if STATUS_FILL.get(d["status_prev"]):
            ws.cell(3 + i, 2).fill = PatternFill("solid", fgColor=STATUS_FILL[d["status_prev"]])
        if STATUS_FILL.get(d["status"]):
            ws.cell(3 + i, 3).fill = PatternFill("solid", fgColor=STATUS_FILL[d["status"]])
    for col, w in zip("ABCD", [8, 30, 30, 70]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A3"


def _write_history_sheet(wb, records):
    """Linha do tempo de status por item (mesma fonte do fluxograma do HTML)."""
    ws = wb.create_sheet("Item History")
    ws.merge_cells("A1:B1")
    ws["A1"] = "Histórico de status por item (Inicial → ... → Atual)"
    ws["A1"].font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F3864")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    for j, h in enumerate(["Item", "Linha do tempo de status"], 1):
        c = ws.cell(2, j, h)
        c.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2E5496")
        c.alignment = Alignment(horizontal="center", vertical="center")
    for i, d in enumerate(records):
        timeline = "  →  ".join(f"{n['s']} ({n['t']})" for n in d["hist"])
        ws.cell(3 + i, 1, d["item"]).font = Font(name="Arial", size=10, bold=True)
        ws.cell(3 + i, 1).alignment = Alignment(horizontal="center", vertical="top")
        c = ws.cell(3 + i, 2, timeline)
        c.font = Font(name="Arial", size=10)
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 110
    ws.freeze_panes = "A3"


# --------------------------------------------------------------------------- #
# Saida HTML
# --------------------------------------------------------------------------- #
def write_html(records, last_updated, source_label):
    tpl = TEMPLATE_HTML.read_text(encoding="utf-8")
    payload = [{
        "item": d["item"], "ojx": d["ojx"], "ajx": d["ajx"], "desc": d["desc"],
        "cert": d["cert"], "category": d["category"],
        "status": d["status"], "status_prev": d["status_prev"],
        "analysis": d["analysis"], "analysis_pt": d["analysis_pt"],
        "resp": d["resp"], "note": d["note"], "closed": d["closed"],
        "hist": d["hist"],
    } for d in records]
    html = (tpl
            .replace("__DATA_JSON__", json.dumps(payload, ensure_ascii=False))
            .replace("__LAST_UPDATED__", last_updated)
            .replace("__ITEM_COUNT__", str(len(records)))
            .replace("__SOURCE_LABEL__", source_label))
    OUT_HTML.write_text(html, encoding="utf-8")


# --------------------------------------------------------------------------- #
def main():
    src = resolve_source()
    gto, meta = read_gto(src)
    generated, gen_short = norm_generated(meta.get("generated", ""))
    last_updated = generated
    source_label = "GTO List of Items"
    if meta.get("generated"):
        source_label += f" · gerado em {generated}"

    current = {k: v["status"] for k, v in gto.items()}
    baseline, snapshots = sync_history(current, source_label, gen_short)
    histories = build_item_histories(snapshots)
    enrich = json.loads(ENRICH_JSON.read_text(encoding="utf-8"))

    records = build_records(gto, baseline, histories, enrich)
    write_excel(records, last_updated, source_label)
    write_html(records, last_updated, source_label)
    consolidate_source(src)

    changed = sum(1 for d in records
                  if d["status_prev"] not in (None, "") and d["status_prev"] != d["status"])
    print(f"OK · fonte: {src.name} · {len(records)} itens · {changed} alterado(s) "
          f"vs baseline · {len(snapshots)} snapshot(s) · atualizacao {last_updated}")
    print(f"  -> {OUT_XLSX.name}")
    print(f"  -> {OUT_HTML.name}")


if __name__ == "__main__":
    main()
