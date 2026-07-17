#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
patch_template.py
==================
Aggiorna il tab "Base" di un file modello Excel (Aldo/Andrea: xlsx o xlsm,
con tabelle pivot gia' costruite) sostituendo SOLO i dati (righe 2 in poi)
con l'output dell'ETL (etl_montature.build()), lasciando invariati:
  - la riga 1 (intestazioni, stile, altezza)
  - tutti gli altri fogli (Modelli, Fornitore, RICIRCOLO, ecc.)
  - le tabelle pivot (xl/pivotTables/*)
  - il progetto VBA (per i file .xlsm)
  - stili, temi, sharedStrings

La cache delle pivot table (pivotCacheDefinition1.xml) viene aggiornata:
  - il riferimento all'intervallo dati (worksheetSource) punta alla nuova
    dimensione del tab Base
  - viene impostato refreshOnLoad="1" cosi' Excel ricalcola automaticamente
    le pivot (valori, elenchi filtro, ecc.) la prima volta che il file viene
    aperto, senza bisogno di ricalcolare nulla in Python.

Le nuove celle di testo vengono scritte come inlineStr (non tocchiamo
sharedStrings.xml, che resta quello originale e puo' restare "sporco":
Excel lo ignora per le celle inlineStr).
"""
import datetime as dt
import re
import zipfile
from xml.sax.saxutils import escape as xml_escape

EXCEL_EPOCH = dt.date(1899, 12, 30)


def _col_letter(n):
    """1 -> A, 2 -> B, ..., 27 -> AA, ..."""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _find_base_sheet_path(z):
    wb_xml = z.read("xl/workbook.xml").decode("utf-8")
    m = re.search(r'<sheet name="Base"[^>]*r:id="(rId\d+)"', wb_xml)
    if not m:
        raise ValueError("Foglio 'Base' non trovato in xl/workbook.xml")
    rid = m.group(1)
    rels_xml = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    m2 = re.search(r'<Relationship Id="%s"[^>]*Target="([^"]+)"' % re.escape(rid), rels_xml)
    if not m2:
        raise ValueError(f"Relationship {rid} non trovata in workbook.xml.rels")
    target = m2.group(1)
    return "xl/" + target if not target.startswith("xl/") else target


def _extract_row1(sheet_xml):
    m = re.search(r'<row r="1"[^>]*>.*?</row>', sheet_xml, re.S)
    if not m:
        raise ValueError("Riga 1 non trovata nel foglio Base")
    return m.group(0)


def _detect_column_styles(sheet_xml):
    """Legge la riga 2 esistente e ritorna, per ogni colonna (lettera),
    lo stile 's' eventualmente presente (o None)."""
    m = re.search(r'<row r="2"[^>]*>(.*?)</row>', sheet_xml, re.S)
    styles = {}
    if not m:
        return styles
    for cm in re.finditer(r'<c r="([A-Z]+)2"([^>]*?)/?>', m.group(1)):
        col_letters, attrs = cm.groups()
        sm = re.search(r's="(\d+)"', attrs)
        if sm:
            styles[col_letters] = sm.group(1)
    return styles


def _excel_serial(d):
    if d is None:
        return None
    if isinstance(d, dt.datetime):
        d = d.date()
    return (d - EXCEL_EPOCH).days


def build_sheet_data_xml(columns, rows, row1_xml, date_style, price_style, pct_style, ricarico_style):
    """Costruisce l'intero blocco <sheetData>...</sheetData> per il foglio
    Base: riga 1 invariata + una riga per ogni elemento di `rows`."""
    n_cols = len(columns)
    col_letters = [_col_letter(i + 1) for i in range(n_cols)]

    date_cols = {"data ult acquisto", "data", "data ult vendita"}
    price_cols = {
        "Prezzo Di Acquisto Scheda Scontato", "Prezzo Di Vendita Scheda Scontato",
        "Prezzo Acquisto Listino Intero", "Prezzo Vendita Listino Intero",
    }
    pct_cols = {"Sconto Acquisto", "Sconto Vendita"}
    ricarico_cols = {"Fattore di RICARICO (su prezzi scontati)"}

    style_by_col = {}
    for i, name in enumerate(columns):
        if name in date_cols and date_style:
            style_by_col[i] = date_style
        elif name in price_cols and price_style:
            style_by_col[i] = price_style
        elif name in pct_cols and pct_style:
            style_by_col[i] = pct_style
        elif name in ricarico_cols and ricarico_style:
            style_by_col[i] = ricarico_style

    parts = [row1_xml]
    for r_idx, row in enumerate(rows, start=2):
        cells = []
        for i, name in enumerate(columns):
            val = row.get(name)
            if val is None or val == "":
                continue
            ref = f"{col_letters[i]}{r_idx}"
            s_attr = f' s="{style_by_col[i]}"' if i in style_by_col else ""
            if name in date_cols:
                serial = _excel_serial(val)
                if serial is None:
                    continue
                cells.append(f'<c r="{ref}"{s_attr}><v>{serial}</v></c>')
            elif isinstance(val, (int, float)):
                cells.append(f'<c r="{ref}"{s_attr}><v>{val}</v></c>')
            else:
                text = xml_escape(str(val))
                cells.append(f'<c r="{ref}"{s_attr} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>')
        parts.append(f'<row r="{r_idx}" spans="1:{n_cols}" x14ac:dyDescent="0.25">' + "".join(cells) + "</row>")
    return "<sheetData>" + "".join(parts) + "</sheetData>"


def patch_workbook(template_path, columns, rows, output_path):
    n_cols = len(columns)
    last_col = _col_letter(n_cols)
    last_row = 1 + len(rows)

    zin = zipfile.ZipFile(template_path, "r")
    base_sheet_path = _find_base_sheet_path(zin)
    sheet_xml = zin.read(base_sheet_path).decode("utf-8")

    row1_xml = _extract_row1(sheet_xml)
    col_styles = _detect_column_styles(sheet_xml)

    # colonne storicamente "data" in entrambi i modelli: AA e/o AC.
    date_style = col_styles.get("AA") or col_styles.get("AC")
    price_style = col_styles.get("AJ")
    pct_style = col_styles.get("AN")
    ricarico_style = col_styles.get("AP")

    new_sheet_data = build_sheet_data_xml(
        columns, rows, row1_xml, date_style, price_style, pct_style, ricarico_style
    )

    new_sheet_xml = re.sub(r"<sheetData>.*?</sheetData>", lambda m: new_sheet_data, sheet_xml, count=1, flags=re.S)
    new_sheet_xml = re.sub(
        r'<dimension ref="[^"]*"/>',
        f'<dimension ref="A1:{last_col}{last_row}"/>',
        new_sheet_xml,
        count=1,
    )

    # ---- pivotCacheDefinition1.xml ----
    pivot_cache_path = "xl/pivotCache/pivotCacheDefinition1.xml"
    pivot_xml = zin.read(pivot_cache_path).decode("utf-8")
    pivot_xml = re.sub(
        r'<worksheetSource ref="[^"]*" sheet="Base"/>',
        f'<worksheetSource ref="A1:{last_col}{last_row}" sheet="Base"/>',
        pivot_xml,
        count=1,
    )
    if "refreshOnLoad=" in pivot_xml:
        pivot_xml = re.sub(r'refreshOnLoad="\d"', 'refreshOnLoad="1"', pivot_xml, count=1)
    else:
        pivot_xml = pivot_xml.replace("<pivotCacheDefinition ", '<pivotCacheDefinition refreshOnLoad="1" ', 1)
    pivot_xml = re.sub(r'recordCount="\d+"', f'recordCount="{len(rows)}"', pivot_xml, count=1)

    # ---- workbook.xml: aggiorna i defined name che puntano a Base!$A$1:$AS$N ----
    wb_xml = zin.read("xl/workbook.xml").decode("utf-8")
    wb_xml = re.sub(
        r"(Base!\$A\$1:\$%s\$)\d+" % re.escape(last_col),
        r"\g<1>" + str(last_row),
        wb_xml,
    )

    zout = zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == base_sheet_path:
            data = new_sheet_xml.encode("utf-8")
        elif item.filename == pivot_cache_path:
            data = pivot_xml.encode("utf-8")
        elif item.filename == "xl/workbook.xml":
            data = wb_xml.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    zin.close()
    return {
        "base_sheet_path": base_sheet_path,
        "last_row": last_row,
        "date_style": date_style,
        "price_style": price_style,
        "pct_style": pct_style,
        "ricarico_style": ricarico_style,
    }
