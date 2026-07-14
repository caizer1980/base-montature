#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
etl_montature.py
=================
Ricostruisce automaticamente il file "Base MONTATURE" a partire dai file
esportati quotidianamente (giacenza, vendite, listini, movimenti, sottoscorta),
replicando la logica di join che prima veniva fatta a mano in Access.

CHIAVE DI JOIN: "codice a barre" + "codice filiale" (es. "41031001" + "G" =
ID_GIACENZA "41031001G"), esattamente come nel file Access originale.

USO:
    python3 etl_montature.py --input-dir "/percorso/Export FOCUS DEPOSITO" \
                              --ref-dir "/percorso/tabelle_riferimento" \
                              --output "Base MONTATURE.xlsx"

I file sorgente (percorso Dropbox assoluto, vedi dizionario FILES qui sotto):
    /Export FOCUS DEPOSITO/BASE/Listini MONTATURE 2017.txt
    /Export FOCUS DEPOSITO/BASE/Sottoscorta MONTATURE.txt
    /Export FOCUS DEPOSITO/BASE/Movimenti ACQUISTO.txt
    /Export FOCUS DEPOSITO/RICERCHE BASE/GIACENZA MONTATURE al 2017.txt
    /Power BI/Dati/VENDUTO/Vendita TUTTI MAGAZZINI.xlsx   (vendite multi-anno,
                                                            solo righe "VEN")

Nota: dal 14/07/2026 i file vivono su un Dropbox diverso da quello originale
(un altro PC, aggiornato ogni notte) e le vendite arrivano da un unico export
Power BI con lo storico di piu' anni, non piu' da due file separati per
anno corrente/precedente.

--input-dir deve contenere, per ogni voce di FILES, un file con il nome
indicato in "local_name" (es. giacenza.txt, vendite.xlsx...). E' app.py che si
occupa di scaricare da Dropbox ogni file al suo "local_name" dentro una
cartella temporanea prima di chiamare build().

Le tabelle di riferimento "manuali" (colonne che in Access erano compilate a
mano e non provengono da nessuno dei file quotidiani) vengono lette da
--ref-dir e AGGIORNATE in automatico con i nuovi articoli trovati (lasciati
vuoti, pronti per essere compilati):
    ref_fornitore2.csv        fornitore -> Fornitore 2
    ref_marchi_attivi.csv     Marchio -> Marchi Attivi
    ref_articolo_manuale.csv  Codice a Barre -> Top, POSIZIONE GRIGLIA,
                               Occhiali con CLIP, Personale, scorta minima

NOTE / IPOTESI (confermate con Salvo il 14/07/2026):
  - Vengono escluse le combinazioni articolo+negozio con giacenza a ZERO e
    NESSUNA vendita negli ultimi due anni solari (anno corrente + anno
    precedente). Es. nel 2026: fuori chi ha giacenza 0 e zero vendite sia nel
    2026 sia nel 2025.
  - Dal file vendite si conta solo "Tipo operazione" = VEN (non ING).
  - "Quantita Disponibile" = quantita magazzino - quantita prenotata (min 0)
  - "Vendita {ANNO-1} CUMULATO" = vendite anno precedente dal 1/1 fino allo
    stesso giorno/mese di oggi (confronto omogeneo anno su anno)
  - "Vendita {ANNO-2}" ora e' calcolata con dati reali (il nuovo file vendite
    copre piu' anni), prima restava sempre vuota
  - "Filiale BIS" = uguale a "Filiale"
"""
import argparse
import csv
import datetime as dt
import os
import sys
from collections import defaultdict

# ---------------------------------------------------------------------------
# CONFIGURAZIONE
# ---------------------------------------------------------------------------

# Ogni file sorgente ha un percorso Dropbox assoluto (da settembre 2026 non
# sono piu' tutti nella stessa cartella: la giacenza/listini/sottoscorta/
# movimenti restano in "Export FOCUS DEPOSITO", le vendite arrivano invece da
# un unico file esportato da Power BI con lo storico di piu' anni).
FILES = {
    "giacenza": {
        "dropbox_path": "/Export FOCUS DEPOSITO/RICERCHE BASE/GIACENZA MONTATURE al 2017.txt",
        "local_name": "giacenza.txt",
        "format": "tsv",
    },
    "listini": {
        "dropbox_path": "/Export FOCUS DEPOSITO/BASE/Listini MONTATURE 2017.txt",
        "local_name": "listini.txt",
        "format": "tsv",
    },
    "sottoscorta": {
        "dropbox_path": "/Export FOCUS DEPOSITO/BASE/Sottoscorta MONTATURE.txt",
        "local_name": "sottoscorta.txt",
        "format": "tsv",
    },
    "movimenti": {
        "dropbox_path": "/Export FOCUS DEPOSITO/BASE/Movimenti ACQUISTO.txt",
        "local_name": "movimenti.txt",
        "format": "tsv",
    },
    "vendite": {
        "dropbox_path": "/Power BI/Dati/VENDUTO/Vendita TUTTI MAGAZZINI.xlsx",
        "local_name": "vendite.xlsx",
        "format": "xlsx",
    },
}

# Elenco filiali "vere" da tenere nel file finale (codice -> nome).
# Estratto dal file Access esistente. Se aprono un nuovo punto vendita,
# aggiungere qui la riga corrispondente.
FILIALI = {
    "A": ".Deposito",
    "A2": "Catania",
    "B": "Paterno",
    "C": "Acireale",
    "D": "Catania",
    "E": "Augusta",
    "G": "Giarre",
    "H": "Lentini",
    "I": "Portali",
    "IR": "Portali",
    "L": "Misterbianco",
    "R": "Paolo",
    "UR": "Paolo Bis",
    "V": "OUTLET",
    "Y": "Portali",
    "Z": "Viale Africa",
}

OUTPUT_COLUMNS = [
    "fornitore", "Fornitore 2", "Marchio", "modello", "Modello CARTIER", "SKU",
    "Modello + Colore + Calibro + CAT", "Modello + Colore + Calibro", "Colore + Calibro",
    "colore", "colore 2", "calibro", "ponte", "materiale", "utente", "Tipo Lenti",
    "Filiale", "Filiale BIS", "quantita magazzino", "Quantita Disponibile",
    "quantita in arrivo", "scorta minima",
    None, None, None, None,  # colonne vendite per anno, nomi dinamici (vedi build)
    "data ult acquisto", "data", "data ult vendita", "Categoria Filtro",
    "POSIZIONE GRIGLIA", "Top", "Marchi Attivi", "Occhiali con CLIP", "Personale",
    "Prezzo Di Acquisto Scheda Scontato", "Prezzo Di Vendita Scheda Scontato",
    "Prezzo Acquisto Listino Intero", "Prezzo Vendita Listino Intero",
    "Sconto Acquisto", "Sconto Vendita", "Fattore di RICARICO (su prezzi scontati)",
    "Codice a Barre", "ID_GIACENZA", "Filiale Focus",
]

# ---------------------------------------------------------------------------
# UTILITY
# ---------------------------------------------------------------------------


def parse_it_number(s):
    """Converte '143,65' -> 143.65. Stringa vuota -> None."""
    if s is None:
        return None
    s = s.strip()
    if s == "":
        return None
    s = s.replace(".", "").replace(",", ".") if "," in s else s
    try:
        return float(s)
    except ValueError:
        return None


def parse_it_date(s):
    """Converte 'gg/mm/aaaa' -> datetime.date. Stringa vuota -> None."""
    if not s:
        return None
    s = s.strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def load_tsv(path):
    """Legge un file .txt esportato (tab-separated, virgolette, encoding latin-1).
    Usato solo per i file piccoli (vendite). Per i file grandi (giacenza,
    listini, sottoscorta, movimenti) si usa iter_tsv per non saturare la RAM."""
    with open(path, encoding="latin-1", newline="") as f:
        reader = csv.reader(f, delimiter="\t", quotechar='"')
        header = next(reader)
        rows = [dict(zip(header, row)) for row in reader]
    return rows


def iter_tsv(path):
    """Ritorna (indice_colonne, reader, filehandle) senza caricare tutto il
    file in memoria. Ogni riga e' una lista posizionale."""
    f = open(path, encoding="latin-1", newline="")
    reader = csv.reader(f, delimiter="\t", quotechar='"')
    header = next(reader)
    idx = {name: i for i, name in enumerate(header)}
    return idx, reader, f


def load_ref_csv(path, key_field):
    d = {}
    if os.path.exists(path):
        with open(path, encoding="utf-8", newline="") as f:
            for row in csv.DictReader(f):
                d[row[key_field]] = row
    return d


def save_ref_csv(path, fieldnames, key_field, data):
    rows = sorted(data.values(), key=lambda r: str(r.get(key_field, "")))
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fieldnames})


# ---------------------------------------------------------------------------
# COSTRUZIONE DEL FILE
# ---------------------------------------------------------------------------


def get(row, idx, name):
    i = idx.get(name)
    return row[i] if i is not None and i < len(row) else ""


def build(input_dir, ref_dir, today=None):
    today = today or dt.date.today()
    year = today.year

    def p(key):
        return os.path.join(input_dir, FILES[key]["local_name"])

    print("Lettura GIACENZA (file guida)...", file=sys.stderr)
    idx_g, reader_g, fh_g = iter_tsv(p("giacenza"))
    giac_rows = []
    barcodes_rilevanti = set()
    for row in reader_g:
        fil_code = row[idx_g["filiale"]]
        if fil_code not in FILIALI:
            continue
        giac_rows.append(row)
        barcodes_rilevanti.add(row[idx_g["codice a barre"]])
    fh_g.close()
    print(f"  righe giacenza rilevanti: {len(giac_rows)}, barcode distinti: {len(barcodes_rilevanti)}", file=sys.stderr)

    print("Lettura LISTINI (solo barcode rilevanti)...", file=sys.stderr)
    listini_by_barcode = {}
    idx_l, reader_l, fh_l = iter_tsv(p("listini"))
    for row in reader_l:
        bc = row[idx_l["codice a barre"]]
        if bc in barcodes_rilevanti:
            listini_by_barcode[bc] = {
                "SKU": get(row, idx_l, "SKU"),
                "colore 2": get(row, idx_l, "colore 2"),
                "ponte": get(row, idx_l, "ponte"),
                "prezzo acquisto": get(row, idx_l, "prezzo acquisto"),
                "prezzo di vendita": get(row, idx_l, "prezzo di vendita"),
            }
    fh_l.close()

    print("Lettura SOTTOSCORTA (solo barcode rilevanti)...", file=sys.stderr)
    sottoscorta_by_barcode = {}
    idx_s, reader_s, fh_s = iter_tsv(p("sottoscorta"))
    for row in reader_s:
        bc = row[idx_s["codice a barre"]]
        if bc in barcodes_rilevanti:
            sottoscorta_by_barcode[bc] = {
                "SKU": get(row, idx_s, "SKU"),
                "colore 2": get(row, idx_s, "colore 2"),
                "ponte": get(row, idx_s, "ponte"),
                "prezzo acquisto": get(row, idx_s, "prezzo acquisto"),
                "prezzo di vendita": get(row, idx_s, "prezzo di vendita"),
            }
    fh_s.close()

    print("Lettura MOVIMENTI ACQUISTO (solo barcode rilevanti)...", file=sys.stderr)
    data_primo_acquisto = {}
    data_ultimo_acquisto = {}
    idx_m, reader_m, fh_m = iter_tsv(p("movimenti"))
    for row in reader_m:
        bc = get(row, idx_m, "codice a barre")
        if bc not in barcodes_rilevanti:
            continue
        if get(row, idx_m, "tipo operazione") != "ACQ":
            continue
        d = parse_it_date(get(row, idx_m, "data"))
        if not d:
            continue
        if bc not in data_primo_acquisto or d < data_primo_acquisto[bc]:
            data_primo_acquisto[bc] = d
        if bc not in data_ultimo_acquisto or d > data_ultimo_acquisto[bc]:
            data_ultimo_acquisto[bc] = d
    fh_m.close()

    print("Lettura VENDITE (file unico multi-anno, solo tipo operazione VEN)...", file=sys.stderr)
    import openpyxl as _openpyxl

    vendite_qta = defaultdict(int)
    vendite_qta_cumulato = defaultdict(int)
    ultima_vendita = {}

    anno_corrente = year
    anno_precedente = year - 1
    anno_meno2 = year - 2

    wb_v = _openpyxl.load_workbook(p("vendite"), read_only=True, data_only=True)
    ws_v = wb_v.active
    rows_v = ws_v.iter_rows(values_only=True)
    header_v = next(rows_v)
    idx_v = {name: i for i, name in enumerate(header_v)}

    def vget(row, name):
        i = idx_v.get(name)
        if i is None or i >= len(row):
            return None
        return row[i]

    n_vendite = 0
    for row in rows_v:
        if vget(row, "Tipo operazione") != "VEN":
            continue
        bc = vget(row, "Codice a barre")
        fil = vget(row, "Filiale")
        if not bc or not fil:
            continue
        d = vget(row, "Data")
        if d is None:
            continue
        d = d.date() if hasattr(d, "date") else d
        raw_qta = vget(row, "Quantità")
        try:
            qta = int(raw_qta or 0)
        except (TypeError, ValueError):
            qta = 0
        n_vendite += 1
        key = (bc, fil, d.year)
        vendite_qta[key] += qta
        cur = ultima_vendita.get((bc, fil))
        if cur is None or d > cur:
            ultima_vendita[(bc, fil)] = d
        if d.year == anno_precedente and (d.month, d.day) <= (today.month, today.day):
            vendite_qta_cumulato[(bc, fil)] += qta
    wb_v.close()
    print(f"  righe vendita (VEN) lette: {n_vendite}", file=sys.stderr)

    col_vendita_meno2 = f"Vendita {anno_meno2}"
    col_vendita_prec = f"Vendita {anno_precedente}"
    col_vendita_prec_cum = f"Vendita {anno_precedente} CUMULATO"
    col_vendita_corr = f"Vendita {anno_corrente}"

    ref_fornitore2 = load_ref_csv(os.path.join(ref_dir, "ref_fornitore2.csv"), "fornitore")
    ref_marchi_attivi = load_ref_csv(os.path.join(ref_dir, "ref_marchi_attivi.csv"), "Marchio")
    ref_articolo = load_ref_csv(os.path.join(ref_dir, "ref_articolo_manuale.csv"), "Codice a Barre")

    new_fornitori = [0]
    new_marchi = [0]
    new_articoli = [0]

    def get_fornitore2(fornitore):
        row = ref_fornitore2.get(fornitore)
        if row is None:
            ref_fornitore2[fornitore] = {"fornitore": fornitore, "Fornitore 2": ""}
            new_fornitori[0] += 1
            return ""
        return row.get("Fornitore 2", "")

    def get_marchi_attivi(marchio):
        row = ref_marchi_attivi.get(marchio)
        if row is None:
            ref_marchi_attivi[marchio] = {"Marchio": marchio, "Marchi Attivi": ""}
            new_marchi[0] += 1
            return ""
        return row.get("Marchi Attivi", "")

    def get_articolo_manuale(barcode, marchio, modello):
        row = ref_articolo.get(barcode)
        if row is None:
            row = {
                "Codice a Barre": barcode, "Marchio": marchio, "modello": modello,
                "Top": "", "POSIZIONE GRIGLIA": "", "Occhiali con CLIP": "",
                "Personale": "", "scorta minima": "",
            }
            ref_articolo[barcode] = row
            new_articoli[0] += 1
        return row

    print("Costruzione righe...", file=sys.stderr)
    out_rows = []
    for grow in giac_rows:
        fil_code = get(grow, idx_g, "filiale")
        barcode = get(grow, idx_g, "codice a barre")
        if not barcode:
            continue

        li = listini_by_barcode.get(barcode, {})
        so = sottoscorta_by_barcode.get(barcode, {})

        modello = get(grow, idx_g, "modello")
        colore = get(grow, idx_g, "colore")
        calibro = get(grow, idx_g, "calibro")
        categoria = get(grow, idx_g, "categoria filtro")
        marchio = get(grow, idx_g, "Marchio")
        fornitore = get(grow, idx_g, "fornitore")

        parts_cat = [x for x in (modello, colore, calibro, categoria) if x]
        parts_no_cat = [x for x in (modello, colore, calibro) if x]
        parts_col_cal = [x for x in (colore, calibro) if x]

        prezzo_acq_listino = parse_it_number(li.get("prezzo acquisto"))
        prezzo_ven_listino = parse_it_number(li.get("prezzo di vendita"))
        prezzo_acq_scheda = parse_it_number(so.get("prezzo acquisto"))
        prezzo_ven_scheda = parse_it_number(so.get("prezzo di vendita"))

        sconto_acq = None
        if prezzo_acq_listino:
            base = prezzo_acq_scheda if prezzo_acq_scheda is not None else prezzo_acq_listino
            sconto_acq = round(100 * (1 - base / prezzo_acq_listino), 6)
        sconto_ven = None
        if prezzo_ven_listino:
            base = prezzo_ven_scheda if prezzo_ven_scheda is not None else prezzo_ven_listino
            sconto_ven = round(100 * (1 - base / prezzo_ven_listino), 6)
        ricarico = None
        if prezzo_acq_scheda:
            ricarico = round((prezzo_ven_scheda or 0) / prezzo_acq_scheda, 6)

        def to_int(s):
            try:
                return int(float((s or "0").replace(",", ".")))
            except ValueError:
                return 0

        q_mag = to_int(get(grow, idx_g, "quantita magazzino"))
        q_prenot = to_int(get(grow, idx_g, "quantita prenotata"))
        q_arrivo = to_int(get(grow, idx_g, "quantita in arrivo"))
        q_disp = max(0, q_mag - q_prenot)

        v_corrente = vendite_qta.get((barcode, fil_code, anno_corrente)) or 0
        v_precedente = vendite_qta.get((barcode, fil_code, anno_precedente)) or 0

        # Filtro: escludi le combinazioni articolo+negozio con giacenza zero
        # E nessuna vendita negli ultimi due anni (anno corrente + anno precedente).
        if q_mag == 0 and v_corrente == 0 and v_precedente == 0:
            continue

        manual = get_articolo_manuale(barcode, marchio, modello)

        row = {
            "fornitore": fornitore,
            "Fornitore 2": get_fornitore2(fornitore),
            "Marchio": marchio,
            "modello": modello,
            "Modello CARTIER": modello if marchio == "CARTIER" else "",
            "SKU": li.get("SKU") or so.get("SKU") or "",
            "Modello + Colore + Calibro + CAT": "; ".join(parts_cat),
            "Modello + Colore + Calibro": "; ".join(parts_no_cat),
            "Colore + Calibro": "; ".join(parts_col_cal),
            "colore": colore,
            "colore 2": li.get("colore 2") or so.get("colore 2") or "",
            "calibro": calibro,
            "ponte": li.get("ponte") or so.get("ponte") or "",
            "materiale": get(grow, idx_g, "materiale"),
            "utente": get(grow, idx_g, "utente"),
            "Tipo Lenti": get(grow, idx_g, "Tipo Lenti"),
            "Filiale": FILIALI[fil_code],
            "Filiale BIS": FILIALI[fil_code],
            "quantita magazzino": q_mag,
            "Quantita Disponibile": q_disp,
            "quantita in arrivo": q_arrivo,
            "scorta minima": manual.get("scorta minima", ""),
            col_vendita_meno2: (vendite_qta.get((barcode, fil_code, anno_meno2)) or ""),
            col_vendita_prec: v_precedente or "",
            col_vendita_prec_cum: vendite_qta_cumulato.get((barcode, fil_code)) or "",
            col_vendita_corr: v_corrente or "",
            "data ult acquisto": data_ultimo_acquisto.get(barcode),
            "data": data_primo_acquisto.get(barcode),
            "data ult vendita": ultima_vendita.get((barcode, fil_code)),
            "Categoria Filtro": categoria,
            "POSIZIONE GRIGLIA": manual.get("POSIZIONE GRIGLIA", ""),
            "Top": manual.get("Top", ""),
            "Marchi Attivi": get_marchi_attivi(marchio),
            "Occhiali con CLIP": manual.get("Occhiali con CLIP", ""),
            "Personale": manual.get("Personale", ""),
            "Prezzo Di Acquisto Scheda Scontato": prezzo_acq_scheda,
            "Prezzo Di Vendita Scheda Scontato": prezzo_ven_scheda,
            "Prezzo Acquisto Listino Intero": prezzo_acq_listino,
            "Prezzo Vendita Listino Intero": prezzo_ven_listino,
            "Sconto Acquisto": sconto_acq,
            "Sconto Vendita": sconto_ven,
            "Fattore di RICARICO (su prezzi scontati)": ricarico,
            "Codice a Barre": barcode,
            "ID_GIACENZA": f"{barcode}{fil_code}",
            "Filiale Focus": fil_code,
        }
        out_rows.append(row)

    columns = [c for c in OUTPUT_COLUMNS if c is not None]
    idx = columns.index("scorta minima") + 1
    columns[idx:idx] = [col_vendita_meno2, col_vendita_prec, col_vendita_prec_cum, col_vendita_corr]

    os.makedirs(ref_dir, exist_ok=True)
    save_ref_csv(os.path.join(ref_dir, "ref_fornitore2.csv"), ["fornitore", "Fornitore 2"], "fornitore", ref_fornitore2)
    save_ref_csv(os.path.join(ref_dir, "ref_marchi_attivi.csv"), ["Marchio", "Marchi Attivi"], "Marchio", ref_marchi_attivi)
    save_ref_csv(
        os.path.join(ref_dir, "ref_articolo_manuale.csv"),
        ["Codice a Barre", "Marchio", "modello", "Top", "POSIZIONE GRIGLIA", "Occhiali con CLIP", "Personale", "scorta minima"],
        "Codice a Barre", ref_articolo,
    )

    print(
        f"Fatto. Righe: {len(out_rows)}. Nuovi fornitori: {new_fornitori[0]}, "
        f"nuovi marchi: {new_marchi[0]}, nuovi articoli: {new_articoli[0]} "
        f"(aggiunti vuoti nelle tabelle di riferimento, da compilare).",
        file=sys.stderr,
    )
    return columns, out_rows


def write_xlsx(columns, rows, output_path):
    import openpyxl
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Base")
    ws.append(columns)
    for row in rows:
        ws.append([row.get(c) for c in columns])
    wb.save(output_path)


def main():
    ap = argparse.ArgumentParser(description="Genera il file Base MONTATURE dai file sorgente.")
    ap.add_argument("--input-dir", required=True, help="Cartella con i file BASE/ e RICERCHE BASE/")
    ap.add_argument("--ref-dir", required=True, help="Cartella con le tabelle di riferimento (CSV)")
    ap.add_argument("--output", required=True, help="Percorso del file xlsx da generare")
    args = ap.parse_args()

    columns, rows = build(args.input_dir, args.ref_dir)
    write_xlsx(columns, rows, args.output)
    print(f"Salvato: {args.output}", file=sys.stderr)


if __name__ == "__main__":
    main()
