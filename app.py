#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py - Pagina web per generare e scaricare il file "Base MONTATURE"
=======================================================================
Applicazione Streamlit con login. Le persone autorizzate (elenco in
st.secrets) accedono con utente/password, premono un pulsante e ottengono
il file Excel aggiornato ad oggi, generato al volo dai file che si trovano
nella cartella Dropbox condivisa (letti tramite l'API di Dropbox, quindi
funziona anche se l'app gira su un server remoto e non sul tuo PC).

CONFIGURAZIONE RICHIESTA (vedi GUIDA_DEPLOY.md):
    In Streamlit Community Cloud -> Settings -> Secrets, incollare:

    DROPBOX_APP_KEY = "..."
    DROPBOX_APP_SECRET = "..."
    DROPBOX_REFRESH_TOKEN = "..."
    REF_TABLES_PATH = "/Export FOCUS DEPOSITO/Programma Base Montature/tabelle_riferimento"

    [utenti]
    salvo = "password-scelta-da-te"
    collega1 = "altra-password"

I percorsi dei 5 file sorgente (giacenza, listini, sottoscorta, movimenti,
vendite) sono definiti direttamente in etl_montature.py (dizionario FILES),
non nei Secrets: cambiano raramente e vivono nel codice per restare
allineati allo script che li legge.
"""
import io
import os
import tempfile
import datetime as dt

import streamlit as st

import etl_montature as etl

st.set_page_config(page_title="Base MONTATURE | Angiolucci Occhiali", page_icon="👓", layout="centered")

# ---------------------------------------------------------------------------
# STILE — allineato alla brand identity di angiolucciocchiali.com
# (colore primario del sito: #c06d49, logo ufficiale, tipografia pulita)
# ---------------------------------------------------------------------------

ANGIOLUCCI_LOGO = "https://www.angiolucciocchiali.com/cdn/shop/files/anglogo_2_1.png?v=1699797699&width=280"
BRAND_COLOR = "#c06d49"
BRAND_COLOR_DARK = "#9c5638"

st.markdown(
    f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Jost:wght@300;400;500;600&display=swap');

        html, body, [class*="css"] {{
            font-family: 'Jost', 'Helvetica Neue', Arial, sans-serif;
        }}

        #MainMenu, footer, header {{visibility: hidden;}}

        /* forza uno sfondo chiaro coerente, indipendentemente dal tema
           chiaro/scuro del browser di chi visita la pagina */
        [data-testid="stAppViewContainer"], [data-testid="stApp"], body {{
            background-color: #faf8f5 !important;
        }}
        [data-testid="stMain"], [data-testid="stMainBlockContainer"], .main {{
            background-color: #faf8f5 !important;
        }}

        .block-container {{
            padding-top: 2rem;
            max-width: 640px;
        }}

        .angiolucci-logo-wrap {{
            text-align: center;
            padding: 0.5rem 0 1.5rem 0;
        }}
        .angiolucci-logo-wrap img {{
            max-width: 220px;
        }}

        h1, h2, h3, .angiolucci-subtitle, p, span, label, div {{
            color: #262019;
        }}
        h1, h2, h3 {{
            font-weight: 500 !important;
            letter-spacing: 0.02em;
            color: #262019 !important;
        }}

        .angiolucci-subtitle {{
            text-align: center;
            color: #8a8378 !important;
            font-size: 0.95rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            margin-top: -0.8rem;
            margin-bottom: 1.5rem;
        }}

        /* etichette dei campi (Nome utente / Password) ben visibili */
        div[data-testid="stForm"] label p {{
            color: #262019 !important;
            font-weight: 600 !important;
            font-size: 0.95rem !important;
        }}

        div.stButton > button,
        div.stDownloadButton > button,
        div[data-testid="stButton"] button,
        div[data-testid="stDownloadButton"] button,
        div[data-testid="stFormSubmitButton"] button,
        button[kind="primary"],
        button[kind="primaryFormSubmit"],
        button[kind="secondaryFormSubmit"] {{
            background-color: {BRAND_COLOR} !important;
            color: #ffffff !important;
            border: none !important;
            border-radius: 4px !important;
            letter-spacing: 0.04em;
            text-transform: uppercase;
            font-size: 0.85rem;
            padding: 0.6rem 1.4rem !important;
            transition: background-color 0.2s ease-in-out;
        }}
        div.stButton > button:hover,
        div.stDownloadButton > button:hover,
        div[data-testid="stButton"] button:hover,
        div[data-testid="stDownloadButton"] button:hover,
        div[data-testid="stFormSubmitButton"] button:hover {{
            background-color: {BRAND_COLOR_DARK} !important;
            color: #ffffff !important;
        }}
        div.stButton > button p,
        div.stDownloadButton > button p,
        div[data-testid="stButton"] button p,
        div[data-testid="stDownloadButton"] button p,
        div[data-testid="stFormSubmitButton"] button p {{
            color: #ffffff !important;
        }}

        div[data-testid="stForm"] {{
            border: 1px solid #ece7df;
            border-radius: 8px;
            padding: 2rem 2rem 1.5rem 2rem;
            background-color: #ffffff;
        }}

        div[data-testid="stTextInput"] input {{
            background-color: #ffffff !important;
            color: #262019 !important;
            border: 1px solid #d8d2c6 !important;
            border-radius: 4px !important;
        }}

        /* icona mostra/nascondi password: discreta invece di quadrato nero */
        div[data-testid="stTextInput"] button,
        [data-testid="stTextInputRevealButton"],
        div[data-testid="stTextInput"] [data-testid*="RevealButton"] {{
            background-color: #ffffff !important;
            border: none !important;
            border-left: 1px solid #d8d2c6 !important;
            border-radius: 0 4px 4px 0 !important;
        }}
        div[data-testid="stTextInput"] button svg,
        [data-testid="stTextInputRevealButton"] svg {{
            fill: #8a8378 !important;
            color: #8a8378 !important;
        }}

        div[data-testid="stSidebar"] {{
            background-color: #fbfaf8 !important;
        }}
        div[data-testid="stSidebar"] * {{
            color: #262019 !important;
        }}

        div[data-testid="stAlert"] p {{
            color: inherit !important;
        }}
    </style>
    <div class="angiolucci-logo-wrap">
        <img src="{ANGIOLUCCI_LOGO}" alt="Angiolucci Occhiali" />
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# LOGIN SEMPLICE
# ---------------------------------------------------------------------------

def check_login():
    if st.session_state.get("logged_in"):
        return True

    st.markdown("<h2 style='text-align:center;'>Base MONTATURE</h2>", unsafe_allow_html=True)
    st.markdown("<p class='angiolucci-subtitle'>Area riservata &middot; accesso</p>", unsafe_allow_html=True)
    with st.form("login"):
        user = st.text_input("Utente")
        pwd = st.text_input("Password", type="password")
        ok = st.form_submit_button("Entra")

    if ok:
        utenti = st.secrets.get("utenti", {})
        if user in utenti and pwd == utenti[user]:
            st.session_state["logged_in"] = True
            st.session_state["utente"] = user
            st.rerun()
        else:
            st.error("Utente o password non corretti.")
    return False


if not check_login():
    st.stop()

st.sidebar.write(f"Connesso come **{st.session_state.get('utente')}**")
if st.sidebar.button("Esci"):
    st.session_state.clear()
    st.rerun()

# ---------------------------------------------------------------------------
# CONNESSIONE A DROPBOX
# ---------------------------------------------------------------------------

@st.cache_resource(show_spinner=False)
def get_dropbox_client():
    import dropbox
    return dropbox.Dropbox(
        app_key=st.secrets["DROPBOX_APP_KEY"],
        app_secret=st.secrets["DROPBOX_APP_SECRET"],
        oauth2_refresh_token=st.secrets["DROPBOX_REFRESH_TOKEN"],
    )


def dropbox_download(dbx, dropbox_path, local_path):
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    try:
        md, res = dbx.files_download(dropbox_path)
        with open(local_path, "wb") as f:
            f.write(res.content)
        return True, None
    except Exception as e:
        return False, str(e)


def dropbox_list_parent(dbx, dropbox_path):
    """Se un file non viene trovato, risale le cartelle finché non trova
    quella più vicina che esiste davvero e ne elenca il contenuto: così si
    capisce a che livello il percorso configurato si discosta da quello
    reale nell'account Dropbox collegato."""
    import posixpath
    path = dropbox_path.rstrip("/")
    last_err = None
    while True:
        parent = posixpath.dirname(path) or "/"
        try:
            res = dbx.files_list_folder(parent)
            names = [e.name for e in res.entries]
            return parent, names
        except Exception as e:
            last_err = e
            if parent == "/" or parent == path:
                return parent, [f"(impossibile leggere anche la cartella radice: {last_err})"]
            path = parent


def dropbox_upload(dbx, local_path, dropbox_path):
    import dropbox as dbx_mod
    with open(local_path, "rb") as f:
        dbx.files_upload(f.read(), dropbox_path, mode=dbx_mod.files.WriteMode.overwrite)


# ---------------------------------------------------------------------------
# GENERAZIONE FILE
# ---------------------------------------------------------------------------

st.markdown("<h2 style='text-align:center;'>Base MONTATURE</h2>", unsafe_allow_html=True)
st.markdown(
    "<p class='angiolucci-subtitle'>Giacenza &middot; vendite &middot; listini &middot; movimenti</p>",
    unsafe_allow_html=True,
)
st.write(
    "Genera il file aggiornato a oggi leggendo giacenza, vendite, listini e "
    "movimenti direttamente dalla cartella Dropbox."
)

if st.button("🔄 Genera file di oggi", type="primary"):
    with st.spinner("Scarico i file sorgente da Dropbox..."):
        dbx = get_dropbox_client()
        ref_subpath = st.secrets.get(
            "REF_TABLES_PATH",
            "/Export FOCUS DEPOSITO/Programma Base Montature/tabelle_riferimento",
        )

        with tempfile.TemporaryDirectory() as tmp:
            input_dir = os.path.join(tmp, "input")
            ref_dir = os.path.join(tmp, "ref")
            os.makedirs(ref_dir, exist_ok=True)

            missing = []
            for key, info in etl.FILES.items():
                ok, err = dropbox_download(dbx, info["dropbox_path"], os.path.join(input_dir, info["local_name"]))
                if not ok:
                    parent, names = dropbox_list_parent(dbx, info["dropbox_path"])
                    trovati = ", ".join(names) if names else "(cartella vuota o non trovata)"
                    missing.append(
                        f'**{key}**: `{info["dropbox_path"]}`\n\n'
                        f'Errore Dropbox: {err}\n\n'
                        f'Contenuto trovato in `{parent}`: {trovati}'
                    )
            if missing:
                st.error(
                    "Non riesco a scaricare questi file da Dropbox:\n\n"
                    + "\n\n---\n\n".join(missing)
                )
                st.stop()

            # tabelle di riferimento: se non esistono ancora su Dropbox, si
            # creano vuote al primo avvio (build() le popola comunque)
            for fname in ("ref_fornitore2.csv", "ref_marchi_attivi.csv", "ref_articolo_manuale.csv"):
                dropbox_download(dbx, f"{ref_subpath}/{fname}", os.path.join(ref_dir, fname))

            with st.spinner("Genero il file (può richiedere qualche minuto)..."):
                columns, rows = etl.build(input_dir, ref_dir)
                out_path = os.path.join(tmp, "Base_MONTATURE.xlsx")
                etl.write_xlsx(columns, rows, out_path)

            with st.spinner("Salvo le tabelle di riferimento aggiornate su Dropbox..."):
                for fname in ("ref_fornitore2.csv", "ref_marchi_attivi.csv", "ref_articolo_manuale.csv"):
                    local = os.path.join(ref_dir, fname)
                    if os.path.exists(local):
                        dropbox_upload(dbx, local, f"{ref_subpath}/{fname}")

            with open(out_path, "rb") as f:
                data = f.read()

            st.session_state["file_data"] = data
            st.session_state["file_rows"] = len(rows)
            st.session_state["file_time"] = dt.datetime.now().strftime("%d/%m/%Y %H:%M")

    st.success(f"File generato: {st.session_state['file_rows']} righe.")

if "file_data" in st.session_state:
    st.download_button(
        label=f"⬇️ Scarica Base MONTATURE ({st.session_state['file_time']})",
        data=st.session_state["file_data"],
        file_name=f"Base MONTATURE al {dt.date.today().strftime('%d %B %Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
