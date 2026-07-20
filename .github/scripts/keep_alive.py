#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
keep_alive.py
=============
Visita l'app "Base MONTATURE" su Streamlit Community Cloud con un browser
headless reale (Playwright), cosi' da tenerla sveglia. Un semplice ping HTTP
(curl, cron-job.org, UptimeRobot in modalita' base) NON funziona: Streamlit
mette prima di ogni app un "cancello" (share.streamlit.io/-/auth/app) che
imposta un cookie di sessione e fa un redirect verso l'app vera. Un client
HTTP che non segue i redirect e non gestisce i cookie resta bloccato la'
e non arriva mai a caricare davvero l'app.

Eseguito periodicamente da .github/workflows/keep-alive.yml.
"""
import sys

from playwright.sync_api import sync_playwright

APP_URL = "https://base-montature-lzjoekyrfzydkhtxm2vpst.streamlit.app/"


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        # Il risveglio a freddo di un'app Streamlit Cloud rimasta addormentata
        # per giorni puo' richiedere diversi minuti (ricreazione del container,
        # reinstallazione delle dipendenze, boot dell'app): timeout molto
        # generosi per non fallire inutilmente. Una volta che questo workflow
        # gira regolarmente ogni 4 ore l'app restera' quasi sempre "calda" e
        # i run successivi saranno molto piu' rapidi.
        SELECTOR_TIMEOUT_MS = 280000
        page.set_default_timeout(SELECTOR_TIMEOUT_MS)

        print(f"Apro {APP_URL} ...")
        page.goto(APP_URL, wait_until="domcontentloaded", timeout=120000)

        # Se l'app risultasse comunque addormentata, Streamlit mostra una
        # pagina con un pulsante per il risveglio manuale: lo clicchiamo
        # se dovesse comparire (di solito con l'app pubblica non serve).
        try:
            wake_button = page.get_by_text("get this app back up", exact=False)
            if wake_button.count() > 0:
                print("Trovato pulsante di risveglio manuale, clicco...")
                wake_button.first.click()
        except Exception:
            pass

        # Aspetta che compaia il form di login della nostra app (prova che
        # il caricamento e' arrivato fino in fondo, non fermato al cancello
        # di autenticazione di Streamlit). Timeout ampio per coprire un
        # risveglio a freddo completo del container.
        try:
            page.wait_for_selector("text=Codice utente", timeout=SELECTOR_TIMEOUT_MS)
        except Exception:
            title = page.title()
            content = page.content()
            browser.close()
            print(f"ERRORE: non ho trovato il form di login entro {SELECTOR_TIMEOUT_MS // 1000}s.")
            print(f"Titolo pagina: {title}")
            print(content[:2000])
            sys.exit(1)

        title = page.title()
        browser.close()
        print(f"OK: app raggiunta e caricata correttamente. Titolo pagina: {title}")


if __name__ == "__main__":
    main()
