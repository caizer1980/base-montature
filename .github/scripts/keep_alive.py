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

Il titolo della pagina risulta sempre corretto (il documento HTML statico
arriva sempre), ma in alcuni run il form di login (componente React dentro
l'app Streamlit) non compare mai entro il timeout: probabile blocco/rallentamento
della connessione WebSocket da parte di sistemi anti-bot davanti a Streamlit
Cloud, che riconoscono le impronte tipiche di Chromium headless pilotato da
automazione (navigator.webdriver, User-Agent "HeadlessChrome", ecc.). Per
questo il browser viene lanciato con alcuni accorgimenti per sembrare un
normale browser reale, e in caso di fallimento vengono salvati screenshot e
log di rete/console utili per la diagnosi.

Eseguito periodicamente da .github/workflows/keep-alive.yml.
"""
import sys

from playwright.sync_api import sync_playwright

APP_URL = "https://base-montature-lzjoekyrfzydkhtxm2vpst.streamlit.app/"

REAL_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            args=[
                "--disable-blink-features=AutomationControlled",
            ]
        )
        context = browser.new_context(
            user_agent=REAL_USER_AGENT,
            locale="it-IT",
            viewport={"width": 1366, "height": 900},
        )
        # Rimuove l'impronta piu' comune usata dai sistemi anti-bot per
        # riconoscere Chromium pilotato da automazione.
        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
        )
        page = context.new_page()

        # Log di diagnosi: richieste di rete fallite e messaggi di errore in
        # console, utili se il form di login non dovesse comparire.
        page.on(
            "requestfailed",
            lambda req: print(
                f"[network] FALLITA: {req.method} {req.url} -> {req.failure}"
            ),
        )
        page.on(
            "console",
            lambda msg: print(f"[console:{msg.type}] {msg.text}")
            if msg.type in ("error", "warning")
            else None,
        )
        page.on(
            "response",
            lambda res: print(f"[response] {res.status} {res.url}")
            if res.status >= 400
            else None,
        )

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
            print(f"ERRORE: non ho trovato il form di login entro {SELECTOR_TIMEOUT_MS // 1000}s.")
            print(f"Titolo pagina: {title}")
            try:
                page.screenshot(path="failure.png", full_page=True)
                print("Screenshot salvato in failure.png (vedi artifact del workflow).")
            except Exception as e:
                print(f"Impossibile salvare lo screenshot: {e}")
            body_text = ""
            try:
                body_text = page.inner_text("body")
            except Exception:
                pass
            print("--- Testo visibile nel <body> al momento del timeout ---")
            print(body_text[:3000] if body_text else "(vuoto)")
            browser.close()
            sys.exit(1)

        title = page.title()
        browser.close()
        print(f"OK: app raggiunta e caricata correttamente. Titolo pagina: {title}")


if __name__ == "__main__":
    main()
