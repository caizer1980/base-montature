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

Diagnosi (run precedenti): il titolo della pagina e il dominio finale
risultano sempre corretti (la richiesta arriva davvero al container
dell'app, oltre il cancello di autenticazione: lo conferma anche una
chiamata all'endpoint /api/v1/app/event/open dell'app stessa, vista nei
log di rete). Questo e' cio' che conta per tenere sveglia l'app: e'
l'evento che resetta il timer di inattivita' di Streamlit Cloud.

In alcuni ambienti (in particolare le shared runner IP di GitHub Actions)
l'interfaccia React che si connette via WebSocket non fa pero' in tempo a
comparire entro il timeout, probabilmente per un rallentamento o blocco
della connessione WebSocket specifico per quell'intervallo di IP. Per
questo il "successo" del ping NON richiede piu' di vedere il form di login
completo: basta che la richiesta abbia raggiunto davvero il dominio
dell'app (oltre il cancello di autenticazione). Il caricamento completo
dell'interfaccia viene comunque tentato come verifica bonus, ma il suo
mancato completamento non fa fallire il job.

Eseguito periodicamente da .github/workflows/keep-alive.yml.
"""
import sys

from playwright.sync_api import sync_playwright

APP_URL = "https://base-montature-lzjoekyrfzydkhtxm2vpst.streamlit.app/"
APP_DOMAIN = "streamlit.app"

REAL_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
)


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(
            args=["--disable-blink-features=AutomationControlled"]
        )
        context = browser.new_context(
            user_agent=REAL_USER_AGENT,
            locale="it-IT",
            viewport={"width": 1366, "height": 900},
        )
        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});"
        )
        page = context.new_page()

        page.on(
            "requestfailed",
            lambda req: print(f"[network] FALLITA: {req.method} {req.url} -> {req.failure}"),
        )
        page.on(
            "console",
            lambda msg: print(f"[console:{msg.type}] {msg.text}")
            if msg.type in ("error", "warning")
            else None,
        )
        page.on(
            "response",
            lambda res: print(f"[response] {res.status} {res.url}") if res.status >= 400 else None,
        )
        page.on("pageerror", lambda err: print(f"[pageerror] {err}"))

        # Il risveglio a freddo di un container Streamlit Cloud rimasto
        # addormentato per giorni puo' richiedere qualche minuto: timeout
        # generoso sul goto iniziale per non fallire inutilmente.
        page.set_default_timeout(120000)

        print(f"Apro {APP_URL} ...")
        page.goto(APP_URL, wait_until="domcontentloaded", timeout=120000)

        final_url = page.url
        title = page.title()
        print(f"URL finale dopo eventuali redirect: {final_url}")
        print(f"Titolo pagina: {title}")

        if APP_DOMAIN not in final_url:
            print("ERRORE: sono rimasto bloccato al cancello di autenticazione di Streamlit (non ho raggiunto l'app vera).")
            try:
                page.screenshot(path="failure.png", full_page=True)
                print("Screenshot salvato in failure.png (vedi artifact del workflow).")
            except Exception as e:
                print(f"Impossibile salvare lo screenshot: {e}")
            browser.close()
            sys.exit(1)

        # Obiettivo raggiunto: la richiesta e' arrivata al container reale
        # dell'app (oltre il cancello di autenticazione). Questo basta a
        # tenerla sveglia, indipendentemente dal fatto che l'interfaccia
        # completa faccia in tempo a caricarsi in questo ambiente.
        print("OK: richiesta arrivata al container dell'app (oltre il cancello di autenticazione). App tenuta sveglia.")

        # Verifica bonus, non bloccante: se il form di login (che richiede
        # una connessione WebSocket riuscita) compare in pochi secondi,
        # conferma che l'interfaccia e' pienamente funzionante da qui.
        try:
            page.wait_for_selector("text=Codice utente", timeout=25000)
            print("Bonus: form di login completamente caricato (connessione WebSocket riuscita).")
        except Exception:
            print(
                "Nota: il form di login non e' comparso entro 25s (probabile rallentamento/blocco "
                "della connessione WebSocket da questo ambiente), ma la richiesta HTTP ha comunque "
                "raggiunto l'app e questo e' sufficiente per tenerla sveglia."
            )
            try:
                page.screenshot(path="failure.png", full_page=True)
                print("Screenshot di diagnosi salvato in failure.png (vedi artifact del workflow).")
            except Exception as e:
                print(f"Impossibile salvare lo screenshot: {e}")

        browser.close()


if __name__ == "__main__":
    main()
