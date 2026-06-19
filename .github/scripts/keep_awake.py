# -*- coding: utf-8 -*-
"""Mantem o app do Streamlit Community Cloud acordado.

Abre a URL num navegador headless (Chromium via Playwright) e, se o app
estiver dormindo, clica no botao "Yes, get this app back up!" para acorda-lo.
A URL pode ser sobrescrita pela variavel de ambiente APP_URL.
"""
import os
import sys

from playwright.sync_api import sync_playwright

URL = os.environ.get("APP_URL", "https://usinai-ops.streamlit.app/")

# Rotulos possiveis do botao mostrado quando o app esta dormindo.
WAKE_LABELS = [
    "get this app back up",
    "Yes, get this app back up",
]


def main() -> int:
    print(f"Abrindo {URL}")
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        try:
            page.goto(URL, wait_until="domcontentloaded", timeout=60000)
        except Exception as exc:  # noqa: BLE001
            print(f"Aviso: goto demorou/falhou ({exc}); seguindo mesmo assim")

        page.wait_for_timeout(5000)

        clicked = False
        for label in WAKE_LABELS:
            btn = page.get_by_role("button", name=label, exact=False)
            if btn.count() > 0:
                print("App estava dormindo -> clicando para acordar")
                btn.first.click()
                clicked = True
                break

        if clicked:
            print("Aguardando o app reiniciar...")
            page.wait_for_timeout(45000)

        try:
            page.wait_for_load_state("networkidle", timeout=60000)
        except Exception:  # noqa: BLE001
            pass

        print("Titulo da pagina:", page.title())
        browser.close()

    print("OK - ping concluido")
    return 0


if __name__ == "__main__":
    sys.exit(main())
