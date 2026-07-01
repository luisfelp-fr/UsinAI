# -*- coding: utf-8 -*-
"""Mantem os apps do Streamlit Community Cloud acordados.

Abre cada URL num navegador headless (Chromium via Playwright) e, se o app
estiver dormindo, clica no botao "Yes, get this app back up!" para acorda-lo.

As URLs podem ser sobrescritas por variaveis de ambiente:
  - APP_URLS: lista separada por virgulas (uma ou mais URLs);
  - APP_URL: uma unica URL (compatibilidade retroativa).
Se nenhuma for definida, usa a lista padrao abaixo.
"""
import os
import sys

from playwright.sync_api import sync_playwright

# Apps mantidos acordados por padrao.
DEFAULT_URLS = [
    "https://usinai-ops.streamlit.app/",
    "https://usinai-espelho.streamlit.app/",
]

# Rotulos possiveis do botao mostrado quando o app esta dormindo.
WAKE_LABELS = [
    "get this app back up",
    "Yes, get this app back up",
]


def _urls() -> list:
    raw = os.environ.get("APP_URLS") or os.environ.get("APP_URL")
    if raw:
        return [u.strip() for u in raw.split(",") if u.strip()]
    return DEFAULT_URLS


def _ping(page, url: str) -> None:
    print(f"Abrindo {url}")
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=60000)
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


def main() -> int:
    urls = _urls()
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        for url in urls:
            _ping(page, url)
        browser.close()

    print(f"OK - ping concluido para {len(urls)} app(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
