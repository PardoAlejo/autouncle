"""
auth.py — Autenticación con Microsoft 365 via Playwright.

Primera vez: abre el navegador (visible) para login con MFA, luego visita
OneDrive y SharePoint para capturar todas las cookies necesarias.
Siguientes veces: headless (invisible), reutiliza sesión guardada (~90 días).
"""

import os
from pathlib import Path
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, BrowserContext

load_dotenv(Path(__file__).parent.parent / ".env")

MS_USERNAME = os.getenv("MS_USERNAME")
MS_PASSWORD = os.getenv("MS_PASSWORD")
SESSION_DIR = Path(__file__).parent.parent / os.getenv("SESSION_DIR", "config/browser_session")
SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL")
ONEDRIVE_BASE_URL = os.getenv("ONEDRIVE_BASE_URL")


def get_context(playwright, headless: bool = True) -> BrowserContext:
    """
    Devuelve un BrowserContext autenticado.
    - Si existe sesión válida: headless (invisible), sin pedir login.
    - Si no existe o expiró: abre navegador visible para MFA,
      visita OneDrive + SharePoint para capturar todas las cookies,
      luego guarda sesión.
    """
    SESSION_DIR.mkdir(parents=True, exist_ok=True)
    session_file = SESSION_DIR / "state.json"

    if session_file.exists():
        print("Reutilizando sesión guardada (headless)...")
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context(storage_state=str(session_file))
        if _session_is_valid(context):
            return context
        print("Sesión expirada, iniciando login...")
        context.close()
        browser.close()

    # Sin sesión válida: login interactivo visible
    print("Abriendo navegador para login con MFA...")
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    _do_login(context)
    _warm_up_cookies(context)
    context.storage_state(path=str(session_file))
    print(f"✓ Sesión guardada en {session_file}")
    return context


def _do_login(context: BrowserContext):
    """Login interactivo con MFA."""
    page = context.new_page()
    page.goto("https://login.microsoftonline.com")

    page.wait_for_selector("input[type='email']", timeout=15000)
    page.fill("input[type='email']", MS_USERNAME)
    page.click("input[type='submit']")

    page.wait_for_selector("input[type='password']", timeout=15000)
    page.fill("input[type='password']", MS_PASSWORD)
    page.click("input[type='submit']")

    print("Esperando aprobación MFA en el teléfono... (máximo 120 segundos)")
    page.wait_for_url(
        lambda url: "login.microsoftonline.com" not in url,
        timeout=120000
    )

    # Aceptar "¿Mantener sesión iniciada?"
    try:
        page.wait_for_selector("#idSIButton9", timeout=5000)
        page.click("#idSIButton9")  # "Sí"
    except Exception:
        pass

    print("✓ Login completado.")
    page.close()


def _warm_up_cookies(context: BrowserContext):
    """
    Visita OneDrive y SharePoint para que el navegador reciba las cookies
    FedAuth/rtFa necesarias para llamadas API sin browser.
    """
    print("Capturando cookies de OneDrive y SharePoint...")
    page = context.new_page()

    page.goto(ONEDRIVE_BASE_URL, wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(2000)

    page.goto(SHAREPOINT_SITE_URL, wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(2000)

    page.close()
    print("✓ Cookies capturadas.")


def _session_is_valid(context: BrowserContext) -> bool:
    """Verifica que la sesión sigue activa (headless)."""
    page = context.new_page()
    try:
        page.goto(SHAREPOINT_SITE_URL, wait_until="domcontentloaded", timeout=20000)
        return "login.microsoftonline.com" not in page.url
    except Exception:
        return False
    finally:
        page.close()


def get_requests_session(playwright_context: BrowserContext):
    """
    Devuelve un requests.Session con las cookies de SharePoint/OneDrive
    para hacer llamadas REST API sin browser.
    """
    import json
    import requests

    state = playwright_context.storage_state()
    session = requests.Session()
    session.headers.update({"Accept": "application/json;odata=verbose"})

    for cookie in state["cookies"]:
        session.cookies.set(
            cookie["name"],
            cookie["value"],
            domain=cookie.get("domain", "").lstrip("."),
        )
    return session


if __name__ == "__main__":
    """Ejecutar directamente para renovar/crear sesión."""
    with sync_playwright() as p:
        context = get_context(p, headless=False)
        print("✓ Autenticación exitosa.")
        context.browser.close()
