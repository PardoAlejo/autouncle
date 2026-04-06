"""
explore.py — Explora OneDrive y SharePoint usando SharePoint REST API + requests.
No necesita browser. Usa las cookies guardadas en config/browser_session/state.json.
"""

import json
import io
import os
import urllib.parse
from pathlib import Path
from dotenv import load_dotenv
import requests
import pandas as pd

load_dotenv(Path(__file__).parent.parent / ".env")

ONEDRIVE_BASE_URL = os.getenv("ONEDRIVE_BASE_URL")
ONEDRIVE_ROOT_FOLDER = os.getenv("ONEDRIVE_ROOT_FOLDER")
SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL")
SESSION_FILE = Path(__file__).parent.parent / os.getenv("SESSION_DIR", "config/browser_session") / "state.json"


def get_session() -> requests.Session:
    """Crea un requests.Session con las cookies guardadas de SharePoint."""
    state = json.loads(SESSION_FILE.read_text())
    session = requests.Session()
    session.headers.update({"Accept": "application/json;odata=verbose"})
    for c in state.get("cookies", []):
        session.cookies.set(c["name"], c["value"], domain=c.get("domain", "").lstrip("."))
    return session


def _sp_url(base_url: str, endpoint: str, path: str) -> str:
    """
    Construye una URL de SharePoint REST API usando el parámetro @p1
    para evitar problemas de escape con rutas que tienen espacios o comillas.
    Ejemplo: /_api/web/GetFolderByServerRelativeUrl(@p1)/Folders?@p1='/mi/ruta'
    """
    encoded_path = urllib.parse.quote(f"'{path}'")
    return f"{base_url}/_api/web/{endpoint}(@p1)/{{}}&@p1={encoded_path}"


def list_folder(session: requests.Session, base_url: str, server_relative_path: str):
    """Lista carpetas y archivos en una ruta. Retorna (folders, files)."""
    p = urllib.parse.quote(f"'{server_relative_path}'")

    r_folders = session.get(f"{base_url}/_api/web/GetFolderByServerRelativeUrl(@p1)/Folders?@p1={p}")
    r_files   = session.get(f"{base_url}/_api/web/GetFolderByServerRelativeUrl(@p1)/Files?@p1={p}")

    folders = r_folders.json().get("d", {}).get("results", []) if r_folders.ok else []
    files   = r_files.json().get("d", {}).get("results", [])   if r_files.ok   else []
    return folders, files


def download_file(session: requests.Session, base_url: str, server_relative_path: str) -> bytes | None:
    """Descarga un archivo desde SharePoint."""
    p = urllib.parse.quote(f"'{server_relative_path}'")
    r = session.get(f"{base_url}/_api/web/GetFileByServerRelativeUrl(@p1)/$value?@p1={p}")
    return r.content if r.ok else None


def explore_onedrive(session: requests.Session):
    print("\n" + "="*60)
    print("EXPLORANDO ONEDRIVE")
    print("="*60)

    user_path = ONEDRIVE_BASE_URL.split("/personal/")[1]
    root = f"/personal/{user_path}{ONEDRIVE_ROOT_FOLDER}"

    folders, files = list_folder(session, ONEDRIVE_BASE_URL, root)
    print(f"\n📁 Raíz ({len(folders)} carpetas, {len(files)} archivos):")

    for f in folders:
        name = f.get("Name", "?")
        path = f.get("ServerRelativeUrl", "")
        print(f"\n  📁 {name}")

        sub_folders, _ = list_folder(session, ONEDRIVE_BASE_URL, path)
        for sf in sub_folders:
            sf_path = sf.get("ServerRelativeUrl", "")
            print(f"      📁 {sf.get('Name')}")

            sub2_folders, _ = list_folder(session, ONEDRIVE_BASE_URL, sf_path)
            for sf2 in sub2_folders:
                sf2_path = sf2.get("ServerRelativeUrl", "")
                print(f"          📁 {sf2.get('Name')}")
                _, bitacoras = list_folder(session, ONEDRIVE_BASE_URL, sf2_path)
                for fi in bitacoras:
                    print(f"              📄 {fi.get('Name')} — {fi.get('TimeLastModified','')}")


def explore_sharepoint(session: requests.Session):
    print("\n" + "="*60)
    print("EXPLORANDO SHAREPOINT DESTINO")
    print("="*60)

    target = "/teams/ETAPASPRODUCTIVAS2025SHP/DOCUMENTOS CERTIFICACIONES 2025/DOCUMENTOS INSTRUCTORES 2025/ADMINISTRACION Y CONTABILIDAD LUIS EDUARDO GONZALEZ"
    folders, files = list_folder(session, SHAREPOINT_SITE_URL, target)

    print(f"\nSub-carpetas ({len(folders)}):")
    for f in folders:
        print(f"  📁 {f.get('Name')}")

    print(f"\nArchivos ({len(files)}):")
    for f in files[:20]:
        print(f"  📄 {f.get('Name')} — {f.get('TimeLastModified','')}")


def read_excel_db(session: requests.Session) -> pd.DataFrame | None:
    print("\n" + "="*60)
    print("LEYENDO BASE DE DATOS (Excel)")
    print("="*60)

    server_relative = "/teams/ETAPASPRODUCTIVAS2025SHP/DOCUMENTOS CERTIFICACIONES 2025/DOCUMENTOS INSTRUCTORES 2025/ADMINISTRACION Y CONTABILIDAD LUIS EDUARDO GONZALEZ/2026 LUIS EDUARDO GONZALEZ.xlsx"
    content = download_file(session, SHAREPOINT_SITE_URL, server_relative)

    if not content:
        print("No se pudo descargar el Excel.")
        return None

    print(f"Descargado: {len(content):,} bytes")
    df_all = pd.read_excel(io.BytesIO(content), sheet_name=None)
    print(f"Hojas: {list(df_all.keys())}")

    target = next((s for s in df_all if "ASIGNACIONES" in s.upper()), None)
    if target:
        df = df_all[target]
        print(f"\nHoja '{target}' — {len(df)} filas")
        print(f"Columnas: {list(df.columns)}")
        print(f"\nPrimeras 5 filas:")
        print(df.head().to_string())
        return df

    return None


if __name__ == "__main__":
    session = get_session()
    explore_onedrive(session)
    explore_sharepoint(session)
    df = read_excel_db(session)
    if df is not None:
        df.to_csv("data/asignaciones_2026.csv", index=False)
        print("\n✓ Guardado en data/asignaciones_2026.csv")
