"""
sharepoint.py — Cliente REST para OneDrive y SharePoint sin browser.
Usa las cookies guardadas (FedAuth/rtFa) para todas las operaciones.
"""

import io
import json
import re
import unicodedata
import urllib.parse
from pathlib import Path
from typing import Optional
import os
import requests
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent.parent / ".env")

ONEDRIVE_BASE_URL = os.getenv("ONEDRIVE_BASE_URL")
ONEDRIVE_ROOT_FOLDER = os.getenv("ONEDRIVE_ROOT_FOLDER")
SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL")
SHAREPOINT_TARGET_LIBRARY = os.getenv("SHAREPOINT_TARGET_LIBRARY")
SESSION_FILE = Path(__file__).parent.parent / os.getenv("SESSION_DIR", "config/browser_session") / "state.json"

DB_SERVER_PATH = (
    "/teams/ETAPASPRODUCTIVAS2025SHP/DOCUMENTOS CERTIFICACIONES 2025"
    "/DOCUMENTOS INSTRUCTORES 2025"
    "/ADMINISTRACION Y CONTABILIDAD LUIS EDUARDO GONZALEZ"
    "/2026 LUIS EDUARDO GONZALEZ.xlsx"
)
DB_SHEET = "ASIGANCIONES 2026"
INSTRUCTOR_NAME = "LUIS EDUARDO GONZALEZ"

# Nombres posibles de la carpeta de bitácoras (fuzzy match)
BITACORA_FOLDER_KEYWORDS = ["bitacora", "bitácora"]


class SessionExpiredError(Exception):
    pass


def _check_auth(r: requests.Response) -> None:
    """Lanza SessionExpiredError si la respuesta indica sesión expirada."""
    if r.status_code in (401, 403) or (
        r.status_code == 200 and "html" in r.headers.get("Content-Type", "") and b"Sign in" in r.content[:500]
    ):
        raise SessionExpiredError("La sesión de SharePoint expiró. Vuelve a hacer login.")

# Caché del Excel de la base de datos (se descarga una vez por proceso)
_excel_db_cache: Optional["pd.DataFrame"] = None  # type: ignore[name-defined]


def get_session() -> requests.Session:
    """Crea un requests.Session con las cookies de SharePoint guardadas."""
    state = json.loads(SESSION_FILE.read_text())
    session = requests.Session()
    session.headers.update({"Accept": "application/json;odata=verbose"})
    for c in state.get("cookies", []):
        session.cookies.set(c["name"], c["value"], domain=c.get("domain", "").lstrip("."))
    return session


def _q(path: str) -> str:
    """Encode una ruta para el parámetro @p1 de SharePoint REST API."""
    return urllib.parse.quote(f"'{path}'")


def list_folders(session: requests.Session, base_url: str, path: str) -> list[dict]:
    r = session.get(f"{base_url}/_api/web/GetFolderByServerRelativeUrl(@p1)/Folders?@p1={_q(path)}")
    _check_auth(r)
    return r.json().get("d", {}).get("results", []) if r.ok else []


def list_files(session: requests.Session, base_url: str, path: str) -> list[dict]:
    r = session.get(f"{base_url}/_api/web/GetFolderByServerRelativeUrl(@p1)/Files?@p1={_q(path)}")
    _check_auth(r)
    return r.json().get("d", {}).get("results", []) if r.ok else []


def download_file(session: requests.Session, base_url: str, server_relative_path: str) -> Optional[bytes]:
    r = session.get(f"{base_url}/_api/web/GetFileByServerRelativeUrl(@p1)/$value?@p1={_q(server_relative_path)}")
    _check_auth(r)
    return r.content if r.ok else None


def convert_to_pdf(session: requests.Session, base_url: str, server_relative_path: str) -> Optional[bytes]:
    """
    Convierte un archivo Office a PDF usando la API de conversión de OneDrive/Microsoft 365.
    El archivo debe estar en OneDrive — no requiere LibreOffice ni Office local.

    Endpoint: /_api/v2.0/drive/root:{path-relativo-al-drive}:/content?format=pdf
    """
    # Extraer la ruta relativa al drive (quitar el prefijo /personal/{user})
    user_prefix = "/" + base_url.split(".com/")[1]  # /personal/lugonzalezs_sena_edu_co
    path_from_root = server_relative_path[len(user_prefix):]  # /Documents/...

    encoded = urllib.parse.quote(path_from_root)
    url = f"{base_url}/_api/v2.0/drive/root:{encoded}:/content?format=pdf"

    r = session.get(url, allow_redirects=True)
    if r.ok and r.content:
        print(f"  ✓ Convertido a PDF via OneDrive ({len(r.content):,} bytes)")
        return r.content
    print(f"  ✗ Conversión PDF falló: {r.status_code}")
    return None


def find_bitacora_folder(session: requests.Session, apprentice_path: str) -> Optional[str]:
    """
    Encuentra la carpeta de bitácoras dentro de la carpeta del aprendiz,
    tolerando variaciones de nombre: "1 - BITÁCORAS", "1. BITACORAS", "BITACORAS", etc.
    """
    folders = list_folders(session, ONEDRIVE_BASE_URL, apprentice_path)
    for f in folders:
        name = f.get("Name", "").lower()
        if any(kw in name for kw in BITACORA_FOLDER_KEYWORDS):
            return f.get("ServerRelativeUrl")
    return None


# ─────────────────────────────────────────────
# HELPERS DE NORMALIZACIÓN Y MATCHING
# ─────────────────────────────────────────────

def _normalize(text: str) -> str:
    """Elimina acentos y convierte a minúsculas para comparaciones."""
    return unicodedata.normalize("NFD", text).encode("ascii", "ignore").decode().lower()


def _name_tokens_set(normalized_text: str) -> set:
    """Extrae tokens significativos (>3 letras, no numéricos) de texto ya normalizado."""
    return {t for t in re.split(r'\W+', normalized_text) if len(t) > 3 and not t.isdigit()}


def _contains_bitacora(folder_name: str) -> bool:
    return "bitac" in _normalize(folder_name)


def _folder_starts_with_ficha(folder_name: str, ficha: str) -> bool:
    """El nombre de la carpeta empieza con el número de ficha (tolerando separadores)."""
    name = folder_name.strip()
    if not name.startswith(ficha):
        return False
    rest = name[len(ficha):]
    return not rest or rest[0] in (" ", "-", "_")


def _best_apprentice_match(folders: list, cc: str, name_tokens: set) -> Optional[dict]:
    """
    Encuentra la carpeta del aprendiz con mejor coincidencia por CC o similitud de nombre.
    name_tokens debe ser el set de tokens canónicos ya calculados (del Excel DB o fallback).
    Usa Jaccard + recall; excluye tokens numéricos para evitar falsos positivos con fichas.
    """
    best_folder = None
    best_score = 0.0

    for folder in folders:
        fn = folder.get("Name", "")
        # CC match = prioridad absoluta
        if cc and cc in fn:
            return folder
        if not name_tokens:
            continue
        folder_tokens = _name_tokens_set(_normalize(fn))
        if not folder_tokens:
            continue
        intersection = name_tokens & folder_tokens
        union = name_tokens | folder_tokens
        jaccard = len(intersection) / len(union)
        recall = len(intersection) / len(name_tokens)
        score = (jaccard + recall) / 2
        if score > best_score:
            best_score = score
            best_folder = folder

    if best_folder:
        print(f"  Mejor match carpeta: '{best_folder.get('Name')}' (score={best_score:.2f})")
    return best_folder if best_score >= 0.5 else None


# ─────────────────────────────────────────────
# BASE DE DATOS EXCEL Y LOOKUP
# ─────────────────────────────────────────────

def _load_excel_db(session: requests.Session) -> "pd.DataFrame":
    """Descarga y cachea el Excel de aprendices. Retorna DataFrame filtrado por instructor."""
    global _excel_db_cache
    if _excel_db_cache is not None:
        return _excel_db_cache

    import pandas as pd
    content = download_file(session, SHAREPOINT_SITE_URL, DB_SERVER_PATH)
    if not content:
        raise RuntimeError("No se pudo descargar la base de datos desde SharePoint.")

    df = pd.read_excel(io.BytesIO(content), sheet_name=DB_SHEET)
    df.columns = [c.replace("\n", " ").strip() for c in df.columns]

    col_instructor = "INSTRUCTOR DE SEGUIMIENTO Escribir letra mayuscula"
    df = df[df[col_instructor].str.upper().str.strip() == INSTRUCTOR_NAME].copy()
    df = df[df["No"].notna()].copy()

    _excel_db_cache = df
    print(f"  DB Excel cargada: {len(df)} aprendices")
    return df


def _lookup_apprentice(session: requests.Session, cc: str, apprentice_name: str) -> dict:
    """
    Busca el aprendiz en el Excel de la DB y retorna un dict con:
      - 'reds': lista de REDs donde buscar (determinada por AREA)
      - 'name_tokens': set de tokens del nombre canónico (APELLIDOS + NOMBRE)
      - 'cc': CC limpio
    Si no se encuentra en el Excel, usa los datos pasados como fallback.
    """
    cc_clean = (cc or "").strip()

    try:
        df = _load_excel_db(session)
        row = None

        # Buscar por CC primero (más confiable)
        if cc_clean:
            mask = df["CC-TI"].astype(str).str.strip() == cc_clean
            if mask.any():
                row = df[mask].iloc[0]

        # Fallback: buscar por similitud de nombre (Jaccard) cuando no hay CC
        if row is None and apprentice_name:
            query_tokens = _name_tokens_set(_normalize(apprentice_name))
            best_score = 0.0
            for _, r in df.iterrows():
                ap = _normalize(str(r.get("APELLIDOS Escribir letra mayuscula", "")))
                nm = _normalize(str(r.get("NOMBRE Escribir letra mayuscula", "")))
                db_tokens = _name_tokens_set(f"{ap} {nm}")
                if not db_tokens or not query_tokens:
                    continue
                inter = query_tokens & db_tokens
                union = query_tokens | db_tokens
                jaccard = len(inter) / len(union)
                recall = len(inter) / len(query_tokens)
                score = (jaccard + recall) / 2
                if score > best_score and score >= 0.5:
                    best_score = score
                    row = r

        if row is not None:
            area = str(row.get("AREA Escribir letra mayuscula", "")).upper()
            apellidos = str(row.get("APELLIDOS Escribir letra mayuscula", ""))
            nombre = str(row.get("NOMBRE Escribir letra mayuscula", ""))
            canonical = f"{apellidos} {nombre}"
            cc_from_db = str(row.get("CC-TI", cc_clean)).strip()

            if "CONTAB" in area:
                reds = ["2. RED CONTABLE"]
            else:
                reds = ["1. RED ADMINISTRACION"]

            name_tokens = _name_tokens_set(_normalize(canonical))
            canonical_str = canonical.strip().upper()
            print(f"  DB lookup: {canonical_str} | área={area} | CC={cc_from_db}")
            return {
                "reds": reds,
                "name_tokens": name_tokens,
                "canonical_name": canonical_str,
                "cc": cc_from_db or cc_clean,
            }

    except Exception as e:
        print(f"  ⚠ DB lookup falló ({e}), usando datos del registro local")

    # Fallback: usar nombre del registro SQLite, RED a determinar por el caller
    return {
        "reds": None,
        "name_tokens": _name_tokens_set(_normalize(apprentice_name)),
        "canonical_name": (apprentice_name or "").strip().upper(),
        "cc": cc_clean,
    }


def _create_subfolder(session: requests.Session, base_url: str,
                      parent_server_relative: str, folder_name: str) -> Optional[str]:
    """
    Crea una carpeta dentro de parent_server_relative.
    Retorna la ruta server-relative de la nueva carpeta, o None si falla.
    """
    r_digest = session.post(
        f"{base_url}/_api/contextinfo",
        headers={"Accept": "application/json;odata=verbose", "Content-Length": "0"}
    )
    if not r_digest.ok:
        print(f"  ✗ contextinfo falló al crear carpeta: {r_digest.status_code}")
        return None
    digest = r_digest.json()["d"]["GetContextWebInformation"]["FormDigestValue"]

    encoded_name = urllib.parse.quote(folder_name)
    url = (f"{base_url}/_api/web/GetFolderByServerRelativeUrl(@p1)"
           f"/Folders/add(url='{encoded_name}')?@p1={_q(parent_server_relative)}")
    r = session.post(url, headers={
        "Accept": "application/json;odata=verbose",
        "X-RequestDigest": digest,
        "Content-Length": "0",
    })
    if r.ok:
        new_path = f"{parent_server_relative}/{folder_name}"
        print(f"  ✓ Carpeta creada: {new_path}")
        return new_path
    print(f"  ✗ No se pudo crear carpeta '{folder_name}': {r.status_code} {r.text[:120]}")
    return None


def find_sharepoint_dest(session: requests.Session,
                         ficha: str, cc: str, apprentice_name: str, area: str) -> Optional[str]:
    """
    Encuentra (o crea) la carpeta de bitácoras del aprendiz en SharePoint.

    Estructura:
      {SHAREPOINT_TARGET_LIBRARY}/
        {1. RED ADMINISTRACION | 2. RED CONTABLE}/
          {ficha folder}/
            {apprentice folder}/        ← se crea si no existe
              BITÁCORAS/                ← se crea si no existe

    Retorna la ruta server-relative de la carpeta de bitácoras.
    """
    # 1. Consultar el Excel para obtener RED, nombre canónico y CC
    lookup = _lookup_apprentice(session, cc, apprentice_name)
    cc_clean = lookup["cc"]
    name_tokens = lookup["name_tokens"]
    canonical_name = lookup["canonical_name"]  # "APELLIDOS NOMBRE" en mayúsculas

    if lookup["reds"]:
        reds = lookup["reds"]
    elif "CONTAB" in (area or "").upper():
        reds = ["2. RED CONTABLE"]
    elif area:
        reds = ["1. RED ADMINISTRACION"]
    else:
        reds = ["1. RED ADMINISTRACION", "2. RED CONTABLE"]

    # 2. Buscar la(s) carpeta(s) de ficha en las REDs
    ficha_folders = []
    for red in reds:
        red_path = f"/teams/ETAPASPRODUCTIVAS2025SHP/{SHAREPOINT_TARGET_LIBRARY}/{red}"
        ficha_folders += [
            f for f in list_folders(session, SHAREPOINT_SITE_URL, red_path)
            if _folder_starts_with_ficha(f.get("Name", ""), ficha)
        ]
    if not ficha_folders:
        print(f"  ✗ No se encontró carpeta de ficha {ficha} en {reds}")
        return None

    # Usar la primera carpeta de ficha encontrada
    ficha_path = ficha_folders[0].get("ServerRelativeUrl", "")

    # 3. Buscar carpeta del aprendiz
    apprentice_folders = list_folders(session, SHAREPOINT_SITE_URL, ficha_path)
    apf = _best_apprentice_match(apprentice_folders, cc_clean, name_tokens)

    if not apf:
        # Crear carpeta del aprendiz: "{ficha} - {APELLIDOS NOMBRE}"
        new_folder_name = f"{ficha} - {canonical_name}"
        apf_path = _create_subfolder(session, SHAREPOINT_SITE_URL, ficha_path, new_folder_name)
        if not apf_path:
            return None
    else:
        apf_path = apf.get("ServerRelativeUrl", "")

    # 4. Buscar subcarpeta de bitácoras
    subfolders = list_folders(session, SHAREPOINT_SITE_URL, apf_path)
    for sf in subfolders:
        if _contains_bitacora(sf.get("Name", "")):
            return sf.get("ServerRelativeUrl")

    # Crear subcarpeta BITÁCORAS
    return _create_subfolder(session, SHAREPOINT_SITE_URL, apf_path, "BITÁCORAS")


def get_onedrive_user_path() -> str:
    user = ONEDRIVE_BASE_URL.split("/personal/")[1]
    return f"/personal/{user}{ONEDRIVE_ROOT_FOLDER}"


def upload_file(session: requests.Session, base_url: str, folder_server_relative: str,
                filename: str, content: bytes) -> bool:
    """
    Sube un archivo a una carpeta de SharePoint/OneDrive.
    folder_server_relative debe ser la ruta completa desde la raíz del sitio,
    ej: /personal/lugonzalezs_sena_edu_co/Documents/Instructor Luis Eduardo
    """
    r_digest = session.post(
        f"{base_url}/_api/contextinfo",
        headers={"Accept": "application/json;odata=verbose", "Content-Length": "0"}
    )
    if not r_digest.ok:
        print(f"  contextinfo failed: {r_digest.status_code}")
        return False
    digest = r_digest.json()["d"]["GetContextWebInformation"]["FormDigestValue"]

    encoded_name = urllib.parse.quote(filename)
    url = (f"{base_url}/_api/web/GetFolderByServerRelativeUrl(@p1)"
           f"/Files/add(overwrite=true,url='{encoded_name}')?@p1={_q(folder_server_relative)}")

    r = session.post(url, data=content, headers={
        "Accept": "application/json;odata=verbose",
        "X-RequestDigest": digest,
        "Content-Type": "application/octet-stream",
    })
    if not r.ok:
        print(f"  upload failed {r.status_code}: {r.text[:200]}")
    return r.ok
