"""
sync.py — Sincronización de bitácoras desde OneDrive vs base de datos.

Lógica:
1. Descarga la base de datos desde SharePoint (filtrada por Luis Eduardo)
2. Recorre las carpetas de OneDrive buscando bitácoras
3. Cruza contra el estado en la DB (Bitácora 1..6 ok/Pendiente)
4. Retorna lista de bitácoras pendientes de revisión con su metadata
"""

import io
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
import pandas as pd
import requests
from dotenv import load_dotenv

from sharepoint import (
    get_session, list_folders, list_files, download_file,
    find_bitacora_folder, get_onedrive_user_path,
    ONEDRIVE_BASE_URL, SHAREPOINT_SITE_URL, SHAREPOINT_TARGET_LIBRARY
)

load_dotenv(Path(__file__).parent.parent / ".env")

DB_EXCEL_URL = os.getenv("DB_EXCEL_URL")
INSTRUCTOR_NAME = "LUIS EDUARDO GONZALEZ"
DB_SERVER_PATH = "/teams/ETAPASPRODUCTIVAS2025SHP/DOCUMENTOS CERTIFICACIONES 2025/DOCUMENTOS INSTRUCTORES 2025/ADMINISTRACION Y CONTABILIDAD LUIS EDUARDO GONZALEZ/2026 LUIS EDUARDO GONZALEZ.xlsx"
DB_SHEET = "ASIGANCIONES 2026"

BITACORA_COLS = [f"Bitácora {i}\nok / Pendiente" for i in range(1, 7)]


@dataclass
class Bitacora:
    """Representa una bitácora encontrada en OneDrive."""
    # Datos del aprendiz
    apprentice_name: str
    cc: str
    ficha: str
    program: str
    area: str          # CONTABILIDAD o ADMINISTRACIÓN
    company: str
    start_date: Optional[str]
    end_date: Optional[str]

    # Datos del archivo
    filename: str
    file_path: str     # server-relative path en OneDrive
    file_ext: str      # xlsx, xls, pdf
    modified: str
    bitacora_num: int  # 1-6, inferido del nombre del archivo

    # Estado en la DB
    db_status: str     # "Pendiente", "ok", "" (no registrada)

    # Seguimientos programados
    seg1_date: Optional[str] = None
    seg2_date: Optional[str] = None
    seg3_date: Optional[str] = None


def load_db(session: requests.Session) -> pd.DataFrame:
    """Descarga y carga la base de datos desde SharePoint."""
    content = download_file(session, SHAREPOINT_SITE_URL, DB_SERVER_PATH)
    if not content:
        raise RuntimeError("No se pudo descargar la base de datos desde SharePoint.")

    df = pd.read_excel(io.BytesIO(content), sheet_name=DB_SHEET)

    # Normalizar nombres de columnas (quitar saltos de línea)
    df.columns = [c.replace("\n", " ").strip() for c in df.columns]

    # Filtrar solo aprendices de Luis Eduardo
    col_instructor = "INSTRUCTOR DE SEGUIMIENTO Escribir letra mayuscula"
    df = df[df[col_instructor].str.upper().str.strip() == INSTRUCTOR_NAME].copy()
    df = df[df["No"].notna()].copy()  # eliminar filas vacías

    print(f"Base de datos: {len(df)} aprendices de {INSTRUCTOR_NAME}")
    return df


def _infer_bitacora_num(filename: str) -> int:
    """Intenta inferir el número de bitácora del nombre del archivo."""
    import re
    # Buscar dígito después de "bitacora", "bitácora", "n°", "#"
    name = filename.lower()
    # Patrones: "bitacora 3", "bitacora3", "bitácora 3", "nro 3", "no.3", "n°3"
    patterns = [
        r"bit[aá]cora\s*[#n°\-]*\s*(\d)",
        r"[#n°]\s*(\d)",
        r"\b(\d)\b",  # fallback: primer dígito suelto
    ]
    for pattern in patterns:
        m = re.search(pattern, name)
        if m:
            num = int(m.group(1))
            if 1 <= num <= 6:
                return num
    return 0  # desconocido


def _normalize_name(name: str) -> str:
    """Normaliza nombre para comparación fuzzy."""
    import unicodedata
    name = unicodedata.normalize("NFD", name)
    name = "".join(c for c in name if unicodedata.category(c) != "Mn")
    return name.upper().strip()


def _find_apprentice_in_db(df: pd.DataFrame, folder_name: str) -> Optional[pd.Series]:
    """
    Busca un aprendiz en la DB por nombre fuzzy-matching contra el nombre de carpeta.
    Las carpetas pueden tener apellido primero, nombre primero, o con tildes.
    """
    folder_norm = _normalize_name(folder_name)

    for _, row in df.iterrows():
        apellido = _normalize_name(str(row.get("APELLIDOS Escribir letra mayuscula", "")))
        nombre = _normalize_name(str(row.get("NOMBRE Escribir letra mayuscula", "")))
        full1 = f"{apellido} {nombre}"
        full2 = f"{nombre} {apellido}"

        # Match exacto o parcial
        if (full1 in folder_norm or full2 in folder_norm or
                folder_norm in full1 or folder_norm in full2):
            return row

        # Match por partes (al menos apellido completo)
        if apellido and apellido in folder_norm:
            return row

    return None


def scan_onedrive(session: requests.Session, df: pd.DataFrame) -> list[Bitacora]:
    """
    Escanea el OneDrive y retorna lista de bitácoras encontradas
    cruzadas contra el estado en la base de datos.
    """
    root = get_onedrive_user_path()
    bitacoras: list[Bitacora] = []

    ficha_folders = list_folders(session, ONEDRIVE_BASE_URL, root)
    print(f"Escaneando {len(ficha_folders)} fichas...")

    for ficha_folder in ficha_folders:
        ficha_path = ficha_folder["ServerRelativeUrl"]
        ficha_name = ficha_folder["Name"]

        # Extraer número de ficha del nombre de la carpeta
        import re
        ficha_match = re.search(r"\d{6,}", ficha_name)
        ficha_num = ficha_match.group(0) if ficha_match else ""

        # Nivel 1: puede haber subcarpeta "APRENDIZ" o ir directo a nombres
        level1_folders = list_folders(session, ONEDRIVE_BASE_URL, ficha_path)

        # Detectar si hay subcarpeta intermedia tipo "APRENDIZ"
        apprentice_folders = []
        for f in level1_folders:
            name = f["Name"].upper()
            if name in ("APRENDIZ", "APRENDICES"):
                # Un nivel más adentro
                sub = list_folders(session, ONEDRIVE_BASE_URL, f["ServerRelativeUrl"])
                apprentice_folders.extend(sub)
            else:
                apprentice_folders.append(f)

        for ap_folder in apprentice_folders:
            ap_path = ap_folder["ServerRelativeUrl"]
            ap_name = ap_folder["Name"]

            # Buscar carpeta de bitácoras
            bitacora_folder = find_bitacora_folder(session, ap_path)
            if not bitacora_folder:
                continue

            # Listar archivos de bitácoras
            files = list_files(session, ONEDRIVE_BASE_URL, bitacora_folder)
            if not files:
                continue

            # Buscar aprendiz en la DB
            db_row = _find_apprentice_in_db(df, ap_name)

            for f in files:
                fname = f["Name"]
                ext = fname.rsplit(".", 1)[-1].lower() if "." in fname else ""
                if ext not in ("xlsx", "xls", "pdf"):
                    continue

                # Saltar archivos que no sean bitácoras (ej: certificados ARL en carpeta equivocada)
                if not any(kw in fname.lower() for kw in ["bitacora", "bitácora", "format"]):
                    continue

                num = _infer_bitacora_num(fname)

                # Estado en DB
                db_status = ""
                if db_row is not None and num > 0:
                    col = f"Bitácora {num} ok / Pendiente"
                    db_status = str(db_row.get(col, "")).strip()

                bitacoras.append(Bitacora(
                    apprentice_name=ap_name,
                    cc=str(db_row.get("CC-TI", "")) if db_row is not None else "",
                    ficha=ficha_num,
                    program=str(db_row.get("PROGRAMA DE FORMACIÓN Escribir letra mayuscula", ficha_name)) if db_row is not None else ficha_name,
                    area=str(db_row.get("AREA Escribir letra mayuscula", "")) if db_row is not None else "",
                    company=str(db_row.get("EMPRESA Escribir letra mayuscula", "")) if db_row is not None else "",
                    start_date=str(db_row.get("FECHA INICIO EP dia de mes de año", "")) if db_row is not None else "",
                    end_date=str(db_row.get("FECHA FIN  dia de mes de año", "")) if db_row is not None else "",
                    filename=fname,
                    file_path=f["ServerRelativeUrl"],
                    file_ext=ext,
                    modified=f.get("TimeLastModified", ""),
                    bitacora_num=num,
                    db_status=db_status,
                    seg1_date=str(db_row.get("1ER SEGUIMIENTO  dia de mes de año", "")) if db_row is not None else "",
                    seg2_date=str(db_row.get("2DO SEGUIMIENTO  dia de mes de año", "")) if db_row is not None else "",
                    seg3_date=str(db_row.get("3er SEGUIMIENTO  dia de mes de año", "")) if db_row is not None else "",
                ))

    return bitacoras


def get_pending_bitacoras(session: requests.Session) -> list[Bitacora]:
    """
    Retorna solo las bitácoras que NO están marcadas como 'ok' en la base de datos.
    Estas son las que tu tío necesita revisar.
    """
    df = load_db(session)
    all_bitacoras = scan_onedrive(session, df)

    pending = [b for b in all_bitacoras if b.db_status.lower() != "ok"]
    ok_count = len(all_bitacoras) - len(pending)

    print(f"\nTotal bitácoras encontradas: {len(all_bitacoras)}")
    print(f"  ✅ Ya revisadas (ok): {ok_count}")
    print(f"  🔴 Pendientes de revisión: {len(pending)}")

    return pending


if __name__ == "__main__":
    session = get_session()
    pending = get_pending_bitacoras(session)

    print("\n--- Bitácoras pendientes ---")
    for b in pending[:20]:
        print(f"  [{b.ficha}] {b.apprentice_name} — Bitácora #{b.bitacora_num} — {b.filename} ({b.db_status or 'sin registro'})")
