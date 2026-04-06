"""
setup_test.py — Crea carpetas y archivos de prueba en OneDrive y SharePoint.
Ejecutar una vez para preparar el entorno de pruebas.
"""

import sys
import urllib.parse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from sharepoint import (
    get_session, upload_file, list_folders,
    ONEDRIVE_BASE_URL, SHAREPOINT_SITE_URL, SHAREPOINT_TARGET_LIBRARY
)

# ── Rutas de prueba ──────────────────────────────────────────
ONEDRIVE_TEST_FOLDER   = "/personal/lugonzalezs_sena_edu_co/Documents/Instructor Luis Eduardo/_TEST"
BITACORAS_TEST_FOLDER  = f"{ONEDRIVE_TEST_FOLDER}/APRENDIZ PRUEBA/1 - BITÁCORAS"
SHAREPOINT_TEST_FOLDER = f"/teams/ETAPASPRODUCTIVAS2025SHP/{SHAREPOINT_TARGET_LIBRARY}/_TEST"

SAMPLE_FILE = Path(__file__).parent.parent / "data/samples/BITACORA- 1 - 22 DICIEMBRE AL 23 DE ENERO.xls.xlsx"
SAMPLE_PDF  = Path(__file__).parent.parent / "data/samples/bitacora_prueba.pdf"


def create_folder(session, base_url: str, parent_path: str, folder_name: str) -> bool:
    """Crea una carpeta en SharePoint/OneDrive."""
    p = urllib.parse.quote(f"'{parent_path}'")
    r_digest = session.post(
        f"{base_url}/_api/contextinfo",
        headers={"Accept": "application/json;odata=verbose", "Content-Length": "0"}
    )
    if not r_digest.ok:
        print(f"  ✗ contextinfo failed")
        return False
    digest = r_digest.json()["d"]["GetContextWebInformation"]["FormDigestValue"]

    url = f"{base_url}/_api/web/GetFolderByServerRelativeUrl(@p1)/Folders/add('{urllib.parse.quote(folder_name)}')?@p1={p}"
    r = session.post(url, headers={
        "Accept": "application/json;odata=verbose",
        "X-RequestDigest": digest,
        "Content-Type": "application/json;odata=verbose",
    })
    if r.ok or "already exists" in r.text.lower() or r.status_code == 409:
        return True
    print(f"  ✗ Error creando carpeta '{folder_name}': {r.status_code} {r.text[:150]}")
    return False


def ensure_folder_path(session, base_url: str, full_path: str):
    """Crea todas las carpetas del path si no existen, nivel por nivel."""
    # Detectar raíz según base_url
    if "sena4-my" in base_url:
        root = "/personal/lugonzalezs_sena_edu_co"
    else:
        root = "/teams/ETAPASPRODUCTIVAS2025SHP"

    relative = full_path[len(root):]  # quitar el prefijo de raíz
    parts = [p for p in relative.split("/") if p]

    current = root
    for part in parts:
        ok = create_folder(session, base_url, current, part)
        status = "✓" if ok else "✗"
        print(f"  {status} {current}/{part}")
        current = f"{current}/{part}"


def create_test_pdf():
    """Crea un PDF de prueba simple si no existe."""
    if SAMPLE_PDF.exists():
        return
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.pagesizes import A4
        c = canvas.Canvas(str(SAMPLE_PDF), pagesize=A4)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(50, 780, "BITÁCORA DE PRUEBA - APRENDIZ PRUEBA")
        c.setFont("Helvetica", 12)
        c.drawString(50, 750, "Nombre: Aprendiz Prueba")
        c.drawString(50, 730, "Ficha: 9999999")
        c.drawString(50, 710, "Período: 01/01/2026 al 31/01/2026")
        c.drawString(50, 690, "Empresa: Empresa de Prueba S.A.S.")
        c.drawString(50, 650, "Actividades realizadas:")
        c.drawString(70, 630, "- Actividad 1: Descripción de prueba")
        c.drawString(70, 610, "- Actividad 2: Otra actividad de prueba")
        c.drawString(50, 560, "Firma aprendiz: ________________")
        c.drawString(50, 520, "Firma jefe inmediato: ________________")
        c.drawString(50, 480, "Firma instructor: ________________")
        c.save()
        print(f"  ✓ PDF de prueba creado: {SAMPLE_PDF}")
    except ImportError:
        print("  ✗ reportlab no disponible, no se creó PDF de prueba")


if __name__ == "__main__":
    session = get_session()

    print("\n=== Creando carpetas en OneDrive ===")
    ensure_folder_path(session, ONEDRIVE_BASE_URL, BITACORAS_TEST_FOLDER)

    print("\n=== Subiendo bitácora Excel de prueba a OneDrive ===")
    if SAMPLE_FILE.exists():
        content = SAMPLE_FILE.read_bytes()
        ok = upload_file(session, ONEDRIVE_BASE_URL, BITACORAS_TEST_FOLDER,
                         SAMPLE_FILE.name, content)
        print(f"  {'✓' if ok else '✗'} {SAMPLE_FILE.name}")
    else:
        print(f"  ✗ No se encontró el archivo de muestra: {SAMPLE_FILE}")

    print("\n=== Creando PDF de prueba y subiéndolo a OneDrive ===")
    create_test_pdf()
    if SAMPLE_PDF.exists():
        content = SAMPLE_PDF.read_bytes()
        ok = upload_file(session, ONEDRIVE_BASE_URL, BITACORAS_TEST_FOLDER,
                         SAMPLE_PDF.name, content)
        print(f"  {'✓' if ok else '✗'} {SAMPLE_PDF.name}")

    print("\n=== Creando carpeta de prueba en SharePoint ===")
    ensure_folder_path(session, SHAREPOINT_SITE_URL, SHAREPOINT_TEST_FOLDER)

    print("\n=== Listo ===")
    print(f"OneDrive prueba : {BITACORAS_TEST_FOLDER}")
    print(f"SharePoint prueba: {SHAREPOINT_TEST_FOLDER}")
