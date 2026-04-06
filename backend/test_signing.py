"""
test_signing.py — Prueba el flujo completo de firma y subida al SharePoint de prueba.

Descarga los archivos de prueba desde OneDrive _TEST,
firma con signing.py y sube al SharePoint _TEST.

Ejecutar: conda run -n autouncle python backend/test_signing.py
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from sharepoint import (
    get_session, download_file, upload_file, list_files,
    ONEDRIVE_BASE_URL, SHAREPOINT_SITE_URL, SHAREPOINT_TARGET_LIBRARY
)
import signing

ONEDRIVE_TEST_PATH = "/personal/lugonzalezs_sena_edu_co/Documents/Instructor Luis Eduardo/_TEST/APRENDIZ PRUEBA/1 - BITÁCORAS"
SHAREPOINT_TEST_FOLDER = f"/teams/ETAPASPRODUCTIVAS2025SHP/{SHAREPOINT_TARGET_LIBRARY}/_TEST"


def test_file(session, filename: str, file_path: str):
    print(f"\n{'='*60}")
    print(f"Probando: {filename}")
    print(f"{'='*60}")

    # 1. Descargar desde OneDrive
    print("  1. Descargando desde OneDrive...")
    content = download_file(session, ONEDRIVE_BASE_URL, file_path)
    if not content:
        print("  ✗ Error descargando el archivo")
        return False
    print(f"  ✓ Descargado ({len(content):,} bytes)")

    # 2. Firmar
    ext = filename.rsplit(".", 1)[-1].lower()
    print(f"  2. Firmando (ext={ext})...")
    try:
        signed_bytes, signed_name = signing.sign_and_export(content, ext, filename)
        print(f"  ✓ Firmado: {signed_name} ({len(signed_bytes):,} bytes)")
    except Exception as e:
        print(f"  ✗ Error firmando: {e}")
        return False

    # 3. Subir a SharePoint _TEST
    print(f"  3. Subiendo a SharePoint _TEST...")
    ok = upload_file(session, SHAREPOINT_SITE_URL, SHAREPOINT_TEST_FOLDER, signed_name, signed_bytes)
    if ok:
        print(f"  ✓ Subido: {signed_name}")
        print(f"     → {SHAREPOINT_TEST_FOLDER}/{signed_name}")
    else:
        print(f"  ✗ Error subiendo a SharePoint")
    return ok


def main():
    session = get_session()

    # Listar archivos en el folder de prueba de OneDrive
    print(f"\nListando archivos en OneDrive _TEST: {ONEDRIVE_TEST_PATH}")
    files = list_files(session, ONEDRIVE_BASE_URL, ONEDRIVE_TEST_PATH)
    if not files:
        print("  ✗ No se encontraron archivos en el folder de prueba")
        print("  Ejecuta primero: conda run -n autouncle python backend/setup_test.py")
        return

    print(f"  Encontrados {len(files)} archivo(s):")
    for f in files:
        print(f"    - {f['Name']} ({f.get('Length', '?')} bytes)")

    results = []
    for f in files:
        ok = test_file(session, f["Name"], f["ServerRelativeUrl"])
        results.append((f["Name"], ok))

    print(f"\n{'='*60}")
    print("RESUMEN:")
    for name, ok in results:
        status = "✓" if ok else "✗"
        print(f"  {status} {name}")
    print(f"\nArchivos firmados subidos a:")
    print(f"  SharePoint: {SHAREPOINT_TEST_FOLDER}")


if __name__ == "__main__":
    main()
