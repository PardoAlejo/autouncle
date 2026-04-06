"""
run.py — Lanzador de AutoUncle.

Uso normal:
    python run.py

Modo test (resetea los registros de prueba a 'pending' antes de arrancar):
    python run.py --test
"""

import sys
import os
import subprocess
import argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "backend"))

REQUIREMENTS = Path(__file__).parent / "requirements.txt"


def ensure_dependencies():
    """Instala dependencias faltantes automáticamente."""
    # Leer paquetes del requirements.txt (ignorar comentarios y líneas vacías)
    packages = [
        line.strip()
        for line in REQUIREMENTS.read_text().splitlines()
        if line.strip() and not line.startswith("#")
    ]

    # Paquetes donde nombre de instalación ≠ nombre de módulo
    import_name = {
        "pillow": "PIL",
        "python-dotenv": "dotenv",
        "python-multipart": "multipart",
        "pypdf": "pypdf",
        "reportlab": "reportlab",
    }

    import importlib
    missing = []
    for pkg in packages:
        install_name = pkg.split("==")[0].split(">=")[0].strip()
        module = import_name.get(install_name, install_name.replace("-", "_"))
        try:
            importlib.import_module(module)
        except ImportError:
            missing.append(pkg)

    if missing:
        print(f"Instalando dependencias faltantes: {', '.join(missing)}")
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])
        print("✓ Dependencias instaladas\n")

    # pywin32 solo en Windows
    if sys.platform == "win32":
        try:
            import win32com.client  # noqa
        except ImportError:
            print("Instalando pywin32 (necesario para Excel → PDF en Windows)...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
            print("✓ pywin32 instalado\n")


TEST_IDS = [
    122,  # Test 1 (PDF)
    98,   # Test 2 (PDF)
    155,  # Test 3 (PDF)
    16,   # Test 4 (PDF)
    162,  # Test 5 (Excel)
    264,  # Test 6 (Excel)
    26,   # Test 7 (Excel)
    71,   # Test 8 (Excel)
]


def reset_test_records():
    from database import init_db, get_conn
    init_db()
    placeholders = ",".join("?" * len(TEST_IDS))
    with get_conn() as conn:
        updated = conn.execute(
            f"UPDATE bitacoras SET status='pending', signed_pdf_path=NULL, reviewed_at=NULL, notes=NULL "
            f"WHERE id IN ({placeholders})",
            TEST_IDS
        ).rowcount
    print(f"[TEST] {updated} registro(s) de prueba reseteados a 'pending'")
    print(f"[TEST] IDs: {TEST_IDS}")


def main():
    ensure_dependencies()

    parser = argparse.ArgumentParser(description="AutoUncle launcher")
    parser.add_argument("--test", action="store_true",
                        help="Resetea registros de prueba a 'pending' antes de arrancar")
    parser.add_argument("--port", type=int, default=8000)
    args = parser.parse_args()

    if args.test:
        reset_test_records()
        os.environ["TEST_IDS"] = ",".join(str(i) for i in TEST_IDS)

    import uvicorn
    uvicorn.run(
        "backend.app:app",
        host="0.0.0.0",
        port=args.port,
        reload=True,
        reload_dirs=["backend", "frontend"],
    )


if __name__ == "__main__":
    main()
