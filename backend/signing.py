"""
signing.py — Inserta la firma del instructor y exporta a PDF.

Estrategia:
  Excel → firma con openpyxl → PDF con LibreOffice headless
  PDF   → firma con pypdf + reportlab

Si LibreOffice no está disponible (ej: desarrollo en Mac),
guarda el Excel firmado como .xlsx en vez de PDF.
"""

import io
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).parent.parent / ".env")

SIGNATURE_IMAGE_PATH = Path(__file__).parent.parent / os.getenv(
    "SIGNATURE_IMAGE_PATH", "config/firma_instructor.png"
)

# Fila donde va la firma del instructor en la bitácora GFPI-F-147
# (fila 69 en Excel = índice 68 base-0, pero openpyxl usa base-1 → row=69)
INSTRUCTOR_SIGNATURE_ROW = 69
INSTRUCTOR_SIGNATURE_COL = 2  # columna B


def libreoffice_available() -> str | None:
    """Retorna el path a LibreOffice si está disponible, None si no."""
    candidates = [
        "libreoffice",
        "soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        if shutil.which(c) or Path(c).exists():
            return c
    return None


def sign_and_export(file_content: bytes, file_ext: str, original_filename: str,
                    sig_x: float = None, sig_y: float = None, sig_page: int = -1) -> tuple[bytes, str]:
    """
    Firma el PDF e retorna (output_bytes, output_filename).
    Los Excel deben convertirse a PDF antes de llamar esta función.
    sig_x, sig_y: coordenadas en puntos PDF donde centrar la firma (origen abajo-izquierda).
    sig_page: índice de página base-0 donde firmar (-1 = última página).
    """
    ext = file_ext.lower().lstrip(".")
    if ext == "pdf":
        return _sign_pdf(file_content, original_filename, sig_x, sig_y, sig_page)
    else:
        raise ValueError(f"Formato no soportado: {ext}. Convertir a PDF primero.")


# ─────────────────────────────────────────────
# EXCEL
# ─────────────────────────────────────────────

def excel_to_pdf(content: bytes) -> bytes:
    """
    Convierte contenido Excel a PDF usando el motor disponible:
    1. Excel COM (Windows + Microsoft Office)
    2. LibreOffice headless
    Lanza RuntimeError si ninguno está disponible.
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        xlsx_path = tmpdir / "bitacora.xlsx"
        pdf_path  = tmpdir / "bitacora.pdf"
        xlsx_path.write_bytes(content)

        # ── Opción 1: Excel COM (Windows + Office) ──────────────
        try:
            import win32com.client
            excel = win32com.client.Dispatch("Excel.Application")
            try:
                excel.Visible = False
                excel.DisplayAlerts = False
            except Exception:
                pass  # En algunos contextos no se puede setear, Excel ya es invisible por defecto
            wb = excel.Workbooks.Open(str(xlsx_path.resolve()))
            wb.ExportAsFixedFormat(0, str(pdf_path.resolve()))  # 0 = xlTypePDF
            wb.Close(False)
            excel.Quit()
            if pdf_path.exists():
                pdf_bytes = pdf_path.read_bytes()
                print(f"  ✓ Excel → PDF via Microsoft Excel ({len(pdf_bytes):,} bytes)")
                return pdf_bytes
        except Exception as e:
            print(f"  · Excel COM no disponible: {e}")

        # ── Opción 2: LibreOffice ────────────────────────────────
        lo = libreoffice_available()
        if lo:
            result = subprocess.run(
                [lo, "--headless", "--convert-to", "pdf",
                 "--outdir", str(tmpdir), str(xlsx_path)],
                capture_output=True, timeout=60
            )
            if result.returncode == 0 and pdf_path.exists():
                pdf_bytes = pdf_path.read_bytes()
                print(f"  ✓ Excel → PDF via LibreOffice ({len(pdf_bytes):,} bytes)")
                return pdf_bytes
            print(f"  · LibreOffice error: {result.stderr.decode()[:200]}")

    raise RuntimeError(
        "No se pudo convertir Excel a PDF. "
        "Se requiere Microsoft Excel (Windows) o LibreOffice instalado."
    )


def _find_bitacora_sheet(wb):
    """Busca la hoja principal de la bitácora."""
    keywords = ["bitacora", "bitácora", "147", "formato"]
    for name in wb.sheetnames:
        if any(k in name.lower() for k in keywords):
            return wb[name]
    return wb.active


# ─────────────────────────────────────────────
# PDF
# ─────────────────────────────────────────────

def _sign_pdf(content: bytes, original_filename: str,
              sig_x: float = None, sig_y: float = None, sig_page: int = -1) -> tuple[bytes, str]:
    """
    Superpone la imagen de firma sobre el PDF.
    sig_x, sig_y: centro de la firma en puntos PDF (origen abajo-izquierda).
                  Si son None, usa posición por defecto (18% desde abajo, 10% desde izquierda).
    sig_page: índice base-0 de la página a firmar (-1 = última).
    """
    import pypdf
    from reportlab.pdfgen import canvas as rl_canvas

    reader = pypdf.PdfReader(io.BytesIO(content))
    n_pages = len(reader.pages)
    target_idx = sig_page if (sig_page is not None and 0 <= sig_page < n_pages) else n_pages - 1
    target_page = reader.pages[target_idx]

    page_width  = float(target_page.mediabox.width)
    page_height = float(target_page.mediabox.height)

    # Tamaño de la firma (proporcional al ancho de página)
    sig_w = page_width * 0.25
    sig_h = 45.0

    # Posición: usar coordenadas del usuario o fallback
    if sig_x is not None and sig_y is not None:
        draw_x = sig_x - sig_w / 2   # centrar horizontalmente en el click
        draw_y = sig_y - sig_h / 2   # centrar verticalmente en el click
    else:
        draw_x = page_width * 0.10
        draw_y = page_height * 0.18

    # Crear overlay del tamaño exacto de la página
    overlay_buf = io.BytesIO()
    c = rl_canvas.Canvas(overlay_buf, pagesize=(page_width, page_height))

    if SIGNATURE_IMAGE_PATH.exists():
        c.drawImage(
            str(SIGNATURE_IMAGE_PATH),
            draw_x, draw_y,
            width=sig_w, height=sig_h,
            preserveAspectRatio=True,
            mask="auto"
        )
        print(f"  ✓ Firma en página {target_idx + 1} en ({draw_x:.0f}, {draw_y:.0f})")
    else:
        c.setFont("Helvetica-Bold", 9)
        c.drawString(draw_x, draw_y, "LUIS EDUARDO GONZALEZ")
        print("  ⚠ Sin imagen de firma — texto insertado en PDF")

    c.save()
    overlay_buf.seek(0)

    # Merge overlay solo sobre la página objetivo
    overlay_reader = pypdf.PdfReader(overlay_buf)
    writer = pypdf.PdfWriter()
    for i, page in enumerate(reader.pages):
        if i == target_idx:
            page.merge_page(overlay_reader.pages[0])
        writer.add_page(page)

    out = io.BytesIO()
    writer.write(out)
    base = original_filename.rsplit(".", 1)[0]
    print(f"  ✓ PDF firmado ({len(out.getvalue()):,} bytes)")
    return out.getvalue(), f"{base}_FIRMADO.pdf"
