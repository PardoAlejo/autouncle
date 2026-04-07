"""
app.py — FastAPI app principal de AutoUncle.
"""

import sys
import os
import io
import threading
import urllib.parse
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, Request, Form, BackgroundTasks
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from dotenv import load_dotenv

sys.path.insert(0, str(Path(__file__).parent))

load_dotenv(Path(__file__).parent.parent / ".env")

from database import init_db, get_all, get_pending, get_stats, update_status, get_conn, upsert_bitacoras
from sharepoint import get_session, list_folders, list_files, download_file, upload_file, find_sharepoint_dest, SessionExpiredError, ONEDRIVE_BASE_URL, SHAREPOINT_SITE_URL, SHAREPOINT_TARGET_LIBRARY
from sync import get_pending_bitacoras, load_db
from tracking import push_to_onedrive, pull_from_onedrive
import signing  # firma y PDF

TEMPLATES_DIR = Path(__file__).parent.parent / "frontend" / "templates"
ONEDRIVE_BASE_URL_ENV = os.getenv("ONEDRIVE_BASE_URL")

app = FastAPI()
templates = Jinja2Templates(directory=str(TEMPLATES_DIR))


@app.exception_handler(SessionExpiredError)
async def session_expired_handler(request: Request, exc: SessionExpiredError):
    return HTMLResponse(
        """
        <html><head><title>Sesión expirada</title>
        <style>body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;background:#f5f5f5}
        .box{background:#fff;border-radius:12px;padding:2.5rem 3rem;text-align:center;box-shadow:0 4px 20px rgba(0,0,0,.1);max-width:420px}
        h2{color:#c62828;margin-top:0} p{color:#555;line-height:1.6}
        .btn{display:inline-block;margin-top:1rem;padding:.7rem 1.6rem;background:#1a73e8;color:#fff;border-radius:6px;text-decoration:none;font-weight:600;font-size:1rem;border:none;cursor:pointer}</style>
        </head><body><div class="box">
        <h2>⚠ Sesión expirada</h2>
        <p>Las credenciales de SharePoint vencieron.<br>Haz clic para abrir el navegador e iniciar sesión con Microsoft.</p>
        <form method="post" action="/reauth">
          <button type="submit" class="btn">🔑 Renovar sesión</button>
        </form>
        </div></body></html>
        """,
        status_code=401,
    )


# ── Re-authentication ──────────────────────────────────────────────────────

_reauth_state: dict = {
    "running": False, "done": False, "error": None,
    "needs_code": False, "submitted_code": None,
}
_reauth_lock = threading.Lock()
_reauth_code_event = threading.Event()


def _run_reauth():
    import auth as auth_module
    from playwright.sync_api import sync_playwright

    try:
        with sync_playwright() as p:
            context = auth_module.get_context(p, headless=False)
            context.browser.close()
        _reauth_state["done"] = True
        _reauth_state["error"] = None
    except Exception as e:
        _reauth_state["error"] = str(e)
    finally:
        _reauth_state["running"] = False


@app.post("/reauth")
def reauth_start():
    with _reauth_lock:
        if not _reauth_state["running"]:
            _reauth_code_event.clear()
            _reauth_state.update({
                "running": True, "done": False, "error": None,
                "needs_code": False, "submitted_code": None,
            })
            threading.Thread(target=_run_reauth, daemon=True).start()
    return HTMLResponse("""
        <html><head><title>Renovando sesión...</title>
        <style>
        body{font-family:sans-serif;display:flex;align-items:center;justify-content:center;
             height:100vh;margin:0;background:#f5f5f5}
        .box{background:#fff;border-radius:12px;padding:2.5rem 3rem;text-align:center;
             box-shadow:0 4px 20px rgba(0,0,0,.1);max-width:440px;width:90%}
        h2{margin-top:0;color:#1a73e8} p{color:#555;line-height:1.6}
        .spinner{font-size:2rem;animation:spin 1.2s linear infinite;display:inline-block;margin-bottom:.5rem}
        @keyframes spin{to{transform:rotate(360deg)}}
        .error{color:#c62828;background:#fff3f3;border:1px solid #ffcdd2;
               border-radius:8px;padding:.7rem;margin-top:1rem}
        input[type=text]{width:100%;padding:.7rem;font-size:1.3rem;text-align:center;
                         letter-spacing:.3rem;border:2px solid #1a73e8;border-radius:8px;
                         box-sizing:border-box;margin:.8rem 0;font-family:monospace}
        button{padding:.65rem 1.8rem;background:#1a73e8;color:#fff;border:none;
               border-radius:8px;font-size:1rem;font-weight:600;cursor:pointer;width:100%}
        </style></head><body><div class="box">
        <div id="spinner-section">
          <div class="spinner">🔄</div>
          <h2>Renovando sesión</h2>
          <p>Se abrió una ventana del navegador.<br><br>
          <strong>Inicia sesión con tu cuenta de Microsoft</strong><br>
          e ingresa el código SMS directamente en esa ventana.<br><br>
          <span style="color:#888;font-size:.9rem">Esta página se actualizará sola cuando termines.</span></p>
        </div>
        <div id="error-section" style="display:none" class="error"></div>
        </div>
        <script>
        function poll() {
          fetch('/reauth/status').then(r => r.json()).then(d => {
            if (d.done) { window.location.href = '/'; return; }
            if (d.error) {
              document.getElementById('spinner-section').style.display = 'none';
              document.getElementById('error-section').style.display = 'block';
              document.getElementById('error-section').textContent = '✗ ' + d.error;
              return;
            }
            setTimeout(poll, 2000);
          });
        }

        setTimeout(poll, 2000);
        </script></body></html>
    """)


@app.post("/reauth/submit-code")
async def reauth_submit_code(request: Request):
    form = await request.form()
    _reauth_state["submitted_code"] = (form.get("code") or "").strip()
    _reauth_code_event.set()
    return JSONResponse({"ok": True})


@app.get("/reauth/status")
def reauth_status():
    return JSONResponse({k: v for k, v in _reauth_state.items()
                         if k not in ("submitted_code",)})

# DB e inicialización al arrancar
@app.on_event("startup")
def startup():
    init_db()
    # Restaurar estado desde OneDrive al arrancar
    try:
        pull_from_onedrive()
    except Exception as e:
        print(f"Warning: no se pudo restaurar tracking desde OneDrive: {e}")


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def get_bitacora_by_id(bitacora_id: int):
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM bitacoras WHERE id = ?", (bitacora_id,)).fetchone()
    return row


def onedrive_web_url(server_relative_path: str) -> str:
    """Convierte un server-relative path a URL web de OneDrive para abrir en browser."""
    encoded = urllib.parse.quote(server_relative_path)
    return f"{ONEDRIVE_BASE_URL_ENV}/_layouts/15/download.aspx?SourceUrl={encoded}"


def onedrive_view_url(server_relative_path: str) -> str:
    """URL para abrir el archivo en el viewer web de OneDrive."""
    encoded = urllib.parse.quote(server_relative_path)
    return f"{ONEDRIVE_BASE_URL_ENV}/_layouts/15/WopiFrame.aspx?sourcedoc={encoded}&action=view"


def sharepoint_dest_folder(area: str) -> str:
    """Determina la subcarpeta de destino en SharePoint según el área del aprendiz."""
    area_upper = (area or "").upper()
    if "CONTAB" in area_upper:
        sub = "2. RED CONTABLE"
    else:
        sub = "1. RED ADMINISTRACION"
    return f"/teams/ETAPASPRODUCTIVAS2025SHP/{SHAREPOINT_TARGET_LIBRARY}/{sub}"


# ─────────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────────

TEST_IDS = [int(i) for i in os.getenv("TEST_IDS", "").split(",") if i.strip()]

# Cache de Excel→PDF para preview (evita convertir dos veces)
_pdf_cache: dict[int, bytes] = {}


@app.get("/", response_class=HTMLResponse)
def dashboard(request: Request,
              status: str = "pending",
              ficha: str = "",
              q: str = "",
              test: str = ""):
    rows = get_all()
    stats = get_stats()
    total = sum(stats.values())

    # Filtro modo test
    if test and TEST_IDS:
        rows = [r for r in rows if r["id"] in TEST_IDS]
    else:
        # Filtros normales
        if status:
            rows = [r for r in rows if r["status"] == status]
        if ficha:
            rows = [r for r in rows if r["ficha"] == ficha]
        if q:
            q_lower = q.lower()
            rows = [r for r in rows if q_lower in r["apprentice_name"].lower()]

    rows.sort(key=lambda r: (r["apprentice_name"] or "", r["bitacora_num"] or 0))

    # Group by apprentice for accordion view
    groups = []
    for r in rows:
        if not groups or groups[-1]["name"] != r["apprentice_name"]:
            groups.append({
                "name": r["apprentice_name"] or "—",
                "ficha": r["ficha"],
                "has_pending": False,
                "rows": [],
            })
        groups[-1]["rows"].append(r)
        if r["status"] == "pending":
            groups[-1]["has_pending"] = True

    all_rows = get_all()
    fichas = sorted(set(r["ficha"] for r in all_rows if r["ficha"]))

    response = templates.TemplateResponse(request, "dashboard.html", {
        "groups": groups,
        "rows": rows,
        "stats": stats,
        "total": total,
        "fichas": fichas,
        "filter_status": status,
        "filter_ficha": ficha,
        "filter_q": q,
        "test_mode": bool(TEST_IDS),
        "filter_test": bool(test),
    })
    response.headers["Cache-Control"] = "no-store"
    return response


@app.get("/bitacora/{bitacora_id}", response_class=HTMLResponse)
def review(request: Request, bitacora_id: int,
           message: str = "", message_type: str = ""):
    b = get_bitacora_by_id(bitacora_id)
    if not b:
        return RedirectResponse("/")

    ext = (b["file_ext"] or "").lower()
    if ext == "pdf":
        # Servir a través de nuestra app — el browser no tiene la sesión de SharePoint
        preview_url = f"/bitacora/{bitacora_id}/download"
    else:
        # Excel: abrir directo en OneDrive (el tío ya está logueado ahí)
        preview_url = onedrive_view_url(b["file_path"])

    # Bitácoras rechazadas anteriores del mismo aprendiz (excluyendo la actual)
    with get_conn() as conn:
        rejected_prev = conn.execute("""
            SELECT filename, notes, reviewed_at, bitacora_num
            FROM bitacoras
            WHERE status = 'rejected'
              AND id != ?
              AND (cc = ? OR (cc IS NULL AND apprentice_name = ?))
            ORDER BY reviewed_at DESC
        """, (bitacora_id, b["cc"], b["apprentice_name"])).fetchall()

    return templates.TemplateResponse(request, "review.html", {
        "b": b,
        "preview_url": preview_url,
        "message": message,
        "message_type": message_type,
        "rejected_prev": rejected_prev,
    })


@app.get("/bitacora/{bitacora_id}/download")
def download_bitacora(bitacora_id: int):
    """Descarga el archivo desde OneDrive y lo sirve al browser."""
    b = get_bitacora_by_id(bitacora_id)
    if not b:
        return RedirectResponse("/")
    session = get_session()
    content = download_file(session, ONEDRIVE_BASE_URL, b["file_path"])
    if not content:
        return HTMLResponse("Error descargando archivo", status_code=500)

    media_type = "application/pdf" if b["file_ext"] == "pdf" else "application/octet-stream"
    return StreamingResponse(
        io.BytesIO(content),
        media_type=media_type,
        headers={"Content-Disposition": f"inline; filename={b['filename']}"}
    )


@app.get("/signature-preview")
def signature_preview():
    """Sirve la imagen de firma para el pin de preview en el visor PDF."""
    from signing import SIGNATURE_IMAGE_PATH
    if SIGNATURE_IMAGE_PATH.exists():
        return StreamingResponse(
            open(SIGNATURE_IMAGE_PATH, "rb"),
            media_type="image/png",
            headers={"Cache-Control": "max-age=3600"}
        )
    return HTMLResponse("", status_code=404)


@app.get("/bitacora/{bitacora_id}/preview-pdf")
def preview_pdf(bitacora_id: int):
    """
    Retorna el PDF listo para mostrar en el visor.
    Para Excel: convierte a PDF (y cachea el resultado para reutilizar al firmar).
    Para PDF: sirve directamente.
    """
    b = get_bitacora_by_id(bitacora_id)
    if not b:
        return HTMLResponse("No encontrado", status_code=404)

    session = get_session()
    ext = (b["file_ext"] or "").lower()

    if ext in ("xlsx", "xls"):
        if bitacora_id not in _pdf_cache:
            content = download_file(session, ONEDRIVE_BASE_URL, b["file_path"])
            if not content:
                return HTMLResponse("Error descargando archivo", status_code=500)
            try:
                _pdf_cache[bitacora_id] = signing.excel_to_pdf(content)
            except RuntimeError as e:
                return HTMLResponse(str(e), status_code=500)
        pdf_bytes = _pdf_cache[bitacora_id]
    else:
        pdf_bytes = download_file(session, ONEDRIVE_BASE_URL, b["file_path"])
        if not pdf_bytes:
            return HTMLResponse("Error descargando archivo", status_code=500)

    return StreamingResponse(
        io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={"Cache-Control": "no-store"}
    )


@app.get("/bitacora/{bitacora_id}/download-signed")
def download_signed(bitacora_id: int):
    """Descarga el archivo firmado desde SharePoint y lo sirve al browser."""
    b = get_bitacora_by_id(bitacora_id)
    if not b or not b["signed_pdf_path"]:
        return HTMLResponse("Archivo firmado no disponible", status_code=404)
    session = get_session()
    content = download_file(session, SHAREPOINT_SITE_URL, b["signed_pdf_path"])
    if not content:
        return HTMLResponse("Error descargando archivo firmado", status_code=500)

    filename = b["signed_pdf_path"].rsplit("/", 1)[-1]
    ext = filename.rsplit(".", 1)[-1].lower()
    media_type = "application/pdf" if ext == "pdf" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return StreamingResponse(
        io.BytesIO(content),
        media_type=media_type,
        headers={"Content-Disposition": f"inline; filename={filename}"}
    )


@app.get("/bitacora/{bitacora_id}/resolve-dest")
def resolve_dest(bitacora_id: int):
    """
    Resuelve la carpeta destino en SharePoint para el aprendiz.
    Retorna JSON con 'path' (ruta server-relative), 'display' (ruta legible) y 'found' (bool).
    """
    b = get_bitacora_by_id(bitacora_id)
    if not b:
        return JSONResponse({"error": "not found"}, status_code=404)

    if TEST_IDS and b["id"] in TEST_IDS:
        path = f"/teams/ETAPASPRODUCTIVAS2025SHP/{SHAREPOINT_TARGET_LIBRARY}/_TEST"
        return JSONResponse({"path": path, "display": "_TEST", "found": True})

    session = get_session()
    path = find_sharepoint_dest(session, b["ficha"], b["cc"], b["apprentice_name"], b["area"])
    if not path:
        path = sharepoint_dest_folder(b["area"])
        found = False
    else:
        found = True

    # Ruta legible: quitar el prefijo largo de SharePoint
    prefix = f"/teams/ETAPASPRODUCTIVAS2025SHP/{SHAREPOINT_TARGET_LIBRARY}/"
    display = path[len(prefix):] if path.startswith(prefix) else path

    return JSONResponse({"path": path, "display": display, "found": found})


@app.post("/bitacora/{bitacora_id}/approve")
def approve(bitacora_id: int, background_tasks: BackgroundTasks,
            notes: str = Form(default=""),
            dest_folder: str = Form(default=""),
            sig_x: Optional[float] = Form(default=None),
            sig_y: Optional[float] = Form(default=None),
            sig_page: int = Form(default=-1)):
    """
    Flujo de aprobación:
    1. Descarga el archivo desde OneDrive
    2. Inserta la firma del instructor
    3. Convierte a PDF
    4. Sube al SharePoint en la carpeta correcta
    5. Marca como aprobada en la DB
    """
    b = get_bitacora_by_id(bitacora_id)
    if not b:
        return RedirectResponse("/", status_code=303)

    session = get_session()

    try:
        # 1. Descargar archivo
        content = download_file(session, ONEDRIVE_BASE_URL, b["file_path"])
        if not content:
            return RedirectResponse(
                f"/bitacora/{bitacora_id}?message=Error+descargando+el+archivo&message_type=error",
                status_code=303
            )

        # 2. Si es Excel, usar PDF cacheado del preview (o convertir si no hay cache)
        ext = (b["file_ext"] or "").lower()
        if ext in ("xlsx", "xls"):
            if bitacora_id in _pdf_cache:
                content = _pdf_cache.pop(bitacora_id)
            else:
                content = signing.excel_to_pdf(content)
            ext = "pdf"

        # 3. Firmar el PDF en las coordenadas que eligió el tío
        pdf_bytes, pdf_name = signing.sign_and_export(
            content, ext, b["filename"],
            sig_x=sig_x, sig_y=sig_y, sig_page=sig_page
        )

        # 4. Usar la carpeta destino confirmada por el usuario
        if not dest_folder:
            # No debería ocurrir si el flujo de UI es correcto, pero por si acaso
            dest_folder = sharepoint_dest_folder(b["area"])
            print(f"  ⚠ dest_folder vacío — usando fallback {dest_folder}")
        ok = upload_file(session, SHAREPOINT_SITE_URL, dest_folder, pdf_name, pdf_bytes)

        if not ok:
            return RedirectResponse(
                f"/bitacora/{bitacora_id}?message=Error+subiendo+a+SharePoint&message_type=error",
                status_code=303
            )

        # 4b. Guardar copia en OneDrive (misma carpeta que el archivo original)
        if TEST_IDS and b["id"] in TEST_IDS:
            user = ONEDRIVE_BASE_URL_ENV.split("/personal/")[1]
            onedrive_folder = f"/personal/{user}{os.getenv('ONEDRIVE_ROOT_FOLDER', '')}/_TEST"
        else:
            onedrive_folder = b["file_path"].rsplit("/", 1)[0]
        od_ok = upload_file(session, ONEDRIVE_BASE_URL, onedrive_folder, pdf_name, pdf_bytes)
        if not od_ok:
            print(f"  ⚠ No se pudo guardar copia en OneDrive ({onedrive_folder}) — continúa igual")

        # 5. Marcar como aprobada
        signed_path = f"{dest_folder}/{pdf_name}"
        with get_conn() as conn:
            conn.execute(
                "UPDATE bitacoras SET status='approved', notes=?, reviewed_at=datetime('now'), signed_pdf_path=? WHERE id=?",
                (notes, signed_path, bitacora_id)
            )

        background_tasks.add_task(push_to_onedrive)
        return RedirectResponse(
            f"/bitacora/{bitacora_id}?message=Bitácora+firmada+y+subida+correctamente&message_type=success",
            status_code=303
        )

    except Exception as e:
        return RedirectResponse(
            f"/bitacora/{bitacora_id}?message=Error:+{str(e)[:80]}&message_type=error",
            status_code=303
        )


@app.post("/bitacora/{bitacora_id}/reject")
def reject(bitacora_id: int, background_tasks: BackgroundTasks, notes: str = Form(default="")):
    update_status(bitacora_id, "rejected", notes)
    background_tasks.add_task(push_to_onedrive)
    return RedirectResponse(
        f"/bitacora/{bitacora_id}?message=Bitácora+marcada+como+rechazada&message_type=info",
        status_code=303
    )


@app.post("/bitacora/{bitacora_id}/already-done")
def already_done(bitacora_id: int, background_tasks: BackgroundTasks):
    update_status(bitacora_id, "approved", "Ya estaba completa")
    background_tasks.add_task(push_to_onedrive)
    return RedirectResponse(
        f"/bitacora/{bitacora_id}?message=Bitácora+marcada+como+aprobada&message_type=success",
        status_code=303
    )


@app.post("/bitacora/{bitacora_id}/skip")
def skip(bitacora_id: int, background_tasks: BackgroundTasks):
    update_status(bitacora_id, "skipped")
    background_tasks.add_task(push_to_onedrive)
    return RedirectResponse("/", status_code=303)


@app.post("/sync")
def sync_now():
    """Sincroniza OneDrive y actualiza la base de datos."""
    try:
        session = get_session()
        bitacoras = get_pending_bitacoras(session)
        new, updated = upsert_bitacoras(bitacoras)
        return RedirectResponse(
            f"/?message=Sync+completado:+{new}+nuevas,+{updated}+actualizadas&message_type=success",
            status_code=303
        )
    except SessionExpiredError:
        raise
    except Exception as e:
        return RedirectResponse(
            f"/?message=Error+en+sync:+{str(e)[:80]}&message_type=error",
            status_code=303
        )
