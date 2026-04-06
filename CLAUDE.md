# AutoUncle — Contexto para Claude

## Qué es este proyecto

Herramienta web para que el instructor Luis Eduardo González (SENA) gestione la revisión de bitácoras de sus ~100 aprendices. Automatiza la descarga desde OneDrive, la firma del instructor y la subida a SharePoint.

## Stack

- **Backend**: Python 3.11, FastAPI, SQLite (vía `database.py`)
- **Frontend**: Jinja2 + HTMX, sin frameworks JS
- **Auth**: Playwright (headless tras primer login MFA), cookies FedAuth/rtFa guardadas en `config/browser_session/state.json`
- **Storage**: OneDrive personal + SharePoint organizacional de SENA, accedidos vía SharePoint REST API con las cookies guardadas
- **Entorno**: conda env llamado `autouncle` — SIEMPRE usar `conda run -n autouncle python` para ejecutar

## Archivos clave

| Archivo | Rol |
|---------|-----|
| `backend/app.py` | FastAPI app — rutas, lógica de aprobación |
| `backend/auth.py` | Login Playwright, manejo de sesión |
| `backend/sharepoint.py` | Cliente REST OneDrive/SharePoint (`list_folders`, `list_files`, `download_file`, `upload_file`) |
| `backend/sync.py` | Escaneo de OneDrive → carga en SQLite |
| `backend/database.py` | SQLite — CRUD de bitácoras, estados |
| `backend/tracking.py` | Sync del historial con `TRACKING_BITACORAS_2026.xlsx` en OneDrive |
| `backend/signing.py` | Inserción de firma en Excel/PDF + conversión a PDF |
| `frontend/templates/` | Jinja2 templates (base, dashboard, review) |
| `config/browser_session/state.json` | Sesión Playwright guardada — no tocar |
| `data/autouncle.db` | SQLite local |
| `.env` | Credenciales — nunca commitear |

## Patrones importantes

### SharePoint REST API
Siempre usar el parámetro `@p1` para escapar rutas con espacios:
```python
p = urllib.parse.quote(f"'{server_relative_path}'")
url = f"{base_url}/_api/web/GetFolderByServerRelativeUrl(@p1)/Files?@p1={p}"
```

### Rutas OneDrive vs SharePoint
- OneDrive personal: `ONEDRIVE_BASE_URL` = `https://sena4-my.sharepoint.com/personal/lugonzalezs_sena_edu_co`
- Ruta server-relative OneDrive: `/personal/lugonzalezs_sena_edu_co/Documents/Instructor Luis Eduardo`  (incluir el prefijo `/personal/...`)
- SharePoint destino: `SHAREPOINT_SITE_URL` = `https://sena4.sharepoint.com/teams/ETAPASPRODUCTIVAS2025SHP`

### Starlette 1.0 — TemplateResponse
La versión instalada es Starlette 1.0. El `request` va como primer argumento, NO dentro del dict:
```python
# Correcto:
templates.TemplateResponse(request, "template.html", {"key": "value"})
# Incorrecto (versiones anteriores):
templates.TemplateResponse("template.html", {"request": request, "key": "value"})
```

### Estados de bitácora
```
pending  → encontrada en OneDrive, sin revisar
approved → firmada y subida a SharePoint
rejected → rechazada con nota
skipped  → ya existía antes de la app
```

### Tracking en OneDrive
- Al arrancar: `pull_from_onedrive()` restaura estados en SQLite
- Al cambiar estado: `push_to_onedrive()` corre en background task
- Archivo: `TRACKING_BITACORAS_2026.xlsx` en raíz de `/Documents/Instructor Luis Eduardo`
- Headers del Excel son los key names del modelo (no texto en español) para que el pull los encuentre

### Estructura de carpetas OneDrive (inconsistente)
Las fichas tienen estructura variable:
```
FICHA → APRENDIZ/ → Nombre/ → 1 - BITÁCORAS/
FICHA → Nombre/ → 1 - BITÁCORAS/   (sin subcarpeta APRENDIZ)
```
El sync maneja esto en `sync.py:scan_onedrive()`.

Los nombres de carpeta de bitácoras varían: "1 - BITÁCORAS", "1. BITACORAS", "BITACORAS", etc.
Fuzzy match por keywords en `sharepoint.py:find_bitacora_folder()`.

### Base de datos Excel (SharePoint)
- Ruta: `DOCUMENTOS INSTRUCTORES 2025/ADMINISTRACION Y CONTABILIDAD LUIS EDUARDO GONZALEZ/2026 LUIS EDUARDO GONZALEZ.xlsx`
- Hoja: `ASIGANCIONES 2026` (typo intencional con C extra)
- Filtrar por columna `INSTRUCTOR DE SEGUIMIENTO Escribir letra mayuscula` == `"LUIS EDUARDO GONZALEZ"`

### Destino SharePoint según área
```python
"CONTABILIDAD" → "2. RED CONTABLE"
otro           → "1. RED ADMINISTRACION"
```

## Cómo arrancar en desarrollo

```bash
conda activate autouncle
cd /Users/pardoga/Projects/autouncle
python -m uvicorn backend.app:app --port 8000 --reload
```

## Tareas pendientes

- [x] `signing.py` — firma en Excel (openpyxl, B69) + fallback xlsx cuando no hay LibreOffice
- [x] `signing.py` — firma en PDF con pypdf/reportlab (overlay en última página)
- [x] Flujo completo de aprobación end-to-end probado con archivos de prueba en _TEST
- [ ] Instalar LibreOffice en la PC de Luis Eduardo para conversión de Excel a PDF
- [ ] Probar con archivos reales (quitar _TEST y usar carpetas del tío)

## Notas de firma

- **Excel**: LibreOffice convierte a PDF primero, luego se firma el PDF — solo se sube PDF a SharePoint
- **Excel sin LibreOffice**: lanza RuntimeError — no hay fallback a xlsx, LibreOffice es requerido
- **PDF**: overlay con reportlab (posición: 10% desde izquierda, 18% desde abajo de la última página)
- **Test**: `backend/test_signing.py` — prueba descarga OneDrive _TEST → firma → subida SharePoint _TEST
