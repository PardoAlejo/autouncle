# AutoUncle

Herramienta de apoyo para que el instructor Luis Eduardo González gestione la revisión de bitácoras de sus ~100 aprendices SENA.

En vez de descargar archivos manualmente, buscarlos en carpetas y firmar uno por uno, la app centraliza todo: muestra las bitácoras pendientes, permite revisarlas, y con un solo click firma y sube el documento al SharePoint correcto.

---

## Flujo de trabajo

```
OneDrive (aprendices suben bitácoras)
         ↓  sync automático
Dashboard — cola de pendientes
         ↓  tío revisa (abre PDF o va a OneDrive para Excel)
Click "Firmar y subir"
         ↓
Firma embebida → PDF → SharePoint (carpeta correcta según área)
         ↓
Historial actualizado en TRACKING_BITACORAS_2026.xlsx (OneDrive)
```

---

## Requisitos

- Python instalado vía **conda** (environment `autouncle`)
- Cuenta Microsoft 365 de SENA (`lugonzalezs@sena.edu.co`)
- Imagen de firma del instructor en `config/firma_instructor.png`

---

## Estructura del proyecto

```
autouncle/
├── backend/
│   ├── app.py          # FastAPI — rutas principales
│   ├── auth.py         # Login Microsoft con Playwright (MFA)
│   ├── database.py     # SQLite local — tracking de estados
│   ├── sharepoint.py   # Cliente REST para OneDrive y SharePoint
│   ├── sync.py         # Escaneo de OneDrive y carga en DB
│   ├── tracking.py     # Sincronización del historial con Excel en OneDrive
│   ├── signing.py      # Firma del instructor + exportación a PDF
│   └── explore.py      # Script de exploración (desarrollo)
├── frontend/
│   └── templates/
│       ├── base.html       # Layout base
│       ├── dashboard.html  # Cola de bitácoras
│       └── review.html     # Pantalla de revisión individual
├── config/
│   ├── firma_instructor.png     # Imagen de firma (NO subir a git)
│   └── browser_session/
│       └── state.json           # Sesión de browser guardada (NO subir a git)
├── data/
│   └── autouncle.db    # Base de datos SQLite local
├── .env                # Credenciales y configuración (NO subir a git)
├── .env.example        # Plantilla de variables de entorno
└── .gitignore
```

---

## Setup inicial (primera vez)

### 1. Crear el entorno conda

```bash
conda create -n autouncle python=3.11 -y
conda activate autouncle
pip install fastapi uvicorn jinja2 python-multipart \
            playwright openpyxl pandas requests python-dotenv \
            pillow reportlab pypdf
playwright install chromium
```

### 2. Configurar variables de entorno

```bash
cp .env.example .env
```

Editar `.env` y completar:

```
MS_USERNAME=lugonzalezs@sena.edu.co
MS_PASSWORD=<contraseña>
```

El resto de variables ya tiene los valores correctos para la cuenta de Luis Eduardo.

### 3. Colocar la imagen de firma

Copiar la imagen de firma del instructor (PNG con fondo transparente) en:

```
config/firma_instructor.png
```

### 4. Hacer login por primera vez (MFA)

```bash
conda activate autouncle
python backend/auth.py
```

Se abrirá el navegador automáticamente. El flujo es:
1. El navegador abre `login.microsoftonline.com`
2. Llena usuario y contraseña automáticamente
3. **El instructor aprueba la notificación MFA en su teléfono**
4. El navegador visita OneDrive y SharePoint para capturar las cookies
5. La sesión se guarda en `config/browser_session/state.json`

> Este paso solo se repite cada ~90 días cuando la sesión expira.

### 5. Cargar bitácoras por primera vez (sync inicial)

```bash
conda activate autouncle
python backend/sync.py
```

Esto escanea todas las carpetas del OneDrive y carga las bitácoras encontradas en la base de datos local.

---

## Uso diario

### Arrancar la app

```bash
conda activate autouncle
cd /ruta/al/proyecto/autouncle
python -m uvicorn backend.app:app --host 0.0.0.0 --port 8000
```

Abrir en el navegador: **http://localhost:8000**

Al arrancar, la app automáticamente restaura el historial desde `TRACKING_BITACORAS_2026.xlsx` en OneDrive.

### Dashboard

- **Pendientes**: bitácoras que aún no han sido revisadas
- **Aprobadas**: firmadas y subidas a SharePoint
- **Rechazadas**: devueltas con observaciones
- **Ya revisadas**: existían antes de la app, marcadas manualmente

Usar el filtro de estado, ficha o nombre para encontrar bitácoras específicas.

### Revisar una bitácora

1. Hacer click en cualquier fila → abre la pantalla de revisión
2. **PDF**: se muestra embebido en la pantalla
3. **Excel**: hacer click en "Abrir en OneDrive ↗", revisar allá, volver a la app
4. Marcar los ítems de la lista de verificación:
   - Firmas presentes (aprendiz + jefe inmediato)
   - Fechas del encabezado correctas
   - Actividades cubren el período completo
   - Evidencias verificadas
5. Agregar observaciones si es necesario
6. Click **"Firmar y subir"** → la app firma, convierte a PDF y sube a SharePoint
7. O click **"Rechazar"** → queda marcada para seguimiento

### Sincronizar nuevas bitácoras

Hacer click en **"🔄 Sincronizar"** en la barra de navegación para buscar archivos nuevos en OneDrive. Solo agrega los nuevos — no toca el estado de los ya revisados.

---

## Historial y persistencia

El estado de cada bitácora se guarda en dos lugares:

| Dónde | Qué | Para qué |
|-------|-----|----------|
| `data/autouncle.db` | SQLite local | Consultas rápidas mientras la app corre |
| `TRACKING_BITACORAS_2026.xlsx` en OneDrive | Excel en la nube | Persistencia — se restaura al arrancar |

Si el instructor cambia de computador:
1. Clonar/copiar el proyecto
2. Repetir el setup (steps 1-4)
3. Arrancar la app — restaura automáticamente el historial desde OneDrive

---

## Renovar sesión (cada ~90 días)

Cuando la sesión expira, la app lo detecta y loguea un warning. Para renovar:

```bash
rm config/browser_session/state.json
python backend/auth.py
```

Se abre el navegador para hacer login con MFA nuevamente.

---

## Dónde quedan los archivos en SharePoint

Los PDF firmados se suben a:

```
ETAPASPRODUCTIVAS2025SHP/
  DOCUMENTOS CERTIFICACIONES 2025/
    DOCUMENTOS INSTRUCTORES 2025/
      ADMINISTRACION Y CONTABILIDAD LUIS EDUARDO GONZALEZ/
        1. RED ADMINISTRACION/    ← aprendices de área Administración
        2. RED CONTABLE/          ← aprendices de área Contabilidad
```
