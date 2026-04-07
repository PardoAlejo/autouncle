@echo off
setlocal EnableDelayedExpansion
title AutoUncle
cd /d "%~dp0"

set VENV_DIR=.venv
set SENTINEL=.setup_done
set REQ_HASH_FILE=.req_hash

echo.
echo  ================================================
echo   AutoUncle — Gestor de Bitacoras
echo  ================================================
echo.

:: ─── Git pull (actualizar app) ───────────────────────────────────────────────
set GIT_EXE=
for %%P in (
    "C:\Program Files\Git\bin\git.exe"
    "C:\Program Files (x86)\Git\bin\git.exe"
    "%LOCALAPPDATA%\Programs\Git\bin\git.exe"
) do if exist %%P if "!GIT_EXE!"=="" set GIT_EXE=%%~P

if "%GIT_EXE%"=="" (
    where git >nul 2>&1 && set GIT_EXE=git
)

if not "%GIT_EXE%"=="" (
    if exist ".git" (
        echo  Actualizando app...
        "!GIT_EXE!" pull --quiet
        if errorlevel 1 (
            echo  Aviso: no se pudo actualizar ^(sin conexion o sin cambios^). Continuando...
        )
    )
) else (
    echo  Git no encontrado, se omite la actualizacion.
)

:: ─── Buscar Python 3 en ubicaciones comunes ──────────────────────────────────
set PYTHON_EXE=
for %%P in (
    "%USERPROFILE%\miniconda3\python.exe"
    "%USERPROFILE%\anaconda3\python.exe"
    "%LOCALAPPDATA%\miniconda3\python.exe"
    "C:\ProgramData\miniconda3\python.exe"
    "C:\ProgramData\anaconda3\python.exe"
    "C:\miniconda3\python.exe"
    "C:\Python311\python.exe"
    "C:\Python310\python.exe"
    "C:\Python39\python.exe"
) do if exist %%P if "!PYTHON_EXE!"=="" set PYTHON_EXE=%%~P

if "!PYTHON_EXE!"=="" (
    where python >nul 2>&1 && set PYTHON_EXE=python
)

if "!PYTHON_EXE!"=="" (
    echo  ERROR: No se encontro Python. Instala Miniconda o Python 3 primero.
    pause
    exit /b 1
)
echo  [1/3] Python encontrado: !PYTHON_EXE!

:: ─── Crear venv si no existe ─────────────────────────────────────────────────
if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo  [2/3] Creando entorno virtual...
    "!PYTHON_EXE!" -m venv %VENV_DIR%
    if errorlevel 1 (
        echo  ERROR: No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
    echo  Entorno creado.
    if exist "%SENTINEL%" del "%SENTINEL%"
    if exist "%REQ_HASH_FILE%" del "%REQ_HASH_FILE%"
) else (
    echo  [2/3] Entorno virtual listo.
)

set VENV_PYTHON=%VENV_DIR%\Scripts\python.exe
set VENV_PIP=%VENV_DIR%\Scripts\pip.exe

:: ─── Instalar/actualizar dependencias si requirements.txt cambio ─────────────
for /f "skip=1 delims=" %%H in (
    'certutil -hashfile requirements.txt MD5 2^>nul'
) do if "!CURRENT_HASH!"=="" set CURRENT_HASH=%%H

set STORED_HASH=
if exist "%REQ_HASH_FILE%" set /p STORED_HASH=<"%REQ_HASH_FILE%"
if not exist "%SENTINEL%" set STORED_HASH=NONE

if /i "!CURRENT_HASH!" neq "!STORED_HASH!" (
    echo  [3/3] Instalando dependencias ^(puede tardar unos minutos^)...
    "!VENV_PIP!" install -r requirements.txt
    if errorlevel 1 (
        echo  ERROR instalando dependencias.
        pause
        exit /b 1
    )
    if not exist "%SENTINEL%" (
        echo  Instalando navegador Chromium...
        "!VENV_PYTHON!" -m playwright install chromium
        if errorlevel 1 (
            echo  ERROR instalando Chromium.
            pause
            exit /b 1
        )
    )
    echo !CURRENT_HASH!>"%REQ_HASH_FILE%"
    echo done>"%SENTINEL%"
    echo  Dependencias actualizadas.
) else (
    echo  [3/3] Dependencias al dia.
)

echo.
echo  Iniciando AutoUncle...
echo  ^(Esta ventana debe quedar abierta mientras usas la app^)
echo.

:: Abrir el navegador despues de 4 segundos (en segundo plano)
start "" cmd /c "timeout /t 4 /nobreak >nul & start http://localhost:8000"

:: Arrancar servidor en primer plano (cerrar esta ventana detiene el servidor)
"!VENV_PYTHON!" run.py

echo.
echo  El servidor se detuvo. Presiona cualquier tecla para cerrar.
pause >nul
