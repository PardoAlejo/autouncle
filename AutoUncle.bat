@echo off
setlocal EnableDelayedExpansion
title AutoUncle
cd /d "%~dp0"

set ENV_NAME=autouncle
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

:: ─── Buscar conda en ubicaciones comunes ─────────────────────────────────────
set CONDA_EXE=
for %%P in (
    "%USERPROFILE%\miniconda3\Scripts\conda.exe"
    "%USERPROFILE%\anaconda3\Scripts\conda.exe"
    "%LOCALAPPDATA%\miniconda3\Scripts\conda.exe"
    "C:\ProgramData\miniconda3\Scripts\conda.exe"
    "C:\ProgramData\anaconda3\Scripts\conda.exe"
    "C:\miniconda3\Scripts\conda.exe"
) do if exist %%P if "!CONDA_EXE!"=="" set CONDA_EXE=%%~P

:: Si no esta, descargar e instalar Miniconda silenciosamente
if "!CONDA_EXE!"=="" (
    echo  [1/3] Instalando Miniconda ^(solo la primera vez^)...
    powershell -NoProfile -Command ^
        "Invoke-WebRequest 'https://repo.anaconda.com/miniconda/Miniconda3-latest-Windows-x86_64.exe' -OutFile '%TEMP%\miniconda_setup.exe'"
    if errorlevel 1 (
        echo  ERROR: No se pudo descargar Miniconda. Verifica tu conexion a internet.
        pause
        exit /b 1
    )
    "%TEMP%\miniconda_setup.exe" /InstallationType=JustMe /AddToPath=0 /RegisterPython=0 /S /D="%USERPROFILE%\miniconda3"
    if errorlevel 1 (
        echo  ERROR: No se pudo instalar Miniconda.
        pause
        exit /b 1
    )
    set CONDA_EXE=%USERPROFILE%\miniconda3\Scripts\conda.exe
    del "%TEMP%\miniconda_setup.exe" 2>nul
    echo  Miniconda instalado.
) else (
    echo  [1/3] Conda encontrado.
)

:: Verificar que el ejecutable existe
if not exist "!CONDA_EXE!" (
    echo  ERROR: No se encontro conda en: !CONDA_EXE!
    pause
    exit /b 1
)

:: ─── Crear entorno si no existe ──────────────────────────────────────────────
"!CONDA_EXE!" env list 2>nul | findstr /B "%ENV_NAME% " >nul
if errorlevel 1 (
    echo  [2/3] Creando entorno Python ^(puede tardar unos minutos^)...
    "!CONDA_EXE!" create -n %ENV_NAME% python=3.11 -y
    if errorlevel 1 (
        echo  ERROR: No se pudo crear el entorno Python.
        pause
        exit /b 1
    )
    echo  Entorno creado.
    if exist "%SENTINEL%" del "%SENTINEL%"
    if exist "%REQ_HASH_FILE%" del "%REQ_HASH_FILE%"
) else (
    echo  [2/3] Entorno Python listo.
)

:: ─── Instalar/actualizar dependencias si requirements.txt cambio ─────────────
for /f "skip=1 delims=" %%H in (
    'certutil -hashfile requirements.txt MD5 2^>nul'
) do if "!CURRENT_HASH!"=="" set CURRENT_HASH=%%H

set STORED_HASH=
if exist "%REQ_HASH_FILE%" set /p STORED_HASH=<"%REQ_HASH_FILE%"
if not exist "%SENTINEL%" set STORED_HASH=NONE

if /i "!CURRENT_HASH!" neq "!STORED_HASH!" (
    echo  [3/3] Instalando dependencias ^(solo cuando hay cambios^)...
    "!CONDA_EXE!" run -n %ENV_NAME% pip install -r requirements.txt
    if errorlevel 1 (
        echo  ERROR instalando dependencias.
        pause
        exit /b 1
    )
    if not exist "%SENTINEL%" (
        echo  Instalando navegador Chromium...
        "!CONDA_EXE!" run -n %ENV_NAME% playwright install chromium
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
"!CONDA_EXE!" run -n %ENV_NAME% python run.py

echo.
echo  El servidor se detuvo. Presiona cualquier tecla para cerrar.
pause >nul
