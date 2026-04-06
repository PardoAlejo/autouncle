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

:: Tambien intentar desde PATH
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

:: ─── Buscar conda en ubicaciones comunes ────────────────────────────────────
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
if "%CONDA_EXE%"=="" (
    echo  [1/3] Instalando Miniconda ^(solo la primera vez^)...
    powershell -NoProfile -Command ^
        "Invoke-WebRequest 'https://repo.anaconda.com/miniconda/Miniconda3-latest-Windows-x86_64.exe' -OutFile '%TEMP%\miniconda_setup.exe'"
    if errorlevel 1 (
        echo  ERROR: No se pudo descargar Miniconda. Verifica tu conexion a internet.
        pause
        exit /b 1
    )
    "%TEMP%\miniconda_setup.exe" /InstallationType=JustMe /AddToPath=0 /RegisterPython=0 /S /D="%USERPROFILE%\miniconda3"
    set CONDA_EXE=%USERPROFILE%\miniconda3\Scripts\conda.exe
    del "%TEMP%\miniconda_setup.exe" 2>nul
    echo  Miniconda instalado.
) else (
    echo  [1/3] Conda encontrado.
)

:: Activar conda para esta sesion
set CONDA_DIR=!CONDA_EXE:\Scripts\conda.exe=!
call "!CONDA_DIR!\Scripts\activate.bat" 2>nul

:: ─── Crear entorno si no existe ─────────────────────────────────────────────
conda env list 2>nul | findstr /B "%ENV_NAME% " >nul
if errorlevel 1 (
    echo  [2/3] Creando entorno Python ^(puede tardar un minuto^)...
    conda create -n %ENV_NAME% python=3.11 -y >nul 2>&1
    echo  Entorno creado.
    :: Forzar instalacion de dependencias al crear env por primera vez
    if exist "%SENTINEL%" del "%SENTINEL%"
) else (
    echo  [2/3] Entorno Python listo.
)

:: ─── Instalar/actualizar dependencias si requirements.txt cambio ────────────
:: Calcular hash actual de requirements.txt
for /f "skip=1 delims=" %%H in (
    'certutil -hashfile requirements.txt MD5 2^>nul'
) do if "!CURRENT_HASH!"=="" set CURRENT_HASH=%%H

set STORED_HASH=
if exist "%REQ_HASH_FILE%" set /p STORED_HASH=<"%REQ_HASH_FILE%"

if not exist "%SENTINEL%" set STORED_HASH=NONE

if /i "!CURRENT_HASH!" neq "!STORED_HASH!" (
    echo  [3/3] Instalando dependencias...
    conda run -n %ENV_NAME% pip install -r requirements.txt >nul 2>&1
    if errorlevel 1 (
        echo  ERROR instalando dependencias.
        pause
        exit /b 1
    )
    if not exist "%SENTINEL%" (
        conda run -n %ENV_NAME% playwright install chromium >nul 2>&1
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
conda run -n %ENV_NAME% python run.py

echo.
echo  El servidor se detuvo. Presiona cualquier tecla para cerrar.
pause >nul
