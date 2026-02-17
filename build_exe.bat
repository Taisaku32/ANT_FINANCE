@echo off
:: ============================================================
::  build_exe.bat  —  Compila Finanzas Personales como .exe
::  Doble clic para ejecutar. Requiere PyInstaller instalado.
:: ============================================================

echo.
echo  === Instalando PyInstaller si no esta instalado ===
C:\Users\juanm\anaconda3\envs\finanzas\python.exe -m pip install pyinstaller --quiet

echo.
echo  === Compilando la aplicacion (puede tardar 2-5 minutos) ===
C:\Users\juanm\anaconda3\envs\finanzas\python.exe -m PyInstaller finanzas.spec --clean --noconfirm

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo  [ERROR] La compilacion fallo. Revisa los mensajes de arriba.
    pause
    exit /b 1
)

echo.
echo  ============================================================
echo   COMPILACION EXITOSA
echo   La aplicacion esta en:
echo   dist\FinanzasPersonales\FinanzasPersonales.exe
echo.
echo   Para distribuir:  copia la carpeta completa
echo                     dist\FinanzasPersonales\
echo   Para acceso rapido: crea un acceso directo de
echo   FinanzasPersonales.exe en el Escritorio.
echo  ============================================================
echo.
pause
