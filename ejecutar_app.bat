@echo off
call "%USERPROFILE%\anaconda3\Scripts\activate.bat" "%USERPROFILE%\anaconda3"
call conda activate finanzas
cd /d "C:\Users\juanm\OneDrive\Escritorio\APPS\FINANZAS STREAMLIT v2"
python main.py
pause
