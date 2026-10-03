@echo off
chcp 65001 >nul
cd /d "%~dp0"

if "%ANTHROPIC_API_KEY%"=="" (
  set /p ANTHROPIC_API_KEY=Pegue su clave de la API de Anthropic ^(sk-ant-...^): 
)

if not exist .venv (
  echo Creando entorno e instalando dependencias...
  py -m venv .venv || goto :error
  call .venv\Scripts\activate.bat
  pip install -r requirements.txt || goto :error
) else (
  call .venv\Scripts\activate.bat
)

start "" http://localhost:8000
python -m uvicorn webapp.main:app --host 127.0.0.1 --port 8000
goto :eof

:error
echo.
echo Algo ha fallado. Compruebe que Python 3.10+ esta instalado ^(https://www.python.org/downloads/^).
pause
