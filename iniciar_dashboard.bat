@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ============================================================
echo   Iniciando Relatorio de Vendas por Hora (Streamlit)
echo ============================================================
echo.

:: Verifica se Python esta instalado (python ou py)
set PY_CMD=
where python >nul 2>nul
if %errorlevel% equ 0 (
    set PY_CMD=python
) else (
    where py >nul 2>nul
    if %errorlevel% equ 0 (
        set PY_CMD=py
    )
)

if "%PY_CMD%"=="" (
    echo [ERRO] Python nao encontrado neste computador.
    echo.
    echo Para instalar rapidamente:
    echo 1. Acesse https://www.python.org/downloads/
    echo 2. Na primeira tela do instalador, marque "Add Python to PATH"
    echo 3. Clique em "Install Now" e depois execute este arquivo novamente.
    echo.
    pause
    exit /b 1
)

:: Cria ambiente virtual local na primeira execucao
if not exist ".venv" (
    echo [1/3] Criando ambiente virtual Python pela primeira vez...
    %PY_CMD% -m venv .venv
)

call .venv\Scriptsctivate.bat

echo [2/3] Verificando e instalando dependencias...
python -m pip install --upgrade pip -q
pip install -r requirements.txt -q

echo [3/3] Abrindo o dashboard no navegador...
echo (Mantenha esta janela aberta enquanto estiver usando o relatorio)
echo.
streamlit run app.py
pause
