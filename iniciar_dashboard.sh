#!/usr/bin/env bash
cd "$(dirname "$0")"

echo "============================================================"
echo "  Iniciando Relatório de Vendas por Hora (Streamlit)"
echo "============================================================"

if ! command -v python3 &> /dev/null; then
    echo "[ERRO] Python 3 não encontrado. Instale via python.org ou Homebrew."
    exit 1
fi

if [ ! -d ".venv" ]; then
    echo "[1/3] Criando ambiente virtual..."
    python3 -m venv .venv
fi

source .venv/bin/activate
echo "[2/3] Instalando dependências..."
pip install --upgrade pip -q
pip install -r requirements.txt -q

echo "[3/3] Abrindo dashboard no navegador..."
streamlit run app.py
