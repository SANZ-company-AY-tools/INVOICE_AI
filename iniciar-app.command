#!/bin/bash
cd "$(dirname "$0")"

echo "================================================"
echo "   🧾 Extractor de Facturas con IA"
echo "================================================"
echo ""

# Check and install poppler if needed
if ! command -v pdftoppm &> /dev/null; then
    echo "📦 Instalando poppler (necesario para PDFs)..."
    brew install poppler
fi

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo "📦 Creando entorno virtual..."
    python3 -m venv venv
fi

# Activate virtual environment
source venv/bin/activate

echo "📦 Instalando dependencias..."
pip install -r requirements.txt --quiet

echo "✅ Todo listo"
echo ""
echo "🚀 Arrancando servidor..."
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "   Abre en tu navegador: http://127.0.0.1:5001"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

# Open browser automatically
open "http://127.0.0.1:5001"

python app.py
