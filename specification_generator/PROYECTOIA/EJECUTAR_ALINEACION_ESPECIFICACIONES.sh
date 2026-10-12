#!/usr/bin/env bash
set -euo pipefail

cd -- "$(dirname -- "${BASH_SOURCE[0]}")"

python_venv=".venv/bin/python"
if [[ ! -x "$python_venv" ]]; then
    if ! command -v python3.11 >/dev/null 2>&1; then
        echo "Error: se requiere Python 3.11 para crear .venv." >&2
        exit 1
    fi
    python3.11 -m venv .venv
fi

"$python_venv" -c 'import openpyxl, pandas, rapidfuzz' >/dev/null 2>&1 || {
    echo "Instalando dependencias necesarias para la alineacion..."
    "$python_venv" -m pip install "openpyxl==3.1.5" "pandas>=2.2,<3" "rapidfuzz>=3.9,<4"
}

exec "$python_venv" alinear_especificaciones_json.py --base "."
