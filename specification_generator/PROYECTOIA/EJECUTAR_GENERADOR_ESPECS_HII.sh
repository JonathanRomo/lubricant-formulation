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

if ! "$python_venv" -c 'import docx' >/dev/null 2>&1; then
    "$python_venv" -m pip install -r requirements.txt
fi

exec "$python_venv" generador_especs_hii.py "$@"
