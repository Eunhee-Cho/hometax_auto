#!/usr/bin/env bash
set -euo pipefail
cd /workspace/hometax_auto
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements-dev.txt
.venv/bin/python -m pip check
.venv/bin/python hometax_vat_downloader.py --input 'INPUT-부가세_빈양식.xlsx' --validate-only
