#!/bin/bash
# Mac 打包脚本：生成 dist/会计工作台.app 和 .dmg（需要 Python 3.12）
set -e
cd "$(dirname "$0")"
PY=${PYTHON:-python3}
[ -d .buildenv ] || "$PY" -m venv .buildenv
.buildenv/bin/python -m pip install --upgrade pip -q
.buildenv/bin/python -m pip install -r requirements-build.txt
PYTHONUTF8=1 .buildenv/bin/python build.py
open dist
