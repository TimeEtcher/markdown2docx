#!/usr/bin/env bash
# md2word Linux 一键打包脚本
set -e
cd "$(dirname "$0")/.."

python3 -m venv .venv
./.venv/bin/pip install -q -r requirements.txt

./.venv/bin/pyinstaller --onefile --windowed --name md2word --collect-data latex2mathml run.py

echo "打包完成: dist/md2word"
