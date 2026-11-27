#!/usr/bin/env bash
# Automail Desk - 前台启动（macOS / Linux）
# 首次运行会创建 .venv 并安装依赖，之后直接启动服务，日志打印在当前终端。
set -euo pipefail
cd "$(dirname "$0")"

PY="${PYTHON:-python3}"
if ! command -v "$PY" >/dev/null 2>&1; then
  echo "[ERROR] 未找到 python3，请先安装 Python 3.9+（https://www.python.org/downloads/）" >&2
  exit 1
fi

if [ ! -d .venv ]; then
  echo "[i] 创建虚拟环境 .venv ..."
  "$PY" -m venv .venv
fi
# shellcheck disable=SC1091
. .venv/bin/activate
python -m pip install -q --upgrade pip
python -m pip install -q -r requirements.txt

if [ ! -f config.json ]; then
  cp config.example.json config.json
  echo "[i] 已生成 config.json，请填写域名、密码与账号池后重新启动。"
fi

exec python server.py
