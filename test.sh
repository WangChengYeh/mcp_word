#!/usr/bin/env bash
set -euo pipefail

# Integration test runner for MCP Word server per SPEC.md and tool.md
# - Launches test.js which spawns the MCP server via stdio and connects a Socket.IO client

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT_DIR"

: "${PORT:=3100}"
KEY_FILE="${KEY_FILE:-$ROOT_DIR/key.pem}"
CERT_FILE="${CERT_FILE:-$ROOT_DIR/cert.pem}"

if [[ ! -f "$KEY_FILE" || ! -f "$CERT_FILE" ]]; then
  echo "Missing TLS materials. Expecting key.pem and cert.pem at project root or set KEY_FILE/CERT_FILE."
  exit 1
fi

export PORT KEY_FILE CERT_FILE
node test.js

