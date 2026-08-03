#!/usr/bin/env sh
set -eu
cd "$(dirname "$0")"
exec python3 -m http.server 5173 --directory frontend
