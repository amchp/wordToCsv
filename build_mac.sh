#!/usr/bin/env bash
# Minimal build script for macOS
# Usage:
#   chmod +x build_mac.sh
#   ./build_mac.sh

set -euo pipefail

# 1) Resolve the REAL pandoc binary (dereferences Homebrew symlink)
PANDOC_PATH=$(python3 - <<'PY'
import shutil, pathlib, sys
p = shutil.which("pandoc")
if not p:
    sys.exit("pandoc not found; install it (e.g. brew install pandoc)")
print(pathlib.Path(p).resolve())
PY
)

echo "Using pandoc at: $PANDOC_PATH"

# 2) Build a .app without a terminal window, bundling pandoc
pyinstaller main.py --name "wordToCsv" --windowed --add-binary "$PANDOC_PATH:pandoc"

echo "Done. App at: dist/wordToCsv.app"

