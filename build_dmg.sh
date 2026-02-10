#!/usr/bin/env bash
set -euo pipefail

# Build a macOS .dmg using Tauri bundler.
# Usage: ./build_dmg.sh [--release|--debug]

MODE="release"
if [[ ${1:-} == "--debug" ]]; then
  MODE="debug"
fi

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "Error: DMG builds require macOS (Darwin)." >&2
  exit 1
fi

# Ensure bun is available (used for frontend build and can run tauri CLI)
command -v bun >/dev/null 2>&1 || { echo "Error: bun not found in PATH." >&2; exit 1; }

# Prefer cargo-tauri if installed, otherwise fall back to bunx tauri
TAURI_CMD=""
if command -v cargo >/dev/null 2>&1 && cargo tauri --version >/dev/null 2>&1; then
  TAURI_CMD="cargo tauri"
else
  TAURI_CMD="bunx tauri"
fi

# Build frontend
bun run build

# Build Tauri bundle
if [[ "$MODE" == "release" ]]; then
  $TAURI_CMD build --bundles dmg
else
  $TAURI_CMD build --bundles dmg --debug
fi

# Show output path(s)
DMG_DIR="src-tauri/target/${MODE}/bundle/dmg"
if [[ -d "$DMG_DIR" ]]; then
  echo "DMG(s) created in: $DMG_DIR"
  ls -1 "$DMG_DIR"/*.dmg 2>/dev/null || true
else
  echo "Warning: expected DMG output dir not found: $DMG_DIR" >&2
fi
