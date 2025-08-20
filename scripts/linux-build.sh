#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")"/.. && pwd)"
BUILD_DIR="$ROOT_DIR/build"

if [[ -n "${VCPKG_ROOT:-}" ]]; then
  cmake -S "$ROOT_DIR" -B "$BUILD_DIR" -DCMAKE_TOOLCHAIN_FILE="$VCPKG_ROOT/scripts/buildsystems/vcpkg.cmake"
else
  cmake -S "$ROOT_DIR" -B "$BUILD_DIR"
fi

cmake --build "$BUILD_DIR" -j


