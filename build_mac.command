#!/bin/zsh
set -euo pipefail

cd "$(dirname "$0")"

APP_NAME="Teq Shipment Docs Generator"
PYTHON_BIN="${PYTHON_BIN:-python3}"

if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
  echo "python3 is required. Install Python 3 first, then run this file again."
  exit 1
fi

if ! xcode-select -p >/dev/null 2>&1 || ! command -v lipo >/dev/null 2>&1; then
  echo "Apple Command Line Developer Tools are required to build the Mac app."
  echo "Run this command, finish the Apple installer, then run this build file again:"
  echo ""
  echo "  xcode-select --install"
  echo ""
  exit 1
fi

if ! command -v pkgbuild >/dev/null 2>&1; then
  echo "pkgbuild is required to create the .pkg installer."
  echo "Install Apple Command Line Developer Tools, then run this build file again:"
  echo ""
  echo "  xcode-select --install"
  echo ""
  exit 1
fi

"$PYTHON_BIN" -m venv .venv-mac
source .venv-mac/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements-mac.txt

rm -rf build "dist/$APP_NAME" "dist/$APP_NAME.app" "dist/$APP_NAME-mac.zip" "dist/$APP_NAME.pkg" "dist/pkg-root"
pyinstaller --clean -y pyinstaller_mac.spec

cd dist
ditto -c -k --sequesterRsrc --keepParent "$APP_NAME.app" "$APP_NAME-mac.zip"

mkdir -p "pkg-root/Applications"
ditto "$APP_NAME.app" "pkg-root/Applications/$APP_NAME.app"
pkgbuild \
  --root "pkg-root" \
  --identifier "com.teqtronix.shipment-docs-generator" \
  --version "1.0.0" \
  --install-location "/" \
  "$APP_NAME.pkg"
rm -rf "pkg-root"

echo ""
echo "Done:"
echo "$(pwd)/$APP_NAME.app"
echo "$(pwd)/$APP_NAME-mac.zip"
echo "$(pwd)/$APP_NAME.pkg"
