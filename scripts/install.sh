#!/usr/bin/env sh
set -eu

REPO="${RUSDOX_REPO:-OthmaneBlial/rusdox}"
VERSION="${RUSDOX_VERSION:-latest}"
DOWNLOAD_BASE="${RUSDOX_DOWNLOAD_BASE:-https://github.com/$REPO/releases}"

if [ -n "${RUSDOX_INSTALL_DIR:-}" ]; then
  INSTALL_DIR="$RUSDOX_INSTALL_DIR"
elif [ -w /usr/local/bin ]; then
  INSTALL_DIR="/usr/local/bin"
else
  INSTALL_DIR="$HOME/.local/bin"
fi

OS="$(uname -s)"
ARCH="$(uname -m)"

case "$OS" in
  Linux)
    OS_TARGET="unknown-linux-gnu"
    case "$ARCH" in
      x86_64|amd64)
        ARCH_TARGET="x86_64"
        ;;
      *)
        echo "Unsupported Linux architecture: $ARCH (supported: x86_64)"
        exit 1
        ;;
    esac
    ;;
  Darwin)
    OS_TARGET="apple-darwin"
    case "$ARCH" in
      x86_64|amd64)
        ARCH_TARGET="x86_64"
        ;;
      arm64|aarch64)
        ARCH_TARGET="aarch64"
        ;;
      *)
        echo "Unsupported macOS architecture: $ARCH (supported: x86_64, arm64)"
        exit 1
        ;;
    esac
    ;;
  *)
    echo "Unsupported OS: $OS"
    exit 1
    ;;
esac

TARGET="$ARCH_TARGET-$OS_TARGET"
ASSET="rusdox-$TARGET.tar.gz"

if [ "$VERSION" = "latest" ]; then
  RELEASE_BASE="$DOWNLOAD_BASE/latest/download"
else
  RELEASE_BASE="$DOWNLOAD_BASE/download/$VERSION"
fi
URL="$RELEASE_BASE/$ASSET"
CHECKSUMS_URL="$RELEASE_BASE/SHA256SUMS"

TMP_DIR="$(mktemp -d)"
cleanup() {
  rm -rf "$TMP_DIR"
}
trap cleanup EXIT INT TERM

ARCHIVE_PATH="$TMP_DIR/$ASSET"
CHECKSUMS_PATH="$TMP_DIR/SHA256SUMS"

echo "Downloading $URL"
if command -v curl >/dev/null 2>&1; then
  curl -fsSL "$URL" -o "$ARCHIVE_PATH"
  curl -fsSL "$CHECKSUMS_URL" -o "$CHECKSUMS_PATH"
elif command -v wget >/dev/null 2>&1; then
  wget -qO "$ARCHIVE_PATH" "$URL"
  wget -qO "$CHECKSUMS_PATH" "$CHECKSUMS_URL"
else
  echo "Need curl or wget to download release binaries."
  exit 1
fi

EXPECTED_SHA256="$(awk -v asset="$ASSET" '$2 == asset || $2 == "*" asset { print $1; exit }' "$CHECKSUMS_PATH")"
if [ -z "$EXPECTED_SHA256" ]; then
  echo "Checksum for $ASSET was not found in SHA256SUMS."
  exit 1
fi

if command -v sha256sum >/dev/null 2>&1; then
  ACTUAL_SHA256="$(sha256sum "$ARCHIVE_PATH" | awk '{print $1}')"
elif command -v shasum >/dev/null 2>&1; then
  ACTUAL_SHA256="$(shasum -a 256 "$ARCHIVE_PATH" | awk '{print $1}')"
else
  echo "Need sha256sum or shasum to verify the release archive."
  exit 1
fi

if [ "$ACTUAL_SHA256" != "$EXPECTED_SHA256" ]; then
  echo "Checksum verification failed for $ASSET."
  exit 1
fi

echo "Verified SHA-256 for $ASSET"

tar -xzf "$ARCHIVE_PATH" -C "$TMP_DIR"

mkdir -p "$INSTALL_DIR"
RUSDOX_BIN="$INSTALL_DIR/rusdox"
install -m 755 "$TMP_DIR/rusdox" "$RUSDOX_BIN"

CONFIG_PATH="$("$RUSDOX_BIN" config path)"
CONFIG_CREATED="false"
if [ ! -f "$CONFIG_PATH" ]; then
  "$RUSDOX_BIN" config init --template >/dev/null
  CONFIG_CREATED="true"
fi

echo "Installed rusdox to $RUSDOX_BIN"
echo "User config: $CONFIG_PATH"
if [ "$CONFIG_CREATED" = "true" ]; then
  echo "Created default config at $CONFIG_PATH"
fi
echo "Customize styling with:"
echo "  rusdox config wizard --level basic"
echo "  rusdox config wizard --level advanced"
echo "Create a project-local override with:"
echo "  rusdox config wizard --path ./rusdox.toml --level basic"
case ":$PATH:" in
  *":$INSTALL_DIR:"*)
    ;;
  *)
    echo "Add this directory to PATH:"
    echo "  export PATH=\"$INSTALL_DIR:\$PATH\""
    ;;
esac
