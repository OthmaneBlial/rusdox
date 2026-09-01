#!/usr/bin/env sh
set -eu

BINARY_PATH="${1:-target/release/rusdox}"
PACKAGE_VERSION="$(sed -n 's/^version = "\([^"]*\)"/\1/p' Cargo.toml | head -n 1)"
VERSION="v$PACKAGE_VERSION"
PORT="${RUSDOX_TEST_PORT:-18731}"

if [ ! -x "$BINARY_PATH" ]; then
  echo "Missing executable RusDox binary: $BINARY_PATH"
  exit 1
fi

case "$(uname -s)" in
  Linux) TARGET="x86_64-unknown-linux-gnu" ;;
  Darwin)
    case "$(uname -m)" in
      arm64|aarch64) TARGET="aarch64-apple-darwin" ;;
      *) TARGET="x86_64-apple-darwin" ;;
    esac
    ;;
  *)
    echo "Installer fixture test supports Linux and macOS."
    exit 1
    ;;
esac

TEST_ROOT="$(mktemp -d)"
SERVER_PID=""
cleanup() {
  if [ -n "$SERVER_PID" ]; then
    kill "$SERVER_PID" 2>/dev/null || true
    wait "$SERVER_PID" 2>/dev/null || true
  fi
  rm -rf "$TEST_ROOT"
}
trap cleanup EXIT INT TERM

RELEASE_DIR="$TEST_ROOT/releases/download/$VERSION"
ASSET="rusdox-$TARGET.tar.gz"
mkdir -p "$RELEASE_DIR/archive" "$TEST_ROOT/home" "$TEST_ROOT/bin"
cp "$BINARY_PATH" "$RELEASE_DIR/archive/rusdox"
tar -czf "$RELEASE_DIR/$ASSET" -C "$RELEASE_DIR/archive" rusdox

if command -v sha256sum >/dev/null 2>&1; then
  (cd "$RELEASE_DIR" && sha256sum "$ASSET" > SHA256SUMS)
else
  (cd "$RELEASE_DIR" && shasum -a 256 "$ASSET" > SHA256SUMS)
fi

python3 -m http.server "$PORT" --bind 127.0.0.1 --directory "$TEST_ROOT/releases" \
  >"$TEST_ROOT/server.log" 2>&1 &
SERVER_PID=$!

READY="false"
attempt=0
while [ "$attempt" -lt 30 ]; do
  if curl -fsS "http://127.0.0.1:$PORT/" >/dev/null 2>&1; then
    READY="true"
    break
  fi
  attempt=$((attempt + 1))
  sleep 0.1
done

if [ "$READY" != "true" ]; then
  echo "Local release fixture did not start."
  exit 1
fi

HOME="$TEST_ROOT/home" \
RUSDOX_VERSION="$VERSION" \
RUSDOX_DOWNLOAD_BASE="http://127.0.0.1:$PORT" \
RUSDOX_INSTALL_DIR="$TEST_ROOT/bin" \
  sh scripts/install.sh

HOME="$TEST_ROOT/home" "$TEST_ROOT/bin/rusdox" --version | grep "$PACKAGE_VERSION"
CONFIG_PATH="$(HOME="$TEST_ROOT/home" "$TEST_ROOT/bin/rusdox" config path)"
if [ -e "$CONFIG_PATH" ]; then
  echo "Installer created an unexpected config at $CONFIG_PATH"
  exit 1
fi
echo "Unix installer fixture passed for $TARGET."
