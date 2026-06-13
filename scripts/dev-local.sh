#!/usr/bin/env zsh
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
CLI="$ROOT_DIR/packages/cli/dist/index.js"
START_FILE_BRIDGE=1
SIDELOAD_MAC=0

usage() {
  cat <<'EOF'
Usage: corepack pnpm dev:local [-- --no-file-bridge] [-- --sideload-mac]

Starts the local Open Workbook development runtime:
  - shared daemon on http://127.0.0.1:37845
  - add-in asset server on http://127.0.0.1:37846
  - native file bridge on http://127.0.0.1:37847, unless disabled

OpenCode should launch the MCP stdio adapter from its config:
  node packages/cli/dist/index.js mcp --agent-name open-workbook

Options:
  --no-file-bridge  Do not start the optional native file bridge.
  --sideload-mac    Copy the generated manifest into Excel's macOS sideload folder before starting.
  -h, --help        Show this help.
EOF
}

for arg in "$@"; do
  case "$arg" in
    --)
      ;;
    --no-file-bridge)
      START_FILE_BRIDGE=0
      ;;
    --sideload-mac)
      SIDELOAD_MAC=1
      ;;
    -h|--help)
      usage
      exit 0
      ;;
    *)
      echo "Unknown option: $arg" >&2
      usage >&2
      exit 1
      ;;
  esac
done

if [[ ! -f "$CLI" ]]; then
  echo "Missing built CLI at $CLI" >&2
  echo "Run: corepack pnpm build" >&2
  exit 1
fi

if [[ "$SIDELOAD_MAC" -eq 1 ]]; then
  node "$CLI" sideload mac
fi

pids=()

cleanup() {
  if [[ "${#pids[@]}" -gt 0 ]]; then
    echo ""
    echo "Stopping Open Workbook development runtime..."
    kill "${pids[@]}" 2>/dev/null || true
  fi
}

trap cleanup EXIT INT TERM

start_process() {
  local label="$1"
  shift
  echo "Starting $label..."
  "$@" &
  pids+=("$!")
}

echo "Open Workbook development runtime"
echo "Repo: $ROOT_DIR"
echo ""

start_process "daemon" node "$CLI" daemon start

if [[ "$START_FILE_BRIDGE" -eq 1 ]]; then
  start_process "file bridge" node "$CLI" file-bridge start
fi

start_process "add-in server" node "$CLI" addin serve

echo ""
echo "Runtime URLs:"
echo "  daemon:      http://127.0.0.1:37845"
echo "  add-in:      http://127.0.0.1:37846/taskpane.html"
echo "  file bridge: http://127.0.0.1:37847"
echo ""
echo "OpenCode MCP command:"
echo "  node $CLI mcp --agent-name open-workbook"
echo ""
echo "Keep this process running while Excel/OpenCode are using Open Workbook."
echo "Press Ctrl-C to stop all child processes."
echo ""

wait
