#!/bin/bash

# WordTrack Startup Script
# 
# This script starts both the proxy server and the add-in debugging session.
# 
# Usage Options:
# 
# Option 1: Use this script (single terminal)
#   ./start.sh
#   - Starts proxy in background, then starts add-in
#   - Press Ctrl+C to stop both
# 
# Option 2: Manual (two terminals - more control)
#   Terminal 1: npm run proxy
#   Terminal 2: npx office-addin-debugging start manifest.xml
#   - Allows restarting add-in without affecting proxy
#   - Better for development/debugging

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo "Starting WordTrack from: $SCRIPT_DIR"

DEFAULT_DOC="$HOME/Downloads/Default.docx"

if [ ! -f "$DEFAULT_DOC" ]; then
  echo "Creating default document: $DEFAULT_DOC"
  if command -v osascript >/dev/null 2>&1; then
    osascript <<EOF 2>/dev/null
tell application "Microsoft Word"
  set newDoc to make new document
  set savePath to POSIX file "$DEFAULT_DOC"
  save active document in savePath
  close active document
end tell
EOF
    if [ ! -f "$DEFAULT_DOC" ]; then
      echo "Could not create Word document automatically. Please create Default.docx in ~/Downloads manually."
      DEFAULT_DOC=""
    fi
  else
    echo "Note: Please create Default.docx manually in ~/Downloads"
    DEFAULT_DOC=""
  fi
fi

if [ -n "$DEFAULT_DOC" ] && [ -f "$DEFAULT_DOC" ]; then
  DEFAULT_DOC_ABS=$(cd "$(dirname "$DEFAULT_DOC")" && pwd)/$(basename "$DEFAULT_DOC")
  DOC_ARG="--document \"$DEFAULT_DOC_ABS\""
  echo "Using document: $DEFAULT_DOC_ABS"
else
  DOC_ARG=""
  echo "Using temporary document (Default.docx not found)"
fi

echo "Checking if proxy server is already running..."
if lsof -Pi :3001 -sTCP:LISTEN -t >/dev/null 2>&1 ; then
  echo "Proxy server already running on port 3001, skipping..."
  PROXY_RUNNING=true
else
  echo "Starting proxy server in background..."
  npm run proxy > /dev/null 2>&1 &
  PROXY_PID=$!
  PROXY_RUNNING=false
  sleep 2
fi

echo "Starting add-in debugging..."
cd "$SCRIPT_DIR"

if [ -n "$DEFAULT_DOC" ] && [ -f "$DEFAULT_DOC" ]; then
  echo "Opening Default.docx first..."
  open -a "Microsoft Word" "$DEFAULT_DOC" 2>/dev/null
  sleep 3
fi

echo "Starting add-in (will sideload into open Word document)..."
npx office-addin-debugging start manifest.xml --no-sideload 2>/dev/null || npx office-addin-debugging start manifest.xml

if [ "$PROXY_RUNNING" = false ]; then
  echo "Stopping proxy server..."
  kill $PROXY_PID 2>/dev/null
fi

