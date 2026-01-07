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
echo "Starting proxy server in background..."
npm run proxy &
PROXY_PID=$!

sleep 2

echo "Starting add-in debugging..."

DEFAULT_DOC="$HOME/Downloads/Default.docx"

if [ ! -f "$DEFAULT_DOC" ]; then
  echo "Creating default document: $DEFAULT_DOC"
  touch "$DEFAULT_DOC"
fi

cd "$SCRIPT_DIR"

npx office-addin-debugging start manifest.xml &
DEBUG_PID=$!

sleep 3

if [ -f "$DEFAULT_DOC" ]; then
  echo "Opening default document: $DEFAULT_DOC"
  open -a "Microsoft Word" "$DEFAULT_DOC" 2>/dev/null || echo "Note: Word should open automatically with the add-in"
fi

wait $DEBUG_PID

echo "Stopping proxy server..."
kill $PROXY_PID 2>/dev/null

