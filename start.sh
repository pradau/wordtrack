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
  echo "Creating empty Word document..."
  osascript -e 'tell application "Microsoft Word" to make new document' 2>/dev/null || \
  echo "Note: Please create Default.docx manually in ~/Downloads if needed"
fi

if [ ! -f "$DEFAULT_DOC" ]; then
  echo "Warning: Default.docx not found, will use temporary document"
  DOC_ARG=""
else
  DOC_ARG="--document \"$DEFAULT_DOC\""
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

if [ -n "$DOC_ARG" ]; then
  eval "npx office-addin-debugging start manifest.xml $DOC_ARG"
else
  npx office-addin-debugging start manifest.xml
fi

if [ "$PROXY_RUNNING" = false ]; then
  echo "Stopping proxy server..."
  kill $PROXY_PID 2>/dev/null
fi

