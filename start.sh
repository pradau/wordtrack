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
  # Note: Keep output visible (don't redirect to /dev/null) so we can see if proxy fails to start
  # The proxy is critical for Claude API calls - if it doesn't start, the add-in will fail
  npm run proxy &
  PROXY_PID=$!
  PROXY_RUNNING=false
  # Wait longer for proxy to fully start before proceeding
  sleep 3
  # Verify proxy is actually running - this helps catch startup issues early
  if ! lsof -Pi :3001 -sTCP:LISTEN -t >/dev/null 2>&1; then
    echo "Warning: Proxy server may not have started properly"
  fi
fi

echo "Starting add-in debugging..."
cd "$SCRIPT_DIR"

echo "Starting add-in (will create temp document with WordTrack)..."
# Start add-in debugging in background so we can hide temp doc and open Default.docx
npx office-addin-debugging start manifest.xml &
ADDIN_PID=$!

# Function to hide temp document and open Default.docx
hide_temp_and_open_default() {
  if [ -n "$DEFAULT_DOC" ] && [ -f "$DEFAULT_DOC" ] && command -v osascript >/dev/null 2>&1; then
    # Wait for add-in to load and create temp document
    # Need to wait longer for add-in to fully initialize (load JS, connect APIs, etc.)
    echo "Waiting for add-in to fully initialize..."
    sleep 6
    
    DEFAULT_DOC_NAME=$(basename "$DEFAULT_DOC")
    DEFAULT_DOC_ABS_PATH="$DEFAULT_DOC_ABS"
    
    # Wait a bit more and check if temp document exists before hiding
    # This ensures the add-in has time to fully load
    sleep 2
    
    # Hide the temp document window
    osascript <<EOF 2>/dev/null
tell application "Microsoft Word"
  activate
  set defaultDocName to "$DEFAULT_DOC_NAME"
  
  -- Find and hide the temp document window (any doc that's not Default.docx)
  set allWindows to every window
  set tempWindow to missing value
  
  repeat with aWindow in allWindows
    try
      set windowDoc to document of aWindow
      set docName to name of windowDoc
      if docName is not defaultDocName then
        set tempWindow to aWindow
        exit repeat
      end if
    on error
      -- Skip windows that don't have a document
    end try
  end repeat
  
  -- Move temp window off-screen (keeps document open for add-in)
  -- Only do this if temp window was found (add-in is loaded)
  if tempWindow is not missing value then
    try
      set position of tempWindow to {-2000, -2000}
      set visible of tempWindow to false
      set collapsed of tempWindow to true
    on error
      -- If that fails, try closing it
      try
        close tempWindow
      end try
    end try
  end if
end tell
EOF
    
    # Now open Default.docx (which will come to front)
    # IMPORTANT: Use Word's internal 'open' command via AppleScript instead of 'open -a'
    # This ensures Default.docx opens in the same Word instance as the temp document,
    # preserving the add-in connection. Using 'open -a' can start a new Word instance
    # or break the proxy connection, causing "Load failed" errors with the Claude API.
    echo "Opening Default.docx..."
    osascript <<EOF 2>/dev/null
tell application "Microsoft Word"
  open POSIX file "$DEFAULT_DOC_ABS_PATH"
  activate
end tell
EOF
    sleep 2
    
    # Ensure Default.docx window is frontmost and keep temp doc hidden
    osascript <<EOF 2>/dev/null
tell application "Microsoft Word"
  activate
  set defaultDocName to "$DEFAULT_DOC_NAME"
  
  set allWindows to every window
  set defaultWindow to missing value
  set tempWindow to missing value
  
  repeat with aWindow in allWindows
    try
      set windowDoc to document of aWindow
      set docName to name of windowDoc
      if docName is defaultDocName then
        set defaultWindow to aWindow
      else
        set tempWindow to aWindow
      end if
    on error
    end try
  end repeat
  
  -- Keep temp window off-screen
  if tempWindow is not missing value then
    try
      set position of tempWindow to {-2000, -2000}
      set visible of tempWindow to false
      set collapsed of tempWindow to true
    on error
      -- If that fails, try closing it
      try
        close tempWindow
      end try
    end try
  end if
  
  -- Bring default window to front
  if defaultWindow is not missing value then
    try
      set visible of defaultWindow to true
      set miniaturized of defaultWindow to false
      set index of defaultWindow to 1
    end try
  end if
end tell
EOF
    
    echo "Default.docx should now be visible with WordTrack loaded in the background."
  fi
}

# Function to continuously monitor and keep temp window minimized
monitor_temp_window() {
  if [ -n "$DEFAULT_DOC" ] && [ -f "$DEFAULT_DOC" ] && command -v osascript >/dev/null 2>&1; then
    DEFAULT_DOC_NAME=$(basename "$DEFAULT_DOC")
    # Keep monitoring while add-in is running
    while kill -0 $ADDIN_PID 2>/dev/null; do
      sleep 1
      osascript <<EOF 2>/dev/null
tell application "Microsoft Word"
  set defaultDocName to "$DEFAULT_DOC_NAME"
  set allWindows to every window
  
  repeat with aWindow in allWindows
    try
      set windowDoc to document of aWindow
      set docName to name of windowDoc
      if docName is not defaultDocName then
        -- Keep temp window off-screen
        try
          set position of aWindow to {-2000, -2000}
          set visible of aWindow to false
          set collapsed of aWindow to true
        on error
          -- If that fails, try closing it
          try
            close aWindow
          end try
        end try
      end if
    on error
    end try
  end repeat
end tell
EOF
    done
  fi
}

# Start background process to hide temp doc and open Default.docx
hide_temp_and_open_default &

# Start continuous monitoring to keep temp window minimized
# Wait for initial setup to complete (6s + 2s wait + 1s for opening = ~9s total)
sleep 9
monitor_temp_window &
MONITOR_PID=$!

# Set up signal handler to clean up on Ctrl+C
cleanup() {
  echo ""
  echo "Shutting down..."
  kill $MONITOR_PID 2>/dev/null
  if [ "$PROXY_RUNNING" = false ]; then
    echo "Stopping proxy server..."
    kill $PROXY_PID 2>/dev/null
  fi
  # Try to kill the add-in process if still running
  kill $ADDIN_PID 2>/dev/null
  exit 0
}

trap cleanup SIGINT SIGTERM

# Wait for add-in process (Ctrl+C will trigger cleanup)
# Note: office-addin-debugging may exit after starting dev server, so we keep waiting
# The dev server runs separately, and we want to keep the proxy running
wait $ADDIN_PID 2>/dev/null || true

# If we get here, the add-in command exited, but dev server may still be running
# Don't stop the proxy - let it keep running for the dev server
# User can press Ctrl+C to stop everything via the signal handler
echo ""
echo "Add-in debugging command completed. Dev server may still be running."
echo "Proxy server is still running on port 3001."
echo "Press Ctrl+C to stop the proxy server and exit."
echo ""

# Keep script alive so proxy keeps running
# Wait for Ctrl+C (handled by trap)
while true; do
  sleep 1
done

