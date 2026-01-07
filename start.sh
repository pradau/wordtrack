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

cd /Users/pradau/Dropbox/ChildrensHospital/IT/scripts/wordtrack

echo "Starting WordTrack..."
echo "Starting proxy server in background..."
npm run proxy &
PROXY_PID=$!

sleep 2

echo "Starting add-in debugging..."
npx office-addin-debugging start manifest.xml

echo "Stopping proxy server..."
kill $PROXY_PID 2>/dev/null

