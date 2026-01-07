#!/bin/bash

cd /Users/pradau/Dropbox/ChildrensHospital/IT/scripts/wordtrack

echo "Starting WordTrack..."
echo "Terminal 1: Starting proxy server..."
npm run proxy &
PROXY_PID=$!

sleep 2

echo "Terminal 2: Starting add-in debugging..."
npx office-addin-debugging start manifest.xml

echo "Stopping proxy server..."
kill $PROXY_PID 2>/dev/null

