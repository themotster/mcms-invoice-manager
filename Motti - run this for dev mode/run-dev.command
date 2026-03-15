#!/bin/bash
# Double-click this file to run dev mode. Terminal will open and start the app—no typing needed.
cd "$(dirname "$0")"
chmod +x run-dev.sh 2>/dev/null
exec ./run-dev.sh
