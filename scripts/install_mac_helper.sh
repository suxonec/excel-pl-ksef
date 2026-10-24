#!/bin/sh
# Install the AppleScript launcher that lets Excel for Mac start ksefiz (one-time, per user).
set -eu
here="$(cd "$(dirname "$0")/.." && pwd)"
dest="$HOME/Library/Application Scripts/com.microsoft.Excel"
mkdir -p "$dest"
osacompile -o "$dest/ksefiz.scpt" "$here/excel/mac/ksefiz.applescript"
echo "installed $dest/ksefiz.scpt"
echo "set Config > helper_path in the workbook to: $here/.venv/bin/ksefiz"
