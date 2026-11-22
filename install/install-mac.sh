#!/bin/sh
# Formula Formatter for Excel - installer for Excel on Mac (Microsoft 365 / Excel 2021+).
# Usage:  curl -fsSL __BASE_URL__install-mac.sh | sh
# Remove: rm ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/excel-formula-formatter.xml
set -eu
WEF="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
mkdir -p "$WEF"
curl -fsSL "__BASE_URL__manifest.xml" -o "$WEF/excel-formula-formatter.xml"
grep -q "__ADDIN_ID__" "$WEF/excel-formula-formatter.xml" || { rm -f "$WEF/excel-formula-formatter.xml"; echo "Unexpected manifest, aborting." >&2; exit 1; }
echo "Installed. Quit Excel completely (Cmd+Q) and start it again."
echo "If the ribbon tab does not appear: Home > Add-ins, and pick Formula Formatter for Excel under Developer Add-ins."
