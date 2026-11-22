#!/bin/sh
# Formula Formatter for Excel - installer for Excel on Mac (Microsoft 365 / Excel 2021+).
# Usage:  curl -fsSL https://kmailmom.github.io/excel-formula-formatter/install-mac.sh | sh
# Remove: rm ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/excel-formula-formatter.xml
set -eu
WEF="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
mkdir -p "$WEF"
curl -fsSL "https://kmailmom.github.io/excel-formula-formatter/manifest.xml" -o "$WEF/excel-formula-formatter.xml"
grep -q "117076a0-2ea1-4038-be03-5425e8b2a84a" "$WEF/excel-formula-formatter.xml" || { rm -f "$WEF/excel-formula-formatter.xml"; echo "Unexpected manifest, aborting." >&2; exit 1; }
echo "Installed. Quit Excel completely (Cmd+Q) and start it again."
echo "If the ribbon tab does not appear: Home > Add-ins, and pick Formula Formatter for Excel under Developer Add-ins."
