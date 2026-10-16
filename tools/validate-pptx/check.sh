#!/usr/bin/env bash
# Validate an exported .pptx the way PowerPoint's own checker would, and show
# only the problems that the edit introduced.
#
#   tools/validate-pptx/check.sh edited.pptx [original.pptx]
#
# Needs the .NET 8 SDK (installed per user with:
#   curl -sSL https://dot.net/v1/dotnet-install.sh | bash -s -- --channel 8.0 --install-dir ~/.dotnet).
set -euo pipefail

here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
export PATH="$HOME/.dotnet:$PATH" DOTNET_CLI_TELEMETRY_OPTOUT=1 DOTNET_NOLOGO=1

if ! command -v dotnet >/dev/null; then
  echo "dotnet not found; see the header of this script for how to install it." >&2
  exit 2
fi

dll="$here/bin/Release/net8.0/validate-pptx.dll"
[ -f "$dll" ] || dotnet build -c Release "$here" >/dev/null

# Compare problems by what they are, not where in the tree they sit.
normalise() { cut -f1,2,4,5 | sort -u; }

edited="$(dotnet "$dll" "$1" 2>/dev/null | normalise || true)"
if [ -n "${2:-}" ]; then
  original="$(dotnet "$dll" "$2" 2>/dev/null | normalise || true)"
  new="$(comm -13 <(echo "$original") <(echo "$edited") | sed '/^$/d')"
else
  new="$(echo "$edited" | sed '/^$/d')"
fi

if [ -z "$new" ]; then
  echo "OK: no problems${2:+ beyond those already in the original}."
else
  echo "$new" | awk -F'\t' '{ printf "%s  %s\n    %s: %s\n", $1, $2, $3, $4 }'
  exit 1
fi
