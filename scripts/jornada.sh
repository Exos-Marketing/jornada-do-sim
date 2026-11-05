#!/usr/bin/env bash
# Roda o motor da jornada-do-sim isolado (-I, não carrega código da pasta atual) com o openpyxl de um ambiente
# efêmero do uv. Funciona igual no bash e no zsh. Uso: bash jornada.sh <subcomando> --config "<config.json>" ...
set -euo pipefail
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$HOME"
exec uv run --quiet --no-project --with openpyxl python -I "$DIR/jornada.py" "$@"
