#!/usr/bin/env bash
#
# Set up the Ubuntu (Excel-free) SheetCopilot evaluation environment.
#
# Installs:
#   * LibreOffice Calc + the python3-uno bridge (apt)  -- charts & pivot tables
#   * openpyxl / pandas / numpy / pyyaml / tqdm (pip)  -- everything else
# and then runs a smoke test that re-evaluates the bundled example logs and
# checks the verdicts against the original Windows evaluator's results.
#
# Usage:
#   bash setup_ubuntu_eval.sh            # install + smoke test
#   bash setup_ubuntu_eval.sh --no-test  # install only
#
# Run from the `agent/` directory. Use sudo if you are not root.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

SUDO=""
if [ "$(id -u)" -ne 0 ]; then
    SUDO="sudo"
fi

echo "==> Installing LibreOffice Calc + python3-uno (apt) ..."
export DEBIAN_FRONTEND=noninteractive
$SUDO apt-get update -qq
$SUDO apt-get install -y --no-install-recommends \
    libreoffice-calc \
    libreoffice-script-provider-python \
    python3-uno

echo "==> Installing Python dependencies (pip) ..."
python3 -m pip install -r "$SCRIPT_DIR/requirements_ubuntu.txt"

echo "==> Verifying toolchain ..."
soffice --version
python3 -c "import uno; print('uno bridge: OK')"
python3 -c "import openpyxl, pandas, numpy, yaml, tqdm; print('python deps: OK')"

if [ "${1:-}" != "--no-test" ]; then
    echo "==> Running smoke test (re-evaluating example logs) ..."
    python3 -m ubuntu_eval.selftest
fi

echo "==> Done. Evaluate your results with:"
echo "    python3 evaluation_ubuntu.py -c config/config.yaml"
