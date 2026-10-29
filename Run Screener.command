#!/bin/bash
# Double-click this file to run the stock screener (ASX + US).
# First run sets up Python packages in a .venv folder here (a minute or two).
# When it finishes, the screener page opens in your browser.
# Results are saved in the screener_output folder next to this file.

cd "$(dirname "$0")" || exit 1
LOG="last_run.log"

run() {
  echo "Daily Stock Ratings - $(date '+%a %d %b %Y %H:%M')"
  echo

  PY=""
  for c in /opt/homebrew/bin/python3 /usr/local/bin/python3 /Library/Frameworks/Python.framework/Versions/Current/bin/python3 /usr/bin/python3; do
    if [ -x "$c" ] && "$c" -c "import sys; assert sys.version_info >= (3, 9)" >/dev/null 2>&1; then PY="$c"; break; fi
  done
  if [ -z "$PY" ]; then
    echo "Python 3.9 or newer wasn't found on this Mac."
    echo "Install it from https://www.python.org/downloads/macos/ then double-click this file again."
    return 1
  fi

  if [ ! -x .venv/bin/python ]; then
    echo "First run: setting up with $("$PY" --version 2>&1) (about a minute)..."
    "$PY" -m venv .venv || return 1
    .venv/bin/python -m pip install --quiet --upgrade pip || return 1
  fi
  echo "Checking packages..."
  .venv/bin/python -m pip install --quiet --upgrade -r requirements.txt || return 1
  echo

  .venv/bin/python -u cmc_screener.py --markets au us "$@"
}

run "$@" 2>&1 | tee "$LOG"
STATUS=${PIPESTATUS[0]}
echo "RUN_FINISHED exit=$STATUS $(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG"
echo
if [ "$STATUS" -eq 0 ]; then
  echo "Finished. You can close this window."
else
  echo "Something went wrong (see the messages above, also saved in $LOG)."
fi
