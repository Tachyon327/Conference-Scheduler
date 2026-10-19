#!/bin/bash
# Double-click to start Conference Scheduler from the Python source.
# The first run creates a private Python environment (.venv) and
# installs the required libraries into it; later runs start at once.
cd "$(dirname "$0")"

if ! command -v python3 &>/dev/null; then
    echo "Python 3 is required."
    echo "Install it from https://www.python.org/downloads/ and run this again."
    echo "Press Enter to close this window."; read -r; exit 1
fi

# Tkinter draws the window. The python.org installer includes it;
# a Homebrew Python needs the separate python-tk package.
if ! python3 -c "import tkinter" &>/dev/null; then
    if command -v brew &>/dev/null; then
        echo "Installing python-tk (needed to draw the window)..."
        brew install python-tk@3.14 || brew install python-tk
    else
        echo "This Python has no Tkinter, which is needed to draw the window."
        echo "Install Python from https://www.python.org/downloads/ and run this again."
        echo "Press Enter to close this window."; read -r; exit 1
    fi
fi

if [ ! -d ".venv" ] || ! .venv/bin/python -c "import tkinter" &>/dev/null; then
    echo "Setting up the Python environment (first run only)..."
    rm -rf .venv
    python3 -m venv .venv
    source .venv/bin/activate
    pip install -r requirements.txt
else
    source .venv/bin/activate
fi

echo "Starting Conference Scheduler..."
python new_conference_scheduler_V1.py
