#!/bin/bash
# Double-click launcher for macOS. Moves into this file's own folder first,
# so it works no matter where the project folder was placed or what the
# current Terminal directory happens to be, then runs the poster. The
# `read` at the end keeps the window open after the script exits (success
# or error) so the output doesn't disappear before it can be read.
cd "$(dirname "$0")"
python3 facebook_direct_poster_v1.py
echo
read -p "Готово. Нажмите Enter, чтобы закрыть окно..."
