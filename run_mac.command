#!/bin/bash
# Double-click launcher for macOS. Moves into this file's own folder first,
# so it works no matter where the project folder was placed or what the
# current Terminal directory happens to be, then runs the poster. The
# `read` at the end keeps the window open after the script exits (success
# or error) so the output doesn't disappear before it can be read.
cd "$(dirname "$0")"
# -d/-i/-s stop display, idle, and (on AC power) system sleep for as long
# as this process runs -- the run that got cut off when the lid closed
# died because macOS suspended the Chrome<->chromedriver connection
# during sleep. NOTE: caffeinate cannot override an actual closed lid
# with no external display attached (Apple's "clamshell sleep" always
# wins) -- keep the lid open (or hook up an external display) while a
# run is in progress.
caffeinate -d -i -s python3 facebook_direct_poster_v1.py
echo
read -p "Готово. Нажмите Enter, чтобы закрыть окно..."
