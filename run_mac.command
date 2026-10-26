#!/bin/bash
# Double-click launcher for macOS. Moves into this file's own folder first,
# so it works no matter where the project folder was placed or what the
# current Terminal directory happens to be, then runs the poster. The
# `read` at the end keeps the window open after the script exits (success
# or error) so the output doesn't disappear before it can be read.
cd "$(dirname "$0")"

# Pull the latest script/groups.xlsx from GitHub before every run, so a
# double-click always uses whatever was last pushed without a manual
# `git pull` first. If this fails (no internet, a local edit conflicting
# with what's upstream) it's reported but not fatal -- the run continues
# with whatever version is already on disk rather than refusing to start.
echo "Checking for updates..."
git pull || echo "Could not update (continuing with the current version)."
echo

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
