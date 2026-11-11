#!/usr/bin/env bash
#
# Build the distributable folder (dist/) containing everything needed to run
# the two helpers without a source checkout:
#
#   clean_aa_backups.py   + requirements-clean-aa.txt
#   run_aa_workbook.py    + requirements-run-aa.txt
#   README.md, LICENSE
#
# Usage: ./build.sh   (safe to run from anywhere; it uses its own directory)

set -euo pipefail

cd "$(dirname "$0")"

# Start from a clean dist/ so it never contains stale files.
rm -rf dist
mkdir -p dist

# Requirements, one per tool.
cp requirements-clean-aa.txt dist/
cp requirements-run-aa.txt  dist/

# Documentation and license.
cp README.md dist/
cp LICENSE   dist/

# The two entry-point scripts.
cp src/clean_aa_backups/clean_aa_backups.py dist/
cp src/run_aa_workbook/run_aa_workbook.py   dist/

echo "Built dist/:"
ls -la dist/
