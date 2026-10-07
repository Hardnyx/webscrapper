#!/usr/bin/env bash
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"
mkdir -p outputs
git ls-files | sort > outputs/filetree.txt
printf 'File inventory: outputs/filetree.txt\n'
