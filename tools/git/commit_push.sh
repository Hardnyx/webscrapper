#!/usr/bin/env bash
set -euo pipefail
cd "$(git rev-parse --show-toplevel)"
git status --short
if [[ $# -gt 0 ]]; then
    message="$1"
else
    read -r -p "Conventional commit message: " message
fi
if [[ ! "$message" =~ ^[a-z]+(\([a-zA-Z0-9_.-]+\))?\!?:\ .+ ]]; then
    echo 'Use a conventional commit message, such as feat(sbs): add a data source.' >&2
    exit 1
fi
git add -A
if git diff --cached --quiet; then
    echo 'No staged changes.'
    exit 0
fi
git commit -m "$message"
git push
