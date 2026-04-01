#!/bin/bash

echo "Current status:"
git status

echo
read -p "Commit message: " MSG

# Default message if empty
if [ -z "$MSG" ]; then
  MSG="Auto: quick update"
fi

echo
echo "Adding files..."
git add .

echo "Committing..."
git commit -m "$MSG"
if [ $? -ne 0 ]; then
  echo
  echo "No commit made. Probably no changes."
  exit 0
fi

echo
echo "Pushing to origin..."
git push

echo