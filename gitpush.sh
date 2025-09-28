#!/bin/bash
# Add, commit, push with a default message if none provided
COMMIT_MSG=${1:-"Update"}
git add .
# If nothing to commit, exit quietly
if git diff --cached --quiet && git diff --quiet; then
  echo "Nothing to commit."
  exit 0
fi
git commit -m "$COMMIT_MSG"
git push