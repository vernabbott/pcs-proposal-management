cat > gitpush.sh << 'EOF'
#!/bin/bash
# Quick helper script to add, commit, and push changes

# Use provided commit message, or fallback to "Update"
COMMIT_MSG=${1:-"Update"}

git add .
git commit -m "$COMMIT_MSG"
git push
EOF