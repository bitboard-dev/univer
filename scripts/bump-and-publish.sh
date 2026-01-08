#!/bin/bash
set -e

# Bitboard Univer Version Bump & Publish Script
# Usage: ./scripts/bump-and-publish.sh [patch|minor|major]

echo "========================================"
echo "Bitboard Univer Version Bump & Publish"
echo "========================================"
echo ""

# Check if we're on a bitboard branch
CURRENT_BRANCH=$(git branch --show-current)
if [[ ! "$CURRENT_BRANCH" =~ ^bitboard- ]]; then
    echo "❌ Error: Not on a bitboard branch (current: $CURRENT_BRANCH)"
    echo "Please checkout a bitboard-* branch first"
    exit 1
fi

echo "✓ On branch: $CURRENT_BRANCH"

# Check for uncommitted changes
if [[ -n $(git status --porcelain) ]]; then
    echo "❌ Error: You have uncommitted changes"
    git status --short
    echo ""
    echo "Please commit or stash your changes first"
    exit 1
fi

echo "✓ No uncommitted changes"

# Get bump type (default: patch)
BUMP_TYPE=${1:-patch}
if [[ ! "$BUMP_TYPE" =~ ^(patch|minor|major)$ ]]; then
    echo "❌ Error: Invalid bump type: $BUMP_TYPE"
    echo "Usage: $0 [patch|minor|major]"
    echo ""
    echo "Examples:"
    echo "  $0 patch   # 0.10.14-bitboard.1 → 0.10.14-bitboard.2"
    echo "  $0 minor   # 0.10.14-bitboard.1 → 0.10.15-bitboard.1"
    echo "  $0 major   # 0.10.14-bitboard.1 → 0.11.0-bitboard.1"
    exit 1
fi

# Get current version
CURRENT_VERSION=$(node -p "require('./package.json').version")
echo "Current version: $CURRENT_VERSION"

# Calculate new version
NEW_VERSION=$(node << ENDNODE
const semver = require('./package.json').version;
const match = semver.match(/^(\d+)\.(\d+)\.(\d+)-bitboard\.(\d+)$/);

if (!match) {
  console.error('Invalid version format');
  process.exit(1);
}

let [, major, minor, patch, bitboard] = match.map(Number);

const bumpType = '$BUMP_TYPE';

if (bumpType === 'patch') {
  bitboard++;
} else if (bumpType === 'minor') {
  patch++;
  bitboard = 1;
} else if (bumpType === 'major') {
  minor++;
  patch = 0;
  bitboard = 1;
}

console.log(\`\${major}.\${minor}.\${patch}-bitboard.\${bitboard}\`);
ENDNODE
)

if [[ -z "$NEW_VERSION" ]]; then
    echo "❌ Error: Failed to calculate new version"
    exit 1
fi

echo "New version:     $NEW_VERSION"
echo ""
read -p "Proceed with version bump? (y/N): " -n 1 -r
echo ""

if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    echo "Aborted"
    exit 0
fi

# Update all package.json files
echo ""
echo "Updating package versions..."

# Escape dots for sed
CURRENT_ESCAPED=$(echo "$CURRENT_VERSION" | sed 's/\./\\./g')
NEW_ESCAPED=$(echo "$NEW_VERSION" | sed 's/\./\\./g')

# Update all packages
find packages -name "package.json" -type f -exec sed -i '' "s/\"version\": \"$CURRENT_ESCAPED\"/\"version\": \"$NEW_ESCAPED\"/g" {} \;

# Update root package.json
sed -i '' "s/\"version\": \"$CURRENT_ESCAPED\"/\"version\": \"$NEW_ESCAPED\"/g" package.json

echo "✓ Updated to version $NEW_VERSION"

# Commit the version bump
echo ""
echo "Committing version bump..."
git add .
git commit -m "chore: bump version to $NEW_VERSION"

echo "✓ Version committed"

# Ask if should publish
echo ""
read -p "Publish to Verdaccio now? (y/N): " -n 1 -r
echo ""

if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    echo "Version bumped but not published"
    echo "Run ./scripts/publish-to-verdaccio.sh when ready"
    exit 0
fi

# Run publish script
echo ""
./scripts/publish-to-verdaccio.sh

echo ""
echo "========================================"
echo "✓ Version bump and publish complete!"
echo "========================================"
echo ""
echo "Don't forget to push your changes:"
echo "  git push origin $CURRENT_BRANCH"
