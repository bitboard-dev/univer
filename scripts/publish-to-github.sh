#!/bin/bash
set -e

# Bitboard Univer Publishing Script
# Publishes custom Univer packages to GitHub Packages

echo "========================================"
echo "Bitboard Univer Publishing Script"
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

# Check for uncommitted changes (excluding .npmrc which is temporary)
UNCOMMITTED=$(git status --porcelain | grep -v "^?? .npmrc" || true)
if [[ -n "$UNCOMMITTED" ]]; then
    echo "❌ Error: You have uncommitted changes"
    echo "$UNCOMMITTED"
    echo ""
    echo "Please commit or stash your changes first"
    exit 1
fi

echo "✓ No uncommitted changes"

# Check for GitHub token
if [[ -z "$UNIVER_PUBLISH_TOKEN" ]]; then
    echo "❌ Error: UNIVER_PUBLISH_TOKEN environment variable not set"
    echo ""
    echo "Create a GitHub token with 'write:packages' permission:"
    echo "  https://github.com/settings/tokens/new"
    echo ""
    echo "Then export it:"
    echo "  export UNIVER_PUBLISH_TOKEN=your_token_here"
    exit 1
fi

echo "✓ GitHub token found"

# Create .npmrc for publishing (only auth, don't redirect registry during install)
echo "Setting up npm authentication..."
cat > .npmrc << EOF
//npm.pkg.github.com/:_authToken=${UNIVER_PUBLISH_TOKEN}
EOF

echo "✓ .npmrc configured"

# Cleanup function to remove .npmrc on exit
cleanup() {
    echo ""
    echo "Cleaning up..."
    rm -f .npmrc
    echo "✓ .npmrc removed"
}
trap cleanup EXIT

# Install dependencies
echo ""
echo "Installing dependencies..."
pnpm install --frozen-lockfile

echo "✓ Dependencies installed"

# Build all packages
echo ""
echo "Building packages..."
pnpm build

echo "✓ Build complete"

# Show what will be published
echo ""
echo "========================================"
echo "Ready to publish the following packages:"
echo "========================================"
VERSION=$(node -p "require('./package.json').version")
echo "Version: $VERSION"
echo ""
find packages -name "package.json" -exec sh -c '
  PKG=$(cat "$1")
  NAME=$(echo "$PKG" | grep "\"name\":" | head -1 | cut -d"\"" -f4)
  PRIVATE=$(echo "$PKG" | grep "\"private\":" | head -1 | cut -d":" -f2 | tr -d ", ")
  if [ "$PRIVATE" != "true" ]; then
    echo "  - $NAME"
  fi
' sh {} \;

echo ""
read -p "Proceed with publishing? (y/N): " -n 1 -r
echo ""

if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    echo "Aborted"
    exit 0
fi

# Publish to GitHub Packages
echo ""
echo "Publishing to GitHub Packages..."
pnpm publish -r --registry https://npm.pkg.github.com --no-git-checks

echo ""
echo "========================================"
echo "✓ Publishing complete!"
echo "========================================"
echo ""
echo "Packages published to:"
echo "  https://github.com/orgs/bitboard-dev/packages"
echo ""
echo "To use in Bitboard, add to .npmrc:"
echo "  @univerjs:registry=https://npm.pkg.github.com"
echo "  //npm.pkg.github.com/:_authToken=\${UNIVER_PUBLISH_TOKEN}"
echo ""
echo "Then update package.json dependencies to: $VERSION"
