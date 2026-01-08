#!/bin/bash
set -e

# Bitboard Univer Publishing Script (Verdaccio)
# Publishes custom Univer packages to Verdaccio registry

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

# Default to localhost if VERDACCIO_URL not set
VERDACCIO_URL=${VERDACCIO_URL:-http://localhost:4873}

echo "✓ Using registry: $VERDACCIO_URL"

# Check if Verdaccio is reachable
if ! curl -sf "$VERDACCIO_URL" > /dev/null; then
    echo "❌ Error: Cannot reach Verdaccio at $VERDACCIO_URL"
    echo ""
    echo "Make sure Verdaccio is running:"
    echo "  npm install -g verdaccio"
    echo "  verdaccio"
    echo ""
    echo "Or set VERDACCIO_URL to your hosted instance:"
    echo "  export VERDACCIO_URL=https://your-verdaccio.com"
    exit 1
fi

echo "✓ Verdaccio is reachable"

# Cleanup function to remove .npmrc on exit
cleanup() {
    echo ""
    echo "Cleaning up..."
    rm -f .npmrc
    echo "✓ .npmrc removed"
}
trap cleanup EXIT

# Create .npmrc for publishing
echo "Setting up npm registry..."
cat > .npmrc << EOF
registry=$VERDACCIO_URL
EOF

echo "✓ .npmrc configured"

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
echo "Registry: $VERDACCIO_URL"
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

# Publish to Verdaccio
echo ""
echo "Publishing to Verdaccio..."
pnpm publish -r --registry "$VERDACCIO_URL" --no-git-checks

echo ""
echo "========================================"
echo "✓ Publishing complete!"
echo "========================================"
echo ""
echo "Packages published to: $VERDACCIO_URL"
echo ""
echo "To use in Bitboard, add to .npmrc:"
echo "  @univerjs:registry=$VERDACCIO_URL"
echo ""
echo "Then update package.json dependencies to: $VERSION"
