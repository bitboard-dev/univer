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

# Check for uncommitted changes
if [[ -n $(git status --porcelain) ]]; then
    echo "❌ Error: You have uncommitted changes"
    git status --short
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

# Create .npmrc for publishing
echo "Setting up npm authentication..."
cat > .npmrc << EOF
@univerjs:registry=https://npm.pkg.github.com
//npm.pkg.github.com/:_authToken=\${UNIVER_PUBLISH_TOKEN}
EOF

echo "✓ .npmrc configured"

# Update all package.json files to use GitHub Packages registry
echo ""
echo "Configuring packages for GitHub Packages..."
PACKAGE_COUNT=$(find packages -name "package.json" | wc -l | tr -d ' ')
echo "Found $PACKAGE_COUNT packages to configure"

# Add registry to publishConfig in all packages
node << 'ENDNODE'
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// Find all package.json files
const packageFiles = execSync('find packages -name "package.json"', { encoding: 'utf-8' })
  .trim()
  .split('\n')
  .filter(Boolean);

let updated = 0;
packageFiles.forEach(file => {
  const content = JSON.parse(fs.readFileSync(file, 'utf-8'));

  if (content.private === true) {
    console.log(`Skipping private package: ${content.name}`);
    return;
  }

  if (!content.publishConfig) {
    content.publishConfig = {};
  }

  // Add registry to publishConfig
  content.publishConfig.registry = 'https://npm.pkg.github.com';

  fs.writeFileSync(file, JSON.stringify(content, null, 2) + '\n');
  updated++;
});

console.log(`✓ Updated ${updated} package.json files`);
ENDNODE

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
