# Bitboard Univer Publishing Scripts

Scripts for managing and publishing the Bitboard custom Univer fork.

## Setup

### 1. Create GitHub Personal Access Token

Create a token with `write:packages` permission:
https://github.com/settings/tokens/new

Permissions needed:
- ✅ `write:packages` - Upload packages to GitHub Package Registry
- ✅ `read:packages` - Download packages from GitHub Package Registry

### 2. Export Token

```bash
export GITHUB_TOKEN=your_token_here
```

Add to your shell profile (~/.zshrc, ~/.bashrc) for persistence:
```bash
echo 'export GITHUB_TOKEN=your_token_here' >> ~/.zshrc
```

## Scripts

### publish-to-github.sh

**Purpose:** Publish current version to GitHub Packages

**Usage:**
```bash
./scripts/publish-to-github.sh
```

**What it does:**
1. Validates you're on a bitboard branch
2. Checks for uncommitted changes
3. Configures packages for GitHub Packages
4. Installs dependencies
5. Builds all packages
6. Shows what will be published
7. Publishes to GitHub Packages

**When to use:**
- First time publishing
- After manually making changes and bumping version
- Re-publishing after a failed publish

### bump-and-publish.sh

**Purpose:** One-stop script to bump version and publish

**Usage:**
```bash
# Bump patch version (0.10.14-bitboard.1 → 0.10.14-bitboard.2)
./scripts/bump-and-publish.sh patch

# Bump minor version (0.10.14-bitboard.1 → 0.10.15-bitboard.1)
./scripts/bump-and-publish.sh minor

# Bump major version (0.10.14-bitboard.1 → 0.11.0-bitboard.1)
./scripts/bump-and-publish.sh major
```

**What it does:**
1. Validates environment
2. Calculates new version
3. Updates all package.json files
4. Commits the version bump
5. Optionally runs publish-to-github.sh

**When to use:**
- Making additional customizations to same upstream version (patch)
- Upgrading to new upstream minor/major version

### Example Workflows

#### First Time Publishing

```bash
# Ensure you're on the right branch
git checkout bitboard-v0.10.14

# Export GitHub token
export GITHUB_TOKEN=ghp_xxxxxxxxxxxx

# Publish
./scripts/publish-to-github.sh
```

#### Making a Change and Republishing

```bash
# Make your changes
vim packages/ui/src/some-file.ts

# Commit changes
git add .
git commit -m "feat: add custom feature"

# Bump version and publish
./scripts/bump-and-publish.sh patch
```

#### Upgrading to New Upstream Version

```bash
# Fetch latest tags
git fetch upstream --tags

# Create new branch from upstream tag
git checkout -b bitboard-v0.11.0 v0.11.0

# Apply version suffix
find packages -name "package.json" -type f -exec sed -i '' 's/"version": "0\.11\.0"/"version": "0.11.0-bitboard.1"/g' {} \;
sed -i '' 's/"version": "0\.11\.0"/"version": "0.11.0-bitboard.1"/g' package.json

# Cherry-pick your customizations
git cherry-pick 947dd7dc6  # Your custom commit from v0.10.14

# Test build
pnpm install
pnpm build

# Commit and publish
git add .
git commit -m "chore: upgrade to v0.11.0 with Bitboard customizations"
./scripts/publish-to-github.sh
```

## Troubleshooting

### "GITHUB_TOKEN not set"
Export your GitHub token:
```bash
export GITHUB_TOKEN=ghp_xxxxxxxxxxxx
```

### "403 Forbidden" during publish
- Check token has `write:packages` permission
- Verify you have access to bitboard-dev org
- Check token hasn't expired

### Build fails
```bash
# Clean and rebuild
rm -rf node_modules pnpm-lock.yaml
pnpm install
pnpm build
```

### Package already exists
GitHub Packages doesn't allow overwriting versions. Bump the version:
```bash
./scripts/bump-and-publish.sh patch
```

## Files Modified by Scripts

- `packages/*/package.json` - Adds GitHub registry to publishConfig
- `.npmrc` - Created temporarily for authentication
- `package.json` - Version updates

## Security Notes

- Never commit `.npmrc` with tokens to git (it's in .gitignore)
- Use environment variables for GITHUB_TOKEN
- Rotate tokens periodically
- Use fine-grained tokens when possible
