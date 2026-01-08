# Bitboard Univer Publishing Scripts

Scripts for managing and publishing the Bitboard custom Univer fork.

## Setup

### 1. Install Verdaccio

```bash
npm install -g verdaccio
```

### 2. Start Verdaccio

```bash
verdaccio
```

Verdaccio will start on `http://localhost:4873`. Keep this terminal open.

### 3. Create User (First Time Only)

In a new terminal:
```bash
npm adduser --registry http://localhost:4873
```

Enter any username/password/email (for local dev).

## Scripts

### publish-to-verdaccio.sh

**Purpose:** Publish current version to Verdaccio registry

**Usage:**
```bash
./scripts/publish-to-verdaccio.sh
```

**What it does:**
1. Validates you're on a bitboard branch
2. Checks for uncommitted changes
3. Verifies Verdaccio is reachable
4. Installs dependencies
5. Builds all packages
6. Shows what will be published
7. Publishes to Verdaccio

**When to use:**
- First time publishing
- After manually making changes and bumping version
- Re-publishing after a failed publish

**Environment Variables:**
- `VERDACCIO_URL` - Registry URL (default: http://localhost:4873)

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
# Start Verdaccio (in separate terminal)
verdaccio

# Create user (first time only)
npm adduser --registry http://localhost:4873

# Ensure you're on the right branch
git checkout bitboard-v0.10.14

# Publish
./scripts/publish-to-verdaccio.sh
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

### "Cannot reach Verdaccio"
Make sure Verdaccio is running:
```bash
verdaccio
```

Or set custom URL:
```bash
export VERDACCIO_URL=https://your-verdaccio.com
```

### "401 Unauthorized" during publish
Re-authenticate with Verdaccio:
```bash
npm adduser --registry http://localhost:4873
```

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

- `.npmrc` - Created temporarily during publish (auto-cleaned)
- `package.json` - Version updates by bump-and-publish.sh

## Important Notes

- Verdaccio must be running for publish to work
- Published packages are stored in `~/.local/share/verdaccio/storage` on Mac
- `.npmrc` is automatically cleaned up after publish
- For production, see `VERDACCIO_SETUP.md` for deployment options
