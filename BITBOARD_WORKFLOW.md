# Bitboard Univer Fork Workflow

This document describes how to maintain and publish the Bitboard custom fork of Univer.

## Repository Setup

**Fork:** `https://github.com/bitboard-dev/univer`
**Upstream:** `https://github.com/dream-num/univer`

### Git Remotes

```bash
origin   → https://github.com/bitboard-dev/univer.git (your fork)
upstream → https://github.com/dream-num/univer.git (upstream)
```

### Branch Strategy

- `dev` - tracks upstream dev branch (don't modify)
- `bitboard-v0.12.3` - your custom branch based on v0.12.3
- Future: `bitboard-v0.13.0` when upgrading

## Custom Changes

Current customizations in `bitboard-v0.12.3`:

1. **Version suffix:** All packages versioned as `0.12.3-bitboard.1`
2. **Toolbar reorganization:** Moved Data menu items to main toolbar
   - Data Validation → Start ribbon
   - Conditional Formatting → Start ribbon
   - Filter, Sort, Find/Replace → Start ribbon
   - Table controls → Start ribbon

## Building Packages

```bash
# Install dependencies
pnpm install

# Build all packages
pnpm build

# Build specific package
pnpm --filter @univerjs/ui build

# Run tests
pnpm test
```

## Publishing to Private npm Registry

### Option 1: Verdaccio (Local/Self-hosted)

1. Install Verdaccio:
```bash
npm install -g verdaccio
```

2. Start Verdaccio:
```bash
verdaccio
# Runs on http://localhost:4873
```

3. Configure npm to use Verdaccio:
```bash
npm set registry http://localhost:4873
```

4. Create user and login:
```bash
npm adduser --registry http://localhost:4873
```

5. Publish all packages:
```bash
pnpm publish -r --registry http://localhost:4873
```

### Option 2: GitHub Packages

1. Create a `.npmrc` file in the project root:
```
@univerjs:registry=https://npm.pkg.github.com
//npm.pkg.github.com/:_authToken=${GITHUB_TOKEN}
```

2. Publish:
```bash
pnpm publish -r --registry https://npm.pkg.github.com --no-git-checks
```

### Option 3: Private npm Registry (npm Enterprise, Artifactory, etc.)

Configure your registry URL and credentials in `.npmrc` or use:
```bash
pnpm publish -r --registry https://your-registry.com
```

## Consuming in Bitboard App

### In your Bitboard app's `.npmrc`:
```
@univerjs:registry=http://localhost:4873
# or your private registry URL
```

### In your Bitboard app's `package.json`:
```json
{
  "dependencies": {
    "@univerjs/core": "0.12.3-bitboard.1",
    "@univerjs/sheets": "0.12.3-bitboard.1",
    "@univerjs/ui": "0.12.3-bitboard.1"
  }
}
```

## Pulling Upstream Updates

When you want to upgrade to a new upstream version (e.g., v0.13.0):

### 1. Fetch latest from upstream
```bash
git fetch upstream
git fetch upstream --tags
```

### 2. Create new custom branch
```bash
git checkout -b bitboard-v0.13.0 v0.13.0
```

### 3. Re-apply version changes
```bash
# Update all package versions to 0.13.0-bitboard.1
find packages -name "package.json" -type f -exec sed -i '' 's/"version": "0\.13\.0"/"version": "0.13.0-bitboard.1"/g' {} \;
sed -i '' 's/"version": "0\.13\.0"/"version": "0.13.0-bitboard.1"/g' package.json
```

### 4. Cherry-pick or manually apply your custom changes
```bash
# Option A: Cherry-pick your custom commits
git cherry-pick f744738a8  # Toolbar changes commit

# Option B: Manually re-apply changes
# Edit the same 6 menu.schema.ts files to move items from RibbonDataGroup to RibbonStartGroup
```

### 5. Test, commit, and publish
```bash
pnpm install
pnpm build
pnpm test

git add .
git commit -m "chore: upgrade to v0.13.0 with Bitboard customizations"

# Push to fork
git push origin bitboard-v0.13.0

# Publish new version
pnpm publish -r --registry http://localhost:4873
```

## Incrementing Custom Version

When making additional changes to the same base version:

```bash
# Update from 0.12.3-bitboard.1 to 0.12.3-bitboard.2
find packages -name "package.json" -type f -exec sed -i '' 's/"version": "0\.12\.3-bitboard\.1"/"version": "0.12.3-bitboard.2"/g' {} \;
sed -i '' 's/"version": "0\.12\.3-bitboard\.1"/"version": "0.12.3-bitboard.2"/g' package.json

git add .
git commit -m "chore: bump to 0.12.3-bitboard.2"
pnpm publish -r --registry http://localhost:4873
```

## Important Commands

```bash
# Check current status
git status
git branch
git log --oneline -10

# View your custom commits
git log upstream/dev..HEAD --oneline

# Push your branch to fork
git push origin bitboard-v0.12.3

# Compare with upstream version
git diff v0.12.3..HEAD

# List modified packages
git diff --name-only v0.12.3..HEAD | grep package.json
```

## Troubleshooting

### Build fails after upgrade
- Check if upstream introduced breaking changes
- Review CHANGELOG.md for migration guides
- May need to update your custom changes

### Publish fails
- Ensure you're authenticated with registry
- Check package.json publishConfig
- Verify version numbers don't conflict

### Conflicts during cherry-pick
- Resolve conflicts manually
- Upstream may have refactored files you modified
- Consider reapplying changes manually instead

## Files Modified in This Fork

Custom changes are in these files:
- All `package.json` files (version updates)
- `packages/sheets-data-validation-ui/src/controllers/menu.schema.ts`
- `packages/sheets-conditional-formatting-ui/src/controllers/menu.schema.ts`
- `packages/sheets-filter-ui/src/controllers/menu.schema.ts`
- `packages/sheets-sort-ui/src/controllers/menu.schema.ts`
- `packages/sheets-table-ui/src/controllers/menu.schema.ts`
- `packages/find-replace/src/controllers/menu.schema.ts`
