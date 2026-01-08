# Verdaccio Setup for Bitboard Univer

Complete guide to setting up Verdaccio as your private npm registry for the custom Univer fork.

## Why Verdaccio?

- ✅ No package scope restrictions (keep `@univerjs/*` names)
- ✅ Simple to set up and run
- ✅ Works locally and can be deployed for production
- ✅ Free and open source

## Local Development Setup

### 1. Install Verdaccio

```bash
npm install -g verdaccio
```

### 2. Start Verdaccio

```bash
verdaccio
```

This starts Verdaccio on `http://localhost:4873`

You'll see output like:
```
 warn --- config file  - /Users/you/.config/verdaccio/config.yaml
 warn --- http address - http://localhost:4873/ - verdaccio/5.x.x
```

Keep this terminal open (Verdaccio needs to stay running).

### 3. Create User Account

In a new terminal:
```bash
npm adduser --registry http://localhost:4873
```

Enter username, password, and email (can be fake for local dev).

### 4. Publish Univer Packages

```bash
cd /Users/ambar/projects/bitboard/univer
./scripts/publish-to-verdaccio.sh
```

The script will:
- Build all packages
- Publish to `http://localhost:4873`
- Show confirmation

### 5. Verify Published Packages

Open http://localhost:4873 in your browser to see all published packages.

## Production Setup (Render Deployment)

For Render to access your packages, Verdaccio needs to be publicly accessible.

### Option A: Deploy Verdaccio to Render

**1. Create `verdaccio-config.yaml`:**

```yaml
storage: /verdaccio/storage
auth:
  htpasswd:
    file: /verdaccio/htpasswd
uplinks:
  npmjs:
    url: https://registry.npmjs.org/
packages:
  '@univerjs/*':
    access: $authenticated
    publish: $authenticated
  '**':
    access: $all
    proxy: npmjs
logs: { type: stdout, format: pretty, level: http }
```

**2. Create Dockerfile:**

```dockerfile
FROM verdaccio/verdaccio:5

COPY verdaccio-config.yaml /verdaccio/conf/config.yaml

EXPOSE 4873
```

**3. Deploy to Render:**
- Create new "Web Service" on Render
- Connect your repo
- Set build command: `docker build -t verdaccio .`
- Set start command: `docker run -p $PORT:4873 verdaccio`
- Add persistent disk at `/verdaccio/storage` (for package storage)

**4. Set Environment Variables in Bitboard:**
```
VERDACCIO_URL=https://your-verdaccio.onrender.com
```

### Option B: Use Render Disk + Background Process

Run Verdaccio as a background process in your Bitboard service:

**1. Update Bitboard's `render.yaml`:**

```yaml
buildCommand: |
  npm install -g verdaccio
  pip install -r requirements.txt
  npm ci
  npm run build:univer
  npm run build:executor

startCommand: |
  verdaccio --config /etc/verdaccio/config.yaml &
  sleep 5
  bash scripts/render-start.sh
```

**2. Create verdaccio config in your Bitboard repo**

This approach runs Verdaccio alongside your Bitboard app, so packages are available during build.

### Option C: Pre-built Packages (Simplest for Render)

**Don't use a registry at all for production:**

1. Build Univer packages locally
2. Commit `node_modules` to git (or use a separate repo)
3. Reference via file path or git URL in package.json

```json
{
  "dependencies": {
    "@univerjs/core": "file:../univer/packages/core",
    "@univerjs/ui": "file:../univer/packages/ui"
  }
}
```

## Using Packages in Bitboard

### Local Development

**1. Create `.npmrc` in Bitboard repo:**
```
@univerjs:registry=http://localhost:4873
```

**2. Update `package.json`:**
```json
{
  "dependencies": {
    "@univerjs/core": "0.10.14-bitboard.1",
    "@univerjs/ui": "0.10.14-bitboard.1"
    // ... all other @univerjs packages
  }
}
```

**3. Install:**
```bash
npm install
```

npm will fetch `@univerjs/*` packages from Verdaccio, everything else from public npm.

### Production (Render)

Add `.npmrc` to your Bitboard repo (committed to git):
```
@univerjs:registry=${VERDACCIO_URL}
```

Then set `VERDACCIO_URL` in Render dashboard.

## Troubleshooting

### Verdaccio won't start
```bash
# Check if port 4873 is already in use
lsof -ti:4873 | xargs kill

# Start again
verdaccio
```

### Can't publish (401 Unauthorized)
```bash
# Re-authenticate
npm adduser --registry http://localhost:4873
```

### Packages not installing in Bitboard
```bash
# Check .npmrc is correct
cat .npmrc

# Verify Verdaccio is running
curl http://localhost:4873

# Check if packages exist
curl http://localhost:4873/@univerjs/core
```

### Reset Everything
```bash
# Kill Verdaccio
pkill -f verdaccio

# Remove storage (deletes all packages)
rm -rf ~/.local/share/verdaccio/storage

# Start fresh
verdaccio
npm adduser --registry http://localhost:4873
```

## Recommended Workflow

**For local development:**
1. Run Verdaccio locally on localhost:4873
2. Publish packages as you make changes
3. Bitboard pulls from localhost:4873

**For production (Render):**
- **Best:** Deploy Verdaccio separately on Render with persistent disk
- **Alternative:** Commit built packages or use git URLs
- **Not recommended:** Run Verdaccio in same service as Bitboard (complicates deployment)

## Configuration Files

Add to Bitboard's `.gitignore`:
```
# Local Verdaccio config
.npmrc
```

Do NOT commit:
- `.npmrc` with localhost URLs
- Verdaccio storage directories

DO commit:
- `.npmrc` with `${VERDACCIO_URL}` env var for production
- Or have Render generate it during build

## Next Steps

1. ✅ Start Verdaccio locally
2. ✅ Publish Univer packages
3. ✅ Test installing in Bitboard locally
4. ⏳ Decide on production registry strategy
5. ⏳ Update Bitboard's package.json and test
