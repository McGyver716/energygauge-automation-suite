# GitHub Authentication Setup Guide

## Option 1: Personal Access Token (Recommended)

### Step 1: Create a Personal Access Token
1. Go to GitHub.com and log into your account
2. Click your profile picture → **Settings**
3. Scroll down to **Developer settings** (left sidebar)
4. Click **Personal access tokens** → **Tokens (classic)**
5. Click **Generate new token** → **Generate new token (classic)**
6. Fill out the form:
   - **Note**: "EnergyGauge Automation Suite"
   - **Expiration**: 90 days (or your preference)
   - **Scopes**: Check "repo" (this gives full repository access)
7. Click **Generate token**
8. **COPY THE TOKEN** - you won't see it again!

### Step 2: Use Token to Push
Open Command Prompt or PowerShell in your project directory and run:

```bash
git push -u origin main
```

When prompted:
- **Username**: McGyver716
- **Password**: [paste your personal access token]

---

## Option 2: GitHub CLI (Alternative)

### Install GitHub CLI
1. Download from: https://cli.github.com/
2. Install and restart terminal
3. Run: `gh auth login`
4. Follow prompts to authenticate
5. Then run: `git push -u origin main`

---

## Option 3: SSH Key (Most Secure)

### Step 1: Generate SSH Key
```bash
ssh-keygen -t ed25519 -C "your.email@example.com"
```

### Step 2: Add to GitHub
1. Copy the public key: `cat ~/.ssh/id_ed25519.pub`
2. Go to GitHub → Settings → SSH and GPG keys
3. Click "New SSH key" and paste the key

### Step 3: Change Remote URL
```bash
git remote set-url origin git@github.com:McGyver716/energygauge-automation-suite.git
git push -u origin main
```

---

## Quick Fix: Force Push with Credentials

If you just want to get it uploaded quickly:

```bash
git push https://McGyver716:YOUR_TOKEN@github.com/McGyver716/energygauge-automation-suite.git main
```

Replace `YOUR_TOKEN` with your personal access token.

---

## What's Ready to Upload:

✅ Complete automation engine (2,226 lines)
✅ Modern GUI with 4 specialized tabs  
✅ OCR processing system
✅ COM interface automation
✅ Windows deployment scripts
✅ Comprehensive documentation
✅ Sample data and templates
✅ Professional Git configuration

The repository is fully prepared and committed locally. Once authenticated, everything will upload to:
**https://github.com/McGyver716/energygauge-automation-suite**