# GitHub Email Authentication Steps

## Important: GitHub Authentication Methods

GitHub **does not** use email/password for Git operations anymore (as of 2021). Here are the current methods:

## Method 1: Personal Access Token (What you need)

### Step 1: Get Your Personal Access Token
1. Go to **GitHub.com** and sign in with your email/password
2. Click your **profile picture** (top right) → **Settings**
3. Scroll down to **Developer settings** (left sidebar)
4. Click **Personal access tokens** → **Tokens (classic)**
5. Click **Generate new token** → **Generate new token (classic)**

### Step 2: Configure the Token
- **Note**: "EnergyGauge Automation Suite"
- **Expiration**: 90 days
- **Scopes**: ✅ Check **"repo"** (full repository access)
- Click **Generate token**
- **COPY THE TOKEN** immediately (you can't see it again!)

### Step 3: Use Token to Push Code
In your project directory, run:
```bash
git push -u origin main
```

When prompted:
- **Username**: McGyver716
- **Password**: [paste the token you copied]

## Method 2: Configure Git with Your Email (Setup only)

First, configure Git with your email:
```bash
git config --global user.email "your.email@example.com"
git config --global user.name "McGyver716"
```

But you'll still need the Personal Access Token from Method 1 to actually push.

## Why Email/Password Doesn't Work:

GitHub stopped accepting passwords for Git operations in August 2021 for security reasons. You must use:
- Personal Access Token (easiest)
- SSH keys (most secure)
- GitHub CLI (most convenient)

## Quick Summary:

1. **Sign in to GitHub.com** with your email/password
2. **Generate a Personal Access Token** (replaces your password for Git)
3. **Use that token** when Git asks for your password

Your complete automation suite is ready to upload once you get the token!