# Detailed Setup Guide for Google Apps Script Local Development

## Overview
This guide will walk you through setting up your local development environment for working with Google Apps Script using clasp (Command Line Apps Script Projects).

## Step-by-Step Installation

### Step 1: Install Node.js

#### Windows Installation:
1. **Download Node.js:**
   - Go to https://nodejs.org/
   - Download the **LTS (Long Term Support)** version (recommended)
   - The installer is typically named something like `node-v20.x.x-x64.msi`

2. **Run the Installer:**
   - Double-click the downloaded `.msi` file
   - Click "Next" through the installation wizard
   - Accept the license agreement
   - Keep the default installation path (usually `C:\Program Files\nodejs\`)
   - Make sure "Add to PATH" is checked
   - Click "Install"

3. **Verify Installation:**
   - Open a **NEW** Command Prompt or PowerShell window
   - Run these commands:
   ```bash
   node --version
   npm --version
   ```
   - You should see version numbers (e.g., `v20.11.0` and `10.2.4`)

**Important:** You must open a NEW terminal window after installation for the PATH changes to take effect.

### Step 2: Install clasp

Once Node.js is installed, install clasp globally:

```bash
npm install -g @google/clasp
```

**What this does:**
- `-g` flag installs clasp globally, making it available from any directory
- Downloads and installs the clasp command-line tool

**Verify Installation:**
```bash
clasp --version
```

You should see something like: `1.8.1` or similar.

### Step 3: Enable Google Apps Script API

Before you can use clasp, you must enable the Apps Script API:

1. **Visit the Settings Page:**
   - Go to: https://script.google.com/home/usersettings

2. **Enable the API:**
   - Find "Google Apps Script API"
   - Toggle it to **ON**

**Why this is needed:** clasp needs API access to interact with your Apps Script projects.

### Step 4: Authenticate clasp

Login to clasp with your Google account:

```bash
clasp login
```

**What happens:**
1. A browser window will open
2. You'll be asked to sign in to your Google account
3. You'll need to authorize clasp to access your Apps Script projects
4. After authorization, you'll see a success message in your terminal

**Troubleshooting:**
- If the browser doesn't open automatically, copy the URL from the terminal and paste it into your browser
- Make sure you're logging in with the Google account that owns the Apps Script project

### Step 5: Clone Your Apps Script Project

Navigate to your project directory and clone the script:

```bash
cd "c:\Users\tsto\OneDrive - Qualcomm\Projects\3GPP\Report"
clasp clone 1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev
```

**What this does:**
- Downloads all script files from Google Apps Script to your local directory
- Creates a `.clasp.json` file with project configuration
- Downloads `appsscript.json` (the manifest file)
- Downloads all `.gs` script files

**Expected Output:**
```
Cloned 3 files.
└─ appsscript.json
└─ Code.gs
└─ Utils.gs
```

## Project Structure After Cloning

After successful cloning, your directory will contain:

```
c:\Users\tsto\OneDrive - Qualcomm\Projects\3GPP\Report\
├── .clasp.json           # Clasp configuration (contains Script ID)
├── appsscript.json       # Apps Script manifest (runtime version, dependencies)
├── Code.gs               # Main script file (or similar name)
├── [Other .gs files]     # Additional script files
├── README.md             # Project documentation
└── SETUP_GUIDE.md        # This file
```

### Important Files:

**`.clasp.json`** - Contains your script ID and root directory:
```json
{
  "scriptId": "1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev",
  "rootDir": "."
}
```

**`appsscript.json`** - Manifest file with project settings:
```json
{
  "timeZone": "America/New_York",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

## Daily Development Workflow

### 1. Pull Latest Changes
Before starting work, pull the latest version from Google:

```bash
clasp pull
```

**When to use:** 
- At the start of your work session
- If you or someone else made changes in the online editor
- To sync your local files with the cloud version

### 2. Edit Files Locally
- Open files in VS Code or your preferred editor
- Make your changes to `.gs` files
- Save your changes

**Tips:**
- VS Code has good JavaScript support
- You can install Apps Script extensions for better syntax highlighting
- Use version control (git) to track your changes

### 3. Push Changes to Google
After making changes locally, push them to Google:

```bash
clasp push
```

**What this does:**
- Uploads all changed files to Google Apps Script
- Overwrites the online version with your local version

**Warning:** This will overwrite the online version. Make sure you've pulled the latest changes first!

### 4. Test Your Changes
Open the script in the online editor to test:

```bash
clasp open
```

**What this does:**
- Opens your Apps Script project in your default browser
- You can run functions and see execution logs
- You can use the debugger

### 5. View Execution Logs
To see logs from recent executions:

```bash
clasp logs
```

**Options:**
```bash
clasp logs --json          # Output in JSON format
clasp logs --open          # Open logs in browser
clasp logs --watch         # Watch logs in real-time
```

## Advanced clasp Commands

### Create a New Version
```bash
clasp version "Description of changes"
```

### Deploy as Web App or API
```bash
clasp deploy --description "Production deployment"
```

### List All Deployments
```bash
clasp deployments
```

### Run a Function
```bash
clasp run functionName
```

### Push Only Specific Files
```bash
clasp push --force
```

## Testing Strategies

### 1. Manual Testing in Apps Script Editor
- Push your changes: `clasp push`
- Open the editor: `clasp open`
- Run functions manually
- Check logs: `clasp logs`

### 2. Unit Testing (Advanced)
For local unit testing, you'll need to:
- Mock Google Apps Script services
- Use a testing framework like Jest
- Create test files in a `tests/` directory

### 3. Debugging
- Use `Logger.log()` in your code
- View logs with `clasp logs`
- Use the Apps Script debugger in the online editor

## Common Issues and Solutions

### Issue: "clasp: command not found" (Windows)
**Cause:** clasp is not in your PATH or not installed globally.

**Solution:**
1. Verify npm installation: `npm --version`
2. Reinstall clasp globally: `npm install -g @google/clasp`
3. Close and reopen your terminal
4. Try again: `clasp --version`

### Issue: "User has not enabled the Apps Script API"
**Cause:** The Apps Script API is not enabled for your account.

**Solution:**
1. Visit: https://script.google.com/home/usersettings
2. Toggle "Google Apps Script API" to ON
3. Try the clasp command again

### Issue: "No credentials found"
**Cause:** You haven't logged in with clasp.

**Solution:**
```bash
clasp login
```

### Issue: "Push failed" or "Pull failed"
**Cause:** Conflicts between local and remote versions.

**Solution:**
```bash
# Force pull to overwrite local files
clasp pull --force

# Or force push to overwrite remote files
clasp push --force
```

**Warning:** Force commands will overwrite files. Make backups first!

### Issue: "Script ID not found"
**Cause:** The `.clasp.json` file is missing or incorrect.

**Solution:**
1. Check if `.clasp.json` exists
2. Verify the scriptId is correct
3. Re-clone if necessary: `clasp clone SCRIPT_ID`

## Best Practices

### 1. Version Control
- Use git to track changes locally
- Create a `.gitignore` file:
```
.clasp.json
node_modules/
.DS_Store
```

### 2. Backup Before Major Changes
```bash
# Pull latest version
clasp pull

# Create a backup
cp Code.gs Code.gs.backup
```

### 3. Test Before Pushing
- Test changes locally if possible
- Use the online editor for final testing
- Check logs for errors

### 4. Document Your Code
- Add comments to explain complex logic
- Update README.md with new features
- Document any bugs you find

### 5. Incremental Changes
- Make small, focused changes
- Push and test frequently
- Don't accumulate too many changes before pushing

## Next Steps

Now that you have clasp set up:

1. ✅ Node.js installed
2. ✅ clasp installed
3. ✅ Apps Script API enabled
4. ✅ Authenticated with Google
5. ⏳ Clone the project
6. ⏳ Review the code
7. ⏳ Identify bugs
8. ⏳ Create tests
9. ⏳ Implement fixes

## Additional Resources

- **clasp GitHub:** https://github.com/google/clasp
- **Apps Script Guides:** https://developers.google.com/apps-script/guides
- **Apps Script Reference:** https://developers.google.com/apps-script/reference
- **Stack Overflow:** Search for "google-apps-script" tag

## Getting Help

If you encounter issues:
1. Check this guide's troubleshooting section
2. Review the clasp documentation
3. Search Stack Overflow
4. Check the Apps Script community forums

## Summary

You now have everything you need to:
- ✅ Work with Apps Script locally
- ✅ Edit code in your preferred editor
- ✅ Push and pull changes
- ✅ Test and debug your scripts
- ✅ Deploy updates

Happy coding!