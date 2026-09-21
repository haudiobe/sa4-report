# Google Apps Script - Report Generation Project

## Project Overview
This project contains the Google Apps Script code for automated report generation related to 3GPP SA4 meeting documentation and data consolidation.

**Script ID:** `1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev`

**Original Document:** https://docs.google.com/document/d/1AKFPylxENnewf0gKPLeZiV2jVKOSWmJKd-X8mZTNvI8/edit

## Prerequisites

### 1. Install Node.js and npm
You need Node.js (which includes npm) to use clasp.

**Download and Install:**
- Visit: https://nodejs.org/
- Download the LTS (Long Term Support) version
- Run the installer and follow the prompts
- Restart your terminal/command prompt after installation

**Verify Installation:**
```bash
node --version
npm --version
```

### 2. Install clasp (Command Line Apps Script Projects)
Once Node.js is installed, install clasp globally:

```bash
npm install -g @google/clasp
```

**Verify Installation:**
```bash
clasp --version
```

## Setup Instructions

### Step 1: Enable Google Apps Script API
1. Visit: https://script.google.com/home/usersettings
2. Turn ON "Google Apps Script API"

### Step 2: Login to clasp
Authenticate clasp with your Google account:

```bash
clasp login
```

This will open a browser window for you to authorize clasp.

### Step 3: Clone the Apps Script Project
Navigate to this project directory and clone the script:

```bash
cd "c:\Users\tsto\OneDrive - Qualcomm\Projects\3GPP\Report"
clasp clone 1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev
```

This will download all script files to your local directory.

## Development Workflow

### Pull Latest Changes from Google
```bash
clasp pull
```

### Edit Files Locally
Edit the `.gs` files in your preferred editor (VS Code, etc.)

### Push Changes to Google
```bash
clasp push
```

### Open in Apps Script Editor
```bash
clasp open
```

### View Logs
```bash
clasp logs
```

## Common Issues and Solutions

### Issue: "npm is not recognized"
**Solution:** Node.js is not installed or not in PATH. Install Node.js and restart your terminal.

### Issue: "clasp: command not found"
**Solution:** clasp is not installed globally. Run: `npm install -g @google/clasp`

### Issue: "User has not enabled the Apps Script API"
**Solution:** Visit https://script.google.com/home/usersettings and enable the API.

## Next Steps

1. Complete Node.js installation
2. Install clasp
3. Clone the script project
4. Review the code for bugs and issues
5. Create test cases
6. Implement improvements

## Resources

- [clasp Documentation](https://github.com/google/clasp)
- [Apps Script Documentation](https://developers.google.com/apps-script)
- [Apps Script API Reference](https://developers.google.com/apps-script/reference)