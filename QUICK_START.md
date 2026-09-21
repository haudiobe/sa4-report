# Quick Start Guide

## For First-Time Setup

### Prerequisites Check
Before starting, verify you have:
- [ ] Windows 10 or 11
- [ ] Internet connection
- [ ] Google account with access to the Apps Script project
- [ ] Administrator rights (for installing Node.js)

### Installation Steps (15-20 minutes)

#### 1. Install Node.js (5 minutes)
```bash
# Download from: https://nodejs.org/
# Install the LTS version
# Restart your terminal after installation

# Verify installation:
node --version
npm --version
```

#### 2. Install clasp (2 minutes)
```bash
npm install -g @google/clasp

# Verify installation:
clasp --version
```

#### 3. Enable Apps Script API (1 minute)
- Visit: https://script.google.com/home/usersettings
- Toggle "Google Apps Script API" to **ON**

#### 4. Login to clasp (2 minutes)
```bash
clasp login
# Follow the browser prompts to authorize
```

#### 5. Clone Your Project (2 minutes)
```bash
cd "c:\Users\tsto\OneDrive - Qualcomm\Projects\3GPP\Report"
clasp clone 1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev
```

**Done!** Your Apps Script files are now in your local directory.

---

## Daily Workflow

### Morning Routine
```bash
# 1. Pull latest changes
clasp pull

# 2. Open in VS Code
code .
```

### Making Changes
```bash
# 1. Edit files in VS Code
# 2. Save your changes

# 3. Push to Google
clasp push

# 4. Test in browser
clasp open
```

### Checking Logs
```bash
# View recent execution logs
clasp logs

# Watch logs in real-time
clasp logs --watch
```

---

## Common Commands Cheat Sheet

### Essential Commands
```bash
clasp pull              # Download latest from Google
clasp push              # Upload changes to Google
clasp open              # Open in browser
clasp logs              # View execution logs
clasp status            # Check what's changed
```

### Advanced Commands
```bash
clasp push --force      # Force push (overwrites remote)
clasp pull --force      # Force pull (overwrites local)
clasp version           # Create a version
clasp deploy            # Deploy the script
clasp deployments       # List all deployments
```

---

## Troubleshooting Quick Fixes

### "npm is not recognized"
**Fix:** Install Node.js and restart terminal

### "clasp: command not found"
**Fix:** Run `npm install -g @google/clasp` and restart terminal

### "User has not enabled the Apps Script API"
**Fix:** Visit https://script.google.com/home/usersettings and enable it

### "No credentials found"
**Fix:** Run `clasp login`

### Push/Pull conflicts
**Fix:** 
```bash
# To keep remote version:
clasp pull --force

# To keep local version:
clasp push --force
```

---

## Project Structure

```
c:\Users\tsto\OneDrive - Qualcomm\Projects\3GPP\Report\
├── .clasp.json              # Project configuration
├── appsscript.json          # Apps Script manifest
├── Code.gs                  # Main script file
├── [Other .gs files]        # Additional scripts
├── README.md                # Full documentation
├── SETUP_GUIDE.md           # Detailed setup guide
├── QUICK_START.md           # This file
├── BUGS_AND_ISSUES.md       # Bug tracker
└── .gitignore               # Git ignore rules
```

---

## Next Steps After Setup

1. **Review the Code**
   ```bash
   # Open all files in VS Code
   code .
   ```

2. **Test Existing Functions**
   ```bash
   clasp open
   # Run functions in the Apps Script editor
   ```

3. **Check for Issues**
   - Review BUGS_AND_ISSUES.md
   - Test all main functions
   - Document any problems you find

4. **Make Your First Change**
   - Edit a file locally
   - Push with `clasp push`
   - Test in browser with `clasp open`

---

## Getting Help

### Documentation
- **Full Setup:** See SETUP_GUIDE.md
- **Project Info:** See README.md
- **Bug Tracking:** See BUGS_AND_ISSUES.md

### Online Resources
- clasp: https://github.com/google/clasp
- Apps Script: https://developers.google.com/apps-script

### Common Questions

**Q: Can I work offline?**
A: You can edit files offline, but need internet to push/pull.

**Q: Will my changes affect the live script immediately?**
A: Yes, `clasp push` updates the script immediately.

**Q: How do I undo a bad push?**
A: Use `clasp pull --force` to restore from Google, or use git to revert.

**Q: Can multiple people work on the same script?**
A: Yes, but coordinate to avoid conflicts. Always pull before pushing.

---

## Safety Tips

1. **Always pull before pushing**
   ```bash
   clasp pull
   # Make changes
   clasp push
   ```

2. **Test before pushing**
   - Test locally if possible
   - Review your changes
   - Push small, incremental changes

3. **Backup important changes**
   ```bash
   # Create a backup
   cp Code.gs Code.gs.backup
   ```

4. **Use version control**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   ```

---

## Success Checklist

After completing setup, you should be able to:
- ✅ Run `clasp --version` successfully
- ✅ See `.clasp.json` in your project directory
- ✅ See `.gs` files in your project directory
- ✅ Run `clasp open` to view in browser
- ✅ Edit files in VS Code
- ✅ Push changes with `clasp push`
- ✅ View logs with `clasp logs`

If all items are checked, you're ready to start developing!

---

## Quick Reference Card

Print or save this for easy reference:

```
┌─────────────────────────────────────────┐
│     CLASP QUICK REFERENCE               │
├─────────────────────────────────────────┤
│ Setup:                                  │
│   npm install -g @google/clasp          │
│   clasp login                           │
│   clasp clone SCRIPT_ID                 │
│                                         │
│ Daily Use:                              │
│   clasp pull        # Get latest        │
│   clasp push        # Upload changes    │
│   clasp open        # Open in browser   │
│   clasp logs        # View logs         │
│                                         │
│ Troubleshooting:                        │
│   clasp pull --force                    │
│   clasp push --force                    │
│   clasp login                           │
│                                         │
│ Script ID:                              │
│   1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0p... │
└─────────────────────────────────────────┘
```

---

**Ready to start? Follow the Installation Steps above!**