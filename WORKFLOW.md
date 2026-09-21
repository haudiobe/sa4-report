# Complete Development Workflow

## Overview
This document provides a comprehensive workflow for developing, testing, and deploying changes to the Google Apps Script project.

---

## Initial Setup (One-Time)

### Step 1: Install Prerequisites
```bash
# 1. Download and install Node.js from https://nodejs.org/
# 2. Restart your terminal
# 3. Verify installation
node --version
npm --version
```

### Step 2: Install and Configure clasp
```bash
# Install clasp globally
npm install -g @google/clasp

# Verify installation
clasp --version

# Enable Apps Script API
# Visit: https://script.google.com/home/usersettings
# Toggle "Google Apps Script API" to ON

# Login to clasp
clasp login
```

### Step 3: Clone the Project
```bash
# Navigate to project directory
cd "c:\Users\tsto\OneDrive - Qualcomm\Projects\3GPP\Report"

# Clone the Apps Script project
clasp clone 1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev

# Verify files were downloaded
dir
```

### Step 4: Initialize Version Control (Optional but Recommended)
```bash
# Initialize git repository
git init

# Add files
git add .

# Create initial commit
git commit -m "Initial commit: Apps Script project setup"
```

---

## Daily Development Workflow

### Morning Routine

```bash
# 1. Navigate to project directory
cd "c:\Users\tsto\OneDrive - Qualcomm\Projects\3GPP\Report"

# 2. Pull latest changes from Google
clasp pull

# 3. Check status
clasp status

# 4. Open project in VS Code
code .
```

### Making Changes

#### Step 1: Plan Your Changes
- Review BUGS_AND_ISSUES.md for known issues
- Identify what needs to be changed
- Consider impact on other parts of the code

#### Step 2: Create a Backup (for major changes)
```bash
# Create backup of files you'll modify
copy Code.gs Code.gs.backup
```

#### Step 3: Make Your Changes
- Edit files in VS Code
- Add comments to explain your changes
- Follow coding best practices
- Save your changes

#### Step 4: Review Your Changes
```bash
# Check what files changed
clasp status

# Review changes in VS Code
# Use Source Control view to see diffs
```

#### Step 5: Test Locally (if possible)
- Review your code for syntax errors
- Check for logical errors
- Verify variable names and function calls

### Pushing Changes

#### Step 1: Push to Google
```bash
# Push your changes
clasp push

# If there are conflicts, you may need to force push
# (Be careful with this!)
clasp push --force
```

#### Step 2: Test in Apps Script Editor
```bash
# Open the script in browser
clasp open

# In the Apps Script editor:
# 1. Select the function to test
# 2. Click "Run"
# 3. Check for errors
# 4. Review execution logs
```

#### Step 3: Check Logs
```bash
# View recent execution logs
clasp logs

# Watch logs in real-time (in a separate terminal)
clasp logs --watch
```

#### Step 4: Verify Results
- Check that the function executed successfully
- Verify the output is correct
- Test edge cases
- Document any issues found

### Committing Changes (if using git)

```bash
# Stage your changes
git add .

# Commit with a descriptive message
git commit -m "Fix: Description of what you fixed"

# Examples of good commit messages:
# git commit -m "Fix: Handle null values in data processing"
# git commit -m "Add: New function for report generation"
# git commit -m "Update: Improve error handling in main function"
```

---

## Testing Workflow

### Manual Testing Process

#### 1. Prepare Test Environment
```bash
# Ensure you have latest code
clasp pull

# Push your changes
clasp push

# Open in browser
clasp open
```

#### 2. Execute Tests
In the Apps Script editor:
1. Select function to test
2. Click "Run"
3. Grant permissions if prompted
4. Wait for execution to complete
5. Review execution logs

#### 3. Document Results
Update BUGS_AND_ISSUES.md with:
- What you tested
- What worked
- What didn't work
- Any errors encountered

#### 4. Check Logs
```bash
# View detailed logs
clasp logs

# Look for:
# - Error messages
# - Unexpected behavior
# - Performance issues
```

### Testing Checklist

For each function you test:
- [ ] Test with valid inputs
- [ ] Test with invalid inputs
- [ ] Test with edge cases (empty, null, very large)
- [ ] Test error handling
- [ ] Check execution time
- [ ] Verify output format
- [ ] Document results

---

## Bug Fixing Workflow

### Step 1: Identify the Bug
- Reproduce the issue
- Document steps to reproduce
- Check logs for error messages
- Add to BUGS_AND_ISSUES.md

### Step 2: Investigate
```bash
# Pull latest code
clasp pull

# Open in VS Code
code .

# Review the relevant code
# Add Logger.log() statements for debugging
```

### Step 3: Fix the Bug
- Make minimal changes to fix the issue
- Add comments explaining the fix
- Consider edge cases
- Update error handling if needed

### Step 4: Test the Fix
```bash
# Push changes
clasp push

# Open in browser
clasp open

# Test the specific scenario that was failing
# Test related scenarios
# Check logs
clasp logs
```

### Step 5: Document the Fix
- Update BUGS_AND_ISSUES.md
- Mark the issue as fixed
- Document the solution
- Commit changes (if using git)

---

## Deployment Workflow

### Pre-Deployment Checklist
- [ ] All tests pass
- [ ] No known critical bugs
- [ ] Code is documented
- [ ] Changes are committed (if using git)
- [ ] Backup of current version exists

### Deployment Steps

#### 1. Create a Version
```bash
# Create a new version with description
clasp version "Description of changes in this version"
```

#### 2. Deploy
```bash
# Create a new deployment
clasp deploy --description "Production deployment - [Date]"

# List all deployments
clasp deployments
```

#### 3. Verify Deployment
- Test the deployed version
- Verify all functionality works
- Check logs for any issues

#### 4. Document Deployment
- Update README.md with version info
- Document any breaking changes
- Notify users if necessary

---

## Rollback Procedure

If something goes wrong:

### Option 1: Revert to Previous Version
```bash
# Pull the last known good version
clasp pull --force

# Push it back
clasp push --force
```

### Option 2: Use Version History
1. Open script in browser: `clasp open`
2. Go to File > Version history
3. Select a previous version
4. Restore it

### Option 3: Use Git (if configured)
```bash
# Revert to previous commit
git revert HEAD

# Push the reverted version
clasp push
```

---

## Collaboration Workflow

### Working with Others

#### Before Starting Work
```bash
# Always pull latest changes first
clasp pull

# Check what changed
git log  # if using git
```

#### While Working
- Communicate with team about what you're working on
- Make small, focused changes
- Push frequently
- Document your changes

#### Resolving Conflicts
```bash
# If someone else pushed changes:

# Option 1: Pull and merge
clasp pull
# Manually resolve conflicts in VS Code
clasp push

# Option 2: Force pull (loses your local changes)
clasp pull --force

# Option 3: Force push (overwrites remote changes)
# Only use if you're sure!
clasp push --force
```

---

## Best Practices

### Code Quality
1. **Write Clear Code**
   - Use descriptive variable names
   - Add comments for complex logic
   - Keep functions small and focused

2. **Error Handling**
   ```javascript
   function myFunction() {
     try {
       // Your code here
     } catch (error) {
       Logger.log('Error in myFunction: ' + error.message);
       // Handle error appropriately
     }
   }
   ```

3. **Logging**
   ```javascript
   Logger.log('Starting process...');
   Logger.log('Processing item: ' + itemName);
   Logger.log('Result: ' + JSON.stringify(result));
   ```

### Version Control
1. **Commit Often**
   - Commit after each logical change
   - Use descriptive commit messages
   - Don't commit broken code

2. **Use Branches** (if using git)
   ```bash
   # Create a feature branch
   git checkout -b feature/new-report-format
   
   # Work on your feature
   # ...
   
   # Merge back to main
   git checkout main
   git merge feature/new-report-format
   ```

### Testing
1. **Test Before Pushing**
   - Review your changes
   - Test locally if possible
   - Think about edge cases

2. **Test After Pushing**
   - Always test in the Apps Script editor
   - Check logs for errors
   - Verify expected behavior

### Documentation
1. **Keep Documentation Updated**
   - Update README.md with new features
   - Document bugs in BUGS_AND_ISSUES.md
   - Add comments to complex code

2. **Document Decisions**
   - Why you made certain choices
   - Trade-offs considered
   - Alternative approaches

---

## Troubleshooting Common Issues

### Issue: Changes Not Appearing
**Solution:**
```bash
# Force push
clasp push --force

# Clear browser cache
# Refresh the Apps Script editor
```

### Issue: Permission Errors
**Solution:**
```bash
# Re-authenticate
clasp login

# Check script permissions in Google
```

### Issue: Execution Timeout
**Solution:**
- Break long operations into smaller chunks
- Use time-based triggers for long processes
- Optimize your code

### Issue: Quota Exceeded
**Solution:**
- Check quota usage in Apps Script dashboard
- Optimize API calls
- Implement caching
- Use batch operations

---

## Quick Reference

### Essential Commands
```bash
clasp pull              # Get latest from Google
clasp push              # Upload to Google
clasp open              # Open in browser
clasp logs              # View logs
clasp status            # Check changes
```

### File Locations
- **Project Root:** `c:\Users\tsto\OneDrive - Qualcomm\Projects\3GPP\Report`
- **Script ID:** `1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev`
- **Documentation:** README.md, SETUP_GUIDE.md, QUICK_START.md
- **Bug Tracking:** BUGS_AND_ISSUES.md
- **Tests:** tests/README.md

### Important Links
- **Apps Script API Settings:** https://script.google.com/home/usersettings
- **Apps Script Editor:** Use `clasp open` command
- **clasp Documentation:** https://github.com/google/clasp
- **Apps Script Docs:** https://developers.google.com/apps-script

---

## Summary

This workflow ensures:
- ✅ Code is properly version controlled
- ✅ Changes are tested before deployment
- ✅ Issues are documented and tracked
- ✅ Team collaboration is smooth
- ✅ Rollback is possible if needed

Follow this workflow consistently for best results!