# Corporate Network Setup Guide

## Issue: SSL Certificate Errors with clasp

You're seeing this error:
```
request to https://oauth2.googleapis.com/token failed, reason: self-signed certificate in certificate chain
```

This occurs because your corporate network uses a proxy with SSL inspection (self-signed certificates).

---

## Solution Options

### Option 1: Disable SSL Verification (Quick Fix - Not Recommended for Production)

This is the fastest solution but less secure. Use only for development.

#### For Current Session:
```bash
# Windows PowerShell
$env:NODE_TLS_REJECT_UNAUTHORIZED="0"
clasp login
clasp clone 1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev
```

#### For Current Session (Command Prompt):
```bash
set NODE_TLS_REJECT_UNAUTHORIZED=0
clasp login
clasp clone 1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev
```

#### Permanently (Add to System Environment Variables):
1. Press `Win + X` and select "System"
2. Click "Advanced system settings"
3. Click "Environment Variables"
4. Under "User variables", click "New"
5. Variable name: `NODE_TLS_REJECT_UNAUTHORIZED`
6. Variable value: `0`
7. Click OK
8. Restart your terminal

---

### Option 2: Configure Corporate Certificate (Recommended)

If your IT department provides a corporate root certificate:

#### Step 1: Get the Certificate
Contact your IT department for:
- Corporate root CA certificate (usually a `.pem` or `.crt` file)
- Or the correct path to the certificate

#### Step 2: Set the Certificate Path
```bash
# Windows PowerShell
$env:NODE_EXTRA_CA_CERTS="C:\path\to\your\corporate-cert.pem"

# Or permanently via Environment Variables:
# Variable name: NODE_EXTRA_CA_CERTS
# Variable value: C:\path\to\your\corporate-cert.pem
```

#### Step 3: Try clasp Again
```bash
clasp login
clasp clone 1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev
```

---

### Option 3: Configure npm to Use Corporate Proxy

If your network requires a proxy:

```bash
# Set proxy settings
npm config set proxy http://proxy.company.com:8080
npm config set https-proxy http://proxy.company.com:8080

# If proxy requires authentication:
npm config set proxy http://username:password@proxy.company.com:8080
npm config set https-proxy http://username:password@proxy.company.com:8080

# Disable strict SSL (if needed)
npm config set strict-ssl false
```

---

## Step-by-Step: Quick Fix Method

### 1. Open PowerShell as Administrator
Right-click PowerShell and select "Run as Administrator"

### 2. Navigate to Project Directory
```powershell
cd "c:\Users\tsto\OneDrive - Qualcomm\Projects\3GPP\Report"
```

### 3. Disable SSL Verification for This Session
```powershell
$env:NODE_TLS_REJECT_UNAUTHORIZED="0"
```

### 4. Login to clasp
```powershell
clasp login
```

This will open a browser window. Authorize the application.

### 5. Clone the Project
```powershell
clasp clone 1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev
```

### 6. Verify Files Were Downloaded
```powershell
dir
```

You should see `.clasp.json`, `appsscript.json`, and `.gs` files.

---

## Permanent Solution Script

Create a file called `setup-clasp-env.ps1`:

```powershell
# setup-clasp-env.ps1
# Run this before using clasp commands

# Disable SSL verification (use with caution)
$env:NODE_TLS_REJECT_UNAUTHORIZED="0"

Write-Host "Environment configured for clasp"
Write-Host "SSL verification disabled for this session"
Write-Host ""
Write-Host "You can now run clasp commands:"
Write-Host "  clasp pull"
Write-Host "  clasp push"
Write-Host "  clasp open"
Write-Host "  clasp logs"
```

**Usage:**
```powershell
# Run this before using clasp
. .\setup-clasp-env.ps1

# Then use clasp normally
clasp pull
clasp push
```

---

## Alternative: Use clasp with Git Bash

If you have Git for Windows installed, Git Bash may handle certificates better:

1. Open Git Bash
2. Navigate to your project:
   ```bash
   cd "/c/Users/tsto/OneDrive - Qualcomm/Projects/3GPP/Report"
   ```
3. Set environment variable:
   ```bash
   export NODE_TLS_REJECT_UNAUTHORIZED=0
   ```
4. Run clasp commands:
   ```bash
   clasp login
   clasp clone 1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev
   ```

---

## Troubleshooting

### Issue: "Warning: Ignoring extra certs from..."
This warning can be ignored. It's just informing you that clasp tried to load a certificate but couldn't find it.

### Issue: Browser doesn't open for login
Copy the URL from the terminal and paste it into your browser manually.

### Issue: Still getting certificate errors
Try all three options in order:
1. Disable SSL verification
2. Configure corporate certificate
3. Configure proxy settings

### Issue: Permission denied
Run PowerShell as Administrator.

---

## Security Considerations

**Important:** Disabling SSL verification (`NODE_TLS_REJECT_UNAUTHORIZED=0`) makes your connection less secure. 

**Best practices:**
1. Only use this in development environments
2. Only disable for the current session, not permanently
3. Contact your IT department for the proper corporate certificate
4. Re-enable SSL verification when not using clasp

---

## Working with Corporate Networks - Daily Workflow

### Morning Setup (PowerShell)
```powershell
# 1. Navigate to project
cd "c:\Users\tsto\OneDrive - Qualcomm\Projects\3GPP\Report"

# 2. Set environment for this session
$env:NODE_TLS_REJECT_UNAUTHORIZED="0"

# 3. Pull latest changes
clasp pull

# 4. Open in VS Code
code .
```

### Before Each clasp Command
If you open a new terminal, remember to set the environment variable:
```powershell
$env:NODE_TLS_REJECT_UNAUTHORIZED="0"
```

---

## Quick Reference

### PowerShell Commands
```powershell
# Set SSL bypass for current session
$env:NODE_TLS_REJECT_UNAUTHORIZED="0"

# Navigate to project
cd "c:\Users\tsto\OneDrive - Qualcomm\Projects\3GPP\Report"

# clasp commands
clasp login
clasp clone 1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev
clasp pull
clasp push
clasp open
clasp logs
```

### Command Prompt Commands
```cmd
# Set SSL bypass for current session
set NODE_TLS_REJECT_UNAUTHORIZED=0

# Navigate to project
cd "c:\Users\tsto\OneDrive - Qualcomm\Projects\3GPP\Report"

# clasp commands (same as above)
```

---

## Contact IT Support

If you need the proper solution, ask your IT department for:
1. Corporate root CA certificate file
2. Proxy server address and port
3. Proxy authentication credentials (if required)
4. Any specific npm or Node.js configuration required

---

## Next Steps

After resolving the certificate issue:

1. ✅ Complete clasp login
2. ✅ Clone the project
3. ✅ Verify files downloaded
4. ✅ Review the code
5. ✅ Identify bugs
6. ✅ Start development

---

## Summary

**Immediate Solution:**
```powershell
$env:NODE_TLS_REJECT_UNAUTHORIZED="0"
clasp login
clasp clone 1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev
```

**Long-term Solution:**
Contact IT for corporate certificate and configure `NODE_EXTRA_CA_CERTS`.

You should now be able to proceed with cloning your Apps Script project!