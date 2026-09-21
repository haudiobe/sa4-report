# setup-clasp-env.ps1
# Helper script to configure environment for clasp in corporate networks
# Run this before using clasp commands

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  clasp Environment Setup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Disable SSL verification for corporate networks
$env:NODE_TLS_REJECT_UNAUTHORIZED="0"

Write-Host "SSL verification disabled for this session" -ForegroundColor Green
Write-Host "  (Required for corporate networks with SSL inspection)" -ForegroundColor Gray
Write-Host ""

# Display current directory
Write-Host "Current Directory:" -ForegroundColor Yellow
Write-Host "  $PWD" -ForegroundColor White
Write-Host ""

# Check if we're in the correct directory
$expectedPath = "c:\Users\tsto\OneDrive - Qualcomm\Projects\3GPP\Report"
if ($PWD.Path -ne $expectedPath) {
    Write-Host "You are not in the project directory" -ForegroundColor Yellow
    Write-Host "  Expected: $expectedPath" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Would you like to navigate there? (Y/N)" -ForegroundColor Yellow
    $response = Read-Host
    if ($response -eq "Y" -or $response -eq "y") {
        Set-Location $expectedPath
        Write-Host "Navigated to project directory" -ForegroundColor Green
        Write-Host ""
    }
}

# Check if clasp is installed
try {
    $claspVersion = clasp --version 2>&1
    Write-Host "clasp is installed: $claspVersion" -ForegroundColor Green
} catch {
    Write-Host "clasp is not installed" -ForegroundColor Red
    Write-Host "  Run: npm install -g @google/clasp" -ForegroundColor Gray
    Write-Host ""
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Available Commands" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Authentication:" -ForegroundColor Yellow
Write-Host "  clasp login                    # Login to Google" -ForegroundColor White
Write-Host ""
Write-Host "Project Setup:" -ForegroundColor Yellow
Write-Host "  clasp clone SCRIPT_ID          # Clone project" -ForegroundColor White
Write-Host ""
Write-Host "Daily Workflow:" -ForegroundColor Yellow
Write-Host "  clasp pull                     # Download latest from Google" -ForegroundColor White
Write-Host "  clasp push                     # Upload changes to Google" -ForegroundColor White
Write-Host "  clasp open                     # Open in browser" -ForegroundColor White
Write-Host "  clasp logs                     # View execution logs" -ForegroundColor White
Write-Host "  clasp status                   # Check what changed" -ForegroundColor White
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if project is already cloned
if (Test-Path ".clasp.json") {
    Write-Host "Project is already cloned" -ForegroundColor Green
    
    # Read and display script ID
    $claspConfig = Get-Content ".clasp.json" | ConvertFrom-Json
    Write-Host "  Script ID: $($claspConfig.scriptId)" -ForegroundColor Gray
    Write-Host ""
    
    Write-Host "Quick Actions:" -ForegroundColor Yellow
    Write-Host "  1. Pull latest:  clasp pull" -ForegroundColor White
    Write-Host "  2. Open editor:  clasp open" -ForegroundColor White
    Write-Host "  3. View logs:    clasp logs" -ForegroundColor White
} else {
    Write-Host "Project not yet cloned" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "To clone the project, run:" -ForegroundColor Yellow
    Write-Host "  clasp clone 1d2z2GPNXwAhVg1ylWJPU2SVyiSfC9b0pKKMDCTxZ238pz_De2qiWSCev" -ForegroundColor White
}

Write-Host ""
Write-Host "Environment is ready! You can now use clasp commands." -ForegroundColor Green
Write-Host ""