# NPOIPlus Push to GitHub Script
$ErrorActionPreference = 'Stop'

Write-Host "NPOIPlus - Pushing to GitHub" -ForegroundColor Cyan
Write-Host ""

# Check if Git is initialized
if (-not (Test-Path ".git")) {
    Write-Host "Initializing Git repository..." -ForegroundColor Yellow
    git init
    Write-Host "Done" -ForegroundColor Green
}

# Configure Git user
Write-Host "Configuring Git user..." -ForegroundColor Yellow
git config user.name "HouseAlwaysWin"
git config user.email "martinwang7963@gmail.com"
Write-Host "Done" -ForegroundColor Green
Write-Host ""

# Add all files
Write-Host "Adding files..." -ForegroundColor Yellow
git add .
Write-Host "Done" -ForegroundColor Green

# Commit
Write-Host "Committing..." -ForegroundColor Yellow
git commit -m "feat: initial commit with complete CI/CD setup"
Write-Host "Done" -ForegroundColor Green
Write-Host ""

# Check remote
$remoteExists = git remote | Select-String "origin"
if (-not $remoteExists) {
    Write-Host "Adding remote..." -ForegroundColor Yellow
    git remote add origin https://github.com/HouseAlwaysWin/NPOIPlus.git
    Write-Host "Done" -ForegroundColor Green
}

# Set main branch
Write-Host "Setting main branch..." -ForegroundColor Yellow
git branch -M main
Write-Host "Done" -ForegroundColor Green
Write-Host ""

# Show remote info
Write-Host "Remote repository:" -ForegroundColor Yellow
git remote -v
Write-Host ""

# Push
Write-Host "Ready to push to GitHub" -ForegroundColor Cyan
Write-Host "Repository: https://github.com/HouseAlwaysWin/NPOIPlus.git" -ForegroundColor Cyan
Write-Host ""

$response = Read-Host "Continue? (y/n)"

if ($response -eq 'y' -or $response -eq 'Y') {
    Write-Host "Pushing..." -ForegroundColor Yellow
    git push -u origin main
    
    Write-Host ""
    Write-Host "SUCCESS!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "1. View repository: https://github.com/HouseAlwaysWin/NPOIPlus" -ForegroundColor White
    Write-Host "2. View CI status: https://github.com/HouseAlwaysWin/NPOIPlus/actions" -ForegroundColor White
    Write-Host "3. IMPORTANT - Set NuGet API Key:" -ForegroundColor Red
    Write-Host "   https://github.com/HouseAlwaysWin/NPOIPlus/settings/secrets/actions" -ForegroundColor White
    Write-Host "   Secret name: NUGET_API_KEY" -ForegroundColor White
} else {
    Write-Host "Cancelled" -ForegroundColor Yellow
}
