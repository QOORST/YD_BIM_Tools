# Git Setup Script
# This script will configure Git and initialize the project repository

Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "  Git Setup and Initialization" -ForegroundColor Cyan
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host ""

# Set Git path
$gitPath = "C:\Program Files\Git\bin\git.exe"

# Check if Git is installed
if (-not (Test-Path $gitPath)) {
    Write-Host "[ERROR] Git not found at: $gitPath" -ForegroundColor Red
    Write-Host "Please make sure Git is installed correctly" -ForegroundColor Yellow
    exit 1
}

Write-Host "[OK] Git is installed: $gitPath" -ForegroundColor Green

# Show Git version
$gitVersion = & $gitPath --version
Write-Host "[OK] $gitVersion" -ForegroundColor Green
Write-Host ""

# Configure Git user information
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "  Step 1: Configure Git User Information" -ForegroundColor Cyan
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host ""

$userName = Read-Host "Enter your Git username (e.g., Your Name)"
$userEmail = Read-Host "Enter your Git email (e.g., your.email@example.com)"

& $gitPath config --global user.name "$userName"
& $gitPath config --global user.email "$userEmail"

Write-Host "[OK] Git user information configured" -ForegroundColor Green
Write-Host "  Username: $userName" -ForegroundColor White
Write-Host "  Email: $userEmail" -ForegroundColor White
Write-Host ""

# Check if already a Git repository
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "  Step 2: Check Git Repository Status" -ForegroundColor Cyan
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host ""

if (Test-Path ".git") {
    Write-Host "[OK] This project is already a Git repository" -ForegroundColor Green

    # Show remote repository
    Write-Host ""
    Write-Host "Remote repository:" -ForegroundColor Cyan
    & $gitPath remote -v

    Write-Host ""
    Write-Host "Current branch:" -ForegroundColor Cyan
    & $gitPath branch

} else {
    Write-Host "[INFO] This project is not yet initialized as a Git repository" -ForegroundColor Yellow
    Write-Host ""

    $initRepo = Read-Host "Do you want to initialize as a Git repository? (Y/N)"

    if ($initRepo -eq "Y" -or $initRepo -eq "y") {
        # Initialize Git repository
        & $gitPath init
        Write-Host "[OK] Git repository initialized" -ForegroundColor Green

        # Create .gitignore
        if (-not (Test-Path ".gitignore")) {
            Write-Host "[INFO] Creating .gitignore file..." -ForegroundColor Cyan

            $gitignoreContent = @"
# Build results
bin/
obj/
Output/
Installer/2024/
Installer/2025/
Installer/2026/
Installer/Resources/
Installer/*.dll
Installer/*.txt

# Visual Studio
.vs/
*.user
*.suo
*.userosscache
*.sln.docstates

# ReSharper
_ReSharper*/
*.DotSettings.user

# User-specific files
*.rsuser
*.suo
*.user
*.userosscache
*.sln.docstates

# Logs
*.log
LOG/

# Temporary files
*.tmp
*.temp
"@
            $gitignoreContent | Out-File -FilePath ".gitignore" -Encoding UTF8

            Write-Host "[OK] .gitignore created" -ForegroundColor Green
        }

        Write-Host ""
        $addRemote = Read-Host "Do you want to add a remote GitHub repository? (Y/N)"

        if ($addRemote -eq "Y" -or $addRemote -eq "y") {
            $remoteUrl = Read-Host "Enter GitHub repository URL (e.g., https://github.com/username/repo.git)"
            & $gitPath remote add origin $remoteUrl
            Write-Host "[OK] Remote repository added: $remoteUrl" -ForegroundColor Green
        }
    }
}

Write-Host ""
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "  Git Setup Complete!" -ForegroundColor Cyan
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host ""
Write-Host "1. Run the publish script:" -ForegroundColor White
Write-Host "   .\publish_v2.2.4.ps1" -ForegroundColor Cyan
Write-Host ""
Write-Host "2. Or manually run Git commands:" -ForegroundColor White
Write-Host "   & ""$gitPath"" status" -ForegroundColor Cyan
Write-Host "   & ""$gitPath"" add ." -ForegroundColor Cyan
Write-Host "   & ""$gitPath"" commit -F COMMIT_MESSAGE_v2.2.4.txt" -ForegroundColor Cyan
Write-Host "   & ""$gitPath"" tag -a v2.2.4 -m ""Release v2.2.4""" -ForegroundColor Cyan
Write-Host "   & ""$gitPath"" push origin main" -ForegroundColor Cyan
Write-Host "   & ""$gitPath"" push origin v2.2.4" -ForegroundColor Cyan
Write-Host ""

