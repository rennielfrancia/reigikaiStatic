# TESTED Deployment Script
# Updates footer timestamps only (no commit hash to prevent infinite loops)

Write-Host "Deploying website updates..." -ForegroundColor Green

# Get current date/time
$currentTime = Get-Date -Format "MMM dd, yyyy 'at' hh:mm tt"
Write-Host "Updating footer timestamps to: $currentTime" -ForegroundColor Cyan

# List of HTML files to update
$htmlFiles = @("index.html", "2025-Las-Vegas-Open.html", "About.html", "Contact.html", "Private-Lessons.html", "Tournament.html")

# Track if any files were modified
$modified = $false

# Update each HTML file
foreach ($file in $htmlFiles) {
    if (Test-Path $file) {
        $content = Get-Content $file -Raw
        $newContent = $content -replace 'Last updated: [^<]*</p>', "Last updated: $currentTime</p>"
        
        if ($content -ne $newContent) {
            Set-Content $file $newContent -NoNewline
            $modified = $true
            Write-Host "Updated $file" -ForegroundColor Yellow
        }
    }
}

if ($modified) {
    Write-Host "Footer timestamps updated!" -ForegroundColor Green
    Write-Host "Committing changes..." -ForegroundColor Cyan
    
    git add .
    git commit -m "Update footer timestamps - $currentTime"
    git push
    
    Write-Host "Deployment complete!" -ForegroundColor Green
} else {
    Write-Host "No files needed updating." -ForegroundColor Yellow
} 