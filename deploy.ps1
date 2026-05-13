$env:PATH = "$env:APPDATA\npm;" + $env:PATH
Set-Location $PSScriptRoot
Write-Host "Deploying..."
& clasp push
if ($LASTEXITCODE -ne 0) { Write-Host "ERROR: push failed"; cmd /c pause; exit 1 }
Write-Host "Updating deployment..."
& clasp deploy -i "AKfycbw5MLHFep_jOQEdAg4_wX8LMGPX7wVL41XbmygqVV794LkZu6Xv-XcRLNAHYqg9bd0fyw" -d "update"
if ($LASTEXITCODE -ne 0) { Write-Host "ERROR: deploy failed"; cmd /c pause; exit 1 }
Write-Host "Deploy Complete!"
cmd /c pause