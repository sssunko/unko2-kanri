$env:PATH = "$env:APPDATA\npm;" + $env:PATH
Set-Location $PSScriptRoot
$PROD_ID = "AKfycbw7rzkd_SuE1I6BNzEjED4Mxl6cnM4wbswIiRiNoPf5zcSS2JcP6YLkfRV21fLc0opU"

Write-Host "Pushing..."
& clasp push --force
if ($LASTEXITCODE -ne 0) { Write-Host "ERROR: push failed"; cmd /c pause; exit 1 }

Write-Host "Updating deployment..."
& clasp deploy -i $PROD_ID -d "update"
if ($LASTEXITCODE -ne 0) { Write-Host "ERROR: deploy failed"; cmd /c pause; exit 1 }

Write-Host "Cleaning old deployments..."
$lines = & clasp deployments 2>&1
foreach ($line in $lines) {
    if ($line -match "^- (AKfycb\S+)") {
        $id = $Matches[1]
        if ($id -ne $PROD_ID -and $line -notmatch "@HEAD") {
            Write-Host "Removing old: $id"
            & clasp undeploy $id
        }
    }
}

Write-Host "Deploy Complete!"
cmd /c pause
