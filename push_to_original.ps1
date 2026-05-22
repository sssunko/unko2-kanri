$DEPLOY_ID = "AKfycbw5MLHFep_jOQEdAg4_wX8LMGPX7wVL41XbmygqVV794LkZu6Xv-XcRLNAHYqg9bd0fyw"
$MAX_KEEP  = 5

Write-Host "[1/3] Push code..."
clasp push --force
if ($LASTEXITCODE -ne 0) { Write-Host "Push failed."; exit 1 }

Write-Host "[2/3] Update deploy..."
clasp deploy --deploymentId $DEPLOY_ID
if ($LASTEXITCODE -ne 0) { Write-Host "Deploy failed."; exit 1 }

Write-Host "[3/3] Clean old deployments..."
$deployLines = (clasp deployments 2>&1) | Where-Object { $_ -match "^- AK" }
$others = $deployLines | Where-Object {
    $_ -notmatch [regex]::Escape($DEPLOY_ID) -and $_ -notmatch "@HEAD"
}
if ($others.Count -gt $MAX_KEEP) {
    $toDelete = $others | Select-Object -First ($others.Count - $MAX_KEEP)
    foreach ($line in $toDelete) {
        $id = ($line -split " ")[1].Trim()
        Write-Host "  Removing: $id"
        clasp undeploy $id 2>&1 | Out-Null
    }
} else {
    Write-Host "  OK ($($others.Count) deployments, no cleanup needed)"
}

Write-Host "Done!"
