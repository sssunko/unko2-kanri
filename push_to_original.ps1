$DEV_ID    = "1weCq7YSieGA6kBnqvcycwdTIwT8t8euhHa6QsD2DM3twYBElTpCTOa6W"
$ORIG_ID   = "1n79omnAcdsEojMRyjnj9-Ic9pIl1-7Nt_HB7Avy0NVFizOSeqt0guqyZ"
$DEPLOY_ID = "AKfycbw5MLHFep_jOQEdAg4_wX8LMGPX7wVL41XbmygqVV794LkZu6Xv-XcRLNAHYqg9bd0fyw"
$MAX_KEEP  = 5
$CLASP_JSON = Join-Path $PSScriptRoot ".clasp.json"

function Restore-DevId {
    (Get-Content $CLASP_JSON -Raw -Encoding UTF8) -replace $ORIG_ID, $DEV_ID |
        Set-Content $CLASP_JSON -Encoding UTF8
}

Write-Host "[1/4] Switch to original SS..."
(Get-Content $CLASP_JSON -Raw -Encoding UTF8) -replace $DEV_ID, $ORIG_ID |
    Set-Content $CLASP_JSON -Encoding UTF8

Write-Host "[2/4] Push code..."
clasp push --force
if ($LASTEXITCODE -ne 0) { Write-Host "Push failed."; Restore-DevId; exit 1 }

Write-Host "[3/4] Update deploy..."
clasp deploy --deploymentId $DEPLOY_ID
if ($LASTEXITCODE -ne 0) { Write-Host "Deploy failed."; Restore-DevId; exit 1 }

Write-Host "[4/4] Clean old deployments..."
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

Restore-DevId
Write-Host "Done! All client apps updated."
