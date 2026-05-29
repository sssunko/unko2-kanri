$env:PATH = "$env:APPDATA\npm;" + $env:PATH
Set-Location $PSScriptRoot
$PROD_ID   = "AKfycbw7rzkd_SuE1I6BNzEjED4Mxl6cnM4wbswIiRiNoPf5zcSS2JcP6YLkfRV21fLc0opU"
$SCRIPT_ID = "1n79omnAcdsEojMRyjnj9-Ic9pIl1-7Nt_HB7Avy0NVFizOSeqt0guqyZ"

# ── バージョン数チェック＆古いもの自動削除 ──────────────────────
Write-Host "Checking versions..."
$verJson = & clasp versions 2>&1 | Out-String
$verNums = [regex]::Matches($verJson, '^\s*(\d+)\s+-', [System.Text.RegularExpressions.RegexOptions]::Multiline) |
           ForEach-Object { [int]$_.Groups[1].Value } | Sort-Object
$verCount = $verNums.Count
Write-Host "Current version count: $verCount"

if ($verCount -ge 190) {
    Write-Host "Too many versions ($verCount). Deleting oldest ones via API..."
    # OAuth tokenをclaspのキャッシュから取得
    $clasprc = "$env:USERPROFILE\.clasprc.json"
    if (Test-Path $clasprc) {
        $token = (Get-Content $clasprc | ConvertFrom-Json).tokens.default.access_token
        $toDelete = $verNums | Select-Object -First ($verCount - 185)
        foreach ($v in $toDelete) {
            $uri = "https://script.googleapis.com/v1/projects/$SCRIPT_ID/versions/$v"
            try {
                Invoke-RestMethod -Uri $uri -Method Delete -Headers @{ Authorization = "Bearer $token" } -ErrorAction SilentlyContinue | Out-Null
                Write-Host "  Deleted version $v"
            } catch {}
        }
    }
}

# ── Push & Deploy ────────────────────────────────────────────────
Write-Host "Pushing..."
& clasp push --force
if ($LASTEXITCODE -ne 0) { Write-Host "ERROR: push failed"; cmd /c pause; exit 1 }

Write-Host "Deploying..."
& clasp deploy -i $PROD_ID -d "update"
if ($LASTEXITCODE -ne 0) { Write-Host "ERROR: deploy failed"; cmd /c pause; exit 1 }

# ── 古い deployment 削除 ─────────────────────────────────────────
Write-Host "Cleaning old deployments..."
$lines = & clasp deployments 2>&1
foreach ($line in $lines) {
    if ($line -match "^- (AKfycb\S+)") {
        $id = $Matches[1]
        if ($id -ne $PROD_ID -and $line -notmatch "@HEAD") {
            Write-Host "Removing old deployment: $id"
            & clasp undeploy $id
        }
    }
}

Write-Host "Deploy Complete!"
cmd /c pause
