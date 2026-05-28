$env:PATH = "$env:APPDATA\npm;" + $env:PATH
Set-Location "$PSScriptRoot\stub_for_clientSS"
Write-Host "テンプレートSS にスタブコードをpushします..."
& clasp push --force
if ($LASTEXITCODE -ne 0) { Write-Host "ERROR: push failed"; cmd /c pause; exit 1 }
Write-Host "完了！テンプレートSS にスタブコードがpushされました。"
Write-Host "以降、このテンプレートをコピーして作られる客SS も自動的にスタブコードになります。"
cmd /c pause
