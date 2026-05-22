Write-Host "[1/1] Push to 修正用SS..."
clasp push --force
if ($LASTEXITCODE -ne 0) { Write-Host "Push failed."; exit 1 }

Write-Host "Done! 元SSへの反映はSSのメニュー「元データに反映」から実行してください。"
