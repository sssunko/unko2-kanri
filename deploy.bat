@echo off
chcp 65001 > nul
echo ====================================
echo GASへコードをデプロイ(アップロード)します...
echo ====================================

call clasp push

echo.
echo Webアプリのデプロイバージョンを更新しています...
call clasp deploy -i AKfycbw5MLHFep_jOQEdAg4_wX8LMGPX7wVL41XbmygqVV794LkZu6Xv-XcRLNAHYqg9bd0fyw

echo.
echo デプロイが完了しました！
pause
