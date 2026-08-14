#!/usr/bin/env node
// deploy.js - ワンコマンドデプロイ
// 使い方: node deploy.js <説明> (例: node deploy.js reloadMenuバグ修正)
// 実行順: check_integrity → build_stub → clasp push → clasp deploy → stub push → appsscript.json更新
'use strict';
const { execSync } = require('child_process');
const fs   = require('fs');
const path = require('path');

const DEPLOY_ID = 'AKfycbw7rzkd_SuE1I6BNzEjED4Mxl6cnM4wbswIiRiNoPf5zcSS2JcP6YLkfRV21fLc0opU';
const ROOT      = __dirname;

const desc = process.argv[2];
if (!desc) {
  console.error('使い方: node deploy.js <説明>\n例: node deploy.js reloadMenuバグ修正');
  process.exit(1);
}

function run(cmd, opts) {
  execSync(cmd, { stdio: 'inherit', cwd: ROOT, ...opts });
}

// バージョン件数チェック
console.log('[0/5] バージョン件数チェック...');
const versions = execSync('clasp versions 2>&1', { cwd: ROOT }).toString();
const vLines   = versions.trim().split('\n').filter(l => /^\d+/.test(l));
if (vLines.length >= 190) {
  console.error(`⚠️ バージョン数が${vLines.length}件（190件以上）。①修正用SS → 拡張機能 → Apps Script → 時計アイコンで古いバージョンを削除してください。`);
  process.exit(1);
}
const lastVer = parseInt(vLines[vLines.length - 1].match(/^(\d+)/)[1]);
const nextVer = lastVer + 1;
console.log(`   現在${vLines.length}件、次バージョン: ${nextVer}`);

// 整合性チェック
console.log('\n[1/5] 整合性チェック...');
run('node check_integrity.js');

// STUB_VERSION_ を次バージョン番号に更新（checkAndRefreshStub の自己更新トリガー）
const mainPath = path.join(ROOT, 'コード.js');
let mainSrc = fs.readFileSync(mainPath, 'utf8');
mainSrc = mainSrc.replace(/^var STUB_VERSION_\s*=\s*'[^']*';/m, `var STUB_VERSION_ = '${nextVer}';`);
fs.writeFileSync(mainPath, mainSrc, 'utf8');
console.log(`✓ STUB_VERSION_ → ${nextVer}`);

// スタブ自動生成
console.log('\n[2/5] スタブ自動生成 (build_stub.js)...');
run('node build_stub.js');

// push + deploy
console.log('\n[3/5] clasp push --force...');
run('clasp push --force');

const deployName = `${nextVer}_${desc}`;
console.log(`\n[4/5] clasp deploy (${deployName})...`);
run(`clasp deploy -i "${DEPLOY_ID}" -d "${deployName}"`);

// stub push + appsscript.json バージョン更新
console.log('\n[5/5] スタブpush & appsscript.json バージョン更新...');
const manifestPath = path.join(ROOT, 'stub_for_clientSS', 'appsscript.json');
const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
manifest.dependencies.libraries.forEach(lib => {
  if (lib.userSymbol === 'UnkouLib') {
    lib.version = String(nextVer);
    lib.developmentMode = false; // 固定バージョン参照：①→②検証→③反映の二重防衛フローを維持
  }
});
fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + '\n', 'utf8');

run('clasp push --force', { cwd: path.join(ROOT, 'stub_for_clientSS') });

console.log(`
============================================
✅ デプロイ完了: ${deployName}
次の操作をしてください:
STEP 2: ①修正用SS → F5 → 📤 テスト客SS（②）に反映
STEP 3: ②客用SS → F5 → 動作確認
STEP 4: ② → 📤 各客に反映
============================================
`);
