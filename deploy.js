#!/usr/bin/env node
// deploy.js - ワンコマンドデプロイ
// 使い方: node deploy.js <説明> (例: node deploy.js reloadMenuバグ修正)
// 実行順: check_integrity → build_stub → clasp push → clasp deploy → stub push → appsscript.json更新 → バージョン一致検証
//
// 【重要：バージョン管理の教訓】
// 過去に「計算値(nextVer)」をstubのappsscript.jsonライブラリバージョンに書き込んでいたため、
// clasp deployが実際に作成したGASバージョン番号と不一致が発生した。
// その結果、①修正用SSは新コードで動くのに②③客用SSは古いコードのまま動く重大バグが起きた。
// 再発防止のため、clasp deployの出力から実バージョン番号を機械的に取得し、
// stubのappsscript.jsonに書き込んだ後、必ず一致を検証してからデプロイ完了とみなす。
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

function capture(cmd, opts) {
  return execSync(cmd, { cwd: ROOT, encoding: 'utf8', ...opts });
}

// バージョン件数チェック
console.log('[0/5] バージョン件数チェック...');
const versions = capture('clasp versions 2>&1');
const vLines   = versions.trim().split('\n').filter(l => /^\d+/.test(l));
if (vLines.length >= 190) {
  console.error(`⚠️ バージョン数が${vLines.length}件（上限200件）。①修正用SS → 拡張機能 → Apps Script → 時計アイコンで古いバージョンを削除してから再実行してください。`);
  process.exit(1);
}
console.log(`   現在${vLines.length}件`);

// 整合性チェック
console.log('\n[1/5] 整合性チェック...');
run('node check_integrity.js');

// STUB_VERSION_ をタイムスタンプで更新（デプロイごとに必ず変化させ、客SS側のスタブ自動更新を確実にトリガーする）
// ※GASバージョン番号と一致させる必要はない。変化すること自体が重要。
const mainPath = path.join(ROOT, 'コード.js');
let mainSrc = fs.readFileSync(mainPath, 'utf8');
const stubVersionTs = String(Date.now());
mainSrc = mainSrc.replace(/^var STUB_VERSION_\s*=\s*'[^']*';/m, `var STUB_VERSION_ = '${stubVersionTs}';`);
fs.writeFileSync(mainPath, mainSrc, 'utf8');
console.log(`✓ STUB_VERSION_ → ${stubVersionTs}`);

// スタブ自動生成（STUB_VERSION_更新後に実行してstubコードに反映させる）
console.log('\n[2/5] スタブ自動生成 (build_stub.js)...');
run('node build_stub.js');

// push（STUB_VERSION_更新済みのコード.jsをライブラリとしてGASに送る）
console.log('\n[3/5] clasp push --force...');
run('clasp push --force');

// deploy（clasp deployの出力から実際に作成されたGASバージョン番号を取得する）
console.log(`\n[4/5] clasp deploy (${desc})...`);
const deployOut = capture(`clasp deploy -i "${DEPLOY_ID}" -d "${desc}"`);
process.stdout.write(deployOut);

// "Deployed AKfyc... @1086" の形式から実バージョン番号を抽出
const verMatch = deployOut.match(/@(\d+)/);
if (!verMatch) {
  console.error('❌ clasp deployの出力からバージョン番号を取得できませんでした。');
  console.error('出力内容: ' + deployOut.trim());
  process.exit(1);
}
const actualVer = parseInt(verMatch[1], 10);
console.log(`✓ 実バージョン番号取得: v${actualVer}`);

// stub appsscript.json を実バージョン番号で更新してpush
// ※ここで nextVer や固定値を使うことは禁止。必ず actualVer を使うこと。
console.log('\n[5/5] スタブpush & appsscript.json バージョン更新...');
const manifestPath = path.join(ROOT, 'stub_for_clientSS', 'appsscript.json');
const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
manifest.dependencies.libraries.forEach(lib => {
  if (lib.userSymbol === 'UnkouLib') {
    lib.version = String(actualVer); // 必ずclasp deployの実バージョン番号を使う
    lib.developmentMode = false;     // 固定バージョン参照：①→②検証→③反映の二重防衛フローを維持
  }
});
fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2) + '\n', 'utf8');
run('clasp push --force', { cwd: path.join(ROOT, 'stub_for_clientSS') });

// 【バージョン一致検証】
// stubのappsscript.jsonを再読込し、実バージョン番号と一致することを必ず確認する。
// 不一致の場合はデプロイ完了とみなさずエラー停止する。
console.log('\n【バージョン一致検証】');
const verifiedManifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
let verifiedVer = null;
verifiedManifest.dependencies.libraries.forEach(lib => {
  if (lib.userSymbol === 'UnkouLib') verifiedVer = parseInt(lib.version, 10);
});
if (verifiedVer !== actualVer) {
  console.error(`❌ バージョン不一致！`);
  console.error(`   ライブラリ実バージョン: v${actualVer}`);
  console.error(`   stub参照バージョン:     v${verifiedVer}`);
  console.error('デプロイ完了とはみなしません。stub_for_clientSS/appsscript.jsonを確認してください。');
  process.exit(1);
}
console.log(`✅ バージョン一致確認: ライブラリv${actualVer} = stub参照v${verifiedVer}`);

console.log(`
============================================
✅ デプロイ完了: v${actualVer}_${desc}
次の操作をしてください:
STEP 2: ①修正用SS → F5 → 📤 テスト客SS（②）に反映
STEP 3: ②客用SS → F5 → 動作確認
STEP 4: ② → 📤 各客に反映
============================================
`);
