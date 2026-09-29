// gas_obfuscate.js — デプロイ時にコード.jsを難読化する（読める本体はローカルに必ず残すフェイルセーフ付き）
// deploy.js から呼ぶ。実行時に戻す処理・鍵は使わない（標準の難読化のみ）。
'use strict';
const fs = require('fs');
const path = require('path');
const JavaScriptObfuscator = require('javascript-obfuscator');

const ROOT = __dirname;
const SRC = path.join(ROOT, 'コード.js');
const BAK = path.join(ROOT, 'コード.js.orig.bak');

// GAS(V8)を壊さない設定。公開関数・プロパティ名は変えない。文字列→関数生成系(自己防衛/デバッグ妨害)はGASが禁止するので必ずオフ。
const OPTIONS = {
  compact: true,
  controlFlowFlattening: false,   // 重くなる・タイムアウト回避のためオフ
  deadCodeInjection: false,
  debugProtection: false,          // Function生成を使うためオフ（GAS禁止）
  disableConsoleOutput: false,
  identifierNamesGenerator: 'mangled', // 関数内の一時変数だけ短縮
  numbersToExpressions: false,
  renameGlobals: false,            // 【必須】UnkouLib.xxx・google.script.run・doGet・トリガー名を保護
  renameProperties: false,         // 【必須】呼び出しキー(プロパティ名)を保護
  selfDefending: false,            // Function生成を使うためオフ（GAS禁止）
  simplify: true,
  splitStrings: false,
  stringArray: true,
  stringArrayEncoding: [],         // 実行時デコードを増やさない（高速優先）
  stringArrayThreshold: 0.75,
  transformObjectKeys: false,      // オブジェクトのキー保護
  unicodeEscapeSequence: false,
  target: 'browser'
};

function recoverIfLeftover() {
  // 前回が途中で落ちて控えが残っていたら、先に元へ戻す（本体が難読化のまま残らないための保険）
  if (fs.existsSync(BAK)) {
    fs.copyFileSync(BAK, SRC);
    fs.unlinkSync(BAK);
    console.log('⚠ 前回の控えを検出→本体を元へ復元しました');
  }
}

function backupAndObfuscate() {
  fs.copyFileSync(SRC, BAK); // バイト単位でそのまま控える（改行コードも保持）
  const clean = fs.readFileSync(SRC, 'utf8');
  const result = JavaScriptObfuscator.obfuscate(clean, OPTIONS).getObfuscatedCode();
  fs.writeFileSync(SRC, result, 'utf8');
  console.log(`✓ コード.js を難読化（元${clean.length}→難読化後${result.length}文字）`);
}

function restore() {
  if (fs.existsSync(BAK)) {
    fs.copyFileSync(BAK, SRC);
    fs.unlinkSync(BAK);
    console.log('✓ コード.js を元のきれいな状態へ復元しました');
  }
}

module.exports = { recoverIfLeftover, backupAndObfuscate, restore, OPTIONS, SRC, BAK };
