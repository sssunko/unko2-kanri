#!/usr/bin/env node
// check_integrity.js
// デプロイ前の静的整合性チェック。
// 1. HTMLのgoogle.script.run呼び出しがスタブに存在するか
// 2. コード.js内の禁止パターン（API関数内でのgetActiveSpreadsheet等）
// 3. AUTO_GENERATED_STUBマーカーの存在確認
'use strict';
const fs   = require('fs');
const path = require('path');

const ROOT = __dirname;
let hasError = false;

function err(msg)  { console.error('❌ ' + msg); hasError = true; }
function warn(msg) { console.warn ('⚠️  ' + msg); }
function ok(msg)   { console.log ('✓  ' + msg); }

// google.script.run.FUNCNAME( のパターンを抽出
// ルール: google.script.run の後、最初に来る .withXxx 以外の .IDENTIFIER( を取る
function extractGsrCalls(src) {
  const called = new Set();
  const lines = src.split('\n');
  const HANDLERS = ['withSuccessHandler', 'withFailureHandler', 'withUserObject'];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!line.includes('google.script.run')) continue;

    // Case 1: same-line direct call  google.script.run.FUNC(
    const direct = line.match(/google\.script\.run\.(\w+)\s*\(/);
    if (direct && !HANDLERS.includes(direct[1])) { called.add(direct[1]); continue; }

    // Case 2: multiline chain — scan ahead for the first non-handler .FUNC(
    for (let j = i + 1; j < Math.min(i + 20, lines.length); j++) {
      const nl = lines[j].trim();
      if (!nl.startsWith('.')) break;
      const m = nl.match(/^\.(\w+)\s*\(/);
      if (!m) continue;
      if (!HANDLERS.includes(m[1])) { called.add(m[1]); break; }
      // Handler line: check if the actual func is chained on the SAME line after the handler
      const inlineAfter = nl.match(/\)\s*\.(\w+)\s*\((?!.*withSuccessHandler|.*withFailureHandler)/);
      if (inlineAfter && !HANDLERS.includes(inlineAfter[1])) { called.add(inlineAfter[1]); break; }
    }
  }
  return called;
}

// 1. HTML→スタブ関数チェック
const HTML_FILES = ['index.html','dispatchDashboard.html','csvImport.html',
                    'etcImport.html','documentPreview.html','plDialog.html'];
const called = new Set();
for (const f of HTML_FILES) {
  const p = path.join(ROOT, f);
  if (!fs.existsSync(p)) continue;
  for (const fn of extractGsrCalls(fs.readFileSync(p, 'utf8'))) called.add(fn);
}

const stubSrc = fs.readFileSync(path.join(ROOT, 'stub_for_clientSS', 'コード.js'), 'utf8');
const stubFns = new Set([...stubSrc.matchAll(/^function\s+(\w+)\s*\(/gm)].map(m => m[1]));

const missing = [...called].filter(f => !stubFns.has(f));
if (missing.length > 0) {
  err('HTMLから呼ばれているがスタブに未定義の関数:\n  ' + missing.join('\n  '));
} else {
  ok('HTML→スタブ関数チェック OK (' + called.size + '件すべて存在)');
}

// 2. 禁止パターンチェック（コード.js）
const mainSrc = fs.readFileSync(path.join(ROOT, 'コード.js'), 'utf8');

const execMatch = mainSrc.match(/function executeStatusSync\b[\s\S]*?^}/m);
if (execMatch && execMatch[0].includes('SpreadsheetApp.getActiveSpreadsheet()')) {
  err('executeStatusSync 内で getActiveSpreadsheet() 使用。ssId引数を使え。');
} else {
  ok('executeStatusSync: getActiveSpreadsheet() 使用なし');
}

const dispatchMatch = mainSrc.match(/function dispatchInstalledEdit\b[\s\S]*?^}/m);
if (dispatchMatch && dispatchMatch[0].includes('ScriptApp.newTrigger')) {
  err('dispatchInstalledEdit 内で ScriptApp.newTrigger 使用。スタブ側の installTriggers に移せ。');
} else {
  ok('dispatchInstalledEdit: ScriptApp.newTrigger 使用なし');
}

// 3. AUTO_GENERATED_STUBマーカー確認
if (!mainSrc.includes('AUTO_GENERATED_STUB_START')) {
  err('getClientStubSource_() に AUTO_GENERATED_STUB_START マーカーがありません。');
} else {
  ok('AUTO_GENERATED_STUB マーカー確認 OK');
}

console.log(hasError ? '\n整合性チェック失敗。デプロイを中止します。' : '\n✅ 整合性チェック全て OK');
process.exit(hasError ? 1 : 0);
