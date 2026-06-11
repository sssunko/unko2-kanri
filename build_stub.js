#!/usr/bin/env node
// build_stub.js
// stub_for_clientSS/コード.js を読み、コード.js 内の getClientStubSource_() を自動再生成する。
// マーカー間（START〜END）の内容を丸ごと置換するため、手動編集は不要・禁止。
'use strict';
const fs   = require('fs');
const path = require('path');

const ROOT      = __dirname;
const STUB_FILE = path.join(ROOT, 'stub_for_clientSS', 'コード.js');
const MAIN_FILE = path.join(ROOT, 'コード.js');
const START     = '  // === AUTO_GENERATED_STUB_START（手動編集禁止：build_stub.js が生成） ===';
const END       = '  // === AUTO_GENERATED_STUB_END ===';

const stubSource = fs.readFileSync(STUB_FILE, 'utf8');
const jsonStr    = JSON.stringify(stubSource);

const replacement = START + '\n  return ' + jsonStr + ';\n' + END;

let mainSource = fs.readFileSync(MAIN_FILE, 'utf8');

const si = mainSource.indexOf(START);
const ei = mainSource.indexOf(END);
if (si === -1 || ei === -1) {
  console.error('ERROR: AUTO_GENERATED_STUB マーカーが コード.js に見つかりません。');
  process.exit(1);
}

mainSource = mainSource.slice(0, si) + replacement + mainSource.slice(ei + END.length);
fs.writeFileSync(MAIN_FILE, mainSource, 'utf8');
console.log('✓ getClientStubSource_() を stub_for_clientSS/コード.js から再生成しました。');
