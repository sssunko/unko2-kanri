// build_journal.js — デプロイ時に コード.js へ「書いたマスの記録」を自動で入れる（読める本体のコード.jsは書き換えない）
// gas_obfuscate.js の backupAndObfuscate から、難読化の前に呼ぶ。決まりはコード.js 4-2k-1 に書いてある。
//
// 入れるもの（acornでコードを読み、文字の位置に差し込むだけ。処理の順番・中身は変えない）
//  ・シート・マスに書く呼び出し（下の METHODS）の受け手を jrMark_(受け手,'名前') で包む
//      例：sh.getRange(2,1).setValue(x) → jrMark_(sh.getRange(2,1),'setValue').setValue(x)
//      jrMark_ は受け手をそのまま返す。受け手がマス・シートでない時（配列の sort など）は何もしない
//  ・copyTo は受け手を jrSrc_、写す先（1つ目の引数）を jrDest_ で包む（写す先は左上の1マスでも、元の大きさ分を記録する）
//  ・公開の関数（名前が _ で終わらない）とトリガーの関数は、中身を try { … } finally { 一番外側なら jrCommit_() } で包む
//      ＝ 1回の処理（メニュー・アプリ・トリガー）が終わった時に、書いたマスだけを写しと「戻る」の控えに写す
//  ・名前が jr で始まる関数（記録・写し・戻る進むの本体）の中には入れない
'use strict';
const acorn = require('acorn');

const METHODS = new Set([
  'setValue','setValues','setFormula','setFormulas','setFormulaR1C1','setFormulasR1C1','setRichTextValue','setRichTextValues',
  'setFontColor','setFontColors','setFontColorObject','setFontColorObjects','setBackground','setBackgrounds','setBackgroundRGB',
  'setBackgroundObject','setBackgroundObjects','setNumberFormat','setNumberFormats','setDataValidation','setDataValidations',
  'setNote','setNotes','setFontWeight','setFontWeights','setFontSize','setFontSizes','setFontFamily','setFontFamilies',
  'setFontStyle','setFontStyles','setFontLine','setFontLines','setHorizontalAlignment','setHorizontalAlignments',
  'setVerticalAlignment','setVerticalAlignments','setWrap','setWraps','setWrapStrategy','setWrapStrategies','setBorder',
  'setTextRotation','setTextRotations','setTextStyle','setTextStyles','setShowHyperlink','setTextDirection','setTextDirections',
  'merge','mergeAcross','mergeVertically','breakApart','insertCheckboxes','removeCheckboxes','check','uncheck',
  'clear','clearContent','clearContents','clearFormat','clearFormats','clearNote','clearNotes','clearDataValidations',
  'clearConditionalFormatRules','deleteCells','insertCells','sort','appendRow',
  'insertRowAfter','insertRowBefore','insertRows','insertRowsAfter','insertRowsBefore',
  'insertColumnAfter','insertColumnBefore','insertColumns','insertColumnsAfter','insertColumnsBefore',
  'deleteRow','deleteRows','deleteColumn','deleteColumns','setColumnWidth','setColumnWidths','setRowHeight','setRowHeights',
  'setRowHeightsForced','hideColumn','hideColumns','showColumns','unhideColumn','hideRow','hideRows','showRows','unhideRow',
  'setConditionalFormatRules','autoResizeColumn','autoResizeColumns','autoResizeRows','setFrozenRows','setFrozenColumns',
  'trimWhitespace','removeDuplicates','randomize','shiftRowGroupDepth','shiftColumnGroupDepth'
]);
// 名前が _ で終わってもトリガーから直接呼ばれる関数（①のトリガー。②③はスタブから公開の関数を呼ぶ）
const TRIGGER_FNS = new Set(['installedOnEdit_', 'onStructureChange_', 'calcDistanceTrigger_', 'runDailyBackup_']);

function transform(src) {
  const ast = acorn.parse(src, { ecmaVersion: 2020 });
  const ins = [];   // { pos, text, kind: 0=閉じ 1=開き, span }
  function open(pos, text, span)  { ins.push({ pos, text, kind: 1, span }); }
  function close(pos, text, span) { ins.push({ pos, text, kind: 0, span }); }

  function visit(node, skip) {
    if (!node || typeof node.type !== 'string') return;
    if (!skip && node.type === 'CallExpression' && node.callee.type === 'MemberExpression' &&
        !node.callee.computed && node.callee.property.type === 'Identifier') {
      const m = node.callee.property.name, obj = node.callee.object;
      if (obj.type !== 'Super') {
        const span = obj.end - obj.start;
        if (m === 'copyTo' && node.arguments.length) {
          const a0 = node.arguments[0];
          open(obj.start, 'jrSrc_(', span);  close(obj.end, ')', span);
          open(a0.start, 'jrDest_(', a0.end - a0.start);  close(a0.end, ')', a0.end - a0.start);
        } else if (METHODS.has(m)) {
          open(obj.start, 'jrMark_(', span);  close(obj.end, ",'" + m + "')", span);
        }
      }
    }
    for (const key of Object.keys(node)) {
      if (key === 'type' || key === 'start' || key === 'end') continue;
      const v = node[key];
      if (Array.isArray(v)) v.forEach(c => { if (c && typeof c.type === 'string') visit(c, skip); });
      else if (v && typeof v.type === 'string') visit(v, skip);
    }
  }

  let wrapped = 0;
  for (const st of ast.body) {
    if (st.type === 'FunctionDeclaration' && st.id) {
      const name = st.id.name;
      const skip = /^jr/.test(name);
      visit(st.body, skip);
      if (!skip && (!/_$/.test(name) || TRIGGER_FNS.has(name))) {
        const span = st.body.end - st.body.start;
        open(st.body.start + 1, 'JR_.d++||(JR_.top=' + JSON.stringify(name) + ');try{', span);
        close(st.body.end - 1, '}finally{if(--JR_.d===0)jrCommit_();}', span);
        wrapped++;
      }
    } else {
      visit(st, false);
    }
  }
  // 同じ位置：閉じるものを先に（内側から）、開くものは後に（外側から）
  ins.sort((a, b) => a.pos - b.pos || a.kind - b.kind || (a.kind === 0 ? a.span - b.span : b.span - a.span));
  let out = '', cur = 0;
  for (const e of ins) { out += src.slice(cur, e.pos) + e.text; cur = e.pos; }
  out += src.slice(cur);
  acorn.parse(out, { ecmaVersion: 2020 });   // 壊れていたらここで止める（送らない）
  const marks = ins.filter(e => e.kind === 1 && /^jr(Mark|Src|Dest)_\(/.test(e.text)).length;
  return { code: out, marks, wrapped };
}

module.exports = { transform, METHODS, TRIGGER_FNS };

if (require.main === module) {
  const fs = require('fs');
  const r = transform(fs.readFileSync(require('path').join(__dirname, 'コード.js'), 'utf8'));
  console.log('記録を入れた書き込み ' + r.marks + ' 件・包んだ関数 ' + r.wrapped + ' 件');
}
