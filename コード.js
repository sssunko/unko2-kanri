// ================================================================
// 運行管理システム Code.gs
// ================================================================
//
// ■ 番号体系
//
// ── 共通・補助 ──────────────────────────
//   1-1  : ID番号取得補助関数（getNextIdNum_）
//   1-2  : Googleドライブのフォルダ取得or作成補助（getOrCreateFolder_）
//   1-3  : 集計表遅延同期ラッパー（delaySyncSummary_）
//   1-4  : 集計表の孤立ID削除（cleanAllOrphanSummary_）
//
// ── スプレッドシート：起動・メニュー ────────────
//   2-1  : メニュー設定（onOpen）
//            - ホ0ーム画面を表示
//            - 集計表再生成
//            - シート再生成（仮日数/給料/%列追加）
//   2-2  : Webアプリ起動（doGet）
//   2-3  : サイドバー表示（showSidebar）
//
// ── スプレッドシート：シート自動処理（onEdit）────
//   3-1  : onEdit本体（シート判定・振り分け）
//   3-2  : 運行シート編集時の処理（onEditUnkou_）
//            - A列ID自動生成（V-XXXX）
//            - I列日付が00:00:00なら現在時刻付加
//            - F列車番入力時にマスタから区分〜携帯番号を自動補完
//            - T列合計高速数式を自動セット
//            - 集計表を同期・孤立ID削除
//   3-3  : 自車専属マスタ編集時の処理（onEditMasterVehicle_）
//            - A列ID自動生成（S-XXXX）
//            - B列運行状態に応じて行色変更
//            - 自車専属運行シートを自動更新
//   3-4  : マスタ（取引先）編集時の処理（onEditMasterCustomer_）
//            - A列ID自動生成（M-XXXX）
//
// ── スプレッドシート：集計表操作 ────────────────
//   4-1  : 集計表再生成（generateSummary）
//            - 運行シートのデータをID単位で集計
//            - 手入力値（距離・ガソリン代・支払・備考・AB〜AD）を保持
//            - AB〜AD列（仮日数・給料・%）をマスタから車番+乗務員名で引き当て
//            - 時刻色付け・利益マイナス赤を再適用
//            - 数式（T列・X列・Z列）を再セット
//            - 支払い再計算（4-4）を自動実行
//   4-2  : 集計表をID単位で同期（syncSummaryForId_）
//            - AB〜AD列を保持＆マスタから引き当て
//            - 時刻色付け・利益マイナス赤を再適用
//
// ── スプレッドシート：シート構造整備 ────────────
//   4-3  : シート再生成（expandAndRefreshSheets）
//            - 自車専属マスタ・自車専属運行に仮日数/給料/%列を追加（なければ）
//            - 自車専属運行シートをマスタから再生成（15列対応）
//   4-4  : 支払い再計算（calculatePaymentAmount）
//            - パターンA: %あり → (売上-合計高速代)×%/100
//            - パターンB: 給料÷仮日数（片方欠けは赤警告）
//            - パターンC: 条件なし → Y列空なら赤警告（手入力値は保持）
//   4-5  : 自車専属運行シート更新内部処理（refreshActiveVehiclesAuto_）
//   4-6  : 自車専属マスタに「運行」列追加（addStatusColumnToMaster）
//
// ── 端末アプリ：起動・紐づけ ────────────────────
//   5-1  : 起動時の初期データ一括取得（getInitialData）
//   5-2  : 紐づけ実行（linkAddress）
//   5-3  : 紐づけ解除（unlinkAddress）
//
// ── 端末アプリ：運行進捗管理 ────────────────────
//   6-1  : 端末の運行進捗を保存（saveRunState）
//   6-2  : 端末の運行進捗を読み込み（loadRunState）
//   6-3  : 端末の運行進捗をクリア（clearRunState）
//
// ── 端末アプリ：運行操作 ────────────────────────
//   7-1  : 今日の行程取得（getTodayRoutes）
//   7-2  : 運行シートへの行作成（createParentRows）
//            - 日付をDate型で書き込み（時刻付き）
//            - LockServiceで同時書き込みによるID重複を防止
//   7-3  : 積完時刻記録（setPickComplete）
//   7-4  : 休憩開始・終了時刻記録（setRest）
//   7-5  : 降完時刻記録（setDropComplete）
//
// ── 端末アプリ：運行一覧・編集 ──────────────────
//   8-1  : 行程データ更新（updateRouteData）
//   8-2  : 運行シート行削除（deleteRunRows）
//   8-3  : 時刻セルクリア（clearTimeCell）
//   8-4  : 運行一覧データ取得（getListData）
//   8-5  : 編集用データ取得（getEditData）
//   8-6  : 編集データ保存（saveEditData）
//   8-7  : 運行データ削除（deleteRunById）
//
// ── 端末アプリ：連絡・ファイル ──────────────────
//   9-1  : 連絡事項保存（saveNotice）
//   9-2  : ファイルアップロード・管理側（uploadFile）
//   9-3  : 端末からの連絡保存（saveTerminalNotice）
//   9-4  : 端末からのファイルアップロード（uploadTerminalFile）
//
// ── 端末アプリ：連絡事項・既読管理 ─────────────
//   10-1 : ホーム用連絡事項取得（getMyNotices）
//   10-2 : ID指定行程取得（getRoutesById）
//   10-3 : 既読管理・既読にする（markAsRead）
//   10-4 : 既読管理・既読一覧取得（getReadNotices）
//
// ■ シート構成
//   設定              : A列=トン数, B列=基準燃費
//   運行(25列)        : ID,区分,会社名,トン数,車種,車番,乗務員名,携帯番号,
//                       日付,荷主名,積地,降地,誘導時刻,積完時刻,休憩開始,休憩終了,降完時刻,
//                       売上,請求高速,実費高速,合計高速,連絡事項,データ,連絡端末,データ端末
//   集計表(31列)      : ID,区分,会社名,トン数,車種,車番,乗務員名,携帯番号,
//                       日付,荷主,積地,降地,誘導時刻,積完時刻,休憩開始,休憩終了,降完時刻,
//                       売上,請求高速代,実費高速代,合計高速代,
//                       距離,燃費,ガソリン代,燃料代,支払い,利益,備考,
//                       仮日数(AC),給料(AD),%(AE)
//   自車専属マスタ(15列): 車両ID,運行状態,区分,会社名,トン数,車種,車番,
//                         乗務員名,携帯番号,アドレス,燃費,備考,仮日数,給料,%
//   自車専属運行(15列)  : 自車専属マスタから運行状態=「運行」のみ抽出
//   マスタ(14列)        : マスタID,会社名,電話,FAX,郵便番号,住所,代表者,
//                         配車担当,銀行名,支店名,種別,番号,名義,備考
// ================================================================


// ================================================================
//  1-1: ID番号取得補助関数
//  指定シートのA列から既存IDの最大番号を取得し+1した値を返す
// ================================================================
function getNextIdNum_(sheet, prefix) {
  var lastRow = sheet.getLastRow();
  var nextNum = 1;
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var k = 0; k < ids.length; k++) {
      var match = String(ids[k][0]).match(/(\d+)$/);
      if (match) {
        var n = parseInt(match[1], 10);
        if (n >= nextNum) nextNum = n + 1;
      }
    }
  }
  return nextNum;
}


// ================================================================
//  1-2: Googleドライブのフォルダ取得or作成（補助）
//  指定名のフォルダが存在すれば返し、なければ作成して返す
// ================================================================
function getOrCreateFolder_(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}


// ================================================================
//  1-3: 集計表遅延同期ラッパー
//  syncSummaryForId_をtry-catchで安全に呼び出す
// ================================================================
function delaySyncSummary_(id) { try { syncSummaryForId_(id); } catch(e) {} }


// ================================================================
//  1-4: 集計表の孤立ID削除
//  運行シートに存在しないIDが集計表にある場合、その行を削除する
// ================================================================
function cleanAllOrphanSummary_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sumSheet = ss.getSheetByName('集計表');
  if (!sumSheet || sumSheet.getLastRow() < 2) return;
  var unkouSheet = ss.getSheetByName('運行');
  if (!unkouSheet) return;
  var unkouData = unkouSheet.getDataRange().getValues();
  var unkouIds = {};
  for (var i = 1; i < unkouData.length; i++) {
    var id = String(unkouData[i][0]||'').trim();
    if (id) unkouIds[id] = true;
  }
  var sumLast = sumSheet.getLastRow();
  var sumIds = sumSheet.getRange(2, 1, sumLast - 1, 1).getValues();
  for (var k = sumIds.length - 1; k >= 0; k--) {
    var sumId = String(sumIds[k][0]||'').trim();
    if (sumId && !unkouIds[sumId]) sumSheet.deleteRow(k + 2);
  }
}


// ================================================================
//  1-5: 金額列へのコンマ書式適用（applyMoneyFormat_）
//  指定シート・行範囲の金額列に #,##0 フォーマットをセットする
// ================================================================
function applyMoneyFormat_(sheet, startRow, numRows, sheetType) {
  if (numRows <= 0) return;
  var fmt = '#,##0';
  var cols = (sheetType === 'unkou')
    ? [18, 19, 20, 21]
    : [18, 19, 20, 21, 24, 25, 26, 27];
  for (var i = 0; i < cols.length; i++) {
    sheet.getRange(startRow, cols[i], numRows, 1).setNumberFormat(fmt);
  }
}


// ================================================================
//  1-6: 時刻列へのM/d HH:mm書式適用（applyDateTimeFormat_）
//  誘導・積完・休憩開始・休憩終了・降完（M〜Q列=13〜17）に書式をセットする
// ================================================================
function applyDateTimeFormat_(sheet, startRow, numRows) {
  if (numRows <= 0) return;
  var fmt = 'M/d HH:mm';
  var cols = [13, 14, 15, 16, 17];
  for (var i = 0; i < cols.length; i++) {
    sheet.getRange(startRow, cols[i], numRows, 1).setNumberFormat(fmt);
  }
}


// ================================================================
//  2-1: メニュー設定（onOpen）
//  スプレッドシート上部に「メニュー」を表示する
//  項目：ホーム画面を表示 / 集計表再生成 / シート再生成
// ================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('メニュー')
    .addItem('ホーム画面を表示', 'showSidebar')
    .addItem('集計表再生成', 'generateSummary')
    .addItem('シート再生成', 'expandAndRefreshSheets')
    .addToUi();
}


// ================================================================
//  2-2: Webアプリ起動（doGet）
//  URLアクセス時にWebアプリとして表示する
// ================================================================
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('運行管理システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// ================================================================
//  2-3: サイドバー表示（showSidebar）
//  スプレッドシートのサイドバーとして表示する
// ================================================================
function showSidebar() {
  var html = HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('ホーム').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}


// ================================================================
//  3-1: onEdit本体
//  編集されたシート名を判定し、対応する処理関数へ振り分ける
// ================================================================
function onEdit(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    var sheetName = sheet.getName();
    if (sheetName === '自車専属マスタ') { onEditMasterVehicle_(sheet, range); return; }
    if (sheetName === 'マスタ') { onEditMasterCustomer_(sheet, range); return; }
    if (sheetName !== '運行') return;
    onEditUnkou_(sheet, range);
  } catch (err) {}
}


// ================================================================
//  3-2: 運行シート編集時の処理（onEditUnkou_）
//  ・A列が空で他列にデータがあればV-XXXXのIDを自動生成
//  ・I列（日付）が00:00:00なら現在時刻を自動付加
//  ・F列（車番）編集時に自車専属マスタから区分〜携帯番号を自動補完
//  ・T列（合計高速）の数式を自動セット
//  ・集計表を同期し、孤立IDを削除
// ================================================================
function onEditUnkou_(sheet, range) {
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  if (startRow <= 1) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  var mData = master ? master.getDataRange().getValues() : [];
  for (var i = 0; i < numRows; i++) {
    var row = startRow + i;
    if (row <= 1) continue;
    var idCell = sheet.getRange(row, 1);
    var currentId = idCell.getValue();
    if (!currentId) {
      var hasData = sheet.getRange(row, 2, 1, 10).getValues()[0].some(function(v) { return v !== ''; });
      if (hasData) {
        var nextNum = getNextIdNum_(sheet, 'V-');
        idCell.setValue('V-' + String(nextNum).padStart(4, '0'));
      }
    }
    var dateCell = sheet.getRange(row, 9);
    var dateVal = dateCell.getValue();
    if (dateVal instanceof Date) {
      var h = dateVal.getHours();
      var m = dateVal.getMinutes();
      var s = dateVal.getSeconds();
      if (h === 0 && m === 0 && s === 0) {
        var now = new Date();
        var merged = new Date(
          dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate(),
          now.getHours(), now.getMinutes(), now.getSeconds()
        );
        dateCell.setValue(merged);
      }
    }
    dateCell.setNumberFormat('yyyy/MM/dd');
    var editedCol = range.getColumn();
    if (editedCol === 6 && range.getNumColumns() === 1) {
      var inputCar = String(sheet.getRange(row, 6).getValue()).trim();
      if (inputCar && mData.length > 1) {
        for (var m2 = 1; m2 < mData.length; m2++) {
          var masterCar = String(mData[m2][6] || '').trim();
          if (masterCar === inputCar || masterCar.indexOf(inputCar) !== -1 || inputCar.indexOf(masterCar) !== -1) {
            sheet.getRange(row, 2).setValue(mData[m2][2]);
            sheet.getRange(row, 3).setValue(mData[m2][3]);
            sheet.getRange(row, 4).setValue(mData[m2][4]);
            sheet.getRange(row, 5).setValue(mData[m2][5]);
            sheet.getRange(row, 6).setValue(masterCar);
            sheet.getRange(row, 7).setValue(mData[m2][7]);
            sheet.getRange(row, 8).setValue(mData[m2][8]);
            break;
          }
        }
      }
    }
    // M/N/O/P/Q列（誘導・積完・休憩・降完時刻）：全角文字・日付なし時刻を正規化して合成
    if ([13, 14, 15, 16, 17].indexOf(editedCol) !== -1) {
      var timeCell = sheet.getRange(row, editedCol);
      var tv = timeCell.getValue();
      var baseDateObj = (sheet.getRange(row, 9).getValue() instanceof Date) ? sheet.getRange(row, 9).getValue() : null;
      var merged = null;
      if (typeof tv === 'string' && tv.trim() !== '') {
        var s = tv.trim().replace(/[：]/g, ':').replace(/[　]/g, ' ');
        var m1 = s.match(/^(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})/);
        if (m1) {
          var yr = baseDateObj ? baseDateObj.getFullYear() : new Date().getFullYear();
          merged = new Date(yr, parseInt(m1[1]) - 1, parseInt(m1[2]), parseInt(m1[3]), parseInt(m1[4]), 0);
        } else {
          var m2 = s.match(/^(\d{1,2}):(\d{2})$/);
          if (m2) {
            var base2 = baseDateObj || new Date();
            merged = new Date(base2.getFullYear(), base2.getMonth(), base2.getDate(),
                              parseInt(m2[1]), parseInt(m2[2]), 0);
          }
        }
      } else if (tv instanceof Date && tv.getFullYear() < 1902) {
        var base3 = baseDateObj || new Date();
        merged = new Date(base3.getFullYear(), base3.getMonth(), base3.getDate(),
                          tv.getHours(), tv.getMinutes(), tv.getSeconds());
      }
      if (merged && !isNaN(merged.getTime())) {
        timeCell.setValue(merged);
        timeCell.setNumberFormat('M/d HH:mm');
      }
    }
    var tollCell = sheet.getRange(row, 21);
    if (!tollCell.getFormula()) {
      tollCell.setFormula('=IF(AND(T' + row + '="",S' + row + '=""),"",T' + row + '-S' + row + ')');
    }
    var newId = sheet.getRange(row, 1).getValue();
    if (newId) syncSummaryForId_(newId);
  }
  applyMoneyFormat_(sheet, startRow, numRows, 'unkou');
  applyDateTimeFormat_(sheet, startRow, numRows);
  cleanAllOrphanSummary_();
}


// ================================================================
//  3-3: 自車専属マスタ編集時の処理（onEditMasterVehicle_）
//  ・A列が空で他列にデータがあればS-XXXXのIDを自動生成
//  ・B列（運行状態）の値に応じて行の背景色を変更
//    運行→薄赤, 待機→薄黄, 故障→薄緑, その他→なし
//  ・自車専属運行シートを自動更新
// ================================================================
function onEditMasterVehicle_(sheet, range) {
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  for (var i = 0; i < numRows; i++) {
    var row = startRow + i;
    if (row <= 1) continue;
    var idCell = sheet.getRange(row, 1);
    if (!idCell.getValue()) {
      var hasData = sheet.getRange(row, 2, 1, 5).getValues()[0].some(function(v) { return v !== ''; });
      if (hasData) {
        var nextNum = getNextIdNum_(sheet, 'S-');
        idCell.setValue('S-' + String(nextNum).padStart(4, '0'));
      }
    }
    var status = String(sheet.getRange(row, 2).getValue()).trim();
    var lastCol = sheet.getLastColumn() || 12;
    var rowRange = sheet.getRange(row, 1, 1, lastCol);
    if (status === '運行') { rowRange.setBackground('#ffcdd2'); }
    else if (status === '待機') { rowRange.setBackground('#fff9c4'); }
    else if (status === '故障') { rowRange.setBackground('#c8e6c9'); }
    else { rowRange.setBackground(null); }
  }
  refreshActiveVehiclesAuto_();
}


// ================================================================
//  3-4: マスタ（取引先）編集時の処理（onEditMasterCustomer_）
//  ・A列が空で他列にデータがあればM-XXXXのIDを自動生成
// ================================================================
function onEditMasterCustomer_(sheet, range) {
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  for (var i = 0; i < numRows; i++) {
    var row = startRow + i;
    if (row <= 1) continue;
    var idCell = sheet.getRange(row, 1);
    if (!idCell.getValue()) {
      var hasData = sheet.getRange(row, 2, 1, 5).getValues()[0].some(function(v) { return v !== ''; });
      if (hasData) {
        var nextNum = getNextIdNum_(sheet, 'M-');
        idCell.setValue('M-' + String(nextNum).padStart(4, '0'));
      }
    }
  }
}


// ================================================================
//  4-1: 集計表再生成（generateSummary）
//  運行シートのデータをID単位で集計し集計表を再生成する
//  ・手入力値（距離・ガソリン代・支払・備考・AB〜AD）を保持
//  ・AB〜AD列は既存値優先、なければマスタから車番+乗務員名で引き当て
//  ・時刻色付け・利益マイナス赤を再適用
//  ・数式（T列・X列・Z列）を再セット
//  ・生成後に支払い再計算（4-4）を自動実行
// ================================================================
function generateSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var unkouSheet = ss.getSheetByName('運行');
  if (!unkouSheet) return;

  var settingSheet = ss.getSheetByName('設定');
  var fuelMap = {};
  if (settingSheet && settingSheet.getLastRow() >= 2) {
    var sVals = settingSheet.getRange(2, 1, settingSheet.getLastRow()-1, 2).getValues();
    for (var s = 0; s < sVals.length; s++) {
      if (sVals[s][0]) fuelMap[String(sVals[s][0]).trim()] = sVals[s][1];
    }
  }

  var master = ss.getSheetByName('自車専属マスタ');
  var payCondMap = {};
  if (master && master.getLastRow() >= 2) {
    var mData = master.getRange(2, 1, master.getLastRow()-1, 15).getValues();
    for (var m = 0; m < mData.length; m++) {
      var mcar  = String(mData[m][6]  || '').trim();
      var mname = String(mData[m][7]  || '').trim();
      var pkey  = mcar + '_' + mname;
      payCondMap[pkey] = {
        kari:   mData[m][12] || '',
        kyuryo: mData[m][13] || '',
        pct:    mData[m][14] || ''
      };
    }
  }

  var sumSheet = ss.getSheetByName('集計表');
  var oldData = {};
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var colCount = sumSheet.getLastColumn();
    var oldRows = sumSheet.getRange(2, 1, sumSheet.getLastRow()-1, Math.max(colCount, 31)).getValues();
    for (var o = 0; o < oldRows.length; o++) {
      var oldId = String(oldRows[o][0] || '').trim();
      if (oldId) {
        oldData[oldId] = {
          distance: oldRows[o][21],
          gas:      oldRows[o][23],
          pay:      oldRows[o][25],
          memo:     oldRows[o][27],
          kari:     oldRows[o][28] || '',
          kyuryo:   oldRows[o][29] || '',
          pct:      oldRows[o][30] || ''
        };
      }
    }
  }
  if (!sumSheet) sumSheet = ss.insertSheet('集計表');

  var header = [
    'ID','区分','会社名','トン数','車種','車番','乗務員名','携帯番号',
    '日付','荷主','積地','降地','誘導時刻','積完時刻','休憩開始','休憩終了','降完時刻',
    '売上','請求(高速代)','実費(高速代)','合計(高速代)',
    '距離','燃費','ガソリン代','燃料代','支払い','利益','備考',
    '仮日数','給料','％'
  ];

  var unkouData = unkouSheet.getDataRange().getValues();
  var idMap = {}, idOrder = [];
  for (var i = 1; i < unkouData.length; i++) {
    var r  = unkouData[i];
    var id = String(r[0] || '').trim();
    if (!id) continue;
    if (!idMap[id]) {
      idMap[id] = {
        id:id, kubun:r[1], company:r[2], tons:r[3], type:r[4], car:r[5],
        name:r[6], tel:r[7], date:r[8], client:r[9],
        picks:[], drops:[],
        guideTime:'',
        pickTime:'', restStart:'', restEnd:'', dropTime:'',
        rawPickTime:null, rawRestStart:null, rawRestEnd:null, rawDropTime:null,
        sales:0, tollReq:0, tollReal:0
      };
      idOrder.push(id);
    }
    var g = idMap[id];
    if (r[10]) g.picks.push(r[10]);
    if (r[11]) g.drops.push(r[11]);
    if (r[12] && !g.guideTime) { g.guideTime = r[12]; }
    if (r[13] && !g.pickTime)  { g.pickTime  = r[13]; g.rawPickTime  = new Date(r[13]); }
    if (r[14] && !g.restStart) { g.restStart = r[14]; g.rawRestStart = new Date(r[14]); }
    if (r[15] && !g.restEnd)   { g.restEnd   = r[15]; g.rawRestEnd   = new Date(r[15]); }
    if (r[16] && !g.dropTime)  { g.dropTime  = r[16]; g.rawDropTime  = new Date(r[16]); }
    g.sales   += Number(r[17]) || 0;
    g.tollReq += Number(r[18]) || 0;
    g.tollReal+= Number(r[19]) || 0;
    if (r[9] && !g.client) g.client = r[9];
  }

  var outRows = [header];
  for (var o2 = 0; o2 < idOrder.length; o2++) {
    var g      = idMap[idOrder[o2]];
    var tonsStr= String(g.tons || '').trim();
    var fuel   = fuelMap[tonsStr] || fuelMap[tonsStr.replace(/[tT]/,'')+'t'] || 3;
    var old    = oldData[g.id] || {};
    var pkey   = String(g.car||'').trim() + '_' + String(g.name||'').trim();
    var pc     = payCondMap[pkey] || {kari:'', kyuryo:'', pct:''};
    var kari   = (old.kari   !== undefined && old.kari   !== '') ? old.kari   : pc.kari;
    var kyuryo = (old.kyuryo !== undefined && old.kyuryo !== '') ? old.kyuryo : pc.kyuryo;
    var pct    = (old.pct    !== undefined && old.pct    !== '') ? old.pct    : pc.pct;

    outRows.push([
      g.id, g.kubun, g.company, g.tons, g.type, g.car, g.name, g.tel,
      g.date, g.client, g.picks.join('・'), g.drops.join('・'),
      g.guideTime||'', g.pickTime, g.restStart, g.restEnd, g.dropTime,
      g.sales||'', g.tollReq||'', g.tollReal||'', '',
      old.distance||'', fuel, old.gas||'', '',
      old.pay||'', '', old.memo||'',
      kari, kyuryo, pct
    ]);
  }

  sumSheet.clear();
  if (outRows.length > 0) {
    sumSheet.getRange(1, 1, outRows.length, 31).setValues(outRows);
    sumSheet.setFrozenRows(1);

    var F = 4*60*60*1000;
    var T = 30*60*1000;

    for (var row = 2; row <= outRows.length; row++) {
      sumSheet.getRange(row, 21).setFormula('=IF(AND(T'+row+'="",S'+row+'=""),"",T'+row+'-S'+row+')');
      sumSheet.getRange(row, 25).setFormula('=IF(OR(V'+row+'="",W'+row+'=""),"",V'+row+'/W'+row+'*X'+row+')');
      sumSheet.getRange(row, 27).setFormula('=IF(AND(R'+row+'="",U'+row+'="",Y'+row+'="",Z'+row+'=""),"",R'+row+'-(U'+row+'+Y'+row+'+Z'+row+'))');

      var g2       = idMap[idOrder[row-2]];
      var keepPay  = outRows[row-1][25] || '';
      var keepDist = outRows[row-1][21] || '';
      var keepGas  = outRows[row-1][23] || '';
      var calcToll = (Number(g2.tollReal)||0)-(Number(g2.tollReq)||0);
      var calcFuel = (Number(keepDist)&&Number(fuel)&&Number(keepGas)) ? (Number(keepDist)/Number(fuel)*Number(keepGas)) : 0;
      var calcProfit = (Number(g2.sales)||0)-(calcToll+calcFuel+(Number(keepPay)||0));

      sumSheet.getRange(row, 1, 1, 31).setBackground(calcProfit < 0 ? '#ffebee' : null);
      sumSheet.getRange(row, 14, 1, 4).setBackground(null);
      if (g2.rawPickTime  && g2.rawRestStart && (g2.rawRestStart-g2.rawPickTime)  > F) { sumSheet.getRange(row,14,1,2).setBackground('#ffd600'); }
      if (g2.rawRestStart && g2.rawRestEnd   && (g2.rawRestEnd  -g2.rawRestStart) < T) { sumSheet.getRange(row,15,1,2).setBackground('#4fc3f7'); }
      if (g2.rawRestEnd   && g2.rawDropTime  && (g2.rawDropTime -g2.rawRestEnd)   > F) { sumSheet.getRange(row,16,1,2).setBackground('#ffd600'); }
    }
    applyMoneyFormat_(sumSheet, 2, outRows.length - 1, 'summary');
    applyDateTimeFormat_(sumSheet, 2, outRows.length - 1);
  }

  calculatePaymentAmount();
}


// ================================================================
//  4-2: 集計表をID単位で同期（syncSummaryForId_）
//  運行シートから対象IDのデータを集計し集計表の該当行を更新する
//  ・AB〜AD列（仮日数・給料・%）を保持＆マスタから引き当て
//  ・時刻色付け・利益マイナス赤を再適用
//  ・数式（T列・X列・Z列）を再セット
// ================================================================
function syncSummaryForId_(targetId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var unkouSheet = ss.getSheetByName('運行');
  if (!unkouSheet) return;
  var sumSheet = ss.getSheetByName('集計表');
  if (!sumSheet || sumSheet.getLastRow() < 1) { generateSummary(); return; }

  var settingSheet = ss.getSheetByName('設定');
  var fuelMap = {};
  if (settingSheet && settingSheet.getLastRow() >= 2) {
    var sVals = settingSheet.getRange(2, 1, settingSheet.getLastRow()-1, 2).getValues();
    for (var s = 0; s < sVals.length; s++) {
      if (sVals[s][0]) fuelMap[String(sVals[s][0]).trim()] = sVals[s][1];
    }
  }

  var master = ss.getSheetByName('自車専属マスタ');
  var payCondMap = {};
  if (master && master.getLastRow() >= 2) {
    var mData = master.getRange(2, 1, master.getLastRow()-1, 15).getValues();
    for (var m = 0; m < mData.length; m++) {
      var mcar  = String(mData[m][6]  || '').trim();
      var mname = String(mData[m][7]  || '').trim();
      payCondMap[mcar+'_'+mname] = {
        kari:   mData[m][12] || '',
        kyuryo: mData[m][13] || '',
        pct:    mData[m][14] || ''
      };
    }
  }

  var unkouData = unkouSheet.getDataRange().getValues();
  var g = null;
  var matchingRows = [];
  var rawPickTime=null, rawRestStart=null, rawRestEnd=null, rawDropTime=null;
  for (var i = 1; i < unkouData.length; i++) {
    var r = unkouData[i];
    if (String(r[0]||'').trim() !== String(targetId).trim()) continue;
    matchingRows.push(i + 1);
    if (!g) { g = { id:String(r[0]).trim(), kubun:r[1], company:r[2], tons:r[3], type:r[4], car:r[5], name:r[6], tel:r[7], date:r[8], client:r[9], picks:[], drops:[], guideTime:'', pickTime:'', restStart:'', restEnd:'', dropTime:'', sales:0, tollReq:0, tollReal:0 }; }
    if (r[10]) g.picks.push(r[10]);
    if (r[11]) g.drops.push(r[11]);
    if (r[12] && !g.guideTime) { g.guideTime = r[12]; }
    if (r[13] && !g.pickTime)  { g.pickTime  = r[13]; rawPickTime  = new Date(r[13]); }
    if (r[14] && !g.restStart) { g.restStart = r[14]; rawRestStart = new Date(r[14]); }
    if (r[15] && !g.restEnd)   { g.restEnd   = r[15]; rawRestEnd   = new Date(r[15]); }
    if (r[16] && !g.dropTime)  { g.dropTime  = r[16]; rawDropTime  = new Date(r[16]); }
    g.sales   += Number(r[17]) || 0;
    g.tollReq += Number(r[18]) || 0;
    g.tollReal+= Number(r[19]) || 0;
  }
  // 運行シートの該当行に書式を確実に適用
  if (matchingRows.length > 0) {
    var minR = matchingRows[0], maxR = matchingRows[matchingRows.length - 1];
    unkouSheet.getRange(minR, 9, maxR - minR + 1, 1).setNumberFormat('yyyy/MM/dd');
    applyDateTimeFormat_(unkouSheet, minR, maxR - minR + 1);
  }

  var sumLast = sumSheet.getLastRow();
  var sumRow  = 0;
  var keepDistance='', keepGas='', keepPay='', keepMemo='';
  var keepKari='', keepKyuryo='', keepPct='';
  if (sumLast >= 2) {
    var colCount = sumSheet.getLastColumn();
    var sumIds   = sumSheet.getRange(2, 1, sumLast-1, Math.max(colCount, 31)).getValues();
    for (var k = 0; k < sumIds.length; k++) {
      if (String(sumIds[k][0]).trim() === String(targetId).trim()) {
        sumRow      = k + 2;
        keepDistance= sumIds[k][21];
        keepGas     = sumIds[k][23];
        keepPay     = sumIds[k][25];
        keepMemo    = sumIds[k][27];
        keepKari    = sumIds[k][28] || '';
        keepKyuryo  = sumIds[k][29] || '';
        keepPct     = sumIds[k][30] || '';
        break;
      }
    }
  }
  if (!g) { if (sumRow > 0) sumSheet.deleteRow(sumRow); return; }

  var tonsStr = String(g.tons||'').trim();
  var fuel    = fuelMap[tonsStr] || fuelMap[tonsStr.replace(/[tT]/,'')+'t'] || 3;

  var pkey   = String(g.car||'').trim()+'_'+String(g.name||'').trim();
  var pc     = payCondMap[pkey] || {kari:'', kyuryo:'', pct:''};
  var kari   = keepKari   !== '' ? keepKari   : pc.kari;
  var kyuryo = keepKyuryo !== '' ? keepKyuryo : pc.kyuryo;
  var pct    = keepPct    !== '' ? keepPct    : pc.pct;

  var rowData = [
    g.id, g.kubun, g.company, g.tons, g.type, g.car, g.name, g.tel,
    g.date, g.client, g.picks.join('・'), g.drops.join('・'),
    g.guideTime||'', g.pickTime, g.restStart, g.restEnd, g.dropTime,
    g.sales||'', g.tollReq||'', g.tollReal||'', '',
    keepDistance, fuel, keepGas, '', keepPay, '', keepMemo,
    kari, kyuryo, pct
  ];

  if (sumRow > 0) {
    sumSheet.getRange(sumRow, 1, 1, 31).setValues([rowData]);
  } else {
    sumRow = sumSheet.getLastRow()+1;
    if (sumRow === 1) {
      var hdr = ['ID','区分','会社名','トン数','車種','車番','乗務員名','携帯番号','日付','荷主','積地','降地','誘導時刻','積完時刻','休憩開始','休憩終了','降完時刻','売上','請求(高速代)','実費(高速代)','合計(高速代)','距離','燃費','ガソリン代','燃料代','支払い','利益','備考','仮日数','給料','％'];
      sumSheet.getRange(1, 1, 1, 31).setValues([hdr]);
      sumSheet.setFrozenRows(1);
      sumRow = 2;
    }
    sumSheet.getRange(sumRow, 1, 1, 31).setValues([rowData]);
  }

  sumSheet.getRange(sumRow, 21).setFormula('=IF(AND(T'+sumRow+'="",S'+sumRow+'=""),"",T'+sumRow+'-S'+sumRow+')');
  sumSheet.getRange(sumRow, 25).setFormula('=IF(OR(V'+sumRow+'="",W'+sumRow+'=""),"",V'+sumRow+'/W'+sumRow+'*X'+sumRow+')');
  sumSheet.getRange(sumRow, 27).setFormula('=IF(AND(R'+sumRow+'="",U'+sumRow+'="",Y'+sumRow+'="",Z'+sumRow+'=""),"",R'+sumRow+'-(U'+sumRow+'+Y'+sumRow+'+Z'+sumRow+'))');

  var calcToll  = (Number(g.tollReal)||0)-(Number(g.tollReq)||0);
  var calcFuel  = (Number(keepDistance)&&Number(fuel)&&Number(keepGas)) ? (Number(keepDistance)/Number(fuel)*Number(keepGas)) : 0;
  var calcProfit= (Number(g.sales)||0)-(calcToll+calcFuel+(Number(keepPay)||0));
  sumSheet.getRange(sumRow, 1, 1, 30).setBackground(calcProfit < 0 ? '#ffebee' : null);

  var F = 4*60*60*1000;
  var T = 30*60*1000;
  sumSheet.getRange(sumRow, 14, 1, 4).setBackground(null);
  if (rawPickTime  && rawRestStart && (rawRestStart-rawPickTime)  > F) { sumSheet.getRange(sumRow,14,1,2).setBackground('#ffd600'); }
  if (rawRestStart && rawRestEnd   && (rawRestEnd  -rawRestStart) < T) { sumSheet.getRange(sumRow,15,1,2).setBackground('#4fc3f7'); }
  if (rawRestEnd   && rawDropTime  && (rawDropTime -rawRestEnd)   > F) { sumSheet.getRange(sumRow,16,1,2).setBackground('#ffd600'); }
  applyMoneyFormat_(sumSheet, sumRow, 1, 'summary');
  applyDateTimeFormat_(sumSheet, sumRow, 1);
}


// ================================================================
//  4-3: シート再生成（expandAndRefreshSheets）
//  メニューの「シート再生成」から呼び出す
//  ・自車専属マスタに仮日数/給料/%列がなければ追加
//  ・自車専属運行シートをマスタから再生成（15列対応）
// ================================================================
function expandAndRefreshSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsToUpdate = ['自車専属マスタ', '自車専属運行'];
  var colsToAdd = ['仮日数', '給料', '％'];

  for (var i = 0; i < sheetsToUpdate.length; i++) {
    var sheet = ss.getSheetByName(sheetsToUpdate[i]);
    if (!sheet) continue;
    var lastCol = sheet.getLastColumn();
    if (lastCol === 0) {
      sheet.getRange(1, 1, 1, colsToAdd.length).setValues([colsToAdd]);
      continue;
    }
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var nextCol = lastCol + 1;
    for (var j = 0; j < colsToAdd.length; j++) {
      if (headers.indexOf(colsToAdd[j]) === -1) {
        sheet.getRange(1, nextCol).setValue(colsToAdd[j]);
        nextCol++;
      }
    }
  }

  refreshActiveVehiclesAuto_();

  var unkouForFmt = ss.getSheetByName('運行');
  if (unkouForFmt && unkouForFmt.getLastRow() >= 2) {
    applyMoneyFormat_(unkouForFmt, 2, unkouForFmt.getLastRow() - 1, 'unkou');
    applyDateTimeFormat_(unkouForFmt, 2, unkouForFmt.getLastRow() - 1);
    unkouForFmt.getRange(2, 9, unkouForFmt.getLastRow() - 1, 1).setNumberFormat('yyyy/MM/dd');
  }
  var sumForFmt = ss.getSheetByName('集計表');
  if (sumForFmt && sumForFmt.getLastRow() >= 2) {
    applyMoneyFormat_(sumForFmt, 2, sumForFmt.getLastRow() - 1, 'summary');
    applyDateTimeFormat_(sumForFmt, 2, sumForFmt.getLastRow() - 1);
  }

  SpreadsheetApp.getUi().alert('シート再生成が完了しました。');
}


// ================================================================
//  4-4: 支払い再計算（calculatePaymentAmount）
//  集計表のAB列(仮日数)・AC列(給料)・AD列(%)からY列(支払い)を計算する
//  ・パターンA: %あり → (売上-合計高速代)×%/100
//  ・パターンB: %なし・給料と仮日数あり → 給料÷仮日数
//              片方欠け → 欠けているセルを赤警告
//  ・パターンC: 条件なし → Y列が空なら赤警告（手入力値は保持）
// ================================================================
function calculatePaymentAmount() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('集計表');
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var data = sheet.getRange(2, 1, lastRow-1, 31).getValues();

  for (var i = 0; i < data.length; i++) {
    var rowNum    = i + 2;
    var sales     = Number(data[i][17]) || 0;
    var totalToll = Number(data[i][20]) || 0;
    var kari      = Number(data[i][28]) || 0;
    var kyuryo    = Number(data[i][29]) || 0;
    var pct       = Number(data[i][30]) || 0;
    var yCell     = sheet.getRange(rowNum, 26);

    yCell.setBackground(null);
    sheet.getRange(rowNum, 29, 1, 3).setBackground(null);

    if (pct > 0) {
      yCell.setValue((sales - totalToll) * (pct / 100));
    } else if (kyuryo > 0 || kari > 0) {
      if (kyuryo > 0 && kari > 0) {
        yCell.setValue(kyuryo / kari);
      } else {
        if (!kyuryo) sheet.getRange(rowNum, 30).setBackground('#f4cccc');
        if (!kari)   sheet.getRange(rowNum, 29).setBackground('#f4cccc');
      }
    } else {
      if (yCell.getValue() === '') yCell.setBackground('#f4cccc');
    }
  }
}


// ================================================================
//  4-5: 自車専属運行シート更新内部処理（refreshActiveVehiclesAuto_）
//  自車専属マスタの運行状態=「運行」の行のみを抽出し
//  自車専属運行シートに15列（A〜O列、仮日数/給料/%含む）で書き出す
// ================================================================
function refreshActiveVehicles() { refreshActiveVehiclesAuto_(); }
function refreshActiveVehiclesAuto_() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master) return;
  var activeSheet = ss.getSheetByName('自車専属運行');
  if (!activeSheet) activeSheet = ss.insertSheet('自車専属運行');

  var header  = ['車両ID','運行状態','区分','会社名','トン数','車種','車番','乗務員名','携帯番号','アドレス','燃費','備考','仮日数','給料','％'];
  var lastRow = master.getLastRow();
  var mData   = lastRow >= 2 ? master.getRange(2, 1, lastRow-1, 15).getValues() : [];
  var outRows = [header];
  for (var i = 0; i < mData.length; i++) {
    if (String(mData[i][1]||'').trim() === '運行') {
      outRows.push([
        mData[i][0],  mData[i][1],  mData[i][2],  mData[i][3],  mData[i][4],
        mData[i][5],  mData[i][6],  mData[i][7],  mData[i][8],  mData[i][9],
        mData[i][10], mData[i][11], mData[i][12], mData[i][13], mData[i][14]
      ]);
    }
  }
  activeSheet.clear();
  if (outRows.length > 0) {
    activeSheet.getRange(1, 1, outRows.length, 15).setValues(outRows);
    activeSheet.setFrozenRows(1);
  }
}


// ================================================================
//  4-6: 自車専属マスタに「運行」列追加（addStatusColumnToMaster）
//  B列が「運行」でなければB列を挿入し全行に「運行」をセットする
// ================================================================
function addStatusColumnToMaster() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('自車専属マスタ');
  if (!sheet) return;
  if (String(sheet.getRange(1, 2).getValue()).trim() === '運行') return;
  sheet.insertColumnBefore(2);
  sheet.getRange(1, 2).setValue('運行');
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) { sheet.getRange(2, 2, lastRow-1, 1).setValue('運行'); }
}


// ================================================================
//  5-1: 起動時の初期データ一括取得（getInitialData）
//  端末保存のアドレスを元にマスタから該当行を一括検索して返す
// ================================================================
function getInitialData() {
  var savedEmail = PropertiesService.getUserProperties().getProperty('linkedEmail');
  var result = { email: savedEmail || "", profile: null };
  if (!savedEmail) return result;
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master) return result;
  var data       = master.getDataRange().getValues();
  var emailLower = savedEmail.toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    if (data[i][9] && String(data[i][9]).toLowerCase().trim() === emailLower) {
      result.profile = {
        company: data[i][3], tons: data[i][4], type: data[i][5],
        carNo:   data[i][6], name: data[i][7], tel:  data[i][8]
      };
      break;
    }
  }
  return result;
}


// ================================================================
//  5-2: 紐づけ実行（linkAddress）
//  入力アドレスを自車専属マスタのJ列と照合し
//  一致したら端末のPropertiesServiceに保存する（シートには書かない）
// ================================================================
function linkAddress(email) {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master) return "エラー：マスタシートなし";
  var rows = master.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][9]).trim() === String(email).trim()) {
      PropertiesService.getUserProperties().setProperty('linkedEmail', email);
      return {
        status: "紐づけOK", email: email,
        company: rows[i][3], tons: rows[i][4], type: rows[i][5],
        carNo:   rows[i][6], name: rows[i][7], tel:  rows[i][8]
      };
    }
  }
  return "エラー：アドレス未登録";
}


// ================================================================
//  5-3: 紐づけ解除（unlinkAddress）
//  端末のPropertiesServiceからアドレス情報を消去する
// ================================================================
function unlinkAddress() {
  PropertiesService.getUserProperties().deleteProperty('linkedEmail');
  return "解除しました";
}


// ================================================================
//  6-1: 端末の運行進捗を保存（saveRunState）
//  picks/drops/rows/pickDone/dropDone/phase/lastPickRow/
//  pickHistory/dropHistoryの9項目をsetPropertiesで一括保存
// ================================================================
function saveRunState(state) {
  var p = PropertiesService.getUserProperties();
  p.setProperties({
    'picks':        JSON.stringify(state.picks        || []),
    'drops':        JSON.stringify(state.drops        || []),
    'rows':         JSON.stringify(state.rows         || []),
    'guideDone':    JSON.stringify(state.guideDone    || []),
    'pickDone':     JSON.stringify(state.pickDone     || []),
    'dropDone':     JSON.stringify(state.dropDone     || []),
    'phase':        state.phase                       || '',
    'lastPickRow':  String(state.lastPickRow          || ''),
    'guideHistory': JSON.stringify(state.guideHistory || []),
    'pickHistory':  JSON.stringify(state.pickHistory  || []),
    'dropHistory':  JSON.stringify(state.dropHistory  || [])
  }, false);
}


// ================================================================
//  6-2: 端末の運行進捗を読み込み（loadRunState）
//  getPropertiesで一括取得して返す
// ================================================================
function loadRunState() {
  var all = PropertiesService.getUserProperties().getProperties();
  return {
    picks:        JSON.parse(all['picks']        || '[]'),
    drops:        JSON.parse(all['drops']        || '[]'),
    rows:         JSON.parse(all['rows']         || '[]'),
    guideDone:    JSON.parse(all['guideDone']    || '[]'),
    pickDone:     JSON.parse(all['pickDone']     || '[]'),
    dropDone:     JSON.parse(all['dropDone']     || '[]'),
    phase:        all['phase']                   || '',
    lastPickRow:  all['lastPickRow']             || '',
    guideHistory: JSON.parse(all['guideHistory'] || '[]'),
    pickHistory:  JSON.parse(all['pickHistory']  || '[]'),
    dropHistory:  JSON.parse(all['dropHistory']  || '[]')
  };
}


// ================================================================
//  6-3: 端末の運行進捗をクリア（clearRunState）
//  linkedEmail（紐づけ）とreadNotices（既読管理）は消さない
//  運行進捗の9項目だけ削除する
// ================================================================
function clearRunState() {
  var p    = PropertiesService.getUserProperties();
  var keys = ['picks','drops','rows','guideDone','pickDone','dropDone','phase','lastPickRow','guideHistory','pickHistory','dropHistory'];
  for (var i = 0; i < keys.length; i++) { p.deleteProperty(keys[i]); }
}


// ================================================================
//  7-1: 今日の行程取得（getTodayRoutes）
//  紐づけアドレスから乗務員名・車番を特定し
//  運行シートから本日分の未完了行程を返す
// ================================================================
function getTodayRoutes() {
  var savedEmail = PropertiesService.getUserProperties().getProperty('linkedEmail');
  if (!savedEmail) return [];
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master) return [];
  var mAll = master.getDataRange().getValues();
  var name = '', car = '';
  for (var j = 1; j < mAll.length; j++) {
    if (String(mAll[j][9]).trim() === savedEmail) {
      name = String(mAll[j][7]).trim();
      car  = String(mAll[j][6]).trim();
      break;
    }
  }
  if (!name || !car) return [];
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return [];
  var all   = sheet.getDataRange().getValues();
  var today = new Date();
  var y = today.getFullYear(), m = today.getMonth(), d = today.getDate();
  var out = [];
  for (var i = 1; i < all.length; i++) {
    var r = all[i];
    if (!r[8]) continue;
    var dv = new Date(r[8]);
    if (dv.getFullYear()!==y || dv.getMonth()!==m || dv.getDate()!==d) continue;
    if (String(r[5]).trim()!==car || String(r[6]).trim()!==name) continue;
    if (r[16]) continue;
    var pickV = (r[10] instanceof Date) ? '' : String(r[10] || '');
    var dropV = (r[11] instanceof Date) ? '' : String(r[11] || '');
    out.push({ row: i+1, pick: pickV, drop: dropV, guideDone: !!r[12], pickDone: !!r[13], dropDone: !!r[16] });
  }
  return out;
}


// ================================================================
//  7-2: 運行シートへの行作成（createParentRows）
//  紐づけアドレスからマスタ情報を取得し運行シートに行程を書き込む
//  ・同じ運行の行程は全て同じIDを付与
//  ・日付をDate型（時刻付き）で書き込む
//  ・LockServiceで同時書き込みによるID重複を防止
// ================================================================
function createParentRows(picks, drops) {
  var savedEmail = PropertiesService.getUserProperties().getProperty('linkedEmail');
  if (!savedEmail) throw new Error('紐づけされていません');

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { throw new Error('混雑中です。少し待ってから再試行してください'); }

  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var master = ss.getSheetByName('自車専属マスタ');
    if (!master) throw new Error('自車専属マスタシートがありません');
    var mAll = master.getDataRange().getValues();
    var info = null;
    for (var j = 1; j < mAll.length; j++) {
      if (String(mAll[j][9]).trim() === savedEmail) {
        info = {
          kubun:mAll[j][2], company:mAll[j][3], tons:mAll[j][4],
          type:mAll[j][5], car:mAll[j][6], name:mAll[j][7], tel:mAll[j][8]
        };
        break;
      }
    }
    if (!info) throw new Error('アドレス未登録');

    var sheet = ss.getSheetByName('運行');
    if (!sheet) throw new Error('運行シートがありません');

    var lastRow  = sheet.getLastRow();
    var nextNum  = getNextIdNum_(sheet, 'V-');
    var now      = new Date();
    var num      = picks.length;
    var startRow = lastRow + 1;
    var sameId   = 'V-' + String(nextNum).padStart(4, '0');

    var rowsData = [];
    for (var i = 0; i < num; i++) {
      rowsData.push([
        sameId, info.kubun, info.company, info.tons, info.type, info.car,
        info.name, info.tel, now, '', picks[i], drops[i],
        '', '', '', '', '', '', '', '', '', '', '', '', ''
      ]);
    }
    sheet.getRange(startRow, 11, num, 2).setNumberFormat('@'); // 積地・降地をテキスト書式に固定
    sheet.getRange(startRow, 1, num, 25).setValues(rowsData);

    var formulas = [];
    for (var i = 0; i < num; i++) {
      var r = startRow + i;
      formulas.push(['=IF(AND(T'+r+'="",S'+r+'=""),"",T'+r+'-S'+r+')']);
    }
    sheet.getRange(startRow, 21, num, 1).setFormulas(formulas);
    sheet.getRange(startRow, 9, num, 1).setNumberFormat('yyyy/MM/dd');
    applyDateTimeFormat_(sheet, startRow, num);

    delaySyncSummary_(sameId);

    var rows = [];
    for (var i = 0; i < num; i++) rows.push(startRow + i);
    return rows;

  } finally {
    lock.releaseLock();
  }
}


// ================================================================
//  7-3: 誘導時刻記録（setGuideComplete）
//  指定行のM列（13列目）に現在時刻を書き込み集計表を同期する
// ================================================================
function setGuideComplete(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  var cell = sheet.getRange(row, 13);
  cell.setValue(new Date());
  cell.setNumberFormat('M/d HH:mm');
  var id = sheet.getRange(row, 1).getValue();
  if (id) delaySyncSummary_(id);
}


// ================================================================
//  7-4: 積完時刻記録（setPickComplete）
//  指定行のN列（14列目）に現在時刻を書き込み集計表を同期する
// ================================================================
function setPickComplete(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  var cell = sheet.getRange(row, 14);
  cell.setValue(new Date());
  cell.setNumberFormat('M/d HH:mm');
  var id = sheet.getRange(row, 1).getValue();
  if (id) delaySyncSummary_(id);
}


// ================================================================
//  7-5: 休憩開始・終了時刻記録（setRest）
//  type='start'→O列(15), type='end'→P列(16) に現在時刻を書き込む
// ================================================================
function setRest(row, type) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  var col = (type === 'start') ? 15 : (type === 'end') ? 16 : 0;
  if (col) {
    var cell = sheet.getRange(row, col);
    cell.setValue(new Date());
    cell.setNumberFormat('M/d HH:mm');
  }
  var id = sheet.getRange(row, 1).getValue();
  if (id) delaySyncSummary_(id);
}


// ================================================================
//  7-6: 降完時刻記録（setDropComplete）
//  指定行のQ列（17列目）に現在時刻を書き込み集計表を同期する
// ================================================================
function setDropComplete(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  var cell = sheet.getRange(row, 17);
  cell.setValue(new Date());
  cell.setNumberFormat('M/d HH:mm');
  var id = sheet.getRange(row, 1).getValue();
  if (id) delaySyncSummary_(id);
}


// ================================================================
//  8-1: 行程データ更新（updateRouteData）
//  戻るボタン用：指定行のK列（積地）・L列（降地）を更新し集計表を同期する
// ================================================================
function updateRouteData(rows, picks, drops) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  if (!sheet) return;
  for (var i = 0; i < rows.length; i++) {
    sheet.getRange(rows[i], 11).setValue(picks[i] || '');
    sheet.getRange(rows[i], 12).setValue(drops[i] || '');
  }
  var id = sheet.getRange(rows[0], 1).getValue();
  if (id) delaySyncSummary_(id);
}


// ================================================================
//  8-2: 運行シート行削除（deleteRunRows）
//  戻るボタン用：指定行番号を降順に削除し集計表を同期する
// ================================================================
function deleteRunRows(rows) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  if (!sheet) return;
  var id = sheet.getRange(rows[0], 1).getValue();
  rows.sort(function(a, b) { return b - a; });
  for (var i = 0; i < rows.length; i++) sheet.deleteRow(rows[i]);
  if (id) delaySyncSummary_(id);
}


// ================================================================
//  8-3: 時刻セルクリア（clearTimeCell）
//  戻るボタン用：指定行・列のセルをクリアし集計表を同期する
// ================================================================
function clearTimeCell(row, col) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  sheet.getRange(row, col).clearContent();
  var id = sheet.getRange(row, 1).getValue();
  if (id) delaySyncSummary_(id);
}


// ================================================================
//  8-4: 運行一覧データ取得（getListData）
//  ・dateSort は運行シートのI列（初回行程登録時刻）を基準にする
//  ・積完時刻・編集で上書きされない
//  ・積完時刻があればdateDispに時刻付加、なければI列の時刻を使用
//  ・集計表から支払い/高速計を取得して返す
// ================================================================
function getListData(year, month) {
  var savedEmail = PropertiesService.getUserProperties().getProperty('linkedEmail');
  if (!savedEmail) return {rows:[], total:{days:0,sales:0,toll:0,pay:0}};

  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  var mAll   = master ? master.getDataRange().getValues() : [];
  var myName = '';
  for (var j = 1; j < mAll.length; j++) {
    if (String(mAll[j][9]).trim().toLowerCase() === savedEmail.toLowerCase()) {
      myName = String(mAll[j][7]).trim();
      break;
    }
  }
  if (!myName) return {rows:[], total:{days:0,sales:0,toll:0,pay:0}};

  var sheet = ss.getSheetByName('運行');
  if (!sheet) return {rows:[], total:{days:0,sales:0,toll:0,pay:0}};

  var all      = sheet.getDataRange().getValues();
  var sumSheet = ss.getSheetByName('集計表');
  var payMap   = {};
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var sumAll = sumSheet.getRange(2, 1, sumSheet.getLastRow()-1, 28).getValues();
    for (var s = 0; s < sumAll.length; s++) {
      var sid = String(sumAll[s][0]||'').trim();
      if (sid) payMap[sid] = {
        pay:      Number(sumAll[s][25]) || 0,
        tollTotal:Number(sumAll[s][20]) || 0
      };
    }
  }

  var idMap = {}, idOrder = [];
  for (var i = 1; i < all.length; i++) {
    var r  = all[i];
    if (!r[8]) continue;
    var dv = r[8] instanceof Date ? r[8] : new Date(r[8]);
    if (isNaN(dv.getTime())) continue;
    var dvYear = dv.getFullYear(), dvMonth = dv.getMonth()+1, dvDate = dv.getDate();
    if (dvYear !== year || dvMonth !== month) continue;
    if (String(r[6]).trim() !== myName) continue;
    var id = String(r[0]||'').trim();
    if (!id) continue;

    if (!idMap[id]) {
      var ds = dvYear+'/'+String(dvMonth).padStart(2,'0')+'/'+String(dvDate).padStart(2,'0');
      // ★dateSort はI列（初回行程登録時刻）を基準に固定・以後変えない
      var baseDateSort = dv.getTime();
      idMap[id] = {
        id:id, car:String(r[5]||'').trim(), date:ds,
        dateSort: baseDateSort,
        dateDisp:'', picks:[], drops:[],
        guideTime:'', pickTime:'', restStart:'', restEnd:'', dropTime:'',
        sales:0, tollReq:0, tollReal:0, tollTotal:0, pay:0,
        notice:r[21]||'', dataUrl:r[22]||'',
        hasNotice:!!(r[21]||r[22]),
        _rawDv: dv
      };
      idOrder.push(id);
    }
    var g = idMap[id];
    if (r[10]) g.picks.push(r[10]);
    if (r[11]) g.drops.push(r[11]);
    if (r[12] && !g.guideTime) {
      var gt = r[12] instanceof Date ? r[12] : new Date(r[12]);
      if (!isNaN(gt.getTime())) g.guideTime = String(gt.getHours()).padStart(2,'0')+':'+String(gt.getMinutes()).padStart(2,'0');
    }
    if (r[13] && !g.pickTime) {
      var pt = r[13] instanceof Date ? r[13] : new Date(r[13]);
      if (!isNaN(pt.getTime())) {
        g.pickTime = String(pt.getHours()).padStart(2,'0')+':'+String(pt.getMinutes()).padStart(2,'0');
        // ★積完時刻でdateSortを上書きしない
      }
    }
    if (r[14] && !g.restStart) {
      var rst = r[14] instanceof Date ? r[14] : new Date(r[14]);
      if (!isNaN(rst.getTime())) g.restStart = String(rst.getHours()).padStart(2,'0')+':'+String(rst.getMinutes()).padStart(2,'0');
    }
    if (r[15] && !g.restEnd) {
      var re2 = r[15] instanceof Date ? r[15] : new Date(r[15]);
      if (!isNaN(re2.getTime())) g.restEnd = String(re2.getHours()).padStart(2,'0')+':'+String(re2.getMinutes()).padStart(2,'0');
    }
    if (r[16] && !g.dropTime) {
      var dt2 = r[16] instanceof Date ? r[16] : new Date(r[16]);
      if (!isNaN(dt2.getTime())) g.dropTime = String(dt2.getHours()).padStart(2,'0')+':'+String(dt2.getMinutes()).padStart(2,'0');
    }
    g.sales   += Number(r[17]) || 0;
    g.tollReq += Number(r[18]) || 0;
    g.tollReal+= Number(r[19]) || 0;
    if (r[21] && !g.notice) g.notice = r[21];
    if (r[22] && !g.dataUrl) g.dataUrl = r[22];
    g.hasNotice = !!(g.notice || g.dataUrl);
  }

  var result = [];
  var totalSales=0, totalToll=0, totalPay=0, dateSet={};
  for (var o = 0; o < idOrder.length; o++) {
    var g  = idMap[idOrder[o]];
    var pm = payMap[g.id] || {pay:0, tollTotal:0};
    g.pay      = pm.pay;
    g.tollTotal= pm.tollTotal;

    // dateDispはI列時刻か積完時刻を表示用に使うが、dateSortは変えない
    var dispTime = g.pickTime;
    if (!dispTime && g._rawDv) {
      var rh = g._rawDv.getHours(), rm = g._rawDv.getMinutes();
      if (rh !== 0 || rm !== 0) {
        dispTime = String(rh).padStart(2,'0')+':'+String(rm).padStart(2,'0');
      }
    }
    g.dateDisp = g.date + (dispTime ? '　'+dispTime : '');

    result.push({
      id:g.id, date:g.date, dateDisp:g.dateDisp, dateSort:g.dateSort,
      car:g.car, pick:g.picks.join('・'), drop:g.drops.join('・'),
      guideTime:g.guideTime, pickTime:g.pickTime, restStart:g.restStart, restEnd:g.restEnd, dropTime:g.dropTime,
      sales:g.sales, tollReq:g.tollReq, tollReal:g.tollReal, tollTotal:g.tollTotal,
      pay:g.pay,
      notice:g.notice, dataUrl:g.dataUrl, hasNotice:g.hasNotice,
      isComplete: !!(g.pickTime && g.restStart && g.restEnd && g.dropTime),
      isNew:      !g.guideTime && !g.pickTime && !g.restStart && !g.restEnd && !g.dropTime
    });
    totalSales += g.sales; totalToll += g.tollTotal; totalPay += g.pay;
    dateSet[g.date] = true;
  }
  result.sort(function(a,b){ return b.dateSort - a.dateSort; });
  return { rows:result, total:{ days:Object.keys(dateSet).length, sales:totalSales, toll:totalToll, pay:totalPay } };
}


// ================================================================
//  8-5: 編集用データ取得（getEditData）
//  運行シートから対象IDの全行を取得して集約する
//  ・同IDの複数行は売上/高速を合算
//  ・積完/休憩/降完は最初に見つかった値を使用
//  ・集計表から合計高速/利益を取得
// ================================================================
function getEditData(id) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return null;

  var sumSheet = ss.getSheetByName('集計表');
  var sumData  = { tollTotal:'', profit:'' };
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var sumAll = sumSheet.getRange(2, 1, sumSheet.getLastRow()-1, 28).getValues();
    for (var s = 0; s < sumAll.length; s++) {
      if (String(sumAll[s][0]||'').trim() === String(id).trim()) {
        sumData.tollTotal = sumAll[s][20] !== '' ? sumAll[s][20] : '';
        sumData.profit    = sumAll[s][26] !== '' ? sumAll[s][26] : '';
        break;
      }
    }
  }

  var all = sheet.getDataRange().getValues();

  // 同IDの全行を収集して集約
  var firstRow = -1;
  var baseData = null;
  var totalSales = 0, totalTollReq = 0, totalTollReal = 0;
  var guideTime='', pickTime='', restStart='', restEnd='', dropTime='';

  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() !== String(id).trim()) continue;
    var r = all[i];

    if (firstRow === -1) {
      firstRow = i + 1;
      baseData = r;
    }

    // 売上・高速は合算
    totalSales   += Number(r[17]) || 0;
    totalTollReq += Number(r[18]) || 0;
    totalTollReal+= Number(r[19]) || 0;

    // 時刻は最初に見つかった値を使用
    if (!guideTime && r[12]) guideTime = Utilities.formatDate(new Date(r[12]),'Asia/Tokyo','HH:mm');
    if (!pickTime  && r[13]) pickTime  = Utilities.formatDate(new Date(r[13]),'Asia/Tokyo','HH:mm');
    if (!restStart && r[14]) restStart = Utilities.formatDate(new Date(r[14]),'Asia/Tokyo','HH:mm');
    if (!restEnd   && r[15]) restEnd   = Utilities.formatDate(new Date(r[15]),'Asia/Tokyo','HH:mm');
    if (!dropTime  && r[16]) dropTime  = Utilities.formatDate(new Date(r[16]),'Asia/Tokyo','HH:mm');
  }

  if (!baseData) return null;

  return {
    row:      firstRow,
    id:       baseData[0],
    kubun:    baseData[1],
    company:  baseData[2],
    tons:     baseData[3],
    type:     baseData[4],
    car:      baseData[5],
    name:     baseData[6],
    tel:      baseData[7],
    date:     baseData[8] ? Utilities.formatDate(new Date(baseData[8]),'Asia/Tokyo','yyyy-MM-dd') : '',
    client:   baseData[9]  || '',
    pick:     baseData[10] || '',
    drop:     baseData[11] || '',
    guideTime: guideTime,
    pickTime:  pickTime,
    restStart: restStart,
    restEnd:   restEnd,
    dropTime:  dropTime,
    sales:    totalSales    || '',
    tollReq:  totalTollReq  || '',
    tollReal: totalTollReal || '',
    tollTotal: sumData.tollTotal,
    notice:   baseData[21] || '',
    dataUrl:  baseData[22] || '',
    termNotice:baseData[23]|| '',
    termData: baseData[24] || '',
    profit:   sumData.profit
  };
}


// ================================================================
//  8-6: 編集データ保存（saveEditData）
//  ・編集可能項目のみ書き込む
//  ・client/pick/dropは空でも上書きしない（削除防止）
//  ・日付はDate型で書き込む（onEditUnkou_の誤発火防止）
//  ・書き込み後に集計表を同期する
// ================================================================
function saveEditData(obj) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  var written = false;
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() !== String(obj.id).trim()) continue;
    var r = i + 1;

    // ★日付はDate型で書き込む（文字列だとonEditが誤発火する）
    if (obj.date) {
      var d = new Date(obj.date);
      // 既存のI列の時刻部分を保持する
      var existing = all[i][8];
      if (existing instanceof Date) {
        d = new Date(d.getFullYear(), d.getMonth(), d.getDate(),
                     existing.getHours(), existing.getMinutes(), existing.getSeconds());
      }
      sheet.getRange(r, 9).setValue(d);
    }

    // ★client/pick/dropは値がある場合のみ書き込む（空で上書きしない）
    if (obj.client) sheet.getRange(r, 10).setValue(obj.client);
    if (obj.pick)   sheet.getRange(r, 11).setValue(obj.pick);
    if (obj.drop)   sheet.getRange(r, 12).setValue(obj.drop);

    var timeFmt = 'M/d HH:mm';
    if (obj.guideTime) { var c13=sheet.getRange(r,13); c13.setValue(new Date(obj.date+' '+obj.guideTime)); c13.setNumberFormat(timeFmt); }
    else               sheet.getRange(r, 13).clearContent();
    if (obj.pickTime)  { var c14=sheet.getRange(r,14); c14.setValue(new Date(obj.date+' '+obj.pickTime));  c14.setNumberFormat(timeFmt); }
    else               sheet.getRange(r, 14).clearContent();
    if (obj.restStart) { var c15=sheet.getRange(r,15); c15.setValue(new Date(obj.date+' '+obj.restStart)); c15.setNumberFormat(timeFmt); }
    else               sheet.getRange(r, 15).clearContent();
    if (obj.restEnd)   { var c16=sheet.getRange(r,16); c16.setValue(new Date(obj.date+' '+obj.restEnd));   c16.setNumberFormat(timeFmt); }
    else               sheet.getRange(r, 16).clearContent();
    if (obj.dropTime)  { var c17=sheet.getRange(r,17); c17.setValue(new Date(obj.date+' '+obj.dropTime));  c17.setNumberFormat(timeFmt); }
    else               sheet.getRange(r, 17).clearContent();

    // 売上・高速は最初の行のみ書き込む（複数行IDの場合の重複防止）
    if (!written) {
      sheet.getRange(r, 18).setValue(obj.sales   || '');
      sheet.getRange(r, 19).setValue(obj.tollReq || '');
      sheet.getRange(r, 20).setValue(obj.tollReal|| '');
      written = true;
    } else {
      // 2行目以降は売上・高速を0にする（合算されるため）
      sheet.getRange(r, 18).setValue('');
      sheet.getRange(r, 19).setValue('');
      sheet.getRange(r, 20).setValue('');
    }

    sheet.getRange(r, 24).setValue(obj.termNotice || '');
  }
  delaySyncSummary_(obj.id);
}


// ================================================================
//  8-7: 運行データ削除（deleteRunById）
//  指定IDに一致する運行シートの全行を削除し集計表を同期する
// ================================================================
function deleteRunById(id) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var all     = sheet.getDataRange().getValues();
  var delRows = [];
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() === String(id).trim()) delRows.push(i+1);
  }
  delRows.sort(function(a,b){ return b-a; });
  for (var i = 0; i < delRows.length; i++) sheet.deleteRow(delRows[i]);
  delaySyncSummary_(id);
}


// ================================================================
//  9-1: 連絡事項保存（saveNotice）
//  指定IDの運行シートU列（21列目）にテキストを書き込む
// ================================================================
function saveNotice(id, text) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() === String(id).trim()) {
      sheet.getRange(i+1, 22).setValue(text); break;
    }
  }
}


// ================================================================
//  9-2: ファイルアップロード・管理側（uploadFile）
//  ファイルをGoogleドライブの「運行データ」フォルダに保存し
//  URLを運行シートのV列（22列目）に書き込む
// ================================================================
function uploadFile(id, fileName, base64Data, mimeType) {
  var folder  = getOrCreateFolder_('運行データ');
  var decoded = Utilities.base64Decode(base64Data);
  var blob    = Utilities.newBlob(decoded, mimeType, fileName);
  var file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url   = file.getUrl();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (sheet) {
    var all = sheet.getDataRange().getValues();
    for (var i = 1; i < all.length; i++) {
      if (String(all[i][0]||'').trim() === String(id).trim()) {
        sheet.getRange(i+1, 23).setValue(url); break;
      }
    }
  }
  return { ok: true, url: url, fileName: fileName };
}


// ================================================================
//  9-3: 端末からの連絡保存（saveTerminalNotice）
//  指定IDの運行シートW列（23列目）にテキストを書き込む
// ================================================================
function saveTerminalNotice(id, text) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() === String(id).trim()) {
      sheet.getRange(i+1, 24).setValue(text); break;
    }
  }
}


// ================================================================
//  9-4: 端末からのファイルアップロード（uploadTerminalFile）
//  ファイルをGoogleドライブの「端末データ」フォルダに保存し
//  URLを運行シートのX列（24列目）に書き込む
// ================================================================
function uploadTerminalFile(id, fileName, base64Data, mimeType) {
  var folder  = getOrCreateFolder_('端末データ');
  var decoded = Utilities.base64Decode(base64Data);
  var blob    = Utilities.newBlob(decoded, mimeType, fileName);
  var file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url   = file.getUrl();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (sheet) {
    var all = sheet.getDataRange().getValues();
    for (var i = 1; i < all.length; i++) {
      if (String(all[i][0]||'').trim() === String(id).trim()) {
        sheet.getRange(i+1, 25).setValue(url); break;
      }
    }
  }
  return { ok: true, url: url, fileName: fileName };
}


// ================================================================
//  10-1: ホーム用連絡事項取得（getMyNotices）
//  紐づけアドレスから乗務員名を特定し未読の連絡事項を最大20件返す
//  U列（連絡事項）またはV列（データURL）がある行が対象
// ================================================================
function getMyNotices() {
  var savedEmail = PropertiesService.getUserProperties().getProperty('linkedEmail');
  if (!savedEmail) return [];
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master) return [];
  var mAll   = master.getDataRange().getValues();
  var myName = '';
  for (var j = 1; j < mAll.length; j++) {
    if (String(mAll[j][9]).trim() === savedEmail) { myName = String(mAll[j][7]).trim(); break; }
  }
  if (!myName) return [];
  var readList = JSON.parse(PropertiesService.getUserProperties().getProperty('readNotices') || '[]');
  var sheet    = ss.getSheetByName('運行');
  if (!sheet) return [];
  var all = sheet.getDataRange().getValues();
  var out = [], seen = {};
  for (var i = 1; i < all.length; i++) {
    var r = all[i];
    if (String(r[6]).trim() !== myName) continue;
    var id = String(r[0]||'').trim();
    if (!id || seen[id]) continue;
    seen[id] = true;
    var notice = r[21]||'', dataUrl = r[22]||'';
    if (!notice && !dataUrl) continue;
    if (readList.indexOf(id) !== -1) continue;
    out.push({ id:id, date: r[8] ? Utilities.formatDate(new Date(r[8]),'Asia/Tokyo','yyyy/MM/dd HH:mm') : '', notice:notice, dataUrl:dataUrl });
  }
  return out.reverse().slice(0, 20);
}


// ================================================================
//  10-2: ID指定行程取得（getRoutesById）
//  指定IDの全行程と進捗状態を返す
//  progress: pick / restStart / restEnd / drop / complete
// ================================================================
function getRoutesById(id) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return { routes:[], progress:'' };
  var all = sheet.getDataRange().getValues();
  var routes = [];
  var allGuideDone=true, anyGuideDone=false;
  var allPickDone=true, anyPickDone=false;
  var hasRestS=false, hasRestE=false;
  var allDropDone=true, anyDropDone=false;
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() !== String(id).trim()) continue;
    var gDone = !!all[i][12], pDone = !!all[i][13], dDone = !!all[i][16];
    var pickVal = (all[i][10] instanceof Date) ? '' : String(all[i][10] || '');
    var dropVal = (all[i][11] instanceof Date) ? '' : String(all[i][11] || '');
    routes.push({ row:i+1, pick:pickVal, drop:dropVal, guideDone:gDone, pickDone:pDone, dropDone:dDone });
    if (!gDone) allGuideDone = false;
    if (gDone)  anyGuideDone = true;
    if (!pDone) allPickDone = false;
    if (pDone)  anyPickDone = true;
    if (all[i][14]) hasRestS = true;
    if (all[i][15]) hasRestE = true;
    if (!dDone) allDropDone = false;
    if (dDone)  anyDropDone = true;
  }
  var progress = 'guide';
  if      (allDropDone && routes.length>0) progress = 'complete';
  else if (anyDropDone)  progress = 'drop';
  else if (hasRestE)     progress = 'drop';
  else if (hasRestS)     progress = 'restEnd';
  else if (allPickDone && routes.length>0) progress = 'restStart';
  else if (anyPickDone)  progress = 'pick';
  else if (allGuideDone && routes.length>0) progress = 'pick';
  else if (anyGuideDone) progress = 'guide';
  return { routes:routes, progress:progress };
}


// ================================================================
//  10-2b: 行番号指定で連絡事項取得（getNoticeByRow）
//  誘導画面に管理側の連絡事項・データURLを表示するために使う
// ================================================================
function getNoticeByRow(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  if (!sheet || !row) return { notice:'', dataUrl:'' };
  var all = sheet.getDataRange().getValues();
  if (all.length < 2) return { notice:'', dataUrl:'' };
  // ヘッダー行で列位置を特定（列追加前後どちらでも対応）
  var headers = all[0];
  var noticeCol = -1, dataCol = -1;
  for (var j = 0; j < headers.length; j++) {
    var h = String(headers[j]).trim();
    if (h === '連絡事項') noticeCol = j;
    if (h === 'データ' || h === 'データURL') dataCol = j;
  }
  if (noticeCol < 0) noticeCol = 21;
  if (dataCol   < 0) dataCol   = 22;
  // 同一IDの行を検索
  var rNum = Number(row) - 1;
  var id = (rNum > 0 && rNum < all.length) ? String(all[rNum][0]||'').trim() : '';
  if (!id) return { notice:'', dataUrl:'' };
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() !== id) continue;
    return { notice: String(all[i][noticeCol]||''), dataUrl: String(all[i][dataCol]||'') };
  }
  return { notice:'', dataUrl:'' };
}


// ================================================================
//  10-3: 既読管理・既読にする（markAsRead）
//  既読にしたIDをPropertiesServiceに保存する（最大200件）
// ================================================================
function markAsRead(id) {
  var p    = PropertiesService.getUserProperties();
  var read = JSON.parse(p.getProperty('readNotices') || '[]');
  if (read.indexOf(id) === -1) {
    read.push(id);
    if (read.length > 200) read = read.slice(-200);
    p.setProperty('readNotices', JSON.stringify(read));
  }
}


// ================================================================
//  10-4: 既読管理・既読一覧取得（getReadNotices）
//  PropertiesServiceから既読IDリストを取得して返す
// ================================================================
function getReadNotices() {
  return JSON.parse(PropertiesService.getUserProperties().getProperty('readNotices') || '[]');
}