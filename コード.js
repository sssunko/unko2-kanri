// ================================================================
// 運行管理システム Code.gs
// ================================================================
// システム概要：
//   Google Apps Script (GAS) 上で動作する運行管理Webアプリのサーバーサイド処理。
//   Google スプレッドシートをデータベースとして使用し、端末（スマートフォン等）から
//   アクセスするWebアプリと連携する。運行データの登録・更新・集計・連絡機能を提供する。
//
// 動作環境：
//   - Google Apps Script（GAS）V8 ランタイム
//   - スプレッドシートにバインドされたコンテナバインドスクリプト
//   - Webアプリとして公開（アクセス権：全員・匿名含む）
//   - 実行アカウント：デプロイユーザー（USER_DEPLOYING）
//
// ================================================================
// ■ 関数番号体系（コード管理用インデックス）
// ================================================================
//
// ── グループ1：共通補助関数 ─────────────────────
//   1-1  : getNextIdNum_(sheet, prefix)
//            各シートのA列から既存IDの最大連番を取得し次の番号を返す
//            ・V-XXXX（運行）/ S-XXXX（マスタ）/ M-XXXX（取引先）に使用
//   1-2  : getOrCreateFolder_(name)
//            Googleドライブに指定名のフォルダを作成または取得して返す
//            ・ファイルアップロード先フォルダの確保に使用
//   1-3  : delaySyncSummary_(id)
//            syncSummaryForId_をtry-catchで囲んで安全に呼び出すラッパー
//            ・onEdit等の短い処理内で集計表同期を呼ぶ際に使用
//   1-4  : cleanAllOrphanSummary_()
//            運行シートに存在しないIDが集計表に残っている場合にその行を削除する
//            ・行削除後やID変更後の集計表クリーンアップに使用
//   1-5  : applyMoneyFormat_(sheet, startRow, numRows, sheetType)
//            指定範囲の金額列（売上・高速・支払等）に #,##0 書式をセットする
//            ・sheetType='unkou'→18〜21列、'summary'→18〜21+24〜27列
//   1-6  : applyDateTimeFormat_(sheet, startRow, numRows)
//            指定範囲の時刻列（誘導〜降完、13〜17列）に 'M/d HH:mm' 書式をセットする
//
// ── グループ2：スプレッドシート起動・表示 ──────────
//   2-1  : onOpen()
//            スプレッドシートを開いた時に実行されるトリガー関数
//            ・上部メニューに「メニュー」を追加（集計表再生成等の操作ボタン）
//            ・convertLegacyAdminDataUrls_を呼び出してW列の古いURL形式を自動変換
//   2-2  : doGet()
//            WebアプリのURLにアクセスした時に実行される関数
//            ・index.htmlをテンプレートとして返しWebアプリを表示する
//   2-3  : showSidebar()
//            スプレッドシート右側のサイドバーとしてWebアプリを表示する
//            ・スプレッドシート上で「ホーム画面を表示」メニューを選んだ時に実行
//   2-4  : showUploadSidebar()
//            運行シートの選択行のW列（管理側データ列）にファイルを直接アップロードするサイドバー
//            ・「📷 写真・ファイル取込」メニューから起動
//            ・GoogleフォトのURLは直接使えないためPCからダウンロードして使うよう案内
//
// ── グループ3：スプレッドシート自動処理（onEdit） ───
//   3-1  : onEdit(e)
//            セル編集時に自動実行されるシンプルトリガー（GAS標準）
//            ・集計表の編集ブロック：距離(V=22)・ガソリン代(X=24)・備考(AB=28)以外は編集不可
//            ・運行シートU列(21)の合計高速は数式を即復元して直接編集を禁止
//            ・運行シートW列(23)にURLを貼り付けた場合はリッチテキストリンクに自動変換
//            ・各シートの専用onEdit関数へ振り分ける
//   3-2  : onEditUnkou_(sheet, range)
//            運行シート編集時の詳細処理
//            ・3-2-1: A列（ID）自動生成 → 他列にデータがあれば V-XXXX 形式で採番
//            ・3-2-2: I列（日付）の時刻補完 → 00:00:00 なら現在時刻を付加
//            ・3-2-3: F列（車番）入力時 → 自車専属マスタから区分〜携帯番号を自動補完
//            ・3-2-4: M〜Q列（時刻列）入力時 → 全角コロン・時刻のみ入力を正規化しDate型で保存
//            ・3-2-5: U列（合計高速）数式を自動セット（=T-S）
//            ・3-2-6: 集計表を該当IDで同期 → 孤立IDを削除
//   3-3  : onEditMasterVehicle_(sheet, range)
//            自車専属マスタ編集時の処理
//            ・3-3-1: A列（ID）自動生成 → S-XXXX 形式
//            ・3-3-2: E列（トン数）変更時 → 設定シートから燃費を自動引き当てK列にセット
//            ・3-3-3: M〜O列（仮日数・給料・%）変更時 → 集計表の該当車番行に即反映・支払再計算
//            ・3-3-4: B列（運行状態）に応じて行の背景色を変更
//              運行=薄赤 / 待機=薄黄 / 故障=薄緑 / それ以外=なし
//            ・3-3-5: 自車専属運行シートを自動更新（運行中の車両のみ抽出）
//   3-4  : onEditMasterCustomer_(sheet, range)
//            取引先マスタ（マスタシート）編集時の処理
//            ・A列（ID）自動生成 → M-XXXX 形式
//
// ── グループ4：集計表・シート構造管理 ───────────────
//   4-1  : generateSummary()
//            集計表シートを運行シートから全件再生成する
//            ・4-1-1: 設定シートからトン数→燃費マップを作成
//            ・4-1-2: 自車専属マスタから車番+乗務員名→仮日数/給料/%マップを作成
//            ・4-1-3: 既存集計表から手入力値（距離・ガソリン代・支払・備考・仮日数等）を退避
//            ・4-1-4: 運行シートをID単位で集約（同IDの複数行は時刻/売上/高速を合算・先勝ち）
//            ・4-1-5: 新しい集計表データを書き込み・数式セット・色付け
//            ・4-1-6: 支払い再計算（4-4）を実行
//            ・4-1-7: W列の旧URL形式をリッチテキストに変換（4-1b）
//   4-1b : convertLegacyAdminDataUrls_()
//            運行シートW列(23)の古い形式のURLをリッチテキストリンクに一括変換する
//            ・プレーンURLセル → setAdminDataRichText_でリッチテキスト化
//            ・リッチテキスト済みセルでノートなし → リンクURLをノートに書き込んで補完
//            ・onOpen・generateSummary の末尾で自動実行
//   4-2  : syncSummaryForId_(targetId)
//            指定IDの行のみ集計表を更新する（リアルタイム同期用）
//            ・4-2-1: 対象IDの運行データを集約
//            ・4-2-2: 仮日数/給料/% はマスタから引き当て（既存値より優先）
//            ・4-2-3: 時刻色付け・利益マイナス赤を再適用
//            ・4-2-4: 数式（T列・X列・Z列）を再セット
//   4-3  : expandAndRefreshSheets()
//            自車専属マスタ・自車専属運行シートの列構成を最新版に整備する
//            ・仮日数/給料/% 列がなければ追加
//            ・自車専属運行シートをマスタの運行中車両から再生成
//   4-4  : calculatePaymentAmount()
//            集計表の支払い列（Z=26列）を計算ルールに従い更新する
//            ・パターンA: AC列(%)あり → (売上 - 合計高速代) × % ÷ 100
//            ・パターンB: AD列(給料)とAC列(仮日数)あり → 給料 ÷ 仮日数
//            ・パターンC: 条件なし → 手入力値を保持（なければ赤背景で警告）
//   4-5  : refreshActiveVehiclesAuto_()
//            自車専属運行シートを自車専属マスタの「運行」中の車両のみで再生成する内部処理
//   4-6  : addStatusColumnToMaster()
//            自車専属マスタのB列に「運行状態」列が存在しない場合に追加するメンテナンス関数
//
// ── グループ5：端末アプリ 起動・紐づけ ──────────────
//   5-1  : getInitialData()
//            端末アプリ起動時に1回だけ呼ばれ、初期表示に必要な全データを一括返却する
//            ・紐づけ済みメールアドレスから乗務員名・車番等を取得
//            ・未読連絡事項リストを取得して返却
//   5-2  : linkAddress(email)
//            端末とメールアドレスを紐づけてPropertiesServiceに保存する
//            ・紐づけ後は端末固有のデータ（運行データ等）を取得できるようになる
//   5-3  : unlinkAddress()
//            紐づけを解除してPropertiesServiceの保存値を削除する
//
// ── グループ6：端末アプリ 運行進捗管理 ──────────────
//   6-1  : saveRunState(state)
//            端末の運行進捗状態（どの行程まで完了したか等）をPropertiesServiceに保存する
//            ・端末を閉じても運行途中から再開できるようにする
//   6-2  : loadRunState()
//            保存された運行進捗状態をPropertiesServiceから読み込んで返す
//   6-3  : clearRunState()
//            保存された運行進捗状態をPropertiesServiceから削除する（運行完了時に実行）
//
// ── グループ7：端末アプリ 運行操作 ──────────────────
//   7-1  : getTodayRoutes()
//            自車専属運行シートから今日の行程一覧を取得して返す
//            ・当日日付の行のみ抽出・荷主/積地/降地/売上等を含む
//   7-2  : createParentRows(routes)
//            行程データを運行シートに新規行として書き込む
//            ・7-2-1: LockServiceでIDの重複採番を防止（同時アクセス対策）
//            ・7-2-2: 日付はDate型で書き込み（文字列だとonEditが誤発火）
//            ・7-2-3: 合計高速の数式をU列にセット
//   7-3  : setPickComplete(id, rowNum, time)
//            指定IDの指定行に積完時刻を記録する
//   7-4  : setRest(id, type, time)
//            指定IDに休憩開始または休憩終了時刻を記録する（type='start'/'end'）
//   7-5  : setDropComplete(id, rowNum, time)
//            指定IDの指定行に降完時刻を記録する
//
// ── グループ8：端末アプリ 運行一覧・編集 ─────────────
//   8-1  : updateRouteData(obj)
//            行程の積地・降地・売上・高速代を運行シートに上書き保存する
//   8-2  : deleteRunRows(id)
//            指定IDの全行を運行シートから削除し集計表も更新する
//   8-3  : clearTimeCell(id, colNum)
//            指定IDの指定列（時刻セル）の内容をクリアする
//   8-4  : getListData(year, month)
//            端末アプリの一覧画面用データを月単位で取得する
//            ・8-4-1: 紐づけメールから乗務員名を特定
//            ・8-4-2: 運行シートを月・乗務員名で絞り込みID単位に集約
//            ・8-4-3: W列(23)のデータURLをノート→リッチテキスト→プレーンの優先順で取得
//            ・8-4-4: 集計表から支払/高速計を引き当て
//            ・8-4-5: 各IDの積完時刻またはI列時刻を表示用に整形
//            ・8-4-6: 月集計（稼働日数・売上・高速・支払）を合算して返却
//   8-5  : getEditData(id)
//            編集モーダル表示用に指定IDの詳細データを取得する
//            ・同IDの複数行は売上/高速を合算、時刻は先勝ち
//            ・W列(23)のデータURLはgetAdminDataUrl_で取得（リッチテキスト対応）
//            ・Y列(25)の端末データURLはgetTerminalUrls_で取得
//   8-6  : saveEditData(obj)
//            編集モーダルで変更された値を運行シートに書き込む
//            ・8-6-1: 日付はDate型で書き込む（既存時刻を保持）
//            ・8-6-2: 荷主名/積地/降地をnullでなければ上書き
//            ・8-6-3: 時刻はDate型で合成して書き込み（空の場合はクリア）
//            ・8-6-4: 売上/高速は複数行IDの場合は先頭行のみ書き込み
//            ・8-6-5: 集計表を該当IDで同期
//   8-6a : setAdminDataRichText_(sheet, rowNum, url)
//            1件のURLをW列にリッチテキストリンク「ファイル1」として書き込む
//   8-6b : setAdminDataRichTextMulti_(sheet, rowNum, urls)
//            複数URLをW列にリッチテキストリンク「ファイル1  ファイル2...」として書き込む
//            ・URLをセルノートにも保存（getNotes()で一括読み取り可能にするため）
//   8-6b-1: importImageToDrive_(url)
//            外部URL（GoogleフォトなどのURL）の画像をGoogleドライブに取り込み公開URLを返す
//            ・外部URLは端末から直接開けない場合があるためDriveに保存して変換
//   8-6b-2: getTerminalUrls_(sheet, rowNum)
//            Y列(25)のリッチテキストからリンクURL一覧を取得して返す
//   8-6b-3: setTerminalUrls_(sheet, rowNum, urls)
//            複数URLをY列(25)にリッチテキストリンク「ファイル1  ファイル2...」として書き込む
//   8-6b-4: getAdminDataUrl_(sheet, rowNum)
//            W列(23)のリッチテキストからURLをカンマ区切り文字列で返す（プレーン値フォールバックあり）
//   8-6c : appendTerminalFile(id, fileName, base64Data, mimeType)
//            Base64データをファイル化してDriveに保存しY列(25)のリッチテキストURLに追記する
//            ・「端末データ」フォルダに保存・誰でも閲覧可能リンクを設定
//   8-7  : deleteRunById(id)
//            指定IDの全行を削除し集計表を再生成する
//
// ── グループ9：端末アプリ 連絡・ファイル ─────────────
//   9-1  : saveNotice(id, notice)
//            指定IDの運行シートV列(22)（管理側連絡事項）にテキストを保存する
//   9-2  : uploadFileToRow(rowNum, fileName, base64Data, mimeType)
//            ファイルをDriveに保存してW列(23)のリッチテキストに追記する（管理側アップロード）
//            ・「運行データ」フォルダに保存
//   9-3  : saveTerminalNotice(id, text)
//            指定IDの運行シートX列(24)（端末側連絡事項）にテキストを保存する
//   9-4  : uploadTerminalFile(id, fileName, base64Data, mimeType)
//            appendTerminalFile(8-6c)のエイリアス（旧バージョン互換）
//
// ── グループ10：端末アプリ 連絡事項・既読管理 ──────────
//   10-1 : getMyNotices()
//            ホーム画面の未読連絡事項一覧を返す
//            ・紐づけメールから乗務員名を特定
//            ・V列(22)=管理側連絡事項またはW列(23)=データURLがある行が対象
//            ・W列のURLはノート→リッチテキスト→プレーンの優先順で取得
//            ・readNoticesリストと照合して既読済みはスキップ
//            ・最新20件を返す
//   10-2 : getRoutesById(id)
//            指定IDの全行程と進捗状態（guide/pick/restStart/restEnd/drop/complete）を返す
//   10-2b: getNoticeByRow(row)
//            行番号を指定して連絡事項とデータURLを返す（誘導画面の連絡表示用）
//   10-3 : markAsRead(id)
//            指定IDを既読済みとしてPropertiesServiceのreadNoticesリストに追加する
//   10-4 : getReadNotices()
//            既読済みIDの一覧をPropertiesServiceから読み込んで返す
//
// ================================================================
// ■ スプレッドシート シート構成
// ================================================================
//
//   【設定シート】（2列）
//     A列: トン数（例: 1t, 2t, 4t）
//     B列: 基準燃費（L/km）
//     ※ トン数に対応する燃費がここから引き当てられる
//
//   【運行シート】（25列）
//     列番号: 1=ID, 2=区分, 3=会社名, 4=トン数, 5=車種, 6=車番, 7=乗務員名, 8=携帯番号,
//             9=日付, 10=荷主名, 11=積地, 12=降地,
//             13=誘導時刻, 14=積完時刻, 15=休憩開始, 16=休憩終了, 17=降完時刻,
//             18=売上, 19=請求高速, 20=実費高速, 21=合計高速（数式: =T-S）,
//             22=連絡事項（管理→端末）, 23=データ（管理側ファイルURLリッチテキスト）,
//             24=連絡端末（端末→管理）, 25=データ端末（端末ファイルURLリッチテキスト）
//     ※ 1つのIDに複数行（行程数分）が紐づく。時刻は各行程ごとではなく先頭行に集約。
//
//   【集計表シート】（31列）
//     列番号: 1〜8=運行シートと同じ基本情報,
//             9〜21=日付〜合計高速（運行シートから集約）,
//             22=距離（手入力）, 23=燃費（設定シート参照）, 24=ガソリン代（手入力）,
//             25=燃料代（数式: =距離÷燃費×ガソリン代）,
//             26=支払い（数式または手入力: 4-4で計算）,
//             27=利益（数式: =売上−(合計高速+燃料代+支払)）,
//             28=備考（手入力）,
//             29=仮日数（自車専属マスタから引き当て）,
//             30=給料（自車専属マスタから引き当て）,
//             31=%（自車専属マスタから引き当て）
//
//   【自車専属マスタシート】（15列）
//     列番号: 1=車両ID(S-XXXX), 2=運行状態, 3=区分, 4=会社名, 5=トン数, 6=車種,
//             7=車番, 8=乗務員名, 9=携帯番号, 10=メールアドレス（端末紐づけ用）,
//             11=燃費, 12=備考, 13=仮日数, 14=給料, 15=%
//     ※ B列(2)の運行状態が「運行」の行のみ自車専属運行シートに抽出される
//
//   【自車専属運行シート】（15列）
//     自車専属マスタから「運行」状態の車両のみ抽出した参照用シート
//
//   【マスタシート（取引先）】（14列）
//     列番号: 1=ID(M-XXXX), 2=会社名, 3=電話, 4=FAX, 5=郵便番号, 6=住所,
//             7=代表者, 8=配車担当, 9=銀行名, 10=支店名, 11=種別, 12=番号, 13=名義, 14=備考
//
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
  var fmt = '#,##0;[RED]#,##0';
  var cols = (sheetType === 'unkou')
    ? [18, 19, 20, 21]
    : [18, 19, 20, 21, 24, 25, 26, 27, 32, 33];
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
//  1-7: 積地（K列=11）休み・有休の背景色設定（applyHolidayRowColors_）
//  積地セルに「休み」または「有休」が含まれる行のK列を灰色（#9e9e9e）に着色する
//  onOpen・generateSummary・onEdit から呼び出す
// ================================================================
function applyHolidayRowColors_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 運行シートのK列(11=積地)：休み/有休はグレー
  var sheet = ss.getSheetByName('運行');
  if (sheet && sheet.getLastRow() >= 2) {
    var lr = sheet.getLastRow();
    var vals = sheet.getRange(2, 11, lr - 1, 1).getValues();
    var bgs = vals.map(function(r) {
      var v = String(r[0] || '');
      return [(v.indexOf('休み') !== -1 || v.indexOf('有休') !== -1) ? '#9e9e9e' : null];
    });
    sheet.getRange(2, 11, lr - 1, 1).setBackgrounds(bgs);
  }
  // 集計表のK列(11=積地)も同様にグレー着色
  var sumSheet = ss.getSheetByName('集計表');
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var slr = sumSheet.getLastRow();
    var svals = sumSheet.getRange(2, 11, slr - 1, 1).getValues();
    var sbgs = svals.map(function(r) {
      var v = String(r[0] || '');
      return [(v.indexOf('休み') !== -1 || v.indexOf('有休') !== -1) ? '#9e9e9e' : null];
    });
    sumSheet.getRange(2, 11, slr - 1, 1).setBackgrounds(sbgs);
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
    .addItem('シート保護設定', 'setupSheetProtection')
    .addItem('📷 写真・ファイル取込', 'showUploadSidebar')
    .addToUi();
  convertLegacyAdminDataUrls_();
  applyHolidayRowColors_();
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
//  2-4: 写真・ファイル取込サイドバー（showUploadSidebar）
//  運行シートの行を選択してメニューから起動 → W列に直接アップロード
// ================================================================
function showUploadSidebar() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row   = sheet.getActiveCell().getRow();
  if (sheet.getName() !== '運行' || row <= 1) {
    SpreadsheetApp.getUi().alert('運行シートのデータ行を選択してから実行してください');
    return;
  }
  var id = String(sheet.getRange(row, 1).getValue()).trim();
  if (!id) { SpreadsheetApp.getUi().alert('IDが空の行です'); return; }

  var body =
    '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>body{font-family:sans-serif;background:#121212;color:#e0e0e0;padding:16px;margin:0;}' +
    'h3{color:#bb86fc;margin:0 0 12px;}p{font-size:13px;color:#aaa;margin:0 0 12px;}' +
    'input[type=file]{display:block;width:100%;margin-bottom:12px;color:#e0e0e0;box-sizing:border-box;}' +
    'button{width:100%;padding:14px;background:#1565c0;color:white;border:none;border-radius:10px;font-size:15px;font-weight:bold;cursor:pointer;margin-bottom:8px;}' +
    'button:active{background:#0d47a1;}' +
    '#msg{font-size:13px;margin-top:10px;min-height:20px;}</style></head>' +
    '<body>' +
    '<h3>📷 ファイル取込</h3>' +
    '<p>行 <b>' + row + '</b>（ID: ' + id + '）のデータ列に追加</p>' +
    '<p>Google フォトの写真は、フォトアプリで写真を長押し→共有→ダウンロードしてからここで選択してください</p>' +
    '<input type="file" id="f" multiple accept="image/*,.pdf,.doc,.docx,.xls,.xlsx">' +
    '<button onclick="upload()">アップロード</button>' +
    '<div id="msg"></div>' +
    '<script>' +
    'function upload(){' +
    '  var files=Array.from(document.getElementById("f").files);' +
    '  if(!files.length){document.getElementById("msg").innerText="ファイルを選択してください";return;}' +
    '  document.getElementById("msg").innerText="アップロード中... 0/"+files.length;' +
    '  var done=0;' +
    '  files.forEach(function(file){' +
    '    if(file.size>20*1024*1024){done++;document.getElementById("msg").innerText=done+"/"+files.length+" 完了";return;}' +
    '    var r=new FileReader();' +
    '    r.onload=function(){' +
    '      var b64=r.result.split(",")[1];' +
    '      google.script.run' +
    '        .withSuccessHandler(function(){done++;document.getElementById("msg").innerText=done===files.length?"✅ 完了！":done+"/"+files.length+" 完了";})' +
    '        .withFailureHandler(function(e){done++;document.getElementById("msg").innerText="エラー："+e.message;})' +
    '        .uploadFileToRow(' + row + ',file.name,b64,file.type);' +
    '    };r.readAsDataURL(file);' +
    '  });' +
    '}' +
    '<\/script></body></html>';

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(body).setTitle('📷 写真・ファイル取込').setWidth(280)
  );
}


// ================================================================
//  3-1: onEdit本体
//  セル編集時にGASが自動実行するシンプルトリガー（イベントオブジェクト e を受け取る）
//  編集されたシート名を判定し、対応する処理関数へ振り分ける
//  ※ シンプルトリガーはユーザーの明示的な編集のみ発火し、GASからの書き込みでは発火しない
// ================================================================
function onEdit(e) {
  try {
    var range     = e.range;
    var sheet     = range.getSheet();
    var sheetName = sheet.getName();
    var col       = range.getColumn();
    var row       = range.getRow();
    var ss        = SpreadsheetApp.getActiveSpreadsheet();

    // ── 3-1-1: 集計表 編集ブロック ──────────────────────────────────
    // 距離(V=22)・ガソリン代(X=24)・支払い(Z=26)・備考(AB=28)以外の列は直接編集を禁止する
    // 禁止列が触れられた場合: IDがあれば集計表を再同期して正しい値に戻す
    //                         IDがなく単一セルなら旧値に戻す
    if (sheetName === '集計表' && row > 1) {
      var allowed = [22, 24, 26, 28, 33]; // V=距離, X=ガソリン代, Z=支払い, AB=備考, AG=その他手当
      var numC = range.getNumColumns(), numR = range.getNumRows();
      var blocked = false;
      for (var c = 0; c < numC; c++) {
        if (allowed.indexOf(col + c) === -1) { blocked = true; break; }
      }
      if (blocked) {
        var bid = String(sheet.getRange(row, 1).getValue() || '').trim();
        if (bid) { try { syncSummaryForId_(bid); } catch(ex) {} }
        else if (numR === 1 && numC === 1) { range.setValue(e.oldValue !== undefined ? e.oldValue : ''); }
        ss.toast('この列は編集できません', '⛔ 保護', 3);
      }
      return;
    }

    // ── 3-1-2: 運行シート U列(21) 合計高速 保護 ──────────────────────
    // 合計高速は「実費高速 - 請求高速」の自動計算列なので直接編集を禁止する
    // 編集されたら即座に数式を復元する
    if (sheetName === '運行' && row > 1 && col === 21) {
      range.setFormula('=IF(AND(T'+row+'="",S'+row+'=""),"",T'+row+'-S'+row+')');
      ss.toast('合計高速は自動計算列です', '⛔ 保護', 3);
      return;
    }

    // ── 3-1-3: 運行シート W列(23) URL自動変換 ───────────────────────
    // W列にURLをペーストすると自動的にリッチテキストリンク（「ファイル1」等）に変換する
    // 改行・全角/半角カンマ・スペース区切りで複数URLに対応
    // DriveやDocs以外のURL（Googleフォト等）はDriveに取り込んで公開URLに変換を試みる
    // 既存のURLはノートに保存されており、新しいURLと結合して上書きすることで追加になる
    if (sheetName === '運行' && row > 1 && col === 23) {
      var val = String(e.value !== undefined ? e.value : range.getValue() || '').trim();
      if (!val) { range.clearNote(); return; }
      var rawUrls = val.split(/[\n,，\s]+/).map(function(u){return u.trim();})
        .filter(function(u){return u.match(/^https?:\/\//);});
      if (rawUrls.length > 0) {
        var converted = rawUrls.map(function(u){
          if (!u.match(/drive\.google\.com|docs\.google\.com/)) {
            var driveUrl = importImageToDrive_(u);
            return driveUrl || u;
          }
          return u;
        });
        var note = range.getNote() || '';
        var existing = note ? note.split('\n').filter(function(u){return u.match(/^https?:\/\//);}) : [];
        // 重複URLを除去してから書き込み
        var merged = existing.concat(converted).filter(function(u, i, arr) { return arr.indexOf(u) === i; });
        setAdminDataRichTextMulti_(sheet, row, merged);
      }
      return;
    }

    // ── 3-1-4: シート別振り分け ──────────────────────────────────────
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
    // A列(1)が空でB〜K列にデータがあれば V-XXXX 形式のIDを自動採番
    var idCell = sheet.getRange(row, 1);
    var currentId = idCell.getValue();
    if (!currentId) {
      var hasData = sheet.getRange(row, 2, 1, 10).getValues()[0].some(function(v) { return v !== ''; });
      if (hasData) {
        var nextNum = getNextIdNum_(sheet, 'V-');
        idCell.setValue('V-' + String(nextNum).padStart(4, '0'));
      }
    }
    // I列(9)の日付：時刻部分が 0:00:00 なら現在時刻を付与（日付のみ入力に対応）
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
    // F列(6)：車番を入力→自車専属マスタと部分一致で他項目を自動補完
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
    // 積地(K=col11) に「休み」「有休」が含まれる場合はセルをグレーに着色
    var pvK = String(sheet.getRange(row, 11).getValue() || '');
    sheet.getRange(row, 11).setBackground(
      (pvK.indexOf('休み') !== -1 || pvK.indexOf('有休') !== -1) ? '#9e9e9e' : null
    );
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
  var numRows  = range.getNumRows();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 設定シートからトン数→燃費マップを取得（正規化: 全角数字→半角, 大文字小文字統一, 't'有無両対応）
  var settingSheet = ss.getSheetByName('設定');
  var fuelMap = {};
  if (settingSheet && settingSheet.getLastRow() >= 2) {
    var sVals = settingSheet.getRange(2, 1, settingSheet.getLastRow()-1, 4).getValues();
    for (var s = 0; s < sVals.length; s++) {
      var skey = String(sVals[s][0]||'').trim()
        .replace(/[０-９]/g, function(c){return String.fromCharCode(c.charCodeAt(0)-0xFEE0);})
        .replace(/[ｔＴ]/g,'t').toLowerCase();
      if (skey && skey !== '有休') {
        var numOnly = skey.replace(/t$/,'');
        fuelMap[skey] = sVals[s][1];          // "4t"
        fuelMap[numOnly] = sVals[s][1];       // "4"
        fuelMap[numOnly+'t'] = sVals[s][1];   // "4t" (念のため重複)
      }
    }
  }
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
    // E列（col5）のトン数に対応する燃費をK列（col11）に自動反映
    var tonsRaw = String(sheet.getRange(row, 5).getValue()).trim();
    if (tonsRaw) {
      var tonsNorm = tonsRaw
        .replace(/[０-９]/g, function(c){return String.fromCharCode(c.charCodeAt(0)-0xFEE0);})
        .replace(/[ｔＴ]/g,'t').toLowerCase();
      var numPart = tonsNorm.replace(/t$/,'');
      var fuel = fuelMap[tonsNorm] || fuelMap[numPart+'t'] || fuelMap[numPart] || '';
      if (fuel !== '' && fuel !== undefined) sheet.getRange(row, 11).setValue(fuel);
    }
    // 仮日数/給料/%が変わったら集計表の該当行（車番+乗務員名一致）に即反映
    var mRow = sheet.getRange(row, 1, 1, 15).getValues()[0];
    var mCar    = String(mRow[6]  || '').trim();
    var mName   = String(mRow[7]  || '').trim();
    var mKari   = mRow[12];
    var mKyuryo = mRow[13];
    var mPct    = mRow[14];
    if (mCar || mName) {
      var sumSheet = ss.getSheetByName('集計表');
      if (sumSheet && sumSheet.getLastRow() >= 2) {
        var sumData = sumSheet.getRange(2, 1, sumSheet.getLastRow()-1, 31).getValues();
        for (var s = 0; s < sumData.length; s++) {
          var sCar  = String(sumData[s][6] || '').trim();
          var sName = String(sumData[s][7] || '').trim();
          if (sCar === mCar && sName === mName) {
            var sRow = s + 2;
            if (mKari   !== '') sumSheet.getRange(sRow, 29).setValue(mKari);
            if (mKyuryo !== '') sumSheet.getRange(sRow, 30).setValue(mKyuryo);
            if (mPct    !== '') sumSheet.getRange(sRow, 31).setValue(mPct);
          }
        }
        calculatePaymentAmount();
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
//  運行シート全件から集計表を一から作り直す（全件対象の重い処理）
//  メニューの「集計表再生成」または自動実行（generateSummary→ボタン）で実行される
//
//  処理の流れ：
//  ① 設定シート → トン数ごとの燃費をマップ化
//  ② 自車専属マスタ → 車番+乗務員名ごとの仮日数/給料/%をマップ化
//  ③ 既存の集計表 → 手入力値（距離・ガソリン代・支払・備考・仮日数等）を退避
//  ④ 運行シートをID単位で集約（同IDの複数行は売上/高速合算・時刻は先勝ち）
//  ⑤ 集計表を全クリアして新しいデータを書き込み
//  ⑥ 各行に数式をセット（U列:合計高速 / Y列:燃料代 / AA列:利益）
//  ⑦ 時刻の間隔異常を色で警告（積完〜休憩4時間超=黄 / 休憩30分未満=水 / 休憩後〜降完4時間超=黄）
//  ⑧ 利益がマイナスの行を薄赤で着色
//  ⑨ 支払い再計算（4-4）を実行して支払額を更新
//  ⑩ W列の旧URL形式をリッチテキストに変換
//  ・生成後に支払い再計算（4-4）を自動実行
// ================================================================
function generateSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var unkouSheet = ss.getSheetByName('運行');
  if (!unkouSheet) return;

  // 設定シートからトン数→燃費マップと有休日額を取得
  var settingSheet = ss.getSheetByName('設定');
  var fuelMap = {};
  var yukyuRate = 0;
  if (settingSheet && settingSheet.getLastRow() >= 2) {
    var sVals = settingSheet.getRange(2, 1, settingSheet.getLastRow()-1, 4).getValues();
    for (var s = 0; s < sVals.length; s++) {
      var skey = String(sVals[s][0]||'').trim();
      if (skey) { fuelMap[skey] = sVals[s][1]; }
      // C列=有休ラベル, D列=有休日額
      if (String(sVals[s][2]||'').trim() === '有休') { yukyuRate = Number(sVals[s][3]) || 0; }
    }
  }

  // 自車専属マスタから 車番+乗務員名 → 仮日数/給料/% の支払条件マップを作成
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

  // 既存の集計表から手入力済みの距離・ガソリン代・支払・備考・支払条件を退避（再生成で消えないように）
  var sumSheet = ss.getSheetByName('集計表');
  var oldData = {};
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var colCount = sumSheet.getLastColumn();
    var oldRows = sumSheet.getRange(2, 1, sumSheet.getLastRow()-1, Math.max(colCount, 33)).getValues();
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
          pct:      oldRows[o][30] || '',
          other:    oldRows[o][32] || ''  // AG=その他手当（手入力保持）
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
    '仮日数','給料','％','有休手当','その他手当'
  ];

  // 運行シートを全行読み込み、ID単位にデータを集約する
  // 同一IDに複数行ある場合（複数行程）は積地/降地を連結、金額は合算
  var unkouData = unkouSheet.getDataRange().getValues();
  var idMap = {}, idOrder = [];
  for (var i = 1; i < unkouData.length; i++) {
    var r  = unkouData[i];
    var id = String(r[0] || '').trim();
    if (!id) continue;
    if (!idMap[id]) {
      idMap[id] = {
        id:id, kubun:r[1], company:r[2], tons:r[3], type:r[4], car:r[5],
        name:r[6], tel:r[7], date:r[8], clients:[],
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
    // 荷主は重複なしで全部収集（行程ごとに荷主が異なる場合も全て・区切りで表示）
    if (r[9]) { var rc = String(r[9]); if (g.clients.indexOf(rc) === -1) g.clients.push(rc); }
    // 時刻は先勝ち（最初に見つかった値を使用）
    if (r[12] && !g.guideTime) { g.guideTime = r[12]; }
    if (r[13] && !g.pickTime)  { g.pickTime  = r[13]; g.rawPickTime  = new Date(r[13]); }
    if (r[14] && !g.restStart) { g.restStart = r[14]; g.rawRestStart = new Date(r[14]); }
    if (r[15] && !g.restEnd)   { g.restEnd   = r[15]; g.rawRestEnd   = new Date(r[15]); }
    if (r[16] && !g.dropTime)  { g.dropTime  = r[16]; g.rawDropTime  = new Date(r[16]); }
    // 売上・高速は複数行程分を合算
    g.sales   += Number(r[17]) || 0;
    g.tollReq += Number(r[18]) || 0;
    g.tollReal+= Number(r[19]) || 0;
  }

  // 集計表に書き出す行データを組み立て
  // 距離・ガソリン代・支払いは退避データを復元（再生成で消えない）
  // 支払条件（仮日数/給料/%）はマスタ優先、次に退避データ
  var outRows = [header];
  for (var o2 = 0; o2 < idOrder.length; o2++) {
    var g      = idMap[idOrder[o2]];
    var tonsStr= String(g.tons || '').trim();
    var fuel   = fuelMap[tonsStr] || fuelMap[tonsStr.replace(/[tT]/,'')+'t'] || 3;
    var old    = oldData[g.id] || {};
    var pkey   = String(g.car||'').trim() + '_' + String(g.name||'').trim();
    var pc     = payCondMap[pkey] || {kari:'', kyuryo:'', pct:''};
    var kari   = (pc.kari   !== undefined && pc.kari   !== '') ? pc.kari   : (old.kari   || '');
    var kyuryo = (pc.kyuryo !== undefined && pc.kyuryo !== '') ? pc.kyuryo : (old.kyuryo || '');
    var pct    = (pc.pct    !== undefined && pc.pct    !== '') ? pc.pct    : (old.pct    || '');

    var gpick = g.picks.join('・'), gdrop = g.drops.join('・');
    var gIsYukyu = gpick.indexOf('有休') !== -1 || gdrop.indexOf('有休') !== -1;
    outRows.push([
      g.id, g.kubun, g.company, g.tons, g.type, g.car, g.name, g.tel,
      g.date, g.clients.join('・'), gpick, gdrop,
      g.guideTime||'', g.pickTime, g.restStart, g.restEnd, g.dropTime,
      g.sales||'', g.tollReq||'', g.tollReal||'', '',
      old.distance||'', fuel, old.gas||'', '',
      old.pay||'', '', old.memo||'',
      kari, kyuryo, pct,
      gIsYukyu ? yukyuRate : '',
      old.other || ''   // AG=その他手当（手入力保持）
    ]);
  }

  // 集計表を全クリアして再書き込み
  sumSheet.clear();
  if (outRows.length > 0) {
    sumSheet.getRange(1, 1, outRows.length, 33).setValues(outRows);
    sumSheet.setFrozenRows(1);

    // 4時間超で黄色（労働時間過超）、30分未満で水色（休憩不足）の判定閾値
    var F = 4*60*60*1000;
    var T = 30*60*1000;

    for (var row = 2; row <= outRows.length; row++) {
      // U列(21)：合計高速代 = 実費(T) - 請求(S)（差額がマイナスなら持ち出し）
      sumSheet.getRange(row, 21).setFormula('=IF(AND(T'+row+'="",S'+row+'=""),"",T'+row+'-S'+row+')');
      // Y列(25)：燃料代 = 距離(V) / 燃費(W) * ガソリン代(X)
      sumSheet.getRange(row, 25).setFormula('=IF(OR(V'+row+'="",W'+row+'=""),"",V'+row+'/W'+row+'*X'+row+')');
      // AA列(27)：利益 = 売上(R) - (合計高速代(U) + 燃料代(Y) + 支払い(Z))
      sumSheet.getRange(row, 27).setFormula('=IF(AND(R'+row+'="",U'+row+'="",Y'+row+'="",Z'+row+'=""),"",R'+row+'-(U'+row+'+Y'+row+'+Z'+row+'))');

      var g2       = idMap[idOrder[row-2]];
      var keepPay  = outRows[row-1][25] || '';
      var keepDist = outRows[row-1][21] || '';
      var keepGas  = outRows[row-1][23] || '';
      var calcToll = (Number(g2.tollReal)||0)-(Number(g2.tollReq)||0);
      var calcFuel = (Number(keepDist)&&Number(fuel)&&Number(keepGas)) ? (Number(keepDist)/Number(fuel)*Number(keepGas)) : 0;
      var calcProfit = (Number(g2.sales)||0)-(calcToll+calcFuel+(Number(keepPay)||0));

      // 利益がマイナスの行は薄赤で警告表示
      sumSheet.getRange(row, 1, 1, 31).setBackground(calcProfit < 0 ? '#ffebee' : null);
      // 積完〜降完間の労働時間・休憩時間を判定して背景色で警告
      sumSheet.getRange(row, 14, 1, 4).setBackground(null);
      if (g2.rawPickTime  && g2.rawRestStart && (g2.rawRestStart-g2.rawPickTime)  > F) { sumSheet.getRange(row,14,1,2).setBackground('#ffd600'); }
      if (g2.rawRestStart && g2.rawRestEnd   && (g2.rawRestEnd  -g2.rawRestStart) < T) { sumSheet.getRange(row,15,1,2).setBackground('#4fc3f7'); }
      if (g2.rawRestEnd   && g2.rawDropTime  && (g2.rawDropTime -g2.rawRestEnd)   > F) { sumSheet.getRange(row,16,1,2).setBackground('#ffd600'); }
    }
    applyMoneyFormat_(sumSheet, 2, outRows.length - 1, 'summary');
    applyDateTimeFormat_(sumSheet, 2, outRows.length - 1);
  }

  calculatePaymentAmount();
  convertLegacyAdminDataUrls_();
  applyHolidayRowColors_();
}


// ================================================================
//  4-1b: 管理側データURLをリッチテキストに一括変換（convertLegacyAdminDataUrls_）
//  運行シートのW列(23)にプレーンURLが残っている行をリッチテキストに変換
// ================================================================
function convertLegacyAdminDataUrls_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastRow = sheet.getLastRow();
  var all = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    var val = String(all[i][22] || '').trim();
    if (val.match(/^https?:\/\//)) {
      setAdminDataRichText_(sheet, i + 1, val);
    } else if (val && !val.match(/^https?:\/\//)) {
      // Already rich text ("ファイル1" etc.) — ensure note is populated
      var cell = sheet.getRange(i + 1, 23);
      if (!cell.getNote()) {
        var rtv = cell.getRichTextValue();
        if (rtv) {
          var runs = rtv.getRuns(), urls = [];
          for (var k = 0; k < runs.length; k++) {
            var lk = runs[k].getLinkUrl();
            if (lk) urls.push(lk);
          }
          if (urls.length > 0) cell.setNote(urls.join('\n'));
        }
      }
    }
  }
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
  var yukyuRate = 0;
  if (settingSheet && settingSheet.getLastRow() >= 2) {
    var sVals = settingSheet.getRange(2, 1, settingSheet.getLastRow()-1, 4).getValues();
    for (var s = 0; s < sVals.length; s++) {
      var skey = String(sVals[s][0]||'').trim();
      if (skey) { fuelMap[skey] = sVals[s][1]; }
      if (String(sVals[s][2]||'').trim() === '有休') { yukyuRate = Number(sVals[s][3]) || 0; }
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
    if (!g) { g = { id:String(r[0]).trim(), kubun:r[1], company:r[2], tons:r[3], type:r[4], car:r[5], name:r[6], tel:r[7], date:r[8], clients:[], picks:[], drops:[], guideTime:'', pickTime:'', restStart:'', restEnd:'', dropTime:'', sales:0, tollReq:0, tollReal:0 }; }
    if (r[10]) g.picks.push(r[10]);
    if (r[11]) g.drops.push(r[11]);
    // 荷主は重複なしで全部収集
    if (r[9]) { var rc = String(r[9]); if (g.clients.indexOf(rc) === -1) g.clients.push(rc); }
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
  var keepKari='', keepKyuryo='', keepPct='', keepOther='';
  if (sumLast >= 2) {
    var colCount = sumSheet.getLastColumn();
    var sumIds   = sumSheet.getRange(2, 1, sumLast-1, Math.max(colCount, 33)).getValues();
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
        keepOther   = sumIds[k][32] || '';  // AG=その他手当
        break;
      }
    }
  }
  if (!g) { if (sumRow > 0) sumSheet.deleteRow(sumRow); return; }

  var tonsStr = String(g.tons||'').trim();
  var fuel    = fuelMap[tonsStr] || fuelMap[tonsStr.replace(/[tT]/,'')+'t'] || 3;

  var pkey   = String(g.car||'').trim()+'_'+String(g.name||'').trim();
  var pc     = payCondMap[pkey] || {kari:'', kyuryo:'', pct:''};
  var kari   = (pc.kari   !== undefined && pc.kari   !== '') ? pc.kari   : keepKari;
  var kyuryo = (pc.kyuryo !== undefined && pc.kyuryo !== '') ? pc.kyuryo : keepKyuryo;
  var pct    = (pc.pct    !== undefined && pc.pct    !== '') ? pc.pct    : keepPct;

  var spick = g.picks.join('・'), sdrop = g.drops.join('・');
  var sIsYukyu = spick.indexOf('有休') !== -1 || sdrop.indexOf('有休') !== -1;
  var sIsYasumi = !sIsYukyu && (spick.indexOf('休み') !== -1 || sdrop.indexOf('休み') !== -1);
  var rowData = [
    g.id, g.kubun, g.company, g.tons, g.type, g.car, g.name, g.tel,
    g.date, g.clients.join('・'), spick, sdrop,
    g.guideTime||'', g.pickTime, g.restStart, g.restEnd, g.dropTime,
    g.sales||'', g.tollReq||'', g.tollReal||'', '',
    keepDistance, fuel, keepGas, '', keepPay, '', keepMemo,
    kari, kyuryo, pct,
    sIsYukyu ? yukyuRate : '',
    keepOther   // AG=その他手当（手入力保持）
  ];

  if (sumRow > 0) {
    sumSheet.getRange(sumRow, 1, 1, 33).setValues([rowData]);
  } else {
    sumRow = sumSheet.getLastRow()+1;
    if (sumRow === 1) {
      var hdr = ['ID','区分','会社名','トン数','車種','車番','乗務員名','携帯番号','日付','荷主','積地','降地','誘導時刻','積完時刻','休憩開始','休憩終了','降完時刻','売上','請求(高速代)','実費(高速代)','合計(高速代)','距離','燃費','ガソリン代','燃料代','支払い','利益','備考','仮日数','給料','％','有休手当','その他手当'];
      sumSheet.getRange(1, 1, 1, 33).setValues([hdr]);
      sumSheet.setFrozenRows(1);
      sumRow = 2;
    }
    sumSheet.getRange(sumRow, 1, 1, 33).setValues([rowData]);
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
  // この行だけの支払い(Z=col26)をインライン計算（全行ループの calculatePaymentAmount を避けて高速化）
  var pctNum    = Number(pct)    || 0;
  var kyuryoNum = Number(kyuryo) || 0;
  var kariNum   = Number(kari)   || 0;
  var thisToll  = (Number(g.tollReal) || 0) - (Number(g.tollReq) || 0);
  var payCell   = sumSheet.getRange(sumRow, 26);
  sumSheet.getRange(sumRow, 29, 1, 3).setBackground(null);
  payCell.setBackground(null);
  var yukyuVal = '';
  if (pctNum > 0) {
    // 歩合制: 有休日のみ有休手当
    payCell.setValue(Math.round(((Number(g.sales) || 0) - thisToll) * pctNum / 100));
    if (sIsYukyu) yukyuVal = yukyuRate;
  } else if (kyuryoNum > 0 && kariNum > 0) {
    // 給料制: 休みはマイナス按分、有休・通常は同じ按分、有休手当なし
    var dailyPay = Math.round(kyuryoNum / kariNum);
    payCell.setValue(sIsYasumi ? -dailyPay : dailyPay);
  } else if (kyuryoNum > 0 || kariNum > 0) {
    if (!kyuryoNum) sumSheet.getRange(sumRow, 30).setBackground('#f4cccc');
    if (!kariNum)   sumSheet.getRange(sumRow, 29).setBackground('#f4cccc');
  } else {
    if (!keepPay) payCell.setBackground('#f4cccc');
  }
  // 有休手当(AF=col32): 歩合制+有休のみ
  sumSheet.getRange(sumRow, 32).setValue(yukyuVal);
  // 集計表のK列(積地)に休み/有休が含まれる場合はグレー着色
  sumSheet.getRange(sumRow, 11).setBackground(
    (spick.indexOf('休み') !== -1 || spick.indexOf('有休') !== -1) ? '#9e9e9e' : null
  );
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

  // 設定シートから有休日額を取得（C列=有休ラベル, D列=日額）
  var yukyuRate = 0;
  var settingSheet = ss.getSheetByName('設定');
  if (settingSheet && settingSheet.getLastRow() >= 2) {
    var sVals = settingSheet.getRange(2, 1, settingSheet.getLastRow()-1, 4).getValues();
    for (var s = 0; s < sVals.length; s++) {
      if (String(sVals[s][2]||'').trim() === '有休') { yukyuRate = Number(sVals[s][3]) || 0; break; }
    }
  }

  var data = sheet.getRange(2, 1, lastRow-1, 32).getValues();
  var yukyuVals = [];

  for (var i = 0; i < data.length; i++) {
    var rowNum    = i + 2;
    var sales     = Number(data[i][17]) || 0;
    var totalToll = Number(data[i][20]) || 0;
    var kari      = Number(data[i][28]) || 0;
    var kyuryo    = Number(data[i][29]) || 0;
    var pct       = Number(data[i][30]) || 0;
    var pick      = String(data[i][10] || '');
    var drop      = String(data[i][11] || '');
    var isYukyu   = pick.indexOf('有休') !== -1 || drop.indexOf('有休') !== -1;
    var isYasumi  = !isYukyu && (pick.indexOf('休み') !== -1 || drop.indexOf('休み') !== -1);
    var yCell     = sheet.getRange(rowNum, 26);

    yCell.setBackground(null);
    sheet.getRange(rowNum, 29, 1, 3).setBackground(null);

    var yukyuVal = '';
    if (pct > 0) {
      // 歩合制: 有休日のみ有休手当
      yCell.setValue(Math.round((sales - totalToll) * (pct / 100)));
      if (isYukyu) yukyuVal = yukyuRate;
    } else if (kyuryo > 0 || kari > 0) {
      if (kyuryo > 0 && kari > 0) {
        // 給料制: 休みはマイナス按分、有休・通常は同じ按分、有休手当なし
        var dailyPay = Math.round(kyuryo / kari);
        yCell.setValue(isYasumi ? -dailyPay : dailyPay);
      } else {
        if (!kyuryo) sheet.getRange(rowNum, 30).setBackground('#f4cccc');
        if (!kari)   sheet.getRange(rowNum, 29).setBackground('#f4cccc');
      }
    } else {
      if (yCell.getValue() === '') yCell.setBackground('#f4cccc');
    }
    yukyuVals.push([yukyuVal]);
  }

  // 有休手当(AF=col32)を一括書込
  if (yukyuVals.length > 0) {
    sheet.getRange(2, 32, yukyuVals.length, 1).setValues(yukyuVals);
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
  // 端末のメールアドレスを確認（未連携なら運行開始不可）
  var savedEmail = PropertiesService.getUserProperties().getProperty('linkedEmail');
  if (!savedEmail) throw new Error('紐づけされていません');

  // 同時に複数端末が運行開始した場合のID重複を防ぐためロック取得
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { throw new Error('混雑中です。少し待ってから再試行してください'); }

  try {
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    // マスタからメールアドレスが一致する乗務員情報を取得
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

    // picks/drops 配列の各要素が1行程（積地と降地がセットになっている）
    // 全行程に同一IDを付与して連番採番
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
    sheet.getRange(startRow, 11, num, 2).setNumberFormat('@'); // 積地・降地をテキスト書式に固定（数値化防止）
    sheet.getRange(startRow, 1, num, 25).setValues(rowsData);

    // U列(21)の高速計算式を各行に設定（合計高速 = 実費T - 請求S）
    var formulas = [];
    for (var i = 0; i < num; i++) {
      var r = startRow + i;
      formulas.push(['=IF(AND(T'+r+'="",S'+r+'=""),"",T'+r+'-S'+r+')']);
    }
    sheet.getRange(startRow, 21, num, 1).setFormulas(formulas);
    sheet.getRange(startRow, 9, num, 1).setNumberFormat('yyyy/MM/dd');
    applyDateTimeFormat_(sheet, startRow, num);

    // 集計表を非同期で同期（遅延実行）
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
//  端末アプリの「一覧」画面に表示するデータを月単位で返す関数
//
//  8-4-1: 紐づけメールアドレス（PropertiesService）から乗務員名を取得
//          → 未紐づけなら空のデータを返す
//  8-4-2: 運行シートを全件読み込み、指定年月・乗務員名で絞り込みID単位に集約
//          → 同一IDの複数行（複数行程）は売上/高速を合算・時刻は先頭行優先
//          → dateSort はI列（初回行程登録時刻）のまま固定（積完時刻では更新しない）
//  8-4-3: W列(23)のデータURLを3段階フォールバックで取得
//          ① セルのノート（getNotes）から → ② リッチテキスト（getRichTextValues）から
//          → ③ プレーン値（getValues）がURLなら直接使用
//  8-4-4: 集計表（payMap）から支払い・高速計を引き当てて各IDに付加
//  8-4-5: dateDisp は積完時刻があれば「yyyy/MM/dd　HH:mm」形式、なければ日付のみ
//  8-4-6: 月合計（稼働日数・売上合計・高速合計・支払合計）を計算して一緒に返す
// ================================================================
function getListData(year, month) {
  // 端末に紐付いたメールアドレスを取得（未連携なら空リストを返す）
  var savedEmail = PropertiesService.getUserProperties().getProperty('linkedEmail');
  if (!savedEmail) return {rows:[], total:{days:0,sales:0,toll:0,pay:0}};

  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  // マスタからメールアドレスが一致する乗務員名を特定
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
  var lastRow  = sheet.getLastRow();
  // W列(23)はリッチテキスト（クリック可能URLラベル）で格納されることがある
  // note→リッチテキストリンク→プレーンテキストの順に3段階フォールバックしてURLを取得
  var notes23  = lastRow >= 2 ? sheet.getRange(2, 23, lastRow-1, 1).getNotes() : [];
  var rtvs23   = lastRow >= 2 ? sheet.getRange(2, 23, lastRow-1, 1).getRichTextValues() : [];
  // 集計表からID単位の金額マップを作成（売上/高速/支払いは集計表の計算済み値を使う）
  var sumSheet = ss.getSheetByName('集計表');
  var payMap   = {};
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var sumAll = sumSheet.getRange(2, 1, sumSheet.getLastRow()-1, 32).getValues();
    for (var s = 0; s < sumAll.length; s++) {
      var sid = String(sumAll[s][0]||'').trim();
      if (sid) payMap[sid] = {
        sales:    Math.round(Number(sumAll[s][17])) || 0,
        tollReq:  Math.round(Number(sumAll[s][18])) || 0,
        tollReal: Math.round(Number(sumAll[s][19])) || 0,
        tollTotal:Math.round(Number(sumAll[s][20])) || 0,
        pay:      Math.round(Number(sumAll[s][25])) || 0,
        yukyu:    Math.round(Number(sumAll[s][31])) || 0,
        other:    Math.round(Number(sumAll[s][32])) || 0
      };
    }
  }

  // 運行シートを走査して指定年月・自分の行だけ抽出、ID単位にデータ集約
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
      var n23 = (notes23[i-1] && notes23[i-1][0]) || '';
      var du23 = n23 ? n23.split('\n').filter(function(u){return u.match(/^https?:\/\//);}) : [];
      if (!du23.length && rtvs23[i-1] && rtvs23[i-1][0]) {
        var rtv23 = rtvs23[i-1][0], rruns = rtv23.getRuns();
        for (var k = 0; k < rruns.length; k++) { var lk=rruns[k].getLinkUrl(); if(lk) du23.push(lk); }
      }
      if (!du23.length) { var pu=String(r[22]||''); if(pu.match(/^https?:\/\//)) du23=[pu]; }
      idMap[id] = {
        id:id, car:String(r[5]||'').trim(), date:ds,
        dateSort: baseDateSort,
        dateDisp:'', picks:[], drops:[],
        guideTime:'', pickTime:'', restStart:'', restEnd:'', dropTime:'',
        sales:0, tollReq:0, tollReal:0, tollTotal:0, pay:0, yukyu:0, other:0,
        notice:r[21]||'', dataUrls:du23, dataUrl:du23[0]||'',
        hasNotice:!!(r[21]||du23.length),
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
    g.hasNotice = !!(g.notice || g.dataUrls.length);
  }

  // 集計表の金額を上書き適用して結果配列を組み立て（日付降順にソート）
  var result = [];
  var totalSales=0, totalToll=0, totalPay=0, totalYukyu=0, totalOther=0, yukyuDays=0, yasumiDays=0, dateSet={};
  for (var o = 0; o < idOrder.length; o++) {
    var g  = idMap[idOrder[o]];
    // 集計表の値で金額列を上書き（スプレッドシートの計算済み数値をそのまま使う）
    var pm = payMap[g.id];
    if (pm) {
      g.sales    = pm.sales;
      g.tollReq  = pm.tollReq;
      g.tollReal = pm.tollReal;
      g.tollTotal= pm.tollTotal;
      g.pay      = pm.pay;
      g.yukyu    = pm.yukyu || 0;
      g.other    = pm.other || 0;
    }

    // dateDispはI列時刻か積完時刻を表示用に使うが、dateSortは変えない
    var dispTime = g.pickTime;
    if (!dispTime && g._rawDv) {
      var rh = g._rawDv.getHours(), rm = g._rawDv.getMinutes();
      if (rh !== 0 || rm !== 0) {
        dispTime = String(rh).padStart(2,'0')+':'+String(rm).padStart(2,'0');
      }
    }
    g.dateDisp = g.date + (dispTime ? '　'+dispTime : '');

    var gpick2 = g.picks.join('・'), gdrop2 = g.drops.join('・');
    var gIsYukyu2 = gpick2.indexOf('有休') !== -1 || gdrop2.indexOf('有休') !== -1;
    var gIsYasumi2 = !gIsYukyu2 && (gpick2.indexOf('休み') !== -1 || gdrop2.indexOf('休み') !== -1);
    if (gIsYukyu2) yukyuDays++;
    if (gIsYasumi2) yasumiDays++;
    result.push({
      id:g.id, date:g.date, dateDisp:g.dateDisp, dateSort:g.dateSort,
      car:g.car, pick:gpick2, drop:gdrop2,
      guideTime:g.guideTime, pickTime:g.pickTime, restStart:g.restStart, restEnd:g.restEnd, dropTime:g.dropTime,
      sales:g.sales, tollReq:g.tollReq, tollReal:g.tollReal, tollTotal:g.tollTotal,
      pay:g.pay, yukyu:g.yukyu, other:g.other,
      notice:g.notice, dataUrl:g.dataUrl, hasNotice:g.hasNotice,
      isComplete: !!(g.pickTime && g.restStart && g.restEnd && g.dropTime),
      isNew:      !g.guideTime && !g.pickTime && !g.restStart && !g.restEnd && !g.dropTime
    });
    totalSales += g.sales; totalToll += g.tollTotal; totalPay += g.pay;
    totalYukyu += g.yukyu || 0;
    totalOther += g.other || 0;
    dateSet[g.date] = true;
  }
  result.sort(function(a,b){ return b.dateSort - a.dateSort; });
  return { rows:result, total:{ days:Object.keys(dateSet).length, sales:totalSales, toll:totalToll, pay:totalPay, yukyu:totalYukyu, other:totalOther, yukyuDays:yukyuDays, yasumiDays:yasumiDays } };
}


// ================================================================
//  8-5: 編集用データ取得（getEditData）
//  編集モーダルに表示する1件分の詳細データを取得して返す
//  ・同一IDの複数行（複数行程）は売上/高速を合算して返す
//  ・時刻（誘導/積完/休憩/降完）は最初に見つかった値を使用（先勝ち）
//  ・集計表から合計高速代・利益を取得して付加
//  ・W列(23)のURL：getAdminDataUrl_（リッチテキスト→URLカンマ区切り）
//  ・Y列(25)のURL：getTerminalUrls_（リッチテキスト→URL配列.join(',')）
// ================================================================
function getEditData(id) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return null;

  var sumSheet = ss.getSheetByName('集計表');
  var sumData  = { tollTotal:'', profit:'', yukyu:'', other:'' };
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var sumAll = sumSheet.getRange(2, 1, sumSheet.getLastRow()-1, 33).getValues();
    for (var s = 0; s < sumAll.length; s++) {
      if (String(sumAll[s][0]||'').trim() === String(id).trim()) {
        sumData.tollTotal = sumAll[s][20] !== '' ? sumAll[s][20] : '';
        sumData.profit    = sumAll[s][26] !== '' ? sumAll[s][26] : '';
        sumData.yukyu     = sumAll[s][31] !== '' ? Math.round(Number(sumAll[s][31])) : '';
        sumData.other     = sumAll[s][32] !== '' ? Math.round(Number(sumAll[s][32])) : '';
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
    dataUrl:  getAdminDataUrl_(sheet, firstRow),
    termNotice:baseData[23]|| '',
    termData: getTerminalUrls_(sheet, firstRow).join(','),
    profit:   sumData.profit,
    yukyu:    sumData.yukyu,
    other:    sumData.other
  };
}


// ================================================================
//  8-6: 編集データ保存（saveEditData）
//  端末アプリの編集モーダルで変更された値を運行シートに書き込む
//  ・日付はDate型で書き込む（文字列だとonEditUnkou_が誤発火するため）
//  ・荷主名/積地/降地は undefined/null でなければ上書きする（空文字でも書き込み可）
//  ・時刻は「日付＋時刻」を合成したDate型で書き込み、空なら clearContent する
//  ・売上/高速は同一IDに複数行ある場合、先頭行のみ書き込み（重複防止）
//  ・書き込み後にdelaySyncSummary_を呼んで集計表を同期する
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

    // client/pick/dropはnullでなければ書き込む（空でも上書き可）
    if (obj.client !== undefined && obj.client !== null) sheet.getRange(r, 10).setValue(obj.client);
    if (obj.pick   !== undefined && obj.pick   !== null) sheet.getRange(r, 11).setValue(obj.pick);
    if (obj.drop   !== undefined && obj.drop   !== null) sheet.getRange(r, 12).setValue(obj.drop);

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
    // sales未指定（閲覧のみ）の場合は上書きしない
    if (!written) {
      if (obj.sales    !== undefined) sheet.getRange(r, 18).setValue(obj.sales    || '');
      if (obj.tollReq  !== undefined) sheet.getRange(r, 19).setValue(obj.tollReq  || '');
      if (obj.tollReal !== undefined) sheet.getRange(r, 20).setValue(obj.tollReal || '');
      written = true;
    } else {
      if (obj.tollReq  !== undefined) sheet.getRange(r, 19).setValue('');
      if (obj.tollReal !== undefined) sheet.getRange(r, 20).setValue('');
    }

    sheet.getRange(r, 24).setValue(obj.termNotice || '');
  }
  delaySyncSummary_(obj.id);
}


// ================================================================
//  8-6b: シート保護設定（setupSheetProtection）
//  集計表: 距離(V=22)・ガソリン代(X=24)・備考(AB=28)以外ロック
//  運行シート: 合計高速(U=21)列のみロック
// ================================================================
function setupSheetProtection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var sumSheet = ss.getSheetByName('集計表');
  if (sumSheet) {
    sumSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p){p.remove();});
    sumSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){p.remove();});
    var sp = sumSheet.protect().setDescription('集計表保護');
    var editableCols = [22, 24, 26, 28, 33]; // V=距離, X=ガソリン代, Z=支払い, AB=備考, AG=その他手当
    sp.setUnprotectedRanges([
      sumSheet.getRange('V2:V2000'),
      sumSheet.getRange('X2:X2000'),
      sumSheet.getRange('Z2:Z2000'),
      sumSheet.getRange('AB2:AB2000'),
      sumSheet.getRange('AG2:AG2000')
    ]);

    // ヘッダー行: 保護列=グレー, 編集可列=グリーン で視覚区別
    var lastCol = Math.max(sumSheet.getLastColumn(), 33);
    sumSheet.getRange(1, 1, 1, lastCol)
      .setBackground('#37474f').setFontColor('#90a4ae').setFontWeight('bold');
    for (var ec = 0; ec < editableCols.length; ec++) {
      sumSheet.getRange(1, editableCols[ec])
        .setBackground('#1b5e20').setFontColor('#a5d6a7').setFontWeight('bold');
    }

    // 編集可列: 列全体に緑の中太枠線を適用
    var lastRow = Math.max(sumSheet.getLastRow(), 2);
    for (var ec2 = 0; ec2 < editableCols.length; ec2++) {
      sumSheet.getRange(1, editableCols[ec2], lastRow, 1)
        .setBorder(null, true, null, true, null, null, '#4caf50', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }

  var unkouSheet = ss.getSheetByName('運行');
  if (unkouSheet) {
    unkouSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p){p.remove();});
    var unkouLastRow = Math.max(unkouSheet.getLastRow(), 2);
    // 合計(高速代)=U列(21): ヘッダーをグレーで色付け
    unkouSheet.getRange(1, 21).setBackground('#37474f').setFontColor('#90a4ae').setFontWeight('bold');
    // U列のデータ全体を薄いグレーで色付け
    unkouSheet.getRange(2, 21, unkouLastRow - 1, 1).setBackground('#eceff1');
  }

  ui.alert('保護設定完了\n■ 集計表\n  編集可: 距離(V)・ガソリン代(X)・支払い(Z)・備考(AB)・その他手当(AG)\n  緑枠＋緑ヘッダー = 編集可 / 灰色ヘッダー = 保護\n■ 運行シート: 合計(高速代)U列を薄いグレーで色付け');
}


// ================================================================
//  8-6b-1: 端末ファイルURL一覧取得（getTerminalUrls_）
//  col25のリッチテキストからリンクURLを配列で返す
// ================================================================
function getTerminalUrls_(sheet, rowNum) {
  var rtv = sheet.getRange(rowNum, 25).getRichTextValue();
  if (!rtv) return [];
  var runs = rtv.getRuns();
  var urls = [];
  for (var i = 0; i < runs.length; i++) {
    var link = runs[i].getLinkUrl();
    if (link) urls.push(link);
  }
  return urls;
}


// ================================================================
//  8-6b-2: 端末ファイルURL一覧書込（setTerminalUrls_）
//  URLをリッチテキスト（ファイル1, ファイル2…）として col25 に書込む
// ================================================================
function setTerminalUrls_(sheet, rowNum, urls) {
  var cell = sheet.getRange(rowNum, 25);
  if (!urls || urls.length === 0) { cell.setValue(''); return; }
  var text = '';
  var runs = [];
  for (var i = 0; i < urls.length; i++) {
    var label = 'ファイル' + (i + 1);
    var start = text.length;
    text += label;
    runs.push({ start: start, end: text.length, url: urls[i] });
    if (i < urls.length - 1) text += '  ';
  }
  var b = SpreadsheetApp.newRichTextValue().setText(text);
  for (var j = 0; j < runs.length; j++) {
    b.setLinkUrl(runs[j].start, runs[j].end, runs[j].url);
  }
  cell.setRichTextValue(b.build());
}


// ================================================================
//  8-6b-0: 画像URLをDriveに取込（importImageToDrive_）
//  公開画像URLをOAuthトークンで取得しDriveにコピーして返す
//  Google フォトのプライベートURLは不可（サイドバーでアップロード推奨）
// ================================================================
function importImageToDrive_(url) {
  try {
    var token = ScriptApp.getOAuthToken();
    var resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
      followRedirects: true
    });
    var code = resp.getResponseCode();
    var ct   = String(resp.getHeaders()['Content-Type'] || '');
    if (code !== 200 || !ct.match(/^image\//)) return null;
    var name = url.split('/').pop().split('?')[0] || 'image.jpg';
    var folder = getOrCreateFolder_('端末データ');
    var file = folder.createFile(resp.getBlob().setName(name));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(e) { return null; }
}


// ================================================================
//  8-6b-3: 管理側ファイルURLをリッチテキストで書込（setAdminDataRichText_）
// ================================================================
function setAdminDataRichText_(sheet, rowNum, url) {
  setAdminDataRichTextMulti_(sheet, rowNum, url ? [url] : []);
}


// ================================================================
//  8-6b-3b: 管理側ファイル複数URLをリッチテキストで書込（setAdminDataRichTextMulti_）
// ================================================================
function setAdminDataRichTextMulti_(sheet, rowNum, urls) {
  var cell = sheet.getRange(rowNum, 23);
  if (!urls || urls.length === 0) { cell.setValue(''); cell.clearNote(); return; }
  // URLをノートに保存（次のペースト時に追記するため）
  cell.setNote(urls.join('\n'));
  var text = '', runs = [];
  for (var i = 0; i < urls.length; i++) {
    var label = 'ファイル' + (i + 1);
    var start = text.length;
    text += label;
    runs.push({ start: start, end: text.length, url: urls[i] });
    if (i < urls.length - 1) text += '  ';
  }
  var b = SpreadsheetApp.newRichTextValue().setText(text);
  for (var j = 0; j < runs.length; j++) b.setLinkUrl(runs[j].start, runs[j].end, runs[j].url);
  cell.setRichTextValue(b.build());
}


// ================================================================
//  8-6b-4: 管理側ファイルURLをリッチテキストから取得（getAdminDataUrl_）
// ================================================================
function getAdminDataUrl_(sheet, rowNum) {
  var rtv = sheet.getRange(rowNum, 23).getRichTextValue();
  if (rtv) {
    var runs = rtv.getRuns(), urls = [];
    for (var i = 0; i < runs.length; i++) {
      var link = runs[i].getLinkUrl();
      if (link) urls.push(link);
    }
    if (urls.length > 0) return urls.join(',');
  }
  return String(sheet.getRange(rowNum, 23).getValue() || '');
}


// ================================================================
//  8-6c: 端末ファイル追加（appendTerminalFile）
//  ファイルをDriveに保存しcol25のリッチテキストURLに追記する
// ================================================================
function appendTerminalFile(id, fileName, base64Data, mimeType) {
  var folder  = getOrCreateFolder_('端末データ');
  var decoded = Utilities.base64Decode(base64Data);
  var blob    = Utilities.newBlob(decoded, mimeType, fileName);
  var file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url   = file.getUrl();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() === String(id).trim()) {
      var urls = getTerminalUrls_(sheet, i + 1);
      urls.push(url);
      setTerminalUrls_(sheet, i + 1, urls);
      break;
    }
  }
  return { ok: true, url: url };
}


// ================================================================
//  8-6c-2: 管理側ファイル追加・削除・差替（ID指定）
// ================================================================
function appendAdminFileById(id, fileName, base64Data, mimeType) {
  var folder  = getOrCreateFolder_('運行データ');
  var decoded = Utilities.base64Decode(base64Data);
  var blob    = Utilities.newBlob(decoded, mimeType, fileName);
  var file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url   = file.getUrl();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() !== String(id).trim()) continue;
    var existing = getAdminDataUrl_(sheet, i+1).split(',').filter(function(u){return u.match(/^https?:\/\//);});
    existing.push(url);
    var deduped = existing.filter(function(u,j,arr){return arr.indexOf(u)===j;});
    setAdminDataRichTextMulti_(sheet, i+1, deduped);
    break;
  }
  return { ok: true, url: url };
}

function deleteAdminFileById(id, urlToDelete) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() !== String(id).trim()) continue;
    var urls = getAdminDataUrl_(sheet, i+1).split(',').filter(function(u){
      return u.match(/^https?:\/\//) && u.trim() !== urlToDelete.trim();
    });
    setAdminDataRichTextMulti_(sheet, i+1, urls);
    break;
  }
}

function replaceAdminFileById(id, oldUrl, fileName, base64Data, mimeType) {
  var folder  = getOrCreateFolder_('運行データ');
  var decoded = Utilities.base64Decode(base64Data);
  var blob    = Utilities.newBlob(decoded, mimeType, fileName);
  var file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var newUrl  = file.getUrl();
  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() !== String(id).trim()) continue;
    var urls = getAdminDataUrl_(sheet, i+1).split(',').filter(function(u){return u.match(/^https?:\/\//);});
    var idx = urls.indexOf(oldUrl);
    if (idx !== -1) urls[idx] = newUrl; else urls.push(newUrl);
    setAdminDataRichTextMulti_(sheet, i+1, urls);
    break;
  }
  return { ok: true, url: newUrl };
}


// ================================================================
//  8-6d: 端末ファイル削除（deleteTerminalFile）
//  col25のリッチテキストURLから指定URLを除去する
// ================================================================
function deleteTerminalFile(id, urlToDelete) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() !== String(id).trim()) continue;
    var urls = getTerminalUrls_(sheet, i + 1)
      .filter(function(u) { return u.trim() !== urlToDelete.trim(); });
    setTerminalUrls_(sheet, i + 1, urls);
    break;
  }
}


// ================================================================
//  8-6e: 端末ファイル差し替え（replaceTerminalFile）
//  col25のリッチテキストURLの指定URLを新URLに置き換える
// ================================================================
function replaceTerminalFile(id, oldUrl, fileName, base64Data, mimeType) {
  var folder  = getOrCreateFolder_('端末データ');
  var decoded = Utilities.base64Decode(base64Data);
  var blob    = Utilities.newBlob(decoded, mimeType, fileName);
  var file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var newUrl = file.getUrl();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() !== String(id).trim()) continue;
    var urls = getTerminalUrls_(sheet, i + 1).map(function(u) {
      return u.trim() === oldUrl.trim() ? newUrl : u;
    });
    setTerminalUrls_(sheet, i + 1, urls);
    break;
  }
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
        setAdminDataRichText_(sheet, i + 1, url); break;
      }
    }
  }
  return { ok: true, url: url, fileName: fileName };
}


// ================================================================
//  9-2b: シートボタン用ファイルアップロードダイアログ（openFileUploadDialog）
//  運行シートで行を選択した状態でボタンを押すとダイアログが開く
// ================================================================
function openFileUploadDialog() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var row   = sheet.getActiveCell().getRow();
  if (sheet.getName() !== '運行' || row <= 1) {
    SpreadsheetApp.getUi().alert('運行シートのデータ行を選択してからボタンを押してください');
    return;
  }
  var id = String(sheet.getRange(row, 1).getValue()).trim();
  if (!id) { SpreadsheetApp.getUi().alert('IDが空の行です'); return; }

  var html = '<html><body style="font-family:sans-serif;padding:20px;background:#1e1e1e;color:#e0e0e0;">' +
    '<p style="margin-bottom:12px;">行 <b>' + row + '</b>（ID: ' + id + '）にアップロード</p>' +
    '<input type="file" id="f" multiple accept="image/*,.pdf,.doc,.docx,.xls,.xlsx"' +
    ' style="color:#e0e0e0;margin-bottom:12px;display:block;"><br>' +
    '<button onclick="upload()" style="padding:10px 24px;background:#1565c0;color:white;border:none;border-radius:8px;font-size:15px;cursor:pointer;">アップロード</button>' +
    '<p id="msg" style="margin-top:12px;color:#aaa;"></p>' +
    '<script>' +
    'function upload(){' +
    '  var files=Array.from(document.getElementById("f").files);' +
    '  if(!files.length){alert("ファイルを選択してください");return;}' +
    '  document.getElementById("msg").innerText="アップロード中...";' +
    '  var done=0;' +
    '  files.forEach(function(file){' +
    '    if(file.size>10*1024*1024){done++;check();return;}' +
    '    var r=new FileReader();' +
    '    r.onload=function(){' +
    '      var b64=r.result.split(",")[1];' +
    '      google.script.run' +
    '        .withSuccessHandler(function(){done++;check();})' +
    '        .withFailureHandler(function(e){document.getElementById("msg").innerText="エラー："+e.message;done++;check();})' +
    '        .uploadFileToRow(' + row + ',file.name,b64,file.type);' +
    '    };' +
    '    r.readAsDataURL(file);' +
    '  });' +
    '  function check(){if(done===files.length){document.getElementById("msg").innerText="完了！";setTimeout(google.script.host.close,800);}}' +
    '}' +
    '<\/script>' +
    '</body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(420).setHeight(220),
    'データアップロード'
  );
}


// ================================================================
//  9-2c: シートボタン用ファイルアップロード処理（uploadFileToRow）
// ================================================================
function uploadFileToRow(rowNum, fileName, base64Data, mimeType) {
  var folder  = getOrCreateFolder_('運行データ');
  var decoded = Utilities.base64Decode(base64Data);
  var blob    = Utilities.newBlob(decoded, mimeType, fileName);
  var file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url   = file.getUrl();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  if (sheet) {
    var existing = getAdminDataUrl_(sheet, rowNum).split(',').filter(function(u){return u.match(/^https?:\/\//);});
    existing.push(url);
    // 重複URLを除去（同じファイルが2度登録されないように）
    var deduped = existing.filter(function(u, i, arr) { return arr.indexOf(u) === i; });
    setAdminDataRichTextMulti_(sheet, rowNum, deduped);
  }
  return { ok: true };
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
  return appendTerminalFile(id, fileName, base64Data, mimeType);
}


// ================================================================
//  10-1: ホーム用連絡事項取得（getMyNotices）
//  端末アプリのホーム画面に表示する未読の連絡事項一覧を返す（最大20件）
//
//  対象行の条件：
//    ・V列(22)=管理側連絡事項 または W列(23)=データURLがある行
//    ・readNoticesリスト（既読済みID）に含まれていない行
//    ・乗務員名が紐づけメールに一致する行
//
//  W列URLの取得優先順：
//    ① getNotes（セルのノート）→ ② getRichTextValues（リッチテキストのリンク）→ ③ getValues（プレーン値）
//
//  返却値：{ id, date, notice, dataUrls[], dataUrl } の配列（最新順・最大20件）
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
  var lr  = sheet.getLastRow();
  var notes23 = lr >= 2 ? sheet.getRange(2, 23, lr-1, 1).getNotes() : [];
  var rtvs23m = lr >= 2 ? sheet.getRange(2, 23, lr-1, 1).getRichTextValues() : [];
  var out = [], seen = {};
  for (var i = 1; i < all.length; i++) {
    var r = all[i];
    if (String(r[6]).trim() !== myName) continue;
    var id = String(r[0]||'').trim();
    if (!id || seen[id]) continue;
    seen[id] = true;
    var notice = String(r[21]||'');
    var n23 = (notes23[i-1] && notes23[i-1][0]) || '';
    var dataUrls = n23 ? n23.split('\n').filter(function(u){return u.match(/^https?:\/\//);}) : [];
    if (!dataUrls.length && rtvs23m[i-1] && rtvs23m[i-1][0]) {
      var rtv23m = rtvs23m[i-1][0], rrunsm = rtv23m.getRuns();
      for (var k = 0; k < rrunsm.length; k++) { var lkm=rrunsm[k].getLinkUrl(); if(lkm) dataUrls.push(lkm); }
    }
    if (!dataUrls.length) { var pu=String(r[22]||''); if(pu.match(/^https?:\/\//)) dataUrls=[pu]; }
    if (!notice && !dataUrls.length) continue;
    if (readList.indexOf(id) !== -1) continue;
    out.push({ id:id, date: r[8] ? Utilities.formatDate(new Date(r[8]),'Asia/Tokyo','yyyy/MM/dd HH:mm') : '', notice:notice, dataUrls:dataUrls, dataUrl:dataUrls[0]||'' });
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
    var notice = String(all[i][noticeCol]||'');
    var cell = sheet.getRange(i+1, dataCol+1);
    var note = cell.getNote() || '';
    var dataUrls = note ? note.split('\n').filter(function(u){return u.match(/^https?:\/\//);}) : [];
    if (!dataUrls.length) {
      var rtv = cell.getRichTextValue();
      if (rtv) { var runs=rtv.getRuns(); for(var k=0;k<runs.length;k++){var lk=runs[k].getLinkUrl();if(lk)dataUrls.push(lk);} }
    }
    if (!dataUrls.length) { var pu=String(all[i][dataCol]||''); if(pu.match(/^https?:\/\//)) dataUrls=[pu]; }
    return { notice:notice, dataUrls:dataUrls, dataUrl:dataUrls[0]||'' };
  }
  return { notice:'', dataUrls:[], dataUrl:'' };
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