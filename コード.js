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
//   4-7  : generateNextMonth()
//            次月の1日〜末日 × 運行中車両 のプレースホルダーIDを運行シートに一括生成する
//            ・積地が空の行には配車漏れ警告色（#fff9c4）を付ける
//            ・生成後に3ヶ月以上のデータがあれば最古月を自動アーカイブ（4-8a呼び出し）
//   4-8  : archiveOldMonth()
//            前月分を別スプレッドシートに値のみで保存して運行シートから削除する（手動実行用）
//   4-8a : archiveOldestMonthIfNeeded_()
//            運行シートに3ヶ月以上のデータがある場合のみ最古月を自動アーカイブする
//   4-8b : archiveMonthData_(ss, year, month, companyName)
//            指定月のデータをアーカイブ用スプレッドシートにコピーし元行を削除する
//            ・「運行管理_アーカイブ/会社名/」フォルダへ移動・メールアドレスで共有設定
//   4-8c : getMonthRows_(sheet, year, month, numCols)
//            指定シートのJ列(10)で絞り込んだ指定年月の行データを返す内部補助関数
//   4-8d : getCompanyName_(ss)
//            自車専属マスタのC列(会社名)からこのスプレッドシートの会社名を返す内部補助関数
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
// ── グループ11：会社セットアップ・配布メール（管理者用）────────────────
//   11-1 : onEditCompanyRegister_(sheet, range)
//            会社登録シートのA+B列 → フォルダ作成・通知メール
//            F+G列（SS URL＋App URL）が揃ったら → 配布メール自動送信
//   11-2 : setupOneCompany_(companyName, adminEmail)
//            1社分のアーカイブフォルダを作成し管理Gmailに編集権限を付与して通知メールを送る
//   11-3 : setupCompanies()
//            会社登録シートを全行処理するか新規作成する（スクリプトエディタ or メニューから実行）
//   11-4 : createUsageSheet()
//            スプレッドシート内「使い方」シートを自動作成（管理者・ドライバー向け2部構成）
//   11-5 : sendDistributionMail_(companyName, adminEmail, ssUrl, appUrl, row, sheet)
//            管理者向け（SS URL＋App URL）＋乗務員向け（App URL＋紐づけ手順）メールを送信
//            乗務員メールは自車専属マスタのJ列から全員分個別送信する
//   11-6 : triggerDistributionMail()
//            会社登録シートのF+G列が揃いH列が未送信の全行に配布メールを一括送信する
//            メニュー「📧 配布メール送信」から手動実行
//
// ================================================================
//
// ================================================================
// ■ 大中小 分類体系（アクセス経路・機能カテゴリ一覧）
// ================================================================
//
//  ▼ 大分類 ── 呼び出し元・アクセス経路で3種に大別
//
//   大A  GAS ↔ HTML 接続ポイント
//        index.html の google.script.run から直接呼ばれる公開関数
//        → グループ5〜10（端末アプリ系・管理アプリ系の全API）
//
//   大B  GAS 内部処理
//        トリガー・内部ヘルパー・他GAS関数からのみ呼ばれる非公開処理
//        → グループ1（共通補助関数）・グループ3（onEdit系）・グループ4（集計表系）
//           + 末尾アンダースコア付きの内部ヘルパー関数
//
//   大C  スプレッドシート 起動・メニュー
//        スプレッドシートを開いたとき・メニュー項目として実行される関数
//        → グループ2（onOpen / doGet / showSidebar 等）
//           + メンテナンス用メニュー項目（setupSheetProtection / insertKanbanColumn）
//
// ──────────────────────────────────────────────────────────────────
//
//  ▼ 中分類 ── 機能グループ（グループ番号と1:1対応）
//
//   中1  補助関数群                 1-1〜1-7   getNextIdNum_, applyMoneyFormat_ など
//   中2  起動・メニュー             2-1〜2-4   onOpen, doGet, showSidebar など
//   中3  スプレッドシート自動処理   3-1〜3-4   onEdit, onEditUnkou_ など
//   中4  集計表・シート操作         4-1〜4-6   generateSummary, syncSummaryForId_ など
//   中5  アプリ初期化・紐づけ       5-1〜5-3   getInitialData, linkAddress など
//   中6  端末 運行進捗管理          6-1〜6-3   saveRunState, loadRunState など
//   中7  端末 運行操作              7-1〜7-6   getTodayRoutes, createParentRows など
//   中8  端末 一覧・編集・ファイル  8-1〜8-7   getListData, saveEditData など
//   中9  端末 連絡・ファイル        9-1〜9-4   saveNotice, uploadFileToRow など
//   中10 端末 既読管理              10-1〜10-4 getMyNotices, markAsRead など
//   中11 会社セットアップ・配布     11-1〜11-6 onEditCompanyRegister_, setupOneCompany_, setupCompanies,
//                                             createUsageSheet, sendDistributionMail_, triggerDistributionMail
//
// ──────────────────────────────────────────────────────────────────
//
//  ▼ 小分類 ── 関数個別番号（既存グループ番号と同一、変更なし）
//    例：小1-1 = getNextIdNum_,  小8-4 = getListData
//
//  ▼ 各関数コメントブロックの読み方（例）
//    // ================================================================
//    //  8-4: 運行一覧データ取得（getListData）  【大A / 中8 / 小8-4】
//    // ================================================================
//    → 「HTMLから呼ばれる（大A） ＞ グループ8（中8） ＞ 番号8-4（小8-4）」
//
// ================================================================
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
//  1-1: ID番号取得補助関数  【大B / 中1 / 小1-1】
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
//  1-2: Googleドライブのフォルダ取得or作成（補助）  【大B / 中1 / 小1-2】
//  指定名のフォルダが存在すれば返し、なければ作成して返す
// ================================================================
function getOrCreateFolder_(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}


// ================================================================
//  1-3: 集計表遅延同期ラッパー  【大B / 中1 / 小1-3】
//  syncSummaryForId_をtry-catchで安全に呼び出す
// ================================================================
function delaySyncSummary_(id, ss) { try { syncSummaryForId_(id, ss); } catch(e) {} }


// ================================================================
//  1-4: 集計表の孤立ID削除  【大B / 中1 / 小1-4】
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
//  1-5: 金額列へのコンマ書式適用（applyMoneyFormat_）  【大B / 中1 / 小1-5】
//  指定シート・行範囲の金額列に #,##0 フォーマットをセットする
// ================================================================
function applyMoneyFormat_(sheet, startRow, numRows, sheetType) {
  if (numRows <= 0) return;
  var fmt = '#,##0;[RED]#,##0';
  var cols = (sheetType === 'unkou')
    ? [19, 20, 21, 22]
    : [19, 20, 21, 22, 23, 25, 26, 27, 28, 29, 34, 35]; // 23=距離(¥なし#,##0を上書き)
  for (var i = 0; i < cols.length; i++) {
    sheet.getRange(startRow, cols[i], numRows, 1).setNumberFormat(fmt);
  }
}


// ================================================================
//  1-6: 時刻列へのM/d HH:mm書式適用（applyDateTimeFormat_）  【大B / 中1 / 小1-6】
//  誘導・積完・休憩開始・休憩終了・降完（M〜Q列=13〜17）に書式をセットする
// ================================================================
function applyDateTimeFormat_(sheet, startRow, numRows) {
  if (numRows <= 0) return;
  var fmt = 'M/d HH:mm';
  var cols = [14, 15, 16, 17, 18];
  var lastCol = sheet.getLastColumn();
  if (sheet.getName() === '集計表' && lastCol >= 36) {
    cols = cols.concat([36, 37]);
  } else if (sheet.getName() === '運行' && lastCol > 0) {
    // 点呼前後列は末尾に動的追加されるため列番号をヘッダーから検索
    var hdrs = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var bi = hdrs.indexOf('点呼前完了'), ai = hdrs.indexOf('点呼後完了');
    if (bi >= 0) cols.push(bi + 1);
    if (ai >= 0) cols.push(ai + 1);
  }
  for (var i = 0; i < cols.length; i++) {
    if (cols[i] <= lastCol) sheet.getRange(startRow, cols[i], numRows, 1).setNumberFormat(fmt);
  }
}


// ================================================================
//  1-7: 積地（L列=12）背景色設定（applyHolidayRowColors_）  【大B / 中1 / 小1-7】
//  運行シート・集計表のL列(12=積地)に以下3パターンで背景色を付ける
//  ・「休み」「有休」含む → グレー (#9e9e9e)
//  ・空 かつ 日付(J列)が今日以降 → 黄色 (#fff9c4)  ← 配車漏れ警告
//  ・その他 → なし (null)
//  onOpen・generateSummary・generateNextMonth・onEdit から呼び出す
// ================================================================
function applyHolidayRowColors_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date(); today.setHours(0, 0, 0, 0);

  // 運行シート: 有休=薄グレー・休み=暗グレーで行全体着色、配車漏れ=L列のみ黄色
  // 通常行の既存背景色は保護（有休/休み色のみ上書き）
  var sheet = ss.getSheetByName('運行');
  if (sheet && sheet.getLastRow() >= 2) {
    var lr      = sheet.getLastRow();
    var lastCol = Math.max(sheet.getLastColumn(), 12);
    var vals    = sheet.getRange(2, 1, lr - 1, 12).getValues(); // A〜L列を読む（A=ID, L=積地）
    var curBgs  = sheet.getRange(2, 1, lr - 1, lastCol).getBackgrounds();
    var bgs2D   = vals.map(function(r, idx) {
      var idV   = String(r[0] || '').trim();
      var pickV = String(r[11] || ''); // L列=積地
      if (pickV.indexOf('有休') !== -1) return new Array(lastCol).fill('#e0e0e0');
      if (pickV.indexOf('休み') !== -1) return new Array(lastCol).fill('#9e9e9e');
      // 通常行: 既存背景を保護し、有休/休み色だけクリア
      var rowArr = curBgs[idx].slice();
      for (var ci = 0; ci < rowArr.length; ci++) {
        if (rowArr[ci] === '#e0e0e0' || rowArr[ci] === '#9e9e9e') rowArr[ci] = null;
      }
      if (pickV === '' && idV !== '') {
        rowArr[11] = '#fff9c4'; // IDがあって積地空→常に黄色
      } else if (pickV !== '' && rowArr[11] === '#fff9c4') {
        rowArr[11] = null; // 積地入力あり→黄色解除
      }
      return rowArr;
    });
    sheet.getRange(2, 1, lr - 1, lastCol).setBackgrounds(bgs2D);
  }

  // 集計表: 同様に行全体着色（警告グレー・保護色を維持）
  var sumSheet = ss.getSheetByName('集計表');
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var slr      = sumSheet.getLastRow();
    var sLastCol = Math.max(sumSheet.getLastColumn(), 12);
    var svals    = sumSheet.getRange(2, 1, slr - 1, 12).getValues(); // A〜L列
    var sCurBgs  = sumSheet.getRange(2, 1, slr - 1, sLastCol).getBackgrounds();
    var sbgs2D   = svals.map(function(r, idx) {
      var sIdV  = String(r[0] || '').trim();
      var pickV = String(r[11] || ''); // L列=積地
      if (pickV.indexOf('有休') !== -1) return new Array(sLastCol).fill('#e0e0e0');
      if (pickV.indexOf('休み') !== -1) return new Array(sLastCol).fill('#9e9e9e');
      // 通常行: 既存背景を保護し、有休/休み色だけクリア
      var rowArr = sCurBgs[idx].slice();
      for (var ci = 0; ci < rowArr.length; ci++) {
        if (rowArr[ci] === '#e0e0e0' || rowArr[ci] === '#9e9e9e') rowArr[ci] = null;
      }
      if (pickV === '' && sIdV !== '') {
        rowArr[11] = '#fff9c4'; // IDがあって積地空→黄色
      } else if (pickV !== '' && rowArr[11] === '#fff9c4') {
        rowArr[11] = null;
      }
      return rowArr;
    });
    sumSheet.getRange(2, 1, slr - 1, sLastCol).setBackgrounds(sbgs2D);
  }

  markIdCollisions_();
}


// ================================================================
//  1-7b: IDが重複している行のA列を赤色でマーク（markIdCollisions_）  【大B / 中1 / 小1-7b】
//  運行シートで「同じID、異なる車番 or 異なる日付」の行を検出して赤く着色
//  正常時（同ID・同車番・同日付）はマークしない
// ================================================================
function markIdCollisions_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  if (!sheet || sheet.getLastRow() < 2) return;
  var lr = sheet.getLastRow();
  var numRows = lr - 1;
  var aVals  = sheet.getRange(2, 1,  numRows, 1).getValues();  // ID
  var fVals  = sheet.getRange(2, 6,  numRows, 1).getValues();  // 車番
  var jVals  = sheet.getRange(2, 10, numRows, 1).getValues();  // 日付
  var curBgs = sheet.getRange(2, 1,  numRows, 1).getBackgrounds();

  // IDごとに車番・日付を収集
  var idMap = {};
  for (var i = 0; i < numRows; i++) {
    var id = String(aVals[i][0] || '').trim();
    if (!id) continue;
    if (!idMap[id]) idMap[id] = { cars: {}, dates: {} };
    var car  = String(fVals[i][0] || '').trim();
    var date = jVals[i][0];
    var dk   = (date instanceof Date) ? Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd') : String(date);
    idMap[id].cars[car]  = true;
    idMap[id].dates[dk]  = true;
  }

  // 車番または日付が2種類以上 → 衝突ID
  var collisionIds = {};
  for (var cid in idMap) {
    var e = idMap[cid];
    if (Object.keys(e.cars).length > 1 || Object.keys(e.dates).length > 1) {
      collisionIds[cid] = true;
    }
  }

  // A列の色のみ更新（衝突→赤、前回赤→解除、それ以外→変更なし）
  var newColors = [];
  var hasChange = false;
  for (var i = 0; i < numRows; i++) {
    var id  = String(aVals[i][0] || '').trim();
    var cur = curBgs[i][0];
    if (id && collisionIds[id]) {
      newColors.push(['#ff1744']);
      if (cur !== '#ff1744') hasChange = true;
    } else if (cur === '#ff1744') {
      newColors.push([null]);
      hasChange = true;
    } else {
      newColors.push([cur]);
    }
  }
  if (hasChange) sheet.getRange(2, 1, numRows, 1).setBackgrounds(newColors);
}


// ================================================================
//  1-8: 運行シートを日付順に並び替え（sortUnkouByDate_）  【大B / 中1 / 小1-8】
//  運行シートのデータ行を J列(日付)昇順 → G列(乗務員名)昇順 でソートする
//  ソート後に V列(22)の数式を行番号に合わせて再セットする
//  generateNextMonth / 手動メニューから呼び出す
// ================================================================
function sortUnkouByDate_(companySsId) {
  var ss    = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet || sheet.getLastRow() < 3) return; // データが1行以下はスキップ

  var lastRow = sheet.getLastRow();
  var numRows = lastRow - 1; // ヘッダー除く行数
  var totalCols = Math.max(sheet.getLastColumn(), 28); // 点呼列(27,28)等も含めて全列処理

  // 全データを値・背景色として一括取得（背景もソート対象にして色がデータと一緒に動く）
  var data   = sheet.getRange(2, 1, numRows, totalCols).getValues();
  var curBgs = sheet.getRange(2, 1, numRows, totalCols).getBackgrounds();

  // リッチテキストが入る列（col24=管理データ, col26=データ端末）を別途取得してリンクを保持
  var rtv24 = sheet.getRange(2, 24, numRows, 1).getRichTextValues();
  var rtv26 = sheet.getRange(2, 26, numRows, 1).getRichTextValues();

  // ソート前インデックス付きで保持（背景色も一緒に持つ）
  var indexed = data.map(function(row, idx) {
    return { row: row, rtv24: rtv24[idx][0], rtv26: rtv26[idx][0], bg: curBgs[idx] };
  });

  // J列(index[9]=日付)昇順 → G列(index[6]=乗務員名)昇順 でソート
  indexed.sort(function(a, b) {
    var da = (a.row[9] instanceof Date) ? a.row[9].getTime() : 0;
    var db = (b.row[9] instanceof Date) ? b.row[9].getTime() : 0;
    if (da !== db) return da - db;
    return String(a.row[6] || '').localeCompare(String(b.row[6] || ''));
  });

  // V列(22, 0-indexed:21)を空にして値を書き戻す
  var writeData = indexed.map(function(item) {
    var row = item.row.slice();
    row[21] = ''; // 合計高速は数式で再セットするため空に
    return row;
  });
  sheet.getRange(2, 1, numRows, totalCols).setValues(writeData);

  // 背景色もソート後の順序で復元（色がデータと一緒に移動する）
  var sortedBgs = indexed.map(function(item) { return item.bg; });
  sheet.getRange(2, 1, numRows, totalCols).setBackgrounds(sortedBgs);

  // リッチテキストをソート後の順序で復元（リンクを維持）
  var newRtv24 = indexed.map(function(item) { return [item.rtv24]; });
  var newRtv26 = indexed.map(function(item) { return [item.rtv26]; });
  sheet.getRange(2, 24, numRows, 1).setRichTextValues(newRtv24);
  sheet.getRange(2, 26, numRows, 1).setRichTextValues(newRtv26);

  // V列(22)の数式を正しい行番号で再セット
  var formulas = [];
  for (var i = 0; i < numRows; i++) {
    var rn = i + 2;
    formulas.push(['=IF(AND(U' + rn + '="",T' + rn + '=""),"",U' + rn + '-T' + rn + ')']);
  }
  sheet.getRange(2, 22, numRows, 1).setFormulas(formulas);

  // 書式を再適用
  sheet.getRange(2, 10, numRows, 1).setNumberFormat('yyyy/MM/dd');
  sheet.getRange(2, 12, numRows, 2).setNumberFormat('@');
  applyMoneyFormat_(sheet, 2, numRows, 'unkou');
  applyDateTimeFormat_(sheet, 2, numRows);
}


// ================================================================
//  1-9: 集計表シートを日付順に並び替え（sortSummaryByDate_）  【大B / 中1 / 小1-9】
//  集計表シートのデータ行を J列(日付)昇順 → G列(乗務員名)昇順 でソートする
//  ソート後に数式列（V/Z/AB）を行番号に合わせて再セットする
// ================================================================
function sortSummaryByDate_(companySsId) {
  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('集計表');
  if (!sheet || sheet.getLastRow() < 3) return;
  var lastRow = sheet.getLastRow();
  var numRows = lastRow - 1;
  var colCount = sheet.getLastColumn();
  var cols = Math.max(colCount, 37);
  var data = sheet.getRange(2, 1, numRows, cols).getValues();
  var bgs  = sheet.getRange(2, 1, numRows, cols).getBackgrounds();
  var indexed = data.map(function(row, idx) { return { row: row, bg: bgs[idx] }; });
  indexed.sort(function(a, b) {
    var da = (a.row[9] instanceof Date) ? a.row[9].getTime() : 0;
    var db = (b.row[9] instanceof Date) ? b.row[9].getTime() : 0;
    if (da !== db) return da - db;
    return String(a.row[6] || '').localeCompare(String(b.row[6] || ''));
  });
  // 値を計算して書き戻す（数式廃止でフィルター後もズレない）
  var writeData = indexed.map(function(item) {
    var row = item.row.slice();
    var tollReq = Number(row[19])||0, tollReal = Number(row[20])||0;
    row[21] = (tollReq === 0 && tollReal === 0) ? '' : tollReal - tollReq;  // V
    row[25] = row[22] ? Math.round(Number(row[22])/(Number(row[23])||3)*(Number(row[24])||0)) : '';  // Z
    var vNS = typeof row[21]==='number'?row[21]:0, zNS = typeof row[25]==='number'?row[25]:0;
    var payNS = Number(row[26])||0, expNS = Number(row[27])||0, salesNS = Number(row[18])||0;
    row[28] = (!salesNS&&!vNS&&!zNS&&!payNS&&!expNS) ? '' : salesNS-(vNS+zNS+payNS+expNS);  // AC
    return row;
  });
  sheet.getRange(2, 1, numRows, cols).setValues(writeData);
  // 背景色もソート後の順序で書き戻す（色がデータと一緒に移動する）
  var sortedBgs = indexed.map(function(item) { return item.bg; });
  sheet.getRange(2, 1, numRows, cols).setBackgrounds(sortedBgs);
  sheet.getRange(2, 10, numRows, 1).setNumberFormat('yyyy/MM/dd');
  applyMoneyFormat_(sheet, 2, numRows, 'summary');
  applyDateTimeFormat_(sheet, 2, numRows);
}


// ================================================================
//  1-10: 運行＋集計表を両方日付順並び替え（sortBothSheetsByDate）  【大B / 中1 / 小1-10】
//  メニューボタンおよびonEditの日付変更時に呼び出す
// ================================================================
function sortBothSheetsByDate() {
  sortUnkouByDate_();
  sortSummaryByDate_();
  SpreadsheetApp.getActiveSpreadsheet().toast('日付順に並び替えました', '🔃 完了', 3);
}


// ================================================================
//  2-1: メニュー設定（onOpen）  【大C / 中2 / 小2-1】
//  ①修正用SS専用メニュー。②客用SS・③各客SSはスタブの onOpen で表示。
// ================================================================
function onOpen() {
  var menu = SpreadsheetApp.getUi().createMenu('メニュー');
  menu
    .addItem('ホーム画面を表示', 'showSidebar')
    .addSeparator()
    .addItem('📅 今月分生成（途中契約）', 'generateCurrentMonth')
    .addItem('📅 翌月分生成（前月アーカイブ）', 'generateNextMonth')
    .addItem('📦 前月分アーカイブ', 'archiveOldMonth')
    .addSeparator()
    .addItem('🔄 メニュー再生成（F5でOK）', 'reloadMenu')
    .addItem('集計表再生成', 'generateSummary')
    .addItem('シート再生成', 'expandAndRefreshSheets')
    .addItem('💴 経費自動入力', 'autoFillExpense')
    .addItem('🔃 日付順並び替え', 'sortBothSheetsByDate')
    .addItem('🆔 ID・車番一括補完', 'fillMissingIdsAndCars')
    .addSeparator()
    .addItem('📷 写真・ファイル取込', 'showUploadSidebar')
    .addItem('📖 使い方シート作成', 'createUsageSheet')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📥 データ読み込み（CSV）')
      .addItem('運行シート', 'showCsvImportDialogUnkou')
      .addItem('自車専属マスタ', 'showCsvImportDialogMaster')
      .addItem('マスタ（取引先）', 'showCsvImportDialogCust')
      .addSeparator()
      .addItem('🗑 空インポート行を削除', 'deleteBlankImportRows'))
    .addSeparator()
    .addItem('シート保護設定', 'setupSheetProtection')
    .addSeparator()
    .addItem('🔧 初期設定', 'installTriggers')
    .addSeparator()
    .addItem('🏢 会社セットアップ実行', 'setupCompanies')
    .addItem('📤 会社SS作成＆メール送信', 'sendCompanySetupEmails')
    .addItem('📧 配布メール送信', 'triggerDistributionMail')
    .addItem('📝 申し込みフォーム作成', 'createSignupForm')
    .addSeparator()
    .addItem('📄 請求書生成', 'showInvoiceDialog')
    .addItem('📄 支払確認書生成', 'showPaymentDialog')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📋 帳票・送信メニュー')
      .addItem('① 発注書・指示書を作成（協力会社・乗務員用）', 'showHatchuDocDialog')
      .addItem('② 車番連絡を作成（荷主用）', 'showShabanDocDialog'))
    .addSeparator()
    .addItem('🔗 チェックした行を配車確定', 'matchAndConfirmDispatch')
    .addSeparator()
    .addItem('📤 テスト客SS（②）に反映', 'syncToTemplateSS')
    .addItem('📤 全客SS（③）に反映',     'syncToAllClientSS');
  menu.addToUi();

  try {
    PropertiesService.getScriptProperties().setProperty(
      'masterSsId', SpreadsheetApp.getActiveSpreadsheet().getId()
    );
  } catch(ex) {}
  convertLegacyAdminDataUrls_();
  applyHolidayRowColors_();
}


// ================================================================
//  2-1b: メニュー再生成（reloadMenu）  【大C / 中2】
//  スタブ更新後にメニューを即時反映させる。onOpenを再実行するだけ。
// ================================================================
function reloadMenu() {
  onOpen();
  SpreadsheetApp.getActiveSpreadsheet().toast('メニューを再生成しました', '🔄', 3);
}


// ================================================================
//  2-2: Webアプリ起動（doGet）  【大C / 中2 / 小2-2】
//  URLアクセス時にWebアプリとして表示する。
//  ?ssId=XXXX パラメータを受け取り、HTMLテンプレートに渡す。
//  初回アクセス時に自分自身のWebアプリURLを自動取得してScript Propertiesに保存する。
//  これにより会社SS作成時のアプリURLが自動で設定される（手動入力不要）。
// ================================================================
function doGet(e) {
  var page  = (e && e.parameter && e.parameter.page)  ? e.parameter.page  : '';
  var ssId  = (e && e.parameter && e.parameter.ssId)  ? e.parameter.ssId  : '';

  // 本番デプロイのURLのみ保存（テンプレートSS等からのアクセスで上書きされないよう限定）
  try {
    var PROD_DEPLOY_ID = 'AKfycbw7rzkd_SuE1I6BNzEjED4Mxl6cnM4wbswIiRiNoPf5zcSS2JcP6YLkfRV21fLc0opU';
    var svcUrl = ScriptApp.getService().getUrl();
    if (svcUrl && svcUrl.indexOf(PROD_DEPLOY_ID) !== -1) {
      PropertiesService.getScriptProperties().setProperty('webAppUrl', svcUrl);
    }
  } catch(ex) {}

  // 契約書ページ（?page=contract）
  if (page === 'contract') {
    var ctmpl = HtmlService.createTemplateFromFile('contract');
    ctmpl.companySsId = ssId;
    ctmpl.companyName = (e && e.parameter && e.parameter.company)    ? e.parameter.company    : '';
    ctmpl.adminEmail  = (e && e.parameter && e.parameter.adminEmail) ? e.parameter.adminEmail : '';
    ctmpl.contractRow = (e && e.parameter && e.parameter.row)        ? e.parameter.row        : '';
    return ctmpl.evaluate()
      .setTitle('利用規約 - 運行管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  var tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.companySsId = ssId;
  return tmpl.evaluate()
    .setTitle('運行管理システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// ================================================================
//  2-2b: ssIdをUserPropertiesに保存（storeCompanySsId）  【大A / 中2 / 小2-2b】
//  index.htmlのwindow.onload時にssIdがあれば呼ばれる。
//  以降 getTargetSS_() でそのSSを開けるようにする。
// ================================================================
function storeCompanySsId(ssId) {
  if (ssId) PropertiesService.getUserProperties().setProperty('linkedSsId', ssId);
}


// ================================================================
//  2-2c: 対象スプレッドシート取得（getTargetSS_）  【大B / 中2 / 小2-2c】
//  乗務員アプリからの呼び出し時は linkedSsId で会社SSを openById で開く。
//  スプレッドシートメニューからの呼び出し時は getActiveSpreadsheet を返す。
// ================================================================
function getTargetSS_(ssId) {
  var id = ssId || PropertiesService.getUserProperties().getProperty('linkedSsId');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch(e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ================================================================
//  2-2b: ドライバー認証（validateDriverEmail_）
//  自車専属マスタのJ列にメールアドレスが存在するか検証する
//  部外者アクセスは例外をスローして即遮断
// ================================================================
function validateDriverEmail_(email, companySsId) {
  if (!email || String(email).trim() === '') throw new Error('認証エラー：メールアドレスがありません');
  var ss     = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master || master.getLastRow() < 2) throw new Error('認証エラー：マスタが存在しません');
  var emails = master.getRange(2, 11, master.getLastRow() - 1, 1).getValues();
  var normalized = String(email).trim().toLowerCase();
  for (var i = 0; i < emails.length; i++) {
    if (String(emails[i][0] || '').trim().toLowerCase() === normalized) return true;
  }
  throw new Error('認証エラー：未登録のアクセスです');
}


// ================================================================
//  2-3: サイドバー表示（showSidebar）  【大C / 中2 / 小2-3】
//  スプレッドシートのサイドバーとして表示する
// ================================================================
function showSidebar() {
  var tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.companySsId = '';
  var html = tmpl.evaluate().setTitle('ホーム').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}


// ================================================================
//  2-4: 写真・ファイル取込サイドバー（showUploadSidebar）  【大C / 中2 / 小2-4】
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
    '  document.querySelector("button").disabled=true;' +
    '  document.getElementById("msg").innerText="アップロード中...";' +
    '  var done=0,total=files.length,urls=[];' +
    '  files.forEach(function(file){' +
    '    if(file.size>20*1024*1024){done++;check("（20MB超のためスキップ）");return;}' +
    '    var r=new FileReader();' +
    '    r.onload=function(){' +
    '      var b64=r.result.split(",")[1];' +
    '      google.script.run' +
    '        .withSuccessHandler(function(res){done++;if(res&&res.url)urls.push(res.url);check("");})' +
    '        .withFailureHandler(function(err){done++;check("エラー："+err.message);})' +
    '        .uploadFileToRow(' + row + ',file.name,b64,file.type);' +
    '    };' +
    '    r.readAsDataURL(file);' +
    '  });' +
    '  function check(note){' +
    '    if(done<total)return;' +
    '    var msg="✅ "+done+"件 完了。SSのW列にリンクが追加されました。"+note;' +
    '    document.getElementById("msg").innerText=msg;' +
    '    document.querySelector("button").disabled=false;' +
    '  }' +
    '}' +
    '<\/script></body></html>';

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(body).setTitle('📷 写真・ファイル取込').setWidth(280)
  );
}


// ================================================================
//  3-1: onEdit本体  【大B / 中3 / 小3-1】
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
    var ss        = sheet.getParent();

    // ── 3-1-1: 集計表 編集ブロック ──────────────────────────────────
    // 距離(V=22)・ガソリン代(X=24)・支払い(Z=26)・備考(AB=28)以外の列は直接編集を禁止する
    // 禁止列が触れられた場合: IDがあれば集計表を再同期して正しい値に戻す
    //                         IDがなく単一セルなら旧値に戻す
    if (sheetName === '集計表' && row > 1) {
      var allowed = [23, 25, 27, 30, 35]; // W=距離, Y=ガソリン代, AA=支払い, AD=備考, AI=その他手当
      var numC = range.getNumColumns(), numR = range.getNumRows();
      var blocked = false;
      for (var c = 0; c < numC; c++) {
        if (allowed.indexOf(col + c) === -1) { blocked = true; break; }
      }
      if (blocked) {
        var bRowId = sheet.getRange(row, 1).getValue();
        if (bRowId) {
          delaySyncSummary_(bRowId, ss); // IDあり→再同期で正しい値・書式に戻す
        } else if (numR === 1 && numC === 1) {
          range.setValue(e.oldValue !== undefined ? e.oldValue : '');
        }
        ss.toast('この列は編集できません', '⛔ 保護', 3);
      } else {
        var rowId = sheet.getRange(row, 1).getValue();
        if (rowId) delaySyncSummary_(rowId, ss);
      }
      return;
    }

    // ── 3-1-2: 運行シート V列(22) 合計高速 保護 ──────────────────────
    // 合計高速は「実費高速 - 請求高速」の自動計算列なので直接編集を禁止する
    // 編集されたら即座に数式を復元する
    if (sheetName === '運行' && row > 1 && col === 22) {
      range.setFormula('=IF(AND(U'+row+'="",T'+row+'=""),"",U'+row+'-T'+row+')');
      ss.toast('合計高速は自動計算列です', '⛔ 保護', 3);
      return;
    }

    // ── 3-1-2b: 運行シート Y列(25)連絡(端末)・Z列(26)データ(端末) 保護 ────
    // 削除・上書きされても元の値に戻す
    if (sheetName === '運行' && row > 1 && (col === 25 || col === 26)) {
      if (col === 25) {
        var oldV25 = (e.oldValue !== undefined) ? String(e.oldValue) : '';
        range.setValue(oldV25);
      } else {
        var restoreUrls = getTerminalUrls_(sheet, row);
        if (restoreUrls.length > 0) {
          setTerminalUrls_(sheet, row, restoreUrls);
        } else {
          range.clearContent();
        }
      }
      ss.toast((col === 25 ? '連絡(端末)' : 'データ(端末)') + 'はアプリからのみ変更できます', '⛔ 保護', 4);
      return;
    }

    // ── 3-1-3: 運行シート X列(24) URL自動変換 ───────────────────────
    // X列にURLをペーストすると自動的にリッチテキストリンク（「ファイル1」等）に変換する
    // 改行・全角/半角カンマ・スペース区切りで複数URLに対応
    // DriveやDocs以外のURL（Googleフォト等）はDriveに取り込んで公開URLに変換を試みる
    // 既存のURLはノートに保存されており、新しいURLと結合して上書きすることで追加になる
    if (sheetName === '運行' && row > 1 && col === 24) {
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
    if (sheetName === '自車専属マスタ') { onEditMasterVehicle_(sheet, range, ss); return; }
    if (sheetName === 'マスタ')         { onEditMasterCustomer_(sheet, range); return; }
    // 情報（マッチング）シートの自動処理：進捗着色・TEL/FAX自動入力
    if (sheetName === '情報')           { onEditJoho_(sheet, range, ss); return; }
    // 会社登録シートの処理はインストール型トリガー（installedOnEdit_）が担当する
    // シンプルトリガーはドライブ/メール/トリガー作成などの認証付き操作が不可のため
    if (sheetName === '会社登録') return;
    if (sheetName !== '運行') return;
    onEditUnkou_(sheet, range, ss);
  } catch (err) {}
}


// ================================================================
//  3-2: 運行シート編集時の処理（onEditUnkou_）  【大B / 中3 / 小3-2】
//  ・A列が空で他列にデータがあればV-XXXXのIDを自動生成
//  ・I列（日付）が00:00:00なら現在時刻を自動付加
//  ・F列（車番）編集時に自車専属マスタから区分〜携帯番号を自動補完
//  ・T列（合計高速）の数式を自動セット
//  ・集計表を同期し、孤立IDを削除
// ================================================================
function onEditUnkou_(sheet, range, ss) {
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var editedCol = range.getColumn();
  if (startRow <= 1) return;
  if (!ss) ss = sheet.getParent();
  var master = ss.getSheetByName('自車専属マスタ');
  var mData = master ? master.getDataRange().getValues() : [];
  // ScriptLockでID採番を排他制御（並列アクセス時のV-番号重複を根絶）
  var idLock = LockService.getScriptLock();
  try { idLock.waitLock(10000); } catch(e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('ロック取得失敗。再度お試しください。', '⚠️', 5);
    return;
  }
  var vNextIdNum = getNextIdNum_(sheet, 'V-');
  for (var pi = 0; pi < numRows; pi++) {
    var prow = startRow + pi;
    if (prow <= 1) continue;
    var pidCell = sheet.getRange(prow, 1);
    if (!pidCell.getValue()) {
      var phd = sheet.getRange(prow, 2, 1, 10).getValues()[0].some(function(v) { return v !== ''; });
      if (phd) { pidCell.setValue('V-' + String(vNextIdNum).padStart(4, '0')); vNextIdNum++; }
    }
  }
  SpreadsheetApp.flush();
  idLock.releaseLock();

  var lastColU = Math.max(sheet.getLastColumn(), 22);
  for (var i = 0; i < numRows; i++) {
    var row = startRow + i;
    if (row <= 1) continue;
    // 行データを1回で一括読み込み（個別getValue廃止）
    var rowData = sheet.getRange(row, 1, 1, lastColU).getValues()[0];
    var currentId = String(rowData[0] || '').trim();
    var dateVal   = rowData[9]; // J列(10)=日付 0-indexed:9

    // J列(10)の日付：時刻部分が 0:00:00 なら現在時刻を付与（日付のみ入力に対応）
    if (dateVal instanceof Date) {
      if (dateVal.getHours() === 0 && dateVal.getMinutes() === 0 && dateVal.getSeconds() === 0) {
        var now = new Date();
        var merged = new Date(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate(),
                              now.getHours(), now.getMinutes(), now.getSeconds());
        sheet.getRange(row, 10).setValue(merged);
        rowData[9] = merged;
      }
    }
    sheet.getRange(row, 10).setNumberFormat('yyyy/MM/dd');

    // F列(6)：車番を入力→自車専属マスタと部分一致で8列一括補完
    if (editedCol === 6 && range.getNumColumns() === 1) {
      var inputCar = String(rowData[5] || '').trim(); // F列 0-indexed:5
      if (inputCar && mData.length > 1) {
        // B〜I列（F=index4を除く）に既存値があれば補完しない
        var hasPreInput = false;
        for (var ci = 0; ci < 8; ci++) {
          if (ci === 4) continue; // F列スキップ
          if (String(rowData[1 + ci] || '').trim() !== '') { hasPreInput = true; break; }
        }
        if (!hasPreInput) {
          for (var m2 = 1; m2 < mData.length; m2++) {
            var masterCar = String(mData[m2][7] || '').trim();
            var masterStatus = String(mData[m2][1] || '').trim();
            if (masterStatus === '故障' || masterStatus === '待機') continue;
            if (masterCar === inputCar || masterCar.indexOf(inputCar) !== -1 || inputCar.indexOf(masterCar) !== -1) {
              // B〜I列を1回のsetValuesで一括書き込み（8個→1回）
              sheet.getRange(row, 2, 1, 8).setValues([[
                mData[m2][2], mData[m2][3], mData[m2][5], mData[m2][6],
                masterCar, mData[m2][8], mData[m2][9], mData[m2][4]
              ]]);
              break;
            }
          }
        }
      }
    }

    // N/O/P/Q/R列（誘導・積完・休憩・降完時刻）：全角文字・日付なし時刻を正規化して合成
    if ([14, 15, 16, 17, 18].indexOf(editedCol) !== -1) {
      var timeCell = sheet.getRange(row, editedCol);
      var tv = timeCell.getValue();
      var baseDateObj = (rowData[9] instanceof Date) ? rowData[9] : null;
      var mergedT = null;
      if (typeof tv === 'string' && tv.trim() !== '') {
        var s = tv.trim().replace(/[：]/g, ':').replace(/[　]/g, ' ');
        var m1 = s.match(/^(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})/);
        if (m1) {
          var yr = baseDateObj ? baseDateObj.getFullYear() : new Date().getFullYear();
          mergedT = new Date(yr, parseInt(m1[1]) - 1, parseInt(m1[2]), parseInt(m1[3]), parseInt(m1[4]), 0);
        } else {
          var m2t = s.match(/^(\d{1,2}):(\d{2})$/);
          if (m2t) {
            var base2 = baseDateObj || new Date();
            mergedT = new Date(base2.getFullYear(), base2.getMonth(), base2.getDate(),
                               parseInt(m2t[1]), parseInt(m2t[2]), 0);
          }
        }
      } else if (tv instanceof Date && tv.getFullYear() < 1902) {
        var base3 = baseDateObj || new Date();
        mergedT = new Date(base3.getFullYear(), base3.getMonth(), base3.getDate(),
                           tv.getHours(), tv.getMinutes(), tv.getSeconds());
      }
      if (mergedT && !isNaN(mergedT.getTime())) {
        timeCell.setValue(mergedT);
        timeCell.setNumberFormat('M/d HH:mm');
        if (range.getNumRows() === 1 && range.getNumColumns() === 1) {
          var rowTimes = sheet.getRange(row, 14, 1, 5).getValues()[0];
          var gv2=rowTimes[0], pv2=rowTimes[1], rsv2=rowTimes[2], rev2=rowTimes[3];
          var gapMsg = null;
          if (editedCol===15 && !gv2) gapMsg='先に誘導時刻を入力してください';
          else if (editedCol===16) gapMsg = !gv2?'先に誘導時刻を入力してください':!pv2?'先に積完時刻を入力してください':null;
          else if (editedCol===17) gapMsg = !gv2?'先に誘導時刻を入力してください':!pv2?'先に積完時刻を入力してください':!rsv2?'先に休憩開始を入力してください':null;
          else if (editedCol===18) gapMsg = !gv2?'先に誘導時刻を入力してください':!pv2?'先に積完時刻を入力してください':!rsv2?'先に休憩開始を入力してください':!rev2?'先に休憩終了を入力してください':null;
          if (gapMsg) {
            timeCell.clearContent();
            SpreadsheetApp.getActiveSpreadsheet().toast(gapMsg, '⛔ 順序エラー', 4);
            continue;
          }
        }
      }
    }

    // 積地(L=col12)の背景色を即座に設定（単一setBackgroundsで確実に反映）
    var pvK = String(rowData[11] || ''); // L列 0-indexed:11
    var rowBgNew = new Array(lastColU).fill(null);
    if (pvK.indexOf('有休') !== -1) {
      rowBgNew = new Array(lastColU).fill('#e0e0e0');
    } else if (pvK.indexOf('休み') !== -1) {
      rowBgNew = new Array(lastColU).fill('#9e9e9e');
    } else if (pvK === '' && currentId) {
      rowBgNew[11] = '#fff9c4'; // L列のみ黄色
    }
    sheet.getRange(row, 1, 1, lastColU).setBackgrounds([rowBgNew]);

    // T列(20)=請求高速 を入力したとき U列(21)=実費高速 が空なら自動コピー（オレンジ色）
    if (editedCol === 20 && (rowData[20] === '' || rowData[20] === null)) {
      if (rowData[19] !== '' && rowData[19] !== null) {
        var uAutoCell = sheet.getRange(row, 21);
        uAutoCell.setValue(rowData[19]);
        uAutoCell.setFontColor('#E65100');
      }
    }
    // U列(21)=実費高速 を手入力したとき黒字に戻す
    if (editedCol === 21) {
      sheet.getRange(row, 21).setFontColor(null);
    }

    // V列(22)の合計高速数式
    if (!sheet.getRange(row, 22).getFormula()) {
      sheet.getRange(row, 22).setFormula('=IF(AND(U' + row + '="",T' + row + '=""),"",U' + row + '-T' + row + ')');
    }
    SpreadsheetApp.flush();
    if (currentId) syncSummaryForId_(currentId, ss);
  }
  applyMoneyFormat_(sheet, startRow, numRows, 'unkou');
  applyDateTimeFormat_(sheet, startRow, numRows);
  // 日付列(J=col10)が編集された場合は両シートをソート（色はsortUnkouByDate_内で一緒に移動）
  if (editedCol === 10) {
    sortUnkouByDate_();
    sortSummaryByDate_();
  }
}


// ================================================================
//  3-3: 自車専属マスタ編集時の処理（onEditMasterVehicle_）  【大B / 中3 / 小3-3】
//  ・A列が空で他列にデータがあればS-XXXXのIDを自動生成
//  ・B列（運行状態）の値に応じて行の背景色を変更
//    運行→薄赤, 待機→薄黄, 故障→薄緑, その他→なし
//  ・自車専属運行シートを自動更新
// ================================================================
function onEditMasterVehicle_(sheet, range, ss) {
  var startRow = range.getRow();
  var numRows  = range.getNumRows();
  if (!ss) ss = sheet.getParent();
  var editedStartCol = range.getColumn();
  var editedEndCol   = editedStartCol + range.getNumColumns() - 1;
  var lastCol = Math.max(sheet.getLastColumn(), 17);

  // ── 先に全行の背景色を一括設定（タイムアウトで途中で終わっても色だけは確実に反映） ──
  // ScriptLockでID採番を排他制御（並列アクセス時のS-番号重複を根絶）
  var sIdLock = LockService.getScriptLock();
  try { sIdLock.waitLock(10000); } catch(e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('ロック取得失敗。再度お試しください。', '⚠️', 5);
    return;
  }
  var sNextIdNum = getNextIdNum_(sheet, 'S-');
  for (var spi = 0; spi < numRows; spi++) {
    var sprow = startRow + spi;
    if (sprow <= 1) continue;
    var spCell = sheet.getRange(sprow, 1);
    if (!spCell.getValue()) {
      var sphd = sheet.getRange(sprow, 2, 1, 5).getValues()[0].some(function(v) { return v !== ''; });
      if (sphd) { spCell.setValue('S-' + String(sNextIdNum).padStart(4, '0')); sNextIdNum++; }
    }
  }
  SpreadsheetApp.flush();
  sIdLock.releaseLock();

  if (numRows > 0 && startRow > 1) {
    var statusVals = sheet.getRange(startRow, 2, numRows, 1).getValues();
    for (var ci = 0; ci < numRows; ci++) {
      var crow = startRow + ci;
      if (crow <= 1) continue;
      var cStatus = String(statusVals[ci][0] || '').trim();
      var cRange  = sheet.getRange(crow, 1, 1, lastCol);
      if      (cStatus === '運行') cRange.setBackground('#ffcdd2');
      else if (cStatus === '待機') cRange.setBackground('#fff9c4');
      else if (cStatus === '故障') cRange.setBackground('#c8e6c9');
      else                         cRange.setBackground(null);
    }
  }

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
    // F列（col6）のトン数に対応する燃費をL列（col12）に自動反映
    var tonsRaw = String(sheet.getRange(row, 6).getValue()).trim();
    if (tonsRaw) {
      var tonsNorm = tonsRaw
        .replace(/[０-９]/g, function(c){return String.fromCharCode(c.charCodeAt(0)-0xFEE0);})
        .replace(/[ｔＴ]/g,'t').toLowerCase();
      var numPart = tonsNorm.replace(/t$/,'');
      var fuel = fuelMap[tonsNorm] || fuelMap[numPart+'t'] || fuelMap[numPart] || '';
      if (fuel !== '' && fuel !== undefined) sheet.getRange(row, 12).setValue(fuel);
    }

    // 行データ読み取り（N=14=仮日数, O=15=給料, P=16=%, Q=17=高速を引く）
    var mRow = sheet.getRange(row, 1, 1, 17).getValues()[0];
    var mCar    = String(mRow[7]  || '').trim();
    var mName   = String(mRow[8]  || '').trim();
    var mKari   = mRow[13]; // N=14
    var mKyuryo = mRow[14]; // O=15
    var mPct    = mRow[15]; // P=16

    // 排他制御: 給料と%は排他。仮日数は経費按分用なので%・給料どちらとも共存可能。
    var kariEdited   = editedStartCol <= 14 && editedEndCol >= 14;
    var kyuryoEdited = editedStartCol <= 15 && editedEndCol >= 15;
    var pctEdited    = editedStartCol <= 16 && editedEndCol >= 16;
    if (kyuryoEdited && mKyuryo !== '') {
      // 給料に入力 → %をクリア
      mPct = '';
      sheet.getRange(row, 16).clearContent();
    } else if (pctEdited && mPct !== '') {
      // %に入力 → 給料のみクリア（仮日数は保持）
      mKyuryo = '';
      sheet.getRange(row, 15).clearContent();
    }
    // 仮日数に入力しても%・給料は変更しない

    // 集計表の該当車番_乗務員名の全行をsyncSummaryForId_で即反映
    if (mCar || mName) {
      var sumSheetM = ss.getSheetByName('集計表');
      if (sumSheetM && sumSheetM.getLastRow() >= 2) {
        var sumIdsM = sumSheetM.getRange(2, 1, sumSheetM.getLastRow()-1, 7).getValues();
        var targetIds = [];
        for (var si = 0; si < sumIdsM.length; si++) {
          var sCar  = String(sumIdsM[si][5] || '').trim(); // F列=車番
          var sName = String(sumIdsM[si][6] || '').trim(); // G列=乗務員名
          if (sCar === mCar && sName === mName) {
            var tid = String(sumIdsM[si][0] || '').trim();
            if (tid) targetIds.push(tid);
          }
        }
        for (var ti = 0; ti < targetIds.length; ti++) {
          try { syncSummaryForId_(targetIds[ti], ss); } catch(e) {}
        }
      }
    }

    // 高速を引く(Q=17): %なし or 仮日数/給料あり → 適用外でグレー
    var pctNum2 = Number(mPct) || 0;
    var kariNum2 = Number(mKari) || 0;
    var kyuNum2  = Number(mKyuryo) || 0;
    var isQGray = (!pctNum2 || kyuNum2 > 0 || kariNum2 > 0);
    var rowBg2 = (function(){
      var st = String(sheet.getRange(row, 2).getValue() || '').trim();
      return st === '運行' ? '#ffcdd2' : st === '待機' ? '#fff9c4' : st === '故障' ? '#c8e6c9' : null;
    })();
    sheet.getRange(row, 17).setBackground(isQGray ? '#e0e0e0' : rowBg2);

    // B列（ステータス）が変更された場合、運行シートを今日以降で同期する
    if (editedStartCol <= 2 && editedEndCol >= 2) {
      var mRowData = sheet.getRange(row, 1, 1, 16).getValues()[0];
      syncVehicleToCurrentMonth_(mRowData, false);
    }
    // 経費列（Q=17〜AE=31）を手入力したら文字色を黒にリセット（自動入力の赤を解除）
    var expS = Math.max(editedStartCol, 17);
    var expE = Math.min(editedEndCol, 31);
    if (expS <= expE) {
      sheet.getRange(row, expS, 1, expE - expS + 1).setFontColor(null);
    }
  }
  refreshActiveVehiclesAuto_();
  applyMasterVehicleWarnings_(sheet);
  try { CacheService.getScriptCache().remove('cfg_master'); } catch(e) {}
}


// ================================================================
//  自車専属マスタ全行の支払条件不備を警告（N=仮日数/O=給料/P=%列のみ一括更新）
// ================================================================
function applyMasterVehicleWarnings_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;
  var lr        = sheet.getLastRow();
  var lc        = Math.max(sheet.getLastColumn(), 32);
  var hdr       = sheet.getRange(1, 1, 1, lc).getValues()[0];
  var data      = sheet.getRange(2, 1, lr - 1, lc).getValues();
  var hasToll   = String(hdr[16] || '').trim() === '高速を引く（引くは〇、引かないは空欄）';
  var expStartI = hasToll ? 17 : 16; // 0-based（payCondMapビルダーと同ロジック）
  var warnBgs   = [];
  for (var r = 0; r < data.length; r++) {
    var kariN  = Number(data[r][13]) || 0;
    var kyuN   = Number(data[r][14]) || 0;
    var pctN   = Number(data[r][15]) || 0;
    var expTot = 0;
    for (var e = expStartI; e <= expStartI + 14; e++) expTot += Number(data[r][e]) || 0;
    var status = String(data[r][1] || '').trim();
    var rowBg  = status === '運行' ? '#ffcdd2' : status === '待機' ? '#fff9c4' : status === '故障' ? '#c8e6c9' : null;
    // 仮日数なし かつ 給料/%/経費のいずれかあり → N列(仮日数)をグレー警告
    var nBg = (!kariN && (kyuN > 0 || pctN > 0 || expTot > 0)) ? '#b0bec5' : rowBg;
    // 仮日数あり かつ 給料も%もなし → O列(給料)とP列(%)を両方グレー警告
    var opBg = (kariN > 0 && !kyuN && !pctN) ? '#b0bec5' : rowBg;
    // %なし または 仮日数/給料あり → Q列(高速を引く)はグレー（適用外）
    var qBg = (!pctN || kyuN > 0 || kariN > 0) ? '#e0e0e0' : rowBg;
    warnBgs.push([nBg, opBg, opBg, qBg]);
  }
  sheet.getRange(2, 14, lr - 1, 4).setBackgrounds(warnBgs); // N,O,P,Q
}


// ================================================================
//  3-3b: 車両ステータス変更時の運行シート同期（syncVehicleToCurrentMonth_）
//  自車専属マスタのB列（運行/故障/待機）変更時に呼ばれる
//  ・今日以降の積地空（未配車）行を削除
//  ・ステータスが「運行」なら今日〜今月末の行を再生成
//  ・skipSort=true のとき並び替え・色付けをスキップ（一括処理用）
// ================================================================
function syncVehicleToCurrentMonth_(veh, skipSort) {
  var carNo  = String(veh[7] || '').trim(); // H列(index7)=車番
  var status = String(veh[1] || '').trim(); // B列(index1)=ステータス
  if (!carNo) return;
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var today = new Date(); today.setHours(0, 0, 0, 0);
  // ① 今日以降の積地空行を削除（過去・積地ありは触らない）
  if (sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
    var toDelete = [];
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][5] || '').trim() !== carNo) continue;
      var d = data[i][9];
      if (!(d instanceof Date)) continue;
      var dm = new Date(d); dm.setHours(0, 0, 0, 0);
      if (dm < today) continue;
      if (String(data[i][11] || '').trim() !== '') continue; // 積地ありは保護
      toDelete.push(i + 2);
    }
    for (var r = toDelete.length - 1; r >= 0; r--) sheet.deleteRow(toDelete[r]);
  }
  // ② 集計表の孤立ID（削除した行のID）を除去
  cleanAllOrphanSummary_();
  // ③ ステータスが「運行」なら今日〜今月末の行を生成（既存行程のある日はスキップ）
  if (status === '運行') {
    var now = new Date();
    var yr = now.getFullYear(), mo = now.getMonth(), startDay = now.getDate();
    var endDay = new Date(yr, mo + 1, 0).getDate();
    // 削除後の現状を読み直して、この車番で既に行がある日付を収集
    var existingDates = {};
    if (sheet.getLastRow() >= 2) {
      var remainData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
      for (var ei = 0; ei < remainData.length; ei++) {
        if (String(remainData[ei][5] || '').trim() !== carNo) continue;
        var ed = remainData[ei][9];
        if (!(ed instanceof Date)) continue;
        var edMid = new Date(ed); edMid.setHours(0, 0, 0, 0);
        existingDates[edMid.getTime()] = true;
      }
    }
    var lock = LockService.getScriptLock();
    try { lock.waitLock(30000); } catch(e) { return; }
    try {
      var insertRow = sheet.getLastRow() + 1;
      var nextNum   = getNextIdNum_(sheet, 'V-');
      var rowsData  = [], formulas = [];
      for (var day = startDay; day <= endDay; day++) {
        var dateMid = new Date(yr, mo, day); dateMid.setHours(0, 0, 0, 0);
        if (existingDates[dateMid.getTime()]) continue; // 既存行程のある日はスキップ
        var rowId = 'V-' + String(nextNum).padStart(4, '0'); nextNum++;
        var rn    = insertRow + rowsData.length;
        rowsData.push([rowId, veh[2], veh[3], veh[5], veh[6], veh[7], veh[8], veh[9],
          veh[4], new Date(yr, mo, day), '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        formulas.push(['=IF(AND(U'+rn+'="",T'+rn+'=""),"",U'+rn+'-T'+rn+')']);
      }
      if (rowsData.length > 0) {
        sheet.getRange(insertRow, 1, rowsData.length, 26).setValues(rowsData);
        sheet.getRange(insertRow, 22, formulas.length, 1).setFormulas(formulas);
        sheet.getRange(insertRow, 10, rowsData.length, 1).setNumberFormat('yyyy/MM/dd');
        sheet.getRange(insertRow, 12, rowsData.length, 2).setNumberFormat('@');
      }
      SpreadsheetApp.flush();
    } finally { lock.releaseLock(); }
  }
  if (!skipSort) { sortUnkouByDate_(); applyHolidayRowColors_(); }
}


// ================================================================
//  3-3c: 全車両を今月分に同期（syncAllVehiclesToCurrentMonth_）
//  expandAndRefreshSheets から呼ばれる全車両一括同期
// ================================================================
function syncAllVehiclesToCurrentMonth_() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master || master.getLastRow() < 2) return;
  var mData = master.getRange(2, 1, master.getLastRow() - 1, 16).getValues();
  for (var v = 0; v < mData.length; v++) {
    syncVehicleToCurrentMonth_(mData[v], true); // 並び替えは最後に1回だけ
  }
  sortUnkouByDate_();
  applyHolidayRowColors_();
  cleanAllOrphanSummary_();
}


// ================================================================
//  3-4: マスタ（取引先）編集時の処理（onEditMasterCustomer_）  【大B / 中3 / 小3-4】
//  ・A列が空で他列にデータがあればM-XXXXのIDを自動生成
// ================================================================
function onEditMasterCustomer_(sheet, range) {
  var startRow = range.getRow();
  var numRows = range.getNumRows();
  // ScriptLockでID採番を排他制御（並列アクセス時のM-番号重複を根絶）
  var mIdLock = LockService.getScriptLock();
  try { mIdLock.waitLock(10000); } catch(e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('ロック取得失敗。再度お試しください。', '⚠️', 5);
    return;
  }
  try {
    var mNextIdNum = getNextIdNum_(sheet, 'M-');
    for (var i = 0; i < numRows; i++) {
      var row = startRow + i;
      if (row <= 1) continue;
      var idCell = sheet.getRange(row, 1);
      if (!idCell.getValue()) {
        var hasData = sheet.getRange(row, 2, 1, 5).getValues()[0].some(function(v) { return v !== ''; });
        if (hasData) {
          idCell.setValue('M-' + String(mNextIdNum).padStart(4, '0'));
          mNextIdNum++;
        }
      }
    }
    SpreadsheetApp.flush();
  } finally {
    mIdLock.releaseLock();
  }
}


// ================================================================
//  3-5: 情報シート編集時の処理（onEditJoho_）  【大B / 中3 / 小3-5】
//  「情報」シートでセルを編集した際に自動実行される内部関数。
//  ① B列（進捗）変更 → 行全体の背景色を自動変更（放置=白/キャンセル・終了=グレー/確定=黄）
//  ② C列（会社名・貨物）変更 → 取引先マスタを検索しTEL(D列)・FAX(E列)を自動入力
//  ③ N列（会社名・車両）変更 → 自車専属マスタ→取引先マスタの順で検索しTEL(O列)・FAX(P列)を自動入力
//
//  M&A向け補足: このトリガーにより「手入力ゼロ」の配車受付が実現できる。
//              貨物側は取引先マスタ（荷主情報）、車両側は自社車両マスタから
//              TEL/FAXを自動補完するため、記録ミスや入力漏れを防止する。
// ================================================================
function onEditJoho_(sheet, range, ss) {
  var col = range.getColumn();
  var row = range.getRow();
  if (row < 2) return; // 1行目はヘッダーなので処理しない

  // ── ① B列（進捗・貨物）変更時: 貨物セクション(A〜M = 列1〜13)のみ着色 ──
  // 車両側(N〜Z)の色は変えない。貨物の状況を独立して管理できるようにする。
  if (col === 2) {
    var prog = String(range.getValue() || '').trim();
    var cargoRange = sheet.getRange(row, 1, 1, 13); // A〜M列
    if (prog === '確定') {
      cargoRange.setBackground('#fff9c4'); // 黄色
    } else if (prog === 'キャンセル' || prog === '終了') {
      cargoRange.setBackground('#eeeeee'); // グレー
    } else {
      cargoRange.setBackground(null); // リセット（白）
    }
    return;
  }

  // ── ①' O列（進捗・車両）変更時: 車両セクション(N〜Z = 列14〜26)のみ着色 ──
  // 貨物側(A〜M)の色は変えない。車両の状況を独立して管理できるようにする。
  if (col === 15) {
    var vprog = String(range.getValue() || '').trim();
    var vehRange = sheet.getRange(row, 14, 1, 13); // N〜Z列
    if (vprog === '確定') {
      vehRange.setBackground('#fff9c4'); // 黄色
    } else if (vprog === 'キャンセル' || vprog === '終了') {
      vehRange.setBackground('#eeeeee'); // グレー
    } else {
      vehRange.setBackground(null); // リセット（白）
    }
    return;
  }

  // ── ② C列（会社名・貨物）変更時: 取引先マスタからTEL/FAX自動入力 ──
  // 荷主名を取引先マスタの「会社名」列で検索し、C=電話→D列、FAX→E列に転記する
  if (col === 3) {
    var cargoCompany = String(range.getValue() || '').trim();
    if (!cargoCompany) return;
    var custSh = ss.getSheetByName('マスタ');
    if (!custSh || custSh.getLastRow() < 2) return;
    var cHdrs = custSh.getRange(1, 1, 1, custSh.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h || '').trim(); });
    var cData = custSh.getRange(2, 1, custSh.getLastRow() - 1, custSh.getLastColumn()).getValues();
    var nIdx = cHdrs.indexOf('会社名'), telIdx = cHdrs.indexOf('電話'), faxIdx = cHdrs.indexOf('FAX');
    for (var i = 0; i < cData.length; i++) {
      if (String(cData[i][nIdx] || '').trim() === cargoCompany) {
        if (telIdx >= 0) sheet.getRange(row, 4).setValue(cData[i][telIdx]); // D列 TEL
        if (faxIdx >= 0) sheet.getRange(row, 5).setValue(cData[i][faxIdx]); // E列 FAX
        break;
      }
    }
    return;
  }

  // ── ③ P列(16)（会社名・車両）変更時: 自社マスタ→取引先マスタの順に検索 ──
  // N=チェック, O=進捗(車両)なのでスキップ。P列(16)が車両側の会社名入力列。
  // 自車専属マスタには携帯番号のみ存在するため、FAXは取引先マスタのみ参照する
  if (col === 16) {
    var vehCompany = String(range.getValue() || '').trim();
    if (!vehCompany) return;

    // 自車専属マスタを先に検索（自社車両の会社名 → 携帯番号をTELとして使用）
    var vmSh = ss.getSheetByName('自車専属マスタ');
    if (vmSh && vmSh.getLastRow() >= 2) {
      var vmHdrs = vmSh.getRange(1, 1, 1, vmSh.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h || '').trim(); });
      var vmData = vmSh.getRange(2, 1, vmSh.getLastRow() - 1, vmSh.getLastColumn()).getValues();
      var vmNIdx = vmHdrs.indexOf('会社名'), vmTelIdx = vmHdrs.indexOf('携帯番号');
      for (var vi = 0; vi < vmData.length; vi++) {
        if (String(vmData[vi][vmNIdx] || '').trim() === vehCompany) {
          if (vmTelIdx >= 0) sheet.getRange(row, 17).setValue(vmData[vi][vmTelIdx]); // Q列 TEL
          break;
        }
      }
    }

    // 取引先マスタでも検索（協力会社の場合、FAXも取得できる）
    var custSh2 = ss.getSheetByName('マスタ');
    if (custSh2 && custSh2.getLastRow() >= 2) {
      var csHdrs = custSh2.getRange(1, 1, 1, custSh2.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h || '').trim(); });
      var csData = custSh2.getRange(2, 1, custSh2.getLastRow() - 1, custSh2.getLastColumn()).getValues();
      var csNIdx = csHdrs.indexOf('会社名'), csTelIdx = csHdrs.indexOf('電話'), csFaxIdx = csHdrs.indexOf('FAX');
      for (var csi = 0; csi < csData.length; csi++) {
        if (String(csData[csi][csNIdx] || '').trim() === vehCompany) {
          if (csTelIdx >= 0) sheet.getRange(row, 17).setValue(csData[csi][csTelIdx]); // Q列 TEL
          if (csFaxIdx >= 0) sheet.getRange(row, 18).setValue(csData[csi][csFaxIdx]); // R列 FAX
          break;
        }
      }
    }
    return;
  }
}


// ================================================================
//  4-1: 集計表再生成（generateSummary）  【大B / 中4 / 小4-1】
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

  // 自車専属マスタから 車番+乗務員名 → 仮日数/給料/%/高速控除/月間経費 の支払条件マップを作成
  var master = ss.getSheetByName('自車専属マスタ');
  var payCondMap = {};
  if (master && master.getLastRow() >= 2) {
    var mReadCols = Math.max(master.getLastColumn(), 31);
    var mData = master.getRange(2, 1, master.getLastRow()-1, mReadCols).getValues();
    // 高速を引く列（col17=idx16）が存在するか判定して経費開始インデックスを決定
    var mHdrRow = master.getRange(1, 1, 1, mReadCols).getValues()[0];
    var mHasTollDeduct = String(mHdrRow[16]||'').trim() === '高速を引く（引くは〇、引かないは空欄）';
    var mExpStart = mHasTollDeduct ? 17 : 16;
    for (var m = 0; m < mData.length; m++) {
      var mcar  = String(mData[m][7]  || '').trim();
      var mname = String(mData[m][8]  || '').trim();
      var pkey  = mcar + '_' + mname;
      var mExp  = 0;
      for (var ei = mExpStart; ei <= mExpStart + 14; ei++) mExp += Number(mData[m][ei]) || 0;
      payCondMap[pkey] = {
        kari:       mData[m][13] || '',
        kyuryo:     mData[m][14] || '',
        pct:        mData[m][15] || '',
        tollDeduct: mHasTollDeduct ? String(mData[m][16] || '').trim() : '',
        expense:    mExp
      };
    }
  }

  // 既存の集計表から手入力済みの距離・ガソリン代・支払・備考・支払条件を退避（再生成で消えないように）
  var sumSheet = ss.getSheetByName('集計表');
  var oldData = {};
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var colCount = sumSheet.getLastColumn();
    var oldRows = sumSheet.getRange(2, 1, sumSheet.getLastRow()-1, Math.max(colCount, 35)).getValues();
    var isNewSumLayout = colCount >= 35; // 経費合計列（AB=28）が追加済みか
    for (var o = 0; o < oldRows.length; o++) {
      var oldId = String(oldRows[o][0] || '').trim();
      if (oldId) {
        oldData[oldId] = {
          distance: oldRows[o][22],
          gas:      oldRows[o][24],
          pay:      oldRows[o][26],
          expense:  isNewSumLayout ? (oldRows[o][27] || '') : '',
          memo:     isNewSumLayout ? oldRows[o][29] : oldRows[o][28],
          kari:     isNewSumLayout ? (oldRows[o][30] || '') : (oldRows[o][29] || ''),
          kyuryo:   isNewSumLayout ? (oldRows[o][31] || '') : (oldRows[o][30] || ''),
          pct:      isNewSumLayout ? (oldRows[o][32] || '') : (oldRows[o][31] || ''),
          other:    isNewSumLayout ? (oldRows[o][34] || '') : (oldRows[o][33] || '')
        };
      }
    }
  }
  if (!sumSheet) sumSheet = ss.insertSheet('集計表');

  var header = [
    'ID','区分','会社名','トン数','車種','車番','乗務員名','携帯番号','看板名',
    '日付','荷主','積地','降地','誘導時刻','積完時刻','休憩開始','休憩終了','降完時刻',
    '売上','請求(高速代)','実費(高速代)','合計(高速代)',
    '距離','燃費','ガソリン代','燃料代','支払い','経費合計','利益','備考',
    '仮日数','給料','％','有休手当','その他手当',
    '点呼前完了','点呼後完了'
  ];

  // 運行シートを全行読み込み、ID単位にデータを集約する
  // 同一IDに複数行ある場合（複数行程）は積地/降地を連結、金額は合算
  var unkouData = unkouSheet.getDataRange().getValues();
  // 運行シートのヘッダー行（row1）から点呼列インデックスを動的取得
  var unkouHdr0 = unkouData.length > 0 ? unkouData[0].map(function(h){ return String(h||'').trim(); }) : [];
  var genInspBCol = unkouHdr0.indexOf('点呼前完了');
  var genInspACol = unkouHdr0.indexOf('点呼後完了');
  var idMap = {}, idOrder = [];
  for (var i = 1; i < unkouData.length; i++) {
    var r  = unkouData[i];
    var id = String(r[0] || '').trim();
    if (!id) continue;
    if (!idMap[id]) {
      idMap[id] = {
        id:id, kubun:r[1], company:r[2], tons:r[3], type:r[4], car:r[5],
        name:r[6], tel:r[7], kanban:r[8], date:r[9], clients:[],
        picks:[], drops:[],
        guideTime:'',
        pickTime:'', restStart:'', restEnd:'', dropTime:'',
        rawPickTime:null, rawRestStart:null, rawRestEnd:null, rawDropTime:null,
        sales:0, tollReq:0, tollReal:0,
        inspBefore:'', inspAfter:''
      };
      idOrder.push(id);
    }
    var g = idMap[id];
    if (r[11]) g.picks.push(r[11]);
    if (r[12]) g.drops.push(r[12]);
    // 荷主は重複なしで全部収集（行程ごとに荷主が異なる場合も全て・区切りで表示）
    if (r[10]) { var rc = String(r[10]); if (g.clients.indexOf(rc) === -1) g.clients.push(rc); }
    // 時刻は先勝ち（最初に見つかった値を使用）
    if (r[13] && !g.guideTime) { g.guideTime = r[13]; }
    if (r[14] && !g.pickTime)  { g.pickTime  = r[14]; g.rawPickTime  = new Date(r[14]); }
    if (r[15] && !g.restStart) { g.restStart = r[15]; g.rawRestStart = new Date(r[15]); }
    if (r[16] && !g.restEnd)   { g.restEnd   = r[16]; g.rawRestEnd   = new Date(r[16]); }
    if (r[17] && !g.dropTime)  { g.dropTime  = r[17]; g.rawDropTime  = new Date(r[17]); }
    // 売上・高速は複数行程分を合算
    g.sales   += Number(r[18]) || 0;
    g.tollReq += Number(r[19]) || 0;
    g.tollReal+= Number(r[20]) || 0;
    // 点呼時刻は先勝ち
    if (genInspBCol >= 0 && !g.inspBefore && r[genInspBCol]) g.inspBefore = r[genInspBCol];
    if (genInspACol >= 0 && !g.inspAfter  && r[genInspACol]) g.inspAfter  = r[genInspACol];
  }

  // 車両・月ごとの実稼働日数をカウント（経費按分用：有休・休みを除く）
  var workDayMap = {};
  for (var wd = 0; wd < idOrder.length; wd++) {
    var gwd  = idMap[idOrder[wd]];
    var gpwd = gwd.picks.join('・'), gdwd = gwd.drops.join('・');
    var isHoli = gpwd.indexOf('有休') !== -1 || gdwd.indexOf('有休') !== -1 ||
                 gpwd.indexOf('休み') !== -1 || gdwd.indexOf('休み') !== -1;
    if (!isHoli && gwd.date instanceof Date) {
      var ymKey = gwd.date.getFullYear() + '_' + gwd.date.getMonth();
      var vKey  = String(gwd.car||'').trim() + '_' + String(gwd.name||'').trim() + '_' + ymKey;
      workDayMap[vKey] = (workDayMap[vKey] || 0) + 1;
    }
  }

  // 月締めロック：再計算範囲の制限日付を取得（設定がなければ全期間）
  var recalcFromStr_ = PropertiesService.getScriptProperties().getProperty('recalcFromDate');
  var recalcFrom_ = null;
  if (recalcFromStr_) {
    var rfp_ = recalcFromStr_.split('/');
    recalcFrom_ = new Date(Number(rfp_[0]), Number(rfp_[1])-1, Number(rfp_[2]));
    recalcFrom_.setHours(0, 0, 0, 0);
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
    var pkey       = String(g.car||'').trim() + '_' + String(g.name||'').trim();
    var hasMasterG = pkey in payCondMap;
    var pc         = payCondMap[pkey] || {kari:'', kyuryo:'', pct:'', tollDeduct:'', expense:0};
    // 月締めロック: 制限日付より前の行はマスタ更新を反映しない（旧集計表の値を保持）
    var rowDate_ = g.date instanceof Date ? g.date : null;
    var useOld_  = recalcFrom_ && rowDate_ && rowDate_ < recalcFrom_;
    // マスタに登録がある場合はマスタの値のみ使う（空は空のまま）
    // マスタに登録がない場合のみ旧集計表の手入力値を使う
    var kari   = (!useOld_ && hasMasterG) ? pc.kari   : (old.kari   || '');
    var kyuryo = (!useOld_ && hasMasterG) ? pc.kyuryo : (old.kyuryo || '');
    var pct    = (!useOld_ && hasMasterG) ? pc.pct    : (old.pct    || '');

    var gpick = g.picks.join('・'), gdrop = g.drops.join('・');
    var gIsYukyu  = gpick.indexOf('有休') !== -1 || gdrop.indexOf('有休') !== -1;
    var gIsYasumi = !gIsYukyu && (gpick.indexOf('休み') !== -1 || gdrop.indexOf('休み') !== -1);
    // 経費按分（有休・休みは0円、仮日数設定時はそれを月間予定稼働日数として使用）
    var gExpense = 0;
    if (!gIsYukyu && !gIsYasumi && g.date instanceof Date) {
      var scheduledDaysG = Number(kari) || 0;
      var pcExpG = pc.expense || 0;
      if (pcExpG > 0 && scheduledDaysG === 0) {
        gExpense = ''; // 仮日数なし・経費あり→空欄(グレー警告)
      } else if (scheduledDaysG > 0) {
        gExpense = Math.round(pcExpG / scheduledDaysG);
      }
    }
    // 月締めロック: 制限日付より前の行は旧集計表の経費値を維持
    if (useOld_ && old.expense !== undefined && old.expense !== '') {
      gExpense = old.expense;
    }
    // V/Z/支払/利益を値で計算（数式廃止でフィルター時のズレを防止）
    var vRow = (g.tollReq === 0 && g.tollReal === 0) ? '' : (Number(g.tollReal)||0)-(Number(g.tollReq)||0);
    var zRow = old.distance ? Math.round(Number(old.distance)/Number(fuel)*(Number(old.gas)||0)) : '';
    var pctNumG = Number(pct)||0, kyuryoNumG = Number(kyuryo)||0, kariNumG = Number(kari)||0;
    var thisTollG = (Number(g.tollReal)||0)-(Number(g.tollReq)||0);
    var payRow, yukyuRow = gIsYukyu ? yukyuRate : '';
    if (pctNumG > 0) {
      var deductTollG = (pc.tollDeduct === '○') ? thisTollG : 0;
      payRow = Math.round(((Number(g.sales)||0)-deductTollG)*pctNumG/100);
    } else if (kyuryoNumG > 0 && kariNumG > 0) {
      payRow = gIsYasumi ? -Math.round(kyuryoNumG/kariNumG) : Math.round(kyuryoNumG/kariNumG);
    } else {
      payRow = old.pay || '';
    }
    var vNG = typeof vRow==='number'?vRow:0, zNG = typeof zRow==='number'?zRow:0;
    var payNG = typeof payRow==='number'?payRow:(Number(old.pay)||0);
    var salesNG = Number(g.sales)||0, expNG = Number(gExpense)||0;
    var acRow = (!salesNG&&!vNG&&!zNG&&!payNG&&!expNG) ? '' : salesNG-(vNG+zNG+payNG+expNG);
    outRows.push([
      g.id, g.kubun, g.company, g.tons, g.type, g.car, g.name, g.tel, g.kanban||g.company,
      g.date, g.clients.join('・'), gpick, gdrop,
      g.guideTime||'', g.pickTime, g.restStart, g.restEnd, g.dropTime,
      g.sales||'', g.tollReq||'', g.tollReal||'', vRow,
      old.distance||'', fuel, old.gas||'', zRow,
      payRow, gExpense, acRow, old.memo||'',
      kari, kyuryo, pct,
      yukyuRow,
      old.other || '',
      g.inspBefore || '', g.inspAfter || ''
    ]);
  }

  // 集計表を再書き込み（値のみクリア・背景リセット、枠線・書式・ヘッダー色は保持）
  var prevLR_ = sumSheet.getLastRow();
  var prevLC_ = Math.max(sumSheet.getLastColumn(), 37);
  sumSheet.clearContents();
  if (prevLR_ >= 2) {
    sumSheet.getRange(2, 1, prevLR_ - 1, prevLC_).setBackground(null);
  }
  if (outRows.length > 0) {
    sumSheet.getRange(1, 1, outRows.length, 37).setValues(outRows);
    // 燃料代（Z=26列）に数式を設定（距離÷燃費×ガソリン代）
    if (outRows.length > 1) {
      var fuelFormulas = [];
      for (var fRow = 2; fRow <= outRows.length; fRow++) {
        fuelFormulas.push(['=IF(OR(W'+fRow+'="",Y'+fRow+'=""),"",ROUND(W'+fRow+'/X'+fRow+'*Y'+fRow+',0))']);
      }
      sumSheet.getRange(2, 26, outRows.length - 1, 1).setFormulas(fuelFormulas);
    }
    sumSheet.setFrozenRows(1);

    // 4時間超で黄色（労働時間過超）、30分未満で水色（休憩不足）の判定閾値
    var F = 4*60*60*1000;
    var T = 30*60*1000;

    for (var row = 2; row <= outRows.length; row++) {
      var g2      = idMap[idOrder[row-2]];
      var rowVN   = typeof outRows[row-1][21]==='number' ? outRows[row-1][21] : 0;
      var rowZN   = typeof outRows[row-1][25]==='number' ? outRows[row-1][25] : 0;
      var rowPayN = typeof outRows[row-1][26]==='number' ? outRows[row-1][26] : 0;
      var rowExpN = Number(outRows[row-1][27])||0;
      var calcProfit = (Number(g2.sales)||0)-(rowVN+rowZN+rowPayN+rowExpN);
      var rowRed = calcProfit < 0 ? '#ffebee' : null;
      sumSheet.getRange(row, 1, 1, 37).setBackground(rowRed);
      sumSheet.getRange(row, 15, 1, 4).setBackground(rowRed);
      if (g2.rawPickTime  && g2.rawRestStart && (g2.rawRestStart-g2.rawPickTime)  > F) { sumSheet.getRange(row,15,1,2).setBackground('#ffd600'); }
      if (g2.rawRestStart && g2.rawRestEnd   && (g2.rawRestEnd  -g2.rawRestStart) < T) { sumSheet.getRange(row,16,1,2).setBackground('#4fc3f7'); }
      if (g2.rawRestEnd   && g2.rawDropTime  && (g2.rawDropTime -g2.rawRestEnd)   > F) { sumSheet.getRange(row,17,1,2).setBackground('#ffd600'); }
      // 支払い条件不備の警告背景色（グレー）
      var pctNR = Number(outRows[row-1][32])||0, kyuRN = Number(outRows[row-1][31])||0, kariRN = Number(outRows[row-1][30])||0;
      sumSheet.getRange(row, 27).setBackground(rowRed);
      sumSheet.getRange(row, 28).setBackground(rowRed);
      sumSheet.getRange(row, 31, 1, 3).setBackground(rowRed);
      if (!pctNR && !(kyuRN > 0 && kariRN > 0)) {
        if (kyuRN > 0 || kariRN > 0) {
          if (!kyuRN)  sumSheet.getRange(row, 32).setBackground('#b0bec5');
          if (!kariRN) sumSheet.getRange(row, 31).setBackground('#b0bec5');
        } else if (!outRows[row-1][26]) {
          sumSheet.getRange(row, 27).setBackground('#b0bec5');
        }
      }
      // 経費合計（AB=28列）: 仮日数なし経費あり→値を空欄でグレー警告
      if (outRows[row-1][27] === '') {
        sumSheet.getRange(row, 28).setBackground('#b0bec5');
      }
    }
    applyMoneyFormat_(sumSheet, 2, outRows.length - 1, 'summary');
    applyDateTimeFormat_(sumSheet, 2, outRows.length - 1);
    sumSheet.getRange(2, 36, outRows.length - 1, 2).setNumberFormat('M/d HH:mm');
    applySumEditableBorders_(sumSheet, 2, outRows.length - 1);
  }
  // 旧データより行数が減った場合、余分な行の書式（枠線・色含む）をクリア
  if (prevLR_ > outRows.length) {
    sumSheet.getRange(outRows.length + 1, 1, prevLR_ - outRows.length, prevLC_).clearFormat();
  }

  convertLegacyAdminDataUrls_();
  applyHolidayRowColors_();

  // フィルターをデータ全列に再設定
  var existingFilter = sumSheet.getFilter();
  if (existingFilter) existingFilter.remove();
  if (sumSheet.getLastRow() >= 1) {
    sumSheet.getRange(1, 1, sumSheet.getLastRow(), sumSheet.getLastColumn()).createFilter();
  }
  // ヘッダー行（1行目）の枠線を確実にクリア（データ行の黄色枠が残らないように）
  sumSheet.getRange(1, 1, 1, Math.max(sumSheet.getLastColumn(), 37)).setBorder(false, false, false, false, false, false);
}


// ================================================================
//  4-1b: 管理側データURLをリッチテキストに一括変換（convertLegacyAdminDataUrls_）  【大B / 中4 / 小4-1b】
//  運行シートのW列(23)にプレーンURLが残っている行をリッチテキストに変換
// ================================================================
function convertLegacyAdminDataUrls_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('運行');
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastRow = sheet.getLastRow();
  var all = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    var val = String(all[i][23] || '').trim();
    if (val.match(/^https?:\/\//)) {
      setAdminDataRichText_(sheet, i + 1, val);
    } else if (val && !val.match(/^https?:\/\//)) {
      // Already rich text ("ファイル1" etc.) — ensure note is populated
      var cell = sheet.getRange(i + 1, 24);
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
//  4-2: 集計表をID単位で同期（syncSummaryForId_）  【大B / 中4 / 小4-2】
//  運行シートから対象IDのデータを集計し集計表の該当行を更新する
//  ・AB〜AD列（仮日数・給料・%）を保持＆マスタから引き当て
//  ・時刻色付け・利益マイナス赤を再適用
//  ・数式（T列・X列・Z列）を再セット
// ================================================================
function syncSummaryForId_(targetId, ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var unkouSheet = ss.getSheetByName('運行');
  if (!unkouSheet) return;
  var sumSheet = ss.getSheetByName('集計表');
  if (!sumSheet || sumSheet.getLastRow() < 1) { generateSummary(); return; }

  // 設定・マスタをCacheServiceで高速化（同一セッション内の連続同期で再読み込みを省く）
  var syncCache = CacheService.getScriptCache();
  var fuelMap = {}, yukyuRate = 0;
  try {
    var cachedCfg = syncCache.get('cfg_setting');
    if (cachedCfg) { var co = JSON.parse(cachedCfg); fuelMap = co.fm; yukyuRate = co.yr; }
  } catch(cErr) {}
  if (!Object.keys(fuelMap).length) {
    var settingSheet = ss.getSheetByName('設定');
    if (settingSheet && settingSheet.getLastRow() >= 2) {
      var sVals = settingSheet.getRange(2, 1, settingSheet.getLastRow()-1, 4).getValues();
      for (var s = 0; s < sVals.length; s++) {
        var skey = String(sVals[s][0]||'').trim();
        if (skey) fuelMap[skey] = sVals[s][1];
        if (String(sVals[s][2]||'').trim() === '有休') yukyuRate = Number(sVals[s][3]) || 0;
      }
    }
    try { syncCache.put('cfg_setting', JSON.stringify({fm:fuelMap,yr:yukyuRate}), 300); } catch(cErr2) {}
  }

  var payCondMap = {};
  try {
    var cachedMst = syncCache.get('cfg_master');
    if (cachedMst) payCondMap = JSON.parse(cachedMst);
  } catch(cErr3) {}
  if (!Object.keys(payCondMap).length) {
    var master = ss.getSheetByName('自車専属マスタ');
    if (master && master.getLastRow() >= 2) {
      var mReadColsS = Math.max(master.getLastColumn(), 31);
      var mAllS = master.getRange(1, 1, master.getLastRow(), mReadColsS).getValues();
      var mHdrRowS = mAllS[0];
      var mHasTollDeductS = String(mHdrRowS[16]||'').trim() === '高速を引く（引くは〇、引かないは空欄）';
      var mExpStartS = mHasTollDeductS ? 17 : 16;
      for (var m = 1; m < mAllS.length; m++) {
        var mcar  = String(mAllS[m][7]  || '').trim();
        var mname = String(mAllS[m][8]  || '').trim();
        var mExp  = 0;
        for (var ei = mExpStartS; ei <= mExpStartS + 14; ei++) mExp += Number(mAllS[m][ei]) || 0;
        payCondMap[mcar+'_'+mname] = {
          kari:       mAllS[m][13] || '',
          kyuryo:     mAllS[m][14] || '',
          pct:        mAllS[m][15] || '',
          tollDeduct: mHasTollDeductS ? String(mAllS[m][16] || '').trim() : '',
          expense:    mExp
        };
      }
    }
    try { syncCache.put('cfg_master', JSON.stringify(payCondMap), 60); } catch(cErr4) {}
  }

  var unkouData = unkouSheet.getDataRange().getValues();
  // ヘッダー行から点呼列インデックスを動的取得
  var uHdr0 = unkouData.length > 0 ? unkouData[0].map(function(h){ return String(h||'').trim(); }) : [];
  var sInspBCol = uHdr0.indexOf('点呼前完了');
  var sInspACol = uHdr0.indexOf('点呼後完了');
  var g = null;
  var matchingRows = [];
  var rawPickTime=null, rawRestStart=null, rawRestEnd=null, rawDropTime=null;
  var inspBeforeTime='', inspAfterTime='';
  for (var i = 1; i < unkouData.length; i++) {
    var r = unkouData[i];
    if (String(r[0]||'').trim() !== String(targetId).trim()) continue;
    matchingRows.push(i + 1);
    if (!g) { g = { id:String(r[0]).trim(), kubun:r[1], company:r[2], tons:r[3], type:r[4], car:r[5], name:r[6], tel:r[7], kanban:r[8], date:r[9], clients:[], picks:[], drops:[], guideTime:'', pickTime:'', restStart:'', restEnd:'', dropTime:'', sales:0, tollReq:0, tollReal:0 }; }
    if (r[11]) g.picks.push(r[11]);
    if (r[12]) g.drops.push(r[12]);
    // 荷主は重複なしで全部収集
    if (r[10]) { var rc = String(r[10]); if (g.clients.indexOf(rc) === -1) g.clients.push(rc); }
    if (r[13] && !g.guideTime) { g.guideTime = r[13]; }
    if (r[14] && !g.pickTime)  { g.pickTime  = r[14]; rawPickTime  = new Date(r[14]); }
    if (r[15] && !g.restStart) { g.restStart = r[15]; rawRestStart = new Date(r[15]); }
    if (r[16] && !g.restEnd)   { g.restEnd   = r[16]; rawRestEnd   = new Date(r[16]); }
    if (r[17] && !g.dropTime)  { g.dropTime  = r[17]; rawDropTime  = new Date(r[17]); }
    g.sales   += Number(r[18]) || 0;
    g.tollReq += Number(r[19]) || 0;
    g.tollReal+= Number(r[20]) || 0;
    if (sInspBCol >= 0 && !inspBeforeTime && r[sInspBCol]) inspBeforeTime = r[sInspBCol];
    if (sInspACol >= 0 && !inspAfterTime  && r[sInspACol]) inspAfterTime  = r[sInspACol];
  }
  // 運行シートの該当行に書式を確実に適用
  if (matchingRows.length > 0) {
    var minR = matchingRows[0], maxR = matchingRows[matchingRows.length - 1];
    unkouSheet.getRange(minR, 10, maxR - minR + 1, 1).setNumberFormat('yyyy/MM/dd');
    applyDateTimeFormat_(unkouSheet, minR, maxR - minR + 1);
  }

  var sumLast = sumSheet.getLastRow();
  var sumRow  = 0;
  var keepDistance='', keepGas='', keepPay='', keepMemo='';
  var keepKari='', keepKyuryo='', keepPct='', keepOther='';
  if (sumLast >= 2) {
    var colCount = sumSheet.getLastColumn();
    var isNewSumLayout = colCount >= 35; // 経費合計列（AB=28）が追加済みか
    var sumIds   = sumSheet.getRange(2, 1, sumLast-1, Math.max(colCount, 37)).getValues();
    for (var k = 0; k < sumIds.length; k++) {
      if (String(sumIds[k][0]).trim() === String(targetId).trim()) {
        sumRow      = k + 2;
        keepDistance= sumIds[k][22];
        keepGas     = sumIds[k][24];
        keepPay     = sumIds[k][26];
        keepMemo    = isNewSumLayout ? sumIds[k][29] : sumIds[k][28];
        keepKari    = isNewSumLayout ? (sumIds[k][30] || '') : (sumIds[k][29] || '');
        keepKyuryo  = isNewSumLayout ? (sumIds[k][31] || '') : (sumIds[k][30] || '');
        keepPct     = isNewSumLayout ? (sumIds[k][32] || '') : (sumIds[k][31] || '');
        keepOther   = isNewSumLayout ? (sumIds[k][34] || '') : (sumIds[k][33] || '');
        break;
      }
    }
  }
  if (!g) { if (sumRow > 0) sumSheet.deleteRow(sumRow); return; }

  var tonsStr = String(g.tons||'').trim();
  var fuel    = fuelMap[tonsStr] || fuelMap[tonsStr.replace(/[tT]/,'')+'t'] || 3;

  var pkey      = String(g.car||'').trim()+'_'+String(g.name||'').trim();
  var hasMaster = pkey in payCondMap;
  var pc        = payCondMap[pkey] || {kari:'', kyuryo:'', pct:'', tollDeduct:'', expense: 0};
  var kari   = hasMaster ? pc.kari   : keepKari;
  var kyuryo = hasMaster ? pc.kyuryo : keepKyuryo;
  var pct    = hasMaster ? pc.pct    : keepPct;

  var spick = g.picks.join('・'), sdrop = g.drops.join('・');
  var sIsYukyu  = spick.indexOf('有休') !== -1 || sdrop.indexOf('有休') !== -1;
  var sIsYasumi = !sIsYukyu && (spick.indexOf('休み') !== -1 || sdrop.indexOf('休み') !== -1);

  // 経費按分計算（有休・休みは0円、実稼働日数で按分）
  var gCar  = String(g.car||'').trim(), gName = String(g.name||'').trim();
  var gDate = g.date instanceof Date ? g.date : new Date(g.date);
  var gYear = gDate.getFullYear(), gMon = gDate.getMonth();
  var uniqueWIds = {};
  for (var ui = 1; ui < unkouData.length; ui++) {
    var ur = unkouData[ui];
    if (String(ur[5]||'').trim() !== gCar || String(ur[6]||'').trim() !== gName) continue;
    var ud = ur[9]; if (!(ud instanceof Date)) continue;
    if (ud.getFullYear() !== gYear || ud.getMonth() !== gMon) continue;
    var upick = String(ur[11]||'').trim();
    if (upick === '' || upick.indexOf('有休') !== -1 || upick.indexOf('休み') !== -1) continue;
    uniqueWIds[String(ur[0]||'').trim()] = true;
  }
  var scheduledDaysS = Number(kari) || 0;
  var expWorkDays = scheduledDaysS > 0 ? scheduledDaysS : Math.max(Object.keys(uniqueWIds).length, 1);
  var pcExp       = (payCondMap[pkey] || {}).expense || 0;
  var expenseVal  = (sIsYukyu || sIsYasumi) ? 0 : (pcExp > 0 && scheduledDaysS === 0 ? '' : Math.round(pcExp / expWorkDays));

  var rowData = [
    g.id, g.kubun, g.company, g.tons, g.type, g.car, g.name, g.tel, g.kanban||g.company,
    g.date, g.clients.join('・'), spick, sdrop,
    g.guideTime||'', g.pickTime, g.restStart, g.restEnd, g.dropTime,
    g.sales||'', g.tollReq||'', g.tollReal||'', '',
    keepDistance, fuel, keepGas, '', keepPay, expenseVal, '', keepMemo,  // AA=支払い, AB=経費合計, AC=利益空, AD=備考
    kari, kyuryo, pct,
    sIsYukyu ? yukyuRate : '',
    keepOther,   // AI=その他手当（手入力保持）
    inspBeforeTime, inspAfterTime  // AJ=点呼前完了, AK=点呼後完了
  ];

  // LockServiceで並行実行による集計表重複挿入を防止
  var sumInsLock = LockService.getScriptLock();
  try { sumInsLock.waitLock(15000); } catch(e) {}
  try {
    // ロック取得後に再度同IDが存在しないか確認（並行実行で先に挿入された場合の対策）
    if (sumRow === 0 && sumSheet.getLastRow() >= 2) {
      var recheckIds = sumSheet.getRange(2, 1, sumSheet.getLastRow()-1, 1).getValues();
      for (var rc = 0; rc < recheckIds.length; rc++) {
        if (String(recheckIds[rc][0]).trim() === String(targetId).trim()) {
          sumRow = rc + 2;
          break;
        }
      }
    }
    if (sumRow > 0) {
      sumSheet.getRange(sumRow, 1, 1, 37).setValues([rowData]);
    } else {
      sumRow = sumSheet.getLastRow()+1;
      if (sumRow === 1) {
        var hdr = ['ID','区分','会社名','トン数','車種','車番','乗務員名','携帯番号','看板名','日付','荷主','積地','降地','誘導時刻','積完時刻','休憩開始','休憩終了','降完時刻','売上','請求(高速代)','実費(高速代)','合計(高速代)','距離','燃費','ガソリン代','燃料代','支払い','経費合計','利益','備考','仮日数','給料','％','有休手当','その他手当','点呼前完了','点呼後完了'];
        sumSheet.getRange(1, 1, 1, 37).setValues([hdr]);
        sumSheet.setFrozenRows(1);
        sumRow = 2;
      }
      sumSheet.getRange(sumRow, 1, 1, 37).setValues([rowData]);
    }
  } finally {
    sumInsLock.releaseLock();
  }

  var vSyncVal = (g.tollReq === 0 && g.tollReal === 0) ? '' : (Number(g.tollReal)||0)-(Number(g.tollReq)||0);
  sumSheet.getRange(sumRow, 22).setValue(vSyncVal);
  // 燃料代（Z=26列）に数式を設定（距離÷燃費×ガソリン代）
  sumSheet.getRange(sumRow, 26).setFormula('=IF(OR(W'+sumRow+'="",Y'+sumRow+'=""),"",ROUND(W'+sumRow+'/X'+sumRow+'*Y'+sumRow+',0))');

  applyMoneyFormat_(sumSheet, sumRow, 1, 'summary');
  applyDateTimeFormat_(sumSheet, sumRow, 1);
  sumSheet.getRange(sumRow, 36, 1, 2).setNumberFormat('M/d HH:mm');
  // この行だけの支払い(AA=col27)をインライン計算
  var pctNum    = Number(pct)    || 0;
  var kyuryoNum = Number(kyuryo) || 0;
  var kariNum   = Number(kari)   || 0;
  var thisToll  = (Number(g.tollReal) || 0) - (Number(g.tollReq) || 0);
  var payCell   = sumSheet.getRange(sumRow, 27);
  sumSheet.getRange(sumRow, 31, 1, 3).setBackground(null);
  payCell.setBackground(null);
  var yukyuVal = '';
  var finalPaySync = null;
  if (pctNum > 0) {
    var tollDeductS = (pc.tollDeduct === '○') ? thisToll : 0;
    finalPaySync = Math.round(((Number(g.sales) || 0) - tollDeductS) * pctNum / 100);
    payCell.setValue(finalPaySync);
    if (sIsYukyu) yukyuVal = yukyuRate;
  } else if (kyuryoNum > 0 && kariNum > 0) {
    var dailyPay = Math.round(kyuryoNum / kariNum);
    finalPaySync = sIsYasumi ? -dailyPay : dailyPay;
    payCell.setValue(finalPaySync);
  } else if (kyuryoNum > 0 || kariNum > 0) {
    // 警告は行背景設定後に適用
  } else {
    // 警告は行背景設定後に適用
  }
  // AC(29)=利益 を値で書き込む（支払い確定後に計算）
  var vSN = typeof vSyncVal==='number' ? vSyncVal : 0;
  var zSN = (Number(keepDistance) > 0 && Number(fuel) > 0 && Number(keepGas) > 0)
    ? Math.round(Number(keepDistance) / Number(fuel) * Number(keepGas)) : 0;
  var resolvedPaySync = finalPaySync !== null ? finalPaySync : (Number(keepPay)||0);
  var salesSync = Number(g.sales)||0;
  var acSyncVal = (!salesSync&&!vSN&&!zSN&&!resolvedPaySync&&!expenseVal) ? '' : salesSync-(vSN+zSN+resolvedPaySync+(Number(expenseVal)||0));
  sumSheet.getRange(sumRow, 29).setValue(acSyncVal);
  // 利益マイナス → 薄赤（有休/休み行はこの後上書き）
  sumSheet.getRange(sumRow, 1, 1, 37).setBackground(typeof acSyncVal==='number' && acSyncVal < 0 ? '#ffebee' : null);
  // 有休手当(AH=col34)
  sumSheet.getRange(sumRow, 34).setValue(yukyuVal);
  // 有休/休み → 行全体をグレー着色
  var sumLastColBg = Math.max(sumSheet.getLastColumn(), 37);
  if (spick.indexOf('有休') !== -1 || sdrop.indexOf('有休') !== -1) {
    sumSheet.getRange(sumRow, 1, 1, sumLastColBg).setBackground('#e0e0e0');
  } else if (spick.indexOf('休み') !== -1 || sdrop.indexOf('休み') !== -1) {
    sumSheet.getRange(sumRow, 1, 1, sumLastColBg).setBackground('#9e9e9e');
  }
  // 430ルール色付け（行全体の背景色設定後に適用して上書きを防ぐ）
  var F430 = 4*60*60*1000, T430 = 30*60*1000;
  var baseRowBg = (typeof acSyncVal === 'number' && acSyncVal < 0) ? '#ffebee' : null;
  if (spick.indexOf('有休') !== -1 || sdrop.indexOf('有休') !== -1) baseRowBg = '#e0e0e0';
  else if (spick.indexOf('休み') !== -1 || sdrop.indexOf('休み') !== -1) baseRowBg = '#9e9e9e';
  sumSheet.getRange(sumRow, 15, 1, 4).setBackground(baseRowBg);
  if (rawPickTime  && rawRestStart && (rawRestStart-rawPickTime)  > F430) { sumSheet.getRange(sumRow,15,1,2).setBackground('#ffd600'); }
  if (rawRestStart && rawRestEnd   && (rawRestEnd  -rawRestStart) < T430) { sumSheet.getRange(sumRow,16,1,2).setBackground('#4fc3f7'); }
  if (rawRestEnd   && rawDropTime  && (rawDropTime -rawRestEnd)   > F430) { sumSheet.getRange(sumRow,17,1,2).setBackground('#ffd600'); }
  // 順序エラー着色（後ろが埋まって前が空の時刻セルをオレンジ表示）
  var hasGuide = !!g.guideTime;
  if (!hasGuide    && (rawPickTime||rawRestStart||rawRestEnd||rawDropTime)) sumSheet.getRange(sumRow,14).setBackground('#ff6d00');
  if (!rawPickTime && (rawRestStart||rawRestEnd||rawDropTime))              sumSheet.getRange(sumRow,15).setBackground('#ff6d00');
  if (!rawRestStart && (rawRestEnd||rawDropTime))                           sumSheet.getRange(sumRow,16).setBackground('#ff6d00');
  if (!rawRestEnd  && rawDropTime)                                          sumSheet.getRange(sumRow,17).setBackground('#ff6d00');
  // 支払い/経費条件不備の警告（行全体の背景色設定の後に適用、グレー）
  if (!sIsYukyu && !sIsYasumi) {
    if (pctNum <= 0 && !(kyuryoNum > 0 && kariNum > 0)) {
      if (kyuryoNum > 0 || kariNum > 0) {
        if (!kyuryoNum) sumSheet.getRange(sumRow, 32).setBackground('#b0bec5');
        if (!kariNum)   sumSheet.getRange(sumRow, 31).setBackground('#b0bec5');
      } else if (!keepPay) {
        sumSheet.getRange(sumRow, 27).setBackground('#b0bec5');
      }
    }
    if (expenseVal === '') {
      sumSheet.getRange(sumRow, 28).setBackground('#b0bec5');
    }
  }

  applySumEditableBorders_(sumSheet, sumRow, 1);

  // フィルターをデータ全列に再設定（列追加時も自動対応）
  var sfExisting = sumSheet.getFilter();
  if (!sfExisting || sumSheet.getLastColumn() > sfExisting.getRange().getLastColumn()) {
    if (sfExisting) sfExisting.remove();
    sumSheet.getRange(1, 1, sumSheet.getLastRow(), sumSheet.getLastColumn()).createFilter();
  }

  // 同じIDの重複行が残っていれば削除（上から2行目以降を後ろから走査して消す）
  var dupLast = sumSheet.getLastRow();
  if (dupLast >= 3) {
    var dupIds = sumSheet.getRange(2, 1, dupLast - 1, 1).getValues();
    for (var di = dupIds.length - 1; di >= 0; di--) {
      if (String(dupIds[di][0]).trim() === String(targetId).trim() && (di + 2) !== sumRow) {
        sumSheet.deleteRow(di + 2);
      }
    }
  }
}


// ================================================================
//  設定シートに業務前点検・業務後点検データがなければデフォルトを挿入
// ================================================================
function ensureSettingItems_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingSheet = ss.getSheetByName('設定');
  if (!settingSheet) return;
  var lastCol = settingSheet.getLastColumn();
  if (lastCol === 0) return;
  var headers = settingSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var bCol = -1, aCol = -1;
  for (var hi = 0; hi < headers.length; hi++) {
    var h = String(headers[hi] || '').trim();
    if (h === '業務前点検') bCol = hi;
    if (h === '業務後点検') aCol = hi;
  }
  if (bCol === -1) { bCol = lastCol; settingSheet.getRange(1, bCol + 1).setValue('業務前点検'); lastCol++; }
  if (aCol === -1) { aCol = lastCol; settingSheet.getRange(1, aCol + 1).setValue('業務後点検'); }
  var sLastRow = settingSheet.getLastRow();
  var existB = sLastRow >= 2 ? settingSheet.getRange(2, bCol + 1, sLastRow - 1, 1).getValues().map(function(r){ return String(r[0]||'').trim(); }).filter(function(v){ return v; }) : [];

  if (existB.length === 0) {
    var defaultBefore = [
      'ブレーキの効き・踏みしろ（エア・液漏れ含む）',
      'タイヤの空気圧・溝の深さ・亀裂や損傷',
      'ホイールナットの緩み・脱落の確認',
      'エンジンオイル・冷却水・ベルト類の確認',
      'バッテリー液・ウォッシャー液の確認',
      '灯火類（ランプ・ウィンカー）の点灯・汚れ',
      'ワイパーの動作・払拭状態',
      'エンジンのかかり具合・異音の確認',
      'ミラーの調整・シートベルトの装着',
      '乗務前点呼・アルコールチェックの実施'
    ];
    for (var di = 0; di < defaultBefore.length; di++) {
      settingSheet.getRange(di + 2, bCol + 1).setValue(defaultBefore[di]);
    }
  }
  var existA = sLastRow >= 2 ? settingSheet.getRange(2, aCol + 1, sLastRow - 1, 1).getValues().map(function(r){ return String(r[0]||'').trim(); }).filter(function(v){ return v; }) : [];

  if (existA.length === 0) {
    var defaultAfter = [
      '車両・積載物の異常の有無（タイヤ・車体等）',
      '事故・ヒヤリハットの有無',
      '道路状況・運行状況の異常の有無',
      '翌乗務員への引き継ぎ事項の有無',
      '運行記録（日報）の提出・乗務後点呼の実施'
    ];
    for (var dj = 0; dj < defaultAfter.length; dj++) {
      settingSheet.getRange(dj + 2, aCol + 1).setValue(defaultAfter[dj]);
    }
  }
}

// ================================================================
//  4-2b: 運行シートのID・車番一括補完（fillMissingIdsAndCars）
//  行程を一括追加したあとにIDや車番が空のまま残っている行を一括補完する
//  ・A列が空でB〜K列にデータがあれば V-XXXX 形式のIDを自動採番
//  ・F(車番)か G(乗務員名) の片方が空の場合、自車専属マスタから補完
//  ・集計表を一括再同期して反映
// ================================================================
function fillMissingIdsAndCars() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('運行シートにデータがありません');
    return;
  }
  var master = ss.getSheetByName('自車専属マスタ');
  var mAll = master ? master.getDataRange().getValues() : [];

  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();

  // ① ID一括採番（ロック内でメモリ更新→A列を一括書き込み）
  var fillLock = LockService.getScriptLock();
  try { fillLock.waitLock(15000); } catch(e) {
    SpreadsheetApp.getUi().alert('ロック取得失敗。しばらく後に再試行してください。');
    return;
  }
  var nextIdNum = getNextIdNum_(sheet, 'V-');
  var idCount = 0;
  for (var fi = 0; fi < data.length; fi++) {
    if (String(data[fi][0] || '').trim()) continue;
    var fhd = false;
    for (var fc = 1; fc <= 10; fc++) {
      if (String(data[fi][fc] || '').trim() !== '') { fhd = true; break; }
    }
    if (!fhd) continue;
    data[fi][0] = 'V-' + String(nextIdNum).padStart(4, '0');
    nextIdNum++;
    idCount++;
  }
  // A列を一括書き込み（ロック保持中）
  if (idCount > 0) {
    sheet.getRange(2, 1, data.length, 1).setValues(data.map(function(r) { return [r[0]]; }));
    SpreadsheetApp.flush();
  }
  fillLock.releaseLock();

  // ② 車番・各フィールド補完（メモリ更新→12列一括書き込み）
  var carCount = 0, syncIds = [];
  for (var i = 0; i < data.length; i++) {
    var id  = String(data[i][0] || '').trim();
    var hasData = false;
    for (var c = 1; c <= 10; c++) {
      if (String(data[i][c] || '').trim() !== '') { hasData = true; break; }
    }
    if (!hasData) continue;

    // F列(車番)かG列(乗務員名)が空の場合、マスタから補完
    var carNo = String(data[i][5] || '').trim();
    var name  = String(data[i][6] || '').trim();
    if ((!carNo || !name) && mAll.length > 1) {
      for (var m = 1; m < mAll.length; m++) {
        var mStatus = String(mAll[m][1] || '').trim();
        if (mStatus === '故障' || mStatus === '待機') continue;
        var mCar  = String(mAll[m][7] || '').trim();
        var mName = String(mAll[m][8] || '').trim();
        var matchByCar  = carNo && mCar === carNo;
        var matchByName = name  && mName === name;
        if (matchByCar || matchByName) {
          if (!carNo) { data[i][5] = mCar;  carCount++; }
          if (!name)  { data[i][6] = mName; carCount++; }
          // 区分/会社名/トン数/車種/携帯/看板名をメモリ上で補完
          var fields = [[2,2],[3,3],[5,4],[6,5],[9,8],[4,9]]; // [masterIdx, sheetCol]
          for (var fj = 0; fj < fields.length; fj++) {
            var mi = fields[fj][0], sc = fields[fj][1];
            if (!String(data[i][sc-1]||'').trim() && String(mAll[m][mi]||'').trim()) {
              data[i][sc-1] = mAll[m][mi];
            }
          }
          break;
        }
      }
    }
    if (id) syncIds.push(id);
  }
  // 12列を一括書き込み
  if (carCount > 0) {
    sheet.getRange(2, 1, data.length, 12).setValues(data);
  }

  // ID重複排除して集計表を再同期
  var uniq = syncIds.filter(function(v,i,a){return a.indexOf(v)===i;});
  for (var si = 0; si < uniq.length; si++) {
    try { syncSummaryForId_(uniq[si]); } catch(e) {}
  }

  sortUnkouByDate_();
  applyHolidayRowColors_();

  SpreadsheetApp.getUi().alert(
    '✅ 完了\n' +
    'ID補完: ' + idCount + '行\n' +
    '車番/乗務員補完: ' + carCount + '箇所\n' +
    '集計表同期: ' + uniq.length + '件'
  );
}


// ================================================================
//  4-2c: 日時入力ダイアログ（showDateTimePicker / setDateTimeToActiveCell_）
//  選択中セルに日時（年月日 + 時刻）を入力するカスタムダイアログ
//  メニュー「📅 日時入力」から起動
// ================================================================
function showDateTimePicker() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell  = sheet.getActiveCell();
  var val   = cell.getValue();
  var cur   = (val instanceof Date)
    ? Utilities.formatDate(val, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm")
    : '';
  var colName = sheet.getRange(1, cell.getColumn()).getValue() || ('列' + cell.getColumn());
  var cellAddr = cell.getA1Notation();

  var html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head>' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font-family:sans-serif;padding:12px;background:#1e1e1e;color:#eee;margin:0}' +
    'label{display:block;margin-bottom:4px;font-size:12px;color:#aaa}' +
    'input{width:100%;box-sizing:border-box;padding:8px;font-size:15px;border-radius:6px;border:1px solid #555;background:#333;color:#fff;margin-bottom:12px}' +
    'button{padding:8px 18px;font-size:14px;border:none;border-radius:6px;cursor:pointer;margin-right:8px}' +
    '.ok{background:#1976d2;color:#fff}.cl{background:#555;color:#fff}' +
    '</style></head><body>' +
    '<label>📍 セル: ' + cellAddr + ' （' + colName + '）</label>' +
    '<label>日時を選択してください</label>' +
    '<input type="datetime-local" id="dt" value="' + cur + '">' +
    '<div><button class="ok" onclick="ok()">✅ セット</button>' +
    '<button class="cl" onclick="google.script.host.close()">キャンセル</button></div>' +
    '<script>function ok(){' +
    'var v=document.getElementById("dt").value;' +
    'if(!v){alert("日時を選択してください");return;}' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
    '.withFailureHandler(function(e){alert(e.message);})' +
    '.setDateTimeToActiveCell_(v);}' +
    '<\/script></body></html>'
  ).setWidth(340).setHeight(175);

  SpreadsheetApp.getUi().showModalDialog(html, '📅 日時入力');
}

function setDateTimeToActiveCell_(isoStr) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cell  = sheet.getActiveCell();
  var d = new Date(isoStr);
  if (isNaN(d.getTime())) return;
  cell.setValue(d);
  cell.setNumberFormat('M/d HH:mm');
}


// ================================================================
//  4-3: シート再生成（expandAndRefreshSheets）  【大B / 中4 / 小4-3】
//  メニューの「シート再生成」から呼び出す
//  ・自車専属マスタに仮日数/給料/%列がなければ追加
//  ・自車専属マスタに15経費列がなければ追加（按分計算用）
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

  // 自車専属マスタに15経費列を追加（なければ）
  var masterExpCols = [
    '車両リース代', '任意保険料', '自賠責保険料', '重量税積立', '車検費積立',
    '整備費積立', 'タイヤ代積立', '修理積立', '駐車場代', 'ETCリース料',
    'カーナビリース料', '通信費', '洗車費', '制服費', 'その他固定費'
  ];
  var masterSheet = ss.getSheetByName('自車専属マスタ');
  if (masterSheet) {
    var mLastCol  = masterSheet.getLastColumn();
    var mHeaders  = mLastCol > 0 ? masterSheet.getRange(1, 1, 1, mLastCol).getValues()[0] : [];
    var mNextCol  = mLastCol + 1;
    for (var ej = 0; ej < masterExpCols.length; ej++) {
      if (mHeaders.indexOf(masterExpCols[ej]) === -1) {
        masterSheet.getRange(1, mNextCol).setValue(masterExpCols[ej]);
        mNextCol++;
      }
    }
  }

  // 自車専属マスタの「高速を引く」列を新名称に統一（リネーム優先・重複削除・なければ新規追加）
  if (masterSheet) {
    var mTollLastCol = masterSheet.getLastColumn();
    var mTollHeaders = mTollLastCol > 0 ? masterSheet.getRange(1, 1, 1, mTollLastCol).getValues()[0] : [];
    var tollColName  = '高速を引く（引くは〇、引かないは空欄）';
    var tollOldName  = '高速を引く（空欄時は引かない）';
    var newIdx = mTollHeaders.indexOf(tollColName);  // 0-based
    var oldIdx = mTollHeaders.indexOf(tollOldName);  // 0-based

    if (newIdx >= 0 && oldIdx >= 0) {
      // 新旧両方ある（前回の誤追加）→ 余分な列（後のほう）を削除して名称統一
      var delCol = Math.max(newIdx, oldIdx) + 1; // 1-based
      masterSheet.deleteColumn(delCol);
      // 残った方が旧名称なら新名称にリネーム
      var keepIdx = Math.min(newIdx, oldIdx);
      var keepHdr = String(mTollHeaders[keepIdx] || '').trim();
      if (keepHdr !== tollColName) masterSheet.getRange(1, keepIdx + 1).setValue(tollColName);
    } else if (oldIdx >= 0) {
      // 旧名称だけある → リネームのみ（列追加しない）
      masterSheet.getRange(1, oldIdx + 1).setValue(tollColName);
    } else if (newIdx === -1) {
      // どちらもない → 新規追加（%列の直後・経費列の前）
      var mTollHdrs2 = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
      var firstExpColPos = mTollHdrs2.indexOf('車両リース代');
      if (firstExpColPos >= 0) {
        masterSheet.insertColumnBefore(firstExpColPos + 1);
        masterSheet.getRange(1, firstExpColPos + 1).setValue(tollColName);
      } else {
        masterSheet.getRange(1, masterSheet.getLastColumn() + 1).setValue(tollColName);
      }
    }
    // newIdx >= 0 かつ oldIdx === -1 → 既に正しい名称で存在、何もしない
  }

  // 運行シートに点呼前完了・点呼後完了列を追加（なければ末尾に追加）―早期に実行してタイムアウトを回避
  var unkouInspSheet = ss.getSheetByName('運行');
  if (unkouInspSheet) {
    var uLastCol = unkouInspSheet.getLastColumn();
    var uRawHdrs = uLastCol > 0 ? unkouInspSheet.getRange(1, 1, 1, uLastCol).getValues()[0] : [];
    var uHdrs = uRawHdrs.map(function(h){ return String(h||'').trim(); });
    var uNext = uLastCol + 1;
    if (uHdrs.indexOf('点呼前完了') === -1) {
      unkouInspSheet.getRange(1, uNext).setValue('点呼前完了');
      unkouInspSheet.getRange(1, uNext).setBackground('#37474f').setFontColor('#90a4ae').setFontWeight('bold');
      uNext++;
    }
    if (uHdrs.indexOf('点呼後完了') === -1) {
      unkouInspSheet.getRange(1, uNext).setValue('点呼後完了');
      unkouInspSheet.getRange(1, uNext).setBackground('#37474f').setFontColor('#90a4ae').setFontWeight('bold');
      uNext++;
    }
    // 点呼後完了の右に帳票関連列を追加（なければ）
    // ※ uNext は最新の「次に追加すべき列」を指している
    uLastCol = unkouInspSheet.getLastColumn();
    uRawHdrs = uLastCol > 0 ? unkouInspSheet.getRange(1, 1, 1, uLastCol).getValues()[0] : [];
    uHdrs    = uRawHdrs.map(function(h){ return String(h||'').trim(); });
    uNext    = uLastCol + 1;
    var docCols = [
      {name:'装備その他',  bg:'#4527a0', fg:'#ede7f6'},
      {name:'発注書・指示書', bg:'#1a237e', fg:'#e8eaf6'},
      {name:'車番連絡',    bg:'#1a237e', fg:'#e8eaf6'},
      {name:'帳票備考',    bg:'#4e342e', fg:'#efebe9'}
    ];
    docCols.forEach(function(col){
      if (uHdrs.indexOf(col.name) === -1) {
        var c = unkouInspSheet.getRange(1, uNext);
        c.setValue(col.name).setBackground(col.bg).setFontColor(col.fg).setFontWeight('bold');
        uNext++;
      }
    });
  }

  // 集計表に点呼前完了・点呼後完了列を追加（なければ末尾に追加）
  var sumInspSheet = ss.getSheetByName('集計表');
  if (sumInspSheet && sumInspSheet.getLastRow() >= 1) {
    var siLastCol = sumInspSheet.getLastColumn();
    var siRawHdrs = siLastCol > 0 ? sumInspSheet.getRange(1, 1, 1, siLastCol).getValues()[0] : [];
    var siHdrs = siRawHdrs.map(function(h){ return String(h||'').trim(); });
    var siNext = siLastCol + 1;
    if (siHdrs.indexOf('点呼前完了') === -1) {
      sumInspSheet.getRange(1, siNext).setValue('点呼前完了');
      siNext++;
    }
    if (siHdrs.indexOf('点呼後完了') === -1) {
      sumInspSheet.getRange(1, siNext).setValue('点呼後完了');
    }
  }

  // 取引先マスタにメールアドレス列を追加（なければ末尾に自動追加）
  var custSheetExp = ss.getSheetByName('マスタ');
  if (custSheetExp) {
    var cLastColExp = custSheetExp.getLastColumn();
    var cHdrsExp = cLastColExp > 0 ? custSheetExp.getRange(1,1,1,cLastColExp).getValues()[0].map(function(h){return String(h||'').trim();}) : [];
    if (cHdrsExp.indexOf('メールアドレス') === -1) {
      custSheetExp.getRange(1, cLastColExp + 1).setValue('メールアドレス').setFontWeight('bold');
    }
  }

  refreshActiveVehiclesAuto_();

  var unkouForFmt = ss.getSheetByName('運行');
  if (unkouForFmt && unkouForFmt.getLastRow() >= 2) {
    applyMoneyFormat_(unkouForFmt, 2, unkouForFmt.getLastRow() - 1, 'unkou');
    applyDateTimeFormat_(unkouForFmt, 2, unkouForFmt.getLastRow() - 1);
    unkouForFmt.getRange(2, 10, unkouForFmt.getLastRow() - 1, 1).setNumberFormat('yyyy/MM/dd');
  }
  var sumForFmt = ss.getSheetByName('集計表');
  if (sumForFmt && sumForFmt.getLastRow() >= 2) {
    applyMoneyFormat_(sumForFmt, 2, sumForFmt.getLastRow() - 1, 'summary');
    applyDateTimeFormat_(sumForFmt, 2, sumForFmt.getLastRow() - 1);
  }

  applySheetColors_();

  // 設定シートに業務前点検・業務後点検データがなければデフォルトを挿入
  ensureSettingItems_(ss);

  // 自車専属マスタ B列（運行状態）に 運行・故障・待機 のドロップダウンを設定
  if (masterSheet && masterSheet.getMaxRows() > 1) {
    var dropRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['運行','故障','待機'], true)
      .setAllowInvalid(false).build();
    masterSheet.getRange(2, 2, masterSheet.getMaxRows() - 1, 1).setDataValidation(dropRule);
  }

  ensureCompanySettingSheet_(ss);

  // ── 情報（マッチング）シートの生成・整備 ────────────────────────────
  // 貨物情報と車両情報を1画面で管理し、チェックボックスで配車を確定するシート
  // M&A向け補足: このシートは「配車マン不要化」の核心機能。
  //              荷主からの依頼（貨物情報）と手配した車両（車両情報）を
  //              同一シートで管理し、チェック→確定で運行シートへ自動登録できる。
  (function() {
    // ── 列構成（A〜Z の26列） ─────────────────────────────────────────
    // [貨物側] A:チェック(貨物) B:進捗(貨物) C:会社名 D:TEL E:FAX F:日付 G:品名
    //          H:トン数 I:車種 J:積地 K:降地 L:金額(売上) M:備考
    // [車両側] N:チェック(車両) O:進捗(車両) P:会社名 Q:TEL R:FAX S:看板名
    //          T:トン数 U:車種 V:車番 W:乗務員名 X:携帯 Y:金額(支払) Z:備考
    //
    // 貨物と車両がそれぞれ独立したチェック列・進捗列を持つ設計。
    // 貨物の進捗(B列)を変えても車両側(N-Z)の色は変わらず、逆も同様。
    // 「3行目貨物チェック＋5行目車両チェック」で異なる行同士のマッチングが可能。
    var JOHO_COLS = 26;
    var johoHdr = [
      'チェック(貨物)', '進捗(貨物)',
      '会社名(貨物)', 'TEL(貨物)', 'FAX(貨物)', '日付', '品名',
      'トン数', '車種', '積地', '降地', '金額(売上)', '備考(貨物)',
      'チェック(車両)', '進捗(車両)',
      '会社名(車両)', 'TEL(車両)', 'FAX(車両)', '看板名',
      'トン数(車両)', '車種(車両)', '車番', '乗務員名', '携帯番号', '金額(支払)', '備考(車両)'
    ];
    var johoSheet = ss.getSheetByName('情報');
    if (!johoSheet) johoSheet = ss.insertSheet('情報');

    // 列数が足りない場合は補完（シートの古いバージョンも安全に更新）
    if (johoSheet.getMaxColumns() < JOHO_COLS) {
      johoSheet.insertColumnsAfter(johoSheet.getMaxColumns(), JOHO_COLS - johoSheet.getMaxColumns());
    }

    // ヘッダー行を常に最新定義で上書き（列追加や名称変更に対応するため毎回書き直す）
    johoSheet.getRange(1, 1, 1, JOHO_COLS).setValues([johoHdr]);
    johoSheet.setFrozenRows(1);

    // ── ヘッダー背景色 ────────────────────────────────────────────────
    // A-B列(貨物チェック・進捗貨物): 濃グレー
    johoSheet.getRange(1, 1, 1, 2)
      .setBackground('#455a64').setFontColor('#ffffff').setFontWeight('bold');
    // C〜M列(貨物情報11列): 濃青
    johoSheet.getRange(1, 3, 1, 11)
      .setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold');
    // N-O列(車両チェック・進捗車両): 濃グレー（貨物側と対称）
    johoSheet.getRange(1, 14, 1, 2)
      .setBackground('#455a64').setFontColor('#ffffff').setFontWeight('bold');
    // P〜Z列(車両情報11列): 濃橙
    johoSheet.getRange(1, 16, 1, 11)
      .setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold');

    // 列幅
    johoSheet.setColumnWidth(1,  50);  // A: チェック(貨物)
    johoSheet.setColumnWidth(2,  70);  // B: 進捗(貨物)
    johoSheet.setColumnWidth(14, 50);  // N: チェック(車両)
    johoSheet.setColumnWidth(15, 70);  // O: 進捗(車両)
    johoSheet.setColumnWidth(6,  90);  // F: 日付
    johoSheet.setColumnWidth(7, 120);  // G: 品名

    // ── データ行バリデーション（2〜501行） ─────────────────────────────
    var dataRows = 500;
    if (johoSheet.getMaxRows() < dataRows + 1) {
      johoSheet.insertRowsAfter(johoSheet.getMaxRows(), dataRows + 1 - johoSheet.getMaxRows());
    }
    var chkRule  = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    var progRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['', 'キャンセル', '終了', '確定'], true)
      .setAllowInvalid(true).build();
    johoSheet.getRange(2, 1,  dataRows, 1).setDataValidation(chkRule);  // A: 貨物チェック
    johoSheet.getRange(2, 2,  dataRows, 1).setDataValidation(progRule); // B: 進捗(貨物)
    johoSheet.getRange(2, 14, dataRows, 1).setDataValidation(chkRule);  // N: 車両チェック
    johoSheet.getRange(2, 15, dataRows, 1).setDataValidation(progRule); // O: 進捗(車両)
    johoSheet.getRange(2, 6,  dataRows, 1).setNumberFormat('yyyy/MM/dd'); // F: 日付
  })();

  applyHolidayRowColors_();
  SpreadsheetApp.getUi().alert('シート再生成が完了しました。');
}


// ================================================================
//  4-3b: 経費自動入力（autoFillExpense）  【大B / 中4 / 小4-3b】
//  自車専属マスタで選択中の行のトン数（F列）を読み
//  トン数別平均値を15経費列（Q〜AE列）に一括セットする
//  ・平均値セル → 文字色を赤にして「自動入力」とわかるようにする
//  ・手入力済みセル（文字色が黒/null）は上書きしない
//  ・トン数不明は4トンを使うか確認（1〜30t対応）
// ================================================================
function autoFillExpense() {
  var ui    = SpreadsheetApp.getUi();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  if (sheet.getName() !== '自車専属マスタ') {
    ui.alert('自車専属マスタシートを開いた状態で実行してください。');
    return;
  }

  var selRange    = sheet.getActiveRange();
  var firstRow    = selRange.getRow();
  var numSelRows  = selRange.getNumRows();
  if (firstRow < 2) { ui.alert('データ行（2行目以降）を選択してください。'); return; }

  // 経費開始列を動的取得（高速を引く列の有無に応じてcol17またはcol18）
  var mLastColAF = sheet.getLastColumn();
  var mHdrAF = mLastColAF > 0 ? sheet.getRange(1, 1, 1, mLastColAF).getValues()[0] : [];
  var firstExpColAF = mHdrAF.indexOf('車両リース代');
  if (firstExpColAF < 0) {
    ui.alert('経費列（車両リース代）が見つかりません。シート再生成を先に実行してください。');
    return;
  }
  var EXP_COL = firstExpColAF + 1; // 1-based
  var EXP_NUM = 15;
  var AUTO_COLOR = '#cc0000';

  // トン数別平均値テーブル（Q〜AE列＝15経費列: 1t〜30t）
  // 列順: 車両リース代,任意保険,自賠責,重量税積立,車検費積立,整備費積立,タイヤ代積立,修理積立,駐車場代,ETCリース,カーナビ,通信費,洗車費,制服費,その他固定費
  var expTable = {
     1: [ 35000, 10000, 1200,  1500,  3000,  7000,  2500,  5000, 10000, 1500, 2000, 2000, 1500, 1000,  3000],
     2: [ 60000, 14000, 1350,  2500,  4500, 13000,  4500, 10000, 15000, 1500, 2000, 2000, 2000, 1000,  5000],
     3: [ 80000, 17000, 1550,  3500,  6000, 16000,  6000, 12000, 17000, 1500, 2000, 2000, 2800, 1000,  7500],
     4: [100000, 20000, 1750,  5000,  7500, 20000,  7500, 15000, 20000, 1500, 2000, 2000, 3500, 1000, 10000],
     5: [110000, 23000, 2000,  6500,  9000, 23000,  9000, 17000, 21000, 1500, 2000, 2000, 4000, 1000, 11000],
     6: [120000, 25000, 2100,  7500, 10500, 26000, 10500, 19000, 22000, 1500, 2000, 2000, 4300, 1000, 12000],
     7: [130000, 27000, 2200,  8500, 11500, 28000, 11500, 22000, 23000, 1500, 2000, 2000, 4500, 1000, 13000],
     8: [140000, 29000, 2300,  9500, 12500, 31000, 12500, 25000, 24000, 1500, 2000, 2000, 4700, 1000, 14000],
     9: [150000, 30000, 2400, 10000, 13500, 33000, 13500, 27000, 24500, 1500, 2000, 2000, 4900, 1000, 14500],
    10: [160000, 32000, 2500, 11000, 15000, 35000, 15000, 30000, 25000, 1500, 2000, 2000, 5000, 1000, 15000],
    11: [165000, 33000, 2600, 12000, 15500, 36000, 16000, 32000, 26000, 1500, 2000, 2000, 5300, 1000, 16000],
    12: [170000, 34000, 2700, 13000, 16500, 37000, 17000, 34000, 27000, 1500, 2000, 2000, 5500, 1000, 17000],
    13: [175000, 35000, 2800, 14000, 17500, 38000, 18000, 36000, 28000, 1500, 2000, 2000, 5800, 1000, 17500],
    14: [180000, 36000, 3000, 15000, 18000, 39000, 19000, 37000, 29000, 1500, 2000, 2000, 6000, 1000, 18000],
    15: [185000, 37000, 3200, 16000, 18500, 40000, 20000, 38000, 30000, 1500, 2000, 2000, 6200, 1000, 18500],
    16: [188000, 38000, 3400, 16500, 19000, 41000, 21000, 39000, 30500, 1500, 2000, 2000, 6400, 1000, 19000],
    17: [190000, 39000, 3600, 17000, 19500, 42000, 22000, 40000, 31000, 1500, 2000, 2000, 6600, 1000, 19500],
    18: [192000, 40000, 3800, 17500, 20000, 44000, 23000, 41000, 32000, 1500, 2000, 2000, 6800, 1000, 19500],
    19: [195000, 41000, 4000, 18000, 20500, 46000, 25000, 42000, 33000, 1500, 2000, 2000, 7000, 1000, 19500],
    20: [200000, 42000, 4500, 18500, 21000, 47000, 27000, 43000, 33500, 1500, 2000, 2000, 7200, 1000, 19500],
    21: [205000, 43000, 4600, 19000, 21000, 48000, 28000, 44000, 34000, 1500, 2000, 2000, 7400, 1000, 20000],
    22: [210000, 43500, 4700, 19500, 21500, 48500, 28500, 44500, 34500, 1500, 2000, 2000, 7500, 1000, 20000],
    23: [215000, 44000, 4800, 19800, 21800, 49000, 29000, 45000, 35000, 1500, 2000, 2000, 7700, 1000, 20000],
    24: [218000, 44300, 4850, 20000, 22000, 49500, 29500, 45500, 35000, 1500, 2000, 2000, 7800, 1000, 20000],
    25: [220000, 44500, 4900, 20200, 22000, 50000, 30000, 45000, 35000, 1500, 2000, 2000, 8000, 1000, 20000],
    26: [222000, 44700, 5000, 20300, 22000, 50000, 30000, 45000, 35000, 1500, 2000, 2000, 8000, 1000, 20000],
    27: [225000, 44800, 5000, 20400, 22000, 50000, 30000, 45000, 35000, 1500, 2000, 2000, 8000, 1000, 20000],
    28: [228000, 45000, 5000, 20500, 22000, 50000, 30000, 45000, 35000, 1500, 2000, 2000, 8000, 1000, 20000],
    29: [230000, 45000, 5000, 20500, 22000, 50000, 30000, 45000, 35000, 1500, 2000, 2000, 8000, 1000, 20000],
    30: [235000, 45000, 5000, 20500, 22000, 50000, 30000, 45000, 35000, 1500, 2000, 2000, 8000, 1000, 20000]
  };

  // トン数を一括読み込み（全角→半角変換）
  var tonsData    = sheet.getRange(firstRow, 6, numSelRows, 1).getValues();
  var existValsAll  = sheet.getRange(firstRow, EXP_COL, numSelRows, EXP_NUM).getValues();
  var fontColsAll   = sheet.getRange(firstRow, EXP_COL, numSelRows, EXP_NUM).getFontColors();

  // トン数を解析し、不明行を検出
  var tNums = [];
  var hasUnknown = false;
  for (var i = 0; i < numSelRows; i++) {
    var raw = String(tonsData[i][0] || '').trim();
    raw = raw.replace(/[０-９]/g, function(c){ return String.fromCharCode(c.charCodeAt(0)-0xFEE0); });
    var tNum = parseInt(raw.replace(/[^0-9]/g, ''), 10);
    tNums.push((isNaN(tNum) || tNum < 1) ? null : Math.min(30, Math.max(1, tNum)));
    if (tNums[i] === null) hasUnknown = true;
  }

  // トン数不明行の扱いを確認（1回だけ）
  var use4tForUnknown = false;
  if (hasUnknown) {
    var res = ui.alert(
      'トン数不明の行があります',
      'F列のトン数が読み取れない行があります。\n4トンの平均値を使用しますか？（「いいえ」の場合はその行をスキップ）',
      ui.ButtonSet.YES_NO
    );
    use4tForUnknown = (res === ui.Button.YES);
    if (use4tForUnknown) {
      for (var i = 0; i < tNums.length; i++) {
        if (tNums[i] === null) tNums[i] = 4;
      }
    }
  }

  // 新しい値・色の2D配列を構築（手入力済みは維持、自動入力・空は上書き）
  var newVals   = [];
  var newColors = [];
  var successCount = 0, skipCount = 0, manualTotal = 0;

  for (var i = 0; i < numSelRows; i++) {
    if (tNums[i] === null) {
      newVals.push(existValsAll[i].slice());
      newColors.push(fontColsAll[i].slice());
      skipCount++;
      continue;
    }
    var vals    = expTable[tNums[i]];
    var newRow  = [];
    var colorRow= [];
    for (var j = 0; j < EXP_NUM; j++) {
      var hasVal = existValsAll[i][j] !== '' && existValsAll[i][j] !== 0;
      var isAuto = String(fontColsAll[i][j] || '').toLowerCase() === AUTO_COLOR;
      if (hasVal && !isAuto) {
        manualTotal++;
        newRow.push(existValsAll[i][j]);
        colorRow.push(fontColsAll[i][j]);
      } else {
        newRow.push(vals[j]);
        colorRow.push(AUTO_COLOR);
      }
    }
    newVals.push(newRow);
    newColors.push(colorRow);
    successCount++;
  }

  // 一括書き込み
  sheet.getRange(firstRow, EXP_COL, numSelRows, EXP_NUM).setValues(newVals);
  sheet.getRange(firstRow, EXP_COL, numSelRows, EXP_NUM).setFontColors(newColors);

  var msg = successCount + '行に平均値を入力しました（赤字）。';
  if (skipCount > 0)   msg += '\n' + skipCount + '行はトン数不明のためスキップしました。';
  if (manualTotal > 0) msg += '\n手入力済み項目（合計' + manualTotal + '件）はそのまま残しました。';
  msg += '\n実態に合わせて修正してください。修正すると黒字に変わります。';
  ui.alert(msg);
}


// ================================================================
//  4-4: 支払い再計算（calculatePaymentAmount）  【大B / 中4 / 小4-4】
//  集計表のAB列(仮日数)・AC列(給料)・AD列(%)からY列(支払い)を計算する
//  ・パターンA: %あり → (売上-合計高速代)×%/100
//  ・パターンB: %なし・給料と仮日数あり → 給料÷仮日数
//              片方欠け → 欠けているセルを赤警告
//  ・パターンC: 条件なし → Y列が空なら赤警告（手入力値は保持）
// ================================================================
function calculatePaymentAmount(companySsId) {
  var ss    = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
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

  // 自車専属マスタから 車番+乗務員名 → 高速を引く 設定を取得
  var tollDeductMap = {};
  var masterForCalc = ss.getSheetByName('自車専属マスタ');
  if (masterForCalc && masterForCalc.getLastRow() >= 2) {
    var mcReadCols = Math.max(masterForCalc.getLastColumn(), 31);
    var mcHdr = masterForCalc.getRange(1, 1, 1, mcReadCols).getValues()[0];
    var mcHasToll = String(mcHdr[16]||'').trim() === '高速を引く（引くは〇、引かないは空欄）';
    if (mcHasToll) {
      var mcData = masterForCalc.getRange(2, 1, masterForCalc.getLastRow()-1, mcReadCols).getValues();
      for (var mc = 0; mc < mcData.length; mc++) {
        var mcKey = String(mcData[mc][7]||'').trim() + '_' + String(mcData[mc][8]||'').trim();
        tollDeductMap[mcKey] = String(mcData[mc][16] || '').trim();
      }
    }
  }

  var data = sheet.getRange(2, 1, lastRow-1, 35).getValues();
  var yukyuVals = [];
  var acVals    = [];

  for (var i = 0; i < data.length; i++) {
    var rowNum    = i + 2;
    var sales     = Number(data[i][18]) || 0;
    var totalToll = Number(data[i][21]) || 0;  // V=col22
    var fuelCost  = Number(data[i][25]) || 0;  // Z=col26=燃料代
    var expense   = Number(data[i][27]) || 0;  // AB=col28=経費合計
    var kari      = Number(data[i][30]) || 0;  // AE=col31
    var kyuryo    = Number(data[i][31]) || 0;  // AF=col32
    var pct       = Number(data[i][32]) || 0;  // AG=col33
    var pick      = String(data[i][11] || '');
    var drop      = String(data[i][12] || '');
    var isYukyu   = pick.indexOf('有休') !== -1 || drop.indexOf('有休') !== -1;
    var isYasumi  = !isYukyu && (pick.indexOf('休み') !== -1 || drop.indexOf('休み') !== -1);
    var yCell     = sheet.getRange(rowNum, 27);

    yCell.setBackground(null);
    sheet.getRange(rowNum, 31, 1, 3).setBackground(null);

    var carNameKey = String(data[i][5]||'').trim() + '_' + String(data[i][6]||'').trim();
    var tollDeductC = tollDeductMap[carNameKey] || '';
    var yukyuVal = '';
    var finalPay;
    if (pct > 0) {
      var effectiveTollC = (tollDeductC === '○') ? totalToll : 0;
      finalPay = Math.round((sales - effectiveTollC) * (pct / 100));
      yCell.setValue(finalPay);
      if (isYukyu) yukyuVal = yukyuRate;
    } else if (kyuryo > 0 || kari > 0) {
      if (kyuryo > 0 && kari > 0) {
        var dailyPay = Math.round(kyuryo / kari);
        finalPay = isYasumi ? -dailyPay : dailyPay;
        yCell.setValue(finalPay);
      } else {
        if (!kyuryo) sheet.getRange(rowNum, 32).setBackground('#f4cccc');
        if (!kari)   sheet.getRange(rowNum, 31).setBackground('#f4cccc');
        finalPay = Number(data[i][26]) || 0;
      }
    } else {
      if (yCell.getValue() === '') yCell.setBackground('#f4cccc');
      finalPay = Number(data[i][26]) || 0;
    }
    yukyuVals.push([yukyuVal]);
    // AC(29)=利益 を更新
    var acAllEmpty = !sales && !totalToll && !fuelCost && !finalPay && !expense;
    acVals.push([acAllEmpty ? '' : sales-(totalToll+fuelCost+finalPay+expense)]);
  }

  if (yukyuVals.length > 0) {
    sheet.getRange(2, 34, yukyuVals.length, 1).setValues(yukyuVals);
    sheet.getRange(2, 29, acVals.length,    1).setValues(acVals);
  }
}


// ================================================================
//  4-5: 自車専属運行シート更新内部処理（refreshActiveVehiclesAuto_）  【大B / 中4 / 小4-5】
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

  var header  = ['車両ID','運行状態','区分','会社名','看板名','トン数','車種','車番','乗務員名','携帯番号','アドレス','燃費','備考','仮日数','給料','％'];
  var lastRow = master.getLastRow();
  var mData   = lastRow >= 2 ? master.getRange(2, 1, lastRow-1, 16).getValues() : [];
  var outRows = [header];
  for (var i = 0; i < mData.length; i++) {
    if (String(mData[i][1]||'').trim() === '運行') {
      outRows.push([
        mData[i][0],  mData[i][1],  mData[i][2],  mData[i][3],  mData[i][4],
        mData[i][5],  mData[i][6],  mData[i][7],  mData[i][8],  mData[i][9],
        mData[i][10], mData[i][11], mData[i][12], mData[i][13], mData[i][14], mData[i][15]
      ]);
    }
  }
  activeSheet.clear();
  if (outRows.length > 0) {
    activeSheet.getRange(1, 1, outRows.length, 16).setValues(outRows);
    activeSheet.setFrozenRows(1);
  }
}


// ================================================================
//  4-6: 自車専属マスタに「運行」列追加（addStatusColumnToMaster）  【大B / 中4 / 小4-6】
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
//  4-6b: 今月分生成（generateCurrentMonth）  【大C / 中4 / 小4-6b】
//  今日〜今月末日 × 運行中車両 のプレースホルダーIDを運行シートに生成する
//  月途中で契約した会社の初期設定時に使用する
// ================================================================
function generateCurrentMonth() {
  var ui    = SpreadsheetApp.getUi();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) { ui.alert('運行シートが見つかりません'); return; }

  var today       = new Date();
  var curYear     = today.getFullYear();
  var curMon      = today.getMonth(); // 0-indexed
  var startDay    = today.getDate();
  var daysInMonth = new Date(curYear, curMon + 1, 0).getDate();

  var master = ss.getSheetByName('自車専属マスタ');
  if (!master || master.getLastRow() < 2) {
    ui.alert('自車専属マスタに車両データがありません');
    return;
  }
  var mData = master.getRange(2, 1, master.getLastRow() - 1, 16).getValues();
  var activeVehicles = [];
  for (var v = 0; v < mData.length; v++) {
    if (String(mData[v][1] || '').trim() === '運行') activeVehicles.push(mData[v]);
  }
  if (activeVehicles.length === 0) {
    ui.alert('運行中の車両がありません。\n自車専属マスタのB列（運行状態）を確認してください。');
    return;
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e) { ui.alert('ロック取得失敗: ' + e.message); return; }

  try {
    var insertRow = sheet.getLastRow() + 1;
    var nextNum   = getNextIdNum_(sheet, 'V-');
    var rowsData  = [];
    var formulas  = [];

    for (var day = startDay; day <= daysInMonth; day++) {
      var dateObj = new Date(curYear, curMon, day);
      for (var v2 = 0; v2 < activeVehicles.length; v2++) {
        var veh   = activeVehicles[v2];
        var rowId = 'V-' + String(nextNum).padStart(4, '0');
        nextNum++;
        var rn = insertRow + rowsData.length;
        rowsData.push([
          rowId,  veh[2], veh[3], veh[5], veh[6], veh[7], veh[8], veh[9],
          veh[4], dateObj, '', '', '',
          '', '', '', '', '',
          '', '', '', '',
          '', '', '', ''
        ]);
        formulas.push(['=IF(AND(U' + rn + '="",T' + rn + '=""),"",U' + rn + '-T' + rn + ')']);
      }
    }

    sheet.getRange(insertRow, 1, rowsData.length, 26).setValues(rowsData);
    sheet.getRange(insertRow, 22, formulas.length, 1).setFormulas(formulas);
    sheet.getRange(insertRow, 10, rowsData.length, 1).setNumberFormat('yyyy/MM/dd');
    sheet.getRange(insertRow, 12, rowsData.length, 2).setNumberFormat('@');
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  sortUnkouByDate_();
  applyHolidayRowColors_();

  var generatedDays = daysInMonth - startDay + 1;
  ui.alert(
    '📅 今月分生成完了\n\n' +
    curYear + '年' + (curMon + 1) + '月 ' + startDay + '日〜' + daysInMonth + '日\n' +
    activeVehicles.length + '台 × ' + generatedDays + '日 = ' +
    (activeVehicles.length * generatedDays) + '行を生成しました。'
  );
}


// ================================================================
//  4-7: 月生成（generateNextMonth）  【大C / 中4 / 小4-7】
//  次月の1日〜末日 × 運行中車両 のプレースホルダーIDを運行シートに一括生成する
//  ・積地が空の行には配車漏れ警告色（#fff9c4）を付ける（1-7呼び出し）
//  ・生成後に運行シートに3ヶ月以上のデータがあれば最古月を自動アーカイブ（4-8a呼び出し）
//  ・重複防止: 次月データが既存なら処理中止
//  ・LockServiceで同時実行によるID重複を防止
// ================================================================
function generateNextMonth() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) { ui.alert('運行シートが見つかりません'); return; }

  // 次月の年・月(0-indexed)を計算
  var today    = new Date();
  var nextYear = (today.getMonth() === 11) ? today.getFullYear() + 1 : today.getFullYear();
  var nextMon  = (today.getMonth() + 1) % 12; // 0-indexed

  // 既に次月データがある場合はスキップ
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var dateVals = sheet.getRange(2, 10, lastRow - 1, 1).getValues();
    for (var i = 0; i < dateVals.length; i++) {
      var dv = dateVals[i][0];
      if (dv instanceof Date && dv.getFullYear() === nextYear && dv.getMonth() === nextMon) {
        ui.alert('📅 月生成\n\n' + nextYear + '年' + (nextMon + 1) + '月分のデータは既に存在します。\n重複生成はできません。');
        return;
      }
    }
  }

  // 運行中の車両一覧を取得（自車専属マスタ B列=運行状態が「運行」の行のみ）
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master || master.getLastRow() < 2) {
    ui.alert('自車専属マスタに車両データがありません');
    return;
  }
  var mData = master.getRange(2, 1, master.getLastRow() - 1, 16).getValues();
  var activeVehicles = [];
  for (var v = 0; v < mData.length; v++) {
    if (String(mData[v][1] || '').trim() === '運行') activeVehicles.push(mData[v]);
  }
  if (activeVehicles.length === 0) {
    ui.alert('運行中の車両がありません。\n自車専属マスタのB列（運行状態）を確認してください。');
    return;
  }

  var daysInMonth = new Date(nextYear, nextMon + 1, 0).getDate();

  // LockServiceでID採番の競合を防止
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e) { ui.alert('ロック取得失敗: ' + e.message); return; }

  try {
    var insertRow = sheet.getLastRow() + 1;
    var nextNum   = getNextIdNum_(sheet, 'V-');
    var rowsData  = [];
    var formulas  = [];

    // 日付ループ（1日〜末日）× 車両ループ
    for (var day = 1; day <= daysInMonth; day++) {
      var dateObj = new Date(nextYear, nextMon, day);
      for (var v2 = 0; v2 < activeVehicles.length; v2++) {
        var veh = activeVehicles[v2];
        var rowId = 'V-' + String(nextNum).padStart(4, '0');
        nextNum++;
        var rn = insertRow + rowsData.length;
        rowsData.push([
          rowId,   veh[2], veh[3], veh[5], veh[6], veh[7], veh[8], veh[9],
          veh[4],  dateObj, '', '', '',
          '', '', '', '', '',  // 誘導〜降完（14-18列）
          '', '', '',          // 売上・請求高速・実費高速（19-21列）
          '',                  // 合計高速（22列: 後で数式セット）
          '', '', '', ''       // 連絡・データURL（23-26列）
        ]);
        formulas.push(['=IF(AND(U' + rn + '="",T' + rn + '=""),"",U' + rn + '-T' + rn + ')']);
      }
    }

    // 一括書き込み
    sheet.getRange(insertRow, 1, rowsData.length, 26).setValues(rowsData);
    sheet.getRange(insertRow, 22, formulas.length, 1).setFormulas(formulas);
    sheet.getRange(insertRow, 10, rowsData.length, 1).setNumberFormat('yyyy/MM/dd');
    // 積地・降地列をテキスト書式（数値化防止）
    sheet.getRange(insertRow, 12, rowsData.length, 2).setNumberFormat('@');
    SpreadsheetApp.flush();

  } finally {
    lock.releaseLock();
  }

  // 日付順にソート（新規行が末尾に追加されているため）
  sortUnkouByDate_();

  // 積地空 + 未来日 = 黄色警告を全体に適用
  applyHolidayRowColors_();

  // 3ヶ月以上データがあれば最古月を自動アーカイブ
  archiveOldestMonthIfNeeded_();

  ui.alert(
    '📅 月生成完了\n\n' +
    nextYear + '年' + (nextMon + 1) + '月分\n' +
    activeVehicles.length + '台 × ' + daysInMonth + '日 = ' +
    (activeVehicles.length * daysInMonth) + '行を生成しました。'
  );
}


// ================================================================
//  4-8: 前月分アーカイブ（archiveOldMonth）  【大C / 中4 / 小4-8】
//  前月（今月-1）の運行シート・集計表を別スプレッドシートに値のみで保存して元行を削除する
//  保存先: Googleドライブ「運行管理_アーカイブ/会社名/YYYY年MM月_会社名」
//  自車専属マスタのJ列（メールアドレス）に登録された全員に編集権限を付与
// ================================================================
function archiveOldMonth() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('運行シートにデータがありません');
    return;
  }

  // 今月のキー（年×100＋月の0-indexed）
  var today = new Date();
  var thisMonKey = today.getFullYear() * 100 + today.getMonth();

  // 運行シートに存在する「今月より前の月」を全件収集
  var dateVals = sheet.getRange(2, 10, sheet.getLastRow() - 1, 1).getValues();
  var monthMap = {};
  for (var i = 0; i < dateVals.length; i++) {
    var dv = dateVals[i][0];
    if (dv instanceof Date) {
      var key = dv.getFullYear() * 100 + dv.getMonth();
      if (key < thisMonKey) monthMap[key] = { year: dv.getFullYear(), month: dv.getMonth() };
    }
  }

  var keys = Object.keys(monthMap).map(Number).sort(); // 古い月から順に
  if (keys.length === 0) {
    ui.alert('📦 前月分アーカイブ\n\n今月より前のデータが見つかりません。\n既にアーカイブ済みか、データがない可能性があります。');
    return;
  }

  var companyName = getCompanyName_(ss);
  var totalArchived = 0;
  var lines = [];
  for (var k = 0; k < keys.length; k++) {
    var m = monthMap[keys[k]];
    var result = archiveMonthData_(ss, m.year, m.month, companyName);
    if (result.archived > 0) {
      totalArchived += result.archived;
      lines.push(m.year + '年' + (m.month + 1) + '月（' + result.archived + '行）');
    }
  }

  if (totalArchived === 0) {
    ui.alert('📦 前月分アーカイブ\n\nアーカイブ対象のデータがありませんでした。');
    return;
  }

  ui.alert(
    '📦 アーカイブ完了\n\n' +
    lines.join('\n') + '\n\n' +
    '合計 ' + totalArchived + '行をアーカイブしました。\n' +
    '保存先: 運行管理_アーカイブ/' + companyName + '/'
  );
}


// ================================================================
//  4-8a: 最古月自動アーカイブ（archiveOldestMonthIfNeeded_）  【大B / 中4 / 小4-8a】
//  運行シートに3ヶ月以上の月データが存在する場合のみ最古月をアーカイブする
//  generateNextMonth から呼び出す（月生成後の自動整理）
// ================================================================
function archiveOldestMonthIfNeeded_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet || sheet.getLastRow() < 2) return;

  // 運行シートに存在する月を収集（キー: YYYYMM形式の数値）
  var dateVals = sheet.getRange(2, 10, sheet.getLastRow() - 1, 1).getValues();
  var monthMap = {};
  for (var i = 0; i < dateVals.length; i++) {
    var dv = dateVals[i][0];
    if (dv instanceof Date) {
      var key = dv.getFullYear() * 100 + dv.getMonth(); // e.g. 202504
      monthMap[key] = { year: dv.getFullYear(), month: dv.getMonth() };
    }
  }

  var keys = Object.keys(monthMap).map(Number).sort();
  if (keys.length < 3) return; // 2ヶ月以内はアーカイブ不要

  // 最古月をアーカイブ
  var oldest = monthMap[keys[0]];
  var companyName = getCompanyName_(ss);
  archiveMonthData_(ss, oldest.year, oldest.month, companyName);
}


// ================================================================
//  4-8b: 月別アーカイブ実行（archiveMonthData_）  【大B / 中4 / 小4-8b】
//  指定年月の運行・集計表データをアーカイブ用スプレッドシートへコピーし元行を削除する
//  処理順序:
//   1. 対象月の行を収集
//   2. 新規スプレッドシート作成・2シート(運行/集計表)に値のみ書き込み
//   3. 「運行管理_アーカイブ/会社名/」フォルダへ移動
//   4. 自車専属マスタJ列(メールアドレス)全員に編集権限を付与
//   5. 運行シート・集計表から対象行を削除
// ================================================================
function archiveMonthData_(ss, year, month, companyName) {
  var monthLabel = year + '年' + (month + 1) + '月';
  var fileName   = monthLabel + '_' + companyName;

  // 運行シートから対象月の行を取得
  var unkouSheet = ss.getSheetByName('運行');
  var unkouRows  = getMonthRows_(unkouSheet, year, month, 26);
  if (unkouRows.length === 0) return { fileName: fileName, archived: 0 };

  // 対象IDを収集（集計表の絞り込みに使用）
  var archiveIds = {};
  for (var i = 0; i < unkouRows.length; i++) {
    var id = String(unkouRows[i].data[0] || '').trim();
    if (id) archiveIds[id] = true;
  }

  // 集計表から対象IDの行を取得
  var sumSheet = ss.getSheetByName('集計表');
  var sumRows  = [];
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var sumVals = sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, 34).getValues();
    for (var j = 0; j < sumVals.length; j++) {
      var sid = String(sumVals[j][0] || '').trim();
      if (archiveIds[sid]) sumRows.push({ rowNum: j + 2, data: sumVals[j] });
    }
  }

  // 新規スプレッドシート作成
  var newSs   = SpreadsheetApp.create(fileName);
  var newFile = DriveApp.getFileById(newSs.getId());

  // 運行シートを書き込み（値のみ）
  var newUnkou   = newSs.getActiveSheet();
  newUnkou.setName('運行');
  var unkouHeader = unkouSheet.getRange(1, 1, 1, 26).getValues()[0];
  var unkouData   = [unkouHeader].concat(unkouRows.map(function(r) { return r.data; }));
  newUnkou.getRange(1, 1, unkouData.length, 26).setValues(unkouData);
  newUnkou.setFrozenRows(1);
  newUnkou.getRange(1, 1, 1, 26).setBackground('#efefef');

  // 集計表シートを書き込み（値のみ）
  if (sumRows.length > 0) {
    var newSum    = newSs.insertSheet('集計表');
    var sumHeader = sumSheet.getRange(1, 1, 1, 34).getValues()[0];
    var sumData   = [sumHeader].concat(sumRows.map(function(r) { return r.data; }));
    newSum.getRange(1, 1, sumData.length, 34).setValues(sumData);
    newSum.setFrozenRows(1);
    newSum.getRange(1, 1, 1, 34).setBackground('#efefef');
  }

  // アーカイブフォルダへ移動（ルートから削除）
  var archiveRoot   = getOrCreateFolder_('運行管理_アーカイブ');
  var subIter       = archiveRoot.getFoldersByName(companyName);
  var companyFolder = subIter.hasNext() ? subIter.next() : archiveRoot.createFolder(companyName);
  companyFolder.addFile(newFile);
  DriveApp.getRootFolder().removeFile(newFile);

  // 自車専属マスタのJ列(index[9]=メールアドレス)に登録された全員に編集権限を付与
  var master = ss.getSheetByName('自車専属マスタ');
  if (master && master.getLastRow() >= 2) {
    var mData = master.getRange(2, 1, master.getLastRow() - 1, 11).getValues();
    var sharedEmails = {};
    for (var m = 0; m < mData.length; m++) {
      var email = String(mData[m][10] || '').trim();
      if (email && email.indexOf('@') !== -1 && !sharedEmails[email]) {
        sharedEmails[email] = true;
        try { newFile.addEditor(email); } catch(e) {}
        // 会社フォルダも共有（ファイルだけでなくフォルダを見えるように）
        try { companyFolder.addEditor(email); } catch(e) {}
      }
    }
  }

  // 元データから対象行を一括削除（連続行をまとめて deleteRows でタイムアウト回避）
  SpreadsheetApp.flush();
  deleteRowsGrouped_(unkouSheet, unkouRows.map(function(r) { return r.rowNum; }));

  if (sumRows.length > 0) {
    deleteRowsGrouped_(sumSheet, sumRows.map(function(r) { return r.rowNum; }));
  }

  return { fileName: fileName, archived: unkouRows.length };
}


// 連続する行番号をまとめて deleteRows で一括削除（高速・タイムアウト対策）
function deleteRowsGrouped_(sheet, rowNums) {
  var nums = rowNums.slice().sort(function(a, b) { return b - a; }); // 降順
  var i = 0;
  while (i < nums.length) {
    var top = nums[i];
    var count = 1;
    while (i + count < nums.length && nums[i + count] === top - count) { count++; }
    sheet.deleteRows(top - count + 1, count);
    i += count;
  }
}


// ================================================================
//  4-8c: 月別行データ抽出（getMonthRows_）  【大B / 中4 / 小4-8c】
//  指定シートのJ列(col10=日付)で絞り込み、指定年月(0-indexed)の行を返す
//  戻り値: [{rowNum: 行番号(1-indexed), data: 行データ配列}] の配列
// ================================================================
function getMonthRows_(sheet, year, month, numCols) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var lastRow = sheet.getLastRow();
  var vals    = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
  var result  = [];
  for (var i = 0; i < vals.length; i++) {
    var dv = vals[i][9]; // col10(J列=日付) → 0-indexed: [9]
    if (dv instanceof Date && dv.getFullYear() === year && dv.getMonth() === month) {
      result.push({ rowNum: i + 2, data: vals[i] });
    }
  }
  return result;
}


// ================================================================
//  4-8d: 会社名取得（getCompanyName_）  【大B / 中4 / 小4-8d】
//  自車専属マスタのD列(col4=会社名)から最初の値を取得する
//  取得できない場合はスプレッドシート名から「運行管理」を除いた文字列を返す
// ================================================================
function getCompanyName_(ss) {
  var master = ss.getSheetByName('自車専属マスタ');
  if (master && master.getLastRow() >= 2) {
    var mData = master.getRange(2, 1, master.getLastRow() - 1, 4).getValues();
    for (var i = 0; i < mData.length; i++) {
      var name = String(mData[i][3] || '').trim(); // col4(D列=会社名) → 0-indexed: [3]
      if (name) return name;
    }
  }
  // フォールバック: スプレッドシート名から取得
  return ss.getName().replace(/\s*運行管理\s*$/, '').trim() || '会社名不明';
}


// ================================================================
//  5-1: 起動時の初期データ一括取得（getInitialData）  【大A / 中5 / 小5-1】
//  端末保存のアドレスを元にマスタから該当行を一括検索して返す
// ================================================================
function getInitialData(hintEmail, companySsId) {
  var savedEmail = hintEmail || '';
  var result = { email: savedEmail, profile: null };
  if (!savedEmail) return result;
  var ss     = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master) return result;
  var data       = master.getDataRange().getValues();
  var emailLower = savedEmail.toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    if (data[i][10] && String(data[i][10]).toLowerCase().trim() === emailLower) {
      result.profile = {
        company: data[i][3], tons: data[i][5], type: data[i][6],
        carNo:   data[i][7], name: data[i][8], tel:  data[i][9]
      };
      break;
    }
  }
  // 設定シートから点検項目を読む（アプリ起動時に1回だけ取得してキャッシュ）
  var settingSheet = ss.getSheetByName('設定');
  result.inspectionBefore = [];
  result.inspectionAfter  = [];
  if (settingSheet && settingSheet.getLastRow() >= 1) {
    var sLastCol = settingSheet.getLastColumn();
    if (sLastCol > 0) {
      var sHeaders = settingSheet.getRange(1, 1, 1, sLastCol).getValues()[0];
      var bCol = -1, aCol = -1;
      for (var shi = 0; shi < sHeaders.length; shi++) {
        var sh = String(sHeaders[shi]||'').trim();
        if (sh === '業務前点検') bCol = shi;
        if (sh === '業務後点検') aCol = shi;
      }
      var sLastRow = settingSheet.getLastRow();
      if (bCol >= 0 && sLastRow >= 2) {
        var bVals = settingSheet.getRange(2, bCol+1, sLastRow-1, 1).getValues();
        for (var bi = 0; bi < bVals.length; bi++) { var bv = String(bVals[bi][0]||'').trim(); if (bv) result.inspectionBefore.push(bv); }
      }
      if (aCol >= 0 && sLastRow >= 2) {
        var aVals = settingSheet.getRange(2, aCol+1, sLastRow-1, 1).getValues();
        for (var ai2 = 0; ai2 < aVals.length; ai2++) { var av = String(aVals[ai2][0]||'').trim(); if (av) result.inspectionAfter.push(av); }
      }
    }
  }
  return result;
}


// ================================================================
//  5-1b: 車番でマスタ検索（getCarInfoByNumber）  【大A / 中5 / 小5-1b】
//  行程入力フォームの車番フィールドからマスタ補完に使用
//  完全一致のみ。マスタにない車番は null を返す
// ================================================================
function getCarInfoByNumber(carNo, companySsId) {
  var ss     = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master || master.getLastRow() < 2) return null;
  var mData = master.getRange(2, 1, master.getLastRow() - 1, 10).getValues();
  for (var i = 0; i < mData.length; i++) {
    var masterCar = String(mData[i][6] || '').trim(); // G列(7)=車番
    if (!masterCar) continue;
    if (masterCar === carNo) {
      return {
        tons:    String(mData[i][4] || '').trim(), // E列(5)=トン数
        type:    String(mData[i][5] || '').trim(), // F列(6)=車種
        company: String(mData[i][3] || '').trim(), // D列(4)=会社名（看板名）
        name:    String(mData[i][7] || '').trim(), // H列(8)=乗務員名
        tel:     String(mData[i][8] || '').trim()  // I列(9)=携帯番号
      };
    }
  }
  return null;
}


// ================================================================
//  5-2: 紐づけ実行（linkAddress）  【大A / 中5 / 小5-2】
//  入力アドレスを自車専属マスタのJ列と照合し
//  一致したら端末のPropertiesServiceに保存する（シートには書かない）
// ================================================================
function linkAddress(email, companySsId) {
  var ss     = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master) return "エラー：マスタシートなし";
  var rows = master.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][10]).trim() === String(email).trim()) {
      var props = PropertiesService.getUserProperties();
      props.setProperty('linkedEmail', email);
      // 紐づけ時に会社SSのIDも保存（次回以降URLなしでも正しいSSを開ける）
      var linkedSsId = props.getProperty('linkedSsId');
      if (!linkedSsId) props.setProperty('linkedSsId', ss.getId());
      return {
        status: "紐づけOK", email: email,
        company: rows[i][3], tons: rows[i][5], type: rows[i][6],
        carNo:   rows[i][7], name: rows[i][8], tel:  rows[i][9]
      };
    }
  }
  return "エラー：アドレス未登録";
}


// ================================================================
//  5-3: 紐づけ解除（unlinkAddress）  【大A / 中5 / 小5-3】
//  端末のPropertiesServiceからアドレス情報を消去する
// ================================================================
function unlinkAddress(companySsId) {
  PropertiesService.getUserProperties().deleteProperty('linkedEmail');
  return "解除しました";
}


// ================================================================
//  6-1: 端末の運行進捗を保存（saveRunState）  【大A / 中6 / 小6-1】
//  picks/drops/rows/pickDone/dropDone/phase/lastPickRow/
//  pickHistory/dropHistoryの9項目をsetPropertiesで一括保存
// ================================================================
function saveRunState(state, email, companySsId) {
  validateDriverEmail_(email, companySsId);
  var p = PropertiesService.getUserProperties();
  p.setProperties({
    'picks':          JSON.stringify(state.picks          || []),
    'drops':          JSON.stringify(state.drops          || []),
    'rows':           JSON.stringify(state.rows           || []),
    'runId':          state.runId                         || '',
    'guideDone':      JSON.stringify(state.guideDone      || []),
    'pickDone':       JSON.stringify(state.pickDone       || []),
    'dropDone':       JSON.stringify(state.dropDone       || []),
    'phase':          state.phase                         || '',
    'lastPickIndex':  (state.lastPickIndex !== null && state.lastPickIndex !== undefined) ? String(state.lastPickIndex) : '',
    'guideHistory':   JSON.stringify(state.guideHistory   || []),
    'pickHistory':    JSON.stringify(state.pickHistory    || []),
    'dropHistory':    JSON.stringify(state.dropHistory    || [])
  }, false);
}


// ================================================================
//  6-2: 端末の運行進捗を読み込み（loadRunState）  【大A / 中6 / 小6-2】
//  getPropertiesで一括取得して返す
// ================================================================
function loadRunState() {
  var all = PropertiesService.getUserProperties().getProperties();
  var lpi = all['lastPickIndex'];
  return {
    picks:         JSON.parse(all['picks']         || '[]'),
    drops:         JSON.parse(all['drops']         || '[]'),
    rows:          JSON.parse(all['rows']          || '[]'),
    runId:         all['runId']                    || '',
    guideDone:     JSON.parse(all['guideDone']     || '[]'),
    pickDone:      JSON.parse(all['pickDone']      || '[]'),
    dropDone:      JSON.parse(all['dropDone']      || '[]'),
    phase:         all['phase']                    || '',
    lastPickIndex: (lpi !== '' && lpi !== undefined && lpi !== null) ? Number(lpi) : null,
    guideHistory:  JSON.parse(all['guideHistory']  || '[]'),
    pickHistory:   JSON.parse(all['pickHistory']   || '[]'),
    dropHistory:   JSON.parse(all['dropHistory']   || '[]')
  };
}


// ================================================================
//  6-3: 端末の運行進捗をクリア（clearRunState）  【大A / 中6 / 小6-3】
//  linkedEmail（紐づけ）とreadNotices（既読管理）は消さない
//  運行進捗の9項目だけ削除する
// ================================================================
function clearRunState(email, companySsId) {
  validateDriverEmail_(email, companySsId);
  var p    = PropertiesService.getUserProperties();
  var keys = ['picks','drops','rows','runId','guideDone','pickDone','dropDone','phase','lastPickIndex','guideHistory','pickHistory','dropHistory'];
  for (var i = 0; i < keys.length; i++) { p.deleteProperty(keys[i]); }
}


// ================================================================
//  7-1: 今日の行程取得（getTodayRoutes）  【大A / 中7 / 小7-1】
//  紐づけアドレスから乗務員名・車番を特定し
//  運行シートから本日分の未完了行程を返す
// ================================================================
function getTodayRoutes(email, companySsId) {
  var savedEmail = email || '';
  if (!savedEmail) return [];
  var ss     = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master) return [];
  var mAll = master.getDataRange().getValues();
  var name = '';
  for (var j = 1; j < mAll.length; j++) {
    if (String(mAll[j][10]).trim() === savedEmail) {
      name = String(mAll[j][8]).trim();
      break;
    }
  }
  if (!name) return [];
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return [];
  var all   = sheet.getDataRange().getValues();
  var today = new Date();
  var y = today.getFullYear(), m = today.getMonth(), d = today.getDate();
  var out = [];
  for (var i = 1; i < all.length; i++) {
    var r = all[i];
    if (!r[9]) continue;
    var dv = new Date(r[9]);
    if (dv.getFullYear()!==y || dv.getMonth()!==m || dv.getDate()!==d) continue;
    if (String(r[6]).trim()!==name) continue;
    if (r[17]) continue;
    var pickV = (r[11] instanceof Date) ? '' : String(r[11] || '');
    var dropV = (r[12] instanceof Date) ? '' : String(r[12] || '');
    out.push({ id: String(r[0]||'').trim(), row: i+1, pick: pickV, drop: dropV, guideDone: !!r[13], pickDone: !!r[14], dropDone: !!r[17] });
  }
  return out;
}


// ================================================================
//  7-1b: IDとルートインデックスで行番号を動的検索（findRowByIdAndIndex_）
//  【大A / 中7 / 小7-1b】
//  SS側で行移動・削除が発生しても現在の正しい行番号を返す
// ================================================================
function findRowByIdAndIndex_(sheet, id, routeIndex) {
  if (!id && id !== 0) return -1;
  if (routeIndex === null || routeIndex === undefined) return -1;
  var all = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0] || '').trim() === String(id).trim()) {
      if (count === Number(routeIndex)) return i + 1; // 1-indexed
      count++;
    }
  }
  return -1;
}


// ================================================================
//  7-2: 運行シートへの行作成（createParentRows）  【大A / 中7 / 小7-2】
//  紐づけアドレスからマスタ情報を取得し運行シートに行程を書き込む
//  ・同じ運行の行程は全て同じIDを付与
//  ・日付をDate型（時刻付き）で書き込む
//  ・LockServiceで同時書き込みによるID重複を防止
//  ★STEP4追加: 今日の未割当プレースホルダー（積地空）があれば新規IDを生成せずそれを使用
//    → スプレッドシートで積地が先入力済みなら別ID生成（既存動作にフォールバック）
//    → 複数行程の場合は先頭行をプレースホルダー更新、追加行程は末尾に新規追加（同ID）
// ================================================================
function createParentRows(picks, drops, dateStr, overrideInfo, companySsId, email) {
  validateDriverEmail_(email, companySsId);
  // 端末のメールアドレスを確認（未連携なら運行開始不可）
  var savedEmail = email || '';
  if (!savedEmail) throw new Error('紐づけされていません');

  // 同時に複数端末が運行開始した場合のID重複を防ぐためロック取得
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e) { throw new Error('混雑中です。少し待ってから再試行してください'); }

  try {
    var ss     = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
    // マスタからメールアドレスが一致する乗務員情報を取得
    var master = ss.getSheetByName('自車専属マスタ');
    if (!master) throw new Error('自車専属マスタシートがありません');
    var mAll = master.getDataRange().getValues();
    var info = null;
    for (var j = 1; j < mAll.length; j++) {
      if (String(mAll[j][10]).trim() === savedEmail) {
        info = {
          kubun:mAll[j][2], company:mAll[j][3], kanban:mAll[j][4],
          tons:mAll[j][5], type:mAll[j][6], car:mAll[j][7],
          name:mAll[j][8], tel:mAll[j][9]
        };
        break;
      }
    }
    if (!info) throw new Error('アドレス未登録');

    // 端末で上書き入力された車両情報を反映（マスタより優先）
    if (overrideInfo) {
      if (overrideInfo.tons   !== undefined && overrideInfo.tons   !== '') info.tons   = overrideInfo.tons;
      if (overrideInfo.type   !== undefined && overrideInfo.type   !== '') info.type   = overrideInfo.type;
      if (overrideInfo.car    !== undefined && overrideInfo.car    !== '') info.car    = overrideInfo.car;
      if (overrideInfo.name   !== undefined && overrideInfo.name   !== '') info.name   = overrideInfo.name;
      if (overrideInfo.tel    !== undefined) info.tel    = overrideInfo.tel;
      if (overrideInfo.kanban !== undefined) info.kanban = overrideInfo.kanban;
    }

    var sheet = ss.getSheetByName('運行');
    if (!sheet) throw new Error('運行シートがありません');

    // 日付: 端末から渡された場合はその日付＋現在時刻を合成（new Date(dateStr)はUTC0時=JST9時になるバグ回避）
    var cur  = new Date();
    var now  = dateStr
      ? (function() { var b = new Date(dateStr); return new Date(b.getFullYear(), b.getMonth(), b.getDate(), cur.getHours(), cur.getMinutes(), cur.getSeconds()); })()
      : cur;
    var nowY = now.getFullYear(), nowM = now.getMonth(), nowD = now.getDate();

    // 今日の未割当プレースホルダーを探す（同ドライバー・同日・積地空）
    // LockService取得済みのため読み取り後の書き込みが排他保証される
    var allRows  = sheet.getDataRange().getValues();
    var phRowNum = -1; // プレースホルダーの1-indexed行番号
    var sameId   = null;
    for (var k = 1; k < allRows.length; k++) {
      var rv  = allRows[k];
      var rid = String(rv[0] || '').trim();
      if (!rid) continue;
      var rDate = rv[9]; // col10(J列=日付) 0-indexed
      if (!(rDate instanceof Date)) continue;
      if (rDate.getFullYear() !== nowY || rDate.getMonth() !== nowM || rDate.getDate() !== nowD) continue;
      if (String(rv[6] || '').trim() !== info.name) continue; // 乗務員名で絞り込み
      if (String(rv[11] || '').trim() === '') {               // 積地空 = 未割当プレースホルダー
        phRowNum = k + 1; // 1-indexed
        sameId   = rid;
        break;
      }
    }

    var lastRow    = sheet.getLastRow();
    var resultRows = [];

    if (phRowNum !== -1) {
      // ── プレースホルダー使用パス ──
      // picks[0]/drops[0] で既存行を上書き（積地・降地のみ。IDや日付・荷主はそのまま）
      sheet.getRange(phRowNum, 12, 1, 2).setNumberFormat('@');
      sheet.getRange(phRowNum, 12).setValue(picks[0] || '');     // 積地(L列)
      sheet.getRange(phRowNum, 13).setValue(drops[0] || '');     // 降地(M列)
      sheet.getRange(phRowNum, 12).setBackground(null);          // 配車漏れ黄色警告を解除
      resultRows.push(phRowNum);

      // picks[1..] の追加行程があれば末尾に新規追加（同一ID）
      if (picks.length > 1) {
        var addStart = lastRow + 1;
        var addData  = [];
        var addDates = [];
        for (var ai = 1; ai < picks.length; ai++) {
          addData.push([
            sameId, info.kubun, info.company, info.tons, info.type, info.car,
            info.name, info.tel, info.kanban || info.company, '', '', picks[ai], drops[ai],
            '', '', '', '', '', '', '', '', '', '', '', '', ''
          ]);
          addDates.push([now]);
        }
        sheet.getRange(addStart, 1, addData.length, 26).setValues(addData);
        sheet.getRange(addStart, 12, addData.length, 2).setNumberFormat('@');
        // 日付は1行ずつ個別setValue（バッチ書き込みが効かない問題を回避）
        for (var ad = 0; ad < addData.length; ad++) {
          sheet.getRange(addStart + ad, 10).setValue(now);
          sheet.getRange(addStart + ad, 10).setNumberFormat('yyyy/MM/dd');
        }
        var addFmls = [];
        for (var af = 0; af < addData.length; af++) {
          var afr = addStart + af;
          addFmls.push(['=IF(AND(U'+afr+'="",T'+afr+'=""),"",U'+afr+'-T'+afr+')']);
        }
        sheet.getRange(addStart, 22, addFmls.length, 1).setFormulas(addFmls);
        applyDateTimeFormat_(sheet, addStart, addData.length);
        for (var ai2 = 0; ai2 < addData.length; ai2++) resultRows.push(addStart + ai2);
      }

    } else {
      // ── 新規ID生成パス（プレースホルダーなし、または積地が既に埋まっている）──
      var nextNum  = getNextIdNum_(sheet, 'V-');
      sameId       = 'V-' + String(nextNum).padStart(4, '0');
      var startRow = lastRow + 1;
      var num      = picks.length;
      var rowsData = [];
      var dateData = [];
      for (var i = 0; i < num; i++) {
        rowsData.push([
          sameId, info.kubun, info.company, info.tons, info.type, info.car,
          info.name, info.tel, info.kanban || info.company, '', '', picks[i], drops[i],
          '', '', '', '', '', '', '', '', '', '', '', '', ''
        ]);
        dateData.push([now]);
      }
      sheet.getRange(startRow, 1, num, 26).setValues(rowsData);
      sheet.getRange(startRow, 12, num, 2).setNumberFormat('@'); // 積地・降地をテキスト書式に固定
      // 日付は1行ずつ個別setValue（バッチ書き込みが効かない問題を回避）
      for (var di = 0; di < num; di++) {
        sheet.getRange(startRow + di, 10).setValue(now);
        sheet.getRange(startRow + di, 10).setNumberFormat('yyyy/MM/dd');
      }
      var formulas = [];
      for (var fi = 0; fi < num; fi++) {
        var fr = startRow + fi;
        formulas.push(['=IF(AND(U'+fr+'="",T'+fr+'=""),"",U'+fr+'-T'+fr+')']);
      }
      sheet.getRange(startRow, 22, num, 1).setFormulas(formulas);
      applyDateTimeFormat_(sheet, startRow, num);
      for (var ri = 0; ri < num; ri++) resultRows.push(startRow + ri);
    }

    // 集計表を非同期で同期（遅延実行）
    delaySyncSummary_(sameId);

    // 日付順にソート（新規行が末尾に追加されているため）
    sortUnkouByDate_(companySsId);

    return { rows: resultRows, id: sameId };

  } finally {
    lock.releaseLock();
  }
}


// ================================================================
//  7-3: 誘導時刻記録（setGuideComplete）  【大A / 中7 / 小7-3】
//  IDとルートインデックスで行を動的検索してN列（14列目）に現在時刻を書き込む
// ================================================================
function setGuideComplete(id, routeIndex, companySsId) {
  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  var row = findRowByIdAndIndex_(sheet, id, routeIndex);
  if (row < 0) return;
  var cell = sheet.getRange(row, 14);
  cell.setValue(new Date());
  cell.setNumberFormat('M/d HH:mm');
  SpreadsheetApp.flush();
  if (id) delaySyncSummary_(id, ss);
}


// ================================================================
//  7-4: 積完時刻記録（setPickComplete）  【大A / 中7 / 小7-4】
//  IDとルートインデックスで行を動的検索してO列（15列目）に現在時刻を書き込む
// ================================================================
function setPickComplete(id, routeIndex, companySsId) {
  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  var row = findRowByIdAndIndex_(sheet, id, routeIndex);
  if (row < 0) return;
  var cell = sheet.getRange(row, 15);
  cell.setValue(new Date());
  cell.setNumberFormat('M/d HH:mm');
  SpreadsheetApp.flush();
  if (id) delaySyncSummary_(id, ss);
}


// ================================================================
//  7-5: 休憩開始・終了時刻記録（setRest）  【大A / 中7 / 小7-5】
//  IDとルートインデックスで行を動的検索して P列(16)/Q列(17) に現在時刻を書き込む
// ================================================================
function setRest(id, routeIndex, type, companySsId) {
  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  var row = findRowByIdAndIndex_(sheet, id, routeIndex);
  if (row < 0) return;
  var col = (type === 'start') ? 16 : (type === 'end') ? 17 : 0;
  if (col) {
    var cell = sheet.getRange(row, col);
    cell.setValue(new Date());
    cell.setNumberFormat('M/d HH:mm');
  }
  SpreadsheetApp.flush();
  if (id) delaySyncSummary_(id, ss);
}


// ================================================================
//  7-6: 降完時刻記録（setDropComplete）  【大A / 中7 / 小7-6】
//  IDとルートインデックスで行を動的検索してR列（18列目）に現在時刻を書き込む
// ================================================================
function setDropComplete(id, routeIndex, companySsId) {
  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  var row = findRowByIdAndIndex_(sheet, id, routeIndex);
  if (row < 0) return;
  var cell = sheet.getRange(row, 18);
  cell.setValue(new Date());
  cell.setNumberFormat('M/d HH:mm');
  SpreadsheetApp.flush();
  if (id) delaySyncSummary_(id, ss);
}


// ================================================================
//  7-7: 状態保存＋時刻記録 一括実行（recordAction）  【大A / 中7 / 小7-7】
//  saveRunState と setXxxComplete を1回のサーバー呼び出しでまとめて実行する
//  actionType: 'guide' / 'pick' / 'restStart' / 'restEnd' / 'drop'
// ================================================================
function recordAction(actionType, id, routeIndex, stateObj, companySsId, email) {
  validateDriverEmail_(email, companySsId);
  saveRunState(stateObj, email, companySsId);
  if      (actionType === 'guide')      setGuideComplete(id, routeIndex, companySsId);
  else if (actionType === 'pick')       setPickComplete(id, routeIndex, companySsId);
  else if (actionType === 'restStart')  setRest(id, routeIndex, 'start', companySsId);
  else if (actionType === 'restEnd')    setRest(id, routeIndex, 'end',   companySsId);
  else if (actionType === 'drop')       setDropComplete(id, routeIndex, companySsId);
  else if (actionType === 'inspBefore') setInspectionComplete_(id, 'before', companySsId);
  else if (actionType === 'inspAfter')  setInspectionComplete_(id, 'after',  companySsId);
}


// ================================================================
//  7-8: 点呼完了時刻記録（setInspectionComplete_）  【大A / 中7 / 小7-8】
//  点呼前/点呼後の完了時刻を運行シートの該当列に書き込む（ID一致の最初の行）
// ================================================================
function setInspectionComplete_(id, type, companySsId) {
  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colName = (type === 'before') ? '点呼前完了' : '点呼後完了';
  var colIdx  = headers.indexOf(colName);
  if (colIdx < 0) return;
  var allRows = sheet.getDataRange().getValues();
  var now = new Date();
  for (var i = 1; i < allRows.length; i++) {
    if (String(allRows[i][0]||'').trim() === String(id).trim()) {
      var cell = sheet.getRange(i + 1, colIdx + 1);
      cell.setValue(now);
      cell.setNumberFormat('M/d HH:mm');
      break;
    }
  }
  SpreadsheetApp.flush();
  if (id) delaySyncSummary_(id, ss);
}


// ================================================================
//  7-9: 点呼時刻クリア（clearInspTime）  【大A / 中7 / 小7-9】
//  戻るボタン用：IDで点呼前/後の完了時刻をクリアして集計表を同期する
// ================================================================
function clearInspTime(id, type, companySsId, email) {
  validateDriverEmail_(email, companySsId);
  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet || !id) return;
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colName = (type === 'before') ? '点呼前完了' : '点呼後完了';
  var colIdx  = headers.indexOf(colName);
  if (colIdx < 0) return;
  var allRows = sheet.getDataRange().getValues();
  for (var i = 1; i < allRows.length; i++) {
    if (String(allRows[i][0]||'').trim() === String(id).trim()) {
      sheet.getRange(i + 1, colIdx + 1).clearContent();
      break;
    }
  }
  SpreadsheetApp.flush();
  if (id) delaySyncSummary_(id, ss);
}


// ================================================================
//  8-1: 行程データ更新（updateRouteData）  【大A / 中8 / 小8-1】
//  戻るボタン用：IDで行を動的検索してL列（積地）・M列（降地）を更新し集計表を同期する
// ================================================================
function updateRouteData(id, picks, drops, companySsId) {
  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  var ri = 0;
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0] || '').trim() !== String(id).trim()) continue;
    if (ri < picks.length) {
      sheet.getRange(i + 1, 12).setValue(picks[ri] || '');
      sheet.getRange(i + 1, 13).setValue(drops[ri] || '');
      ri++;
    }
  }
  if (id) delaySyncSummary_(id, ss);
}


// ================================================================
//  8-2: 運行シート行削除（deleteRunRows）  【大A / 中8 / 小8-2】
//  戻るボタン用：IDで全行を動的検索して降順に削除し集計表を同期する
// ================================================================
function deleteRunRows(id, companySsId, email) {
  validateDriverEmail_(email, companySsId);
  if (!id) return;
  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  var del = [];
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0] || '').trim() === String(id).trim()) del.push(i + 1);
  }
  del.sort(function(a, b) { return b - a; });
  for (var j = 0; j < del.length; j++) sheet.deleteRow(del[j]);
  if (id) delaySyncSummary_(id, ss);
}


// ================================================================
//  8-3: 時刻セルクリア（clearTimeCell）  【大A / 中8 / 小8-3】
//  戻るボタン用：IDとルートインデックスで行を動的検索して指定列をクリアし集計表を同期する
// ================================================================
function clearTimeCell(id, routeIndex, col, companySsId, email) {
  validateDriverEmail_(email, companySsId);
  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  var row = findRowByIdAndIndex_(sheet, id, routeIndex);
  if (row < 0) return;
  sheet.getRange(row, col).clearContent();
  if (id) delaySyncSummary_(id, ss);
}


// ================================================================
//  8-4: 運行一覧データ取得（getListData）  【大A / 中8 / 小8-4】
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
function getListData(year, month, companySsId, email) {
  // 端末に紐付いたメールアドレスを取得（未連携なら空リストを返す）
  var savedEmail = email || '';
  if (!savedEmail) return {rows:[], total:{days:0,sales:0,toll:0,pay:0}};

  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  // マスタからメールアドレスが一致する乗務員名を特定（60秒キャッシュ）
  var listCache = CacheService.getUserCache();
  var myName = '';
  var emailKey = 'driver_name_' + savedEmail.toLowerCase().replace(/[^a-z0-9]/g,'_');
  try { myName = listCache.get(emailKey) || ''; } catch(e) {}
  if (!myName) {
    var master = ss.getSheetByName('自車専属マスタ');
    var mAll   = master ? master.getDataRange().getValues() : [];
    for (var j = 1; j < mAll.length; j++) {
      if (String(mAll[j][10]).trim().toLowerCase() === savedEmail.toLowerCase()) {
        myName = String(mAll[j][8]).trim();
        break;
      }
    }
    if (myName) { try { listCache.put(emailKey, myName, 60); } catch(e) {} }
  }
  if (!myName) return {rows:[], total:{days:0,sales:0,toll:0,pay:0}};

  var sheet = ss.getSheetByName('運行');
  if (!sheet) return {rows:[], total:{days:0,sales:0,toll:0,pay:0}};

  var all      = sheet.getDataRange().getValues();
  var lastRow  = sheet.getLastRow();
  // W列(23)はリッチテキスト（クリック可能URLラベル）で格納されることがある
  // note→リッチテキストリンク→プレーンテキストの順に3段階フォールバックしてURLを取得
  var notes23  = lastRow >= 2 ? sheet.getRange(2, 24, lastRow-1, 1).getNotes() : [];
  var rtvs23   = lastRow >= 2 ? sheet.getRange(2, 24, lastRow-1, 1).getRichTextValues() : [];
  // 集計表からID単位の金額マップを作成（売上/高速/支払いは集計表の計算済み値を使う）
  var sumSheet = ss.getSheetByName('集計表');
  var payMap   = {};
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var sumAll = sumSheet.getRange(2, 1, sumSheet.getLastRow()-1, 37).getValues();
    for (var s = 0; s < sumAll.length; s++) {
      var sid = String(sumAll[s][0]||'').trim();
      if (sid) payMap[sid] = {
        sales:    Math.round(Number(sumAll[s][18])) || 0,
        tollReq:  Math.round(Number(sumAll[s][19])) || 0,
        tollReal: Math.round(Number(sumAll[s][20])) || 0,
        tollTotal:Math.round(Number(sumAll[s][21])) || 0,
        pay:      Math.round(Number(sumAll[s][26])) || 0,
        expense:  Math.round(Number(sumAll[s][27])) || 0,  // col28=AB=経費合計
        yukyu:    Math.round(Number(sumAll[s][33])) || 0,  // col34=AH=有休手当
        other:    Math.round(Number(sumAll[s][34])) || 0,  // col35=AI=その他手当
        inspBefore: sumAll[s][35] || '',  // col36=AJ=点呼前完了
        inspAfter:  sumAll[s][36] || ''   // col37=AK=点呼後完了
      };
    }
  }

  // 運行シートを走査して指定年月・自分の行だけ抽出、ID単位にデータ集約
  var idMap = {}, idOrder = [];
  for (var i = 1; i < all.length; i++) {
    var r  = all[i];
    if (!r[9]) continue;
    var dv = r[9] instanceof Date ? r[9] : new Date(r[9]);
    if (isNaN(dv.getTime())) continue;
    var dvYear = dv.getFullYear(), dvMonth = dv.getMonth()+1, dvDate = dv.getDate();
    if (dvYear !== year || dvMonth !== month) continue;
    if (String(r[6]).trim() !== myName) continue;
    var id = String(r[0]||'').trim();
    if (!id) continue;

    if (!idMap[id]) {
      var ds = dvYear+'/'+String(dvMonth).padStart(2,'0')+'/'+String(dvDate).padStart(2,'0');
      // ★dateSort はJ列（初回行程登録時刻）を基準に固定・以後変えない
      var baseDateSort = dv.getTime();
      var n23 = (notes23[i-1] && notes23[i-1][0]) || '';
      var du23 = n23 ? n23.split('\n').filter(function(u){return u.match(/^https?:\/\//);}) : [];
      if (!du23.length && rtvs23[i-1] && rtvs23[i-1][0]) {
        var rtv23 = rtvs23[i-1][0], rruns = rtv23.getRuns();
        for (var k = 0; k < rruns.length; k++) { var lk=rruns[k].getLinkUrl(); if(lk) du23.push(lk); }
      }
      if (!du23.length) { var pu=String(r[23]||''); if(pu.match(/^https?:\/\//)) du23=[pu]; }
      idMap[id] = {
        id:id, car:String(r[5]||'').trim(), date:ds,
        dateSort: baseDateSort,
        dateDisp:'', picks:[], drops:[],
        guideTime:'', pickTime:'', restStart:'', restEnd:'', dropTime:'',
        sales:0, tollReq:0, tollReal:0, tollTotal:0, pay:0, expense:0, yukyu:0, other:0,
        inspBefore:'', inspAfter:'',
        notice:r[22]||'', dataUrls:du23, dataUrl:du23[0]||'',
        hasNotice:!!(r[22]||du23.length),
        _rawDv: dv
      };
      idOrder.push(id);
    }
    var g = idMap[id];
    if (r[11]) g.picks.push(r[11]);
    if (r[12]) g.drops.push(r[12]);
    if (r[13] && !g.guideTime) {
      var gt = r[13] instanceof Date ? r[13] : new Date(r[13]);
      if (!isNaN(gt.getTime())) g.guideTime = String(gt.getHours()).padStart(2,'0')+':'+String(gt.getMinutes()).padStart(2,'0');
    }
    if (r[14] && !g.pickTime) {
      var pt = r[14] instanceof Date ? r[14] : new Date(r[14]);
      if (!isNaN(pt.getTime())) {
        g.pickTime = String(pt.getHours()).padStart(2,'0')+':'+String(pt.getMinutes()).padStart(2,'0');
        // ★積完時刻でdateSortを上書きしない
      }
    }
    if (r[15] && !g.restStart) {
      var rst = r[15] instanceof Date ? r[15] : new Date(r[15]);
      if (!isNaN(rst.getTime())) g.restStart = String(rst.getHours()).padStart(2,'0')+':'+String(rst.getMinutes()).padStart(2,'0');
    }
    if (r[16] && !g.restEnd) {
      var re2 = r[16] instanceof Date ? r[16] : new Date(r[16]);
      if (!isNaN(re2.getTime())) g.restEnd = String(re2.getHours()).padStart(2,'0')+':'+String(re2.getMinutes()).padStart(2,'0');
    }
    if (r[17] && !g.dropTime) {
      var dt2 = r[17] instanceof Date ? r[17] : new Date(r[17]);
      if (!isNaN(dt2.getTime())) g.dropTime = String(dt2.getHours()).padStart(2,'0')+':'+String(dt2.getMinutes()).padStart(2,'0');
    }
    g.sales   += Number(r[18]) || 0;
    g.tollReq += Number(r[19]) || 0;
    g.tollReal+= Number(r[20]) || 0;
    if (r[22] && !g.notice) g.notice = r[22];
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
      g.expense  = pm.expense || 0;
      g.yukyu    = pm.yukyu || 0;
      g.other    = pm.other || 0;
      if (pm.inspBefore) { var ibt = pm.inspBefore instanceof Date ? pm.inspBefore : new Date(pm.inspBefore); if (!isNaN(ibt.getTime())) g.inspBefore = String(ibt.getHours()).padStart(2,'0')+':'+String(ibt.getMinutes()).padStart(2,'0'); }
      if (pm.inspAfter)  { var iat = pm.inspAfter  instanceof Date ? pm.inspAfter  : new Date(pm.inspAfter);  if (!isNaN(iat.getTime())) g.inspAfter  = String(iat.getHours()).padStart(2,'0')+':'+String(iat.getMinutes()).padStart(2,'0'); }
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
      pay:g.pay, expense:g.expense, yukyu:g.yukyu, other:g.other,
      inspBefore:g.inspBefore, inspAfter:g.inspAfter,
      notice:g.notice, dataUrl:g.dataUrl, hasNotice:g.hasNotice,
      isComplete: !!(g.guideTime && g.pickTime && g.restStart && g.restEnd && g.dropTime && g.inspBefore && g.inspAfter),
      isNew:      !g.guideTime && !g.pickTime && !g.restStart && !g.restEnd && !g.dropTime && !g.inspBefore && !g.inspAfter
    });
    totalSales += g.sales; totalToll += g.tollTotal; totalPay += g.pay;
    totalYukyu += g.yukyu || 0;
    totalOther += g.other || 0;
    if (g.picks.join('').trim() !== '') dateSet[g.date] = true; // 積地空（未配車）はノーカウント
  }
  result.sort(function(a,b){ return b.dateSort - a.dateSort; });
  return { rows:result, total:{ days:Object.keys(dateSet).length, sales:totalSales, toll:totalToll, pay:totalPay, yukyu:totalYukyu, other:totalOther, yukyuDays:yukyuDays, yasumiDays:yasumiDays } };
}


// ================================================================
//  8-5: 編集用データ取得（getEditData）  【大A / 中8 / 小8-5】
//  編集モーダルに表示する1件分の詳細データを取得して返す
//  ・同一IDの複数行（複数行程）は売上/高速を合算して返す
//  ・時刻（誘導/積完/休憩/降完）は最初に見つかった値を使用（先勝ち）
//  ・集計表から合計高速代・利益を取得して付加
//  ・W列(23)のURL：getAdminDataUrl_（リッチテキスト→URLカンマ区切り）
//  ・Y列(25)のURL：getTerminalUrls_（リッチテキスト→URL配列.join(',')）
// ================================================================
function getEditData(id, companySsId, email) {
  validateDriverEmail_(email, companySsId);
  var ss    = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return null;

  var sumSheet = ss.getSheetByName('集計表');
  var sumData  = { tollTotal:'', profit:'', yukyu:'', other:'', inspBefore:'', inspAfter:'' };
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var sumAll = sumSheet.getRange(2, 1, sumSheet.getLastRow()-1, 37).getValues();
    for (var s = 0; s < sumAll.length; s++) {
      if (String(sumAll[s][0]||'').trim() === String(id).trim()) {
        sumData.tollTotal = sumAll[s][21] !== '' ? sumAll[s][21] : '';
        sumData.profit    = sumAll[s][27] !== '' ? sumAll[s][27] : '';
        sumData.yukyu     = sumAll[s][33] !== '' ? Math.round(Number(sumAll[s][33])) : '';  // AH=有休手当
        sumData.other     = sumAll[s][34] !== '' ? Math.round(Number(sumAll[s][34])) : '';  // AI=その他手当
        var ibRaw = sumAll[s][35]; // AJ=点呼前完了
        var iaRaw = sumAll[s][36]; // AK=点呼後完了
        if (ibRaw) { var ibd = ibRaw instanceof Date ? ibRaw : new Date(ibRaw); if (!isNaN(ibd.getTime())) sumData.inspBefore = String(ibd.getHours()).padStart(2,'0')+':'+String(ibd.getMinutes()).padStart(2,'0'); }
        if (iaRaw) { var iad = iaRaw instanceof Date ? iaRaw : new Date(iaRaw); if (!isNaN(iad.getTime())) sumData.inspAfter  = String(iad.getHours()).padStart(2,'0')+':'+String(iad.getMinutes()).padStart(2,'0'); }
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
    totalSales   += Number(r[18]) || 0;
    totalTollReq += Number(r[19]) || 0;
    totalTollReal+= Number(r[20]) || 0;

    // 時刻は最初に見つかった値を使用
    if (!guideTime && r[13]) guideTime = Utilities.formatDate(new Date(r[13]),'Asia/Tokyo','HH:mm');
    if (!pickTime  && r[14]) pickTime  = Utilities.formatDate(new Date(r[14]),'Asia/Tokyo','HH:mm');
    if (!restStart && r[15]) restStart = Utilities.formatDate(new Date(r[15]),'Asia/Tokyo','HH:mm');
    if (!restEnd   && r[16]) restEnd   = Utilities.formatDate(new Date(r[16]),'Asia/Tokyo','HH:mm');
    if (!dropTime  && r[17]) dropTime  = Utilities.formatDate(new Date(r[17]),'Asia/Tokyo','HH:mm');
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
    kanban:   baseData[8] || baseData[2],
    date:     baseData[9] ? Utilities.formatDate(new Date(baseData[9]),'Asia/Tokyo','yyyy-MM-dd') : '',
    client:   baseData[10] || '',
    pick:     baseData[11] || '',
    drop:     baseData[12] || '',
    guideTime: guideTime,
    pickTime:  pickTime,
    restStart: restStart,
    restEnd:   restEnd,
    dropTime:  dropTime,
    sales:    totalSales    || '',
    tollReq:  totalTollReq  || '',
    tollReal: totalTollReal || '',
    tollTotal: sumData.tollTotal,
    notice:   baseData[22] || '',
    dataUrl:  getAdminDataUrl_(sheet, firstRow),
    termNotice:baseData[24]|| '',
    termData: getTerminalUrls_(sheet, firstRow).join(','),
    profit:     sumData.profit,
    yukyu:      sumData.yukyu,
    other:      sumData.other,
    inspBefore: sumData.inspBefore || '',
    inspAfter:  sumData.inspAfter  || ''
  };
}


// ================================================================
//  8-6: 編集データ保存（saveEditData）  【大A / 中8 / 小8-6】
//  端末アプリの編集モーダルで変更された値を運行シートに書き込む
//  ・日付はDate型で書き込む（文字列だとonEditUnkou_が誤発火するため）
//  ・荷主名/積地/降地は undefined/null でなければ上書きする（空文字でも書き込み可）
//  ・時刻は「日付＋時刻」を合成したDate型で書き込み、空なら clearContent する
//  ・売上/高速は同一IDに複数行ある場合、先頭行のみ書き込み（重複防止）
//  ・書き込み後にdelaySyncSummary_を呼んで集計表を同期する
// ================================================================
function saveEditData(obj, companySsId, email) {
  validateDriverEmail_(email, companySsId);
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try {
  var ss    = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  // origIdが提供されIDが変更された場合は元IDで行を検索する
  var idChanged = !!(obj.origId && obj.id && String(obj.origId).trim() !== String(obj.id).trim());
  var searchId  = idChanged ? obj.origId : obj.id;
  // ヘッダーから点呼列インデックスを取得
  var uHdr0     = all[0].map(function(h){ return String(h||'').trim(); });
  var inspBColIdx = uHdr0.indexOf('点呼前完了');
  var inspAColIdx = uHdr0.indexOf('点呼後完了');
  var written = false;
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() !== String(searchId).trim()) continue;
    var r = i + 1;

    // IDが変更された場合はA列を更新
    if (idChanged) sheet.getRange(r, 1).setValue(obj.id);

    // 車両情報（同一IDの全行に書き込む）
    if (obj.kubun   !== undefined) sheet.getRange(r, 2).setValue(obj.kubun   || '');
    if (obj.company !== undefined) sheet.getRange(r, 3).setValue(obj.company || '');
    if (obj.tons    !== undefined) sheet.getRange(r, 4).setValue(obj.tons    || '');
    if (obj.type    !== undefined) sheet.getRange(r, 5).setValue(obj.type    || '');
    if (obj.car     !== undefined) sheet.getRange(r, 6).setValue(obj.car     || '');
    if (obj.name    !== undefined) sheet.getRange(r, 7).setValue(obj.name    || '');
    if (obj.tel     !== undefined) sheet.getRange(r, 8).setValue(obj.tel     || '');
    if (obj.kanban  !== undefined) sheet.getRange(r, 9).setValue(obj.kanban  || '');

    // ★日付はDate型で書き込む（文字列だとonEditが誤発火する）
    if (obj.date) {
      var d = new Date(obj.date);
      // 既存のJ列の時刻部分を保持する
      var existing = all[i][9];
      if (existing instanceof Date) {
        d = new Date(d.getFullYear(), d.getMonth(), d.getDate(),
                     existing.getHours(), existing.getMinutes(), existing.getSeconds());
      }
      sheet.getRange(r, 10).setValue(d);
    }

    // client/pick/dropはnullでなければ書き込む（空でも上書き可）
    if (obj.client !== undefined && obj.client !== null) sheet.getRange(r, 11).setValue(obj.client);
    if (obj.pick   !== undefined && obj.pick   !== null) sheet.getRange(r, 12).setValue(obj.pick);
    if (obj.drop   !== undefined && obj.drop   !== null) sheet.getRange(r, 13).setValue(obj.drop);

    var timeFmt = 'M/d HH:mm';
    var rowDateVal = all[i][9];
    var dateStr = obj.date || (rowDateVal instanceof Date ? Utilities.formatDate(rowDateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(rowDateVal || ''));
    if (obj.guideTime !== undefined) { if (obj.guideTime) { var c14=sheet.getRange(r,14); c14.setValue(new Date(dateStr+' '+obj.guideTime)); c14.setNumberFormat(timeFmt); } else sheet.getRange(r,14).clearContent(); }
    if (obj.pickTime  !== undefined) { if (obj.pickTime)  { var c15=sheet.getRange(r,15); c15.setValue(new Date(dateStr+' '+obj.pickTime));  c15.setNumberFormat(timeFmt); } else sheet.getRange(r,15).clearContent(); }
    if (obj.restStart !== undefined) { if (obj.restStart) { var c16=sheet.getRange(r,16); c16.setValue(new Date(dateStr+' '+obj.restStart)); c16.setNumberFormat(timeFmt); } else sheet.getRange(r,16).clearContent(); }
    if (obj.restEnd   !== undefined) { if (obj.restEnd)   { var c17=sheet.getRange(r,17); c17.setValue(new Date(dateStr+' '+obj.restEnd));   c17.setNumberFormat(timeFmt); } else sheet.getRange(r,17).clearContent(); }
    if (obj.dropTime  !== undefined) { if (obj.dropTime)  { var c18=sheet.getRange(r,18); c18.setValue(new Date(dateStr+' '+obj.dropTime));  c18.setNumberFormat(timeFmt); } else sheet.getRange(r,18).clearContent(); }

    // 売上・高速・点呼前後は最初の行のみ書き込む
    if (!written) {
      if (obj.sales    !== undefined) sheet.getRange(r, 19).setValue(obj.sales    || '');
      if (obj.tollReq  !== undefined) sheet.getRange(r, 20).setValue(obj.tollReq  || '');
      if (obj.tollReal !== undefined) sheet.getRange(r, 21).setValue(obj.tollReal || '');
      var timeFmt2 = 'M/d HH:mm';
      if (obj.inspBefore !== undefined && inspBColIdx >= 0) {
        if (obj.inspBefore) {
          var bd = all[i][9]; var inspBDate = bd instanceof Date ? new Date(bd) : new Date(bd);
          var bp = String(obj.inspBefore).split(':'); if (bp.length >= 2) inspBDate.setHours(parseInt(bp[0]), parseInt(bp[1]), 0, 0);
          sheet.getRange(r, inspBColIdx+1).setValue(inspBDate).setNumberFormat(timeFmt2);
        } else { sheet.getRange(r, inspBColIdx+1).clearContent(); }
      }
      if (obj.inspAfter !== undefined && inspAColIdx >= 0) {
        if (obj.inspAfter) {
          var ad = all[i][9]; var inspADate = ad instanceof Date ? new Date(ad) : new Date(ad);
          var ap = String(obj.inspAfter).split(':'); if (ap.length >= 2) inspADate.setHours(parseInt(ap[0]), parseInt(ap[1]), 0, 0);
          sheet.getRange(r, inspAColIdx+1).setValue(inspADate).setNumberFormat(timeFmt2);
        } else { sheet.getRange(r, inspAColIdx+1).clearContent(); }
      }
      written = true;
    } else {
      if (obj.tollReq  !== undefined) sheet.getRange(r, 20).setValue('');
      if (obj.tollReal !== undefined) sheet.getRange(r, 21).setValue('');
    }

    sheet.getRange(r, 25).setValue(obj.termNotice !== undefined ? obj.termNotice : (all[i][24] || ''));
  }
  // その他手当と（必要なら）IDを集計表に反映
  if (obj.other !== undefined || idChanged) {
    var sumSheetX = ss.getSheetByName('集計表');
    if (sumSheetX && sumSheetX.getLastRow() >= 2) {
      var sumIdsX = sumSheetX.getRange(2, 1, sumSheetX.getLastRow()-1, 1).getValues();
      for (var sx = 0; sx < sumIdsX.length; sx++) {
        if (String(sumIdsX[sx][0]).trim() === String(searchId).trim()) {
          if (obj.other !== undefined) sumSheetX.getRange(sx+2, 35).setValue(Number(obj.other) || 0);
          if (idChanged) sumSheetX.getRange(sx+2, 1).setValue(obj.id);
          break;
        }
      }
    }
  }
  delaySyncSummary_(obj.id);
  try { applyHolidayRowColors_(); } catch(e) {}
  } finally {
    lock.releaseLock();
  }
}


// ================================================================
//  8-6c: 端末連絡保存（saveTermNoticeByDriver）  【大A / 中8 / 小8-6c】
//  端末アプリの一覧編集モーダルから連絡(端末)(Y列=25)のみを書き込む
// ================================================================
function saveTermNoticeByDriver(id, termNotice, companySsId) {
  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() !== String(id).trim()) continue;
    sheet.getRange(i + 1, 25).setValue(termNotice || '');
  }
}


// ================================================================
//  8-6b: シート保護設定（setupSheetProtection）  【大C / 中8 / 小8-6b】
//  集計表: 距離(V=22)・ガソリン代(X=24)・備考(AB=28)以外ロック
//  運行シート: 合計高速(U=21)列のみロック
// ================================================================
function applySheetColors_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sumSheet = ss.getSheetByName('集計表');
  if (sumSheet) {
    var editableCols = [23, 25, 27, 30, 35];
    var lastCol = Math.max(sumSheet.getLastColumn(), 34);
    sumSheet.getRange(1, 1, 1, lastCol)
      .setBackground('#37474f').setFontColor('#90a4ae').setFontWeight('bold');
    for (var ec = 0; ec < editableCols.length; ec++) {
      sumSheet.getRange(1, editableCols[ec])
        .setBackground('#1b5e20').setFontColor('#a5d6a7').setFontWeight('bold');
    }
    var lastRow = Math.max(sumSheet.getLastRow(), 2);
    for (var ec2 = 0; ec2 < editableCols.length; ec2++) {
      sumSheet.getRange(1, editableCols[ec2], lastRow, 1)
        .setBorder(null, true, null, true, null, null, '#4caf50', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }
  var unkouSheet = ss.getSheetByName('運行');
  if (unkouSheet) {
    var unkouLastRow = Math.max(unkouSheet.getLastRow(), 2);
    var protectedUnkouCols = [22, 25, 26];
    for (var pc = 0; pc < protectedUnkouCols.length; pc++) {
      var pcol = protectedUnkouCols[pc];
      unkouSheet.getRange(1, pcol).setBackground('#37474f').setFontColor('#90a4ae').setFontWeight('bold');
      unkouSheet.getRange(2, pcol, unkouLastRow - 1, 1).setBackground('#eceff1');
    }
    // T列(col20=請求高速)ヘッダーは編集可のため保護色をリセット（データ行はapplyHolidayRowColors_に委ねる）
    unkouSheet.getRange(1, 20).setBackground(null).setFontColor(null).setFontWeight('normal');
    // 点呼前後列：ロックなし・入力可のため白にする
    var uLastCol = unkouSheet.getLastColumn();
    if (uLastCol > 0) {
      var uHdrs = unkouSheet.getRange(1, 1, 1, uLastCol).getValues()[0];
      var ubIdx = uHdrs.indexOf('点呼前完了'), uaIdx = uHdrs.indexOf('点呼後完了');
      if (ubIdx >= 0) {
        unkouSheet.getRange(1, ubIdx + 1).setBackground(null).setFontColor(null).setFontWeight('normal');
        if (unkouLastRow > 1) unkouSheet.getRange(2, ubIdx + 1, unkouLastRow - 1, 1).setBackground(null);
      }
      if (uaIdx >= 0) {
        unkouSheet.getRange(1, uaIdx + 1).setBackground(null).setFontColor(null).setFontWeight('normal');
        if (unkouLastRow > 1) unkouSheet.getRange(2, uaIdx + 1, unkouLastRow - 1, 1).setBackground(null);
      }
    }
  }
  // 有休/休み行の着色を再適用（保護色設定で上書きされないよう最後に実行）
  applyHolidayRowColors_();
}

function applySumEditableBorders_(sumSheet, startRow, numRows) {
  if (!sumSheet || numRows < 1 || startRow <= 1) return;
  var editCols = [23, 25, 27, 30, 35]; // W=距離, Y=ガソリン代, AA=支払い, AD=備考, AI=その他手当
  for (var c = 0; c < editCols.length; c++) {
    // 先に右隣の左枠をクリア（後でeditColの右枠が最後に書かれて勝つ）
    sumSheet.getRange(startRow, editCols[c] + 1, numRows, 1)
      .setBorder(null, false, null, null, null, null);
    sumSheet.getRange(1, editCols[c] + 1, 1, 1)
      .setBorder(null, false, null, null, null, null);
    sumSheet.getRange(startRow, editCols[c], numRows, 1)
      .setBorder(true, true, true, true, null, true, '#ffd600', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
}

function setupSheetProtection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var sumSheet = ss.getSheetByName('集計表');
  if (sumSheet) {
    sumSheet.setFrozenColumns(1);
    sumSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p){p.remove();});
    sumSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){p.remove();});
    var sp = sumSheet.protect().setDescription('集計表保護');
    sp.setUnprotectedRanges([
      sumSheet.getRange('W2:W2000'),
      sumSheet.getRange('Y2:Y2000'),
      sumSheet.getRange('AA2:AA2000'),
      sumSheet.getRange('AD2:AD2000'),
      sumSheet.getRange('AI2:AI2000')
    ]);
    // 1行目の黄色枠滲み（editCols+1の左枠）をクリア
    [24, 26, 28, 31, 36].forEach(function(bc) {
      sumSheet.getRange(1, bc).setBorder(null, false, null, null, null, null);
    });
  }
  var unkouSheet = ss.getSheetByName('運行');
  if (unkouSheet) {
    unkouSheet.setFrozenColumns(1);
    unkouSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p){p.remove();});
    // Y列(連絡端末)・Z列(データ端末)：シートレベルで保護（手動入力をGoogleが直接ブロック）
    // GASスクリプトはオーナー権限で実行するため書き込み可
    var termProtect = unkouSheet.getRange('Y2:Z').protect().setDescription('端末列保護');
    termProtect.removeEditors(termProtect.getEditors());
    if (termProtect.canDomainEdit()) termProtect.setDomainEdit(false);
  }
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var spLr = sumSheet.getLastRow();
    var spIds = sumSheet.getRange(2, 1, spLr - 1, 1).getValues();
    var spEditCols = [23, 25, 27, 30, 35];
    for (var spi = 0; spi < spIds.length; spi++) {
      if (!spIds[spi][0]) continue;
      for (var spc = 0; spc < spEditCols.length; spc++) {
        sumSheet.getRange(spi + 2, spEditCols[spc])
          .setBorder(true, true, true, true, null, null, '#ffd600', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    }
  }
  applySheetColors_(ss);
  ui.alert('保護設定完了\n■ 集計表\n  編集可: 距離(W)・ガソリン代(Y)・支払い(AA)・備考(AD)・その他手当(AI)\n  黄色枠線＝手入力可 / 灰色ヘッダー = 保護\n■ 運行シート: V(合計高速)・Y(連絡端末)・Z(データ端末) をグレー着色');
}


// ================================================================
//  8-6b-0b: 看板名列を既存シートに挿入（insertKanbanColumn）  【大C / 中8 / 小8-6b-0b】
//  メニューから1回だけ実行する。I列（col9）に看板名を追加し
//  既存行は会社名(C列)をデフォルト値として埋める。
// ================================================================
function insertKanbanColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // 運行シートに挿入
  var unkouSheet = ss.getSheetByName('運行');
  if (unkouSheet) {
    var unkouHeader = unkouSheet.getLastRow() >= 1 ? String(unkouSheet.getRange(1, 9).getValue()).trim() : '';
    if (unkouHeader === '看板名') {
      ui.alert('運行シートにはすでに看板名列があります。');
    } else {
      unkouSheet.insertColumnBefore(9);
      unkouSheet.getRange(1, 9).setValue('看板名');
      var ulr = unkouSheet.getLastRow();
      if (ulr >= 2) {
        var companyVals = unkouSheet.getRange(2, 3, ulr - 1, 1).getValues();
        var kanbanVals = companyVals.map(function(r) { return [r[0] || '']; });
        unkouSheet.getRange(2, 9, ulr - 1, 1).setValues(kanbanVals);
      }
    }
  }

  // 集計表に挿入
  var sumSheet = ss.getSheetByName('集計表');
  if (sumSheet) {
    var sumHeader = sumSheet.getLastRow() >= 1 ? String(sumSheet.getRange(1, 9).getValue()).trim() : '';
    if (sumHeader === '看板名') {
      ui.alert('集計表にはすでに看板名列があります。');
    } else {
      sumSheet.insertColumnBefore(9);
      sumSheet.getRange(1, 9).setValue('看板名');
      var slr = sumSheet.getLastRow();
      if (slr >= 2) {
        var sCompanyVals = sumSheet.getRange(2, 3, slr - 1, 1).getValues();
        var sKanbanVals = sCompanyVals.map(function(r) { return [r[0] || '']; });
        sumSheet.getRange(2, 9, slr - 1, 1).setValues(sKanbanVals);
      }
    }
  }

  ui.alert('看板名列の挿入が完了しました。\n次に「シート保護設定」を実行してください。');
}


// ================================================================
//  8-6b-1: 端末ファイルURL一覧取得（getTerminalUrls_）  【大B / 中8 / 小8-6b-1】
//  col25のリッチテキストからリンクURLを配列で返す
// ================================================================
function getTerminalUrls_(sheet, rowNum) {
  var cell = sheet.getRange(rowNum, 26);
  var urls = [];

  // 1. リッチテキストのrunからURL取得（アプリ経由でアップした場合）
  var rtv = cell.getRichTextValue();
  if (rtv) {
    var runs = rtv.getRuns();
    for (var i = 0; i < runs.length; i++) {
      var link = runs[i].getLinkUrl();
      if (link && urls.indexOf(link) === -1) urls.push(link);
    }
  }

  // 2. セルノートからURL取得
  if (urls.length === 0) {
    var note = cell.getNote();
    if (note) {
      note.split('\n').forEach(function(u) {
        u = u.trim();
        if (u.match(/^https?:\/\//) && urls.indexOf(u) === -1) urls.push(u);
      });
    }
  }

  // 3. HYPERLINK数式からURL取得（SS上でHYPERLINK関数を使った場合）
  if (urls.length === 0) {
    var formula = cell.getFormula();
    if (formula) {
      var re = /HYPERLINK\s*\(\s*"([^"]+)"/gi;
      var m;
      while ((m = re.exec(formula)) !== null) {
        if (m[1].match(/^https?:\/\//) && urls.indexOf(m[1]) === -1) urls.push(m[1]);
      }
    }
  }

  // 4. プレーンテキストURL（直接URLを入力した場合）
  if (urls.length === 0) {
    var plain = String(cell.getValue() || '').trim();
    if (plain) {
      plain.split(/[\n,]/).forEach(function(u) {
        u = u.trim();
        if (u.match(/^https?:\/\//) && urls.indexOf(u) === -1) urls.push(u);
      });
    }
  }

  return urls;
}


// ================================================================
//  8-6b-2: 端末ファイルURL一覧書込（setTerminalUrls_）  【大B / 中8 / 小8-6b-2】
//  URLをリッチテキスト（ファイル1, ファイル2…）として col25 に書込む
// ================================================================
function setTerminalUrls_(sheet, rowNum, urls) {
  var cell = sheet.getRange(rowNum, 26);
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
  cell.setNote(urls.join('\n'));
}


// ================================================================
//  8-6b-0: 画像URLをDriveに取込（importImageToDrive_）  【大B / 中8 / 小8-6b-0】
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
//  8-6b-3: 管理側ファイルURLをリッチテキストで書込（setAdminDataRichText_）  【大B / 中8 / 小8-6b-3】
// ================================================================
function setAdminDataRichText_(sheet, rowNum, url) {
  setAdminDataRichTextMulti_(sheet, rowNum, url ? [url] : []);
}


// ================================================================
//  8-6b-3b: 管理側ファイル複数URLをリッチテキストで書込（setAdminDataRichTextMulti_）  【大B / 中8 / 小8-6b-3b】
// ================================================================
function setAdminDataRichTextMulti_(sheet, rowNum, urls) {
  var cell = sheet.getRange(rowNum, 24);
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
//  8-6b-4: 管理側ファイルURLをリッチテキストから取得（getAdminDataUrl_）  【大B / 中8 / 小8-6b-4】
// ================================================================
function getAdminDataUrl_(sheet, rowNum) {
  var rtv = sheet.getRange(rowNum, 24).getRichTextValue();
  if (rtv) {
    var runs = rtv.getRuns(), urls = [];
    for (var i = 0; i < runs.length; i++) {
      var link = runs[i].getLinkUrl();
      if (link) urls.push(link);
    }
    if (urls.length > 0) return urls.join(',');
  }
  return String(sheet.getRange(rowNum, 24).getValue() || '');
}


// ================================================================
//  8-6c: 端末ファイル追加（appendTerminalFile）  【大A / 中8 / 小8-6c】
//  ファイルをDriveに保存しcol25のリッチテキストURLに追記する
// ================================================================
function appendTerminalFile(id, fileName, base64Data, mimeType, companySsId, email) {
  validateDriverEmail_(email, companySsId);
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try {
    var folder  = getOrCreateFolder_('端末データ');
    var decoded = Utilities.base64Decode(base64Data);
    var blob    = Utilities.newBlob(decoded, mimeType, fileName);
    var file    = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url   = file.getUrl();
    var ss    = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('運行');
    if (!sheet) return { ok: false };
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
  } finally {
    lock.releaseLock();
  }
}


// ================================================================
//  8-6c-2: 管理側ファイル追加・削除・差替（ID指定）  【大A / 中8 / 小8-6c-2】
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
//  8-6d: 端末ファイル削除（deleteTerminalFile）  【大A / 中8 / 小8-6d】
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
//  8-6e: 端末ファイル差し替え（replaceTerminalFile）  【大A / 中8 / 小8-6e】
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
//  8-7: 運行データ削除（deleteRunById）  【大A / 中8 / 小8-7】
//  指定IDに一致する運行シートの全行を削除し集計表を同期する
// ================================================================
function deleteRunById(id, companySsId, email) {
  validateDriverEmail_(email, companySsId);
  var ss    = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var all     = sheet.getDataRange().getValues();
  var delRows = [];
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() === String(id).trim()) delRows.push(i+1);
  }
  delRows.sort(function(a,b){ return b-a; });
  for (var i = 0; i < delRows.length; i++) sheet.deleteRow(delRows[i]);
  delaySyncSummary_(id, ss);
}


// ================================================================
//  9-1: 連絡事項保存（saveNotice）  【大A / 中9 / 小9-1】
//  指定IDの運行シートU列（21列目）にテキストを書き込む
// ================================================================
function saveNotice(id, text, companySsId, email) {
  validateDriverEmail_(email, companySsId);
  var ss    = getTargetSS_(companySsId);
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() === String(id).trim()) {
      sheet.getRange(i+1, 23).setValue(text); break;
    }
  }
}


// ================================================================
//  9-2: ファイルアップロード・管理側（uploadFile）  【大A / 中9 / 小9-2】
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
//  9-2b: シートボタン用ファイルアップロードダイアログ（openFileUploadDialog）  【大C / 中9 / 小9-2b】
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
    '<button id="upbtn" onclick="upload()" style="padding:10px 24px;background:#1565c0;color:white;border:none;border-radius:8px;font-size:15px;cursor:pointer;">アップロード</button>' +
    '<p id="msg" style="margin-top:12px;color:#aaa;"></p>' +
    '<script>' +
    'function upload(){' +
    '  var files=Array.from(document.getElementById("f").files);' +
    '  if(!files.length){alert("ファイルを選択してください");return;}' +
    '  document.getElementById("upbtn").disabled=true;' +
    '  document.getElementById("msg").innerText="読み込み中...";' +
    '  var queued=0,done=0,total=files.length;' +
    '  files.forEach(function(file){' +
    '    if(file.size>10*1024*1024){queued++;done++;check();return;}' +
    '    var r=new FileReader();' +
    '    r.onload=function(){' +
    '      var b64=r.result.split(",")[1];' +
    '      queued++;' +
    '      google.script.run' +
    '        .withSuccessHandler(function(){done++;check();})' +
    '        .withFailureHandler(function(err){done++;document.getElementById("msg").innerText="エラー："+err.message;})' +
    '        .queueFileUpload(' + row + ',file.name,b64,file.type);' +
    '    };' +
    '    r.readAsDataURL(file);' +
    '  });' +
    '  function check(){if(done===total){document.getElementById("msg").innerText="✅ バックグラウンドでアップロード中。1〜2分後にSSを確認してください。";setTimeout(google.script.host.close,2500);}}' +
    '}' +
    '<\/script>' +
    '</body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(420).setHeight(220),
    'データアップロード'
  );
}


// ================================================================
//  9-2c: シートボタン用ファイルアップロード処理（uploadFileToRow）  【大A / 中9 / 小9-2c】
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
    SpreadsheetApp.flush();
  }
  return { ok: true, url: url };
}


// ================================================================
//  9-2d: アップロードキュー登録（queueFileUpload）  【大A / 中9 / 小9-2d】
//  base64をDriveの一時ファイルに保存し、時間トリガーで後処理させる
// ================================================================
function queueFileUpload(rowNum, fileName, base64Data, mimeType) {
  var folder = getOrCreateFolder_('_upload_queue_');
  var queueId = 'uq_' + Date.now() + '_' + Math.floor(Math.random() * 10000);
  var tempFile = folder.createFile(Utilities.newBlob(base64Data, 'text/plain', queueId + '.txt'));
  PropertiesService.getScriptProperties().setProperty(queueId, JSON.stringify({
    tempFileId: tempFile.getId(), rowNum: rowNum, fileName: fileName, mimeType: mimeType
  }));
  var existing = ScriptApp.getProjectTriggers().filter(function(t) {
    return t.getHandlerFunction() === 'processUploadQueue';
  });
  if (existing.length === 0) {
    ScriptApp.newTrigger('processUploadQueue').timeBased().after(90000).create();
  }
  return { queued: true };
}


// ================================================================
//  9-2e: アップロードキュー処理（processUploadQueue）  【大A / 中9 / 小9-2e】
//  時間トリガーから呼ばれ、キューに積まれたファイルを実際にアップロードする
// ================================================================
function processUploadQueue() {
  var props = PropertiesService.getScriptProperties();
  var allProps = props.getProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  var folder = getOrCreateFolder_('運行データ');
  for (var key in allProps) {
    if (key.indexOf('uq_') !== 0) continue;
    var data;
    try { data = JSON.parse(allProps[key]); } catch (e2) { props.deleteProperty(key); continue; }
    try {
      var tempFile = DriveApp.getFileById(data.tempFileId);
      var base64Str = tempFile.getBlob().getDataAsString();
      var decoded = Utilities.base64Decode(base64Str);
      var blob = Utilities.newBlob(decoded, data.mimeType, data.fileName);
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var url = file.getUrl();
      if (sheet) {
        var existing2 = getAdminDataUrl_(sheet, data.rowNum).split(',').filter(function(u) { return u.match(/^https?:\/\//); });
        existing2.push(url);
        var deduped = existing2.filter(function(u, i, arr) { return arr.indexOf(u) === i; });
        setAdminDataRichTextMulti_(sheet, data.rowNum, deduped);
      }
      tempFile.setTrashed(true);
    } catch (e3) {}
    props.deleteProperty(key);
  }
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processUploadQueue') {
      try { ScriptApp.deleteTrigger(t); } catch (e4) {}
    }
  });
}


// ================================================================
//  9-3: 端末からの連絡保存（saveTerminalNotice）  【大A / 中9 / 小9-3】
//  指定IDの運行シートW列（23列目）にテキストを書き込む
// ================================================================
function saveTerminalNotice(id, text, companySsId, email) {
  validateDriverEmail_(email, companySsId);
  var ss    = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() === String(id).trim()) {
      sheet.getRange(i+1, 25).setValue(text); break;
    }
  }
}


// ================================================================
//  9-4: 端末からのファイルアップロード（uploadTerminalFile）  【大A / 中9 / 小9-4】
//  ファイルをGoogleドライブの「端末データ」フォルダに保存し
//  URLを運行シートのX列（24列目）に書き込む
// ================================================================
function uploadTerminalFile(id, fileName, base64Data, mimeType) {
  return appendTerminalFile(id, fileName, base64Data, mimeType);
}


// ================================================================
//  10-1: ホーム用連絡事項取得（getMyNotices）  【大A / 中10 / 小10-1】
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
function getMyNotices(companySsId, email) {
  validateDriverEmail_(email, companySsId);
  var savedEmail = email || '';
  if (!savedEmail) return [];
  var ss     = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master) return [];
  var mAll   = master.getDataRange().getValues();
  var myName = '';
  for (var j = 1; j < mAll.length; j++) {
    if (String(mAll[j][10]).trim() === savedEmail) { myName = String(mAll[j][8]).trim(); break; }
  }
  if (!myName) return [];
  var readKey  = 'readNotices_' + savedEmail;
  var readList = JSON.parse(PropertiesService.getScriptProperties().getProperty(readKey) || '[]');
  var sheet    = ss.getSheetByName('運行');
  if (!sheet) return [];
  var all = sheet.getDataRange().getValues();
  var lr  = sheet.getLastRow();
  var notes23 = lr >= 2 ? sheet.getRange(2, 24, lr-1, 1).getNotes() : [];
  var rtvs23m = lr >= 2 ? sheet.getRange(2, 24, lr-1, 1).getRichTextValues() : [];
  var out = [], seen = {};
  for (var i = 1; i < all.length; i++) {
    var r = all[i];
    if (String(r[6]).trim() !== myName) continue;
    var id = String(r[0]||'').trim();
    if (!id || seen[id]) continue;
    seen[id] = true;
    var notice = String(r[22]||'');
    var n23 = (notes23[i-1] && notes23[i-1][0]) || '';
    var dataUrls = n23 ? n23.split('\n').filter(function(u){return u.match(/^https?:\/\//);}) : [];
    if (!dataUrls.length && rtvs23m[i-1] && rtvs23m[i-1][0]) {
      var rtv23m = rtvs23m[i-1][0], rrunsm = rtv23m.getRuns();
      for (var k = 0; k < rrunsm.length; k++) { var lkm=rrunsm[k].getLinkUrl(); if(lkm) dataUrls.push(lkm); }
    }
    if (!dataUrls.length) { var pu=String(r[23]||''); if(pu.match(/^https?:\/\//)) dataUrls=[pu]; }
    if (!notice && !dataUrls.length) continue;
    if (readList.indexOf(id) !== -1) continue;
    out.push({ id:id, date: r[9] ? Utilities.formatDate(new Date(r[9]),'Asia/Tokyo','yyyy/MM/dd HH:mm') : '', notice:notice, dataUrls:dataUrls, dataUrl:dataUrls[0]||'' });
  }
  return out.reverse().slice(0, 20);
}


// ================================================================
//  10-2: ID指定行程取得（getRoutesById）  【大A / 中10 / 小10-2】
//  指定IDの全行程と進捗状態を返す
//  progress: pick / restStart / restEnd / drop / complete
// ================================================================
function getRoutesById(id, companySsId, email) {
  validateDriverEmail_(email, companySsId);
  var ss    = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return { routes:[], progress:'', firstGap:null };
  var all = sheet.getDataRange().getValues();
  var uHdr = all.length > 0 ? all[0].map(function(h){return String(h||'').trim();}) : [];
  var inspAIdx = uHdr.indexOf('点呼後完了');
  var routes = [];
  var allGuideDone=true, anyGuideDone=false;
  var allPickDone=true, anyPickDone=false;
  var hasRestS=false, hasRestE=false;
  var allDropDone=true, anyDropDone=false;
  var hasInspAfter=false;
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() !== String(id).trim()) continue;
    var gDone = !!all[i][13], pDone = !!all[i][14], dDone = !!all[i][17];
    var pickVal = (all[i][11] instanceof Date) ? '' : String(all[i][11] || '');
    var dropVal = (all[i][12] instanceof Date) ? '' : String(all[i][12] || '');
    routes.push({ row:i+1, pick:pickVal, drop:dropVal, guideDone:gDone, pickDone:pDone, dropDone:dDone });
    if (!gDone) allGuideDone = false;
    if (gDone)  anyGuideDone = true;
    if (!pDone) allPickDone = false;
    if (pDone)  anyPickDone = true;
    if (all[i][15]) hasRestS = true;
    if (all[i][16]) hasRestE = true;
    if (!dDone) allDropDone = false;
    if (dDone)  anyDropDone = true;
    if (inspAIdx >= 0 && all[i][inspAIdx]) hasInspAfter = true;
  }
  var progress = 'guide';
  if      (allDropDone && routes.length>0) progress = hasInspAfter ? 'complete' : 'inspAfter';
  else if (anyDropDone)  progress = 'drop';
  else if (hasRestE)     progress = 'drop';
  else if (hasRestS)     progress = 'restEnd';
  else if (allPickDone && routes.length>0) progress = 'restStart';
  else if (anyPickDone)  progress = 'pick';
  else if (allGuideDone && routes.length>0) progress = 'pick';
  else if (anyGuideDone) progress = 'guide';
  var firstGap = null;
  if (!anyGuideDone && (anyPickDone||hasRestS||hasRestE||anyDropDone)) firstGap = 'guide';
  else if (!anyPickDone && (hasRestS||hasRestE||anyDropDone)) firstGap = 'pick';
  else if (!hasRestS && (hasRestE||anyDropDone)) firstGap = 'restStart';
  else if (!hasRestE && anyDropDone) firstGap = 'restEnd';
  if (!firstGap && allDropDone && routes.length>0 && !hasInspAfter) firstGap = 'inspAfter';
  return { routes:routes, progress:progress, firstGap:firstGap };
}


// ================================================================
//  10-2b: 行番号指定で連絡事項取得（getNoticeByRow）  【大A / 中10 / 小10-2b】
//  誘導画面に管理側の連絡事項・データURLを表示するために使う
// ================================================================
function getNoticeByRow(id, companySsId, email) {
  validateDriverEmail_(email, companySsId);
  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet || !id) return { notice:'', dataUrls:[], dataUrl:'' };
  var all = sheet.getDataRange().getValues();
  if (all.length < 2) return { notice:'', dataUrls:[], dataUrl:'' };
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
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() !== String(id).trim()) continue;
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
//  10-3: 既読管理・既読にする（markAsRead）  【大A / 中10 / 小10-3】
//  既読にしたIDをPropertiesServiceに保存する（最大200件）
// ================================================================
function markAsRead(id, email) {
  var readKey = email ? ('readNotices_' + email) : 'readNotices';
  var p    = PropertiesService.getScriptProperties();
  var read = JSON.parse(p.getProperty(readKey) || '[]');
  if (read.indexOf(id) === -1) {
    read.push(id);
    if (read.length > 200) read = read.slice(-200);
    p.setProperty(readKey, JSON.stringify(read));
  }
}


// ================================================================
//  10-4: 既読管理・既読一覧取得（getReadNotices）  【大A / 中10 / 小10-4】
//  PropertiesServiceから既読IDリストを取得して返す
// ================================================================
function getReadNotices(email) {
  var readKey = email ? ('readNotices_' + email) : 'readNotices';
  return JSON.parse(PropertiesService.getScriptProperties().getProperty(readKey) || '[]');
}


// ================================================================
//  11-1: 会社登録シートのonEditハンドラ（onEditCompanyRegister_）  【大B / 中11 / 小11-1】
//  ・A列(会社名)+B列(管理Gmail) が揃いC列が空 → フォルダ作成・通知メール送信
//  ・F列(スプレッドシートURL)+G列(WebアプリURL) が両方揃いH列が空 → 配布メール自動送信
//  C列: セットアップ状態（済/エラー）  D列: 実行日時  E列: フォルダURL
//  F列: スプレッドシートURL  G列: WebアプリURL  H列: 配布メール送信状態
//  onEdit(e) から呼び出す（シート名='会社登録' の場合）
// ================================================================
function onEditCompanyRegister_(sheet, range) {
  var row = range.getRow();
  if (row <= 1) return; // ヘッダー行スキップ

  var col = range.getColumn();

  // ── A列 or B列が編集された → installedOnEdit_ がフルセットアップを実行する ──
  // （シンプルトリガーには認証付き処理が不可のため、ここでは何もしない）
  if (col === 1 || col === 2) return;

  // ── F列 or G列が編集された → 両方揃ったら配布メール自動送信 ──
  if (col === 6 || col === 7) {
    var ssUrl      = String(sheet.getRange(row, 6).getValue() || '').trim();
    var appUrl     = String(sheet.getRange(row, 7).getValue() || '').trim();
    var mailStatus = String(sheet.getRange(row, 8).getValue() || '').trim();
    if (!ssUrl || !appUrl) return; // 両方揃うまで待つ
    if (mailStatus && mailStatus.indexOf('送信済') !== -1) return; // 送信済はスキップ
    var cName = String(sheet.getRange(row, 1).getValue() || '').trim();
    var aEmail = String(sheet.getRange(row, 2).getValue() || '').trim();
    if (!cName || !aEmail || aEmail.indexOf('@') === -1) return;
    try {
      sendDistributionMail_(cName, aEmail, ssUrl, appUrl, row, sheet);
    } catch(e) {
      sheet.getRange(row, 8).setValue('エラー: ' + e.message).setBackground('#ffcdd2');
    }
  }
}


// ================================================================
//  11-2: 1社分のセットアップ実行（setupOneCompany_）  【大B / 中11 / 小11-2】
//  「運行管理_アーカイブ/会社名/」フォルダを作成し管理Gmailに編集権限を付与する。
//  suppressEmail=true のとき通知メールを送らない（processNewCompany_ から呼ぶ場合）。
//  戻り値: { folderUrl, folderId }
// ================================================================
function setupOneCompany_(companyName, adminEmail, suppressEmail) {
  // 運行管理_アーカイブ/ を作成or取得
  var rootFolder    = getOrCreateFolder_('運行管理_アーカイブ');
  // 会社名サブフォルダを作成or取得
  var subIter       = rootFolder.getFoldersByName(companyName);
  var companyFolder = subIter.hasNext() ? subIter.next() : rootFolder.createFolder(companyName);

  // 管理Gmailに編集権限を付与（フォルダ）
  try { companyFolder.addEditor(adminEmail); } catch(e) {}

  var folderUrl = companyFolder.getUrl();
  var folderId  = companyFolder.getId();

  // processNewCompany_ から呼ばれた場合は個別メールを送らない
  // （sendDistributionMail_ で1通にまとめて送信するため）
  if (!suppressEmail) {
    var subject = '[運行管理] ' + companyName + ' のご利用準備が整いました';
    var body =
      companyName + ' 担当者様\n\n' +
      '運行管理システムのフォルダを作成しました。\n\n' +
      '■ 共有フォルダ（運行データ・アーカイブの保管場所）\n' +
      folderUrl + '\n\n' +
      'スプレッドシートとアプリのURLは別途お知らせします。';
    try { GmailApp.sendEmail(adminEmail, subject, body); } catch(e) {}
  }

  return { folderUrl: folderUrl, folderId: folderId };
}


// ================================================================
//  11-3: 全会社セットアップ実行（setupCompanies）  【大C / 中11 / 小11-3】
//  会社登録シートを読み込んで未処理の全社を一括セットアップする
//  会社登録シートがなければ新規作成して案内を表示する
//  スクリプトエディタまたは管理者メニューから手動実行
// ================================================================
function setupCompanies() {
  var ui    = SpreadsheetApp.getUi();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('会社登録');

  // 会社登録シートがなければ新規作成
  if (!sheet) {
    sheet = ss.insertSheet('会社登録');
    var header = ['会社名', '管理用Gmail', 'セットアップ状態', '実行日時', 'フォルダURL',
                  'スプレッドシートURL', 'WebアプリURL', '配布メール送信状態'];
    sheet.getRange(1, 1, 1, 8).setValues([header]);
    sheet.getRange(1, 1, 1, 8).setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 160);
    sheet.setColumnWidth(4, 160);
    sheet.setColumnWidth(5, 350);
    sheet.setColumnWidth(6, 350);
    sheet.setColumnWidth(7, 350);
    sheet.setColumnWidth(8, 200);
    // サンプル行（社名A・社名B）
    sheet.getRange(2, 1, 2, 2).setValues([
      ['社名A', '（管理用Gmailを入力）'],
      ['社名B', '（管理用Gmailを入力）']
    ]);
    ui.alert(
      '「会社登録」シートを作成しました。\n\n' +
      '① B列（管理用Gmail）に各社担当者のGmailを入力 → フォルダ自動作成\n' +
      '② F列（スプレッドシートURL）+ G列（WebアプリURL）を入力\n' +
      '   → 管理者・乗務員に配布メールを自動送信します。'
    );
    return;
  }

  // 既存シートにF〜H列ヘッダーが未追加なら追加
  var existingHeader = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (existingHeader.length < 6 || !existingHeader[5]) {
    sheet.getRange(1, 6).setValue('スプレッドシートURL').setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold');
    sheet.getRange(1, 7).setValue('WebアプリURL').setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold');
    sheet.getRange(1, 8).setValue('配布メール送信状態').setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setColumnWidth(6, 350);
    sheet.setColumnWidth(7, 350);
    sheet.setColumnWidth(8, 200);
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('会社登録シートにデータがありません'); return; }

  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var done = 0, skip = 0, errCount = 0;

  for (var i = 0; i < data.length; i++) {
    var row         = i + 2;
    var companyName = String(data[i][0] || '').trim();
    var adminEmail  = String(data[i][1] || '').trim();
    var status      = String(data[i][2] || '').trim();

    if (!companyName || !adminEmail || adminEmail.indexOf('@') === -1) { skip++; continue; }
    if (status === '済') { skip++; continue; }

    try {
      var result = setupOneCompany_(companyName, adminEmail);
      sheet.getRange(row, 3).setValue('済').setBackground('#c8e6c9');
      sheet.getRange(row, 4).setValue(new Date());
      sheet.getRange(row, 5).setValue(result.folderUrl);
      done++;
    } catch(e) {
      sheet.getRange(row, 3).setValue('エラー: ' + e.message).setBackground('#ffcdd2');
      errCount++;
    }
  }

  ui.alert(
    'セットアップ完了\n\n' +
    '処理済: ' + done + '社\n' +
    'スキップ（済 or 未入力）: ' + skip + '社\n' +
    'エラー: ' + errCount + '社'
  );
}


// ================================================================
//  11-4: スプレッドシート内「使い方」シート自動作成（createUsageSheet）  【大C / 中11 / 小11-4】
//  スプレッドシートを開いたまま確認できる操作手順書シートを作成する
//  管理者向け（スプレッドシート操作）とドライバー向け（アプリ操作）の2部構成
//  スクリプトエディタまたは管理者メニューから手動実行
// ================================================================
function createUsageSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName('使い方');
  if (sheet) { ss.deleteSheet(sheet); }
  sheet = ss.insertSheet('使い方');
  ss.moveActiveSheet(1); // 先頭に移動

  // ── レイアウト設定 ──────────────────────────
  sheet.setColumnWidth(1, 30);   // A: 余白
  sheet.setColumnWidth(2, 220);  // B: 項目
  sheet.setColumnWidth(3, 400);  // C: 内容
  sheet.setColumnWidth(4, 30);   // D: 余白

  var row = 1;

  function title(text, bgColor) {
    sheet.getRange(row, 1, 1, 4).merge().setValue(text)
      .setBackground(bgColor || '#1565c0').setFontColor('#ffffff')
      .setFontSize(16).setFontWeight('bold').setVerticalAlignment('middle');
    sheet.setRowHeight(row, 50);
    row++;
  }
  function section(text, bgColor) {
    sheet.getRange(row, 2, 1, 2).merge().setValue('▶ ' + text)
      .setBackground(bgColor || '#1976d2').setFontColor('#ffffff')
      .setFontSize(13).setFontWeight('bold');
    sheet.setRowHeight(row, 36);
    row++;
  }
  function item(label, value, labelBg) {
    sheet.getRange(row, 2).setValue(label)
      .setBackground(labelBg || '#e3f2fd').setFontSize(12).setFontWeight('bold')
      .setVerticalAlignment('top').setWrap(true);
    sheet.getRange(row, 3).setValue(value)
      .setFontSize(12).setVerticalAlignment('top').setWrap(true);
    sheet.setRowHeight(row, 60);
    row++;
  }
  function note(text) {
    sheet.getRange(row, 2, 1, 2).merge().setValue('  ' + text)
      .setBackground('#fff9c4').setFontSize(11).setFontColor('#5d4037').setWrap(true);
    sheet.setRowHeight(row, 40);
    row++;
  }
  function spacer() {
    sheet.setRowHeight(row, 14);
    row++;
  }

  // ══════════════════════════════════════
  // タイトル
  // ══════════════════════════════════════
  title('　運行管理システム　使い方ガイド', '#1a237e');
  spacer();

  // ══════════════════════════════════════
  // PART 1: 管理者向け（スプレッドシート操作）
  // ══════════════════════════════════════
  title('　【管理者向け】 スプレッドシート操作', '#1b5e20');
  spacer();

  section('メニューの使い方（上の「メニュー」をクリック）', '#2e7d32');
  item('ホーム画面を表示', 'スプレッドシート右側にアプリ画面を表示します。\nドライバーの連絡・データ確認に使います。', '#c8e6c9');
  item('集計表再生成', '集計表の内容がおかしい時に押します。\n運行シートのデータから自動で再計算します。', '#c8e6c9');
  item('シート再生成', '自車専属マスタ・自車専属運行シートを最新状態に整備します。\n新しい車両を追加した後などに使います。', '#c8e6c9');
  item('📅 月生成', '翌月分の運行予定（プレースホルダー）を一括作成します。\n毎月25日〜月末頃に押してください。\n積地が空の行は黄色で表示されます（配車漏れ警告）。', '#c8e6c9');
  item('📦 前月分アーカイブ', '前月のデータを別ファイルに保存して運行シートをスッキリさせます。\n月生成時に自動実行されますが、手動でも使えます。', '#c8e6c9');
  item('💴 経費自動入力', '自車専属マスタで行を選択して実行します。\nトン数に応じた経費の平均値をQ列以降に自動入力します。', '#c8e6c9');
  item('🔃 日付順並び替え', '運行シートと集計表を両方日付順に並び替えます。', '#c8e6c9');
  item('📷 写真・ファイル取込', '選択した行にPC上の写真・ファイルを直接添付します。', '#c8e6c9');
  item('📖 使い方シート作成', 'この使い方シートを再作成します。', '#c8e6c9');
  spacer();

  section('月次の運用フロー', '#2e7d32');
  item('① 月末（25日頃）', '「📅 月生成」を押して翌月分の行を一括作成', '#c8e6c9');
  item('② 配車確定後', '生成された行（黄色=未配車）に積地・降地を入力\n→ 黄色が消えて配車完了', '#c8e6c9');
  item('③ 翌月になったら', '「📦 前月分アーカイブ」（または月生成時に自動実行）', '#c8e6c9');
  note('⚠ 月生成は翌月分が存在しない場合のみ実行できます（重複防止）');
  spacer();

  section('自車専属マスタの管理', '#2e7d32');
  item('運行状態の変更', 'B列の値を変更\n「運行」= 月生成の対象  /  「待機」「故障」= 対象外', '#c8e6c9');
  item('新しい車両追加', 'マスタに1行追加 → ID（S-XXXX）が自動生成されます', '#c8e6c9');
  item('給料・歩合の設定', 'N列=仮日数、O列=給料、P列=% を入力すると集計表に自動反映', '#c8e6c9');
  spacer();

  section('データ一括読込（インポート）のルール', '#2e7d32');
  item('読込の仕組み', 'システムの項目（列）に合わせて読み込みます。\nファイルを開いたあと、各列が何に対応するかを画面上で確認・修正できます。', '#c8e6c9');
  item('Excelに無い項目がある場合', 'お使いのExcelに無い項目（日付など）があれば、読み込む前にExcel側に列を追加して入力しておいてください。', '#c8e6c9');
  item('不要な列があっても問題なし', '不要な列がExcelにあっても無視されるのでそのまま読込可能です。', '#c8e6c9');
  spacer();

  spacer();

  // ══════════════════════════════════════
  // PART 2: ドライバー向け（アプリ操作）
  // ══════════════════════════════════════
  title('　【ドライバー向け】 アプリ操作', '#4a148c');
  spacer();

  section('最初に1回だけやること', '#6a1b9a');
  item('① URLを開く', '担当者から受け取ったURLをスマホのブラウザで開く\n（ホーム画面に追加しておくと便利）', '#e1bee7');
  item('② アドレス紐づけ', '画面下の「アドレス紐づけ」欄に自分のGmailアドレスを入力して「紐づけ」を押す\n→「紐づけOK」と表示されれば完了', '#e1bee7');
  note('✅ 紐づけは一度やれば次回から自動でログインされます');
  spacer();

  section('毎日の使い方', '#6a1b9a');
  item('① 運行開始', '「運行開始」ボタンを押す\n積地・降地を入力して「スタート」', '#e1bee7');
  item('② 誘導', '現場に向けて出発するとき「誘導」ボタンを押す', '#e1bee7');
  item('③ 積完', '荷物の積み込みが終わったとき「積完」ボタンを押す', '#e1bee7');
  item('④ 休憩（あれば）', '休憩に入るとき「休憩開始」、終わったら「休憩終了」を押す', '#e1bee7');
  item('⑤ 降完', '荷降ろしが終わったとき「降完」ボタンを押す\n→ 1日の記録が完了', '#e1bee7');
  spacer();

  section('困った時の対処', '#6a1b9a');
  item('アプリを閉じてしまった', 'URLを開き直せば途中から再開できます', '#e1bee7');
  item('積地・降地を間違えた', '「運行一覧」→ 該当行の「編集」ボタンで修正できます', '#e1bee7');
  item('ボタンを押し間違えた', '「戻る」ボタンで1つ前の操作に戻れます', '#e1bee7');
  item('画面が固まった', 'ブラウザを更新（リロード）してください\n途中の状態は保存されています', '#e1bee7');
  item('黄色の「未配車」が表示されている', '積地がまだ決まっていない日です\n担当者に確認してください', '#e1bee7');
  spacer();

  section('一覧画面のボタンの意味', '#6a1b9a');
  item('黄色「運行開始」', 'まだ始めていない仕事。押すと運行を開始できます', '#e1bee7');
  item('オレンジ「運行再開」', '途中で止まっている仕事。続きから再開できます', '#e1bee7');
  item('グレー「完了」', 'その日の仕事が終わっています', '#e1bee7');
  item('グレー青「運行無」', '休み・有休の日です', '#e1bee7');
  item('「連絡」ボタン', '担当者からのメッセージを確認できます', '#e1bee7');
  item('「編集」ボタン', '積地・降地・売上などを後から修正できます', '#e1bee7');
  spacer();

  // フッター
  sheet.getRange(row, 1, 1, 4).merge()
    .setValue('このシートは「createUsageSheet」を実行すると再作成できます')
    .setBackground('#263238').setFontColor('#546e7a').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(row, 30);

  sheet.setFrozenRows(0);
  sheet.setTabColor('#ffd600');

  ui.alert('「使い方」シートを作成しました。\n先頭タブに移動しています。');
}


// ================================================================
//  11-5: 配布メール自動送信（sendDistributionMail_）  【大B / 中11 / 小11-5】
//  管理者向け（スプレッドシートURL＋アプリURL）と乗務員向け（アプリURL＋紐づけ手順）の
//  2種類のメールを送信する。乗務員メールは自車専属マスタのJ列から全員分個別送信。
//  onEditCompanyRegister_（F/G列入力時）または triggerDistributionMail から呼び出す。
// ================================================================
function sendDistributionMail_(companyName, adminEmail, ssUrl, appUrl, row, sheet) {
  // 乗務員リストは③各客SSのマスタから取得する
  var clientSsMatch = ssUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
  var clientSs = null;
  if (clientSsMatch) {
    try { clientSs = SpreadsheetApp.openById(clientSsMatch[1]); } catch(e) {}
  }
  var adminSent = 0, driverSent = 0;

  // ── 管理者向けメール ──────────────────────────────────
  var adminSubject = '[運行管理] ' + companyName + ' 運行管理システムのご案内';
  var adminBody =
    companyName + ' ご担当者様\n\n' +
    'このたびは運行管理システムをご利用いただきありがとうございます。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 運行管理スプレッドシート（PC・タブレット推奨）\n' +
    ssUrl + '\n' +
    '　→ 配車・運行データの入力・月生成・集計表の確認はこちら\n\n' +
    '■ 運行管理アプリ（乗務員用 スマートフォン推奨）\n' +
    appUrl + '\n' +
    '　→ 乗務員がスマートフォンから運行状況を入力するアプリです\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    '【スプレッドシートの使い方】\n' +
    '・このスプレッドシートに乗務員のメールアドレスを入力してください\n' +
    '・「自車専属マスタ」タブのJ列（メールアドレス）に1名ずつ入力します\n\n' +
    '【乗務員への配布方法】\n' +
    '・各乗務員に上記アプリURLを共有してください\n' +
    '・初回は「紐づけ設定」でメールアドレスを登録するだけで使えます\n\n' +
    'ご不明な点はお気軽にお問い合わせください。\n' +
    'よろしくお願いいたします。';
  try {
    GmailApp.sendEmail(adminEmail, adminSubject, adminBody);
    adminSent++;
  } catch(e) {}

  // ── 乗務員向けメール（③各客SSの自車専属マスタ J列の全アドレスに個別送信）──
  var master = clientSs ? clientSs.getSheetByName('自車専属マスタ') : null;
  if (master && master.getLastRow() >= 2) {
    var masterData = master.getRange(2, 1, master.getLastRow() - 1, 10).getValues();
    for (var i = 0; i < masterData.length; i++) {
      var driverEmail = String(masterData[i][9] || '').trim(); // J列(index9)=メールアドレス
      var driverName  = String(masterData[i][7] || '').trim(); // H列(index7)=乗務員名
      if (!driverEmail || driverEmail.indexOf('@') === -1) continue;

      var driverSubject = '[運行管理] 運行管理アプリのご案内';
      var driverBody =
        (driverName ? driverName + ' さん\n\n' : '') +
        'お疲れさまです。\n\n' +
        '運行管理アプリをご案内します。\n\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
        '■ 運行管理アプリ（以下のURLをブックマーク登録してください）\n' +
        appUrl + '\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
        '【最初にやること（初回のみ・1分で完了）】\n' +
        '① 上のURLをスマートフォンで開く\n' +
        '② 画面下の「紐づけ設定」をタップ\n' +
        '③ このメールアドレス（' + driverEmail + '）を入力して「登録」\n' +
        '④ 「登録しました」と表示されたら設定完了！\n\n' +
        '次回からは同じURLを開くだけで、自動的にあなたの運行データが表示されます。\n\n' +
        '何かご不明な点があれば担当者にご連絡ください。';
      try {
        GmailApp.sendEmail(driverEmail, driverSubject, driverBody);
        driverSent++;
      } catch(e) {}
    }
  }

  // H列にステータスを記録
  if (row && sheet) {
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
    sheet.getRange(row, 8).setValue('送信済(' + now + ')').setBackground('#c8e6c9');
  }

  return { adminSent: adminSent, driverSent: driverSent };
}


// ================================================================
//  11-6: メニューから配布メール一括送信（triggerDistributionMail）  【大C / 中11 / 小11-6】
//  会社登録シートの F列(SS URL)＋G列(AppURL) が揃いH列が未送信の全行にメール送信する
//  メニュー「📧 配布メール送信」から手動実行
// ================================================================
function triggerDistributionMail() {
  var ui    = SpreadsheetApp.getUi();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('会社登録');

  if (!sheet) {
    ui.alert('「会社登録」シートがありません。\n先に「🏢 会社セットアップ実行」を実行してください。');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('会社登録シートにデータがありません。'); return; }

  var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var totalAdmin = 0, totalDriver = 0, skip = 0, errCount = 0;

  for (var i = 0; i < data.length; i++) {
    var row         = i + 2;
    var companyName = String(data[i][0] || '').trim();
    var adminEmail  = String(data[i][1] || '').trim();
    var ssUrl       = String(data[i][5] || '').trim(); // F列
    var appUrl      = String(data[i][6] || '').trim(); // G列
    var mailStatus  = String(data[i][7] || '').trim(); // H列

    if (!companyName || !adminEmail || adminEmail.indexOf('@') === -1) { skip++; continue; }
    if (!ssUrl || !appUrl) { skip++; continue; }
    if (mailStatus.indexOf('送信済') !== -1) { skip++; continue; }

    try {
      var result = sendDistributionMail_(companyName, adminEmail, ssUrl, appUrl, row, sheet);
      totalAdmin  += result.adminSent;
      totalDriver += result.driverSent;
    } catch(e) {
      sheet.getRange(row, 8).setValue('エラー: ' + e.message).setBackground('#ffcdd2');
      errCount++;
    }
  }

  ui.alert(
    '配布メール送信完了\n\n' +
    '管理者向け: ' + totalAdmin + '通\n' +
    '乗務員向け: ' + totalDriver + '通\n' +
    'スキップ（URL未入力 or 送信済）: ' + skip + '社\n' +
    'エラー: ' + errCount + '社'
  );
}


// ================================================================
//  12-1: WebアプリのベースURL取得（getWebAppBaseUrl_）  【大B / 中12 / 小12-1】
//  ① Script Propertiesに保存済みのURLを返す（doGetアクセス時に自動保存される）。
//  ② 未保存の場合は ScriptApp.getService().getUrl() で自動取得して保存する。
//  ③ それでも取得できない場合のみ空文字を返す（メニューからの手動設定は不要）。
// ================================================================
function getWebAppBaseUrl_() {
  var PROD_DEPLOY_ID = 'AKfycbw7rzkd_SuE1I6BNzEjED4Mxl6cnM4wbswIiRiNoPf5zcSS2JcP6YLkfRV21fLc0opU';
  var PROD_URL = 'https://script.google.com/macros/s/' + PROD_DEPLOY_ID + '/exec';
  var props = PropertiesService.getScriptProperties();
  var stored = props.getProperty('webAppUrl') || '';
  // 本番デプロイIDが含まれるURLのみ有効とする
  if (stored && stored.indexOf(PROD_DEPLOY_ID) !== -1) return stored;
  // どんな状況でも本番URLを返す（テンプレートSS等で上書きされた値は無視）
  props.setProperty('webAppUrl', PROD_URL);
  return PROD_URL;
}


// ================================================================
//  12-2: WebアプリURLをScript Propertiesに保存（setWebAppUrl）  【大C / 中12 / 小12-2】
//  メニュー「⚙ WebアプリURLを設定」から手動実行。デプロイURLを1回入力するだけでOK。
// ================================================================
function setWebAppUrl() {
  var ui  = SpreadsheetApp.getUi();
  var cur = getWebAppBaseUrl_();
  var res = ui.prompt(
    '⚙ WebアプリURL設定',
    'Google Apps ScriptのデプロイURL（exec で終わるURL）を入力してください。\n\n' +
    '現在の設定: ' + (cur || '（未設定）') + '\n\n' +
    '例: https://script.google.com/macros/s/XXXXXXXXXX/exec',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var url = res.getResponseText().trim();
  if (!url || url.indexOf('https://') !== 0) {
    ui.alert('URLが正しくありません。https:// で始まるURLを入力してください。');
    return;
  }
  // /exec の後ろに余分な文字があれば /exec までトリム
  var execIdx = url.indexOf('/exec');
  if (execIdx !== -1) url = url.substring(0, execIdx + 5);
  PropertiesService.getScriptProperties().setProperty('webAppUrl', url);
  ui.alert('保存しました！\n\n' + url + '\n\nこれ以降、会社SS作成時にこのURLが自動で使われます。');
}


// getOrCreateCompanyTemplate_ は syncToTemplateSS に統合済みのため削除


// ================================================================
//  12-3d: 新規SSのバインドスクリプトID取得（getNewSsScriptId_）  【大B / 中12】
//  DriveApp でSSの子ファイルからApps Scriptを検索してスクリプトIDを返す
// ================================================================
function getNewSsScriptId_(ssId) {
  try {
    var token = ScriptApp.getOAuthToken();
    var q = encodeURIComponent('mimeType="application/vnd.google-apps.script" and "' + ssId + '" in parents');
    var resp = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files?q=' + q + '&fields=files(id,name)&supportsAllDrives=true&includeItemsFromAllDrives=true',
      { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (resp.getResponseCode() !== 200) return null;
    var data = JSON.parse(resp.getContentText());
    return (data.files && data.files.length > 0) ? data.files[0].id : null;
  } catch(e) { return null; }
}


// ================================================================
//  12-3: 会社専用スプレッドシートを作成（createCompanySpreadsheet_）  【大B / 中12 / 小12-3】
//  ②客用SS（スタブコード入り）をコピーして各客SSを作る。
//  コピー元は Script Properties の clientTemplateSsId。未設定ならハードコードIDを使用。
//  __TEMPLATE_SS__ を削除して __COMPANY_SS__ に会社名をセット。
//  targetFolderId が指定されたそのフォルダに移動する（未指定なら「運行管理_会社別」）。
//  戻り値: { ssId, ssUrl }
// ================================================================
function createCompanySpreadsheet_(companyName, adminEmail, targetFolderId) {
  var props = PropertiesService.getScriptProperties();
  var templateSsId = props.getProperty('clientTemplateSsId') || '1NBtosd_MN8KcboV_4OXTrY8WqcE3TJwpxdA_nASmTOo';
  var destFolder = targetFolderId
    ? DriveApp.getFolderById(targetFolderId)
    : getOrCreateFolder_('運行管理_会社別');

  var templateFile = DriveApp.getFileById(templateSsId);
  var newFile = templateFile.makeCopy(companyName + ' 運行管理', destFolder);
  try { newFile.addEditor(adminEmail); } catch(e) {}

  var newSs = SpreadsheetApp.openById(newFile.getId());

  // __TEMPLATE_SS__ マーカーを削除（コピー元②のマーカーが引き継がれるため）
  var tmplMarker = newSs.getSheetByName('__TEMPLATE_SS__');
  if (tmplMarker && newSs.getSheets().length > 1) {
    try { newSs.deleteSheet(tmplMarker); } catch(e) {}
  }

  // __COMPANY_SS__ マーカーに会社名をセット（なければ作成）
  var marker = newSs.getSheetByName('__COMPANY_SS__') || newSs.insertSheet('__COMPANY_SS__');
  marker.getRange(1, 1).setValue(companyName);
  if (!marker.isSheetHidden()) marker.hideSheet();

  // 不要シートを削除（管理者専用シートが引き継がれていた場合）
  var validNames = ['運行','集計表','自車専属マスタ','自車専属運行','マスタ','設定','__COMPANY_SS__'];
  newSs.getSheets().forEach(function(s) {
    if (validNames.indexOf(s.getName()) === -1 && newSs.getSheets().length > 1) {
      try { newSs.deleteSheet(s); } catch(e) {}
    }
  });

  // 全シートをテストデータ初期状態にリセット（②のデータが混入しないよう必須）
  initClientSSSheets_(newSs, companyName);

  // 点検項目が未設定なら初期値をセット（業務前点検・業務後点検）
  ensureSettingItems_(newSs);

  // ③のスクリプトIDを取得して WebApp を自動デプロイ
  var newScriptId = getNewSsScriptId_(newSs.getId());
  props.setProperty('scriptId_' + newSs.getId(), newScriptId || '');
  var deployResult = deployClientWebApp_(newSs.getId(), companyName, newScriptId);
  var webAppUrl = deployResult ? deployResult.webAppUrl : '';
  var finalScriptId = (deployResult && deployResult.scriptId) ? deployResult.scriptId : (newScriptId || '');

  return { ssId: newSs.getId(), ssUrl: newSs.getUrl(), scriptId: finalScriptId, webAppUrl: webAppUrl || '' };
}


// ================================================================
//  12-3a: 各客SS用 WebApp を自動デプロイ（deployClientWebApp_）
//  Apps Script API でスクリプトにスタブコードを書き込み WebApp をデプロイして URL を返す
// ================================================================
function deployClientWebApp_(ssId, companyName, existingScriptId, libVersion) {
  try {
    var token   = ScriptApp.getOAuthToken();
    var hdrs    = { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' };
    var apiBase = 'https://script.googleapis.com/v1/projects';

    // スクリプトID確定（引数で渡された既存ID → なければ新規作成）
    var scriptId = existingScriptId || null;
    if (!scriptId) {
      var cr = UrlFetchApp.fetch(apiBase, {
        method: 'POST', headers: hdrs,
        payload: JSON.stringify({ title: companyName + '運行管理', parentId: ssId }),
        muteHttpExceptions: true
      });
      var crData = JSON.parse(cr.getContentText());
      scriptId = crData.scriptId;
      if (!scriptId) {
        var crErr = (crData.error && crData.error.message) ? crData.error.message.slice(0, 100) : cr.getContentText().slice(0, 100);
        return { error: 'projects.create ' + cr.getResponseCode() + ': ' + crErr };
      }
    }

    // マニフェスト（appsscript.json）
    var libVer = libVersion ? String(libVersion)
      : (PropertiesService.getScriptProperties().getProperty('approvedLibVersion')
      || PropertiesService.getScriptProperties().getProperty('pinnedLibVersion')
      || '308');
    var manifest = JSON.stringify({
      timeZone: 'Asia/Tokyo',
      dependencies: { libraries: [{ userSymbol: 'UnkouLib',
        libraryId: '1n79omnAcdsEojMRyjnj9-Ic9pIl1-7Nt_HB7Avy0NVFizOSeqt0guqyZ',
        version: libVer, developmentMode: true }] },
      webapp: { executeAs: 'USER_DEPLOYING', access: 'ANYONE_ANONYMOUS' },
      oauthScopes: [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive',
        'https://www.googleapis.com/auth/script.external_request',
        'https://www.googleapis.com/auth/script.scriptapp',
        'https://www.googleapis.com/auth/gmail.send',
        'https://www.googleapis.com/auth/script.container.ui'
      ],
      exceptionLogging: 'STACKDRIVER', runtimeVersion: 'V8'
    });

    // スタブコード書き込み
    UrlFetchApp.fetch(apiBase + '/' + scriptId + '/content', {
      method: 'PUT', headers: hdrs,
      payload: JSON.stringify({ files: [
        { name: 'appsscript', type: 'JSON',      source: manifest },
        { name: 'コード',      type: 'SERVER_JS', source: getClientStubSource_() }
      ]}),
      muteHttpExceptions: true
    });

    // WebApp デプロイ作成
    var dr = UrlFetchApp.fetch(apiBase + '/' + scriptId + '/deployments', {
      method: 'POST', headers: hdrs,
      payload: JSON.stringify({ description: companyName + '_WebApp', manifestFileName: 'appsscript' }),
      muteHttpExceptions: true
    });
    var drData = JSON.parse(dr.getContentText());
    if (!drData.deploymentId) return null;

    var webAppUrl = 'https://script.google.com/macros/s/' + drData.deploymentId + '/exec';
    PropertiesService.getScriptProperties().setProperty('scriptId_' + ssId, scriptId);
    return { scriptId: scriptId, webAppUrl: webAppUrl };
  } catch(e) { return null; }
}


// スタブコードのソース文字列（stub_for_clientSS/コード.js と同一内容）
function getClientStubSource_() {
  return [
    "function onOpen(){var ss=SpreadsheetApp.getActiveSpreadsheet();var isTemplate=ss.getSheetByName('__TEMPLATE_SS__')!==null;var ui=SpreadsheetApp.getUi();var menu=ui.createMenu('メニュー');menu.addItem('ホーム画面を表示','showSidebar').addSeparator().addItem('📅 今月分生成（途中契約）','generateCurrentMonth').addItem('📅 翌月分生成（前月アーカイブ）','generateNextMonth').addItem('📦 前月分アーカイブ','archiveOldMonth').addSeparator().addItem('📄 請求書生成','showInvoiceDialog').addItem('📄 支払確認書生成','showPaymentDialog').addSeparator().addItem('🔄 メニュー再生成','reloadMenu').addItem('集計表再生成','generateSummary').addItem('シート再生成','expandAndRefreshSheets').addItem('💴 経費自動入力','autoFillExpense').addItem('🔃 日付順並び替え','sortBothSheetsByDate').addItem('🆔 ID・車番一括補完','fillMissingIdsAndCars').addSeparator().addItem('📷 写真・ファイル取込','showUploadSidebar').addItem('📖 使い方シート作成','createUsageSheet').addSeparator().addSubMenu(ui.createMenu('📥 データ読み込み（CSV）').addItem('運行シート','showCsvImportDialogUnkou').addItem('自車専属マスタ','showCsvImportDialogMaster').addItem('マスタ（取引先）','showCsvImportDialogCust').addSeparator().addItem('🗑 空インポート行を削除','deleteBlankImportRows')).addSeparator().addSubMenu(ui.createMenu('📋 帳票・送信メニュー').addItem('① 発注書・指示書を作成（協力会社・乗務員用）','showHatchuDocDialog').addItem('② 車番連絡を作成（荷主用）','showShabanDocDialog')).addSeparator().addItem('🔗 チェックした行を配車確定','matchAndConfirmDispatch');if(isTemplate){menu.addSeparator().addItem('📤 各客に反映','syncToAllClientSS');}menu.addToUi();try{UnkouLib.convertLegacyAdminDataUrls_();}catch(e){}try{UnkouLib.applyHolidayRowColors_();}catch(e){}}",
    "function doGet(e){return UnkouLib.doGet(e);}",
    "function onEdit(e){return UnkouLib.onEdit(e);}",
    "function installedOnEdit_(e){return UnkouLib.installedOnEdit_(e);}",
    "function reloadMenu(){onOpen();SpreadsheetApp.getActiveSpreadsheet().toast('メニューを再生成しました','🔄',3);}",
    "function showSidebar(){return UnkouLib.showSidebar();}",
    "function showUploadSidebar(){return UnkouLib.showUploadSidebar();}",
    "function generateCurrentMonth(){return UnkouLib.generateCurrentMonth();}",
    "function generateNextMonth(){return UnkouLib.generateNextMonth();}",
    "function archiveOldMonth(){return UnkouLib.archiveOldMonth();}",
    "function generateSummary(){return UnkouLib.generateSummary();}",
    "function expandAndRefreshSheets(){return UnkouLib.expandAndRefreshSheets();}",
    "function autoFillExpense(){return UnkouLib.autoFillExpense();}",
    "function sortBothSheetsByDate(){return UnkouLib.sortBothSheetsByDate();}",
    "function fillMissingIdsAndCars(){return UnkouLib.fillMissingIdsAndCars();}",
    "function createUsageSheet(){return UnkouLib.createUsageSheet();}",
    "function installTriggers(){return UnkouLib.installTriggers();}",
    "function setRecalcChoice(a){return UnkouLib.setRecalcChoice(a);}",
    "function syncToAllClientSS(){return UnkouLib.syncToAllClientSS();}",
    "function storeCompanySsId(a){return UnkouLib.storeCompanySsId(a);}",
    "function getInitialData(a,b){return UnkouLib.getInitialData(a,b);}",
    "function linkAddress(a,b){return UnkouLib.linkAddress(a,b);}",
    "function unlinkAddress(a){return UnkouLib.unlinkAddress(a);}",
    "function saveRunState(a,b,c){return UnkouLib.saveRunState(a,b,c);}",
    "function loadRunState(){return UnkouLib.loadRunState();}",
    "function clearRunState(a,b){return UnkouLib.clearRunState(a,b);}",
    "function getTodayRoutes(a,b){return UnkouLib.getTodayRoutes(a,b);}",
    "function createParentRows(a,b,c,d,e,f){return UnkouLib.createParentRows(a,b,c,d,e,f);}",
    "function setPickComplete(a,b,c){return UnkouLib.setPickComplete(a,b,c);}",
    "function setRest(a,b,c,d){return UnkouLib.setRest(a,b,c,d);}",
    "function setDropComplete(a,b,c){return UnkouLib.setDropComplete(a,b,c);}",
    "function updateRouteData(a,b,c,d){return UnkouLib.updateRouteData(a,b,c,d);}",
    "function deleteRunRows(a,b,c){return UnkouLib.deleteRunRows(a,b,c);}",
    "function clearTimeCell(a,b,c,d,e){return UnkouLib.clearTimeCell(a,b,c,d,e);}",
    "function getListData(a,b,c,d){return UnkouLib.getListData(a,b,c,d);}",
    "function getEditData(a,b,c){return UnkouLib.getEditData(a,b,c);}",
    "function saveEditData(a,b,c){return UnkouLib.saveEditData(a,b,c);}",
    "function appendTerminalFile(a,b,c,d,e,f){return UnkouLib.appendTerminalFile(a,b,c,d,e,f);}",
    "function deleteRunById(a,b,c){return UnkouLib.deleteRunById(a,b,c);}",
    "function saveNotice(a,b,c,d){return UnkouLib.saveNotice(a,b,c,d);}",
    "function uploadFileToRow(a,b,c,d){return UnkouLib.uploadFileToRow(a,b,c,d);}",
    "function saveTerminalNotice(a,b,c,d){return UnkouLib.saveTerminalNotice(a,b,c,d);}",
    "function uploadTerminalFile(a,b,c,d){return UnkouLib.uploadTerminalFile(a,b,c,d);}",
    "function getMyNotices(a,b){return UnkouLib.getMyNotices(a,b);}",
    "function getRoutesById(a,b,c){return UnkouLib.getRoutesById(a,b,c);}",
    "function getNoticeByRow(a,b,c){return UnkouLib.getNoticeByRow(a,b,c);}",
    "function markAsRead(a,b){return UnkouLib.markAsRead(a,b);}",
    "function getReadNotices(a){return UnkouLib.getReadNotices(a);}",
    "function agreeContract(a,b,c,d){return UnkouLib.agreeContract(a,b,c,d);}",
    "function queueFileUpload(a,b,c,d){return UnkouLib.queueFileUpload(a,b,c,d);}",
    "function recordAction(a,b,c,d,e,f){return UnkouLib.recordAction(a,b,c,d,e,f);}",
    "function clearInspTime(a,b,c,d){return UnkouLib.clearInspTime(a,b,c,d);}",
    "function showInvoiceDialog(){return UnkouLib.showInvoiceDialog();}",
    "function generateInvoiceSheet(a,b,c,d){return UnkouLib.generateInvoiceSheet(a,b,c,d);}",
    "function showPaymentDialog(){return UnkouLib.showPaymentDialog();}",
    "function generatePaymentSheet(a,b,c,d,e){return UnkouLib.generatePaymentSheet(a,b,c,d,e);}",
    "function matchAndConfirmDispatch(){return UnkouLib.matchAndConfirmDispatch();}",
    "function showCsvImportDialogUnkou(){return UnkouLib.showCsvImportDialogUnkou();}",
    "function showCsvImportDialogMaster(){return UnkouLib.showCsvImportDialogMaster();}",
    "function showCsvImportDialogCust(){return UnkouLib.showCsvImportDialogCust();}",
    "function deleteBlankImportRows(){return UnkouLib.deleteBlankImportRows();}",
    "function getImportDictionary(a,b){return UnkouLib.getImportDictionary(a,b);}",
    "function importBulkRows(a,b,c){return UnkouLib.importBulkRows(a,b,c);}",
    "function saveImportAliases(a,b,c){return UnkouLib.saveImportAliases(a,b,c);}",
    "function showHatchuDocDialog(){return UnkouLib.showHatchuDocDialog();}",
    "function showShabanDocDialog(){return UnkouLib.showShabanDocDialog();}",
    "function sendDocumentEmail(a,b,c){return UnkouLib.sendDocumentEmail(a,b,c);}",
    "function markDocumentIssued(a,b){return UnkouLib.markDocumentIssued(a,b);}"
  ].join('\n');
}


// ================================================================
//  12-3b: クライアントSSシート初期化共通処理（initClientSSSheets_）
//  createCompanySpreadsheet_・initTemplateSS_ の両方から呼ばれる。
//  全シートを正しいヘッダー＋テストデータで再構築し、不要シートを削除する。
// ================================================================
function initClientSSSheets_(ss, companyName) {
  var unkouHeader = [
    'ID','区分','会社名','トン数','車種','車番','乗務員名','携帯番号','看板名','日付',
    '荷主','積地','降地','誘導時刻','積完時刻','休憩開始','休憩終了','降完時刻',
    '売上','請求(高速代)','実費(高速代)','合計(高速代)','備考','管理データ',
    '連絡(端末)','データ(端末)','点呼前完了','点呼後完了'
  ];
  var sumHeader = [
    'ID','区分','会社名','トン数','車種','車番','乗務員名','携帯番号','看板名',
    '日付','荷主','積地','降地','誘導時刻','積完時刻','休憩開始','休憩終了','降完時刻',
    '売上','請求(高速代)','実費(高速代)','合計(高速代)',
    '距離','燃費','ガソリン代','燃料代','支払い','経費合計','利益','備考',
    '仮日数','給料','％','有休手当','その他手当','点呼前完了','点呼後完了'
  ];
  var masterHeader = [
    '車両ID','運行状態','区分','会社名','看板名','トン数','車種','車番','乗務員名','携帯番号',
    'アドレス','燃費','備考','仮日数','給料','％','高速を引く（引くは〇、引かないは空欄）',
    '車両リース代','任意保険料','自賠責保険料','重量税積立','車検費積立',
    '整備費積立','タイヤ代積立','修理積立','駐車場代','ETCリース料',
    'カーナビリース料','通信費','洗車費','制服費','その他固定費'
  ];
  var activeHeader = [
    '車両ID','運行状態','区分','会社名','看板名','トン数','車種','車番',
    '乗務員名','携帯番号','アドレス','燃費','備考','仮日数','給料','％'
  ];
  var custHeader    = ['マスタID','会社名','電話','FAX','郵便番号','住所','代表者','配車担当','銀行名','支店名','種別','番号','名義','備考','インボイス登録番号','インボイス発行者名（自社名）','メールアドレス'];
  var settingHeader = ['トン数','基準燃費','有休設定','有休金額','業務前点検','業務後点検'];

  function resetSheet(name, header) {
    var s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    if (s.getMaxColumns() < header.length) {
      s.insertColumnsAfter(s.getMaxColumns(), header.length - s.getMaxColumns());
    }
    s.getRange(1, 1, 1, header.length).setValues([header]);
    if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
    s.setFrozenRows(1);
    return s;
  }

  var unkouSheet   = resetSheet('運行',         unkouHeader);
  var sumSheet     = resetSheet('集計表',        sumHeader);
  var masterSheet  = resetSheet('自車専属マスタ', masterHeader);
                     resetSheet('自車専属運行',  activeHeader);
  var custSheet    = resetSheet('マスタ',        custHeader);
  var settingSheet = resetSheet('設定',          settingHeader);

  // テストデータ（各シート）
  unkouSheet.getRange(2, 1, 1, 28).setValues([[
    'V-0001','自車','テスト商事','4t','平車','品川1234','山田太郎','090-0000-0000','テスト看板',new Date(),
    'テスト荷主','テスト積地','テスト降地','09:00','10:00','12:00','13:00','15:00',
    50000,1000,1000,2000,'テスト備考','','','','',''
  ]]);
  unkouSheet.getRange(2, 10).setNumberFormat('yyyy/MM/dd');

  sumSheet.getRange(2, 1, 1, 37).setValues([[
    'V-0001','自車','テスト商事','4t','平車','品川1234','山田太郎','090-0000-0000','テスト看板',
    new Date(),'テスト荷主','テスト積地','テスト降地','09:00','10:00','12:00','13:00','15:00',
    50000,1000,1000,2000,100,6.5,1538,1538,0,1538,50462,'テスト備考',25,200000,'','0','','',''
  ]]);
  sumSheet.getRange(2, 10).setNumberFormat('yyyy/MM/dd');

  // 給料行・%行を別々に作成（給料と%はどちらか一方のみ有効）
  masterSheet.getRange(2, 1, 2, 32).setValues([
    ['S-0001','運行','自車','テスト商事','テスト看板（給料）','4t','平車','品川1234','山田太郎','090-0000-0000',
     'test1@example.com',6.5,'テスト備考',25,200000,'','〇',
     30000,15000,5000,10000,8000,5000,5000,3000,5000,3000,2000,2000,1000,1000,5000],
    ['S-0002','待機','自車','テスト商事','テスト看板（％）','4t','平車','品川5678','鈴木花子','090-1111-1111',
     'test2@example.com',6.5,'テスト備考（%）',25,'',75,'',
     30000,15000,5000,10000,8000,5000,5000,3000,5000,3000,2000,2000,1000,1000,5000]
  ]);
  // 自車専属マスタ B列（運行状態）に 運行・故障・待機 のドロップダウンを設定
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['運行','故障','待機'], true)
    .setAllowInvalid(false)
    .build();
  masterSheet.getRange(2, 2, masterSheet.getMaxRows() - 1, 1).setDataValidation(statusRule);

  custSheet.getRange(2, 1, 1, 16).setValues([[
    'M-0001','テスト荷主','03-0000-0000','03-1111-1111','100-0001',
    '東京都テスト市テスト1-1-1','テスト代表','テスト配車担当',
    'テスト銀行','テスト支店','普通','1234567','テスト商事',
    'テスト備考','T1234567890123','テスト商事'
  ]]);

  settingSheet.getRange(2, 1, 25, 6).setValues([
    ['1t',  6.5, '有休', 8000, 'ブレーキの効き・踏みしろ（エア・液漏れ含む）', '車両・積載物の異常の有無（タイヤ・車体等）'],
    ['2t',  6.5, '',    '',   'タイヤの空気圧・溝の深さ・亀裂や損傷',         '事故・ヒヤリハットの有無'],
    ['3t',  4.5, '',    '',   'ホイールナットの緩み・脱落の確認',             '道路状況・運行状況の異常の有無'],
    ['4t',  4.5, '',    '',   'エンジンオイル・冷却水・ベルト類の確認',       '翌乗務員への引き継ぎ事項の有無'],
    ['5t',  3.5, '',    '',   'バッテリー液・ウォッシャー液の確認',           '運行記録（日報）の提出・乗務後点呼の実施'],
    ['6t',  3.5, '',    '',   '灯火類（ランプ・ウィンカー）の点灯・汚れ',     ''],
    ['7t',  3.5, '',    '',   'ワイパーの動作・払拭状態',                     ''],
    ['8t',  3.5, '',    '',   'エンジンのかかり具合・異音の確認',             ''],
    ['9t',  2.5, '',    '',   'ミラーの調整・シートベルトの装着',             ''],
    ['10t', 2.5, '',    '',   '乗務前点呼・アルコールチェックの実施',         ''],
    ['11t', 2.5, '',    '',   '',                                             ''],
    ['12t', 2.5, '',    '',   '',                                             ''],
    ['13t', 2.5, '',    '',   '',                                             ''],
    ['14t', 2.5, '',    '',   '',                                             ''],
    ['15t', 2.5, '',    '',   '',                                             ''],
    ['16t', 2.5, '',    '',   '',                                             ''],
    ['17t', 2.5, '',    '',   '',                                             ''],
    ['18t', 2.5, '',    '',   '',                                             ''],
    ['19t', 2.5, '',    '',   '',                                             ''],
    ['20t', 2.5, '',    '',   '',                                             ''],
    ['21t', 2.5, '',    '',   '',                                             ''],
    ['22t', 2.5, '',    '',   '',                                             ''],
    ['23t', 2.5, '',    '',   '',                                             ''],
    ['24t', 2.5, '',    '',   '',                                             ''],
    ['25t', 2.5, '',    '',   '',                                             '']
  ]);

  // __COMPANY_SS__ マーカー（非表示）→ onOpen で客用メニュー表示のトリガー
  var marker = ss.getSheetByName('__COMPANY_SS__') || ss.insertSheet('__COMPANY_SS__');
  marker.getRange(1, 1).setValue(companyName || 'テンプレート');
  if (!marker.isSheetHidden()) marker.hideSheet();

  // 不要シートを削除（取引先マスタ・会社登録・Sheet1 等）
  var validNames = ['運行','集計表','自車専属マスタ','自車専属運行','マスタ','設定','__COMPANY_SS__'];
  ss.getSheets().forEach(function(s) {
    if (validNames.indexOf(s.getName()) === -1 && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch(e) {}
    }
  });

  // 自車専属マスタの各行に運行状態に応じた背景色を付ける（onEditは発火しないので手動で実行）
  var mLastRow = masterSheet.getLastRow();
  if (mLastRow >= 2) {
    var mLastCol = masterSheet.getLastColumn();
    var mAllData = masterSheet.getRange(2, 1, mLastRow - 1, mLastCol).getValues();
    for (var mi = 0; mi < mAllData.length; mi++) {
      var mStatus = String(mAllData[mi][1] || '').trim();
      var mRowRange = masterSheet.getRange(mi + 2, 1, 1, mLastCol);
      if      (mStatus === '運行') { mRowRange.setBackground('#ffcdd2'); }
      else if (mStatus === '待機') { mRowRange.setBackground('#fff9c4'); }
      else if (mStatus === '故障') { mRowRange.setBackground('#c8e6c9'); }
      else                         { mRowRange.setBackground(null); }
    }

    // 自車専属運行を「運行」状態の行のみで再生成（refreshActiveVehiclesAuto_はgetActiveSpreadsheet依存のため直接実行）
    var aSheet = ss.getSheetByName('自車専属運行');
    if (aSheet) {
      var aHeader = ['車両ID','運行状態','区分','会社名','看板名','トン数','車種','車番','乗務員名','携帯番号','アドレス','燃費','備考','仮日数','給料','％'];
      var aRows = [aHeader];
      for (var ai = 0; ai < mAllData.length; ai++) {
        if (String(mAllData[ai][1] || '').trim() === '運行') {
          aRows.push(mAllData[ai].slice(0, 16));
        }
      }
      aSheet.clear();
      aSheet.getRange(1, 1, aRows.length, 16).setValues(aRows);
      aSheet.setFrozenRows(1);
    }
  }

  applySheetColors_(ss);
  ensureSettingItems_(ss);
}


// ================================================================
//  12-3c-1: ライブラリバージョン作成（createLibraryVersion_）  【大B / 中12】
//  Script API でライブラリの新バージョンを作成して番号を返す
// ================================================================
function createLibraryVersion_(description) {
  try {
    var token   = ScriptApp.getOAuthToken();
    var scriptId = ScriptApp.getScriptId();
    var hdrs    = { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' };
    var apiBase = 'https://script.googleapis.com/v1/projects/' + scriptId;

    // 古いバージョンを自動削除（190個超えたら最古のものを削除）
    try {
      var listResp = UrlFetchApp.fetch(apiBase + '/versions?pageSize=200', { headers: hdrs, muteHttpExceptions: true });
      var versions = JSON.parse(listResp.getContentText()).versions || [];
      if (versions.length >= 190) {
        // 番号の小さい順（古い順）にソートして古いものから削除
        versions.sort(function(a, b) { return a.versionNumber - b.versionNumber; });
        var toDelete = versions.slice(0, versions.length - 185); // 185個残して古いもの全削除
        toDelete.forEach(function(v) {
          UrlFetchApp.fetch(apiBase + '/versions/' + v.versionNumber,
            { method: 'delete', headers: hdrs, muteHttpExceptions: true });
        });
      }
    } catch(e2) {}

    var resp = UrlFetchApp.fetch(apiBase + '/versions', {
      method: 'post',
      headers: hdrs,
      payload: JSON.stringify({ description: description || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') }),
      muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    return data.versionNumber || null;
  } catch(e) { return null; }
}


// ================================================================
//  12-3c-2: スタブのライブラリバージョン更新（updateStubVersion_）  【大B / 中12】
//  Script API でスタブの appsscript.json 内ライブラリバージョンを書き換える
// ================================================================
function updateStubVersion_(stubScriptId, versionNumber, useDevMode) {
  try {
    var token = ScriptApp.getOAuthToken();
    var libId = ScriptApp.getScriptId();
    var getResp = UrlFetchApp.fetch(
      'https://script.googleapis.com/v1/projects/' + stubScriptId + '/content',
      { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
    );
    if (getResp.getResponseCode() !== 200) {
      return { ok: false, error: 'GET ' + getResp.getResponseCode() + ': ' + getResp.getContentText().slice(0, 200) };
    }
    var content = JSON.parse(getResp.getContentText());
    var files = content.files || [];
    for (var i = 0; i < files.length; i++) {
      if (files[i].name !== 'appsscript') continue;
      var manifest = JSON.parse(files[i].source);
      var libs = (manifest.dependencies && manifest.dependencies.libraries) || [];
      for (var j = 0; j < libs.length; j++) {
        if (libs[j].libraryId === libId) {
          libs[j].version = String(versionNumber);
          libs[j].developmentMode = false; // 常にfalse固定
        }
      }
      files[i].source = JSON.stringify(manifest, null, 2);
      break;
    }
    // スタブコード（SERVER_JS）も最新版に差し替え
    var stubSource = getClientStubSource_();
    var codeFound = false;
    for (var k = 0; k < files.length; k++) {
      if (files[k].type === 'SERVER_JS') {
        files[k].source = stubSource;
        codeFound = true;
        break;
      }
    }
    if (!codeFound) {
      files.push({ name: 'コード', type: 'SERVER_JS', source: stubSource });
    }
    var putResp = UrlFetchApp.fetch(
      'https://script.googleapis.com/v1/projects/' + stubScriptId + '/content',
      {
        method: 'put',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify({ files: files }),
        muteHttpExceptions: true
      }
    );
    if (putResp.getResponseCode() !== 200) {
      return { ok: false, error: 'PUT ' + putResp.getResponseCode() + ': ' + putResp.getContentText().slice(0, 200) };
    }
    return { ok: true, error: '' };
  } catch(e) {
    return { ok: false, error: e.message || String(e) };
  }
}


// ================================================================
//  12-3c: ①修正用SS→②客用SSに反映（syncToTemplateSS）
//  メニュー「📤 テスト客SSに反映」から実行。
//  ① 新しいライブラリバージョンを作成 → ②のスタブを固定バージョンに更新
//     （押すまでコード変更が②に波及しない。押して初めて反映）
//  ② ②のシート構成・テストデータ・設定を全て最新化
// ================================================================
// ================================================================
//  13-1: 請求書生成ダイアログ（showInvoiceDialog）
// ================================================================
function showInvoiceDialog() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var mSh  = ss.getSheetByName('マスタ');
  var cos  = [];
  if (mSh && mSh.getLastRow() >= 2) {
    mSh.getRange(2, 2, mSh.getLastRow()-1, 1).getValues()
      .forEach(function(r){ if(r[0]) cos.push(String(r[0])); });
  }
  var today    = new Date();
  var firstDay = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth(), 1),     'Asia/Tokyo', 'yyyy-MM-dd');
  var lastDay  = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth()+1, 0),   'Asia/Tokyo', 'yyyy-MM-dd');
  var opts = cos.map(function(c){ return '<option>'+c+'</option>'; }).join('');
  var html = '<html><body style="font-family:sans-serif;padding:16px;font-size:13px;">'
    +'<table style="border-collapse:collapse;width:100%">'
    +'<tr><td style="padding:6px">宛先（荷主）</td><td><select id="co" style="width:200px">'+opts+'</select></td></tr>'
    +'<tr><td style="padding:6px">期間</td><td><input type="date" id="f" value="'+firstDay+'" style="width:130px"> 〜 <input type="date" id="t" value="'+lastDay+'" style="width:130px"></td></tr>'
    +'<tr><td style="padding:6px">消費税率(%)</td><td><input type="number" id="tax" value="10" style="width:60px"></td></tr>'
    +'</table>'
    +'<br><button onclick="g()" style="background:#1565c0;color:#fff;padding:8px 24px;border:none;border-radius:6px;cursor:pointer;font-size:13px">生成</button>'
    +'<span id="m" style="margin-left:10px;color:#888"></span>'
    +'<script>function g(){document.getElementById("m").innerText="生成中...";'
    +'google.script.run.withSuccessHandler(function(){document.getElementById("m").innerText="完了";setTimeout(function(){google.script.host.close();},800);})'
    +'.withFailureHandler(function(e){document.getElementById("m").innerText="エラー: "+(e.message||e);document.getElementById("m").style.color="red";})'
    +'.generateInvoiceSheet(document.getElementById("co").value,document.getElementById("f").value,document.getElementById("t").value,Number(document.getElementById("tax").value));}'
    +'</script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(620).setHeight(260), '請求書生成');
}


// ================================================================
//  13-2: 支払確認書生成ダイアログ（showPaymentDialog）
// ================================================================
function showPaymentDialog() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sumSh = ss.getSheetByName('集計表');
  var companies = [], cars = [], names = [];
  if (sumSh && sumSh.getLastRow() >= 2) {
    var sv = sumSh.getRange(2, 1, sumSh.getLastRow()-1, 7).getValues();
    sv.forEach(function(r) {
      var co = String(r[2]||'').trim(), ca = String(r[5]||'').trim(), nm = String(r[6]||'').trim();
      if (co && companies.indexOf(co) === -1) companies.push(co);
      if (ca && cars.indexOf(ca) === -1) cars.push(ca);
      if (nm && names.indexOf(nm) === -1) names.push(nm);
    });
  }
  var today    = new Date();
  var firstDay = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth(), 1),   'Asia/Tokyo', 'yyyy-MM-dd');
  var lastDay  = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth()+1, 0), 'Asia/Tokyo', 'yyyy-MM-dd');
  function mkOpts(arr){ return '<option value="">(全て)</option>'+arr.map(function(v){ return '<option>'+v+'</option>'; }).join(''); }
  var html = '<html><body style="font-family:sans-serif;padding:16px;font-size:13px;">'
    +'<table style="border-collapse:collapse;width:100%">'
    +'<tr><td style="padding:6px">会社名</td><td><select id="co" style="width:200px">'+mkOpts(companies)+'</select></td></tr>'
    +'<tr><td style="padding:6px">車番</td><td><select id="ca" style="width:200px">'+mkOpts(cars)+'</select></td></tr>'
    +'<tr><td style="padding:6px">乗務員名</td><td><select id="nm" style="width:200px">'+mkOpts(names)+'</select></td></tr>'
    +'<tr><td style="padding:6px">期間</td><td><input type="date" id="f" value="'+firstDay+'" style="width:130px"> 〜 <input type="date" id="t" value="'+lastDay+'" style="width:130px"></td></tr>'
    +'</table>'
    +'<br><button onclick="g()" style="background:#1b5e20;color:#fff;padding:8px 24px;border:none;border-radius:6px;cursor:pointer;font-size:13px">生成</button>'
    +'<span id="m" style="margin-left:10px;color:#888"></span>'
    +'<script>function g(){document.getElementById("m").innerText="生成中...";'
    +'google.script.run'
    +'.withSuccessHandler(function(){document.getElementById("m").innerText="完了";setTimeout(function(){google.script.host.close();},800);})'
    +'.withFailureHandler(function(e){document.getElementById("m").innerText="エラー: "+(e.message||e);document.getElementById("m").style.color="red";})'
    +'.generatePaymentSheet(document.getElementById("co").value,document.getElementById("ca").value,document.getElementById("nm").value,document.getElementById("f").value,document.getElementById("t").value);}'
    +'</script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(620).setHeight(280), '支払確認書生成');
}


// ================================================================
//  13-3: 書類用連番採番（getNextDocNum_）
// ================================================================
function getNextDocNum_(type) {
  var props = PropertiesService.getScriptProperties();
  var ym    = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMM');
  var key   = 'docnum_' + type + '_' + ym;
  var n     = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(n));
  return ym + '-' + String(n).padStart(4, '0');
}


// ================================================================
//  13-4: 請求書シート生成（generateInvoiceSheet）
// ================================================================
function generateInvoiceSheet(company, dateFrom, dateTo, taxRate) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sumSh    = ss.getSheetByName('集計表');
  if (!sumSh || sumSh.getLastRow() < 2) { ss.toast('集計表にデータがありません', '⚠️', 4); return; }

  // 期間フィルタ
  var from = new Date(dateFrom + 'T00:00:00');
  var to   = new Date(dateTo   + 'T23:59:59');

  // 集計表読み込み・絞り込み（荷主 + 日付）
  var rows = sumSh.getRange(2, 1, sumSh.getLastRow()-1, 37).getValues();
  var items = rows.filter(function(r) {
    var d = r[9]; if (!(d instanceof Date)) return false;
    var client = String(r[10]||'').trim();
    return client.indexOf(company.trim()) !== -1 && d >= from && d <= to
      && String(r[11]||'').trim() !== '' // 積地あり（実稼働行のみ）
      && String(r[11]||'').indexOf('有休') === -1 && String(r[11]||'').indexOf('休み') === -1;
  }).sort(function(a,b){ return a[9]-b[9]; });

  // 荷主FAX
  var mSh   = ss.getSheetByName('マスタ'), faxNo = '';
  if (mSh && mSh.getLastRow() >= 2) {
    var mVals = mSh.getRange(2,1,mSh.getLastRow()-1,6).getValues();
    for (var mi=0;mi<mVals.length;mi++){
      if (String(mVals[mi][1]||'').trim()===company.trim()){ faxNo=String(mVals[mi][3]||''); break; }
    }
  }

  // マスタシートのO列(15)・P列(16)からインボイス情報を取得（記載あれば適格請求書形式）
  var invoiceRegNum = '', invoiceName = '';
  var mInvSh = ss.getSheetByName('マスタ');
  if (mInvSh) {
    if (!mInvSh.getRange(1, 15).getValue()) {
      mInvSh.getRange(1, 15).setValue('インボイス登録番号').setFontWeight('bold');
      mInvSh.getRange(1, 16).setValue('インボイス発行者名（自社名）').setFontWeight('bold');
    }
    if (mInvSh.getLastRow() >= 2) {
      var invRows = mInvSh.getRange(2, 15, mInvSh.getLastRow()-1, 2).getValues();
      for (var ir = 0; ir < invRows.length; ir++) {
        var regNum = String(invRows[ir][0]||'').trim();
        var issuer = String(invRows[ir][1]||'').trim();
        if (regNum && issuer) { invoiceRegNum = regNum; invoiceName = issuer; break; }
      }
    }
  }
  var hasInvoice = invoiceRegNum !== '' && invoiceName !== '';

  // 書類シート準備
  var sh = ss.getSheetByName('請求書') || ss.insertSheet('請求書');
  sh.clear(); sh.clearFormats(); sh.clearNotes();
  [40,70,130,130,70,80,80].forEach(function(w,i){ sh.setColumnWidth(i+1,w); });

  var today   = new Date();
  var docNum  = 'R-' + getNextDocNum_('inv');
  var issued  = Utilities.formatDate(today,'Asia/Tokyo','yyyy年MM月dd日');
  var pFrom   = Utilities.formatDate(from,'Asia/Tokyo','yyyy年M月d日');
  var pTo     = Utilities.formatDate(to,  'Asia/Tokyo','yyyy年M月d日');
  var tax     = Number(taxRate)||10;

  var totalSales = 0, totalToll = 0;
  items.forEach(function(r){ totalSales+=Number(r[18])||0; totalToll+=Number(r[19])||0; });
  var taxAmt   = Math.round(totalSales*tax/100);
  var grandTotal = totalSales + totalToll + taxAmt;

  var R = 1;
  function cell(r,c){ return sh.getRange(r,c); }
  function merge(r,c,nr,nc){ return sh.getRange(r,c,nr,nc).merge(); }

  // タイトル
  merge(R,1,1,7).setValue('請　求　書').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  R++;
  if (hasInvoice) {
    merge(R,1,1,4).setValue(invoiceName).setFontWeight('bold');
    R++;
    merge(R,1,1,4).setValue('登録番号: '+invoiceRegNum).setFontColor('#555');
    R++;
  }
  if (faxNo) {
    cell(R,6).setValue('FAX').setHorizontalAlignment('right').setFontColor('#555');
    cell(R,7).setValue(faxNo);
    R++;
  }
  R++;
  // 宛先
  merge(R,1,1,4).setValue(company+' 御中').setFontSize(14).setFontWeight('bold');
  R++;
  merge(R,1,1,7).setValue('期間：'+pFrom+' 〜 '+pTo).setFontColor('#444');
  R+=2;
  // ご請求金額ボックス
  merge(R,1,1,7).setValue('ご請求金額：¥'+grandTotal.toLocaleString()+'（税込）')
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#e3f2fd').setFontColor('#0d47a1')
    .setBorder(true,true,true,true,null,null,'#1565c0',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  R+=2;
  // テーブルヘッダー
  sh.getRange(R,1,1,7).setValues([['No.','日付','積地','降地','車番','売上','高速代']])
    .setBackground('#1565c0').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  R++;
  // 明細
  items.forEach(function(r,i){
    var dStr = Utilities.formatDate(r[9],'Asia/Tokyo','M/d');
    var bg   = i%2===0 ? null : '#f5f5f5';
    sh.getRange(R,1,1,7).setValues([[i+1,dStr,String(r[11]||''),String(r[12]||''),String(r[5]||''),Number(r[18])||0,Number(r[19])||0]]);
    sh.getRange(R,1,1,7).setBackground(bg).setHorizontalAlignment('left');
    sh.getRange(R,6,1,2).setNumberFormat('#,##0');
    R++;
  });
  // 集計行
  R++;
  merge(R,3,1,3).setValue('売上合計').setHorizontalAlignment('right').setFontWeight('bold');
  merge(R,6,1,2).setValue(totalSales).setNumberFormat('#,##0').setHorizontalAlignment('right').setFontWeight('bold');
  R++;
  merge(R,3,1,3).setValue('高速代合計').setHorizontalAlignment('right').setFontColor('#555');
  merge(R,6,1,2).setValue(totalToll).setNumberFormat('#,##0').setHorizontalAlignment('right').setFontColor('#555');
  R++;
  merge(R,3,1,3).setValue('消費税（'+tax+'%）').setHorizontalAlignment('right').setFontColor('#555');
  merge(R,6,1,2).setValue(taxAmt).setNumberFormat('#,##0').setHorizontalAlignment('right').setFontColor('#555');
  R++;
  sh.getRange(R,3,1,5).setBackground('#e8f5e9').setBorder(true,true,true,true,null,null,'#2e7d32',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  merge(R,3,1,3).setValue('合計').setHorizontalAlignment('right').setFontWeight('bold').setFontSize(12);
  merge(R,6,1,2).setValue(grandTotal).setNumberFormat('#,##0').setHorizontalAlignment('right').setFontWeight('bold').setFontSize(12);
  R+=2;
  merge(R,1,1,7).setValue('※ご請求に関するお問い合わせはご連絡ください。').setFontColor('#999').setFontSize(9);

  sh.setHiddenGridlines(true);
  ss.setActiveSheet(sh);
  ss.toast('請求書を生成しました（'+items.length+'件）', '完了', 4);
}


// ================================================================
//  13-5: 支払確認書シート生成（generatePaymentSheet）
// ================================================================
function generatePaymentSheet(company, carNo, driverName, dateFrom, dateTo) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sumSh = ss.getSheetByName('集計表');
  if (!sumSh || sumSh.getLastRow() < 2) { ss.toast('集計表にデータがありません', '⚠️', 4); return; }

  var from = new Date(dateFrom+'T00:00:00');
  var to   = new Date(dateTo  +'T23:59:59');

  var rows = sumSh.getRange(2, 1, sumSh.getLastRow()-1, 37).getValues();
  var items = rows.filter(function(r){
    var d = r[9]; if(!(d instanceof Date)) return false;
    if (d < from || d > to) return false;
    if (company    && String(r[2]||'').trim() !== company.trim())    return false;
    if (carNo      && String(r[5]||'').trim() !== carNo.trim())      return false;
    if (driverName && String(r[6]||'').trim() !== driverName.trim()) return false;
    if (!String(r[11]||'').trim()) return false;
    if (String(r[11]||'').indexOf('有休') !== -1) return false;
    if (String(r[11]||'').indexOf('休み') !== -1) return false;
    return true;
  }).sort(function(a,b){ return a[9]-b[9]; });

  var sh = ss.getSheetByName('支払確認書') || ss.insertSheet('支払確認書');
  sh.clear(); sh.clearFormats(); sh.clearNotes();
  [40,70,120,120,100,70,80,80].forEach(function(w,i){ sh.setColumnWidth(i+1,w); });

  var today  = new Date();
  var docNum = 'S-' + getNextDocNum_('pay');
  var issued = Utilities.formatDate(today,'Asia/Tokyo','yyyy年MM月dd日');
  var pFrom  = Utilities.formatDate(from,'Asia/Tokyo','yyyy年M月d日');
  var pTo    = Utilities.formatDate(to,  'Asia/Tokyo','yyyy年M月d日');

  var totalPay = 0, totalToll = 0;
  items.forEach(function(r){ totalPay+=Number(r[26])||0; totalToll+=Number(r[20])||0; });
  var grandTotal = totalPay + totalToll;

  var titleName = driverName || carNo || company || '全乗務員';
  var carNos = [];
  items.forEach(function(r){ var c = String(r[5]||'').trim(); if(c && carNos.indexOf(c)===-1) carNos.push(c); });

  var R = 1;
  function merge(r,c,nr,nc){ return sh.getRange(r,c,nr,nc).merge(); }
  function cell(r,c){ return sh.getRange(r,c); }

  merge(R,1,1,8).setValue('支　払　確　認　書').setFontSize(18).setFontWeight('bold').setHorizontalAlignment('center');
  R+=2;
  merge(R,1,1,6).setValue(titleName+' 様').setFontSize(14).setFontWeight('bold');
  R++;
  if (company) { merge(R,1,1,6).setValue('会社名：'+company).setFontColor('#444'); R++; }
  if (carNos.length > 0) { merge(R,1,1,6).setValue('車番：'+carNos.join(' / ')).setFontColor('#444'); R++; }
  merge(R,1,1,8).setValue('期間：'+pFrom+' 〜 '+pTo).setFontColor('#444');
  R+=2;
  merge(R,1,1,8).setValue('お支払金額：¥'+grandTotal.toLocaleString())
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackground('#e8f5e9').setFontColor('#1b5e20')
    .setBorder(true,true,true,true,null,null,'#2e7d32',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  R+=2;
  sh.getRange(R,1,1,8).setValues([['No.','日付','積地','降地','看板名','車番','支払い','実費高速']])
    .setBackground('#1b5e20').setFontColor('#fff').setFontWeight('bold').setHorizontalAlignment('center');
  R++;
  items.forEach(function(r,i){
    var dStr = Utilities.formatDate(r[9],'Asia/Tokyo','M/d');
    var bg   = i%2===0 ? null : '#f5f5f5';
    sh.getRange(R,1,1,8).setValues([[i+1,dStr,String(r[11]||''),String(r[12]||''),String(r[8]||''),String(r[5]||''),Number(r[26])||0,Number(r[20])||0]]);
    sh.getRange(R,1,1,8).setBackground(bg).setHorizontalAlignment('left');
    sh.getRange(R,7,1,2).setNumberFormat('#,##0');
    R++;
  });
  R++;
  merge(R,4,1,3).setValue('支払合計').setHorizontalAlignment('right').setFontWeight('bold');
  merge(R,7,1,2).setValue(totalPay).setNumberFormat('#,##0').setHorizontalAlignment('right').setFontWeight('bold');
  R++;
  merge(R,4,1,3).setValue('実費高速計').setHorizontalAlignment('right').setFontColor('#555');
  merge(R,7,1,2).setValue(totalToll).setNumberFormat('#,##0').setHorizontalAlignment('right').setFontColor('#555');
  R++;
  sh.getRange(R,4,1,5).setBackground('#e3f2fd').setBorder(true,true,true,true,null,null,'#1565c0',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  merge(R,4,1,3).setValue('合計').setHorizontalAlignment('right').setFontWeight('bold').setFontSize(12);
  merge(R,7,1,2).setValue(grandTotal).setNumberFormat('#,##0').setHorizontalAlignment('right').setFontWeight('bold').setFontSize(12);

  sh.setHiddenGridlines(true);
  ss.setActiveSheet(sh);
  ss.toast('支払確認書を生成しました（'+items.length+'件）', '完了', 4);
}


// ================================================================
function syncToTemplateSS() {
  var props = PropertiesService.getScriptProperties();
  var TEMPLATE_SCRIPT_ID = '19CfyUPhldzSccj05xo-sn4Xh78fCHAHDVJtGyKdDGQkO1D4wZWFEnZCT';
  var templateSsId = props.getProperty('clientTemplateSsId') || '1NBtosd_MN8KcboV_4OXTrY8WqcE3TJwpxdA_nASmTOo';
  var masterSs = SpreadsheetApp.getActiveSpreadsheet();
  var tgtSs    = SpreadsheetApp.openById(templateSsId);
  DriveApp.getFileById(tgtSs.getId()).setName('客用');

  // ① バージョン作成（60分以内に作成済みなら再利用してバージョン消費を節約）
  var lastVerTime = Number(props.getProperty('lastLibVersionTime') || 0);
  var lastVerNum  = Number(props.getProperty('lastLibVersionNum')  || 0);
  var newVersion;
  if (lastVerNum > 0 && (Date.now() - lastVerTime) < 60 * 60 * 1000) {
    newVersion = lastVerNum;
  } else {
    newVersion = createLibraryVersion_(
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') + ' テスト客SS反映'
    );
    if (newVersion) {
      props.setProperty('lastLibVersionTime', String(Date.now()));
      props.setProperty('lastLibVersionNum',  String(newVersion));
    }
  }
  if (newVersion) {
    var tmplStubResult = updateStubVersion_(TEMPLATE_SCRIPT_ID, newVersion);
    if (tmplStubResult && tmplStubResult.ok) {
      props.setProperty('approvedLibVersion', String(newVersion));
    } else {
      try { SpreadsheetApp.getUi().alert('②客用SSへのスタブ更新に失敗しました。\n' + (tmplStubResult ? tmplStubResult.error : '不明')); } catch(e) {}
      return;
    }
  }

  // ② ヘッダー行（1行目）の値・書式のみ最新化。データ行は絶対触らない。
  var businessSheets = ['運行', '集計表', '自車専属マスタ', '自車専属運行', 'マスタ', '設定'];
  for (var si = 0; si < businessSheets.length; si++) {
    var sheetName = businessSheets[si];
    var srcSheet  = masterSs.getSheetByName(sheetName);
    if (!srcSheet || srcSheet.getLastColumn() === 0) continue;
    var tgtSheet = tgtSs.getSheetByName(sheetName);
    if (!tgtSheet) tgtSheet = tgtSs.insertSheet(sheetName);
    var srcCols = srcSheet.getLastColumn();
    tgtSheet.getRange(1, 1, 1, srcCols).setValues(srcSheet.getRange(1, 1, 1, srcCols).getValues());
    tgtSheet.getRange(1, 1, 1, srcCols).setBackgrounds(srcSheet.getRange(1, 1, 1, srcCols).getBackgrounds());
    tgtSheet.getRange(1, 1, 1, srcCols).setFontColors(srcSheet.getRange(1, 1, 1, srcCols).getFontColors());
    tgtSheet.getRange(1, 1, 1, srcCols).setFontWeights(srcSheet.getRange(1, 1, 1, srcCols).getFontWeights());
    tgtSheet.setFrozenRows(1);
  }

  // __COMPANY_SS__ を削除して __TEMPLATE_SS__ マーカーに切り替え
  var companyMarker = tgtSs.getSheetByName('__COMPANY_SS__');
  if (companyMarker && tgtSs.getSheets().length > 1) {
    try { tgtSs.deleteSheet(companyMarker); } catch(e) {}
  }
  var tmplMarker = tgtSs.getSheetByName('__TEMPLATE_SS__') || tgtSs.insertSheet('__TEMPLATE_SS__');
  tmplMarker.getRange(1, 1).setValue('客用テスト');
  tmplMarker.getRange(1, 2).setValue(masterSs.getId());
  if (newVersion) tmplMarker.getRange(1, 3).setValue(newVersion);
  if (!tmplMarker.isSheetHidden()) tmplMarker.hideSheet();

  // 保護をすべて削除
  tgtSs.getSheets().forEach(function(s) {
    s.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) { p.remove(); });
    s.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); });
  });

  ensureSettingItems_(tgtSs);

  props.setProperty('clientTemplateSsId', tgtSs.getId());

  var msg = 'テスト客SSへの反映が完了しました。';
  if (newVersion) msg += '\nライブラリバージョン: ' + newVersion;
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}


// ================================================================
//  12-4: 新規会社フルセットアップ（processNewCompany_）  【大B / 中12 / 小12-4】
//  ① 共有フォルダ（運行管理_アーカイブ/会社名/）を作成して管理Gmailに共有
//  ② コードなしSSをそのフォルダ内に作成
//  ③ アプリURL（WebアプリURL?ssId=会社SS_ID）を生成
//  ④ 契約書URL（?page=contract&ssId=...&company=...&row=...）を生成
//  ⑤ 会社登録シートにSSURL・AppURL・フォルダURL・契約書URLを記録（I列）
//  ⑥ 管理Gmail宛に「契約書確認のお願い」メールを送信
//  メニュー「📤 会社SS作成＆メール送信」から手動実行。
// ================================================================
function processNewCompany_(companyName, adminEmail) {
  var ss;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(ex) {}
  if (!ss) {
    var mid = PropertiesService.getScriptProperties().getProperty('masterSsId');
    if (mid) ss = SpreadsheetApp.openById(mid);
  }
  if (!ss) throw new Error('マスターSSが見つかりません');
  var regSsId = ss.getId(); // 会社登録シートがあるSS（修正用SS）のID

  var regSheet = ss.getSheetByName('会社登録');

  // 会社登録シートのI列・J列ヘッダーを確保
  if (regSheet && regSheet.getLastRow() >= 1) {
    var hdrColCount = regSheet.getLastColumn();
    if (hdrColCount < 9)  regSheet.getRange(1, 9).setValue('契約書URL');
    if (hdrColCount < 10) regSheet.getRange(1, 10).setValue('同意時刻');
    if (hdrColCount < 11) regSheet.getRange(1, 11).setValue('スクリプトID');
  }

  // ① 共有フォルダを作成（メール送信なし）
  var folderResult = setupOneCompany_(companyName, adminEmail, true);
  var folderUrl = folderResult.folderUrl;
  var folderId  = folderResult.folderId;

  // ② ②客用SSをコピーして③各客SS作成（スタブコードのみ・心臓部コードは含まれない）
  var ssResult   = createCompanySpreadsheet_(companyName, adminEmail, folderId);
  var ssUrl      = ssResult.ssUrl;
  var ssId       = ssResult.ssId;
  var clientScriptId = ssResult.scriptId || '';

  // ③ アプリURL = 会社SS独自のWebAppURL（ssId不要・独立URL）
  var appUrl = ssResult.webAppUrl || (getWebAppBaseUrl_() + '?ssId=' + ssId);

  // ④ 会社登録シートの行番号を特定
  var targetRow = -1;
  if (regSheet && regSheet.getLastRow() >= 2) {
    var rows = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === companyName) { targetRow = i + 2; break; }
    }
  }

  // ④ 契約書URL = 会社SS独自WebAppURL + page=contract（ssId不要）
  var contractUrl = '[WebアプリURL未設定]';
  if (ssResult.webAppUrl) {
    contractUrl = ssResult.webAppUrl + '?page=contract' +
      '&company='    + encodeURIComponent(companyName) +
      '&adminEmail=' + encodeURIComponent(adminEmail) +
      (targetRow > 0 ? '&row=' + targetRow : '');
  } else {
    var baseUrl = getWebAppBaseUrl_();
    if (baseUrl) {
      contractUrl = baseUrl + '?page=contract' +
        '&ssId='       + encodeURIComponent(ssId) +
        '&company='    + encodeURIComponent(companyName) +
        '&adminEmail=' + encodeURIComponent(adminEmail) +
        (targetRow > 0 ? '&row=' + targetRow : '');
    }
  }

  // ⑤ 会社登録シートに記録
  if (targetRow > 0) {
    regSheet.getRange(targetRow, 3).setValue('契約書送信済').setBackground('#fff9c4').clearNote();
    regSheet.getRange(targetRow, 4).setValue(new Date());
    regSheet.getRange(targetRow, 5).setValue(folderUrl);
    regSheet.getRange(targetRow, 6).setValue(ssUrl);
    regSheet.getRange(targetRow, 7).setValue(appUrl);
    regSheet.getRange(targetRow, 9).setValue(contractUrl);
    if (clientScriptId) regSheet.getRange(targetRow, 11).setValue(clientScriptId);
  }

  // ⑥ 管理Gmail宛に「契約書確認のお願い」メールを送信
  var subject = '[運行管理] ' + companyName + ' 利用規約への同意のお願い';
  var body =
    companyName + ' ご担当者様\n\n' +
    'このたびは運行管理システムへのお申し込みありがとうございます。\n\n' +
    '下記URLより利用規約をご確認いただき、「同意する」ボタンを押してください。\n' +
    '同意完了後に、スプレッドシートおよびアプリのURLをメールでお送りします。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 利用規約・同意フォーム\n' +
    contractUrl + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    '何かご不明な点があればお気軽にお問い合わせください。\n' +
    'よろしくお願いいたします。';
  try {
    GmailApp.sendEmail(adminEmail, subject, body);
    if (targetRow > 0) {
      regSheet.getRange(targetRow, 8).setValue('契約書メール送信済(' +
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') + ')')
        .setBackground('#fff9c4');
    }
  } catch(e) {}

  return { ssId: ssId, ssUrl: ssUrl, appUrl: appUrl, folderUrl: folderUrl, contractUrl: contractUrl };
}


// ================================================================
//  12-4b: 契約書同意処理（agreeContract）  【大B / 中12 / 小12-4b】
//  contract.html の「同意する」ボタンから google.script.run で呼ばれる。
//  ① 会社登録シートのJ列(10)に同意時刻を記録・C列を「同意済」に更新
//  ② 管理Gmail宛にSS URLとアプリURLをメール送信
//  ③ H列(8)に送信済ステータスを記録
// ================================================================
function agreeContract(ssId, companyName, adminEmail, contractRow) {
  var masterSsId = PropertiesService.getScriptProperties().getProperty('masterSsId');
  if (!masterSsId) throw new Error('マスターSSが見つかりません');

  var ss = SpreadsheetApp.openById(masterSsId);
  var regSheet = ss.getSheetByName('会社登録');
  if (!regSheet) throw new Error('会社登録シートが見つかりません');

  var row = parseInt(contractRow, 10);
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');

  var ssUrl  = '';
  var appUrl = '';
  if (row >= 2) {
    ssUrl  = String(regSheet.getRange(row, 6).getValue() || '');
    appUrl = String(regSheet.getRange(row, 7).getValue() || '');
    regSheet.getRange(row, 10).setValue(now).setBackground('#c8e6c9');
    regSheet.getRange(row, 3).setValue('同意済').setBackground('#c8e6c9');
  }

  if (adminEmail && adminEmail.indexOf('@') !== -1) {
    var subject = '[運行管理] ' + companyName + ' 運行管理システム利用開始のご案内';
    var body =
      companyName + ' ご担当者様\n\n' +
      '利用規約へのご同意ありがとうございます。\n\n' +
      '以下のURLよりご利用を開始いただけます。\n\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      '■ 運行管理スプレッドシート（PC・タブレット推奨）\n' +
      ssUrl + '\n' +
      '　→ 配車・運行データの入力・月生成・集計表の確認はこちら\n\n' +
      '■ 運行管理アプリ（乗務員用 スマートフォン推奨）\n' +
      appUrl + '\n' +
      '　→ 乗務員がスマートフォンから運行状況を入力するアプリです\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
      '【スプレッドシートの使い方】\n' +
      '・「自車専属マスタ」タブのK列（アドレス）に乗務員メールを入力してください\n\n' +
      '【乗務員への配布方法】\n' +
      '・各乗務員に上記アプリURLを共有してください\n' +
      '・初回は「紐づけ設定」でメールアドレスを登録するだけで使えます\n\n' +
      'ご不明な点はお気軽にお問い合わせください。\n' +
      'よろしくお願いいたします。';
    GmailApp.sendEmail(adminEmail, subject, body);
  }

  if (row >= 2) {
    regSheet.getRange(row, 8).setValue('送信済(' + now + ')').setBackground('#c8e6c9');
  }

  return { success: true, ssUrl: ssUrl, appUrl: appUrl };
}


// ================================================================
//  12-5: 全未処理会社のSS作成＆メール送信（sendCompanySetupEmails）  【大C / 中12 / 小12-5】
//  会社登録シートのF列(SS URL)が空の行を対象に processNewCompany_ を実行する。
//  メニュー「📤 会社SS作成＆メール送信」から手動実行。
// ================================================================
function sendCompanySetupEmails() {
  var ui    = SpreadsheetApp.getUi();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('会社登録');

  if (!sheet) {
    ui.alert('「会社登録」シートがありません。\n先に「🏢 会社セットアップ実行」を実行してください。');
    return;
  }

  // WebアプリURL未設定の場合は確認
  if (!getWebAppBaseUrl_()) {
    var ans = ui.alert(
      'WebアプリURL未設定',
      '先にメニュー「⚙ WebアプリURLを設定」でURLを登録してください。\n\n' +
      'このまま続けるとメール内のアプリURLが「[WebアプリURL未設定]」になります。続けますか？',
      ui.ButtonSet.YES_NO
    );
    if (ans !== ui.Button.YES) return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { ui.alert('会社登録シートにデータがありません。'); return; }

  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var done = 0, skip = 0, errCount = 0;

  for (var i = 0; i < data.length; i++) {
    var companyName = String(data[i][0] || '').trim();
    var adminEmail  = String(data[i][1] || '').trim();
    var ssUrlExist  = String(data[i][5] || '').trim(); // F列: SS URL

    if (!companyName || !adminEmail || adminEmail.indexOf('@') === -1) { skip++; continue; }
    if (ssUrlExist) { skip++; continue; } // 既にSS作成済みならスキップ

    try {
      processNewCompany_(companyName, adminEmail);
      done++;
    } catch(e) {
      sheet.getRange(i + 2, 3).setValue('エラー: ' + e.message).setBackground('#ffcdd2');
      errCount++;
    }
  }

  ui.alert(
    'SS作成＆メール送信 完了\n\n' +
    '処理: ' + done + '社\n' +
    'スキップ（SS作成済み or 未入力）: ' + skip + '社\n' +
    'エラー: ' + errCount + '社'
  );
}


// ================================================================
//  12-6: 申し込みフォーム作成（createSignupForm）  【大C / 中12 / 小12-6】
//  Google フォームを作成して申し込み受付を自動化する。
//  フォーム送信時に onFormSubmit_ が自動実行されるようトリガーも設定する。
//  メニュー「📝 申し込みフォーム作成」から1回だけ実行する。
// ================================================================
function createSignupForm() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var form = FormApp.create('運行管理システム 申し込みフォーム');
  form.setDescription(
    '運行管理システムへのお申し込みはこちらからどうぞ。\n' +
    '送信後、スプレッドシートとアプリのURLをメールでお送りします。'
  );
  form.setConfirmationMessage(
    'お申し込みありがとうございます！\n' +
    '担当者より数日以内にメールにてご連絡いたします。'
  );

  form.addTextItem().setTitle('会社名').setRequired(true);
  form.addTextItem()
    .setTitle('担当者メールアドレス（Gmail）')
    .setHelpText('Gmailアドレスをご入力ください。スプレッドシートをこのアドレスに共有します。')
    .setRequired(true);
  form.addTextItem().setTitle('担当者名').setRequired(true);
  form.addTextItem().setTitle('電話番号');
  form.addParagraphTextItem().setTitle('ご質問・メモ（任意）');

  // このスプレッドシートを回答先に設定
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // 既存トリガーを削除してから再登録
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onFormSubmit_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmit_').forSpreadsheet(ss).onFormSubmit().create();

  var formUrl = form.getPublishedUrl();
  var editUrl = form.getEditUrl();

  ui.alert(
    '申し込みフォームを作成しました！\n\n' +
    '■ 客先に渡すURL（申し込みリンク）:\n' + formUrl + '\n\n' +
    '■ フォーム編集URL（管理用）:\n' + editUrl + '\n\n' +
    '▶ このURLを送るだけで、申し込み→SS作成→メール送信が全自動になります。'
  );
  return formUrl;
}


// ================================================================
//  12-7: フォーム送信時の自動処理（onFormSubmit_）  【大B / 中12 / 小12-7】
//  申し込みフォームが送信されたら会社登録シートに追記し、SS作成＋メール送信を自動実行する。
//  createSignupForm() が設定したトリガーから自動実行される。
// ================================================================
function onFormSubmit_(e) {
  try {
    var responses   = e.namedValues;
    var companyName = String((responses['会社名'] || [''])[0]).trim();
    var adminEmail  = String((responses['担当者メールアドレス（Gmail）'] || [''])[0]).trim().toLowerCase();

    if (!companyName || !adminEmail || adminEmail.indexOf('@') === -1) return;

    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName('会社登録');
    if (!regSheet) return;

    // 重複チェック（同じ会社名が既にあればスキップ）
    var existing = regSheet.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (String(existing[i][0]).trim() === companyName) return;
    }

    // 会社登録シートに新規行を追加
    var newRow = regSheet.getLastRow() + 1;
    regSheet.getRange(newRow, 1).setValue(companyName);
    regSheet.getRange(newRow, 2).setValue(adminEmail);

    // フルセットアップ実行（SS作成 → フォルダ → メール）
    processNewCompany_(companyName, adminEmail);

  } catch(err) {
    // フォームトリガーのエラーはサイレント
  }
}


// ================================================================
//  12-8: キュー処理（processPendingCompanies_）  【大B / 中12 / 小12-8】
//  onEditCompanyRegister_ がキューに積んだ会社セットアップをまとめて処理する。
//  時間トリガー（60秒後）から自動実行される。getActiveSpreadsheet が使えないため
//  masterSsId を Script Properties から取得して openById で開く。
// ================================================================
function processPendingCompanies_() {
  var props = PropertiesService.getScriptProperties();
  var queue = JSON.parse(props.getProperty('companySetupQueue') || '[]');

  // キューを即座にクリア（二重処理防止）
  props.deleteProperty('companySetupQueue');

  if (queue.length === 0) {
    // トリガー削除だけして終了
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'processPendingCompanies_') ScriptApp.deleteTrigger(t);
    });
    return;
  }

  var masterSsId = props.getProperty('masterSsId');
  var ss = masterSsId ? SpreadsheetApp.openById(masterSsId) : null;
  var regSheet   = ss ? ss.getSheetByName('会社登録') : null;

  queue.forEach(function(item) {
    try {
      processNewCompany_(item.companyName, item.adminEmail);
    } catch(e) {
      // 失敗した場合はシートにエラーを記録
      if (regSheet) {
        var rows = regSheet.getDataRange().getValues();
        for (var i = 1; i < rows.length; i++) {
          if (String(rows[i][0]).trim() === item.companyName) {
            regSheet.getRange(i + 1, 3).setValue('エラー: ' + e.message).setBackground('#ffcdd2').clearNote();
            break;
          }
        }
      }
    }
  });

  // 使い終わったトリガーを削除
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processPendingCompanies_') ScriptApp.deleteTrigger(t);
  });
}


// ================================================================
//  13-1: インストール型トリガーのセットアップ（installTriggers）  【大C / 中13 / 小13-1】
//  最初に1回だけ実行する。これ以降は会社登録シートのA+B列入力で完全自動化される。
//  メニュー「🔧 初期設定（最初に1回だけ押す）」から実行。
//  ※ シンプルトリガー(onEdit)は認証付きサービス（Drive/Gmail等）が使えないため
//    インストール型トリガー(installedOnEdit_)を別途登録する必要がある。
// ================================================================
function installTriggers() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // マスターSS IDを保存（時間トリガーから openById で使う）
  PropertiesService.getScriptProperties().setProperty('masterSsId', ss.getId());

  // WebアプリURLを手入力で保存（自動取得は信頼性が低いため常にプロンプト表示）
  var props = PropertiesService.getScriptProperties();
  var currentUrl = props.getProperty('webAppUrl') || '';
  var resp = ui.prompt(
    'WebアプリURLを確認・入力してください',
    'deploy.bat 実行後に表示されるURLを貼り付けてください。\n' +
    '形式: https://script.google.com/macros/s/XXXXX/exec\n\n' +
    '現在の設定値: ' + (currentUrl || '（未設定）'),
    ui.ButtonSet.OK_CANCEL
  );
  var savedUrl = '';
  if (resp.getSelectedButton() === ui.Button.OK) {
    var inputUrl = resp.getResponseText().trim();
    if (inputUrl && inputUrl.indexOf('script.google.com/macros') !== -1) savedUrl = inputUrl;
  }
  if (!savedUrl) savedUrl = currentUrl; // キャンセル時は現在値を維持

  if (savedUrl) props.setProperty('webAppUrl', savedUrl);
  else props.deleteProperty('webAppUrl');

  // 既存の installedOnEdit_ トリガーを全て削除してから再登録
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'installedOnEdit_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('installedOnEdit_')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  var finalUrl = props.getProperty('webAppUrl') || '（未設定 — URLを貼り付けてから再実行してください）';
  ui.alert(
    '初期設定が完了しました！\n\n' +
    'WebアプリURL:\n' + finalUrl + '\n\n' +
    '「会社登録」シートのA列（会社名）とB列（Gmail）を入力するだけで\n' +
    'スプレッドシート作成・フォルダ作成・メール送信が自動で行われます。'
  );
}


// ================================================================
//  13-2: インストール型onEditトリガー（installedOnEdit_）  【大B / 中13 / 小13-2】
//  会社登録シートのA+B列入力 → processNewCompany_() を自動実行する。
//  インストール型トリガーはDrive/Gmail/ScriptApp等の認証付きサービスが使用可能。
//  実行時間制限もシンプルトリガーの30秒ではなく6分まで利用できる。
//  installTriggers() で登録済みの場合のみ発火する。
// ================================================================
function installedOnEdit_(e) {
  try {
    var range     = e.range;
    var sheet     = range.getSheet();
    var sheetName = sheet.getName();
    var row       = range.getRow();
    var col       = range.getColumn();

    // ── 自車専属マスタ or 設定シート編集時：再計算範囲ポップアップ ──────
    if (sheetName === '自車専属マスタ' || sheetName === '設定') {
      var mProps = PropertiesService.getScriptProperties();
      if (mProps.getProperty('recalcFromDateSet')) {
        SpreadsheetApp.getActiveSpreadsheet().toast('集計表再生成を押してください', '📋 マスタ変更済み', 5);
        return;
      }
      var mToday = new Date();
      var mTodayStr = Utilities.formatDate(mToday, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      var mFirst = new Date(mToday.getFullYear(), mToday.getMonth(), 1);
      var mFirstStr = Utilities.formatDate(mFirst, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      var mHtml = HtmlService.createHtmlOutput(
        '<style>' +
        'body{font-family:sans-serif;padding:14px;text-align:center;margin:0}' +
        'p{margin:0 0 12px;font-size:13px}' +
        'button{display:block;width:100%;padding:10px 8px;margin:6px 0;font-size:13px;' +
        'cursor:pointer;border:1px solid #ccc;border-radius:4px;background:#fff}' +
        'button:hover{background:#f0f0f0}.cancel{color:#999}' +
        '</style>' +
        '<p>集計表を再計算する範囲を選んでください。</p>' +
        '<button onclick="go(\'today\')">本日以降（' + mTodayStr + '〜）</button>' +
        '<button onclick="go(\'month\')">今月以降（' + mFirstStr + '〜）</button>' +
        '<button onclick="go(\'all\')">全期間（制限なし・従来通り）</button>' +
        '<button class="cancel" onclick="go(\'cancel\')">キャンセル</button>' +
        '<script>function go(v){' +
        'google.script.run.withSuccessHandler(function(){google.script.host.close();}).setRecalcChoice(v);' +
        '}<\/script>'
      ).setWidth(300).setHeight(235);
      SpreadsheetApp.getUi().showModalDialog(mHtml, 'マスタ／設定が変更されました');
      return;
    }

    if (sheetName !== '会社登録' || row <= 1) return;

    // ── A列 or B列: 会社名+Gmail が揃ったらフルセットアップ ──────────
    if (col === 1 || col === 2) {
      var companyName = String(sheet.getRange(row, 1).getValue() || '').trim();
      var adminEmail  = String(sheet.getRange(row, 2).getValue() || '').trim();
      var status      = String(sheet.getRange(row, 3).getValue() || '').trim();

      if (!companyName || !adminEmail || adminEmail.indexOf('@') === -1) return;
      // 未処理（空）またはエラーのみ実行。処理済ステータスは全てスキップ
      if (status !== '' && status.indexOf('エラー') !== 0) return;

      sheet.getRange(row, 3).setValue('処理中...').setBackground('#fff9c4');
      try {
        processNewCompany_(companyName, adminEmail);
      } catch(err) {
        sheet.getRange(row, 3).setValue('エラー: ' + err.message).setBackground('#ffcdd2');
      }
      return;
    }

    // ── F列 or G列: SS URL + App URL が揃ったら配布メール送信 ──────────
    if (col === 6 || col === 7) {
      var ssUrl      = String(sheet.getRange(row, 6).getValue() || '').trim();
      var appUrl     = String(sheet.getRange(row, 7).getValue() || '').trim();
      var mailStatus = String(sheet.getRange(row, 8).getValue() || '').trim();
      if (!ssUrl || !appUrl) return;
      if (mailStatus.indexOf('送信済') !== -1) return;
      var cName  = String(sheet.getRange(row, 1).getValue() || '').trim();
      var aEmail = String(sheet.getRange(row, 2).getValue() || '').trim();
      if (!cName || !aEmail || aEmail.indexOf('@') === -1) return;
      try {
        sendDistributionMail_(cName, aEmail, ssUrl, appUrl, row, sheet);
      } catch(err) {
        sheet.getRange(row, 8).setValue('エラー: ' + err.message).setBackground('#ffcdd2');
      }
    }
  } catch(ex) {}
}


// ================================================================
//  13-2b: マスタ変更時の再計算範囲保存（setRecalcChoice_）
//  showModalDialogのボタンから呼ばれる。choiceを受け取りScriptPropertiesに保存。
// ================================================================
function setRecalcChoice(choice) {
  if (choice === 'cancel') return;
  var mToday = new Date();
  var fromDate = '';
  if (choice === 'today') {
    fromDate = Utilities.formatDate(mToday, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  } else if (choice === 'month') {
    var mFirst = new Date(mToday.getFullYear(), mToday.getMonth(), 1);
    fromDate = Utilities.formatDate(mFirst, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  }
  // 'all' → fromDate = ''（全期間・制限なし）
  var props = PropertiesService.getScriptProperties();
  props.setProperty('recalcFromDate', fromDate);
  props.setProperty('recalcFromDateSet', '1');
  SpreadsheetApp.getActiveSpreadsheet().toast('集計表再生成を押してください', '📋 マスタ変更', 5);
}


// ================================================================
//  14-1: 修正用SSを作成（createDevSs）
//  今のSS（元SS）をコピーして「修正用_運行管理」を作成する。
//  作成後、修正用SSでメニュー「🆔 スクリプトID確認」を実行してIDをVSCODE担当者に伝える。
// ================================================================
function createDevSs() {
  var ui   = SpreadsheetApp.getUi();
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());

  var existing = DriveApp.getFilesByName('修正用_運行管理');
  if (existing.hasNext()) {
    var ans = ui.alert(
      '修正用SSが既に存在します',
      '「修正用_運行管理」が既にドライブにあります。\n新たに作り直しますか？',
      ui.ButtonSet.YES_NO
    );
    if (ans !== ui.Button.YES) return;
  }

  var copy    = file.makeCopy('修正用_運行管理');
  var copyUrl = 'https://docs.google.com/spreadsheets/d/' + copy.getId() + '/edit';

  PropertiesService.getScriptProperties().setProperty('originalSsId', ss.getId());

  ui.alert(
    '修正用SSを作成しました！\n\n' +
    '修正用SS URL:\n' + copyUrl + '\n\n' +
    '次の手順：\n' +
    '① 上のURLを開く\n' +
    '② 「メニュー」→「🆔 スクリプトID確認」を実行\n' +
    '③ 表示されたIDをVSCODE担当者に伝える\n' +
    '（以降このVSCODEが修正用SSと連動します）'
  );
}


// ================================================================
//  14-2: このSSのスクリプトIDを確認（showMyScriptId）
//  修正用SSを開いてこれを実行するとIDが表示される。
//  表示されたIDをVSCODE担当者（.clasp.json）に設定してもらう。
// ================================================================
function showMyScriptId() {
  var id = ScriptApp.getScriptId();
  SpreadsheetApp.getUi().alert(
    'このSSのスクリプトID:\n\n' +
    id + '\n\n' +
    'このIDをVSCODE担当者に伝えてください。\n' +
    '（.clasp.jsonに設定します）'
  );
}


// ================================================================
//  14-3: ②客用SS→全③各客SSにヘッダー・設定を反映（syncToAllClientSS）
//  ②客用SSのメニュー「📤 各客に反映」から実行。データ行は一切消さない。
//  ②の __TEMPLATE_SS__ シートのB1から①修正用SSのIDを取得し会社登録シートを参照する。
// ================================================================
function syncToAllClientSS() {
  var activeSs = SpreadsheetApp.getActiveSpreadsheet();

  // ①修正用SSから直接呼ばれた場合と②客用SSから呼ばれた場合の両方に対応する。
  // ②客用SSには __TEMPLATE_SS__ シートがあり、そこに①修正用SSのIDが記録されている。
  // ①修正用SSには __TEMPLATE_SS__ シートがないため、自分自身が修正用SSとして振る舞う。
  var tmplMarker    = activeSs.getSheetByName('__TEMPLATE_SS__');
  var masterSs, masterSsId, approvedVersion;

  if (tmplMarker) {
    // ② 客用SSから実行: __TEMPLATE_SS__ のB1から①修正用SSのIDを取得
    masterSsId = tmplMarker.getRange(1, 2).getValue();
    if (!masterSsId) {
      try { SpreadsheetApp.getUi().alert('①修正用SSのIDが未設定です。①で「テスト客SSに反映」を先に実行してください。'); } catch(e) {}
      return;
    }
    masterSs       = SpreadsheetApp.openById(masterSsId);
    approvedVersion = tmplMarker.getRange(1, 3).getValue() || null;
  } else {
    // ① 修正用SSから直接実行: 自分自身が修正用SS
    masterSs        = activeSs;
    masterSsId      = activeSs.getId();
    approvedVersion = PropertiesService.getScriptProperties().getProperty('approvedLibVersion') || null;
  }

  // approvedVersionが取れない場合のフォールバック（lastLibVersionNum → 設定なければ自動作成）
  if (!approvedVersion) {
    var props2 = PropertiesService.getScriptProperties();
    approvedVersion = props2.getProperty('lastLibVersionNum') || null;
    if (!approvedVersion) {
      // バージョンが一切なければここで新規作成して保存
      var newVer = createLibraryVersion_('syncToAllClientSS自動作成');
      if (newVer) {
        approvedVersion = String(newVer);
        props2.setProperty('approvedLibVersion', approvedVersion);
        props2.setProperty('lastLibVersionNum',  approvedVersion);
        if (tmplMarker) tmplMarker.getRange(1, 3).setValue(approvedVersion);
      }
    }
  }

  var regSheet = masterSs.getSheetByName('会社登録');
  if (!regSheet || regSheet.getLastRow() < 2) {
    try { SpreadsheetApp.getUi().alert('会社登録シートにデータがありません。'); } catch(e) {}
    return;
  }

  var businessSheets = ['運行', '集計表', '自車専属マスタ', '自車専属運行', 'マスタ', '設定'];
  var lastCol = regSheet.getLastColumn();
  var rows = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, Math.max(lastCol, 11)).getValues();
  var successCount = 0;
  var errorNames   = [];

  for (var i = 0; i < rows.length; i++) {
    var companyName  = String(rows[i][0]).trim();
    var ssUrl        = String(rows[i][5]).trim();  // F列: SS URL
    var clientScriptId = rows[i][10] ? String(rows[i][10]).trim() : ''; // K列: スクリプトID
    if (!companyName || !ssUrl) continue;

    var match = ssUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) continue;
    var clientSsId = match[1];

    try {
      var clientSs = SpreadsheetApp.openById(clientSsId);

      // ヘッダー行を①から反映（データ行は触らない）
      for (var si = 0; si < businessSheets.length; si++) {
        var sheetName = businessSheets[si];
        var srcSheet  = masterSs.getSheetByName(sheetName);
        if (!srcSheet || srcSheet.getLastColumn() === 0) continue;
        var tgtSheet = clientSs.getSheetByName(sheetName);
        if (!tgtSheet) tgtSheet = clientSs.insertSheet(sheetName);
        var srcCols = srcSheet.getLastColumn();
        tgtSheet.getRange(1, 1, 1, srcCols).setValues(srcSheet.getRange(1, 1, 1, srcCols).getValues());
        tgtSheet.getRange(1, 1, 1, srcCols).setBackgrounds(srcSheet.getRange(1, 1, 1, srcCols).getBackgrounds());
        tgtSheet.getRange(1, 1, 1, srcCols).setFontColors(srcSheet.getRange(1, 1, 1, srcCols).getFontColors());
        tgtSheet.getRange(1, 1, 1, srcCols).setFontWeights(srcSheet.getRange(1, 1, 1, srcCols).getFontWeights());
        tgtSheet.setFrozenRows(1);
      }
      ensureSettingItems_(clientSs);

      // 自車専属マスタ B列（運行状態）にドロップダウンを適用
      var mSheet = clientSs.getSheetByName('自車専属マスタ');
      if (mSheet && mSheet.getMaxRows() > 1) {
        var sv = SpreadsheetApp.newDataValidation()
          .requireValueInList(['運行','故障','待機'], true)
          .setAllowInvalid(false).build();
        mSheet.getRange(2, 2, mSheet.getMaxRows() - 1, 1).setDataValidation(sv);
      }

      // 不要シートを削除（マスタ点検項目など）
      // 情報・自社設定は新機能シートのため保持する
      var validClientSheets = [
        '運行','集計表','自車専属マスタ','自車専属運行','マスタ','設定',
        '情報','自社設定','__COMPANY_SS__'
      ];
      clientSs.getSheets().forEach(function(s) {
        if (validClientSheets.indexOf(s.getName()) === -1 && clientSs.getSheets().length > 1) {
          try { clientSs.deleteSheet(s); } catch(e) {}
        }
      });

      // K列にスクリプトIDがなければDrive APIで自動検索してK列に保存
      if (!clientScriptId) {
        clientScriptId = getNewSsScriptId_(clientSsId) || '';
        if (clientScriptId) regSheet.getRange(i + 2, 11).setValue(clientScriptId);
      }

      var stubOk = false;
      if (clientScriptId) {
        // スクリプトIDあり → スタブを最新化
        var stubResult = updateStubVersion_(clientScriptId, approvedVersion || '', false);
        stubOk = stubResult && stubResult.ok;
        if (!stubOk) {
          errorNames.push(companyName + '（API更新失敗: ' + (stubResult ? stubResult.error : '不明') + '）');
        }
      } else if (approvedVersion) {
        // スクリプトIDがどうしても取れない場合のみ新規デプロイ
        var deployResult = deployClientWebApp_(clientSsId, companyName, null, approvedVersion);
        if (deployResult && deployResult.scriptId) {
          regSheet.getRange(i + 2, 11).setValue(deployResult.scriptId);
          regSheet.getRange(i + 2, 7).setValue(deployResult.webAppUrl);
          stubOk = true;
        } else {
          var dErr = (deployResult && deployResult.error) ? deployResult.error : '詳細不明';
          errorNames.push(companyName + '（スクリプトID取得失敗・API: ' + dErr + '）');
        }
      } else {
        // スクリプトIDもapprovedVersionも取れない
        errorNames.push(companyName + '（スクリプトID未登録・バージョン未設定。①で「テスト客SSに反映」を先に実行してください）');
      }

      if (stubOk) successCount++;
    } catch(e) {
      errorNames.push(companyName + '（例外: ' + (e.message || String(e)) + '）');
    }
  }

  var msg = successCount + '社への反映が完了しました。';
  if (approvedVersion) msg += '\nコードバージョン: ' + approvedVersion;
  if (errorNames.length > 0) msg += '\n失敗: ' + errorNames.join(', ');
  msg += '\n\n各客SSでF5を押すとメニューが更新されます。';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}

function checkScopeAuth() {
  var token = ScriptApp.getOAuthToken();
  Logger.log('スコープ承認OK: ' + token.substring(0, 20) + '...');
}

// ================================================================
//  診断：各客SSのスクリプトID取得失敗原因を調べる（diagClientApi_）
//  スクリプトエディタから直接実行 → アラートで結果表示
// ================================================================
function diagClientApi() {
  var masterSs = SpreadsheetApp.getActiveSpreadsheet();
  var regSheet = masterSs.getSheetByName('会社登録');
  if (!regSheet || regSheet.getLastRow() < 2) {
    Logger.log('会社登録シートにデータがありません');
    return;
  }
  var token = ScriptApp.getOAuthToken();
  var rows = regSheet.getRange(2, 1, Math.min(3, regSheet.getLastRow() - 1), 11).getValues();

  for (var i = 0; i < rows.length; i++) {
    var companyName = String(rows[i][0]).trim();
    var ssUrl = String(rows[i][5]).trim();
    if (!companyName || !ssUrl) continue;
    var match = ssUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) { Logger.log(companyName + ': URLパース失敗'); continue; }
    var ssId = match[1];

    // Drive API テスト
    var q = encodeURIComponent('mimeType="application/vnd.google-apps.script" and "' + ssId + '" in parents');
    var drResp = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files?q=' + q + '&fields=files(id,name)&supportsAllDrives=true&includeItemsFromAllDrives=true',
      { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
    );
    Logger.log('[' + companyName + '] Drive:' + drResp.getResponseCode() + ' ' + drResp.getContentText().slice(0, 200));

    // Script API テスト（自分のプロジェクトが読めるか確認）
    if (i === 0) {
      var selfResp = UrlFetchApp.fetch(
        'https://script.googleapis.com/v1/projects/' + ScriptApp.getScriptId(),
        { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
      );
      Logger.log('[ScriptAPI-self] ' + selfResp.getResponseCode() + ' ' + selfResp.getContentText().slice(0, 200));
    }
  }
  Logger.log('diagClientApi 完了');
}


// ================================================================
// ■ グループ13：CSV・Excelデータ一括読込
// ================================================================
//
//   13-1  : showCsvImportDialogUnkou()
//             運行シート用の一括読込ダイアログを開く（メニュー項目）
//   13-1b : showCsvImportDialogMaster()
//             自車専属マスタ用の一括読込ダイアログを開く
//   13-1c : showCsvImportDialogCust()
//             マスタ（取引先）用の一括読込ダイアログを開く
//   13-2  : showCsvImportDialog_(sheetType)
//             種別を受け取りcsvImport.htmlをモーダルで表示する共通処理
//             現在のSSのIDをテンプレートに埋め込み誤SS登録を防止する
//   13-3  : getImportDictionary(sheetType, companySsId)
//             設定シートH列の辞書データを読み込んで返す（HTML→GAS API）
//             辞書未作成なら initImportDictionary_ でデフォルトを自動生成する
//   13-4  : importBulkRows(sheetType, mappedRows, companySsId)
//             マッピング済み行データを対象シートに一括登録する（HTML→GAS API）
//             IDはすべてシステム採番（V-/S-/M-）、LockServiceで排他制御
//   13-5  : buildSheetRow_(sheetType, id, fieldMap, ss)
//             fieldIdマップから各シートの列構成に合わせた配列を返す内部補助
//             運行シートは車番/乗務員名で自車専属マスタを引いて区分等を自動補完
//   13-6  : initImportDictionary_(ss)
//             設定シートH列に辞書ヘッダーとデフォルトエントリを自動生成する
//             H1が既に「【辞書】種別」なら即返却（2重初期化防止）
//   13-7  : parseImportDate_(v)
//             SheetJSシリアル値・ISO文字列・Dateを受け取りDate型に変換する補助
//   13-8  : toImportNum_(v)
//             インポート値を数値変換。空・変換不可は空文字を返す補助
//
// ================================================================


// ================================================================
//  13-1: CSV/Excel一括読込ダイアログ表示（運行）  【大C / 中13 / 小13-1】
// ================================================================
function showCsvImportDialogUnkou() { showCsvImportDialog_('unkou'); }


// ================================================================
//  13-1b: CSV/Excel一括読込ダイアログ表示（自車専属マスタ）  【大C / 中13 / 小13-1b】
// ================================================================
function showCsvImportDialogMaster() { showCsvImportDialog_('master'); }


// ================================================================
//  13-1c: CSV/Excel一括読込ダイアログ表示（マスタ取引先）  【大C / 中13 / 小13-1c】
// ================================================================
function showCsvImportDialogCust() { showCsvImportDialog_('cust'); }


// ================================================================
//  13-2: インポートダイアログ共通表示（showCsvImportDialog_）  【大C / 中13 / 小13-2】
//  sheetTypeを受け取りcsvImport.htmlをモーダルダイアログとして表示する
//  現在のSSのIDをテンプレートに渡してサーバー側が正しいSSを開けるようにする
// ================================================================
function showCsvImportDialog_(sheetType) {
  var titles = { unkou: '運行シート', master: '自車専属マスタ', cust: 'マスタ（取引先）' };
  var title  = '📥 データ読み込み ─ ' + (titles[sheetType] || sheetType);
  var tmpl   = HtmlService.createTemplateFromFile('csvImport');
  tmpl.sheetType    = sheetType;
  tmpl.currentSsId  = SpreadsheetApp.getActiveSpreadsheet().getId();
  var html = tmpl.evaluate().setWidth(880).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, title);
}


// ================================================================
//  13-3: 辞書データ取得（getImportDictionary）  【大A / 中13 / 小13-3】
//  設定シートH〜K列に保存された辞書を返す（HTML側 google.script.run から呼ぶ）
//  辞書が未作成なら initImportDictionary_ でデフォルトエントリを自動生成する
// ================================================================
function getImportDictionary(sheetType, companySsId) {
  var ss = getTargetSS_(companySsId);
  initImportDictionary_(ss);
  var setting = ss.getSheetByName('設定');
  if (!setting) return [];
  var lr = setting.getLastRow();
  var lc = setting.getLastColumn();
  if (lr < 2 || lc < 8) return [];
  var numCols = Math.max(lc - 7, 4);
  var rows    = setting.getRange(2, 8, lr - 1, numCols).getValues();
  var result  = [];
  for (var i = 0; i < rows.length; i++) {
    var type     = String(rows[i][0] || '').trim();
    if (type !== sheetType) continue;
    var fieldId  = String(rows[i][1] || '').trim();
    var dispName = String(rows[i][2] || '').trim();
    var aliasStr = String(rows[i][3] || '').trim();
    if (!fieldId) continue;
    var aliases = aliasStr
      ? aliasStr.split(',').map(function(a) { return a.trim(); }).filter(Boolean)
      : [];
    result.push({ fieldId: fieldId, dispName: dispName, aliases: aliases });
  }
  return result;
}


// ================================================================
//  13-4: データ一括登録（importBulkRows）  【大A / 中13 / 小13-4】
//  HTMLから受け取ったマッピング済み行データを対象シートに一括登録する
//  IDはすべてシステム採番（V-/S-/M-）、LockServiceで排他制御して重複防止
//  重複判定なし：運送業の2回戦（同人同所）を正しく扱うためすべて新規追加
// ================================================================
function importBulkRows(sheetType, mappedRows, companySsId) {
  if (!mappedRows || mappedRows.length === 0) return { ok: 0 };
  var ss         = getTargetSS_(companySsId);
  var sheetNames = { unkou: '運行', master: '自車専属マスタ', cust: 'マスタ' };
  var prefixes   = { unkou: 'V',   master: 'S',              cust: 'M'    };
  var sheetName  = sheetNames[sheetType];
  var prefix     = prefixes[sheetType];
  if (!sheetName) throw new Error('不明なシート種別: ' + sheetType);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(sheetName + 'シートが見つかりません');

  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(e) {
    throw new Error('ロック取得タイムアウト。再度お試しください。');
  }
  try {
    var nextNum   = getNextIdNum_(sheet, prefix);
    var writeRows = [];
    for (var i = 0; i < mappedRows.length; i++) {
      var id = prefix + '-' + ('0000' + (nextNum + i)).slice(-4);
      writeRows.push(buildSheetRow_(sheetType, id, mappedRows[i], ss));
    }
    var startRow = Math.max(sheet.getLastRow() + 1, 2);
    var numCols  = writeRows[0].length;
    sheet.getRange(startRow, 1, writeRows.length, numCols).setValues(writeRows);

    if (sheetType === 'unkou') {
      sheet.getRange(startRow, 10, writeRows.length, 1).setNumberFormat('yyyy/MM/dd');
      applyMoneyFormat_(sheet, startRow, writeRows.length, 'unkou');
      var formulas = [];
      for (var r = 0; r < writeRows.length; r++) {
        var rn = startRow + r;
        formulas.push(['=IF(AND(U' + rn + '="",T' + rn + '=""),"",U' + rn + '-T' + rn + ')']);
      }
      sheet.getRange(startRow, 22, writeRows.length, 1).setFormulas(formulas);
      // 集計表を即時同期（インポートした全IDを更新）
      for (var i = 0; i < writeRows.length; i++) {
        delaySyncSummary_(writeRows[i][0], ss);
      }
      // 支払いが含まれる場合は集計表のAA列(col27)に直接書き込む
      var paymentMap = {};
      for (var i = 0; i < mappedRows.length; i++) {
        var pay = toImportNum_(mappedRows[i]['payment']);
        if (pay !== '' && pay > 0) paymentMap[writeRows[i][0]] = pay;
      }
      if (Object.keys(paymentMap).length > 0) {
        var sumSheet = ss.getSheetByName('集計表');
        if (sumSheet && sumSheet.getLastRow() >= 2) {
          var sumIds = sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, 1).getValues();
          for (var si = 0; si < sumIds.length; si++) {
            var sid = String(sumIds[si][0] || '').trim();
            if (paymentMap[sid] !== undefined) {
              sumSheet.getRange(si + 2, 27).setValue(paymentMap[sid]); // AA列=支払い
            }
          }
        }
      }
      // 日付ソート・利益計算（手動入力と同じ状態にする）
      sortUnkouByDate_(companySsId);
      sortSummaryByDate_(companySsId);
      calculatePaymentAmount(companySsId);
    }
    return { ok: writeRows.length };
  } finally {
    lock.releaseLock();
  }
}


// ================================================================
//  13-5: シート行データ構築（buildSheetRow_）  【大B / 中13 / 小13-5】
//  fieldIdとその値のマップを受け取り各シートの列順に合わせた配列を返す
//  運行シートは車番または乗務員名で自車専属マスタを検索し区分・会社名等を補完する
// ================================================================
function buildSheetRow_(sheetType, id, fieldMap, ss) {

  if (sheetType === 'unkou') {
    var car    = String(fieldMap['car']    || '').trim();
    var driver = String(fieldMap['driver'] || '').trim();
    var division = '', company = '', ton = '', carType = '', phone = '';
    if (car || driver) {
      var master = ss.getSheetByName('自車専属マスタ');
      if (master && master.getLastRow() >= 2) {
        var mData = master.getRange(2, 1, master.getLastRow() - 1, 11).getValues();
        for (var m = 0; m < mData.length; m++) {
          var mCar    = String(mData[m][7]  || '').trim(); // H列(8): 車番
          var mDriver = String(mData[m][8]  || '').trim(); // I列(9): 乗務員名
          if ((car && mCar === car) || (driver && mDriver === driver)) {
            division = String(mData[m][2] || '');          // C列(3): 区分
            company  = String(mData[m][3] || '');          // D列(4): 会社名
            ton      = String(mData[m][5] || '');          // F列(6): トン数
            carType  = String(mData[m][6] || '');          // G列(7): 車種
            phone    = String(mData[m][9]  || '');         // J列(10): 携帯番号
            if (!driver) driver = mDriver;
            if (!car)    car    = mCar;
            break;
          }
        }
      }
    }
    var dateVal = parseImportDate_(fieldMap['date']);
    // 運行シート 28列（unkouHeader参照）
    return [
      id,                                                  //  1: ID
      division || String(fieldMap['division'] || ''),      //  2: 区分
      company  || String(fieldMap['company']  || ''),      //  3: 会社名
      ton      || String(fieldMap['ton']      || ''),      //  4: トン数
      carType  || String(fieldMap['carType']  || ''),      //  5: 車種
      car,                                                 //  6: 車番
      driver,                                              //  7: 乗務員名
      phone    || String(fieldMap['phone']    || ''),      //  8: 携帯番号
      String(fieldMap['signboard'] || ''),                 //  9: 看板名
      dateVal || '',                                       // 10: 日付
      String(fieldMap['client']    || ''),                 // 11: 荷主
      String(fieldMap['pickPlace'] || ''),                 // 12: 積地
      String(fieldMap['dropPlace'] || ''),                 // 13: 降地
      '', '', '', '', '',                                  // 14-18: 時刻（空）
      toImportNum_(fieldMap['sales']),                     // 19: 売上
      toImportNum_(fieldMap['tollReq']),                   // 20: 請求高速
      toImportNum_(fieldMap['tollReal']),                  // 21: 実費高速
      '',                                                  // 22: 合計高速（数式で後セット）
      String(fieldMap['memo'] || ''),                      // 23: 備考
      '', '', '', '', ''                                   // 24-28: 管理データ等（空）
    ];
  }

  if (sheetType === 'master') {
    // 自車専属マスタ 32列（masterHeader参照）
    return [
      id,                                                  //  1: 車両ID
      '待機',                                              //  2: 運行状態（デフォルト）
      String(fieldMap['division']  || ''),                 //  3: 区分
      String(fieldMap['company']   || ''),                 //  4: 会社名
      String(fieldMap['signboard'] || ''),                 //  5: 看板名
      String(fieldMap['ton']       || ''),                 //  6: トン数
      String(fieldMap['carType']   || ''),                 //  7: 車種
      String(fieldMap['car']       || ''),                 //  8: 車番
      String(fieldMap['driver']    || ''),                 //  9: 乗務員名
      String(fieldMap['phone']     || ''),                 // 10: 携帯番号
      String(fieldMap['email']     || ''),                 // 11: アドレス
      String(fieldMap['fuel']      || ''),                 // 12: 燃費
      String(fieldMap['memo']      || ''),                 // 13: 備考
      '', '', '',                                          // 14-16: 仮日数・給料・%
      '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '' // 17-32: 各種費用
    ];
  }

  if (sheetType === 'cust') {
    // マスタ（取引先）17列（custHeader参照）
    return [
      id,                                                  //  1: マスタID
      String(fieldMap['company'] || ''),                   //  2: 会社名
      String(fieldMap['tel']     || ''),                   //  3: 電話
      String(fieldMap['fax']     || ''),                   //  4: FAX
      String(fieldMap['zip']     || ''),                   //  5: 郵便番号
      String(fieldMap['address'] || ''),                   //  6: 住所
      String(fieldMap['rep']     || ''),                   //  7: 代表者
      String(fieldMap['contact'] || ''),                   //  8: 配車担当
      '', '', '', '', '',                                  //  9-13: 銀行等
      String(fieldMap['memo']    || ''),                   // 14: 備考
      '', '', ''                                           // 15-17: インボイス等
    ];
  }
  return [id];
}


// ================================================================
//  13-6: 辞書初期化（initImportDictionary_）  【大B / 中13 / 小13-6】
// ================================================================
function initImportDictionary_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var setting = ss.getSheetByName('設定');
  if (!setting) return;

  // ★修正：列数が足りない場合は、先に列を追加してエラーを防ぐ
  var neededCols = 11;
  if (setting.getMaxColumns() < neededCols) {
    setting.insertColumnsAfter(setting.getMaxColumns(), neededCols - setting.getMaxColumns());
  }

  // v3: 支払い追加のため再初期化
  if (String(setting.getRange(1, 8).getValue()).trim() === '【辞書v3】種別') return;

  var defaults = [
    ['unkou', 'division',  '区分',          '区分,自車区分,車両区分'],
    ['unkou', 'company',   '会社名',        '会社名,所属,協力会社,事業者名,法人名,会社'],
    ['unkou', 'ton',       'トン数',        'トン数,t数,積載,クラス,積載量'],
    ['unkou', 'carType',   '車種',          '車種,形状,ボディ,車体形状'],
    ['unkou', 'car',       '車番',          '車番,号車,車両番号,車両,ナンバー,登録番号'],
    ['unkou', 'driver',    '乗務員名',      '乗務員,ドライバー,乗務員名,運転手,氏名,担当者'],
    ['unkou', 'phone',     '携帯番号',      '携帯番号,連絡先,電話番号,スマホ,TEL,携帯'],
    ['unkou', 'signboard', '看板名',        '看板名,看板,取引先名,ブランド名,荷主略称'],
    ['unkou', 'date',      '日付',          '日付,月日,運行日,配車日,作業日,納期,輸送日'],
    ['unkou', 'client',    '荷主名',        '荷主,荷主名,依頼主,発注元,委託元'],
    ['unkou', 'pickPlace', '積地',          '積地,発地,積込先,出発地,引取場所,メーカー,集荷先,仕入先'],
    ['unkou', 'dropPlace', '降地',          '降地,着地,納品先,到着地,配送先,現場,納入先,配達先'],
    ['unkou', 'sales',     '売上',          '売上,運賃,金額,請求額,単価,受注金額,請求金額,支払運賃'],
    ['unkou', 'tollReq',   '請求高速',      '請求高速,請求高速代,高速請求,高速代（請求）,高速請求代,請求(高速)'],
    ['unkou', 'tollReal',  '実費高速',      '実費高速,実費高速代,高速実費,高速代（実費）,高速実費代,実費(高速)'],
    ['unkou', 'payment',   '支払い',        '支払い,支払,支払額,給与,手取り,運転手支払'],
    ['unkou', 'memo',      '備考',          '備考,メモ,特記,注意事項'],
    ['master', 'company',  '会社名',        '会社名,所属,協力会社,事業者名,法人名'],
    ['master', 'ton',      'トン数',        'トン数,t数,積載,クラス,積載量'],
    ['master', 'carType',  '車種',          '車種,形状,ボディ,車体形状'],
    ['master', 'car',      '車番',          '車番,登録番号,ナンバー,車両番号,プレート'],
    ['master', 'driver',   '乗務員名',      '乗務員名,氏名,ドライバー,運転手,担当者'],
    ['master', 'phone',    '携帯番号',      '携帯番号,連絡先,電話番号,スマホ,TEL'],
    ['master', 'email',    'メールアドレス','メール,メールアドレス,mail,email,アドレス'],
    ['master', 'division', '区分',          '区分,自車区分,所属区分'],
    ['master', 'fuel',     '燃費',          '燃費,燃料消費率'],
    ['master', 'memo',     '備考',          '備考,メモ'],
    ['cust', 'company',    '会社名',        '会社名,荷主名,取引先,顧客名,得意先,荷主,客先'],
    ['cust', 'tel',        '電話',          '電話,TEL,連絡先,電話番号,代表電話'],
    ['cust', 'fax',        'FAX',           'FAX,ファックス,FAX番号'],
    ['cust', 'zip',        '郵便番号',      '郵便番号,〒,zip'],
    ['cust', 'address',    '住所',          '住所,所在地,住所1,本社住所,所在'],
    ['cust', 'rep',        '代表者',        '代表者,代表,社長,責任者'],
    ['cust', 'contact',    '配車担当',      '配車担当,担当者,担当,配車係'],
    ['cust', 'memo',       '備考',          '備考,メモ,特記']
  ];

  var neededRows = defaults.length + 1;
  if (setting.getMaxRows() < neededRows) {
    setting.insertRowsAfter(setting.getMaxRows(), neededRows - setting.getMaxRows());
  }
  setting.getRange(1, 8, 1, 4).setValues([['【辞書v3】種別', 'フィールドID', '表示名', 'エイリアス（カンマ区切り・自由に追加可）']]);
  setting.getRange(1, 8, 1, 4).setFontWeight('bold').setBackground('#e3f2fd');
  setting.getRange(2, 8, defaults.length, 4).setValues(defaults);
}


// ================================================================
//  13-7: インポート用日付変換（parseImportDate_）  【大B / 中13 / 小13-7】
//  SheetJSシリアル値・ISO文字列・DateオブジェクトをDate型に変換する
//  変換不可な値はnullを返す
// ================================================================
function parseImportDate_(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  var n = Number(v);
  if (!isNaN(n) && n > 40000) {
    return new Date(Math.round((n - 25569) * 86400 * 1000));
  }
  var d = new Date(String(v));
  return isNaN(d.getTime()) ? null : d;
}


// ================================================================
//  13-8: インポート数値変換補助（toImportNum_）  【大B / 中13 / 小13-8】
//  数値に変換できれば数値を、空・変換不可なら空文字を返す
// ================================================================
function toImportNum_(v) {
  if (v === null || v === undefined || v === '') return '';
  var n = Number(v);
  return isNaN(n) ? '' : n;
}


// ================================================================
//  13-0: 空インポート行一括削除（deleteBlankImportRows）  【大C / 中13 / 小13-0】
//  運行シートにIDだけあって日付・車番・売上が全部空の行を削除する
//  テスト失敗時の一括クリアに使う（メニューから実行）
// ================================================================
function deleteBlankImportRows() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet || sheet.getLastRow() < 2) { ss.toast('削除対象なし', '✅', 3); return; }
  var lr   = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lr - 1, 22).getValues();
  var dels = [];
  for (var i = data.length - 1; i >= 0; i--) {
    var id    = String(data[i][0]  || '').trim();
    var car   = String(data[i][5]  || '').trim();   // F列: 車番
    var date  = String(data[i][9]  || '').trim();   // J列: 日付
    var pick  = String(data[i][11] || '').trim();   // L列: 積地
    var sales = String(data[i][18] || '').trim();   // S列: 売上
    if (/^[VSM]-\d+$/.test(id) && !car && !date && !pick && !sales) {
      dels.push(i + 2);
    }
  }
  if (dels.length === 0) { ss.toast('削除対象の空行なし', '✅', 3); return; }
  dels.forEach(function(r) { sheet.deleteRow(r); });
  ss.toast(dels.length + '行の空インポート行を削除しました', '🗑', 4);
}


// ================================================================
//  13-9: 辞書エイリアス自動保存（saveImportAliases）  【大A / 中13 / 小13-9】
//  手動マッピングで使われたExcel列名を設定シートの辞書に自動追記する
//  次回インポート時に同じExcelを使うと自動マッピングされるようになる
//  既に同じ内容（大文字小文字・空白無視）が登録済みの場合はスキップする
// ================================================================
function saveImportAliases(sheetType, newMappings, companySsId) {
  if (!newMappings || newMappings.length === 0) return;
  var ss      = getTargetSS_(companySsId);
  var setting = ss.getSheetByName('設定');
  if (!setting || setting.getLastRow() < 2) return;
  var lr   = setting.getLastRow();
  var rows = setting.getRange(2, 8, lr - 1, 4).getValues();
  for (var m = 0; m < newMappings.length; m++) {
    var nm       = newMappings[m];
    var newAlias = String(nm.alias || '').trim();
    if (!newAlias) continue;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim() !== sheetType)  continue;
      if (String(rows[i][1] || '').trim() !== nm.fieldId) continue;
      var aliasStr = String(rows[i][3] || '').trim();
      var aliases  = aliasStr ? aliasStr.split(',').map(function(a) { return a.trim(); }) : [];
      var already  = aliases.some(function(a) {
        return a.toLowerCase().replace(/\s/g, '') === newAlias.toLowerCase().replace(/\s/g, '');
      });
      if (!already) {
        aliases.push(newAlias);
        setting.getRange(i + 2, 11, 1, 1).setValue(aliases.join(','));
      }
      break;
    }
  }
}


// ================================================================
//  14-1: 自社設定シート自動生成（ensureCompanySettingSheet_）
// ================================================================
// ================================================================
//  14-4: 帳票発行済マーク（markDocumentIssued）
// ================================================================
function markDocumentIssued(rowId, docType) {
  if (!rowId) return;
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastCol = sheet.getLastColumn();
  var hdrs    = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h||'').trim(); });
  var colName = (docType === 'hatchu') ? '発注書・指示書' : '車番連絡';
  var colIdx  = hdrs.indexOf(colName); // 0-based
  if (colIdx === -1) return;
  var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]||'').trim() === rowId) {
      sheet.getRange(i + 2, colIdx + 1).setValue('済');
      SpreadsheetApp.getActiveSpreadsheet().toast(
        colName + ' を発行済にしました（ID: ' + rowId + '）', '✅', 3
      );
      return;
    }
  }
}


// ================================================================
//  14-5: 帳票メール／FAX送信（sendDocumentEmail）
// ================================================================
function sendDocumentEmail(docData, docType, method) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var custSheet = ss.getSheetByName('マスタ');
  if (!custSheet || custSheet.getLastRow() < 2) return { ok: false, msg: 'マスタシートが見つかりません。' };

  var cLastCol = custSheet.getLastColumn();
  var cHdrs = custSheet.getRange(1,1,1,cLastCol).getValues()[0].map(function(h){return String(h||'').trim();});
  var cData = custSheet.getRange(2,1,custSheet.getLastRow()-1,cLastCol).getValues();
  var nIdx     = cHdrs.indexOf('会社名');
  var emailIdx = cHdrs.indexOf('メールアドレス');
  var faxIdx   = cHdrs.indexOf('FAX');

  var clientName = String(docData.client || '').trim();
  if (!clientName) return { ok: false, msg: '荷主名が取得できません。' };

  var custRow = null;
  for (var i = 0; i < cData.length; i++) {
    if (String(cData[i][nIdx]||'').trim() === clientName) { custRow = cData[i]; break; }
  }
  if (!custRow) return { ok: false, msg: '「' + clientName + '」がマスタに見つかりません。' };

  if (method === 'email') {
    var email = (emailIdx >= 0) ? String(custRow[emailIdx]||'').trim() : '';
    if (!email) return { ok: false, msg: 'メールアドレスが未登録です。\nマスタシートのQ列に登録してください。' };
    var docName = (docType === 'hatchu') ? '発注書・指示書' : '車番連絡';
    var subject = docName + '【' + (docData.dateStr||'') + '】' + (docData.selfName||'');
    var body = clientName + ' 御中\n\n下記の通りご連絡申し上げます。\n\n'
      + '運行日：' + (docData.dateStr||'') + '（' + (docData.wdStr||'') + '）\n'
      + '積　地：' + (docData.pickPlace||'') + '\n'
      + '降　地：' + (docData.dropPlace||'') + '\n'
      + '車　番：' + (docData.car||'') + '\n'
      + '乗務員：' + (docData.driver||'') + '\n\n'
      + (docData.selfName||'');
    GmailApp.sendEmail(email, subject, body);
    markDocumentIssued(docData.id, docType);
    return { ok: true, msg: email + ' に送信しました。' };

  } else if (method === 'fax') {
    var fax = (faxIdx >= 0) ? String(custRow[faxIdx]||'').trim() : '';
    if (!fax) return { ok: false, msg: 'FAX番号が未登録です。\nマスタシートのD列に登録してください。' };
    Logger.log('[FAX送信予約] 宛先:' + fax + ' 書類:' + docType + ' ID:' + docData.id);
    markDocumentIssued(docData.id, docType);
    return { ok: true, msg: 'FAX送信を予約しました。（宛先：' + fax + '）' };
  }

  return { ok: false, msg: '不明な送信方法です。' };
}


// ================================================================
//  15-1: 配車確定・合体処理（matchAndConfirmDispatch）  【大C / 中15 / 小15-1】
//
//  「情報」シートでA列（チェックボックス）にチェックを入れた行を
//  メニューから「🔗 チェックした行を配車確定」で実行する。
//
//  【処理パターン】
//  ・1行チェック: 貨物単独または車両単独として運行シートに登録
//  ・2行チェック: 貨物行と車両行をマッチング→1つの運行IDで運行シートに登録
//  ・3行以上: エラーを出して中断（同時処理は2行まで）
//
//  【車種不一致チェック】
//  屋根ありグループ（W・箱・幌等）と屋根なしグループ（平・ユニック等）が
//  混在する場合、担当者に確認ダイアログを出して正しい車種を入力させる。
//
//  【確定後の処理】
//  ・対象行のチェックを解除
//  ・進捗列を「確定」に変更
//  ・行を黄色に着色（onEditが発火しなくても確実に着色）
//  ・集計表を同期（delaySyncSummary_）
//
//  M&A向け補足: この関数が「情報シート→運行シート」への自動転記の核心。
//              従来は担当者が手動で運行シートに入力していたが、
//              この機能により入力工数を大幅削減・入力ミスをゼロにできる。
// ================================================================
function matchAndConfirmDispatch() {
  var ui   = SpreadsheetApp.getUi();
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var joho = ss.getSheetByName('情報');
  if (!joho || joho.getLastRow() < 2) {
    ui.alert('「情報」シートにデータがありません。\nメニュー→シート再生成で情報シートを作成してください。');
    return;
  }

  // ── A列(貨物チェック)とN列(車両チェック)を独立して収集 ──────────────
  // 列構成（0ベース）: A=0:チェック(貨物) / N=13:チェック(車両)
  var lastRow = joho.getLastRow();
  var allData = joho.getRange(2, 1, lastRow - 1, 26).getValues();

  var cargoRows = []; // A列(index 0)が true の行 → 貨物として使用
  var vehRows   = []; // N列(index 13)が true の行 → 車両として使用
  for (var i = 0; i < allData.length; i++) {
    var rn = i + 2;
    if (allData[i][0]  === true) cargoRows.push({ rowNum: rn, data: allData[i] });
    if (allData[i][13] === true) vehRows.push(  { rowNum: rn, data: allData[i] });
  }

  // ── 選択バリデーション ────────────────────────────────────────────
  // 対応パターン:
  //   貨物1・車両0 → 貨物単独受託
  //   貨物0・車両1 → 車両単独配車
  //   貨物1・車両1 → 1対1マッチング
  //   貨物2・車両1 → 2行程（1人の乗務員が2つの荷物を運ぶ）← 同一IDで2行追加
  // 「車両2以上」「貨物3以上」は非対応
  if (cargoRows.length === 0 && vehRows.length === 0) {
    ui.alert('行を選択してください。\nA列（貨物チェック）またはN列（車両チェック）にチェックを入れてから実行してください。');
    return;
  }
  if (cargoRows.length > 2) {
    ui.alert('貨物チェック（A列）は2行までにしてください。\n3行程以上は個別に実行してください。'); return;
  }
  if (vehRows.length > 1) {
    ui.alert('車両チェック（N列）は1行だけにしてください。'); return;
  }
  if (cargoRows.length === 2 && vehRows.length === 0) {
    ui.alert('貨物が2行選択されています。\n対応する車両行のN列もチェックしてから実行してください。'); return;
  }

  // ── 運行シートの準備 ────────────────────────────────────────────
  var unkou    = ss.getSheetByName('運行');
  if (!unkou) { ui.alert('運行シートが見つかりません。'); return; }
  var uLastCol = unkou.getLastColumn();
  var uHdrs    = unkou.getRange(1, 1, 1, uLastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  function uIdx(name) { return uHdrs.indexOf(name); }

  var newRow = [];
  for (var n = 0; n < uLastCol; n++) newRow.push('');

  // ── パターン分岐 ────────────────────────────────────────────────
  // ① 同一行に両方チェック → 1行で貨物+車両両方入力されている場合の単独確定
  // ② 貨物のみ → 貨物情報だけで運行登録
  // ③ 車両のみ → 車両情報だけで運行登録
  // ④ 別行でそれぞれチェック → 異なる行をマッチングして合体

  var vehData    = vehRows.length > 0 ? vehRows[0].data : null;
  var groupA     = ['W', '箱', '幌', 'アコーディオン', 'Wトレ']; // 屋根あり
  var groupB     = ['平', 'ユニック', '平トレ'];                    // 屋根なし

  // ── 車種不一致チェック（貨物側と車両側の屋根グループが違う場合に確認） ──
  // 2行程の場合は貨物1件目でチェックし、決定した車種を両行程に適用する
  function checkCarType(cData, vData) {
    if (!cData || !vData) return vData ? String(vData[19]||'').trim() : (cData ? String(cData[8]||'').trim() : null);
    var cType = String(cData[8]  || '').trim(); // I列=車種(貨物要求)
    var vType = String(vData[20] || '').trim(); // U列(index20)=車種(車両) ※26列構成
    var cGrp  = groupA.indexOf(cType) >= 0 ? 'A' : (groupB.indexOf(cType) >= 0 ? 'B' : '');
    var vGrp  = groupA.indexOf(vType) >= 0 ? 'A' : (groupB.indexOf(vType) >= 0 ? 'B' : '');
    var resolved = vType || cType;
    if (cGrp && vGrp && cGrp !== vGrp) {
      var res = ui.prompt(
        '【警告】車種不一致',
        '屋根の有無が異なります。\n貨物要求: ' + cType + ' / 実車両: ' + vType + '\n\n登録する車種を入力してOKを押してください。',
        ui.ButtonSet.OK_CANCEL
      );
      if (res.getSelectedButton() !== ui.Button.OK) return null; // null=中断シグナル
      resolved = res.getResponseText().trim() || resolved;
    }
    return resolved;
  }

  // ── 運行シートへの追記（内部処理） ──────────────────────────────────
  // presetId を渡すと採番せずそのIDで登録（2行程で同一IDを共有するため）
  var registeredIds = [];
  function addUnkouRow(cData, vData, finalType, presetId) {
    var row = [];
    for (var n = 0; n < uLastCol; n++) row.push('');
    buildJohoNewRow_(row, uIdx, cData, vData, finalType);
    var nid = presetId || ('V-' + String(getNextIdNum_(unkou, 'V-')).padStart(4, '0'));
    if (uIdx('ID') >= 0) row[uIdx('ID')] = nid;
    var ins = unkou.getLastRow() + 1;
    unkou.getRange(ins, 1, 1, uLastCol).setValues([row]);
    if (uIdx('日付') >= 0) unkou.getRange(ins, uIdx('日付') + 1).setNumberFormat('yyyy/MM/dd');
    if (registeredIds.indexOf(nid) === -1) registeredIds.push(nid);
    try { delaySyncSummary_(nid, ss); } catch(e) {}
    return nid;
  }

  // ── パターン分岐 ──────────────────────────────────────────────────
  if (cargoRows.length === 2 && vehRows.length === 1) {
    // 【2行程パターン】1台の車両が2つの貨物を順番に運ぶ
    // IDを先に1つだけ採番し、2行とも同じIDで登録する
    var ft = checkCarType(cargoRows[0].data, vehData);
    if (ft === null) return;
    var sharedId = 'V-' + String(getNextIdNum_(unkou, 'V-')).padStart(4, '0');
    addUnkouRow(cargoRows[0].data, vehData, ft, sharedId); // 1行程目
    addUnkouRow(cargoRows[1].data, vehData, ft, sharedId); // 2行程目（同じID）

  } else {
    // 【通常パターン】1対1マッチング・貨物単独・車両単独
    var cargoData = cargoRows.length > 0 ? cargoRows[0].data : null;
    var isSameRow = cargoData && vehData && (cargoRows[0].rowNum === vehRows[0].rowNum);
    var ft2 = (cargoData && vehData && !isSameRow) ? checkCarType(cargoData, vehData) : (vehData ? String(vehData[20]||'').trim() : null);
    if (ft2 === null && cargoData && vehData && !isSameRow) return;
    addUnkouRow(cargoData, vehData, ft2);
  }

  // ── 確定後の後処理 ────────────────────────────────────────────────
  // 【重要】貨物行はA-M（貨物セクション）のみ、車両行はN-Z（車両セクション）のみ着色する。
  // cargoRowsとvehRowsを別々に処理することで「貨物が確定しても車両側の色は変わらない」を実現。
  // 同一行に両方チェックがある場合は両セクションとも着色される。
  for (var ci = 0; ci < cargoRows.length; ci++) {
    var cr = cargoRows[ci].rowNum;
    joho.getRange(cr, 1).setValue(false);                  // A: 貨物チェック解除
    joho.getRange(cr, 2).setValue('確定');                 // B: 進捗(貨物)=確定
    joho.getRange(cr, 1, 1, 13).setBackground('#fff9c4'); // A〜M: 貨物セクションのみ黄色
  }
  if (vehRows.length > 0) {
    var vr2 = vehRows[0].rowNum;
    joho.getRange(vr2, 14).setValue(false);                 // N: 車両チェック解除
    joho.getRange(vr2, 15).setValue('確定');                // O: 進捗(車両)=確定
    joho.getRange(vr2, 14, 1, 13).setBackground('#fff9c4'); // N〜Z: 車両セクションのみ黄色
  }

  var idMsg = registeredIds.length > 1
    ? registeredIds[0] + ' の2行程（同一ID・2行）'
    : registeredIds[0];
  ui.alert('✅ 配車確定\n\n' + idMsg + ' を運行シートに登録しました。\n\n情報シートの対象行を「確定」（黄色）にしました。');
}


// ================================================================
//  15-2: 情報シート→運行行データ組み立て（buildJohoNewRow_）  【大B / 中15 / 小15-2】
//
//  matchAndConfirmDispatch の内部補助関数。
//  情報シートの貨物行・車両行の各列データを読み取り、
//  運行シートのヘッダー列名に合わせて newRow 配列に値をセットする。
//
//  【マージルール】
//  ・日付・積地・降地・荷主・売上・備考 → 貨物行（cargoRow）を優先
//  ・会社名・トン数・車種・車番・乗務員名・携帯・看板名・支払 → 車両行（vehRow）を優先
//  ・overrideType が指定されている場合、車種はその値で上書き（不一致確認ダイアログの入力値）
//
//  M&A向け補足: この関数を修正することで、運行シートへの転記ルールを柔軟に変更できる。
//              例えば「トン数は貨物側を優先」に変えたい場合はこの関数のみ修正すればよい。
// ================================================================
function buildJohoNewRow_(newRow, uIdx, cargoRow, vehRow, overrideType) {
  // colName: 運行シートの列名。alt: 別名（旧フォーマット対応）。val: セットする値
  // 運行シートのヘッダーが「荷主」か「荷主名」かでSSのバージョンによって異なるため
  // 両方試して見つかった方にセットする（alt は省略可）
  function set(colName, val, alt) {
    if (val === null || val === undefined || String(val).trim() === '') return;
    var i = uIdx(colName);
    if (i < 0 && alt) i = uIdx(alt); // 第一候補がなければ別名で再検索
    if (i < 0) return;
    newRow[i] = val;
  }

  // ── 貨物情報からのセット ───────────────────────────────────────────
  // 情報シートの列インデックス（0ベース）: A=0 B=1 C=2 D=3 E=4 F=5 G=6 H=7 I=8 J=9 K=10 L=11 M=12
  if (cargoRow) {
    set('荷主',   cargoRow[2], '荷主名'); // C: 会社名(貨物) → 運行シートの荷主/荷主名列へ
    set('日付',   cargoRow[5]);           // F: 日付
    set('車種',   cargoRow[8]);           // I: 車種（貨物要求）→ 後で車両側が上書き
    set('積地',   cargoRow[9]);           // J: 積込地
    set('降地',   cargoRow[10]);          // K: 降ろし地
    set('売上',   cargoRow[11]);          // L: 金額(売上)
    set('備考',   cargoRow[12]);          // M: 備考(貨物)
  }

  // ── 車両情報からのセット ───────────────────────────────────────────
  // 情報シート26列構成での車両セクションのインデックス（0ベース）:
  // N=13:チェック(車両) O=14:進捗(車両) P=15:会社名 Q=16:TEL R=17:FAX S=18:看板名
  // T=19:トン数 U=20:車種 V=21:車番 W=22:乗務員名 X=23:携帯 Y=24:金額(支払) Z=25:備考
  if (vehRow) {
    set('会社名',   vehRow[15]); // P: 会社名(車両) → 協力会社名
    set('看板名',   vehRow[18]); // S: 看板名
    set('トン数',   vehRow[19]); // T: トン数(車両) ← 車両実績を優先
    set('車種',     vehRow[20]); // U: 車種(車両)   ← 貨物側を上書き
    set('車番',     vehRow[21]); // V: 車番
    set('乗務員名', vehRow[22]); // W: 乗務員名
    set('携帯番号', vehRow[23]); // X: 携帯番号
    set('支払い',   vehRow[24]); // Y: 金額(支払)
  }

  // 車種不一致ダイアログで入力された値で最終上書き
  if (overrideType) set('車種', overrideType);
}


function ensureCompanySettingSheet_(ss) {
  var SHEET_NAME = '自社設定';
  var items = ['会社名','郵便番号','住所','電話番号','FAX番号','担当者名','インボイス登録番号'];
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  var lastRow = sheet.getLastRow();
  var existingItems = lastRow > 0
    ? sheet.getRange(1, 1, lastRow, 1).getValues().map(function(r){ return String(r[0]||'').trim(); })
    : [];
  for (var i = 0; i < items.length; i++) {
    if (existingItems.indexOf(items[i]) === -1) {
      var nextRow = sheet.getLastRow() + 1;
      sheet.getRange(nextRow, 1).setValue(items[i]);
    }
  }
  var numRows = Math.max(sheet.getLastRow(), items.length);
  sheet.getRange(1, 1, numRows, 1)
    .setBackground('#cfd8dc').setFontWeight('bold').setFontSize(11).setVerticalAlignment('middle');
  sheet.getRange(1, 1, numRows, 2)
    .setBorder(true, true, true, true, true, true, '#78909c', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 340);
  sheet.setRowHeights(1, numRows, 28);
}


// ================================================================
//  14-2: 帳票ダイアログ表示
// ================================================================
function showHatchuDocDialog() {
  var data = getDocumentData_('sum'); // 発注書・指示書: 集計表から（金額=支払い）
  if (!data) return;
  var tmpl = HtmlService.createTemplateFromFile('documentPreview');
  tmpl.docType = 'hatchu';
  tmpl.docData = JSON.stringify(data);
  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(780).setHeight(660), '発注書・指示書'
  );
}

function showShabanDocDialog() {
  var data = getDocumentData_('unkou'); // 車番連絡: 運行シートから（金額=売上）
  if (!data) return;
  var tmpl = HtmlService.createTemplateFromFile('documentPreview');
  tmpl.docType = 'shaban';
  tmpl.docData = JSON.stringify(data);
  SpreadsheetApp.getUi().showModalDialog(
    tmpl.evaluate().setWidth(780).setHeight(580), '車番連絡'
  );
}


// ================================================================
//  14-3: アクティブ行のデータ取得（getDocumentData_）
// ================================================================
function getDocumentData_(source) {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var activeSheet = ss.getActiveSheet();
  var activeCell  = activeSheet.getActiveRange();
  var cellValue   = String(activeCell.getValue() || '').trim();

  // 運行シートからデータ取得
  var unkouSheet = ss.getSheetByName('運行');
  if (!unkouSheet) { ui.alert('運行シートが見つかりません。'); return null; }
  var uLastCol = unkouSheet.getLastColumn();
  var uHdrs    = unkouSheet.getRange(1, 1, 1, uLastCol).getValues()[0].map(function(h){ return String(h||'').trim(); });
  var uAll     = unkouSheet.getDataRange().getValues();

  var rowData = null;
  var rowId   = '';

  // ① クリックしたセルの値がIDパターン（V-XXXX等）ならそれで検索
  if (/^[VvSsMm]-\d+$/.test(cellValue)) {
    rowId = cellValue;
    for (var i = 1; i < uAll.length; i++) {
      if (String(uAll[i][0]||'').trim() === rowId) { rowData = uAll[i]; break; }
    }
  }

  // ② IDが取れなければ、アクティブ行のA列から取得（運行シート上の行選択にも対応）
  if (!rowData) {
    var activeRow = activeCell.getRow();
    if (activeRow < 2) { ui.alert('運行シートのIDセルか、データ行を選択して実行してください。'); return null; }
    if (activeSheet.getName() === '運行') {
      rowData = uAll[activeRow - 1];
      rowId   = rowData ? String(rowData[0]||'').trim() : '';
    } else {
      rowId = String(activeSheet.getRange(activeRow, 1).getValue() || '').trim();
      if (/^[VvSsMm]-\d+$/.test(rowId)) {
        for (var j = 1; j < uAll.length; j++) {
          if (String(uAll[j][0]||'').trim() === rowId) { rowData = uAll[j]; break; }
        }
      }
    }
  }

  if (!rowData || !rowId) { ui.alert('IDが取得できません。\nIDセル（V-0001等）をクリックしてから実行してください。'); return null; }

  function uVal(name) {
    var idx = uHdrs.indexOf(name);
    return idx >= 0 ? rowData[idx] : '';
  }

  // 集計表フォールバック（運行シートのヘッダー旧形式対応）
  var sumFallbackRow = null, sumFallbackHdrs = [];
  (function(){
    var sf = ss.getSheetByName('集計表');
    if (!sf || sf.getLastRow() < 2) return;
    sumFallbackHdrs = sf.getRange(1,1,1,sf.getLastColumn()).getValues()[0].map(function(h){return String(h||'').trim();});
    var sfData = sf.getRange(2,1,sf.getLastRow()-1,sf.getLastColumn()).getValues();
    for (var si2 = 0; si2 < sfData.length; si2++) {
      if (String(sfData[si2][0]||'').trim() === rowId) { sumFallbackRow = sfData[si2]; break; }
    }
  })();
  function sVal(name) {
    if (!sumFallbackRow) return '';
    var idx = sumFallbackHdrs.indexOf(name);
    return idx >= 0 ? sumFallbackRow[idx] : '';
  }
  function anyVal() {
    for (var ai = 0; ai < arguments.length; ai++) {
      var v = uVal(arguments[ai]); if (v !== '' && v !== null && v !== undefined) return v;
      v = sVal(arguments[ai]); if (v !== '' && v !== null && v !== undefined) return v;
    }
    return '';
  }

  // 日付フォーマット
  var WD = ['日','月','火','水','木','金','土'];
  var dateRaw = uVal('日付');
  var dateStr = '', wdStr = '', dropDateStr = '', dropWdStr = '';
  var baseDateObj = null;
  if (dateRaw instanceof Date) {
    baseDateObj = dateRaw;
    dateStr = Utilities.formatDate(dateRaw, 'Asia/Tokyo', 'M/d');
    wdStr   = WD[dateRaw.getDay()];
  } else if (dateRaw) {
    var d2 = new Date(dateRaw);
    if (!isNaN(d2.getTime())) {
      baseDateObj = d2;
      dateStr = Utilities.formatDate(d2, 'Asia/Tokyo', 'M/d');
      wdStr   = WD[d2.getDay()];
    } else {
      dateStr = String(dateRaw);
    }
  }
  if (baseDateObj) {
    var dropDate = new Date(baseDateObj.getFullYear(), baseDateObj.getMonth(), baseDateObj.getDate() + 1);
    dropDateStr = Utilities.formatDate(dropDate, 'Asia/Tokyo', 'M/d');
    dropWdStr   = WD[dropDate.getDay()];
  }

  // 自社設定シートから情報取得
  var self = {};
  var selfSheet = ss.getSheetByName('自社設定');
  if (selfSheet && selfSheet.getLastRow() >= 1) {
    selfSheet.getDataRange().getValues().forEach(function(r){
      if (r[0]) self[String(r[0]).trim()] = String(r[1]||'').trim();
    });
  }

  // 取引先マスタから荷主情報取得（'荷主'/'荷主名' 両ヘッダー対応）
  var clientInfo = {};
  var clientName = String(anyVal('荷主', '荷主名') || '').trim();
  var custSheet = ss.getSheetByName('マスタ');
  if (custSheet && custSheet.getLastRow() >= 2 && clientName) {
    var cHdrs = custSheet.getRange(1,1,1,custSheet.getLastColumn()).getValues()[0].map(function(h){return String(h||'').trim();});
    var cData = custSheet.getRange(2,1,custSheet.getLastRow()-1,custSheet.getLastColumn()).getValues();
    var nIdx  = cHdrs.indexOf('会社名');
    for (var ci = 0; ci < cData.length; ci++) {
      if (String(cData[ci][nIdx]||'').trim() === clientName) {
        cHdrs.forEach(function(h,idx){ clientInfo[h] = String(cData[ci][idx]||'').trim(); });
        break;
      }
    }
  }

  // 発注書（source==='sum'）の場合、集計表から支払い金額を取得
  var paymentFromSum = '';
  if (source === 'sum') {
    var sumSheet2 = ss.getSheetByName('集計表');
    if (sumSheet2 && sumSheet2.getLastRow() >= 2) {
      var sHdrs = sumSheet2.getRange(1,1,1,sumSheet2.getLastColumn()).getValues()[0].map(function(h){return String(h||'').trim();});
      var sData = sumSheet2.getRange(2,1,sumSheet2.getLastRow()-1,sumSheet2.getLastColumn()).getValues();
      var sPayIdx = sHdrs.indexOf('支払い');
      if (sPayIdx >= 0) {
        for (var si = 0; si < sData.length; si++) {
          if (String(sData[si][0]||'').trim() === rowId) { paymentFromSum = String(sData[si][sPayIdx]||''); break; }
        }
      }
    }
  }

  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年M月d日');

  return {
    id:          rowId,
    dateStr:     dateStr,
    wdStr:       wdStr,
    dropDateStr: dropDateStr,
    dropWdStr:   dropWdStr,
    company:     String(anyVal('会社名')           || ''),
    ton:         String(anyVal('トン数', 'トン')   || ''),
    carType:     String(anyVal('車種')             || ''),
    car:         String(anyVal('車番')             || ''),
    driver:      String(anyVal('乗務員名', '乗務員')|| ''),
    phone:       String(anyVal('携帯番号', '電話') || ''),
    signboard:   String(anyVal('看板名', '看板')   || ''),
    client:      clientName,
    pickPlace:   String(anyVal('積地', '積込地')   || ''),
    dropPlace:   String(anyVal('降地', '降ろし地') || ''),
    sales:       String(anyVal('売上')             || ''),
    payment:     source === 'sum' ? paymentFromSum : String(anyVal('支払い') || ''),
    tollReq:     String(anyVal('請求(高速代)', '請求高速') || ''),
    memo:        String(anyVal('備考')                    || ''),
    sobiSonota:  String(anyVal('装備その他', '装備')      || ''),
    clientTel:   clientInfo['電話']        || '',
    clientAddr:  clientInfo['住所']        || '',
    clientEmail: clientInfo['メールアドレス'] || '',
    clientFax:   clientInfo['FAX']         || '',
    today:       today,
    selfName:    self['会社名']     || '',
    selfZip:     self['郵便番号']   || '',
    selfAddr:    self['住所']       || '',
    selfTel:     self['電話番号']   || '',
    selfFax:     self['FAX番号']    || '',
    selfPerson:  self['担当者名']   || ''
  };
}