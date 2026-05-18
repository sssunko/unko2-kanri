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
function delaySyncSummary_(id) { try { syncSummaryForId_(id); } catch(e) {} }


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
    : [19, 20, 21, 22, 25, 26, 27, 28, 33, 34];
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
  for (var i = 0; i < cols.length; i++) {
    sheet.getRange(startRow, cols[i], numRows, 1).setNumberFormat(fmt);
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

  // 運行シートのL列(12=積地)
  var sheet = ss.getSheetByName('運行');
  if (sheet && sheet.getLastRow() >= 2) {
    var lr   = sheet.getLastRow();
    // J列(10)=日付 と L列(12)=積地 を一括読み込み（col10〜12 の3列）
    var vals = sheet.getRange(2, 10, lr - 1, 3).getValues();
    var bgs  = vals.map(function(r) {
      var dateV = r[0];           // col10: 日付
      var pickV = String(r[2] || ''); // col12: 積地
      if (pickV.indexOf('休み') !== -1 || pickV.indexOf('有休') !== -1) return ['#9e9e9e'];
      if (pickV === '' && dateV instanceof Date) {
        var d = new Date(dateV); d.setHours(0, 0, 0, 0);
        if (d >= today) return ['#fff9c4'];
      }
      return [null];
    });
    sheet.getRange(2, 12, lr - 1, 1).setBackgrounds(bgs);
  }

  // 集計表のL列(12=積地)も同様
  var sumSheet = ss.getSheetByName('集計表');
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var slr   = sumSheet.getLastRow();
    var svals = sumSheet.getRange(2, 10, slr - 1, 3).getValues();
    var sbgs  = svals.map(function(r) {
      var dateV = r[0];
      var pickV = String(r[2] || '');
      if (pickV.indexOf('休み') !== -1 || pickV.indexOf('有休') !== -1) return ['#9e9e9e'];
      if (pickV === '' && dateV instanceof Date) {
        var d = new Date(dateV); d.setHours(0, 0, 0, 0);
        if (d >= today) return ['#fff9c4'];
      }
      return [null];
    });
    sumSheet.getRange(2, 12, slr - 1, 1).setBackgrounds(sbgs);
  }
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

  // 全データを値として一括取得（数式は値に変換される）
  var data = sheet.getRange(2, 1, numRows, 26).getValues();

  // J列(index[9]=日付)昇順 → G列(index[6]=乗務員名)昇順 でソート
  data.sort(function(a, b) {
    var da = (a[9] instanceof Date) ? a[9].getTime() : 0;
    var db = (b[9] instanceof Date) ? b[9].getTime() : 0;
    if (da !== db) return da - db;
    return String(a[6] || '').localeCompare(String(b[6] || ''));
  });

  // V列(22, 0-indexed:21)を空にして値を書き戻す
  var writeData = data.map(function(r) {
    var row = r.slice();
    row[21] = ''; // 合計高速は数式で再セットするため空に
    return row;
  });
  sheet.getRange(2, 1, numRows, 26).setValues(writeData);

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
  var cols = Math.max(colCount, 34);
  var data = sheet.getRange(2, 1, numRows, cols).getValues();
  data.sort(function(a, b) {
    var da = (a[9] instanceof Date) ? a[9].getTime() : 0;
    var db = (b[9] instanceof Date) ? b[9].getTime() : 0;
    if (da !== db) return da - db;
    return String(a[6] || '').localeCompare(String(b[6] || ''));
  });
  // 数式列を空にして値を書き戻す
  var writeData = data.map(function(r) {
    var row = r.slice();
    row[21] = ''; row[25] = ''; row[27] = ''; // V/Z/AB列
    return row;
  });
  sheet.getRange(2, 1, numRows, cols).setValues(writeData);
  // 数式を一括再セット
  var f22 = [], f26 = [], f28 = [];
  for (var i = 0; i < numRows; i++) {
    var rn = i + 2;
    f22.push(['=IF(AND(U'+rn+'="",T'+rn+'=""),"",U'+rn+'-T'+rn+')']);
    f26.push(['=IF(OR(W'+rn+'="",X'+rn+'=""),"",W'+rn+'/X'+rn+'*Y'+rn+')']);
    f28.push(['=IF(AND(S'+rn+'="",V'+rn+'="",Z'+rn+'="",AA'+rn+'=""),"",S'+rn+'-(V'+rn+'+Z'+rn+'+AA'+rn+'))']);
  }
  sheet.getRange(2, 22, numRows, 1).setFormulas(f22);
  sheet.getRange(2, 26, numRows, 1).setFormulas(f26);
  sheet.getRange(2, 28, numRows, 1).setFormulas(f28);
  sheet.getRange(2, 10, numRows, 1).setNumberFormat('yyyy/MM/dd');
  applyMoneyFormat_(sheet, 2, numRows, 'summary');
  applyDateTimeFormat_(sheet, 2, numRows);
}


// ================================================================
//  2-1: メニュー設定（onOpen）  【大C / 中2 / 小2-1】
//  スプレッドシート上部に「メニュー」を表示する（客先配布用）
//  項目：ホーム画面を表示 / 集計表再生成 / シート再生成 /
//        月生成 / 前月分アーカイブ / 写真・ファイル取込
// ================================================================
function onOpen() {
  // '__COMPANY_SS__' シートがあれば客先配布SSとして客用メニューを表示
  // マスターSSにはこのシートがないため常に管理者メニューを表示
  var isAdmin = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('__COMPANY_SS__') === null;

  var menu = SpreadsheetApp.getUi().createMenu('メニュー')
    .addItem('ホーム画面を表示', 'showSidebar')
    .addSeparator()
    .addItem('📅 今月分生成（途中契約）', 'generateCurrentMonth')
    .addItem('📅 翌月分生成', 'generateNextMonth')
    .addItem('📦 前月分アーカイブ', 'archiveOldMonth')
    .addSeparator()
    .addItem('集計表再生成', 'generateSummary')
    .addItem('シート再生成', 'expandAndRefreshSheets')
    .addItem('💴 経費自動入力', 'autoFillExpense')
    .addItem('🔃 日付順並び替え', 'sortUnkouByDate_')
    .addSeparator()
    .addItem('📷 写真・ファイル取込', 'showUploadSidebar')
    .addItem('📖 使い方シート作成', 'createUsageSheet');

  if (isAdmin) {
    menu
      .addSeparator()
      .addItem('シート保護設定', 'setupSheetProtection')
      .addSeparator()
      .addItem('🏢 会社セットアップ実行', 'setupCompanies')
      .addItem('📤 会社SS作成＆メール送信', 'sendCompanySetupEmails')
      .addItem('📧 配布メール送信', 'triggerDistributionMail')
      .addItem('📝 申し込みフォーム作成', 'createSignupForm')
      .addSeparator()
      .addItem('🔧 初期設定', 'installTriggers');
  }

  menu.addToUi();

  // マスターSSのIDを保存（時間トリガーからopenByIdで開くために使う）
  try {
    PropertiesService.getScriptProperties().setProperty(
      'masterSsId', SpreadsheetApp.getActiveSpreadsheet().getId()
    );
  } catch(ex) {}
  convertLegacyAdminDataUrls_();
  applyHolidayRowColors_();
}


// ================================================================
//  2-2: Webアプリ起動（doGet）  【大C / 中2 / 小2-2】
//  URLアクセス時にWebアプリとして表示する。
//  ?ssId=XXXX パラメータを受け取り、HTMLテンプレートに渡す。
//  初回アクセス時に自分自身のWebアプリURLを自動取得してScript Propertiesに保存する。
//  これにより会社SS作成時のアプリURLが自動で設定される（手動入力不要）。
// ================================================================
function doGet(e) {
  var ssId = (e && e.parameter && e.parameter.ssId) ? e.parameter.ssId : '';

  // WebアプリURLを常に最新に更新（deployするたびに正しいURLで上書き）
  try {
    var svcUrl = ScriptApp.getService().getUrl();
    if (svcUrl && svcUrl.indexOf('script.google.com/macros') !== -1) {
      PropertiesService.getScriptProperties().setProperty('webAppUrl', svcUrl);
    }
  } catch(ex) {}

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
//  2-3: サイドバー表示（showSidebar）  【大C / 中2 / 小2-3】
//  スプレッドシートのサイドバーとして表示する
// ================================================================
function showSidebar() {
  var html = HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('ホーム').setWidth(400);
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
    var ss        = SpreadsheetApp.getActiveSpreadsheet();

    // ── 3-1-1: 集計表 編集ブロック ──────────────────────────────────
    // 距離(V=22)・ガソリン代(X=24)・支払い(Z=26)・備考(AB=28)以外の列は直接編集を禁止する
    // 禁止列が触れられた場合: IDがあれば集計表を再同期して正しい値に戻す
    //                         IDがなく単一セルなら旧値に戻す
    if (sheetName === '集計表' && row > 1) {
      var allowed = [23, 25, 27, 29, 34]; // W=距離, Y=ガソリン代, AA=支払い, AC=備考, AH=その他手当
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

    // ── 3-1-2: 運行シート V列(22) 合計高速 保護 ──────────────────────
    // 合計高速は「実費高速 - 請求高速」の自動計算列なので直接編集を禁止する
    // 編集されたら即座に数式を復元する
    if (sheetName === '運行' && row > 1 && col === 22) {
      range.setFormula('=IF(AND(U'+row+'="",T'+row+'=""),"",U'+row+'-T'+row+')');
      ss.toast('合計高速は自動計算列です', '⛔ 保護', 3);
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
    if (sheetName === '自車専属マスタ') { onEditMasterVehicle_(sheet, range); return; }
    if (sheetName === 'マスタ')         { onEditMasterCustomer_(sheet, range); return; }
    // 会社登録シートの処理はインストール型トリガー（installedOnEdit_）が担当する
    // シンプルトリガーはドライブ/メール/トリガー作成などの認証付き操作が不可のため
    if (sheetName === '会社登録') return;
    if (sheetName !== '運行') return;
    onEditUnkou_(sheet, range);
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
    // J列(10)の日付：時刻部分が 0:00:00 なら現在時刻を付与（日付のみ入力に対応）
    var dateCell = sheet.getRange(row, 10);
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
    // Y列(25)連絡(端末)・Z列(26)データ(端末)：削除・複数セル貼り付けはOK、手入力はブロック
    // GAS側書き込み（saveTermNoticeByDriver 等）はシンプルトリガー非発火のため影響なし
    if (editedCol === 25 || editedCol === 26) {
      var isEmptyEdit = (e.value === undefined || String(e.value || '').trim() === '');
      var isPasteEdit = (range.getNumColumns() > 1 || range.getNumRows() > 1);
      if (!isEmptyEdit && !isPasteEdit) {
        var colLabel = editedCol === 25 ? '連絡(端末)' : 'データ(端末)';
        range.setValue(e.oldValue !== undefined ? e.oldValue : '');
        SpreadsheetApp.getActiveSpreadsheet().toast(colLabel + 'はアプリからのみ入力できます（削除・貼り付けはOK）', '⛔ 保護', 4);
      }
      continue;
    }
    // F列(6)：車番を入力→自車専属マスタと部分一致で他項目を自動補完
    if (editedCol === 6 && range.getNumColumns() === 1) {
      var inputCar = String(sheet.getRange(row, 6).getValue()).trim();
      if (inputCar && mData.length > 1) {
        // B〜I列（F=col6を除く）にすでに値があれば補完しない（先入力優先）
        var bToI = sheet.getRange(row, 2, 1, 8).getValues()[0];
        var hasPreInput = false;
        for (var ci = 0; ci < bToI.length; ci++) {
          if (ci === 4) continue; // index4=F列(col6)はスキップ
          if (String(bToI[ci] || '').trim() !== '') { hasPreInput = true; break; }
        }
        if (!hasPreInput) {
          for (var m2 = 1; m2 < mData.length; m2++) {
            var masterCar = String(mData[m2][7] || '').trim();
            if (masterCar === inputCar || masterCar.indexOf(inputCar) !== -1 || inputCar.indexOf(masterCar) !== -1) {
              sheet.getRange(row, 2).setValue(mData[m2][2]);
              sheet.getRange(row, 3).setValue(mData[m2][3]);
              sheet.getRange(row, 4).setValue(mData[m2][5]);
              sheet.getRange(row, 5).setValue(mData[m2][6]);
              sheet.getRange(row, 6).setValue(masterCar);
              sheet.getRange(row, 7).setValue(mData[m2][8]);
              sheet.getRange(row, 8).setValue(mData[m2][9]);
              sheet.getRange(row, 9).setValue(mData[m2][4]);
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
      var baseDateObj = (sheet.getRange(row, 10).getValue() instanceof Date) ? sheet.getRange(row, 10).getValue() : null;
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
    // 積地(L=col12)の背景色: 休み/有休→グレー、空+未来日→黄色(配車漏れ警告)、それ以外→なし
    var pvK   = String(sheet.getRange(row, 12).getValue() || '');
    var dateV = sheet.getRange(row, 10).getValue();
    var pickBg;
    if (pvK.indexOf('休み') !== -1 || pvK.indexOf('有休') !== -1) {
      pickBg = '#9e9e9e';
    } else if (pvK === '' && dateV instanceof Date) {
      var todayMid = new Date(); todayMid.setHours(0, 0, 0, 0);
      var dateMid  = new Date(dateV); dateMid.setHours(0, 0, 0, 0);
      pickBg = (dateMid >= todayMid) ? '#fff9c4' : null;
    } else {
      pickBg = null;
    }
    sheet.getRange(row, 12).setBackground(pickBg);
    var tollCell = sheet.getRange(row, 22);
    if (!tollCell.getFormula()) {
      tollCell.setFormula('=IF(AND(U' + row + '="",T' + row + '=""),"",U' + row + '-T' + row + ')');
    }
    var newId = sheet.getRange(row, 1).getValue();
    if (newId) syncSummaryForId_(newId);
  }
  applyMoneyFormat_(sheet, startRow, numRows, 'unkou');
  applyDateTimeFormat_(sheet, startRow, numRows);
  cleanAllOrphanSummary_();
}


// ================================================================
//  3-3: 自車専属マスタ編集時の処理（onEditMasterVehicle_）  【大B / 中3 / 小3-3】
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
    // 仮日数/給料/%が変わったら集計表の該当行（車番+乗務員名一致）に即反映
    var mRow = sheet.getRange(row, 1, 1, 16).getValues()[0];
    var mCar    = String(mRow[7]  || '').trim();
    var mName   = String(mRow[8]  || '').trim();
    var mKari   = mRow[13];
    var mKyuryo = mRow[14];
    var mPct    = mRow[15];
    if (mCar || mName) {
      var sumSheet = ss.getSheetByName('集計表');
      if (sumSheet && sumSheet.getLastRow() >= 2) {
        var sumData = sumSheet.getRange(2, 1, sumSheet.getLastRow()-1, 35).getValues();
        for (var s = 0; s < sumData.length; s++) {
          var sCar  = String(sumData[s][6] || '').trim();
          var sName = String(sumData[s][7] || '').trim();
          if (sCar === mCar && sName === mName) {
            var sRow = s + 2;
            if (mKari   !== '') sumSheet.getRange(sRow, 31).setValue(mKari);   // AE=仮日数
            if (mKyuryo !== '') sumSheet.getRange(sRow, 32).setValue(mKyuryo); // AF=給料
            if (mPct    !== '') sumSheet.getRange(sRow, 33).setValue(mPct);    // AG=%
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
    // B列（ステータス）が変更された場合、運行シートを今日以降で同期する
    var editedStartCol = range.getColumn();
    var editedEndCol   = editedStartCol + range.getNumColumns() - 1;
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
  // ③ ステータスが「運行」なら今日〜今月末の行を生成
  if (status === '運行') {
    var now = new Date();
    var yr = now.getFullYear(), mo = now.getMonth(), startDay = now.getDate();
    var endDay = new Date(yr, mo + 1, 0).getDate();
    var lock = LockService.getDocumentLock();
    try { lock.waitLock(30000); } catch(e) { return; }
    try {
      var insertRow = sheet.getLastRow() + 1;
      var nextNum   = getNextIdNum_(sheet, 'V-');
      var rowsData  = [], formulas = [];
      for (var day = startDay; day <= endDay; day++) {
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

  // 自車専属マスタから 車番+乗務員名 → 仮日数/給料/%/月間経費 の支払条件マップを作成
  var master = ss.getSheetByName('自車専属マスタ');
  var payCondMap = {};
  if (master && master.getLastRow() >= 2) {
    var mData = master.getRange(2, 1, master.getLastRow()-1, 31).getValues();
    for (var m = 0; m < mData.length; m++) {
      var mcar  = String(mData[m][7]  || '').trim();
      var mname = String(mData[m][8]  || '').trim();
      var pkey  = mcar + '_' + mname;
      var mExp  = 0;
      for (var ei = 16; ei <= 30; ei++) mExp += Number(mData[m][ei]) || 0;
      payCondMap[pkey] = {
        kari:    mData[m][13] || '',
        kyuryo:  mData[m][14] || '',
        pct:     mData[m][15] || '',
        expense: mExp
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
        name:r[6], tel:r[7], kanban:r[8], date:r[9], clients:[],
        picks:[], drops:[],
        guideTime:'',
        pickTime:'', restStart:'', restEnd:'', dropTime:'',
        rawPickTime:null, rawRestStart:null, rawRestEnd:null, rawDropTime:null,
        sales:0, tollReq:0, tollReal:0
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
    var pc     = payCondMap[pkey] || {kari:'', kyuryo:'', pct:'', expense:0};
    var kari   = (pc.kari   !== undefined && pc.kari   !== '') ? pc.kari   : (old.kari   || '');
    var kyuryo = (pc.kyuryo !== undefined && pc.kyuryo !== '') ? pc.kyuryo : (old.kyuryo || '');
    var pct    = (pc.pct    !== undefined && pc.pct    !== '') ? pc.pct    : (old.pct    || '');

    var gpick = g.picks.join('・'), gdrop = g.drops.join('・');
    var gIsYukyu  = gpick.indexOf('有休') !== -1 || gdrop.indexOf('有休') !== -1;
    var gIsYasumi = !gIsYukyu && (gpick.indexOf('休み') !== -1 || gdrop.indexOf('休み') !== -1);
    // 経費按分（有休・休みは0円、実稼働日数で按分）
    var gExpense = 0;
    if (!gIsYukyu && !gIsYasumi && g.date instanceof Date) {
      var gYmKey   = g.date.getFullYear() + '_' + g.date.getMonth();
      var gVKey    = String(g.car||'').trim() + '_' + String(g.name||'').trim() + '_' + gYmKey;
      var gWorkDays= workDayMap[gVKey] || 1;
      gExpense     = Math.round((pc.expense || 0) / gWorkDays);
    }
    outRows.push([
      g.id, g.kubun, g.company, g.tons, g.type, g.car, g.name, g.tel, g.kanban||g.company,
      g.date, g.clients.join('・'), gpick, gdrop,
      g.guideTime||'', g.pickTime, g.restStart, g.restEnd, g.dropTime,
      g.sales||'', g.tollReq||'', g.tollReal||'', '',
      old.distance||'', fuel, old.gas||'', '',
      old.pay||'', gExpense, '', old.memo||'',   // AA=支払い, AB=経費合計, AC=利益空, AD=備考
      kari, kyuryo, pct,
      gIsYukyu ? yukyuRate : '',
      old.other || ''   // AI=その他手当（手入力保持）
    ]);
  }

  // 集計表を全クリアして再書き込み
  sumSheet.clear();
  if (outRows.length > 0) {
    sumSheet.getRange(1, 1, outRows.length, 35).setValues(outRows);
    sumSheet.setFrozenRows(1);

    // 4時間超で黄色（労働時間過超）、30分未満で水色（休憩不足）の判定閾値
    var F = 4*60*60*1000;
    var T = 30*60*1000;

    for (var row = 2; row <= outRows.length; row++) {
      // V列(22)：合計高速代 = 実費(U) - 請求(T)（差額がマイナスなら持ち出し）
      sumSheet.getRange(row, 22).setFormula('=IF(AND(U'+row+'="",T'+row+'=""),"",U'+row+'-T'+row+')');
      // Z列(26)：燃料代 = 距離(W) / 燃費(X) * ガソリン代(Y)
      sumSheet.getRange(row, 26).setFormula('=IF(OR(W'+row+'="",X'+row+'=""),"",W'+row+'/X'+row+'*Y'+row+')');
      // AC列(29)：利益 = 売上(S) - (合計高速(V) + 燃料代(Z) + 支払(AA) + 経費合計(AB))
      sumSheet.getRange(row, 29).setFormula('=IF(AND(S'+row+'="",V'+row+'="",Z'+row+'="",AA'+row+'="",AB'+row+'=""),"",S'+row+'-(V'+row+'+Z'+row+'+AA'+row+'+AB'+row+'))');

      var g2        = idMap[idOrder[row-2]];
      var keepPay   = outRows[row-1][26] || '';
      var keepDist  = outRows[row-1][22] || '';
      var keepGas   = outRows[row-1][24] || '';
      var rowExpense= outRows[row-1][27] || 0; // AB=経費合計
      var calcToll  = (Number(g2.tollReal)||0)-(Number(g2.tollReq)||0);
      var calcFuel  = (Number(keepDist)&&Number(fuel)&&Number(keepGas)) ? (Number(keepDist)/Number(fuel)*Number(keepGas)) : 0;
      var calcProfit= (Number(g2.sales)||0)-(calcToll+calcFuel+(Number(keepPay)||0)+(Number(rowExpense)||0));

      // 利益がマイナスの行は薄赤で警告表示
      sumSheet.getRange(row, 1, 1, 33).setBackground(calcProfit < 0 ? '#ffebee' : null);
      // 積完〜降完間の労働時間・休憩時間を判定して背景色で警告
      sumSheet.getRange(row, 15, 1, 4).setBackground(null);
      if (g2.rawPickTime  && g2.rawRestStart && (g2.rawRestStart-g2.rawPickTime)  > F) { sumSheet.getRange(row,15,1,2).setBackground('#ffd600'); }
      if (g2.rawRestStart && g2.rawRestEnd   && (g2.rawRestEnd  -g2.rawRestStart) < T) { sumSheet.getRange(row,16,1,2).setBackground('#4fc3f7'); }
      if (g2.rawRestEnd   && g2.rawDropTime  && (g2.rawDropTime -g2.rawRestEnd)   > F) { sumSheet.getRange(row,17,1,2).setBackground('#ffd600'); }
    }
    applyMoneyFormat_(sumSheet, 2, outRows.length - 1, 'summary');
    applyDateTimeFormat_(sumSheet, 2, outRows.length - 1);
  }

  calculatePaymentAmount();
  convertLegacyAdminDataUrls_();
  applyHolidayRowColors_();
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
    var mData = master.getRange(2, 1, master.getLastRow()-1, 31).getValues();
    for (var m = 0; m < mData.length; m++) {
      var mcar  = String(mData[m][7]  || '').trim();
      var mname = String(mData[m][8]  || '').trim();
      var mExp  = 0;
      for (var ei = 16; ei <= 30; ei++) mExp += Number(mData[m][ei]) || 0;
      payCondMap[mcar+'_'+mname] = {
        kari:    mData[m][13] || '',
        kyuryo:  mData[m][14] || '',
        pct:     mData[m][15] || '',
        expense: mExp
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
    var sumIds   = sumSheet.getRange(2, 1, sumLast-1, Math.max(colCount, 35)).getValues();
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

  var pkey   = String(g.car||'').trim()+'_'+String(g.name||'').trim();
  var pc     = payCondMap[pkey] || {kari:'', kyuryo:'', pct:''};
  var kari   = (pc.kari   !== undefined && pc.kari   !== '') ? pc.kari   : keepKari;
  var kyuryo = (pc.kyuryo !== undefined && pc.kyuryo !== '') ? pc.kyuryo : keepKyuryo;
  var pct    = (pc.pct    !== undefined && pc.pct    !== '') ? pc.pct    : keepPct;

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
  var expWorkDays = Math.max(Object.keys(uniqueWIds).length, 1);
  var pcExp       = (payCondMap[pkey] || {}).expense || 0;
  var expenseVal  = (sIsYukyu || sIsYasumi) ? 0 : Math.round(pcExp / expWorkDays);

  var rowData = [
    g.id, g.kubun, g.company, g.tons, g.type, g.car, g.name, g.tel, g.kanban||g.company,
    g.date, g.clients.join('・'), spick, sdrop,
    g.guideTime||'', g.pickTime, g.restStart, g.restEnd, g.dropTime,
    g.sales||'', g.tollReq||'', g.tollReal||'', '',
    keepDistance, fuel, keepGas, '', keepPay, expenseVal, '', keepMemo,  // AA=支払い, AB=経費合計, AC=利益空, AD=備考
    kari, kyuryo, pct,
    sIsYukyu ? yukyuRate : '',
    keepOther   // AI=その他手当（手入力保持）
  ];

  if (sumRow > 0) {
    sumSheet.getRange(sumRow, 1, 1, 35).setValues([rowData]);
  } else {
    sumRow = sumSheet.getLastRow()+1;
    if (sumRow === 1) {
      var hdr = ['ID','区分','会社名','トン数','車種','車番','乗務員名','携帯番号','看板名','日付','荷主','積地','降地','誘導時刻','積完時刻','休憩開始','休憩終了','降完時刻','売上','請求(高速代)','実費(高速代)','合計(高速代)','距離','燃費','ガソリン代','燃料代','支払い','経費合計','利益','備考','仮日数','給料','％','有休手当','その他手当'];
      sumSheet.getRange(1, 1, 1, 35).setValues([hdr]);
      sumSheet.setFrozenRows(1);
      sumRow = 2;
    }
    sumSheet.getRange(sumRow, 1, 1, 35).setValues([rowData]);
  }

  sumSheet.getRange(sumRow, 22).setFormula('=IF(AND(U'+sumRow+'="",T'+sumRow+'=""),"",U'+sumRow+'-T'+sumRow+')');
  sumSheet.getRange(sumRow, 26).setFormula('=IF(OR(W'+sumRow+'="",X'+sumRow+'=""),"",W'+sumRow+'/X'+sumRow+'*Y'+sumRow+')');
  sumSheet.getRange(sumRow, 29).setFormula('=IF(AND(S'+sumRow+'="",V'+sumRow+'="",Z'+sumRow+'="",AA'+sumRow+'="",AB'+sumRow+'=""),"",S'+sumRow+'-(V'+sumRow+'+Z'+sumRow+'+AA'+sumRow+'+AB'+sumRow+'))');

  var calcToll  = (Number(g.tollReal)||0)-(Number(g.tollReq)||0);
  var calcFuel  = (Number(keepDistance)&&Number(fuel)&&Number(keepGas)) ? (Number(keepDistance)/Number(fuel)*Number(keepGas)) : 0;
  var calcProfit= (Number(g.sales)||0)-(calcToll+calcFuel+(Number(keepPay)||0)+(Number(expenseVal)||0));
  sumSheet.getRange(sumRow, 1, 1, 32).setBackground(calcProfit < 0 ? '#ffebee' : null);

  var F = 4*60*60*1000;
  var T = 30*60*1000;
  sumSheet.getRange(sumRow, 15, 1, 4).setBackground(null);
  if (rawPickTime  && rawRestStart && (rawRestStart-rawPickTime)  > F) { sumSheet.getRange(sumRow,15,1,2).setBackground('#ffd600'); }
  if (rawRestStart && rawRestEnd   && (rawRestEnd  -rawRestStart) < T) { sumSheet.getRange(sumRow,16,1,2).setBackground('#4fc3f7'); }
  if (rawRestEnd   && rawDropTime  && (rawDropTime -rawRestEnd)   > F) { sumSheet.getRange(sumRow,17,1,2).setBackground('#ffd600'); }
  applyMoneyFormat_(sumSheet, sumRow, 1, 'summary');
  applyDateTimeFormat_(sumSheet, sumRow, 1);
  // この行だけの支払い(AA=col27)をインライン計算（全行ループの calculatePaymentAmount を避けて高速化）
  var pctNum    = Number(pct)    || 0;
  var kyuryoNum = Number(kyuryo) || 0;
  var kariNum   = Number(kari)   || 0;
  var thisToll  = (Number(g.tollReal) || 0) - (Number(g.tollReq) || 0);
  var payCell   = sumSheet.getRange(sumRow, 27);
  sumSheet.getRange(sumRow, 31, 1, 3).setBackground(null); // AE〜AG = 仮日数/給料/%
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
    if (!kyuryoNum) sumSheet.getRange(sumRow, 32).setBackground('#f4cccc'); // AF=給料
    if (!kariNum)   sumSheet.getRange(sumRow, 31).setBackground('#f4cccc'); // AE=仮日数
  } else {
    if (!keepPay) payCell.setBackground('#f4cccc');
  }
  // 有休手当(AH=col34): 歩合制+有休のみ
  sumSheet.getRange(sumRow, 34).setValue(yukyuVal);
  // 集計表のL列(12=積地)に休み/有休が含まれる場合はグレー着色
  sumSheet.getRange(sumRow, 12).setBackground(
    (spick.indexOf('休み') !== -1 || spick.indexOf('有休') !== -1) ? '#9e9e9e' : null
  );
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

  refreshActiveVehiclesAuto_();
  syncAllVehiclesToCurrentMonth_(); // 車両ステータスに応じて運行シートを今月分で同期

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

  var row = sheet.getActiveRange().getRow();
  if (row < 2) { ui.alert('データ行（2行目以降）を選択してください。'); return; }

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

  // トン数を読んで数値に変換（全角→半角、"t"等を除去）
  var tonsRaw = String(sheet.getRange(row, 6).getValue() || '').trim();
  tonsRaw = tonsRaw.replace(/[０-９]/g, function(c){ return String.fromCharCode(c.charCodeAt(0)-0xFEE0); });
  var tNum = parseInt(tonsRaw.replace(/[^0-9]/g, ''), 10);

  if (isNaN(tNum) || tNum < 1) {
    var res = ui.alert(
      'トン数が不明です',
      'F列のトン数が読み取れませんでした。\n4トンの平均値を入力しますか？',
      ui.ButtonSet.YES_NO
    );
    if (res !== ui.Button.YES) return;
    tNum = 4;
  }
  // 1〜30の範囲にクランプ
  tNum = Math.min(30, Math.max(1, tNum));
  var vals = expTable[tNum];

  // 手入力済みセルの判定：文字色が赤（#cc0000）でなく値が入っているセルは手入力→スキップ
  var expRange    = sheet.getRange(row, 17, 1, 15);
  var existVals   = expRange.getValues()[0];
  var fontColors  = expRange.getFontColors()[0];
  var AUTO_COLOR  = '#cc0000';

  var manualCount = 0;
  for (var k = 0; k < 15; k++) {
    var hasVal   = existVals[k] !== '' && existVals[k] !== 0;
    var isAuto   = String(fontColors[k] || '').toLowerCase() === AUTO_COLOR;
    if (hasVal && !isAuto) manualCount++;
  }
  if (manualCount > 0) {
    ui.alert('手入力済みの項目（' + manualCount + '件）はそのまま残します。\n赤字の項目のみ平均値で更新します。');
  }

  // セルごとに判定して書き込み（手入力済みはスキップ）
  for (var j = 0; j < 15; j++) {
    var hasVal2 = existVals[j] !== '' && existVals[j] !== 0;
    var isAuto2 = String(fontColors[j] || '').toLowerCase() === AUTO_COLOR;
    if (!hasVal2 || isAuto2) {
      sheet.getRange(row, 17 + j).setValue(vals[j]).setFontColor(AUTO_COLOR);
    }
  }

  ui.alert(tNum + 'トンの平均値を入力しました（赤字）。\n実態に合わせて修正してください。修正すると黒字に変わります。');
}


// ================================================================
//  4-4: 支払い再計算（calculatePaymentAmount）  【大B / 中4 / 小4-4】
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

  var data = sheet.getRange(2, 1, lastRow-1, 35).getValues();
  var yukyuVals = [];

  for (var i = 0; i < data.length; i++) {
    var rowNum    = i + 2;
    var sales     = Number(data[i][18]) || 0;
    var totalToll = Number(data[i][21]) || 0;
    var kari      = Number(data[i][30]) || 0;  // AE=col31=index30=仮日数
    var kyuryo    = Number(data[i][31]) || 0;  // AF=col32=index31=給料
    var pct       = Number(data[i][32]) || 0;  // AG=col33=index32=%
    var pick      = String(data[i][11] || '');
    var drop      = String(data[i][12] || '');
    var isYukyu   = pick.indexOf('有休') !== -1 || drop.indexOf('有休') !== -1;
    var isYasumi  = !isYukyu && (pick.indexOf('休み') !== -1 || drop.indexOf('休み') !== -1);
    var yCell     = sheet.getRange(rowNum, 27); // AA=col27=支払い

    yCell.setBackground(null);
    sheet.getRange(rowNum, 31, 1, 3).setBackground(null); // AE〜AG=仮日数/給料/%

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
        if (!kyuryo) sheet.getRange(rowNum, 32).setBackground('#f4cccc'); // AF=給料
        if (!kari)   sheet.getRange(rowNum, 31).setBackground('#f4cccc'); // AE=仮日数
      }
    } else {
      if (yCell.getValue() === '') yCell.setBackground('#f4cccc');
    }
    yukyuVals.push([yukyuVal]);
  }

  // 有休手当(AH=col34)を一括書込
  if (yukyuVals.length > 0) {
    sheet.getRange(2, 34, yukyuVals.length, 1).setValues(yukyuVals);
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

  var lock = LockService.getDocumentLock();
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
  var lock = LockService.getDocumentLock();
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

  // 前月の年・月(0-indexed)を計算
  var today    = new Date();
  var prevYear = (today.getMonth() === 0) ? today.getFullYear() - 1 : today.getFullYear();
  var prevMon  = (today.getMonth() === 0) ? 11 : today.getMonth() - 1;

  // 前月データの存在確認
  var prevRows = getMonthRows_(sheet, prevYear, prevMon, 26);
  if (prevRows.length === 0) {
    ui.alert(
      '📦 前月分アーカイブ\n\n' +
      prevYear + '年' + (prevMon + 1) + '月のデータが見つかりません。\n既にアーカイブ済みか、データがない可能性があります。'
    );
    return;
  }

  var companyName = getCompanyName_(ss);
  var result = archiveMonthData_(ss, prevYear, prevMon, companyName);

  ui.alert(
    '📦 アーカイブ完了\n\n' +
    prevYear + '年' + (prevMon + 1) + '月分\n' +
    result.archived + '行をアーカイブしました。\n\nファイル名: ' + result.fileName +
    '\n保存先: 運行管理_アーカイブ/' + companyName + '/'
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

  // 元データから対象行を削除（行番号を降順にして下から削除）
  var unkouNums = unkouRows.map(function(r) { return r.rowNum; }).sort(function(a, b) { return b - a; });
  for (var d = 0; d < unkouNums.length; d++) { unkouSheet.deleteRow(unkouNums[d]); }

  if (sumRows.length > 0) {
    var sumNums = sumRows.map(function(r) { return r.rowNum; }).sort(function(a, b) { return b - a; });
    for (var e = 0; e < sumNums.length; e++) { sumSheet.deleteRow(sumNums[e]); }
  }

  return { fileName: fileName, archived: unkouRows.length };
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
  var savedEmail = hintEmail || PropertiesService.getUserProperties().getProperty('linkedEmail');
  var result = { email: savedEmail || "", profile: null };
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
function saveRunState(state) {
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
function clearRunState() {
  var p    = PropertiesService.getUserProperties();
  var keys = ['picks','drops','rows','runId','guideDone','pickDone','dropDone','phase','lastPickIndex','guideHistory','pickHistory','dropHistory'];
  for (var i = 0; i < keys.length; i++) { p.deleteProperty(keys[i]); }
}


// ================================================================
//  7-1: 今日の行程取得（getTodayRoutes）  【大A / 中7 / 小7-1】
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
function createParentRows(picks, drops, dateStr, overrideInfo, companySsId) {
  // 端末のメールアドレスを確認（未連携なら運行開始不可）
  var savedEmail = PropertiesService.getUserProperties().getProperty('linkedEmail');
  if (!savedEmail) throw new Error('紐づけされていません');

  // 同時に複数端末が運行開始した場合のID重複を防ぐためロック取得
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) { throw new Error('混雑中です。少し待ってから再試行してください'); }

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

    // 日付: 端末から渡された場合はその日付を使用（省略時は現在時刻）
    var now  = dateStr ? new Date(dateStr) : new Date();
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
    sortSummaryByDate_(companySsId);

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
  if (id) delaySyncSummary_(id);
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
  if (id) delaySyncSummary_(id);
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
  if (id) delaySyncSummary_(id);
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
  if (id) delaySyncSummary_(id);
}


// ================================================================
//  7-7: 状態保存＋時刻記録 一括実行（recordAction）  【大A / 中7 / 小7-7】
//  saveRunState と setXxxComplete を1回のサーバー呼び出しでまとめて実行する
//  actionType: 'guide' / 'pick' / 'restStart' / 'restEnd' / 'drop'
// ================================================================
function recordAction(actionType, id, routeIndex, stateObj, companySsId) {
  saveRunState(stateObj);
  if      (actionType === 'guide')     setGuideComplete(id, routeIndex, companySsId);
  else if (actionType === 'pick')      setPickComplete(id, routeIndex, companySsId);
  else if (actionType === 'restStart') setRest(id, routeIndex, 'start', companySsId);
  else if (actionType === 'restEnd')   setRest(id, routeIndex, 'end',   companySsId);
  else if (actionType === 'drop')      setDropComplete(id, routeIndex, companySsId);
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
  if (id) delaySyncSummary_(id);
}


// ================================================================
//  8-2: 運行シート行削除（deleteRunRows）  【大A / 中8 / 小8-2】
//  戻るボタン用：IDで全行を動的検索して降順に削除し集計表を同期する
// ================================================================
function deleteRunRows(id, companySsId) {
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
  if (id) delaySyncSummary_(id);
}


// ================================================================
//  8-3: 時刻セルクリア（clearTimeCell）  【大A / 中8 / 小8-3】
//  戻るボタン用：IDとルートインデックスで行を動的検索して指定列をクリアし集計表を同期する
// ================================================================
function clearTimeCell(id, routeIndex, col, companySsId) {
  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  var row = findRowByIdAndIndex_(sheet, id, routeIndex);
  if (row < 0) return;
  sheet.getRange(row, col).clearContent();
  if (id) delaySyncSummary_(id);
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
    if (String(mAll[j][10]).trim().toLowerCase() === savedEmail.toLowerCase()) {
      myName = String(mAll[j][8]).trim();
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
  var notes23  = lastRow >= 2 ? sheet.getRange(2, 24, lastRow-1, 1).getNotes() : [];
  var rtvs23   = lastRow >= 2 ? sheet.getRange(2, 24, lastRow-1, 1).getRichTextValues() : [];
  // 集計表からID単位の金額マップを作成（売上/高速/支払いは集計表の計算済み値を使う）
  var sumSheet = ss.getSheetByName('集計表');
  var payMap   = {};
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var sumAll = sumSheet.getRange(2, 1, sumSheet.getLastRow()-1, 34).getValues();
    for (var s = 0; s < sumAll.length; s++) {
      var sid = String(sumAll[s][0]||'').trim();
      if (sid) payMap[sid] = {
        sales:    Math.round(Number(sumAll[s][18])) || 0,
        tollReq:  Math.round(Number(sumAll[s][19])) || 0,
        tollReal: Math.round(Number(sumAll[s][20])) || 0,
        tollTotal:Math.round(Number(sumAll[s][21])) || 0,
        pay:      Math.round(Number(sumAll[s][26])) || 0,
        yukyu:    Math.round(Number(sumAll[s][32])) || 0,
        other:    Math.round(Number(sumAll[s][33])) || 0
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
        sales:0, tollReq:0, tollReal:0, tollTotal:0, pay:0, yukyu:0, other:0,
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
      isComplete: !!(g.guideTime && g.pickTime && g.restStart && g.restEnd && g.dropTime),
      isNew:      !g.guideTime && !g.pickTime && !g.restStart && !g.restEnd && !g.dropTime
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
function getEditData(id) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return null;

  var sumSheet = ss.getSheetByName('集計表');
  var sumData  = { tollTotal:'', profit:'', yukyu:'', other:'' };
  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var sumAll = sumSheet.getRange(2, 1, sumSheet.getLastRow()-1, 34).getValues();
    for (var s = 0; s < sumAll.length; s++) {
      if (String(sumAll[s][0]||'').trim() === String(id).trim()) {
        sumData.tollTotal = sumAll[s][21] !== '' ? sumAll[s][21] : '';
        sumData.profit    = sumAll[s][27] !== '' ? sumAll[s][27] : '';
        sumData.yukyu     = sumAll[s][32] !== '' ? Math.round(Number(sumAll[s][32])) : '';
        sumData.other     = sumAll[s][33] !== '' ? Math.round(Number(sumAll[s][33])) : '';
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
    profit:   sumData.profit,
    yukyu:    sumData.yukyu,
    other:    sumData.other
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
function saveEditData(obj) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var all = sheet.getDataRange().getValues();
  var written = false;
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0]||'').trim() !== String(obj.id).trim()) continue;
    var r = i + 1;

    // 車両情報（同一IDの全行に書き込む）
    if (obj.tons   !== undefined) sheet.getRange(r, 4).setValue(obj.tons   || '');
    if (obj.type   !== undefined) sheet.getRange(r, 5).setValue(obj.type   || '');
    if (obj.car    !== undefined) sheet.getRange(r, 6).setValue(obj.car    || '');
    if (obj.name   !== undefined) sheet.getRange(r, 7).setValue(obj.name   || '');
    if (obj.tel    !== undefined) sheet.getRange(r, 8).setValue(obj.tel    || '');
    if (obj.kanban !== undefined) sheet.getRange(r, 9).setValue(obj.kanban || '');

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
    if (obj.guideTime) { var c14=sheet.getRange(r,14); c14.setValue(new Date(obj.date+' '+obj.guideTime)); c14.setNumberFormat(timeFmt); }
    else               sheet.getRange(r, 14).clearContent();
    if (obj.pickTime)  { var c15=sheet.getRange(r,15); c15.setValue(new Date(obj.date+' '+obj.pickTime));  c15.setNumberFormat(timeFmt); }
    else               sheet.getRange(r, 15).clearContent();
    if (obj.restStart) { var c16=sheet.getRange(r,16); c16.setValue(new Date(obj.date+' '+obj.restStart)); c16.setNumberFormat(timeFmt); }
    else               sheet.getRange(r, 16).clearContent();
    if (obj.restEnd)   { var c17=sheet.getRange(r,17); c17.setValue(new Date(obj.date+' '+obj.restEnd));   c17.setNumberFormat(timeFmt); }
    else               sheet.getRange(r, 17).clearContent();
    if (obj.dropTime)  { var c18=sheet.getRange(r,18); c18.setValue(new Date(obj.date+' '+obj.dropTime));  c18.setNumberFormat(timeFmt); }
    else               sheet.getRange(r, 18).clearContent();

    // 売上・高速は最初の行のみ書き込む（複数行IDの場合の重複防止）
    // sales未指定（閲覧のみ）の場合は上書きしない
    if (!written) {
      if (obj.sales    !== undefined) sheet.getRange(r, 19).setValue(obj.sales    || '');
      if (obj.tollReq  !== undefined) sheet.getRange(r, 20).setValue(obj.tollReq  || '');
      if (obj.tollReal !== undefined) sheet.getRange(r, 21).setValue(obj.tollReal || '');
      written = true;
    } else {
      if (obj.tollReq  !== undefined) sheet.getRange(r, 20).setValue('');
      if (obj.tollReal !== undefined) sheet.getRange(r, 21).setValue('');
    }

    sheet.getRange(r, 25).setValue(obj.termNotice || '');
  }
  delaySyncSummary_(obj.id);
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
function setupSheetProtection() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var sumSheet = ss.getSheetByName('集計表');
  if (sumSheet) {
    sumSheet.setFrozenColumns(1);
    sumSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p){p.remove();});
    sumSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){p.remove();});
    var sp = sumSheet.protect().setDescription('集計表保護');
    var editableCols = [23, 25, 27, 29, 34]; // W=距離, Y=ガソリン代, AA=支払い, AC=備考, AH=その他手当
    sp.setUnprotectedRanges([
      sumSheet.getRange('W2:W2000'),
      sumSheet.getRange('Y2:Y2000'),
      sumSheet.getRange('AA2:AA2000'),
      sumSheet.getRange('AC2:AC2000'),
      sumSheet.getRange('AH2:AH2000')
    ]);

    // ヘッダー行: 保護列=グレー, 編集可列=グリーン で視覚区別
    var lastCol = Math.max(sumSheet.getLastColumn(), 34);
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
    unkouSheet.setFrozenColumns(1);
    unkouSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p){p.remove();});
    var unkouLastRow = Math.max(unkouSheet.getLastRow(), 2);
    // 保護列: V(22)=合計高速, Y(25)=連絡(端末), Z(26)=データ(端末) → グレー着色
    var protectedUnkouCols = [22, 25, 26];
    for (var pc = 0; pc < protectedUnkouCols.length; pc++) {
      var pcol = protectedUnkouCols[pc];
      unkouSheet.getRange(1, pcol).setBackground('#37474f').setFontColor('#90a4ae').setFontWeight('bold');
      unkouSheet.getRange(2, pcol, unkouLastRow - 1, 1).setBackground('#eceff1');
    }
  }

  ui.alert('保護設定完了\n■ 集計表\n  編集可: 距離(W)・ガソリン代(Y)・支払い(AA)・備考(AC)・その他手当(AH)\n  緑枠＋緑ヘッダー = 編集可 / 灰色ヘッダー = 保護\n■ 運行シート: V(合計高速)・Y(連絡端末)・Z(データ端末) をグレー着色');
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
//  9-1: 連絡事項保存（saveNotice）  【大A / 中9 / 小9-1】
//  指定IDの運行シートU列（21列目）にテキストを書き込む
// ================================================================
function saveNotice(id, text) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
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
  }
  return { ok: true };
}


// ================================================================
//  9-3: 端末からの連絡保存（saveTerminalNotice）  【大A / 中9 / 小9-3】
//  指定IDの運行シートW列（23列目）にテキストを書き込む
// ================================================================
function saveTerminalNotice(id, text) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
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
function getMyNotices() {
  var savedEmail = PropertiesService.getUserProperties().getProperty('linkedEmail');
  if (!savedEmail) return [];
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master) return [];
  var mAll   = master.getDataRange().getValues();
  var myName = '';
  for (var j = 1; j < mAll.length; j++) {
    if (String(mAll[j][10]).trim() === savedEmail) { myName = String(mAll[j][8]).trim(); break; }
  }
  if (!myName) return [];
  var readList = JSON.parse(PropertiesService.getUserProperties().getProperty('readNotices') || '[]');
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
//  10-2b: 行番号指定で連絡事項取得（getNoticeByRow）  【大A / 中10 / 小10-2b】
//  誘導画面に管理側の連絡事項・データURLを表示するために使う
// ================================================================
function getNoticeByRow(id, companySsId) {
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
//  10-4: 既読管理・既読一覧取得（getReadNotices）  【大A / 中10 / 小10-4】
//  PropertiesServiceから既読IDリストを取得して返す
// ================================================================
function getReadNotices() {
  return JSON.parse(PropertiesService.getUserProperties().getProperty('readNotices') || '[]');
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
  item('📷 写真・ファイル取込', '選択した行にPC上の写真・ファイルを直接添付します。', '#c8e6c9');
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
  item('給料・歩合の設定', 'M列=仮日数、N列=給料、O列=% を入力すると集計表に自動反映', '#c8e6c9');
  spacer();

  section('会社の追加セットアップ', '#2e7d32');
  item('新しい会社を追加', '「会社登録」シートにA列=会社名、B列=管理Gmail を入力\n→ 自動でアーカイブフォルダが作成され通知メールが届きます', '#c8e6c9');
  item('手動でまとめて実行', 'スクリプトエディタから「setupCompanies」を実行', '#c8e6c9');
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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

  // ── 乗務員向けメール（自車専属マスタ J列の全アドレスに個別送信）──
  var master = ss.getSheetByName('自車専属マスタ');
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
  var props = PropertiesService.getScriptProperties();
  var stored = props.getProperty('webAppUrl') || '';
  // script.google.com/macros で始まるURLのみ有効（Driveファイルや不正なURLを除外）
  if (stored && stored.indexOf('script.google.com/macros') !== -1) return stored;
  try {
    var svcUrl = ScriptApp.getService().getUrl();
    if (svcUrl && svcUrl.indexOf('script.google.com/macros') !== -1) {
      props.setProperty('webAppUrl', svcUrl);
      return svcUrl;
    }
  } catch(e) {}
  return '';
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


// ================================================================
//  12-3a: 会社配布用テンプレートSSを取得または作成（getOrCreateCompanyTemplate_）
//  マスターSSをコピーして全シートのデータ行を削除した「テンプレート」を1回だけ作成する。
//  以降は このテンプレートを makeCopy() して各会社SSを作成する。
//  テンプレートにはメニュー・全機能がそのまま引き継がれる。
// ================================================================
function getOrCreateCompanyTemplate_() {
  var props = PropertiesService.getScriptProperties();
  var templateId = props.getProperty('companyTemplateId');
  if (templateId) {
    try {
      var existingFile = DriveApp.getFileById(templateId);
      // マーカーシートがなければ古いテンプレートなので作り直す
      var existingSs = SpreadsheetApp.openById(templateId);
      if (existingSs.getSheetByName('__COMPANY_SS__') !== null) return existingFile;
      // 古いテンプレートを削除して再作成
      existingFile.setTrashed(true);
    } catch(e) {}
    props.deleteProperty('companyTemplateId');
  }

  // マスターSS取得
  var masterSs;
  try { masterSs = SpreadsheetApp.getActiveSpreadsheet(); } catch(ex) {}
  if (!masterSs) {
    var mid = props.getProperty('masterSsId');
    if (mid) masterSs = SpreadsheetApp.openById(mid);
  }
  if (!masterSs) throw new Error('マスターSSが見つかりません');

  // マスターをコピーしてテンプレート作成
  var masterFile   = DriveApp.getFileById(masterSs.getId());
  var templateFile = masterFile.makeCopy('【テンプレート】運行管理_配布用');
  var templateSs   = SpreadsheetApp.openById(templateFile.getId());

  // 各シートのデータ行のみ削除（ヘッダー行は残す）
  // 会社登録・使い方シートは不要なので削除
  var removeNames = ['会社登録', '使い方'];
  templateSs.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (removeNames.indexOf(name) !== -1) {
      if (templateSs.getSheets().length > 1) {
        try { templateSs.deleteSheet(sheet); } catch(e) {}
      }
    } else if (sheet.getLastRow() > 1) {
      sheet.deleteRows(2, sheet.getLastRow() - 1);
    }
  });

  // 客先SSであることを示す隠しマーカーシートを追加
  // このシートがあると onOpen() が客用メニューを表示する（マスターSSにはない）
  var marker = templateSs.insertSheet('__COMPANY_SS__');
  marker.hideSheet();

  props.setProperty('companyTemplateId', templateFile.getId());
  return templateFile;
}


// ================================================================
//  12-3: 会社専用スプレッドシートを作成（createCompanySpreadsheet_）  【大B / 中12 / 小12-3】
//  テンプレートSS（マスターのデータなしコピー）を makeCopy() して会社SSを作成する。
//  テンプレートにはメニュー・月生成・集計等の全機能が含まれる。
//  targetFolderId が指定されたそのフォルダに移動する（未指定なら「運行管理_会社別」）。
//  戻り値: { ssId, ssUrl }
// ================================================================
function createCompanySpreadsheet_(companyName, adminEmail, targetFolderId) {
  // テンプレートSSを取得（なければ自動作成）
  var templateFile = getOrCreateCompanyTemplate_();

  // テンプレートをコピーして会社SS作成
  var newFile = templateFile.makeCopy(companyName + ' 運行管理');

  // 管理Gmailに編集権限を付与
  try { newFile.addEditor(adminEmail); } catch(e) {}

  // 指定フォルダ（または「運行管理_会社別」）に移動
  var destFolder = targetFolderId
    ? DriveApp.getFolderById(targetFolderId)
    : getOrCreateFolder_('運行管理_会社別');
  destFolder.addFile(newFile);
  try { DriveApp.getRootFolder().removeFile(newFile); } catch(e) {}

  var newSs = SpreadsheetApp.openById(newFile.getId());
  return { ssId: newSs.getId(), ssUrl: newSs.getUrl() };
}


// ================================================================
//  12-4: 新規会社フルセットアップ（processNewCompany_）  【大B / 中12 / 小12-4】
//  ① 共有フォルダ（運行管理_アーカイブ/会社名/）を作成して管理Gmailに共有
//  ② そのフォルダの中にコードなしSSを作成（開けばスプレッドシートが見える）
//  ③ アプリURL（WebアプリURL?ssId=会社SS_ID）を生成
//  ④ 会社登録シートに各URLを記録
//  ⑤ 管理者に配布メール1通だけ送信（フォルダ通知メールは送らない）
//  メニュー「📤 会社SS作成＆メール送信」または時間トリガー（processPendingCompanies_）から呼ぶ。
// ================================================================
function processNewCompany_(companyName, adminEmail) {
  // マスターSS取得（メニュー実行 or 時間トリガーの両方に対応）
  var ss;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(ex) {}
  if (!ss) {
    var mid = PropertiesService.getScriptProperties().getProperty('masterSsId');
    if (mid) ss = SpreadsheetApp.openById(mid);
  }
  if (!ss) throw new Error('マスターSSが見つかりません');

  var regSheet = ss.getSheetByName('会社登録');

  // ① 共有フォルダを作成（メール送信なし）
  //    フォルダパス: 運行管理_アーカイブ/会社名/
  var folderResult = setupOneCompany_(companyName, adminEmail, true); // suppressEmail=true
  var folderUrl = folderResult.folderUrl;
  var folderId  = folderResult.folderId;

  // ② コードなしSSをその共有フォルダ内に作成
  //    → 相手が共有フォルダを開くとスプレッドシートが見える
  var ssResult = createCompanySpreadsheet_(companyName, adminEmail, folderId);
  var ssUrl    = ssResult.ssUrl;
  var ssId     = ssResult.ssId;

  // ③ アプリURL = ベースURL?ssId=会社のSS_ID
  var webAppUrl = getWebAppBaseUrl_();
  var appUrl    = webAppUrl ? (webAppUrl + '?ssId=' + ssId) : '[WebアプリURL未設定]';

  // ④ 会社登録シートにSSURL・AppURL・フォルダURLを記録
  var targetRow = -1;
  if (regSheet && regSheet.getLastRow() >= 2) {
    var rows = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === companyName) { targetRow = i + 2; break; }
    }
  }
  if (targetRow > 0) {
    regSheet.getRange(targetRow, 3).setValue('済').setBackground('#c8e6c9').clearNote();
    regSheet.getRange(targetRow, 4).setValue(new Date());
    regSheet.getRange(targetRow, 5).setValue(folderUrl);
    regSheet.getRange(targetRow, 6).setValue(ssUrl);
    regSheet.getRange(targetRow, 7).setValue(appUrl);
  }

  // ⑤ 配布メール1通だけ送信（管理者向け + 乗務員向け）
  sendDistributionMail_(
    companyName, adminEmail, ssUrl, appUrl,
    targetRow > 0 ? targetRow : null,
    targetRow > 0 ? regSheet : null
  );

  return { ssId: ssId, ssUrl: ssUrl, appUrl: appUrl, folderUrl: folderUrl };
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

    if (sheetName !== '会社登録' || row <= 1) return;

    // ── A列 or B列: 会社名+Gmail が揃ったらフルセットアップ ──────────
    if (col === 1 || col === 2) {
      var companyName = String(sheet.getRange(row, 1).getValue() || '').trim();
      var adminEmail  = String(sheet.getRange(row, 2).getValue() || '').trim();
      var status      = String(sheet.getRange(row, 3).getValue() || '').trim();

      if (!companyName || !adminEmail || adminEmail.indexOf('@') === -1) return;
      if (status === '済' || status.indexOf('エラー') === 0) return;

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