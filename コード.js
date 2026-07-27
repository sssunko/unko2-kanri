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
//            ・?action=agree → 同意処理を実行してお礼ページを返す（google.script.run不使用）
//            ・?page=contract → 利用規約・同意フォームページを返す
//            ・?page=parent → 管理画面ページを返す
//            ・それ以外 → index.htmlをテンプレートとして返しWebアプリを表示する
//   2-3  : showSidebar()
//            スプレッドシート右側のサイドバーとしてWebアプリを表示する
//            ・スプレッドシート上で「ホーム画面を表示」メニューを選んだ時に実行
//   2-4  : showUploadSidebar()
//            運行シートの選択行のW列（管理側データ列）にファイルを直接アップロードするサイドバー
//            ・「📷 写真・ファイル取込」メニューから起動
//            ・GoogleフォトのURLは直接使えないためPCからダウンロードして使うよう案内
//   2-5  : showDispatchDashboard()
//            メニュー「配車ダッシュボード」からサイドバーで当日の配車状況（積地あり/空欄）を表示
//   2-2b : storeCompanySsId(ssId)
//            ssIdをUserPropertiesに保存（doGet後に端末が自分のSSを識別するために使う）
//   2-2c : validateDriverEmail_(email, companySsId)
//            自車専属マスタJ列のメールアドレスを検証して部外者アクセスを遮断する内部認証関数
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
//   3-3b : syncVehicleToCurrentMonth_(veh, skipSort, applyDate, ss)
//            B列（運行状態）変更時：適用日以降の空行を削除→「運行」なら当月分の行を再生成
//   3-3c : syncAllVehiclesToCurrentMonth_()
//            expandAndRefreshSheetsから呼ばれる全車両一括版（3-3bをループで呼ぶ）
//   3-4  : onEditMasterCustomer_(sheet, range)
//            取引先マスタ（マスタシート）編集時の処理
//            ・A列（ID）自動生成 → M-XXXX 形式
//   3-5a : registerJohoRowToUnkou_(johoSheet, rowNum, confirmedCol, ss)
//            配車板のB列またはO列が「確定」になった時に運行シートへ行を即時登録
//   3-5b : refreshJohoColors_(ss)
//            配車板全行の進捗色を一括再適用（コピペ後・シート再生成後の復元用）
//   3-5c : afterSaveJoho(ssId, rowNum, col1Based) / afterSaveJohoFull(ssId, rowNum)
//            親アプリのインライン保存後に呼ばれ onEditJoho_ 相当の処理をSS側で実行
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
//   4-2b : fillMissingIdsAndCars()
//            メニュー「ID・車番一括補完」: IDや車番が空の行を一括補完し集計表を再同期する
//   4-2c : showDateTimePicker() / setDateTimeToActiveCell_()
//            メニュー「日時入力」: 選択セルに年月日+時刻を入力するカスタムダイアログ
//   4-3b : autoFillExpense()
//            メニュー「経費自動入力」: トン数別平均値テーブルから集計表の経費を一括自動入力
//   4-8d : getCompanyName_(ss)
//            自車専属マスタのC列(会社名)からこのスプレッドシートの会社名を返す内部補助関数
//   4-9  : calcDistanceManual()
//            メニュー「距離計算（未計算分）」: 積地・降地あり・距離未設定の行をMaps APIで一括計算
//   4-10 : resolveAmbiguousAddresses()
//            メニュー「住所確認（確認待ち分）」: 複数候補ルートを1件ずつダイアログで選択し距離マスタに確定登録
//   4-11 : generateAuditSheet()
//            メニュー「監査用表生成」: 集計表から改善基準告示コンプライアンス確認表（監査用シート）を生成
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
//   8-6a : saveTermNoticeByDriver(id, termNotice, companySsId)
//            端末アプリの一覧編集モーダルから連絡(端末)(Y列=25)のみを書き込む
//   8-6b : setupSheetProtection()
//            集計表を全体保護（W・Y・AA・AD・AI列のデータ行のみ編集可）し運行シートも保護する
//            ・集計表：スプレッドシートレベル保護＋編集可列をunprotectedRangesで指定
//            ・運行シート：Y列(連絡端末)・Z列(データ端末)をlockして直接入力を防止
//   8-6b-0a: insertKanbanColumn()
//            看板名列(I列=col9)が未存在の場合に運行シート・集計表に列を挿入し
//            既存行は会社名(C列)をデフォルト値として埋める（メニューから1回のみ実行）
//   8-6b-1: getTerminalUrls_(sheet, rowNum)
//            Y列(25)のリッチテキストからリンクURL一覧を配列で取得して返す
//   8-6b-2: setTerminalUrls_(sheet, rowNum, urls)
//            複数URLをY列(25)にリッチテキストリンク「ファイル1  ファイル2...」として書き込む
//            ・URLをセルノートにも保存（getNotes()で一括読み取り可能にするため）
//   8-6b-2b: importImageToDrive_(url)
//            外部URL（GoogleフォトなどのURL）の画像をOAuthトークンで取得しDriveにコピー
//            ・Googleフォトのプライベート URL は不可（サイドバーアップロード推奨）
//            ※ ファイル上はb-1/b-2の後に配置されているが補助関数として番号2bとする
//   8-6b-3: setAdminDataRichText_(sheet, rowNum, url)
//            1件のURLをW列(23)にリッチテキストリンク「ファイル1」として書き込む
//   8-6b-3b: setAdminDataRichTextMulti_(sheet, rowNum, urls)
//            複数URLをW列(23)にリッチテキストリンク「ファイル1  ファイル2...」として書き込む
//            ・URLをセルノートにも保存（getNotes()で一括読み取り可能にするため）
//   8-6b-4: getAdminDataUrl_(sheet, rowNum)
//            W列(23)のリッチテキストからURLをカンマ区切り文字列で返す（プレーン値フォールバックあり）
//   8-6c : appendTerminalFile(id, fileName, base64Data, mimeType)
//            Base64データをファイル化してDriveに保存しY列(25)のリッチテキストURLに追記する
//            ・「端末データ」フォルダに保存・誰でも閲覧可能リンクを設定
//   8-6c-2: appendAdminFileById / deleteAdminFileById / replaceAdminFileById
//            管理側（W列=23）ファイルの追加・削除・差し替えをIDで指定して実行する
//   8-6d : deleteTerminalFile(id, fileUrl, companySsId)
//            端末データ（Y列=25）から指定URLのファイルを削除しDriveからも削除する
//   8-6e : replaceTerminalFile(id, oldUrl, fileName, base64Data, mimeType, companySsId)
//            端末データ（Y列=25）の既存URLを新ファイルで差し替える
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
//   中2  起動・メニュー             2-1〜2-5,2-2b/c   onOpen, doGet, showSidebar, showUploadSidebar, showDispatchDashboard など
//   中3  スプレッドシート自動処理   3-1〜3-5c  onEdit, onEditUnkou_, onEditMasterVehicle_, syncVehicleToCurrentMonth_,
//                                             afterSaveJoho など
//   中4  集計表・シート操作         4-1〜4-11  generateSummary, syncSummaryForId_, fillMissingIdsAndCars,
//                                             calcDistanceManual, resolveAmbiguousAddresses, generateAuditSheet など
//   中5  アプリ初期化・紐づけ       5-1〜5-3   getInitialData, linkAddress など
//   中6  端末 運行進捗管理          6-1〜6-3   saveRunState, loadRunState など
//   中7  端末 運行操作              7-1〜7-6   getTodayRoutes, createParentRows など
//   中8  端末 一覧・編集・ファイル  8-1〜8-7   getListData, saveEditData など
//   中9  端末 連絡・ファイル        9-1〜9-4   saveNotice, uploadFileToRow など
//   中10 端末 既読管理              10-1〜10-4 getMyNotices, markAsRead など
//   中11 会社セットアップ・配布     11-1〜11-7 onEditCompanyRegister_, setupOneCompany_, setupCompanies,
//                                             createUsageSheet, sendDistributionMail_, triggerDistributionMail,
//                                             createManualSheet(11-7)
//   中12 会社専用SS作成・管理       12-1〜12-10 getWebAppBaseUrl_, setWebAppUrl, getNewSsScriptId_(12-3a),
//                                             deployClientWebApp_(12-3a-2), createCompanySpreadsheet_(12-3),
//                                             initClientSSSheets_(12-3b), syncToTemplateSS(12-3c),
//                                             createLibraryVersion_(12-3c-1), updateStubVersion_(12-3c-2),
//                                             processNewCompany_(12-4), agreeContract(12-4b),
//                                             sendCompanySetupEmails(12-5), createSignupForm(12-6),
//                                             onFormSubmit_(12-7), processPendingCompanies_(12-8),
//                                             createDevSs(12-9), showMyScriptId(12-10)
//   中13 CSVインポート              13-0〜13-15 showCsvImportDialog_, importBulkRows, buildSheetRow_,
//                                             getImportDictionary, saveImportAliases,
//                                             showEtcImportDialog, prepareEtcImport, executeEtcImport など
//   中14 トリガー・反映・帳票       14-1〜14-5 installTriggers(14-1), ensureInstalledTrigger_(14-1a),
//                                             installedOnEdit_(14-2), dispatchInstalledEdit(14-2b),
//                                             showHatchuDocDialog(14-2c), showShabanDocDialog(14-2d),
//                                             getDocumentData_(14-2e), syncToAllClientSS(14-3),
//                                             markDocumentIssued(14-4), sendDocumentEmail(14-5)
//   中15 配車確定                   15-1〜15-2 matchAndConfirmDispatch, cancelDispatch, buildJohoNewRow_
//   中16 受領書・請求書・支払確認書  16-1〜16-7 showUketorishoDialog(16-1), generateUketorishoSheet(16-2),
//                                             showInvoiceDialog(16-3), showPaymentDialog(16-4),
//                                             getNextDocNum_(16-5), generateInvoiceSheet(16-6),
//                                             generatePaymentSheet(16-7)
//   中17 PL管理                     17-1〜17-7b showPlDialog, getPlFilterOptions, generatePl,
//                                             buildPlBreakdown_, buildPlSheet_, getFixedCosts_,
//                                             apportionFixedCosts_, exportPlJournalCsv, exportPlBundle,
//                                             exportSheetAsCsvBase64, exportSelectedSheetsAsExcel,
//                                             showExportDialog, initFixedCostMaster,
//                                             updatePlApportionment_(17-7a), setRecalcChoice_(17-7b)
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
  var maxNum = 0;
  if (lastRow >= 2) {
    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var k = 0; k < ids.length; k++) {
      var match = String(ids[k][0]).match(/(\d+)$/);
      if (match) {
        var n = parseInt(match[1], 10);
        if (n > maxNum) maxNum = n;
      }
    }
  }
  // 完全通し番号：発行済み最大番号をSS内の開発者メタデータに記録し、
  // 行削除・アーカイブ後も一度使った番号は二度と再利用しない
  var stored = 0;
  try {
    var mdKey = 'MAX_ID_NUM_' + String(prefix || '').replace(/-$/, '');
    var md = sheet.getParent().createDeveloperMetadataFinder().withKey(mdKey).find();
    if (md.length > 0) stored = parseInt(md[0].getValue(), 10) || 0;
  } catch (e) {}
  var next = Math.max(maxNum, stored) + 1;
  commitLastId_(sheet, prefix, next);
  return next;
}

// 発行済み最大番号の帳簿更新（現在の記録より大きい時だけ書き込む・失敗しても採番は継続）
function commitLastId_(sheet, prefix, lastUsedNum) {
  if (!lastUsedNum || lastUsedNum <= 0) return;
  try {
    var key = 'MAX_ID_NUM_' + String(prefix || '').replace(/-$/, '');
    var ss = sheet.getParent();
    var md = ss.createDeveloperMetadataFinder().withKey(key).find();
    if (md.length > 0) {
      if ((parseInt(md[0].getValue(), 10) || 0) < lastUsedNum) md[0].setValue(String(lastUsedNum));
    } else {
      ss.addDeveloperMetadata(key, String(lastUsedNum));
    }
  } catch (e) {}
}

// トン数を「4t」形式に正規化。全角数字・大文字T対応。数字のみ入力→t付加。
function normalizeTons_(val) {
  var s = String(val || '').trim()
    .replace(/[０-９]/g, function(c){return String.fromCharCode(c.charCodeAt(0)-0xFEE0);})
    .replace(/[ｔＴ]/g, 't');
  if (!s) return val;
  var m = s.match(/^([\d.]+)(?:[tT]|トン|ｔｏｎ|ton)?$/);
  if (m) return m[1] + 't';
  return val;
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
function delaySyncSummary_(id, ss) { try { syncSummaryForId_(id, ss); } catch(e) { logError_('delaySyncSummary_', e); } }

// エラーログをシート "_ErrorLog_" に追記（シートがなければ自動作成）
function logError_(context, e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('_ErrorLog_');
    if (!sh) {
      sh = ss.insertSheet('_ErrorLog_');
      sh.getRange(1, 1, 1, 3).setValues([['日時', '関数名', 'エラーメッセージ']]);
      sh.setFrozenRows(1);
      sh.hideSheet();
    }
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    sh.appendRow([now, String(context), String(e && e.message ? e.message : e)]);
  } catch(e2) {}
}


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
function applyHolidayRowColors_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
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
      // 通常行: 既存背景を保護し、有休/休み色と保護色を一旦クリア（残骸・IDなし行の保護色もF5で消える）
      var rowArr = curBgs[idx].slice();
      for (var ci = 0; ci < rowArr.length; ci++) {
        if (rowArr[ci] === '#e0e0e0' || rowArr[ci] === '#9e9e9e' || rowArr[ci] === '#eceff1') rowArr[ci] = null;
      }
      if (pickV === '' && idV !== '') {
        rowArr[11] = '#fff9c4'; // IDがあって積地空→常に黄色
      } else if (pickV !== '' && rowArr[11] === '#fff9c4') {
        rowArr[11] = null; // 積地入力あり→黄色解除
      }
      // 保護列（V=22, Y=25, Z=26）の保護色をIDあり行に自動復元（F5だけで直る）
      if (idV !== '') {
        if (lastCol >= 22) rowArr[21] = '#eceff1';
        if (lastCol >= 25) rowArr[24] = '#eceff1';
        if (lastCol >= 26) rowArr[25] = '#eceff1';
      }
      return rowArr;
    });
    sheet.getRange(2, 1, lr - 1, lastCol).setBackgrounds(bgs2D);
    // データ最終行より下の残骸色（過去バージョンの塗り跡）もF5で自動除去
    var maxR = sheet.getMaxRows();
    if (maxR > lr) {
      var tailBgs = sheet.getRange(lr + 1, 1, maxR - lr, lastCol).getBackgrounds();
      var tailDirty = false;
      for (var tr = 0; tr < tailBgs.length; tr++) {
        for (var tc = 0; tc < tailBgs[tr].length; tc++) {
          var tv = tailBgs[tr][tc];
          if (tv === '#eceff1' || tv === '#e0e0e0' || tv === '#9e9e9e' || tv === '#fff9c4') {
            tailBgs[tr][tc] = null;
            tailDirty = true;
          }
        }
      }
      if (tailDirty) sheet.getRange(lr + 1, 1, maxR - lr, lastCol).setBackgrounds(tailBgs);
    }
    sheet.getRange(1, 12).setNote('🟡 薄黄: 配車漏れ（IDあり積地空）');
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

  // 自車専属マスタ: B列ヘッダーに行色の説明ノートを設定
  var masterSheet = ss.getSheetByName('自車専属マスタ');
  if (masterSheet) {
    masterSheet.getRange(1, 2).setNote('🔴 淡い赤: 運行\n🟢 淡い緑: 故障\n🟡 薄黄: 待機');
  }

  applyExpiryWarningColors_(ss);
  markIdCollisions_(ss);
}
function applyHolidayRowColors() { applyHolidayRowColors_(); }


// ================================================================
//  1-7b: IDが重複している行のA列を赤色でマーク（markIdCollisions_）  【大B / 中1 / 小1-7b】
//  運行シートで「同じID、異なる車番 or 異なる日付」の行を検出して赤く着色
//  正常時（同ID・同車番・同日付）はマークしない
// ================================================================
function markIdCollisions_(ss) {
  var sheet = (ss || SpreadsheetApp.getActiveSpreadsheet()).getSheetByName('運行');
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
//  1-7c: 期限アラート色付け（applyExpiryWarningColors_）  【大B / 中1 / 小1-7c】
//  自車専属マスタの4期限列（免許証・安全教育・健康診断・適性診断）を参照し
//  運行シートのA列（ID）を乗務員ごとの最も近い期限に応じて着色する。
//  淡い赤=期限超過 / 淡い青=当日 / 淡い緑=7日以内
//  有休・休み行(グレー)・ID衝突行(#ff1744)には上書きしない。
// ================================================================
function applyExpiryWarningColors_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet       = ss.getSheetByName('運行');
  var masterSheet = ss.getSheetByName('自車専属マスタ');
  if (!sheet || !masterSheet || sheet.getLastRow() < 2 || masterSheet.getLastRow() < 2) return;

  var mHeaders = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
  var nameColM = mHeaders.indexOf('乗務員名');
  var expiCols = [
    mHeaders.indexOf('免許証有効期限'),
    mHeaders.indexOf('安全教育次回予定日'),
    mHeaders.indexOf('健康診断次回予定日'),
    mHeaders.indexOf('適性診断次回予定日')
  ].filter(function(c) { return c >= 0; });
  if (nameColM < 0 || expiCols.length === 0) return;

  var today = new Date(); today.setHours(0, 0, 0, 0);
  var mData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, masterSheet.getLastColumn()).getValues();

  var expiryMap = {};
  for (var mi = 0; mi < mData.length; mi++) {
    var mRow = mData[mi];
    var dName = String(mRow[nameColM] || '').trim();
    if (!dName) continue;
    var minDays = null;
    for (var ci = 0; ci < expiCols.length; ci++) {
      var dt = mRow[expiCols[ci]];
      if (!(dt instanceof Date) || isNaN(dt.getTime())) continue;
      var exp = new Date(dt); exp.setHours(0, 0, 0, 0);
      var days = Math.round((exp.getTime() - today.getTime()) / 86400000);
      if (minDays === null || days < minDays) minDays = days;
    }
    if (minDays !== null) expiryMap[dName] = minDays;
  }
  if (Object.keys(expiryMap).length === 0) return;

  var lr = sheet.getLastRow() - 1;
  var aVals  = sheet.getRange(2, 1, lr, 1).getValues();
  var gVals  = sheet.getRange(2, 7, lr, 1).getValues();
  var curBgs = sheet.getRange(2, 1, lr, 1).getBackgrounds();
  var skipSet = { '#e0e0e0': 1, '#9e9e9e': 1, '#ff1744': 1 };

  var newBgs = [], changed = false;
  for (var i = 0; i < lr; i++) {
    var id  = String(aVals[i][0] || '').trim();
    var drv = String(gVals[i][0] || '').trim();
    var cur = curBgs[i][0];
    if (!id || skipSet[cur] || !expiryMap.hasOwnProperty(drv)) {
      newBgs.push([cur]); continue;
    }
    var d = expiryMap[drv];
    var nc = d < 0 ? '#ffcdd2' : d === 0 ? '#bbdefb' : d <= 7 ? '#c8e6c9' : cur;
    if (nc !== cur) changed = true;
    newBgs.push([nc]);
  }
  if (changed) sheet.getRange(2, 1, lr, 1).setBackgrounds(newBgs);

  // A1セルに色凡例をメモとして設定（常に最新状態に保つ）
  sheet.getRange(1, 1).setNote(
    '■ 背景色の意味\n' +
    '〔ID列（A列）〕\n' +
    '🔴 濃い赤: ID衝突（同IDで車番/日付が不一致）\n' +
    '🔴 淡い赤: 資格期限切れ\n' +
    '🔵 淡い青: 本日が期限\n' +
    '🟢 淡い緑: 7日以内（免許/教育/健診/適性）\n' +
    '〔行全体〕\n' +
    '⚪ 薄グレー: 有休行\n' +
    '⚫ 濃いグレー: 休み行'
  );
}
function applyExpiryWarningColors() { applyExpiryWarningColors_(); }


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
    var da = (a.row[9] instanceof Date) ? a.row[9].getTime() : Infinity;
    var db = (b.row[9] instanceof Date) ? b.row[9].getTime() : Infinity;
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
    formulas.push(['=IF(T' + rn + '=U' + rn + ',"",T' + rn + '-U' + rn + ')']);
  }
  sheet.getRange(2, 22, numRows, 1).setFormulas(formulas);

  // 書式を再適用
  sheet.getRange(2, 10, numRows, 1).setNumberFormat('yyyy/MM/dd');
  sheet.getRange(2, 12, numRows, 2).setNumberFormat('@');
  applyMoneyFormat_(sheet, 2, numRows, 'unkou');
  applyDateTimeFormat_(sheet, 2, numRows);
  // ソート後、ID（A列）・日付（J列）が共に空の完全空白行を物理削除
  var blankRows = [];
  for (var bi = 0; bi < writeData.length; bi++) {
    var bId = String(writeData[bi][0] || '').trim();
    var bDate = writeData[bi][9];
    if (!bId && !(bDate instanceof Date) && !String(bDate || '').trim()) blankRows.push(bi + 2);
  }
  if (blankRows.length > 0) deleteRowsGrouped_(sheet, blankRows);
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
//  ①修正用SS専用メニュー（全機能フル搭載）。②客用SS・③各客SSはスタブの onOpen で表示。
//  GASはサブメニューの2階層が上限のため、大カテゴリ＋直接項目の構成で実装。
//  大カテゴリ: 🏠毎日の配車業務 / 📥データ読み込み / 📨帳票・送信 / 📒経理・出力
//             / 📊PL管理 / 🗓月またぎ処理 / ⚙️システム設定・保守 / 🏢管理者専用
// ================================================================
function onOpen(e) {
  // サイレント自動トリガー再構築（FULL権限時のみ有効・LIMITED時はtry-catchで自動スキップ）
  try {
    var _tgSs0 = SpreadsheetApp.getActiveSpreadsheet();
    var _sf = ['installedOnEdit_','onStructureChange_','checkMasterExpiries','onOpen','checkExpiryDates','calcDistanceTrigger_'];
    ScriptApp.getUserTriggers(_tgSs0).forEach(function(t) {
      if (_sf.indexOf(t.getHandlerFunction()) !== -1) { try { ScriptApp.deleteTrigger(t); } catch(ex) {} }
    });
    ScriptApp.newTrigger('installedOnEdit_').forSpreadsheet(_tgSs0).onEdit().create();
    ScriptApp.newTrigger('onStructureChange_').forSpreadsheet(_tgSs0).onChange().create();
    ScriptApp.newTrigger('calcDistanceTrigger_').timeBased().atHour(0).everyDays(1).create();
  } catch(_ex0) {}
  var ui   = SpreadsheetApp.getUi();
  var menu = ui.createMenu('メニュー');
  menu
    // ── 🏠 毎日の配車業務（配車・日次作業） ──────────────────────────
    .addSubMenu(ui.createMenu('🏠 毎日の配車業務')
      .addItem('🏠 ホーム画面を表示',           'showSidebar')
      .addItem('🚚 配車ダッシュボード',          'showDispatchDashboard')
      .addItem('🔗 チェックした行を配車確定',    'matchAndConfirmDispatch')
      .addItem('🔓 選択行のマッチング解除',      'cancelDispatch')
      .addItem('🔧 配車板 列ズレ修復',          'repairJohoSheet')
      .addItem('🆔 ID・車番一括補完',            'fillMissingIdsAndCars')
      .addItem('📏 距離計算（未計算分）',         'calcDistanceManual')
      .addItem('📍 住所確認（確認待ち分）',       'resolveAmbiguousAddresses')
      .addItem('🔃 日付順並び替え',             'sortBothSheetsByDate')
      .addItem('💴 経費自動入力',               'autoFillExpense')
      .addItem('📷 写真・ファイル取込',          'showUploadSidebar'))
    // ── 📥 データ読み込み（CSV） ──────────────────────────────────────
    .addSubMenu(ui.createMenu('📥 データ読み込み（CSV）')
      .addItem('🚛 運行シート',                 'showCsvImportDialogUnkou')
      .addItem('🗄 自車専属マスタ',              'showCsvImportDialogMaster')
      .addItem('📇 マスタ（取引先）',             'showCsvImportDialogCust')
      .addSeparator()
      .addItem('⛽ ETC利用明細',                'showEtcImportDialog')
      .addSeparator()
      .addItem('🗑 空インポート行を削除',         'deleteBlankImportRows'))
    // ── 📨 帳票・送信メニュー ──────────────────────────────────────────
    .addSubMenu(ui.createMenu('📨 帳票・送信メニュー')
      .addItem('① 発注書・指示書を作成（協力会社・乗務員用）', 'showHatchuDocDialog')
      .addItem('② 車番連絡を作成（荷主用）',       'showShabanDocDialog')
      .addSeparator()
      .addItem('🗒 受領書の耳生成',              'showUketorishoDialog'))
    // ── 📒 経理・出力 ──────────────────────────────────────────────────
    .addSubMenu(ui.createMenu('📒 経理・出力')
      .addItem('🧾 請求書生成',                 'showInvoiceDialog')
      .addItem('💳 支払確認書生成',              'showPaymentDialog')
      .addItem('📋 監査用表生成',               'generateAuditSheet')
      .addItem('💾 CSV・Excel出力',             'showExportDialog'))
    // ── 📊 PL管理 ──────────────────────────────────────────────────────
    .addSubMenu(ui.createMenu('📊 PL管理')
      .addItem('📈 PL作成',                    'showPlDialog')
      .addItem('🗃 PL設定初期化',               'initFixedCostMaster'))
    // ── 🗓 月またぎ処理 ────────────────────────────────────────────────
    .addSubMenu(ui.createMenu('🗓 月またぎ処理')
      .addItem('📆 今月分生成（途中契約）',       'generateCurrentMonth')
      .addItem('📅 翌月分生成（前月アーカイブ）',  'generateNextMonth')
      .addItem('📦 前月分アーカイブ',            'archiveOldMonth'))
    // ── ⚙️ システム設定・保守 ──────────────────────────────────────────
    .addSubMenu(ui.createMenu('⚙️ システム設定・保守')
      .addItem('🔧 初期設定',                   'installTriggers')
      .addItem('🔄 メニュー再生成（メニューが消えたら押す）', 'reloadMenu')
      .addItem('🗂 シート再生成',               'expandAndRefreshSheets')
      .addItem('🔑 ヘッダー復旧＋保護',          'restoreHeaders')
      .addItem('📃 集計表再生成',               'generateSummary')
      .addItem('🛡 シート保護設定',             'setupSheetProtection')
      .addItem('🔓 保護を全解除',               'removeAllProtections')
      .addItem('📖 使い方シート作成',            'createUsageSheet')
      .addItem('📘 説明書作成',                 'createManualSheet')
      .addItem('📋 サポートテンプレ作成',        'createSupportSheet')
      .addItem('🗾 距離マスタ 主要地データ投入',  'initDistanceMasterMajorCities'))
    // ── 🏢 管理者・セットアップ（開発専用） ────────────────────────────
    .addSubMenu(ui.createMenu('🏢 管理者・セットアップ（開発専用）')
      .addItem('🏗 会社セットアップ実行',         'setupCompanies')
      .addItem('📬 会社SS作成＆メール送信',       'sendCompanySetupEmails')
      .addItem('📧 配布メール送信',              'triggerDistributionMail')
      .addItem('📝 申し込みフォーム作成',         'createSignupForm')
      .addSeparator()
      .addItem('📡 テスト客SS（②）に反映',       'syncToTemplateSS')
      .addItem('🌐 全客SS（③）に反映',           'syncToAllClientSS'))
    // ── 🔄 バックアップ・復旧 ────────────────────────────────────────
    .addSubMenu(ui.createMenu('🔄 バックアップ・復旧')
      .addItem('🔄 バックアップから復旧（会社選択）', 'openRestoreDialog')
      .addItem('⏱ バックアップタイマー設定（毎日3時）', 'setupBackupTrigger')
      .addItem('📅 月次自動生成タイマー設定（毎月20日）', 'setupMonthlyTrigger'));
  menu.addToUi();

  try {
    PropertiesService.getScriptProperties().setProperty(
      'masterSsId', SpreadsheetApp.getActiveSpreadsheet().getId()
    );
  } catch(ex) {}
  try { convertLegacyAdminDataUrls_(); } catch(ex) {}
  try { applyHolidayRowColors_(); } catch(ex) {}
  try {
    var _hideSs = SpreadsheetApp.getActiveSpreadsheet();
    ['指示先履歴', '指示先ID別'].forEach(function(n) {
      var sh = _hideSs.getSheetByName(n);
      if (sh && !sh.isSheetHidden()) sh.hideSheet();
    });
  } catch(ex) {}
  try { showExpiryAlert(); } catch(ex) {}
  try { applyExpiryWarningColors_(); } catch(ex) {}
  try {
    var _bkProps = PropertiesService.getDocumentProperties();
    var _bkLast  = Number(_bkProps.getProperty('LAST_BACKUP_TS') || 0);
    if (Date.now() - _bkLast > 24 * 60 * 60 * 1000) {
      backupAllSheets_(SpreadsheetApp.getActiveSpreadsheet());
      _bkProps.setProperty('LAST_BACKUP_TS', String(Date.now()));
    }
  } catch(ex) {}
  try {
    var _errSs = SpreadsheetApp.getActiveSpreadsheet();
    var _errSh = _errSs.getSheetByName('_ErrorLog_');
    if (_errSh) {
      var _a1 = String(_errSh.getRange(1, 1).getValue());
      if (_a1.indexOf('⚠️ 要確認') === 0) {
        SpreadsheetApp.getUi().alert(_a1);
        _errSh.getRange(1, 1).setValue('日時');
      }
    }
  } catch(ex) {}
}


// ================================================================
//  客SS・テンプレートSS用メニュー構築（buildClientMenu）
//  スタブの onOpen から呼ばれる。メニュー定義をライブラリに置くことで
//  「各客に反映」でライブラリバージョンを更新するだけでメニューが追随する。
// ================================================================
function buildClientMenu() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('メニュー');
  menu
    // ── 🏠 毎日の配車業務（配車・日次作業） ──────────────────────────
    .addSubMenu(ui.createMenu('🏠 毎日の配車業務')
      .addItem('🏠 ホーム画面を表示',           'showSidebar')
      .addItem('🚚 配車ダッシュボード',          'showDispatchDashboard')
      .addItem('🔗 チェックした行を配車確定',    'matchAndConfirmDispatch')
      .addItem('🔓 選択行のマッチング解除',      'cancelDispatch')
      .addItem('🔧 配車板 列ズレ修復',          'repairJohoSheet')
      .addItem('🆔 ID・車番一括補完',            'fillMissingIdsAndCars')
      .addItem('📏 距離計算（未計算分）',         'calcDistanceManual')
      .addItem('📍 住所確認（確認待ち分）',       'resolveAmbiguousAddresses')
      .addItem('🔃 日付順並び替え',             'sortBothSheetsByDate')
      .addItem('💴 経費自動入力',               'autoFillExpense')
      .addItem('📷 写真・ファイル取込',          'showUploadSidebar'))
    // ── 📥 データ読み込み（CSV） ──────────────────────────────────────
    .addSubMenu(ui.createMenu('📥 データ読み込み（CSV）')
      .addItem('🚛 運行シート',                 'showCsvImportDialogUnkou')
      .addItem('🗄 自車専属マスタ',              'showCsvImportDialogMaster')
      .addItem('📇 マスタ（取引先）',             'showCsvImportDialogCust')
      .addSeparator()
      .addItem('⛽ ETC利用明細',                'showEtcImportDialog')
      .addSeparator()
      .addItem('🗑 空インポート行を削除',         'deleteBlankImportRows'))
    // ── 📨 帳票・送信メニュー ──────────────────────────────────────────
    .addSubMenu(ui.createMenu('📨 帳票・送信メニュー')
      .addItem('① 発注書・指示書を作成（協力会社・乗務員用）', 'showHatchuDocDialog')
      .addItem('② 車番連絡を作成（荷主用）',       'showShabanDocDialog')
      .addSeparator()
      .addItem('🗒 受領書の耳生成',              'showUketorishoDialog'))
    // ── 📒 経理・出力 ──────────────────────────────────────────────────
    .addSubMenu(ui.createMenu('📒 経理・出力')
      .addItem('🧾 請求書生成',                 'showInvoiceDialog')
      .addItem('💳 支払確認書生成',              'showPaymentDialog')
      .addItem('📋 監査用表生成',               'generateAuditSheet')
      .addItem('💾 CSV・Excel出力',             'showExportDialog'))
    // ── 📊 PL管理 ──────────────────────────────────────────────────────
    .addSubMenu(ui.createMenu('📊 PL管理')
      .addItem('📈 PL作成',                    'showPlDialog')
      .addItem('🗃 PL設定初期化',               'initFixedCostMaster'))
    // ── 🗓 月またぎ処理 ────────────────────────────────────────────────
    .addSubMenu(ui.createMenu('🗓 月またぎ処理')
      .addItem('📆 今月分生成（途中契約）',       'generateCurrentMonth')
      .addItem('📅 翌月分生成（前月アーカイブ）',  'generateNextMonth')
      .addItem('📦 前月分アーカイブ',            'archiveOldMonth'))
    // ── ⚙️ システム設定・保守 ──────────────────────────────────────────
    .addSubMenu(ui.createMenu('⚙️ システム設定・保守')
      .addItem('🔧 初期設定',                   'installTriggers')
      .addItem('🔄 メニュー再生成（メニューが消えたら押す）', 'reloadMenu')
      .addItem('🗂 シート再生成',               'expandAndRefreshSheets')
      .addItem('🔑 ヘッダー復旧＋保護',          'restoreHeaders')
      .addItem('📃 集計表再生成',               'generateSummary')
      .addItem('🛡 シート保護設定',             'setupSheetProtection')
      .addItem('🔓 保護を全解除',               'removeAllProtections')
      .addItem('📖 使い方シート作成',            'createUsageSheet')
      .addItem('📘 説明書作成',                 'createManualSheet')
      .addItem('📋 サポートテンプレ作成',        'createSupportSheet')
      .addItem('🗾 距離マスタ 主要地データ投入',  'initDistanceMasterMajorCities')
      .addSeparator()
      .addItem('🧹 不正トリガー削除＋システムシート非表示化', 'cleanupStaleTriggers')
      .addSeparator()
      .addItem('🔄 バックアップから復旧',         'openRestoreDialog'));
  menu.addToUi();
}


// ================================================================
//  2-5: 配車ダッシュボード表示（showDispatchDashboard）  【大C / 中2 / 小2-5】
//  メニュー「配車ダッシュボード」からサイドバーで当日の配車状況を一覧表示する
// ================================================================
function showDispatchDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('dispatchDashboard')
    .setTitle('配車ダッシュボード')
    .setWidth(520);
  SpreadsheetApp.getUi().showSidebar(html);
}


// ================================================================
//  配車ダッシュボード用データ取得（getDispatchDashboardData）
//  本日分の運行シートで「積地が空欄（配車未確定）」の行だけを返す。
//  配車漏れチェック専用。
// ================================================================
function getDispatchDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var unkouSheet = ss.getSheetByName('運行');

  var today    = new Date(); today.setHours(0, 0, 0, 0);
  var tomorrow = new Date(today.getTime() + 86400000);

  var gaps = [];

  if (unkouSheet && unkouSheet.getLastRow() > 1) {
    var uCols = Math.min(unkouSheet.getLastColumn(), 13);
    var uData = unkouSheet.getRange(2, 1, unkouSheet.getLastRow() - 1, uCols).getValues();

    for (var i = 0; i < uData.length; i++) {
      var r    = uData[i];
      var id   = String(r[0] || '').trim();   // A=ID
      var dt   = r[9];                          // J=日付
      var car  = String(r[5] || '').trim();   // F=車番
      var drv  = String(r[6] || '').trim();   // G=乗務員名
      var cust = String(r[10] || '').trim();  // K=荷主
      var pick = String(r[11] || '').trim();  // L=積地

      if (!id) continue;                                        // IDなし（空行）は無視
      if (!(dt instanceof Date) || dt < today || dt >= tomorrow) continue; // 今日以外は無視
      if (pick !== '') continue;                               // 積地あり = 配車済みは無視

      gaps.push({ id: id, car: car, driver: drv, customer: cust, rowNum: i + 2 });
    }
  }

  var youbi = ['日', '月', '火', '水', '木', '金', '土'][today.getDay()];
  var dateLabel = Utilities.formatDate(today, 'Asia/Tokyo', 'M月d日') + '（' + youbi + '）';

  return {
    gaps: gaps,
    count: gaps.length,
    date: Utilities.formatDate(today, 'Asia/Tokyo', 'M/d'),
    dateLabel: dateLabel,
    timestamp: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'H:mm:ss')
  };
}


// ================================================================
//  2-1b: メニュー再生成（reloadMenu）  【大C / 中2 / 小2-1b】
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
  var page   = (e && e.parameter && e.parameter.page)   ? e.parameter.page   : '';
  var ssId   = (e && e.parameter && e.parameter.ssId)   ? e.parameter.ssId   : '';
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

  // 本番デプロイのURLのみ保存（テンプレートSS等からのアクセスで上書きされないよう限定）
  try {
    var PROD_DEPLOY_ID = 'AKfycbw7rzkd_SuE1I6BNzEjED4Mxl6cnM4wbswIiRiNoPf5zcSS2JcP6YLkfRV21fLc0opU';
    var svcUrl = ScriptApp.getService().getUrl();
    if (svcUrl && svcUrl.indexOf(PROD_DEPLOY_ID) !== -1) {
      PropertiesService.getScriptProperties().setProperty('webAppUrl', svcUrl);
    }
  } catch(ex) {}

  // 同意処理（?action=agree）：google.script.run不使用でGoogleセキュリティ通知を回避
  if (action === 'agree') {
    // サーバー側バリデーション：agreed=true がない場合はフロント改ざん等として弾く
    if (!e || !e.parameter || e.parameter.agreed !== 'true') {
      return HtmlService.createHtmlOutput(
        '<!DOCTYPE html><html><head><meta charset="utf-8">' +
        '<meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
        '<body style="font-family:sans-serif;padding:24px;text-align:center;">' +
        '<p>同意チェックが確認できませんでした。前のページに戻って再度お試しください。</p>' +
        '</body></html>'
      ).setTitle('エラー - 運行管理システム');
    }
    var agreeCompany = (e && e.parameter && e.parameter.company) ? e.parameter.company : '';
    var agreeRow     = (e && e.parameter && e.parameter.row)     ? e.parameter.row     : '';
    var agreeOk = false;
    try {
      agreeContract(ssId, agreeCompany, '', agreeRow, '');
      agreeOk = true;
    } catch(ex2) {}
    var thanksHtml =
      '<!DOCTYPE html><html><head><meta charset="utf-8">' +
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<title>同意完了 - 運行管理システム</title>' +
      '<style>*{box-sizing:border-box;}body{font-family:-apple-system,BlinkMacSystemFont,\'Hiragino Sans\',sans-serif;' +
      'background:#f5f5f5;margin:0;padding:16px;min-height:100vh;display:flex;align-items:center;justify-content:center;}' +
      '.card{background:#fff;border-radius:8px;padding:32px 24px;max-width:480px;width:100%;' +
      'box-shadow:0 2px 8px rgba(0,0,0,.12);text-align:center;}' +
      '.icon{font-size:52px;margin-bottom:8px;}' +
      'h2{margin:0 0 12px;}p{font-size:14px;color:#555;margin:0;line-height:1.7;}</style></head>' +
      '<body><div class="card">' +
      (agreeOk
        ? '<div class="icon">✅</div><h2 style="color:#188038;">同意が完了しました</h2>' +
          '<p>ご同意ありがとうございます。<br>' +
          'スプレッドシートとアプリのURLを<br>' +
          'メールでお送りしましたのでご確認ください。</p>'
        : '<div class="icon">⚠️</div><h2 style="color:#c62828;">エラーが発生しました</h2>' +
          '<p>もう一度お試しいただくか、<br>担当者にお問い合わせください。</p>') +
      '</div></body></html>';
    return HtmlService.createHtmlOutput(thanksHtml)
      .setTitle('同意完了 - 運行管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 契約書ページ（?page=contract）
  if (page === 'contract') {
    var ctmpl = HtmlService.createTemplateFromFile('contract');
    ctmpl.companySsId  = ssId;
    ctmpl.companyName  = (e && e.parameter && e.parameter.company)     ? e.parameter.company     : '';
    ctmpl.contractRow  = (e && e.parameter && e.parameter.row) ? e.parameter.row : '';
    ctmpl.appUrl       = PropertiesService.getScriptProperties().getProperty('webAppUrl') || ScriptApp.getService().getUrl();
    return ctmpl.evaluate()
      .setTitle('利用規約 - 運行管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 利用規約ページ（?page=terms）
  if (page === 'terms') {
    return HtmlService.createHtmlOutputFromFile('terms')
      .setTitle('利用規約 - 運行管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // プライバシーポリシーページ（?page=privacy）
  if (page === 'privacy') {
    return HtmlService.createHtmlOutputFromFile('privacy')
      .setTitle('プライバシーポリシー - 運行管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 管理画面ページ（?page=parent）
  if (page === 'parent') {
    var ptmpl = HtmlService.createTemplateFromFile('parent_app');
    ptmpl.companySsId = ssId;
    return ptmpl.evaluate()
      .setTitle('運行管理システム 管理画面')
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
//  2-2d: ドライバー認証（validateDriverEmail_）  【大B / 中2 / 小2-2d】
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
      range.setFormula('=IF(U'+row+'=T'+row+',"",T'+row+'-U'+row+')');
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
    if (sheetName === 'PL設定')        { onEditPlSettings_(sheet, range); return; }
    if (sheetName === 'マスタ')         { onEditMasterCustomer_(sheet, range); return; }
    // 配車板（マッチング）シートの自動処理：進捗着色・TEL/FAX自動入力
    if (sheetName === '配車板')           { onEditJoho_(sheet, range, ss); return; }
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
  var editedEndCol = editedCol + range.getNumColumns() - 1;
  if (startRow <= 1) return;
  if (!ss) ss = sheet.getParent();

  var lastColU = Math.max(sheet.getLastColumn(), 22);
  var _oeUnkouSyncIds = [];
  var _allRowsData = sheet.getRange(startRow, 1, numRows, lastColU).getValues(); // 一括先読み（個別読みのキャッシュずれ防止）

  // ── ID採番：IDが必要な行がある時だけロック取得・A列読み込みを行う ──
  var _needIdIdx = [];
  for (var pi = 0; pi < numRows; pi++) {
    if (startRow + pi <= 1) continue;
    var _pr = _allRowsData[pi];
    if (!_pr[0]) {
      for (var pc = 1; pc <= 10; pc++) { // B〜K列
        if (_pr[pc] !== '') { _needIdIdx.push(pi); break; }
      }
    }
  }
  var _idAssignedIdx = {};
  if (_needIdIdx.length > 0) {
    // ScriptLockでID採番を排他制御（並列アクセス時のV-番号重複を根絶）
    var idLock = LockService.getScriptLock();
    try { idLock.waitLock(10000); } catch(e) {
      SpreadsheetApp.getActiveSpreadsheet().toast('ロック取得失敗。再度お試しください。', '⚠️', 5);
      return;
    }
    var vNextIdNum = getNextIdNum_(sheet, 'V-');
    for (var ni = 0; ni < _needIdIdx.length; ni++) {
      var nIdx = _needIdIdx[ni];
      var pidCell = sheet.getRange(startRow + nIdx, 1);
      var pidCur = pidCell.getValue(); // ロック待ちの間に他プロセスが採番した可能性があるため直前再確認
      if (!pidCur) {
        var newId = 'V-' + String(vNextIdNum).padStart(4, '0');
        pidCell.setValue(newId);
        _allRowsData[nIdx][0] = newId;
        _idAssignedIdx[nIdx] = true;
        vNextIdNum++;
      } else {
        _allRowsData[nIdx][0] = pidCur;
      }
    }
    commitLastId_(sheet, 'V-', vNextIdNum - 1);
    SpreadsheetApp.flush();
    idLock.releaseLock();
  }

  // D列(4)=トン数 一括正規化（D列が編集範囲に含まれる時だけ）
  if (startRow >= 2 && editedCol <= 4 && editedEndCol >= 4) {
    var _uTons = _allRowsData.map(function(r) { return [normalizeTons_(r[3])]; }); // D=index3
    var _uChg  = _uTons.some(function(v, i) { return String(v[0]) !== String(_allRowsData[i][3]); });
    if (_uChg) sheet.getRange(startRow, 4, numRows, 1).setValues(_uTons);
  }
  // E列(5)=車種 小文字統一（全角Ｗ→半角・大文字→小文字。E列が編集範囲に含まれる時だけ）
  if (startRow >= 2 && editedCol <= 5 && editedEndCol >= 5) {
    var _uType = _allRowsData.map(function(r) {
      var tv = String(r[4] || '');
      var tn = tv.replace(/[Ａ-Ｚａ-ｚ]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }).toUpperCase();
      return [tn];
    });
    var _uTypeChg = _uType.some(function(v, i) { return String(v[0]) !== String(_allRowsData[i][4]); });
    if (_uTypeChg) {
      sheet.getRange(startRow, 5, numRows, 1).setValues(_uType);
      for (var _ti = 0; _ti < numRows; _ti++) _allRowsData[_ti][4] = _uType[_ti][0];
    }
  }
  // 点呼前完了・点呼後完了の列番号を動的取得（該当しうる列（23列目以降）が編集された時だけヘッダーを読む）
  var _timeCols = [14, 15, 16, 17, 18];
  if (editedEndCol >= 23 || editedCol === 18 || editedCol === 14) {
    var _unkHdr0 = sheet.getRange(1, 1, 1, lastColU).getValues()[0];
    var _inspBColNum = _unkHdr0.indexOf('点呼前完了') + 1; // 1-based、見つからなければ0
    var _inspAColNum = _unkHdr0.indexOf('点呼後完了') + 1;
    if (_inspBColNum > 0) _timeCols.push(_inspBColNum);
    if (_inspAColNum > 0) _timeCols.push(_inspAColNum);
  }
  // 車番補完・J列キャッシュクリア用マスタ読み込み
  var _fInRange = (editedCol <= 6 && editedEndCol >= 6);
  var _jInRange = (editedCol <= 10 && editedEndCol >= 10);
  var mData = [];
  if (_fInRange || _jInRange) {
    var master = ss.getSheetByName('自車専属マスタ');
    mData = master ? master.getDataRange().getValues() : [];
  }
  for (var i = 0; i < numRows; i++) {
    var row = startRow + i;
    if (row <= 1) continue;
    var rowData = _allRowsData[i]; // 一括先読みデータを使用
    var currentId = String(rowData[0] || '').trim();
    var dateVal   = rowData[9]; // J列(10)=日付 0-indexed:9

    // J列(10)の日付：時刻部分が 0:00:00 なら現在時刻を付与（J列が編集範囲に含まれる時だけ）
    if (editedCol <= 10 && editedEndCol >= 10) {
      if (typeof dateVal === 'string' && dateVal.trim() !== '') {
        var jStr = dateVal
          .replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); })
          .replace(/[／]/g, '/')
          .replace(/\s/g, '');
        var jM = jStr.match(/^(\d{1,2})\/(\d{1,2})$/);
        if (jM) {
          var nowJ = new Date();
          dateVal = new Date(nowJ.getFullYear(), parseInt(jM[1]) - 1, parseInt(jM[2]),
                             nowJ.getHours(), nowJ.getMinutes(), nowJ.getSeconds());
          sheet.getRange(row, 10).setValue(dateVal);
          rowData[9] = dateVal;
        }
      }
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
      // J列更新後、該当ドライバーのリストキャッシュをクリア（アプリに即時反映）
      var _jDrName = String(rowData[6] || '').trim();
      for (var _jm = 1; _jm < mData.length; _jm++) {
        if (String(mData[_jm][8] || '').trim() === _jDrName) {
          var _jEmail = String(mData[_jm][10] || '').trim();
          if (_jEmail) { clearListCache_(_jEmail); }
          break;
        }
      }
    }

    // F列(6)：車番を入力→自車専属マスタと部分一致で8列一括補完
    // range.getNumColumns()制限を撤廃：複数列貼り付け・Ctrl+Enter一括入力にも対応
    if (_fInRange) {
      var rawCar   = String(rowData[5] || '').trim(); // F列 0-indexed:5
      var inputCar = rawCar.replace(/[０-９Ａ-Ｚａ-ｚ]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }); // 全角英数→半角
      if (inputCar !== rawCar) { // 正規化した値をセルに書き戻す（１０１→101が見た目にも反映される）
        sheet.getRange(row, 6).setValue(inputCar);
        rowData[5] = inputCar;
      }
      var inputCarDigits = (inputCar.match(/\d+/g) || []).join('');
      if (inputCar && mData.length > 1) {
        // B〜I列（F=車番=index5を除く）に何か入力済みなら手入力を保護してマスタ補完しない
        // 例：会社名・乗務員名など別会社の同じ車番がある場合に自動上書きを防ぐ
        var hasOtherInput = [1,2,3,4,6,7,8].some(function(idx) {
          return String(rowData[idx] || '').trim() !== '';
        });
        if (!hasOtherInput) {
          for (var m2 = 1; m2 < mData.length; m2++) {
            var masterCar = String(mData[m2][7] || '').trim().replace(/[０-９Ａ-Ｚａ-ｚ]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
            var masterStatus = String(mData[m2][1] || '').trim();
            if (masterStatus === '故障' || masterStatus === '待機') continue;
            // マッチ条件：完全一致 or 数字部分の完全一致（101→大阪か101は補完、1→101は補完しない）
            var masterCarDigits = (masterCar.match(/\d+/g) || []).join('');
            if (masterCar === inputCar || (inputCarDigits !== '' && inputCarDigits === masterCarDigits)) {
              // B〜I列を1回のsetValuesで一括書き込み（8個→1回）
              sheet.getRange(row, 2, 1, 8).setValues([[
                mData[m2][2], mData[m2][3], mData[m2][5], mData[m2][6],
                masterCar, mData[m2][8], mData[m2][9], mData[m2][4]
              ]]);
              // 以降の処理がバグらないようメモリ上の配列も同期する
              rowData[1] = mData[m2][2]; rowData[2] = mData[m2][3];
              rowData[3] = mData[m2][5]; rowData[4] = mData[m2][6];
              rowData[5] = masterCar;    rowData[6] = mData[m2][8];
              rowData[7] = mData[m2][9]; rowData[8] = mData[m2][4];
              break;
            }
          }
        }
      }
    }

    // N〜R列（誘導・積完・休憩・降完時刻）および点呼前後完了：全角・日付なし時刻を正規化して合成
    if (_timeCols.indexOf(editedCol) !== -1) {
      var timeCell = sheet.getRange(row, editedCol);
      var tv = timeCell.getValue();
      var baseDateObj = (rowData[9] instanceof Date) ? rowData[9] : null;
      var mergedT = null;
      if (typeof tv === 'string' && tv.trim() !== '') {
        var s = tv.trim().replace(/[：]/g, ':').replace(/[　]/g, ' ').replace(/[０-９]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
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
          var bv2 = (_inspBColNum > 0) ? sheet.getRange(row, _inspBColNum).getValue() : null;
          var av2 = (_inspAColNum > 0) ? sheet.getRange(row, _inspAColNum).getValue() : null;
          var gapMsg = null;
          if (editedCol===14 && _inspBColNum > 0 && !bv2) gapMsg = '先に点呼前完了を入力してください';
          else if (editedCol===15 && !gv2) gapMsg='先に誘導時刻を入力してください';
          else if (editedCol===16) gapMsg = !gv2?'先に誘導時刻を入力してください':!pv2?'先に積完時刻を入力してください':null;
          else if (editedCol===17) gapMsg = !gv2?'先に誘導時刻を入力してください':!pv2?'先に積完時刻を入力してください':!rsv2?'先に休憩開始を入力してください':null;
          else if (editedCol===18) gapMsg = !gv2?'先に誘導時刻を入力してください':!pv2?'先に積完時刻を入力してください':!rsv2?'先に休憩開始を入力してください':!rev2?'先に休憩終了を入力してください':(_inspAColNum>0&&!av2)?'先に点呼後完了を入力してください':null;
          else if (_inspAColNum > 0 && editedCol===_inspAColNum) gapMsg = !gv2?'先に誘導時刻を入力してください':!pv2?'先に積完時刻を入力してください':!rsv2?'先に休憩開始を入力してください':!rev2?'先に休憩終了を入力してください':null;
          if (gapMsg) {
            timeCell.clearContent();
            SpreadsheetApp.getActiveSpreadsheet().toast(gapMsg, '⛔ 順序エラー', 4);
            continue;
          }
        }
      }
    }

    // 積地(L=col12)の背景色を即座に設定（L列編集・A列編集・ID新規採番の時だけ書き込み）
    if ((editedCol <= 12 && editedEndCol >= 12) || (editedCol <= 1 && editedEndCol >= 1) || _idAssignedIdx[i]) {
      var pvK = String(rowData[11] || ''); // L列 0-indexed:11
      var rowBgNew;
      if (pvK.indexOf('有休') !== -1) {
        rowBgNew = new Array(lastColU).fill('#e0e0e0');
      } else if (pvK.indexOf('休み') !== -1) {
        rowBgNew = new Array(lastColU).fill('#9e9e9e');
      } else {
        // 通常行: 既存背景（期限色など）を保護し、有休/休み色・保護色を一旦クリアして塗り直す
        rowBgNew = sheet.getRange(row, 1, 1, lastColU).getBackgrounds()[0];
        for (var bci = 0; bci < rowBgNew.length; bci++) {
          if (rowBgNew[bci] === '#e0e0e0' || rowBgNew[bci] === '#9e9e9e' || rowBgNew[bci] === '#eceff1') rowBgNew[bci] = null;
        }
        if (pvK === '' && currentId) {
          rowBgNew[11] = '#fff9c4'; // L列のみ黄色
        } else if (pvK !== '' && rowBgNew[11] === '#fff9c4') {
          rowBgNew[11] = null; // 積地入力あり→黄色解除
        }
        // 保護列（V=22, Y=25, Z=26）の保護色をIDあり行に維持・復元
        if (currentId) {
          if (lastColU >= 22) rowBgNew[21] = '#eceff1';
          if (lastColU >= 25) rowBgNew[24] = '#eceff1';
          if (lastColU >= 26) rowBgNew[25] = '#eceff1';
        }
      }
      sheet.getRange(row, 1, 1, lastColU).setBackgrounds([rowBgNew]);
    }

    // T列(20)=請求高速 を入力したとき U列(21)=実費高速 が空なら自動コピー（オレンジ色）
    if (editedCol === 20 && (rowData[20] === '' || rowData[20] === null)) {
      if (rowData[19] !== '' && rowData[19] !== null) {
        var uAutoCell = sheet.getRange(row, 21);
        uAutoCell.setValue(rowData[19]);
        uAutoCell.setFontColor('#E65100');
      }
    }
    // U列(21)=実費高速 を手入力したとき黒字に戻す／空欄に戻したときT列を再コピー
    if (editedCol === 21) {
      var uCell21 = sheet.getRange(row, 21);
      if ((rowData[20] === '' || rowData[20] === null) && (rowData[19] !== '' && rowData[19] !== null)) {
        uCell21.setValue(rowData[19]);
        uCell21.setFontColor('#E65100');
      } else {
        uCell21.setFontColor(null);
      }
    }

    // V列(22)の合計高速数式（T/U列編集・ID新規採番の時だけチェック）
    if (_idAssignedIdx[i] || (editedCol <= 21 && editedEndCol >= 20)) {
      if (!sheet.getRange(row, 22).getFormula()) {
        sheet.getRange(row, 22).setFormula('=IF(U' + row + '=T' + row + ',"",T' + row + '-U' + row + ')');
      }
    }
    if (currentId && _oeUnkouSyncIds.indexOf(currentId) === -1) _oeUnkouSyncIds.push(currentId);
  }
  // F列補完後のE列（車種）大文字・半角一括変換
  if (_fInRange && startRow >= 2) {
    var _postFType = sheet.getRange(startRow, 5, numRows, 1).getValues();
    var _postFTypeNew = _postFType.map(function(r) {
      var tv = String(r[0] || '');
      return [tv.replace(/[Ａ-Ｚａ-ｚ]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }).toUpperCase()];
    });
    if (_postFTypeNew.some(function(v, i) { return String(v[0]) !== String(_postFType[i][0]); })) {
      sheet.getRange(startRow, 5, numRows, 1).setValues(_postFTypeNew);
    }
  }
  // 金額書式：金額列(S〜V=19〜22)が編集範囲に含まれる時、またはID新規採番時だけ
  if ((editedCol <= 22 && editedEndCol >= 19) || _needIdIdx.length > 0) {
    applyMoneyFormat_(sheet, startRow, numRows, 'unkou');
  }
  // 時刻書式：時刻列(N〜R=14〜18)または点呼列(23列目以降)が編集範囲に含まれる時だけ
  if ((editedCol <= 18 && editedEndCol >= 14) || editedEndCol >= 23) {
    applyDateTimeFormat_(sheet, startRow, numRows);
  }
  // 日付列(J=col10)が編集された場合は両シートをソート（色はsortUnkouByDate_内で一緒に移動）
  if (editedCol === 10) {
    sortUnkouByDate_();
    sortSummaryByDate_();
  }
  // 集計表syncはinstalledOnEdit_（30分トリガー）で処理するためここでは行わない
}


// ================================================================
//  3-2b: 自車専属マスタ車種列の大文字・半角一括クレンジング（cleanMasterCarType）
//  1回限りの手動実行用。G列（車種）を大文字・半角に統一する。
// ================================================================
function cleanMasterCarType() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master || master.getLastRow() < 2) return;
  var lastRow = master.getLastRow();
  var vals = master.getRange(2, 7, lastRow - 1, 1).getValues();
  var newVals = vals.map(function(r) {
    var tv = String(r[0] || '');
    return [tv.replace(/[Ａ-Ｚａ-ｚ]/g, function(c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); }).toUpperCase()];
  });
  master.getRange(2, 7, lastRow - 1, 1).setValues(newVals);
  ss.toast('自車専属マスタの車種を大文字・半角に変換しました', '✅', 5);
}


// ================================================================
//  3-3: 自車専属マスタ編集時の処理（onEditMasterVehicle_）  【大B / 中3 / 小3-3】
//  ・A列が空で他列にデータがあればS-XXXXのIDを自動生成
// ================================================================
//  PL設定シート編集時：B列（月額）に手入力したらフォントを黒にする
//  緑=PL設定初期化で入れた目安値、黒=自分で入力した実績値
// ================================================================
function onEditPlSettings_(sheet, range) {
  var row = range.getRow();
  if (row < 2) return;
  // B列（月額）を手入力したら黒字に（緑は目安値の印）
  var startCol = range.getColumn();
  var endCol   = startCol + range.getNumColumns() - 1;
  if (startCol <= 2 && endCol >= 2) {
    var numRows = range.getNumRows();
    for (var i = 0; i < numRows; i++) {
      var cell = sheet.getRange(row + i, 2);
      if (String(cell.getFontColor()).toLowerCase() === '#1a9a50') {
        cell.setFontColor('#000000');
      }
    }
  }
}


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
  commitLastId_(sheet, 'S-', sNextIdNum - 1);
  SpreadsheetApp.flush();
  sIdLock.releaseLock();

  if (numRows > 0 && startRow > 1) {
    var statusVals = sheet.getRange(startRow, 2, numRows, 1).getValues();
    var rowBgColors = [];
    for (var ci = 0; ci < numRows; ci++) {
      var cStatus = String(statusVals[ci][0] || '').trim();
      var bg = cStatus === '運行' ? '#ffcdd2' : cStatus === '待機' ? '#fff9c4' : cStatus === '故障' ? '#c8e6c9' : null;
      rowBgColors.push(Array(lastCol).fill(bg));
    }
    sheet.getRange(startRow, 1, numRows, lastCol).setBackgrounds(rowBgColors);
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
  // ── バッチ先読み（100行一括貼り付けでも燃費・トン数正規化が確実に動くように）──
  var _bStart = Math.max(startRow, 2);
  var _bRows  = (startRow + numRows - 1) - _bStart + 1;
  var _bData  = _bRows > 0 ? sheet.getRange(_bStart, 1, _bRows, 17).getValues() : [];
  var _bFuel  = _bRows > 0 ? sheet.getRange(_bStart, 12, _bRows, 1).getValues() : [];
  var _newTons = [], _newFuel = [];
  for (var bi = 0; bi < _bRows; bi++) {
    var bTonsRaw = String(_bData[bi][5] || '').trim(); // F=index5
    var normT    = normalizeTons_(bTonsRaw);
    _newTons.push([normT]);
    var tNrm = normT.replace(/[０-９]/g, function(c){return String.fromCharCode(c.charCodeAt(0)-0xFEE0);}).replace(/[ｔＴ]/g,'t').toLowerCase();
    var nPrt = tNrm.replace(/t$/,'');
    var fv   = fuelMap[tNrm] || fuelMap[nPrt+'t'] || fuelMap[nPrt] || '';
    _newFuel.push([fv !== '' ? fv : _bFuel[bi][0]]);
  }
  if (_bRows > 0) {
    sheet.getRange(_bStart, 6,  _bRows, 1).setValues(_newTons);
    sheet.getRange(_bStart, 12, _bRows, 1).setValues(_newFuel);
    SpreadsheetApp.flush();
  }
  for (var i = 0; i < numRows; i++) {
    var row = startRow + i;
    if (row <= 1) continue;
    var _bi = row - _bStart;
    // 行データ（バッチ先読みから取得）
    var mRow = (_bi >= 0 && _bi < _bData.length) ? _bData[_bi] : sheet.getRange(row, 1, 1, 17).getValues()[0];
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

    // Q列(高速を引く)はグレーにしない（給料制でも高速を引く運用があるため常に入力可）
    var rowBg2 = (function(){
      var st = String(sheet.getRange(row, 2).getValue() || '').trim();
      return st === '運行' ? '#ffcdd2' : st === '待機' ? '#fff9c4' : st === '故障' ? '#c8e6c9' : null;
    })();
    sheet.getRange(row, 17).setBackground(rowBg2);

    // 経費列（Q=17〜AE=31）を手入力したら文字色を黒にリセット（自動入力の赤を解除）
    var expS = Math.max(editedStartCol, 17);
    var expE = Math.min(editedEndCol, 31);
    if (expS <= expE) {
      sheet.getRange(row, expS, 1, expE - expS + 1).setFontColor(null);
    }
  }
  refreshActiveVehiclesAuto_(ss);
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
    // Q列(高速を引く)はグレーにしない（給料制でも会社によって高速を引く場合があるため常に入力可）
    var qBg = rowBg;
    warnBgs.push([nBg, opBg, opBg, qBg]);
  }
  sheet.getRange(2, 14, lr - 1, 4).setBackgrounds(warnBgs); // N,O,P,Q
}


// ================================================================
//  3-3b: 車両ステータス変更時の運行シート同期（syncVehicleToCurrentMonth_）  【大B / 中3 / 小3-3b】
//  自車専属マスタのB列（運行/故障/待機）変更時に呼ばれる
//  ・今日以降の積地空（未配車）行を削除
//  ・ステータスが「運行」なら今日〜今月末の行を再生成
//  ・skipSort=true のとき並び替え・色付けをスキップ（一括処理用）
//  [MOD-v1.2] 引数 applyDate を追加。起点日以降の空行削除と生成を行う
// ================================================================
function syncVehicleToCurrentMonth_(veh, skipSort, applyDate, ss) {
  var carNo  = String(veh[7] || '').trim(); // H列(index7)=車番
  var status = String(veh[1] || '').trim(); // B列(index1)=ステータス
  if (!carNo) return;
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;

  var targetDate = applyDate instanceof Date ? applyDate : new Date();
  targetDate.setHours(0, 0, 0, 0);

  // ① 適用日以降の空行を削除（過去行・データあり行は一切触らない）
  //    空行 = 荷主(K=10)〜売上(S=18) が全て空の行
  if (sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 19).getValues();
    var toDelete = [], toDeleteIds = [];
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][5] || '').trim() !== carNo) continue;
      var d = data[i][9];
      var dm;
      if (d instanceof Date) {
        dm = new Date(d);
      } else if (typeof d === 'number' && d > 0) {
        dm = new Date(Math.round((d - 25569) * 86400 * 1000));
      } else if (typeof d === 'string' && d.trim()) {
        dm = new Date(d.trim());
      } else {
        continue;
      }
      if (!dm || isNaN(dm.getTime())) continue;
      dm.setHours(0, 0, 0, 0);
      if (dm < targetDate) continue;
      // 荷主(K=10)〜売上(S=18)のどれかに値があれば保護（絶対に消さない）
      var hasData = false;
      for (var ci = 10; ci <= 18; ci++) {
        if (String(data[i][ci] || '').trim() !== '') { hasData = true; break; }
      }
      if (hasData) continue;
      toDelete.push(i + 2);
      var rowId = String(data[i][0] || '').trim();
      if (rowId) toDeleteIds.push(rowId);
    }
    if (toDelete.length > 0) {
      deleteRowsGrouped_(sheet, toDelete);
      // 削除したIDだけ集計表からピンポイント削除（全リビルド不要）
      var sumSheet = ss.getSheetByName('集計表');
      if (sumSheet && sumSheet.getLastRow() >= 2 && toDeleteIds.length > 0) {
        var sumData = sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, 1).getValues();
        var sumToDelete = [];
        for (var si = 0; si < sumData.length; si++) {
          if (toDeleteIds.indexOf(String(sumData[si][0] || '').trim()) !== -1) sumToDelete.push(si + 2);
        }
        if (sumToDelete.length > 0) deleteRowsGrouped_(sumSheet, sumToDelete);
      }
    }
  }

  // ③ ステータスが「運行」なら起点日〜今月末の行を生成（既存行程のある日はスキップ）
  if (status === '運行') {
    var yr = targetDate.getFullYear(), mo = targetDate.getMonth(), startDay = targetDate.getDate();
    var endDay = new Date(yr, mo + 1, 0).getDate();
    var todayRef = new Date();
    var inCurMo = (yr === todayRef.getFullYear() && mo === todayRef.getMonth());
    var genNext = inCurMo && (todayRef.getDate() >= 20);
    var nxtYr = 0, nxtMo = 0, nxtEnd = 0;
    if (genNext) {
      var nm1 = new Date(yr, mo + 1, 1);
      nxtYr = nm1.getFullYear(); nxtMo = nm1.getMonth();
      nxtEnd = new Date(yr, mo + 2, 0).getDate();
    }
    var lock = LockService.getDocumentLock();
    try { lock.waitLock(30000); } catch(e) { return; }
    try {
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
      var insertRow = sheet.getLastRow() + 1;
      var nextNum   = getNextIdNum_(sheet, 'V-');
      var rowsData  = [], formulas = [];
      for (var day = startDay; day <= endDay; day++) {
        var dateMid = new Date(yr, mo, day); dateMid.setHours(0, 0, 0, 0);
        if (existingDates[dateMid.getTime()]) continue;
        var rowId = 'V-' + String(nextNum).padStart(4, '0'); nextNum++;
        var rn    = insertRow + rowsData.length;
        rowsData.push([rowId, veh[2], veh[3], veh[5], veh[6], veh[7], veh[8], veh[9],
          veh[4], new Date(yr, mo, day), '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
        formulas.push(['=IF(AND(U'+rn+'="",T'+rn+'=""),"",U'+rn+'-T'+rn+')']);
      }
      if (genNext) {
        for (var nday = 1; nday <= nxtEnd; nday++) {
          var ndm = new Date(nxtYr, nxtMo, nday); ndm.setHours(0, 0, 0, 0);
          if (existingDates[ndm.getTime()]) continue;
          var nrid = 'V-' + String(nextNum).padStart(4, '0'); nextNum++;
          var nrn  = insertRow + rowsData.length;
          rowsData.push([nrid, veh[2], veh[3], veh[5], veh[6], veh[7], veh[8], veh[9],
            veh[4], new Date(nxtYr, nxtMo, nday), '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
          formulas.push(['=IF(AND(U'+nrn+'="",T'+nrn+'=""),"",U'+nrn+'-T'+nrn+')']);
        }
      }
      if (rowsData.length > 0) {
        sheet.getRange(insertRow, 1, rowsData.length, 26).setValues(rowsData);
        sheet.getRange(insertRow, 22, formulas.length, 1).setFormulas(formulas);
        sheet.getRange(insertRow, 10, rowsData.length, 1).setNumberFormat('yyyy/MM/dd');
        sheet.getRange(insertRow, 12, rowsData.length, 2).setNumberFormat('@');
      }
      commitLastId_(sheet, 'V-', nextNum - 1);
      SpreadsheetApp.flush();
    } finally { lock.releaseLock(); }
  }
  // 運行（行追加あり）の時だけソート・色付けを実行。待機/故障は削除のみなので不要
  if (!skipSort && status === '運行') { sortUnkouByDate_(ss.getId()); applyHolidayRowColors_(ss); }
}


// ================================================================
//  3-3c: 全車両を今月分に同期（syncAllVehiclesToCurrentMonth_）  【大B / 中3 / 小3-3c】
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
    commitLastId_(sheet, 'M-', mNextIdNum - 1);
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
  var col      = range.getColumn();
  var numCols  = range.getNumColumns();
  var startRow = range.getRow();
  var numRows  = range.getNumRows();

  // 有効データ行（ヘッダー行=1行目を除く）を計算
  var effStart = Math.max(startRow, 2);
  var effEnd   = startRow + numRows - 1;
  if (effStart > effEnd) return;
  var effRows  = effEnd - effStart + 1;

  // ── 進捗色の再適用（全編集に対して常に実施） ──
  // 編集がコピペか単セル入力かに関わらず、B列・P列の現在値で色を確定する
  var bVals = sheet.getRange(effStart, 2,  effRows, 1).getValues();
  var oVals = sheet.getRange(effStart, 16, effRows, 1).getValues(); // P列: 進捗(車両)
  var cargoBgs = [], vehBgs = [];
  for (var ri = 0; ri < effRows; ri++) {
    var cp = String(bVals[ri][0] || '').trim();
    var vp = String(oVals[ri][0] || '').trim();
    var bc_ri = (cp === '確定' && vp === '確定');
    cargoBgs.push(Array(14).fill(bc_ri ? '#f8bbd0' : (cp === '確定' ? '#fff9c4' : (cp === 'キャンセル' || cp === '終了') ? '#eeeeee' : null)));
    vehBgs.push  (Array(15).fill(bc_ri ? '#f8bbd0' : (vp === '確定' ? '#fff9c4' : (vp === 'キャンセル' || vp === '終了') ? '#eeeeee' : null)));
  }
  sheet.getRange(effStart, 1,  effRows, 14).setBackgrounds(cargoBgs); // A〜N(貨物セクション)
  sheet.getRange(effStart, 15, effRows, 15).setBackgrounds(vehBgs);   // O〜AC(車両セクション)

  // ── トン数自動変換: H(8)=貨物トン数, V(22)=車両トン数 ──
  var _tonsEnd = col + numCols - 1;
  [8, 22].forEach(function(tc) {
    if (tc < col || tc > _tonsEnd) return;
    var _tvOld = sheet.getRange(effStart, tc, effRows, 1).getValues();
    var _tvNew = _tvOld.map(function(r) { return [normalizeTons_(r[0])]; });
    var _tvChg = _tvNew.some(function(v, i) { return String(v[0]) !== String(_tvOld[i][0]); });
    if (_tvChg) sheet.getRange(effStart, tc, effRows, 1).setValues(_tvNew);
  });

  // コピペ（複数列一括編集）はTEL/FAX自動入力・運行登録不要なので着色のみで終了
  if (numCols > 1) return;

  // ── ② 確定→即運行登録（B列またはP列を'確定'に変えた時） ──
  if ((col === 2 || col === 16) && effRows === 1) {
    var newProg = String(range.getValue() || '').trim();
    if (newProg === '確定') {
      // col===2(B=貨物確定)→N列(14)確認 / col===16(P=車両確定)→AC列(29)確認
      var idCol      = (col === 2) ? 14 : 29;
      var existingId = String(sheet.getRange(effStart, idCol).getValue() || '').trim();
      if (!existingId) registerJohoRowToUnkou_(sheet, effStart, col, ss);
    }
  }

  // ── ② C列（会社名・貨物）変更時: 取引先マスタからTEL/FAX自動入力（複数行対応） ──
  if (col === 3) {
    var custSh = ss.getSheetByName('マスタ');
    if (!custSh || custSh.getLastRow() < 2) return;
    var cHdrs = custSh.getRange(1, 1, 1, custSh.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h || '').trim(); });
    var cData = custSh.getRange(2, 1, custSh.getLastRow() - 1, custSh.getLastColumn()).getValues();
    var nIdx = cHdrs.indexOf('会社名'), telIdx = cHdrs.indexOf('電話'), faxIdx = cHdrs.indexOf('FAX');
    var cellVals = range.getValues();
    for (var ri = 0; ri < numRows; ri++) {
      var row = startRow + ri;
      if (row < 2) continue;
      var cargoCompany = String(cellVals[ri][0] || '').trim();
      if (!cargoCompany) continue;
      for (var i = 0; i < cData.length; i++) {
        if (String(cData[i][nIdx] || '').trim() === cargoCompany) {
          if (telIdx >= 0) sheet.getRange(row, 4).setValue(cData[i][telIdx]); // D列 TEL
          if (faxIdx >= 0) sheet.getRange(row, 5).setValue(cData[i][faxIdx]); // E列 FAX
          break;
        }
      }
    }
    return;
  }

  // ── ③ Q列(17)（会社名・車両）変更時: 自社マスタ→取引先マスタの順に検索（複数行対応） ──
  if (col === 17) {
    var vmSh = ss.getSheetByName('自車専属マスタ');
    var vmData, vmNIdx, vmTelIdx;
    if (vmSh && vmSh.getLastRow() >= 2) {
      var vmHdrs = vmSh.getRange(1, 1, 1, vmSh.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h || '').trim(); });
      vmData   = vmSh.getRange(2, 1, vmSh.getLastRow() - 1, vmSh.getLastColumn()).getValues();
      vmNIdx   = vmHdrs.indexOf('会社名');
      vmTelIdx = vmHdrs.indexOf('携帯番号');
    }
    var custSh2 = ss.getSheetByName('マスタ');
    var csData, csNIdx, csTelIdx, csFaxIdx;
    if (custSh2 && custSh2.getLastRow() >= 2) {
      var csHdrs = custSh2.getRange(1, 1, 1, custSh2.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h || '').trim(); });
      csData   = custSh2.getRange(2, 1, custSh2.getLastRow() - 1, custSh2.getLastColumn()).getValues();
      csNIdx   = csHdrs.indexOf('会社名');
      csTelIdx = csHdrs.indexOf('電話');
      csFaxIdx = csHdrs.indexOf('FAX');
    }
    var cellVals2 = range.getValues();
    for (var ri2 = 0; ri2 < numRows; ri2++) {
      var row2 = startRow + ri2;
      if (row2 < 2) continue;
      var vehCompany = String(cellVals2[ri2][0] || '').trim();
      if (!vehCompany) continue;
      var vmFound = false;
      if (vmData) {
        for (var vi = 0; vi < vmData.length; vi++) {
          if (String(vmData[vi][vmNIdx] || '').trim() === vehCompany) {
            if (vmTelIdx >= 0) sheet.getRange(row2, 18).setValue(vmData[vi][vmTelIdx]); // R列 TEL
            vmFound = true; break;
          }
        }
      }
      if (csData) {
        for (var csi = 0; csi < csData.length; csi++) {
          if (String(csData[csi][csNIdx] || '').trim() === vehCompany) {
            if (!vmFound && csTelIdx >= 0) sheet.getRange(row2, 18).setValue(csData[csi][csTelIdx]); // R列 TEL（自社マスタで未発見の場合のみ）
            if (csFaxIdx >= 0) sheet.getRange(row2, 19).setValue(csData[csi][csFaxIdx]); // S列 FAX
            break;
          }
        }
      }
    }
    return;
  }
}

// ================================================================
//  3-5a: 情報シート1行を運行シートに即登録（registerJohoRowToUnkou_）  【大B / 中3 / 小3-5a】
//  B列またはO列が'確定'に変わった時にonEditJoho_から呼ばれる
// ================================================================
function registerJohoRowToUnkou_(johoSheet, rowNum, confirmedCol, ss) {
  var rowData  = johoSheet.getRange(rowNum, 1, 1, 29).getValues()[0];
  var unkou    = ss.getSheetByName('運行');
  if (!unkou) return;
  var uLastCol = unkou.getLastColumn();
  if (uLastCol < 1) return;
  var uHdrs    = unkou.getRange(1, 1, 1, uLastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  function uIdx(name) { return uHdrs.indexOf(name); }

  var newRow = [];
  for (var n = 0; n < uLastCol; n++) newRow.push('');

  // B確定→貨物側(A-M)のみ / P確定→車両側(O-AA)のみ を渡す（反対側は null）
  var cargoData = (confirmedCol === 2)  ? rowData : null;
  var vehData   = (confirmedCol === 16) ? rowData : null;
  buildJohoNewRow_(newRow, uIdx, cargoData, vehData, null);

  var nid = 'V-' + String(getNextIdNum_(unkou, 'V-')).padStart(4, '0');
  if (uIdx('ID') >= 0) newRow[uIdx('ID')] = nid;

  var ins = unkou.getLastRow() + 1;
  unkou.getRange(ins, 1, 1, uLastCol).setValues([newRow]);
  if (uIdx('日付') >= 0) unkou.getRange(ins, uIdx('日付') + 1).setNumberFormat('yyyy/MM/dd');

  // N列(14)=貨物登録ID / AC列(29)=車両登録ID に書き分けて「どちら側の確定か」を明確にする
  var idCol = (confirmedCol === 2) ? 14 : 29;
  johoSheet.getRange(rowNum, idCol).setValue(nid);
  try { delaySyncSummary_(nid, ss); } catch(e) {}
}

// ================================================================
//  3-5b: 情報シート全行の進捗色を一括再適用（refreshJohoColors_）  【大B / 中3 / 小3-5b】
//  コピペ後・シート再生成後に正しい進捗色を復元する
// ================================================================
function refreshJohoColors_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('配車板');
  if (!sheet || sheet.getLastRow() < 2) return;
  var numRows = sheet.getLastRow() - 1;
  var bVals   = sheet.getRange(2, 2,  numRows, 1).getValues(); // B列: 進捗(貨物)
  var oVals   = sheet.getRange(2, 16, numRows, 1).getValues(); // P列: 進捗(車両)
  var cargoBgs = [], vehBgs = [];
  for (var i = 0; i < numRows; i++) {
    var cp = String(bVals[i][0] || '').trim();
    var vp = String(oVals[i][0] || '').trim();
    var bc_i = (cp === '確定' && vp === '確定');
    cargoBgs.push(Array(14).fill(bc_i ? '#f8bbd0' : (cp === '確定' ? '#fff9c4' : (cp === 'キャンセル' || cp === '終了') ? '#eeeeee' : null)));
    vehBgs.push(Array(15).fill(bc_i ? '#f8bbd0' : (vp === '確定' ? '#fff9c4' : (vp === 'キャンセル' || vp === '終了') ? '#eeeeee' : null)));
  }
  sheet.getRange(2, 1,  numRows, 14).setBackgrounds(cargoBgs); // A〜N(貨物セクション)
  sheet.getRange(2, 15, numRows, 15).setBackgrounds(vehBgs);   // O〜AC(車両セクション)
}

// ================================================================
//  3-5c: 配車板保存後処理（afterSaveJoho / afterSaveJohoFull）  【大A / 中3 / 小3-5c】
//  親アプリのインライン保存後に呼ばれ、onEditJoho_相当の処理をSS側で実行する
//  行色変更・会社名→TEL/FAX補完・確定→運行シート自動登録 を担う
//  col1Based: 変更されたセルの列番号（1始まり）
// ================================================================
function afterSaveJoho(ssId, rowNum, col1Based) {
  var ss = getTargetSS_(ssId);
  var sheet = ss.getSheetByName('配車板');
  if (!sheet || rowNum < 2) return { ok: false };
  onEditJoho_(sheet, sheet.getRange(rowNum, col1Based), ss);
  return { ok: true };
}
// 編集モーダル用: 貨物・車両両方の進捗チェックを一括実行
function afterSaveJohoFull(ssId, rowNum) {
  var ss = getTargetSS_(ssId);
  var sheet = ss.getSheetByName('配車板');
  if (!sheet || rowNum < 2) return { ok: false };
  onEditJoho_(sheet, sheet.getRange(rowNum, 2),  ss); // 貨物進捗（色 + ID + 運行登録）
  onEditJoho_(sheet, sheet.getRange(rowNum, 16), ss); // 車両進捗（色 + ID + 運行登録）
  return { ok: true };
}
// 親アプリ 配車板追加専用: 日付列が空の先頭行に書き込む（appendRowはチェックボックスで行502に飛ぶため使用禁止）
function appendJohoRow(rowData, ssId) {
  var ss = getTargetSS_(ssId);
  var sheet = ss.getSheetByName('配車板');
  if (!sheet) return { ok: false, msg: '配車板シートが見つかりません' };
  // チェックボックス列は必ず false（空文字を書くと表示が壊れる）
  if (rowData.length > 0)  rowData[0]  = false;
  if (rowData.length > 14) rowData[14] = false;
  var lastRow = sheet.getLastRow();
  var newRowNum = -1;
  if (lastRow >= 2) {
    var dateVals    = sheet.getRange(2, 6,  lastRow - 1, 1).getValues(); // F: 日付(貨物)
    var vehDateVals = sheet.getRange(2, 20, lastRow - 1, 1).getValues(); // T: 日付(車両)
    for (var ri = 0; ri < dateVals.length; ri++) {
      var dv  = dateVals[ri][0];
      var vdv = vehDateVals[ri][0];
      var cEmpty = (dv  === null || dv  === undefined || String(dv).trim()  === '');
      var vEmpty = (vdv === null || vdv === undefined || String(vdv).trim() === '');
      if (cEmpty && vEmpty) { newRowNum = ri + 2; break; }
    }
  }
  if (newRowNum === -1) {
    // 全行使用済みの場合のみ末尾追加
    sheet.appendRow(rowData);
    newRowNum = sheet.getLastRow();
  } else {
    sheet.getRange(newRowNum, 1, 1, rowData.length).setValues([rowData]);
  }
  onEditJoho_(sheet, sheet.getRange(newRowNum, 2),  ss); // B: 進捗(貨物) → 行色
  onEditJoho_(sheet, sheet.getRange(newRowNum, 16), ss); // P: 進捗(車両) → 行色
  onEditJoho_(sheet, sheet.getRange(newRowNum, 3),  ss); // C: 会社名(貨物) → TEL/FAX補完
  onEditJoho_(sheet, sheet.getRange(newRowNum, 17), ss); // Q: 会社名(車両) → TEL/FAX補完
  return { ok: true, rowNum: newRowNum };
}

// ================================================================
//  親アプリ: 会社名→TEL/FAX照会（lookupCompanyContact）
//  追加フォームで会社名入力時にTEL/FAXをリアルタイム補完するために使用
// ================================================================
function lookupCompanyContact(ssId, companyName, side) {
  var ss  = getTargetSS_(ssId);
  var res = { tel: '', fax: '' };
  if (!companyName) return res;
  var name = String(companyName).trim();
  if (!name) return res;

  if (side === 'vehicle') {
    var vmSh = ss.getSheetByName('自車専属マスタ');
    if (vmSh && vmSh.getLastRow() >= 2) {
      var vmHdrs = vmSh.getRange(1, 1, 1, vmSh.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h || '').trim(); });
      var vmData = vmSh.getRange(2, 1, vmSh.getLastRow() - 1, vmSh.getLastColumn()).getValues();
      var vni = vmHdrs.indexOf('会社名'), vti = vmHdrs.indexOf('携帯番号');
      for (var vi = 0; vi < vmData.length; vi++) {
        if (String(vmData[vi][vni] || '').trim() === name) {
          if (vti >= 0) res.tel = String(vmData[vi][vti] || '');
          break;
        }
      }
    }
  }

  var custSh = ss.getSheetByName('マスタ');
  if (custSh && custSh.getLastRow() >= 2) {
    var cHdrs = custSh.getRange(1, 1, 1, custSh.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h || '').trim(); });
    var cData = custSh.getRange(2, 1, custSh.getLastRow() - 1, custSh.getLastColumn()).getValues();
    var ni = cHdrs.indexOf('会社名'), ti = cHdrs.indexOf('電話'), fi = cHdrs.indexOf('FAX');
    for (var ci = 0; ci < cData.length; ci++) {
      if (String(cData[ci][ni] || '').trim() === name) {
        if (!res.tel && ti >= 0) res.tel = String(cData[ci][ti] || '');
        if (fi >= 0) res.fax = String(cData[ci][fi] || '');
        break;
      }
    }
  }
  return res;
}

// ================================================================
//  親アプリ: 配車確定（matchAndConfirmDispatchFromApp）
//  matchAndConfirmDispatch のWebApp版（getUi/getActiveSpreadsheet 不使用）
//  車種不一致時は { needsTypeConfirm:true, cargoType, vehType } を返し
//  クライアント側でユーザーに確認させて resolvedType 付きで再呼び出しする
// ================================================================
function matchAndConfirmDispatchFromApp(ssId, resolvedType) {
  var ss   = getTargetSS_(ssId);
  var joho = ss.getSheetByName('配車板');
  if (!joho || joho.getLastRow() < 2)
    return { ok: false, error: '配車板にデータがありません' };

  var lastRow = joho.getLastRow();
  var allData = joho.getRange(2, 1, lastRow - 1, 29).getValues();
  var cargoRows = [], vehRows = [];
  for (var i = 0; i < allData.length; i++) {
    var rn = i + 2;
    if (allData[i][0]  === true) cargoRows.push({ rowNum: rn, data: allData[i] });
    if (allData[i][14] === true) vehRows.push(  { rowNum: rn, data: allData[i] });
  }

  if (cargoRows.length === 0 && vehRows.length === 0)
    return { ok: false, error: 'A列（貨物）またはO列（車両）にチェックを入れてください' };
  if (cargoRows.length > 2)
    return { ok: false, error: '貨物チェックは2行までにしてください' };
  if (vehRows.length > 1)
    return { ok: false, error: '車両チェックは1行だけにしてください' };
  if (cargoRows.length === 2 && vehRows.length === 0)
    return { ok: false, error: '貨物が2行あります。車両行のO列もチェックしてください' };
  // ピンク行（両側確定済み）は再マッチング不可
  var allChk2 = cargoRows.concat(vehRows.filter(function(vr) {
    return !cargoRows.some(function(cr) { return cr.rowNum === vr.rowNum; });
  }));
  for (var vx2 = 0; vx2 < allChk2.length; vx2++) {
    var rdx = allChk2[vx2].data;
    if (String(rdx[1]).trim() === '確定' && String(rdx[15]).trim() === '確定')
      return { ok: false, error: '配車確定済み（ピンク）の行は再マッチングできません（行' + allChk2[vx2].rowNum + '）' };
  }

  var unkou = ss.getSheetByName('運行');
  if (!unkou) return { ok: false, error: '運行シートが見つかりません' };
  var uLastCol = unkou.getLastColumn();
  var uHdrs    = unkou.getRange(1, 1, 1, uLastCol).getValues()[0]
    .map(function(h) { return String(h || '').trim(); });
  function uIdx2(name) { return uHdrs.indexOf(name); }

  var vehData = vehRows.length > 0 ? vehRows[0].data : null;
  var groupA  = ['W', '箱', '幌', 'アコーディオン', 'Wトレ'];
  var groupB  = ['平', 'ユニック', '平トレ'];

  function checkType2(cData, vData) {
    var cType = cData ? String(cData[8]  || '').trim() : '';
    var vType = vData ? String(vData[22] || '').trim() : '';
    if (resolvedType !== null && resolvedType !== undefined && String(resolvedType).trim() !== '') return String(resolvedType).trim();
    if (!cData || !vData) return vType || cType;
    var cGrp = groupA.indexOf(cType) >= 0 ? 'A' : (groupB.indexOf(cType) >= 0 ? 'B' : '');
    var vGrp = groupA.indexOf(vType) >= 0 ? 'A' : (groupB.indexOf(vType) >= 0 ? 'B' : '');
    if (cGrp && vGrp && cGrp !== vGrp)
      return { needsTypeConfirm: true, cargoType: cType, vehType: vType };
    return vType || cType;
  }

  // ── 運行シートへの追記（既存IDあれば更新、なければ新規） ────────────
  var registeredIds2 = [];
  function addRow2(cData, vData, fType, presetId) {
    var nid = presetId || ('V-' + String(getNextIdNum_(unkou, 'V-')).padStart(4, '0'));
    var idColIdx2 = uIdx2('ID');
    var existRow2 = -1;
    if (presetId && idColIdx2 >= 0 && unkou.getLastRow() >= 2) {
      var srch2 = unkou.getRange(2, idColIdx2 + 1, unkou.getLastRow() - 1, 1).getValues();
      for (var ek2 = 0; ek2 < srch2.length; ek2++) {
        if (String(srch2[ek2][0]).trim() === nid) { existRow2 = ek2 + 2; break; }
      }
    }
    if (existRow2 > 0) {
      var updRow2 = unkou.getRange(existRow2, 1, 1, uLastCol).getValues()[0];
      buildJohoNewRow_(updRow2, uIdx2, cData, vData, typeof fType === 'string' ? fType : null);
      unkou.getRange(existRow2, 1, 1, uLastCol).setValues([updRow2]);
    } else {
      var newRow2 = [];
      for (var n2 = 0; n2 < uLastCol; n2++) newRow2.push('');
      buildJohoNewRow_(newRow2, uIdx2, cData, vData, typeof fType === 'string' ? fType : null);
      if (uIdx2('ID') >= 0) newRow2[uIdx2('ID')] = nid;
      var ins2 = unkou.getLastRow() + 1;
      unkou.getRange(ins2, 1, 1, uLastCol).setValues([newRow2]);
      if (uIdx2('日付') >= 0) unkou.getRange(ins2, uIdx2('日付') + 1).setNumberFormat('yyyy/MM/dd');
    }
    if (registeredIds2.indexOf(nid) === -1) registeredIds2.push(nid);
    try { delaySyncSummary_(nid, ss); } catch(e) {}
    return nid;
  }

  // ── 既存IDを優先取得 ──────────────────────────────────────────────
  function detectExistId2_(cData, vData) {
    if (cData && String(cData[13] || '').trim()) return String(cData[13]).trim();
    if (vData && String(vData[28] || '').trim()) return String(vData[28]).trim();
    return '';
  }

  var ft2;
  if (cargoRows.length === 2 && vehRows.length === 1) {
    ft2 = checkType2(cargoRows[0].data, vehData);
    if (ft2 && ft2.needsTypeConfirm) return ft2;
    var existId2A = detectExistId2_(cargoRows[0].data, null) || detectExistId2_(cargoRows[1].data, null) || detectExistId2_(null, vehData);
    var sharedId2 = existId2A || ('V-' + String(getNextIdNum_(unkou, 'V-')).padStart(4, '0'));
    addRow2(cargoRows[0].data, vehData, ft2, sharedId2);
    addRow2(cargoRows[1].data, vehData, ft2, sharedId2);
  } else {
    var cargoData2 = cargoRows.length > 0 ? cargoRows[0].data : null;
    var isSameRow2 = cargoData2 && vehData && (cargoRows[0].rowNum === vehRows[0].rowNum);
    ft2 = (cargoData2 && vehData && !isSameRow2) ? checkType2(cargoData2, vehData) : (vehData ? String(vehData[22] || '').trim() : null);
    if (ft2 && ft2.needsTypeConfirm) return ft2;
    var existId2B = detectExistId2_(cargoData2, vehData);
    addRow2(cargoData2, vehData, ft2, existId2B || undefined);
  }

  // ── 全マッチ行を両側確定（B・P列）＋ピンク着色 ──────────────────────
  var regId2 = registeredIds2.length > 0 ? registeredIds2[0] : '';
  var allMatchRows2 = [];
  for (var cw2 = 0; cw2 < cargoRows.length; cw2++) {
    if (allMatchRows2.indexOf(cargoRows[cw2].rowNum) === -1) allMatchRows2.push(cargoRows[cw2].rowNum);
  }
  if (vehRows.length > 0 && allMatchRows2.indexOf(vehRows[0].rowNum) === -1) allMatchRows2.push(vehRows[0].rowNum);
  for (var mi2 = 0; mi2 < allMatchRows2.length; mi2++) {
    var mr2 = allMatchRows2[mi2];
    if (regId2) { joho.getRange(mr2, 14).setValue(regId2); joho.getRange(mr2, 29).setValue(regId2); }
    joho.getRange(mr2, 1).setValue(false);  joho.getRange(mr2, 2).setValue('確定');
    joho.getRange(mr2, 15).setValue(false); joho.getRange(mr2, 16).setValue('確定');
    joho.getRange(mr2, 1,  1, 14).setBackground('#f8bbd0');
    joho.getRange(mr2, 15, 1, 15).setBackground('#f8bbd0');
  }
  var idMsg2 = registeredIds2.length > 1 ? registeredIds2[0] + ' の2行程' : registeredIds2[0];
  return { ok: true, msg: idMsg2 + ' を運行シートに登録しました' };
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
function generateSummary(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
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
    var mMgrIdx = mHdrRow.indexOf('担当管理者');
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
        expense:    mExp,
        manager:    mMgrIdx >= 0 ? String(mData[m][mMgrIdx] || '').trim() : ''
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
    '点呼前完了','点呼後完了','拘束時間(h)','点呼前担当者','点呼後担当者'
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
    // 点呼前後完了を Date に正規化（文字列・全角コロン対応）して集計表に書き込む
    var iBefore_g = normInspTime_(g.inspBefore, g.date);
    var iAfter_g  = normInspTime_(g.inspAfter,  g.date);
    // 拘束時間計算（AL=col38）
    // ① 点呼前〜点呼後が揃っていれば法律定義通り
    // ② なければ誘導〜降完で代替（全行表示のため）
    var kosokuH_g = '';
    if (iBefore_g && iAfter_g && iAfter_g > iBefore_g) {
      kosokuH_g = Math.round((iAfter_g.getTime() - iBefore_g.getTime()) / 36000) / 100;
    } else {
      var gT_g = g.guideTime instanceof Date ? g.guideTime : (g.guideTime ? new Date(g.guideTime) : null);
      if (gT_g && g.rawDropTime && !isNaN(gT_g.getTime()) && g.rawDropTime > gT_g) {
        kosokuH_g = Math.round((g.rawDropTime.getTime() - gT_g.getTime()) / 36000) / 100;
      }
    }
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
      iBefore_g || '', iAfter_g || '',  // 点呼前後完了（Dateに正規化済み）
      kosokuH_g,
      iBefore_g ? (pc.manager || '') : '', iAfter_g ? (pc.manager || '') : ''  // 点呼前担当者, 点呼後担当者（時刻あり時のみ）
    ]);
  }

  // 集計表を再書き込み（値のみクリア・背景リセット、枠線・書式・ヘッダー色は保持）
  var prevLR_ = sumSheet.getLastRow();
  var prevLC_ = Math.max(sumSheet.getLastColumn(), 40);
  sumSheet.clearContents();
  if (prevLR_ >= 2) {
    sumSheet.getRange(2, 1, prevLR_ - 1, prevLC_).setBackground(null);
  }
  if (outRows.length > 0) {
    sumSheet.getRange(1, 1, outRows.length, 40).setValues(outRows);
    // 燃料代（Z=26列）に数式を設定（距離÷燃費×ガソリン代）
    if (outRows.length > 1) {
      var fuelFormulas = [];
      for (var fRow = 2; fRow <= outRows.length; fRow++) {
        fuelFormulas.push(['=IF(OR(W'+fRow+'="",Y'+fRow+'=""),"",ROUND(W'+fRow+'/X'+fRow+'*Y'+fRow+',0))']);
      }
      sumSheet.getRange(2, 26, outRows.length - 1, 1).setFormulas(fuelFormulas);
    }
    sumSheet.setFrozenRows(1);
    // ヘッダー行の色を全40列に明示的に再設定（clearContentsで書式保持されるが新列は暗背景+黒文字で不可視になるため上書き）
    sumSheet.getRange(1, 1, 1, 40).setBackground('#37474f').setFontColor('#90a4ae').setFontWeight('bold');
    var _hdrEditCols = [23, 25, 27, 30, 35];
    for (var _hec = 0; _hec < _hdrEditCols.length; _hec++) {
      sumSheet.getRange(1, _hdrEditCols[_hec]).setBackground('#1b5e20').setFontColor('#a5d6a7').setFontWeight('bold');
    }

    // 4時間超で黄色（労働時間過超）、30分未満で水色（休憩不足）の判定閾値
    var F = 4*60*60*1000;
    var T = 30*60*1000;

    // 背景色を numDataRows×40 の配列で一括計算し setBackgrounds を1回だけ呼ぶ
    var numDataRows_ = outRows.length - 1;
    var bgColors_ = [];
    for (var row = 2; row <= outRows.length; row++) {
      var g2      = idMap[idOrder[row-2]];
      var rowVN   = typeof outRows[row-1][21]==='number' ? outRows[row-1][21] : 0;
      var rowZN   = typeof outRows[row-1][25]==='number' ? outRows[row-1][25] : 0;
      var rowPayN = typeof outRows[row-1][26]==='number' ? outRows[row-1][26] : 0;
      var rowExpN = Number(outRows[row-1][27])||0;
      var calcProfit = (Number(g2.sales)||0)-(rowVN+rowZN+rowPayN+rowExpN);
      var rowRed = calcProfit < 0 ? '#ffebee' : null;
      var rowBg_ = [];
      for (var c_ = 0; c_ < 40; c_++) rowBg_.push(rowRed);   // base: 全40列
      rowBg_[14] = rowRed; rowBg_[15] = rowRed; rowBg_[16] = rowRed; rowBg_[17] = rowRed; // 15-18上書き
      if (g2.rawPickTime  && g2.rawRestStart && (g2.rawRestStart-g2.rawPickTime)  > F) { rowBg_[14]='#ffd600'; rowBg_[15]='#ffd600'; }
      if (g2.rawRestStart && g2.rawRestEnd   && (g2.rawRestEnd  -g2.rawRestStart) < T) { rowBg_[15]='#4fc3f7'; rowBg_[16]='#4fc3f7'; }
      if (g2.rawRestEnd   && g2.rawDropTime  && (g2.rawDropTime -g2.rawRestEnd)   > F) { rowBg_[16]='#ffd600'; rowBg_[17]='#ffd600'; }
      // 拘束時間13時間超え：AL列（col38）をオレンジ（利益マイナスの赤と区別）
      var kH_chk = outRows[row-1][37];
      if (typeof kH_chk === 'number' && kH_chk > 13) { rowBg_[37] = '#ff9800'; }
      // 支払い条件不備の警告背景色（グレー）
      var pctNR = Number(outRows[row-1][32])||0, kyuRN = Number(outRows[row-1][31])||0, kariRN = Number(outRows[row-1][30])||0;
      rowBg_[26] = rowRed; rowBg_[27] = rowRed; // 27,28上書き
      rowBg_[30] = rowRed; rowBg_[31] = rowRed; rowBg_[32] = rowRed; // 31-33上書き
      if (!pctNR && !(kyuRN > 0 && kariRN > 0)) {
        if (kyuRN > 0 || kariRN > 0) {
          if (!kyuRN)  rowBg_[31] = '#b0bec5';
          if (!kariRN) rowBg_[30] = '#b0bec5';
        } else if (!outRows[row-1][26]) {
          rowBg_[26] = '#b0bec5';
        }
      }
      // 経費合計（AB=28列）: 仮日数なし経費あり→値を空欄でグレー警告
      if (outRows[row-1][27] === '') { rowBg_[27] = '#b0bec5'; }
      bgColors_.push(rowBg_);
    }
    if (numDataRows_ > 0) sumSheet.getRange(2, 1, numDataRows_, 40).setBackgrounds(bgColors_);
    applyMoneyFormat_(sumSheet, 2, outRows.length - 1, 'summary');
    applyDateTimeFormat_(sumSheet, 2, outRows.length - 1);
    sumSheet.getRange(2, 36, outRows.length - 1, 2).setNumberFormat('M/d HH:mm');
    // 拘束時間（AL=col38）は数値書式（隣の日時列の書式が滲むため明示設定）
    sumSheet.getRange(2, 38, outRows.length - 1, 1).setNumberFormat('0.00');
    applySumEditableBorders_(sumSheet, 2, outRows.length - 1);
  }
  // 旧データより行数が減った場合、余分な行の書式（枠線・色含む）をクリア
  if (prevLR_ > outRows.length) {
    sumSheet.getRange(outRows.length + 1, 1, prevLR_ - outRows.length, prevLC_).clearFormat();
  }

  convertLegacyAdminDataUrls_();
  applyHolidayRowColors_(ss);

  // フィルターをデータ全列に再設定
  var existingFilter = sumSheet.getFilter();
  if (existingFilter) existingFilter.remove();
  if (sumSheet.getLastRow() >= 1) {
    sumSheet.getRange(1, 1, sumSheet.getLastRow(), sumSheet.getLastColumn()).createFilter();
  }
  // ヘッダー行（1行目）の枠線を確実にクリア（データ行の黄色枠が残らないように）
  sumSheet.getRange(1, 1, 1, Math.max(sumSheet.getLastColumn(), 40)).setBorder(false, false, false, false, false, false);
  // 再生成完了 → 次にマスタ編集したとき「いつから？」ダイアログが再表示されるようリセット
  PropertiesService.getScriptProperties().deleteProperty('recalcFromDateSet');
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
function convertLegacyAdminDataUrls() { convertLegacyAdminDataUrls_(); }


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
  if (!sumSheet || sumSheet.getLastRow() < 1) { generateSummary(ss); return; }

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
      var mMgrIdxS = mHdrRowS.indexOf('担当管理者');
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
          expense:    mExp,
          manager:    mMgrIdxS >= 0 ? String(mAllS[m][mMgrIdxS] || '').trim() : ''
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
    var sumIds   = sumSheet.getRange(2, 1, sumLast-1, Math.max(colCount, 38)).getValues();
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

  // 点呼前後完了を Date に正規化（文字列・全角コロン対応）して集計表に書き込む
  var iBefore_s = normInspTime_(inspBeforeTime, g.date);
  var iAfter_s  = normInspTime_(inspAfterTime,  g.date);
  // 拘束時間計算（AL=col38）
  // ① 点呼前〜点呼後が揃っていれば法律定義通り
  // ② なければ誘導〜降完で代替（全行表示のため）
  var kosokuH_s = '';
  if (iBefore_s && iAfter_s && iAfter_s > iBefore_s) {
    kosokuH_s = Math.round((iAfter_s.getTime() - iBefore_s.getTime()) / 36000) / 100;
  } else {
    var gT_s = g.guideTime instanceof Date ? g.guideTime : (g.guideTime ? new Date(g.guideTime) : null);
    if (gT_s && rawDropTime && !isNaN(gT_s.getTime()) && rawDropTime > gT_s) {
      kosokuH_s = Math.round((rawDropTime.getTime() - gT_s.getTime()) / 36000) / 100;
    }
  }
  var rowData = [
    g.id, g.kubun, g.company, g.tons, g.type, g.car, g.name, g.tel, g.kanban||g.company,
    g.date, g.clients.join('・'), spick, sdrop,
    g.guideTime||'', g.pickTime, g.restStart, g.restEnd, g.dropTime,
    g.sales||'', g.tollReq||'', g.tollReal||'', '',
    keepDistance, fuel, keepGas, '', keepPay, expenseVal, '', keepMemo,  // AA=支払い, AB=経費合計, AC=利益空, AD=備考
    kari, kyuryo, pct,
    sIsYukyu ? yukyuRate : '',
    keepOther,   // AI=その他手当（手入力保持）
    iBefore_s || '', iAfter_s || '',  // 点呼前完了, 点呼後完了（Dateに正規化済み）
    kosokuH_s,  // 拘束時間(h)
    iBefore_s ? (pc.manager || '') : '', iAfter_s ? (pc.manager || '') : ''  // 点呼前担当者, 点呼後担当者（時刻あり時のみ）
  ];

  // LockServiceで並行実行による集計表重複挿入を防止
  // _lockAcq_ フラグで「ロック取得できた場合のみ release」にし、
  // waitLock タイムアウト時の releaseLock() 例外による途中終了を防ぐ
  var sumInsLock = LockService.getScriptLock();
  var _lockAcq_ = false;
  try {
    sumInsLock.waitLock(10000);
    _lockAcq_ = true;
  } catch(e) {
    logError_('syncSumSheet_', 'ロック取得に失敗しました（タイムアウト）');
    throw new Error('現在他の処理が実行中です。しばらく待ってから再度お試しください。');
  }
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
      sumSheet.getRange(sumRow, 1, 1, 40).setValues([rowData]);
      SpreadsheetApp.flush();
    } else {
      sumRow = sumSheet.getLastRow()+1;
      if (sumRow === 1) {
        var hdr = ['ID','区分','会社名','トン数','車種','車番','乗務員名','携帯番号','看板名','日付','荷主','積地','降地','誘導時刻','積完時刻','休憩開始','休憩終了','降完時刻','売上','請求(高速代)','実費(高速代)','合計(高速代)','距離','燃費','ガソリン代','燃料代','支払い','経費合計','利益','備考','仮日数','給料','％','有休手当','その他手当','点呼前完了','点呼後完了','拘束時間(h)','点呼前担当者','点呼後担当者'];
        sumSheet.getRange(1, 1, 1, 40).setValues([hdr]);
        sumSheet.setFrozenRows(1);
        sumRow = 2;
      }
      sumSheet.getRange(sumRow, 1, 1, 40).setValues([rowData]);
      SpreadsheetApp.flush();
    }
  } finally {
    if (_lockAcq_) sumInsLock.releaseLock();
  }

  var vSyncVal = (g.tollReq === 0 && g.tollReal === 0) ? '' : (Number(g.tollReal)||0)-(Number(g.tollReq)||0);
  sumSheet.getRange(sumRow, 22).setValue(vSyncVal);
  // 燃料代（Z=26列）に数式を設定（距離÷燃費×ガソリン代）
  sumSheet.getRange(sumRow, 26).setFormula('=IF(OR(W'+sumRow+'="",Y'+sumRow+'=""),"",ROUND(W'+sumRow+'/X'+sumRow+'*Y'+sumRow+',0))');

  applyMoneyFormat_(sumSheet, sumRow, 1, 'summary');
  applyDateTimeFormat_(sumSheet, sumRow, 1);
  sumSheet.getRange(sumRow, 36, 1, 2).setNumberFormat('M/d HH:mm');
  // 拘束時間（AL=col38）は数値書式（隣の日時列の書式が滲むため明示設定）
  sumSheet.getRange(sumRow, 38).setNumberFormat('0.00');
  // この行だけの支払い(AA=col27)をインライン計算
  var pctNum    = Number(pct)    || 0;
  var kyuryoNum = Number(kyuryo) || 0;
  var kariNum   = Number(kari)   || 0;
  var thisToll  = (Number(g.tollReal) || 0) - (Number(g.tollReq) || 0);
  // ── setValue まとめて実行（setBackground より先に確定）──────────────────
  var payCell   = sumSheet.getRange(sumRow, 27);
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
  }
  // AC(29)=利益
  var vSN = typeof vSyncVal==='number' ? vSyncVal : 0;
  var zSN = (Number(keepDistance) > 0 && Number(fuel) > 0 && Number(keepGas) > 0)
    ? Math.round(Number(keepDistance) / Number(fuel) * Number(keepGas)) : 0;
  var resolvedPaySync = finalPaySync !== null ? finalPaySync : (Number(keepPay)||0);
  var salesSync = Number(g.sales)||0;
  var acSyncVal = (!salesSync&&!vSN&&!zSN&&!resolvedPaySync&&!expenseVal) ? '' : salesSync-(vSN+zSN+resolvedPaySync+(Number(expenseVal)||0));
  sumSheet.getRange(sumRow, 29).setValue(acSyncVal);
  sumSheet.getRange(sumRow, 34).setValue(yukyuVal);
  // ── 背景色を1配列に積んで1回のsetBackgroundsで書き込む（API呼び出しを削減）──
  var F430 = 4*60*60*1000, T430 = 30*60*1000;
  var numBgCols = Math.max(sumSheet.getLastColumn(), 38);
  var baseRowBg = null;
  if      (spick.indexOf('有休') !== -1 || sdrop.indexOf('有休') !== -1) baseRowBg = '#e0e0e0';
  else if (spick.indexOf('休み') !== -1 || sdrop.indexOf('休み') !== -1) baseRowBg = '#9e9e9e';
  else if (typeof acSyncVal === 'number' && acSyncVal < 0)               baseRowBg = '#ffebee';
  var rowBgArr = [];
  for (var _bi = 0; _bi < numBgCols; _bi++) rowBgArr.push(baseRowBg);
  // 430ルール（0-based: col14=13, col15=14, col16=15, col17=16, col18=17）
  if (rawPickTime  && rawRestStart && (rawRestStart-rawPickTime)  > F430) { rowBgArr[14]='#ffd600'; rowBgArr[15]='#ffd600'; }
  if (rawRestStart && rawRestEnd   && (rawRestEnd  -rawRestStart) < T430) { rowBgArr[15]='#4fc3f7'; rowBgArr[16]='#4fc3f7'; }
  if (rawRestEnd   && rawDropTime  && (rawDropTime -rawRestEnd)   > F430) { rowBgArr[16]='#ffd600'; rowBgArr[17]='#ffd600'; }
  // 順序エラー着色
  var hasGuide = !!g.guideTime;
  if (!hasGuide    && (rawPickTime||rawRestStart||rawRestEnd||rawDropTime)) rowBgArr[13]='#ff6d00';
  if (!rawPickTime && (rawRestStart||rawRestEnd||rawDropTime))              rowBgArr[14]='#ff6d00';
  if (!rawRestStart && (rawRestEnd||rawDropTime))                           rowBgArr[15]='#ff6d00';
  if (!rawRestEnd  && rawDropTime)                                          rowBgArr[16]='#ff6d00';
  // 支払い/経費条件不備の警告（グレー）
  if (!sIsYukyu && !sIsYasumi) {
    if (pctNum <= 0 && !(kyuryoNum > 0 && kariNum > 0)) {
      if (kyuryoNum > 0 || kariNum > 0) {
        if (!kyuryoNum) rowBgArr[31]='#b0bec5'; // col32
        if (!kariNum)   rowBgArr[30]='#b0bec5'; // col31
      } else if (!keepPay) {
        rowBgArr[26]='#b0bec5'; // col27
      }
    }
    if (expenseVal === '') rowBgArr[27]='#b0bec5'; // col28
  }
  // 拘束時間13時間超え：AL列（col38, index37）をオレンジ（利益マイナスの赤と区別）
  if (typeof kosokuH_s === 'number' && kosokuH_s > 13) rowBgArr[37] = '#ff9800';
  sumSheet.getRange(sumRow, 1, 1, numBgCols).setBackgrounds([rowBgArr]);

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
  var bRaw = sLastRow >= 2 ? settingSheet.getRange(2, bCol + 1, sLastRow - 1, 1).getValues() : [];
  var existB = bRaw.map(function(r){ return String(r[0]||'').trim(); }).filter(function(v){ return v; });

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
      'アルコールチェックの実施'
    ];
    for (var di = 0; di < defaultBefore.length; di++) {
      settingSheet.getRange(di + 2, bCol + 1).setValue(defaultBefore[di]);
    }
  } else {
    // 旧名称リネーム（乗務前点呼・ の prefix を削除）
    for (var bri = 0; bri < bRaw.length; bri++) {
      var bv = String(bRaw[bri][0]||'').trim();
      if (bv === '乗務前点呼・アルコールチェックの実施') {
        settingSheet.getRange(bri + 2, bCol + 1).setValue('アルコールチェックの実施');
        break;
      }
    }
    var hasAlcB = existB.some(function(v){ return v.indexOf('アルコールチェック') !== -1; });
    if (!hasAlcB) {
      settingSheet.getRange(existB.length + 2, bCol + 1).setValue('アルコールチェックの実施');
    }
  }

  // AFTER列: シート全体のlastRowから読み、空白行を詰めて書き直す
  var fullLastRow = settingSheet.getLastRow();
  var aAllRaw = fullLastRow >= 2 ? settingSheet.getRange(2, aCol + 1, fullLastRow - 1, 1).getValues() : [];
  var existA = [];
  for (var ari = 0; ari < aAllRaw.length; ari++) {
    var av2 = String(aAllRaw[ari][0] || '').trim();
    if (av2 === '乗務後点呼・アルコールチェック（帰庫後）の実施') av2 = 'アルコールチェックの実施（帰庫後）';
    if (av2) existA.push(av2);
  }

  if (existA.length === 0) {
    existA = [
      '車両・積載物の異常の有無（タイヤ・車体等）',
      '事故・ヒヤリハットの有無',
      '道路状況・運行状況の異常の有無',
      '翌乗務員への引き継ぎ事項の有無',
      '運行記録（日報）の提出・乗務後点呼の実施',
      'アルコールチェックの実施（帰庫後）'
    ];
  }

  var hasAlcA = existA.some(function(v){ return v.indexOf('アルコールチェック') !== -1; });
  if (!hasAlcA) existA.push('アルコールチェックの実施（帰庫後）');

  // 空白行を詰めてAFTER列を書き直す
  if (fullLastRow >= 2) {
    settingSheet.getRange(2, aCol + 1, fullLastRow - 1, 1).clearContent();
  }
  for (var dj = 0; dj < existA.length; dj++) {
    settingSheet.getRange(dj + 2, aCol + 1).setValue(existA[dj]);
  }
}

// ================================================================
//  4-2b: 運行シートのID・車番一括補完（fillMissingIdsAndCars）  【大C / 中4 / 小4-2b】
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
  commitLastId_(sheet, 'V-', nextIdNum - 1);
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
//  4-2c: 日時入力ダイアログ（showDateTimePicker / setDateTimeToActiveCell_）  【大C / 中4 / 小4-2c】
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
  var safeColName = String(colName).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
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
    '<label>📍 セル: ' + cellAddr + ' （' + safeColName + '）</label>' +
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
//  ヘッダー行復旧＋保護
// ================================================================
function getSheetHeaderDef_(sheetName) {
  var defs = {
    '運行': ['ID','区分','会社名','トン数','車種','車番','乗務員名','携帯番号','看板名','日付','荷主','積地','降地','誘導時刻','積完時刻','休憩開始','休憩終了','降完時刻','売上','請求(高速代)','実費(高速代)','合計(高速代)','備考','管理データ','連絡(端末)','データ(端末)','点呼前完了','点呼後完了'],
    '集計表': ['ID','区分','会社名','トン数','車種','車番','乗務員名','携帯番号','看板名','日付','荷主','積地','降地','誘導時刻','積完時刻','休憩開始','休憩終了','降完時刻','売上','請求(高速代)','実費(高速代)','合計(高速代)','距離','燃費','ガソリン代','燃料代','支払い','経費合計','利益','備考','仮日数','給料','％','有休手当','その他手当','点呼前完了','点呼後完了','拘束時間(h)','点呼前担当者','点呼後担当者'],
    '自車専属マスタ': ['車両ID','運行状態','区分','会社名','看板名','トン数','車種','車番','乗務員名','携帯番号','アドレス','燃費','備考','仮日数','給料','％','高速を引く（引くは〇、引かないは空欄）','車両リース代','任意保険料','自賠責保険料','重量税積立','車検費積立','整備費積立','タイヤ代積立','修理積立','駐車場代','ETCリース料','カーナビリース料','通信費','洗車費','制服費','その他固定費','免許証有効期限','安全教育次回予定日','健康診断次回予定日','適性診断次回予定日','担当管理者'],
    '自車専属運行': ['車両ID','運行状態','区分','会社名','看板名','トン数','車種','車番','乗務員名','携帯番号','アドレス','燃費','備考','仮日数','給料','％'],
    'マスタ': ['マスタID','会社名','電話','FAX','郵便番号','住所','代表者','配車担当','銀行名','支店名','種別','番号','名義','備考','インボイス登録番号','インボイス発行者名（自社名）','メールアドレス'],
    '設定': ['トン数','基準燃費','有休設定','有休金額','業務前点検','業務後点検']
  };
  return defs[sheetName] || null;
}

function restoreAndProtectHeaders_(ss) {
  var names = ['運行','集計表','自車専属マスタ','自車専属運行','マスタ','設定'];
  names.forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (!s) return;
    var hdr = getSheetHeaderDef_(name);
    if (!hdr) return;
    if (s.getMaxColumns() < hdr.length) {
      s.insertColumnsAfter(s.getMaxColumns(), hdr.length - s.getMaxColumns());
    }
    s.getRange(1, 1, 1, hdr.length).setValues([hdr]).setBackground('#efefef');
    s.setFrozenRows(1);
    s.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) {
      if (p.getDescription() === 'ヘッダー行保護') p.remove();
    });
    var protRange = s.getRange(1, 1, 1, Math.max(hdr.length, s.getLastColumn()));
    var prot = protRange.protect().setDescription('ヘッダー行保護');
    prot.removeEditors(prot.getEditors());
    if (prot.canDomainEdit()) prot.setDomainEdit(false);
  });
}

function restoreHeaders() {
  restoreAndProtectHeaders_(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('ヘッダー復旧＋保護が完了しました');
}

// 列削除・並び替え後にデータごと正規列順へ復元する（bkSheet があれば削除列のデータも復元）
function restoreSheetColumnOrder_(sheet, canonHdr, bkSheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    if (sheet.getMaxColumns() < canonHdr.length)
      sheet.insertColumnsAfter(sheet.getMaxColumns(), canonHdr.length - sheet.getMaxColumns());
    sheet.getRange(1, 1, 1, canonHdr.length).setValues([canonHdr]);
    return;
  }
  var rows = Math.max(lastRow, 1);
  var allData = sheet.getRange(1, 1, rows, lastCol).getValues();
  var curHdrs = allData[0].map(function(h) { return String(h || '').trim(); });

  // バックアップシートのデータを一括読み込み（列削除後の復元用）
  var bkData = null;
  var bkHdrs = [];
  if (bkSheet && bkSheet.getLastRow() > 0 && bkSheet.getLastColumn() > 0) {
    bkData = bkSheet.getRange(1, 1, bkSheet.getLastRow(), bkSheet.getLastColumn()).getValues();
    bkHdrs = bkData[0].map(function(h) { return String(h || '').trim(); });
  }

  var newData = allData.map(function() { return []; });
  var used = {};
  canonHdr.forEach(function(name) {
    var src = curHdrs.indexOf(name);
    var bkSrc = bkHdrs.indexOf(name);
    for (var r = 0; r < allData.length; r++) {
      if (src >= 0) {
        newData[r].push(allData[r][src]);
      } else if (bkSrc >= 0 && bkData && r < bkData.length) {
        newData[r].push(bkData[r][bkSrc]); // バックアップから復元
      } else {
        newData[r].push(r === 0 ? name : '');
      }
    }
    if (src >= 0) used[src] = true;
  });
  // 正規定義外の余剰列は末尾に温存
  for (var ci = 0; ci < lastCol; ci++) {
    if (!used[ci]) {
      for (var r = 0; r < allData.length; r++) newData[r].push(allData[r][ci]);
    }
  }
  var nc = newData[0].length;
  if (sheet.getMaxColumns() < nc)
    sheet.insertColumnsAfter(sheet.getMaxColumns(), nc - sheet.getMaxColumns());
  else if (sheet.getMaxColumns() > nc)
    sheet.deleteColumns(nc + 1, sheet.getMaxColumns() - nc);
  sheet.getRange(1, 1, newData.length, nc).setValues(newData);
}

// onChange トリガーハンドラ（列削除・並び替え・シート追加削除を検知して復元）
function dispatchStructureChange(e) {
  try {
    var ct = e.changeType;
    if (['INSERT_ROW','REMOVE_ROW'].indexOf(ct) !== -1) return;
    var cache = CacheService.getScriptCache();
    if (cache.get('__structureRestoring')) return;
    cache.put('__structureRestoring', '1', 60);
    try {
      var ss = e.source;

      // ①修正用SS判定: __COMPANY_SS__マーカーがないSSが①（シート追加・削除とも完全自由）
      var isClientSs = !!ss.getSheetByName('__COMPANY_SS__');

      // シートが削除されたら _BK_ から全データを即復元（②③のみ）
      if (ct === 'REMOVE_GRID') {
        if (!isClientSs) return; // ①修正用SSは削除自由
        var curNames = ss.getSheets().map(function(s) { try { return s.getName(); } catch(e2) { return ''; } });
        ss.getSheets().forEach(function(bkSheet) {
          try {
            var bkName = bkSheet.getName();
            if (bkName.indexOf('_BK_') !== 0) return;
            var orig = bkName.slice(4);
            if (curNames.indexOf(orig) !== -1) return; // まだ存在する
            var ns = ss.insertSheet(orig);
            var lr = bkSheet.getLastRow(), lc = bkSheet.getLastColumn();
            if (lr > 0 && lc > 0) {
              if (ns.getMaxRows() < lr) ns.insertRowsAfter(ns.getMaxRows(), lr - ns.getMaxRows());
              if (ns.getMaxColumns() < lc) ns.insertColumnsAfter(ns.getMaxColumns(), lc - ns.getMaxColumns());
              ns.getRange(1, 1, lr, lc).setValues(bkSheet.getRange(1, 1, lr, lc).getValues());
            }
            ss.toast('「' + orig + '」は削除できません。元に戻しました', '🔄', 5);
          } catch(e2) {}
        });
        return;
      }

      // ユーザーが追加した不正なシートを即削除（②③のみ）
      if (ct === 'INSERT_GRID') {
        if (!isClientSs) return; // ①修正用SSはシート追加自由
        var allowed = [
          '運行','集計表','自車専属マスタ','自車専属運行','マスタ','設定','メモ',
          '使い方','説明書','配車板','距離マスタ','受領書_耳','受領',
          'PL','PL設定','仕訳表','監査用','自社設定','管理者',
          '請求書','支払確認書','サポート','_ErrorLog_',
          '指示先履歴','指示先ID別'
        ];
        ss.getSheets().forEach(function(sh) {
          try {
            var sname = sh.getName();
            if (sname.indexOf('_BK_') === 0) return;
            if (sname.indexOf('__') === 0) return;
            if (allowed.indexOf(sname) !== -1) return;
            if (sname.indexOf('メモ') === 0) return;
            ss.deleteSheet(sh);
            ss.toast('シートの追加はできません', '🚫', 4);
          } catch(e2) {}
        });
        return;
      }

      // 列並び替え・削除を検知して復元
      var restored = false;
      var names = ['運行','集計表','自車専属マスタ','自車専属運行','マスタ','設定'];
      names.forEach(function(name) {
        var s = ss.getSheetByName(name);
        if (!s) return;
        var canon = getSheetHeaderDef_(name);
        if (!canon) return;
        var lc = s.getLastColumn();
        if (lc === 0) return;
        var cur = s.getRange(1, 1, 1, lc).getValues()[0].map(function(h) { return String(h || '').trim(); });
        var bad = (lc < canon.length);
        if (!bad) {
          for (var i = 0; i < canon.length; i++) {
            if (cur[i] !== canon[i]) { bad = true; break; }
          }
        }
        if (bad) {
          var bkSh = ss.getSheetByName('_BK_' + name);
          restoreSheetColumnOrder_(s, canon, bkSh);
          restored = true;
        }
      });
      if (restored) ss.toast('列構成を元に戻しました', '🔄', 4);
    } finally {
      cache.remove('__structureRestoring');
    }
  } catch(ex) {}
}

function onStructureChange_(e) { dispatchStructureChange(e); }

function setNoSheetGuard_() {
  PropertiesService.getDocumentProperties().setProperty('__NO_SHEET_GUARD__', 'true');
  SpreadsheetApp.getActiveSpreadsheet().toast('このSSのシート追加保護を解除しました', '✅', 3);
}

// ================================================================
//  列バックアップ（隠しシート _BK_xxx）
//  列削除後にデータを復元するための事前バックアップ機構
// ================================================================
function getOrCreateBackupSheet_(ss, sheetName, canon) {
  var bkName = '_BK_' + sheetName;
  var bk = ss.getSheetByName(bkName);
  if (!bk) {
    bk = ss.insertSheet(bkName);
    bk.hideSheet();
    if (bk.getMaxColumns() < canon.length)
      bk.insertColumnsAfter(bk.getMaxColumns(), canon.length - bk.getMaxColumns());
    bk.getRange(1, 1, 1, canon.length).setValues([canon]);
  }
  return bk;
}

function backupSheetRow_(ss, sheetName, rowIdx) {
  var canon = getSheetHeaderDef_(sheetName);
  if (!canon) return;
  var s = ss.getSheetByName(sheetName);
  if (!s || rowIdx < 1 || rowIdx > s.getLastRow()) return;
  var lastCol = s.getLastColumn();
  if (lastCol === 0) return;
  var rowData = s.getRange(rowIdx, 1, 1, lastCol).getValues()[0];
  var curHdrs = s.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h||'').trim(); });
  var canonRow = canon.map(function(name) {
    var src = curHdrs.indexOf(name);
    return src >= 0 ? rowData[src] : '';
  });
  var bk = getOrCreateBackupSheet_(ss, sheetName, canon);
  if (bk.getMaxRows() < rowIdx) bk.insertRowsAfter(bk.getMaxRows(), rowIdx - bk.getMaxRows());
  if (bk.getMaxColumns() < canon.length) bk.insertColumnsAfter(bk.getMaxColumns(), canon.length - bk.getMaxColumns());
  bk.getRange(rowIdx, 1, 1, canonRow.length).setValues([canonRow]);
}

function fullBackupSheet_(ss, sheetName) {
  var s = ss.getSheetByName(sheetName);
  if (!s) return;
  var canon = getSheetHeaderDef_(sheetName);
  if (!canon) return;
  var lastRow = s.getLastRow();
  if (lastRow === 0) return;
  var lastCol = s.getLastColumn();
  if (lastCol === 0) return;
  var allData = s.getRange(1, 1, lastRow, lastCol).getValues();
  var curHdrs = allData[0].map(function(h) { return String(h||'').trim(); });
  var newData = allData.map(function(row, ri) {
    return canon.map(function(name) {
      var src = curHdrs.indexOf(name);
      return src >= 0 ? row[src] : (ri === 0 ? name : '');
    });
  });
  var bk = getOrCreateBackupSheet_(ss, sheetName, canon);
  if (bk.getLastRow() > 0) bk.clearContents();
  if (bk.getMaxRows() < lastRow) bk.insertRowsAfter(bk.getMaxRows(), lastRow - bk.getMaxRows());
  if (bk.getMaxColumns() < canon.length) bk.insertColumnsAfter(bk.getMaxColumns(), canon.length - bk.getMaxColumns());
  bk.getRange(1, 1, newData.length, canon.length).setValues(newData);
}

function backupAllSheets_(ss) {
  ['運行','集計表','自車専属マスタ','自車専属運行','マスタ','設定'].forEach(function(name) {
    try { fullBackupSheet_(ss, name); } catch(ex) {}
  });
}

function backupAllSheets() {
  backupAllSheets_(SpreadsheetApp.getActiveSpreadsheet());
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
  // ヘッダー修復の前に列順・データを正規位置へ復元（並び替え・削除後のズレを解消）
  ['運行','集計表','自車専属マスタ','自車専属運行','マスタ','設定'].forEach(function(sname) {
    var s = ss.getSheetByName(sname);
    if (!s || s.getLastColumn() === 0) return;
    var canon = getSheetHeaderDef_(sname);
    if (!canon) return;
    restoreSheetColumnOrder_(s, canon, ss.getSheetByName('_BK_' + sname));
  });
  restoreAndProtectHeaders_(ss);
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

  // 自車専属マスタにコンプライアンス管理4列を追加（なければ）
  var masterCompCols = ['免許証有効期限', '安全教育次回予定日', '健康診断次回予定日', '適性診断次回予定日'];
  if (masterSheet) {
    var mCLastCol = masterSheet.getLastColumn();
    var mCHeaders = mCLastCol > 0 ? masterSheet.getRange(1, 1, 1, mCLastCol).getValues()[0] : [];
    var mCNextCol = mCLastCol + 1;
    for (var cj = 0; cj < masterCompCols.length; cj++) {
      if (mCHeaders.indexOf(masterCompCols[cj]) === -1) {
        masterSheet.getRange(1, mCNextCol).setValue(masterCompCols[cj]);
        mCNextCol++;
      }
    }
  }

  // 自車専属マスタに「担当管理者」列を追加（なければ）
  if (masterSheet) {
    var mMgrLastCol = masterSheet.getLastColumn();
    var mMgrHeaders = mMgrLastCol > 0 ? masterSheet.getRange(1, 1, 1, mMgrLastCol).getValues()[0] : [];
    if (mMgrHeaders.indexOf('担当管理者') === -1) {
      masterSheet.getRange(1, mMgrLastCol + 1).setValue('担当管理者');
    }
  }

  // コンプライアンス4列・担当管理者の重複列を後方から自動削除（ensureSettingItems_の多重実行で生じた二重追加を修正）
  if (masterSheet) {
    var _dedupTargets = ['免許証有効期限','安全教育次回予定日','健康診断次回予定日','適性診断次回予定日','担当管理者'];
    var _dedupFound = true;
    while (_dedupFound) {
      _dedupFound = false;
      var _dHdrs = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
      for (var _di = 0; _di < _dedupTargets.length; _di++) {
        var _first = _dHdrs.indexOf(_dedupTargets[_di]);
        if (_first === -1) continue;
        var _last = _dHdrs.lastIndexOf(_dedupTargets[_di]);
        if (_last === _first) continue;
        masterSheet.deleteColumn(_last + 1); // 後方の重複列を削除
        _dedupFound = true;
        break; // 列番号が変わったので再スキャン
      }
    }
  }

  // 管理者シートを作成（なければ）
  var adminMgrSheet = ss.getSheetByName('管理者');
  if (!adminMgrSheet) {
    adminMgrSheet = ss.insertSheet('管理者');
    adminMgrSheet.getRange(1, 1, 1, 3).setValues([['管理者名', 'メールアドレス', '備考']]);
    adminMgrSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#37474f').setFontColor('#ffffff');
  }

  // 自車専属マスタの燃費を設定シートから一括再計算（一括貼り付け後にシート再生成でも確実に反映）
  var _earFuelSetting = ss.getSheetByName('設定');
  var _earFuelMap = {};
  if (_earFuelSetting && _earFuelSetting.getLastRow() >= 2) {
    var _earSVals = _earFuelSetting.getRange(2, 1, _earFuelSetting.getLastRow()-1, 4).getValues();
    for (var _ese = 0; _ese < _earSVals.length; _ese++) {
      var _esKey = String(_earSVals[_ese][0]||'').trim()
        .replace(/[０-９]/g, function(c){return String.fromCharCode(c.charCodeAt(0)-0xFEE0);})
        .replace(/[ｔＴ]/g,'t').toLowerCase();
      if (_esKey && _esKey !== '有休') {
        var _esNum = _esKey.replace(/t$/,'');
        _earFuelMap[_esKey]       = _earSVals[_ese][1];
        _earFuelMap[_esNum]       = _earSVals[_ese][1];
        _earFuelMap[_esNum + 't'] = _earSVals[_ese][1];
      }
    }
  }
  if (masterSheet && masterSheet.getLastRow() >= 2) {
    var _earLR      = masterSheet.getLastRow();
    var _earTons    = masterSheet.getRange(2, 6, _earLR - 1, 1).getValues();  // F=トン数
    var _earCurFuel = masterSheet.getRange(2, 12, _earLR - 1, 1).getValues(); // L=燃費（現在値）
    var _earNewFuel = [];
    for (var _efi = 0; _efi < _earTons.length; _efi++) {
      var _eft  = String(_earTons[_efi][0] || '').trim();
      if (_eft) {
        var _efn  = normalizeTons_(_eft);
        var _efnL = _efn.replace(/[０-９]/g, function(c){return String.fromCharCode(c.charCodeAt(0)-0xFEE0);}).replace(/[ｔＴ]/g,'t').toLowerCase();
        var _efp  = _efnL.replace(/t$/,'');
        var _efv  = _earFuelMap[_efnL] || _earFuelMap[_efp+'t'] || _earFuelMap[_efp] || '';
        _earNewFuel.push([_efv !== '' ? _efv : _earCurFuel[_efi][0]]);
      } else {
        _earNewFuel.push([_earCurFuel[_efi][0]]);
      }
    }
    masterSheet.getRange(2, 12, _earLR - 1, 1).setValues(_earNewFuel);
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
      {name:'受領書',      bg:'#006064', fg:'#e0f7fa'},
      {name:'帳票備考',    bg:'#4e342e', fg:'#efebe9'}
    ];
    docCols.forEach(function(col){
      if (uHdrs.indexOf(col.name) === -1) {
        var c = unkouInspSheet.getRange(1, uNext);
        c.setValue(col.name).setBackground(col.bg).setFontColor(col.fg).setFontWeight('bold');
        uNext++;
      }
    });
    // 受領書が帳票備考より後ろにある場合は帳票備考の直前に移動する
    uLastCol = unkouInspSheet.getLastColumn();
    uRawHdrs = uLastCol > 0 ? unkouInspSheet.getRange(1, 1, 1, uLastCol).getValues()[0] : [];
    uHdrs    = uRawHdrs.map(function(h){ return String(h||'').trim(); });
    var ukPos = uHdrs.indexOf('受領書') + 1;  // 1-based
    var bPos  = uHdrs.indexOf('帳票備考') + 1; // 1-based
    if (ukPos > 0 && bPos > 0 && ukPos > bPos) {
      // 帳票備考の前に新列を挿入して受領書ヘッダー＋データをコピーし、元列を削除
      var _ukRows = unkouInspSheet.getLastRow() - 1;
      var _ukData = _ukRows > 0 ? unkouInspSheet.getRange(2, ukPos, _ukRows, 1).getValues() : null;
      unkouInspSheet.insertColumnBefore(bPos);
      unkouInspSheet.getRange(1, bPos).setValue('受領書')
        .setBackground('#006064').setFontColor('#e0f7fa').setFontWeight('bold');
      if (_ukData && _ukRows > 0) {
        unkouInspSheet.getRange(2, bPos, _ukRows, 1).setValues(_ukData);
      }
      unkouInspSheet.deleteColumn(ukPos + 1); // 挿入で1列右にずれているので+1
    }
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

  applySheetColors_(ss);

  // 設定シートに業務前点検・業務後点検データがなければデフォルトを挿入
  ensureSettingItems_(ss);

  // 自車専属マスタ B列（運行状態）に 運行・故障・待機 のドロップダウンを設定
  if (masterSheet && masterSheet.getMaxRows() > 1) {
    var dropRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['運行','故障','待機'], true)
      .setAllowInvalid(false).build();
    masterSheet.getRange(2, 2, masterSheet.getMaxRows() - 1, 1).setDataValidation(dropRule);
  }

  // 自車専属マスタ 経費列・日付列のフォーマット適用
  if (masterSheet && masterSheet.getLastRow() >= 2) {
    var mFmtLastCol = masterSheet.getLastColumn();
    var mFmtHdrs    = masterSheet.getRange(1, 1, 1, mFmtLastCol).getValues()[0];
    var mFmtRows    = masterSheet.getLastRow() - 1;
    // 給料列: #,##0
    var kyuryoIdx = mFmtHdrs.indexOf('給料');
    if (kyuryoIdx >= 0) {
      masterSheet.getRange(2, kyuryoIdx + 1, mFmtRows, 1).setNumberFormat('#,##0');
    }
    // 経費15列（車両リース代〜その他固定費）: #,##0
    var expStartIdx = mFmtHdrs.indexOf('車両リース代');
    var expEndIdx   = mFmtHdrs.indexOf('その他固定費');
    if (expStartIdx >= 0 && expEndIdx >= expStartIdx) {
      masterSheet.getRange(2, expStartIdx + 1, mFmtRows, expEndIdx - expStartIdx + 1)
        .setNumberFormat('#,##0');
    }
    // コンプライアンス4列（免許証有効期限〜適性診断次回予定日）: yyyy/M/d
    var compDateCols = ['免許証有効期限','安全教育次回予定日','健康診断次回予定日','適性診断次回予定日'];
    compDateCols.forEach(function(colName) {
      var ci = mFmtHdrs.indexOf(colName);
      if (ci >= 0) masterSheet.getRange(2, ci + 1, mFmtRows, 1).setNumberFormat('yyyy/M/d');
    });
    // AF（その他固定費）〜AM（担当管理者）のヘッダー行を黒背景・緑文字に統一
    var _afIdx = mFmtHdrs.indexOf('その他固定費');
    var _amIdx = mFmtHdrs.indexOf('担当管理者');
    if (_afIdx >= 0 && _amIdx >= _afIdx) {
      masterSheet.getRange(1, _afIdx + 1, 1, _amIdx - _afIdx + 1)
        .setBackground('#212121').setFontColor('#00e676').setFontWeight('bold');
    }
  }

  ensureCompanySettingSheet_(ss);

  // ── 情報（マッチング）シートの生成・整備 ────────────────────────────
  // 貨物情報と車両情報を1画面で管理し、チェックボックスで配車を確定するシート
  // M&A向け補足: このシートは「配車マン不要化」の核心機能。
  //              荷主からの依頼（貨物情報）と手配した車両（車両情報）を
  //              同一シートで管理し、チェック→確定で運行シートへ自動登録できる。
  (function() {
    // ── 列構成（A〜AB の28列） ────────────────────────────────────────
    // [貨物側] A:チェック(貨物) B:進捗(貨物) C:会社名 D:TEL E:FAX F:日付 G:品名
    //          H:トン数 I:車種 J:積地 K:降地 L:金額(売上) M:備考(貨物) N:貨物登録ID
    // [車両側] O:チェック(車両) P:進捗(車両) Q:会社名 R:TEL S:FAX T:看板名
    //          U:トン数(車両) V:車種(車両) W:車番 X:乗務員名 Y:携帯 Z:金額(支払) AA:備考(車両) AB:車両登録ID
    //
    // 貨物と車両がそれぞれ独立したチェック列・進捗列を持つ設計。
    // 貨物の進捗(B列)を変えても車両側(O-AA)の色は変わらず、逆も同様。
    // 「3行目貨物チェック＋5行目車両チェック」で異なる行同士のマッチングが可能。
    var JOHO_COLS = 29;
    var johoHdr = [
      'チェック(貨物)', '進捗(貨物)',
      '会社名(貨物)', 'TEL(貨物)', 'FAX(貨物)', '日付', '品名',
      'トン数', '車種', '積地', '降地', '金額(売上)', '備考(貨物)',
      '貨物登録ID',
      'チェック(車両)', '進捗(車両)',
      '会社名(車両)', 'TEL(車両)', 'FAX(車両)', '日付(車両)', '看板名',
      'トン数(車両)', '車種(車両)', '車番', '乗務員名', '携帯番号', '金額(支払)', '備考(車両)',
      '車両登録ID'
    ];
    // 旧名称「情報」から「配車板」へ移行（既存SSのシートタブ名を変更）
    var oldJoho = ss.getSheetByName('情報');
    if (oldJoho) oldJoho.setName('配車板');
    var johoSheet = ss.getSheetByName('配車板');
    if (!johoSheet) johoSheet = ss.insertSheet('配車板');

    // ── 日付(車両)列の移行（旧28列→新29列）: T列(20)が看板名なら挿入 ──
    if (johoSheet.getMaxColumns() >= 20) {
      var _col20hdr = String(johoSheet.getRange(1, 20, 1, 1).getValue() || '').trim();
      if (_col20hdr !== '日付(車両)') {
        johoSheet.insertColumnAfter(19); // 旧T(看板名)の前に日付(車両)を挿入
      }
    }

    // 列数が足りない場合は補完（シートの古いバージョンも安全に更新）
    if (johoSheet.getMaxColumns() < JOHO_COLS) {
      johoSheet.insertColumnsAfter(johoSheet.getMaxColumns(), JOHO_COLS - johoSheet.getMaxColumns());
    }

    // ── 旧N構造の判定はヘッダー更新前に実施（更新後は常に新構造になるため） ──
    var _col15HdrBefore = johoSheet.getLastColumn() >= 15
      ? String(johoSheet.getRange(1, 15, 1, 1).getValue() || '').trim() : '';
    var _needsNMig = (_col15HdrBefore !== 'チェック(車両)') && (_col15HdrBefore !== '');

    // ヘッダー行を常に最新定義で上書き（列追加や名称変更に対応するため毎回書き直す）
    johoSheet.getRange(1, 1, 1, JOHO_COLS).setValues([johoHdr]);
    johoSheet.setFrozenRows(1);

    // ── 旧構造からの自動移行（O列ヘッダーが旧名のシートのみ実行） ──
    // 旧構造: N=チェック(bool), O=進捗(確定等), P=会社名... AA=貨物ID, AB=車両ID
    // 新構造: N=貨物ID, O=チェック(bool), P=進捗(確定等), Q=会社名... AA=備考, AB=車両ID
    var lastMigRow = johoSheet.getLastRow();
    if (_needsNMig && lastMigRow >= 2) {
      var col15vals = johoSheet.getRange(2, 15, lastMigRow - 1, 1).getValues();
      var needsMig = false;
      var migKw = {確定: true, キャンセル: true, 終了: true};
      for (var mi = 0; mi < col15vals.length; mi++) {
        var mv = String(col15vals[mi][0] || '').trim();
        if (mv && migKw[mv]) { needsMig = true; break; }
      }
      if (needsMig) {
        var oldCargoIds = johoSheet.getRange(2, 27, lastMigRow - 1, 1).getValues(); // 旧AA=貨物ID
        for (var sc = 26; sc >= 14; sc--) {
          johoSheet.getRange(2, sc + 1, lastMigRow - 1, 1)
            .setValues(johoSheet.getRange(2, sc, lastMigRow - 1, 1).getValues());
        }
        johoSheet.getRange(2, 14, lastMigRow - 1, 1).setValues(oldCargoIds);
        SpreadsheetApp.flush();
      }
    }

    // N(14=貨物登録ID)に紛れ込んだboolean値を除去（列シフト起因の残留データ）
    // ※V-xxxx文字列（運行ID）は正常なので触らない。booleanのみクリア。
    (function() {
      var _nLR = johoSheet.getLastRow() - 1;
      if (_nLR < 1) return;
      var _nD = johoSheet.getRange(2, 14, _nLR, 1).getValues();
      var _nChg = false;
      for (var _ni = 0; _ni < _nD.length; _ni++) {
        if (_nD[_ni][0] === true || _nD[_ni][0] === false) { _nD[_ni][0] = ''; _nChg = true; }
      }
      if (_nChg) johoSheet.getRange(2, 14, _nLR, 1).setValues(_nD);
    })();

    // ── ヘッダー背景色（荷物側/車両側の2色のみ） ───────────────────────
    // A〜N列(貨物側14列): 濃青
    johoSheet.getRange(1, 1, 1, 14)
      .setBackground('#1565c0').setFontColor('#ffffff').setFontWeight('bold');
    // O〜AC列(車両側15列): 濃橙
    johoSheet.getRange(1, 15, 1, 15)
      .setBackground('#e65100').setFontColor('#ffffff').setFontWeight('bold');

    // 列幅
    johoSheet.setColumnWidth(1,  50);  // A: チェック(貨物)
    johoSheet.setColumnWidth(2,  70);  // B: 進捗(貨物)
    johoSheet.setColumnWidth(14, 80);  // N: 貨物登録ID
    johoSheet.setColumnWidth(15, 50);  // O: チェック(車両)
    johoSheet.setColumnWidth(16, 70);  // P: 進捗(車両)
    johoSheet.setColumnWidth(6,  90);  // F: 日付
    johoSheet.setColumnWidth(7, 120);  // G: 品名
    johoSheet.setColumnWidth(20, 90);  // T: 日付(車両)
    johoSheet.setColumnWidth(29, 80);  // AC: 車両登録ID

    // ── データ行バリデーション（2〜501行） ─────────────────────────────
    var dataRows = 500;
    if (johoSheet.getMaxRows() < dataRows + 1) {
      johoSheet.insertRowsAfter(johoSheet.getMaxRows(), dataRows + 1 - johoSheet.getMaxRows());
    }
    // 列操作後に誤った位置へシフトした残留バリデーションを除去してから正規列のみ再設定
    johoSheet.getRange(2, 3,  dataRows, 12).clearDataValidations(); // C〜N（貨物中間列）
    johoSheet.getRange(2, 17, dataRows, 13).clearDataValidations(); // Q〜AC（車両中間列）
    var chkRule  = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    var progRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['', 'キャンセル', '終了', '確定'], true)
      .setAllowInvalid(true).build();
    johoSheet.getRange(2, 1,  dataRows, 1).setDataValidation(chkRule);  // A: 貨物チェック
    johoSheet.getRange(2, 2,  dataRows, 1).setDataValidation(progRule); // B: 進捗(貨物)
    johoSheet.getRange(2, 15, dataRows, 1).setDataValidation(chkRule);  // O: 車両チェック
    johoSheet.getRange(2, 16, dataRows, 1).setDataValidation(progRule); // P: 進捗(車両)
    johoSheet.getRange(2, 6,  dataRows, 1).setNumberFormat('yyyy/MM/dd'); // F: 日付
    johoSheet.getRange(2, 20, dataRows, 1).setNumberFormat('yyyy/MM/dd'); // T: 日付(車両)
    // ヘッダーなしの余分な列（AD以降）を削除
    var _johoLastCol = johoSheet.getLastColumn();
    if (_johoLastCol > JOHO_COLS) {
      var _extraHdrs = johoSheet.getRange(1, JOHO_COLS + 1, 1, _johoLastCol - JOHO_COLS).getValues()[0];
      var _allEmpty = _extraHdrs.every(function(h) { return String(h || '').trim() === ''; });
      if (_allEmpty) johoSheet.deleteColumns(JOHO_COLS + 1, _johoLastCol - JOHO_COLS);
    }
  })();

  applyHolidayRowColors_();
  refreshJohoColors_(ss);
  if (!ss.getSheetByName('メモ')) ss.insertSheet('メモ');
  ['指示先履歴', '指示先ID別'].forEach(function(hname) {
    var hsh = ss.getSheetByName(hname);
    if (hsh && !hsh.isSheetHidden()) hsh.hideSheet();
  });
  SpreadsheetApp.getUi().alert('シート再生成が完了しました。');
}


// ================================================================
//  4-3c: 配車板 列ズレ修復（repairJohoSheet）  【大B / 中4 / 小4-3c】
//  移行コードの誤実行により配車板のO〜AC列が1つ右にズレた場合の修復関数
//  P列(16)にbool値(TRUE/FALSE)が入っていれば列ズレと判定し、左に1シフトして元に戻す
// ================================================================
function repairJohoSheet() {
  var ui    = SpreadsheetApp.getUi();
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('配車板');
  if (!sheet || sheet.getLastRow() < 2) {
    ui.alert('配車板にデータがありません。'); return;
  }
  var numRows = sheet.getLastRow() - 1;

  // P列(16)にbool値があれば列ズレ確定（本来はtrue/false文字列ではなく進捗テキスト）
  var pVals = sheet.getRange(2, 16, numRows, 1).getValues();
  var boolCount = 0;
  for (var i = 0; i < pVals.length; i++) {
    var v = pVals[i][0];
    if (v === true || v === false) boolCount++;
  }
  if (boolCount > 0) {
    var res = ui.alert('配車板 列ズレ修復',
      'P列にチェックボックス値が' + boolCount + '行検出されました。\nN〜AA列(14〜27)を1列左にシフトして元の配置に戻します。\n\nよろしいですか？',
      ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;

    // 現在のN(14)を退避（誤移行で金額(支払)が入っている → 修復後にAA(27)へ戻す）
    var savedN14 = sheet.getRange(2, 14, numRows, 1).getValues();

    // O〜AA (15〜27) を N〜Z (14〜26) へ左シフト
    for (var sc = 14; sc <= 26; sc++) {
      sheet.getRange(2, sc, numRows, 1)
        .setValues(sheet.getRange(2, sc + 1, numRows, 1).getValues());
    }

    // 退避値を AA(27=金額(支払)) に書き戻す
    sheet.getRange(2, 27, numRows, 1).setValues(savedN14);
    SpreadsheetApp.flush();
  }

  // 残留バリデーションをクリアしてから正規列のみ再設定（列ズレの有無に関わらず実施）
  sheet.getRange(2, 3,  numRows, 12).clearDataValidations(); // C〜N
  sheet.getRange(2, 17, numRows, 13).clearDataValidations(); // Q〜AC
  var chkRule  = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  var progRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['', 'キャンセル', '終了', '確定'], true)
    .setAllowInvalid(true).build();
  sheet.getRange(2, 1,  numRows, 1).setDataValidation(chkRule);
  sheet.getRange(2, 2,  numRows, 1).setDataValidation(progRule);
  sheet.getRange(2, 15, numRows, 1).setDataValidation(chkRule);
  sheet.getRange(2, 16, numRows, 1).setDataValidation(progRule);
  sheet.getRange(2, 6,  numRows, 1).setNumberFormat('yyyy/MM/dd');
  sheet.getRange(2, 20, numRows, 1).setNumberFormat('yyyy/MM/dd');

  refreshJohoColors_(ss);
  ui.alert(boolCount > 0
    ? '✅ 配車板の列ズレを修復しました。内容を確認してください。'
    : '✅ 配車板のバリデーション設定を修正しました。内容を確認してください。');
}


// ================================================================
//  4-3a-2: PL按分列を全行自動更新（refreshPlApportionColumn_）  【大B / 中4 / 小4-3a-2】
//  自車専属マスタのB列（状態）変更時にonEditから呼ばれ、AG列の按分額を全行再計算する
// ================================================================
function refreshPlApportionColumn_(ss, sheet) {
  var plSh = ss.getSheetByName('PL設定');
  if (!plSh || plSh.getLastRow() < 2) return;

  // PL設定から月額合計を計算（PL含入フラグ=FALSEは除外）
  var plData  = plSh.getRange(2, 1, plSh.getLastRow() - 1, 5).getValues();
  var plTotal = 0;
  for (var pi = 0; pi < plData.length; pi++) {
    if (!String(plData[pi][0] || '').trim()) continue;
    var flag = plData[pi][4];
    if (flag === false || String(flag).toUpperCase() === 'FALSE') continue;
    plTotal += Number(plData[pi][1]) || 0;
  }
  if (plTotal === 0) return;

  var masterLR = sheet.getLastRow();
  if (masterLR < 2) return;

  // 運行台数をカウント
  var statusAll  = sheet.getRange(2, 2, masterLR - 1, 1).getValues();
  var activeCars = 0;
  for (var si = 0; si < statusAll.length; si++) {
    if (String(statusAll[si][0] || '').trim() === '運行') activeCars++;
  }
  if (activeCars < 1) return;

  var perCar = Math.round(plTotal / activeCars);

  // AG列を動的取得（車両リース代の列 + 15）
  var hdVals = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var expIdx = hdVals.indexOf('車両リース代');
  if (expIdx < 0) return;
  var agCol = expIdx + 16; // 1-based: (expIdx+1) + 15

  // ヘッダーがなければ書く
  if (String(sheet.getRange(1, agCol).getValue()).trim() === '') {
    sheet.getRange(1, agCol).setValue('PL設定按分（参照）')
      .setBackground('#1a2a3a').setFontColor('#00ff88').setFontWeight('bold');
    sheet.setColumnWidth(agCol, 140);
  }

  // 全行を一括更新
  for (var ri = 0; ri < masterLR - 1; ri++) {
    var agCell = sheet.getRange(ri + 2, agCol);
    if (String(statusAll[ri][0] || '').trim() === '運行') {
      agCell.setValue(perCar).setFontColor('#1a9a50').setNumberFormat('#,##0');
    } else {
      agCell.clearContent().clearFormat();
    }
  }
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
  var AUTO_COLOR = '#1a9a50'; // 自動入力値=緑（手入力=黒で区別）

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

  // 手入力（黒字）セルの有無を確認
  var hasManual = false;
  for (var i = 0; i < numSelRows; i++) {
    if (tNums[i] === null) continue;
    for (var j = 0; j < EXP_NUM; j++) {
      var hasVal = existValsAll[i][j] !== '' && existValsAll[i][j] !== 0;
      var fc_ = String(fontColsAll[i][j] || '').toLowerCase();
      var isAuto = fc_ === AUTO_COLOR || fc_ === '#cc0000';
      if (hasVal && !isAuto) { hasManual = true; break; }
    }
    if (hasManual) break;
  }

  var resetAll = false;
  if (hasManual) {
    var resetRes = ui.alert(
      '手入力データが存在します',
      '手入力を初期値（緑字）にリセットしますか？\n\n[はい]：全て基準値で上書き（緑字）\n[いいえ]：手入力は維持、空欄・緑字のみ上書き\n[キャンセル]：中止',
      ui.ButtonSet.YES_NO_CANCEL
    );
    if (resetRes === ui.Button.CANCEL) return;
    resetAll = (resetRes === ui.Button.YES);
  }

  // 新しい値・色の2D配列を構築
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
      var fc_ = String(fontColsAll[i][j] || '').toLowerCase();
      var isAuto = fc_ === AUTO_COLOR || fc_ === '#cc0000';
      if (!resetAll && hasVal && !isAuto) {
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

  // AG列（EXP_COL+EXP_NUM）にPL設定按分額を計算して書き込む（参照用）
  // 稼働台数 = マスタ全体で運行状態が「運行」の行数のみカウント（故障・待機は除外）
  var plSettingsSheet = ss.getSheetByName('PL設定');
  var plTotalPerCar = 0;
  if (plSettingsSheet && plSettingsSheet.getLastRow() >= 2) {
    var plData    = plSettingsSheet.getRange(2, 1, plSettingsSheet.getLastRow() - 1, 5).getValues();
    var plTotal   = 0;
    var masterLR  = sheet.getLastRow();
    var activeCars = 0;
    if (masterLR >= 2) {
      var statusVals = sheet.getRange(2, 2, masterLR - 1, 1).getValues();
      for (var si = 0; si < statusVals.length; si++) {
        if (String(statusVals[si][0] || '').trim() === '運行') activeCars++;
      }
    }
    if (activeCars < 1) activeCars = 1;
    for (var pi = 0; pi < plData.length; pi++) {
      var plName   = String(plData[pi][0] || '').trim();
      var plAmt    = Number(plData[pi][1]) || 0;
      var plFlag   = plData[pi][4];
      if (!plName) continue;
      if (plFlag === false || String(plFlag).toUpperCase() === 'FALSE') continue;
      plTotal += plAmt;
    }
    plTotalPerCar = Math.round(plTotal / activeCars);
  }
  var plAgCol = EXP_COL + EXP_NUM; // AG列（1-based）
  // AG列のヘッダーがなければ追加
  if (String(sheet.getRange(1, plAgCol).getValue()).trim() === '') {
    sheet.getRange(1, plAgCol).setValue('PL設定按分（参照）')
      .setBackground('#1a2a3a').setFontColor('#00ff88').setFontWeight('bold');
    sheet.setColumnWidth(plAgCol, 140);
  }
  // 選択行の運行状態を一括取得（B列）
  var selStatusVals = sheet.getRange(firstRow, 2, numSelRows, 1).getValues();
  for (var ai = 0; ai < numSelRows; ai++) {
    var rowStatus = String(selStatusVals[ai][0] || '').trim();
    var agCell = sheet.getRange(firstRow + ai, plAgCol);
    if (rowStatus === '運行') {
      // 運行中のみ按分額を記入（緑）
      agCell.setValue(plTotalPerCar).setFontColor('#1a9a50').setNumberFormat('#,##0');
    } else {
      // 故障・待機はクリア
      agCell.clearContent().clearFormat();
    }
  }

  var msg = successCount + '行に平均値を入力しました（緑字）。';
  if (skipCount > 0)   msg += '\n' + skipCount + '行はトン数不明のためスキップしました。';
  if (manualTotal > 0) msg += '\n手入力済み項目（合計' + manualTotal + '件）はそのまま残しました。';
  msg += '\n実態に合わせて修正してください。修正すると黒字に変わります。';
  if (plTotalPerCar > 0) msg += '\nPL設定按分（参照）: 1台あたり ' + plTotalPerCar.toLocaleString() + '円/月';
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
function refreshActiveVehiclesAuto_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
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
    commitLastId_(sheet, 'V-', nextNum - 1);
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
    commitLastId_(sheet, 'V-', nextNum - 1);
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
function archiveOldestMonthIfNeeded_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
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
  var result = { email: savedEmail, profile: null, expiryWarning: null };
  if (!savedEmail) return result;
  var ss     = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master) return result;
  var data       = master.getDataRange().getValues();
  var headers    = data[0] || [];
  var emailLower = savedEmail.toLowerCase().trim();

  // 期限列のインデックスを動的に取得
  var expiCols = [
    headers.indexOf('免許証有効期限'),
    headers.indexOf('安全教育次回予定日'),
    headers.indexOf('健康診断次回予定日'),
    headers.indexOf('適性診断次回予定日')
  ].filter(function(c) { return c >= 0; });

  var today = new Date(); today.setHours(0, 0, 0, 0);

  for (var i = 1; i < data.length; i++) {
    if (data[i][10] && String(data[i][10]).toLowerCase().trim() === emailLower) {
      result.profile = {
        company: data[i][3], tons: data[i][5], type: data[i][6],
        carNo:   data[i][7], name: data[i][8], tel:  data[i][9]
      };
      // 期限チェック（最も近い期限日数を計算）
      var minDays = null;
      var minLabel = '';
      for (var ci = 0; ci < expiCols.length; ci++) {
        var dt = data[i][expiCols[ci]];
        if (!(dt instanceof Date) || isNaN(dt.getTime())) continue;
        var exp = new Date(dt); exp.setHours(0, 0, 0, 0);
        var days = Math.round((exp.getTime() - today.getTime()) / 86400000);
        if (minDays === null || days < minDays) {
          minDays = days;
          minLabel = String(headers[expiCols[ci]]);
        }
      }
      if (minDays !== null && minDays <= 7) {
        var status = minDays < 0 ? 'expired' : minDays === 0 ? 'today' : 'soon';
        result.expiryWarning = { minDays: minDays, label: minLabel, status: status };
      }
      var allExpiry = [];
      for (var ci2 = 0; ci2 < expiCols.length; ci2++) {
        var dt3 = data[i][expiCols[ci2]], label3 = String(headers[expiCols[ci2]]);
        if (dt3 instanceof Date && !isNaN(dt3.getTime())) {
          var exp3 = new Date(dt3); exp3.setHours(0,0,0,0);
          var days3 = Math.round((exp3.getTime()-today.getTime())/86400000);
          var ds3 = exp3.getFullYear()+'/'+String(exp3.getMonth()+1).padStart(2,'0')+'/'+String(exp3.getDate()).padStart(2,'0');
          var st3 = days3<0?'expired':days3<=7?'soon':'ok';
          allExpiry.push({label:label3, days:days3, dateStr:ds3, status:st3});
        } else {
          allExpiry.push({label:label3, days:null, dateStr:'', status:'none'});
        }
      }
      result.expiryDates = allExpiry;
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
  function safeJ_(k, d) { try { return JSON.parse(all[k] || d); } catch(e) { return JSON.parse(d); } }
  var lpi = all['lastPickIndex'];
  return {
    picks:         safeJ_('picks',        '[]'),
    drops:         safeJ_('drops',        '[]'),
    rows:          safeJ_('rows',         '[]'),
    runId:         all['runId']           || '',
    guideDone:     safeJ_('guideDone',    '[]'),
    pickDone:      safeJ_('pickDone',     '[]'),
    dropDone:      safeJ_('dropDone',     '[]'),
    phase:         all['phase']           || '',
    lastPickIndex: (lpi !== '' && lpi !== undefined && lpi !== null) ? Number(lpi) : null,
    guideHistory:  safeJ_('guideHistory', '[]'),
    pickHistory:   safeJ_('pickHistory',  '[]'),
    dropHistory:   safeJ_('dropHistory',  '[]')
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
    delaySyncSummary_(sameId, ss);
    SpreadsheetApp.flush(); // 集計表への書き込みを確定してから距離マスタ照合

    // 距離マスタを照合して集計表W列に即時反映（なければMaps APIで即時計算）
    lookupAndSetDistanceAfterCreate_(ss, sameId, picks, drops);

    // 日付順にソート（新規行が末尾に追加されているため）
    sortUnkouByDate_(companySsId);

    clearListCache_(email);
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
  // 集計同期はアプリからの後追い呼び出し（syncSummary）で実行（応答を待たせない）
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
}


// ================================================================
//  7-7: 状態保存＋時刻記録 一括実行（recordAction）  【大A / 中7 / 小7-7】
//  saveRunState と setXxxComplete を1回のサーバー呼び出しでまとめて実行する
//  actionType: 'guide' / 'pick' / 'restStart' / 'restEnd' / 'drop'
// ================================================================
function recordAction(actionType, id, routeIndex, stateObj, companySsId, email) {
  validateDriverEmail_(email, companySsId);
  if (stateObj) saveRunState(stateObj, email, companySsId);
  var now = new Date();
  var hhmm = String(now.getHours()).padStart(2,'0')+':'+String(now.getMinutes()).padStart(2,'0');
  var timeField = null;
  if      (actionType === 'guide')      { setGuideComplete(id, routeIndex, companySsId); timeField = 'guideTime'; }
  else if (actionType === 'pick')       { setPickComplete(id, routeIndex, companySsId);  timeField = 'pickTime'; }
  else if (actionType === 'restStart')  { setRest(id, routeIndex, 'start', companySsId); timeField = 'restStart'; }
  else if (actionType === 'restEnd')    { setRest(id, routeIndex, 'end',   companySsId); timeField = 'restEnd'; }
  else if (actionType === 'drop')       { setDropComplete(id, routeIndex, companySsId);  timeField = 'dropTime'; }
  else if (actionType === 'inspBefore') { setInspectionComplete_(id, 'before', companySsId); timeField = 'inspBefore'; }
  else if (actionType === 'inspAfter')  { setInspectionComplete_(id, 'after',  companySsId); timeField = 'inspAfter'; }
  else if (actionType === 'syncSummary') {
    // アプリからの後追い同期専用：時刻書き込みの応答を待たせないための分離呼び出し
    var ssS = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
    if (id) delaySyncSummary_(id, ssS);
  }
  clearListCache_(email);
  return timeField ? {field: timeField, time: hhmm} : null;
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
}


// ================================================================
//  8-1: 行程データ更新（updateRouteData）  【大A / 中8 / 小8-1】
//  戻るボタン用：IDで行を動的検索してL列（積地）・M列（降地）を更新し集計表を同期する
// ================================================================
function updateRouteData(id, picks, drops, companySsId, email) {
  if (email) validateDriverEmail_(email, companySsId);
  var ss = companySsId ? getTargetSS_(companySsId) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return null;
  var all = sheet.getDataRange().getValues();
  var ri = 0;
  var rows = [];
  var firstRowData = null;
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][0] || '').trim() !== String(id).trim()) continue;
    if (!firstRowData) firstRowData = all[i];
    if (ri < picks.length) {
      sheet.getRange(i + 1, 12).setValue(picks[ri] || '');
      sheet.getRange(i + 1, 13).setValue(drops[ri] || '');
      rows.push(i + 1);
      ri++;
    }
  }
  // 行程が増えた分は同IDで末尾に追加（既存行は絶対に削除しない・上書きしない）
  if (firstRowData && ri < picks.length) {
    var addStart = sheet.getLastRow() + 1;
    var addData  = [];
    for (var a = ri; a < picks.length; a++) {
      addData.push([
        id, firstRowData[1], firstRowData[2], firstRowData[3], firstRowData[4], firstRowData[5],
        firstRowData[6], firstRowData[7], firstRowData[8], firstRowData[9], firstRowData[10],
        picks[a] || '', drops[a] || '', '', '', '', '', '', '', '', '', '', '', '', '', ''
      ]);
    }
    sheet.getRange(addStart, 1, addData.length, 26).setValues(addData);
    sheet.getRange(addStart, 12, addData.length, 2).setNumberFormat('@');
    if (firstRowData[9] instanceof Date) sheet.getRange(addStart, 10, addData.length, 1).setNumberFormat('yyyy/MM/dd');
    for (var b = 0; b < addData.length; b++) rows.push(addStart + b);
  }
  SpreadsheetApp.flush();
  if (id) delaySyncSummary_(id, ss);
  return { rows: rows, id: id };
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
  clearListCache_(email);
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

  // 30秒キャッシュ（毎回フルシート読みを回避。書き込み操作後はclearListCache_で削除）
  var _ldKey = 'ld_' + savedEmail.toLowerCase() + '_' + year + '_' + month;
  var _ldSC  = CacheService.getScriptCache();
  try { var _ldCached = _ldSC.get(_ldKey); if (_ldCached) return JSON.parse(_ldCached); } catch(e) {}

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
        guideTime:'', pickTime:'', restStart:'', restEnd:'', dropTime:'', allDropsDone:true,
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
    if (!r[17]) g.allDropsDone = false;
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
      isComplete: !!(g.guideTime && g.pickTime && g.restStart && g.restEnd && g.allDropsDone && g.inspBefore && g.inspAfter),
      isNew:      !g.guideTime && !g.pickTime && !g.restStart && !g.restEnd && !g.dropTime && !g.inspBefore && !g.inspAfter
    });
    totalSales += g.sales; totalToll += g.tollTotal; totalPay += g.pay;
    totalYukyu += g.yukyu || 0;
    totalOther += g.other || 0;
    if (g.picks.join('').trim() !== '') dateSet[g.date] = true; // 積地空（未配車）はノーカウント
  }
  result.sort(function(a,b){ return b.dateSort - a.dateSort; });
  var _ldResult = { rows:result, total:{ days:Object.keys(dateSet).length, sales:totalSales, toll:totalToll, pay:totalPay, yukyu:totalYukyu, other:totalOther, yukyuDays:yukyuDays, yasumiDays:yasumiDays } };
  try { _ldSC.put(_ldKey, JSON.stringify(_ldResult), 30); } catch(e) {}
  return _ldResult;
}


// ================================================================
//  clearListCache_: 一覧データキャッシュを削除（書き込み操作後に呼ぶ）
// ================================================================
function clearListCache_(email) {
  if (!email) return;
  var sc = CacheService.getScriptCache();
  var now = new Date(), yr = now.getFullYear(), mo = now.getMonth() + 1;
  var prevMo = mo > 1 ? mo - 1 : 12, prevYr = mo > 1 ? yr : yr - 1;
  try { sc.remove('ld_' + email.toLowerCase() + '_' + yr + '_' + mo); } catch(e) {}
  try { sc.remove('ld_' + email.toLowerCase() + '_' + prevYr + '_' + prevMo); } catch(e) {}
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
  var gotLock = lock.tryLock(30000);
  if (!gotLock) {
    logError_('saveEditData', 'ロック取得に失敗しました（タイムアウト）');
    throw new Error('現在他の処理が実行中です。しばらく待ってから再度お試しください。');
  }
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

    // 行ごとの項目（荷主・積地・降地・時刻）は最初の行のみ書き込む
    // ※同IDの2行目以降は別行程（2箇所積み下ろし等）のため絶対に上書きしない
    if (!written) {
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
    }

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
  try { applyHolidayRowColors_(ss); } catch(e) {}
  clearListCache_(email);
  } finally {
    lock.releaseLock();
  }
}


// ================================================================
//  8-6a: 端末連絡保存（saveTermNoticeByDriver）  【大A / 中8 / 小8-6a】
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
    var lastCol = Math.max(sumSheet.getLastColumn(), 40);
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
    // IDあり行のみ保護色を塗る（IDなし行はnull）
    var unkouIdVals = unkouSheet.getRange(2, 1, unkouLastRow - 1, 1).getValues();
    var protBgs = unkouIdVals.map(function(r) {
      return [String(r[0] || '').trim() !== '' ? '#eceff1' : null];
    });
    for (var pc = 0; pc < protectedUnkouCols.length; pc++) {
      var pcol = protectedUnkouCols[pc];
      unkouSheet.getRange(1, pcol).setBackground('#37474f').setFontColor('#90a4ae').setFontWeight('bold');
      unkouSheet.getRange(2, pcol, unkouLastRow - 1, 1).setBackgrounds(protBgs);
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
  applyHolidayRowColors_(ss);
}

// ================================================================
//  点呼前後完了の値を Date に正規化する（全角コロン・日付なし形式に対応）
//  運行シートの値（Date or 文字列）を集計表に書く前に必ず通す
// ================================================================
function normInspTime_(v, baseDate) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (!v) return null;
  var s = String(v).trim().replace(/[：]/g, ':').replace(/[　]/g, ' ');
  // "M/d H:mm" 形式（例: "6/6 21:00"）
  var mx = s.match(/^(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})/);
  if (mx) {
    var yr = (baseDate instanceof Date) ? baseDate.getFullYear() : new Date().getFullYear();
    var d  = new Date(yr, parseInt(mx[1])-1, parseInt(mx[2]), parseInt(mx[3]), parseInt(mx[4]), 0);
    return isNaN(d.getTime()) ? null : d;
  }
  // "H:mm" のみ形式（例: "21:00"）
  var mx2 = s.match(/^(\d{1,2}):(\d{2})$/);
  if (mx2 && baseDate instanceof Date) {
    var d2 = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate(), parseInt(mx2[1]), parseInt(mx2[2]), 0);
    return isNaN(d2.getTime()) ? null : d2;
  }
  return null;
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

function removeAllProtections() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var n = 0;
  ss.getSheets().forEach(function(s) {
    s.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) { p.remove(); n++; });
    s.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p) { p.remove(); n++; });
  });
  SpreadsheetApp.getUi().alert('保護を ' + n + ' 件解除しました');
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
//  8-6b-0a: 看板名列を既存シートに挿入（insertKanbanColumn）  【大C / 中8 / 小8-6b-0a】
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
//  8-6b-2b: 画像URLをDriveに取込（importImageToDrive_）  【大B / 中8 / 小8-6b-2b】
//  公開画像URLをOAuthトークンで取得しDriveにコピーして返す
//  Google フォトのプライベートURLは不可（サイドバーでアップロード推奨）
//  ※ ファイル上は8-6b-1/2の後に配置されているが、補助関数として番号2bとする
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
  var gotLock = lock.tryLock(30000);
  if (!gotLock) {
    logError_('appendTerminalFile', 'ロック取得に失敗しました（タイムアウト）');
    throw new Error('現在他の処理が実行中です。しばらく待ってから再度お試しください。');
  }
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
//  8-6c-a: 端末ファイル追加（管理者用・メール認証なし）  【大A / 中8 / 小8-6c-a】
// ================================================================
function appendTerminalFileAdmin(id, fileName, base64Data, mimeType, companySsId) {
  var lock = LockService.getScriptLock();
  var gotLock = lock.tryLock(30000);
  if (!gotLock) {
    logError_('appendTerminalFileAdmin', 'ロック取得に失敗しました（タイムアウト）');
    throw new Error('現在他の処理が実行中です。しばらく待ってから再度お試しください。');
  }
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
function appendAdminFileById(id, fileName, base64Data, mimeType, companySsId) {
  var folder  = getOrCreateFolder_('運行データ');
  var decoded = Utilities.base64Decode(base64Data);
  var blob    = Utilities.newBlob(decoded, mimeType, fileName);
  var file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url   = file.getUrl();
  var sheet = getTargetSS_(companySsId).getSheetByName('運行');
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

function deleteAdminFileById(id, urlToDelete, companySsId) {
  var sheet = getTargetSS_(companySsId).getSheetByName('運行');
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

function replaceAdminFileById(id, oldUrl, fileName, base64Data, mimeType, companySsId) {
  var folder  = getOrCreateFolder_('運行データ');
  var decoded = Utilities.base64Decode(base64Data);
  var blob    = Utilities.newBlob(decoded, mimeType, fileName);
  var file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var newUrl  = file.getUrl();
  var sheet   = getTargetSS_(companySsId).getSheetByName('運行');
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
function deleteTerminalFile(id, urlToDelete, companySsId) {
  var sheet = getTargetSS_(companySsId).getSheetByName('運行');
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
function replaceTerminalFile(id, oldUrl, fileName, base64Data, mimeType, companySsId) {
  var folder  = getOrCreateFolder_('端末データ');
  var decoded = Utilities.base64Decode(base64Data);
  var blob    = Utilities.newBlob(decoded, mimeType, fileName);
  var file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var newUrl = file.getUrl();
  var sheet = getTargetSS_(companySsId).getSheetByName('運行');
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
  clearListCache_(email);
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
  var ssId = '';
  try { ssId = SpreadsheetApp.getActiveSpreadsheet().getId(); } catch(e) {}
  var tempFile = folder.createFile(Utilities.newBlob(base64Data, 'text/plain', queueId + '.txt'));
  PropertiesService.getScriptProperties().setProperty(queueId, JSON.stringify({
    tempFileId: tempFile.getId(), rowNum: rowNum, fileName: fileName, mimeType: mimeType, ssId: ssId
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
      var qSs = data.ssId ? SpreadsheetApp.openById(data.ssId) : null;
      var sheet = qSs ? qSs.getSheetByName('運行') : null;
      if (sheet) {
        var existing2 = getAdminDataUrl_(sheet, data.rowNum).split(',').filter(function(u) { return u.match(/^https?:\/\//); });
        existing2.push(url);
        var deduped = existing2.filter(function(u, i, arr) { return arr.indexOf(u) === i; });
        setAdminDataRichTextMulti_(sheet, data.rowNum, deduped);
      }
      tempFile.setTrashed(true);
      props.deleteProperty(key);
    } catch (e3) {
      var retryCount = (data.retryCount || 0) + 1;
      var errMsg = (e3 && e3.message ? e3.message : String(e3));
      if (retryCount >= 3) {
        logError_('processUploadQueue', '3回リトライしたが失敗しました。ファイル: ' + (data.fileName || '') + ' / 最終エラー内容: ' + errMsg);
        try {
          var failSs = data.ssId ? SpreadsheetApp.openById(data.ssId) : null;
          if (failSs) markUploadFailureNotice_(failSs);
        } catch (e6) {}
        props.deleteProperty(key);
      } else {
        logError_('processUploadQueue', 'アップロード失敗（' + retryCount + '回目）。ファイル: ' + (data.fileName || '') + ' / エラー内容: ' + errMsg);
        data.retryCount = retryCount;
        try { props.setProperty(key, JSON.stringify(data)); } catch (e5) {}
      }
    }
  }
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'processUploadQueue') {
      try { ScriptApp.deleteTrigger(t); } catch (e4) {}
    }
  });
}

function markUploadFailureNotice_(ss) {
  var sh = ss.getSheetByName('_ErrorLog_');
  if (sh) sh.getRange(1, 1).setValue('⚠️ 要確認：ファイルアップロード失敗（詳細は下の行を確認）');
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
function setupOneCompany_(companyName, adminEmail, suppressEmail, skipEditor) {
  // 運行管理_アーカイブ/ を作成or取得
  var rootFolder    = getOrCreateFolder_('運行管理_アーカイブ');
  // 会社名サブフォルダを作成or取得
  var subIter       = rootFolder.getFoldersByName(companyName);
  var companyFolder = subIter.hasNext() ? subIter.next() : rootFolder.createFolder(companyName);

  // 管理Gmailに編集権限を付与（フォルダ）：skipEditor=trueなら同意後に行う
  if (!skipEditor) { try { companyFolder.addEditor(adminEmail); } catch(e) {} }

  var folderUrl = companyFolder.getUrl();
  var folderId  = companyFolder.getId();

  // processNewCompany_ から呼ばれた場合は個別メールを送らない
  // （sendDistributionMail_ で1通にまとめて送信するため）
  if (!suppressEmail) {
    var subject = '[運行管理] ' + companyName + ' のご利用準備が整いました';
    var body =
      companyName + ' ご担当者様\n\n' +
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

  sheet.setColumnWidth(1, 16);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 460);
  sheet.setColumnWidth(4, 16);

  var sheetId = sheet.getSheetId();
  var TOC_ROWS = 12;
  var row = TOC_ROWS + 1;
  var sections = [];

  function mainTitle(text) {
    sheet.getRange(row,1,1,4).merge().setValue(text)
      .setBackground('#0f1923').setFontColor('#ffffff')
      .setFontSize(16).setFontWeight('bold')
      .setVerticalAlignment('middle').setHorizontalAlignment('center');
    sheet.setRowHeight(row, 48); row++;
  }
  function title(text) {
    sections.push({ label: text.replace(/^[　\s]+|[　\s]+$/g, ''), row: row });
    sheet.getRange(row,1,1,4).merge().setValue(text)
      .setBackground('#1a2f3d').setFontColor('#ffffff')
      .setFontSize(13).setFontWeight('bold').setVerticalAlignment('middle');
    sheet.setRowHeight(row, 40); row++;
  }
  function section(text) {
    sheet.getRange(row,2,1,2).merge().setValue(text)
      .setBackground('#2c4356').setFontColor('#ffffff')
      .setFontSize(11).setFontWeight('bold');
    sheet.setRowHeight(row, 30); row++;
  }
  function colRow(col, name, desc, bg) {
    sheet.getRange(row,2).setValue(col + '列　' + name)
      .setBackground(bg).setFontSize(10).setFontWeight('bold')
      .setVerticalAlignment('top').setWrap(true);
    sheet.getRange(row,3).setValue(desc)
      .setFontSize(10).setVerticalAlignment('top').setWrap(true);
    sheet.setRowHeight(row, 40); row++;
  }
  function note(text) {
    var rng = sheet.getRange(row,2,1,2).merge();
    rng.setValue(text)
      .setBackground('#fdf1e0').setFontSize(10).setFontColor('#5d4037')
      .setWrap(true).setVerticalAlignment('top');
    rng.setBorder(false, true, false, false, false, false,
      '#c9a84c', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.setRowHeight(row, 36); row++;
  }
  function menuRow(menu, desc, when_) {
    sheet.getRange(row,2).setValue(menu)
      .setBackground('#eef1f3').setFontSize(10).setFontWeight('bold')
      .setVerticalAlignment('top').setWrap(true);
    sheet.getRange(row,3).setValue(desc + '\n使いどき：' + when_)
      .setFontSize(10).setVerticalAlignment('top').setWrap(true);
    sheet.setRowHeight(row, 52); row++;
  }
  function linkRow(from_, what_, when_) {
    sheet.getRange(row,2).setValue(from_)
      .setBackground('#eef1f3').setFontSize(10).setFontWeight('bold')
      .setFontColor('#0f1923').setVerticalAlignment('top').setWrap(true);
    sheet.getRange(row,3).setValue(what_ + '\nタイミング：' + when_)
      .setFontSize(10).setVerticalAlignment('top').setWrap(true);
    sheet.setRowHeight(row, 52); row++;
  }
  function sp() {
    sheet.getRange(row, 2, 1, 2)
      .setBorder(false, false, true, false, false, false,
        '#d0d5da', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(row, 10); row++;
  }

  // ── タイトル ───────────────────────────────────
  mainTitle('運行シート　使い方');
  sp();

  // ── ① 基本操作 ────────────────────────────────
  title('　① 基本操作　— どこに何を入力するか');
  sp();

  section('■ 手入力する列');
  colRow('F', '車番',
    '自車専属マスタに登録された車番を入力すると、B〜I列（区分・会社名・トン数・車種・乗務員名・携帯番号・看板名）が自動補完される\n\n▶ マスタの情報を使いたいとき → 車番（F列）を必ず最初に入力すること\n\n⚠ 自車専属マスタにない情報を入力したい場合は、意図的にB〜I列（F列の車番以外）を先に入力することで自動補完を防げます',
    '#f5f0e0');
  colRow('J', '日付',
    '日付のみ入力すると現在時刻が自動付与（例：2026/06/01 → 2026/06/01 09:23）',
    '#f5f0e0');
  colRow('K', '荷主', '手入力', '#f5f0e0');
  colRow('L', '積地',
    '手入力\n「有休」と入力 → 行全体が薄グレー（#e0e0e0）に変わる。経費0で計算\n「休み」と入力 → 行全体が濃グレー（#9e9e9e）に変わる。経費0で計算\n空欄のまま（IDあり）→ L列のみ黄色（配車漏れ警告）',
    '#f5f0e0');
  colRow('M', '降地', '手入力', '#f5f0e0');
  colRow('N', '誘導時刻',
    '手入力。どんな形式でも自動変換（「8:30」「0830」「８：３０」すべてOK）\n⚠ 時刻は必ず順番通りに入力すること（下の「時刻入力の順序」を参照）',
    '#f5f0e0');
  colRow('O', '積完時刻', '手入力（N列・誘導の入力後でないと入力できない）', '#f5f0e0');
  colRow('P', '休憩開始', '手入力（O列・積完の入力後でないと入力できない）', '#f5f0e0');
  colRow('Q', '休憩終了', '手入力（P列・休憩開始の入力後でないと入力できない）', '#f5f0e0');
  colRow('R', '降完時刻', '手入力（Q列・休憩終了の入力後でないと入力できない）', '#f5f0e0');
  colRow('S', '売上', '手入力', '#f5f0e0');
  colRow('T', '請求（高速代）',
    '手入力。入力するとU列（実費高速代）が空の場合に自動でコピーされる（オレンジ色）\n実費が請求と違う場合はU列を直接上書きしてください',
    '#f5f0e0');
  colRow('U', '実費（高速代）',
    'T列から自動コピー（オレンジ色）。実費が違う場合は上書きすると黒字に戻る\n実費が本当に0円の場合は「0」を明示的に入力すること',
    '#f5f0e0');
  colRow('W', '備考', '手入力', '#f5f0e0');
  sp();

  section('■ 自動で入る列（手入力不要）');
  colRow('A', 'ID',
    '自動採番（V-XXXX形式）。B列以降に入力が始まると自動で採番される\n手入力不要、触らない',
    '#eef1f3');
  colRow('B〜I', '区分〜看板名',
    'F列（車番）入力後に自車専属マスタから一括補完\n⚠ マスタの運行状態が「故障」「待機」の車両は補完対象外',
    '#eef1f3');
  colRow('D', 'トン数',
    '補完後も手入力でOK。「4t」「4T」「4トン」などは自動で「4」に正規化',
    '#eef1f3');
  colRow('V', '合計（高速代）',
    '「実費（U列）－請求（T列）」の数式が自動でセットされる\n手入力しない（上書きすると数式が消える）',
    '#eef1f3');
  sp();

  section('■ アプリが管理する列（触らない）');
  colRow('AA', '🔒 点呼前完了',
    'ドライバーがアプリで点呼前チェックを完了すると自動記録される',
    '#e8e8e8');
  colRow('AB', '🔒 点呼後完了',
    'ドライバーがアプリで点呼後チェックを完了すると自動記録される',
    '#e8e8e8');
  colRow('Y・Z', '🔒 連絡（端末）・データ（端末）',
    'アプリ側のデータ管理列。触らない',
    '#e8e8e8');
  colRow('X', '🔒 管理データ',
    'システム内部管理用。触らない',
    '#e8e8e8');
  sp();

  section('■ 保護されている場所');
  colRow('1行目', '🔒 ヘッダー行',
    'シート保護設定後は編集不可。誤って編集しても自動で元のヘッダーに復元される',
    '#e8e8e8');
  sp();

  note('⚠ 時刻入力の順序ルール：N誘導 → O積完 → P休憩開始 → Q休憩終了 → R降完 の順でしか入力できない。前の時刻が空のまま次を入力しようとすると自動でクリアされてトースト警告が出る');
  note('⚠ 車番の自動補完について：自車専属マスタに登録されている車両の情報を使いたい場合は、必ず車番（F列）を先に入力すること。B〜I列（F列の車番以外）を先に入力すると、同じ車番でも自動補完されない。自車専属マスタにない情報を入力したい場合は意図的にB〜I列（F列の車番以外）を先に入力することで補完を防げる');
  sp();

  // ── ② メニュー ────────────────────────────────
  title('　② 運行シートで使うメニュー');
  sp();
  section('■ 🏠 毎日の配車業務');
  menuRow('🚚 配車ダッシュボード',
    '積地が未入力の行を一覧表示',
    '配車漏れをチェックしたいとき');
  menuRow('📏 距離計算（未計算分）',
    '積地・降地が入っている行の中で距離が未計算のものを一括計算（Google Maps使用）\n毎日0時に未計算分が自動でも実行されます（SSを開いていなくても動く）',
    '距離が入っていない行があるとき・すぐに計算したいとき');
  menuRow('📍 住所確認（確認待ち分）',
    '地名があいまいで距離計算が止まっている行を確認・修正',
    '距離計算が通らない行があるとき');
  menuRow('🆔 ID・車番一括補完',
    'IDや車番が抜けている行を手動で一括補完',
    'まとめて補完したいとき');
  menuRow('🔃 日付順並び替え',
    '手動で日付順に強制並び替え（日付変更時は自動で動く）',
    '並び順がずれたとき');
  menuRow('💴 経費自動入力',
    'トン数別の平均値テーブルをもとに経費を自動入力',
    '経費をまとめて入れたいとき');
  sp();
  section('■ 📥 データ読み込み（CSV）');
  menuRow('🚛 運行シート',
    'CSV・ExcelのデータをA列以降に一括取込。列マッピングと別名（エイリアス）を保存できる',
    '配車会社からのデータを使うとき');
  sp();
  section('■ 📨 帳票・送信メニュー');
  menuRow('① 発注書・指示書（協力会社・乗務員用）',
    '協力会社・乗務員への指示書を作成してメール/FAX送信',
    '発注・指示書を送るとき');
  menuRow('② 車番連絡（荷主用）',
    '荷主への車番連絡を作成してメール/FAX送信',
    '車番が確定したとき');
  sp();
  section('■ 🗓 月またぎ処理');
  menuRow('📅 翌月分生成（前月アーカイブ）',
    '翌月の運行行を一括生成し、当月分をアーカイブ',
    '月末（25日頃）');
  menuRow('📆 今月分生成（途中契約）',
    '月途中から利用開始するときに今月分の運行行を生成',
    '月途中から使い始めるとき');
  sp();

  // ── ③ 他シートとの連携 ─────────────────────────
  title('　③ 他シートとの連携');
  sp();
  section('■ この運行シートに情報が来る');
  linkRow('自車専属マスタ',
    'F列（車番）をトリガーに、区分・会社名・トン数・車種・乗務員名・携帯番号・看板名（B〜I列）を自動補完',
    'F列（車番）に入力したとき');
  linkRow('配車板',
    'チェックした行が「車番・日付・荷主」などとセットで運行シートに追加される',
    '配車板で「配車確定」を実行したとき');
  linkRow('ドライバーアプリ',
    '誘導〜降完の時刻（N〜R列）、点呼前後完了（AA・AB列）が自動記録される',
    'ドライバーがアプリで各操作をしたとき');
  sp();
  section('■ 運行シートから情報が飛ぶ');
  linkRow('集計表',
    'ID単位の運行データ（売上・高速代・距離等）が自動同期される\n集計表では経費・利益・給料の計算に使われる',
    '運行シートを編集したとき（自動）');
  sp();

  // フッター
  sheet.getRange(row,1,1,4).merge()
    .setValue('⚙️ システム設定・保守 →「📖 使い方シート作成」で最新版に更新できます')
    .setBackground('#0f1923').setFontColor('#c9a84c').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(row, 28);

  // ── 目次 ────────────────────────────────────────
  // Row 1: ヘッダー
  sheet.getRange(1,1,1,4).merge().setValue('📋　目次')
    .setBackground('#0f1923').setFontColor('#ffffff')
    .setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(1, 36);
  // Row 2: ガイド文
  sheet.getRange(2,2,1,2).merge()
    .setValue('🔍 このシートを初めて開いた方へ：下の項目をクリックすると該当箇所へ直接移動できます')
    .setBackground('#f5f0e0').setFontColor('#0f1923').setFontSize(10)
    .setVerticalAlignment('middle').setWrap(true);
  sheet.setRowHeight(2, 30);
  // Row 3: 小スペーサー
  sheet.setRowHeight(3, 4);
  // Row 4: 親ラベル（非リンク）
  sheet.getRange(4,2,1,2).merge()
    .setValue('運行シート 使い方')
    .setBackground('#1a2f3d').setFontColor('#c9a84c')
    .setFontSize(11).setFontWeight('bold').setVerticalAlignment('middle');
  sheet.setRowHeight(4, 28);
  // Row 5-7: サブリンク（インデント付き）
  for (var ti = 0; ti < sections.length; ti++) {
    var tr = ti + 5;
    sheet.getRange(tr,2,1,2).merge()
      .setFormula('=HYPERLINK("#gid=' + sheetId + '&range=A' + sections[ti].row + '","　　▷  ' + sections[ti].label + '")')
      .setFontColor('#0f1923').setFontSize(11)
      .setBackground('#f8f9fa').setVerticalAlignment('middle');
    sheet.setRowHeight(tr, 30);
  }
  // Row 8-12: パディング
  for (var ri = 8; ri <= 12; ri++) { sheet.setRowHeight(ri, 8); }

  sheet.setFrozenRows(0);
  sheet.setTabColor('#c9a84c');
  ui.alert('「使い方」シートを作成しました。');
}


// ================================================================
//  11-7: 詳細説明書シート作成（createManualSheet）  【大C / 中11 / 小11-7】
// ================================================================
function createManualSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName('説明書');
  if (sheet) { ss.deleteSheet(sheet); }
  sheet = ss.insertSheet('説明書');
  ss.setActiveSheet(sheet);

  sheet.setColumnWidth(1, 16);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 440);
  sheet.setColumnWidth(4, 16);

  var sheetId = sheet.getSheetId();
  var TOC_ROWS = 14;
  var row = TOC_ROWS + 1;
  var sections = [];

  function mainTitle(text) {
    sheet.getRange(row,1,1,4).merge().setValue(text)
      .setBackground('#0f1923').setFontColor('#ffffff')
      .setFontSize(16).setFontWeight('bold').setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(row, 48); row++;
  }
  function title(text) {
    sections.push({ label: text.replace(/^[　\s]+|[　\s]+$/g, ''), row: row });
    sheet.getRange(row,1,1,4).merge().setValue(text)
      .setBackground('#1a2f3d').setFontColor('#ffffff')
      .setFontSize(13).setFontWeight('bold').setVerticalAlignment('middle');
    sheet.setRowHeight(row, 40); row++;
  }
  function section(text) {
    sheet.getRange(row,2,1,2).merge().setValue('▶ '+text)
      .setBackground('#2c4356').setFontColor('#ffffff')
      .setFontSize(11).setFontWeight('bold');
    sheet.setRowHeight(row, 30); row++;
  }
  function item(label, val, bg) {
    sheet.getRange(row,2).setValue(label)
      .setBackground('#eef1f3').setFontSize(11).setFontWeight('bold')
      .setVerticalAlignment('top').setWrap(true);
    sheet.getRange(row,3).setValue(val)
      .setFontSize(11).setVerticalAlignment('top').setWrap(true)
      .setBackground('#ffffff');
    sheet.setRowHeight(row, 56); row++;
  }
  function feature(num, name, desc) {
    sheet.getRange(row,2).setValue(num + '  ' + name)
      .setBackground('#eef1f3').setFontSize(10).setFontWeight('bold')
      .setVerticalAlignment('top').setWrap(true);
    sheet.getRange(row,3).setValue(desc)
      .setFontSize(10).setVerticalAlignment('top').setWrap(true);
    sheet.setRowHeight(row, 44); row++;
  }
  function catHead(text) {
    sheet.getRange(row,2,1,2).merge().setValue(text)
      .setBackground('#eef1f3').setFontSize(10).setFontWeight('bold')
      .setFontColor('#0f1923');
    sheet.setRowHeight(row, 26); row++;
  }
  function warn(text) {
    var rng = sheet.getRange(row,2,1,2).merge();
    rng.setValue('⚠ '+text)
      .setBackground('#fdf1e0').setFontSize(10).setFontColor('#5d4037')
      .setWrap(true).setVerticalAlignment('top');
    rng.setBorder(false, true, false, false, false, false,
      '#c9a84c', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.setRowHeight(row, 36); row++;
  }
  function sp() {
    sheet.getRange(row, 2, 1, 2)
      .setBorder(false, false, true, false, false, false,
        '#d0d5da', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(row, 10); row++;
  }

  mainTitle('運行管理システム　説明書');
  sp();

  // ── 1. このシステムでできること ──────────────
  title('　1. このシステムでできること');
  sp();
  section('概要');
  item('システムの説明',
    '貨物運送会社の運行管理業務をGoogleスプレッドシートで一元管理するシステムです\n\n' +
    '・配車・運行データの入力と管理\n' +
    '・乗務員のスマートフォンから運行状況をリアルタイム入力\n' +
    '・集計・請求書・支払確認書・損益計算書（PL）の自動生成\n' +
    '・発注書・指示書・車番連絡の作成とメール/FAX送信\n' +
    '・改善基準告示コンプライアンス確認表（監査用表）の自動生成\n' +
    '・毎日のバックアップ自動保存（30日分）');
  item('動作環境',
    'スプレッドシート（管理者）：PCまたはタブレット推奨\n' +
    'ドライバーアプリ：スマートフォンのブラウザで利用（インストール不要）\n' +
    'インターネット接続が必要です');
  sp();

  // ── 2. ご利用開始時に受け取るもの ─────────────
  title('　2. ご利用開始時に受け取るもの');
  sp();
  section('ご案内メールに記載された内容');
  item('運行管理スプレッドシート（管理者用）',
    '管理者が直接データを入力・確認するGoogleスプレッドシートのURLです\n配車・集計・帳票生成など全ての管理業務をここで行います\n推奨端末：PC・タブレット');
  item('管理画面アプリURL（管理者・配車係用）',
    '管理者・配車係がアプリ画面から操作できる管理ツールのURLです\nシート閲覧・編集・帳票生成・月次処理などが行えます');
  item('乗務員アプリURL（ドライバー用）',
    '各ドライバーに配布するアプリのURLです\nスマートフォンのブラウザから利用し、運行記録（誘導・積完・休憩・降完）を入力します\nインストール不要');
  sp();

  section('初回セットアップ手順（管理者の作業）');
  item('STEP 1　初期設定の実行',
    'スプレッドシートを開く → メニュー →「⚙️ システム設定・保守」→「🔧 初期設定」\n→ 列の自動保護・編集検知トリガーが有効になります\n⚠ これを押すまで自動復元・一部の自動機能が動作しません');
  item('STEP 2　シート保護の設定',
    'メニュー →「⚙️ システム設定・保守」→「🛡 シート保護設定」\n→ ヘッダー行（1行目）が編集不可になります（項目名の誤削除防止）');
  item('STEP 3　自社情報の入力',
    '「自社設定」シートに会社名・住所・電話番号・代表者名を入力してください\n→ 請求書・支払確認書の自社情報欄に自動で反映されます');
  item('STEP 4　マスタの登録',
    '「自車専属マスタ」：自社の車両・乗務員情報を入力\n「マスタ（取引先）」：荷主・協力会社の情報とメールアドレスを入力\n「設定」シート：トン数別の燃費・点呼チェックリストを確認・修正');
  warn('STEP 1・2 は最初の1回のみ実行してください。新しいスプレッドシートを開くたびに1回ずつ実行してください');
  sp();

  // ── 3. 各シートの役割 ─────────────────────────
  title('　3. 各シートの役割');
  sp();
  section('メインシート');
  item('配車板',
    '荷物情報（左側・青ヘッダー）と車両情報（右側・橙ヘッダー）を1行に入力し、配車を確定するシートです\nA列のチェックを入れて「配車確定」を押すと運行シートに自動登録されます');
  item('運行シート',
    '全ての運行記録を管理するメインシートです\n1つの運行ID（V-XXXX）に対して、行程数分の行が紐づきます\n入力するだけで多くの処理が自動で行われます（詳細は機能一覧を参照）');
  item('集計表',
    '運行シートをID単位で集約し、経費・利益・給料計算を行うシートです\n運行シートを更新すると自動で同期されます\n手入力できる列（黄色枠線）：距離・ガソリン代・支払い・備考・仮日数・その他手当');
  sp();
  section('マスタ・設定シート');
  item('自車専属マスタ',
    '自社所有車両・乗務員の情報を管理します\nB列（運行状態）が「運行」の行のみ翌月分生成の対象になります\n免許証・健康診断などの期限もここで管理します（期限が近づくとSS起動時に警告）');
  item('マスタ（取引先）',
    '荷主・協力会社の情報を管理します\n請求書・帳票の送付先メール/FAXアドレスもここに登録します');
  item('設定シート',
    'トン数別の燃費・有休設定・アプリの点呼チェックリスト項目を管理します\nE列（業務前点検）・F列（業務後点検）を変更するとアプリの点呼項目も変わります\n会社ごとにカスタマイズできます');
  item('自社設定シート',
    '会社名・住所・電話番号・代表者名を登録します\n→ 請求書・支払確認書の会社情報欄に自動反映されます');
  sp();

  // ── 4. 機能一覧（26機能） ──────────────────────
  title('　4. 機能一覧（精査済み 26機能）');
  sp();
  catHead('【入力補助】セルに入力するだけで自動処理される');
  feature('1', 'ID自動採番',
    'B列以降に入力するとA列のID（V-XXXX形式）が自動採番されます\n排他ロックで同時アクセス時の重複を防止');
  feature('2', 'トン数正規化',
    'D列に「4t」「4T」「4トン」等で入力しても「4」に自動変換');
  feature('3', '車番自動補完',
    'F列（車番）を入力すると自車専属マスタと照合して区分・会社名・乗務員名など8項目を一括補完');
  feature('4', '日付への現在時刻付与',
    'J列（日付）に日付だけ入力すると現在時刻が自動付与');
  feature('5', '時刻の正規化と日付合成',
    'N〜R列（誘導〜降完時刻）と点呼前後完了列の入力を統一フォーマットに自動変換');
  sp();
  catHead('【自動計算】入力値から自動計算・同期');
  feature('6', '書式自動設定',
    '売上〜合計高速（S〜V列）に金額書式（カンマ区切り）、時刻列に時刻書式を自動設定');
  feature('7', '請求高速→実費高速の自動コピー',
    'T列（請求高速代）を入力するとU列（実費高速代）が空の場合に自動コピー（オレンジ文字）\n実費が違う場合はU列を直接上書きしてください');
  feature('8', '合計高速代の数式自動セット',
    'V列（合計高速代）に「実費−請求」の数式を自動セット（会社の実質負担分）');
  feature('9', '日付変更時の自動ソート',
    'J列（日付）を変更すると運行シート・集計表が自動で日付順に並び替わる');
  feature('10', '積地・降地→距離自動計算',
    'L列（積地）またはM列（降地）を入力すると積地〜降地の距離をGoogle Mapsで自動計算\n毎日0時に未計算分を自動一括計算（SSを開いていなくても実行）\nAPIの一時障害は最大3回まで自動リトライ');
  feature('11', '集計表への自動同期',
    '運行シートを編集すると集計表に自動で反映（ID単位で随時同期）');
  sp();
  catHead('【警告・保護】異常の検知と自動修復');
  feature('12', '時刻入力の順序チェック',
    '誘導→積完→休憩開始→休憩終了→降完の順序が前後したら入力をキャンセルしてトースト警告');
  feature('13', '資格期限の警告ポップアップ（SS起動時）',
    'SS起動時に免許証・安全教育・健康診断・適性診断の期限切れ・60日以内をアラートで通知');
  feature('14', 'ID重複・車番不一致の警告',
    '同じIDで車番または日付が食い違う行をA列赤（#ff1744）でマーク');
  feature('15', 'ヘッダー行（1行目）の自動復元',
    '1行目を誤って編集すると正規ヘッダーに自動復元。必須シートが削除された場合も即時バックアップから復元');
  sp();
  catHead('【色表示】セルや行の色で状態を視覚化');
  feature('16', '有休行の色付け（薄グレー）',
    '積地列に「有休」と入力すると行全体が薄グレー（#e0e0e0）に変わる');
  feature('17', '休み行の色付け（濃グレー）',
    '積地列に「休み」と入力すると行全体が濃グレー（#9e9e9e）に変わる');
  feature('18', '資格期限アラート色（A列）',
    'SS起動時に自車専属マスタの4期限を参照してA列を着色\n淡い赤=期限切れ / 淡い青=当日 / 淡い緑=7日以内');
  feature('19', 'A1セルへの色凡例メモ',
    'A1セルにセルの色の意味を自動でメモ書き（期限アラート色の説明）');
  sp();
  catHead('【メニュー操作】手動で実行するメニュー項目');
  feature('20', 'ID・車番の一括補完（手動）',
    'メニュー→「ID・車番一括補完」でIDや車番が抜けている行を手動で一括補完');
  feature('21', '手動の日付順並び替え',
    'メニュー→「日付順並び替え」で強制ソート実行');
  sp();
  catHead('【帳票・メール送信】書類の作成と送信');
  feature('22', 'ファイル添付（集計・請求データ）',
    '集計・請求データに画像やPDFなどのファイルを添付する');
  feature('23', '発注書・指示書のメール送信',
    'メニュー→「発注書・指示書送信」で指示書をメール/FAX送信');
  feature('24', '車番連絡のメール送信',
    'メニュー→「車番連絡送信」で荷主に車番連絡をメール/FAX送信');
  sp();
  catHead('【データ取込・配車確定】外部データの取込と配車確定');
  feature('25', '配車表CSV/Excelの取込',
    'メニュー→「CSVをExcelに」で配車会社からのCSV・Excelを運行シートに一括取込\n列マッピングと別名（エイリアス）を保存できる');
  feature('26', '配車板からの配車確定',
    '配車板でチェックした行を運行シートに一括登録');
  sp();

  // ── 5. バックアップ・データ保護 ──────────────
  title('　5. バックアップとデータ保護');
  sp();
  section('自動バックアップ（毎日3時）');
  item('バックアップの仕組み',
    '【① 外部バックアップ（毎日深夜3時）】\nGoogleドライブへ自動保存\n保存場所：マイドライブ →「運行管理バックアップ」→「会社名」フォルダ\nファイル名：yyyy-MM-dd_会社名　最大30日分を自動保持\n\n【② 内部バックアップ（SS起動時・24時間に1回）】\nSS内の _BK_ シートを最新状態に更新\n同日中に何度起動しても1回のみ実行（Drive容量の急増を防止）');
  item('データ復旧の手順',
    'メニュー →「🔄 バックアップ・復旧」→「🔄 バックアップから復旧」\n→ 復旧日を選択して実行\n⚠ 現在のデータは上書きされます');
  item('その他の復旧方法',
    '① Googleスプレッドシートの「バージョン履歴」（ファイル → バージョン履歴）\n② 自動生成される_BK_シートからのコピー（SS内に常時保存）');
  sp();

  // ── 6. 法令対応 ───────────────────────────────
  title('　6. 法令対応の範囲（点呼記録について）');
  sp();
  section('監査用表と点呼記録の区別');
  item('監査用表の位置づけ',
    '「監査用表」は「業務管理・改善基準告示コンプライアンス確認表」です\n法令上の点呼記録（輸送安全規則第7条）そのものではありません\n行政監査・労基署対応では「監査用表＋点呼記録（アプリデータ）」の2点セットで対応してください');
  item('アプリの点呼機能で対応済みの項目',
    '・アルコールチェックの実施（業務前・帰庫後）\n・日常点検（ブレーキ・タイヤ・エンジンオイル等）\n・点呼の実施確認');
  item('記録の保管義務',
    '乗務記録・点呼記録ともに1年間の保管義務があります\nアプリに記録されたデータはスプレッドシートに自動保存されます');
  warn('法令上は点呼記録に「点呼者名・点呼方法（対面/電話/IT）」も必要です。小規模会社では実務上認められるケースが多いですが、厳密には現バージョンでは未対応です。');
  sp();

  // フッター
  sheet.getRange(row,1,1,4).merge()
    .setValue('メニュー →「⚙️ システム設定・保守」→「📘 説明書作成」で最新版に更新できます')
    .setBackground('#0f1923').setFontColor('#c9a84c').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(row, 28);

  // ─── 目次を先頭に書き込む ─────────────────────
  sheet.getRange(1,1,1,4).merge().setValue('📋  目次（クリックで各セクションへ移動）')
    .setBackground('#0f1923').setFontColor('#ffffff')
    .setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(1, 36);
  sheet.getRange(2,2,1,2).merge()
    .setValue('🔍 このシートを初めて開いた方へ：下の項目をクリックすると該当箇所へ直接移動できます')
    .setBackground('#f5f0e0').setFontColor('#0f1923').setFontSize(10)
    .setVerticalAlignment('middle').setWrap(true);
  sheet.setRowHeight(2, 30);
  sheet.setRowHeight(3, 4);
  sheet.getRange(4,2,1,2).merge()
    .setValue('運行管理システム 説明書')
    .setBackground('#1a2f3d').setFontColor('#c9a84c')
    .setFontSize(11).setFontWeight('bold').setVerticalAlignment('middle');
  sheet.setRowHeight(4, 28);
  for (var ti = 0; ti < sections.length; ti++) {
    var tr = ti + 5;
    if (tr > TOC_ROWS - 1) break;
    sheet.getRange(tr,2,1,2).merge()
      .setFormula('=HYPERLINK("#gid=' + sheetId + '&range=A' + sections[ti].row + '","　　▷  ' + sections[ti].label + '")')
      .setFontColor('#0f1923').setFontSize(11)
      .setBackground('#f8f9fa').setVerticalAlignment('middle');
    sheet.setRowHeight(tr, 30);
  }
  for (var ri = sections.length + 5; ri <= TOC_ROWS; ri++) {
    sheet.setRowHeight(ri, 8);
  }

  sheet.setFrozenRows(0);
  sheet.setTabColor('#c9a84c');
  ui.alert('「説明書」シートを作成しました。');
}


// ================================================================
//  11-8: サポート返信テンプレートシート作成（createSupportSheet）  【大C / 中11 / 小11-8】
// ================================================================
function createSupportSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName('サポート');
  if (sheet) { ss.deleteSheet(sheet); }
  sheet = ss.insertSheet('サポート');
  ss.setActiveSheet(sheet);

  sheet.setColumnWidth(1, 16);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 440);
  sheet.setColumnWidth(4, 16);

  var sheetId = sheet.getSheetId();
  var TOC_ROWS = 14;
  var row = TOC_ROWS + 1;
  var sections = [];

  function mainTitle(text) {
    sheet.getRange(row,1,1,4).merge().setValue(text)
      .setBackground('#0f1923').setFontColor('#ffffff')
      .setFontSize(16).setFontWeight('bold').setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(row, 48); row++;
  }
  function title(text) {
    sections.push({ label: text.replace(/^[　\s]+|[　\s]+$/g, ''), row: row });
    sheet.getRange(row,1,1,4).merge().setValue(text)
      .setBackground('#1a2f3d').setFontColor('#ffffff')
      .setFontSize(13).setFontWeight('bold').setVerticalAlignment('middle');
    sheet.setRowHeight(row, 40); row++;
  }
  function section(text) {
    sheet.getRange(row,2,1,2).merge().setValue('▶ '+text)
      .setBackground('#2c4356').setFontColor('#ffffff')
      .setFontSize(11).setFontWeight('bold');
    sheet.setRowHeight(row, 30); row++;
  }
  function qa(q, cause, sol, bg) {
    var label = 'Q. ' + q + '\n原因：' + cause;
    var answer = '対処：' + sol;
    sheet.getRange(row,2).setValue(label)
      .setBackground('#eef1f3').setFontSize(10).setFontWeight('bold')
      .setVerticalAlignment('top').setWrap(true);
    sheet.getRange(row,3).setValue(answer)
      .setFontSize(10).setVerticalAlignment('top').setWrap(true);
    sheet.setRowHeight(row, 72); row++;
  }
  function tmpl(label, content, bg) {
    sheet.getRange(row,2).setValue(label)
      .setBackground('#eef1f3').setFontSize(10).setFontWeight('bold')
      .setVerticalAlignment('top').setWrap(true);
    sheet.getRange(row,3).setValue(content)
      .setFontSize(10).setVerticalAlignment('top').setWrap(true).setFontFamily('Courier New');
    sheet.setRowHeight(row, 90); row++;
  }
  function tip(text) {
    sheet.getRange(row,2,1,2).merge().setValue('💡 ' + text)
      .setBackground('#e8f0e8').setFontSize(10).setFontColor('#2e5c2e').setWrap(true);
    sheet.setRowHeight(row, 28); row++;
  }
  function warn(text) {
    var rng = sheet.getRange(row,2,1,2).merge();
    rng.setValue('⚠ ' + text)
      .setBackground('#fdf1e0').setFontSize(10).setFontColor('#5d4037')
      .setWrap(true).setVerticalAlignment('top');
    rng.setBorder(false, true, false, false, false, false,
      '#c9a84c', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    sheet.setRowHeight(row, 28); row++;
  }
  function sp() {
    sheet.getRange(row, 2, 1, 2)
      .setBorder(false, false, true, false, false, false,
        '#d0d5da', SpreadsheetApp.BorderStyle.SOLID);
    sheet.setRowHeight(row, 10); row++;
  }

  mainTitle('サポートガイド　よくある質問と対処法');
  sp();

  // ── まず試すこと（鉄板3手順） ──────────────
  title('　まず試すこと（鉄板3手順）');
  sp();
  section('何かおかしいと思ったらまずこれを試す');
  qa('メニューが消えた',
    'ページ更新するだけで戻ることがほとんど',
    '① F5キー（またはブラウザの更新ボタン）でページを更新する\n② それでも出ない場合：メニュー →「⚙️ システム設定・保守」→「🔄 メニュー再生成」',
    '#e8eaf6');
  qa('集計の数字がおかしい',
    '運行シートの更新が集計表に反映されていない',
    'メニュー →「⚙️ システム設定・保守」→「📃 集計表再生成」を実行',
    '#e8eaf6');
  qa('シートの列がズレた・消えた',
    '誤操作で列を移動・削除した',
    'メニュー →「⚙️ システム設定・保守」→「🗂 シート再生成」を実行\nデータはそのままで列の並びが元に戻ります',
    '#e8eaf6');
  tip('上記3手順で解決するケースが8割以上です。まずここから試してください');
  sp();

  // ── アプリ系トラブル ─────────────────────────
  title('　アプリ系トラブル');
  sp();
  section('アプリが開かない・表示がおかしい');
  qa('アプリが開かない・真っ白',
    'ブラウザキャッシュ・URLの開き方・通信状態の問題',
    '① ホーム画面のアイコンから開いているか確認（URLを直打ちしない）\n② アプリを完全に閉じて再起動\n③ Wi-Fi・モバイル通信を切り替えて再試行\n④ ブラウザのキャッシュを削除してから再起動',
    '#bbdefb');
  qa('ログインできない・別の会社の画面が出る',
    'Googleアカウントが複数登録されている・URLが古い',
    '① 登録したGmailアドレスのGoogleアカウントに切り替えてから再度開く\n② 最新のアプリURLでホーム画面に追加し直す\n③ アドレス紐づけを再実行（紐づけ画面にGmailアドレスを入力）',
    '#bbdefb');
  qa('入力したデータが画面に反映されない',
    '通信遅延・画面の更新が必要',
    '① 画面を下に引っ張ってリフレッシュ\n② 2〜3分待ってから再確認（データは正常に保存されていることがほとんど）\n③ それでも表示されない場合はスプレッドシート側で直接確認',
    '#bbdefb');
  qa('アプリを途中で閉じてしまった',
    '途中で閉じてもデータは保存済み',
    '同じURLを開き直せば途中から再開できます\n一覧画面で「運行再開」ボタンを押してください',
    '#bbdefb');
  qa('入力を間違えた',
    '修正機能あり',
    '一覧画面 →「編集」ボタン → 積地・降地・売上などを修正して保存',
    '#bbdefb');
  qa('アプリで「処理に失敗しました」とエラーが出た',
    '通信状況の問題または一時的なサーバーエラー',
    '① Wi-Fi・モバイル通信を切り替えて再試行\n② アプリを閉じて再起動し、同じ操作を再度実行\n③ ファイル送信の場合、10MB以内のファイルか確認（超過するとエラー）\n④ 繰り返す場合は操作内容とエラー文を開発者に連絡',
    '#bbdefb');
  sp();

  // ── SS系トラブル ─────────────────────────────
  title('　スプレッドシート系トラブル');
  sp();
  section('シートの表示・操作系');
  qa('シートを誤って削除してしまった',
    '必須シートは自動復元される',
    '① 1〜3秒待つ（自動で復元されます）\n② 復元されない場合：メニュー →「🔄 バックアップ・復旧」→「🔄 バックアップから復旧」\n③ どうしても解決しない場合は開発者に連絡（スプレッドシートへのアクセス権をください）',
    '#e1bee7');
  qa('運行シートの色がおかしい（有休なのに赤い等）',
    '色の優先順位の処理タイミングがズレた',
    'F5キーでページを更新するか\nメニュー →「⚙️ システム設定・保守」→「🗂 シート再生成」を実行',
    '#e1bee7');
  qa('A列（ID）が赤くなっている行がある',
    '同じIDで車番または日付が食い違っている（ID衝突）',
    '赤くなっている行のID・車番・日付を確認して正しい値に修正してください\n修正すると赤色が消えます',
    '#e1bee7');
  qa('日付を変更しても並び替わらない',
    '初期設定のトリガーが未設定',
    'メニュー →「⚙️ システム設定・保守」→「🔧 初期設定」を実行してください\nその後 F5 で更新すると自動ソートが有効になります',
    '#e1bee7');
  sp();
  section('入力・計算系');
  qa('車番を入力しても乗務員名が補完されない',
    '自車専属マスタとの部分一致が見つからない・車番の表記が違う',
    '① 自車専属マスタのF列（車番）と入力した車番の表記が一致しているか確認\n② マスタのB列（運行状態）が「故障」または「待機」になっていないか確認',
    '#e1bee7');
  qa('高速代を入力したが実費と請求が違うのに同じ値になる',
    '自動コピー機能の動作',
    'T列（請求高速）に入力すると U列（実費高速）が空の場合は自動コピーされます\n実費が違う場合は U列を直接上書きしてください\n実費が本当に0円の場合は U列に「0」を入力してください',
    '#e1bee7');
  qa('距離が計算されない行がある',
    '降完時刻が空・地名が一意に特定できない',
    '① 降完時刻（R列）が入力されているか確認\n② 積地・降地の地名をより具体的に入力してから「距離計算（未計算分）」を実行\n③ 毎日0時に未計算分が自動実行されます。翌朝確認でOKなら放置でも可',
    '#e1bee7');
  sp();

  // ── メール・帳票系トラブル ────────────────────
  title('　メール・帳票系トラブル');
  sp();
  section('メール送受信のトラブル');
  qa('送信したメールが相手に届かない',
    '迷惑メールフォルダ振り分け・メールアドレスの登録ミス',
    '① 受信者に迷惑メールフォルダを確認してもらう\n② マスタ（取引先）シートのメールアドレスに誤りがないか確認\n③ テスト送信として自分のアドレスに送って正常に届くか確認',
    '#ffe0b2');
  qa('送信ボタンを押してもエラーになる',
    'メールアドレス未登録・Gmail送信上限超過',
    '① マスタ（取引先）シートにメールアドレスが登録されているか確認\n② 1日の送信上限（Gmailは500通/日）に達していないか確認\n③ エラーメッセージの内容を開発者に転送してください',
    '#ffe0b2');
  qa('帳票の内容（会社名・住所など）が違う',
    '自社設定シートの情報が古い',
    '「自社設定」シートの会社名・住所・電話番号・代表者名を正しい内容に更新してください',
    '#ffe0b2');
  sp();

  // ── データ復旧 ───────────────────────────────
  title('　データ復旧の手順（3段階）');
  sp();
  section('状況に応じた復旧方法');
  qa('特定の行を間違えて削除・変更した',
    '操作直後なら最速はバージョン履歴',
    '① ファイル →「バージョン履歴」→「バージョン履歴を表示」\n→ 変更前のバージョンを選んで「この版を復元」\n② 毎日3時の自動バックアップを開いて該当行をコピペ',
    '#e0f7fa');
  qa('シート全体のデータが壊れた',
    'バックアップからの一括復旧が確実',
    'メニュー →「🔄 バックアップ・復旧」→「🔄 バックアップから復旧」\n→ 会社を選んで復旧日を選択して実行\n⚠ 現在のデータは上書きされます',
    '#e0f7fa');
  qa('シートのタブ（シート自体）が消えた',
    '自動復元または_BK_シートからのコピー',
    '① 1〜3秒待つ（自動復元される）\n② 「非表示のシートを表示」から _BK_シートを確認してコピー\n③ バックアップから復旧',
    '#e0f7fa');
  warn('復旧操作は現在のデータを上書きします。必ず複数の方法を把握した上で実行してください');
  sp();

  // ── 初期設定関連 ─────────────────────────────
  title('　初期設定・セットアップ関連');
  sp();
  section('初回セットアップの確認');
  qa('初期設定を押したことがない',
    '初期設定は新規SS開設時に1回必須',
    'メニュー →「⚙️ システム設定・保守」→「🔧 初期設定」を実行\nその後 F5 で更新してください',
    '#cfd8dc');
  qa('シート保護設定を押したことがない',
    'ヘッダー行が保護されない',
    'メニュー →「⚙️ システム設定・保守」→「🛡 シート保護設定」を実行',
    '#cfd8dc');
  qa('乗務員アプリのURLを紛失した',
    '最初のご案内メールに記載されていた',
    '① 最初に届いた案内メールを確認\n② スプレッドシートを開いた状態で管理者に問い合わせ（URLの再送が可能）',
    '#cfd8dc');
  sp();

  // ── 問い合わせ返信テンプレート ───────────────
  title('　問い合わせ返信テンプレート');
  sp();
  section('基本テンプレート');
  tmpl('📬 受付完了返信',
    '件名：【受付完了】○○について\n\n○○様\n\nお問い合わせいただきありがとうございます。\nご連絡内容を確認のうえ、順次対応いたします。\n解決まで今しばらくお待ちいただけますでしょうか。\nよろしくお願いいたします。',
    '#c8e6c9');
  tmpl('⏳ 対応中返信',
    '件名：Re: ○○について（対応中）\n\n○○様\n\nただいま内容を確認・対応中です。\n進捗があり次第、改めてご連絡いたします。\nよろしくお願いいたします。',
    '#c8e6c9');
  tmpl('✅ 解決確認返信',
    '件名：Re: 【解決確認】○○について\n\n○○様\n\n問題が解決されたとのこと、安心いたしました。\n引き続きご不明な点がございましたら、いつでもご連絡ください。\nよろしくお願いいたします。',
    '#c8e6c9');
  tmpl('🚨 エスカレーション（現地確認が必要な場合）',
    '件名：Re: ○○について（確認のお願い）\n\n○○様\n\n状況をより詳しく確認させていただきたいです。\n以下の対応をお願いできますでしょうか。\n\n① スプレッドシートへの編集権限を共有いただく\n（スプレッドシート画面右上「共有」→ メールアドレスを追加）\n\n権限をいただいた後、直接確認して対応いたします。\nよろしくお願いいたします。',
    '#c8e6c9');
  sp();

  section('対応チェックリスト');
  sheet.getRange(row,2,1,2).merge()
    .setValue(
      '□ 問い合わせ内容を正確に把握した\n' +
      '□ 受付完了の返信を送った（緊急なら即対応）\n' +
      '□ 「鉄板3手順」で解決できる内容か確認した\n' +
      '□ アプリ系 → 再起動・URL・アドレス紐づけを確認\n' +
      '□ シート系 → シート再生成・集計再生成を案内\n' +
      '□ メール系 → 迷惑メールフォルダ・アドレスを確認\n' +
      '□ 解決しない場合 → スプレッドシートへのアクセス権をもらって直接確認\n' +
      '□ 解決済み確認の返信を送った'
    )
    .setBackground('#f8f9fa').setFontSize(10).setWrap(true)
    .setVerticalAlignment('top');
  sheet.setRowHeight(row, 124); row++;
  sp();

  // フッター
  sheet.getRange(row,1,1,4).merge()
    .setValue('メニュー →「⚙️ システム設定・保守」→「📋 サポートテンプレ作成」で最新版に更新できます')
    .setBackground('#0f1923').setFontColor('#c9a84c').setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(row, 28);

  // ─── 目次を先頭に書き込む ─────────────────────
  sheet.getRange(1,1,1,4).merge().setValue('📋  目次（クリックで各セクションへ移動）')
    .setBackground('#0f1923').setFontColor('#ffffff')
    .setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setRowHeight(1, 36);
  sheet.getRange(2,2,1,2).merge()
    .setValue('🔍 このシートを初めて開いた方へ：下の項目をクリックすると該当箇所へ直接移動できます')
    .setBackground('#f5f0e0').setFontColor('#0f1923').setFontSize(10)
    .setVerticalAlignment('middle').setWrap(true);
  sheet.setRowHeight(2, 30);
  sheet.setRowHeight(3, 4);
  sheet.getRange(4,2,1,2).merge()
    .setValue('サポートガイド よくある質問と対処法')
    .setBackground('#1a2f3d').setFontColor('#c9a84c')
    .setFontSize(11).setFontWeight('bold').setVerticalAlignment('middle');
  sheet.setRowHeight(4, 28);
  for (var ti = 0; ti < sections.length; ti++) {
    var tr = ti + 5;
    if (tr > TOC_ROWS - 1) break;
    sheet.getRange(tr,2,1,2).merge()
      .setFormula('=HYPERLINK("#gid=' + sheetId + '&range=A' + sections[ti].row + '","　　▷  ' + sections[ti].label + '")')
      .setFontColor('#0f1923').setFontSize(11)
      .setBackground('#f8f9fa').setVerticalAlignment('middle');
    sheet.setRowHeight(tr, 30);
  }
  for (var ri = sections.length + 5; ri <= TOC_ROWS; ri++) {
    sheet.setRowHeight(ri, 8);
  }

  sheet.setFrozenRows(0);
  sheet.setTabColor('#c9a84c');
  ui.alert('「サポート」シートを作成しました。');
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
  // 管理画面（親アプリ）URL：子アプリURLに ?page=parent&ssId=【SSID】 を付与
  var ssIdForParent  = clientSsMatch ? clientSsMatch[1] : '';
  var parentAppUrl   = ssIdForParent ? appUrl + '?page=parent&ssId=' + ssIdForParent : appUrl + '?page=parent';
  var adminSent = 0, driverSent = 0;

  // ── 管理者向けメール ──────────────────────────────────
  var adminSubject = '[運行管理] ' + companyName + ' 運行管理システムのご案内';
  var adminBody =
    companyName + ' ご担当者様\n\n' +
    'このたびは運行管理システムをご利用いただきありがとうございます。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 運行管理スプレッドシート（PC・タブレット推奨）\n' +
    ssUrl + '\n' +
    '　→ 配車・運行データの直接入力・確認はこちら\n\n' +
    '■ 管理画面アプリ（配車係・管理者用）\n' +
    parentAppUrl + '\n' +
    '　→ アプリ上でシート閲覧・編集・帳票生成・月次処理が行えます\n\n' +
    '■ 乗務員アプリ（乗務員用 スマートフォン推奨）\n' +
    appUrl + '\n' +
    '　→ 乗務員がスマートフォンから運行状況を入力するアプリです\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    '【スプレッドシートの使い方】\n' +
    '・このスプレッドシートに乗務員のメールアドレスを入力してください\n' +
    '・「自車専属マスタ」タブのJ列（メールアドレス）に1名ずつ入力します\n\n' +
    '【乗務員への配布方法】\n' +
    '・各乗務員に上記「乗務員アプリ」URLを共有してください\n' +
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
//  12-3a: 新規SSのバインドスクリプトID取得（getNewSsScriptId_）  【大B / 中12 / 小12-3a】
//  DriveApp でSSの子ファイルからApps Scriptを検索してスクリプトIDを返す
//  ※ createCompanySpreadsheet_(12-3)の補助関数。12-3本体より前に配置されている
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
//  12-3a-2: 各客SS用 WebApp を自動デプロイ（deployClientWebApp_）  【大B / 中12 / 小12-3a-2】
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
        version: libVer, developmentMode: false }] },
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
    var putResp = UrlFetchApp.fetch(apiBase + '/' + scriptId + '/content', {
      method: 'PUT', headers: hdrs,
      payload: JSON.stringify({ files: [
        { name: 'appsscript', type: 'JSON',      source: manifest },
        { name: 'コード',      type: 'SERVER_JS', source: getClientStubSource_() }
      ]}),
      muteHttpExceptions: true
    });
    if (putResp.getResponseCode() !== 200) {
      return { error: 'content PUT ' + putResp.getResponseCode() + ': ' + putResp.getContentText().slice(0, 120) };
    }

    // バージョン作成（デプロイに必要）
    var verResp = UrlFetchApp.fetch(apiBase + '/' + scriptId + '/versions', {
      method: 'POST', headers: hdrs,
      payload: JSON.stringify({ description: companyName + '_auto' }),
      muteHttpExceptions: true
    });
    var verData = JSON.parse(verResp.getContentText());
    var versionNum = verData.versionNumber || null;

    // WebApp デプロイ作成
    var deployPayload = { description: companyName + '_WebApp', manifestFileName: 'appsscript' };
    if (versionNum) deployPayload.versionNumber = versionNum;
    var dr = UrlFetchApp.fetch(apiBase + '/' + scriptId + '/deployments', {
      method: 'POST', headers: hdrs,
      payload: JSON.stringify(deployPayload),
      muteHttpExceptions: true
    });
    var drData = JSON.parse(dr.getContentText());
    if (!drData.deploymentId) {
      return { error: 'deploy ' + dr.getResponseCode() + ': ' + dr.getContentText().slice(0, 120) };
    }

    var webAppUrl = 'https://script.google.com/macros/s/' + drData.deploymentId + '/exec';
    PropertiesService.getScriptProperties().setProperty('scriptId_' + ssId, scriptId);
    return { scriptId: scriptId, webAppUrl: webAppUrl };
  } catch(e) { return { error: 'EXCEPTION: ' + (e.message || String(e)) }; }
}


// ================================================================
//  P-1: 管理画面用シート一覧取得（getParentSheets）  【大A / 中P / 小P-1】
//  親アプリのホーム画面でSSの全シートをカテゴリ別に分類して返す。
//  分類に当てはまらないシートは「その他」カテゴリに追加する。
// ================================================================
function getParentSheets(companySsId) {
  var ss = getTargetSS_(companySsId);
  var sheets = ss.getSheets();
  var CATS = [
    { name: '🚚 配車・運行',   test: function(n) { return /^運行$|^集計表$|^自車専属運行$|^配車板$|主計表/.test(n); } },
    { name: '👤 自車・マスタ', test: function(n) { return /^自車専属マスタ$|^マスタ$|^距離マスタ$/.test(n); } },
    { name: '📊 PL・経理',    test: function(n) { return /^PL|^監査用$/.test(n); } },
    { name: '📋 帳票・出力',  test: function(n) { return /^仕分表$|^請求書$|^支払確認書$|^受領者_耳$/.test(n); } },
    { name: '⚙️ 設定・情報',  test: function(n) { return /^設定$|^PL設定$|^管理者$|^使い方$/.test(n); } },
  ];
  var DISPLAY = { '受領者_耳': '受領書の耳' };
  var result = CATS.map(function(c) { return { name: c.name, sheets: [] }; });
  var other = { name: '📋 その他', sheets: [] };
  sheets.forEach(function(s) {
    var name = s.getName();
    var item = { name: name, display: DISPLAY[name] || name };
    var matched = false;
    for (var i = 0; i < CATS.length; i++) {
      if (CATS[i].test(name)) { result[i].sheets.push(item); matched = true; break; }
    }
    if (!matched) other.sheets.push(item);
  });
  if (other.sheets.length > 0) result.push(other);
  return result.filter(function(c) { return c.sheets.length > 0; });
}


// ================================================================
//  P-4: 管理者紐づけ登録（linkAdminEmail）  【大A / 中P / 小P-4】
//  入力メールアドレスを管理者シートのB列と照合し、一致すればA列（名前）を取得して
//  UserPropertiesに保存する。照合失敗時はエラーを返す。
// ================================================================
function linkAdminEmail(email, companySsId) {
  if (!email || email.indexOf('@') === -1) return { ok: false, msg: 'メールアドレスが正しくありません' };
  var ss = getTargetSS_(companySsId);
  var adminSheet = ss.getSheetByName('管理者');
  if (!adminSheet) return { ok: false, msg: '管理者シートが見つかりません。シート再生成を実行してください' };
  var lastRow = adminSheet.getLastRow();
  if (lastRow < 2) return { ok: false, msg: '管理者シートに登録がありません' };
  var data = adminSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var foundName = null;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === email.trim().toLowerCase()) {
      foundName = String(data[i][0]).trim();
      break;
    }
  }
  if (foundName === null) return { ok: false, msg: '管理者シートにこのメールアドレスが登録されていません' };
  var prefix = 'admin_' + (companySsId || 'default') + '_';
  var props = PropertiesService.getUserProperties();
  props.setProperty(prefix + 'email', email.trim());
  props.setProperty(prefix + 'name',  foundName);
  return { ok: true, email: email.trim(), name: foundName };
}


// ================================================================
//  P-5: 管理者情報取得（getLinkedAdminEmail）  【大A / 中P / 小P-5】
//  親アプリ起動時にUserPropertiesから管理者の名前・メールを返す。
// ================================================================
function getLinkedAdminEmail(companySsId) {
  var prefix = 'admin_' + (companySsId || 'default') + '_';
  var props = PropertiesService.getUserProperties();
  return {
    email: props.getProperty(prefix + 'email') || '',
    name:  props.getProperty(prefix + 'name')  || '',
  };
}


// ================================================================
//  P-2: 管理画面用シートデータ取得（getSheetTableData）  【大A / 中P / 小P-2】
//  指定シートの1行目をヘッダー、2行目以降をデータとして返す。
//  最大200行まで取得（大容量シートへの負荷対策）。
// ================================================================
function getSheetTableData(sheetName, companySsId) {
  var ss = getTargetSS_(companySsId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { headers: [], rows: [], backgrounds: [], fontColors: [], sheetName: sheetName };
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { headers: [], rows: [], backgrounds: [], fontColors: [], sheetName: sheetName };
  var limit = (sheetName === '配車板') ? Math.min(lastRow, 502) : Math.min(lastRow, 201);
  var range = sheet.getRange(1, 1, limit, lastCol);
  var vals = range.getDisplayValues();
  var bgs  = range.getBackgrounds();
  var fcs  = range.getFontColors();
  var dropdowns = {};
  if (lastRow >= 2) {
    try {
      var dvRow = sheet.getRange(2, 1, 1, lastCol).getDataValidations()[0];
      dvRow.forEach(function(dv, ci) {
        if (dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          dropdowns[ci] = dv.getCriteriaValues()[0];
        }
      });
    } catch(e) {}
  }
  return { headers: vals[0], rows: vals.slice(1), backgrounds: bgs.slice(1), fontColors: fcs.slice(1), dropdowns: dropdowns, sheetName: sheetName };
}


// ================================================================
//  P-3: 管理画面用シート行保存（saveSheetRowData）  【大A / 中P / 小P-3】
//  親アプリ編集モーダルから保存。rowIndex は2以上（ヘッダー行=1を除く）。
// ================================================================
function saveSheetRowData(sheetName, rowIndex, rowData, companySsId) {
  if (rowIndex < 2) return { ok: false, msg: 'ヘッダー行は編集不可' };
  var ss = getTargetSS_(companySsId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { ok: false, msg: 'シートが見つかりません: ' + sheetName };
  sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  return { ok: true };
}


// ================================================================
//  P-6: 行追加（appendSheetRow）  【大A / 中P / 小P-6】
//  指定シートの末尾に1行追加する。
// ================================================================
function appendSheetRow(sheetName, rowData, companySsId) {
  var ss = getTargetSS_(companySsId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { ok: false, msg: 'シートが見つかりません: ' + sheetName };
  sheet.appendRow(rowData);
  return { ok: true };
}


// ================================================================
//  P-7: 行削除（deleteSheetRow）  【大A / 中P / 小P-7】
//  指定シートの指定行を削除する（ヘッダー行=1は不可）。
// ================================================================
function deleteSheetRow(sheetName, rowIndex, companySsId) {
  if (rowIndex < 2) return { ok: false, msg: 'ヘッダー行は削除できません' };
  var ss = getTargetSS_(companySsId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { ok: false, msg: 'シートが見つかりません: ' + sheetName };
  sheet.deleteRow(rowIndex);
  return { ok: true };
}


// スタブコードのソース文字列（stub_for_clientSS/コード.js から build_stub.js が自動生成）
function getClientStubSource_() {
  // === AUTO_GENERATED_STUB_START（手動編集禁止：build_stub.js が生成） ===
  return "// 客SS・テンプレートSS用スタブ（実装はライブラリ UnkouLib にある）\n// ②客用SS・③各客SS 共通。メニュー定義はライブラリ（buildClientMenu）に集約済み。\n// スタブは公開関数の転送のみ担当。反映ボタンは①修正用SSのみ。\nfunction onOpen(e) {\n  // サイレント自動トリガー再構築（FULL権限時のみ有効・LIMITED時はtry-catchで自動スキップ）\n  try {\n    var _ss0 = SpreadsheetApp.getActiveSpreadsheet();\n    var _sf = ['installedOnEdit_','onStructureChange_','checkMasterExpiries','onOpen','checkExpiryDates','calcDistanceTrigger_'];\n    ScriptApp.getUserTriggers(_ss0).forEach(function(t) {\n      if (_sf.indexOf(t.getHandlerFunction()) !== -1) { try { ScriptApp.deleteTrigger(t); } catch(ex) {} }\n    });\n    ScriptApp.newTrigger('installedOnEdit_').forSpreadsheet(_ss0).onEdit().create();\n    ScriptApp.newTrigger('onStructureChange_').forSpreadsheet(_ss0).onChange().create();\n    ScriptApp.newTrigger('calcDistanceTrigger_').timeBased().atHour(0).everyDays(1).create();\n  } catch(_ex0) {}\n  // 通常パス（LIMITED では上記は無害スキップ済み）\n  UnkouLib.buildClientMenu();\n  try { UnkouLib.convertLegacyAdminDataUrls(); } catch(e) {}\n  try { UnkouLib.applyHolidayRowColors(); } catch(e) {}\n  try {\n    var _hideSs = SpreadsheetApp.getActiveSpreadsheet();\n    ['指示先履歴', '指示先ID別'].forEach(function(n) {\n      var sh = _hideSs.getSheetByName(n);\n      if (sh && !sh.isSheetHidden()) sh.hideSheet();\n    });\n  } catch(e) {}\n  try {\n    var _epDp = PropertiesService.getDocumentProperties();\n    var _epTs = Number(_epDp.getProperty('EXPIRY_POPUP_TS') || 0);\n    if (Date.now() - _epTs >= 30000) {\n      _epDp.setProperty('EXPIRY_POPUP_TS', String(Date.now()));\n      UnkouLib.showExpiryAlert();\n    }\n  } catch(_epEx) {}\n  try { UnkouLib.applyExpiryWarningColors(); } catch(e) {}\n  try {\n    var _bkProps = PropertiesService.getDocumentProperties();\n    var _bkLast  = Number(_bkProps.getProperty('LAST_BACKUP_TS') || 0);\n    if (Date.now() - _bkLast > 24 * 60 * 60 * 1000) {\n      UnkouLib.backupAllSheets();\n      _bkProps.setProperty('LAST_BACKUP_TS', String(Date.now()));\n    }\n  } catch(e) {}\n  try {\n    var _ss2 = SpreadsheetApp.getActiveSpreadsheet();\n    var _errSh = _ss2.getSheetByName('_ErrorLog_');\n    if (_errSh) {\n      var _a1 = String(_errSh.getRange(1, 1).getValue());\n      if (_a1.indexOf('⚠️ 要確認') === 0) {\n        SpreadsheetApp.getUi().alert(_a1);\n        _errSh.getRange(1, 1).setValue('日時');\n      }\n    }\n  } catch(e) {}\n}\n\nfunction doGet(e)            { return UnkouLib.doGet(e); }\nfunction onEdit(e)           { return UnkouLib.onEdit(e); }\nfunction installedOnEdit_(e) {\n  var _FLAG = 'ZOMBIE_CLEANED_V792';\n  var _dp = PropertiesService.getDocumentProperties();\n  if (!_dp.getProperty(_FLAG)) {\n    var _lck = LockService.getDocumentLock();\n    if (!_lck.tryLock(3000)) return;\n    try {\n      if (!_dp.getProperty(_FLAG)) {\n        var _ss1 = e.source;\n        ScriptApp.getUserTriggers(_ss1).forEach(function(t) { try { ScriptApp.deleteTrigger(t); } catch(ex) {} });\n        ScriptApp.newTrigger('installedOnEdit_').forSpreadsheet(_ss1).onEdit().create();\n        ScriptApp.newTrigger('onStructureChange_').forSpreadsheet(_ss1).onChange().create();\n        ScriptApp.newTrigger('calcDistanceTrigger_').timeBased().atHour(0).everyDays(1).create();\n        _dp.setProperty(_FLAG, '1');\n      }\n    } finally { _lck.releaseLock(); }\n  }\n  var r = UnkouLib.dispatchInstalledEdit(e);\n  if (r && r.html) {\n    SpreadsheetApp.getUi().showModalDialog(\n      HtmlService.createHtmlOutput(r.html).setWidth(r.width || 300).setHeight(r.height || 290),\n      r.title || ''\n    );\n  }\n}\n\n// ── 画面表示 ──────────────────────────────────────────────────────────\nfunction showSidebar()            { return UnkouLib.showSidebar(); }\nfunction showUploadSidebar()      { return UnkouLib.showUploadSidebar(); }\n// ライブラリ経由だとライブラリのonOpen()（①メニュー）が実行されるためローカル実装\nfunction reloadMenu() { UnkouLib.buildClientMenu(); SpreadsheetApp.getActiveSpreadsheet().toast('メニューを再生成しました', '🔄', 3); }\n\n// ── 月次処理 ──────────────────────────────────────────────────────────\nfunction generateCurrentMonth()   { return UnkouLib.generateCurrentMonth(); }\nfunction generateNextMonth()      { return UnkouLib.generateNextMonth(); }\nfunction archiveOldMonth()        { return UnkouLib.archiveOldMonth(); }\n\n// ── シート管理 ────────────────────────────────────────────────────────\nfunction generateSummary()        { return UnkouLib.generateSummary(); }\nfunction calcDistanceManual()              { return UnkouLib.calcDistanceManual(); }\nfunction resolveAmbiguousAddresses()      { return UnkouLib.resolveAmbiguousAddresses(); }\nfunction receiveAddressChoice(s)          { return UnkouLib.receiveAddressChoice(s); }\nfunction initDistanceMasterMajorCities()  { return UnkouLib.initDistanceMasterMajorCities(); }\nfunction expandAndRefreshSheets() { return UnkouLib.expandAndRefreshSheets(); }\nfunction restoreHeaders()         { return UnkouLib.restoreHeaders(); }\nfunction autoFillExpense()        { return UnkouLib.autoFillExpense(); }\nfunction sortBothSheetsByDate()   { return UnkouLib.sortBothSheetsByDate(); }\nfunction fillMissingIdsAndCars()  { return UnkouLib.fillMissingIdsAndCars(); }\nfunction createUsageSheet()       { return UnkouLib.createUsageSheet(); }\nfunction createManualSheet()      { return UnkouLib.createManualSheet(); }\nfunction createSupportSheet()     { return UnkouLib.createSupportSheet(); }\nfunction setupSheetProtection()   { return UnkouLib.setupSheetProtection(); }\nfunction showExportDialog()             { return UnkouLib.showExportDialog(); }\nfunction exportSheetAsCsvBase64(a)      { return UnkouLib.exportSheetAsCsvBase64(a); }\nfunction exportSelectedSheetsAsExcel(a) { return UnkouLib.exportSelectedSheetsAsExcel(a); }\nfunction exportPlBundle(a)              { return UnkouLib.exportPlBundle(a); }\n// installTriggersはライブラリ経由にするとScriptAppが①を向くためローカル実装\nfunction installTriggers() {\n  var ss = SpreadsheetApp.getActiveSpreadsheet();\n  // 全バインドスクリプト横断で全インストール済みトリガーを強制削除してから3本だけ再登録\n  ScriptApp.getUserTriggers(ss).forEach(function(t) {\n    try { ScriptApp.deleteTrigger(t); } catch(e) {}\n  });\n  ScriptApp.newTrigger('installedOnEdit_').forSpreadsheet(ss).onEdit().create();\n  ScriptApp.newTrigger('onStructureChange_').forSpreadsheet(ss).onChange().create();\n  ScriptApp.newTrigger('calcDistanceTrigger_').timeBased().atHour(0).everyDays(1).create();\n  ss.toast('初期設定完了（ステータス変更ポップアップが有効になりました）', '✓', 3);\n}\n\nfunction calcDistanceTrigger_() {\n  try {\n    var parents = DriveApp.getFileById(ScriptApp.getScriptId()).getParents();\n    if (!parents.hasNext()) return;\n    UnkouLib.calcDistanceForSS(parents.next().getId());\n  } catch(e) {}\n}\nfunction onStructureChange_(e)  { UnkouLib.dispatchStructureChange(e); }\nfunction setRecalcChoice(a)       { return UnkouLib.setRecalcChoice(a); }\nfunction executeStatusSync(a,b,c){ return UnkouLib.executeStatusSync(a,b,c); }\nfunction syncToAllClientSS()      { return UnkouLib.syncToAllClientSS(); }\n\n// ── CSVインポート ─────────────────────────────────────────────────────\nfunction showCsvImportDialogUnkou()      { return UnkouLib.showCsvImportDialogUnkou(); }\nfunction showCsvImportDialogMaster()     { return UnkouLib.showCsvImportDialogMaster(); }\nfunction showCsvImportDialogCust()       { return UnkouLib.showCsvImportDialogCust(); }\nfunction showEtcImportDialog()           { return UnkouLib.showEtcImportDialog(); }\nfunction prepareEtcImport(a,b,c)         { return UnkouLib.prepareEtcImport(a,b,c); }\nfunction executeEtcImport(a,b,c,d)       { return UnkouLib.executeEtcImport(a,b,c,d); }\nfunction deleteBlankImportRows()         { return UnkouLib.deleteBlankImportRows(); }\nfunction getImportDictionary(a,b)        { return UnkouLib.getImportDictionary(a,b); }\nfunction importBulkRows(a,b,c)           { return UnkouLib.importBulkRows(a,b,c); }\nfunction saveImportAliases(a,b,c)        { return UnkouLib.saveImportAliases(a,b,c); }\n\n// ── 帳票・送信 ────────────────────────────────────────────────────────\nfunction showHatchuDocDialog()           { return UnkouLib.showHatchuDocDialog(); }\nfunction showShabanDocDialog()           { return UnkouLib.showShabanDocDialog(); }\nfunction showUketorishoDialog()          { return UnkouLib.showUketorishoDialog(); }\nfunction generateUketorishoSheet(a)      { return UnkouLib.generateUketorishoSheet(a); }\nfunction sendDocumentEmail(a,b,c)        { return UnkouLib.sendDocumentEmail(a,b,c); }\nfunction markDocumentIssued(a,b)         { return UnkouLib.markDocumentIssued(a,b); }\nfunction getShijisakiHistory(a,b)        { return UnkouLib.getShijisakiHistory(a,b); }\nfunction saveShijisakiHistory(a,b,c)     { return UnkouLib.saveShijisakiHistory(a,b,c); }\nfunction getShijisakiByRowId(a,b)           { return UnkouLib.getShijisakiByRowId(a,b); }\nfunction saveShijisakiByRowId(a,b,c,d)     { return UnkouLib.saveShijisakiByRowId(a,b,c,d); }\nfunction deleteShijisakiHistory(a,b,c,d,e,f){ return UnkouLib.deleteShijisakiHistory(a,b,c,d,e,f); }\nfunction showPlDialog()                  { return UnkouLib.showPlDialog(); }\nfunction getPlFilterOptions()            { return UnkouLib.getPlFilterOptions(); }\nfunction generatePl(a)                   { return UnkouLib.generatePl(a); }\nfunction exportPlJournalCsv()            { return UnkouLib.exportPlJournalCsv(); }\nfunction initFixedCostMaster()           { return UnkouLib.initFixedCostMaster(); }\n\n// ── 請求書・支払確認書 ────────────────────────────────────────────────\nfunction showInvoiceDialog()             { return UnkouLib.showInvoiceDialog(); }\nfunction generateInvoiceSheet(a,b,c,d)   { return UnkouLib.generateInvoiceSheet(a,b,c,d); }\nfunction showPaymentDialog()             { return UnkouLib.showPaymentDialog(); }\nfunction generatePaymentSheet(a,b,c,d,e) { return UnkouLib.generatePaymentSheet(a,b,c,d,e); }\n\n// ── 情報シート・配車確定 ──────────────────────────────────────────────\nfunction matchAndConfirmDispatch()       { return UnkouLib.matchAndConfirmDispatch(); }\nfunction cancelDispatch()               { return UnkouLib.cancelDispatch(); }\nfunction repairJohoSheet()              { return UnkouLib.repairJohoSheet(); }\nfunction generateAuditSheet()           { return UnkouLib.generateAuditSheet(); }\n// 古いインストール済みトリガー経由の発火（引数あり）は即return（多重ポップアップ封じ）\nfunction checkMasterExpiries(e)         { return; }  // デコイ：ゾンビトリガー空振り\nfunction showDispatchDashboard()        { return UnkouLib.showDispatchDashboard(); }\nfunction getDispatchDashboardData()     { return UnkouLib.getDispatchDashboardData(); }\n\n// ── アプリ連携（端末↔SS） ────────────────────────────────────────────\nfunction storeCompanySsId(a)              { return UnkouLib.storeCompanySsId(a); }\nfunction getInitialData(a,b)              { return UnkouLib.getInitialData(a,b); }\nfunction linkAddress(a,b)                 { return UnkouLib.linkAddress(a,b); }\nfunction unlinkAddress(a)                 { return UnkouLib.unlinkAddress(a); }\nfunction saveRunState(a,b,c)              { return UnkouLib.saveRunState(a,b,c); }\nfunction loadRunState()                   { return UnkouLib.loadRunState(); }\nfunction clearRunState(a,b)               { return UnkouLib.clearRunState(a,b); }\nfunction getTodayRoutes(a,b)              { return UnkouLib.getTodayRoutes(a,b); }\nfunction createParentRows(a,b,c,d,e,f)   { return UnkouLib.createParentRows(a,b,c,d,e,f); }\nfunction setPickComplete(a,b,c)           { return UnkouLib.setPickComplete(a,b,c); }\nfunction setRest(a,b,c,d)                { return UnkouLib.setRest(a,b,c,d); }\nfunction setDropComplete(a,b,c)           { return UnkouLib.setDropComplete(a,b,c); }\nfunction updateRouteData(a,b,c,d)         { return UnkouLib.updateRouteData(a,b,c,d); }\nfunction deleteRunRows(a,b,c)             { return UnkouLib.deleteRunRows(a,b,c); }\nfunction clearTimeCell(a,b,c,d,e)         { return UnkouLib.clearTimeCell(a,b,c,d,e); }\nfunction getListData(a,b,c,d)             { return UnkouLib.getListData(a,b,c,d); }\nfunction getEditData(a,b,c)               { return UnkouLib.getEditData(a,b,c); }\nfunction saveEditData(a,b,c)              { return UnkouLib.saveEditData(a,b,c); }\nfunction appendTerminalFile(a,b,c,d,e,f) { return UnkouLib.appendTerminalFile(a,b,c,d,e,f); }\nfunction deleteRunById(a,b,c)             { return UnkouLib.deleteRunById(a,b,c); }\nfunction saveNotice(a,b,c,d)             { return UnkouLib.saveNotice(a,b,c,d); }\nfunction uploadFileToRow(a,b,c,d)         { return UnkouLib.uploadFileToRow(a,b,c,d); }\nfunction saveTerminalNotice(a,b,c,d)      { return UnkouLib.saveTerminalNotice(a,b,c,d); }\nfunction uploadTerminalFile(a,b,c,d)      { return UnkouLib.uploadTerminalFile(a,b,c,d); }\nfunction getMyNotices(a,b)               { return UnkouLib.getMyNotices(a,b); }\nfunction getRoutesById(a,b,c)             { return UnkouLib.getRoutesById(a,b,c); }\nfunction getNoticeByRow(a,b,c)            { return UnkouLib.getNoticeByRow(a,b,c); }\nfunction markAsRead(a,b)                  { return UnkouLib.markAsRead(a,b); }\nfunction getReadNotices(a)               { return UnkouLib.getReadNotices(a); }\nfunction agreeContract(a,b,c,d,e)        { return UnkouLib.agreeContract(a,b,c,d,e); }\nfunction queueFileUpload(a,b,c,d)        { return UnkouLib.queueFileUpload(a,b,c,d); }\nfunction recordAction(a,b,c,d,e,f)       { return UnkouLib.recordAction(a,b,c,d,e,f); }\nfunction clearInspTime(a,b,c,d)          { return UnkouLib.clearInspTime(a,b,c,d); }\nfunction getCarInfoByNumber(a,b)         { return UnkouLib.getCarInfoByNumber(a,b); }\nfunction deleteTerminalFile(a,b,c)       { return UnkouLib.deleteTerminalFile(a,b,c); }\nfunction replaceTerminalFile(a,b,c,d,e,f){ return UnkouLib.replaceTerminalFile(a,b,c,d,e,f); }\nfunction appendTerminalFileAdmin(a,b,c,d,e){ return UnkouLib.appendTerminalFileAdmin(a,b,c,d,e); }\nfunction saveTermNoticeByDriver(a,b,c)   { return UnkouLib.saveTermNoticeByDriver(a,b,c); }\nfunction appendAdminFileById(a,b,c,d,e)  { return UnkouLib.appendAdminFileById(a,b,c,d,e); }\nfunction deleteAdminFileById(a,b,c)      { return UnkouLib.deleteAdminFileById(a,b,c); }\nfunction replaceAdminFileById(a,b,c,d,e,f){ return UnkouLib.replaceAdminFileById(a,b,c,d,e,f); }\n\n// ── 管理画面（親アプリ）────────────────────────────────────────────────\nfunction getParentSheets(a)            { return UnkouLib.getParentSheets(a); }\nfunction getSheetTableData(a,b)        { return UnkouLib.getSheetTableData(a,b); }\nfunction saveSheetRowData(a,b,c,d)     { return UnkouLib.saveSheetRowData(a,b,c,d); }\nfunction appendSheetRow(a,b,c)         { return UnkouLib.appendSheetRow(a,b,c); }\nfunction deleteSheetRow(a,b,c)         { return UnkouLib.deleteSheetRow(a,b,c); }\nfunction afterSaveJoho(a,b,c)          { return UnkouLib.afterSaveJoho(a,b,c); }\nfunction afterSaveJohoFull(a,b)        { return UnkouLib.afterSaveJohoFull(a,b); }\nfunction appendJohoRow(a,b)            { return UnkouLib.appendJohoRow(a,b); }\nfunction linkAdminEmail(a,b)           { return UnkouLib.linkAdminEmail(a,b); }\nfunction getLinkedAdminEmail(a)        { return UnkouLib.getLinkedAdminEmail(a); }\nfunction removeAllProtections()        { return UnkouLib.removeAllProtections(); }\n\n// ── バックアップ・復旧 ────────────────────────────────────────────────\nfunction openRestoreDialog()           { return UnkouLib.openRestoreDialog(); }\nfunction executeRestore(a,b)           { return UnkouLib.executeRestore(a,b); }\n\n// ── 保守ユーティリティ（ローカル実装：ScriptApp・SpreadsheetApp は呼び出し元SS文脈で動かす必要あり）────\nfunction cleanupStaleTriggers() {\n  var ss       = SpreadsheetApp.getActiveSpreadsheet();\n  var staleFns = ['checkMasterExpiries', 'onOpen', 'checkExpiryDates'];\n  var removed  = 0;\n  ScriptApp.getUserTriggers(ss).forEach(function(t) {\n    if (staleFns.indexOf(t.getHandlerFunction()) !== -1) {\n      try { ScriptApp.deleteTrigger(t); removed++; } catch(e) {}\n    }\n  });\n  ['指示先履歴', '指示先ID別'].forEach(function(name) {\n    var sh = ss.getSheetByName(name);\n    if (sh && !sh.isSheetHidden()) { try { sh.hideSheet(); } catch(e) {} }\n  });\n  SpreadsheetApp.getUi().alert(\n    '✅ クリーンアップ完了\\n\\n' +\n    '・削除したトリガー：' + removed + '件\\n' +\n    '・システムシート（指示先履歴・指示先ID別）を非表示にしました'\n  );\n}\n";
  // === AUTO_GENERATED_STUB_END ===
}


// ================================================================
//  12-3b: クライアントSSシート初期化共通処理（initClientSSSheets_）  【大B / 中12 / 小12-3b】
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
    '仮日数','給料','％','有休手当','その他手当','点呼前完了','点呼後完了','拘束時間(h)','点呼前担当者','点呼後担当者'
  ];
  var masterHeader = [
    '車両ID','運行状態','区分','会社名','看板名','トン数','車種','車番','乗務員名','携帯番号',
    'アドレス','燃費','備考','仮日数','給料','％','高速を引く（引くは〇、引かないは空欄）',
    '車両リース代','任意保険料','自賠責保険料','重量税積立','車検費積立',
    '整備費積立','タイヤ代積立','修理積立','駐車場代','ETCリース料',
    'カーナビリース料','通信費','洗車費','制服費','その他固定費',
    '免許証有効期限','安全教育次回予定日','健康診断次回予定日','適性診断次回予定日','担当管理者'
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
  var validNames = ['運行','集計表','自車専属マスタ','自車専属運行','マスタ','設定','__COMPANY_SS__','メモ'];
  ss.getSheets().forEach(function(s) {
    if (validNames.indexOf(s.getName()) === -1 && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch(e) {}
    }
  });
  if (!ss.getSheetByName('メモ')) ss.insertSheet('メモ');

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
//
//  ★ M&A後の引き継ぎメモ ★
//  この関数は「①修正用SS（ライブラリ本体）のバージョンスナップショット」を作成する。
//  ScriptApp.getScriptId() はGASの仕様で「library として呼ばれると呼び出し元のID」を返すため、
//  Script Properties の 'ownScriptId' を優先参照して①のIDを正確に取得する。
//  'ownScriptId' は syncToTemplateSS（STEP2「テスト客SS反映」）実行時に自動保存される。
//
//  ★ Googleアカウントごと譲渡する場合（推奨）★
//    → ①②③のID・URLは一切変わらないため、このコードを修正する必要なし。
//      新管理者が clasp login で再認証するだけ。
//
//  ★ GASプロジェクトを新規作成して移行する場合 ★
//    → 以下の4か所を新しいIDに更新する必要がある（詳細は syncToTemplateSS のコメント参照）：
//      [1] .clasp.json の scriptId
//      [2] stub_for_clientSS/appsscript.json の libraryId
//      [3] syncToTemplateSS の TEMPLATE_SCRIPT_ID と templateSsId
//      [4] Script Properties の 'ownScriptId' と 'clientTemplateSsId'（①から「テスト客SS反映」実行で自動更新）
// ================================================================
function createLibraryVersion_(description) {
  try {
    var token   = ScriptApp.getOAuthToken();
    // ScriptApp.getScriptId() はライブラリとして呼ばれると呼び出し元のIDを返すため
    // syncToTemplateSS が保存した①本体のIDを Script Properties から優先参照する
    var props2  = PropertiesService.getScriptProperties();
    var scriptId = props2.getProperty('ownScriptId') || ScriptApp.getScriptId();
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
//
//  ★ 重要バグ修正メモ（2026-06-10）★
//  旧実装は `ScriptApp.getScriptId()` でライブラリIDを取得し libraryId と比較していたが、
//  GASの仕様で「libraryとして呼ばれると呼び出し元（②）のIDを返す」ため、
//  ②から「各客に反映」を実行したとき libraryId（①のID）と一致せずバージョンが更新されなかった。
//  → `userSymbol === 'UnkouLib'` に変更することでIDに依存しない比較に修正済み。
//
//  ★ M&A後の引き継ぎメモ ★
//  userSymbol は appsscript.json の "userSymbol": "UnkouLib" と一致させる文字列。
//  新管理者がライブラリのuserSymbolを変更した場合はこの文字列も合わせて変更すること。
//  ライブラリID（libraryId）が変わっても userSymbol が 'UnkouLib' のままなら修正不要。
// ================================================================
function updateStubVersion_(stubScriptId, versionNumber, useDevMode) {
  try {
    var token = ScriptApp.getOAuthToken();
    // ScriptApp.getScriptId() はライブラリとして呼ばれると呼び出し元のIDを返すため削除
    // → userSymbol 'UnkouLib' で一致させることでIDに依存しない比較に変更
    // GETせず固定2ファイルで完全上書き（クリーンインストール）
    // ③各客SSは常に appsscript.json + コード.js の2ファイルのみ。②と100%同一を保証。
    var files = [
      { name: 'appsscript', type: 'JSON', source: JSON.stringify({
        timeZone: 'Asia/Tokyo',
        dependencies: { libraries: [{ userSymbol: 'UnkouLib',
          libraryId: '1n79omnAcdsEojMRyjnj9-Ic9pIl1-7Nt_HB7Avy0NVFizOSeqt0guqyZ',
          version: String(versionNumber), developmentMode: true }] },
        oauthScopes: ['https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/script.external_request',
          'https://www.googleapis.com/auth/script.scriptapp',
          'https://www.googleapis.com/auth/gmail.send',
          'https://www.googleapis.com/auth/script.container.ui',
          'https://www.googleapis.com/auth/script.projects',
          'https://www.googleapis.com/auth/script.deployments'],
        webapp: { executeAs: 'USER_DEPLOYING', access: 'ANYONE_ANONYMOUS' },
        exceptionLogging: 'STACKDRIVER', runtimeVersion: 'V8'
      }, null, 2) },
      { name: 'コード', type: 'SERVER_JS', source: getClientStubSource_() }
    ];
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
//  ── 請求書・支払確認書（16-3〜16-7）──
//  syncToTemplateSS（12-3c）本体は下記に定義（M&Aメモ含む）
// ================================================================
// ================================================================
//  16-3: 請求書生成ダイアログ（showInvoiceDialog）  【大C / 中16 / 小16-3】
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
//  16-4: 支払確認書生成ダイアログ（showPaymentDialog）  【大C / 中16 / 小16-4】
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
//  16-5: 書類用連番採番（getNextDocNum_）  【大B / 中16 / 小16-5】
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
//  16-6: 請求書シート生成（generateInvoiceSheet）  【大A / 中16 / 小16-6】
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
//  16-7: 支払確認書シート生成（generatePaymentSheet）  【大A / 中16 / 小16-7】
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
//  12-3c: ①修正用SS→②客用SSに反映（syncToTemplateSS）  【大C / 中12 / 小12-3c】
//
//  ★ M&A後の引き継ぎメモ ★
//  この関数は①から直接実行（メニュー「テスト客SS反映」）することが必要。
//  ②のメニューから実行しても動かない（①のメニューにしか存在しない）。
//
//  ★ Googleアカウントごと譲渡する場合（推奨）★
//    → TEMPLATE_SCRIPT_ID・templateSsId はそのまま使用可能。変更不要。
//
//  ★ GASプロジェクトを新規作成して移行する場合 ★
//    TEMPLATE_SCRIPT_ID → 新②客用SS の Apps Script ID（スクリプトエディタURLの /projects/XXXX 部分）
//    templateSsId       → 新②客用SS の スプレッドシートID（SS URLの /d/XXXX 部分）
//    Script Properties の 'clientTemplateSsId' も更新（または削除してデフォルト値を変更）
// ================================================================
function syncToTemplateSS() {
  var props = PropertiesService.getScriptProperties();
  var TEMPLATE_SCRIPT_ID = '19CfyUPhldzSccj05xo-sn4Xh78fCHAHDVJtGyKdDGQkO1D4wZWFEnZCT';
  var templateSsId = props.getProperty('clientTemplateSsId') || '1NBtosd_MN8KcboV_4OXTrY8WqcE3TJwpxdA_nASmTOo';
  var masterSs = SpreadsheetApp.getActiveSpreadsheet();
  var tgtSs    = SpreadsheetApp.openById(templateSsId);
  DriveApp.getFileById(tgtSs.getId()).setName('客用');

  // ①修正用SS自身のスクリプトIDを保存（createLibraryVersion_がlibrary経由で呼ばれても正しいIDを参照できるよう）
  // この関数は必ず①から直接実行されるため、ScriptApp.getScriptId()は①のIDを正しく返す
  props.setProperty('ownScriptId', ScriptApp.getScriptId());

  // ② 毎回新バージョン作成（デプロイ後の最新コードを確実に反映するため）
  var newVersion = createLibraryVersion_(
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm') + ' テスト客SS反映'
  );
  if (newVersion) {
    props.setProperty('lastLibVersionTime', String(Date.now()));
    props.setProperty('lastLibVersionNum',  String(newVersion));
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
  var businessSheets = ['運行', '集計表', '自車専属マスタ', '自車専属運行', 'マスタ', '設定',
                        '配車板', '自社設定', '距離マスタ', 'PL設定', '管理者'];
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
    // ①の列数を超えた余分な列を消去（②③が勝手に列を増やしても①の構造に強制一致させる）
    var _tgtCols = tgtSheet.getLastColumn();
    if (_tgtCols > srcCols) {
      tgtSheet.getRange(1, srcCols + 1, 1, _tgtCols - srcCols).clearContent().clearFormat();
    }
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
  initImportDictionary_(tgtSs);

  // 自車専属マスタのコンプライアンス日付列フォーマットとAF〜AMヘッダー色を②にも適用
  (function() {
    var _ms = tgtSs.getSheetByName('自車専属マスタ');
    if (!_ms || _ms.getLastRow() < 2) return;
    var _mh = _ms.getRange(1, 1, 1, _ms.getLastColumn()).getValues()[0];
    var _mr = _ms.getLastRow() - 1;
    ['免許証有効期限','安全教育次回予定日','健康診断次回予定日','適性診断次回予定日'].forEach(function(cn) {
      var ci = _mh.indexOf(cn);
      if (ci >= 0) _ms.getRange(2, ci + 1, _mr, 1).setNumberFormat('yyyy/M/d');
    });
    var _afI = _mh.indexOf('その他固定費'), _amI = _mh.indexOf('担当管理者');
    if (_afI >= 0 && _amI >= _afI) {
      _ms.getRange(1, _afI + 1, 1, _amI - _afI + 1).setBackground('#212121').setFontColor('#00e676').setFontWeight('bold');
    }
  })();

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

  // 会社登録シートのヘッダーを確保（ヘッダー名検索で列ズレに強く）
  if (regSheet && regSheet.getLastRow() >= 1) {
    var hdrColCount = regSheet.getLastColumn();
    if (hdrColCount < 9)  regSheet.getRange(1, 9).setValue('契約書URL');
    if (hdrColCount < 10) regSheet.getRange(1, 10).setValue('同意時刻');
    // スクリプトIDヘッダーがなければ末尾に追加（固定列に書かない）
    var regHdrs_ = regSheet.getRange(1, 1, 1, regSheet.getLastColumn()).getValues()[0].map(function(h){return String(h||'').trim();});
    if (regHdrs_.indexOf('スクリプトID') === -1) {
      regSheet.getRange(1, regSheet.getLastColumn() + 1).setValue('スクリプトID');
    }
  }

  // ① 共有フォルダを作成（メール送信なし）
  var folderResult = setupOneCompany_(companyName, adminEmail, true, true);
  var folderUrl = folderResult.folderUrl;
  var folderId  = folderResult.folderId;

  // ② ②客用SSをコピーして③各客SS作成（スタブコードのみ・心臓部コードは含まれない）
  var ssResult   = createCompanySpreadsheet_(companyName, adminEmail, folderId);
  var ssUrl      = ssResult.ssUrl;
  var ssId       = ssResult.ssId;
  var clientScriptId = ssResult.scriptId || '';

  // ③ アプリURL = 会社SS独自のWebAppURL（ssId不要・独立URL）
  var appUrl = getWebAppBaseUrl_() + '?ssId=' + encodeURIComponent(ssId);

  // ④ 会社登録シートの行番号を特定
  var targetRow = -1;
  if (regSheet && regSheet.getLastRow() >= 2) {
    var rows = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === companyName) { targetRow = i + 2; break; }
    }
  }

  // ④ 契約書URL = 常に①のURL（安定動作）
  var masterSsId_ = PropertiesService.getScriptProperties().getProperty('masterSsId') || ss.getId();
  var contractUrl = '[WebアプリURL未設定]';
  var baseUrl = getWebAppBaseUrl_();
  if (baseUrl) {
    contractUrl = baseUrl + '?page=contract' +
      '&ssId='    + encodeURIComponent(ssId) +
      '&company=' + encodeURIComponent(companyName) +
      (targetRow > 0 ? '&row=' + targetRow : '');
  }

  // 客SSにmasterSsIdとcontractRowを保存（doGetのaction=agreeで①②どちらからでも正しいSSを特定するため）
  if (ssId && targetRow > 0 && masterSsId_) {
    try {
      var metaClientSs_ = SpreadsheetApp.openById(ssId);
      var metaSheet_ = metaClientSs_.getSheetByName('__AGREE_META__') || metaClientSs_.insertSheet('__AGREE_META__');
      metaSheet_.getRange(1, 1).setValue(masterSsId_);
      metaSheet_.getRange(1, 2).setValue(targetRow);
      if (!metaSheet_.isSheetHidden()) metaSheet_.hideSheet();
    } catch(e_) {}
  }

  // ⑤ 会社登録シートに記録
  if (targetRow > 0) {
    regSheet.getRange(targetRow, 3).setValue('契約書送信済').setBackground('#fff9c4').clearNote();
    regSheet.getRange(targetRow, 4).setValue(new Date());
    regSheet.getRange(targetRow, 5).setValue(folderUrl);
    regSheet.getRange(targetRow, 6).setValue(ssUrl);
    regSheet.getRange(targetRow, 7).setValue(appUrl);
    regSheet.getRange(targetRow, 9).setValue(contractUrl);
    if (clientScriptId) {
      var sidHdrs_ = regSheet.getRange(1, 1, 1, regSheet.getLastColumn()).getValues()[0].map(function(h){return String(h||'').trim();});
      var sidCol_ = sidHdrs_.indexOf('スクリプトID');
      if (sidCol_ >= 0) regSheet.getRange(targetRow, sidCol_ + 1).setValue(clientScriptId);
    }
  }

  // ⑥ 管理Gmail宛に「契約書確認のお願い」メールを送信
  var subject = '[運行管理] ' + companyName + ' 利用規約への同意のお願い';
  var body =
    companyName + ' ご担当者様\n\n' +
    'このたびは運行管理システムへのお申し込みありがとうございます。\n\n' +
    '下記より利用規約をご確認いただき、「同意する」ボタンを押してください。\n' +
    '同意完了後に、スプレッドシートおよびアプリのURLをメールでお送りします。\n\n' +
    '▶ 利用規約・同意フォームはこちら\n' + contractUrl + '\n\n' +
    '何かご不明な点があればお気軽にお問い合わせください。\n' +
    'よろしくお願いいたします。';
  var htmlBody =
    '<html><head><meta charset="utf-8"></head><body>' +
    '<div style="font-family:sans-serif;font-size:14px;color:#212121;max-width:520px;">' +
    '<p>' + companyName + ' ご担当者様</p>' +
    '<p>このたびは運行管理システムへのお申し込みありがとうございます。</p>' +
    '<p>下記より利用規約をご確認いただき、「同意する」ボタンを押してください。<br>' +
    '同意完了後に、スプレッドシートおよびアプリのURLをメールでお送りします。</p>' +
    '<p style="margin:24px 0;">' +
    '<a href="' + contractUrl + '" style="background:#1565c0;color:#ffffff;padding:12px 24px;' +
    'text-decoration:none;border-radius:4px;display:inline-block;font-size:15px;">▶ 利用規約・同意フォームはこちら</a></p>' +
    '<p>何かご不明な点があればお気軽にお問い合わせください。<br>よろしくお願いいたします。</p>' +
    '</div></body></html>';
  try {
    GmailApp.sendEmail(adminEmail, subject, body, { htmlBody: htmlBody });
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
//  doGet(?action=agree) から呼ばれる。google.script.runは使用しない（Googleセキュリティ通知回避のため）。
//  masterSsId はURLパラメータに含めず PropertiesService から取得する（URL機密情報除去のため）。
//  adminEmail はURLパラメータに含めず会社登録シートのB列から取得する（同上）。
//  ① 会社登録シートのJ列(10)に同意時刻を記録・C列を「同意済」に更新
//  ② 管理Gmail宛にSS URLとアプリURLをメール送信
//  ③ H列(8)に送信済ステータスを記録
// ================================================================
function agreeContract(ssId, companyName, adminEmail, contractRow, masterSsIdParam) {
  var masterSsId = masterSsIdParam || '';

  // 客SSの__AGREE_META__シートからmasterSsIdとcontractRowを取得（①doGet経由時はScript Propertiesが違うSSを指すため）
  if (!masterSsId && ssId) {
    try {
      var agreeMetaSs_ = SpreadsheetApp.openById(ssId);
      var agreeMetaSheet_ = agreeMetaSs_.getSheetByName('__AGREE_META__');
      if (agreeMetaSheet_) {
        masterSsId = String(agreeMetaSheet_.getRange(1, 1).getValue() || '');
        var metaRow_ = String(agreeMetaSheet_.getRange(1, 2).getValue() || '');
        if (metaRow_ && (!contractRow || contractRow === '')) contractRow = metaRow_;
      }
    } catch(e_) {}
  }

  // フォールバック: Script Propertiesから取得
  if (!masterSsId) masterSsId = PropertiesService.getScriptProperties().getProperty('masterSsId');
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
    // adminEmailがURLに含まれなくなったのでシートのB列から読む
    if (!adminEmail || adminEmail.indexOf('@') === -1) {
      adminEmail = String(regSheet.getRange(row, 2).getValue() || '');
    }
    regSheet.getRange(row, 10).setValue(now).setBackground('#c8e6c9');
    regSheet.getRange(row, 3).setValue('同意済').setBackground('#c8e6c9');
  }
  if (!ssUrl && ssId) ssUrl  = 'https://docs.google.com/spreadsheets/d/' + ssId;
  if (!appUrl && ssId) appUrl = getWebAppBaseUrl_() + '?ssId=' + encodeURIComponent(ssId);

  if (adminEmail && adminEmail.indexOf('@') !== -1) {
    var subject = '[運行管理] ' + companyName + ' 運行管理システム利用開始のご案内';
    var body =
      companyName + ' ご担当者様\n\n' +
      '利用規約へのご同意ありがとうございます。\n\n' +
      '以下よりご利用を開始いただけます。\n\n' +
      '■ 運行管理スプレッドシート（PC・タブレット推奨）\n' + ssUrl + '\n\n' +
      '■ 運行管理アプリ（乗務員用 スマートフォン推奨）\n' + appUrl + '\n\n' +
      '【スプレッドシートの使い方】\n' +
      '・「自車専属マスタ」タブのK列（アドレス）に乗務員メールを入力してください\n\n' +
      '【乗務員への配布方法】\n' +
      '・各乗務員に上記アプリURLを共有してください\n' +
      'ご不明な点はお気軽にお問い合わせください。\n' +
      'よろしくお願いいたします。';
    var safeCompanyName = companyName.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    var htmlBody2 =
      '<html><head><meta charset="utf-8"></head><body>' +
      '<div style="font-family:sans-serif;font-size:14px;color:#212121;max-width:520px;">' +
      '<p>' + safeCompanyName + ' ご担当者様</p>' +
      '<p>利用規約へのご同意ありがとうございます。<br>以下よりご利用を開始いただけます。</p>' +
      '<table style="border-collapse:collapse;width:100%;margin:16px 0;">' +
      '<tr><td style="padding:12px;background:#e3f2fd;border-left:4px solid #1565c0;vertical-align:top;">' +
      '<strong>運行管理スプレッドシート</strong><br><span style="font-size:12px;color:#555;">PC・タブレット推奨 ／ 配車・運行データの入力・集計</span><br><br>' +
      '<a href="' + ssUrl + '" style="background:#1565c0;color:#fff;padding:8px 16px;text-decoration:none;border-radius:4px;display:inline-block;">スプレッドシートを開く</a>' +
      '</td></tr>' +
      '<tr><td style="padding:12px;background:#e8f5e9;border-left:4px solid #2e7d32;vertical-align:top;margin-top:8px;">' +
      '<strong>運行管理アプリ</strong><br><span style="font-size:12px;color:#555;">スマートフォン推奨 ／ 乗務員が運行状況を入力</span><br><br>' +
      '<a href="' + appUrl + '" style="background:#2e7d32;color:#fff;padding:8px 16px;text-decoration:none;border-radius:4px;display:inline-block;">アプリを開く</a>' +
      '</td></tr></table>' +
      '<p><strong>【スプレッドシートの使い方】</strong><br>「自車専属マスタ」タブのK列（アドレス）に乗務員メールを入力してください</p>' +
      '<p><strong>【乗務員への配布方法】</strong><br>各乗務員に上記アプリURLを共有してください<br>初回は「紐づけ設定」でメールアドレスを登録するだけで使えます</p>' +
      '<p>ご不明な点はお気軽にお問い合わせください。<br>よろしくお願いいたします。</p>' +
      '</div></body></html>';
    GmailApp.sendEmail(adminEmail, subject, body, { htmlBody: htmlBody2 });
  }

  if (row >= 2) {
    regSheet.getRange(row, 8).setValue('送信済(' + now + ')').setBackground('#c8e6c9');
  }

  // 同意後にSSとフォルダを共有（Drive通知メールなし）
  if (ssId && adminEmail) {
    try {
      var oauthToken_  = ScriptApp.getOAuthToken();
      var reqHeaders_  = { Authorization: 'Bearer ' + oauthToken_, 'Content-Type': 'application/json' };
      var permPayload_ = JSON.stringify({ role: 'writer', type: 'user', emailAddress: adminEmail });
      UrlFetchApp.fetch(
        'https://www.googleapis.com/drive/v3/files/' + ssId + '/permissions?sendNotificationEmail=false',
        { method: 'post', headers: reqHeaders_, payload: permPayload_, muteHttpExceptions: true }
      );
      var parentIter_ = DriveApp.getFileById(ssId).getParents();
      if (parentIter_.hasNext()) {
        UrlFetchApp.fetch(
          'https://www.googleapis.com/drive/v3/files/' + parentIter_.next().getId() + '/permissions?sendNotificationEmail=false',
          { method: 'post', headers: reqHeaders_, payload: permPayload_, muteHttpExceptions: true }
        );
      }
    } catch(e) {}
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
//  [ADD] 距離マスタ・自動距離計算 ユーティリティ群
//  距離マスタシートへのキャッシュ照合・Maps API計算・集計表W列反映を担う。
//  行程入力時にキャッシュヒットすれば即時反映。未ヒットは13時トリガーorメニューで補完。
// ================================================================

// 地名文字列の先頭の丸数字（①〜⑩）を解析して順序と地名を返す
function parseLocation_(str) {
  str = String(str || '').trim();
  var circled = '①②③④⑤⑥⑦⑧⑨⑩';
  if (str.length > 0 && circled.indexOf(str.charAt(0)) >= 0) {
    return { order: circled.indexOf(str.charAt(0)) + 1, location: str.slice(1).trim() };
  }
  return { order: 0, location: str };
}

// picks/drops 配列からルートキーを生成（①②順・なければ積地→降地の入力順）
function buildRouteKeyFromPicksDrops_(picks, drops) {
  var entries = [];
  for (var i = 0; i < picks.length; i++) {
    var p = parseLocation_(picks[i] || '');
    var d = parseLocation_(drops[i] || '');
    if (p.location) entries.push({ order: p.order || 999, location: p.location });
    if (d.location) entries.push({ order: d.order || 999, location: d.location });
  }
  if (!entries.length) return '';
  var hasNum = entries.some(function(e) { return e.order < 999; });
  if (hasNum) entries.sort(function(a, b) { return a.order - b.order; });
  return entries.map(function(e) { return e.location; }).join('→');
}

// 運行シートデータから積完/降完の時刻順でルートキーを生成（トリガー・手動ボタン用）
function buildTimeOrderedRouteKey_(unkouData, targetId) {
  var events = [];
  for (var i = 1; i < unkouData.length; i++) {
    var r = unkouData[i];
    if (String(r[0] || '').trim() !== targetId) continue;
    var pickTime = r[14]; // O列=積完時刻
    var pickLoc  = parseLocation_(String(r[11] || '')).location;
    if (pickTime instanceof Date && pickLoc) events.push({ time: pickTime, loc: pickLoc });
    var dropTime = r[17]; // R列=降完時刻
    var dropLoc  = parseLocation_(String(r[12] || '')).location;
    if (dropTime instanceof Date && dropLoc) events.push({ time: dropTime, loc: dropLoc });
  }
  if (!events.length) return null;
  events.sort(function(a, b) { return a.time - b.time; });
  return events.map(function(e) { return e.loc; }).join('→');
}

// 距離マスタシートを取得または作成（なければ新規作成してヘッダーを書く）
function getOrCreateDistanceMasterSheet_(ss) {
  var sh = ss.getSheetByName('距離マスタ');
  if (!sh) {
    sh = ss.insertSheet('距離マスタ');
    sh.getRange(1, 1, 1, 4).setValues([['ルートキー', '距離(km)', '計算方法', '更新日']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// 距離マスタを {ルートキー: km} のマップとして一括ロード（API呼び出し1回）
function loadDistanceMasterMap_(ss) {
  var sh  = getOrCreateDistanceMasterSheet_(ss);
  var map = {};
  var lr  = sh.getLastRow();
  if (lr < 2) return map;
  var data = sh.getRange(2, 1, lr - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    var key = String(data[i][0] || '').trim();
    var km  = Number(data[i][1]) || 0;
    if (key && km) map[key] = km;
  }
  return map;
}

// 距離マスタに新規ルートを一括追記（既存キーはスキップ・setValuesで一括書き込み）
function appendDistanceMasterRows_(ss, newRoutes) {
  if (!newRoutes || !newRoutes.length) return;
  var sh    = getOrCreateDistanceMasterSheet_(ss);
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  var rows  = newRoutes.map(function(r) { return [r.key, r.km, r.method, today]; });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
}

// Google Maps Directions API で経由地込みの走行距離を計算して km を返す
function calcDistanceMapsApi_(locations) {
  if (!locations || locations.length < 2) return null;
  var MAX_RETRY = 3;
  for (var attempt = 0; attempt < MAX_RETRY; attempt++) {
    try {
      var finder = Maps.newDirectionFinder().setMode(Maps.DirectionFinder.Mode.DRIVING);
      finder.setOrigin(locations[0]);
      for (var i = 1; i < locations.length - 1; i++) finder.addWaypoint(locations[i]);
      finder.setDestination(locations[locations.length - 1]);
      var dirs = finder.getDirections();
      if (!dirs || !dirs.routes || !dirs.routes.length) return null;
      var total = 0;
      dirs.routes[0].legs.forEach(function(leg) { total += leg.distance.value; });
      return Math.round(total / 1000);
    } catch(e) {
      if (attempt < MAX_RETRY - 1) Utilities.sleep(1000);
    }
  }
  return null;
}

// 集計表W列（距離・col23）に値をセット。isAuto=trueで緑文字（API計算済み）
function setDistanceInSummary_(ss, id, km, isAuto) {
  var sh = ss.getSheetByName('集計表');
  if (!sh || sh.getLastRow() < 2) return;
  var ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() !== String(id).trim()) continue;
    var cell = sh.getRange(i + 2, 23);
    cell.setValue(km);
    if (isAuto) cell.setFontColor('#1a9a50');
    return;
  }
}

// 行程作成直後に距離マスタを照合して即時反映（ヒットしなければ無処理・失敗してもOK）
function lookupAndSetDistanceAfterCreate_(ss, id, picks, drops) {
  try {
    var routeKey = buildRouteKeyFromPicksDrops_(picks, drops);
    if (!routeKey) return;
    var map = loadDistanceMasterMap_(ss);
    if (map[routeKey] !== undefined) {
      setDistanceInSummary_(ss, id, map[routeKey], true);
      return;
    }
    // マスタにない場合はMaps APIで即時計算してマスタ追加＋即時反映
    var locs = routeKey.split('→');
    if (locs.length < 2) return;
    var km = calcDistanceMapsApi_(locs);
    if (km !== null) {
      setDistanceInSummary_(ss, id, km, true);
      appendDistanceMasterRows_(ss, [{ key: routeKey, km: km, method: 'API' }]);
    }
  } catch(e) {}
}

// 降完あり・距離未入力の行を一括処理してW列に反映（トリガー・手動ボタン共通）
function calcDistanceForAllPending_(ss) {
  var sumSh   = ss.getSheetByName('集計表');
  var unkouSh = ss.getSheetByName('運行');
  if (!sumSh || !unkouSh || sumSh.getLastRow() < 2) return 0;

  var numRows  = sumSh.getLastRow() - 1;
  var sumCols  = Math.max(sumSh.getLastColumn(), 37);
  var sumData  = sumSh.getRange(2, 1, numRows, sumCols).getValues();
  var sumFonts = sumSh.getRange(2, 1, numRows, sumCols).getFontColors();
  var unkouData = unkouSh.getDataRange().getValues();
  var masterMap = loadDistanceMasterMap_(ss);

  // エラー・確認待ちのルートキーを skipSet に収集（再試行しない）
  var skipSet = {};
  var dmSh = ss.getSheetByName('距離マスタ');
  if (dmSh && dmSh.getLastRow() >= 2) {
    var dmFull = dmSh.getRange(2, 1, dmSh.getLastRow() - 1, 3).getValues();
    for (var si = 0; si < dmFull.length; si++) {
      var sk = String(dmFull[si][0] || '').trim();
      var sm = String(dmFull[si][2] || '').trim();
      if (sk && (sm === 'エラー(住所不明等)' || sm === '住所確認待ち')) skipSet[sk] = true;
    }
  }

  var GREEN         = '#1a9a50';
  var newRoutes     = [];
  var errorRoutes   = [];  // 候補0件：-1フラグ
  var pendingRoutes = [];  // 候補2件以上：住所確認待ち
  var processed     = 0;
  var apiCalls      = 0;

  for (var i = 0; i < sumData.length; i++) {
    var id = String(sumData[i][0] || '').trim();
    if (!id) continue;

    // 緑（API計算済み）はスキップ。空・黒（手入力）は全て処理して緑に上書き
    var curDist = sumData[i][22];
    var curFont = String(sumFonts[i][22] || '').toLowerCase();
    var alreadyGreen = (curFont === '#1a9a50') && curDist !== '' && curDist !== null && curDist !== undefined;
    if (alreadyGreen) continue;

    // 積地・降地がなければスキップ（空行）
    var pickStr = String(sumData[i][11] || '').trim(); // L列=積地
    var dropStr = String(sumData[i][12] || '').trim(); // M列=降地
    if (!pickStr || !dropStr) continue;

    // STEP1: 積地/降地のシンプルキーでマスタ照合（降完不要）
    var pickArr   = pickStr.split('・');
    var dropArr   = dropStr.split('・');
    var simpleKey = buildRouteKeyFromPicksDrops_(pickArr, dropArr);
    if (simpleKey && skipSet[simpleKey]) continue;  // エラー・確認待ちはスキップ
    if (simpleKey && masterMap[simpleKey] !== undefined && masterMap[simpleKey] > 0) {
      sumData[i][22]  = masterMap[simpleKey];
      sumFonts[i][22] = GREEN;
      processed++;
      continue;
    }

    // STEP2: 降完ありなら時刻順キーでもマスタ照合
    var routeKey = null;
    if (sumData[i][17]) {
      routeKey = buildTimeOrderedRouteKey_(unkouData, id);
      if (routeKey && skipSet[routeKey]) continue;  // エラー・確認待ちはスキップ
      if (routeKey && masterMap[routeKey] !== undefined && masterMap[routeKey] > 0) {
        sumData[i][22]  = masterMap[routeKey];
        sumFonts[i][22] = GREEN;
        processed++;
        continue;
      }
    }

    // STEP3: Geocoding確認→問題なければMaps API計算
    var apiKey  = routeKey || simpleKey;
    var apiLocs = apiKey ? apiKey.split('→') : null;
    if (!apiLocs || apiLocs.length < 2) continue;
    if (skipSet[apiKey]) continue;

    var geoStatus = checkGeocode_(apiLocs);
    apiCalls++;
    if (geoStatus === 'error') {
      errorRoutes.push({ key: apiKey, km: -1, method: 'エラー(住所不明等)' });
      skipSet[apiKey] = true;
      if (apiCalls % 10 === 0) Utilities.sleep(500);
      continue;
    }
    if (geoStatus === 'pending') {
      pendingRoutes.push({ key: apiKey, km: '', method: '住所確認待ち' });
      skipSet[apiKey] = true;
      if (apiCalls % 10 === 0) Utilities.sleep(500);
      continue;
    }

    var km = calcDistanceMapsApi_(apiLocs);
    if (km !== null) {
      sumData[i][22]    = km;
      sumFonts[i][22]   = GREEN;
      masterMap[apiKey] = km;
      newRoutes.push({ key: apiKey, km: km, method: 'API' });
      processed++;
    }
    apiCalls++;
    if (apiCalls % 10 === 0) Utilities.sleep(500);
  }

  // W列のみ一括書き込み（setValues + setFontColors各1回）
  if (processed > 0) {
    var distVals  = sumData.map(function(r) { return [r[22]]; });
    var distFonts = sumFonts.map(function(r) { return [r[22]]; });
    sumSh.getRange(2, 23, numRows, 1).setValues(distVals);
    sumSh.getRange(2, 23, numRows, 1).setFontColors(distFonts);
  }
  if (newRoutes.length)     appendDistanceMasterRows_(ss, newRoutes);
  if (errorRoutes.length)   { appendDistanceMasterRows_(ss, errorRoutes);  applyDistanceMasterErrorStyle_(ss); }
  if (pendingRoutes.length) appendDistanceMasterRows_(ss, pendingRoutes);
  return processed;
}

// Geocoding APIで地名候補数を確認（'ok'=全地名1件・'error'=0件あり・'pending'=2件以上あり）
function checkGeocode_(locs) {
  var geocoder = Maps.newGeocoder().setRegion('JP');
  for (var i = 0; i < locs.length; i++) {
    var loc = String(locs[i] || '').trim();
    if (!loc) continue;
    try {
      var res   = geocoder.geocode(loc);
      var count = (res && res.results) ? res.results.length : 0;
      if (count === 0) return 'error';
      if (count >= 2) return 'pending';
    } catch(e) { return 'error'; }
  }
  return 'ok';
}

// 距離マスタのエラー行（計算方法="エラー(住所不明等)"）に背景色 #e1bee7 を設定
function applyDistanceMasterErrorStyle_(ss) {
  var sh = ss.getSheetByName('距離マスタ');
  if (!sh || sh.getLastRow() < 2) return;
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  var backgrounds = data.map(function(row) {
    var color = String(row[2] || '').trim() === 'エラー(住所不明等)' ? '#e1bee7' : null;
    return [color, color, color, color];
  });
  sh.getRange(2, 1, data.length, 4).setBackgrounds(backgrounds);
}

// 距離マスタの指定キー行を上書き更新（住所選択後の確定処理で使用）
function updateDistanceMasterRow_(ss, key, km, method) {
  var sh = ss.getSheetByName('距離マスタ');
  if (!sh || sh.getLastRow() < 2) return;
  var keys  = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  for (var i = 0; i < keys.length; i++) {
    if (String(keys[i][0] || '').trim() !== String(key).trim()) continue;
    sh.getRange(i + 2, 2, 1, 3).setValues([[km, method, today]]);
    sh.getRange(i + 2, 1, 1, 4).setBackground(null);
    return;
  }
}

// ================================================================
//  4-9: 住所確認（resolveAmbiguousAddresses）  【大C / 中4 / 小4-9】
//  メニュー「住所確認（確認待ち分）」: 距離マスタの「住所確認待ち」ルートを1件ずつダイアログで候補選択・確定登録
// ================================================================
function resolveAmbiguousAddresses() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('距離マスタ');
  if (!sh || sh.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('住所確認待ちのルートはありません。');
    return;
  }
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  var pendingKeys = [];
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][2] || '').trim() === '住所確認待ち') {
      pendingKeys.push(String(data[i][0] || '').trim());
    }
  }
  if (pendingKeys.length === 0) {
    SpreadsheetApp.getUi().alert('住所確認待ちのルートはありません。');
    return;
  }
  var routeKey = pendingKeys[0];
  var remaining = pendingKeys.length;
  var locs = routeKey.split('→');
  var geocoder = Maps.newGeocoder().setRegion('JP');
  var locInfos = [];
  for (var li = 0; li < locs.length; li++) {
    var loc = String(locs[li] || '').trim();
    var candidates = [];
    try {
      var res = geocoder.geocode(loc);
      if (res && res.results) {
        for (var ci = 0; ci < res.results.length; ci++) {
          candidates.push(res.results[ci].formatted_address);
        }
      }
    } catch(e) {}
    locInfos.push({ loc: loc, candidates: candidates });
  }
  PropertiesService.getScriptProperties().setProperty(
    '_addrDialogState_', JSON.stringify({ routeKey: routeKey, locs: locs })
  );
  // 全地名が1件に絞れた → ダイアログなしで直接計算して終了
  var hasAmbiguous = false;
  for (var li2 = 0; li2 < locInfos.length; li2++) {
    if (locInfos[li2].candidates.length >= 2) { hasAmbiguous = true; break; }
  }
  if (!hasAmbiguous) { receiveAddressChoice('{}'); return; }

  var html = '<html><head><meta charset="utf-8">' +
    '<style>body{font-family:sans-serif;font-size:13px;padding:12px;margin:0;}' +
    'h3{font-size:14px;margin:0 0 6px;}' +
    '.sub{color:#555;font-size:12px;margin:0 0 10px;}' +
    '.lb{margin-bottom:10px;padding:8px;background:#f5f5f5;border-radius:4px;}' +
    '.ln{font-weight:bold;margin-bottom:5px;}' +
    'label{display:block;margin:3px 0;cursor:pointer;}' +
    'button{padding:8px 16px;border-radius:4px;border:none;cursor:pointer;font-size:13px;margin-top:10px;}' +
    '.ok{background:#1565c0;color:#fff;margin-right:8px;}.sk{background:#757575;color:#fff;}' +
    '</style></head><body>' +
    '<h3>住所確認（残り' + remaining + '件）</h3>' +
    '<p class="sub">ルート: <strong>' + routeKey + '</strong></p>';
  for (var li3 = 0; li3 < locInfos.length; li3++) {
    var info = locInfos[li3];
    if (info.candidates.length <= 1) continue;
    html += '<div class="lb"><div class="ln">「' + info.loc + '」の候補</div>';
    for (var ci3 = 0; ci3 < info.candidates.length; ci3++) {
      html += '<label><input type="radio" name="loc_' + li3 + '" value="' +
        info.candidates[ci3].replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;') +
        '"' + (ci3 === 0 ? ' checked' : '') + '> ' + info.candidates[ci3] + '</label>';
    }
    html += '</div>';
  }
  html += '<button class="ok" onclick="submit()">この住所で決定</button>' +
    '<button class="sk" onclick="google.script.host.close()">スキップ</button>' +
    '<script>function submit(){' +
    'var sel={};' +
    'document.querySelectorAll("input[type=radio]:checked").forEach(function(r){' +
    'var m=r.name.match(/loc_(\\d+)/);if(m)sel[parseInt(m[1])]=r.value;});' +
    'google.script.run.withSuccessHandler(function(){google.script.host.close();})' +
    '.receiveAddressChoice(JSON.stringify(sel));}' +
    '<\/script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(500).setHeight(420), '住所確認'
  );
}

// 住所選択ダイアログの結果を受け取り距離計算・マスタ更新（google.script.runから呼ばれる）
function receiveAddressChoice(selectionsJson) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stateStr = PropertiesService.getScriptProperties().getProperty('_addrDialogState_');
  if (!stateStr) return;
  PropertiesService.getScriptProperties().deleteProperty('_addrDialogState_');
  var state = JSON.parse(stateStr);
  var locs  = state.locs;
  var sel   = JSON.parse(selectionsJson || '{}');
  var resolvedLocs = locs.slice();
  for (var k in sel) {
    var idx = parseInt(k, 10);
    if (!isNaN(idx) && sel[k]) resolvedLocs[idx] = sel[k];
  }
  var km        = calcDistanceMapsApi_(resolvedLocs);
  var newKm     = km !== null ? km : -1;
  var newMethod = km !== null ? 'API(住所選択)' : 'エラー(住所不明等)';
  updateDistanceMasterRow_(ss, state.routeKey, newKm, newMethod);
  if (km === null) applyDistanceMasterErrorStyle_(ss);
}

// 毎日13時に自動実行されるトリガー関数
function calcDistanceTrigger_() {
  var ssId = PropertiesService.getScriptProperties().getProperty('masterSsId');
  var ss   = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
  calcDistanceForAllPending_(ss);
}

function calcDistanceForSS(ssId) {
  if (!ssId) return 0;
  var ss = SpreadsheetApp.openById(ssId);
  return calcDistanceForAllPending_(ss);
}

// ================================================================
//  4-10: 距離計算手動実行（calcDistanceManual）  【大C / 中4 / 小4-10】
//  メニュー「距離計算（未計算分）」: 積地・降地あり・距離未設定の行をMaps APIで一括計算して集計表に反映
// ================================================================
function calcDistanceManual() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cnt = calcDistanceForAllPending_(ss);
  SpreadsheetApp.getUi().alert(cnt + '件の距離を計算・反映しました。\n\n※ 降完時刻が空の行はスキップします。');
}

// 全国主要都市ペアの距離をMaps APIで一括取得して距離マスタに登録する初期データ投入
function initDistanceMasterMajorCities() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    '全国主要ルートをMaps APIで計算して距離マスタに登録します。\n（約40〜50件・2〜3分かかります）\n\nよろしいですか？',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;

  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var masterMap = loadDistanceMasterMap_(ss);

  // ハブ都市（幹線）
  var hubs = [
    '大阪市', '東京都', '名古屋市', '福岡市', '広島市', '仙台市', '札幌市'
  ];
  // 主要地方都市
  var regionals = [
    '神戸市', '京都市', '岡山市', '高松市', '松山市', '高知市', '徳島市',
    '和歌山市', '奈良市', '津市', '大津市', '金沢市', '富山市', '新潟市',
    '静岡市', '横浜市', 'さいたま市', '千葉市', '宇都宮市', '水戸市',
    '北九州市', '熊本市', '鹿児島市', '長崎市', '大分市', '宮崎市',
    '山口市', '鳥取市', '松江市', '山形市', '盛岡市', '青森市',
    '八尾市', '吹田市', '堺市', '東大阪市', '尼崎市'
  ];
  var allCities = hubs.concat(regionals);

  // 登録するペアを組み立て（ハブ→全都市 + ハブ間双方向）
  var pairs = [];
  for (var h = 0; h < hubs.length; h++) {
    for (var c = 0; c < allCities.length; c++) {
      if (hubs[h] === allCities[c]) continue;
      var key  = hubs[h] + '→' + allCities[c];
      var keyR = allCities[c] + '→' + hubs[h];
      if (masterMap[key]  === undefined) pairs.push({ from: hubs[h],   to: allCities[c], key: key });
      if (masterMap[keyR] === undefined) pairs.push({ from: allCities[c], to: hubs[h],   key: keyR });
    }
  }

  var newRoutes = [];
  var apiCount  = 0;
  var totalSaved = 0;
  for (var i = 0; i < pairs.length; i++) {
    var km = calcDistanceMapsApi_([pairs[i].from, pairs[i].to]);
    if (km !== null) {
      newRoutes.push({ key: pairs[i].key, km: km, method: 'API' });
      masterMap[pairs[i].key] = km;
    }
    apiCount++;
    // 20件ごとに途中保存（タイムアウトしても進捗が消えない）
    if (newRoutes.length > 0 && apiCount % 20 === 0) {
      appendDistanceMasterRows_(ss, newRoutes);
      totalSaved += newRoutes.length;
      newRoutes = [];
    }
    if (apiCount % 10 === 0) Utilities.sleep(500);
    if (apiCount >= 200) break; // 6分タイムアウト回避
  }

  if (newRoutes.length) { appendDistanceMasterRows_(ss, newRoutes); totalSaved += newRoutes.length; }
  ui.alert(totalSaved + '件の主要ルートを距離マスタに登録しました。\n\n残りがある場合はもう一度実行してください。');
}


// ================================================================
//  14-1: インストール型トリガーのセットアップ（installTriggers）  【大C / 中14 / 小14-1】
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

  // 全バインドスクリプト横断で全インストール済みトリガーを強制削除してから3本だけ再登録
  ScriptApp.getUserTriggers(ss).forEach(function(t) {
    try { ScriptApp.deleteTrigger(t); } catch(e) {}
  });
  ScriptApp.newTrigger('installedOnEdit_')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  ScriptApp.newTrigger('onStructureChange_')
    .forSpreadsheet(ss)
    .onChange()
    .create();
  ScriptApp.newTrigger('calcDistanceTrigger_')
    .timeBased()
    .atHour(0)
    .everyDays(1)
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
//  17-7a: PL設定按分を全マスタ行に即時反映（updatePlApportionment_）  【大B / 中17 / 小17-7a】
// ================================================================
function updatePlApportionment_(ss) {
  var plSheet     = ss.getSheetByName('PL設定');
  var masterSheet = ss.getSheetByName('自車専属マスタ');
  if (!plSheet || !masterSheet || masterSheet.getLastRow() < 2) return;

  // PL設定合計（フラグONのみ）
  var plTotal = 0;
  if (plSheet.getLastRow() >= 2) {
    var plData = plSheet.getRange(2, 1, plSheet.getLastRow() - 1, 5).getValues();
    for (var pi = 0; pi < plData.length; pi++) {
      var plName = String(plData[pi][0] || '').trim();
      var plAmt  = Number(plData[pi][1]) || 0;
      var plFlag = plData[pi][4];
      if (!plName) continue;
      if (plFlag === false || String(plFlag).toUpperCase() === 'FALSE') continue;
      plTotal += plAmt;
    }
  }

  // 稼働台数（運行のみ）
  var masterLR  = masterSheet.getLastRow();
  var statusVals = masterSheet.getRange(2, 2, masterLR - 1, 1).getValues();
  var activeCars = 0;
  for (var si = 0; si < statusVals.length; si++) {
    if (String(statusVals[si][0] || '').trim() === '運行') activeCars++;
  }
  if (activeCars < 1) activeCars = 1;
  var plTotalPerCar = Math.round(plTotal / activeCars);

  // AG列をヘッダー名で特定（存在しなければ何もしない）
  var lastCol = masterSheet.getLastColumn();
  var hdrs = masterSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h||'').trim(); });
  var agColIdx = hdrs.indexOf('PL設定按分（参照）');
  if (agColIdx === -1) return;
  var agCol = agColIdx + 1; // 1-based

  // [MOD-v1.3] ループ内setValue廃止→配列構築後に一括setValues/setFontColors/setNumberFormatsへ変更
  var numRows    = statusVals.length;
  var writeVals  = [];
  var writeFonts = [];
  var writeFmts  = [];
  for (var i = 0; i < numRows; i++) {
    var rowStatus = String(statusVals[i][0] || '').trim();
    if (rowStatus === '運行') {
      writeVals.push([plTotalPerCar]);
      writeFonts.push(['#1a9a50']);
      writeFmts.push(['#,##0']);
    } else {
      writeVals.push(['']);
      writeFonts.push(['#000000']);
      writeFmts.push(['']);
    }
  }
  var agRange = masterSheet.getRange(2, agCol, numRows, 1);
  agRange.setValues(writeVals);
  agRange.setFontColors(writeFonts);
  agRange.setNumberFormats(writeFmts);
}

// ================================================================
//  14-2: インストール型onEditトリガー（installedOnEdit_）  【大B / 中14 / 小14-2】
//  会社登録シートのA+B列入力 → processNewCompany_() を自動実行する。
//  インストール型トリガーはDrive/Gmail/ScriptApp等の認証付きサービスが使用可能。
//  実行時間制限もシンプルトリガーの30秒ではなく6分まで利用できる。
//  installTriggers() で登録済みの場合のみ発火する。
// ================================================================
// ================================================================
//  14-2b: インストール型トリガーのディスパッチャ（dispatchInstalledEdit）  【大B / 中14 / 小14-2b】
//  installedOnEdit_ から分離した公開関数。②③スタブからも呼べる（アンダースコアなし）。
//  UI表示（showModalDialog）はスタブ側で行うため、ポップアップ案件は htmlなどを返す。
//  非UI案件は直接処理してnullを返す。
// ================================================================
function dispatchInstalledEdit(e) {
  try {
    var range     = e.range;
    var sheet     = range.getSheet();
    var sheetName = sheet.getName();
    var row       = range.getRow();
    var col       = range.getColumn();

    // ヘッダー行（1行目）が編集された場合は正規ヘッダーに自動復元
    if (row === 1) {
      var canonHdr = getSheetHeaderDef_(sheetName);
      if (canonHdr) {
        if (sheet.getMaxColumns() < canonHdr.length) {
          sheet.insertColumnsAfter(sheet.getMaxColumns(), canonHdr.length - sheet.getMaxColumns());
        }
        sheet.getRange(1, 1, 1, canonHdr.length).setValues([canonHdr]);
      } else if (e.oldValue !== undefined && e.oldValue !== null) {
        range.setValue(e.oldValue);
      }
      return null;
    }

    // 自車専属マスタのB列（運行状態）編集時：適用日確認ポップアップ
    if (sheetName === '自車専属マスタ' && col === 2 && row >= 2) {
      // 複数行同時編集に対応：編集された全行番号を収集
      // e.source はトリガー発火元のSSを確実に返す（ライブラリ内でgetActiveSpreadsheetは使えない）
      var ssId = e.source.getId();
      var numEditRows = range.getNumRows();
      var editedRows = [];
      for (var rIdx = 0; rIdx < numEditRows; rIdx++) {
        if (row + rIdx >= 2) editedRows.push(row + rIdx);
      }
      if (editedRows.length === 0) return null;
      var rowsJson = JSON.stringify(editedRows);
      var mToday = new Date();
      var mTodayStr = Utilities.formatDate(mToday, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      var mFirst = new Date(mToday.getFullYear(), mToday.getMonth(), 1);
      var mFirstStr = Utilities.formatDate(mFirst, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      var countNote = editedRows.length > 1 ? '<br><small>（' + editedRows.length + '台まとめて適用）</small>' : '';
      var htmlStr =
        '<style>' +
        'body{font-family:sans-serif;padding:16px;text-align:center;margin:0;overflow:hidden}' +
        'p{margin:0 0 14px;font-size:13px;line-height:1.5}' +
        'small{color:#888;font-size:11px}' +
        'button{display:block;width:100%;padding:11px 8px;margin:7px 0;font-size:13px;' +
        'cursor:pointer;border:1px solid #ccc;border-radius:4px;background:#fff}' +
        'button:hover:not(:disabled){background:#f0f0f0}' +
        'button:disabled{opacity:.5;cursor:default}' +
        '.cancel{color:#999}' +
        '</style>' +
        '<p>ステータス変更をいつから適用しますか？' + countNote + '<br>（未配車の空行のみ整理します）</p>' +
        '<button onclick="go(\'today\')">本日以降（' + mTodayStr + '〜）</button>' +
        '<button onclick="go(\'month\')">今月以降（' + mFirstStr + '〜）</button>' +
        '<button onclick="go(\'all\')">全期間（制限なし）</button>' +
        '<button class="cancel" onclick="google.script.host.close()">キャンセル</button>' +
        '<script>' +
        'var _rows=' + rowsJson + ';' +
        'var _ssId="' + ssId + '";' +
        'function go(v){' +
        '  document.querySelectorAll("button").forEach(function(b){b.disabled=true;});' +
        '  google.script.run' +
        '    .withSuccessHandler(function(){google.script.host.close();})' +
        '    .withFailureHandler(function(e){alert("エラー: "+(e.message||e));document.querySelectorAll("button").forEach(function(b){b.disabled=false;});})' +
        '    .executeStatusSync(JSON.stringify(_rows),v,_ssId);' +
        '}' +
        '<\/script>';
      return { html: htmlStr, title: 'いつから適用しますか？', width: 300, height: 290 };
    }

    // PL設定シートが変更されたら自車専属マスタのAG列を即時更新
    if (sheetName === 'PL設定' && row >= 2) {
      updatePlApportionment_(e.source);
      return null;
    }

    // 運行シート編集：集計表を30分トリガー（installedOnEdit_）内でsync
    if (sheetName === '運行' && row >= 2) {
      var ss_unk = e.source;
      var allData_unk = sheet.getDataRange().getValues();
      // 編集された行のIDを収集
      var syncIds_set = {}, syncIds_arr = [];
      var numRows_unk = range.getNumRows();
      for (var ri_unk = 0; ri_unk < numRows_unk; ri_unk++) {
        var erow_unk = row + ri_unk;
        if (erow_unk < 2 || erow_unk > allData_unk.length) continue;
        var eid_unk = String(allData_unk[erow_unk - 1][0] || '').trim();
        if (eid_unk && !syncIds_set[eid_unk]) { syncIds_set[eid_unk] = true; syncIds_arr.push(eid_unk); }
      }
      if (syncIds_arr.length > 3) {
        generateSummary(ss_unk);
      } else {
        for (var si_unk = 0; si_unk < syncIds_arr.length; si_unk++) {
          try { syncSummaryForId_(syncIds_arr[si_unk], ss_unk); } catch(e_unk) {}
        }
      }
      // L列（積地）またはM列（降地）の場合は距離も即時反映
      if (col === 12 || col === 13) {
        var id_dist = String(sheet.getRange(row, 1).getValue() || '').trim();
        if (id_dist) {
          var picksArr_ = [], dropsArr_ = [];
          for (var ri_ = 1; ri_ < allData_unk.length; ri_++) {
            if (String(allData_unk[ri_][0] || '').trim() !== id_dist) continue;
            var pv_ = String(allData_unk[ri_][11] || '').trim();
            var dv_ = String(allData_unk[ri_][12] || '').trim();
            if (pv_) picksArr_.push(pv_);
            if (dv_) dropsArr_.push(dv_);
          }
          if (picksArr_.length && dropsArr_.length) {
            lookupAndSetDistanceAfterCreate_(ss_unk, id_dist, picksArr_, dropsArr_);
          }
        }
      }
      // 編集行をバックアップ（列削除後のデータ復元用）
      try {
        var bkRows_ = range.getNumRows();
        for (var bkR_ = 0; bkR_ < bkRows_; bkR_++) backupSheetRow_(ss_unk, '運行', row + bkR_);
      } catch(bkEx_) {}
      return null;
    }

    // 正規定義シートのデータ行をバックアップ（列削除後のデータ復元用）
    if (row >= 2 && getSheetHeaderDef_(sheetName)) {
      try {
        var bkN_ = range.getNumRows();
        for (var bkI_ = 0; bkI_ < bkN_; bkI_++) backupSheetRow_(e.source, sheetName, row + bkI_);
      } catch(bkEx2_) {}
    }

    if (sheetName !== '会社登録' || row <= 1) return null;

    // A列 or B列: 会社名+Gmail が揃ったらフルセットアップ
    if (col === 1 || col === 2) {
      var companyName = String(sheet.getRange(row, 1).getValue() || '').trim();
      var adminEmail  = String(sheet.getRange(row, 2).getValue() || '').trim();
      var status      = String(sheet.getRange(row, 3).getValue() || '').trim();
      if (!companyName || !adminEmail || adminEmail.indexOf('@') === -1) return null;
      if (status !== '' && status.indexOf('エラー') !== 0) return null;
      sheet.getRange(row, 3).setValue('処理中...').setBackground('#fff9c4');
      try {
        processNewCompany_(companyName, adminEmail);
      } catch(err) {
        sheet.getRange(row, 3).setValue('エラー: ' + err.message).setBackground('#ffcdd2');
      }
      return null;
    }

    // F列 or G列: SS URL + App URL が揃ったら配布メール送信
    if (col === 6 || col === 7) {
      var ssUrl      = String(sheet.getRange(row, 6).getValue() || '').trim();
      var appUrl     = String(sheet.getRange(row, 7).getValue() || '').trim();
      var mailStatus = String(sheet.getRange(row, 8).getValue() || '').trim();
      if (!ssUrl || !appUrl) return null;
      if (mailStatus.indexOf('送信済') !== -1) return null;
      var cName  = String(sheet.getRange(row, 1).getValue() || '').trim();
      var aEmail = String(sheet.getRange(row, 2).getValue() || '').trim();
      if (!cName || !aEmail || aEmail.indexOf('@') === -1) return null;
      try {
        sendDistributionMail_(cName, aEmail, ssUrl, appUrl, row, sheet);
      } catch(err) {
        sheet.getRange(row, 8).setValue('エラー: ' + err.message).setBackground('#ffcdd2');
      }
    }
  } catch(ex) {}
  return null;
}


function installedOnEdit_(e) {
  var result = dispatchInstalledEdit(e);
  if (result && result.html) {
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutput(result.html).setWidth(result.width).setHeight(result.height),
      result.title
    );
  }
}


// ================================================================
//  17-7b: マスタ変更時の再計算範囲保存（setRecalcChoice）  【大A / 中17 / 小17-7b】
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
}


// ================================================================
//  12-9: 修正用SSを作成（createDevSs）  【大C / 中12 / 小12-9】
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
//  12-10: このSSのスクリプトIDを確認（showMyScriptId）  【大C / 中12 / 小12-10】
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
//
//  ★ M&A後の引き継ぎメモ ★
//  この関数は①からも②からも実行可能。動作の違いは以下の通り：
//  ・①から実行 → activeSpreadsheetが①、__TEMPLATE_SS__なし → elseブランチ → ScriptPropertiesから approvedLibVersion 取得
//  ・②から実行 → activeSpreadsheetが②、__TEMPLATE_SS__あり → ifブランチ → __TEMPLATE_SS__ C1から取得
//
//  ★ 注意：必ず STEP2「テスト客SS反映」を先に実行してから STEP4「各客に反映」を実行すること ★
//  STEP2 を省くと approvedVersion が古いままになり、③のライブラリが旧バージョンを参照し続ける。
//
//  ★ 各客SSへのコード配布の仕組み ★
//  updateStubVersion_() が Script API 経由で各③のスクリプトを以下の2点更新する：
//  [1] appsscript.json → UnkouLibのバージョン番号を最新に書き換え
//  [2] コード.js（SERVER_JS）→ getClientStubSource_() の最新内容に書き換え
//  これにより③のF5だけでメニュー・スタブ関数が全て最新になる。
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

  // ヘッダー行からスクリプトID列を動的に探す（列が増減しても壊れないよう）
  var lastCol = regSheet.getLastColumn();
  var regHdrs = regSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h||'').trim(); });
  var scriptIdColIdx = regHdrs.indexOf('スクリプトID'); // 0-indexed。-1なら未設定

  var businessSheets = ['運行', '集計表', '自車専属マスタ', '自車専属運行', 'マスタ', '設定',
                        '配車板', '自社設定', '距離マスタ', 'PL設定', '管理者'];
  var fetchCols = Math.max(lastCol, 13);
  var rows = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, fetchCols).getValues();
  var successCount = 0;
  var errorNames   = [];
  var neutralizedTotal = 0;
  // 全バインドスクリプトの対応表を一度だけ構築（SSID→スクリプトID配列）
  var ssScriptMap = buildSsScriptMap_();

  for (var i = 0; i < rows.length; i++) {
    var companyName  = String(rows[i][0]).trim();
    var ssUrl        = String(rows[i][5]).trim();  // F列: SS URL
    // スクリプトIDをヘッダー名で動的に探す（K列固定をやめる）
    var clientScriptId = scriptIdColIdx >= 0 ? String(rows[i][scriptIdColIdx] || '').trim() : '';
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
        // ①の列数を超えた余分な列を消去（②③が勝手に列を増やしても①の構造に強制一致させる）
        var _tgtCols2 = tgtSheet.getLastColumn();
        if (_tgtCols2 > srcCols) {
          tgtSheet.getRange(1, srcCols + 1, 1, _tgtCols2 - srcCols).clearContent().clearFormat();
        }
        tgtSheet.setFrozenRows(1);
      }
      ensureSettingItems_(clientSs);
      initImportDictionary_(clientSs);

      // 自車専属マスタ B列（運行状態）にドロップダウンを適用
      var mSheet = clientSs.getSheetByName('自車専属マスタ');
      if (mSheet && mSheet.getMaxRows() > 1) {
        var sv = SpreadsheetApp.newDataValidation()
          .requireValueInList(['運行','故障','待機'], true)
          .setAllowInvalid(false).build();
        mSheet.getRange(2, 2, mSheet.getMaxRows() - 1, 1).setDataValidation(sv);
      }

      // 自車専属マスタのコンプライアンス日付列フォーマットとAF〜AMヘッダー色を③にも適用
      if (mSheet && mSheet.getLastRow() >= 2) {
        var _cmh = mSheet.getRange(1, 1, 1, mSheet.getLastColumn()).getValues()[0];
        var _cmr = mSheet.getLastRow() - 1;
        ['免許証有効期限','安全教育次回予定日','健康診断次回予定日','適性診断次回予定日'].forEach(function(cn) {
          var ci = _cmh.indexOf(cn);
          if (ci >= 0) mSheet.getRange(2, ci + 1, _cmr, 1).setNumberFormat('yyyy/M/d');
        });
        var _cafI = _cmh.indexOf('その他固定費'), _camI = _cmh.indexOf('担当管理者');
        if (_cafI >= 0 && _camI >= _cafI) {
          mSheet.getRange(1, _cafI + 1, 1, _camI - _cafI + 1).setBackground('#212121').setFontColor('#00e676').setFontWeight('bold');
        }
      }

      // 不要シートを削除（マスタ点検項目など）、ただし業務シートは保護
      var validClientSheets = [
        '運行','集計表','自車専属マスタ','自車専属運行','マスタ','設定','メモ',
        '使い方','説明書','サポート','配車板','距離マスタ','受領書_耳','受領',
        'PL','PL設定','仕訳表','監査用','自社設定','管理者',
        '請求書','支払確認書','_ErrorLog_','指示先履歴','指示先ID別','__COMPANY_SS__'
      ];
      clientSs.getSheets().forEach(function(s) {
        var sn = s.getName();
        // PL または PL_で始まる月次PLシートも保護
        var isValid = validClientSheets.indexOf(sn) >= 0 || sn === 'PL' || sn.indexOf('PL_') === 0;
        if (!isValid && clientSs.getSheets().length > 1) {
          try { clientSs.deleteSheet(s); } catch(e) {}
        }
      });

      // projects.getのparentId照合による確実な全列挙（Drive親検索は廃止）
      var allScriptIds = ssScriptMap[clientSsId] || [];
      var webAppUrlCell = String(rows[i][6] || '').trim();  // G列: WebアプリURL
      var primaryId = '';
      var primaryConfirmed = false;  // 正スクリプトを証拠付きで確定できたか
      if (allScriptIds.length === 1) {
        primaryId = allScriptIds[0];
        primaryConfirmed = true;
      } else if (allScriptIds.length > 1) {
        // 会社登録G列のWebアプリURL（＝実際に動いているWebApp）を持つスクリプトが正
        primaryId = findWebAppScriptId_(allScriptIds, webAppUrlCell);
        if (primaryId) {
          primaryConfirmed = true;
        } else if (clientScriptId && allScriptIds.indexOf(clientScriptId) >= 0) {
          primaryId = clientScriptId;
          primaryConfirmed = true;
        } else {
          primaryId = allScriptIds[0];  // 確定できない場合は更新のみ行い、無効化はしない（安全側）
        }
      }
      if (primaryId) {
        if (primaryId !== clientScriptId) {
          var sidCol1 = scriptIdColIdx >= 0 ? scriptIdColIdx + 1 : regSheet.getLastColumn() + 1;
          regSheet.getRange(i + 2, sidCol1).setValue(primaryId);
        }
        clientScriptId = primaryId;
      }
      // Drive検索が空の場合は既存値をそのまま使用

      var stubOk = false;
      if (clientScriptId) {
        // スクリプトIDあり → スタブを最新化
        var stubResult = updateStubVersion_(clientScriptId, approvedVersion || '', false);
        stubOk = stubResult && stubResult.ok;
        if (!stubOk) {
          errorNames.push(companyName + '（API更新失敗: ' + (stubResult ? stubResult.error : '不明') + '）');
        }
        // 正スクリプトを確定できた場合のみ、余分な旧スクリプトを空コードで無効化
        if (stubOk && primaryConfirmed) {
          for (var ni = 0; ni < allScriptIds.length; ni++) {
            if (allScriptIds[ni] === clientScriptId) continue;
            if (neutralizeScript_(allScriptIds[ni])) neutralizedTotal++;
            else errorNames.push(companyName + '（余分スクリプト無効化失敗）');
          }
        }
      } else if (approvedVersion) {
        // スクリプトIDがどうしても取れない場合のみ新規デプロイ
        var deployResult = deployClientWebApp_(clientSsId, companyName, null, approvedVersion);
        if (deployResult && deployResult.scriptId) {
          if (scriptIdColIdx >= 0) regSheet.getRange(i + 2, scriptIdColIdx + 1).setValue(deployResult.scriptId);
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
  msg += '\n無効化した余分スクリプト：' + neutralizedTotal + '個';
  if (errorNames.length > 0) msg += '\n失敗: ' + errorNames.join(', ');
  msg += '\n\n各客SSでF5を押すとメニューが更新されます。';
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}

// ================================================================
//  14-3a: 全バインドスクリプトの対応表を作る（buildSsScriptMap_）
//  syncToAllClientSS専用。Drive親検索（"SSID" in parents）は結果が欠落することが
//  あるため廃止し、実行者から見える全スクリプトファイルを広域列挙 →
//  Apps Script API projects.get の parentId（バインド先SSのID）で確実に照合する。
//  戻り値: { <SSのID>: [scriptId, ...], ... }
// ================================================================
function buildSsScriptMap_() {
  var token = ScriptApp.getOAuthToken();
  var fileIds = [];
  try {
    var pageToken = '';
    do {
      var url = 'https://www.googleapis.com/drive/v3/files?q=' +
        encodeURIComponent("mimeType='application/vnd.google-apps.script' and trashed=false") +
        '&fields=nextPageToken,files(id)&pageSize=100&supportsAllDrives=true&includeItemsFromAllDrives=true' +
        (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
      var resp = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) break;
      var data = JSON.parse(resp.getContentText());
      (data.files || []).forEach(function(f) { fileIds.push(f.id); });
      pageToken = data.nextPageToken || '';
    } while (pageToken && fileIds.length < 500);
  } catch(e) {}

  var map = {};
  // projects.get を20件ずつ並列実行（parentId = バインド先SSのID）
  for (var i = 0; i < fileIds.length; i += 20) {
    var reqs = fileIds.slice(i, i + 20).map(function(id) {
      return { url: 'https://script.googleapis.com/v1/projects/' + id,
               headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true };
    });
    try {
      var resps = UrlFetchApp.fetchAll(reqs);
      for (var j = 0; j < resps.length; j++) {
        try {
          if (resps[j].getResponseCode() !== 200) continue;
          var p = JSON.parse(resps[j].getContentText());
          if (p.parentId && p.scriptId) {
            if (!map[p.parentId]) map[p.parentId] = [];
            map[p.parentId].push(p.scriptId);
          }
        } catch(e2) {}
      }
    } catch(e3) {}
  }
  return map;
}

// ================================================================
//  14-3b: WebアプリURL（会社登録G列）のデプロイIDを持つスクリプトを正として特定（findWebAppScriptId_）
//  複数バインドスクリプトがある場合、実際にWebAppが動いている方を正スクリプトと判定する。
// ================================================================
function findWebAppScriptId_(scriptIds, webAppUrl) {
  var depId = (String(webAppUrl || '').match(/\/s\/([A-Za-z0-9_-]+)\/exec/) || [])[1] || '';
  var token = ScriptApp.getOAuthToken();
  var fallback = '';
  for (var i = 0; i < scriptIds.length; i++) {
    try {
      var resp = UrlFetchApp.fetch(
        'https://script.googleapis.com/v1/projects/' + scriptIds[i] + '/deployments?pageSize=50',
        { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
      );
      if (resp.getResponseCode() !== 200) continue;
      var deps = JSON.parse(resp.getContentText()).deployments || [];
      for (var j = 0; j < deps.length; j++) {
        if (depId && deps[j].deploymentId === depId) return scriptIds[i];  // G列URLと一致→確定
        var eps = deps[j].entryPoints || [];
        for (var k = 0; k < eps.length; k++) {
          if (eps[k].entryPointType === 'WEB_APP' && !fallback) fallback = scriptIds[i];
        }
      }
    } catch(e) {}
  }
  return fallback;
}

// ================================================================
//  14-3c: 余分なバインドスクリプトを無効化（neutralizeScript_）
//  ライブラリ依存なし・全処理を空関数化した最小コードに上書きする。
//  バインドスクリプトはAPI削除不可のため無効化で対応。
//  残存インストール済みトリガーが発火しても空関数のため何も起きない
//  （関数未定義エラーによる通知メールも発生しない）。
// ================================================================
function neutralizeScript_(scriptId) {
  try {
    var token = ScriptApp.getOAuthToken();
    var manifest = JSON.stringify({ timeZone: 'Asia/Tokyo', dependencies: {}, exceptionLogging: 'STACKDRIVER', runtimeVersion: 'V8' });
    var src =
      '// 無効化済み（自己消去型）：発火するたびに自スクリプトの残存トリガーを全削除\n' +
      'function _selfClean_() {\n' +
      '  try { ScriptApp.getProjectTriggers().forEach(function(t) { try { ScriptApp.deleteTrigger(t); } catch(ex) {} }); } catch(ex) {}\n' +
      '}\n' +
      'function onOpen(e) { _selfClean_(); }\n' +
      'function onEdit(e) { _selfClean_(); }\n' +
      'function installedOnEdit_(e) { _selfClean_(); }\n' +
      'function onStructureChange_(e) { _selfClean_(); }\n' +
      'function checkMasterExpiries(e) { _selfClean_(); }\n' +
      'function checkExpiryDates(e) { _selfClean_(); }\n' +
      'function calcDistanceTrigger_(e) { _selfClean_(); }\n' +
      'function onFormSubmit_(e) { _selfClean_(); }\n' +
      'function processUploadQueue(e) { _selfClean_(); }\n' +
      'function runDailyBackup_(e) { _selfClean_(); }\n' +
      'function scheduledGenerateNextMonth_(e) { _selfClean_(); }\n';
    var resp = UrlFetchApp.fetch(
      'https://script.googleapis.com/v1/projects/' + scriptId + '/content',
      { method: 'put',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify({ files: [
          { name: 'appsscript', type: 'JSON', source: manifest },
          { name: 'コード', type: 'SERVER_JS', source: src }
        ]}),
        muteHttpExceptions: true }
    );
    return resp.getResponseCode() === 200;
  } catch(e) { return false; }
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
      var id = prefix + '-' + String(nextNum + i).padStart(4, '0'); // padStartは4桁超も切り捨てない
      writeRows.push(buildSheetRow_(sheetType, id, mappedRows[i], ss));
    }
    commitLastId_(sheet, prefix, nextNum + mappedRows.length - 1);
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
      // 集計表を一括再生成（行ごとの個別syncより大幅に高速）
      generateSummary(ss);
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

  // H1ヘッダー + H2以降にデータがある場合のみスキップ（反映でH1だけコピーされた状態では初期化を実行する）
  if (String(setting.getRange(1, 8).getValue()).trim() === '【辞書v3】種別'
      && setting.getLastRow() >= 2
      && String(setting.getRange(2, 8).getValue()).trim() !== '') return;

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
//  14-4: 帳票発行済マーク（markDocumentIssued）  【大A / 中14 / 小14-4】
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
      sheet.getRange(i + 2, colIdx + 1).setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd HH:mm'));
      SpreadsheetApp.getActiveSpreadsheet().toast(
        colName + ' を発行済にしました（ID: ' + rowId + '）', '✅', 3
      );
      return;
    }
  }
}


// ================================================================
//  14-5: 帳票メール／FAX送信（sendDocumentEmail）  【大A / 中14 / 小14-5】
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
    if (docData.shijisaki) { try { saveShijisakiHistory_(String(docData.client||''), docData.shijisaki, ss); } catch(e) {} }
    return { ok: true, msg: email + ' に送信しました。' };

  } else if (method === 'fax') {
    var fax = (faxIdx >= 0) ? String(custRow[faxIdx]||'').trim() : '';
    if (!fax) return { ok: false, msg: 'FAX番号が未登録です。\nマスタシートのD列に登録してください。' };
    Logger.log('[FAX送信予約] 宛先:' + fax + ' 書類:' + docType + ' ID:' + docData.id);
    markDocumentIssued(docData.id, docType);
    if (docData.shijisaki) { try { saveShijisakiHistory_(String(docData.client||''), docData.shijisaki, ss); } catch(e) {} }
    return { ok: true, msg: 'FAX送信を予約しました。（宛先：' + fax + '）' };
  }

  return { ok: false, msg: '不明な送信方法です。' };
}


// ================================================================
//  14-6: 指示先履歴取得（getShijisakiHistory）  【大A / 中14 / 小14-6】
//  発注書ダイアログの「📋 履歴」ボタンから google.script.run 経由で呼ばれる
//  荷主名に紐づく全件を返す（最終使用日の降順）
// ================================================================
function getShijisakiHistory(clientName, ssId) {
  var ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('指示先履歴');
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  var results = [];
  data.forEach(function(r) {
    if (String(r[0]||'').trim() !== String(clientName||'').trim()) return;
    var lastUsed = '';
    if (r[5]) {
      try { lastUsed = Utilities.formatDate(r[5] instanceof Date ? r[5] : new Date(r[5]), 'Asia/Tokyo', 'yyyy/MM/dd'); } catch(ex) {}
    }
    results.push({ company: String(r[1]||''), tel: String(r[2]||''), person: String(r[3]||''), addr: String(r[4]||''), lastUsed: lastUsed });
  });
  results.sort(function(a, b) { return (b.lastUsed||'').localeCompare(a.lastUsed||''); });
  return results;
}

// ================================================================
//  14-7: 指示先履歴保存・公開ラッパー（saveShijisakiHistory）  【大A / 中14 / 小14-7】
//  印刷時など sendDocumentEmail を経由しない場合に HTML から直接呼ばれる
// ================================================================
function saveShijisakiHistory(clientName, shijisaki, ssId) {
  var ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
  saveShijisakiHistory_(clientName, shijisaki, ss);
}

// ================================================================
//  14-7i: 指示先履歴保存・内部実装（saveShijisakiHistory_）  【大B / 中14 / 小14-7i】
//  重複チェック: 荷主名＋会社名＋電話＋担当者＋住所が同一なら最終使用日を更新のみ
// ================================================================
function saveShijisakiHistory_(clientName, shijiData, ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var company = String(shijiData.company||'').trim();
  var tel     = String(shijiData.tel    ||'').trim();
  var person  = String(shijiData.person ||'').trim();
  var addr    = String(shijiData.addr   ||'').trim();
  if (!company && !tel && !person && !addr) return;
  var sh = ss.getSheetByName('指示先履歴');
  if (!sh) {
    sh = ss.insertSheet('指示先履歴');
    sh.getRange(1, 1, 1, 6).setValues([['荷主名', '指示先会社名', '電話番号', '担当者名', '住所', '最終使用日']]);
    sh.hideSheet();
  }
  var cName = String(clientName||'').trim();
  var now = new Date();
  if (sh.getLastRow() >= 2) {
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]||'').trim() === cName &&
          String(rows[i][1]||'').trim() === company &&
          String(rows[i][2]||'').trim() === tel &&
          String(rows[i][3]||'').trim() === person &&
          String(rows[i][4]||'').trim() === addr) {
        sh.getRange(i + 2, 6).setValue(now);
        return;
      }
    }
  }
  sh.appendRow([cName, company, tel, person, addr, now]);
}


// ================================================================
//  14-8: 行ID別指示先取得（getShijisakiByRowId）  【大A / 中14 / 小14-8】
//  発注書ダイアログを開いた時にそのIDの保存済み指示先情報を返す
// ================================================================
function getShijisakiByRowId(rowId, ssId) {
  var ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('指示先ID別');
  if (!sh || sh.getLastRow() < 2) return null;
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  var rid = String(rowId||'').trim();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]||'').trim() === rid) {
      return { company: String(data[i][1]||''), tel: String(data[i][2]||''), person: String(data[i][3]||''), addr: String(data[i][4]||'') };
    }
  }
  return null;
}

// ================================================================
//  14-9: 行ID別指示先保存（saveShijisakiByRowId）  【大A / 中14 / 小14-9】
//  入力中にリアルタイム保存（0.8秒デバウンス）＋指示先履歴にも反映
// ================================================================
function saveShijisakiByRowId(rowId, shijiData, clientName, ssId) {
  var ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
  var company = String(shijiData.company||'').trim();
  var tel     = String(shijiData.tel    ||'').trim();
  var person  = String(shijiData.person ||'').trim();
  var addr    = String(shijiData.addr   ||'').trim();
  var rid     = String(rowId||'').trim();

  var sh = ss.getSheetByName('指示先ID別');
  if (!sh) {
    sh = ss.insertSheet('指示先ID別');
    sh.getRange(1, 1, 1, 5).setValues([['行ID', '指示先会社名', '電話番号', '担当者名', '住所']]);
    sh.hideSheet();
  }
  var updated = false;
  if (sh.getLastRow() >= 2) {
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]||'').trim() === rid) {
        sh.getRange(i + 2, 2, 1, 4).setValues([[company, tel, person, addr]]);
        updated = true;
        break;
      }
    }
  }
  if (!updated) sh.appendRow([rid, company, tel, person, addr]);

  if (company || tel || person || addr) {
    saveShijisakiHistory_(String(clientName||''), shijiData, ss);
  }
}


// ================================================================
//  14-10: 指示先履歴削除（deleteShijisakiHistory）  【大A / 中14 / 小14-10】
//  履歴モーダルの「削除」ボタンから google.script.run 経由で呼ばれる
// ================================================================
function deleteShijisakiHistory(clientName, company, tel, person, addr, ssId) {
  var ss = ssId ? SpreadsheetApp.openById(ssId) : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('指示先履歴');
  if (!sh || sh.getLastRow() < 2) return;
  var cName = String(clientName||'').trim();
  var cComp = String(company    ||'').trim();
  var cTel  = String(tel        ||'').trim();
  var cPer  = String(person     ||'').trim();
  var cAddr = String(addr       ||'').trim();
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]||'').trim() === cName &&
        String(rows[i][1]||'').trim() === cComp &&
        String(rows[i][2]||'').trim() === cTel  &&
        String(rows[i][3]||'').trim() === cPer  &&
        String(rows[i][4]||'').trim() === cAddr) {
      sh.deleteRow(i + 2);
      return;
    }
  }
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
  var joho = ss.getSheetByName('配車板');
  if (!joho || joho.getLastRow() < 2) {
    ui.alert('「配車板」シートにデータがありません。\nメニュー→シート再生成で配車板シートを作成してください。');
    return;
  }

  // ── A列(貨物チェック)とO列(車両チェック)を独立して収集 ──────────────
  // 列構成（0ベース）: A=0:チェック(貨物) / N=13:貨物登録ID / O=14:チェック(車両)
  var lastRow = joho.getLastRow();
  var allData = joho.getRange(2, 1, lastRow - 1, 29).getValues();

  var cargoRows = []; // A列(index 0)が true の行 → 貨物として使用
  var vehRows   = []; // O列(index 14)が true の行 → 車両として使用
  for (var i = 0; i < allData.length; i++) {
    var rn = i + 2;
    if (allData[i][0]  === true) cargoRows.push({ rowNum: rn, data: allData[i] });
    if (allData[i][14] === true) vehRows.push(  { rowNum: rn, data: allData[i] });
  }

  // ── 選択バリデーション ────────────────────────────────────────────
  if (cargoRows.length === 0 && vehRows.length === 0) {
    ui.alert('行を選択してください。\nA列（貨物チェック）またはO列（車両チェック）にチェックを入れてから実行してください。');
    return;
  }
  if (cargoRows.length > 2) {
    ui.alert('貨物チェック（A列）は2行までにしてください。\n3行程以上は個別に実行してください。'); return;
  }
  if (vehRows.length > 1) {
    ui.alert('車両チェック（O列）は1行だけにしてください。'); return;
  }
  if (cargoRows.length === 2 && vehRows.length === 0) {
    ui.alert('貨物が2行選択されています。\n対応する車両行のN列もチェックしてから実行してください。'); return;
  }
  // ピンク行（両側確定済み）は再マッチング不可
  var allCheckedRows = cargoRows.concat(vehRows.filter(function(vr) {
    return !cargoRows.some(function(cr) { return cr.rowNum === vr.rowNum; });
  }));
  for (var vx = 0; vx < allCheckedRows.length; vx++) {
    var rd = allCheckedRows[vx].data;
    if (String(rd[1]).trim() === '確定' && String(rd[15]).trim() === '確定') {
      ui.alert('配車確定済み（ピンク）の行は再マッチングできません。\nチェックを外してください（行' + allCheckedRows[vx].rowNum + '）。');
      return;
    }
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
    if (!cData || !vData) return vData ? String(vData[22]||'').trim() : (cData ? String(cData[8]||'').trim() : null);
    var cType = String(cData[8]  || '').trim(); // I列=車種(貨物要求)
    var vType = String(vData[22] || '').trim(); // W列(index22)=車種(車両) ※29列構成
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

  // ── 運行シートへの追記（既存IDあれば更新、なければ新規） ────────────
  var registeredIds = [];
  function addUnkouRow(cData, vData, finalType, presetId) {
    var nid = presetId || ('V-' + String(getNextIdNum_(unkou, 'V-')).padStart(4, '0'));
    var idColIdx = uIdx('ID');
    var existRow = -1;
    if (presetId && idColIdx >= 0 && unkou.getLastRow() >= 2) {
      var srch = unkou.getRange(2, idColIdx + 1, unkou.getLastRow() - 1, 1).getValues();
      for (var ek = 0; ek < srch.length; ek++) {
        if (String(srch[ek][0]).trim() === nid) { existRow = ek + 2; break; }
      }
    }
    if (existRow > 0) {
      var updRow = unkou.getRange(existRow, 1, 1, uLastCol).getValues()[0];
      buildJohoNewRow_(updRow, uIdx, cData, vData, finalType);
      unkou.getRange(existRow, 1, 1, uLastCol).setValues([updRow]);
    } else {
      var row = [];
      for (var n = 0; n < uLastCol; n++) row.push('');
      buildJohoNewRow_(row, uIdx, cData, vData, finalType);
      if (uIdx('ID') >= 0) row[uIdx('ID')] = nid;
      var ins = unkou.getLastRow() + 1;
      unkou.getRange(ins, 1, 1, uLastCol).setValues([row]);
      if (uIdx('日付') >= 0) unkou.getRange(ins, uIdx('日付') + 1).setNumberFormat('yyyy/MM/dd');
    }
    if (registeredIds.indexOf(nid) === -1) registeredIds.push(nid);
    try { delaySyncSummary_(nid, ss); } catch(e) {}
    return nid;
  }

  // ── 既存IDを優先取得（片方確定済みのIDを引き継ぐ） ────────────────
  function detectExistId_(cData, vData) {
    if (cData && String(cData[13] || '').trim()) return String(cData[13]).trim();
    if (vData && String(vData[28] || '').trim()) return String(vData[28]).trim();
    return '';
  }

  // ── パターン分岐 ──────────────────────────────────────────────────
  if (cargoRows.length === 2 && vehRows.length === 1) {
    var ft = checkCarType(cargoRows[0].data, vehData);
    if (ft === null) return;
    var existIdA = detectExistId_(cargoRows[0].data, null) || detectExistId_(cargoRows[1].data, null) || detectExistId_(null, vehData);
    var sharedId = existIdA || ('V-' + String(getNextIdNum_(unkou, 'V-')).padStart(4, '0'));
    addUnkouRow(cargoRows[0].data, vehData, ft, sharedId);
    addUnkouRow(cargoRows[1].data, vehData, ft, sharedId);
  } else {
    var cargoData = cargoRows.length > 0 ? cargoRows[0].data : null;
    var isSameRow = cargoData && vehData && (cargoRows[0].rowNum === vehRows[0].rowNum);
    var ft2 = (cargoData && vehData && !isSameRow) ? checkCarType(cargoData, vehData) : (vehData ? String(vehData[22]||'').trim() : null);
    if (ft2 === null && cargoData && vehData && !isSameRow) return;
    var existIdB = detectExistId_(cargoData, vehData);
    addUnkouRow(cargoData, vehData, ft2, existIdB || undefined);
  }

  // ── 全マッチ行を両側確定（B・P列）＋ピンク着色 ──────────────────────
  var regId = registeredIds.length > 0 ? registeredIds[0] : '';
  var allMatchRows = [];
  for (var cw = 0; cw < cargoRows.length; cw++) {
    if (allMatchRows.indexOf(cargoRows[cw].rowNum) === -1) allMatchRows.push(cargoRows[cw].rowNum);
  }
  if (vehRows.length > 0 && allMatchRows.indexOf(vehRows[0].rowNum) === -1) allMatchRows.push(vehRows[0].rowNum);
  for (var mi = 0; mi < allMatchRows.length; mi++) {
    var mr = allMatchRows[mi];
    if (regId) { joho.getRange(mr, 14).setValue(regId); joho.getRange(mr, 29).setValue(regId); }
    joho.getRange(mr, 1).setValue(false);  joho.getRange(mr, 2).setValue('確定');
    joho.getRange(mr, 15).setValue(false); joho.getRange(mr, 16).setValue('確定');
    joho.getRange(mr, 1,  1, 14).setBackground('#f8bbd0');
    joho.getRange(mr, 15, 1, 15).setBackground('#f8bbd0');
  }

  var idMsg = registeredIds.length > 1
    ? registeredIds[0] + ' の2行程（同一ID・2行）'
    : registeredIds[0];
  ui.alert('✅ 配車確定\n\n' + idMsg + ' を運行シートに登録しました。\n\n情報シートの対象行をピンク（両側確定）にしました。');
}


// ================================================================
//  15-1b: マッチング解除（cancelDispatch）  【大C / 中15 / 小15-1b】
//  情報シートで選択中の行のAA列（登録ID）を読み取り、
//  運行シート・集計表から該当行を削除して情報シートを初期化する
// ================================================================
function cancelDispatch() {
  var ui       = SpreadsheetApp.getUi();
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var joho     = ss.getSheetByName('配車板');
  if (!joho) { ui.alert('配車板シートが見つかりません。'); return; }

  var curSheet = ss.getActiveSheet();
  var sel      = ss.getActiveRange();
  var sRow     = sel.getRow();
  var eRow     = sRow + sel.getNumRows() - 1;
  if (sRow < 2) { ui.alert('データ行を選択してから実行してください。'); return; }

  var ids = {};

  if (curSheet.getName() === '配車板') {
    // 配車板シートから: N列(14=貨物登録ID) と AB列(28=車両登録ID) 両方確認
    var johoLast0 = joho.getLastRow();
    for (var r = sRow; r <= Math.min(eRow, johoLast0); r++) {
      var aa = String(joho.getRange(r, 14).getValue() || '').trim(); // N=貨物登録ID
      var ab = String(joho.getRange(r, 29).getValue() || '').trim(); // AC=車両登録ID
      if (aa) ids[aa] = true;
      if (ab) ids[ab] = true;
    }
  } else if (curSheet.getName() === '運行') {
    // 運行シートから: A列(1=ID)を直接読む
    for (var r2 = sRow; r2 <= eRow; r2++) {
      var rid = String(curSheet.getRange(r2, 1).getValue() || '').trim();
      if (rid) ids[rid] = true;
    }
  } else {
    ui.alert('情報シートまたは運行シートで行を選択してから実行してください。'); return;
  }

  if (Object.keys(ids).length === 0) {
    ui.alert('選択行に登録IDがありません。\nすでに解除済みか、まだ運行登録されていない行です。'); return;
  }

  var idList = Object.keys(ids);
  var res = ui.alert(
    'マッチング解除',
    '以下のIDを運行シート・集計表から削除します。\n' + idList.join(', ') + '\n\n情報シートの対象行は未確定状態に戻します。よろしいですか？',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  // 運行シートから削除（後ろから削除してインデックスずれを防ぐ）
  var unkou = ss.getSheetByName('運行');
  if (unkou && unkou.getLastRow() >= 2) {
    var uData = unkou.getRange(2, 1, unkou.getLastRow() - 1, 1).getValues();
    for (var ud = uData.length - 1; ud >= 0; ud--) {
      if (ids[String(uData[ud][0] || '').trim()]) unkou.deleteRow(ud + 2);
    }
  }

  // 集計表から削除
  var sumSh = ss.getSheetByName('集計表');
  if (sumSh && sumSh.getLastRow() >= 2) {
    var sData = sumSh.getRange(2, 1, sumSh.getLastRow() - 1, 1).getValues();
    for (var sd = sData.length - 1; sd >= 0; sd--) {
      if (ids[String(sData[sd][0] || '').trim()]) sumSh.deleteRow(sd + 2);
    }
  }

  // 情報シートのN(14=貨物登録ID)またはAB(28=車両登録ID)が一致する全行をリセット
  var johoLast = joho.getLastRow();
  if (johoLast >= 2) {
    var nVals  = joho.getRange(2, 14, johoLast - 1, 1).getValues(); // N=貨物登録ID
    var abVals = joho.getRange(2, 29, johoLast - 1, 1).getValues(); // AC=車両登録ID
    for (var ji = 0; ji < nVals.length; ji++) {
      var aaId = String(nVals[ji][0]  || '').trim();
      var abId = String(abVals[ji][0] || '').trim();
      if (ids[aaId] || ids[abId]) {
        var jr = ji + 2;
        if (ids[aaId]) { joho.getRange(jr, 2).clearContent(); joho.getRange(jr, 1,  1, 14).setBackground(null); joho.getRange(jr, 14).clearContent(); }
        if (ids[abId]) { joho.getRange(jr, 16).clearContent(); joho.getRange(jr, 15, 1, 15).setBackground(null); joho.getRange(jr, 29).clearContent(); }
      }
    }
  }

  ui.alert('✅ マッチング解除完了\n\n' + idList.join(', ') + ' を運行・集計表から削除しました。');
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
    set('トン数', cargoRow[7]);           // H: トン数(貨物要求) → 車両側で上書きされる場合も
    set('車種',   cargoRow[8]);           // I: 車種（貨物要求）→ 後で車両側が上書き
    set('積地',   cargoRow[9]);           // J: 積込地
    set('降地',   cargoRow[10]);          // K: 降ろし地
    set('売上',   cargoRow[11]);          // L: 金額(売上)
    set('備考',   cargoRow[12]);          // M: 備考(貨物)
  }

  // ── 車両情報からのセット ───────────────────────────────────────────
  // 情報シート29列構成での車両セクションのインデックス（0ベース）:
  // N=13:貨物登録ID O=14:チェック(車両) P=15:進捗(車両) Q=16:会社名 R=17:TEL S=18:FAX
  // T=19:日付(車両) U=20:看板名 V=21:トン数(車両) W=22:車種(車両) X=23:車番 Y=24:乗務員名 Z=25:携帯 AA=26:金額(支払) AB=27:備考
  if (vehRow) {
    set('会社名',   vehRow[16]); // Q: 会社名(車両) → 協力会社名
    set('看板名',   vehRow[20]); // U: 看板名
    set('トン数',   vehRow[21]); // V: トン数(車両) ← 車両実績を優先（貨物側を上書き）
    set('車種',     vehRow[22]); // W: 車種(車両)   ← 貨物側を上書き
    set('車番',     vehRow[23]); // X: 車番
    set('乗務員名', vehRow[24]); // Y: 乗務員名
    set('携帯番号', vehRow[25]); // Z: 携帯番号
    set('支払い',   vehRow[26]); // AA: 金額(支払)
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
//  14-2c: 発注書・指示書ダイアログ表示（showHatchuDocDialog）  【大C / 中14 / 小14-2c】
//  メニュー「発注書・指示書」: 選択行のデータで発注書・指示書プレビューダイアログを開く
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

// ================================================================
//  14-2d: 車番連絡ダイアログ表示（showShabanDocDialog）  【大C / 中14 / 小14-2d】
//  メニュー「車番連絡」: 選択行のデータで車番連絡プレビューダイアログを開く
// ================================================================
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
//  14-2e: アクティブ行のデータ取得（getDocumentData_）  【大B / 中14 / 小14-2e】
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
    ssId:        ss.getId(),
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


// ================================================================
//  16-1: 受領書の耳生成ダイアログ（showUketorishoDialog）  【大C / 中16 / 小16-1】
//  ドロップダウン形式で絞り込み条件を選んでシート生成する
// ================================================================
function showUketorishoDialog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('運行');
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('運行シートにデータがありません。');
    return;
  }

  // 運行シートから選択肢を収集
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  var clients = {}, companies = {}, dates = {}, cars = {}, drivers = {}, picks = {}, drops = {};
  data.forEach(function(r) {
    var id = String(r[0]||'').trim();
    if (!id) return;
    var pick = String(r[11]||'').trim();
    if (!pick || pick.indexOf('有休') !== -1 || pick.indexOf('休み') !== -1) return;
    if (r[10]) clients[String(r[10]).trim()] = true;
    if (r[2])  companies[String(r[2]).trim()] = true;
    if (r[9] instanceof Date) {
      dates[Utilities.formatDate(r[9], 'Asia/Tokyo', 'yyyy/MM/dd')] = true;
    }
    if (r[5]) cars[String(r[5]).trim()] = true;
    if (r[6]) drivers[String(r[6]).trim()] = true;
    if (r[11]) picks[String(r[11]).trim()] = true;
    if (r[12]) drops[String(r[12]).trim()] = true;
  });

  function mkOpts(obj) {
    return '<option value="">(全て)</option>'
      + Object.keys(obj).sort().map(function(v) {
          return '<option value="' + v.replace(/"/g,'&quot;') + '">' + v + '</option>';
        }).join('');
  }

  var html = '<html><body style="font-family:sans-serif;padding:16px;font-size:13px;">'
    + '<table style="border-collapse:collapse;width:100%">'
    + '<tr><td style="padding:6px">荷主名</td><td><select id="cl" style="width:200px">' + mkOpts(clients)   + '</select></td></tr>'
    + '<tr><td style="padding:6px">会社名</td><td><select id="co" style="width:200px">' + mkOpts(companies) + '</select></td></tr>'
    + '<tr><td style="padding:6px">日付</td><td><select id="dt" style="width:200px">'   + mkOpts(dates)     + '</select></td></tr>'
    + '<tr><td style="padding:6px">車番</td><td><select id="ca" style="width:200px">'   + mkOpts(cars)      + '</select></td></tr>'
    + '<tr><td style="padding:6px">乗務員名</td><td><select id="dr" style="width:200px">' + mkOpts(drivers) + '</select></td></tr>'
    + '<tr><td style="padding:6px">積地</td><td><select id="pk" style="width:200px">'   + mkOpts(picks)     + '</select></td></tr>'
    + '<tr><td style="padding:6px">降地</td><td><select id="dp" style="width:200px">'   + mkOpts(drops)     + '</select></td></tr>'
    + '</table>'
    + '<br><button onclick="g()" style="background:#006064;color:#fff;padding:8px 24px;border:none;border-radius:6px;cursor:pointer;font-size:13px">生成</button>'
    + '<br><br><button id="pb" style="display:none;background:#1565c0;color:#fff;padding:10px 28px;border:none;border-radius:6px;cursor:pointer;font-size:14px;font-weight:bold">🖨 印刷用PDFを開く</button>'
    + '<br><span id="m" style="margin-top:8px;display:inline-block;color:#888;font-size:12px"></span>'
    + '<script>function g(){'
    + '  document.getElementById("m").innerText="生成中...";'
    + '  var f={'
    + '    client:document.getElementById("cl").value,'
    + '    company:document.getElementById("co").value,'
    + '    date:document.getElementById("dt").value,'
    + '    car:document.getElementById("ca").value,'
    + '    driver:document.getElementById("dr").value,'
    + '    pick:document.getElementById("pk").value,'
    + '    drop:document.getElementById("dp").value'
    + '  };'
    + '  google.script.run'
    + '    .withSuccessHandler(function(r){'
    + '      if(r && r.pdfUrl){'
    + '        document.getElementById("m").innerText=r.msg;'
    + '        var b=document.getElementById("pb");b.style.display="inline-block";'
    + '        b.onclick=function(){window.open(r.pdfUrl,"_blank");};'
    + '      } else { document.getElementById("m").innerText=(r&&r.msg)||r||"完了"; }'
    + '    })'
    + '    .withFailureHandler(function(e){document.getElementById("m").innerText="エラー: "+(e.message||e);})'
    + '    .generateUketorishoSheet(f);'
    + '}'
    + '<\/script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(420).setHeight(350),
    '🗒 受領書の耳生成'
  );
}


// ================================================================
//  16-2: 受領書耳シート生成（generateUketorishoSheet）  【大A / 中16 / 小16-2】
//  絞り込み条件に合う行を運行シートから抽出し、A4横2列×5行=10件/ページで
//  「受領書_耳」シートに印刷用レイアウトを生成する
// ================================================================
function generateUketorishoSheet(filters) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var unkou = ss.getSheetByName('運行');
  if (!unkou || unkou.getLastRow() < 2) return '運行シートにデータがありません。';

  var lastCol = unkou.getLastColumn();
  var all = unkou.getRange(2, 1, unkou.getLastRow() - 1, Math.max(lastCol, 13)).getValues();

  // 自社設定から自社名取得
  var selfName = '';
  var selfSheet = ss.getSheetByName('自社設定');
  if (selfSheet && selfSheet.getLastRow() >= 1) {
    selfSheet.getDataRange().getValues().forEach(function(r) {
      if (String(r[0]||'').trim() === '会社名') selfName = String(r[1]||'').trim();
    });
  }

  // 絞り込み（空文字 = 全対象）
  function matchFilter(filterVal, val) {
    if (!filterVal) return true;
    return String(val||'').trim() === filterVal;
  }

  var rows = [];
  var seen = {};
  all.forEach(function(r) {
    var id = String(r[0]||'').trim();
    if (!id || seen[id]) return;
    var dateStr = r[9] instanceof Date ? Utilities.formatDate(r[9], 'Asia/Tokyo', 'yyyy/MM/dd') : String(r[9]||'');
    var client  = String(r[10]||'').trim();
    var company = String(r[2] ||'').trim();
    var car     = String(r[5] ||'').trim();
    var driver  = String(r[6] ||'').trim();
    var pick    = String(r[11]||'').trim();
    var drop    = String(r[12]||'').trim();
    if (!pick || pick.indexOf('有休') !== -1 || pick.indexOf('休み') !== -1) return;
    if (!matchFilter(filters.client,   client))  return;
    if (!matchFilter(filters.company,  company)) return;
    if (!matchFilter(filters.date,     dateStr)) return;
    if (!matchFilter(filters.car,      car))     return;
    if (!matchFilter(filters.driver,   driver))  return;
    if (!matchFilter(filters.pick,     pick))    return;
    if (!matchFilter(filters.drop,     drop))    return;
    seen[id] = true;
    var WD = ['日','月','火','水','木','金','土'];
    var wdStr = r[9] instanceof Date ? WD[r[9].getDay()] : '';
    var signboard = String(r[8]||'').trim() || company;
    rows.push({ id: id, client: client, company: company, signboard: signboard,
      dateStr: dateStr, wdStr: wdStr, car: car, driver: driver, pick: pick, drop: drop });
  });

  // 日付昇順ソート
  rows.sort(function(a, b) { return a.dateStr.localeCompare(b.dateStr); });

  if (rows.length === 0) return '対象データが見つかりません。絞り込み条件を確認してください。';

  // シート準備
  var shName = '受領書_耳';
  var sh = ss.getSheetByName(shName);
  if (sh) { sh.clear(); sh.clearFormats(); } else { sh = ss.insertSheet(shName); }

  // レイアウト定数: A4縦 3列×7行 = 21件/ページ・余白5mm
  // fitw=true, 余白0.20inch(5mm)。プリンタ印刷可能領域に確実に収める。
  // 15列 × 53px = 795px → scale=0.949 → 使用可高さ≈1143px → 7段×163px=1141px ✓
  var COLS_PER_ROW  = 3;
  var ROWS_PER_PAGE = 7;
  var PER_PAGE      = COLS_PER_ROW * ROWS_PER_PAGE; // 21

  // 15列（3耳×5列）の列幅を53pxに統一（fitw=false時にA4幅793pxをほぼカバー）
  for (var ci = 1; ci <= 15; ci++) { sh.setColumnWidth(ci, 53); }

  var CELL_ROWS = 8;  // 1耳あたりのスプレッドシート行数
  var curRow = 1;

  for (var idx = 0; idx < rows.length; idx += PER_PAGE) {
    var pageRows = rows.slice(idx, idx + PER_PAGE);

    // ページ内を3列×8行に配置
    for (var pi = 0; pi < pageRows.length; pi++) {
      var item    = pageRows[pi];
      var colSlot = pi % COLS_PER_ROW;               // 0/1/2
      var rowSlot = Math.floor(pi / COLS_PER_ROW);
      var c1      = colSlot * 5 + 1;                 // 1/6/11
      var r1      = curRow + rowSlot * CELL_ROWS;

      var BD = '#006064';
      var BS = SpreadsheetApp.BorderStyle.SOLID;

      // 行1: 荷主名 御中  [1耳=23+20×6+16=159px×7=1113px ≈ 余白5mm込み使用可高1143px ✓]
      sh.getRange(r1, c1, 1, 5).merge()
        .setValue((item.client || '　') + '　御中')
        .setFontSize(12).setFontWeight('bold')
        .setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBackground('#e0f7fa').setFontColor('#006064')
        .setBorder(true, true, false, true, false, false, BD, BS);
      sh.setRowHeight(r1, 23);

      // 行2: 積込日
      var r2 = r1 + 1;
      sh.getRange(r2, c1, 1, 5).merge()
        .setValue('積込日　' + item.dateStr + '（' + item.wdStr + '）')
        .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(false, true, false, true, false, false, BD, BS);
      sh.setRowHeight(r2, 20);

      // 行3: 会社名／看板名
      var r3 = r1 + 2;
      sh.getRange(r3, c1, 1, 2).merge().setValue('会社名')
        .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#f0f0f0')
        .setBorder(false, true, false, false, false, false, BD, BS);
      sh.getRange(r3, c1+2, 1, 3).merge().setValue(item.signboard || item.company)
        .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(false, false, false, true, false, false, BD, BS);
      sh.setRowHeight(r3, 20);

      // 行4: 車番
      var r4 = r1 + 3;
      sh.getRange(r4, c1, 1, 2).merge().setValue('車番')
        .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#f0f0f0')
        .setBorder(false, true, false, false, false, false, BD, BS);
      sh.getRange(r4, c1+2, 1, 3).merge().setValue(item.car)
        .setFontSize(11).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(false, false, false, true, false, false, BD, BS);
      sh.setRowHeight(r4, 20);

      // 行5: 乗務員
      var r5 = r1 + 4;
      sh.getRange(r5, c1, 1, 2).merge().setValue('乗務員')
        .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#f0f0f0')
        .setBorder(false, true, false, false, false, false, BD, BS);
      sh.getRange(r5, c1+2, 1, 3).merge().setValue(item.driver)
        .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setBorder(false, false, false, true, false, false, BD, BS);
      sh.setRowHeight(r5, 20);

      // 行6: 積地
      var r6 = r1 + 5;
      sh.getRange(r6, c1, 1, 2).merge().setValue('積地')
        .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#e8f5e9')
        .setBorder(false, true, false, false, false, false, BD, BS);
      sh.getRange(r6, c1+2, 1, 3).merge().setValue(item.pick)
        .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true)
        .setBorder(false, false, false, true, false, false, BD, BS);
      sh.setRowHeight(r6, 20);

      // 行7: 降地
      var r7 = r1 + 6;
      sh.getRange(r7, c1, 1, 2).merge().setValue('降地')
        .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground('#fce4ec')
        .setBorder(false, true, false, false, false, false, BD, BS);
      sh.getRange(r7, c1+2, 1, 3).merge().setValue(item.drop)
        .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true)
        .setBorder(false, false, false, true, false, false, BD, BS);
      sh.setRowHeight(r7, 20);

      // 行8: 自社名（右下）
      var r8 = r1 + 7;
      sh.getRange(r8, c1, 1, 5).merge().setValue(selfName)
        .setFontSize(8).setHorizontalAlignment('right').setVerticalAlignment('middle').setFontColor('#777')
        .setBorder(true, true, true, true, false, false, BD, BS);
      sh.setRowHeight(r8, 16);
    }

    curRow += ROWS_PER_PAGE * CELL_ROWS; // ページ間余白なし
  }

  // 受領書列に「済」を記録
  var uHdrs = unkou.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h||'').trim(); });
  var ukCol = uHdrs.indexOf('受領書');
  if (ukCol >= 0) {
    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd HH:mm');
    var allIds = unkou.getRange(2, 1, unkou.getLastRow() - 1, 1).getValues();
    for (var ri = 0; ri < allIds.length; ri++) {
      var rid = String(allIds[ri][0]||'').trim();
      if (seen[rid]) {
        unkou.getRange(ri + 2, ukCol + 1).setValue(now);
      }
    }
  }

  sh.setHiddenGridlines(false);
  ss.setActiveSheet(sh);

  // 余白なし・A4縦・幅フィット・グリッド線なし のPDF出力URL
  var totalRows = curRow - 1;
  var pdfUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export'
    + '?format=pdf'
    + '&gid=' + sh.getSheetId()
    + '&size=A4'
    + '&portrait=false'
    + '&fitw=true'
    + '&gridlines=false'
    + '&top_margin=0.20'
    + '&bottom_margin=0.20'
    + '&left_margin=0.20'
    + '&right_margin=0.20'
    + '&sheetnames=false'
    + '&printtitle=false'
    + '&pagenumbers=false'
    + '&attachment=false';

  return { msg: rows.length + '件の耳を生成しました。「印刷用PDFを開く」ボタンをクリックして印刷してください。', pdfUrl: pdfUrl };
}


// ================================================================
// ■ グループ17：PL（損益計算書）生成エンジン
// ================================================================
//   17-1  : showPlDialog()           - PLフィルタモーダル表示
//   17-2  : getPlFilterOptions()     - フィルタ選択肢取得
//   17-3  : generatePl(filters)      - PL生成メイン
//   17-3a : buildPlBreakdown_()      - 表示単位別内訳生成
//   17-3b : buildPlSheet_()          - PLシート整形出力
//   17-4  : getFixedCosts_()         - 固定費マスタ取得
//   17-5  : apportionFixedCosts_()   - 固定費按分計算
//   17-6  : exportPlJournalCsv()     - 弥生会計互換仕訳CSV出力
//   17-6a : buildJournalEntry_()     - 仕訳行生成補助
//   17-7  : initFixedCostMaster()    - PL設定シート初期化
// ================================================================


// ================================================================
//  17-1: PLフィルタモーダル表示（showPlDialog）  【大C / 中17 / 小17-1】
// ================================================================
function showPlDialog() {
  var tmpl = HtmlService.createTemplateFromFile('plDialog');
  var html = tmpl.evaluate().setWidth(680).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 PL（損益計算書）作成');
}


// ================================================================
//  17-2: PLフィルタ選択肢取得（getPlFilterOptions）  【大A / 中17 / 小17-2】
//  集計表・自車専属マスタから荷主・会社・車番・乗務員の選択肢を返す
// ================================================================
function getPlFilterOptions() {
  var ss         = SpreadsheetApp.getActiveSpreadsheet();
  var sumSheet   = ss.getSheetByName('集計表');
  var masterSheet= ss.getSheetByName('自車専属マスタ');

  var clients = [], companies = [], cars = [], drivers = [];

  if (sumSheet && sumSheet.getLastRow() >= 2) {
    var sumData = sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, 11).getValues();
    for (var i = 0; i < sumData.length; i++) {
      var co = String(sumData[i][2]  || '').trim(); // C列: 会社名
      var cl = String(sumData[i][10] || '').trim(); // K列: 荷主
      if (co && companies.indexOf(co) === -1) companies.push(co);
      if (cl && clients.indexOf(cl)   === -1) clients.push(cl);
    }
  }

  if (masterSheet && masterSheet.getLastRow() >= 2) {
    var mData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 9).getValues();
    for (var j = 0; j < mData.length; j++) {
      var car    = String(mData[j][7] || '').trim(); // H列: 車番
      var driver = String(mData[j][8] || '').trim(); // I列: 乗務員名
      if (car    && cars.indexOf(car)       === -1) cars.push(car);
      if (driver && drivers.indexOf(driver) === -1) drivers.push(driver);
    }
  }

  var today    = new Date();
  var firstDay = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth(), 1),   'Asia/Tokyo', 'yyyy-MM-dd');
  var lastDay  = Utilities.formatDate(new Date(today.getFullYear(), today.getMonth() + 1, 0), 'Asia/Tokyo', 'yyyy-MM-dd');

  return {
    clients:     clients.sort(),
    companies:   companies.sort(),
    cars:        cars.sort(),
    drivers:     drivers.sort(),
    defaultFrom: firstDay,
    defaultTo:   lastDay
  };
}


// ================================================================
//  17-3: PL生成メイン（generatePl）  【大A / 中17 / 小17-3】
//  フィルタ条件で集計表を絞り込み、固定費按分を加えてPLシートを出力する
//  戻り値: { ok, msg, sheetName }
// ================================================================
function generatePl(filters) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sumSheet  = ss.getSheetByName('集計表');
  if (!sumSheet || sumSheet.getLastRow() < 2) {
    return { ok: false, msg: '集計表にデータがありません。' };
  }

  var fromDate         = filters.from ? new Date(filters.from + 'T00:00:00') : null;
  var toDate           = filters.to   ? new Date(filters.to   + 'T23:59:59') : null;
  var unit             = filters.unit || 'monthly';
  var filterClients    = filters.clients   || [];
  var filterCompanies  = filters.companies || [];
  var filterCars       = filters.cars      || [];
  var filterDrivers    = filters.drivers   || [];

  var colCount = Math.max(sumSheet.getLastColumn(), 37);
  var sumData  = sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, colCount).getValues();

  var filteredRows = [];
  for (var i = 0; i < sumData.length; i++) {
    var r       = sumData[i];
    var rowDate = r[9]; // J列: 日付
    if (!(rowDate instanceof Date) || isNaN(rowDate.getTime())) continue;
    if (fromDate && rowDate < fromDate) continue;
    if (toDate   && rowDate > toDate)   continue;

    var client  = String(r[10] || '').trim(); // K列: 荷主
    var company = String(r[2]  || '').trim(); // C列: 会社名
    var car     = String(r[5]  || '').trim(); // F列: 車番
    var driver  = String(r[6]  || '').trim(); // G列: 乗務員名

    if (filterClients.length   > 0 && filterClients.indexOf(client)     === -1) continue;
    if (filterCompanies.length > 0 && filterCompanies.indexOf(company)  === -1) continue;
    if (filterCars.length      > 0 && filterCars.indexOf(car)           === -1) continue;
    if (filterDrivers.length   > 0 && filterDrivers.indexOf(driver)     === -1) continue;

    var pick     = String(r[11] || '').trim();
    var isHoliday= pick.indexOf('有休') !== -1 || pick.indexOf('休み') !== -1;

    filteredRows.push({
      date:      rowDate,
      company:   company,
      car:       car,
      driver:    driver,
      client:    client,
      sales:     Number(r[18]) || 0, // T: 売上
      tollReq:   Number(r[19]) || 0, // U: 請求高速
      tollReal:  Number(r[20]) || 0, // V: 実費高速
      fuel:      Number(r[25]) || 0, // Z: 燃料代
      payment:   Number(r[26]) || 0, // AA: 支払い
      expense:   Number(r[27]) || 0, // AB: 経費合計
      profit:    Number(r[28]) || 0, // AC: 利益
      yukyu:     Number(r[33]) || 0, // AH: 有休手当
      other:     Number(r[34]) || 0, // AI: その他手当
      isHoliday: isHoliday
    });
  }

  if (filteredRows.length === 0) {
    return { ok: false, msg: '条件に合致するデータが見つかりません。' };
  }

  // 稼働統計（固定費按分用）
  var activeDateSet = {}, activeCarSet = {};
  for (var ri = 0; ri < filteredRows.length; ri++) {
    var row = filteredRows[ri];
    if (!row.isHoliday && row.sales > 0) {
      activeDateSet[Utilities.formatDate(row.date, 'Asia/Tokyo', 'yyyy/MM/dd')] = true;
      activeCarSet[row.car] = true;
    }
  }
  var activeDays     = Object.keys(activeDateSet).length;
  var activeVehicles = Object.keys(activeCarSet).length;
  var totalDays      = 30;
  if (fromDate && toDate) {
    totalDays = Math.round((toDate.getTime() - fromDate.getTime()) / 86400000) + 1;
  }

  // 売上・変動費を集計
  var totalSales = 0, totalTollReq = 0, totalTollReal = 0;
  var totalFuel  = 0, totalPayment = 0, totalExpense  = 0;
  var totalYukyu = 0, totalOther  = 0;
  for (var fi = 0; fi < filteredRows.length; fi++) {
    var fr      = filteredRows[fi];
    totalSales   += fr.sales;
    totalTollReq += fr.tollReq;
    totalTollReal+= fr.tollReal;
    totalFuel    += fr.fuel;
    totalPayment += fr.payment;
    totalExpense += fr.expense;
    totalYukyu   += fr.yukyu;
    totalOther   += fr.other;
  }

  // 固定費取得・按分
  var fixedCosts       = getFixedCosts_(ss);
  var apportioned      = apportionFixedCosts_(fixedCosts, activeDays, activeVehicles, totalDays);
  var totalFixed       = 0;
  for (var fc = 0; fc < apportioned.length; fc++) totalFixed += apportioned[fc].amount;

  var totalVariable    = totalPayment + totalTollReal + totalFuel + totalExpense + totalYukyu + totalOther;
  var grossProfit      = (totalSales + totalTollReq) - totalVariable;
  var operatingProfit  = grossProfit - totalFixed;

  var breakdown = buildPlBreakdown_(filteredRows, unit);

  var sheetName = buildPlSheet_(ss, {
    from: filters.from, to: filters.to, unit: unit,
    totalSales: totalSales, totalTollReq: totalTollReq, totalTollReal: totalTollReal,
    totalFuel: totalFuel, totalPayment: totalPayment, totalExpense: totalExpense,
    totalYukyu: totalYukyu, totalOther: totalOther, totalVariable: totalVariable,
    grossProfit: grossProfit, fixedCosts: apportioned, totalFixed: totalFixed,
    operatingProfit: operatingProfit,
    activeDays: activeDays, activeVehicles: activeVehicles, totalRows: filteredRows.length,
    breakdown: breakdown
  });

  return { ok: true, msg: filteredRows.length + '件のデータからPLを生成しました。', sheetName: sheetName };
}


// ================================================================
//  17-3a: 表示単位別内訳生成（buildPlBreakdown_）  【大B / 中17 / 小17-3a】
// ================================================================
function buildPlBreakdown_(rows, unit) {
  var groups = {};
  for (var i = 0; i < rows.length; i++) {
    var r   = rows[i];
    var key;
    if (unit === 'daily') {
      key = Utilities.formatDate(r.date, 'Asia/Tokyo', 'yyyy/MM/dd');
    } else if (unit === 'weekly') {
      var day  = r.date.getDay();
      var diff = (day === 0) ? -6 : 1 - day;
      var mon  = new Date(r.date.getFullYear(), r.date.getMonth(), r.date.getDate() + diff);
      key = Utilities.formatDate(mon, 'Asia/Tokyo', 'yyyy/MM/dd') + '〜';
    } else {
      key = Utilities.formatDate(r.date, 'Asia/Tokyo', 'yyyy年MM月');
    }
    if (!groups[key]) groups[key] = { key: key, sales: 0, tollReal: 0, fuel: 0, payment: 0, expense: 0, profit: 0 };
    groups[key].sales   += r.sales;
    groups[key].tollReal+= r.tollReal;
    groups[key].fuel    += r.fuel;
    groups[key].payment += r.payment;
    groups[key].expense += r.expense;
    groups[key].profit  += r.profit;
  }
  var result = [];
  Object.keys(groups).sort().forEach(function(k) { result.push(groups[k]); });
  return result;
}


// ================================================================
//  17-3b: PLシート整形出力（buildPlSheet_）  【大B / 中17 / 小17-3b】
//  黒背景・ライムグリーンアクセントのPLシートを生成してシート名を返す
// ================================================================
function buildPlSheet_(ss, d) {
  var sheetName = 'PL';
  var existing  = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  var sh = ss.insertSheet(sheetName);

  sh.setColumnWidth(1, 24);   // A: インデント
  sh.setColumnWidth(2, 230);  // B: 項目名
  sh.setColumnWidth(3, 140);  // C: 金額
  sh.setColumnWidth(4, 24);   // D: 余白
  sh.setColumnWidth(5, 120);  // E: 内訳期間
  sh.setColumnWidth(6, 100);  // F: 内訳売上
  sh.setColumnWidth(7, 100);  // G: 内訳支払
  sh.setColumnWidth(8, 100);  // H: 内訳燃料
  sh.setColumnWidth(9, 110);  // I: 内訳利益

  var C_BG      = '#0a0a0a';
  var C_CARD    = '#111827';
  var C_CARD2   = '#0d1117';
  var C_ACCENT  = '#00ff88';
  var C_DIM     = '#888888';
  var C_TEXT    = '#e0e0e0';
  var C_RED     = '#ff4444';
  var C_BORDER  = '#1e3a2a';
  var BS        = SpreadsheetApp.BorderStyle.SOLID;
  var R         = 1;

  function bgRow(r1, numR, numC, bg) {
    sh.getRange(r1, 1, numR, numC).setBackground(bg);
  }

  function titleRow(txt) {
    sh.getRange(R, 1, 1, 9).merge()
      .setValue(txt).setBackground('#0a1628').setFontColor(C_ACCENT)
      .setFontWeight('bold').setFontSize(16).setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.setRowHeight(R, 44); R++;
  }

  function infoRow(txt) {
    sh.getRange(R, 1, 1, 9).merge()
      .setValue(txt).setBackground('#0d1117').setFontColor(C_DIM)
      .setFontSize(10).setHorizontalAlignment('center');
    sh.setRowHeight(R, 20); R++;
  }

  function sectionHead(txt) {
    sh.getRange(R, 1, 1, 9).merge()
      .setValue('  ' + txt).setBackground('#0d2137').setFontColor(C_ACCENT)
      .setFontWeight('bold').setFontSize(11);
    sh.setRowHeight(R, 26); R++;
  }

  function itemRow(label, amount, indent) {
    var bg = indent ? C_CARD : C_CARD2;
    sh.getRange(R, 2).setValue(indent ? '    ' + label : '  ' + label)
      .setBackground(bg).setFontColor(C_TEXT).setFontSize(12);
    sh.getRange(R, 3).setValue(amount)
      .setNumberFormat('¥#,##0;[RED]-¥#,##0')
      .setBackground(bg).setFontColor(amount < 0 ? C_RED : C_TEXT)
      .setFontSize(12).setHorizontalAlignment('right');
    sh.setRowHeight(R, 22); R++;
  }

  function totalRow(label, amount) {
    sh.getRange(R, 2).setValue('  ' + label)
      .setBackground('#0a2a1a').setFontColor(C_ACCENT).setFontWeight('bold').setFontSize(12);
    sh.getRange(R, 3).setValue(amount)
      .setNumberFormat('¥#,##0;[RED]-¥#,##0')
      .setBackground('#0a2a1a').setFontColor(amount < 0 ? C_RED : C_ACCENT)
      .setFontWeight('bold').setFontSize(12).setHorizontalAlignment('right');
    sh.getRange(R, 2, 1, 2).setBorder(null, null, true, null, null, null, C_BORDER, BS);
    sh.setRowHeight(R, 24); R++;
  }

  function profitRow(label, amount, big) {
    var pBg = big
      ? (amount >= 0 ? '#0b4a25' : '#4a0b0b')
      : (amount >= 0 ? '#0a2a18' : '#2a0a0a');
    var pFg = amount >= 0 ? C_ACCENT : C_RED;
    var fz  = big ? 15 : 13;
    var sign = amount >= 0 ? '¥' : '-¥';
    sh.getRange(R, 1, 1, 9).merge()
      .setValue(label + '　　' + sign + Math.abs(amount).toLocaleString())
      .setBackground(pBg).setFontColor(pFg).setFontWeight('bold')
      .setFontSize(fz).setHorizontalAlignment('center');
    // 上下に横線を入れて区切りを明確にする（★の代替）
    sh.getRange(R, 1, 1, 9).setBorder(
      true, null, true, null, null, null,
      amount >= 0 ? '#1a5f30' : '#5f1a1a',
      SpreadsheetApp.BorderStyle.MEDIUM
    );
    sh.setRowHeight(R, big ? 42 : 36); R++;
  }

  function blank(h) { sh.setRowHeight(R, h || 8); R++; }

  // ── タイトル ──────────────────────────────────────────────────
  titleRow('📊  PL（損益計算書）');
  infoRow('期間: ' + (d.from || '—') + '  〜  ' + (d.to || '—') +
          '    稼働 ' + d.activeDays + '日 / ' + d.activeVehicles + '台    対象 ' + d.totalRows + '件');
  blank(10);

  // ── 収益 ──────────────────────────────────────────────────────
  sectionHead('▶ 収益');
  itemRow('運賃収入（売上）',       d.totalSales,    true);
  itemRow('高速代収入（請求分）',   d.totalTollReq,  true);
  totalRow('売上合計',              d.totalSales + d.totalTollReq);
  blank();

  // ── 変動費 ────────────────────────────────────────────────────
  sectionHead('▶ 変動費（売上原価）');
  itemRow('支払運賃（乗務員）',     d.totalPayment,  true);
  itemRow('高速代（実費）',         d.totalTollReal, true);
  itemRow('燃料費',                 d.totalFuel,     true);
  if (d.totalExpense > 0) itemRow('月次経費按分',    d.totalExpense, true);
  if (d.totalYukyu   > 0) itemRow('有休手当',        d.totalYukyu,   true);
  if (d.totalOther   > 0) itemRow('その他手当',      d.totalOther,   true);
  totalRow('変動費合計',            d.totalVariable);
  blank();

  // ── 売上総利益 ────────────────────────────────────────────────
  profitRow('売上総利益（粗利）', d.grossProfit, false);
  blank();

  // ── 固定費 ────────────────────────────────────────────────────
  sectionHead('▶ 固定費（按分後）');
  if (d.fixedCosts && d.fixedCosts.length > 0) {
    for (var fc = 0; fc < d.fixedCosts.length; fc++) {
      var fci = d.fixedCosts[fc];
      itemRow(fci.name + '  (' + fci.method + ')', fci.amount, true);
    }
  } else {
    itemRow('（PL設定シートで固定費を登録してください）', 0, true);
  }
  totalRow('固定費合計', d.totalFixed);
  blank();

  // ── 営業利益 ──────────────────────────────────────────────────
  profitRow('営業利益', d.operatingProfit, true);
  blank(16);

  // ── 内訳テーブル ──────────────────────────────────────────────
  if (d.breakdown && d.breakdown.length > 0) {
    var unitLbl = d.unit === 'daily' ? '日次' : d.unit === 'weekly' ? '週次' : '月次';
    sectionHead('▶ ' + unitLbl + '内訳');
    var bHdrs = ['期間', '売上', '支払運賃', '燃料費', '利益'];
    var bCols  = [5, 6, 7, 8, 9];
    for (var bhi = 0; bhi < bHdrs.length; bhi++) {
      sh.getRange(R, bCols[bhi]).setValue(bHdrs[bhi])
        .setBackground('#1a2a3a').setFontColor(C_ACCENT).setFontWeight('bold').setFontSize(10)
        .setHorizontalAlignment('center');
    }
    sh.setRowHeight(R, 22); R++;
    for (var bi = 0; bi < d.breakdown.length; bi++) {
      var bd  = d.breakdown[bi];
      var bBg = bi % 2 === 0 ? C_CARD : C_CARD2;
      sh.getRange(R, 5).setValue(String(bd.key)).setNumberFormat('@').setBackground(bBg).setFontColor(C_TEXT).setFontSize(10);
      [[6, bd.sales],[7, bd.payment + bd.expense],[8, bd.fuel],[9, bd.profit]].forEach(function(kv) {
        sh.getRange(R, kv[0]).setValue(kv[1])
          .setNumberFormat('#,##0').setBackground(bBg)
          .setFontColor(kv[0] === 9 ? (kv[1] >= 0 ? '#00cc66' : C_RED) : C_TEXT)
          .setFontSize(10).setHorizontalAlignment('right');
      });
      sh.setRowHeight(R, 20); R++;
    }
  }

  // シート全体に黒背景を下敷き
  sh.getRange(1, 1, R - 1, 9).setBackground(C_BG);
  // 各行の個別背景設定が上書きされないよう、再度全体に黒背景（Googleは後勝ちなので問題なし）

  sh.setHiddenGridlines(true);
  ss.setActiveSheet(sh);
  return sheetName;
}


// ================================================================
//  17-4: 固定費データ取得（getFixedCosts_）  【大B / 中17 / 小17-4】
//  「PL設定」シートから固定費マスタを読み込む
//  PL含入フラグ=FALSE の項目（社長給与等）は除外する
// ================================================================
function getFixedCosts_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('PL設定');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data   = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  var result = [];
  for (var i = 0; i < data.length; i++) {
    var name    = String(data[i][0] || '').trim(); // A: 費目名
    var monthly = Number(data[i][1]) || 0;          // B: 月額
    var method  = String(data[i][2] || '一定').trim(); // C: 按分方式
    var account = String(data[i][3] || '').trim();  // D: 勘定科目
    var plFlag  = data[i][4];                       // E: PL含入フラグ
    if (!name) continue;
    if (plFlag === false || String(plFlag).toUpperCase() === 'FALSE') continue; // OFF→除外
    result.push({ name: name, monthly: monthly, method: method, account: account });
  }
  return result;
}


// ================================================================
//  17-5: 固定費按分計算（apportionFixedCosts_）  【大B / 中17 / 小17-5】
//  按分方式:
//    「一定」    → 月額をそのまま使用
//    「車両台数」→ 月額 × 稼働車両台数 ÷ 全車両台数
//    「稼働日数」→ 月額 × 稼働日数 ÷ 暦日数
// ================================================================
function apportionFixedCosts_(fixedCosts, activeDays, activeVehicles, totalDays) {
  var ss            = SpreadsheetApp.getActiveSpreadsheet();
  var master        = ss.getSheetByName('自車専属マスタ');
  var totalVehicles = master && master.getLastRow() > 1 ? master.getLastRow() - 1 : 1;
  var result        = [];
  for (var i = 0; i < fixedCosts.length; i++) {
    var fc      = fixedCosts[i];
    var amount  = fc.monthly;
    if (fc.method === '車両台数') {
      amount = Math.round(fc.monthly * activeVehicles / Math.max(1, totalVehicles));
    } else if (fc.method === '稼働日数') {
      amount = Math.round(fc.monthly * activeDays / Math.max(1, totalDays));
    }
    result.push({ name: fc.name, monthly: fc.monthly, amount: amount, method: fc.method, account: fc.account });
  }
  return result;
}


// ================================================================
//  17-6: 仕訳CSV出力（exportPlJournalCsv）  【大A / 中17 / 小17-6】
//  直近生成のPLシートを弥生会計互換CSVとしてDriveに保存して返す
//  戻り値: { ok, msg, url }
// ================================================================
function exportPlJournalCsv() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var plSheets = ss.getSheets().filter(function(s) {
    return s.getName() === 'PL' || s.getName().indexOf('PL_') === 0;
  }).sort(function(a, b) { return b.getName().localeCompare(a.getName()); });

  if (plSheets.length === 0) {
    return { ok: false, msg: 'PLシートがありません。先に「PL作成」を実行してください。' };
  }
  var plSheet = plSheets[0];
  if (plSheet.getLastRow() < 2) {
    return { ok: false, msg: 'PLシートにデータがありません。' };
  }

  var data = plSheet.getRange(1, 1, plSheet.getLastRow(), 3).getValues();
  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');

  // 弥生会計CSVヘッダー（公式25列フォーマット準拠）
  var csvRows = [[
    '識別フラグ','伝票No.','決算','取引日付',
    '借方勘定科目','借方補助科目','借方部門','借方税区分','借方金額','借方税金額',
    '貸方勘定科目','貸方補助科目','貸方部門','貸方税区分','貸方金額','貸方税金額',
    '摘要','番号','期日','タイプ','生成元','仕訳メモ','付箋1','付箋2','調整'
  ]];

  var vNo = 1;
  for (var i = 0; i < data.length; i++) {
    var rawLabel = String(data[i][1] || '');
    var label    = rawLabel.trim();
    var amount   = Number(data[i][2]) || 0;
    if (!label || amount === 0) continue;
    // trim前のrawLabelでインデント判定（trimするとスペースが消えて常にfalseになるため）
    if (rawLabel.charAt(0) === ' ' || rawLabel.charAt(0) === '　') {
      var entry = buildJournalEntry_(label, amount, today, vNo);
      if (entry) { csvRows.push(entry); vNo++; }
    }
  }

  if (csvRows.length <= 1) {
    return { ok: false, msg: '仕訳データが生成できませんでした。PLシートをご確認ください。' };
  }

  var csvText = '﻿'; // BOM（Excel文字化け防止）
  for (var ri = 0; ri < csvRows.length; ri++) {
    csvText += csvRows[ri].map(function(cell) {
      var s = String(cell === null || cell === undefined ? '' : cell);
      return (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1)
        ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(',') + '\r\n';
  }

  // 仕訳表シートにも書き込む（確認・個別CSV出力用）
  var jSheet = ss.getSheetByName('仕訳表');
  if (jSheet) ss.deleteSheet(jSheet);
  jSheet = ss.insertSheet('仕訳表');
  jSheet.getRange(1, 1, csvRows.length, csvRows[0].length).setValues(csvRows);
  jSheet.getRange(1, 1, 1, csvRows[0].length)
    .setBackground('#1a2a3a').setFontColor('#00ff88').setFontWeight('bold').setFontSize(10);
  jSheet.setFrozenRows(1);

  var fileName = plSheet.getName() + '_仕訳.csv';
  // Base64エンコードしてクライアントに返す（ブラウザ直接DL用）
  var base64csv = Utilities.base64Encode(csvText, Utilities.Charset.UTF_8);

  return { ok: true, msg: (csvRows.length - 1) + '件の仕訳を生成しました。', base64csv: base64csv, fileName: fileName };
}


// ================================================================
//  17-6a-2: PL表+仕訳CSV ZIP出力（exportPlBundle）  【大C / 中17 / 小17-6a-2】
//  opts = { includeJournal: bool, includePl: bool }
//  選択した内容をBOM付きCSVにしてZIPで返す（plDialog.html の📦ZIPボタン用）
// ================================================================
function exportPlBundle(opts) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var plSheets = ss.getSheets().filter(function(s) {
    return s.getName() === 'PL' || s.getName().indexOf('PL_') === 0;
  }).sort(function(a, b) { return b.getName().localeCompare(a.getName()); });

  if (plSheets.length === 0) {
    return { ok: false, msg: 'PLシートがありません。先に「📈 PL生成」を実行してください。' };
  }
  var plSheet = plSheets[0];
  var plName  = plSheet.getName();
  var today   = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  var blobs   = [];
  var msgParts = [];

  // 仕訳CSV
  if (opts && opts.includeJournal) {
    if (plSheet.getLastRow() < 2) return { ok: false, msg: 'PLシートにデータがありません。' };
    var jFmt = (opts.journalFormat === 'mf') ? 'mf' : 'yayoi';
    var data = plSheet.getRange(1, 1, plSheet.getLastRow(), 3).getValues();
    var csvRows = jFmt === 'mf'
      ? [['取引No','取引日',
          '借方勘定科目','借方補助科目','借方部門','借方取引先','借方税区分','借方インボイス','借方金額(円)','借方税額',
          '貸方勘定科目','貸方補助科目','貸方部門','貸方取引先','貸方税区分','貸方インボイス','貸方金額(円)','貸方税額',
          '摘要','仕訳メモ','タグ','MF仕訳タイプ','決算整理仕訳','作成日時','作成者','最終更新日時','最終更新者']]
      : [['識別フラグ','伝票No.','決算','取引日付',
          '借方勘定科目','借方補助科目','借方部門','借方税区分','借方金額','借方税金額',
          '貸方勘定科目','貸方補助科目','貸方部門','貸方税区分','貸方金額','貸方税金額',
          '摘要','番号','期日','タイプ','生成元','仕訳メモ','付箋1','付箋2','調整']];
    var vNo = 1;
    for (var i = 0; i < data.length; i++) {
      var rawLabel = String(data[i][1] || '');
      var label    = rawLabel.trim();
      var amount   = Number(data[i][2]) || 0;
      if (!label || amount === 0) continue;
      if (rawLabel.charAt(0) === ' ' || rawLabel.charAt(0) === '　') {
        var entry = jFmt === 'mf'
          ? buildJournalEntryMF_(label, amount, today)
          : buildJournalEntry_(label, amount, today, vNo);
        if (entry) { csvRows.push(entry); vNo++; }
      }
    }
    if (csvRows.length <= 1) return { ok: false, msg: '仕訳データが生成できませんでした。PLシートをご確認ください。' };
    var jText = '﻿';
    for (var jr = 0; jr < csvRows.length; jr++) {
      jText += csvRows[jr].map(function(c) {
        var s = String(c === null || c === undefined ? '' : c);
        return (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1)
          ? '"' + s.replace(/"/g, '""') + '"' : s;
      }).join(',') + '\r\n';
    }
    // 仕訳表シートに書き込む
    var jSheet = ss.getSheetByName('仕訳表');
    if (jSheet) ss.deleteSheet(jSheet);
    jSheet = ss.insertSheet('仕訳表');
    jSheet.getRange(1, 1, csvRows.length, csvRows[0].length).setValues(csvRows);
    jSheet.getRange(1, 1, 1, csvRows[0].length)
      .setBackground('#1a2a3a').setFontColor('#00ff88').setFontWeight('bold').setFontSize(10);
    jSheet.setFrozenRows(1);
    var jSuffix = jFmt === 'mf' ? '_仕訳MF.csv' : '_仕訳弥生.csv';
    blobs.push(Utilities.newBlob(jText, 'text/csv', plName + jSuffix));
    msgParts.push((csvRows.length - 1) + '件の仕訳CSV');
  }

  // PL表CSV
  if (opts && opts.includePl) {
    if (plSheet.getLastRow() < 1) return { ok: false, msg: 'PLシートが空です。' };
    var plData = plSheet.getRange(1, 1, plSheet.getLastRow(), Math.max(plSheet.getLastColumn(), 1)).getValues();
    var tz = Session.getScriptTimeZone();
    var plText = '﻿';
    for (var pi = 0; pi < plData.length; pi++) {
      plText += plData[pi].map(function(cell) {
        var s;
        if (cell instanceof Date && !isNaN(cell.getTime())) {
          var yr = parseInt(Utilities.formatDate(cell, tz, 'yyyy'), 10);
          s = (yr <= 1900) ? Utilities.formatDate(cell, tz, 'H:mm') : Utilities.formatDate(cell, tz, 'yyyy/MM/dd');
        } else {
          s = String(cell === null || cell === undefined ? '' : cell);
        }
        return (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1)
          ? '"' + s.replace(/"/g, '""') + '"' : s;
      }).join(',') + '\r\n';
    }
    blobs.push(Utilities.newBlob(plText, 'text/csv', plName + '_PL表.csv'));
    msgParts.push('PL表CSV');
  }

  if (blobs.length === 0) return { ok: false, msg: '出力内容が選択されていません。' };

  var zipName = plName + '_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '.zip';
  var zipBlob = Utilities.zip(blobs, zipName);
  return {
    ok: true,
    base64: Utilities.base64Encode(zipBlob.getBytes()),
    fileName: zipName,
    msg: msgParts.join(' + ') + ' をZIPにしました。'
  };
}


// ================================================================
//  17-6a: 仕訳行生成補助  【大B / 中17 / 小17-6a】
//  resolveJournalMapping_ : 費目名→勘定科目と税区分キー(sale/purchase/exempt)を返す
//  buildJournalEntry_     : 弥生会計形式25列を返す（公式フォーマット準拠）
//  buildJournalEntryMF_   : マネーフォワード クラウド会計形式27列を返す（公式フォーマット準拠）
// ================================================================
function resolveJournalMapping_(label) {
  // dt/ct = 'sale'(課税売上) / 'purchase'(課税仕入) / 'exempt'(対象外)
  // 売上系: 借方(売掛金)=exempt、貸方(売上高)=sale
  // 費用系: 借方(費用科目)=purchase、貸方(未払金)=exempt
  // 不課税: 給与・保険・税金・会費など両側exempt
  var MAP = [
    { key: '運賃収入',      debit: '売掛金',         credit: '売上高',         dt: 'exempt',   ct: 'sale'    },
    { key: '売上',          debit: '売掛金',         credit: '売上高',         dt: 'exempt',   ct: 'sale'    },
    { key: '高速代収入',    debit: '売掛金',         credit: '売上高',         dt: 'exempt',   ct: 'sale'    },
    { key: '支払運賃',      debit: '支払運賃',       credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '乗務員',        debit: '支払運賃',       credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '高速代（実費）',debit: '高速道路料金',   credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '燃料費',        debit: '燃料費',         credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '月次経費',      debit: '諸経費',         credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '有休手当',      debit: '法定福利費',     credit: '未払費用',       dt: 'exempt',   ct: 'exempt'  },
    { key: 'その他手当',    debit: '給与手当',       credit: '未払費用',       dt: 'exempt',   ct: 'exempt'  },
    { key: '家賃',          debit: '地代家賃',       credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '駐車場',        debit: '地代家賃',       credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '電気',          debit: '水道光熱費',     credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '水道',          debit: '水道光熱費',     credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '通信費',        debit: '通信費',         credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '保険',          debit: '損害保険料',     credit: '未払金',         dt: 'exempt',   ct: 'exempt'  },
    { key: 'リース',        debit: 'リース料',       credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '減価償却',      debit: '減価償却費',     credit: '減価償却累計額', dt: 'exempt',   ct: 'exempt'  },
    { key: '修繕',          debit: '修繕費',         credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '消耗品',        debit: '消耗品費',       credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '重量税',        debit: '租税公課',       credit: '未払金',         dt: 'exempt',   ct: 'exempt'  },
    { key: 'タイヤ',        debit: '消耗品費',       credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '洗車',          debit: '消耗品費',       credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '制服',          debit: '消耗品費',       credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '税理士',        debit: '支払手数料',     credit: '未払金',         dt: 'purchase', ct: 'exempt'  },
    { key: '安全協会',      debit: '諸会費',         credit: '未払金',         dt: 'exempt',   ct: 'exempt'  }
  ];
  for (var m = 0; m < MAP.length; m++) {
    if (label.indexOf(MAP[m].key) !== -1) return MAP[m];
  }
  return { debit: '諸経費', credit: '未払金', dt: 'purchase', ct: 'exempt' };
}

function calcJournalTax_(amount, key) {
  return (key === 'sale' || key === 'purchase') ? Math.round(amount * 10 / 110) : 0;
}

// 弥生会計形式（25列・公式サポート情報準拠）
// 識別フラグ,伝票No.,決算,取引日付,
// 借方勘定科目,借方補助科目,借方部門,借方税区分,借方金額,借方税金額,
// 貸方勘定科目,貸方補助科目,貸方部門,貸方税区分,貸方金額,貸方税金額,
// 摘要,番号,期日,タイプ,生成元,仕訳メモ,付箋1,付箋2,調整
// 税区分（税込み入力）: 課税売上込10% / 課対仕入込10% / 対象外
function buildJournalEntry_(label, amount, date, vNo) {
  var mp  = resolveJournalMapping_(label);
  var TAX = { sale: '課税売上込10%', purchase: '課対仕入込10%', exempt: '対象外' };
  return [
    '2000', String(vNo), '', date,
    mp.debit,  '', '', TAX[mp.dt], String(amount), String(calcJournalTax_(amount, mp.dt)),
    mp.credit, '', '', TAX[mp.ct], String(amount), String(calcJournalTax_(amount, mp.ct)),
    label, '', '', '', '', '', '', '', ''
  ];
}

// マネーフォワード クラウド会計形式（27列・公式サポート情報準拠）
// 取引No,取引日,
// 借方勘定科目,借方補助科目,借方部門,借方取引先,借方税区分,借方インボイス,借方金額(円),借方税額,
// 貸方勘定科目,貸方補助科目,貸方部門,貸方取引先,貸方税区分,貸方インボイス,貸方金額(円),貸方税額,
// 摘要,仕訳メモ,タグ,MF仕訳タイプ,決算整理仕訳,作成日時,作成者,最終更新日時,最終更新者
// 税区分（正式名称）: 課税売上 10% / 課税仕入 10% / 対象外
function buildJournalEntryMF_(label, amount, date) {
  var mp  = resolveJournalMapping_(label);
  var TAX = { sale: '課税売上 10%', purchase: '課税仕入 10%', exempt: '対象外' };
  return [
    '', date,
    mp.debit,  '', '', '', TAX[mp.dt], '', String(amount), String(calcJournalTax_(amount, mp.dt)),
    mp.credit, '', '', '', TAX[mp.ct], '', String(amount), String(calcJournalTax_(amount, mp.ct)),
    label, '', '', '', '', '', '', '', ''
  ];
}


// ================================================================
//  17-6b-1: シートCSV取得（exportSheetAsCsvBase64）  【大C / 中17 / 小17-6b-1】
//  指定シートのデータをBOM付きUTF-8 CSVにしてBase64で返す（ブラウザDL用）
// ================================================================
function exportSheetAsCsvBase64(sheetName) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 1) return { ok: false };
  var data = sheet.getRange(1, 1, sheet.getLastRow(), Math.max(sheet.getLastColumn(), 1)).getValues();
  var tz   = Session.getScriptTimeZone();
  var csv  = '﻿'; // BOM
  for (var i = 0; i < data.length; i++) {
    csv += data[i].map(function(cell) {
      var s;
      if (cell instanceof Date && !isNaN(cell.getTime())) {
        var yr = parseInt(Utilities.formatDate(cell, tz, 'yyyy'), 10);
        s = (yr <= 1900)
          ? Utilities.formatDate(cell, tz, 'H:mm')
          : Utilities.formatDate(cell, tz, 'yyyy/MM/dd');
      } else {
        s = String(cell === null || cell === undefined ? '' : cell);
      }
      return (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1)
        ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(',') + '\r\n';
  }
  return { ok: true, base64: Utilities.base64Encode(csv, Utilities.Charset.UTF_8), name: sheetName };
}


// ================================================================
//  17-6b-2: 選択シートExcel取得（exportSelectedSheetsAsExcel）  【大C / 中17 / 小17-6b-2】
//  指定シートのみ一時SSにコピーしてXLSX化・Base64で返す（ブラウザDL用）
// ================================================================
function exportSelectedSheetsAsExcel(sheetNames) {
  if (!sheetNames || sheetNames.length === 0) return { ok: false, msg: 'シートが選択されていません' };
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var tempSs = SpreadsheetApp.create('__temp_export__');
  var tempId = tempSs.getId();
  try {
    var copied = 0;
    for (var i = 0; i < sheetNames.length; i++) {
      var sheet = ss.getSheetByName(sheetNames[i]);
      if (sheet) { sheet.copyTo(tempSs).setName(sheetNames[i]); copied++; }
    }
    if (copied === 0) return { ok: false, msg: 'コピーできるシートがありませんでした' };
    tempSs.getSheets().forEach(function(s) {
      if (sheetNames.indexOf(s.getName()) === -1) {
        try { tempSs.deleteSheet(s); } catch(e) {}
      }
    });
    var url = 'https://docs.google.com/spreadsheets/d/' + tempId + '/export?format=xlsx';
    var res  = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
    var b64  = Utilities.base64Encode(res.getContent());
    var fname = 'export_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd') + '.xlsx';
    return { ok: true, base64: b64, fileName: fname };
  } finally {
    try { DriveApp.getFileById(tempId).setTrashed(true); } catch(e) {}
  }
}


// ================================================================
//  17-6b: CSV・Excel出力ダイアログ（showExportDialog）  【大C / 中17 / 小17-6b】
//  シートごとにチェックボックスでExcel対象を選択、CSV個別DLも可能
// ================================================================
function showExportDialog() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var rows = sheets.map(function(s) {
    var name = s.getName().replace(/\\/g, '\\\\').replace(/'/g, "\\'");
    return '<tr>'
      + '<td style="padding:4px 6px;border-bottom:1px solid #eee;width:20px">'
      + '<input type="checkbox" class="sheetCk" value="' + name + '" checked style="width:14px;height:14px;cursor:pointer;accent-color:#1565c0">'
      + '</td>'
      + '<td style="padding:4px 8px;border-bottom:1px solid #eee;font-size:13px">' + s.getName() + '</td>'
      + '<td style="padding:4px 6px;border-bottom:1px solid #eee">'
      + '<button onclick="dlCsv(\'' + name + '\')" style="background:#1565c0;color:#fff;padding:3px 10px;border:none;border-radius:4px;font-size:12px;cursor:pointer">CSV</button>'
      + '</td></tr>';
  }).join('');
  var html = '<html><body style="font-family:sans-serif;padding:16px;margin:0">'
    + '<script>'
    + 'function dlCsv(name){'
    +   'var btn=event.target; btn.disabled=true; btn.textContent="取得中…";'
    +   'google.script.run'
    +     '.withSuccessHandler(function(r){'
    +       'btn.disabled=false; btn.textContent="CSV";'
    +       'if(!r||!r.ok){alert("取得失敗");return;}'
    +       'var a=document.createElement("a");'
    +       'a.href="data:text/csv;charset=utf-8;base64,"+r.base64;'
    +       'a.download=r.name+".csv";'
    +       'document.body.appendChild(a);a.click();document.body.removeChild(a);'
    +     '})'
    +     '.withFailureHandler(function(e){btn.disabled=false;btn.textContent="CSV";alert(e.message);})'
    +     '.exportSheetAsCsvBase64(name);'
    + '}'
    + 'function selectAll(v){document.querySelectorAll(".sheetCk").forEach(function(c){c.checked=v;});}'
    + 'function dlExcel(){'
    +   'var names=[];'
    +   'document.querySelectorAll(".sheetCk:checked").forEach(function(c){names.push(c.value);});'
    +   'if(names.length===0){alert("シートを選択してください");return;}'
    +   'var btn=document.getElementById("btnXl"); btn.disabled=true; btn.textContent="生成中…（数秒かかります）";'
    +   'google.script.run'
    +     '.withSuccessHandler(function(r){'
    +       'btn.disabled=false; btn.textContent="📊 選択シートをExcel DL";'
    +       'if(!r||!r.ok){alert(r&&r.msg||"失敗");return;}'
    +       'var a=document.createElement("a");'
    +       'a.href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,"+r.base64;'
    +       'a.download=r.fileName;'
    +       'document.body.appendChild(a);a.click();document.body.removeChild(a);'
    +     '})'
    +     '.withFailureHandler(function(e){btn.disabled=false;btn.textContent="📊 選択シートをExcel DL";alert(e.message);})'
    +     '.exportSelectedSheetsAsExcel(names);'
    + '}'
    + '<\/script>'
    + '<h3 style="margin:0 0 10px;color:#1565c0">📤 データ出力</h3>'
    + '<div style="margin-bottom:8px">'
    + '<button id="btnXl" onclick="dlExcel()" style="background:#1b5e20;color:white;padding:8px 16px;border:none;border-radius:6px;font-size:13px;font-weight:bold;cursor:pointer;width:100%">📊 選択シートをExcel DL</button>'
    + '</div>'
    + '<div style="margin-bottom:6px;font-size:11px;color:#555">'
    + '☑ チェックでExcel対象 ／ <a href="#" onclick="selectAll(true);return false" style="color:#1565c0">全選択</a>'
    + ' | <a href="#" onclick="selectAll(false);return false" style="color:#1565c0">全解除</a>'
    + '</div>'
    + '<hr style="border:none;border-top:1px solid #eee;margin:4px 0 8px">'
    + '<p style="font-size:11px;color:#555;margin:0 0 4px">シート別CSV（BOM付きUTF-8）:</p>'
    + '<table style="border-collapse:collapse;width:100%">' + rows + '</table>'
    + '</body></html>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(400).setHeight(Math.min(160 + sheets.length * 34, 580)),
    '📤 データ出力'
  );
}


// ================================================================
//  17-7: 固定費マスタ初期化（initFixedCostMaster）  【大C / 中17 / 小17-7】
//  「PL設定」シートを初期構造・サンプルデータで作成する
//  PL含入フラグ=FALSE の行（社長給与等）はPL集計から自動除外される
// ================================================================
function initFixedCostMaster() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName('PL設定');

  var fullReset = true;
  if (sheet) {
    var initRes = ui.alert(
      'PL設定シートが既に存在します',
      '[はい]：全て上書き（追加項目も消えます）\n[いいえ]：デフォルト項目の金額のみ台数連動で更新（追加項目を保持）\n[キャンセル]：中止',
      ui.ButtonSet.YES_NO_CANCEL
    );
    if (initRes === ui.Button.CANCEL) return;
    fullReset = (initRes === ui.Button.YES);
    if (fullReset) { sheet.clear(); sheet.clearFormats(); }
  } else {
    sheet = ss.insertSheet('PL設定');
  }

  var header = ['費目名', '月額（円）', '按分方式', '勘定科目', 'PL含入フラグ'];
  if (fullReset) {
    sheet.getRange(1, 1, 1, 5).setValues([header])
      .setBackground('#1a2a3a').setFontColor('#00ff88').setFontWeight('bold').setFontSize(12);
  }

  // 自車専属マスタから「運行」台数をカウント（B列が「運行」の行のみ）
  var n = 1;
  var masterSh = ss.getSheetByName('自車専属マスタ');
  if (masterSh && masterSh.getLastRow() >= 2) {
    var mStatus = masterSh.getRange(2, 2, masterSh.getLastRow() - 1, 1).getValues();
    var mCount  = 0;
    for (var mi = 0; mi < mStatus.length; mi++) {
      if (String(mStatus[mi][0] || '').trim() === '運行') mCount++;
    }
    if (mCount > 0) n = mCount;
  }

  var defaults = [
    ['家賃',                          n * 30000,           '一定',     '地代家賃',         true  ],
    ['駐車場代',                      n * 15000,           '車両台数', '地代家賃',         true  ],
    ['電気代',                        8000 + n * 2000,     '一定',     '水道光熱費',       true  ],
    ['水道代',                        3000,                '一定',     '水道光熱費',       true  ],
    ['通信費（携帯）',                n * 5000,            '車両台数', '通信費',           true  ],
    ['通信費（固定回線・ネット）',     5000,               '一定',     '通信費',           true  ],
    ['損害保険料（任意）',            n * 20000,           '車両台数', '損害保険料',       true  ],
    ['車両リース代',                  n * 80000,           '車両台数', 'リース料',         true  ],
    ['事務用品費',                    3000,                '一定',     '消耗品費',         true  ],
    ['修繕費積立',                    n * 10000,           '車両台数', '修繕費',           true  ],
    ['減価償却費',                    n * 50000,           '車両台数', '減価償却費',       true  ],
    ['社長給与',                      300000,              '一定',     '役員報酬',         false ],  // ← PLに含めない
    ['法定福利費',                    n * 8000,            '車両台数', '法定福利費',       true  ],
    ['自賠責保険料積立',              n * 3000,            '車両台数', '損害保険料',       true  ],
    ['重量税積立',                    n * 5000,            '車両台数', '租税公課',         true  ],
    ['車検費積立',                    n * 8000,            '車両台数', '修繕費',           true  ],
    ['タイヤ代積立',                  n * 5000,            '車両台数', '消耗品費',         true  ],
    ['ETCリース料',                   n * 1500,            '車両台数', 'リース料',         true  ],
    ['カーナビリース料',               n * 2000,           '車両台数', 'リース料',         true  ],
    ['洗車費',                        n * 1000,            '車両台数', '消耗品費',         true  ],
    ['制服費',                        n * 1000,            '車両台数', '消耗品費',         true  ],
    ['税理士顧問料',                  30000,               '一定',     '支払手数料',       true  ],
    ['安全協会費',                    2000,                '一定',     '諸会費',           true  ]
  ];
  // [いいえ]選択時：デフォルト項目の金額だけ更新して終了（追加項目はそのまま）
  if (!fullReset) {
    var existLR = sheet.getLastRow();
    if (existLR >= 2) {
      var existNms = sheet.getRange(2, 1, existLR - 1, 1).getValues();
      var nameToRow = {};
      for (var ei = 0; ei < existNms.length; ei++) {
        var eName = String(existNms[ei][0] || '').trim();
        if (eName) nameToRow[eName] = ei + 2;
      }
      for (var di = 0; di < defaults.length; di++) {
        var dRow = nameToRow[defaults[di][0]];
        if (dRow !== undefined) {
          sheet.getRange(dRow, 2).setValue(defaults[di][1]).setFontColor('#1a9a50').setNumberFormat('#,##0');
        }
      }
    }
    ui.alert('PL設定の金額を更新しました。（稼働台数: ' + n + '台）\n追加項目はそのまま保持されています。');
    return;
  }

  sheet.getRange(2, 1, defaults.length, 5).setValues(defaults);
  // B列（月額）はすべて緑（稼働台数連動の目安値）
  sheet.getRange(2, 2, defaults.length, 1).setFontColor('#1a9a50');

  // 月額列に数値書式
  sheet.getRange(2, 2, defaults.length, 1).setNumberFormat('#,##0');
  // 按分方式ドロップダウン
  sheet.getRange(2, 3, defaults.length, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['一定','車両台数','稼働日数'], true).setAllowInvalid(false).build()
  );
  // PL含入フラグ チェックボックス
  sheet.getRange(2, 5, defaults.length, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build()
  );
  // 社長給与行を赤背景でわかりやすく
  var plFlagCol = sheet.getRange(2, 1, defaults.length, 5).getValues();
  for (var pi = 0; pi < plFlagCol.length; pi++) {
    if (plFlagCol[pi][4] === false) {
      sheet.getRange(pi + 2, 1, 1, 5).setBackground('#2a0a0a').setFontColor('#ff6666');
    }
  }

  sheet.setColumnWidth(1, 220); sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 100); sheet.setColumnWidth(4, 130); sheet.setColumnWidth(5, 100);
  sheet.setFrozenRows(1);
  ss.setActiveSheet(sheet);

  ui.alert(
    'PL設定シートを作成しました。（稼働台数: ' + n + '台 で金額を設定）\n\n' +
    '【按分方式の説明】\n' +
    '・一定    → 月額をそのまま使用\n' +
    '・車両台数 → 月額 × 稼働車両数 ÷ 全車両数\n' +
    '・稼働日数 → 月額 × 稼働日数 ÷ 暦日数\n\n' +
    '【PL含入フラグ】\n' +
    '・OFFにした費目はPLに含まれません（社長給与など秘匿項目に使用）\n' +
    '・赤背景行が現在OFFです。'
  );
}


// ================================================================
//  [ADD-v1.2] ポップアップから呼ばれ、指定された日付を起点に運行シートを同期する
// ================================================================
function executeStatusSync(rowsParam, choice, ssId) {
  if (choice === 'cancel') return;
  // rowsParam はJSON配列文字列または数値（後方互換）
  var rows = typeof rowsParam === 'string' ? JSON.parse(rowsParam) : [rowsParam];
  // getActiveSpreadsheetはAPIから呼ばれる時に正しいSSを返さないため getTargetSS_ を使う
  var ss = getTargetSS_(ssId);
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master) return;

  var applyDate;
  var now = new Date();
  if (choice === 'today') {
    applyDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  } else if (choice === 'month') {
    applyDate = new Date(now.getFullYear(), now.getMonth(), 1);
  } else {
    applyDate = new Date(2000, 0, 1);
  }

  for (var i = 0; i < rows.length; i++) {
    var mRowData = master.getRange(rows[i], 1, 1, 16).getValues()[0];
    // 最後の1台だけソートを実行（それ以前はskipSort=true）
    syncVehicleToCurrentMonth_(mRowData, i < rows.length - 1, applyDate, ss);
  }
}


// ================================================================
// ■ グループ13追記：ETC利用明細 CSV読込
// ================================================================

// ================================================================
//  13-10: ETC利用明細インポートダイアログ表示（showEtcImportDialog）  【大C / 中13 / 小13-10】
// ================================================================
function showEtcImportDialog() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  if (!ss.getSheetByName('集計表')) {
    var linked = PropertiesService.getUserProperties().getProperty('linkedSsId');
    if (linked) ssId = linked;
  }
  var tmpl = HtmlService.createTemplateFromFile('etcImport');
  tmpl.currentSsId = ssId;
  var html = tmpl.evaluate().setWidth(860).setHeight(540);
  SpreadsheetApp.getUi().showModalDialog(html, '⛽ ETC利用明細 読み込み');
}


// ================================================================
//  13-11: ETC照合準備（prepareEtcImport）  【大A / 中13 / 小13-11】
//  CSVテキストと列設定を受け取り、ETC行パース・重複車番チェック・手入力チェックを返す
//  conflicts: 1つのETC車番に集計表の車両が複数ある場合のリスト
//  carResolution: 重複なし（1対1）の場合の自動解決マップ
// ================================================================
function prepareEtcImport(csvText, colConfig, companySsId) {
  var ss      = getTargetSS_(companySsId);
  var unkouSh = ss.getSheetByName('運行');
  if (!unkouSh) throw new Error('運行シートが見つかりません');

  // CSVパース
  var lines = csvText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  if (lines.length < 2) return { etcRows: [], conflicts: [], carResolution: {}, manualCount: 0 };

  // colConfigがnullまたは未指定の場合はヘッダー行から自動検出
  if (!colConfig) {
    var headers = parseEtcCsvLine_(lines[0]);
    colConfig = detectEtcColumns_(headers);
    if (colConfig.colError) return { etcRows: [], conflicts: [], carResolution: {}, manualCount: 0, colError: colConfig.colError };
  }

  var etcRows = [];
  for (var i = 1; i < lines.length; i++) {
    var cols = parseEtcCsvLine_(lines[i]);
    if (!cols || cols.length < 2) continue;
    var rawCar    = String(cols[colConfig.colCarNum] || '').trim();
    var rawDateTo = String(cols[colConfig.colDateTo]  || '').trim();
    var rawTimeTo = colConfig.colTimeTo >= 0 ? String(cols[colConfig.colTimeTo] || '').trim() : '';
    var rawAmt    = String(cols[colConfig.colAmount]  || '').trim()
                      .replace(/,/g,'').replace(/[－−‐]/g,'-');
    if (!rawCar || !rawDateTo) continue;
    var amount = parseInt(rawAmt, 10);
    if (isNaN(amount) || amount <= 0) continue;
    var dt = parseEtcDateTime_(rawDateTo, rawTimeTo);
    if (!dt) continue;
    var carNum = extractCarNum_(rawCar);
    if (!carNum) continue;
    etcRows.push({ carNum: carNum, dt: dt.getTime(), amount: amount });
  }
  if (etcRows.length === 0) return { etcRows: [], conflicts: [], carResolution: {}, manualCount: 0, noData: true };

  // 運行シートの車番リストを取得（末尾数字 → 実車番リスト）
  var lr = unkouSh.getLastRow();
  if (lr < 2) return { etcRows: etcRows, conflicts: [], carResolution: {}, manualCount: 0 };
  var numCols  = 21;
  var sumData  = unkouSh.getRange(2, 1, lr - 1, numCols).getValues();
  var sumFonts = unkouSh.getRange(2, 21, lr - 1, 1).getFontColors();

  var carNumMap = {};
  for (var r = 0; r < sumData.length; r++) {
    var car = String(sumData[r][5] || '').trim();
    if (!car) continue;
    var num = extractCarNum_(car);
    if (!num) continue;
    if (!carNumMap[num]) carNumMap[num] = [];
    if (carNumMap[num].indexOf(car) < 0) carNumMap[num].push(car);
  }

  // 重複チェック & 自動解決マップ構築
  var conflictsMap  = {};
  var carResolution = {};
  var usedNums      = {};
  for (var j = 0; j < etcRows.length; j++) {
    var cn = etcRows[j].carNum;
    if (usedNums[cn]) continue;
    usedNums[cn] = true;
    var cands = carNumMap[cn] || [];
    if (cands.length === 0) {
      // 集計表に一致なし → スキップ扱い（空マップのまま）
    } else if (cands.length === 1) {
      carResolution[cn] = cands[0];
    } else {
      conflictsMap[cn] = cands;
    }
  }
  var conflicts = [];
  for (var key in conflictsMap) conflicts.push({ etcCar: key, candidates: conflictsMap[key] });

  // 手入力チェック（U列フォントカラーが#1a9a50でなく値がある行）
  var GREEN       = '#1a9a50';
  var manualCount = 0;
  for (var r2 = 0; r2 < sumData.length; r2++) {
    var uVal = sumData[r2][20];
    var uFc  = String(sumFonts[r2][0] || '').toLowerCase();
    if (uVal !== '' && uVal !== null && uFc !== GREEN) manualCount++;
  }

  return { etcRows: etcRows, conflicts: conflicts, carResolution: carResolution, manualCount: manualCount };
}


// ================================================================
//  13-12: ETC照合実行（executeEtcImport）  【大A / 中13 / 小13-12】
//  etcRows・車番解決マップ・手入力上書きフラグを受け取り集計表U列に書き込む
//  照合範囲：前行程の降完時刻 〜 この行程の降完時刻（降完なしはスキップ）
// ================================================================
function executeEtcImport(etcRows, carResolution, overwriteManual, companySsId) {
  var ss      = getTargetSS_(companySsId);
  var unkouSh = ss.getSheetByName('運行');
  if (!unkouSh) throw new Error('運行シートが見つかりません');

  var lr = unkouSh.getLastRow();
  if (lr < 2) return { updated: 0, skipped: 0 };
  var numCols  = 21; // A〜U列（U=実費高速）
  var unkouData = unkouSh.getRange(2, 1, lr - 1, numCols).getValues();
  var unkouFonts = unkouSh.getRange(2, 21, lr - 1, 1).getFontColors();
  var GREEN = '#1a9a50';
  var tz    = Session.getScriptTimeZone();

  // 車番+日付 → 運行シート行インデックスのマップ
  var carDateRows = {}; // key: "realCar|yyyy/M/d" → [rowIdx, ...]
  for (var r = 0; r < unkouData.length; r++) {
    var car = String(unkouData[r][5] || '').trim();
    if (!car) continue;
    var dv = unkouData[r][9]; // 日付
    if (!dv || !(dv instanceof Date)) continue;
    var ds  = Utilities.formatDate(dv, tz, 'yyyy/M/d');
    var key = car + '|' + ds;
    if (!carDateRows[key]) carDateRows[key] = [];
    carDateRows[key].push(r);
  }

  // ETCレコードを 車番+日付 単位で合算
  var etcAccum = {}; // key: "realCar|yyyy/M/d" → amount合計
  for (var e = 0; e < etcRows.length; e++) {
    var realCar = carResolution[etcRows[e].carNum];
    if (!realCar) continue;
    var etcDs = Utilities.formatDate(new Date(etcRows[e].dt), tz, 'yyyy/M/d');
    var key   = realCar + '|' + etcDs;
    etcAccum[key] = (etcAccum[key] || 0) + etcRows[e].amount;
  }

  // 運行シートU列（col 21）に書き込み、更新IDを収集
  var updated = 0, skipped = 0;
  var updatedIds = [];
  for (var key2 in etcAccum) {
    var rowIdxs = carDateRows[key2];
    if (!rowIdxs || rowIdxs.length === 0) continue;
    var totalAmt = etcAccum[key2];
    if (totalAmt <= 0) continue;
    var firstIdx = rowIdxs[0];
    var curVal   = unkouData[firstIdx][20];
    var curFc    = String(unkouFonts[firstIdx][0] || '').toLowerCase();
    var isManual = (curVal !== '' && curVal !== null && curVal !== 0 && curFc !== GREEN);
    if (isManual && !overwriteManual) { skipped++; continue; }
    unkouSh.getRange(firstIdx + 2, 21).setValue(totalAmt).setFontColor(GREEN).setNumberFormat('#,##0');
    var rowId = String(unkouData[firstIdx][0] || '').trim();
    if (rowId && updatedIds.indexOf(rowId) < 0) updatedIds.push(rowId);
    updated++;
  }

  // 更新したIDの集計表行を再同期（運行シートの値を集計表に反映）
  for (var u = 0; u < updatedIds.length; u++) {
    try { syncSummaryForId_(updatedIds[u], ss); } catch(e) {}
  }

  return { updated: updated, skipped: skipped };
}


// ================================================================
//  13-13: 車番末尾数字抽出（extractCarNum_）  【大B / 中13 / 小13-13】
//  「奈良お102」→「102」, 「大阪100か101」→「101」, 「55」→「55」
// ================================================================
function extractCarNum_(carStr) {
  var s = String(carStr || '').trim().replace(/\s/g, '');
  var m = s.match(/(\d+)$/);
  return m ? String(parseInt(m[1], 10)) : '';
}


// ================================================================
//  13-13b: ETC列自動検出（detectEtcColumns_）  【大B / 中13 / 小13-13b】
//  ヘッダー行の文字列から出口日付・出口時刻・料金・車番の列番号を自動判定する
// ================================================================
function detectEtcColumns_(headers) {
  var colDateTo = -1, colTimeTo = -1, colAmount = -1, colCarNum = -1;
  for (var i = 0; i < headers.length; i++) {
    var lh = headers[i].replace(/\s/g, '');
    // 出口（至）日付：「利用年月日（至）」「至」を含み時刻でないもの
    if (colDateTo < 0 && /利用年月日.*至|至.*利用年月日/.test(lh)) colDateTo = i;
    // 出口（至）時刻：「時分（至）」
    if (colTimeTo < 0 && /時分.*至|至.*時分/.test(lh)) colTimeTo = i;
    // 通行料金
    if (colAmount < 0 && /通行料金/.test(lh)) colAmount = i;
    // 車両番号（K列相当）：「車両番号」「車番」「車輌番号」
    if (colCarNum < 0 && /車両番号|車番|車輌番号/.test(lh)) colCarNum = i;
  }
  // フォールバック：「至」がヘッダーに単独で含まれる列を日付として使用
  if (colDateTo < 0) {
    for (var j = 0; j < headers.length; j++) {
      if (/至/.test(headers[j]) && !/時分/.test(headers[j])) { colDateTo = j; break; }
    }
  }
  var missing = [];
  if (colDateTo < 0) missing.push('出口日付（利用年月日（至））');
  if (colAmount < 0) missing.push('通行料金');
  if (colCarNum < 0) missing.push('車両番号');
  if (missing.length > 0) return { colError: '以下の列が見つかりません: ' + missing.join(' / ') };
  return { colDateTo: colDateTo, colTimeTo: colTimeTo, colAmount: colAmount, colCarNum: colCarNum };
}


// ================================================================
//  13-14: ETC CSV行パース補助（parseEtcCsvLine_）  【大B / 中13 / 小13-14】
// ================================================================
function parseEtcCsvLine_(line) {
  var result = [], cur = '', inQ = false;
  for (var i = 0; i < line.length; i++) {
    var c = line[i];
    if (c === '"') {
      if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
      else { inQ = !inQ; }
    } else if (c === ',' && !inQ) {
      result.push(cur); cur = '';
    } else {
      cur += c;
    }
  }
  result.push(cur);
  return result;
}


// ================================================================
//  13-15: ETC日時パース補助（parseEtcDateTime_）  【大B / 中13 / 小13-15】
//  「2026/6/1」+「7:54」→ Date。タイムゾーン Asia/Tokyo で解釈
// ================================================================
function parseEtcDateTime_(dateStr, timeStr) {
  var ds = String(dateStr || '').trim();
  var ts = String(timeStr || '').trim() || '00:00';
  var dp = ds.replace(/\//g, '-').split('-');
  if (dp.length !== 3) return null;
  var pad = function(n) { return (parseInt(n, 10) < 10 ? '0' : '') + parseInt(n, 10); };
  var y = parseInt(dp[0], 10);
  if (y < 100) y += 2000;  // 2桁年（26→2026）対応
  var iso = y + '-' + pad(dp[1]) + '-' + pad(dp[2]);
  var tp  = ts.split(':');
  var tIso = pad(tp[0] || '0') + ':' + pad(tp[1] || '0') + ':00';
  var dt  = new Date(iso + 'T' + tIso + '+09:00');
  return isNaN(dt.getTime()) ? null : dt;
}


// ================================================================
//  期限アラート（showExpiryAlert）
//  onOpen時に自車専属マスタの免許・教育・健診・適性の期限が近い乗務員を通知
//  checkMasterExpiries/checkExpiryDatesはゾンビトリガー向けの空デコイ
// ================================================================
// デコイ：ゾンビトリガーが発火しても何も起きない
function checkMasterExpiries(e) { return; }
function checkExpiryDates(e) { return; }

function showExpiryAlert() {
  // Layer1: キャッシュ高速チェック（ユーザー横断・ゾンビトリガーも物理ブロック）
  try { if (CacheService.getDocumentCache().get('EXPIRY_POPUP_SHOWN')) return; } catch(_ex0) {}
  // Layer2: DocumentLock + PropertiesService（競合回避・確実1回保証）
  try {
    var _dLock = LockService.getDocumentLock();
    if (!_dLock.tryLock(500)) return;
    var _dProps = PropertiesService.getDocumentProperties();
    var _dTs = Number(_dProps.getProperty('EXPIRY_POPUP_TS') || 0);
    if (Date.now() - _dTs < 30000) { _dLock.releaseLock(); return; }
    _dProps.setProperty('EXPIRY_POPUP_TS', String(Date.now()));
    try { CacheService.getDocumentCache().put('EXPIRY_POPUP_SHOWN', '1', 25); } catch(_ex2) {}
    _dLock.releaseLock();
  } catch(_ex) {}
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName('自車専属マスタ');
  if (!masterSheet || masterSheet.getLastRow() < 2) return;

  var headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
  var nameCol    = headers.indexOf('乗務員名');
  var licenseCol = headers.indexOf('免許証有効期限');
  var eduCol     = headers.indexOf('安全教育次回予定日');
  var healthCol  = headers.indexOf('健康診断次回予定日');
  var fitCol     = headers.indexOf('適性診断次回予定日');
  if (licenseCol === -1 && eduCol === -1 && healthCol === -1 && fitCol === -1) return;

  var data = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, masterSheet.getLastColumn()).getValues();
  var today = new Date(); today.setHours(0, 0, 0, 0);
  var day60 = new Date(today.getTime() + 60 * 24 * 60 * 60 * 1000);
  var day30 = new Date(today.getTime() + 30 * 24 * 60 * 60 * 1000);

  var checks = [
    { col: licenseCol, label: '免許証有効期限',     limit: day60 },
    { col: eduCol,     label: '安全教育次回予定日',  limit: day30 },
    { col: healthCol,  label: '健康診断次回予定日',  limit: day30 },
    { col: fitCol,     label: '適性診断次回予定日',  limit: day30 }
  ];

  var warnings = [];
  for (var i = 0; i < data.length; i++) {
    var name = nameCol >= 0 ? String(data[i][nameCol] || '').trim() : '';
    if (!name) continue;
    for (var j = 0; j < checks.length; j++) {
      var c = checks[j];
      if (c.col < 0) continue;
      var d = data[i][c.col];
      if (!d || !(d instanceof Date) || isNaN(d.getTime())) continue;
      var dd = new Date(d); dd.setHours(0, 0, 0, 0);
      if (dd <= c.limit) {
        var dStr  = Utilities.formatDate(dd, 'Asia/Tokyo', 'M/d');
        var label = dd < today ? '期限切れ' : 'まもなく期限';
        warnings.push('【' + label + '】' + name + '  ' + c.label + ' : ' + dStr);
      }
    }
  }

  if (warnings.length > 0) {
    SpreadsheetApp.getUi().alert('⚠ 期限アラート\n\n' + warnings.join('\n') + '\n\n自車専属マスタで確認・更新してください。');
  }
}


// ================================================================
//  4-11: 監査用表生成（generateAuditSheet）  【大C / 中4 / 小4-11】
//  メニュー「監査用表生成」: 集計表から改善基準告示コンプライアンス確認表（監査用シート）を月次小計付きで生成
// ================================================================
function generateAuditSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sumSheet = ss.getSheetByName('集計表');
  if (!sumSheet || sumSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('集計表にデータがありません');
    return;
  }

  // 集計表の列マップ（0-based）:
  // 0=ID,1=区分,2=会社名,3=トン数,4=車種,5=車番,6=乗務員名,7=携帯番号,8=看板名
  // 9=日付,10=荷主,11=積地,12=降地,13=誘導時刻,14=積完時刻,15=休憩開始,16=休憩終了,17=降完時刻
  // 18=売上,19=請求高速,20=実費高速,21=合計高速,22=距離,23=燃費,24=ガソリン代,25=燃料代
  // 26=支払い,27=経費合計,28=利益,29=備考,30=仮日数,31=給料,32=%,33=有休手当,34=その他手当
  // 35=点呼前完了,36=点呼後完了
  var auditCols = [
    { name: '日付',         src: 9  },
    { name: '区分',         src: 1  },
    { name: '会社名',       src: 2  },
    { name: '車種',         src: 4  },
    { name: '車番',         src: 5  },
    { name: '乗務員名',     src: 6  },
    { name: '荷主',         src: 10 },
    { name: '積地',         src: 11 },
    { name: '降地',         src: 12 },
    { name: 'トン数',       src: 3  },
    { name: '距離(km)',     src: 22 },
    { name: '点呼前完了',   src: 35 },
    { name: '誘導時刻',     src: 13 },
    { name: '積完時刻',     src: 14 },
    { name: '休憩開始',     src: 15 },
    { name: '休憩終了',     src: 16 },
    { name: '降完時刻',     src: 17 },
    { name: '点呼後完了',   src: 36 },
    { name: '拘束時間(h)', src: -1 },
    { name: '売上',         src: 18 },
    { name: '高速代(合計)', src: 21 },
    { name: '燃料代',       src: 25 },
    { name: '支払い',       src: 26 },
    { name: '経費合計',     src: 27 },
    { name: '利益',         src: 28 },
    { name: '有休手当',     src: 33 },
    { name: 'その他手当',   src: 34 },
    { name: '備考',         src: 29 },
    { name: '警告内容',     src: -2 }
  ];
  var numCols = auditCols.length;

  var lastRow = sumSheet.getLastRow();
  var lastCol = sumSheet.getLastColumn();
  var rawData = sumSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var rows = rawData.filter(function(r) { return r[9] instanceof Date && r[9].getTime() > 0; });
  rows.sort(function(a, b) { return a[9] - b[9]; });

  var auditSheet = ss.getSheetByName('監査用');
  if (!auditSheet) {
    auditSheet = ss.insertSheet('監査用');
  } else {
    auditSheet.clearContents();
    auditSheet.clearFormats();
  }

  var outputRows = [];
  outputRows.push(auditCols.map(function(c) { return c.name; }));

  var prevMonth = '';
  var mt = { u: 0, h: 0, n: 0, p: 0, k: 0, r: 0, y: 0, o: 0 };
  // 各データ行のシート行番号と元データを保持（時刻色付けに使用）
  var dataRowNums = [];

  function pushMonthTotal(month) {
    var tr = new Array(numCols).fill('');
    tr[0] = month + ' 合計';
    tr[19] = mt.u; tr[20] = mt.h; tr[21] = mt.n; tr[22] = mt.p;
    tr[23] = mt.k; tr[24] = mt.r; tr[25] = mt.y; tr[26] = mt.o;
    outputRows.push(tr);
  }

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var monthKey = Utilities.formatDate(r[9], 'Asia/Tokyo', 'yyyy/MM');
    if (prevMonth && monthKey !== prevMonth) {
      pushMonthTotal(prevMonth);
      mt = { u: 0, h: 0, n: 0, p: 0, k: 0, r: 0, g: 0, y: 0, o: 0 };
    }
    prevMonth = monthKey;

    var kosokunJikan = '';
    var tenkoMae = r[35]; var tenkoGo = r[36]; // 点呼前完了・点呼後完了（改善基準告示：始業〜終業点呼）
    if (tenkoMae instanceof Date && !isNaN(tenkoMae.getTime()) && tenkoGo instanceof Date && !isNaN(tenkoGo.getTime()) && tenkoGo > tenkoMae) {
      kosokunJikan = Math.round((tenkoGo.getTime() - tenkoMae.getTime()) / 36000) / 100;
    }

    // 警告内容を生成（時刻異常・拘束時間超えの理由を文字で記録）
    var warnArr = [];
    var F_w = 4 * 60 * 60 * 1000; var T_w = 30 * 60 * 1000;
    var pw = r[14], ksw = r[15], kew = r[16], dw = r[17];
    if (pw instanceof Date && ksw instanceof Date && !isNaN(pw.getTime()) && !isNaN(ksw.getTime()) && (ksw.getTime()-pw.getTime()) > F_w) warnArr.push('積完〜休憩4h超');
    if (ksw instanceof Date && kew instanceof Date && !isNaN(ksw.getTime()) && !isNaN(kew.getTime()) && (kew.getTime()-ksw.getTime()) < T_w) warnArr.push('休憩30分未満');
    if (kew instanceof Date && dw instanceof Date && !isNaN(kew.getTime()) && !isNaN(dw.getTime()) && (dw.getTime()-kew.getTime()) > F_w) warnArr.push('休憩後〜降完4h超');
    if (kosokunJikan !== '' && Number(kosokunJikan) > 13) warnArr.push('拘束13h超');
    var warnStr = warnArr.join(' / ');

    outputRows.push(auditCols.map(function(c) {
      if (c.src === -1) return kosokunJikan;
      if (c.src === -2) return warnStr;
      return (r[c.src] !== undefined && r[c.src] !== null) ? r[c.src] : '';
    }));
    dataRowNums.push({ sheetRow: outputRows.length, r: r, kosoku: kosokunJikan }); // setValues後の色付け用

    mt.u += Number(r[18]) || 0;
    mt.h += Number(r[21]) || 0;
    mt.n += Number(r[25]) || 0;
    mt.p += Number(r[26]) || 0;
    mt.k += Number(r[27]) || 0;
    mt.r += Number(r[28]) || 0;
    mt.y += Number(r[33]) || 0;
    mt.o += Number(r[34]) || 0;
  }
  if (prevMonth) pushMonthTotal(prevMonth);

  if (outputRows.length > 0) {
    auditSheet.getRange(1, 1, outputRows.length, numCols).setValues(outputRows);
  }

  // 時刻間隔異常の色付け（集計表と同基準）
  // 積完〜休憩開始 4時間超=黄 / 休憩時間 30分未満=水 / 休憩終了〜降完 4時間超=黄
  var F_a = 4 * 60 * 60 * 1000; // 4時間（労働時間過超の閾値）
  var T_a = 30 * 60 * 1000;     // 30分（休憩不足の閾値）
  for (var ci = 0; ci < dataRowNums.length; ci++) {
    var sr   = dataRowNums[ci].sheetRow;
    var dr   = dataRowNums[ci].r;
    var pick = dr[14]; var ks_a = dr[15]; var ke_a = dr[16]; var drop = dr[17];
    if (pick instanceof Date && ks_a instanceof Date && !isNaN(pick.getTime()) && !isNaN(ks_a.getTime()) && (ks_a.getTime() - pick.getTime()) > F_a) {
      auditSheet.getRange(sr, 14, 1, 2).setBackground('#ffd600'); // 積完・休憩開始を黄色
    }
    if (ks_a instanceof Date && ke_a instanceof Date && !isNaN(ks_a.getTime()) && !isNaN(ke_a.getTime()) && (ke_a.getTime() - ks_a.getTime()) < T_a) {
      auditSheet.getRange(sr, 15, 1, 2).setBackground('#4fc3f7'); // 休憩開始・終了を水色
    }
    if (ke_a instanceof Date && drop instanceof Date && !isNaN(ke_a.getTime()) && !isNaN(drop.getTime()) && (drop.getTime() - ke_a.getTime()) > F_a) {
      auditSheet.getRange(sr, 16, 1, 2).setBackground('#ffd600'); // 休憩終了・降完を黄色
    }
    // 拘束時間13時間超え：拘束時間セル（col 19）をオレンジ（利益マイナスの赤と区別）
    if (dataRowNums[ci].kosoku !== '' && Number(dataRowNums[ci].kosoku) > 13) {
      auditSheet.getRange(sr, 19, 1, 1).setBackground('#ff9800');
    }
  }

  // ヘッダー書式
  auditSheet.getRange(1, 1, 1, numCols)
    .setBackground('#1a3d6b').setFontColor('#ffffff').setFontWeight('bold');
  auditSheet.setFrozenRows(1);

  var dataRows = outputRows.length - 1;
  if (dataRows > 0) {
    // 日付列書式
    auditSheet.getRange(2, 1, dataRows, 1).setNumberFormat('yyyy/MM/dd');
    // 点呼前完了（col12）・点呼後完了（col18）の日時書式
    auditSheet.getRange(2, 12, dataRows, 1).setNumberFormat('M/d H:mm');
    auditSheet.getRange(2, 18, dataRows, 1).setNumberFormat('M/d H:mm');
    // 時刻列書式（誘導〜降完: 13〜17列目 1-based）
    auditSheet.getRange(2, 13, dataRows, 5).setNumberFormat('M/d H:mm');
    // 金額列書式（売上〜その他手当: 20〜27列目）
    auditSheet.getRange(2, 20, dataRows, 8).setNumberFormat('#,##0');
  }

  // 月次合計行ハイライト
  for (var mr = 2; mr <= outputRows.length; mr++) {
    if (String(outputRows[mr - 1][0]).indexOf('合計') !== -1) {
      auditSheet.getRange(mr, 1, 1, numCols)
        .setBackground('#fff9c4').setFontWeight('bold');
    }
  }

  auditSheet.autoResizeColumns(1, numCols);
  ss.setActiveSheet(auditSheet);
  SpreadsheetApp.getUi().alert('監査用表を生成しました（' + rows.length + '件 / 月次合計付き）');
}


// ================================================================
//  No.4 自動バックアップ・復旧機能
//  ・runDailyBackup_()          : タイムトリガーから毎日3時に実行（①修正用SSで設定）
//  ・setupBackupTrigger()       : ①修正用SSメニューからトリガー設定
//  ・openRestoreDialog()        : メニューから直接呼ぶ（①=会社選択→復旧 / ③=直接復旧）
//  ・getBackupListForRestore(a) : 復旧ダイアログから google.script.run 経由
//  ・executeRestore(a, b)       : 復旧ダイアログから google.script.run 経由
// ================================================================
var BACKUP_FOLDER_NAME_ = '運行管理バックアップ';
var BACKUP_KEEP_DAYS_   = 30;

function getOrCreateBackupRoot_() {
  var it = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME_);
  return it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME_);
}

function getOrCreateCompanyBackupFolder_(companyName) {
  var root = getOrCreateBackupRoot_();
  var it   = root.getFoldersByName(companyName);
  return it.hasNext() ? it.next() : root.createFolder(companyName);
}

// タイムトリガーから毎日呼ばれる（①修正用SSで setupBackupTrigger() を1回実行して設定）
function runDailyBackup_() {
  var adminSsId = PropertiesService.getScriptProperties().getProperty('masterSsId');
  if (!adminSsId) { Logger.log('runDailyBackup_: masterSsId未設定。①で setupBackupTrigger を実行してください。'); return; }
  var ss       = SpreadsheetApp.openById(adminSsId);
  var regSheet = ss.getSheetByName('会社登録');
  if (!regSheet || regSheet.getLastRow() < 2) return;

  var today   = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var lastRow = regSheet.getLastRow();
  var rows    = regSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var cutoff  = new Date(); cutoff.setDate(cutoff.getDate() - BACKUP_KEEP_DAYS_);

  rows.forEach(function(row) {
    var companyName = String(row[0] || '').trim();
    var ssUrl       = String(row[5] || '').trim();
    if (!companyName || !ssUrl) return;
    var match = ssUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!match) return;
    try {
      var folder   = getOrCreateCompanyBackupFolder_(companyName);
      var fileName = today + '_' + companyName;
      // 同日分が既にあれば削除してから再作成
      var existing = folder.getFilesByName(fileName);
      while (existing.hasNext()) existing.next().setTrashed(true);
      DriveApp.getFileById(match[1]).makeCopy(fileName, folder);
      // 保持期間を超えた古いファイルを削除
      var all = folder.getFiles();
      while (all.hasNext()) {
        var f = all.next();
        if (f.getDateCreated() < cutoff) f.setTrashed(true);
      }
    } catch(e) {
      Logger.log('Backup error [' + companyName + ']: ' + e.message);
    }
  });
}

// ①修正用SSのメニューから実行してタイムトリガーを設定する
function setupBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runDailyBackup_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runDailyBackup_').timeBased().atHour(3).everyDays(1).create();
  SpreadsheetApp.getUi().alert('バックアップタイマーを設定しました\n毎日深夜3時に③各客SSを自動バックアップします\n（マイドライブ→運行管理バックアップ フォルダに保存）');
}

// 毎月20日 0:00 に全客SSの翌月分を自動生成する（①修正用SSで setupMonthlyTrigger を1回実行して登録）
function scheduledGenerateNextMonth_() {
  var adminSsId = PropertiesService.getScriptProperties().getProperty('masterSsId');
  if (!adminSsId) { Logger.log('scheduledGenerateNextMonth_: masterSsId未設定'); return; }
  var adminSs  = SpreadsheetApp.openById(adminSsId);
  var regSheet = adminSs.getSheetByName('会社登録');
  if (!regSheet || regSheet.getLastRow() < 2) return;
  var rows = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, 6).getValues();
  rows.forEach(function(row) {
    var companyName = String(row[0] || '').trim();
    var ssUrl       = String(row[5] || '').trim();
    if (!companyName || !ssUrl) return;
    var m = ssUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!m) return;
    try {
      generateNextMonthSilent_(SpreadsheetApp.openById(m[1]));
    } catch(e) {
      Logger.log('月次自動生成エラー [' + companyName + ']: ' + e.message);
    }
  });
}

// UIなし版 generateNextMonth（時間トリガーから呼ぶ・ss を引数で受け取る）
function generateNextMonthSilent_(ss) {
  var sheet = ss.getSheetByName('運行');
  if (!sheet) return;
  var today    = new Date();
  var nextYear = (today.getMonth() === 11) ? today.getFullYear() + 1 : today.getFullYear();
  var nextMon  = (today.getMonth() + 1) % 12;
  var lastRow  = sheet.getLastRow();
  if (lastRow >= 2) {
    var dateVals = sheet.getRange(2, 10, lastRow - 1, 1).getValues();
    for (var i = 0; i < dateVals.length; i++) {
      var dv = dateVals[i][0];
      if (dv instanceof Date && dv.getFullYear() === nextYear && dv.getMonth() === nextMon) return;
    }
  }
  var master = ss.getSheetByName('自車専属マスタ');
  if (!master || master.getLastRow() < 2) return;
  var mData          = master.getRange(2, 1, master.getLastRow() - 1, 16).getValues();
  var activeVehicles = mData.filter(function(r) { return String(r[1] || '').trim() === '運行'; });
  if (activeVehicles.length === 0) return;
  var daysInMonth = new Date(nextYear, nextMon + 1, 0).getDate();
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(e) { return; }
  try {
    var insertRow = sheet.getLastRow() + 1;
    var nextNum   = getNextIdNum_(sheet, 'V-');
    var rowsData  = [], formulas = [];
    for (var day = 1; day <= daysInMonth; day++) {
      var dateObj = new Date(nextYear, nextMon, day);
      for (var v = 0; v < activeVehicles.length; v++) {
        var veh   = activeVehicles[v];
        var rowId = 'V-' + String(nextNum).padStart(4, '0');
        nextNum++;
        var rn = insertRow + rowsData.length;
        rowsData.push([
          rowId,  veh[2], veh[3], veh[5], veh[6], veh[7], veh[8], veh[9],
          veh[4], dateObj, '', '', '',
          '', '', '', '', '',
          '', '', '',
          '',
          '', '', '', ''
        ]);
        formulas.push(['=IF(AND(U' + rn + '="",T' + rn + '=""),"",U' + rn + '-T' + rn + ')']);
      }
    }
    sheet.getRange(insertRow, 1, rowsData.length, 26).setValues(rowsData);
    sheet.getRange(insertRow, 22, formulas.length, 1).setFormulas(formulas);
    sheet.getRange(insertRow, 10, rowsData.length, 1).setNumberFormat('yyyy/MM/dd');
    sheet.getRange(insertRow, 12, rowsData.length, 2).setNumberFormat('@');
    commitLastId_(sheet, 'V-', nextNum - 1);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  sortUnkouByDate_(ss.getId());
  archiveOldestMonthIfNeeded_(ss);
}

// ①修正用SSのメニューから実行してタイムトリガーを設定する（1回だけ押せばOK）
function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'scheduledGenerateNextMonth_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scheduledGenerateNextMonth_').timeBased().onMonthDay(20).atHour(0).create();
  SpreadsheetApp.getUi().alert('月次自動生成タイマーを設定しました\n毎月20日 0:00 に全客SSへ翌月分を自動生成します');
}

// メニューから直接呼ばれる（getActiveSpreadsheet使用可）
// ①修正用SS→会社選択ダイアログ、③客SS→直接バックアップ一覧ダイアログ
function openRestoreDialog() {
  var ss     = SpreadsheetApp.getActiveSpreadsheet();
  var ssId   = ss.getId();
  var ssName = ss.getName();
  if (ss.getSheetByName('会社登録')) {
    _showCompanySelectorForRestore_();
  } else {
    _showRestoreDialog_(ssId, ssName);
  }
}

function _showCompanySelectorForRestore_() {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var regSheet = ss.getSheetByName('会社登録');
  if (!regSheet || regSheet.getLastRow() < 2) { SpreadsheetApp.getUi().alert('会社登録シートにデータがありません'); return; }
  var rows = regSheet.getRange(2, 1, regSheet.getLastRow() - 1, 6).getValues();
  var opts = '';
  rows.forEach(function(row) {
    var name  = String(row[0] || '').trim();
    var ssUrl = String(row[5] || '').trim();
    if (!name || !ssUrl) return;
    var m = ssUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!m) return;
    opts += '<option value="' + m[1] + '|' + name.replace(/[<>"]/g, '') + '">' + name + '</option>';
  });
  if (!opts) { SpreadsheetApp.getUi().alert('会社データがありません'); return; }

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>body{font-family:sans-serif;padding:16px;background:#1e1e1e;color:#e0e0e0;margin:0;font-size:14px;}' +
    'select,button{width:100%;padding:10px;border-radius:8px;font-size:14px;box-sizing:border-box;margin-bottom:10px;}' +
    'select{background:#2c2c2c;color:white;border:1px solid #444;}' +
    '.btn-next{background:#1565c0;border:none;color:white;cursor:pointer;font-weight:bold;}' +
    '.btn-do{background:#c62828;border:none;color:white;cursor:pointer;font-weight:bold;}' +
    '.warn{color:#ff9800;font-size:12px;}</style></head><body>' +
    '<p style="margin-bottom:8px;">復旧する会社を選んでください</p>' +
    '<select id="sel">' + opts + '</select>' +
    '<button class="btn-next" onclick="next()">次へ →</button>' +
    '<div id="step2"></div>' +
    '<script>' +
    'function next(){' +
    '  var v=document.getElementById("sel").value.split("|");' +
    '  document.getElementById("step2").innerHTML="<p style=\'color:#aaa;font-size:12px;\'>読み込み中...</p>";' +
    '  google.script.run.withSuccessHandler(function(list){showList(list,v[0],v[1]);}).getBackupListForRestore(v[0]);' +
    '}' +
    'function showList(list,sid,name){' +
    '  if(!list||!list.length){document.getElementById("step2").innerHTML="<p style=\'color:#ff5252;\'>バックアップがありません。先にタイマー設定を実行してください。</p>";return;}' +
    '  var h="<hr style=\'border-color:#333;\'><p style=\'margin-bottom:6px;\'>"+name+"　復旧日を選択</p>";' +
    '  h+="<select id=\'b\'>";' +
    '  list.forEach(function(b){h+="<option value=\'"+b.id+"\'>"+b.name+"</option>";});' +
    '  h+="</select><p class=\'warn\'>⚠ 現在のデータが上書きされます</p>";' +
    '  h+="<button class=\'btn-do\' onclick=\'restore(\\\""+sid+"\\\")\'> 🔄 復旧実行</button>";' +
    '  document.getElementById("step2").innerHTML=h;' +
    '}' +
    'function restore(sid){' +
    '  var bid=document.getElementById("b").value;' +
    '  if(!confirm("復旧しますか？\\n現在のデータが上書きされます"))return;' +
    '  google.script.run.withSuccessHandler(function(){alert("復旧完了しました");google.script.host.close();}).executeRestore(bid,sid);' +
    '}' +
    '</script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(370).setHeight(300),
    '🔄 バックアップから復旧'
  );
}

function _showRestoreDialog_(ssId, ssName) {
  var root = getOrCreateBackupRoot_();
  var it   = root.getFoldersByName(ssName);
  var list = [];
  if (it.hasNext()) {
    var folder = it.next(); var files = folder.getFiles();
    while (files.hasNext()) { var f = files.next(); list.push({ id: f.getId(), name: f.getName() }); }
    list.sort(function(a, b) { return a.name > b.name ? -1 : 1; });
  }

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>body{font-family:sans-serif;padding:16px;background:#1e1e1e;color:#e0e0e0;margin:0;font-size:14px;}' +
    'select,button{width:100%;padding:10px;border-radius:8px;font-size:14px;box-sizing:border-box;margin-bottom:10px;}' +
    'select{background:#2c2c2c;color:white;border:1px solid #444;}' +
    'button{background:#c62828;border:none;color:white;cursor:pointer;font-weight:bold;}' +
    '.warn{color:#ff9800;font-size:12px;}</style></head><body>';

  if (list.length === 0) {
    html += '<p style="color:#ff5252;">バックアップがありません。<br>①修正用SSで「バックアップタイマー設定」を実行してください。</p>';
  } else {
    var opts = list.map(function(b) { return '<option value="' + b.id + '">' + b.name + '</option>'; }).join('');
    html += '<p style="margin-bottom:8px;">' + ssName + '</p>';
    html += '<select id="b">' + opts + '</select>';
    html += '<p class="warn">⚠ 現在のデータが上書きされます</p>';
    html += '<button onclick="restore()">🔄 復旧実行</button>';
    html += '<script>function restore(){var bid=document.getElementById("b").value;' +
      'if(!confirm("復旧しますか？\\n現在のデータが上書きされます"))return;' +
      'google.script.run.withSuccessHandler(function(){alert("復旧完了しました");google.script.host.close();})' +
      '.executeRestore(bid,"' + ssId + '");}</script>';
  }
  html += '</body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(360).setHeight(240),
    '🔄 バックアップから復旧'
  );
}

// google.script.run 経由で呼ばれる（①の会社選択ダイアログから）
function getBackupListForRestore(ssId) {
  var ss     = getTargetSS_(ssId);
  var ssName = ss.getName();
  var root   = getOrCreateBackupRoot_();
  var it     = root.getFoldersByName(ssName);
  if (!it.hasNext()) return [];
  var folder = it.next(); var files = folder.getFiles(); var list = [];
  while (files.hasNext()) { var f = files.next(); list.push({ id: f.getId(), name: f.getName() }); }
  list.sort(function(a, b) { return a.name > b.name ? -1 : 1; });
  return list;
}

// google.script.run 経由で呼ばれる（復旧ダイアログから）
// バックアップSSのデータ行を対象SSに上書きコピー（ヘッダー行は変えない）
function executeRestore(backupFileId, targetSsId) {
  var targetSs = getTargetSS_(targetSsId);
  var backupSs = SpreadsheetApp.openById(backupFileId);
  var sheets   = backupSs.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var srcSheet  = sheets[i];
    var sheetName = srcSheet.getName();
    if (sheetName.charAt(0) === '_') continue; // _BK_等の内部シートはスキップ
    var tgtSheet = targetSs.getSheetByName(sheetName);
    if (!tgtSheet) continue;
    var lastRow = srcSheet.getLastRow();
    var lastCol = srcSheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) continue;
    var data       = srcSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var tgtLastRow = tgtSheet.getLastRow();
    if (tgtLastRow > 1) tgtSheet.getRange(2, 1, tgtLastRow - 1, tgtSheet.getLastColumn()).clearContent();
    tgtSheet.getRange(2, 1, data.length, lastCol).setValues(data);
  }
  return true;
}
