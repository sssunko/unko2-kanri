# 運行管理システム 技術仕様・開発ガイド

---

## システム概要

- プラットフォーム: Google Apps Script (GAS) V8 + SpreadsheetApp + PropertiesService
- ローカルパス: `c:\gas\unko2-kanri\`
- デプロイ: `node deploy.js "説明"` によるワンコマンド自動化

---

## SS・スクリプト・デプロイID一覧

| 役割 | SS ID | WebApp デプロイID |
|------|-------|-----------------|
| ①開発用SS（ライブラリ） | .clasp.json 参照 | AKfycbw7rzkd_SuE1I6BNzEjED4Mxl6cnM4wbswIiRiNoPf5zcSS2JcP6YLkfRV21fLc0opU |
| ②検証用SS | 1NBtosd_MN8KcboV_4OXTrY8WqcE3TJwpxdA_nASmTOo | AKfycbxP6x0cdhr8WzUaP-u_XlspPwY9EGvh8D3qleTVeOKRxawZFmV6rabbbR1ROHYzTvhD |
| ③本番SS群 | 会社登録シートF列 | 会社登録シートG列 |

---

## アーキテクチャ（3層構成）

```text
①開発用SS（UnkouLib ライブラリ）
  │
  ├─ 「📤 テスト客SS（②）に反映」ボタン
  ▼
②検証用SS（スタブ経由でライブラリ呼び出し）
  │
  ├─ 「📤 各客に反映」ボタン
  ▼
③本番SS群（顧客ごとに独立したSpreadsheet）
```

- 各SSは独自WebApp URLで完全独立（ssId不要）
- ②③はスタブコードのみを持ち、処理はすべてライブラリ（①）に委譲
- ライブラリのバージョンは「反映」ボタンを押した時点で固定される
- コードを修正しても反映ボタンを押すまで②③には波及しない（意図的設計）

---

## ファイル構成

```
コード.js                 ← メインライブラリ（GAS・全処理の実装）
index.html               ← フロントエンド（モバイルアプリUI）
stub_for_clientSS/
  コード.js               ← ②③用スタブ（このファイルのみ編集）
  appsscript.json        ← deploy.js が自動更新
build_stub.js            ← getClientStubSource_() 自動生成ツール（Node.js）
check_integrity.js       ← デプロイ前整合性チェック（Node.js）
deploy.js                ← ワンコマンドデプロイスクリプト（Node.js）
```

---

## デプロイ手順

### ワンコマンドデプロイ

```
cd c:\gas\unko2-kanri
node deploy.js "修正内容の説明"
```

以下を自動実行する：

1. 整合性チェック（HTMLのスタブ未定義関数を検出・エラー時は停止）
2. `getClientStubSource_()` を `stub_for_clientSS/コード.js` から自動再生成
3. `clasp push --force` → ①のコードを更新
4. `clasp deploy` → WebAppを更新（次バージョン番号を自動付与）
5. スタブpush → ②のスタブを更新
6. `stub_for_clientSS/appsscript.json` のバージョンを自動更新

個別の clasp コマンドは使用しない。必ずこのスクリプト経由でデプロイすること。

### デプロイ後の手順（手動操作）

```
STEP 2: ①開発用SS → F5 → 「📤 テスト客SS（②）に反映」
         ※ 新しいライブラリバージョンが作成される。省略すると②③に旧コードが入る。
STEP 3: ②検証用SS → F5 → 動作確認
STEP 4: ②（または①）→ 「📤 各客に反映」
```

デプロイ後は必ず上記STEP 2〜4の操作指示をユーザーに提示すること。

### バージョン管理

- `clasp versions` でバージョン件数を確認
- 190件以上になったら、デプロイ前にユーザーへ削除を依頼する
  - 削除場所: ①開発用SS → 拡張機能 → Apps Script → 左メニューの時計アイコン（バージョン管理）
- デプロイ名の形式: `番号_説明`（例: `490_reloadMenuバグ修正`）

---

## 設計原則

### データフロー原則

- フロントエンド（HTML）はサーバー関数の呼び出しと表示のみを担当
- 処理はすべてGAS側（SS）で実行し、結果値を返す
- UIの応答性を保つため、重い処理をフロントエンドに持たせない

### SS参照規約

- `companySsId` を受け取る関数は必ず `getTargetSS_(companySsId)` でSS取得
- `google.script.run` 経由で呼ばれる関数内で `SpreadsheetApp.getActiveSpreadsheet()` を使用しない
  - `ssId` を引数で受け取るか、インストール済みトリガーなら `e.source` を使用
- インストール済みトリガーからのUI表示（`showModalDialog`）はスタブ側で実施
- トリガー登録（`ScriptApp.newTrigger`）はライブラリではなくスタブ側で実施

### スタブ管理規約

- スタブの修正は `stub_for_clientSS/コード.js` のみで行う
- `getClientStubSource_()` 内の自動生成領域（`AUTO_GENERATED_STUB_START〜END` 間）は手動編集禁止
- 新しい公開関数を追加した場合は `stub_for_clientSS/コード.js` にスタブを追加し、`node deploy.js` を実行
- `check_integrity.js` がHTMLから呼ばれているがスタブに未定義の関数を自動検出する

### データ保護規約

- 「反映」ボタンはシート構成・書式のみを更新する。データ行の削除・変更は行わない
- 新規顧客SS作成時は `createCompanySpreadsheet_` → `initClientSSSheets_` の順で呼ぶ（検証用SSのデータ混入防止）
- 新規顧客SSのテストデータ仕様:
  - 自車専属マスタは2行必須（1行目: 給料行、2行目: ％行）
  - 給料と％を同一行に記載しない

---

## 開発規約

### 作業規約

1. 依頼範囲外のコードは変更しない
2. 実装前に変更内容を説明し、確認を得てから実施する
3. 不明点は推測で進めず確認する
4. デプロイは必ず `node deploy.js` を使用する
5. デプロイ後は必ずSTEP 2〜4の操作指示をユーザーに提示する
6. 修正完了の定義: ①②③すべてで同じ動作を確認できる状態

### 開発セッション管理

- `コード.js` は150行以下に分割して読む（大容量ファイルのコンテキスト管理）
- 1セッションで扱うタスクは1件を基本とする
- セッション開始時は本ドキュメントを読んでからタスクを確認する

---

## 製品ロードマップ

| # | 内容 | 状態 |
|---|------|------|
| 1 | システム安定化（排他ロック・認証） | ✅ 完了 |
| 2 | データ一括読込（配車表CSV/Excel対応） | ✅ 完了 |
| 3 | 自動送信機能（指示書・車番連絡） | ✅ 完了 |
| 4 | 技術引き継ぎ資料・マニュアル整備 | 進行中 |
| 5 | 利用規約・法務対応 | 未着手 |
| 6 | 追加機能（PL自動生成・ETC明細連携等） | 未着手 |
