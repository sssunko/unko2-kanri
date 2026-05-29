# 運行管理システム 引き継ぎ文書

---

## システム概要
- Google Apps Script (GAS) V8 + SpreadsheetApp + PropertiesService
- ローカルファイル: `c:\gas\unko2-kanri\コード.js` / `index.html`
- デプロイ: `npx clasp push` → 既存デプロイIDを指定して更新（古いデプロイは自動削除）
- バージョン上限対策: `clasp deploy -i {ID} -d "番号_内容"` で上書き更新のみ

---

## SS・スクリプト・デプロイID一覧

| 役割 | SS ID | スクリプトID | WebApp デプロイID |
|------|-------|-------------|-----------------|
| ①修正用SS（ライブラリ・心臓部） | `.clasp.json` 参照 | `1n79omnAcdsEojMRyjnj9-Ic9pIl1-7Nt_HB7Avy0NVFizOSeqt0guqyZ` | `AKfycbw7rzkd_SuE1I6BNzEjED4Mxl6cnM4wbswIiRiNoPf5zcSS2JcP6YLkfRV21fLc0opU` |
| ②客用SS（テスト・テンプレート） | `1NBtosd_MN8KcboV_4OXTrY8WqcE3TJwpxdA_nASmTOo` | `19CfyUPhldzSccj05xo-sn4Xh78fCHAHDVJtGyKdDGQkO1D4wZWFEnZCT` | `AKfycbxP6x0cdhr8WzUaP-u_XlspPwY9EGvh8D3qleTVeOKRxawZFmV6rabbbR1ROHYzTvhD` |
| ③各客SS | 会社登録シートのF列参照 | 会社登録シートのK列参照 | 会社登録シートのG列参照 |

---

## アーキテクチャ（3層構成）

```
①修正用SS（心臓部・ライブラリ UnkouLib として公開）
    ↓ メニュー「📤 テスト客SSに反映」（シート構成・書式のみ・データ消さない）
②客用SS（テスト専用・スタブのみ）
    ↓ メニュー「📤 各客に反映」（シート構成・書式のみ・データ消さない）
③各客SS × N社（本番・スタブのみ）
```

**各SSは独自のWebApp URLを持つ（ssId不要・完全独立）**
- ①: `...exec`（管理者用）
- ②: `...exec`（テスト用）
- ③: 会社登録時に Apps Script API で自動デプロイ・各社固有URL

**ライブラリバージョン管理**
- ②③のスタブは固定バージョン（現在 `260`）を参照
- `「📤 テスト客SSに反映」` を押すと新バージョン作成 → ②のスタブを更新
- ②の `「📤 各客に反映」` で全③を同じバージョンに更新
- ①でコードをpushしても②③には自動で変わらない（手動反映必須）

---

## ローカルファイル構成

| ファイル/フォルダ | 役割 |
|-----------------|------|
| `コード.js` | サーバーサイド全処理（①にデプロイ） |
| `index.html` | フロントエンド（CSS+HTML+JS 全部1ファイル） |
| `appsscript.json` | ①のマニフェスト（WebApp設定・スコープ） |
| `stub_for_clientSS/` | ②客用SSのスタブ（clasp push 先は②のスクリプトID） |
| `stub_clientSS.js` | stub_for_clientSS/コード.js のフラットコピー |
| `deploy.ps1` | push + deploy + 古いバージョン削除を1発実行 |
| `push_stub_clientSS.ps1` | ②へのスタブpush専用スクリプト |

---

## デプロイ手順

### ①修正用SSへのデプロイ
```powershell
# c:\gas\unko2-kanri\ で実行
$PROD_ID = "AKfycbw7rzkd_SuE1I6BNzEjED4Mxl6cnM4wbswIiRiNoPf5zcSS2JcP6YLkfRV21fLc0opU"
clasp push --force
clasp deploy -i $PROD_ID -d "番号_内容"
```

### ②客用SSへのスタブpush
```powershell
cd stub_for_clientSS
clasp push --force
```

---

## 永続ルール（絶対守れ・毎回確認）

### 作業ルール
1. **依頼してない箇所は絶対壊すな・消すな・変えるな**
2. 実走前に必ずやる事を日本語で説明 → ユーザーOK後に実走
3. 全部確認してこっちがOK言ったら実走して
4. 不明点は聞け
5. 毎回このタスク一覧を出して状況確認してから作業開始
6. 毎回履歴（コード・git log）見て確認してから作業しろ
7. `npx clasp push` まで自分でやれ。ユーザーはF5押すだけ

### 説明ルール
- **コードで話すな。依頼内容に沿って日本語で話せ**
- 修正内容の説明禁止。操作手順（何を押すか）を毎回書け
- 要約・まとめ文は不要。結果と次の操作だけ書け

### 設計ルール
- ボタン遷移は1個ずつ（絶対遷移最優先）
- 処理はSSに送り裏で処理 → SSから値で返す（アプリを絶対重くするな）
- アプリはSSから値で受け取って表示するだけ

### SS紐づけ大原則（絶対守れ）
- ①②③各SSとアプリは1対1。違うSSのデータが別のSSに入ることはあり得ない
- `companySsId` を受け取っている関数は必ず `getTargetSS_(companySsId)` で対象SSを取得
- `SpreadsheetApp.getActiveSpreadsheet()` をAPIとして呼ばれる関数で使うな（常に①を返す）
- 反映ボタンはシート構成・書式のみ更新。データ行は絶対消さない

### デプロイルール
- **デプロイ名は「番号_依頼内容（簡潔）」で命名する**
- 例: `270_行追加時刻9時固定バグ修正` / `272_アーカイブ削除一括化`

---

## 最初にやること（新チャット開始時）

1. このHANDOVER.mdを最初から最後まで読む
2. 下記タスクリストの「今」の項目だけを抽出してユーザーに見せる
3. ユーザーにOKもらってから作業開始（コード全部読むのは必要になった時だけ）

---

## タスクリスト

| # | 内容 | 状態 |
|---|------|------|
| 1 | 集計表の自動反映修正（flush順序 + zSyncVal） | 済 |
| 2 | 運行再開ボタンの判定修正（点呼前・点呼後追加） | 済 |
| 3 | 各SS独自WebApp化（②③が①を通らないアーキテクチャ） | 済 |
| 4 | HANDOVER.md更新（元SS記述を新構成に） | 済 |
| 5 | 💴 経費自動入力ボタン実装（自車専属マスタ・トン数別平均値） | 未 |
| 6 | 説明書作成 | 未 |
