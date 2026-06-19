# 成績送信用 GAS（mado_score_gas v2.0）

英検ループ訓練アプリの「先生に送信」から届く成績を、Google スプレッドシートに記録するための Google Apps Script です。

## 役割
アプリ（`index.html`）が送る JSON を受け取り、スプレッドシートに1行ずつ追記します。

**受け取る JSON（`text/plain` で POST）**
```json
{ "cls":"3A", "num":"5", "name":"山田", "ver":"v0.23",
  "grade":"3級", "mode":"フル", "score":60, "total":93,
  "pct":65, "pass":"合格" }
```

**記録カラム（v2.0で 級・モード・正答率・合否 を追加）**

| 日時 | 組 | 番号 | 名前 | 級 | モード | 正解 | 問題数 | 正答率(%) | 合否 | 版 |
|---|---|---|---|---|---|---|---|---|---|---|

## セットアップ手順
1. 成績を貯めるスプレッドシートを用意（新規でOK）。
2. メニュー「**拡張機能 → Apps Script**」を開く。
3. 下のコードを貼り付けて保存。
4. 「**デプロイ → 新しいデプロイ → ウェブアプリ**」
   - 実行するユーザー：**自分**
   - アクセスできるユーザー：**全員（匿名含む）**
   - → 表示される **ウェブアプリURL** をコピー。
5. `index.html` 内の `var GAS_URL = "..."` にそのURLを貼る。

> 💡 既存のGASプロジェクトを使う場合は、コードを差し替えて「**デプロイを管理 → 編集 → バージョン：新バージョン**」で更新すると **URLが変わらず**、`index.html` の変更が不要です。

> ✅ 動作確認：ウェブアプリURLをブラウザで開いて `{"result":"ok",...}` が返ればデプロイ成功。

## コード

```javascript
/**
 * mado_score_gas_v2_0.gs  |  英検ループ訓練 — 成績受信用 GAS  |  v2.0
 */

// ===== 設定 =====
var SHEET_ID   = "";        // 空ならこのスクリプトが紐づくスプレッドシートを使用。別管理なら開きたいシートのIDを入れる
var SHEET_NAME = "成績";    // 書き込むシート名（無ければ自動作成）

// ===== エンドポイント =====
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // 同時送信での行ずれを防ぐ

    if (!e || !e.postData || !e.postData.contents) {
      return json_({ result: "error", message: "no body" });
    }
    var data;
    try { data = JSON.parse(e.postData.contents); }
    catch (err) { return json_({ result: "error", message: "bad json" }); }

    var cls = str_(data.cls), num = str_(data.num);
    if (!cls || !num) {
      return json_({ result: "error", message: "cls and num are required" });
    }

    var sheet = getSheet_();
    ensureHeaders_(sheet);

    var now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([
      now,
      cls,
      num,
      str_(data.name),
      str_(data.grade),
      str_(data.mode),
      num_(data.score),
      num_(data.total),
      num_(data.pct),
      str_(data.pass),
      str_(data.ver)
    ]);

    return json_({ result: "ok" });
  } catch (err) {
    return json_({ result: "error", message: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// 動作確認用（ブラウザでURLを開くと {"result":"ok",...} が返る）
function doGet() {
  return json_({ result: "ok", service: "mado_score_gas", ver: "2.0" });
}

// ===== ヘルパー =====
function getSheet_() {
  var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("スプレッドシートが見つかりません。SHEET_ID を設定するか、スプレッドシートに紐づくプロジェクトで実行してください。");
  }
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  return sh;
}

function ensureHeaders_(sh) {
  if (sh.getLastRow() === 0) {
    sh.appendRow(["日時", "組", "番号", "名前", "級", "モード", "正解", "問題数", "正答率(%)", "合否", "版"]);
    sh.setFrozenRows(1);
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function str_(v) { return (v === null || v === undefined) ? "" : String(v); }
function num_(v) { var n = Number(v); return isNaN(n) ? "" : n; }
```
