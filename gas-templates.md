# GAS（Google Apps Script）追加コード

`external-match.html` と `stats.html` の他校試合機能を有効化するために、Apps Script エディタに以下を追記してください。

> **注意**: 既存の `doGet(e)` / `saveMatch` などのコードは触らず、**追加・修正** していきます。

---

## 1. `doGet` のディスパッチに新アクションを追加

既存の `doGet(e)` 関数の `switch` / `if` 分岐に以下を追加：

```js
function doGet(e) {
  const action = e.parameter.action;
  const type = e.parameter.type;  // ← 追加：'all' | 'practice' | 'external'

  // ... 既存の case ...

  if (action === 'saveExternalMatch') {
    return saveExternalMatch(JSON.parse(e.parameter.data));
  }

  // 既存の取得系も type パラメータを渡す
  if (action === 'getPersonalStats') return getPersonalStats(type);
  if (action === 'getMatchDetail')   return getMatchDetail(type);
  if (action === 'getParentStats')   return getParentStats(type);
  if (action === 'getParentDetail')  return getParentDetail(type);

  // ... 既存の他のハンドラ ...
}
```

---

## 2. `saveExternalMatch` 関数を追加

```js
function saveExternalMatch(data) {
  // matches シートを開く（既存 saveMatch と同じシート想定）
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('matches');

  // 1行 = 1試合の記録。シート列の構成は既存 saveMatch と合わせる + type 列を持つ前提。
  // 既存シートに type 列がなければ末尾に追加してください。
  // 例: type | date | opponent | teamA(or members) | teamB | winner | halves(JSON) | hits(JSON) | parents | substitutions

  sheet.appendRow([
    'external',                                  // type
    data.date,                                   // 試合日
    data.opponent,                               // 対戦校
    JSON.stringify(data.teamMembers),            // 自チームメンバー
    '',                                          // 相手チーム（不要なので空）
    data.winner,                                 // 'us' | 'them' | 'draw'
    JSON.stringify(data.halves),                 // 前半・後半の詳細
    JSON.stringify(mergedHits(data.halves)),     // 全半合計のヒット数
    JSON.stringify(data.parents),                // 保護者
    JSON.stringify(data.substitutions || []),    // 交代記録
    data.totalOurBest,                           // 自チーム合計ベスト
    data.totalTheirBest                          // 相手合計ベスト
  ]);

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function mergedHits(halves) {
  const out = {};
  halves.forEach(h => {
    Object.keys(h.hits || {}).forEach(name => {
      out[name] = (out[name] || 0) + h.hits[name];
    });
  });
  return out;
}
```

---

## 3. 既存 `saveMatch` に `type='practice'` を保存

`stats.html` のフィルタを正しく効かせるため、既存の練習試合保存にも種別を入れます。

```js
function saveMatch(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('matches');

  sheet.appendRow([
    'practice',                  // ← 追加: type
    new Date(),                  // date
    '',                          // opponent（練習なので空）
    JSON.stringify(data.teamA),
    JSON.stringify(data.teamB),
    data.winner,
    // ... 以降は既存と同じ形でOK
  ]);

  // ...
}
```

> シートのカラム順は既存環境に合わせて並べ替えてください。**重要なのは1列目に `type` を入れること**（または末尾でも可）。フィルタはこの列で分岐します。

---

## 4. 取得系（`getPersonalStats` 等）に type フィルタを追加

```js
function getPersonalStats(type) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('matches');
  const rows = sheet.getDataRange().getValues();
  // ヘッダー行をスキップ
  const data = rows.slice(1);

  // type フィルタ
  const filtered = data.filter(row => {
    const rowType = row[0] || 'practice'; // 既存データは practice 扱い
    if (!type || type === 'all') return true;
    return rowType === type;
  });

  // 以降は既存集計ロジックをそのまま使う（filtered を入力に）
  // 例: 子どもごとの win/lose/draw/survive/total を集計
  // ...

  return ContentService.createTextOutput(JSON.stringify({ stats: result }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

同じパターンで `getMatchDetail` / `getParentStats` / `getParentDetail` も `type` 引数を受け取り、`filtered` だけ集計するよう変更してください。

---

## 5. デプロイ

修正後：

1. Apps Script エディタ右上の **「デプロイ」 → 「デプロイを管理」**
2. 既存デプロイの右にある鉛筆アイコンで **「新しいバージョン」** を選んで更新
3. デプロイURLは変わらないので、フロントの `GAS_URL` 定数は触らなくてOK

---

## 6. 確認

- `external-match.html` で1試合記録 → スプレッドシートに `type=external` の行が追加されているか
- 既存 `dodgeball.html` で1試合記録 → `type=practice` の行が追加されているか
- `stats.html` で「すべて／練習のみ／他校試合のみ」フィルタを切替 → 数値が変わるか

問題が出たら、ブラウザの開発者ツール（Network タブ）で実際のレスポンスを見てください。
