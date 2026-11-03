// 献立ノート（献立コスト計算Webアプリ）用 Google Apps Script
//
// 使い方
//  1. 新しい Google スプレッドシートを作る（名前は自由。例：献立ノート）
//  2. 「拡張機能」→「Apps Script」を開き、このコードを全部貼り付けて保存
//  3. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
//       次のユーザーとして実行：自分 / アクセスできるユーザー：全員
//  4. 表示された「ウェブアプリのURL」をアプリの「設定」に貼り付ける
//
// シートは最初のアクセス時に自動で作られます（買い物・料理記録・レシピ・献立予定・休日・表記ゆれ辞書）。
// 1行目の見出しは変更しないでください。

const TABLES = {
  purchases: { sheet: '買い物',       cols: ['id', 'date', 'name', 'qty', 'unit', 'price', 'updatedAt'] },
  cooks:     { sheet: '料理記録',     cols: ['id', 'date', 'slot', 'dish', 'ingredient', 'qty', 'updatedAt'] },
  recipes:   { sheet: 'レシピ',       cols: ['id', 'dish', 'ingredient', 'qty', 'unit', 'updatedAt'] },
  plans:     { sheet: '献立予定',     cols: ['id', 'date', 'slot', 'dish1', 'dish2', 'dish3', 'budget', 'updatedAt'] },
  holidays:  { sheet: '休日',         cols: ['id', 'date', 'updatedAt'] },
  aliases:   { sheet: '表記ゆれ辞書', cols: ['id', 'alias', 'name', 'updatedAt'] },
  seedPrices:{ sheet: '初期単価',     cols: ['id', 'name', 'unit', 'price', 'updatedAt'] }
};

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'all';
  if (action === 'ping') return json_({ ok: true, app: 'kondate' });
  const out = { ok: true };
  Object.keys(TABLES).forEach(function (t) { out[t] = readTable_(t); });
  return json_(out);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'upsert') {
      upsert_(data.table, data.rows || []);
    } else if (data.action === 'delete') {
      deleteIds_(data.table, data.ids || []);
    } else if (data.action === 'replaceDish') {
      replaceDish_(data.dish, data.rows || []);
    } else {
      return json_({ ok: false, error: '不明な操作です: ' + data.action });
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'doPost失敗: ' + (err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// ---------- 内部処理 ----------
function sheet_(table) {
  const def = TABLES[table];
  if (!def) throw new Error('不明なテーブルです: ' + table);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(def.sheet);
  if (!sh) {
    sh = ss.insertSheet(def.sheet);
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight('bold');
    // 日付やIDが勝手に日付型・数値型に変換されないよう、文字列として扱う列を書式設定
    def.cols.forEach(function (c, i) {
      if (['id', 'date', 'name', 'unit', 'slot', 'dish', 'ingredient', 'dish1', 'dish2', 'dish3', 'alias'].indexOf(c) >= 0) {
        sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@');
      }
    });
    sh.setFrozenRows(1);
  }
  return sh;
}

function cellOut_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return v;
}

function readTable_(table) {
  const sh = sheet_(table);
  const values = sh.getDataRange().getValues();
  const headers = values.shift() || [];
  return values
    .filter(function (r) { return r[0] !== '' && r[0] !== null; })
    .map(function (r) {
      const o = {};
      headers.forEach(function (h, i) { o[h] = cellOut_(r[i]); });
      o.id = String(o.id);
      return o;
    });
}

function idIndex_(sh) {
  const last = sh.getLastRow();
  const map = {};
  if (last < 2) return map;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  ids.forEach(function (r, i) { map[String(cellOut_(r[0]))] = i + 2; });
  return map;
}

function upsert_(table, rows) {
  if (!rows.length) return;
  const def = TABLES[table];
  const sh = sheet_(table);
  const map = idIndex_(sh);
  const now = new Date().toISOString();
  const appends = [];
  rows.forEach(function (row) {
    row.updatedAt = now;
    const line = def.cols.map(function (c) { return row[c] === undefined || row[c] === null ? '' : row[c]; });
    const at = map[String(row.id)];
    if (at) sh.getRange(at, 1, 1, line.length).setValues([line]);
    else appends.push(line);
  });
  if (appends.length) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, def.cols.length).setValues(appends);
  }
}

function deleteIds_(table, ids) {
  if (!ids.length) return;
  const sh = sheet_(table);
  const map = idIndex_(sh);
  const rowsToDelete = ids.map(function (id) { return map[String(id)]; })
    .filter(function (r) { return r; })
    .sort(function (a, b) { return b - a; });
  rowsToDelete.forEach(function (r) { sh.deleteRow(r); });
}

function replaceDish_(dish, rows) {
  const sh = sheet_('recipes');
  const last = sh.getLastRow();
  if (last >= 2) {
    const dishes = sh.getRange(2, 2, last - 1, 1).getValues();
    for (let i = dishes.length - 1; i >= 0; i--) {
      if (String(dishes[i][0]) === String(dish)) sh.deleteRow(i + 2);
    }
  }
  upsert_('recipes', rows);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 動作確認用：GASエディタでこの関数を選んで実行すると、シートが全部作られます。
function setup() {
  Object.keys(TABLES).forEach(function (t) { sheet_(t); });
}
