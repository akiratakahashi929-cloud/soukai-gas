/**
 * ApiHandlers.gs — GAS Web App APIハンドラー群
 * ─────────────────────────────────────────────────────────────
 * Next.jsアプリからの読み書きリクエストを処理する。
 * Code.gs の doGet / doPost から呼び出される。
 * ─────────────────────────────────────────────────────────────
 */

// ============================================================
// GET: action=getData — 全シートデータをJSONで返す
// ============================================================
function handleGetData(ss) {
  try {
    var result = {};

    // 設定シート: key-value オブジェクトとして返す
    var settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    if (settingsSheet && settingsSheet.getLastRow() >= 2) {
      var sd = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 2).getValues();
      var settingsObj = {};
      sd.forEach(function(r) {
        if (r[0]) settingsObj[String(r[0]).trim()] = String(r[1] || '').trim();
      });
      result.settings = settingsObj;
    } else {
      result.settings = {};
    }

    // 各シートを2次元配列で返す
    result.activity        = sheetToJson(ss, SHEET_NAMES.ACTIVITY,       7);
    result.prefectural     = sheetToJson(ss, SHEET_NAMES.PREFECTURAL,    4);
    result.org             = sheetToJson(ss, SHEET_NAMES.ORG,            3);
    result.members         = sheetToJson(ss, SHEET_NAMES.MEMBERS,        9);
    result.income_summary  = sheetDisplayJson(ss, SHEET_NAMES.INCOME_SUMMARY,  5);
    result.expense_summary = sheetDisplayJson(ss, SHEET_NAMES.EXPENSE_SUMMARY, 5);
    result.assets          = sheetDisplayJson(ss, SHEET_NAMES.ASSETS,         3);
    result.budget_in       = sheetDisplayJson(ss, SHEET_NAMES.BUDGET_IN,      5);
    result.budget_out      = sheetDisplayJson(ss, SHEET_NAMES.BUDGET_OUT,     5);
    result.plan            = sheetToJson(ss, SHEET_NAMES.PLAN,          2);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return errorJson('handleGetData: ' + err.toString());
  }
}

function sheetToJson(ss, name, cols) {
  var s = ss.getSheetByName(name);
  if (!s || s.getLastRow() < 2) return [];
  var n = Math.max(s.getLastRow() - 1, 1);
  return s.getRange(2, 1, n, cols).getValues();
}

function sheetDisplayJson(ss, name, cols) {
  var s = ss.getSheetByName(name);
  if (!s || s.getLastRow() < 2) return [];
  var n = Math.max(s.getLastRow() - 1, 1);
  return s.getRange(2, 1, n, cols).getDisplayValues();
}

// ============================================================
// POST: saveSettings — 設定シートの値を更新
// ============================================================
function handleSaveSettings(ss, data) {
  try {
    var s = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    if (!s) return errorJson('設定シートが見つかりません');
    var settings = data.settings || {};
    if (s.getLastRow() >= 2) {
      var rows = s.getRange(2, 1, s.getLastRow() - 1, 2).getValues();
      rows.forEach(function(r, i) {
        var key = String(r[0]).trim();
        if (key && settings[key] !== undefined) {
          s.getRange(i + 2, 2).setValue(settings[key]);
        }
      });
    }
    SpreadsheetApp.flush();
    return okJson('saveSettings');
  } catch (err) {
    return errorJson('saveSettings: ' + err.toString());
  }
}

// ============================================================
// POST: saveActivity — 活動報告シートを置換
// ============================================================
function handleSaveActivity(ss, data) {
  return replaceSheetRows(ss, SHEET_NAMES.ACTIVITY, data.rows, 7, 'saveActivity');
}

// ============================================================
// POST: savePrefectural — 県青年部報告シートを置換
// ============================================================
function handleSavePrefectural(ss, data) {
  return replaceSheetRows(ss, SHEET_NAMES.PREFECTURAL, data.rows, 4, 'savePrefectural');
}

// ============================================================
// POST: saveOrg — 組織情報シートを置換
// ============================================================
function handleSaveOrg(ss, data) {
  return replaceSheetRows(ss, SHEET_NAMES.ORG, data.rows, 3, 'saveOrg');
}

// ============================================================
// POST: saveMembers — 会員名簿シートを置換
// ============================================================
function handleSaveMembers(ss, data) {
  return replaceSheetRows(ss, SHEET_NAMES.MEMBERS, data.rows, 9, 'saveMembers');
}

// ============================================================
// POST: savePlan — 事業計画シートを置換
// ============================================================
function handleSavePlan(ss, data) {
  return replaceSheetRows(ss, SHEET_NAMES.PLAN, data.rows, 2, 'savePlan');
}

// ============================================================
// POST: saveFinancial — 決算集計シートの「決算額」列を直接更新
//   ※ SUMIF数式は保持したまま、数値セルのみ上書き
// ============================================================
function handleSaveFinancial(ss, data) {
  try {
    if (data.income && Array.isArray(data.income)) {
      updateSummaryActuals(ss, SHEET_NAMES.INCOME_SUMMARY, data.income);
    }
    if (data.expense && Array.isArray(data.expense)) {
      updateSummaryActuals(ss, SHEET_NAMES.EXPENSE_SUMMARY, data.expense);
    }
    SpreadsheetApp.flush();
    return okJson('saveFinancial');
  } catch (err) {
    return errorJson('saveFinancial: ' + err.toString());
  }
}

function updateSummaryActuals(ss, sheetName, items) {
  var s = ss.getSheetByName(sheetName);
  if (!s || s.getLastRow() < 2) return;
  var existing = s.getRange(2, 1, s.getLastRow() - 1, 3).getValues();
  items.forEach(function(item) {
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][0]).trim() === String(item.category).trim()) {
        s.getRange(i + 2, 3).setValue(Number(item.actual) || 0); // C列＝決算額
        break;
      }
    }
  });
}

// ============================================================
// POST: saveAssets — 財産目録の現金・預金・負債を更新
//   data.rows = [genkin, futsu, teiki, miharai]
// ============================================================
function handleSaveAssets(ss, data) {
  try {
    var s = ss.getSheetByName(SHEET_NAMES.ASSETS);
    if (!s || s.getLastRow() < 2) return errorJson('財産目録シートが見つかりません');

    var rows = data.rows || [0, 0, 0, 0];
    var map = {
      '現金':     Number(rows[0]) || 0,
      '普通預金': Number(rows[1]) || 0,
      '定期預金': Number(rows[2]) || 0,
      '流動負債': Number(rows[3]) || 0,
      '未払い金': Number(rows[3]) || 0,
    };

    var existing = s.getRange(2, 1, s.getLastRow() - 1, 2).getValues();
    existing.forEach(function(r, i) {
      var key = String(r[0]).trim();
      if (map[key] !== undefined) {
        s.getRange(i + 2, 2).setValue(map[key]);
      }
    });
    SpreadsheetApp.flush();
    return okJson('saveAssets');
  } catch (err) {
    return errorJson('saveAssets: ' + err.toString());
  }
}

// ============================================================
// 共通: ヘッダー行を残してデータ行を全置換
// ============================================================
function replaceSheetRows(ss, sheetName, rows, numCols, actionName) {
  try {
    var s = ss.getSheetByName(sheetName);
    if (!s) return errorJson(sheetName + 'シートが見つかりません');

    if (!rows || !Array.isArray(rows) || rows.length === 0) {
      return okJson(actionName + ' (no rows)');
    }

    // ヘッダー(1行目)を保持して2行目以降をクリア
    var lastRow = s.getLastRow();
    if (lastRow > 1) {
      var lastCol = Math.max(s.getLastColumn(), numCols);
      s.getRange(2, 1, lastRow - 1, lastCol).clearContent();
    }

    // numCols列に正規化して書き込み
    var writeRows = rows.map(function(r) {
      var row = Array.isArray(r) ? r.slice() : [r];
      while (row.length < numCols) row.push('');
      return row.slice(0, numCols);
    });

    s.getRange(2, 1, writeRows.length, numCols).setValues(writeRows);
    SpreadsheetApp.flush();
    return okJson(actionName);
  } catch (err) {
    return errorJson(actionName + ': ' + err.toString());
  }
}

// ============================================================
// レスポンスユーティリティ
// ============================================================
function okJson(action) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', action: action }))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorJson(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
