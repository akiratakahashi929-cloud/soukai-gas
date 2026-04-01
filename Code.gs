// ============================================================
// 総会資料自動生成システム v4
// Code.gs — エントリポイント + スプレッドシート管理
// SS_ID: 1OEN6G_AeeLEfkApeuJWOJLkvSfiJJwiAce9NKWPZso8
// ============================================================

var SS_ID = '1OEN6G_AeeLEfkApeuJWOJLkvSfiJJwiAce9NKWPZso8';

var SHEET_NAMES = {
  SETTINGS:       '設定',
  ACTIVITY:       '活動報告',
  PREFECTURAL:    '県青年部報告',
  INCOME_SUMMARY: '決算_収入',
  EXPENSE_SUMMARY:'決算_支出',
  INCOME_DETAIL:  '収入_明細',
  EXPENSE_DETAIL: '支出_明細',
  ASSETS:         '財産目録',
  BUDGET_IN:      '予算_収入',
  BUDGET_OUT:     '予算_支出',
  PLAN:           '事業計画',
  ORG:            '組織情報',
  MEMBERS:        '会員名簿',
  RULES:          '規約',
  DASHBOARD:      'ダッシュボード'
};

// ============================================================
// メニュー
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('総会資料')
    .addItem('シートを初期化',     'initializeAllSheets').addSeparator()
    .addItem('明細から集計更新',   'refreshSummary')
    .addItem('入力チェック',       'validateAllSheets')
    .addItem('HTMLプレビュー',     'openHtmlPreview')
    .addItem('PDF出力',            'generatePdf')
    .addToUi();
}

function openHtmlPreview() {
  var html = generateSoukaiHtml();
  var ui = SpreadsheetApp.getUi();
  var output = HtmlService.createHtmlOutput(html).setWidth(900).setHeight(700);
  ui.showModalDialog(output, '総会資料プレビュー');
}

// ============================================================
// Web App エントリポイント
// ============================================================

/**
 * GET: デフォルト → HTML総会資料を返す
 *      action=html → HTML総会資料
 *      action=addTransaction → 取引データ追加
 *      action=status → API疎通確認
 */
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';

    if (!action || action === 'html') {
      var html = generateSoukaiHtml();
      return HtmlService.createHtmlOutput(html)
        .setTitle('総会資料')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // ── アプリ → SS データ取得 ──
    if (action === 'getData') {
      var ss = SpreadsheetApp.openById(SS_ID);
      return handleGetData(ss);
    }

    if (action === 'addTransaction') {
      var ss = SpreadsheetApp.openById(SS_ID);
      var data = JSON.parse(e.parameter.data);
      return processData(ss, data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'Soukai GAS System v4' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<pre style="color:red;padding:20px;font-size:12px">エラーが発生しました:\n' +
      String(err) + '</pre>'
    ).setTitle('エラー');
  }
}

/**
 * POST: 取引データ一括追加
 * Body JSON: { type:"income"|"expense", transactions:[{date,category,description,amount,...}] }
 */
function doPost(e) {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var data = JSON.parse(e.postData.contents);
    var action = data.action || '';

    // ── アプリ → SS 書き込みルーティング ──
    if (action === 'saveSettings')    return handleSaveSettings(ss, data);
    if (action === 'saveActivity')    return handleSaveActivity(ss, data);
    if (action === 'savePrefectural') return handleSavePrefectural(ss, data);
    if (action === 'saveOrg')         return handleSaveOrg(ss, data);
    if (action === 'saveMembers')     return handleSaveMembers(ss, data);
    if (action === 'savePlan')        return handleSavePlan(ss, data);
    if (action === 'saveFinancial')   return handleSaveFinancial(ss, data);
    if (action === 'saveAssets')      return handleSaveAssets(ss, data);

    // 後方互換: 既存の取引明細追加（action=addTransaction等）
    return processData(ss, data);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function processData(ss, data) {
  var type = data.type;
  var sheetName = (type === 'income') ? SHEET_NAMES.INCOME_DETAIL : SHEET_NAMES.EXPENSE_DETAIL;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Sheet not found: ' + sheetName }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var txs = data.transactions || [];
  txs.forEach(function(tx) {
    sheet.appendRow([
      tx.date, tx.category, tx.description, tx.amount,
      tx.recipient || '', tx.note || ''
    ]);
  });
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', count: txs.length }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// シート初期化
// ============================================================
function initializeAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var r = ui.alert('確認', '全シートを初期化しますか？\n※既存データは上書きされます。', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;

  setupSettingsSheet(ss);
  setupActivitySheet(ss);
  setupPrefecturalSheet(ss);
  setupIncomeDetail(ss);
  setupExpenseDetail(ss);
  setupIncomeSummary(ss);
  setupExpenseSummary(ss);
  setupAssetsSheet(ss);
  setupBudgetInSheet(ss);
  setupBudgetOutSheet(ss);
  setupPlanSheet(ss);
  setupOrgSheet(ss);
  setupMembersSheet(ss);
  setupRulesSheet(ss);
  setupDashboard(ss);

  ui.alert('完了', '全シートの初期化が完了しました。', ui.ButtonSet.OK);
}

// ── 設定シート ──
function setupSettingsSheet(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.SETTINGS); s.clear();
  var d = [
    ['項目',           '値',                       '説明'],
    ['年度',           '令和7年度',                 '当年度（例: 令和7年度）'],
    ['回数',           '39',                        '総会回数（算用数字）'],
    ['総会開催日',     '令和7年5月10日（土）',       ''],
    ['開会時刻',       '16時00分',                   ''],
    ['会場',           '料亭 角家 三芳町上富',       ''],
    ['会計年度開始',   '令和6年4月1日',              '前年度開始（報告対象）'],
    ['会計年度終了',   '令和7年3月31日',             '前年度終了（報告対象）'],
    ['次年度開始',     '令和7年4月1日',              '今年度開始（計画対象）'],
    ['次年度終了',     '令和8年3月31日',             '今年度終了（計画対象）'],
    ['会計担当',       '宇田川 圭太',                ''],
    ['監査担当',       '佐々木 義人',                ''],
    ['監査日',         '令和7年4月○日',             ''],
    ['報告日',         '令和7年5月10日',             '決算・財産目録の報告日']
  ];
  s.getRange(1, 1, d.length, 3).setValues(d);
  styleH(s, 3);
  s.setColumnWidth(1, 120); s.setColumnWidth(2, 200); s.setColumnWidth(3, 200);
}

// ── 活動報告シート ──
function setupActivitySheet(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.ACTIVITY); s.clear();
  s.getRange(1, 1, 1, 7).setValues([['年', '月日', '支部事業名', '場所', '担当会社', '県事業名', '県場所']]);
  styleH(s, 7);
  s.setColumnWidth(1, 60); s.setColumnWidth(2, 80); s.setColumnWidth(3, 200);
  s.setColumnWidth(4, 120); s.setColumnWidth(5, 100); s.setColumnWidth(6, 200); s.setColumnWidth(7, 120);
}

// ── 県青年部報告シート ──
function setupPrefecturalSheet(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.PREFECTURAL); s.clear();
  s.getRange(1, 1, 1, 4).setValues([['年', '月日', '事業名', '場所']]);
  styleH(s, 4);
  s.setColumnWidth(1, 60); s.setColumnWidth(2, 80); s.setColumnWidth(3, 250); s.setColumnWidth(4, 200);
}

// ── 収入明細シート ──
function setupIncomeDetail(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.INCOME_DETAIL); s.clear();
  s.getRange(1, 1, 1, 6).setValues([['日付', '科目', '内容', '金額', '相手先', '備考']]);
  styleH(s, 6);
  s.getRange('D:D').setNumberFormat('#,##0');
}

// ── 支出明細シート ──
function setupExpenseDetail(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.EXPENSE_DETAIL); s.clear();
  s.getRange(1, 1, 1, 6).setValues([['日付', '科目', '内容', '金額', '相手先', '備考']]);
  styleH(s, 6);
  s.getRange('D:D').setNumberFormat('#,##0');
}

// ── 決算_収入シート（SUMIFで集計） ──
function setupIncomeSummary(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.INCOME_SUMMARY); s.clear();
  s.getRange(1, 1, 1, 5).setValues([['科目', '予算額', '決算額', '差　異', '摘　要']]);
  styleH(s, 5);
  var items = [
    ['会費',             1100000, '', '', '年会費及入会金'],
    ['臨時会費',         200000,  '', '', '総会等各社負担分'],
    ['支部補助金',       400000,  '', '', '実務交流会等（支部より）'],
    ['イベント費',       800000,  '', '', 'トラックの日活動費等（支部より）'],
    ['安協補助金',       200000,  '', '', '交通安全啓発活動等'],
    ['雑収入',           100000,  '', '', '寸志等'],
    ['預金利息',         0,       '', '', ''],
    ['繰越金',           2598489, '', '', ''],
    ['合計',             '',      '', '', '']
  ];
  s.getRange(2, 1, items.length, 5).setValues(items);
  var lr = items.length + 1;
  var ds = SHEET_NAMES.INCOME_DETAIL;
  for (var r = 2; r < lr; r++) {
    s.getRange(r, 3).setFormula("=IFERROR(SUMIF('" + ds + "'!B:B,A" + r + ",'" + ds + "'!D:D),0)");
    s.getRange(r, 4).setFormula('=IF(C' + r + '=0,"",C' + r + '-B' + r + ')');
  }
  s.getRange(lr, 2).setFormula('=SUM(B2:B' + (lr - 1) + ')');
  s.getRange(lr, 3).setFormula('=SUM(C2:C' + (lr - 1) + ')');
  s.getRange(lr, 4).setFormula('=C' + lr + '-B' + lr);
  s.getRange(lr, 1, 1, 5).setFontWeight('bold').setBackground('#FFF2CC');
  fmtCur(s, [2, 3, 4], 2, lr);
}

// ── 決算_支出シート（SUMIFで集計） ──
function setupExpenseSummary(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.EXPENSE_SUMMARY); s.clear();
  s.getRange(1, 1, 1, 5).setValues([['科目', '予算額', '決算額', '差　異', '摘　要']]);
  styleH(s, 5);
  var items = [
    ['研修活動費',         600000,  '', '', '親睦旅行、研修会等'],
    ['事業活動費',         300000,  '', '', 'トラックの日活動等'],
    ['実務交流会活動費',   300000,  '', '', '実務交流会等'],
    ['事業費',             500000,  '', '', '交通安全啓発活動等'],
    ['会場費',             350000,  '', '', '総会、新年会等'],
    ['県青年部対応費',     400000,  '', '', '県青年部等参加費'],
    ['40周年記念積立金',   100000,  '', '', ''],
    ['通信費',             10000,   '', '', '切手、振込手数料等'],
    ['慶弔費',             50000,   '', '', '御祝、香典等'],
    ['交通費',             100000,  '', '', '役員交通費'],
    ['事務費',             240000,  '', '', '事務用品、事務手数料等'],
    ['雑費',               10000,   '', '', '役員会等'],
    ['広告活動費',         200000,  '', '', ''],
    ['予備費',             0,       '', '', ''],
    ['合計',               '',      '', '', '']
  ];
  s.getRange(2, 1, items.length, 5).setValues(items);
  var lr = items.length + 1;
  var ds = SHEET_NAMES.EXPENSE_DETAIL;
  for (var r = 2; r < lr; r++) {
    s.getRange(r, 3).setFormula("=IFERROR(SUMIF('" + ds + "'!B:B,A" + r + ",'" + ds + "'!D:D),0)");
    s.getRange(r, 4).setFormula('=IF(C' + r + '=0,"",C' + r + '-B' + r + ')');
  }
  s.getRange(lr, 2).setFormula('=SUM(B2:B' + (lr - 1) + ')');
  s.getRange(lr, 3).setFormula('=SUM(C2:C' + (lr - 1) + ')');
  s.getRange(lr, 4).setFormula('=C' + lr + '-B' + lr);
  s.getRange(lr, 1, 1, 5).setFontWeight('bold').setBackground('#FFF2CC');
  fmtCur(s, [2, 3, 4], 2, lr);
}

// ── 財産目録シート ──
function setupAssetsSheet(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.ASSETS); s.clear();
  s.getRange(1, 1, 1, 3).setValues([['科目', '金額（円）', '適　用']]);
  styleH(s, 3);
  var d = [
    ['【資産の部】', '',         ''],
    ['現金',         0,          ''],
    ['普通預金',     0,          '武蔵野銀行'],
    ['定期預金',     0,          '武蔵野銀行（40周年事業積立金）'],
    ['流動資産合計', '',         ''],
    ['資産合計',     '',         ''],
    ['【負債の部】', '',         ''],
    ['流動負債',     0,          ''],
    ['未払い金',     0,          ''],
    ['流動負債合計', '',         ''],
    ['負債合計',     '',         ''],
    ['正味資産',     '',         '']
  ];
  s.getRange(2, 1, d.length, 3).setValues(d);
  // 合計数式
  s.getRange(6, 2).setFormula('=SUM(B3:B5)');   // 流動資産合計
  s.getRange(7, 2).setFormula('=B6');             // 資産合計
  s.getRange(11, 2).setFormula('=SUM(B9:B10)');  // 流動負債合計
  s.getRange(12, 2).setFormula('=B11');           // 負債合計
  s.getRange(13, 2).setFormula('=B7-B12');        // 正味資産
  fmtCur(s, [2], 2, 13);
}

// ── 予算_収入シート ──
function setupBudgetInSheet(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.BUDGET_IN); s.clear();
  s.getRange(1, 1, 1, 5).setValues([['科目', '前年度予算額', '予算額', '増　減', '摘　要']]);
  styleH(s, 5);
  var items = [
    ['会費',       1100000, 550000,  '', '年会費及入会金'],
    ['臨時会費',   200000,  500000,  '', '総会等各社負担分'],
    ['支部補助金', 400000,  400000,  '', '実務交流会等（支部より）'],
    ['イベント費', 800000,  800000,  '', 'トラックの日活動費等（支部より）'],
    ['安協補助金', 200000,  200000,  '', '交通安全啓発活動等'],
    ['雑収入',     100000,  100000,  '', '寸志等'],
    ['預金利息',   0,       0,       '', ''],
    ['繰越金',     2598489, 5130056, '', ''],
    ['合計',       '',      '',      '', '']
  ];
  s.getRange(2, 1, items.length, 5).setValues(items);
  var lr = items.length + 1;
  for (var r = 2; r < lr; r++) { s.getRange(r, 4).setFormula('=C' + r + '-B' + r); }
  s.getRange(lr, 2).setFormula('=SUM(B2:B' + (lr - 1) + ')');
  s.getRange(lr, 3).setFormula('=SUM(C2:C' + (lr - 1) + ')');
  s.getRange(lr, 4).setFormula('=C' + lr + '-B' + lr);
  s.getRange(lr, 1, 1, 5).setFontWeight('bold').setBackground('#FFF2CC');
  fmtCur(s, [2, 3, 4], 2, lr);
}

// ── 予算_支出シート ──
function setupBudgetOutSheet(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.BUDGET_OUT); s.clear();
  s.getRange(1, 1, 1, 5).setValues([['科目', '前年度予算額', '予算額', '増　減', '摘　要']]);
  styleH(s, 5);
  var items = [
    ['研修活動費',       600000,  1000000, '', '親睦旅行、研修会等'],
    ['事業活動費',       300000,  500000,  '', 'トラックの日活動等'],
    ['実務交流活動費',   300000,  400000,  '', '実務交流会等'],
    ['事業費',           500000,  500000,  '', '交通安全啓発活動等'],
    ['会場費',           350000,  350000,  '', '総会、新年会等'],
    ['県青年部対応費',   400000,  500000,  '', '県青年部等参加費'],
    ['40周年記念積立金', 100000,  100000,  '', ''],
    ['通信費',           10000,   10000,   '', '切手、振込手数料等'],
    ['慶弔費',           50000,   50000,   '', '御祝、香典等'],
    ['交通費',           100000,  200000,  '', '役員交通費'],
    ['事務費',           240000,  400000,  '', '役員名刺代、役員会議等'],
    ['雑費',             10000,   10000,   '', ''],
    ['広告活動費',       200000,  200000,  '', ''],
    ['予備費',           2238489, 3460056, '', ''],
    ['合計',             '',      '',      '', '']
  ];
  s.getRange(2, 1, items.length, 5).setValues(items);
  var lr = items.length + 1;
  for (var r = 2; r < lr; r++) { s.getRange(r, 4).setFormula('=C' + r + '-B' + r); }
  s.getRange(lr, 2).setFormula('=SUM(B2:B' + (lr - 1) + ')');
  s.getRange(lr, 3).setFormula('=SUM(C2:C' + (lr - 1) + ')');
  s.getRange(lr, 4).setFormula('=C' + lr + '-B' + lr);
  s.getRange(lr, 1, 1, 5).setFontWeight('bold').setBackground('#FFF2CC');
  fmtCur(s, [2, 3, 4], 2, lr);
}

// ── 事業計画シート ──
function setupPlanSheet(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.PLAN); s.clear();
  s.getRange(1, 1, 1, 2).setValues([['番号', '事業計画項目']]);
  styleH(s, 2);
  var items = [
    ['1', '交通安全啓発活動'],
    ['2', 'トラックの日　ＰＲ活動'],
    ['3', '青年部会員の拡大'],
    ['4', '交流会及び研修会'],
    ['5', '親睦会及び慰労会'],
    ['6', '関東トラック協会　研修会'],
    ['7', '啓発活動反省会及び卒業式']
  ];
  s.getRange(2, 1, items.length, 2).setValues(items);
  s.setColumnWidth(1, 60); s.setColumnWidth(2, 300);
}

// ── 組織情報シート ──
function setupOrgSheet(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.ORG); s.clear();
  s.getRange(1, 1, 1, 3).setValues([['役職', '氏名', '備考（兼務等）']]);
  styleH(s, 3);
  var d = [
    ['顧問',       '石川 稔大',    ''],
    ['会長',       '直井 咲子',    ''],
    ['副会長',     '久保 真康',    '県担当兼務'],
    ['会計監査',   '佐々木 義人',  ''],
    ['会計',       '宇田川 圭太',  ''],
    ['事務局',     '高橋 晟',      ''],
    ['事務局補佐', '大脇 裕次郎',  ''],
    ['事業委員長', '新井 臣大朗',  '']
  ];
  s.getRange(2, 1, d.length, 3).setValues(d);
  s.setColumnWidth(1, 100); s.setColumnWidth(2, 120); s.setColumnWidth(3, 150);
}

// ── 会員名簿シート ──
function setupMembersSheet(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.MEMBERS); s.clear();
  s.getRange(1, 1, 1, 9).setValues([
    ['No', '役職', '氏名', '会社名', '〒', '住所', 'TEL', 'FAX', '備考']
  ]);
  styleH(s, 9);
  // サンプルデータ（役職者 + 部会員）
  var d = [
    [1,  '相談役',   '小泉 保雄',    '小泉運輸（株）',              '359-0002', '埼玉県所沢市中富1400-1',         '04(2943)4221', '04(2943)4222', ''],
    [2,  '顧問',     '石川 稔大',    '（株）石川興業運輸',           '354-0045', '埼玉県入間郡三芳町上富1995',     '049(257)4949', '049(257)4959', ''],
    [3,  '会長',     '直井 咲子',    '（株）韋駄天',                 '354-0002', '埼玉県富士見市上南畑222',        '049(268)0221', '049(268)0222', ''],
    [4,  '副会長',   '久保 真康',    '(株)ランドポート',              '354-0044', '埼玉県入間郡三芳町北永井797-3',  '049(259)3850', '049(259)7732', '県担当兼務'],
    [5,  '会計監査', '佐々木 義人',  '（有）最上運輸',               '359-0012', '埼玉県所沢市坂之下780-1',        '04(2944)5916', '04(2944)5917', ''],
    [6,  '会計',     '宇田川 圭太',  '(有)アクセル',                  '359-0045', '埼玉県入間郡三芳町上富2040-2',   '049(257)0680', '049(257)0681', ''],
    [7,  '事務局',   '高橋 晟',      '高橋運送（株）',               '359-0014', '埼玉県所沢市亀ケ谷55-1',         '04(2951)3300', '04(2951)3400', ''],
    [8,  '事業委員長','新井 臣大朗', '（株）石川興業運輸',           '354-0045', '埼玉県入間郡三芳町上富1995',     '049(257)4949', '049(257)4959', ''],
    [9,  '事務局補佐','大脇 裕次郎', '（株）コイデ運輸',              '359-1142', '埼玉県所沢市上新井1-46-1',       '04(2924)8661', '04(2928)3096', ''],
    [10, '部会員',   '早坂 幸泰',    '（株）イーエム・アイ',          '359-0023', '埼玉県所沢市東所沢和田3-14-2',   '04(2945)4843', '04(2945)4705', ''],
    [11, '部会員',   '高谷 修嗣',    'ＳＨＵ（株）',                  '354-0045', '埼玉県入間郡三芳町上富1711-27',  '049(293)6530', '049(293)6531', ''],
    [12, '部会員',   '三上 敦史',    '小泉運輸（株）',               '359-0002', '埼玉県所沢市中富1400-2',         '04(2943)4221', '04(2943)4222', ''],
    [13, '部会員',   '佐藤 孝',      '(株)旭',                        '359-0012', '埼玉県所沢市坂之下1078-1',       '04(2945)1500', '04(2944)6215', ''],
    [14, '部会員',   '関根 隆之',    'マルタケ運輸(株)',              '359-0011', '埼玉県所沢市南永井619-16',       '04(2944)3939', '04(2944)5945', ''],
    [15, '部会員',   '牛山 弘規',    '（有）オー・アンド・ユー',      '359-0011', '埼玉県所沢市南永井1005-4',       '04(2944)1351', '04(2944)9525', ''],
    [16, '部会員',   '中村 昌弘',    '（株）新興運輸',               '359-0021', '埼玉県所沢市東所沢2-36-7',       '04(2941)6267', '04(2941)6268', ''],
    [17, '部会員',   '野澤 純一',    '小泉運輸（株）',               '359-0002', '埼玉県所沢市中富1400-2',         '04(2943)4221', '04(2943)4222', ''],
    [18, '部会員',   '松井 佑介',    '（株）松井カンパニー',          '359-0011', '埼玉県所沢市南永井123-1',        '04(2944)8276', '04(2946)3444', ''],
    [19, '部会員',   '阿部 傑',      '(株)豊運輸',                    '359-0024', '埼玉県所沢市下安松901-2',        '04(2945)1861', '04(2951)1804', ''],
    [20, '部会員',   '大森 一憲',    '（株）エイチイム',              '354-0045', '埼玉県入間郡三芳町上富1818-1',   '049(265)6500', '049(265)6040', ''],
    [21, '部会員',   '長谷部 浩一',  '(有)レッカーオートルック',      '359-0013', '埼玉県所沢市城468-1',            '04(2944)3591', '04(2945)4564', ''],
    [22, '部会員',   '鈴木 一也',    '磐城土建工業(株)',              '359-0011', '埼玉県所沢市南永井1033',         '04(2944)6191', '04(2944)7439', ''],
    [23, '部会員',   '土田 武士',    '関東冷凍運輸（有）',            '359-1142', '埼玉県所沢市上新井1-36-8',       '04(2945)7377', '04(2951)6559', ''],
    [24, '部会員',   '斎藤 晃司',    '（株）石川興業運輸',           '354-0045', '埼玉県入間郡三芳町上富1995',     '049(257)4949', '049(257)4959', ''],
    [25, '部会員',   '岩﨑 鷹拓',    '（株）Ｉ・Ｌ・O',               '359-0011', '埼玉県所沢市南永井222-1',        '04(2946)3399', '04(2946)7799', '']
  ];
  s.getRange(2, 1, d.length, 9).setValues(d);
}

// ── 規約シート ──
function setupRulesSheet(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.RULES); s.clear();
  s.getRange(1, 1, 1, 3).setValues([['条番号', '見出し', '条文']]);
  styleH(s, 3);
  s.appendRow(['第１条', '名称', '埼玉県トラック協会所沢支部青年部会と称するものとする。以下本会という。']);
  s.appendRow(['第２条', '目的', '本会は、トラック事業の後継者及び幹部社員の育成と親睦をはかり、健全なる事業の発展に努め、トラック事業及び支部発展のために寄与するものである。']);
}

// ── ダッシュボードシート ──
function setupDashboard(ss) {
  var s = getOrCreate(ss, SHEET_NAMES.DASHBOARD); s.clear();
  s.getRange(1, 1).setValue('総会資料 入力ダッシュボード').setFontSize(16).setFontWeight('bold');
  s.getRange(3, 1).setValue('このスプレッドシートにデータを入力すると、総会資料HTMLが自動生成されます。').setFontSize(11);
  s.getRange(5, 1, 1, 3).setValues([['シート名', '入力内容', '備考']]);
  s.getRange(5, 1, 1, 3).setBackground('#4A86C8').setFontColor('#FFFFFF').setFontWeight('bold');
  var info = [
    ['設定',         '総会基本情報',               '年度、開催日、会場など'],
    ['活動報告',     '支部活動履歴',               '年、月日、事業名、場所、担当会社'],
    ['県青年部報告', '県青年部事業履歴',            '年、月日、事業名、場所'],
    ['収入_明細',    '収入取引明細',               '日付、科目、内容、金額'],
    ['支出_明細',    '支出取引明細',               '日付、科目、内容、金額'],
    ['決算_収入',    '収入集計（自動集計）',        'SUMIFで自動計算'],
    ['決算_支出',    '支出集計（自動集計）',        'SUMIFで自動計算'],
    ['財産目録',     '資産・負債情報',             '現金、預金残高など'],
    ['予算_収入',    '来年度収入予算',             ''],
    ['予算_支出',    '来年度支出予算',             ''],
    ['事業計画',     '来年度事業計画',             '番号と計画項目'],
    ['組織情報',     '役員・役職情報',             '役職名と氏名'],
    ['会員名簿',     '会員情報',                   'No、役職、氏名、会社名など']
  ];
  s.getRange(6, 1, info.length, 3).setValues(info);
  s.setColumnWidth(1, 120); s.setColumnWidth(2, 180); s.setColumnWidth(3, 220);
}

// ============================================================
// 集計・検証・PDF
// ============================================================
function refreshSummary() {
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('集計更新完了。', SpreadsheetApp.getUi().ButtonSet.OK);
}

function validateAllSheets() {
  var errs = runValidation(SpreadsheetApp.getActiveSpreadsheet());
  var ui = SpreadsheetApp.getUi();
  if (errs.length === 0) {
    ui.alert('OK', '入力チェック完了：問題ありません。', ui.ButtonSet.OK);
  } else {
    ui.alert('入力エラー', errs.join('\n'), ui.ButtonSet.OK);
  }
}

function runValidation(ss) {
  var errs = [];
  var st = readSettings(ss);
  if (!st['年度']) errs.push('【設定】年度が未入力です');
  if (!st['回数']) errs.push('【設定】回数が未入力です');
  if (!st['総会開催日']) errs.push('【設定】総会開催日が未入力です');
  if (!st['会計担当']) errs.push('【設定】会計担当が未入力です');
  var incSheet = ss.getSheetByName(SHEET_NAMES.INCOME_SUMMARY);
  var expSheet = ss.getSheetByName(SHEET_NAMES.EXPENSE_SUMMARY);
  if (!incSheet || incSheet.getLastRow() < 2) errs.push('【決算_収入】データがありません');
  if (!expSheet || expSheet.getLastRow() < 2) errs.push('【決算_支出】データがありません');
  return errs;
}

function generatePdf() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var st = readSettings(ss);
  var html = generateSoukaiHtml();
  // HTML→Blob経由でPDF（注: GASのHtmlService経由PDF生成）
  var blob = Utilities.newBlob(html, 'text/html', st['年度'] + '_総会資料.html');
  var file = DriveApp.createFile(blob);
  var ui = SpreadsheetApp.getUi();
  ui.alert('HTML保存完了', 'Driveに保存しました。\n印刷はブラウザのHTMLからPDF印刷を使用してください。\n' + file.getUrl(), ui.ButtonSet.OK);
}

// ============================================================
// ユーティリティ
// ============================================================

/**
 * 設定シートからkey-valueマップを返す
 * @param {Spreadsheet} ss
 * @returns {Object} settings map
 */
function readSettings(ss) {
  var s = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!s || s.getLastRow() < 2) return {};
  var d = s.getRange(2, 1, s.getLastRow() - 1, 2).getValues();
  var m = {};
  d.forEach(function(r) { if (r[0]) m[String(r[0]).trim()] = String(r[1]).trim(); });
  return m;
}

function getOrCreate(ss, name) {
  var s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  return s;
}

function styleH(s, cols) {
  s.getRange(1, 1, 1, cols)
    .setFontWeight('bold')
    .setBackground('#E8EAF6')
    .setHorizontalAlignment('center');
  s.setFrozenRows(1);
}

function fmtCur(s, cols, startRow, endRow) {
  cols.forEach(function(c) {
    s.getRange(startRow, c, endRow - startRow + 1, 1).setNumberFormat('#,##0');
  });
}

function addC(body, text, sz, bold) {
  var p = body.appendParagraph(text);
  p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  p.setFontSize(sz || 11);
  if (bold) p.setBold(true);
  return p;
}
