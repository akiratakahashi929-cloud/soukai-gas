/**
 * SoukaiHtmlGen.gs
 * ─────────────────────────────────────────────────────────────
 * スプレッドシート（SS_ID）のデータを読み込み、
 * soukai_template_demo.html の書式・CSS・レイアウトで
 * 総会資料HTMLを動的生成する。
 *
 * ※ SS_ID / SHEET_NAMES / readSettings() は Code.gs で定義済み。
 * ─────────────────────────────────────────────────────────────
 */

// ============================================================
// ユーティリティ
// ============================================================

function esc(v) {
  if (v === null || v === undefined) return '';
  return String(v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function tv(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function fmtNum(v) {
  if (v === null || v === undefined || v === '') return '';
  var s = String(v).replace(/,/g, '');
  var n = Number(s);
  if (isNaN(n)) return esc(v);
  return n.toLocaleString('ja-JP');
}

/** シートのデータ行を2次元配列で返す（ヘッダー除く） */
function getSheetRows(ss, sheetName, numCols) {
  var s = ss.getSheetByName(sheetName);
  if (!s || s.getLastRow() < 2) return [];
  var rows = Math.max(s.getLastRow() - 1, 1);
  var cols = numCols || Math.max(s.getLastColumn(), 1);
  return s.getRange(2, 1, rows, cols).getValues();
}

/** 表示値（フォーマット済み）でシートのデータ行を返す */
function getSheetDisplayRows(ss, sheetName, numCols) {
  var s = ss.getSheetByName(sheetName);
  if (!s || s.getLastRow() < 2) return [];
  var rows = Math.max(s.getLastRow() - 1, 1);
  var cols = numCols || Math.max(s.getLastColumn(), 1);
  return s.getRange(2, 1, rows, cols).getDisplayValues();
}

// ============================================================
// メイン生成関数
// ============================================================

function generateSoukaiHtml() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var st = readSettings(ss);

  // ── 年度計算 ──
  var nengo = st['年度'] || '令和7年度';
  var m = nengo.match(/令和(\d+)年度/);
  var curYear  = m ? parseInt(m[1]) : 7;
  var prevYear = curYear - 1;

  var prevNendo       = '令和' + prevYear + '年度';
  var currentNendo    = '令和' + curYear  + '年度';
  var prevStart       = st['会計年度開始'] || ('令和' + prevYear + '年4月1日');
  var prevEnd         = st['会計年度終了'] || ('令和' + curYear  + '年3月31日');
  var nextStart       = st['次年度開始']   || ('令和' + curYear  + '年4月1日');
  var nextEnd         = st['次年度終了']   || ('令和' + (curYear + 1) + '年3月31日');

  var pages = [
    buildCoverPage(st),
    buildAgendaPage(st, prevNendo, currentNendo),
    buildActivityPage(ss, prevNendo, prevStart, prevEnd),
    buildPrefecturalPage(ss, prevNendo, prevStart, prevEnd),
    buildFinancialPage(ss, st, prevNendo, prevStart, prevEnd),
    buildAssetPage(ss, st, prevNendo),
    buildBusinessPlanPage(ss, currentNendo, nextStart, nextEnd),
    buildBudgetPage(ss, currentNendo, nextStart, nextEnd),
    buildOrgPage(ss, currentNendo),
    buildMemberPage(ss),
    buildBylaws1Page(),
    buildBylaws2Page(),
    buildBylaws3Page()
  ];

  return getHtmlHead() + pages.join('\n') + '\n</body>\n</html>';
}

// ============================================================
// HTML Head + CSS（テンプレートと完全一致）
// ============================================================

function getHtmlHead() {
  return '<!DOCTYPE html>\n<html lang="ja">\n<head>\n' +
    '<meta charset="UTF-8">\n' +
    '<title>総会資料</title>\n' +
    '<style>\n' + getCSS() + '\n</style>\n' +
    '</head>\n<body>\n';
}

function getCSS() {
  return [
    "@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@400;700&family=Noto+Sans+JP:wght@400;600;700&display=swap');",
    "*{margin:0;padding:0;box-sizing:border-box}",
    "@page{size:A4 portrait;margin:20mm 18mm 20mm 22mm}",
    'body{font-family:"Noto Serif JP","MS 明朝",serif;font-size:10.5pt;color:#1a1a1a;background:#fff}',
    '.h-main{font-size:16pt;font-weight:700;font-family:"Noto Sans JP",sans-serif;text-align:center;margin-bottom:12pt}',
    '.h-title{font-size:13pt;font-weight:700;font-family:"Noto Sans JP",sans-serif}',
    '.h-sub{font-size:11pt;font-weight:700}',
    '.page{width:100%;min-height:257mm;page-break-after:always;position:relative;padding:0}',
    '.page:last-child{page-break-after:auto}',
    '@media screen{',
    '  body{background:#b0b0b0}',
    '  .page{background:#fff;max-width:210mm;margin:20px auto;padding:20mm 18mm 20mm 22mm;box-shadow:0 3px 16px rgba(0,0,0,.25);border-radius:2px}',
    '}',
    '.page-no{position:absolute;bottom:0;right:0;font-size:9pt;color:#666}',
    'table{width:100%;border-collapse:collapse;font-size:9.5pt;margin-bottom:10pt}',
    'th,td{border:1px solid #333;padding:4pt 6pt;vertical-align:middle}',
    'th{background:#E8EAF6;font-family:"Noto Sans JP",sans-serif;font-weight:700;font-size:9pt;text-align:center}',
    '.tr{text-align:right}.tc{text-align:center}',
    '.row-total{font-weight:700;border-top:2px solid #333;background:#FFF8E1}',
    '.row-section{font-weight:700;background:#E3F2FD}',
    '.tbl-sm td,.tbl-sm th{font-size:8pt;padding:2pt 4pt}',
    '.cover-page{display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;min-height:257mm}',
    '.cover-info{font-size:11pt;margin-bottom:5pt}',
    '.cover-title{font-size:24pt;font-weight:700;letter-spacing:.4em;margin:50pt 0 14pt;font-family:"Noto Serif JP",serif}',
    '.cover-period{font-size:12pt;margin-bottom:70pt;color:#333}',
    '.cover-org{font-size:14pt;letter-spacing:.6em;margin-bottom:5pt}',
    '.agenda-wrap{display:flex;flex-direction:column;justify-content:center;align-items:center;min-height:215mm}',
    '.agenda-inner{width:360pt}',
    '.agenda-title{font-size:18pt;font-weight:700;letter-spacing:.5em;text-align:center;margin-bottom:32pt;font-family:"Noto Sans JP",sans-serif}',
    '.ag-item{font-size:13pt;font-weight:700;margin-bottom:14pt;display:flex;gap:16pt;align-items:flex-start}',
    '.ag-num{min-width:28pt;white-space:nowrap}',
    '.ag-sub{padding-left:44pt;margin-top:8pt}',
    '.ag-subitem{font-size:12pt;font-weight:400;margin-bottom:9pt;display:flex;gap:12pt}',
    '.ag-subnum{min-width:72pt;color:#555}',
    '.plan-wrap{display:flex;flex-direction:column;justify-content:center;align-items:center;min-height:210mm}',
    '.plan-inner{width:320pt}',
    '.plan-section{font-size:14pt;font-weight:700;margin-bottom:24pt;font-family:"Noto Sans JP",sans-serif}',
    '.plan-item{font-size:13pt;margin-bottom:18pt;display:flex;gap:12pt;align-items:flex-start}',
    '.plan-num{min-width:32pt;color:#333;flex-shrink:0}',
    '.rule-block{display:flex;gap:12pt;margin-bottom:7pt;font-size:10pt;line-height:1.9;align-items:flex-start}',
    '.rule-no{font-weight:700;white-space:nowrap;min-width:52pt}',
    '.rule-body{flex:1}',
    '.rule-heading{font-size:10.5pt;font-weight:700;margin:16pt 0 4pt;color:#1a237e;border-left:4px solid #1a237e;padding-left:6pt;}',
    '.sig{margin-top:14pt;font-size:9.5pt;line-height:2.1}',
    '.period-note{text-align:right;font-size:9pt;color:#555;margin-bottom:7pt}'
  ].join('\n');
}

// ============================================================
// 表紙
// ============================================================

function buildCoverPage(st) {
  var kaisaiDate = esc(st['総会開催日'] || '');
  var venue      = esc(st['会場'] || '');
  var kaijTime   = esc(st['開会時刻'] || '');
  var kaisu      = esc(st['回数'] || '');
  var prevStart  = esc(st['会計年度開始'] || '');
  var prevEnd    = esc(st['会計年度終了'] || '');

  // 回数を全角に変換して見出し文字列を生成
  var titleText  = '第　' + kaisu + '　回　定　期　総　会';

  return [
    '<div class="page cover-page">',
    '  <div>',
    '    <div class="cover-info">' + kaisaiDate + '</div>',
    '    <div class="cover-info">於　' + venue + '</div>',
    '    <div class="cover-info">' + kaijTime + '　開会</div>',
    '    <div class="cover-title">' + titleText + '</div>',
    '    <div class="cover-period">（ ' + prevStart + ' 〜 ' + prevEnd + ' ）</div>',
    '    <div class="cover-org">埼 玉 県 ト ラ ッ ク 協 会</div>',
    '    <div class="cover-org">所 沢 支 部 青 年 部 会</div>',
    '  </div>',
    '</div>'
  ].join('\n');
}

// ============================================================
// P1 総会次第
// ============================================================

function buildAgendaPage(st, prevNendo, currentNendo) {
  var proposals = [
    { num: '第１号議案', text: prevNendo + '　事業報告' },
    { num: '第２号議案', text: prevNendo + '　決算報告' },
    { num: '第３号議案', text: currentNendo + '　事業計画（案）' },
    { num: '第４号議案', text: currentNendo + '　予算案' }
  ];

  var subItems = proposals.map(function(p) {
    return '          <div class="ag-subitem"><span class="ag-subnum">' +
      esc(p.num) + '</span><span>' + esc(p.text) + '</span></div>';
  }).join('\n');

  return [
    '<div class="page">',
    '  <div class="agenda-wrap">',
    '    <div class="agenda-inner">',
    '      <div class="agenda-title">総　会　次　第</div>',
    '      <div class="ag-item"><span class="ag-num">１、</span><span>開会の辞</span></div>',
    '      <div class="ag-item"><span class="ag-num">２、</span><span>支部長挨拶</span></div>',
    '      <div class="ag-item"><span class="ag-num">３、</span><span>議長選出</span></div>',
    '      <div class="ag-item"><span class="ag-num">４、</span><span>定足数確認</span></div>',
    '      <div class="ag-item"><span class="ag-num">５、</span><span>議　　事',
    '        <div class="ag-sub">',
    subItems,
    '        </div>',
    '      </span></div>',
    '      <div class="ag-item"><span class="ag-num">６、</span><span>閉会の辞</span></div>',
    '    </div>',
    '  </div>',
    '  <div class="page-no">P1</div>',
    '</div>'
  ].join('\n');
}

// ============================================================
// P2 活動報告
// ============================================================

function buildActivityPage(ss, prevNendo, prevStart, prevEnd) {
  // 列: 0=年, 1=月日, 2=支部事業名, 3=場所, 4=担当会社, 5=県事業名, 6=県場所
  var rows = getSheetRows(ss, SHEET_NAMES.ACTIVITY, 7);

  var tableRows = '';
  var lastYear  = '';
  rows.forEach(function(r) {
    var year    = tv(r[0]);
    var day     = tv(r[1]);
    var jigyou  = tv(r[2]);
    var basho   = tv(r[3]);
    var tanto   = tv(r[4]);

    if (!day && !jigyou) return; // 空行スキップ

    if (year) lastYear = year;
    var dateStr = lastYear ? lastYear + (day ? day : '') : day;

    tableRows +=
      '<tr>' +
      '<td class="tc">' + esc(dateStr) + '</td>' +
      '<td>' + esc(jigyou) + '</td>' +
      '<td>' + esc(basho)  + '</td>' +
      '<td>' + esc(tanto)  + '</td>' +
      '</tr>\n';
  });

  if (!tableRows) tableRows = '<tr><td colspan="4" class="tc">データなし</td></tr>\n';

  return [
    '<div class="page">',
    '  <div class="h-sub" style="text-align:center;margin-bottom:2pt">埼玉県トラック協会 所沢支部 青年部会</div>',
    '  <div class="h-main">' + esc(prevNendo) + '　活動内容</div>',
    '  <div class="period-note">' + esc(prevStart) + ' 〜 ' + esc(prevEnd) + '</div>',
    '  <table>',
    '    <thead><tr>',
    '      <th style="width:90pt">月　日</th>',
    '      <th>支部事業名</th>',
    '      <th style="width:70pt">場所</th>',
    '      <th style="width:90pt">担当会社</th>',
    '    </tr></thead>',
    '    <tbody>',
    tableRows,
    '    </tbody>',
    '  </table>',
    '  <div class="page-no">P2</div>',
    '</div>'
  ].join('\n');
}

// ============================================================
// P3 県青年部事業報告
// ============================================================

function buildPrefecturalPage(ss, prevNendo, prevStart, prevEnd) {
  // 列: 0=年, 1=月日, 2=事業名, 3=場所
  var rows = getSheetRows(ss, SHEET_NAMES.PREFECTURAL, 4);

  var tableRows = '';
  var lastYear  = '';
  rows.forEach(function(r) {
    var year   = tv(r[0]);
    var day    = tv(r[1]);
    var jigyou = tv(r[2]);
    var basho  = tv(r[3]);

    if (!day && !jigyou) return;

    if (year) lastYear = year;
    var dateStr = lastYear ? lastYear + (day ? day : '') : day;

    tableRows +=
      '<tr>' +
      '<td class="tc">' + esc(dateStr) + '</td>' +
      '<td>' + esc(jigyou) + '</td>' +
      '<td>' + esc(basho)  + '</td>' +
      '</tr>\n';
  });

  if (!tableRows) tableRows = '<tr><td colspan="3" class="tc">データなし</td></tr>\n';

  return [
    '<div class="page">',
    '  <div class="h-sub" style="text-align:center;margin-bottom:2pt">埼玉県トラック協会 所沢支部 青年部会</div>',
    '  <div class="h-main">' + esc(prevNendo) + '　県青年部会事業報告</div>',
    '  <div class="period-note">' + esc(prevStart) + ' 〜 ' + esc(prevEnd) + '</div>',
    '  <table>',
    '    <thead><tr>',
    '      <th style="width:90pt">月　日</th>',
    '      <th>事　業　名</th>',
    '      <th style="width:120pt">場　所</th>',
    '    </tr></thead>',
    '    <tbody>',
    tableRows,
    '    </tbody>',
    '  </table>',
    '  <div class="page-no">P3</div>',
    '</div>'
  ].join('\n');
}

// ============================================================
// P4 収支決算報告書
// ============================================================

function buildFinancialPage(ss, st, prevNendo, prevStart, prevEnd) {
  // 決算_収入: 科目, 予算額, 決算額, 差異, 摘要
  var incRows = getSheetDisplayRows(ss, SHEET_NAMES.INCOME_SUMMARY, 5);
  // 決算_支出: 科目, 予算額, 決算額, 差異, 摘要
  var expRows = getSheetDisplayRows(ss, SHEET_NAMES.EXPENSE_SUMMARY, 5);

  // 収入テーブル
  var incHtml = buildFinTable(incRows);

  // 支出テーブル
  var expHtml = buildFinTable(expRows);

  // 収支差額行（収入合計 - 支出合計）
  var incTotalRow = findTotalRow(incRows);
  var expTotalRow = findTotalRow(expRows);
  var netBudget   = calcNet(incTotalRow, 1, expTotalRow, 1);
  var netActual   = calcNet(incTotalRow, 2, expTotalRow, 2);
  var netDiff     = calcNet(incTotalRow, 3, expTotalRow, 3);

  var sigDate    = esc(st['報告日'] || '');
  var accountant = esc(st['会計担当'] || '');
  var auditor    = esc(st['監査担当'] || '');
  var auditDate  = esc(st['監査日'] || '');

  return [
    '<div class="page">',
    '  <div class="h-main">' + esc(prevNendo) + '　青年部収支決算報告書</div>',
    '  <div class="h-title" style="margin:4pt 0 3pt">【収入の部】</div>',
    '  <div class="period-note">' + esc(prevStart) + ' 〜 ' + esc(prevEnd) + '</div>',
    '  <table>',
    '    <thead><tr>',
    '      <th>科　目</th>',
    '      <th style="width:80pt">予算額（円）</th>',
    '      <th style="width:80pt">決算額（円）</th>',
    '      <th style="width:72pt">差　額（円）</th>',
    '      <th>備考</th>',
    '    </tr></thead>',
    '    <tbody>' + incHtml + '</tbody>',
    '  </table>',
    '  <div class="h-title" style="margin:4pt 0 3pt">【支出の部】</div>',
    '  <table>',
    '    <thead><tr>',
    '      <th>科　目</th>',
    '      <th style="width:80pt">予算額（円）</th>',
    '      <th style="width:80pt">決算額（円）</th>',
    '      <th style="width:72pt">差　額（円）</th>',
    '      <th>備考</th>',
    '    </tr></thead>',
    '    <tbody>' + expHtml + '</tbody>',
    '  </table>',
    '  <table style="margin-top:6pt">',
    '    <tbody>',
    '      <tr class="row-total" style="background:#E8F5E9">',
    '        <td>収支差額（次年度繰越金）</td>',
    '        <td class="tr" style="width:80pt">' + esc(netBudget)  + '</td>',
    '        <td class="tr" style="width:80pt">' + esc(netActual)  + '</td>',
    '        <td class="tr" style="width:72pt">' + esc(netDiff)    + '</td>',
    '        <td></td>',
    '      </tr>',
    '    </tbody>',
    '  </table>',
    '  <div class="sig">',
    '    上記の通りご報告いたします。<br>',
    '    ' + sigDate + '　　　　会計　' + accountant + '<br><br>',
    '    上記各項について監査した結果正確であることを認めます。<br>',
    '    ' + auditDate + '　　　　会計監査　' + auditor,
    '  </div>',
    '  <div class="page-no">P4</div>',
    '</div>'
  ].join('\n');
}

/** 決算テーブルのtbody行HTMLを生成 */
function buildFinTable(rows) {
  var html = '';
  rows.forEach(function(r, i) {
    var label  = tv(r[0]);
    var budget = tv(r[1]);
    var actual = tv(r[2]);
    var diff   = tv(r[3]);
    var note   = tv(r[4]);
    if (!label) return;
    var isTotal = (label === '合計' || i === rows.length - 1);
    var cls = isTotal ? ' class="row-total"' : '';
    html +=
      '<tr' + cls + '>' +
      '<td>' + esc(label)  + '</td>' +
      '<td class="tr">' + esc(budget) + '</td>' +
      '<td class="tr">' + esc(actual) + '</td>' +
      '<td class="tr">' + esc(diff)   + '</td>' +
      '<td>' + esc(note)   + '</td>' +
      '</tr>\n';
  });
  return html;
}

/** 合計行を探す（'合計'ラベルか最終行） */
function findTotalRow(rows) {
  for (var i = 0; i < rows.length; i++) {
    if (tv(rows[i][0]) === '合計') return rows[i];
  }
  return rows.length > 0 ? rows[rows.length - 1] : null;
}

/** 表示値の差額を計算 */
function calcNet(incRow, col, expRow, col2) {
  if (!incRow || !expRow) return '';
  var inc = Number(String(tv(incRow[col])).replace(/,/g, ''));
  var exp = Number(String(tv(expRow[col2 !== undefined ? col2 : col])).replace(/,/g, ''));
  if (isNaN(inc) || isNaN(exp)) return '';
  return (inc - exp).toLocaleString('ja-JP');
}

// ============================================================
// P5 財産目録
// ============================================================

function buildAssetPage(ss, st, prevNendo) {
  // 列: 0=科目, 1=金額, 2=適用
  var rows = getSheetDisplayRows(ss, SHEET_NAMES.ASSETS, 3);

  var tableRows = '';
  rows.forEach(function(r) {
    var label = tv(r[0]);
    var kin   = tv(r[1]);
    var tekiy = tv(r[2]);
    if (!label && !kin) return;

    var isSection = label.indexOf('【') === 0 || label.indexOf('部】') > -1;
    var isTotal   = label.indexOf('合計') > -1 || label === '正味資産';
    var cls = isSection ? 'class="row-section"' : (isTotal ? 'class="row-total"' : '');

    tableRows +=
      '<tr ' + cls + '>' +
      '<td>' + esc(label) + '</td>' +
      '<td class="tr">' + esc(kin) + '</td>' +
      '<td>' + esc(tekiy) + '</td>' +
      '</tr>\n';
  });

  if (!tableRows) tableRows = '<tr><td colspan="3" class="tc">データなし</td></tr>\n';

  var asOfDate   = esc(st['会計年度終了'] || '');
  var sigDate    = esc(st['報告日'] || '');
  var accountant = esc(st['会計担当'] || '');
  var auditor    = esc(st['監査担当'] || '');
  var auditDate  = esc(st['監査日'] || '');

  return [
    '<div class="page">',
    '  <div class="h-main">財　産　目　録</div>',
    '  <div class="period-note">' + asOfDate + '現在</div>',
    '  <table>',
    '    <thead><tr>',
    '      <th>科　目</th>',
    '      <th style="width:120pt">金額（円）</th>',
    '      <th>適　用</th>',
    '    </tr></thead>',
    '    <tbody>',
    tableRows,
    '    </tbody>',
    '  </table>',
    '  <div class="sig">',
    '    上記の通りご報告いたします。<br>',
    '    ' + sigDate + '　　会計　' + accountant + '<br><br>',
    '    上記各項について監査した結果正確であることを認めます。<br>',
    '    ' + auditDate + '　　会計監査　' + auditor,
    '  </div>',
    '  <div class="page-no">P5</div>',
    '</div>'
  ].join('\n');
}

// ============================================================
// P6 事業計画（案）
// ============================================================

function buildBusinessPlanPage(ss, currentNendo, nextStart, nextEnd) {
  // 列: 0=番号, 1=事業計画項目
  var rows = getSheetRows(ss, SHEET_NAMES.PLAN, 2);

  var items = '';
  var nums = ['（１）','（２）','（３）','（４）','（５）','（６）','（７）','（８）','（９）','（１０）'];
  rows.forEach(function(r, i) {
    var text = tv(r[1]);
    if (!text) return;
    var numLabel = nums[i] || ('（' + (i + 1) + '）');
    items +=
      '<div class="plan-item">' +
      '<span class="plan-num">' + esc(numLabel) + '</span>' +
      '<span>' + esc(text) + '</span>' +
      '</div>\n';
  });

  if (!items) items = '<div class="plan-item"><span>データなし</span></div>';

  return [
    '<div class="page">',
    '  <div class="h-main">' + esc(currentNendo) + '　事業計画（案）</div>',
    '  <div class="period-note">' + esc(nextStart) + ' 〜 ' + esc(nextEnd) + '</div>',
    '  <div class="plan-wrap">',
    '    <div class="plan-inner">',
    '      <div class="plan-section">１・事業計画</div>',
    items,
    '    </div>',
    '  </div>',
    '  <div class="page-no">P6</div>',
    '</div>'
  ].join('\n');
}

// ============================================================
// P7 収支予算（案）
// ============================================================

function buildBudgetPage(ss, currentNendo, nextStart, nextEnd) {
  // 予算_収入: 科目, 前年度予算額, 予算額, 増減, 摘要
  var incRows = getSheetDisplayRows(ss, SHEET_NAMES.BUDGET_IN, 5);
  // 予算_支出: 科目, 前年度予算額, 予算額, 増減, 摘要
  var expRows = getSheetDisplayRows(ss, SHEET_NAMES.BUDGET_OUT, 5);

  var incHtml = buildBudgetTable(incRows);
  var expHtml = buildBudgetTable(expRows);

  return [
    '<div class="page">',
    '  <div class="h-main">' + esc(currentNendo) + '　青年部収支予算（案）</div>',
    '  <div class="h-title" style="margin:4pt 0 3pt">【収入の部】</div>',
    '  <div class="period-note">' + esc(nextStart) + ' 〜 ' + esc(nextEnd) + '</div>',
    '  <table>',
    '    <thead><tr>',
    '      <th>科　目</th>',
    '      <th style="width:80pt">前年度予算額</th>',
    '      <th style="width:80pt">予算額（円）</th>',
    '      <th style="width:72pt">増　減</th>',
    '      <th>摘　要</th>',
    '    </tr></thead>',
    '    <tbody>' + incHtml + '</tbody>',
    '  </table>',
    '  <div class="h-title" style="margin:4pt 0 3pt">【支出の部】</div>',
    '  <table>',
    '    <thead><tr>',
    '      <th>科　目</th>',
    '      <th style="width:80pt">前年度予算額</th>',
    '      <th style="width:80pt">予算額（円）</th>',
    '      <th style="width:72pt">増　減</th>',
    '      <th>摘　要</th>',
    '    </tr></thead>',
    '    <tbody>' + expHtml + '</tbody>',
    '  </table>',
    '  <div class="page-no">P7</div>',
    '</div>'
  ].join('\n');
}

function buildBudgetTable(rows) {
  var html = '';
  rows.forEach(function(r, i) {
    var label   = tv(r[0]);
    var prevBud = tv(r[1]);
    var budget  = tv(r[2]);
    var diff    = tv(r[3]);
    var note    = tv(r[4]);
    if (!label) return;
    var isTotal = (label === '合計' || i === rows.length - 1);
    var cls = isTotal ? ' class="row-total"' : '';
    html +=
      '<tr' + cls + '>' +
      '<td>' + esc(label)   + '</td>' +
      '<td class="tr">' + esc(prevBud) + '</td>' +
      '<td class="tr">' + esc(budget)  + '</td>' +
      '<td class="tr">' + esc(diff)    + '</td>' +
      '<td>' + esc(note)    + '</td>' +
      '</tr>\n';
  });
  return html;
}

// ============================================================
// P8 組織図
// ============================================================

function buildOrgPage(ss, currentNendo) {
  // 組織情報: 役職, 氏名, 備考
  var orgRows = getSheetRows(ss, SHEET_NAMES.ORG, 3);
  var org = {};
  orgRows.forEach(function(r) {
    var role = tv(r[0]);
    if (role) org[role] = { name: tv(r[1]), note: tv(r[2]) };
  });

  // 会員名簿から部会員を取得
  var memRows = getSheetRows(ss, SHEET_NAMES.MEMBERS, 3);
  var fukaiin = [];
  memRows.forEach(function(r) {
    if (tv(r[1]) === '部会員' && tv(r[2])) fukaiin.push(tv(r[2]));
  });

  var get = function(role) { return org[role] ? org[role].name : ''; };
  var getNote = function(role) { return org[role] ? org[role].note : ''; };

  var advisorName      = get('顧問');
  var presidentName    = get('会長');
  var vicePresName     = get('副会長');
  var vicePresNote     = getNote('副会長');
  var auditorName      = get('会計監査');
  var accountantName   = get('会計');
  var officeName       = get('事務局');
  var subOfficeName    = get('事務局補佐');
  var committeeName    = get('事業委員長');

  var advisorBlock = '';
  if (advisorName) {
    advisorBlock = [
      '<div style="position:relative;width:100%;height:32px;display:flex;justify-content:center">',
      '  <div style="width:1.5px;height:100%;background:#333"></div>',
      '  <div style="position:absolute;top:16px;left:50%;width:100px;height:1.5px;background:#333"></div>',
      '  <div style="position:absolute;top:4px;left:calc(50% + 100px);border:1px solid #880e4f;padding:2px 10px;font-size:9pt;background:#fce4ec;text-align:center;white-space:nowrap">',
      '    顧　問<br><span style="font-weight:400">' + esc(advisorName) + '</span>',
      '  </div>',
      '</div>'
    ].join('\n');
  } else {
    advisorBlock = '<div style="width:1.5px;height:32px;background:#333;margin:0 auto"></div>';
  }

  var viceLabel = vicePresNote ? esc(vicePresNote) : '副会長';
  var committeeBlock = '';
  if (committeeName) {
    committeeBlock = [
      '<div style="width:1.5px;height:18px;background:#333;margin:0 auto"></div>',
      '<div style="border:1px solid #333;background:#fff;padding:4px 20px;font-weight:700;font-size:9.5pt;text-align:center;display:inline-block">',
      '  事業委員長<br><span style="font-weight:400;font-size:8.5pt;margin-top:3px;display:block">' + esc(committeeName) + '</span>',
      '</div>'
    ].join('\n');
  }

  var membersGrid = fukaiin.map(function(name) {
    return '<div style="border:1px solid #666;padding:4px;text-align:center;font-size:8.5pt">' + esc(name) + '</div>';
  }).join('\n');

  return [
    '<div class="page">',
    '  <div class="h-main">' + esc(currentNendo) + '　青年部組織図</div>',
    '  <div style="display:flex;flex-direction:column;align-items:center;font-family:\'Noto Sans JP\',sans-serif;font-size:9pt;padding:2pt 0">',

    '    <!-- 総会 -->',
    '    <div style="border:1.5px solid #111;padding:4px 36px;font-size:10.5pt;font-weight:700;letter-spacing:.25em;background:#fff">総　会</div>',

    advisorBlock,

    '    <!-- 会長 -->',
    '    <div style="border:1.5px solid #111;padding:4px 28px;font-size:10.5pt;font-weight:700;background:#fff">',
    '      会　長　' + esc(presidentName),
    '    </div>',

    '    <div style="width:1.5px;height:24px;background:#333;margin:0 auto"></div>',

    '    <!-- 役員会ブロック（十字ハブ） -->',
    '    <div style="position:relative;width:420px;text-align:center;padding:26px 0;margin-bottom:2px">',
    '      <div style="position:absolute;top:0;bottom:0;left:50%;width:1.5px;background:#333;margin-left:-0.75px;z-index:1"></div>',
    '      <div style="position:absolute;top:50%;left:50px;right:50px;height:1.5px;background:#333;margin-top:-0.75px;z-index:1"></div>',
    '      <div style="position:absolute;top:0;bottom:0;left:50px;width:1.5px;background:#333;margin-left:-0.75px;z-index:1"></div>',
    '      <div style="position:absolute;top:0;bottom:0;right:50px;width:1.5px;background:#333;margin-right:-0.75px;z-index:1"></div>',
    '      <!-- 中央: 役員会 -->',
    '      <div style="position:relative;z-index:2;display:inline-block;border:1.5px solid #111;background:#fff;padding:6px 20px;font-weight:700;letter-spacing:.1em;font-size:10pt">役員会</div>',
    '      <!-- 左上: 事務局 -->',
    '      <div style="position:absolute;top:-10px;left:50px;transform:translateX(-50%);z-index:2;border:1px solid #333;background:#fff;padding:4px 12px;text-align:center;min-width:105px;font-weight:700;font-size:9.5pt">',
    '        事務局',
    '        <div style="font-weight:400;font-size:8.5pt;margin-top:3px">' + esc(officeName) + '</div>',
    '      </div>',
    '      <!-- 左下: 事務局補佐 -->',
    '      <div style="position:absolute;bottom:-10px;left:50px;transform:translateX(-50%);z-index:2;border:1px solid #333;background:#fff;padding:4px 12px;text-align:center;min-width:105px;font-weight:700;font-size:9.5pt">',
    '        事務局補佐',
    '        <div style="font-weight:400;font-size:8.5pt;margin-top:3px">' + esc(subOfficeName) + '</div>',
    '      </div>',
    '      <!-- 右上: 会計監査 -->',
    '      <div style="position:absolute;top:-10px;right:50px;transform:translateX(50%);z-index:2;border:1px solid #333;background:#fff;padding:4px 12px;text-align:center;min-width:105px;font-weight:700;font-size:9.5pt">',
    '        会計監査',
    '        <div style="font-weight:400;font-size:8.5pt;margin-top:3px">' + esc(auditorName) + '</div>',
    '      </div>',
    '      <!-- 右下: 会計 -->',
    '      <div style="position:absolute;bottom:-10px;right:50px;transform:translateX(50%);z-index:2;border:1px solid #333;background:#fff;padding:4px 12px;text-align:center;min-width:105px;font-weight:700;font-size:9.5pt">',
    '        会　計',
    '        <div style="font-weight:400;font-size:8.5pt;margin-top:3px">' + esc(accountantName) + '</div>',
    '      </div>',
    '    </div>',

    '    <div style="width:1.5px;height:24px;background:#333;margin:0 auto"></div>',

    '    <!-- 副会長 -->',
    '    <div style="border:1px solid #333;background:#fff;padding:4px 16px;font-weight:700;font-size:9.5pt;text-align:center">',
    '      ' + viceLabel,
    '      <div style="font-weight:400;font-size:8.5pt;margin-top:3px">' + esc(vicePresName) + '</div>',
    '    </div>',

    committeeBlock,

    '    <!-- 部会員 -->',
    '    <div style="width:100%;border-top:1.5px solid #ccc;margin-top:16px;padding-top:10px">',
    '      <div style="text-align:center;font-size:10pt;font-weight:700;margin-bottom:8px;letter-spacing:.2em">部　会　員</div>',
    '      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:4px">',
    membersGrid || '<div style="grid-column:1/-1;text-align:center;font-size:9pt">部会員データなし</div>',
    '      </div>',
    '    </div>',

    '  </div>',
    '  <div class="page-no">P8</div>',
    '</div>'
  ].join('\n');
}

// ============================================================
// P9-10 会員名簿
// ============================================================

function buildMemberPage(ss) {
  // 列: 0=No, 1=役職, 2=氏名, 3=会社名, 4=〒, 5=住所, 6=TEL, 7=FAX, 8=備考
  var rows = getSheetRows(ss, SHEET_NAMES.MEMBERS, 9);

  var tableRows = '';
  rows.forEach(function(r) {
    var no      = tv(r[0]);
    var role    = tv(r[1]);
    var name    = tv(r[2]);
    var company = tv(r[3]);
    var zip     = tv(r[4]);
    var addr    = tv(r[5]);
    var tel     = tv(r[6]);

    if (!no && !name) return;

    var fullAddr = zip ? '〒' + zip + '　' + addr : addr;

    tableRows +=
      '<tr>' +
      '<td class="tc">' + esc(no)      + '</td>' +
      '<td class="tc">' + esc(role)    + '</td>' +
      '<td>' + esc(name)    + '</td>' +
      '<td>' + esc(company) + '</td>' +
      '<td>' + esc(fullAddr)+ '</td>' +
      '<td>' + esc(tel)     + '</td>' +
      '</tr>\n';
  });

  if (!tableRows) tableRows = '<tr><td colspan="6" class="tc">データなし</td></tr>\n';

  return [
    '<div class="page">',
    '  <div class="h-main">会員名簿</div>',
    '  <table class="tbl-sm">',
    '    <thead><tr>',
    '      <th style="width:26pt">No</th>',
    '      <th style="width:56pt">役職</th>',
    '      <th style="width:66pt">氏名</th>',
    '      <th>会社名</th>',
    '      <th>住　所</th>',
    '      <th style="width:80pt">TEL</th>',
    '    </tr></thead>',
    '    <tbody>',
    tableRows,
    '    </tbody>',
    '  </table>',
    '  <div class="page-no">P9-10</div>',
    '</div>'
  ].join('\n');
}

// ============================================================
// P11 会則①（静的）
// ============================================================

function buildBylaws1Page() {
  return [
    '<div class="page">',
    '  <div class="h-main" style="margin-bottom:2pt">埼玉県トラック協会所沢支部青年部会</div>',
    '  <div class="h-main">会　　　則</div>',
    '  <div style="font-size:10pt;text-align:center;color:#555;margin-bottom:14pt">令和7年度</div>',
    '  <hr style="border:none;border-top:1px solid #999;margin-bottom:12pt">',
    '  <div class="rule-heading">名称</div>',
    '  <div class="rule-block"><span class="rule-no">第１条</span><span class="rule-body">埼玉県トラック協会所沢支部青年部会と称するものとする。以下本会という。</span></div>',
    '  <div class="rule-heading">目的</div>',
    '  <div class="rule-block"><span class="rule-no">第２条</span><span class="rule-body">本会は、トラック事業の後継者及び幹部社員の育成と親睦をはかり、健全なる事業の発展に努め、トラック事業及び支部発展のために寄与するものである。</span></div>',
    '  <div class="rule-heading">会員</div>',
    '  <div class="rule-block"><span class="rule-no">第３条</span><span class="rule-body">本会は埼玉県トラック協会所沢支部の事業所（原則として支部在籍６ヶ月以上）の後継者又は事業所責任者によって組織し運営する。入会は各事業所１名以上若干名とする。又本会の趣旨に賛同する法人又は個人を協賛会員とする事が出来る。</span></div>',
    '  <div class="rule-heading">限定</div>',
    '  <div class="rule-block"><span class="rule-no">第４条</span><span class="rule-body">会員の年齢は、満２０歳より原則５０歳までとする。</span></div>',
    '  <div class="rule-heading">会費</div>',
    '  <div class="rule-block"><span class="rule-no">第５条</span><span class="rule-body">会費は月々２千円とする。但し各社２人目以降は半額とする。その他に必要に応じ特別会費を徴収する事が出来る。</span></div>',
    '  <div class="rule-heading">役員</div>',
    '  <div class="rule-block"><span class="rule-no">第６条</span><span class="rule-body">本会は次の役を置く。（１）会長１名　（２）副会長若干名　（３）県青年部担当１名　（４）委員長若干名　（５）事務局１名　（６）会計１名　（７）会計監査１名</span></div>',
    '  <div class="rule-heading">役員の選任</div>',
    '  <div class="rule-block"><span class="rule-no">第７条</span><span class="rule-body">１，本会の会長は、総会において３分の２以上の多数をもって選任される。２，会長以下役員は、会長が指名する。</span></div>',
    '  <div class="rule-heading">役員の任期</div>',
    '  <div class="rule-block"><span class="rule-no">第８条</span><span class="rule-body">１，役員の任期は２年とする。但し再任を妨げない。２，補充された役員は前任者の残任期間とする。３，役員は任期満了でも後任者の就任があるまでは、その職務を行うものとする。</span></div>',
    '  <div class="rule-heading">顧問</div>',
    '  <div class="rule-block"><span class="rule-no">第９条</span><span class="rule-body">１，本会に顧問を若干名置く事が出来る。２，顧問は役員会において選任する。３，顧問は会長の要請により各議会に出席し本会の重要事項について意見を述べる事が出来る。</span></div>',
    '  <div class="rule-heading">会議</div>',
    '  <div class="rule-block"><span class="rule-no">第１０条</span><span class="rule-body">定例会議は３ヶ月に１回とする。</span></div>',
    '  <div class="page-no">P11</div>',
    '</div>'
  ].join('\n');
}

// ============================================================
// P12 会則②（静的）
// ============================================================

function buildBylaws2Page() {
  return [
    '<div class="page">',
    '  <div class="h-main" style="margin-bottom:14pt">会則（続き）</div>',
    '  <div class="rule-heading">総会</div>',
    '  <div class="rule-block"><span class="rule-no">第１１条</span><span class="rule-body">１，通常総会は年１回とし、下記項目を議題とする。（１）事業報告及び会計収支決算報告　（２）事業案及び予算案の審議　（３）本会則の変更　（４）その他の重要事項　２，臨時総会は役員会において必要と認めた時に召集する。</span></div>',
    '  <div class="rule-heading">召集</div>',
    '  <div class="rule-block"><span class="rule-no">第１２条</span><span class="rule-body">総会は会長が召集する。会長が欠席の場合は会長の要請により副会長が召集する。</span></div>',
    '  <div class="rule-heading">定例会</div>',
    '  <div class="rule-block"><span class="rule-no">第１３条</span><span class="rule-body">定例会はその都度、議題を持ちよって討議する。</span></div>',
    '  <div class="rule-heading">役員会</div>',
    '  <div class="rule-block"><span class="rule-no">第１４条</span><span class="rule-body">役員会は必要に応じ会長が召集し、次の事項を審議する。（１）総会に提出すべき議案　（２）重要事項及び事業計画　（３）総会を招集する緊急議案　（４）その他会長が必要と認めた事項</span></div>',
    '  <div class="rule-heading">議決</div>',
    '  <div class="rule-block"><span class="rule-no">第１５条</span><span class="rule-body">議決は過半数以上の会員の出席により運営される。また、出席会員の過半数の賛同を得て議決する。</span></div>',
    '  <div class="rule-heading">年度</div>',
    '  <div class="rule-block"><span class="rule-no">第１６条</span><span class="rule-body">本会の会計年度は４月１日に始まり３月３１日に終わる。</span></div>',
    '  <div class="rule-heading">除名</div>',
    '  <div class="rule-block"><span class="rule-no">第１７条</span><span class="rule-body">１，秩序を乱し本会の目的に反する行為を行なった者に対しては、役員会に於いて除名する。２，通年を通して、青年部活動に対して参加意欲に欠く者に対しては、役員会に於いて除名する事が出来る。</span></div>',
    '  <div class="rule-heading">脱会</div>',
    '  <div class="rule-block"><span class="rule-no">第１８条</span><span class="rule-body">１，会員が移動その他の事由により脱退しようとする時はその旨を会長に届け出て役員会の承認を得るものとする。２，前項の場合未納の会費は完納しなければならない。既納の会費は払い戻さない。</span></div>',
    '  <div class="rule-heading">委員会</div>',
    '  <div class="rule-block"><span class="rule-no">第１９条</span><span class="rule-body">本会には委員会を置く事が出来る。</span></div>',
    '  <div class="rule-heading">会則の成立</div>',
    '  <div class="rule-block"><span class="rule-no">第２０条</span><span class="rule-body">本会の会則は総会の３分の２以上の決をもって成立する。</span></div>',
    '  <div class="rule-heading">その他</div>',
    '  <div class="rule-block"><span class="rule-no">第２１条</span><span class="rule-body">本規約に定めない事項については、総会において決議するものとする。</span></div>',
    '  <div class="rule-block"><span class="rule-no">第２２条</span><span class="rule-body">本規約は、令和６年４月１日より実施する。</span></div>',
    '  <div style="margin-top:18pt;text-align:right;font-size:9pt;color:#555">附則　本規約は、令和６年４月１日より実施する。</div>',
    '  <div class="page-no">P12</div>',
    '</div>'
  ].join('\n');
}

// ============================================================
// P13 慶弔規定（静的）
// ============================================================

function buildBylaws3Page() {
  return [
    '<div class="page">',
    '  <div class="h-main" style="margin-bottom:2pt">埼玉県トラック協会所沢支部青年部会</div>',
    '  <div class="h-main">慶　弔　規　定</div>',
    '  <hr style="border:none;border-top:1px solid #999;margin-bottom:10pt">',
    '  <div class="rule-block"><span class="rule-no">第１条</span><span class="rule-body">（目的）この定義は、埼玉県トラック協会所沢支部青年部会の会員の平等な立場で公平かつ有意義な関係を維持する事を目的とする。</span></div>',
    '  <div class="rule-block"><span class="rule-no">第２条</span><span class="rule-body">（定義）この定義は、次の各号に定めるところによる。但し役員会で決定し定例総会において承認されたものに限る。</span></div>',
    '  <table style="margin:8pt 0">',
    '    <thead><tr>',
    '      <th style="width:110pt">区　分</th>',
    '      <th style="width:80pt">対　象</th>',
    '      <th style="width:90pt">金　額</th>',
    '      <th>備　考</th>',
    '    </tr></thead>',
    '    <tbody>',
    '      <tr><td>１・会員の結婚</td><td></td><td class="tr">３０，０００円</td><td>祝電１通</td></tr>',
    '      <tr><td>２・会員実子の誕生</td><td></td><td class="tr">１０，０００円</td><td></td></tr>',
    '      <tr><td>３・傷病見舞い</td><td></td><td class="tr">１０，０００円</td><td></td></tr>',
    '      <tr><td rowspan="4">４・弔慰</td><td>会員本人</td><td class="tr">５０，０００円</td><td>弔電１通、花輪１基又は生花１盛</td></tr>',
    '      <tr><td>会員企業代表者</td><td class="tr">３０，０００円</td><td>弔電１通、花輪１基又は生花１盛</td></tr>',
    '      <tr><td>配偶者</td><td class="tr">３０，０００円</td><td>弔電１通、花輪１基又は生花１盛</td></tr>',
    '      <tr><td>実子・会員両親</td><td class="tr">１０，０００〜２０，０００円</td><td>弔電１通、花輪１基又は生花１盛</td></tr>',
    '      <tr><td rowspan="3">５・脱会者</td><td>年齢到達</td><td class="tr">記念品</td><td></td></tr>',
    '      <tr><td>転勤</td><td class="tr">記念品</td><td></td></tr>',
    '      <tr><td>円満退職</td><td class="tr">記念品</td><td></td></tr>',
    '      <tr><td>６・その他</td><td></td><td></td><td></td></tr>',
    '    </tbody>',
    '  </table>',
    '  <div class="rule-block"><span class="rule-no">第３条</span><span class="rule-body">（特別事由）この規定に定め無き事項は、会長・副会長の決定による。但し会員に報告しなければならない。</span></div>',
    '  <div class="rule-block"><span class="rule-no">第４条</span><span class="rule-body">（会計経理）本規定の会計は、青年部会会計が行なう。</span></div>',
    '  <div class="rule-block"><span class="rule-no">第５条</span><span class="rule-body">（施行）本規定は令和６年４月１日より実施する。</span></div>',
    '  <div class="rule-block"><span class="rule-no">第６条</span><span class="rule-body">（付則）支部担当者の慶弔に対しては奉仕する。</span></div>',
    '  <div class="page-no">P13</div>',
    '</div>'
  ].join('\n');
}
