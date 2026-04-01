/**
 * SoukaiHtmlGen.gs
 * 総会資料 HTML ジェネレーター
 * スプレッドシート ID: 1GF86ve7gkpuhSDmQBu-rIgmv8UdzLUTsC5xH5GjRUw0
 */

// ─────────────────────────────────────────────
//  ユーティリティ
// ─────────────────────────────────────────────

/** HTML エスケープ */
function esc(v) {
  if (v === null || v === undefined) return '';
  return String(v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** 値を文字列にトリム */
function trimVal(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

/** 数値をカンマ区切り書式に */
function fmtNum(v) {
  if (v === null || v === undefined || v === '') return '';
  var n = Number(String(v).replace(/,/g, ''));
  if (isNaN(n)) return esc(v);
  return n.toLocaleString();
}

// ─────────────────────────────────────────────
//  Web App エントリーポイント
// ─────────────────────────────────────────────

function doGet(e) {
  try {
    var html = generateSoukaiHtml();
    return HtmlService.createHtmlOutput(html)
      .setTitle('総会資料')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<pre style="color:red;white-space:pre-wrap">エラーが発生しました:\n' +
      esc(err.message || String(err)) + '</pre>'
    );
  }
}

// ─────────────────────────────────────────────
//  メイン生成関数
// ─────────────────────────────────────────────

function generateSoukaiHtml() {
  var ss = SpreadsheetApp.openById('1GF86ve7gkpuhSDmQBu-rIgmv8UdzLUTsC5xH5GjRUw0');

  var parts = [];
  parts.push(getHtmlHead());
  parts.push('<body>');
  parts.push(buildCoverPage(ss));
  parts.push(buildAgendaPage(ss));
  parts.push(buildActivityReportPage(ss));
  parts.push(buildPrefecturalReportPage(ss));
  parts.push(buildFinancialReportPage(ss));
  parts.push(buildAssetInventoryPage(ss));
  parts.push(buildBusinessPlanPage(ss));
  parts.push(buildBudgetPage(ss));
  parts.push(buildOrgChartPage(ss));
  parts.push(buildMemberListPage(ss));
  parts.push(buildBylaws1Page());
  parts.push(buildBylaws2Page());
  parts.push(buildBylaws3Page());
  parts.push('</body></html>');

  return parts.join('\n');
}

// ─────────────────────────────────────────────
//  HTML ヘッド
// ─────────────────────────────────────────────

function getHtmlHead() {
  return '<!DOCTYPE html>\n<html lang="ja">\n<head>\n' +
    '<meta charset="UTF-8">\n' +
    '<title>総会資料 — 所沢支部青年部会</title>\n' +
    '<style>\n' + getCSS() + '\n</style>\n' +
    '</head>';
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
    '@media screen{body{background:#b0b0b0}.page{background:#fff;max-width:210mm;margin:20px auto;padding:20mm 18mm 20mm 22mm;box-shadow:0 3px 16px rgba(0,0,0,.25);border-radius:2px}}',
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

// ─────────────────────────────────────────────
//  P0: 表紙
// ─────────────────────────────────────────────

function buildCoverPage(ss) {
  var sheet = ss.getSheetByName('表紙');
  if (!sheet) return '<!-- 表紙シートが見つかりません -->';

  var data = sheet.getDataRange().getValues();

  // Row インデックスは 0-based
  var date   = trimVal(data[1][6]);   // Row2, ColG (index 1,6)
  var place  = trimVal(data[2][6]);   // Row3, ColG
  var time   = trimVal(data[3][5]);   // Row4, ColF (index 3,5)
  var title  = trimVal(data[15][0]);  // Row16, ColA (index 15,0)
  var period = trimVal(data[17][0]);  // Row18, ColA (index 17,0)
  var org1   = trimVal(data[47][2]);  // Row48, ColC (index 47,2)
  var org2   = trimVal(data[48][2]);  // Row49, ColC (index 48,2)

  var html = [];
  html.push('<!-- ===== 表紙 ===== -->');
  html.push('<div class="page cover-page">');
  html.push('  <div>');
  if (date)   html.push('    <div class="cover-info">'  + esc(date)   + '</div>');
  if (place)  html.push('    <div class="cover-info">'  + esc(place)  + '</div>');
  if (time)   html.push('    <div class="cover-info">'  + esc(time)   + '</div>');
  if (title)  html.push('    <div class="cover-title">' + esc(title)  + '</div>');
  if (period) html.push('    <div class="cover-period">'+ esc(period) + '</div>');
  if (org1)   html.push('    <div class="cover-org">'   + esc(org1)   + '</div>');
  if (org2)   html.push('    <div class="cover-org">'   + esc(org2)   + '</div>');
  html.push('  </div>');
  html.push('</div>');

  return html.join('\n');
}

// ─────────────────────────────────────────────
//  P1: 総会次第
// ─────────────────────────────────────────────

function buildAgendaPage(ss) {
  var sheet = ss.getSheetByName('P1\u3000総会次第');
  if (!sheet) return '<!-- P1シートが見つかりません -->';

  var data = sheet.getDataRange().getValues();

  // 議案行を収集: ColC (index 2) に "議案" を含む行
  var proposals = [];
  for (var i = 0; i < data.length; i++) {
    var num  = trimVal(data[i][2]);
    var text = trimVal(data[i][4]);
    if (num.indexOf('議案') !== -1 && text !== '') {
      proposals.push({ num: num, text: text });
    }
  }

  // 固定の次第項目（議案番号以外）
  var fixedItems = [
    { num: '１、', text: '開会の辞' },
    { num: '２、', text: '支部長挨拶' },
    { num: '３、', text: '議長選出' },
    { num: '４、', text: '定足数確認' }
  ];

  var html = [];
  html.push('<!-- ===== P1 総会次第 ===== -->');
  html.push('<div class="page">');
  html.push('  <div class="agenda-wrap">');
  html.push('    <div class="agenda-inner">');
  html.push('      <div class="agenda-title">総　会　次　第</div>');

  // 固定項目
  fixedItems.forEach(function(item) {
    html.push('      <div class="ag-item"><span class="ag-num">' + esc(item.num) + '</span><span>' + esc(item.text) + '</span></div>');
  });

  // 議事 + 議案
  html.push('      <div class="ag-item"><span class="ag-num">５、</span><span>議　　事');
  if (proposals.length > 0) {
    html.push('        <div class="ag-sub">');
    proposals.forEach(function(p) {
      html.push('          <div class="ag-subitem"><span class="ag-subnum">' + esc(p.num) + '</span><span>' + esc(p.text) + '</span></div>');
    });
    html.push('        </div>');
  }
  html.push('      </span></div>');

  html.push('      <div class="ag-item"><span class="ag-num">６、</span><span>閉会の辞</span></div>');
  html.push('    </div>');
  html.push('  </div>');
  html.push('  <div class="page-no">P1</div>');
  html.push('</div>');

  return html.join('\n');
}

// ─────────────────────────────────────────────
//  P2: 活動報告
// ─────────────────────────────────────────────

function buildActivityReportPage(ss) {
  var sheet = ss.getSheetByName('P2\u3000活動報告');
  if (!sheet) return '<!-- P2シートが見つかりません -->';

  var data = sheet.getDataRange().getValues();

  // 期間テキスト: Row4 (index 3), ColI (index 8)
  var periodText = trimVal(data[3][8]);

  // ページタイトル (H-main): 決まった行のタイトルがあれば使う、なければデフォルト
  // スプレッドシートに専用タイトルセルがない場合はデフォルト文字列を使用
  var pageTitle = '活動内容';
  // Row1 (index 0) ColA などにタイトルがあれば取得
  var titleCandidate = trimVal(data[0][0]);
  if (titleCandidate !== '') pageTitle = titleCandidate;

  var html = [];
  html.push('<!-- ===== P2 活動報告 ===== -->');
  html.push('<div class="page">');
  html.push('  <div class="h-sub" style="text-align:center;margin-bottom:2pt">埼玉県トラック協会 所沢支部 青年部会</div>');
  html.push('  <div class="h-main">' + esc(pageTitle) + '</div>');
  if (periodText) {
    html.push('  <div class="period-note">' + esc(periodText) + '</div>');
  }
  html.push('  <table>');
  html.push('    <thead><tr><th style="width:90pt">月　日</th><th>支部事業名</th><th style="width:90pt">担当会社</th></tr></thead>');
  html.push('    <tbody>');

  var lastYear = '';
  // データ行は index 4 (Row5) から
  for (var i = 4; i < data.length; i++) {
    var row   = data[i];
    var year  = trimVal(row[0]);
    var date  = trimVal(row[1]);
    var event = trimVal(row[5]);
    var comp  = trimVal(row[8]);

    // 日付または事業名が空ならスキップ
    if (date === '' || event === '') continue;

    // 年を繰り越す
    if (year !== '') lastYear = year;

    // 月日表示: 年があれば "年X月X日"、なければ月日のみ
    var displayDate = (lastYear !== '' && date.indexOf('年') === -1)
      ? lastYear + date
      : date;

    html.push('      <tr><td class="tc">' + esc(displayDate) + '</td><td>' + esc(event) + '</td><td>' + esc(comp) + '</td></tr>');
  }

  html.push('    </tbody>');
  html.push('  </table>');
  html.push('  <div class="page-no">P2</div>');
  html.push('</div>');

  return html.join('\n');
}

// ─────────────────────────────────────────────
//  P3: 県青年部事業報告
// ─────────────────────────────────────────────

function buildPrefecturalReportPage(ss) {
  var sheet = ss.getSheetByName('P3\u3000県青年部事業報告');
  if (!sheet) return '<!-- P3シートが見つかりません -->';

  var data = sheet.getDataRange().getValues();

  // ページタイトル
  var pageTitle = '県青年部会事業報告';
  var titleCandidate = trimVal(data[0][0]);
  if (titleCandidate !== '') pageTitle = titleCandidate;

  var html = [];
  html.push('<!-- ===== P3 県青年部事業報告 ===== -->');
  html.push('<div class="page">');
  html.push('  <div class="h-sub" style="text-align:center;margin-bottom:2pt">埼玉県トラック協会 所沢支部 青年部会</div>');
  html.push('  <div class="h-main">' + esc(pageTitle) + '</div>');
  html.push('  <table>');
  html.push('    <thead><tr><th style="width:90pt">月　日</th><th>事　業　名</th><th style="width:120pt">場　所</th></tr></thead>');
  html.push('    <tbody>');

  var lastYear = '';
  for (var i = 4; i < data.length; i++) {
    var row   = data[i];
    var year  = trimVal(row[0]);
    var date  = trimVal(row[1]);
    var event = trimVal(row[5]);
    var place = trimVal(row[8]);

    if (date === '' || event === '') continue;

    if (year !== '') lastYear = year;

    var displayDate = (lastYear !== '' && date.indexOf('年') === -1)
      ? lastYear + date
      : date;

    html.push('      <tr><td class="tc">' + esc(displayDate) + '</td><td>' + esc(event) + '</td><td>' + esc(place) + '</td></tr>');
  }

  html.push('    </tbody>');
  html.push('  </table>');
  html.push('  <div class="page-no">P3</div>');
  html.push('</div>');

  return html.join('\n');
}

// ─────────────────────────────────────────────
//  P4: 決算報告
// ─────────────────────────────────────────────

function buildFinancialReportPage(ss) {
  var sheet = ss.getSheetByName('P4\u3000決算報告');
  if (!sheet) return '<!-- P4シートが見つかりません -->';

  var data = sheet.getDataRange().getValues();

  var title      = trimVal(data[0][0]);
  var periodText = trimVal(data[2][2]);

  // 収入ヘッダー: Row4 (index 3)
  var incHdrs = [
    trimVal(data[3][0]) || '科　目',
    trimVal(data[3][1]) || '予算額（円）',
    trimVal(data[3][2]) || '決算額（円）',
    trimVal(data[3][3]) || '差　額（円）',
    trimVal(data[3][4]) || '備考'
  ];

  // 収入データ: index 4 以降、ColA が空になるまで
  var incRows = [];
  var i = 4;
  while (i < data.length && trimVal(data[i][0]) !== '') {
    incRows.push(data[i]);
    i++;
  }

  // 支出の部を検索
  var expSectionRow = -1;
  for (var j = i; j < data.length; j++) {
    var cellVal = trimVal(data[j][0]);
    if (cellVal.indexOf('支出の部') !== -1 || cellVal.indexOf('【支出') !== -1) {
      expSectionRow = j;
      break;
    }
  }

  var expHdrs = ['科　目', '予算額（円）', '決算額（円）', '差　額（円）', '備考'];
  var expRows = [];
  var sigRows = [];

  if (expSectionRow !== -1) {
    // 支出ヘッダーは支出の部の次の行
    var expHdrRow = expSectionRow + 1;
    if (expHdrRow < data.length) {
      expHdrs = [
        trimVal(data[expHdrRow][0]) || '科　目',
        trimVal(data[expHdrRow][1]) || '予算額（円）',
        trimVal(data[expHdrRow][2]) || '決算額（円）',
        trimVal(data[expHdrRow][3]) || '差　額（円）',
        trimVal(data[expHdrRow][4]) || '備考'
      ];
    }
    // 支出データ
    var k = expHdrRow + 1;
    while (k < data.length) {
      var ca = trimVal(data[k][0]);
      if (ca === '') { k++; continue; }
      if (ca.indexOf('上記') !== -1) {
        // 上記以降は署名行
        for (var s = k; s < data.length; s++) {
          var sigLine = trimVal(data[s][0]);
          if (sigLine !== '') sigRows.push(sigLine);
        }
        break;
      }
      expRows.push(data[k]);
      k++;
    }
  }

  // 収支差額行 (最後の収入行 - 最後の支出行から計算。シートから取るのが望ましいが
  // シート構造が不定のため、収入最終行・支出最終行からラベルを使う)
  // 収支差額行はシートの "上記" より前の特殊行を探す
  // ここでは incRows の最後・expRows の最後を合計行として扱い、差額行は別途生成
  var lastIncRow = incRows.length > 0 ? incRows[incRows.length - 1] : null;
  var lastExpRow = expRows.length > 0 ? expRows[expRows.length - 1] : null;

  // 収支差額の行: シートから探す
  var diffRow = null;
  for (var d = 0; d < data.length; d++) {
    var dl = trimVal(data[d][0]);
    if (dl.indexOf('収支差額') !== -1 || dl.indexOf('次年度繰越') !== -1) {
      diffRow = data[d];
      break;
    }
  }

  var html = [];
  html.push('<!-- ===== P4 決算報告 ===== -->');
  html.push('<div class="page">');
  html.push('  <div class="h-main">' + esc(title || '決算報告書') + '</div>');

  // 収入の部
  html.push('  <div class="h-title" style="margin:4pt 0 3pt">【収入の部】</div>');
  html.push('  <table>');
  html.push('    <thead><tr>' +
    '<th>' + esc(incHdrs[0]) + '</th>' +
    '<th style="width:80pt">' + esc(incHdrs[1]) + '</th>' +
    '<th style="width:80pt">' + esc(incHdrs[2]) + '</th>' +
    '<th style="width:72pt">' + esc(incHdrs[3]) + '</th>' +
    '<th>' + esc(incHdrs[4]) + '</th>' +
    '</tr></thead>');
  html.push('    <tbody>');
  incRows.forEach(function(r, idx) {
    var cls = (idx === incRows.length - 1) ? ' class="row-total"' : '';
    html.push('      <tr' + cls + '><td>' + esc(trimVal(r[0])) + '</td>' +
      '<td class="tr">' + esc(fmtNum(r[1])) + '</td>' +
      '<td class="tr">' + esc(fmtNum(r[2])) + '</td>' +
      '<td class="tr">' + esc(fmtNum(r[3])) + '</td>' +
      '<td>' + esc(trimVal(r[4])) + '</td></tr>');
  });
  html.push('    </tbody>');
  html.push('  </table>');

  // 支出の部
  html.push('  <div class="h-title" style="margin:4pt 0 3pt">【支出の部】</div>');
  html.push('  <table>');
  html.push('    <thead><tr>' +
    '<th>' + esc(expHdrs[0]) + '</th>' +
    '<th style="width:80pt">' + esc(expHdrs[1]) + '</th>' +
    '<th style="width:80pt">' + esc(expHdrs[2]) + '</th>' +
    '<th style="width:72pt">' + esc(expHdrs[3]) + '</th>' +
    '<th>' + esc(expHdrs[4]) + '</th>' +
    '</tr></thead>');
  html.push('    <tbody>');
  expRows.forEach(function(r, idx) {
    var cls = (idx === expRows.length - 1) ? ' class="row-total"' : '';
    html.push('      <tr' + cls + '><td>' + esc(trimVal(r[0])) + '</td>' +
      '<td class="tr">' + esc(fmtNum(r[1])) + '</td>' +
      '<td class="tr">' + esc(fmtNum(r[2])) + '</td>' +
      '<td class="tr">' + esc(fmtNum(r[3])) + '</td>' +
      '<td>' + esc(trimVal(r[4])) + '</td></tr>');
  });
  html.push('    </tbody>');
  html.push('  </table>');

  // 収支差額行
  if (diffRow) {
    html.push('  <table style="margin-top:6pt"><tr class="row-total" style="background:#E8F5E9">' +
      '<td>' + esc(trimVal(diffRow[0])) + '</td>' +
      '<td class="tr" style="width:80pt">' + esc(fmtNum(diffRow[1])) + '</td>' +
      '<td class="tr" style="width:80pt">' + esc(fmtNum(diffRow[2])) + '</td>' +
      '<td class="tr" style="width:72pt">' + esc(fmtNum(diffRow[3])) + '</td>' +
      '<td>' + esc(trimVal(diffRow[4])) + '</td></tr></table>');
  }

  // 署名
  if (sigRows.length > 0) {
    html.push('  <div class="sig">' + sigRows.map(esc).join('<br>') + '</div>');
  }

  html.push('  <div class="page-no">P4</div>');
  html.push('</div>');

  return html.join('\n');
}

// ─────────────────────────────────────────────
//  P5: 財産目録
// ─────────────────────────────────────────────

function buildAssetInventoryPage(ss) {
  var sheet = ss.getSheetByName('P5\u3000財産目録');
  if (!sheet) return '<!-- P5シートが見つかりません -->';

  var data = sheet.getDataRange().getValues();

  var title  = trimVal(data[0][0]) || '財　産　目　録';
  var dateStr = trimVal(data[1][2]);

  var html = [];
  html.push('<!-- ===== P5 財産目録 ===== -->');
  html.push('<div class="page">');
  html.push('  <div class="h-main">' + esc(title) + '</div>');
  if (dateStr) {
    html.push('  <div class="period-note">' + esc(dateStr) + '</div>');
  }
  html.push('  <table>');
  html.push('    <thead><tr><th>科　目</th><th style="width:120pt">金額（円）</th><th>適　用</th></tr></thead>');
  html.push('    <tbody>');

  var sigRows = [];
  for (var i = 4; i < data.length; i++) {
    var label  = trimVal(data[i][0]);
    var amount = trimVal(data[i][1]);
    var note   = trimVal(data[i][2]);

    if (label === '') continue;

    if (label.indexOf('上記') !== -1) {
      // 上記以降は署名
      for (var s = i; s < data.length; s++) {
        var sl = trimVal(data[s][0]);
        if (sl !== '') sigRows.push(sl);
      }
      break;
    }

    var rowCls = '';
    if (label.indexOf('部】') !== -1) {
      rowCls = ' class="row-section"';
    } else if (label.indexOf('合計') !== -1 || label.indexOf('正味資産') !== -1) {
      // 正味資産は緑背景
      if (label.indexOf('正味資産') !== -1) {
        rowCls = ' class="row-total" style="background:#E8F5E9"';
      } else {
        rowCls = ' class="row-total"';
      }
    }

    html.push('      <tr' + rowCls + '><td>' + esc(label) + '</td>' +
      '<td class="tr">' + esc(fmtNum(amount)) + '</td>' +
      '<td>' + esc(note) + '</td></tr>');
  }

  html.push('    </tbody>');
  html.push('  </table>');

  if (sigRows.length > 0) {
    html.push('  <div class="sig">' + sigRows.map(esc).join('<br>') + '</div>');
  }

  html.push('  <div class="page-no">P5</div>');
  html.push('</div>');

  return html.join('\n');
}

// ─────────────────────────────────────────────
//  P6: 事業計画案
// ─────────────────────────────────────────────

function buildBusinessPlanPage(ss) {
  var sheet = ss.getSheetByName('P6\u3000事業計画案');
  if (!sheet) return '<!-- P6シートが見つかりません -->';

  var data = sheet.getDataRange().getValues();

  var title        = trimVal(data[0][0]) || '事業計画（案）';
  var sectionTitle = trimVal(data[5][0]) || '１・事業計画';

  // 計画項目: ColC (index 2) で /^（[^）]+）/ にマッチする値
  var planItems = [];
  var planPattern = /^（[^）]+）/;
  for (var i = 0; i < data.length; i++) {
    var val = trimVal(data[i][2]);
    if (planPattern.test(val)) {
      // 番号部分と内容部分を分割
      var match = val.match(/^（[^）]+）/);
      var num   = match ? match[0] : '';
      var text  = val.substring(num.length).trim();
      planItems.push({ num: num, text: text });
    }
  }

  var html = [];
  html.push('<!-- ===== P6 事業計画（案） ===== -->');
  html.push('<div class="page">');
  html.push('  <div class="h-main">' + esc(title) + '</div>');
  html.push('  <div class="plan-wrap">');
  html.push('    <div class="plan-inner">');
  html.push('      <div class="plan-section">' + esc(sectionTitle) + '</div>');

  planItems.forEach(function(item) {
    html.push('      <div class="plan-item"><span class="plan-num">' + esc(item.num) + '</span><span>' + esc(item.text) + '</span></div>');
  });

  html.push('    </div>');
  html.push('  </div>');
  html.push('  <div class="page-no">P6</div>');
  html.push('</div>');

  return html.join('\n');
}

// ─────────────────────────────────────────────
//  P7: 予算案
// ─────────────────────────────────────────────

function buildBudgetPage(ss) {
  var sheet = ss.getSheetByName('P7\u3000予算案');
  if (!sheet) return '<!-- P7シートが見つかりません -->';

  var data = sheet.getDataRange().getValues();

  var title = trimVal(data[0][0]) || '予算（案）';

  // 収入ヘッダー: Row4 (index 3)
  var incHdrs = [
    trimVal(data[3][0]) || '科　目',
    trimVal(data[3][1]) || '前年度予算額',
    trimVal(data[3][2]) || '予算額（円）',
    trimVal(data[3][3]) || '増　減',
    trimVal(data[3][4]) || '摘　要'
  ];

  // 収入データ: index 4 以降、ColA が空になるまで
  var incRows = [];
  var i = 4;
  while (i < data.length && trimVal(data[i][0]) !== '') {
    incRows.push(data[i]);
    i++;
  }

  // 支出の部を検索
  var expSectionRow = -1;
  for (var j = i; j < data.length; j++) {
    var cv = trimVal(data[j][0]);
    if (cv.indexOf('支出の部') !== -1 || cv.indexOf('【支出') !== -1) {
      expSectionRow = j;
      break;
    }
  }

  var expHdrs = ['科　目', '前年度予算額', '予算額（円）', '増　減', '摘　要'];
  var expRows = [];

  if (expSectionRow !== -1) {
    var expHdrRow = expSectionRow + 1;
    if (expHdrRow < data.length) {
      expHdrs = [
        trimVal(data[expHdrRow][0]) || '科　目',
        trimVal(data[expHdrRow][1]) || '前年度予算額',
        trimVal(data[expHdrRow][2]) || '予算額（円）',
        trimVal(data[expHdrRow][3]) || '増　減',
        trimVal(data[expHdrRow][4]) || '摘　要'
      ];
    }
    var k = expHdrRow + 1;
    while (k < data.length) {
      var ca = trimVal(data[k][0]);
      if (ca === '') { k++; continue; }
      if (ca.indexOf('上記') !== -1) break;
      expRows.push(data[k]);
      k++;
    }
  }

  var html = [];
  html.push('<!-- ===== P7 予算案 ===== -->');
  html.push('<div class="page">');
  html.push('  <div class="h-main">' + esc(title) + '</div>');

  // 収入の部
  html.push('  <div class="h-title" style="margin:4pt 0 3pt">【収入の部】</div>');
  html.push('  <table>');
  html.push('    <thead><tr>' +
    '<th>' + esc(incHdrs[0]) + '</th>' +
    '<th style="width:80pt">' + esc(incHdrs[1]) + '</th>' +
    '<th style="width:80pt">' + esc(incHdrs[2]) + '</th>' +
    '<th style="width:72pt">' + esc(incHdrs[3]) + '</th>' +
    '<th>' + esc(incHdrs[4]) + '</th>' +
    '</tr></thead>');
  html.push('    <tbody>');
  incRows.forEach(function(r, idx) {
    var cls = (idx === incRows.length - 1) ? ' class="row-total"' : '';
    html.push('      <tr' + cls + '><td>' + esc(trimVal(r[0])) + '</td>' +
      '<td class="tr">' + esc(fmtNum(r[1])) + '</td>' +
      '<td class="tr">' + esc(fmtNum(r[2])) + '</td>' +
      '<td class="tr">' + esc(fmtNum(r[3])) + '</td>' +
      '<td>' + esc(trimVal(r[4])) + '</td></tr>');
  });
  html.push('    </tbody>');
  html.push('  </table>');

  // 支出の部
  html.push('  <div class="h-title" style="margin:4pt 0 3pt">【支出の部】</div>');
  html.push('  <table>');
  html.push('    <thead><tr>' +
    '<th>' + esc(expHdrs[0]) + '</th>' +
    '<th style="width:80pt">' + esc(expHdrs[1]) + '</th>' +
    '<th style="width:80pt">' + esc(expHdrs[2]) + '</th>' +
    '<th style="width:72pt">' + esc(expHdrs[3]) + '</th>' +
    '<th>' + esc(expHdrs[4]) + '</th>' +
    '</tr></thead>');
  html.push('    <tbody>');
  expRows.forEach(function(r, idx) {
    var cls = (idx === expRows.length - 1) ? ' class="row-total"' : '';
    html.push('      <tr' + cls + '><td>' + esc(trimVal(r[0])) + '</td>' +
      '<td class="tr">' + esc(fmtNum(r[1])) + '</td>' +
      '<td class="tr">' + esc(fmtNum(r[2])) + '</td>' +
      '<td class="tr">' + esc(fmtNum(r[3])) + '</td>' +
      '<td>' + esc(trimVal(r[4])) + '</td></tr>');
  });
  html.push('    </tbody>');
  html.push('  </table>');

  html.push('  <div class="page-no">P7</div>');
  html.push('</div>');

  return html.join('\n');
}

// ─────────────────────────────────────────────
//  P8: 組織図
// ─────────────────────────────────────────────

function buildOrgChartPage(ss) {
  var sheet = ss.getSheetByName('P8\u3000組織図');
  if (!sheet) return '<!-- P8シートが見つかりません -->';

  // GAS は 1-indexed
  function gc(row, col) {
    var v = sheet.getRange(row, col).getValue();
    return trimVal(v);
  }

  var advisorRole      = gc(10, 16);  // P10
  var advisorName      = gc(10, 19);  // S10
  var presidentRole    = gc(13, 11);  // K13
  var presidentName    = gc(13, 13);  // M13
  var officeRole       = gc(18, 4);   // D18
  var officeName       = gc(18, 8);   // H18
  var auditorName      = gc(18, 13);  // M18
  var auditorRole      = gc(18, 17);  // Q18
  var subOfficeRole    = gc(23, 4);   // D23
  var subOfficeName    = gc(23, 8);   // H23
  var accountantName   = gc(23, 13);  // M23
  var accountantRole   = gc(23, 17);  // Q23
  var vpRole           = gc(31, 10);  // J31
  var vpName           = gc(31, 14);  // N31
  var ccRole           = gc(39, 10);  // J39
  var ccName           = gc(42, 10);  // J42

  // 部会員グリッド
  var memberCoords = [
    [50,4],[50,8],[50,12],[50,16],
    [52,4],[52,8],[52,12],[52,16],
    [54,4],[54,8],[54,12],[54,16],
    [56,4],[56,8],[56,12],[56,16]
  ];
  var members = memberCoords.map(function(coord) {
    return gc(coord[0], coord[1]);
  }).filter(function(v) { return v !== ''; });

  // ページタイトル
  var pageTitle = '組織図';
  try {
    var t = gc(1, 1);
    if (t !== '') pageTitle = t;
  } catch(e) {}

  var html = [];
  html.push('<!-- ===== P8 組織図 ===== -->');
  html.push('<div class="page">');
  html.push('  <div class="h-main">令和度　青年部組織図</div>');
  html.push('  <div style="display:flex;flex-direction:column;align-items:center;font-family:\'Noto Sans JP\',sans-serif;font-size:9pt;padding:2pt 0">');

  // 総会ボックス
  html.push('    <div style="border:1.5px solid #111;padding:4px 36px;font-size:10.5pt;font-weight:700;letter-spacing:.25em;background:#fff">総　会</div>');

  // 顧問ブロック
  if (advisorName) {
    html.push('    <div style="position:relative;width:100%;height:32px;display:flex;justify-content:center">');
    html.push('      <div style="width:1.5px;height:100%;background:#333"></div>');
    html.push('      <div style="position:absolute;top:16px;left:50%;width:100px;height:1.5px;background:#333"></div>');
    html.push('      <div style="position:absolute;top:4px;left:calc(50% + 100px);border:1px solid #880e4f;padding:2px 10px;font-size:9pt;background:#fce4ec;text-align:center;white-space:nowrap">');
    html.push('        ' + esc(advisorRole || '顧　問') + '<br>');
    html.push('        <span style="font-weight:400">' + esc(advisorName) + '</span>');
    html.push('      </div>');
    html.push('    </div>');
  } else {
    html.push('    <div style="width:1.5px;height:32px;background:#333"></div>');
  }

  // 会長ボックス
  html.push('    <div style="border:1.5px solid #111;padding:4px 28px;font-size:10.5pt;font-weight:700;background:#fff">');
  html.push('      ' + esc((presidentRole || '会　長') + (presidentName ? '　' + presidentName : '')));
  html.push('    </div>');

  // 縦線
  html.push('    <div style="width:1.5px;height:24px;background:#333"></div>');

  // 役員会十字ハブ
  html.push('    <div style="position:relative;width:420px;text-align:center;padding:26px 0;margin-bottom:2px">');
  html.push('      <div style="position:absolute;top:0;bottom:0;left:50%;width:1.5px;background:#333;margin-left:-0.75px;z-index:1"></div>');
  html.push('      <div style="position:absolute;top:50%;left:50px;right:50px;height:1.5px;background:#333;margin-top:-0.75px;z-index:1"></div>');
  html.push('      <div style="position:absolute;top:0;bottom:0;left:50px;width:1.5px;background:#333;margin-left:-0.75px;z-index:1"></div>');
  html.push('      <div style="position:absolute;top:0;bottom:0;right:50px;width:1.5px;background:#333;margin-right:-0.75px;z-index:1"></div>');
  html.push('      <div style="position:relative;z-index:2;display:inline-block;border:1.5px solid #111;background:#fff;padding:6px 20px;font-weight:700;letter-spacing:.1em;font-size:10pt">役員会</div>');

  // 左上：事務局
  html.push('      <div style="position:absolute;top:-10px;left:50px;transform:translateX(-50%);z-index:2;border:1px solid #333;background:#fff;padding:4px 12px;text-align:center;min-width:105px;font-weight:700;font-size:9.5pt">');
  html.push('        ' + esc(officeRole || '事務局'));
  html.push('        <div style="font-weight:400;font-size:8.5pt;margin-top:3px">' + esc(officeName) + '</div>');
  html.push('      </div>');

  // 左下：事務局補佐
  html.push('      <div style="position:absolute;bottom:-10px;left:50px;transform:translateX(-50%);z-index:2;border:1px solid #333;background:#fff;padding:4px 12px;text-align:center;min-width:105px;font-weight:700;font-size:9.5pt">');
  html.push('        ' + esc(subOfficeRole || '事務局補佐'));
  html.push('        <div style="font-weight:400;font-size:8.5pt;margin-top:3px">' + esc(subOfficeName) + '</div>');
  html.push('      </div>');

  // 右上：会計監査
  html.push('      <div style="position:absolute;top:-10px;right:50px;transform:translateX(50%);z-index:2;border:1px solid #333;background:#fff;padding:4px 12px;text-align:center;min-width:105px;font-weight:700;font-size:9.5pt">');
  html.push('        ' + esc(auditorRole || '会計監査'));
  html.push('        <div style="font-weight:400;font-size:8.5pt;margin-top:3px">' + esc(auditorName) + '</div>');
  html.push('      </div>');

  // 右下：会計
  html.push('      <div style="position:absolute;bottom:-10px;right:50px;transform:translateX(50%);z-index:2;border:1px solid #333;background:#fff;padding:4px 12px;text-align:center;min-width:105px;font-weight:700;font-size:9.5pt">');
  html.push('        ' + esc(accountantRole || '会　計'));
  html.push('        <div style="font-weight:400;font-size:8.5pt;margin-top:3px">' + esc(accountantName) + '</div>');
  html.push('      </div>');

  html.push('    </div>');

  // 縦線
  html.push('    <div style="width:1.5px;height:24px;background:#333"></div>');

  // 副会長
  if (vpName) {
    html.push('    <div style="border:1px solid #333;background:#fff;padding:4px 16px;font-weight:700;font-size:9.5pt;text-align:center">');
    html.push('      ' + esc(vpRole || '副会長'));
    html.push('      <div style="font-weight:400;font-size:8.5pt;margin-top:3px">' + esc(vpName) + '</div>');
    html.push('    </div>');
    html.push('    <div style="width:1.5px;height:18px;background:#333"></div>');
  }

  // 事業委員長
  if (ccName) {
    html.push('    <div style="border:1px solid #333;background:#fff;padding:4px 20px;font-weight:700;font-size:9.5pt;text-align:center">');
    html.push('      ' + esc(ccRole || '事業委員長'));
    html.push('      <div style="font-weight:400;font-size:8.5pt;margin-top:3px">' + esc(ccName) + '</div>');
    html.push('    </div>');
  }

  // 部会員セクション
  html.push('    <div style="width:100%;border-top:1.5px solid #ccc;margin-top:16px;padding-top:10px">');
  html.push('      <div style="text-align:center;font-size:10pt;font-weight:700;margin-bottom:8px;letter-spacing:.2em">部　会　員</div>');
  html.push('      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:4px">');
  members.forEach(function(m) {
    html.push('        <div style="border:1px solid #666;padding:4px;text-align:center;font-size:8.5pt">' + esc(m) + '</div>');
  });
  html.push('      </div>');
  html.push('    </div>');

  html.push('  </div>');
  html.push('  <div class="page-no">P8</div>');
  html.push('</div>');

  return html.join('\n');
}

// ─────────────────────────────────────────────
//  P9-10: 会員名簿
// ─────────────────────────────────────────────

function buildMemberListPage(ss) {
  var sheet = ss.getSheetByName('P9\u3000会員名簿①');
  if (!sheet) return '<!-- P9シートが見つかりません -->';

  var data = sheet.getDataRange().getValues();

  var html = [];
  html.push('<!-- ===== P9-10 会員名簿 ===== -->');
  html.push('<div class="page">');
  html.push('  <div class="h-main">会員名簿</div>');
  html.push('  <table class="tbl-sm">');
  html.push('    <thead><tr>' +
    '<th style="width:26pt">No</th>' +
    '<th style="width:56pt">役職</th>' +
    '<th style="width:66pt">氏名</th>' +
    '<th>会社名</th>' +
    '<th>住　所</th>' +
    '<th style="width:80pt">TEL</th>' +
    '</tr></thead>');
  html.push('    <tbody>');

  // 会員データは Row3 (index 2) から 2行1セット
  var i = 2;
  while (i + 1 < data.length) {
    var row1 = data[i];
    var row2 = data[i + 1];

    var no      = trimVal(row1[0]);
    var role    = trimVal(row1[1]);
    var name    = trimVal(row1[2]);
    var company = trimVal(row1[4]);
    var zip     = trimVal(row1[7]);
    var address = trimVal(row1[8]);

    // No と名前が両方空ならスキップ
    if (no === '' && name === '') {
      i += 2;
      continue;
    }

    // TEL: 2行目の ColH (index 7)
    var telRaw = trimVal(row2[7]);
    // "TEL " プレフィックスを除去、" ・ FAX ..." サフィックスを除去
    var tel = telRaw.replace(/^TEL\s*/i, '').replace(/\s*・\s*FAX.*/i, '').trim();

    // 住所: 郵便番号 + 住所
    var fullAddress = '';
    if (zip !== '' && address !== '') {
      fullAddress = '〒' + zip + ' ' + address;
    } else if (address !== '') {
      fullAddress = address;
    } else if (zip !== '') {
      fullAddress = '〒' + zip;
    }

    html.push('      <tr>' +
      '<td class="tc">' + esc(no) + '</td>' +
      '<td class="tc">' + esc(role) + '</td>' +
      '<td>' + esc(name) + '</td>' +
      '<td>' + esc(company) + '</td>' +
      '<td>' + esc(fullAddress) + '</td>' +
      '<td>' + esc(tel) + '</td>' +
      '</tr>');

    i += 2;
  }

  html.push('    </tbody>');
  html.push('  </table>');
  html.push('  <div class="page-no">P9-10</div>');
  html.push('</div>');

  return html.join('\n');
}

// ─────────────────────────────────────────────
//  P11: 会則①（静的）
// ─────────────────────────────────────────────

function buildBylaws1Page() {
  return [
    '<!-- ===== P11 規約① ===== -->',
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

// ─────────────────────────────────────────────
//  P12: 会則②（静的）
// ─────────────────────────────────────────────

function buildBylaws2Page() {
  return [
    '<!-- ===== P12 規約② ===== -->',
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

// ─────────────────────────────────────────────
//  P13: 慶弔規定（静的）
// ─────────────────────────────────────────────

function buildBylaws3Page() {
  return [
    '<!-- ===== P13 慶弔規定 ===== -->',
    '<div class="page">',
    '  <div class="h-main" style="margin-bottom:2pt">埼玉県トラック協会所沢支部青年部会</div>',
    '  <div class="h-main">慶　弔　規　定</div>',
    '  <hr style="border:none;border-top:1px solid #999;margin-bottom:10pt">',
    '  <div class="rule-block"><span class="rule-no">第１条</span><span class="rule-body">（目的）この定義は、埼玉県トラック協会所沢支部青年部会の会員の平等な立場で公平かつ有意義な関係を維持する事を目的とする。</span></div>',
    '  <div class="rule-block"><span class="rule-no">第２条</span><span class="rule-body">（定義）この定義は、次の各号に定めるところによる。但し役員会で決定し定例総会において承認されたものに限る。</span></div>',
    '  <table style="margin:8pt 0">',
    '    <thead><tr><th style="width:110pt">区　分</th><th style="width:80pt">対　象</th><th style="width:90pt">金　額</th><th>備　考</th></tr></thead>',
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
