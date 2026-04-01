// ============================================================
// 青年部 テンプレート×AI 文書自動生成システム
// Code.gs - エントリポイント + ルーティング
// ============================================================

var DOC_ID = '1mmktQr0dQdmZZI3bVzvELwVSzIMue_T8pYb4DApop-4';
var SS_ID  = '1GF86ve7gkpuhSDmQBu-rIgmv8UdzLUTsC5xH5GjRUw0';

// ── メイン：テンプレートから文書生成 ──
function generateDocument(templateName, userInput) {
  templateName = templateName || '定期総会';
  userInput = userInput || '';

  var template = getTemplate(templateName);
  if (!template) {
    Logger.log('テンプレートが見つかりません: ' + templateName);
    return;
  }

  var variables = {};
  if (userInput && userInput.length > 0) {
    var prompt = getPromptTemplate(templateName);
    variables = callGeminiForExtraction(prompt, userInput);
  } else {
    variables = template.defaults || {};
  }

  var doc = DocumentApp.openById(DOC_ID);
  var body = doc.getBody();
  body.clear();
  body.setMarginTop(56.7); body.setMarginBottom(56.7);
  body.setMarginLeft(56.7); body.setMarginRight(56.7);

  renderSections(body, template.sections, variables);
  doc.saveAndClose();

  Logger.log('文書生成完了: ' + doc.getUrl());
  return doc.getUrl();
}

// ── 定期総会：清書版を直接生成（データ埋め込み済み） ──
function generateSoukaiSeiSho() {
  var template = getTemplate('定期総会');
  var doc = DocumentApp.openById(DOC_ID);
  var body = doc.getBody();
  body.clear();
  body.setMarginTop(56.7); body.setMarginBottom(56.7);
  body.setMarginLeft(56.7); body.setMarginRight(56.7);

  renderSections(body, template.sections, template.defaults);
  doc.saveAndClose();

  Logger.log('定期総会 清書版 生成完了: ' + doc.getUrl());
  return doc.getUrl();
}

// ── 定期総会：テンプレ版を生成（プレースホルダー入り） ──
function generateSoukaiTemplate() {
  var template = getTemplate('定期総会');
  var doc = DocumentApp.openById(DOC_ID);
  var body = doc.getBody();
  body.clear();
  body.setMarginTop(56.7); body.setMarginBottom(56.7);
  body.setMarginLeft(56.7); body.setMarginRight(56.7);

  var placeholders = {};
  for (var key in template.defaults) {
    placeholders[key] = '{{' + key + '}}';
  }
  renderSections(body, template.sections, placeholders);
  doc.saveAndClose();

  Logger.log('定期総会 テンプレ版 生成完了: ' + doc.getUrl());
  return doc.getUrl();
}

// ── PDF出力 ──
function exportToPdf() {
  var doc = DocumentApp.openById(DOC_ID);
  var blob = doc.getAs('application/pdf');
  blob.setName(doc.getName() + '.pdf');
  var file = DriveApp.createFile(blob);
  Logger.log('PDF生成完了: ' + file.getUrl());
  return file.getUrl();
}
