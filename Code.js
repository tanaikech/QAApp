function onOpen() {
  SpreadsheetApp.getUi().createMenu("sample").addItem("Open sidebar", "openSidebar").addToUi();
}

function openSidebar() {
  const html = HtmlService.createTemplateFromFile("index");
  html.javascript = javascript_.toString().match(/^function javascript_\(\) {([\s\S\w]*)}$/)[1];
  SpreadsheetApp.getUi().showSidebar(html.evaluate().setTitle("QAApp"));
}

function manageChunks() {
  const QA = new QAApp();
  QA.manageChunks();
}

function clearCorpus() {
  const QA = new QAApp();
  QA.clearCorpus();
}

function getChunksAndUpdateDataSheet() {
  const QA = new QAApp();
  QA.getChunksAndUpdateDataSheet();
}

function generateAnswer(prompt = "") {
  const QA = new QAApp();
  return QA.generateAnswer({ prompt });
}
