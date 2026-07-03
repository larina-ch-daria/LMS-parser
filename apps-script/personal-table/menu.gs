// ============================================================
// PERSONAL TABLE · MENU
// Bound-скрипт личной таблицы. Меню QA Pipeline и тонкие обёртки
// над библиотекой QALib. Тяжёлая логика — в QALib, стадия 1
// (загрузка из API) — в loader.gs этой же таблицы.
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("QA Pipeline")
    .addItem("Запустить пайплайн", "runFullPipeline")
    .addSeparator()
    .addItem("Построить заново репорт (без новой выгрузки, при Drive error)", "buildLevelQAView")
    .addSeparator()
    .addItem("Проверка одного урока", "buildLevelQASingleView")
    .addToUi();
}

function runFullPipeline() {
  runLoader(true);
  QALib.runFullPipeline();
}

function runActualLevelLoader() { runLoader(false); }
function buildLmsPairs()                { QALib.buildLmsPairs(); }
function buildConfigDiffsFromLmsPairs() { QALib.buildConfigDiffsFromLmsPairs(); }
function buildRawParsedNodes()          { QALib.buildRawParsedNodes(); }
function buildLevelQAView()             { QALib.buildLevelQAView(); }

function buildLevelQASingleView() {
  runLoader(true);
  QALib.buildLevelQASingleView();
}
