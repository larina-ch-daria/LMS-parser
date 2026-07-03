// ============================================================
// QALIB PIPELINE RUNNER
// Прогон стадий 2–5 (пары → конфиги → ноды → отчёт).
// Стадия 1 (загрузка из API) живёт в личной таблице — loader.gs.
// Вызывается из личной таблицы как QALib.runFullPipeline().
// ============================================================

function runFullPipeline(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  var steps = [
    { name: '2/5 Ищу уровням пары',  fn: function() { buildLmsPairs(ss); } },
    { name: '3/5 Изучаю конфиги',    fn: function() { buildConfigDiffsFromLmsPairs(ss); } },
    { name: '4/5 Разбираю теги',     fn: function() { buildRawParsedNodes(ss); } },
    { name: '5/5 Создаю репорт',     fn: function() { buildLevelQAView(ss); } }
  ];

  for (var i = 0; i < steps.length; i++) {
    var step = steps[i];
    ss.toast(step.name + '...', 'QA Pipeline', -1);
    try {
      step.fn();
      ss.toast(step.name + ' done', 'QA Pipeline', 3);
    } catch (err) {
      SpreadsheetApp.getUi().alert('Error: ' + step.name, err.message, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
  }

  ss.toast('All steps completed!', 'QA Pipeline', 5);
}
