// ─── QALIB CONFIG (ШАБЛОН) ───────────────────────────────────
// В библиотеке QALib PropertiesService читает СВОИ script properties
// (общие для всех пользователей). Нужен только ID общей корневой папки
// результатов. Скопируй в Properties.gs (он в .gitignore) и подставь
// значение, либо один раз запусти setupQalibConfig().
//   QA_RESULTS_ROOT_FOLDER_ID — ID корневой папки Диска для QA-отчётов

function setupQalibConfig() {
  PropertiesService.getScriptProperties().setProperties({
    QA_RESULTS_ROOT_FOLDER_ID: 'YOUR_RESULTS_ROOT_FOLDER_ID'
  });
}
