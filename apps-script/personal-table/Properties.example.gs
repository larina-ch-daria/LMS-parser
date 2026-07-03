// ─── ЛИЧНАЯ ТАБЛИЦА · CONFIG (ШАБЛОН) ────────────────────────
// Читается loader.gs через Script Properties личной таблицы.
// Секреты в репозиторий не коммитим. Скопируй в Properties.gs
// (он в .gitignore) и подставь значения, либо запусти setupConfig().
//   SCRIPT_URL     — URL персонального вебхука (Web App)
//   CLIENT_ID / CLIENT_SECRET — OAuth для API alg.academy
//   REDIRECT_URI   — дефолт https://api.alg.academy
//   SCOPE          — дефолт level
//   STATE          — дефолт 1

function setupConfig() {
  PropertiesService.getScriptProperties().setProperties({
    SCRIPT_URL:    'https://script.google.com/macros/s/YOUR_WEBHOOK_DEPLOY_ID/exec',
    CLIENT_ID:     'YOUR_OAUTH_CLIENT_ID',
    CLIENT_SECRET: 'YOUR_OAUTH_CLIENT_SECRET',
    REDIRECT_URI:  'https://api.alg.academy',
    SCOPE:         'level',
    STATE:         '1'
  });
}
