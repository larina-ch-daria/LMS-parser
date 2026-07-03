// ============================================================
// QA DRIVE
// Создание QA-таблицы в личной папке владельца на Google Drive.
// Определение владельца из названия таблицы, создание папки,
// прогресс-тосты, диалог с результатом.
// ============================================================

function updateQaProgress_(ss, message, step, total) {
  let prefix = 'QA';
  if (step && total) prefix += ' [' + step + '/' + total + ']';
  ss.toast(message, prefix, 5);
  Logger.log(prefix + ' ' + message);
}

function showQaResultDialog_(url, fileName) {
  const safeUrl = String(url || '');
  const safeFileName = String(fileName || 'Открыть таблицу')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;padding:16px 18px;">' +
      '<div style="font-size:16px;font-weight:700;margin-bottom:10px;">Готово</div>' +
      '<div style="margin-bottom:12px;">QA-таблица успешно создана.</div>' +
      '<div style="margin-bottom:14px;">' +
        '<a href="' + safeUrl + '" target="_blank" ' +
        'style="color:#1155cc;text-decoration:underline;font-weight:600;">' +
          safeFileName +
        '</a>' +
      '</div>' +
      '<div>' +
        '<button onclick="google.script.host.close()" ' +
        'style="padding:8px 14px;border:none;border-radius:8px;cursor:pointer;">Закрыть</button>' +
      '</div>' +
    '</div>'
  ).setWidth(420).setHeight(180);

  SpreadsheetApp.getUi().showModalDialog(html, 'Результат');
}

// ID корневой папки результатов хранится в Script Properties, не в коде.
const QA_RESULTS_ROOT_FOLDER_ID_ =
  PropertiesService.getScriptProperties().getProperty('QA_RESULTS_ROOT_FOLDER_ID') ||
  'YOUR_RESULTS_ROOT_FOLDER_ID';

function getOwnerNameFromSpreadsheetTitle_(ss) {
  const title = ss.getName();
  const matches = title.match(/\(([^()]+)\)/g);
  if (!matches || !matches.length) {
    throw new Error(
      'Не удалось определить имя владельца из названия таблицы. ' +
      'Ожидаю имя в скобках, например: "Название таблицы (Имя)".'
    );
  }
  const last = matches[matches.length - 1];
  return String(last || '').replace(/[()]/g, '').trim();
}

function getOrCreatePersonalResultsFolder_(ownerName) {
  const root = DriveApp.getFolderById(QA_RESULTS_ROOT_FOLDER_ID_);
  const folderName = 'Результаты ' + ownerName;
  const folders = root.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return root.createFolder(folderName);
}

function createQaSpreadsheetInPersonalFolder_(ss, lessonName) {
  const ownerName = getOwnerNameFromSpreadsheetTitle_(ss);
  updateQaProgress_(ss, 'Определён владелец: ' + ownerName, 1, 8);
  const folder = getOrCreatePersonalResultsFolder_(ownerName);
  updateQaProgress_(ss, 'Найдена или создана личная папка', 2, 8);
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const fileName = sanitizeSpreadsheetFileName_(lessonName + '_' + dateStr);
  const newSs = SpreadsheetApp.create(fileName);
  const file = DriveApp.getFileById(newSs.getId());
  file.moveTo(folder);
  updateQaProgress_(ss, 'Создан новый файл: ' + fileName, 3, 8);
  return { spreadsheet: newSs, url: newSs.getUrl(), fileName: fileName };
}

function sanitizeSpreadsheetFileName_(name) {
  let s = String(name || '').trim();
  s = s.replace(/[\\\/:*?"<>|#\[\]]/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();
  if (!s) s = 'QA_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (s.length > 180) s = s.substring(0, 180).trim();
  return s;
}

function removeDefaultSheetIfNeeded_(ss) {
  const sheets = ss.getSheets();
  if (sheets.length <= 1) return;
  const defaultSheet = sheets.find(sh => sh.getName() === 'Sheet1');
  if (defaultSheet) ss.deleteSheet(defaultSheet);
}

function sanitizeForFileNamePart_(s) {
  return String(s || '')
    .replace(/[\/\\:*?"<>|]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 50);
}

// Собирает "Курс_Урок_Язык" (дата добавится в createQaSpreadsheetInPersonalFolder_).
function buildQaFileBaseName_(lessonName, lmsLevelsData) {
  let course = '';
  if (lmsLevelsData && lmsLevelsData.length) {
    const withCourse = lmsLevelsData.find(r => clean_(r.courseTitle));
    if (withCourse) course = sanitizeForFileNamePart_(withCourse.courseTitle);
  }

  const lang = QA_CMP_LABEL_ && QA_CMP_LABEL_ !== QA_BASE_LABEL_
    ? QA_BASE_LABEL_ + '-' + QA_CMP_LABEL_
    : QA_BASE_LABEL_;

  return [course, lessonName, lang].filter(Boolean).join('_');
}
