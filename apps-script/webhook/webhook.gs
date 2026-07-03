// ============================================================
// WEBHOOK (Apps Script Web App)
// Отдельный проект, развёрнутый как Web App. Tampermonkey шлёт сюда
// POST с уровнями урока; данные пишутся в LMS_LEVELS / RAW_API
// личной таблицы. GET ?action=levels отдаёт список level_id.
// ============================================================

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';

  if (action === 'levels') {
    return getLevelsJson_();
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      actions: ['?action=levels']
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    const action = payload.action || '';

    if (action === 'saveLessonLevels') {
      return saveLessonLevels_(payload);
    }

    if (action === 'clearLessonLevels') {
      return clearLessonLevels_();
    }

    return saveRawApiPayload_(payload);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: err.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveRawApiPayload_(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('RAW_API') || ss.insertSheet('RAW_API');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['timestamp', 'level_id', 'status', 'action', 'data']);
  }

  sheet.appendRow([
    new Date(),
    payload.id || '',
    payload.status || '',
    payload.action || '',
    JSON.stringify(payload.data || payload)
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      handler: 'saveRawApiPayload_',
      action: payload.action || null
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function saveLessonLevels_(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('LMS_LEVELS') || ss.insertSheet('LMS_LEVELS');
  const rows = Array.isArray(payload.rows) ? payload.rows : [];

  ensureLessonLevelsHeader_(sheet);

  if (!rows.length) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'ok',
        handler: 'saveLessonLevels_',
        inserted: 0
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const values = rows.map(r => [
    new Date(),

    r.lessonId || '',
    r.taskId || '',
    r.track || '',
    r.taskTitle || '',
    r.levelKind || '',
    r.orderInTask || '',
    r.taskLevelId || '',
    r.mainLevelId || '',
    r.multiLevelId || '',
    r.parentTaskLevelId || '',
    r.parentMainLevelId || '',
    r.parentLevelTitle || '',
    r.levelUuid || '',
    r.levelTitle || '',
    r.pageUrl || '',

    r.lessonTitle || '',
    r.lessonGuid || '',
    r.lessonStatus || '',
    r.lessonNote || '',
    r.courseTitle || '',
    r.courseUrl || '',
    r.courseUuid || '',
    r.courseLanguage || '',
    r.courseLocale || '',
    r.lessonPositionInCourse || '',
    r.courseLessonsTotal || '',
    r.msoStatus || '',
    r.msoEnabled || '',
    r.publicName || '',
    r.hasPublicName || '',
    r.pageTitle || '',
    r.isBonus || '',
    r.isTheory || '',
    r.isQuiz || '',
    r.lessonMaterials || '',
    r.lessonVideoUrl || '',
  ]);

  sheet
    .getRange(sheet.getLastRow() + 1, 1, values.length, values[0].length)
    .setValues(values);

  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      handler: 'saveLessonLevels_',
      inserted: values.length
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function clearLessonLevels_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('LMS_LEVELS') || ss.insertSheet('LMS_LEVELS');

  sheet.clearContents();
  sheet.getRange(1, 1, 1, getLessonLevelsHeaders_().length).setValues([getLessonLevelsHeaders_()]);

  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      handler: 'clearLessonLevels_'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getLevelsJson_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Levels');

  if (!sheet) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Sheet "Levels" not found'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'ok',
        items: []
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  const items = values
    .flat()
    .map(String)
    .map(s => s.trim())
    .filter(Boolean);

  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      items: items
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── HELPERS FOR LMS_LEVELS ──────────────────────────────────

function getLessonLevelsHeaders_() {
  return [
    'timestamp',
    'lessonId',
    'taskId',
    'track',
    'taskTitle',
    'levelKind',
    'orderInTask',
    'taskLevelId',
    'mainLevelId',
    'multiLevelId',
    'parentTaskLevelId',
    'parentMainLevelId',
    'parentLevelTitle',
    'levelUuid',
    'levelTitle',
    'pageUrl',

    'lessonTitle',
    'lessonGuid',
    'lessonStatus',
    'lessonNote',
    'courseTitle',
    'courseUrl',
    'courseUuid',
    'courseLanguage',
    'courseLocale',
    'lessonPositionInCourse',
    'courseLessonsTotal',
    'msoStatus',
    'msoEnabled',
    'publicName',
    'hasPublicName',
    'pageTitle',
    'isBonus',
    'isTheory',
    'isQuiz',
    'lessonMaterials',
    'lessonVideoUrl',
  ];
}

function ensureLessonLevelsHeader_(sheet) {
  const requiredHeaders = getLessonLevelsHeaders_();

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return;
  }

  const currentWidth = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet.getRange(1, 1, 1, currentWidth).getValues()[0].map(function (x) {
    return String(x || '').trim();
  });

  const needsRewrite =
    currentHeaders.length < requiredHeaders.length ||
    requiredHeaders.some(function (h, i) {
      return currentHeaders[i] !== h;
    });

  if (!needsRewrite) return;

  const dataLastRow = sheet.getLastRow();
  const dataLastCol = sheet.getLastColumn();

  let oldData = [];
  if (dataLastRow > 1 && dataLastCol > 0) {
    oldData = sheet.getRange(2, 1, dataLastRow - 1, dataLastCol).getValues();
  }

  const headerIndexMap = {};
  currentHeaders.forEach(function (h, i) {
    if (h) headerIndexMap[h] = i;
  });

  sheet.clearContents();
  sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);

  if (!oldData.length) return;

  const migrated = oldData.map(function (row) {
    return requiredHeaders.map(function (header) {
      const oldIdx = Object.prototype.hasOwnProperty.call(headerIndexMap, header)
        ? headerIndexMap[header]
        : -1;
      return oldIdx >= 0 ? row[oldIdx] : '';
    });
  });

  sheet.getRange(2, 1, migrated.length, requiredHeaders.length).setValues(migrated);
}
