// ============================================================
// PERSONAL TABLE · LOADER (стадия 1)
// OAuth к API alg.academy, батчевая загрузка каждого уровня,
// запись в RAW_API и Results. Rate limiting + ретраи.
// Конфиг (SCRIPT_URL / CLIENT_ID / CLIENT_SECRET / …) — в Script Properties.
// ============================================================

function runLoader(silent) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    throw new Error('Скрипт уже выполняется.');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var config = getConfig_();
    var sheets = prepareLoaderSheets_();

    logRunStart_(sheets.resultsSheet);

    ss.toast('Загружаю список уровней...', '1/5 Load levels', -1);
    var levelIds = loadLevels_(config);
    if (!levelIds || !Array.isArray(levelIds) || levelIds.length === 0) {
      throw new Error('Levels list is empty');
    }

    ss.toast('Авторизация...', '1/5 Load levels', -1);
    var code = authorize_(config);
    var accessToken = getToken_(config, code);

    ss.toast('Загружаю уровни: 0 / ' + levelIds.length, '1/5 Load levels', -1);

    var resultsBuffer = [];
    var rawApiBuffer = [];
    var batchSize = 5;

    for (var i = 0; i < levelIds.length; i++) {
      var levelId = levelIds[i];

      if (i % 10 === 0) {
        ss.toast(
          'Загружаю уровни: ' + i + ' / ' + levelIds.length,
          '1/5 Load levels',
          -1
        );
      }

      Utilities.sleep(5000);

      var result = fetchLevelWithRetry_(config, levelId, accessToken, 3);

      resultsBuffer.push([
        new Date(),
        levelId,
        result.status,
        buildShortResultInfo_(result.data)
      ]);

      // Полный JSON только в RAW_API. Слишком большой уровень усекаем маркером.
      var rawJson = safeStringify_(result.data);
      var MAX_CELL = 49000;
      if (rawJson.length > MAX_CELL) {
        Logger.log('Level ' + levelId + ' payload too big: ' + rawJson.length + ' chars, truncating');
        rawJson = JSON.stringify({
          __TRUNCATED__: true,
          __original_length__: rawJson.length,
          __note__: 'Уровень слишком большой для одной ячейки (>50k). Проверить вручную.',
          id: (result.data && result.data.data && result.data.data.id) || levelId
        });
      }
      rawApiBuffer.push([
        new Date(),
        levelId || '',
        result.status || '',
        '',
        rawJson
      ]);
      Logger.log('Saved level: ' + levelId + ', status=' + result.status);

      if (
        resultsBuffer.length >= batchSize ||
        rawApiBuffer.length >= batchSize ||
        i === levelIds.length - 1
      ) {
        appendRowsBatch_(sheets.resultsSheet, resultsBuffer);
        appendRowsBatch_(sheets.rawApiSheet, rawApiBuffer);

        resultsBuffer = [];
        rawApiBuffer = [];

        SpreadsheetApp.flush();
      }

      if (i < levelIds.length - 1) {
        Logger.log('Next level: ' + levelIds[i + 1]);
      } else {
        Logger.log('DONE');
      }
    }

    if (!silent) {
      SpreadsheetApp.getUi().alert(
        'Готово',
        'Обработка завершена. Уровней обработано: ' + levelIds.length,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    ss.toast('Уровней загружено: ' + levelIds.length, '1/5 Load levels ✓', 5);

  } catch (err) {
    Logger.log('ERROR: ' + err.message);

    if (!silent) {
      SpreadsheetApp.getUi().alert(
        'Ошибка',
        err.message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }

    throw err;
  } finally {
    lock.releaseLock();
  }
}

function getConfig_() {
  var props = PropertiesService.getScriptProperties();

  var config = {
    script_url: props.getProperty('SCRIPT_URL'),
    client_id: props.getProperty('CLIENT_ID'),
    client_secret: props.getProperty('CLIENT_SECRET'),
    redirect_uri: props.getProperty('REDIRECT_URI') || 'https://api.alg.academy',
    scope: props.getProperty('SCOPE') || 'level',
    state: props.getProperty('STATE') || '1'
  };

  if (!config.script_url) throw new Error('Не задан SCRIPT_URL. Запустите "Setup config".');
  if (!config.client_id) throw new Error('Не задан CLIENT_ID. Запустите "Setup config".');
  if (!config.client_secret) throw new Error('Не задан CLIENT_SECRET. Запустите "Setup config".');

  return config;
}

function logRunStart_(sheet) {
  sheet.appendRow([new Date(), '__RUN_START__', '', '']);
}

function toFormBody_(obj) {
  var parts = [];
  for (var key in obj) {
    if (Object.prototype.hasOwnProperty.call(obj, key)) {
      parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(String(obj[key])));
    }
  }
  return parts.join('&');
}

function loadLevels_(config) {
  var response = withFetchRetry_('Load levels (webhook)', function() {
    return UrlFetchApp.fetch(config.script_url + '?action=levels', {
      method: 'get',
      muteHttpExceptions: true
    });
  });

  var status = response.getResponseCode();
  var text = response.getContentText();

  if (status !== 200) {
    throw new Error('Response is not JSON. Status: ' + status + '. Body: ' + text);
  }

  var json;
  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new Error('Response is not JSON. Status: ' + status + '. Body: ' + text);
  }

  if (!json.items || !Array.isArray(json.items)) {
    throw new Error('No level IDs received. Body: ' + JSON.stringify(json));
  }

  if (json.items.length === 0) {
    throw new Error('Levels list is empty');
  }

  Logger.log('Loaded levels: ' + json.items.length);
  Logger.log('First level: ' + json.items[0]);

  return json.items;
}

function authorize_(config) {
  var payload = toFormBody_({
    client_id: config.client_id,
    client_secret: config.client_secret,
    response_type: 'code',
    scope: 'level',
    redirect_uri: 'https://api.alg.academy',
    state: '1'
  });

  var response = UrlFetchApp.fetch('https://api.alg.academy/oauth/authorize', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: payload,
    muteHttpExceptions: true
  });

  var text = response.getContentText();
  var json;

  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new Error('Authorize response is not JSON. Body: ' + text);
  }

  if (!json.data || !json.data.code) {
    throw new Error('No code in authorize response. Body: ' + text);
  }

  Logger.log('Authorize code saved');

  return json.data.code;
}

function getToken_(config, code) {
  var payload = toFormBody_({
    client_id: config.client_id,
    client_secret: config.client_secret,
    response_type: 'code',
    scope: config.scope,
    redirect_uri: config.redirect_uri,
    state: config.state,
    grant_type: 'authorization_code',
    code: code
  });

  var response = UrlFetchApp.fetch('https://api.alg.academy/oauth/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: payload,
    muteHttpExceptions: true
  });

  var text = response.getContentText();
  var json;

  try {
    json = JSON.parse(text);
  } catch (e) {
    throw new Error('Token response is not JSON. Body: ' + text);
  }

  if (!json.data || !json.data.access_token) {
    throw new Error('No access_token in token response. Body: ' + text);
  }

  var rawToken = json.data.access_token;
  var base64Token = Utilities.base64Encode(rawToken);

  Logger.log('Base64 token ready');

  return base64Token;
}

function fetchLevel_(config, levelId, accessToken) {
  var response = UrlFetchApp.fetch(
    'https://api.alg.academy/level/v1/default/view/' + encodeURIComponent(levelId),
    {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + accessToken,
        Accept: 'application/json'
      },
      muteHttpExceptions: true
    }
  );

  var status = response.getResponseCode();
  var text = response.getContentText();

  var responseData;
  try {
    responseData = JSON.parse(text);
  } catch (e) {
    responseData = { raw: text };
  }

  return {
    status: status,
    data: responseData
  };
}

function saveLevelResult_(config, levelId, status, data) {
  var payload = JSON.stringify({
    id: levelId,
    status: status,
    data: data
  });

  var response = UrlFetchApp.fetch(config.script_url, {
    method: 'post',
    contentType: 'application/json',
    payload: payload,
    muteHttpExceptions: true
  });

  Logger.log(
    'saveLevelResult_ response: ' +
    response.getResponseCode() + ' ' + response.getContentText()
  );
}

function saveRawApiRowDirect_(levelId, status, data, action) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('RAW_API');

  if (!sheet) {
    sheet = ss.insertSheet('RAW_API');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['timestamp', 'level_id', 'status', 'action', 'data']);
  }

  var dataJson;
  try {
    dataJson = JSON.stringify(data);
  } catch (e) {
    dataJson = JSON.stringify({ raw: String(data) });
  }

  sheet.appendRow([new Date(), levelId || '', status || '', action || '', dataJson]);
}

// ─── батчевая загрузка / ретраи ──────────────────────────────

function prepareLoaderSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var resultsSheet = ss.getSheetByName('Results');
  if (!resultsSheet) {
    resultsSheet = ss.insertSheet('Results');
  }
  if (resultsSheet.getLastRow() === 0) {
    resultsSheet.appendRow(['timestamp', 'level_id', 'status', 'info']);
  }

  var rawApiSheet = ss.getSheetByName('RAW_API');
  if (!rawApiSheet) {
    rawApiSheet = ss.insertSheet('RAW_API');
  }
  if (rawApiSheet.getLastRow() === 0) {
    rawApiSheet.appendRow(['timestamp', 'level_id', 'status', 'action', 'data']);
  }

  return {
    ss: ss,
    resultsSheet: resultsSheet,
    rawApiSheet: rawApiSheet
  };
}

function safeStringify_(data) {
  try {
    return JSON.stringify(data);
  } catch (e) {
    return JSON.stringify({ raw: String(data) });
  }
}

function fetchLevelWithRetry_(config, levelId, accessToken, maxAttempts) {
  var attempt = 0;
  var lastError = null;

  while (attempt < maxAttempts) {
    attempt++;

    try {
      var result = fetchLevel_(config, levelId, accessToken);

      if (result.status === 429 || result.status >= 500) {
        if (attempt < maxAttempts) {
          var waitMs = attempt * 5000;
          Logger.log(
            'Retry for level ' + levelId +
            ', attempt ' + attempt +
            ', status=' + result.status +
            ', wait=' + waitMs + 'ms'
          );
          Utilities.sleep(waitMs);
          continue;
        }
      }

      return result;

    } catch (e) {
      lastError = e;

      if (attempt < maxAttempts) {
        var backoff = attempt * 5000;
        Logger.log(
          'Fetch exception for level ' + levelId +
          ', attempt ' + attempt +
          ': ' + e.message +
          '. Waiting ' + backoff + 'ms'
        );
        Utilities.sleep(backoff);
      }
    }
  }

  if (lastError) throw lastError;

  throw new Error('Failed to fetch level ' + levelId);
}

function withFetchRetry_(label, fn, maxAttempts) {
  maxAttempts = maxAttempts || 3;
  for (var attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return fn();
    } catch (e) {
      var msg = String(e && e.message || e);
      if (attempt < maxAttempts) {
        var waitMs = attempt * 8000;
        Logger.log('Retry ' + label + ' attempt ' + attempt + ': ' + msg + ' (wait ' + waitMs + 'ms)');
        Utilities.sleep(waitMs);
        continue;
      }
      throw new Error(
        label + ': превышена квота Google или временная ошибка после ' + maxAttempts +
        ' попыток. Подожди 1-2 минуты и запусти пайплайн снова. (' + msg + ')'
      );
    }
  }
}

function appendRowsBatch_(sheet, rows) {
  if (!rows || !rows.length) return;

  var startRow = sheet.getLastRow() + 1;
  var startCol = 1;
  var numRows = rows.length;
  var numCols = rows[0].length;

  sheet.getRange(startRow, startCol, numRows, numCols).setValues(rows);
}

function buildShortResultInfo_(data) {
  if (!data) return '';

  try {
    var root = data.data || data;
    var parts = [];

    if (root.id != null) parts.push('id=' + root.id);
    if (root.title) parts.push('title=' + root.title);
    if (root.type) parts.push('type=' + root.type);
    if (root.updatedAt) parts.push('updatedAt=' + root.updatedAt);

    return parts.join('; ');
  } catch (e) {
    return '';
  }
}
