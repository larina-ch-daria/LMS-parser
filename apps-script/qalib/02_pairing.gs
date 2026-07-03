// ============================================================
// STAGE 3 — CONFIG DIFF
// Читает LMS_PAIRS + RAW_API, сравнивает конфиги каждой пары
// уровней и классифицирует расхождения. Пишет CONFIG_DIFFS
// и PAIR_SUMMARY.
// ============================================================

function buildConfigDiffsFromLmsPairs(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  const pairsSheet = ss.getSheetByName('LMS_PAIRS');
  const rawSheet = ss.getSheetByName('RAW_API');

  if (!pairsSheet) throw new Error('Лист LMS_PAIRS не найден');
  if (!rawSheet) throw new Error('Лист RAW_API не найден');

  const diffsSheet = ss.getSheetByName('CONFIG_DIFFS') || ss.insertSheet('CONFIG_DIFFS');
  const summarySheet = ss.getSheetByName('PAIR_SUMMARY') || ss.insertSheet('PAIR_SUMMARY');

  const rawMap = buildRawApiMapForPairs_(rawSheet);
  const pairs = readLmsPairsCfgDiff_(pairsSheet);

  logDuplicateTaskIdsCfgDiff_(pairs);

  const diffRows = [];
  const summaryRows = [];

  pairs.forEach(pair => {
    const ruPayload = pair.ru_mainLevelId ? rawMap[pair.ru_mainLevelId] : null;
    const azPayload = pair.az_mainLevelId ? rawMap[pair.az_mainLevelId] : null;

    const pairDiffs = [];

    if (!pair.ru_mainLevelId || !pair.az_mainLevelId) {
      pairDiffs.push(makeDiffRowFromPairCfgDiff_(
        pair,
        '',
        'PAIR_MISSING',
        pair.ru_mainLevelId || '',
        pair.az_mainLevelId || '',
        'Одна из сторон отсутствует в LMS_PAIRS'
      ));
    } else if (!ruPayload || !azPayload) {
      pairDiffs.push(makeDiffRowFromPairCfgDiff_(
        pair,
        '',
        'RAW_MISSING',
        ruPayload ? 'FOUND' : 'NOT_FOUND',
        azPayload ? 'FOUND' : 'NOT_FOUND',
        'Для одной из сторон нет JSON в RAW_API'
      ));
    } else {
      comparePairedPayloadsCfgDiff_(pair, ruPayload, azPayload, pairDiffs);
    }

    const counts = summarizeDiffTypesCfgDiff_(pairDiffs);

    let overallStatus = 'OK';
    if (counts.structure_alert_count > 0) overallStatus = 'HAS_STRUCTURE_ALERTS';
    else if (counts.pair_missing_count > 0 || counts.raw_missing_count > 0) overallStatus = 'INCOMPLETE';
    else if (counts.text_diff_count > 0) overallStatus = 'TEXT_ONLY_DIFFS';

    summaryRows.push([
      pair.taskSeq,
      pair.orderInTask,
      pair.pairKey,
      pair.originalPairKey,
      pair.taskId,
      pair.taskTitle,
      pair.levelKind,
      pair.ru_mainLevelId,
      pair.ru_levelTitle,
      pair.az_mainLevelId,
      pair.az_levelTitle,
      pair.pairStatus || '',
      pair.pairConfidence || '',
      pair.pairingNote || '',
      counts.match_count,
      counts.text_diff_count,
      counts.structure_alert_count,
      counts.pair_missing_count,
      counts.raw_missing_count,
      overallStatus
    ]);

    pairDiffs.forEach(r => diffRows.push(r));
  });

  writeConfigDiffsSheetCfgDiff_(diffsSheet, diffRows);
  writePairSummarySheetCfgDiff_(summarySheet, summaryRows);
}

function readLmsPairsCfgDiff_(pairsSheet) {
  const values = pairsSheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h).trim());
  const rows = values.slice(1);

  const idx = indexMapCfgDiff_(headers, [
    'taskSeq',
    'orderInTask',
    'pairKey',

    'ru_taskId',
    'ru_taskTitle',
    'ru_mainLevelId',
    'ru_levelTitle',
    'ru_levelKind',

    'az_taskId',
    'az_taskTitle',
    'az_mainLevelId',
    'az_levelTitle',
    'az_levelKind',

    'pairStatus',
    'pairConfidence',
    'pairingNote'
  ]);

  return rows.map(r => {
    const ru_taskId = clean_(r[idx.ru_taskId]);
    const ru_taskTitle = clean_(r[idx.ru_taskTitle]);
    const ru_levelKind = clean_(r[idx.ru_levelKind]);

    const az_taskId = clean_(r[idx.az_taskId]);
    const az_taskTitle = clean_(r[idx.az_taskTitle]);
    const az_levelKind = clean_(r[idx.az_levelKind]);

    const taskSeq = clean_(r[idx.taskSeq]);
    const orderInTask = clean_(r[idx.orderInTask]);
    const ru_mainLevelId = clean_(r[idx.ru_mainLevelId]);
    const az_mainLevelId = clean_(r[idx.az_mainLevelId]);

    const taskId = ru_taskId || az_taskId;
    const originalPairKey = clean_(r[idx.pairKey]);

    const stablePairKey = buildStablePairKeyCfgDiff_({
      taskId: taskId,
      ru_mainLevelId: ru_mainLevelId,
      az_mainLevelId: az_mainLevelId,
      fallbackPairKey: originalPairKey
    });

    return {
      taskSeq: taskSeq,
      orderInTask: orderInTask,
      pairKey: stablePairKey,
      originalPairKey: originalPairKey,

      taskId: taskId,
      taskTitle: ru_taskTitle || az_taskTitle,
      levelKind: ru_levelKind || az_levelKind,

      ru_mainLevelId: ru_mainLevelId,
      ru_levelTitle: clean_(r[idx.ru_levelTitle]),

      az_mainLevelId: az_mainLevelId,
      az_levelTitle: clean_(r[idx.az_levelTitle]),

      pairStatus: clean_(r[idx.pairStatus]),
      pairConfidence: clean_(r[idx.pairConfidence]),
      pairingNote: clean_(r[idx.pairingNote])
    };
  });
}

function buildStablePairKeyCfgDiff_(obj) {
  const taskId = clean_(obj.taskId);
  const ruId = clean_(obj.ru_mainLevelId);
  const azId = clean_(obj.az_mainLevelId);
  const fallbackPairKey = clean_(obj.fallbackPairKey);

  return [
    taskId || 'NO_TASK',
    ruId || 'NO_RU',
    azId || 'NO_AZ',
    fallbackPairKey || 'NO_PAIRKEY'
  ].join(' | ');
}

function comparePairedPayloadsCfgDiff_(pair, ruPayload, azPayload, outRows) {
  const ru = ruPayload && ruPayload.data ? ruPayload.data : {};
  const az = azPayload && azPayload.data ? azPayload.data : {};

  compareFieldFromPairCfgDiff_(pair, outRows, 'type', ru.type, az.type, false);
  compareFieldFromPairCfgDiff_(pair, outRows, 'levelScore', ru.levelScore, az.levelScore, false);
  compareFieldFromPairCfgDiff_(pair, outRows, 'isMulti', ru.isMulti, az.isMulti, false);
  compareFieldFromPairCfgDiff_(pair, outRows, 'isAutocheck', ru.isAutocheck, az.isAutocheck, false);
  compareFieldFromPairCfgDiff_(pair, outRows, 'treePosition', ru.treePosition, az.treePosition, false);

  compareFieldFromPairCfgDiff_(pair, outRows, 'title', ru.title, az.title, true);

  compareFieldFromPairCfgDiff_(
    pair, outRows,
    'children.length',
    safeArray_(ru.children).length,
    safeArray_(az.children).length,
    false
  );

  compareFieldFromPairCfgDiff_(
    pair, outRows,
    'config.type',
    getPath_(ru, 'config.type'),
    getPath_(az, 'config.type'),
    false
  );

  compareFieldFromPairCfgDiff_(
    pair, outRows,
    'config.bigSize',
    getPath_(ru, 'config.bigSize'),
    getPath_(az, 'config.bigSize'),
    false
  );

  compareFieldFromPairCfgDiff_(
    pair, outRows,
    'config.mechanic.problemType',
    firstNonEmpty_([
      getPath_(ru, 'config.mechanic.problemType'),
      getPath_(ru, 'children.0.config.mechanic.problemType'),
      getPath_(ru, 'children.0.children.0.config.mechanic.problemType')
    ]),
    firstNonEmpty_([
      getPath_(az, 'config.mechanic.problemType'),
      getPath_(az, 'children.0.config.mechanic.problemType'),
      getPath_(az, 'children.0.children.0.config.mechanic.problemType')
    ]),
    false
  );

  compareFieldFromPairCfgDiff_(
    pair, outRows,
    'config.mechanic.needRandomize',
    firstNonEmpty_([
      getPath_(ru, 'config.mechanic.needRandomize'),
      getPath_(ru, 'children.0.config.mechanic.needRandomize'),
      getPath_(ru, 'children.0.children.0.config.mechanic.needRandomize')
    ]),
    firstNonEmpty_([
      getPath_(az, 'config.mechanic.needRandomize'),
      getPath_(az, 'children.0.config.mechanic.needRandomize'),
      getPath_(az, 'children.0.children.0.config.mechanic.needRandomize')
    ]),
    false
  );

  compareFieldFromPairCfgDiff_(
    pair, outRows,
    'description',
    extractDescriptionCfgDiff_(ru),
    extractDescriptionCfgDiff_(az),
    true
  );

  compareFieldFromPairCfgDiff_(
    pair, outRows,
    'answers.length',
    extractAnswersCfgDiff_(ru).length,
    extractAnswersCfgDiff_(az).length,
    false
  );

  compareFieldFromPairCfgDiff_(
    pair, outRows,
    'images.length',
    extractImagesCfgDiff_(ru).length,
    extractImagesCfgDiff_(az).length,
    false
  );

  compareFieldFromPairCfgDiff_(
    pair, outRows,
    'google_slides_id',
    extractSlidesIdCfgDiff_(clean_(getPath_(ru, 'config.content'))),
    extractSlidesIdCfgDiff_(clean_(getPath_(az, 'config.content'))),
    true
  );

  compareMessagesCfgDiff_(pair, outRows, ru, az);
  compareAnswersByIndexCfgDiff_(pair, outRows, ru, az);
}

function compareMessagesCfgDiff_(pair, outRows, ru, az) {
  const keys = [
    'congrat',
    'fail',
    'hint',
    'question',
    'description',
    'title',
    'msgs',
    'congrat_repeat'
  ];

  keys.forEach(key => {
    compareFieldFromPairCfgDiff_(
      pair,
      outRows,
      'messages.' + key,
      getPath_(ru, 'messages.' + key),
      getPath_(az, 'messages.' + key),
      true
    );
  });
}

function compareAnswersByIndexCfgDiff_(pair, outRows, ru, az) {
  const ruAnswers = extractAnswersCfgDiff_(ru);
  const azAnswers = extractAnswersCfgDiff_(az);
  const maxLen = Math.max(ruAnswers.length, azAnswers.length);

  for (let i = 0; i < maxLen; i++) {
    const r = ruAnswers[i] || {};
    const a = azAnswers[i] || {};

    compareFieldFromPairCfgDiff_(pair, outRows, 'answers[' + i + '].type', r.type, a.type, false);
    compareFieldFromPairCfgDiff_(pair, outRows, 'answers[' + i + '].isCorrect', r.isCorrect, a.isCorrect, false);
    compareFieldFromPairCfgDiff_(pair, outRows, 'answers[' + i + '].text', r.text, a.text, true);
  }
}

function compareFieldFromPairCfgDiff_(pair, outRows, path, ruValue, azValue, textDifferenceAllowed) {
  const ruNorm = normalizeCompareValueCfgDiff_(ruValue);
  const azNorm = normalizeCompareValueCfgDiff_(azValue);

  let diffType = 'MATCH';
  let note = '';

  if (ruNorm === azNorm) {
    diffType = 'MATCH';
  } else if (textDifferenceAllowed) {
    diffType = 'TEXT_DIFF_OK';
  } else {
    diffType = 'STRUCTURE_DIFF_ALERT';
    note = 'Проверь расхождение';
  }

  outRows.push(makeDiffRowFromPairCfgDiff_(pair, path, diffType, ruNorm, azNorm, note));
}

function makeDiffRowFromPairCfgDiff_(pair, path, diffType, ruValue, azValue, note) {
  return [
    pair.taskSeq,
    pair.orderInTask,
    pair.pairKey,
    pair.originalPairKey,
    pair.taskId,
    pair.taskTitle,
    pair.levelKind,
    pair.ru_mainLevelId,
    pair.ru_levelTitle,
    pair.az_mainLevelId,
    pair.az_levelTitle,
    pair.pairStatus || '',
    pair.pairConfidence || '',
    pair.pairingNote || '',
    path,
    diffType,
    safeJsonStringify_(ruValue),
    safeJsonStringify_(azValue),
    note
  ];
}

function summarizeDiffTypesCfgDiff_(rows) {
  const out = {
    match_count: 0,
    text_diff_count: 0,
    structure_alert_count: 0,
    pair_missing_count: 0,
    raw_missing_count: 0
  };

  rows.forEach(r => {
    const type = r[15];
    if (type === 'MATCH') out.match_count++;
    else if (type === 'TEXT_DIFF_OK') out.text_diff_count++;
    else if (type === 'STRUCTURE_DIFF_ALERT') out.structure_alert_count++;
    else if (type === 'PAIR_MISSING') out.pair_missing_count++;
    else if (type === 'RAW_MISSING') out.raw_missing_count++;
  });

  return out;
}

function writeConfigDiffsSheetCfgDiff_(sheet, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 19).setValues([[
    'taskSeq',
    'orderInTask',
    'pairKey',
    'original_pairKey',
    'taskId',
    'taskTitle',
    'levelKind',
    'ru_mainLevelId',
    'ru_levelTitle',
    'az_mainLevelId',
    'az_levelTitle',
    'pairStatus',
    'pairConfidence',
    'pairingNote',
    'field_path',
    'diff_type',
    'ru_value',
    'az_value',
    'note'
  ]]);

  if (!rows.length) return;
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function writePairSummarySheetCfgDiff_(sheet, rows) {
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 20).setValues([[
    'taskSeq',
    'orderInTask',
    'pairKey',
    'original_pairKey',
    'taskId',
    'taskTitle',
    'levelKind',
    'ru_mainLevelId',
    'ru_levelTitle',
    'az_mainLevelId',
    'az_levelTitle',
    'pairStatus',
    'pairConfidence',
    'pairingNote',
    'match_count',
    'text_diff_count',
    'structure_alert_count',
    'pair_missing_count',
    'raw_missing_count',
    'overall_status'
  ]]);

  if (!rows.length) return;
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function extractDescriptionCfgDiff_(data) {
  return firstNonEmpty_([
    getPath_(data, 'children.0.children.0.config.mechanic.description'),
    getPath_(data, 'children.0.config.mechanic.description'),
    getPath_(data, 'config.mechanic.description'),
    getPath_(data, 'quiz3Data.problem.description'),
    getPath_(data, 'quiz3Data.data.template')
  ]);
}

function extractAnswersCfgDiff_(data) {
  return safeArray_(firstNonEmpty_([
    getPath_(data, 'children.0.children.0.config.mechanic.answers'),
    getPath_(data, 'children.0.config.mechanic.answers'),
    getPath_(data, 'config.mechanic.answers')
  ]));
}

function extractImagesCfgDiff_(data) {
  return safeArray_(firstNonEmpty_([
    getPath_(data, 'children.0.children.0.config.mechanic.images'),
    getPath_(data, 'children.0.config.mechanic.images'),
    getPath_(data, 'config.mechanic.images'),
    getPath_(data, 'comics')
  ]));
}

function extractSlidesIdCfgDiff_(html) {
  const match = clean_(html).match(/presentation\/d\/([^/'"\s]+)/i);
  return match ? match[1] : '';
}

function normalizeCompareValueCfgDiff_(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value.trim();
  if (typeof value === 'object') return JSON.stringify(value);
  return String(value);
}

function indexMapCfgDiff_(headers, names) {
  const map = {};
  names.forEach(name => {
    map[name] = headers.indexOf(name);
    if (map[name] === -1) {
      throw new Error('Не найдена колонка: ' + name);
    }
  });
  return map;
}

function logDuplicateTaskIdsCfgDiff_(pairs) {
  const seen = {};

  pairs.forEach(p => {
    const taskId = clean_(p.taskId);
    if (!taskId) return;

    if (!seen[taskId]) seen[taskId] = [];
    seen[taskId].push({
      taskSeq: p.taskSeq,
      orderInTask: p.orderInTask,
      pairKey: p.pairKey,
      ru_mainLevelId: p.ru_mainLevelId,
      az_mainLevelId: p.az_mainLevelId,
      pairStatus: p.pairStatus,
      pairConfidence: p.pairConfidence
    });
  });

  Object.keys(seen).forEach(taskId => {
    if (seen[taskId].length > 1) {
      Logger.log('Duplicate taskId in CONFIG_DIFFS source: ' + taskId + ' -> ' + JSON.stringify(seen[taskId]));
    }
  });
}
