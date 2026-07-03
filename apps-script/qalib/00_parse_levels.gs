// ============================================================
// STAGE 2 — PARSE & PAIR
// Читает LMS_LEVELS + RAW_API, определяет локаль каждой строки,
// выстраивает иерархию заданий и спаривает уровни базовой и
// сравниваемой локали. Пишет LMS_PAIRS.
// ============================================================

function buildLmsPairs(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  const levelsSheet = ss.getSheetByName('LMS_LEVELS');
  const rawSheet = ss.getSheetByName('RAW_API');
  if (!levelsSheet) throw new Error('Лист LMS_LEVELS не найден');
  if (!rawSheet) throw new Error('Лист RAW_API не найден');

  const outSheet = ss.getSheetByName('LMS_PAIRS') || ss.insertSheet('LMS_PAIRS');

  const levels = readLmsLevelsForPairs_(levelsSheet);
  logLocalesForPairs_(levels);
  if (!levels.length) throw new Error('LMS_LEVELS пуст');

  const pairLocales = getPairLocalesForLmsPairs_(levels);
  const baseLocale = pairLocales.baseLocale;
  const compareLocale = pairLocales.compareLocale;

  Logger.log('LMS_PAIRS locale pair: base=' + baseLocale + ', compare=' + compareLocale);

  const rawMap = buildRawApiMapForPairs_(rawSheet);

  const baseLevels = enrichHierarchyOrdinalsForPairs_(
    levels.filter(r => r.locale === baseLocale)
  );

  const compareLevels = enrichHierarchyOrdinalsForPairs_(
    levels.filter(r => r.locale === compareLocale)
  );

  if (!baseLevels.length) {
    throw new Error('Не найдены строки для base locale=' + baseLocale + ' в LMS_LEVELS');
  }

  if (!compareLevels.length) {
    throw new Error('Не найдены строки для compare locale=' + compareLocale + ' в LMS_LEVELS');
  }

  const taskOrdinals = collectTaskOrdinalsForPairs_(baseLevels, compareLevels);
  const outRows = [];

  taskOrdinals.forEach(taskOrdinal => {
    const baseGroup = baseLevels
      .filter(r => String(r.taskOrdinal) === String(taskOrdinal))
      .sort(sortByHierarchyForPairs_);

    const compareGroup = compareLevels
      .filter(r => String(r.taskOrdinal) === String(taskOrdinal))
      .sort(sortByHierarchyForPairs_);

    const aligned = alignLevelPairsWithinTaskSeq_(baseGroup, compareGroup, rawMap);

    aligned.forEach(pair => {
      const base = pair.ru || {};
      const compare = pair.az || {};

      const orderInTask =
        clean_(base.orderInTask) ||
        clean_(compare.orderInTask) ||
        '';

      const taskSeq =
        clean_(base.taskSeq) ||
        clean_(compare.taskSeq) ||
        String(taskOrdinal || '');

      const pairKey = buildStablePairKeyForLmsPairs_({
        taskSeq: taskSeq,
        orderInTask: orderInTask,
        ru_mainLevelId: base.mainLevelId,
        az_mainLevelId: compare.mainLevelId,
        ru_taskId: base.taskId,
        az_taskId: compare.taskId
      });

      outRows.push([
        taskSeq,
        orderInTask,
        pairKey,

        baseLocale,
        compareLocale,

        pair.pairStatus || '',
        pair.pairConfidence || '',
        pair.pairingScore != null ? pair.pairingScore : '',
        pair.pairingReason || '',
        pair.pairingNote || '',

        base.locale || '',
        base.taskId || '',
        base.taskTitle || '',
        base.lessonId || '',
        base.lessonTitle || '',
        base.levelKind || '',
        base.mainLevelId || '',
        base.levelTitle || '',
        base.levelUuid || '',
        base.pageUrl || '',
        base.track || '',
        base.lessonMaterials || '',
        base.lessonVideoUrl  || '',

        compare.locale || '',
        compare.taskId || '',
        compare.taskTitle || '',
        compare.lessonId || '',
        compare.lessonTitle || '',
        compare.levelKind || '',
        compare.mainLevelId || '',
        compare.levelTitle || '',
        compare.levelUuid || '',
        compare.pageUrl || '',
        compare.track || '',
        compare.lessonMaterials || '',
        compare.lessonVideoUrl  || '',
      ]);
    });
  });

  writeLmsPairsSheet_(outSheet, outRows);
}

function readLmsLevelsForPairs_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h || '').trim());
  const rows = values.slice(1);

  const allRows = rows.map((row, rowIndex) => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i] != null ? String(row[i]) : '';
    });

    const level = {
      _rowIndex: rowIndex + 2,
      timestamp: obj.timestamp || '',
      lessonId: obj.lessonId || '',
      taskId: obj.taskId || '',
      track: obj.track || '',
      taskTitle: obj.taskTitle || '',
      levelKind: obj.levelKind || '',
      orderInTask: obj.orderInTask || '',
      taskLevelId: obj.taskLevelId || '',
      mainLevelId: obj.mainLevelId || '',
      multiLevelId: obj.multiLevelId || '',
      parentTaskLevelId: obj.parentTaskLevelId || '',
      parentMainLevelId: obj.parentMainLevelId || '',
      parentLevelTitle: obj.parentLevelTitle || '',
      levelUuid: obj.levelUuid || '',
      levelTitle: obj.levelTitle || '',
      pageUrl: obj.pageUrl || '',
      lessonTitle: obj.lessonTitle || '',
      lessonGuid: obj.lessonGuid || '',
      lessonStatus: obj.lessonStatus || '',
      courseTitle: obj.courseTitle || '',
      courseUrl: obj.courseUrl || '',
      courseUuid: obj.courseUuid || '',
      courseLanguage: obj.courseLanguage || '',
      courseLocale: obj.courseLocale || '',
      msoStatus: obj.msoStatus || '',
      msoEnabled: obj.msoEnabled || '',
      publicName: obj.publicName || '',
      hasPublicName: obj.hasPublicName || '',
      pageTitle: obj.pageTitle || '',
      lessonMaterials: obj.lessonMaterials || '',
      lessonVideoUrl:  obj.lessonVideoUrl  || '',
      lessonPositionInCourse: obj.lessonPositionInCourse || '',
      lessonTotalInCourse: obj.lessonTotalInCourse || '',
      lessonNote: obj.lessonNote || '',
      isBonus: obj.isBonus || '',
      isTheory: obj.isTheory || '',
      isQuiz: obj.isQuiz || '',
    };

    level.locale = detectLocaleForLevelRow_(level);
    level.taskSeq = extractTaskSeqFromOrder_(level.orderInTask);

    return level;
  }).filter(r => {
    return clean_(r.mainLevelId) && clean_(r.locale);
  });

  const seenLevelIds = {};
  return allRows.filter(function(r) {
    if (seenLevelIds[r.mainLevelId]) return false;
    seenLevelIds[r.mainLevelId] = true;
    return true;
  });
}

function writeLmsPairsSheet_(sheet, rows) {
  const header = [
    'taskSeq',
    'orderInTask',
    'pairKey',

    'base_locale',
    'compare_locale',

    'pairStatus',
    'pairConfidence',
    'pairingScore',
    'pairingReason',
    'pairingNote',

    'ru_locale',
    'ru_taskId',
    'ru_taskTitle',
    'ru_lessonId',
    'ru_lessonTitle',
    'ru_levelKind',
    'ru_mainLevelId',
    'ru_levelTitle',
    'ru_levelUuid',
    'ru_pageUrl',
    'ru_track',
    'ru_lessonMaterials',
    'ru_lessonVideoUrl',

    'az_locale',
    'az_taskId',
    'az_taskTitle',
    'az_lessonId',
    'az_lessonTitle',
    'az_levelKind',
    'az_mainLevelId',
    'az_levelTitle',
    'az_levelUuid',
    'az_pageUrl',
    'az_track',
    'az_lessonMaterials',
    'az_lessonVideoUrl',
  ];

  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  const widths = [
    70, 90, 240,
    90, 90,
    140, 110, 90, 180, 240,
    80, 90, 220, 90, 220, 100, 100, 220, 180, 220, 80, 200, 200,
    80, 90, 220, 90, 220, 100, 100, 220, 180, 220, 80, 200, 200,
  ];

  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  sheet.setFrozenRows(1);
  if (sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, sheet.getLastRow(), header.length)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment('top');

    sheet.getRange(1, 1, 1, header.length)
      .setFontWeight('bold')
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff');
  }
}

function detectLocaleForLevelRow_(row) {
  const fromLocale = normalizeLocaleForPairs_(row.courseLocale);
  if (fromLocale) return fromLocale;

  const fromLanguage = normalizeLocaleForPairs_(row.courseLanguage);
  if (fromLanguage) return fromLanguage;

  const probe = [
    row.track,
    row.courseTitle,
    row.lessonTitle,
    row.taskTitle,
    row.levelTitle,
    row.pageTitle,
    row.courseUrl,
    row.pageUrl
  ].join(' | ').toLowerCase();

  if (/\bru\b|russian|рус|русский|ru track|track ru/.test(probe)) return 'RU';
  if (/\baz\b|azer|azərbay|azərbaycan|track az|az track/.test(probe)) return 'AZ';
  if (/\ben\b|english|англ|ingilis|track en|en track/.test(probe)) return 'EN';
  if (/\bid\b|indonesian|bahasa indonesia|индонез/i.test(probe)) return 'ID';
  if (/\bes\b|spanish|español|испан/i.test(probe)) return 'ES';

  const rawText = [
    row.lessonTitle,
    row.taskTitle,
    row.levelTitle,
    row.pageTitle
  ].join(' ');

  if (/[ƏəĞğİıÖöŞşÜüÇç]/.test(rawText)) return 'AZ';
  if (/[А-Яа-яЁё]/.test(rawText)) return 'RU';

  return '';
}

function normalizeLocaleForPairs_(value) {
  const s = clean_(value).toLowerCase();
  if (!s) return '';

  if (/^ru\b|^ru-/.test(s)) return 'RU';
  if (/^az\b|^az-/.test(s)) return 'AZ';
  if (/^en\b|^en-/.test(s)) return 'EN';
  if (/^id\b|^id-/.test(s)) return 'ID';
  if (/^es\b|^es-/.test(s)) return 'ES';

  if (s === 'русский' || /russian|рус/.test(s)) return 'RU';
  if (s === 'azərbaycan' || /azer|azərbay/.test(s)) return 'AZ';
  if (s === 'english' || /ingilis|англ/.test(s)) return 'EN';
  if (s === 'bahasa indonesia' || /indones/i.test(s)) return 'ID';
  if (s === 'español' || /spanish|испан/.test(s)) return 'ES';

  return '';
}

function extractTaskSeqFromOrder_(orderInTask) {
  const s = String(orderInTask || '').trim();
  if (!s) return '';
  const m = s.match(/^(\d+)/);
  return m ? m[1] : s;
}

function enrichHierarchyOrdinalsForPairs_(levels) {
  const rows = (Array.isArray(levels) ? levels.slice() : []).sort((a, b) => {
    return (a._rowIndex || 0) - (b._rowIndex || 0);
  });

  let taskOrdinal = 0;
  let prevTaskId = '';
  let prevTaskSeq = '';

  const taskSeenMains = {};
  const taskSeenSubs = {};

  rows.forEach(row => {
    const taskId = clean_(row.taskId);
    const taskSeq = clean_(row.taskSeq);
    const taskBoundaryChanged =
      (taskId && taskId !== prevTaskId) ||
      (!taskId && taskSeq !== prevTaskSeq && taskSeq !== '');

    if (taskBoundaryChanged || taskOrdinal === 0) {
      taskOrdinal++;
      prevTaskId = taskId;
      prevTaskSeq = taskSeq;

      if (!taskSeenMains[taskOrdinal]) taskSeenMains[taskOrdinal] = {};
      if (!taskSeenSubs[taskOrdinal]) taskSeenSubs[taskOrdinal] = {};
    }

    row.taskOrdinal = taskOrdinal;

    const taskMainMap = taskSeenMains[taskOrdinal];
    const taskSubMap = taskSeenSubs[taskOrdinal];

    const mainKey = buildMainKeyForPairs_(row);
    if (!taskMainMap[mainKey]) {
      taskMainMap[mainKey] = Object.keys(taskMainMap).length + 1;
    }
    row.mainOrdinal = taskMainMap[mainKey];

    const subKey = buildSubKeyForPairs_(row);
    if (!taskSubMap[mainKey]) taskSubMap[mainKey] = {};

    if (!taskSubMap[mainKey][subKey]) {
      taskSubMap[mainKey][subKey] = Object.keys(taskSubMap[mainKey]).length + 1;
    }
    row.subOrdinal = taskSubMap[mainKey][subKey];

    row.hierarchyKey =
      padNum__(row.taskOrdinal, 4) + '|' +
      padNum__(row.mainOrdinal, 4) + '|' +
      padNum__(row.subOrdinal, 4);
  });

  return rows;
}

function alignLevelPairsWithinTaskSeq_(ruGroup, azGroup, rawMap) {
  const ru = Array.isArray(ruGroup) ? ruGroup.slice() : [];
  const az = Array.isArray(azGroup) ? azGroup.slice() : [];

  if (!ru.length && !az.length) return [];

  const best = findBestLevelAssignmentWithinTaskSeq_(ru, az, rawMap);
  const pairs = [];
  const usedAz = {};

  for (let i = 0; i < ru.length; i++) {
    const ruRow = ru[i];
    const azIndex = best.assignment[i];

    if (azIndex === -1 || azIndex === undefined || !az[azIndex]) {
      pairs.push({
        ru: ruRow, az: {},
        pairStatus: 'COMPARE_LEVEL_MISSING',
        pairConfidence: 'LOW',
        pairingScore: '', pairingReason: '',
        pairingNote: 'Для base-уровня не найдена пара на compare-стороне'
      });
      continue;
    }

    usedAz[azIndex] = true;
    const azRow = az[azIndex];
    const scored = scoreLevelPairSimilarity_(ruRow, azRow, rawMap);
    const sameMain = String(ruRow.mainOrdinal || '') === String(azRow.mainOrdinal || '');
    const sameSub  = String(ruRow.subOrdinal  || '') === String(azRow.subOrdinal  || '');

    let pairStatus = 'MATCHED';
    let pairConfidence = 'HIGH';
    let pairingNote = '';

    if (!sameMain || !sameSub) {
      pairStatus = 'REORDERED_MATCH';
      pairConfidence = 'TENTATIVE';
      pairingNote = 'Уровень сопоставлен не по ожидаемой позиции в иерархии';
    } else if (scored.score < 12) {
      pairStatus = 'WEAK_MATCH';
      pairConfidence = 'TENTATIVE';
      pairingNote = 'Пара выбрана по похожести, но совпадение слабое';
    }

    pairs.push({
      ru: ruRow, az: azRow,
      pairStatus, pairConfidence,
      pairingScore: scored.score,
      pairingReason: scored.reason,
      pairingNote
    });
  }

  for (let j = 0; j < az.length; j++) {
    if (usedAz[j]) continue;
    pairs.push({
      ru: {}, az: az[j],
      pairStatus: 'BASE_LEVEL_MISSING',
      pairConfidence: 'LOW',
      pairingScore: '', pairingReason: '',
      pairingNote: 'Лишний compare-уровень без пары'
    });
  }

  return pairs;
}

function findBestLevelAssignmentWithinTaskSeq_(ruGroup, azGroup, rawMap) {
  const ru = Array.isArray(ruGroup) ? ruGroup : [];
  const az = Array.isArray(azGroup) ? azGroup : [];
  const assignment = [];
  const usedAz = {};

  for (let i = 0; i < ru.length; i++) {
    let bestScore = -999999;
    let bestIndex = -1;

    for (let j = 0; j < az.length; j++) {
      if (usedAz[j]) continue;
      const score = scoreLevelPairSimilarity_(ru[i], az[j], rawMap).score;
      const finalScore = score - Math.abs(i - j) * 2;
      if (finalScore > bestScore) { bestScore = finalScore; bestIndex = j; }
    }

    if (bestScore < -10) {
      assignment.push(-1);
    } else {
      assignment.push(bestIndex);
      if (bestIndex !== -1) usedAz[bestIndex] = true;
    }
  }

  return { score: 0, assignment };
}

function scoreLevelPairSimilarity_(ruRow, azRow, rawMap) {
  const reasons = [];
  let score = 0;

  const ruPayload = rawMap[String(ruRow.mainLevelId || '').trim()] || null;
  const azPayload = rawMap[String(azRow.mainLevelId || '').trim()] || null;

  if (clean_(ruRow.levelKind) && clean_(azRow.levelKind)) {
    if (clean_(ruRow.levelKind) === clean_(azRow.levelKind)) { score += 4; reasons.push('levelKind'); }
    else score -= 3;
  }

  const ruOrder = parseOrderForPairs_(ruRow.orderInTask);
  const azOrder = parseOrderForPairs_(azRow.orderInTask);
  const orderDelta = Math.abs(ruOrder - azOrder);

  if (ruOrder === azOrder) { score += 8; reasons.push('order'); }
  else if (orderDelta <= 0.01) { score += 6; reasons.push('order~'); }
  else if (orderDelta <= 0.11) { score += 3; reasons.push('order_close'); }
  else if (orderDelta <= 1) score -= 2;
  else score -= 8;

  if (String(ruRow.mainOrdinal || '') === String(azRow.mainOrdinal || '')) { score += 8; reasons.push('mainOrdinal'); }
  else score -= 6;

  if (String(ruRow.subOrdinal || '') === String(azRow.subOrdinal || '')) { score += 5; reasons.push('subOrdinal'); }
  else score -= 3;

  if (!ruPayload || !azPayload) return { score, reason: reasons.join(', ') };

  const ruType = clean_(ruPayload.type), azType = clean_(azPayload.type);
  if (ruType && azType) { if (ruType === azType) { score += 6; reasons.push('type'); } else score -= 5; }

  const ruCfgType = clean_(getPath_(ruPayload, 'config.type'));
  const azCfgType = clean_(getPath_(azPayload, 'config.type'));
  if (ruCfgType && azCfgType) { if (ruCfgType === azCfgType) { score += 8; reasons.push('config.type'); } else score -= 6; }

  const ruMech = clean_(firstNonEmpty_([
    getPath_(ruPayload, 'config.mechanic.problemType'),
    getPath_(ruPayload, 'children.0.config.mechanic.problemType')
  ]));
  const azMech = clean_(firstNonEmpty_([
    getPath_(azPayload, 'config.mechanic.problemType'),
    getPath_(azPayload, 'children.0.config.mechanic.problemType')
  ]));
  if (ruMech && azMech) { if (ruMech === azMech) { score += 10; reasons.push('mechanic'); } else score -= 8; }

  return { score, reason: reasons.join(', ') };
}

function buildMainKeyForPairs_(row) {
  return firstNonEmpty_([
    clean_(row.mainLevelId),
    clean_(row.parentMainLevelId),
    clean_(row.taskLevelId),
    clean_(row.orderInTask),
    'ROW_' + clean_(row._rowIndex)
  ]);
}

function buildSubKeyForPairs_(row) {
  return firstNonEmpty_([
    clean_(row.multiLevelId),
    clean_(row.taskLevelId),
    clean_(row.orderInTask),
    clean_(row.levelUuid),
    'ROW_' + clean_(row._rowIndex)
  ]);
}

function padNum__(n, len) {
  const s = String(parseInt(n, 10) || 0);
  return ('0000000000' + s).slice(-len);
}

function sortByHierarchyForPairs_(a, b) {
  const ta = parseInt(a.taskOrdinal || '0', 10) || 0;
  const tb = parseInt(b.taskOrdinal || '0', 10) || 0;
  if (ta !== tb) return ta - tb;

  const ma = parseInt(a.mainOrdinal || '0', 10) || 0;
  const mb = parseInt(b.mainOrdinal || '0', 10) || 0;
  if (ma !== mb) return ma - mb;

  const sa = parseInt(a.subOrdinal || '0', 10) || 0;
  const sb = parseInt(b.subOrdinal || '0', 10) || 0;
  if (sa !== sb) return sa - sb;

  return (a._rowIndex || 0) - (b._rowIndex || 0);
}

function collectTaskOrdinalsForPairs_(ruLevels, azLevels) {
  const seen = {};
  const out = [];

  ruLevels.concat(azLevels)
    .sort((a, b) => (a._rowIndex || 0) - (b._rowIndex || 0))
    .forEach(r => {
      const key = String(r.taskOrdinal || '').trim();
      if (!key || seen[key]) return;
      seen[key] = true;
      out.push(key);
    });

  return out.sort((a, b) => {
    return (parseInt(a, 10) || 0) - (parseInt(b, 10) || 0);
  });
}

function buildStablePairKeyForLmsPairs_(obj) {
  return [
    clean_(obj.taskSeq) || 'NO_TASKSEQ',
    clean_(obj.orderInTask) || 'NO_ORDER',
    clean_(obj.ru_mainLevelId) || 'NO_RU',
    clean_(obj.az_mainLevelId) || 'NO_AZ',
    clean_(obj.ru_taskId || obj.az_taskId) || 'NO_TASKID'
  ].join(' | ');
}

function logLocalesForPairs_(levels) {
  const stats = {};

  levels.forEach(r => {
    const key = [
      clean_(r.courseLanguage) || 'NO_LANG',
      clean_(r.courseLocale) || 'NO_LOCALE',
      clean_(r.locale) || 'NO_DETECTED'
    ].join(' | ');

    stats[key] = (stats[key] || 0) + 1;
  });

  Object.keys(stats).sort().forEach(key => {
    Logger.log('LMS_LEVELS locale stats: ' + key + ' -> ' + stats[key]);
  });
}

function parseOrderForPairs_(v) {
  const s = String(v || '').trim().replace(',', '.');
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function getPairLocalesForLmsPairs_(levels) {
  const stats = {};

  (levels || []).forEach(r => {
    const locale = clean_(r.locale);
    if (!locale) return;
    stats[locale] = (stats[locale] || 0) + 1;
  });

  const locales = Object.keys(stats);
  if (locales.length < 2) {
    throw new Error(
      'Для сравнения нужно минимум 2 локали. Найдено: ' +
      (locales.length ? locales.join(', ') : '0')
    );
  }

  if (locales.length > 2) {
    Logger.log(
      'Внимание: найдено больше 2 локалей: ' +
      JSON.stringify(stats) +
      '. Будут выбраны только 2.'
    );
  }

  const sorted = locales.sort((a, b) => {
    if (a === 'RU') return -1;
    if (b === 'RU') return 1;

    const byCount = (stats[b] || 0) - (stats[a] || 0);
    if (byCount !== 0) return byCount;

    return a.localeCompare(b);
  });

  const baseLocale = sorted[0];
  const compareLocale = sorted.find(x => x !== baseLocale);

  if (!baseLocale || !compareLocale) {
    throw new Error('Не удалось определить пару локалей для сравнения');
  }

  return {
    baseLocale: baseLocale,
    compareLocale: compareLocale
  };
}
