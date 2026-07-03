// ============================================================
// QA SINGLE
// QA-отчёт для одного урока без сравнения с локализацией.
// Логика аналогична buildLevelQAView, но без второй колонки.
// ============================================================

function buildLevelQASingleView(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  try {
    updateQaProgress_(ss, 'Старт сборки QA (single)', 1, 10);

    updateQaProgress_(ss, 'Разбираю теги уровней', 2, 10);
    buildRawParsedNodes(ss);
    SpreadsheetApp.flush();
    Utilities.sleep(1000);

    updateQaProgress_(ss, 'Строю пары', 3, 10);
    buildLmsPairsSingle_(ss);
    SpreadsheetApp.flush();
    Utilities.sleep(500);

    const pairsSheet = ss.getSheetByName('LMS_PAIRS');
    const nodesSheet = ss.getSheetByName('RAW_PARSED_NODES');
    const lmsLevelsSheet = ss.getSheetByName('LMS_LEVELS');
    const issuesSheet = ss.getSheetByName('LEVEL_ISSUES');

    if (!pairsSheet) throw new Error('Лист LMS_PAIRS не найден');
    if (!nodesSheet) throw new Error('Лист RAW_PARSED_NODES не найден');
    if (!lmsLevelsSheet) throw new Error('Лист LMS_LEVELS не найден');

    updateQaProgress_(ss, 'Читаю данные из листов', 4, 10);

    const pairsData = getSheetData_(pairsSheet);
    setQaLocaleLabelsFromPairs_(pairsData);
    const nodesSheet2 = SpreadsheetApp.openById(ss.getId()).getSheetByName('RAW_PARSED_NODES');
    const nodesData = getSheetData_(nodesSheet2);
    const lmsLevelsData = getSheetData_(lmsLevelsSheet);
    const issuesData = issuesSheet ? getSheetData_(issuesSheet) : [];

    const nodesByLevel = buildNodesByLevel_(nodesData);
    const levelMetaByMainLevelId = buildLevelMetaByMainLevelId_(lmsLevelsData);

    const issuesByPair = {};
    issuesData.forEach(iss => {
      issuesByPair[String(iss.pairKey || '')] = iss;
    });

    const rows = [];
    const rowGroups = [];

    const lessonMetas = pickLessonMetasFromPairs_(pairsData, levelMetaByMainLevelId);
    addLessonConfigBlockSingle_(rows, lessonMetas.ruMeta);
    const ruMainIdsForMso = {};
    pairsData.forEach(function(p){ const id=clean_(p.ru_mainLevelId); if(id) ruMainIdsForMso[id]=true; });
    addMsoScoringBlock_(rows, lessonMetas.ruMeta, lmsLevelsData, nodesByLevel, ruMainIdsForMso);
    updateQaProgress_(ss, 'Собираю строки QA-листа', 5, 10);

    let lastTaskKey = null;
    pairsData.forEach((pair, pairIndex) => {
      if (pairIndex > 0 && pairIndex % 20 === 0) {
        updateQaProgress_(ss, 'Обработано уровней: ' + pairIndex, 5, 10);
      }

      const ruId = String(pair.ru_mainLevelId || '').trim();
      if (!ruId) return;

      const ruLevelMeta = levelMetaByMainLevelId[ruId] || null;
      const ruTaskId = clean_(ruLevelMeta && ruLevelMeta.taskId);
      const currentTaskKey = ruTaskId;

      if (currentTaskKey !== lastTaskKey) {
        lastTaskKey = currentTaskKey;

        const ruTaskTitle = clean_(ruLevelMeta && ruLevelMeta.taskTitle);
        rows.push({
          cells: ['📋 Задание', ruTaskTitle || '—'],
          fmt: 'taskheader'
        });

        const ruTaskType = detectTaskTypeForQa_(ruLevelMeta);
        rows.push({ cells: ['Тип задачи', ruTaskType], fmt: 'field' });
      }

      const ruNodes = dedupeNodeObjects_((nodesByLevel[ruId] || []).slice()).sort(sortNodesStable_);
      attachChildrenForQa_(ruNodes);

      const pairIssue = issuesByPair[String(pair.pairKey || '')] || {};
      const issueColor = String(pairIssue.status_color || '').toUpperCase();

      let statusIcon = '🟢';
      if (issueColor === 'RED') statusIcon = '🔴';
      else if (issueColor === 'YELLOW') statusIcon = '🟡';

      const levelHeaderRowIndex = rows.length;

      rows.push({
        cells: [
          'Уровень ' + clean_(pair.orderInTask) + '  ' + statusIcon,
          pair.ru_levelTitle || (ruLevelMeta && ruLevelMeta.levelTitle) || ''
        ],
        fmt: 'header'
      });

      const detailStart = rows.length;

      rows.push({ cells: ['ID', ruId], fmt: 'field' });
      rows.push({
        cells: [
          'Ссылка на уровень',
          formatDisplayValueForQa_(buildLevelUrlFromMeta_(ruLevelMeta) || clean_(pair.ru_pageUrl))
        ],
        fmt: 'field'
      });

      const ruRoot = findMainLevelNode_(ruNodes);
      if (ruRoot && ruRoot.type === '__TOO_BIG__') {
        rows.push({
          cells: ['⚠ Уровень не загружен', clean_(ruRoot.description)],
          fmt: 'critical'
        });
      }
      const ruCasingIssues = checkConfigKeyCasing_(ruRoot && ruRoot.config_json);
      if (ruCasingIssues.length) {
        rows.push({
          cells: [
            '⚠ Теги конфига',
            ruCasingIssues.map(function(i) {
              return i.kind === 'casing'
                ? '⚠ регистр: ' + i.found + ' → ' + i.expected
                : '? похоже на: ' + i.found + ' → ' + i.expected;
            }).join('\n')
          ],
          fmt: 'critical'
        });
      }

      if (ruRoot) {
        const rr = ruRoot;

        addSingleRow_(rows, 'Механика', rr.mechanic_type || rr.type);
        addSingleRow_(rows, 'Баллы', rr.levelScore);
        addSingleRow_(rows, 'Автопроверка', rr.isAutocheck);
        addSingleRow_(rows, 'isAllowFail', rr.isAllowFail);
        addSingleRow_(rows, 'force_enabled', rr.force_enabled);
        addSingleRow_(rows, 'Защита от копирования', detectCopyProtectionForQa_(ruNodes));

        const ruRootTests = extractPythonCodeTestsForQa_(rr.config_json);
        if (ruRootTests.testInput) addSingleRow_(rows, 'testInput', ruRootTests.testInput);
        if (ruRootTests.testOutput) addSingleRow_(rows, 'testOutput', ruRootTests.testOutput);

        if (rr.description) addSingleTextRow_(rows, 'Описание (root)', rr.description);
        if (rr.inputs_json) addSingleRow_(rows, 'Inputs (root)', formatJson_(rr.inputs_json));

        const ruRootImages = collectNodeImagesForQa_(rr);
        if (ruRootImages.length) {
          rows.push({
            cells: ['Изображения (root)', ruRootImages.map(formatImageLabelForQa_).join('\n')],
            fmt: 'img',
            ruImgs: ruRootImages
          });
        }

        if (hasEmbeddedLinks_(rr.embedded_links_json || '[]')) {
          addSingleEmbedRow_(rows, 'Встроенные ссылки (root)', rr.embedded_links_json);
        }

        const isUploaderPair =
          clean_(rr.type) === 'uploader' || clean_(rr.config_type) === 'uploader';

        if (isUploaderPair) {
          addSingleRow_(rows, 'Задачник есть', normalizeBoolText_(rr.tasklist_exists));
          addSingleRow_(rows, 'Задачник виден студенту', normalizeBoolText_(rr.tasklist_visible));
          addSingleRow_(rows, 'Проверка задачника', rr.tasklist_check_type);
          if (rr.uploader_description) addSingleTextRow_(rows, 'Описание uploader', rr.uploader_description);
          if (rr.tasklist_text) addSingleTextRow_(rows, 'Текст задачника', rr.tasklist_text);
          if (hasNamedLinks_(rr.tasklist_links_json || '[]')) {
            addSingleEmbedRow_(rows, 'Ссылки в задачнике', rr.tasklist_links_json);
          }
          if (hasNamedLinks_(rr.level_links_json || '[]')) {
            addSingleEmbedRow_(rows, 'Ссылки уровня', rr.level_links_json);
          }
        }

        addConfigFlagsSingle_(rows, rr.config_json, rr);
      }

      const ruTasks = extractComparableTaskNodes_(ruNodes);

      ruTasks.forEach((rn, i) => {
        rows.push({ cells: ['Задание ' + (i + 1), rn.type || '—'], fmt: 'subheader' });

        if (rn.mechanic_type) addSingleRow_(rows, '  Механика', rn.mechanic_type);
        addSingleRow_(rows, '  Автопроверка', rn.isAutocheck);
        if (rn.description) addSingleTextRow_(rows, '  Описание', rn.description);
        if (rn.note) addSingleTextRow_(rows, '  Note / условие', rn.note);
        if (rn.text) addSingleTextRow_(rows, '  Шаблон (text)', rn.text);
        if (rn.template) addSingleTextRow_(rows, '  Template', rn.template);
        const ruGridCheck = extractCoordinateGridForQa_(rn);
        if (rn.checkExpression && !ruGridCheck) {
          addSingleRow_(rows, '  checkExpression', rn.checkExpression);
        }
        if (rn.verificationType) addSingleRow_(rows, '  Верификация', rn.verificationType);
        if (rn.variables_json) addSingleRow_(rows, '  Переменные', formatJson_(rn.variables_json));

        const ruAnswerStruct = extractAnswersStructure_(rn.config_json);
        if (ruAnswerStruct) addSingleRow_(rows, '  Ответы (структура)', ruAnswerStruct);

        const ruAnswers = extractAnswersWithLinksForQa_(rn.config_json);
        if (ruAnswers.length) {
          rows.push({
            cells: ['  Ответы (текст)', ruAnswers.map(x => x.label).join('\n')],
            fmt: 'field',
            ruAnswerLinks: ruAnswers
          });
        }

        const ruTaskImages = dedupeImagesForQa_(
          collectNodeImagesForQa_(rn).concat(extractMatchingColumnImagesForQa_(rn.config_json))
        );
        ruTaskImages.forEach(function(img, idx) {
          rows.push({
            cells: [idx === 0 ? '  Изображения' : '  ', formatImageLabelForQa_(img)],
            fmt: 'img',
            ruImgs: [img]
          });
        });

        if (hasEmbeddedLinks_(rn.embedded_links_json || '[]')) {
          addSingleEmbedRow_(rows, '  Встроенные ссылки', rn.embedded_links_json);
        }

        if (rn.hints_count && rn.hints_count !== '0') {
          addSingleRow_(rows, '  Подсказки', rn.hints_count);
        }

        const ruColumnsText = formatMatchingColumnsForQa_(rn.config_json);
        if (ruColumnsText) {
          rows.push({ cells: ['  Колонки', ruColumnsText], fmt: 'tl_ok' });
        }

        const ruCategorization = extractMathCategorizationMappingForQa_(rn);
        if (ruCategorization) addSingleTextRow_(rows, '  Категории и ответы', ruCategorization);

        const ruGrid = extractCoordinateGridForQa_(rn);
        if (ruGrid) addSingleTextRow_(rows, '  Сетка (предметы и ответы)', ruGrid);

        const ruTypingPattern = extractTypingPatternForQa_(rn.config_json);
        if (ruTypingPattern) addSingleRow_(rows, '  Текст тренажёра', ruTypingPattern);

        const ruTaskTests = extractPythonCodeTestsForQa_(rn.config_json);
        if (ruTaskTests.testInput) addSingleRow_(rows, '  testInput', ruTaskTests.testInput);
        if (ruTaskTests.testOutput) addSingleRow_(rows, '  testOutput', ruTaskTests.testOutput);

        addConfigFlagsSingle_(rows, rn.config_json, rn);
      });

      const detailEndBeforeSep = rows.length - 1;

      const finalStatusIcon = getLevelHeaderStatusIconSingle_(
        rows, detailStart, detailEndBeforeSep, statusIcon
      );
      rows[levelHeaderRowIndex].cells[0] =
        'Уровень ' + clean_(pair.orderInTask) + '  ' + finalStatusIcon;

      rows.push({ cells: ['', ''], fmt: 'sep' });

      const detailEnd = rows.length - 1;
      rowGroups.push({
        start: detailStart,
        end: detailEnd,
        collapsed: finalStatusIcon === '🟢'
      });
    });

    updateQaProgress_(ss, 'Создаю файл', 6, 10);

    const lessonName = getLessonName_(pairsData, nodesData, issuesData, lmsLevelsData);
    const fileBaseName = buildQaFileBaseName_(lessonName, lmsLevelsData);
    const targetInfo = createQaSpreadsheetInPersonalFolder_(ss, fileBaseName + '_single');
    const targetSs = targetInfo.spreadsheet;

    updateQaProgress_(ss, 'Записываю лист', 7, 10);

    const sheetName = makeUniqueSheetName_(targetSs, 'LEVEL_QA_' + lessonName);
    const sheet = targetSs.insertSheet(sheetName);

    const allData = [['Поле', 'Значение']].concat(rows.map(r => r.cells));
    const formats = ['colheader'].concat(rows.map(r => r.fmt));

    if (allData.length > 0) {
      sheet.getRange(1, 1, allData.length, 2).setValues(allData);
    }

    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 700);

    const S = getQaStyles_();

    for (let i = 0; i < formats.length; i++) {
      const style = S[formats[i]] || S.field;
      sheet.getRange(i + 1, 1, 1, 2)
        .setBackground(style.bg).setFontColor(style.fg)
        .setFontWeight(style.bold ? 'bold' : 'normal')
        .setFontSize(style.size)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setVerticalAlignment('top');
    }

    updateQaProgress_(ss, 'Проставляю ссылки и группировки', 8, 10);

    rows.forEach((r, i) => {
      if ((r.fmt === 'img' || r.fmt === 'img_diff') && r.ruImgs) {
        sheet.setRowHeight(i + 2, 150);
        setImageLinks_(sheet, i + 2, 2, r.ruImgs);
      }
      if (r.ruEmbeds) setEmbeddedLinks_(sheet, i + 2, 2, r.ruEmbeds);
      if (r.ruAnswerLinks) setAnswerLinks_(sheet, i + 2, 2, r.ruAnswerLinks);
      if (r._ruMaterials) setMaterialLinks_(sheet, i + 2, 2, r._ruMaterials);
      if (r._ruVideo) {
        sheet.getRange(i + 2, 2)
          .setFormula('=HYPERLINK("' + escapeForFormula_(r._ruVideo) + '","▶ Видео")');
      }
    });

    rowGroups.forEach(g => {
      const s = g.start + 2, e = g.end + 2;
      if (e >= s) sheet.getRange(s, 1, e - s + 1, 1).shiftRowGroupDepth(1);
    });
    rowGroups.forEach(g => {
      try {
        const gr = sheet.getRowGroupAt(g.start + 2, 1);
        if (gr && g.collapsed) gr.collapse();
      } catch(e) {}
    });

    sheet.setFrozenRows(1);

    updateQaProgress_(ss, 'Создаю BRIEF и текстовый листы', 9, 10);
    buildLevelQaBriefSheetSingle_(targetSs, lessonName, rows, S);
    buildLevelTextsSheetSingle_(targetSs, lessonName, pairsData, nodesByLevel);

    removeDefaultSheetIfNeeded_(targetSs);
    SpreadsheetApp.flush();

    updateQaProgress_(ss, 'Готово', 10, 10);
    showQaResultDialog_(targetInfo.url, targetInfo.fileName);

    return targetInfo.url;

  } catch (err) {
    Logger.log('Ошибка buildLevelQASingleView: ' + err.message + '\n' + (err.stack || ''));
    ss.toast('Ошибка: ' + err.message, 'QA', 10);
    throw err;
  }
}

// ─── PAIRS SINGLE ────────────────────────────────────────────

function buildLmsPairsSingle_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();

  const levelsSheet = ss.getSheetByName('LMS_LEVELS');
  const rawSheet = ss.getSheetByName('RAW_API');
  if (!levelsSheet) throw new Error('Лист LMS_LEVELS не найден');
  if (!rawSheet) throw new Error('Лист RAW_API не найден');

  const outSheet = ss.getSheetByName('LMS_PAIRS') || ss.insertSheet('LMS_PAIRS');

  const levels = readLmsLevelsForPairs_(levelsSheet);
  if (!levels.length) throw new Error('LMS_LEVELS пуст');

  enrichHierarchyOrdinalsForPairs_(levels);
  SpreadsheetApp.flush();

  const taskOrdinals = collectTaskOrdinalsForPairs_(levels, []);
  const outRows = [];

  taskOrdinals.forEach(taskOrdinal => {
    const group = levels
      .filter(r => String(r.taskOrdinal) === String(taskOrdinal))
      .sort(sortByHierarchyForPairs_);

    group.forEach(row => {
      const pairKey = buildStablePairKeyForLmsPairs_({
        taskSeq: row.taskSeq || String(taskOrdinal),
        orderInTask: row.orderInTask,
        ru_mainLevelId: row.mainLevelId,
        az_mainLevelId: '',
        ru_taskId: row.taskId,
        az_taskId: ''
      });

      outRows.push([
        row.taskSeq || String(taskOrdinal),
        row.orderInTask || '',
        pairKey,

        row.locale || '',
        '',

        'SINGLE', 'HIGH', '', '', '',

        row.locale || '',
        row.taskId || '',
        row.taskTitle || '',
        row.lessonId || '',
        row.lessonTitle || '',
        row.levelKind || '',
        row.mainLevelId || '',
        row.levelTitle || '',
        row.levelUuid || '',
        row.pageUrl || '',
        row.track || '',
        row.lessonMaterials || '',
        row.lessonVideoUrl || '',

        '', '', '', '', '', '', '', '', '', '', '', '', ''
      ]);
    });
  });

  writeLmsPairsSheet_(outSheet, outRows);
  SpreadsheetApp.flush();
}

// ─── BRIEF SINGLE ────────────────────────────────────────────

function buildLevelQaBriefSheetSingle_(ss, lessonName, rows, S) {
  const briefSheetName = makeUniqueSheetName_(ss, 'LEVEL_QA_BRIEF_' + lessonName);
  const sheet = ss.insertSheet(briefSheetName);

  const briefRows = [];
  const formats = [];

  briefRows.push(['Пункт', 'Статус', 'Механика', 'Комментарий']);
  formats.push('colheader');

  rows.forEach(function(r) {
    if (!r || !r.cells) return;

    const label = clean_(r.cells[0]);
    const fmt = clean_(r.fmt);

    if (fmt === 'taskheader') {
      briefRows.push([label, '📋', '—', clean_(r.cells[1])]);
      formats.push('taskheader');
      return;
    }

    if (fmt === 'header') {
      const iconMatch = label.match(/[🔴🟡🟢🟠]/);
      const statusIcon = iconMatch ? iconMatch[0] : '🟢';
      briefRows.push([label, statusIcon, '—', clean_(r.cells[1])]);
      formats.push(
        statusIcon === '🔴' ? 'critical' :
        statusIcon === '🟡' ? 'check' :
        'field'
      );
      return;
    }

    if (fmt === 'sep') {
      briefRows.push(['', '', '', '']);
      formats.push('sep');
      return;
    }

    const isProblem =
      fmt === 'critical' || fmt === 'check' || fmt === 'info' ||
      fmt === 'tl_bad' || fmt === 'tl_warn' || fmt === 'tl_embed' ||
      fmt === 'img_diff' || fmt === 'embed' || fmt === 'embed_diff';

    if (!isProblem || !label) return;

    const value = clean_(r.cells[1]);
    const short = shortenForBrief_(value, 140);
    const icon =
      fmt === 'critical' || fmt === 'tl_bad' ? '🔴' :
      fmt === 'embed' || fmt === 'tl_embed' ? '🔎' : '🟠';
    const briefFmt =
      fmt === 'critical' || fmt === 'tl_bad' ? 'critical' :
      fmt === 'embed' || fmt === 'tl_embed' ? 'info' : 'check';

    briefRows.push(['   ↳ ' + label, icon, '', short]);
    formats.push(briefFmt);
  });

  sheet.getRange(1, 1, briefRows.length, 4).setValues(briefRows);

  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 70);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 760);

  for (let i = 0; i < formats.length; i++) {
    const style = S[formats[i]] || S.field;
    sheet.getRange(i + 1, 1, 1, 4)
      .setBackground(style.bg).setFontColor(style.fg)
      .setFontWeight(style.bold ? 'bold' : 'normal')
      .setFontSize(style.size)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment('top');
  }

  sheet.getRange(1, 2, briefRows.length, 1).setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  return sheet.getName();
}

// ─── TEXTS SINGLE ────────────────────────────────────────────

function buildLevelTextsSheetSingle_(ss, lessonName, pairsData, nodesByLevel) {
  const sheetName = makeUniqueSheetName_(ss, 'LEVEL_TEXTS_' + lessonName);
  const sheet = ss.insertSheet(sheetName);

  const header = ['Задание', 'Уровень', 'Поле', 'Текст'];
  const rows = [header];

  let taskCounter = 0;
  let lastTaskKey = null;

  pairsData.forEach(function(pair) {
    const ruId = String(pair.ru_mainLevelId || '').trim();
    if (!ruId) return;

    const ruTaskId = clean_(pair.ru_taskId);
    if (ruTaskId !== lastTaskKey) {
      lastTaskKey = ruTaskId;
      taskCounter++;
    }

    const ruNodes = (nodesByLevel[ruId] || []).slice().sort(sortNodesStable_);

    const taskLabel = clean_(pair.ru_taskTitle)
      ? 'Задание ' + taskCounter + '\n' + clean_(pair.ru_taskTitle)
      : 'Задание ' + taskCounter;

    const order = clean_(pair.orderInTask);
    const levelLabel = order
      ? order + '. ' + clean_(pair.ru_levelTitle)
      : clean_(pair.ru_levelTitle);

    const fields = collectTextFieldsSingle_(ruNodes);

    fields.forEach(function(field) {
      if (!field.text) return;
      rows.push([taskLabel, levelLabel, field.label, field.text]);
    });
  });

  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, 4).setValues(rows);
  }

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 560);

  sheet.getRange(1, 1, 1, 4)
    .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');

  sheet.getRange(1, 1, rows.length, 4)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setVerticalAlignment('top');

  sheet.setFrozenRows(1);

  return sheet.getName();
}

function collectTextFieldsSingle_(ruNodes) {
  attachChildrenForQa_(ruNodes);

  const ruRoot = findMainLevelNode_(ruNodes);
  const ruTasks = extractComparableTaskNodes_(ruNodes);
  const fields = [];

  if (ruRoot) {
    const desc = clean_(ruRoot.description);
    if (desc) fields.push({ label: 'Описание', text: stripHtml_(desc) });

    if (clean_(ruRoot.uploader_description)) {
      fields.push({ label: 'Описание uploader', text: stripHtml_(ruRoot.uploader_description) });
    }

    if (clean_(ruRoot.tasklist_text)) {
      fields.push({ label: 'Текст задачника', text: clean_(ruRoot.tasklist_text) });
    }

    const question = extractQuestionTextForTextsSheet_(ruRoot.config_json);
    if (question) fields.push({ label: 'Вопрос', text: stripHtml_(question) });

    const answers = extractAnswersTextForTextsSheet_(ruRoot.config_json);
    if (answers) fields.push({ label: 'Ответы', text: answers });

    const rootColumns = formatMatchingColumnsForQa_(ruRoot.config_json);
    if (rootColumns) fields.push({ label: 'Колонки', text: rootColumns });

    const rootCategorization = extractMathCategorizationMappingForQa_(ruRoot);
    if (rootCategorization) fields.push({ label: 'Категории и ответы', text: rootCategorization });

    const rootPattern = extractTypingPatternForQa_(ruRoot.config_json);
    if (rootPattern) fields.push({ label: 'Условие / код (pattern)', text: rootPattern });

    const rootTests = extractPythonCodeTestsForQa_(ruRoot.config_json);
    if (rootTests.testInput) fields.push({ label: 'testInput', text: rootTests.testInput });
    if (rootTests.testOutput) fields.push({ label: 'testOutput', text: rootTests.testOutput });
  }

  ruTasks.forEach(function(rn, idx) {
    const num = idx + 1;

    if (rn.description) fields.push({ label: 'Задание ' + num + ' — описание',  text: stripHtml_(rn.description) });
    if (rn.note)        fields.push({ label: 'Задание ' + num + ' — условие',    text: stripHtml_(rn.note) });
    if (rn.text)        fields.push({ label: 'Задание ' + num + ' — текст',      text: stripHtml_(rn.text) });
    if (rn.template)    fields.push({ label: 'Задание ' + num + ' — шаблон',     text: stripHtml_(rn.template) });

    const question = extractQuestionTextForTextsSheet_(rn.config_json);
    if (question) fields.push({ label: 'Задание ' + num + ' — вопрос', text: stripHtml_(question) });

    const answers = extractAnswersTextForTextsSheet_(rn.config_json);
    if (answers) fields.push({ label: 'Задание ' + num + ' — ответы', text: answers });

    const columns = formatMatchingColumnsForQa_(rn.config_json);
    if (columns) fields.push({ label: 'Задание ' + num + ' — колонки', text: columns });

    const categorization = extractMathCategorizationMappingForQa_(rn);
    if (categorization) fields.push({ label: 'Задание ' + num + ' — категории', text: categorization });

    const grid = extractCoordinateGridForQa_(rn);
    if (grid) fields.push({ label: 'Задание ' + num + ' — сетка (предметы и ответы)', text: grid });

    const typing = extractTypingPatternForQa_(rn.config_json);
    if (typing) fields.push({ label: 'Задание ' + num + ' — тренажёр', text: typing });
  });

  return fields;
}

// ─── SEVERITY ROWS SINGLE ────────────────────────────────────

function addLessonConfigBlockSingle_(rows, ruMeta) {
  rows.push({ cells: ['Название урока', clean_(ruMeta.lessonTitle) || '—'], fmt: 'field' });
  rows.push({
    cells: ['Ссылка на урок', formatDisplayValueForQa_(buildLessonUrlFromMeta_(ruMeta))],
    fmt: 'field'
  });
  rows.push({ cells: ['Курс', clean_(ruMeta.courseTitle) || '—'], fmt: 'field' });
  rows.push({
    cells: ['Ссылка на курс', formatDisplayValueForQa_(clean_(ruMeta.courseUrl))],
    fmt: 'field'
  });
  addSingleRow_(rows, 'МСО', ruMeta.msoStatus);
  addSingleRow_(rows, 'Статус урока', ruMeta.lessonStatus);
  addSingleRow_(rows, 'Публичное имя', ruMeta.publicName);
  addSingleRow_(rows, 'Есть публичное имя', normalizeBoolText_(ruMeta.hasPublicName));

  if (ruMeta.lessonNote) {
    rows.push({ cells: ['Заметка урока', clean_(ruMeta.lessonNote)], fmt: 'field' });
  }

  const materials = clean_(ruMeta.lessonMaterials);
  const video = clean_(ruMeta.lessonVideoUrl);
  if (materials) {
    rows.push({ cells: ['📎 Материалы урока', materials], fmt: 'embed', _ruMaterials: materials });
  }
  if (video) {
    rows.push({ cells: ['🎬 Видео урока', video], fmt: 'embed', _ruVideo: video });
  }

  rows.push({ cells: ['', ''], fmt: 'sep' });
}

function addSingleRow_(rows, label, value) {
  const v = formatDisplayValueForQa_(clean_(value));
  rows.push({ cells: [label, v || '—'], fmt: 'field' });
}

function addSingleTextRow_(rows, label, html) {
  const text = stripHtml_(html);
  if (!text) return;
  rows.push({ cells: [label, text], fmt: 'tl_ok' });
}

function addSingleEmbedRow_(rows, label, json) {
  const embeds = parseEmbeddedLinks_(json);
  if (!embeds.length) return;
  rows.push({
    cells: [label, embeds.map(i => i.name).join('\n')],
    fmt: 'embed',
    ruEmbeds: embeds
  });
}

function addConfigFlagsSingle_(rows, cfgJson, node) {
  const cfg = safeParse_(cfgJson);
  if (!cfg && !node) return;

  CONFIG_FLAGS.forEach(flag => {
    if (flag.path === 'mechanic.testInput' || flag.path === 'mechanic.testOutput') return;
    const rv = getConfigFlagValueForQa_(cfg, node, flag.path);
    if (rv === undefined || rv === null || String(rv).trim() === '') return;
    const rvs = safeJsonStringify_(rv);
    rows.push({ cells: ['  ⚙ ' + flag.label, rvs], fmt: 'field' });
  });
}

function getLevelHeaderStatusIconSingle_(rows, startIndex, endIndex, fallbackIcon) {
  let hasRed = false;
  let hasYellow = false;

  for (let i = startIndex; i <= endIndex; i++) {
    const row = rows[i];
    if (!row) continue;
    const fmt = String(row.fmt || '').trim();
    const icon = String((row.cells && row.cells[1]) || '').trim();

    if (fmt === 'critical' || fmt === 'tl_bad' || icon === '🔴') {
      hasRed = true;
      break;
    }
    if (
      fmt === 'check' || fmt === 'tl_warn' || fmt === 'embed_diff' ||
      fmt === 'img_diff' || fmt === 'info' ||
      icon === '🟠' || icon === '🟡' || icon === '🔎'
    ) {
      hasYellow = true;
    }
  }

  if (hasRed) return '🔴';
  if (hasYellow) return '🟡';
  return fallbackIcon || '🟢';
}
