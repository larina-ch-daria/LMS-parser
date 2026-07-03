// ============================================================
// QA TEXTS
// Лист с текстами уровней для проверки перевода: каждое текстовое
// поле уровня — отдельная строка (описание, текст задания,
// шаблон, ответы, колонки, категории, сетка, тренажёр).
// ============================================================

function buildLevelTextsSheet_(ss, lessonName, pairsData, nodesByLevel) {
  const sheetName = makeUniqueSheetName_(ss, 'LEVEL_TEXTS_' + lessonName);
  const sheet = ss.insertSheet(sheetName);

  const header = ['Задание', 'Уровень', 'Поле', QA_BASE_LABEL_, QA_CMP_LABEL_];
  const rows = [header];

  let taskCounter = 0;
  let lastTaskKey = null;

  pairsData.forEach(function(pair) {
    const ruId = String(pair.ru_mainLevelId || '').trim();
    const azId = String(pair.az_mainLevelId || '').trim();
    if (!ruId && !azId) return;

    const ruTaskId = clean_(pair.ru_taskId);
    const azTaskId = clean_(pair.az_taskId);
    const currentTaskKey = ruTaskId + '|' + azTaskId;

    if (currentTaskKey !== lastTaskKey) {
      lastTaskKey = currentTaskKey;
      taskCounter++;
    }

    const ruNodes = (nodesByLevel[ruId] || []).slice().sort(sortNodesStable_);
    const azNodes = (nodesByLevel[azId] || []).slice().sort(sortNodesStable_);

    const taskLabel = buildTextsTaskLabel_(pair, taskCounter);
    const levelLabel = buildTextsLevelLabel_(pair);

    const fields = collectTextFields_(ruNodes, azNodes);

    fields.forEach(function(field) {
      if (!field.ru && !field.az) return;
      rows.push([taskLabel, levelLabel, field.label, field.ru || '', field.az || '']);
    });
  });

  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, 5).setValues(rows);
  }

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 420);
  sheet.setColumnWidth(5, 420);

  sheet.getRange(1, 1, 1, 5)
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sheet.getRange(1, 1, rows.length, 5)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setVerticalAlignment('top');

  sheet.setFrozenRows(1);

  return sheet.getName();
}

function buildTextsTaskLabel_(pair, taskCounter) {
  const ruTitle = clean_(pair.ru_taskTitle);
  const azTitle = clean_(pair.az_taskTitle);

  const titlePart = (ruTitle && azTitle && ruTitle !== azTitle)
    ? QA_BASE_LABEL_ + ': ' + ruTitle + '\n' + QA_CMP_LABEL_ + ': ' + azTitle
    : ruTitle || azTitle || '';

  const prefix = 'Задание ' + taskCounter;
  return titlePart ? prefix + '\n' + titlePart : prefix;
}

function buildTextsLevelLabel_(pair) {
  const order = clean_(pair.orderInTask);
  const ruTitle = clean_(pair.ru_levelTitle);
  const azTitle = clean_(pair.az_levelTitle);

  const titlePart = (ruTitle && azTitle && ruTitle !== azTitle)
    ? QA_BASE_LABEL_ + ': ' + ruTitle + '\n' + QA_CMP_LABEL_ + ': ' + azTitle
    : ruTitle || azTitle || '';

  return order ? order + '. ' + titlePart : titlePart;
}

function collectTextFields_(ruNodes, azNodes) {
  attachChildrenForQa_(ruNodes);
  attachChildrenForQa_(azNodes);

  const ruRoot = findMainLevelNode_(ruNodes);
  const azRoot = findMainLevelNode_(azNodes);

  const ruTasks = extractComparableTaskNodes_(ruNodes);
  const azTasks = extractComparableTaskNodes_(azNodes);
  const aligned = alignComparableTaskPairs_(ruTasks, azTasks);

  const fields = [];

  if (ruRoot || azRoot) {
    const rr = ruRoot || {};
    const ar = azRoot || {};

    const ruRootDesc = clean_(rr.description);
    const azRootDesc = clean_(ar.description);
    if (ruRootDesc || azRootDesc) {
      fields.push({ label: 'Описание', ru: stripHtml_(ruRootDesc), az: stripHtml_(azRootDesc) });
    }

    const ruUploaderDesc = clean_(rr.uploader_description);
    const azUploaderDesc = clean_(ar.uploader_description);
    if (ruUploaderDesc || azUploaderDesc) {
      fields.push({ label: 'Описание uploader', ru: stripHtml_(ruUploaderDesc), az: stripHtml_(azUploaderDesc) });
    }

    const ruTasklistText = clean_(rr.tasklist_text);
    const azTasklistText = clean_(ar.tasklist_text);
    if (ruTasklistText || azTasklistText) {
      fields.push({ label: 'Текст задачника', ru: ruTasklistText, az: azTasklistText });
    }

    const ruRootQuestion = extractQuestionTextForTextsSheet_(rr.config_json);
    const azRootQuestion = extractQuestionTextForTextsSheet_(ar.config_json);
    if (ruRootQuestion || azRootQuestion) {
      fields.push({ label: 'Вопрос', ru: stripHtml_(ruRootQuestion), az: stripHtml_(azRootQuestion) });
    }

    const ruRootAnswers = extractAnswersTextForTextsSheet_(rr.config_json);
    const azRootAnswers = extractAnswersTextForTextsSheet_(ar.config_json);
    if (ruRootAnswers || azRootAnswers) {
      fields.push({ label: 'Ответы', ru: ruRootAnswers, az: azRootAnswers });
    }

    const ruRootColumns = formatMatchingColumnsForQa_(rr.config_json);
    const azRootColumns = formatMatchingColumnsForQa_(ar.config_json);
    if (ruRootColumns || azRootColumns) {
      fields.push({ label: 'Колонки', ru: ruRootColumns, az: azRootColumns });
    }

    const ruRootCategorization = extractMathCategorizationMappingForQa_(rr);
    const azRootCategorization = extractMathCategorizationMappingForQa_(ar);
    if (ruRootCategorization || azRootCategorization) {
      fields.push({ label: 'Категории и ответы', ru: ruRootCategorization, az: azRootCategorization });
    }

    // python_code: условие в cfg.pattern, тесты — в cfg.testInput/testOutput
    const ruPattern = extractTypingPatternForQa_(rr.config_json);
    const azPattern = extractTypingPatternForQa_(ar.config_json);
    if (ruPattern || azPattern) {
      fields.push({ label: 'Условие / код (pattern)', ru: ruPattern, az: azPattern });
    }

    const ruTests = extractPythonCodeTestsForQa_(rr.config_json);
    const azTests = extractPythonCodeTestsForQa_(ar.config_json);
    if (ruTests.testInput || azTests.testInput) {
      fields.push({ label: 'testInput', ru: ruTests.testInput, az: azTests.testInput });
    }
    if (ruTests.testOutput || azTests.testOutput) {
      fields.push({ label: 'testOutput', ru: ruTests.testOutput, az: azTests.testOutput });
    }
  }

  aligned.forEach(function(pairing, idx) {
    const rn = pairing.ru || {};
    const an = pairing.az || {};
    const taskNum = idx + 1;

    if (rn.description || an.description) {
      fields.push({
        label: 'Задание ' + taskNum + ' — описание',
        ru: stripHtml_(rn.description || ''),
        az: stripHtml_(an.description || '')
      });
    }

    if (rn.note || an.note) {
      fields.push({
        label: 'Задание ' + taskNum + ' — условие',
        ru: stripHtml_(rn.note || ''),
        az: stripHtml_(an.note || '')
      });
    }

    if (rn.text || an.text) {
      fields.push({
        label: 'Задание ' + taskNum + ' — текст',
        ru: stripHtml_(rn.text || ''),
        az: stripHtml_(an.text || '')
      });
    }

    if (rn.template || an.template) {
      fields.push({
        label: 'Задание ' + taskNum + ' — шаблон',
        ru: stripHtml_(rn.template || ''),
        az: stripHtml_(an.template || '')
      });
    }

    const ruQuestion = extractQuestionTextForTextsSheet_(rn.config_json);
    const azQuestion = extractQuestionTextForTextsSheet_(an.config_json);
    if (ruQuestion || azQuestion) {
      fields.push({
        label: 'Задание ' + taskNum + ' — вопрос',
        ru: stripHtml_(ruQuestion),
        az: stripHtml_(azQuestion)
      });
    }

    const ruAnswers = extractAnswersTextForTextsSheet_(rn.config_json);
    const azAnswers = extractAnswersTextForTextsSheet_(an.config_json);
    if (ruAnswers || azAnswers) {
      fields.push({
        label: 'Задание ' + taskNum + ' — ответы',
        ru: ruAnswers,
        az: azAnswers
      });
    }

    const ruColumns = formatMatchingColumnsForQa_(rn.config_json);
    const azColumns = formatMatchingColumnsForQa_(an.config_json);
    if (ruColumns || azColumns) {
      fields.push({
        label: 'Задание ' + taskNum + ' — колонки',
        ru: ruColumns,
        az: azColumns
      });
    }

    const ruCategorization = extractMathCategorizationMappingForQa_(rn);
    const azCategorization = extractMathCategorizationMappingForQa_(an);
    if (ruCategorization || azCategorization) {
      fields.push({
        label: 'Задание ' + taskNum + ' — категории',
        ru: ruCategorization,
        az: azCategorization
      });
    }

    const ruGrid = extractCoordinateGridForQa_(rn);
    const azGrid = extractCoordinateGridForQa_(an);
    if (ruGrid || azGrid) {
      fields.push({
        label: 'Задание ' + taskNum + ' — сетка (предметы и ответы)',
        ru: ruGrid,
        az: azGrid
      });
    }

    const ruTyping = extractTypingPatternForQa_(rn.config_json);
    const azTyping = extractTypingPatternForQa_(an.config_json);
    if (ruTyping || azTyping) {
      fields.push({
        label: 'Задание ' + taskNum + ' — тренажёр',
        ru: ruTyping,
        az: azTyping
      });
    }
  });

  return fields;
}

function extractQuestionTextForTextsSheet_(cfgJson) {
  if (!cfgJson) return '';
  try {
    const cfg = JSON.parse(cfgJson);
    return clean_(
      getPath_(cfg, 'multipleChoice.questionText') ||
      getPath_(cfg, 'mechanic.question') ||
      ''
    );
  } catch (e) {
    return '';
  }
}

function extractAnswersTextForTextsSheet_(cfgJson) {
  if (!cfgJson) return '';
  try {
    const cfg = JSON.parse(cfgJson);
    const answers =
      (cfg.mechanic && cfg.mechanic.answers) ||
      (cfg.multipleChoice && cfg.multipleChoice.options);

    if (!Array.isArray(answers) || !answers.length) return '';

    return answers.map(function(x, idx) {
      const text = stripAnswerTextForQa_(x && (x.text || x.content || x.title || ''));
      const mark = x && x.isCorrect ? '✓' : '✗';
      return (idx + 1) + '. ' + mark + ' ' + text;
    }).join('\n');
  } catch (e) {
    return '';
  }
}

// Координатная сетка с перетаскиванием: предметы (elements[].html, type=text-dnd)
// + ответы из дочерних problem (mechanic.name / checkExpression / answer).
function extractCoordinateGridForQa_(node) {
  if (!node) return '';

  const cfg = safeParse_(node.config_json) || {};
  const mech = cfg.mechanic || {};

  const elements = (mech && !Array.isArray(mech) && Array.isArray(mech.elements))
    ? mech.elements : [];

  const items = elements
    .filter(function (el) { return el && el.type === 'text-dnd'; })
    .map(function (el, i) {
      const label = stripHtml_(clean_(el.html)) || clean_(el.name) || String(i + 1);
      return (i + 1) + '. ' + label;
    });

  const childProblems = safeArray_(node._childrenRaw)
    .filter(function (ch) { return clean_(ch.type) === 'problem'; });

  const answers = childProblems.map(function (ch) {
    const chCfg = safeParse_(ch.config_json) || {};
    const chMech = chCfg.mechanic || {};
    const name = clean_(chMech.name);
    const expr = clean_(chMech.checkExpression) || clean_(ch.checkExpression);
    const ans = clean_(chMech.answer);
    if (expr) return expr;
    if (name && ans) return name + '=' + ans;
    return name || ans;
  }).filter(Boolean);

  if (!items.length && !answers.length) return '';

  const parts = [];
  if (items.length) parts.push('Предметы: ' + items.join('  '));
  if (answers.length) parts.push('Ответы: ' + answers.join(', '));
  return parts.join('\n');
}
