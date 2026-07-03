// ============================================================
// STAGE 5 — QA MAIN
// Основная логика построения QA-отчёта: сборка строк LEVEL_QA
// и LEVEL_QA_BRIEF, сопоставление уровней (пэринг), проверка
// переводов, работа с изображениями, ссылками и конфиг-флагами.
// ============================================================

function buildLevelQAView(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  try {
    updateQaProgress_(ss, 'Старт сборки QA', 1, 10);
    const sourceSs = ss;

    const pairsSheet = sourceSs.getSheetByName('LMS_PAIRS');
    const nodesSheet = sourceSs.getSheetByName('RAW_PARSED_NODES');
    const lmsLevelsSheet = sourceSs.getSheetByName('LMS_LEVELS');
    const issuesSheet = sourceSs.getSheetByName('LEVEL_ISSUES');

    if (!pairsSheet) throw new Error('Лист LMS_PAIRS не найден');
    if (!nodesSheet) throw new Error('Лист RAW_PARSED_NODES не найден');
    if (!lmsLevelsSheet) throw new Error('Лист LMS_LEVELS не найден');

    updateQaProgress_(ss, 'Читаю данные из листов', 2, 10);

    const pairsData = getSheetData_(pairsSheet);
    setQaLocaleLabelsFromPairs_(pairsData);
    const nodesData = getSheetData_(nodesSheet);
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
    const briefItems = [];

    const lessonMetas = pickLessonMetasFromPairs_(pairsData, levelMetaByMainLevelId);
    addLessonConfigBlock_(rows, lessonMetas.ruMeta, lessonMetas.azMeta);
    const lessonCourseEntries = collectLessonCourseEntriesBySide_(lmsLevelsData, pairsData);
    addLessonCoursesBlockBySide_(rows, lessonCourseEntries.ru, lessonCourseEntries.az);

    const ruMainIdsForMso = {};
    pairsData.forEach(function(p){ const id=clean_(p.ru_mainLevelId); if(id) ruMainIdsForMso[id]=true; });
    addMsoScoringBlock_(rows, lessonMetas.ruMeta, lmsLevelsData, nodesByLevel, ruMainIdsForMso);
    const lessonBriefInfo = buildLessonBriefInfo_(lessonMetas.ruMeta, lessonMetas.azMeta);
    updateQaProgress_(ss, 'Собираю строки основного QA-листа', 3, 10);

    let lastTaskKey = null;
    let briefTaskCounter = 0;
    let briefLevelCounterInTask = 0;
    let currentBriefTask = null;

    pairsData.forEach((pair, pairIndex) => {
      if (pairIndex > 0 && pairIndex % 20 === 0) {
        updateQaProgress_(ss, 'Обработано уровней: ' + pairIndex, 4, 10);
      }

      const ruId = String(pair.ru_mainLevelId || '').trim();
      const azId = String(pair.az_mainLevelId || '').trim();
      if (!ruId && !azId) return;

      const ruLevelMeta = levelMetaByMainLevelId[ruId] || null;
      const azLevelMeta = levelMetaByMainLevelId[azId] || null;

      const ruTaskId = clean_(ruLevelMeta && ruLevelMeta.taskId);
      const azTaskId = clean_(azLevelMeta && azLevelMeta.taskId);
      const currentTaskKey = ruTaskId + '|' + azTaskId;

      if (currentTaskKey !== lastTaskKey) {
        lastTaskKey = currentTaskKey;

        const ruTaskTitle = clean_(ruLevelMeta && ruLevelMeta.taskTitle);
        const azTaskTitle = clean_(azLevelMeta && azLevelMeta.taskTitle);

        rows.push({
          cells: ['📋 Задание', '', ruTaskTitle || '—', azTaskTitle || '—'],
          fmt: 'taskheader'
        });

        const ruTaskType = detectTaskTypeForQa_(ruLevelMeta);
        const azTaskType = detectTaskTypeForQa_(azLevelMeta);

        addCritical_(rows, 'Тип задачи', ruTaskType, azTaskType);

        briefTaskCounter++;
        briefLevelCounterInTask = 0;

        currentBriefTask = {
          taskNumber: briefTaskCounter,
          ruTaskTitle: ruTaskTitle || '—',
          azTaskTitle: azTaskTitle || '—',
          levels: []
        };
        briefItems.push(currentBriefTask);
      }

      const ruNodes = dedupeNodeObjects_((nodesByLevel[ruId] || []).slice()).sort(sortNodesStable_);
      const azNodes = dedupeNodeObjects_((nodesByLevel[azId] || []).slice()).sort(sortNodesStable_);
      attachChildrenForQa_(ruNodes);
      attachChildrenForQa_(azNodes);

      const pairIssue = issuesByPair[String(pair.pairKey || '')] || {};
      const issueColor = String(pairIssue.status_color || '').toUpperCase();

      const pairStatus = clean_(pair.pairStatus);
      const pairConfidence = clean_(pair.pairConfidence);
      const pairingNote = clean_(pair.pairingNote);
      const pairingReason = clean_(pair.pairingReason);

      let statusIcon = '🟢';
      if (issueColor === 'RED') statusIcon = '🔴';
      else if (issueColor === 'YELLOW') statusIcon = '🟡';
      else if (pairStatus === 'REORDERED_MATCH' || pairConfidence === 'TENTATIVE') statusIcon = '🟠';

      const levelHeaderRowIndex = rows.length;

      rows.push({
        cells: [
          'Уровень ' + clean_(pair.orderInTask) + '  ' + statusIcon,
          '',
          pair.ru_levelTitle || (ruLevelMeta && ruLevelMeta.levelTitle) || '',
          pair.az_levelTitle || (azLevelMeta && azLevelMeta.levelTitle) || ''
        ],
        fmt: 'header'
      });

      const detailStart = rows.length;

      rows.push({
        cells: ['ID', '', formatDisplayValueForQa_(ruId), formatDisplayValueForQa_(azId)],
        fmt: 'field'
      });

      rows.push({
        cells: [
          'Ссылка на уровень',
          '',
          formatDisplayValueForQa_(buildLevelUrlFromMeta_(ruLevelMeta) || clean_(pair.ru_pageUrl)),
          formatDisplayValueForQa_(buildLevelUrlFromMeta_(azLevelMeta) || clean_(pair.az_pageUrl))
        ],
        fmt: 'field'
      });

      const showPairRow =
        (pairStatus && pairStatus !== 'MATCHED') ||
        (pairConfidence && pairConfidence !== 'HIGH') ||
        clean_(pairingNote);

      if (showPairRow) {
        rows.push({
          cells: [
            'Пара уровней',
            '🟠',
            formatDisplayValueForQa_([pairStatus, pairConfidence].filter(Boolean).join(' / ')),
            formatDisplayValueForQa_([pairingNote, pairingReason].filter(Boolean).join('\n'))
          ],
          fmt: 'check'
        });
      }

      const ruRoot = findMainLevelNode_(ruNodes);
      const azRoot = findMainLevelNode_(azNodes);
      if ((ruRoot && ruRoot.type === '__TOO_BIG__') || (azRoot && azRoot.type === '__TOO_BIG__')) {
        rows.push({
          cells: ['⚠ Уровень не загружен', '🔴',
            ruRoot && ruRoot.type === '__TOO_BIG__' ? clean_(ruRoot.description) : '—',
            azRoot && azRoot.type === '__TOO_BIG__' ? clean_(azRoot.description) : '—'],
          fmt: 'critical'
        });
      }

      const ruCasingIssues = checkConfigKeyCasing_(ruRoot && ruRoot.config_json);
      const azCasingIssues = checkConfigKeyCasing_(azRoot && azRoot.config_json);

      if (ruCasingIssues.length || azCasingIssues.length) {
        const formatCasingIssues_ = function(issues) {
          if (!issues.length) return '—';
          return issues.map(function(i) {
            return i.kind === 'casing'
              ? '⚠ регистр: ' + i.found + ' → ' + i.expected
              : '? похоже на: ' + i.found + ' → ' + i.expected;
          }).join('\n');
        };

        rows.push({
          cells: [
            '⚠ Теги конфига',
            '🔴',
            formatCasingIssues_(ruCasingIssues),
            formatCasingIssues_(azCasingIssues)
          ],
          fmt: 'critical'
        });
      }
      if (ruRoot || azRoot) {
        const rr = ruRoot || {};
        const ar = azRoot || {};

        addCritical_(rows, 'Механика', rr.mechanic_type || rr.type, ar.mechanic_type || ar.type);
        addCritical_(rows, 'Баллы', rr.levelScore, ar.levelScore);
        addCritical_(rows, 'Автопроверка', rr.isAutocheck, ar.isAutocheck);
        addCritical_(rows, 'isAllowFail', rr.isAllowFail, ar.isAllowFail);
        addCritical_(rows, 'force_enabled', rr.force_enabled, ar.force_enabled);
        addCheck_(rows, 'Защита от копирования', detectCopyProtectionForQa_(ruNodes), detectCopyProtectionForQa_(azNodes));

        const ruRootTests = extractPythonCodeTestsForQa_(rr.config_json);
        const azRootTests = extractPythonCodeTestsForQa_(ar.config_json);

        if (ruRootTests.testInput || azRootTests.testInput) {
          addCritical_(rows, 'testInput', ruRootTests.testInput, azRootTests.testInput);
        }

        if (ruRootTests.testOutput || azRootTests.testOutput) {
          addCritical_(rows, 'testOutput', ruRootTests.testOutput, azRootTests.testOutput);
        }

        addTranslationRow_(rows, 'Описание (root)', rr.description, ar.description);

        if (rr.inputs_json || ar.inputs_json) {
          addCritical_(rows, 'Inputs (root)', formatJson_(rr.inputs_json), formatJson_(ar.inputs_json));
        }

        const ruRootImages = collectNodeImagesForQa_(rr);
        const azRootImages = collectNodeImagesForQa_(ar);

        if (ruRootImages.length || azRootImages.length) {
          addImageRowFromItems_(
            rows,
            buildImageRowLabelForQa_('Изображения (root)', ruRootImages, azRootImages),
            ruRootImages,
            azRootImages
          );
        }

        const ruRootEmbedJson = rr.embedded_links_json || '[]';
        const azRootEmbedJson = ar.embedded_links_json || '[]';
        if (hasEmbeddedLinks_(ruRootEmbedJson) || hasEmbeddedLinks_(azRootEmbedJson)) {
          addEmbeddedLinksRow_(rows, 'Встроенные ссылки (root)', ruRootEmbedJson, azRootEmbedJson);
        }

        const isUploaderPair =
          clean_(rr.type) === 'uploader' || clean_(ar.type) === 'uploader' ||
          clean_(rr.config_type) === 'uploader' || clean_(ar.config_type) === 'uploader';

        if (isUploaderPair) {
          addCritical_(rows, 'Задачник есть', normalizeBoolText_(rr.tasklist_exists), normalizeBoolText_(ar.tasklist_exists));
          addCritical_(rows, 'Задачник виден студенту', normalizeBoolText_(rr.tasklist_visible), normalizeBoolText_(ar.tasklist_visible));
          addCritical_(rows, 'Проверка задачника', rr.tasklist_check_type, ar.tasklist_check_type);

          if (rr.uploader_description || ar.uploader_description) {
            addTranslationRow_(rows, 'Описание uploader', rr.uploader_description, ar.uploader_description);
          }

          if (rr.tasklist_text || ar.tasklist_text) {
            addTranslationRow_(rows, 'Текст задачника', rr.tasklist_text, ar.tasklist_text);
          }

          const ruTaskLinksJson = rr.tasklist_links_json || '[]';
          const azTaskLinksJson = ar.tasklist_links_json || '[]';
          if (hasNamedLinks_(ruTaskLinksJson) || hasNamedLinks_(azTaskLinksJson)) {
            addNamedLinksRow_(rows, 'Ссылки в задачнике', ruTaskLinksJson, azTaskLinksJson);
          }

          const ruTaskImagesJson = rr.tasklist_images_json || '[]';
          const azTaskImagesJson = ar.tasklist_images_json || '[]';
          if (hasNamedLinks_(ruTaskImagesJson) || hasNamedLinks_(azTaskImagesJson)) {
            addNamedImageRow_(rows, 'Изображения в задачнике', ruTaskImagesJson, azTaskImagesJson);
          }

          const ruLevelLinksJson = rr.level_links_json || '[]';
          const azLevelLinksJson = ar.level_links_json || '[]';
          if (hasNamedLinks_(ruLevelLinksJson) || hasNamedLinks_(azLevelLinksJson)) {
            addNamedLinksRow_(rows, 'Ссылки уровня', ruLevelLinksJson, azLevelLinksJson);
          }

          const ruLevelMatJson = rr.level_materials_json || '[]';
          const azLevelMatJson = ar.level_materials_json || '[]';

          if (hasNamedLinks_(ruLevelMatJson) || hasNamedLinks_(azLevelMatJson)) {
            const ruLevelMatItems = parseNamedLinks_(ruLevelMatJson);
            const azLevelMatItems = parseNamedLinks_(azLevelMatJson);

            const ruSplit = splitNamedLinksAndImagesForQa_(ruLevelMatItems);
            const azSplit = splitNamedLinksAndImagesForQa_(azLevelMatItems);

            if (ruSplit.links.length || azSplit.links.length) {
              addNamedLinksRowFromItems_(rows, 'Материалы уровня', ruSplit.links, azSplit.links);
            }

            if (ruSplit.images.length || azSplit.images.length) {
              addNamedImageRowFromItems_(rows, 'Изображения уровня', ruSplit.images, azSplit.images);
            }
          }
        }

        addConfigFlags_(rows, rr.config_json, ar.config_json, rr, ar, ruNodes, azNodes);
      }

      const ruTasks = extractComparableTaskNodes_(ruNodes);
      const azTasks = extractComparableTaskNodes_(azNodes);
      const alignedTaskPairs = alignComparableTaskPairs_(ruTasks, azTasks);

      for (let i = 0; i < alignedTaskPairs.length; i++) {
        const pairing = alignedTaskPairs[i];
        const rn = pairing.ru || {};
        const an = pairing.az || {};

        let taskHeader = 'Задание ' + (i + 1);
        if (pairing.isTentative) taskHeader += '  🟠';
        if (pairing.isReordered) taskHeader += '  ↕';

        rows.push({
          cells: [taskHeader, '', rn.type || '—', an.type || '—'],
          fmt: 'subheader'
        });

        if (pairing.isTentative || pairing.isReordered) {
          if (pairing.note) {
            rows.push({
              cells: ['  Пэринг', '🟠', pairing.ruHint || '—', pairing.note],
              fmt: 'check'
            });
          }

          if (clean_(rn.path) || clean_(an.path)) {
            addCheck_(rows, '  path', rn.path, an.path);
          }
        }

        if (clean_(rn.mechanic_type) || clean_(an.mechanic_type)) {
          addCritical_(rows, '  Механика', rn.mechanic_type, an.mechanic_type);
        }

        addCritical_(rows, '  Автопроверка', rn.isAutocheck, an.isAutocheck);

        if (rn.description || an.description) {
          addTranslationRow_(rows, '  Описание', rn.description, an.description);
        }

        if (rn.note || an.note) {
          addTranslationRow_(rows, '  Note / условие', rn.note, an.note);
        }

        if (rn.text || an.text) {
          const label = isAnswerFieldNode_(rn, an) ? '  Поле ответа' : '  Шаблон (text)';
          addTranslationRow_(rows, label, rn.text, an.text);
        }

        if (rn.template || an.template) {
          addTranslationRow_(rows, '  Template', rn.template, an.template);
        }

        const ruGridCheck = extractCoordinateGridForQa_(rn);
        const azGridCheck = extractCoordinateGridForQa_(an);
        const isGridNode = !!(ruGridCheck || azGridCheck);

        if ((rn.checkExpression || an.checkExpression) && !isGridNode) {
          addCritical_(rows, '  checkExpression', rn.checkExpression, an.checkExpression);
        }

        if (rn.verificationType || an.verificationType) {
          addCritical_(rows, '  Верификация', rn.verificationType, an.verificationType);
        }

        if (rn.variables_json || an.variables_json) {
          addCritical_(rows, '  Переменные', formatJson_(rn.variables_json), formatJson_(an.variables_json));
        }

        const ruAnswerStruct = extractAnswersStructure_(rn.config_json);
        const azAnswerStruct = extractAnswersStructure_(an.config_json);
        if (ruAnswerStruct || azAnswerStruct) {
          addCritical_(rows, '  Ответы (структура)', ruAnswerStruct, azAnswerStruct);
        }

        addAnswerLinksRow_(rows, '  Ответы (текст)', rn.config_json, an.config_json);

        const ruTaskImages = dedupeImagesForQa_(
          collectNodeImagesForQa_(rn).concat(extractMatchingColumnImagesForQa_(rn.config_json))
        );
        const azTaskImages = dedupeImagesForQa_(
          collectNodeImagesForQa_(an).concat(extractMatchingColumnImagesForQa_(an.config_json))
        );
        if (ruTaskImages.length || azTaskImages.length) {
          addImageRowFromItems_(
            rows,
            buildImageRowLabelForQa_('  Изображения', ruTaskImages, azTaskImages),
            ruTaskImages,
            azTaskImages
          );
        }
        const ruEmbedJson = rn.embedded_links_json || '[]';
        const azEmbedJson = an.embedded_links_json || '[]';
        if (hasEmbeddedLinks_(ruEmbedJson) || hasEmbeddedLinks_(azEmbedJson)) {
          addEmbeddedLinksRow_(rows, '  Встроенные ссылки', ruEmbedJson, azEmbedJson);
        }

        if ((rn.hints_count && rn.hints_count !== '0') || (an.hints_count && an.hints_count !== '0')) {
          addCheck_(rows, '  Подсказки', rn.hints_count || '0', an.hints_count || '0');
        }

        const ruColumnsText = formatMatchingColumnsForQa_(rn.config_json);
        const azColumnsText = formatMatchingColumnsForQa_(an.config_json);
        if (ruColumnsText || azColumnsText) {
          addTranslationRow_(rows, '  Колонки', ruColumnsText, azColumnsText);
        }

        const ruCategorization = extractMathCategorizationMappingForQa_(rn);
        const azCategorization = extractMathCategorizationMappingForQa_(an);
        if (ruCategorization || azCategorization) {
          addTranslationRow_(rows, '  Категории и ответы', ruCategorization, azCategorization);
        }

        const ruTypingPattern = extractTypingPatternForQa_(rn.config_json);
        const azTypingPattern = extractTypingPatternForQa_(an.config_json);
        if (ruTypingPattern || azTypingPattern) {
          addTranslationRow_(rows, '  Текст тренажёра', ruTypingPattern, azTypingPattern);
        }

        const ruGrid = extractCoordinateGridForQa_(rn);
        const azGrid = extractCoordinateGridForQa_(an);
        if (ruGrid || azGrid) {
          addTranslationRow_(rows, '  Сетка (предметы и ответы)', ruGrid, azGrid);
        }

        const ruTaskTests = extractPythonCodeTestsForQa_(rn.config_json);
        const azTaskTests = extractPythonCodeTestsForQa_(an.config_json);

        if (ruTaskTests.testInput || azTaskTests.testInput) {
          addCritical_(rows, '  testInput', ruTaskTests.testInput, azTaskTests.testInput);
        }

        if (ruTaskTests.testOutput || azTaskTests.testOutput) {
          addCritical_(rows, '  testOutput', ruTaskTests.testOutput, azTaskTests.testOutput);
        }

        addConfigFlags_(rows, rn.config_json, an.config_json, rn, an, ruNodes, azNodes);
      }

      const detailEndBeforeSep = rows.length - 1;

      const finalStatusIcon = getLevelHeaderStatusIconByRows_(
        rows,
        detailStart,
        detailEndBeforeSep,
        statusIcon
      );

      rows[levelHeaderRowIndex].cells[0] =
        'Уровень ' + clean_(pair.orderInTask) + '  ' + finalStatusIcon;

      rows.push({ cells: ['', '', '', ''], fmt: 'sep' });

      const detailEnd = rows.length - 1;

      rowGroups.push({
        start: detailStart,
        end: detailEnd,
        collapsed: finalStatusIcon === '🟢'
      });

      briefLevelCounterInTask++;

      const rootMechanicRu = clean_(ruRoot && (ruRoot.mechanic_type || ruRoot.type));
      const rootMechanicAz = clean_(azRoot && (azRoot.mechanic_type || azRoot.type));
      const levelMechanic = formatMechanicPairForBrief_(rootMechanicRu, rootMechanicAz);

      const levelTitleRu = clean_(pair.ru_levelTitle || (ruLevelMeta && ruLevelMeta.levelTitle) || '');
      const levelTitleAz = clean_(pair.az_levelTitle || (azLevelMeta && azLevelMeta.levelTitle) || '');

      const briefIssues = collectLevelProblemsFromRows_(rows, detailStart, detailEndBeforeSep);

      if (currentBriefTask) {
        currentBriefTask.levels.push({
          levelNumber: briefLevelCounterInTask,
          orderInTask: clean_(pair.orderInTask) || String(briefLevelCounterInTask),
          statusIcon: finalStatusIcon,
          ruLevelTitle: levelTitleRu || '—',
          azLevelTitle: levelTitleAz || '—',
          mechanic: levelMechanic || '—',
          issues: briefIssues
        });
      }
    });

    updateQaProgress_(ss, 'Определяю имя урока и создаю новый файл', 5, 10);

    const lessonName = getLessonName_(pairsData, nodesData, issuesData, lmsLevelsData);
    const fileBaseName = buildQaFileBaseName_(lessonName, lmsLevelsData);
    const targetInfo = createQaSpreadsheetInPersonalFolder_(ss, fileBaseName);
    const targetSs = targetInfo.spreadsheet;

    updateQaProgress_(ss, 'Записываю основной QA-лист', 6, 10);

    const sheetName = makeUniqueSheetName_(targetSs, 'LEVEL_QA_' + lessonName);
    const sheet = targetSs.insertSheet(sheetName);

    const allData = [['Поле', '', QA_BASE_LABEL_, QA_CMP_LABEL_]].concat(rows.map(r => r.cells));
    const formats = ['colheader'].concat(rows.map(r => r.fmt));

    if (allData.length > 0) {
      sheet.getRange(1, 1, allData.length, 4).setValues(allData);
    }

    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 40);
    sheet.setColumnWidth(3, 560);
    sheet.setColumnWidth(4, 560);

    const S = getQaStyles_();

    for (let i = 0; i < formats.length; i++) {
      const style = S[formats[i]] || S.field;
      sheet.getRange(i + 1, 1, 1, 4)
        .setBackground(style.bg)
        .setFontColor(style.fg)
        .setFontWeight(style.bold ? 'bold' : 'normal')
        .setFontSize(style.size)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setVerticalAlignment('top');
    }

    sheet.getRange(1, 2, allData.length, 1).setHorizontalAlignment('center');

    for (let r = 1; r <= allData.length; r++) {
      const label = String(allData[r - 1][0] || '').trim();

      for (let c = 3; c <= 4; c++) {
        const val = String(allData[r - 1][c - 1] || '').trim();

        const isCenteredValue =
          val === '—' ||
          val === '✅' ||
          val === '❌';

        const isScoreRow = label === 'Баллы';
        const isTaskTypeRow = label === 'Тип задачи';

        if (isCenteredValue || isScoreRow || isTaskTypeRow) {
          sheet.getRange(r, c).setHorizontalAlignment('center');
        }
      }

      if (/^(Задание \d+|Уровень )/i.test(label)) {
        sheet.getRange(r, 1).setHorizontalAlignment('center');
      }
    }

    updateQaProgress_(ss, 'Проставляю ссылки, изображения и группировки', 7, 10);

    rows.forEach((r, i) => {
      if (r.fmt === 'img' || r.fmt === 'img_diff') {
        sheet.setRowHeight(i + 2, 150);
        setImageLinks_(sheet, i + 2, 3, r.ruImgs);
        setImageLinks_(sheet, i + 2, 4, r.azImgs);
      }

      if (r.ruEmbeds || r.azEmbeds) {
        setEmbeddedLinks_(sheet, i + 2, 3, r.ruEmbeds);
        setEmbeddedLinks_(sheet, i + 2, 4, r.azEmbeds);
      }

      if (r.ruAnswerLinks || r.azAnswerLinks) {
        setAnswerLinks_(sheet, i + 2, 3, r.ruAnswerLinks || []);
        setAnswerLinks_(sheet, i + 2, 4, r.azAnswerLinks || []);
      }

      if (r.ruTextLinks && r.ruTextLinks.length > 0) {
        setInlineLinks_(sheet, i + 2, 3, r.cells[2], r.ruTextLinks);
      }
      if (r.azTextLinks && r.azTextLinks.length > 0) {
        setInlineLinks_(sheet, i + 2, 4, r.cells[3], r.azTextLinks);
      }

      if (r._ruMaterials) setMaterialLinks_(sheet, i + 2, 3, r._ruMaterials);
      if (r._azMaterials) setMaterialLinks_(sheet, i + 2, 4, r._azMaterials);
      if (r._ruVideo) {
        sheet.getRange(i + 2, 3)
          .setFormula('=HYPERLINK("' + escapeForFormula_(r._ruVideo) + '","▶ Видео")');
      }
      if (r._azVideo) {
        sheet.getRange(i + 2, 4)
          .setFormula('=HYPERLINK("' + escapeForFormula_(r._azVideo) + '","▶ Видео")');
      }
    });

    rowGroups.forEach(g => {
      const s = g.start + 2;
      const e = g.end + 2;
      if (e >= s) {
        sheet.getRange(s, 1, e - s + 1, 1).shiftRowGroupDepth(1);
      }
    });

    rowGroups.forEach(g => {
      try {
        const gr = sheet.getRowGroupAt(g.start + 2, 1);
        if (gr && g.collapsed) {
          gr.collapse();
        }
      } catch (e) {}
    });

    sheet.setFrozenRows(1);

    updateQaProgress_(ss, 'Создаю краткий BRIEF-лист', 8, 10);
    buildLevelQaBriefSheet_(targetSs, lessonName, briefItems, S, lessonBriefInfo);
    buildLevelTextsSheet_(targetSs, lessonName, pairsData, nodesByLevel);

    updateQaProgress_(ss, 'Финальная очистка', 9, 10);
    removeDefaultSheetIfNeeded_(targetSs);

    SpreadsheetApp.flush();

    updateQaProgress_(ss, 'Готово', 10, 10);
    showQaResultDialog_(targetInfo.url, targetInfo.fileName);

    return targetInfo.url;

  } catch (err) {
    Logger.log('Ошибка buildLevelQAView: ' + err.message + '\n' + (err.stack || ''));
    ss.toast('Ошибка: ' + err.message, 'QA', 10);
    throw err;
  }
}

// ─── BRIEF SHEET ─────────────────────────────────────────────

function buildLevelQaBriefSheet_(ss, lessonName, briefItems, S, lessonBriefInfo) {
  const briefSheetName = makeUniqueSheetName_(ss, 'LEVEL_QA_BRIEF_' + lessonName);
  const sheet = ss.insertSheet(briefSheetName);

  const rows = [];
  const formats = [];

  rows.push(['Пункт', 'Статус', 'Механика', 'Комментарий']);
  formats.push('colheader');
  const lessonInfo = lessonBriefInfo || { statusIcon: '🟢', issues: [] };

  rows.push([
    '0. Урок ' + lessonInfo.statusIcon,
    lessonInfo.statusIcon,
    '—',
    lessonInfo.issues.length ? 'Есть расхождения в конфиге урока' : 'Расхождений нет'
  ]);
  formats.push(
    lessonInfo.statusIcon === '🔴' ? 'critical' :
    lessonInfo.statusIcon === '🟡' ? 'check' :
    'field'
  );

  (lessonInfo.issues || []).forEach(function(issue) {
    rows.push([
      '   ↳ ' + issue.label,
      issue.icon || '',
      '',
      issue.comment || ''
    ]);
    formats.push(issue.fmt || 'check');
  });

  rows.push(['', '', '', '']);
  formats.push('sep');

  if (!briefItems || !briefItems.length) {
    rows.push(['Нет данных', '', '', '']);
    formats.push('field');
  } else {
    briefItems.forEach(task => {
      const taskLabel = task.taskNumber + '. Задание';
      const taskTitle = buildBriefTaskTitle_(task.ruTaskTitle, task.azTaskTitle);

      rows.push([taskLabel, '📋', '', taskTitle]);
      formats.push('taskheader');

      (task.levels || []).forEach(level => {
        const levelIndex = task.taskNumber + '.' + level.levelNumber;
        const levelLabel = levelIndex + ' Уровень ' + level.orderInTask;
        const levelTitle = buildBriefLevelTitle_(level.ruLevelTitle, level.azLevelTitle);

        if (level.statusIcon === '🟢') {
          rows.push([
            levelLabel + ' ' + level.statusIcon,
            level.statusIcon,
            level.mechanic || '—',
            levelTitle
          ]);
          formats.push('field');
        } else {
          rows.push([
            levelLabel + ' ' + level.statusIcon,
            level.statusIcon,
            level.mechanic || '—',
            levelTitle
          ]);
          formats.push(level.statusIcon === '🔴' ? 'critical' : 'check');

          (level.issues || []).forEach(issue => {
            rows.push([
              '   ↳ ' + issue.label,
              issue.icon || '',
              '',
              issue.comment || ''
            ]);
            formats.push(issue.fmt || 'check');
          });
        }
      });

      rows.push(['', '', '', '']);
      formats.push('sep');
    });
  }

  sheet.getRange(1, 1, rows.length, 4).setValues(rows);

  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 70);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 760);

  for (let i = 0; i < formats.length; i++) {
    const style = S[formats[i]] || S.field;
    sheet.getRange(i + 1, 1, 1, 4)
      .setBackground(style.bg)
      .setFontColor(style.fg)
      .setFontWeight(style.bold ? 'bold' : 'normal')
      .setFontSize(style.size)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment('top');
  }

  sheet.getRange(1, 1, rows.length, 4).setVerticalAlignment('top');
  sheet.getRange(1, 2, rows.length, 1).setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  return sheet.getName();
}

function buildBriefTaskTitle_(ruTaskTitle, azTaskTitle) {
  const ru = clean_(ruTaskTitle);
  const az = clean_(azTaskTitle);
  if (ru && az && ru !== az) return QA_BASE_LABEL_ + ': ' + ru + '\n' + QA_CMP_LABEL_ + ': ' + az;
  return ru || az || '—';
}

function buildBriefLevelTitle_(ruLevelTitle, azLevelTitle) {
  const ru = clean_(ruLevelTitle);
  const az = clean_(azLevelTitle);
  if (ru && az && ru !== az) return QA_BASE_LABEL_ + ': ' + ru + '\n' + QA_CMP_LABEL_ + ': ' + az;
  return ru || az || '—';
}

function formatMechanicPairForBrief_(ruMech, azMech) {
  const ru = clean_(ruMech);
  const az = clean_(azMech);

  if (ru && az && ru !== az) return ru + ' / ' + az;
  return ru || az || '—';
}

function collectLevelProblemsFromRows_(rows, startIndex, endIndex) {
  const issues = [];

  for (let i = startIndex; i <= endIndex; i++) {
    const row = rows[i];
    if (!row || !row.cells) continue;

    const label = clean_(row.cells[0]);
    const icon = clean_(row.cells[1]);
    const ruVal = clean_(row.cells[2]);
    const azVal = clean_(row.cells[3]);
    const fmt = clean_(row.fmt);

    if (!label) continue;
    if (fmt === 'sep') continue;

    const isProblem =
      fmt === 'critical' ||
      fmt === 'check' ||
      fmt === 'info' ||
      fmt === 'tl_bad' ||
      fmt === 'tl_warn' ||
      fmt === 'tl_embed' ||
      fmt === 'img_diff' ||
      fmt === 'embed' ||
      fmt === 'embed_diff' ||
      icon === '🔴' ||
      icon === '🟠' ||
      icon === '🟡' ||
      icon === '🔎';

    if (!isProblem) continue;

    const briefComment = buildBriefIssueComment_(label, ruVal, azVal, fmt, icon);

    let briefIcon = icon;
    let briefFmt = 'check';

    if (fmt === 'critical' || fmt === 'tl_bad') {
      briefIcon = briefIcon || '🔴';
      briefFmt = 'critical';
    } else if (fmt === 'tl_embed' || fmt === 'embed') {
      briefIcon = briefIcon || '🟡';
      briefFmt = 'info';
    } else if (fmt === 'embed_diff') {
      briefIcon = briefIcon || '🟠';
      briefFmt = 'check';
    } else if (fmt === 'info') {
      briefIcon = briefIcon || '🟡';
      briefFmt = 'info';
    } else {
      briefIcon = briefIcon || '🟠';
      briefFmt = 'check';
    }

    issues.push({
      label: label,
      icon: briefIcon,
      comment: briefComment,
      fmt: briefFmt
    });
  }

  return dedupeBriefIssues_(issues);
}

function buildBriefIssueComment_(label, ruVal, azVal, fmt, icon) {
  const shortRu = shortenForBrief_(ruVal, 140);
  const shortAz = shortenForBrief_(azVal, 140);
  const tag = function(ru, az) {
    return QA_BASE_LABEL_ + '[' + ru + '] / ' + QA_CMP_LABEL_ + '[' + az + ']';
  };

  if (fmt === 'img_diff') return 'Проверить изображения: ' + tag(shortRu, shortAz);
  if (fmt === 'embed_diff') return 'Проверить ссылки/материалы: ' + tag(shortRu, shortAz);
  if (fmt === 'embed') return 'Есть материалы/внешние ссылки уровня, нужна ручная проверка';
  if (fmt === 'tl_embed') return 'Есть встроенные ссылки в тексте, нужна ручная проверка';
  if (fmt === 'tl_bad' || fmt === 'tl_warn') return 'Проверить перевод: ' + tag(shortRu, shortAz);

  if (label === 'Пара уровней' || /Пэринг/i.test(label) || /path/i.test(label)) {
    return [shortRu, shortAz].filter(Boolean).join(' | ');
  }

  return tag(shortRu, shortAz);
}

function shortenForBrief_(s, maxLen) {
  const text = clean_(s).replace(/\n+/g, ' ⏎ ');
  if (!text) return '—';
  if (text.length <= maxLen) return text;
  return text.substring(0, maxLen - 1) + '…';
}

function dedupeBriefIssues_(issues) {
  const out = [];
  const seen = {};

  (issues || []).forEach(item => {
    const key = [
      clean_(item.label),
      clean_(item.icon),
      clean_(item.comment)
    ].join('||');

    if (seen[key]) return;
    seen[key] = true;
    out.push(item);
  });

  return out;
}

// ─── BUILD HELPERS ───────────────────────────────────────────

function buildNodesByLevel_(nodesData) {
  const byLevel = {};
  nodesData.forEach(n => {
    const lid = String(n.mainLevelId || '').trim();
    if (!lid) return;
    if (!byLevel[lid]) byLevel[lid] = [];
    byLevel[lid].push(n);
  });
  return byLevel;
}

function buildLevelMetaByMainLevelId_(lmsLevelsData) {
  const out = {};
  lmsLevelsData.forEach(r => {
    const mainLevelId = String(r.mainLevelId || '').trim();
    if (!mainLevelId) return;
    if (!out[mainLevelId]) out[mainLevelId] = r;
  });
  return out;
}

function pickLessonMetasFromPairs_(pairsData, levelMetaByMainLevelId) {
  let ruMeta = null;
  let azMeta = null;

  for (let i = 0; i < pairsData.length; i++) {
    const pair = pairsData[i];
    const ruId = String(pair.ru_mainLevelId || '').trim();
    const azId = String(pair.az_mainLevelId || '').trim();

    if (!ruMeta && ruId && levelMetaByMainLevelId[ruId]) ruMeta = levelMetaByMainLevelId[ruId];
    if (!azMeta && azId && levelMetaByMainLevelId[azId]) azMeta = levelMetaByMainLevelId[azId];
    if (ruMeta && azMeta) break;
  }

  return { ruMeta: ruMeta || {}, azMeta: azMeta || {} };
}

function addLessonConfigBlock_(rows, ruMeta, azMeta) {
  addTranslationRow_(rows, 'Название урока', ruMeta.lessonTitle, azMeta.lessonTitle);

  rows.push({
    cells: [
      'Ссылка на урок',
      '',
      formatDisplayValueForQa_(buildLessonUrlFromMeta_(ruMeta)),
      formatDisplayValueForQa_(buildLessonUrlFromMeta_(azMeta))
    ],
    fmt: 'field'
  });

  addTranslationRow_(rows, 'Курс', ruMeta.courseTitle, azMeta.courseTitle);

  rows.push({
    cells: [
      'Ссылка на курс',
      '',
      formatDisplayValueForQa_(clean_(ruMeta.courseUrl)),
      formatDisplayValueForQa_(clean_(azMeta.courseUrl))
    ],
    fmt: 'field'
  });

  addCritical_(rows, 'Позиция в курсе',
    clean_(ruMeta.lessonPositionInCourse) || '—',
    clean_(azMeta.lessonPositionInCourse) || '—'
  );

  addCheck_(rows, 'МСО', ruMeta.msoStatus, azMeta.msoStatus);
  addTranslationRow_(rows, 'Публичное имя', ruMeta.publicName, azMeta.publicName);
  addCheck_(rows, 'Есть публичное имя', normalizeBoolText_(ruMeta.hasPublicName), normalizeBoolText_(azMeta.hasPublicName));
  addCheck_(rows, 'Статус урока', ruMeta.lessonStatus, azMeta.lessonStatus);
  addNoteRow_(rows, 'Заметка урока', ruMeta.lessonNote, azMeta.lessonNote);

  const ruMaterials = clean_(ruMeta.lessonMaterials);
  const azMaterials = clean_(azMeta.lessonMaterials);
  const ruVideo     = clean_(ruMeta.lessonVideoUrl);
  const azVideo     = clean_(azMeta.lessonVideoUrl);

  if (ruMaterials || azMaterials) {
    const ruLinks = parseMaterialsToLinks_(ruMaterials);
    const azLinks = parseMaterialsToLinks_(azMaterials);

    const countDiff = ruLinks.length !== azLinks.length;
    const icon = countDiff ? '🔴' : (ruMaterials || azMaterials ? '🔎' : '');
    const fmt  = countDiff ? 'critical' : 'embed';

    rows.push({
      cells: ['📎 Материалы урока', icon, ruMaterials || '—', azMaterials || '—'],
      fmt: fmt,
      _ruMaterials: ruMaterials,
      _azMaterials: azMaterials
    });
  }

  if (ruVideo || azVideo) {
    const oneMissing  = (!ruVideo && azVideo) || (ruVideo && !azVideo);
    const icon = oneMissing ? '🟠' : '🔎';
    const fmt  = oneMissing ? 'check' : 'embed';

    rows.push({
      cells: ['🎬 Видео урока', icon, ruVideo || '—', azVideo || '—'],
      fmt: fmt,
      _ruVideo: ruVideo,
      _azVideo: azVideo
    });
  }

  rows.push({ cells: ['', '', '', ''], fmt: 'sep' });
}

function formatLessonPosition_(position, total) {
  const p = clean_(position);
  const t = clean_(total);
  if (!p && !t) return '—';
  if (p && t) return p + ' / ' + t;
  if (p) return p;
  return '? / ' + t;
}

function addNoteRow_(rows, label, ruVal, azVal) {
  const rvRaw = clean_(ruVal);
  const avRaw = clean_(azVal);

  if (!rvRaw && !avRaw) return;

  var extractNums = function (s) {
    return (s.match(/\d+/g) || []).join(',');
  };

  var ruNums = extractNums(rvRaw);
  var azNums = extractNums(avRaw);

  var icon = '';
  var fmt = 'field';

  if (rvRaw && !avRaw) {
    icon = '🟠';
    fmt = 'check';
  } else if (!rvRaw && avRaw) {
    icon = '🟠';
    fmt = 'check';
  } else if (ruNums !== azNums) {
    icon = '🟠';
    fmt = 'check';
  }

  var note = '';
  if (ruNums !== azNums && ruNums && azNums) {
    note = '\n[числа: ' + QA_BASE_LABEL_ + '[' + ruNums + '] ' + QA_CMP_LABEL_ + '[' + azNums + ']]';
  }

  rows.push({
    cells: [
      label,
      icon,
      formatDisplayValueForQa_(rvRaw),
      formatDisplayValueForQa_(avRaw) + note
    ],
    fmt: fmt
  });
}

function buildLessonUrlFromMeta_(meta) {
  const pageUrl = clean_(meta && meta.pageUrl);
  if (pageUrl) return pageUrl;

  const lessonGuid = clean_(meta && meta.lessonGuid);
  if (lessonGuid) return 'https://lms.alg.academy/lesson/view/' + lessonGuid;

  return '—';
}

function buildLevelUrlFromMeta_(meta) {
  const levelUuid = clean_(meta && meta.levelUuid);
  if (levelUuid) return 'https://lms.alg.academy/level/update/' + levelUuid;
  return '—';
}

function normalizeBoolText_(v) {
  const s = String(v || '').trim().toUpperCase();
  if (s === 'TRUE') return 'Да';
  if (s === 'FALSE') return 'Нет';
  return String(v || '');
}

function dedupeNodeObjects_(nodes) {
  const out = [];
  const seen = {};

  nodes.forEach(n => {
    const key = [
      clean_(n.mainLevelId),
      clean_(n.nodeId),
      clean_(n.path),
      clean_(n.type),
      clean_(n.config_json)
    ].join('||');

    if (seen[key]) return;
    seen[key] = true;
    out.push(n);
  });

  return out;
}

function sortNodesStable_(a, b) {
  const ai = parseInt(a.visitIndex || '0', 10);
  const bi = parseInt(b.visitIndex || '0', 10);
  if (ai !== bi) return ai - bi;

  const ap = String(a.path || '');
  const bp = String(b.path || '');
  if (ap !== bp) return ap < bp ? -1 : 1;

  const ad = parseInt(a.depth || '0', 10);
  const bd = parseInt(b.depth || '0', 10);
  if (ad !== bd) return ad - bd;

  return String(a.nodeId || '').localeCompare(String(b.nodeId || ''));
}

function findRootNode_(nodes) {
  return nodes.find(n => String(n.depth || '0') === '0') || null;
}

function findMainLevelNode_(nodes) {
  if (!nodes || !nodes.length) return null;

  const exactRoot = nodes.find(n => String(n.depth || '0') === '0');
  if (exactRoot) return exactRoot;

  const withFlags = nodes.find(n =>
    String(n.isAllowFail || '').trim() !== '' ||
    String(n.force_enabled || '').trim() !== '' ||
    String(n.isMulti || '').trim() !== '' ||
    String(n.bigSize || '').trim() !== '' ||
    String(n.keyboard || '').trim() !== '' ||
    String(n.keyboard_lang || '').trim() !== '' ||
    String(n.language_code || '').trim() !== ''
  );
  if (withFlags) return withFlags;

  const mathNode = nodes.find(n => String(n.type || '').trim() === 'math');
  if (mathNode) return mathNode;

  return findRootNode_(nodes);
}

function extractComparableTaskNodes_(nodes) {
  const root = findMainLevelNode_(nodes);
  const rootType = clean_(root && root.type);
  const rootCfgType = clean_(root && root.config_type);

  // Для контейнерных root-уровней не выводим внутренние ноды как отдельные задания
  if (
    rootType === 'uploader' || rootCfgType === 'uploader' ||
    rootType === 'pdf' || rootCfgType === 'pdf' ||
    rootType === 'presentation' || rootCfgType === 'presentation'
  ) {
    return [];
  }

  return nodes.filter(n => {
    const depth = parseInt(n.depth || '0', 10);
    if (depth <= 0) return false;

    const type = String(n.type || '').trim();
    const cfgType = String(n.config_type || '').trim();

    if (type === 'problem') return false;
    if (type === 'math-section') return false;
    if (cfgType === 'problem') return false;
    if (cfgType === 'math-section') return false;

    return true;
  }).sort(sortNodesStable_);
}

function isAnswerFieldNode_(rn, an) {
  const types = [
    String(rn.type || ''),
    String(an.type || ''),
    String(rn.config_type || ''),
    String(an.config_type || '')
  ].join('|');

  return /math-fill-blanks-v2|typing|fill-blanks/i.test(types);
}

// ─── TASK MATCHING ───────────────────────────────────────────

function alignComparableTaskPairs_(ruTasks, azTasks) {
  const ru = Array.isArray(ruTasks) ? ruTasks.slice() : [];
  const az = Array.isArray(azTasks) ? azTasks.slice() : [];

  if (!ru.length && !az.length) return [];

  const best = findBestTaskAssignment_(ru, az);
  const pairs = [];
  const usedAz = {};

  for (let i = 0; i < ru.length; i++) {
    const rn = ru[i];
    const azIndex = best.assignment[i];

    if (azIndex === -1 || azIndex === undefined || !az[azIndex]) {
      pairs.push({
        ru: rn,
        az: {},
        isTentative: true,
        isReordered: false,
        note: 'Пара не найдена на стороне АЗ',
        ruHint: buildTaskHint_(rn)
      });
      continue;
    }

    usedAz[azIndex] = true;

    const an = az[azIndex];
    const scored = scoreTaskPairSimilarity_(rn, an);
    const exactByIndex = i === azIndex;
    const isReordered = !exactByIndex;
    const isTentative = scored.score < 10 || isReordered;

    let note = '';
    if (isReordered) note = 'Похоже на перестановку подуровней';
    if (isTentative && !isReordered) note = 'Предположительный пэринг по похожим признакам';
    if (scored.reason) note = note ? note + ' (' + scored.reason + ')' : scored.reason;

    pairs.push({
      ru: rn,
      az: an,
      isTentative: isTentative,
      isReordered: isReordered,
      note: note,
      ruHint: buildTaskHint_(rn)
    });
  }

  for (let j = 0; j < az.length; j++) {
    if (usedAz[j]) continue;

    pairs.push({
      ru: {},
      az: az[j],
      isTentative: true,
      isReordered: false,
      note: 'Лишняя нода на стороне АЗ без пары',
      ruHint: '—'
    });
  }

  return pairs;
}

function findBestTaskAssignment_(ruTasks, azTasks) {
  const ru = Array.isArray(ruTasks) ? ruTasks : [];
  const az = Array.isArray(azTasks) ? azTasks : [];
  const assignment = [];
  const usedAz = {};

  for (let i = 0; i < ru.length; i++) {
    let bestScore = -999999;
    let bestIndex = -1;

    for (let j = 0; j < az.length; j++) {
      if (usedAz[j]) continue;

      const score = scoreTaskPairSimilarity_(ru[i], az[j]).score;
      const finalScore = score - Math.abs(i - j) * 2;

      if (finalScore > bestScore) {
        bestScore = finalScore;
        bestIndex = j;
      }
    }

    if (bestScore < -10) {
      assignment.push(-1);
    } else {
      assignment.push(bestIndex);
      if (bestIndex !== -1) usedAz[bestIndex] = true;
    }
  }

  return {
    score: 0,
    assignment: assignment
  };
}

function scoreTaskPairSimilarity_(rn, an) {
  let score = 0;
  const reasons = [];

  const rType = clean_(rn.type);
  const aType = clean_(an.type);
  if (rType && aType) {
    if (rType === aType) { score += 8; reasons.push('type'); }
    else score -= 6;
  }

  const rCfgType = clean_(rn.config_type);
  const aCfgType = clean_(an.config_type);
  if (rCfgType && aCfgType) {
    if (rCfgType === aCfgType) { score += 10; reasons.push('config_type'); }
    else score -= 8;
  }

  const rMech = clean_(rn.mechanic_type);
  const aMech = clean_(an.mechanic_type);
  if (rMech && aMech) {
    if (rMech === aMech) { score += 12; reasons.push('mechanic'); }
    else score -= 8;
  }

  const rAnswerKind = getAnswerKindSignatureForQa_(rn);
  const aAnswerKind = getAnswerKindSignatureForQa_(an);
  if (rAnswerKind || aAnswerKind) {
    if (rAnswerKind === aAnswerKind) { score += 10; reasons.push('answer_kind'); }
    else score -= 8;
  }

  const rVer = clean_(rn.verificationType);
  const aVer = clean_(an.verificationType);
  if (rVer && aVer) {
    if (rVer === aVer) { score += 6; reasons.push('verification'); }
    else score -= 4;
  }

  const rHasCheck = clean_(rn.checkExpression) ? '1' : '0';
  const aHasCheck = clean_(an.checkExpression) ? '1' : '0';
  if (rHasCheck === aHasCheck) score += 4;
  else score -= 3;

  const rAns = extractAnswersStructure_(rn.config_json);
  const aAns = extractAnswersStructure_(an.config_json);
  if (rAns || aAns) {
    if (rAns === aAns) { score += 9; reasons.push('answers'); }
    else score -= 5;
  }

  const rCols = getColumnsSignatureForQa_(rn.config_json);
  const aCols = getColumnsSignatureForQa_(an.config_json);
  if (rCols || aCols) {
    if (rCols === aCols) { score += 8; reasons.push('columns'); }
    else score -= 4;
  }

  const rTyping = extractTypingPatternForQa_(rn.config_json);
  const aTyping = extractTypingPatternForQa_(an.config_json);
  if (rTyping || aTyping) {
    if (normalizeLongTextKey_(rTyping) === normalizeLongTextKey_(aTyping)) { score += 8; reasons.push('typing'); }
    else score -= 5;
  }

  const rImgCount = countImagesForQa_(rn.images_json);
  const aImgCount = countImagesForQa_(an.images_json);
  if (rImgCount === aImgCount) {
    score += 4;
    if (rImgCount > 0) reasons.push('images');
  } else {
    score -= 3;
  }

  const rHints = parseInt(rn.hints_count || '0', 10) || 0;
  const aHints = parseInt(an.hints_count || '0', 10) || 0;
  if (rHints === aHints) score += 2;
  else score -= 1;

  const rAuto = clean_(rn.isAutocheck);
  const aAuto = clean_(an.isAutocheck);
  if (rAuto && aAuto) {
    if (rAuto === aAuto) score += 2;
    else score -= 2;
  }

  return {
    score: score,
    reason: reasons.join(', ')
  };
}

function getAnswerKindSignatureForQa_(node) {
  if (!node) return '';

  const cfg = safeParse_(node.config_json) || {};

  const hasAnswers =
    Array.isArray(cfg.mechanic && cfg.mechanic.answers) ||
    Array.isArray(cfg.multipleChoice && cfg.multipleChoice.options);

  const hasCheckExpression = !!clean_(node.checkExpression);
  const hasVerification = !!clean_(node.verificationType);
  const hasImages = countImagesForQa_(node.images_json) > 0;
  const hasTextField = isAnswerFieldNode_(node, node);
  const hasColumns = !!getColumnsSignatureForQa_(node.config_json);
  const hasTyping = !!extractTypingPatternForQa_(node.config_json);

  return [
    hasAnswers ? 'answers' : '',
    hasCheckExpression ? 'check' : '',
    hasVerification ? 'verify' : '',
    hasImages ? 'images' : '',
    hasTextField ? 'field' : '',
    hasColumns ? 'columns' : '',
    hasTyping ? 'typing' : ''
  ].filter(Boolean).join('|');
}

function buildTaskHint_(n) {
  if (!n) return '';
  const parts = [];
  if (clean_(n.type)) parts.push(clean_(n.type));
  if (clean_(n.config_type) && clean_(n.config_type) !== clean_(n.type)) parts.push(clean_(n.config_type));
  if (clean_(n.mechanic_type)) parts.push(clean_(n.mechanic_type));
  return parts.join(' / ');
}

function countImagesForQa_(imagesJson) {
  try {
    const arr = JSON.parse(imagesJson || '[]');
    return Array.isArray(arr) ? arr.length : 0;
  } catch (e) {
    return 0;
  }
}

function getColumnsSignatureForQa_(cfgJson) {
  const cfg = safeParse_(cfgJson);
  const cols = getPath_(cfg, 'mechanic.columns');
  if (!Array.isArray(cols) || !cols.length) return '';

  return cols.map(col => {
    const type = clean_(col && col.type);
    const items = Array.isArray(col && col.items) ? col.items.length : 0;
    return type + ':' + items;
  }).join('|');
}

function normalizeLongTextKey_(s) {
  return String(s || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

// ─── SEVERITY ROWS ───────────────────────────────────────────

function addCritical_(rows, label, ruVal, azVal) {
  const rvRaw = clean_(ruVal);
  const avRaw = clean_(azVal);

  const rv = formatDisplayValueForQa_(rvRaw);
  const av = formatDisplayValueForQa_(avRaw);

  if (rvRaw === avRaw) {
    rows.push({ cells: [label, '', rv, av], fmt: 'field' });
    return;
  }

  rows.push({ cells: [label, '🔴', rv, av], fmt: 'critical' });
}

function addCheck_(rows, label, ruVal, azVal) {
  const rvRaw = clean_(ruVal);
  const avRaw = clean_(azVal);

  const rv = formatDisplayValueForQa_(rvRaw);
  const av = formatDisplayValueForQa_(avRaw);

  if (rvRaw === avRaw) {
    rows.push({ cells: [label, '', rv, av], fmt: 'field' });
    return;
  }

  rows.push({ cells: [label, '🟠', rv, av], fmt: 'check' });
}

// ─── CONFIG FLAGS ────────────────────────────────────────────

const CONFIG_FLAGS = [
  { path: 'force_enabled',                         label: 'force_enabled',        sev: 'critical' },
  { path: 'isAllowFail',                           label: 'isAllowFail',          sev: 'critical' },
  { path: 'attempts', label: 'attempts', sev: 'critical' },
  { path: 'mechanic.timer',                        label: 'Таймер',               sev: 'critical' },
  { path: 'mechanic.subtype',                      label: 'Подтип',               sev: 'critical' },
  { path: 'mechanic.verificationType',             label: 'Верификация (cfg)',    sev: 'critical' },
  { path: 'mechanic.needRandomize',                label: 'needRandomize',        sev: 'check' },
  { path: 'mechanic.shouldBeHiddenAfterSolving',   label: 'hideAfterSolving',     sev: 'check' },
  { path: 'mechanic.imageSize',                    label: 'imageSize',            sev: 'check' },
  { path: 'mechanic.testInput',                    label: 'testInput',            sev: 'critical' },
  { path: 'mechanic.testOutput',                   label: 'testOutput',           sev: 'critical' },
  { path: 'language_code',                         label: 'language_code',        sev: 'critical' },
  { path: 'bigSize',                               label: 'bigSize',              sev: 'check' },
  { path: 'pattern',                               label: 'pattern (typing)',     sev: 'check' },
  { path: 'multipleChoice.answerType',             label: 'answerType',           sev: 'critical' },
  { path: 'keyboard',                              label: 'keyboard',             sev: 'critical' },
  { path: 'keyboard_lang',                         label: 'keyboard_lang',        sev: 'critical' }
];

const KNOWN_CONFIG_KEYS_ = [
  'isAllowFail', 'force_enabled', 'isAutocheck', 'isMulti', 'bigSize',
  'keyboard', 'keyboard_lang', 'language_code', 'pattern', 'attempts',
  'timerType', 'timerSeconds', 'treePosition', 'verificationType',
  'needRandomize', 'shouldBeHiddenAfterSolving', 'imageSize', 'subtype',
  'answerType', 'levelScore', 'testInput', 'testOutput'
];

function levenshtein_(a, b) {
  const m = a.length, n = b.length;
  const dp = [];
  for (let i = 0; i <= m; i++) {
    dp[i] = [i];
    for (let j = 1; j <= n; j++) {
      dp[i][j] = i === 0 ? j
        : j === 0 ? i
        : a[i-1] === b[j-1] ? dp[i-1][j-1]
        : 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
    }
  }
  return dp[m][n];
}

function checkConfigKeyCasing_(cfgJson) {
  const cfg = safeParse_(cfgJson);
  if (!cfg || typeof cfg !== 'object') return [];

  // Ключи глубиной до 2: корень, config, mechanic, multipleChoice
  const keysToCheck = [];
  const addKeys = function(obj, depth) {
    if (!obj || typeof obj !== 'object' || depth > 1) return;
    Object.keys(obj).forEach(function(k) {
      keysToCheck.push(k);
      addKeys(obj[k], depth + 1);
    });
  };
  addKeys(cfg, 0);

  const issues = [];
  const seen = {};

  keysToCheck.forEach(function(key) {
    if (seen[key]) return;
    seen[key] = true;

    if (KNOWN_CONFIG_KEYS_.indexOf(key) !== -1) return;

    const keyLower = key.toLowerCase();

    KNOWN_CONFIG_KEYS_.forEach(function(known) {
      if (known.toLowerCase() === keyLower && known !== key) {
        issues.push({ found: key, expected: known, kind: 'casing' });
        return;
      }

      // Порог Левенштейна зависит от длины: короткие теги строже
      const threshold = known.length <= 6 ? 1 : 2;
      if (levenshtein_(key.toLowerCase(), known.toLowerCase()) <= threshold) {
        issues.push({ found: key, expected: known, kind: 'typo' });
      }
    });
  });

  return issues;
}

function addFieldOrCheckAlways_(rows, label, ruVal, azVal) {
  const rvRaw = clean_(ruVal);
  const avRaw = clean_(azVal);

  const rv = formatDisplayValueForQa_(rvRaw);
  const av = formatDisplayValueForQa_(avRaw);

  if (rvRaw === avRaw) {
    rows.push({ cells: [label, '', rv, av], fmt: 'field' });
    return;
  }

  rows.push({ cells: [label, '🟠', rv, av], fmt: 'check' });
}

function addConfigFlags_(rows, ruCfgJson, azCfgJson, ruNode, azNode) {
  const ruCfg = safeParse_(ruCfgJson);
  const azCfg = safeParse_(azCfgJson);

  if (!ruCfg && !azCfg && !ruNode && !azNode) return;

  const isKeyboardTrainer =
    isKeyboardTrainerNodeForQa_(ruNode, ruCfg) ||
    isKeyboardTrainerNodeForQa_(azNode, azCfg);

  const alwaysShowTypingFlags = [
    { path: 'pattern', label: '  ⚙ pattern (typing)' },
    { path: 'language_code', label: '  ⚙ language_code' },
    { path: 'keyboard_lang', label: '  ⚙ keyboard_lang' }
  ];

  const keyboardLikeFlags = {
    keyboard: true,
    keyboard_lang: true,
    language_code: true,
    'pattern (typing)': true
  };

  if (isKeyboardTrainer) {
    alwaysShowTypingFlags.forEach(function(flag) {
      const rv = getConfigFlagValueForQa_(ruCfg, ruNode, flag.path);
      const av = getConfigFlagValueForQa_(azCfg, azNode, flag.path);

      const rvs = safeJsonStringify_(rv);
      const avs = safeJsonStringify_(av);

      if (!clean_(rvs) && !clean_(avs)) return;

      addFieldOrCheckAlways_(rows, flag.label, rvs, avs);
    });
  }

  CONFIG_FLAGS.forEach(flag => {
    // testInput/testOutput уже выводятся отдельной строкой — не дублируем
    if (flag.path === 'mechanic.testInput' || flag.path === 'mechanic.testOutput') return;
    if (
      isKeyboardTrainer &&
      (flag.path === 'pattern' || flag.path === 'language_code' || flag.path === 'keyboard_lang')
    ) {
      return;
    }

    const rv = getConfigFlagValueForQa_(ruCfg, ruNode, flag.path);
    const av = getConfigFlagValueForQa_(azCfg, azNode, flag.path);

    if (rv === undefined && av === undefined) return;

    const rvs = safeJsonStringify_(rv);
    const avs = safeJsonStringify_(av);

    if (rvs === avs) return;

    const isKeyboardLike = !!keyboardLikeFlags[flag.label];

    if (isKeyboardTrainer && isKeyboardLike) {
      addCheck_(rows, '  ⚙ ' + flag.label, rvs, avs);
      return;
    }

    if (flag.sev === 'critical') {
      addCritical_(rows, '  ⚙ ' + flag.label, rvs, avs);
    } else {
      addCheck_(rows, '  ⚙ ' + flag.label, rvs, avs);
    }
  });
}

function safeParse_(s) {
  try {
    return s ? JSON.parse(s) : null;
  } catch (e) {
    return null;
  }
}

// ─── TRANSLATION CHECK ───────────────────────────────────────

function addTranslationRow_(rows, label, ruHtml, azHtml) {
  const ruNoImg = stripImagesFromHtml_(ruHtml);
  const azNoImg = stripImagesFromHtml_(azHtml);

  const ruParsed = stripHtmlWithLinks_(ruNoImg);
  const azParsed = stripHtmlWithLinks_(azNoImg);

  const rv = ruParsed.text;
  const av = azParsed.text;

  if (!rv && !av) return;

  const hasLinks = ruParsed.links.length > 0 || azParsed.links.length > 0;
  const checks = checkTranslation_(rv, av);

  let icon = '';
  let fmt = 'tl_ok';

  if (hasLinks) {
    icon = '🟡';
    fmt = 'tl_embed';
  }

  if (checks.length > 0) {
    const hasCrit = checks.some(c => c.indexOf('🔴') === 0);
    icon = hasCrit ? '🔴' : '🟠';
    fmt = hasCrit ? 'tl_bad' : 'tl_warn';
  }

  const note = checks.length > 0 ? '\n[' + checks.join('; ') + ']' : '';

  rows.push({
    cells: [label, icon, rv || '—', (av || '—') + note],
    fmt: fmt,
    ruTextLinks: ruParsed.links,
    azTextLinks: azParsed.links
  });
}

function checkTranslation_(ru, az) {
  const issues = [];

  if (!ru && az) {
    issues.push('🔴 ' + QA_BASE_LABEL_ + ' пусто');
    return issues;
  }
  if (ru && !az) {
    issues.push('🔴 ' + QA_CMP_LABEL_ + ' пусто');
    return issues;
  }
  if (!ru && !az) return issues;

  const extractNums = function (s) {
    return (String(s || '').match(/\b\d+(?:[.,]\d+)?\b/g) || []).sort().join(',');
  };

  const ruNums = extractNums(ru);
  const azNums = extractNums(az);
  if (ruNums !== azNums) {
    issues.push('🔴 числа: ' + QA_BASE_LABEL_ + '[' + ruNums + '] ' + QA_CMP_LABEL_ + '[' + azNums + ']');
  }

  const ruPh = (String(ru || '').match(/\{\{[^}]+\}\}/g) || []).sort().join(',');
  const azPh = (String(az || '').match(/\{\{[^}]+\}\}/g) || []).sort().join(',');
  if (ruPh !== azPh) {
    issues.push('🔴 плейсхолдеры: ' + QA_BASE_LABEL_ + '[' + ruPh + '] ' + QA_CMP_LABEL_ + '[' + azPh + ']');
  }

  const ratio = ru.length / Math.max(az.length, 1);
  if (ratio > 3 || ratio < 0.33) {
    issues.push('🟠 длина: ' + QA_BASE_LABEL_ + '=' + ru.length + ' ' + QA_CMP_LABEL_ + '=' + az.length);
  }

  const ruOps = (String(ru || '').match(/[=<>!]+/g) || []).sort().join(',');
  const azOps = (String(az || '').match(/[=<>!]+/g) || []).sort().join(',');
  if (ruOps !== azOps && ruOps) {
    issues.push('🟠 операторы различаются');
  }

  const extractArithOps = function(s) {
    return (String(s || '').match(/\/\/|[*\/%]/g) || []).sort().join(',');
  };
  const ruArith = extractArithOps(ru);
  const azArith = extractArithOps(az);
  if (ruArith !== azArith && (ruArith || azArith)) {
    issues.push('🟠 операторы: ' + QA_BASE_LABEL_ + '[' + ruArith + '] ' + QA_CMP_LABEL_ + '[' + azArith + ']');
  }

  return issues;
}

// ─── IMAGES / EMBEDDED LINKS ─────────────────────────────────

const IMG_BASE = 'https://lms.alg.academy';

function hasEmbeddedLinks_(json) {
  try {
    const arr = JSON.parse(json || '[]');
    return Array.isArray(arr) && arr.length > 0;
  } catch (e) {
    return false;
  }
}

function parseEmbeddedLinks_(json) {
  if (!json || json === '[]') return [];
  try {
    const arr = JSON.parse(json);
    if (!Array.isArray(arr)) return [];
    return arr.map((url, idx) => ({
      name: buildEmbeddedLinkLabel_(url, idx + 1),
      url: normalizeUrlForSheetFormula_(String(url || '').trim())
    })).filter(x => x.url);
  } catch (e) {
    return [];
  }
}

function addEmbeddedLinksRow_(rows, label, ruJson, azJson) {
  const ruEmbeds = parseEmbeddedLinks_(ruJson);
  const azEmbeds = parseEmbeddedLinks_(azJson);

  const ruText = ruEmbeds.map(i => i.name).join('\n') || '—';
  const azText = azEmbeds.map(i => i.name).join('\n') || '—';

  const ruUrls = ruEmbeds.map(i => i.url).join('\n');
  const azUrls = azEmbeds.map(i => i.url).join('\n');

  const sameCount = ruEmbeds.length === azEmbeds.length;
  const sameUrls = ruUrls === azUrls;

  let icon = '🔎';
  let fmt = 'embed';

  if (!sameCount || !sameUrls) {
    icon = '🟠';
    fmt = 'embed_diff';
  }

  rows.push({
    cells: [label, icon, ruText, azText],
    fmt: fmt,
    ruEmbeds: ruEmbeds,
    azEmbeds: azEmbeds
  });
}

function formatMatchingColumnsForQa_(cfgJson) {
  const cfg = safeParse_(cfgJson);
  const cols = getPath_(cfg, 'mechanic.columns');

  if (!Array.isArray(cols) || !cols.length) return '';

  return cols.map((col, colIdx) => {
    const type = String((col && col.type) || '');
    const items = Array.isArray(col && col.items) ? col.items : [];

    const lines = items.map((item, itemIdx) => {
      const content = item && item.content;

      if (type === 'image') {
        if (content && typeof content === 'object') {
          return (itemIdx + 1) + '. ' + (
            clean_(content.name) ||
            getFileNameFromUrl_(content.url) ||
            'image'
          );
        }
        return (itemIdx + 1) + '. image';
      }

      const textContent =
        typeof content === 'string'
          ? stripAnswerTextForQa_(content)
          : stripAnswerTextForQa_(clean_(content));

      return (itemIdx + 1) + '. ' + textContent;
    });

    return 'Колонка ' + (colIdx + 1) + ' [' + type + ']:\n' + lines.join('\n');
  }).join('\n\n');
}

// ─── ANSWERS ─────────────────────────────────────────────────

function extractAnswersStructure_(c) {
  if (!c) return '';

  try {
    const cfg = JSON.parse(c);

    const answers =
      (cfg.mechanic && cfg.mechanic.answers) ||
      (cfg.multipleChoice && cfg.multipleChoice.options);

    if (!Array.isArray(answers) || !answers.length) return '';

    const total = answers.length;
    const correct = answers.filter(x => x && x.isCorrect).length;
    const mask = answers.map(x => x && x.isCorrect ? '1' : '0').join('');

    return 'count=' + total + '; correct=' + correct + '; mask=' + mask;
  } catch (e) {
    return '';
  }
}

// ─── UTILS ───────────────────────────────────────────────────

function getSheetData_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h || '').trim());

  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      if (!h) return;
      obj[h] = row[i] != null ? String(row[i]) : '';
    });
    return obj;
  });
}

function formatJson_(s) {
  if (!s || s === '{}' || s === '[]' || s === '') return '';

  try {
    const o = JSON.parse(s);

    if (Array.isArray(o)) {
      return o.map(function (x) {
        return typeof x === 'object' ? JSON.stringify(x) : String(x);
      }).join(', ');
    }

    if (typeof o === 'object' && o !== null) {
      return Object.keys(o).map(function (k) {
        const v = o[k];
        return k + '=' + (typeof v === 'object' ? JSON.stringify(v) : String(v));
      }).join('; ');
    }

    return String(o);
  } catch (e) {
    return String(s).substring(0, 500);
  }
}

function extractTypingPatternForQa_(cfgJson) {
  const cfg = safeParse_(cfgJson);
  const pattern = getPath_(cfg, 'pattern');
  return pattern ? String(pattern) : '';
}

function detectCopyProtectionForQa_(nodes) {
  const list = Array.isArray(nodes) ? nodes : [nodes];
  const re = /user-select\s*:\s*none/i;
  return list.some(function(n) {
    if (!n) return false;
    return [n.config_json, n.note, n.description, n.text, n.template]
      .some(function(f) { return re.test(String(f || '')); });
  }) ? 'TRUE' : 'FALSE';
}

function extractPythonCodeTestsForQa_(cfgJson) {
  const cfg = safeParse_(cfgJson);
  if (!cfg) {
    return { testInput: '', testOutput: '' };
  }

  return {
    testInput: safeJsonStringify_(getPath_(cfg, 'testInput')),
    testOutput: safeJsonStringify_(getPath_(cfg, 'testOutput'))
  };
}

function extractMathCategorizationMappingForQa_(node) {
  if (!node) return '';

  const cfg = safeParse_(node.config_json) || {};
  const mechanic = cfg.mechanic || {};

  const items = Array.isArray(mechanic.problems) ? mechanic.problems : [];
  const itemMap = {};

  items.forEach(item => {
    const id = clean_(item && item.id);
    if (!id) return;

    itemMap[id] = stripHtml_(
      item && (item.text || item.content || item.title || item.name || id)
    ) || id;
  });

  const childProblems = safeArray_(node._childrenRaw);

  const categories = childProblems
    .filter(ch => clean_(ch.type) === 'problem')
    .map(ch => {
      const childCfg = safeParse_(ch.config_json) || {};
      const childMech = childCfg.mechanic || {};

      if (clean_(childMech.type) !== 'category') return null;

      const title = stripHtml_(childMech.text || '');
      const ids = Array.isArray(childMech.problems) ? childMech.problems : [];
      const values = ids.map(id => itemMap[clean_(id)] || clean_(id)).filter(Boolean);

      return {
        title: title || 'Без названия',
        values: values
      };
    })
    .filter(Boolean);

  if (!categories.length) return '';

  return categories.map((cat, idx) => {
    return (idx + 1) + '. ' + cat.title + ' → ' + cat.values.join(', ');
  }).join('\n');
}

function attachChildrenForQa_(nodes) {
  if (!Array.isArray(nodes) || !nodes.length) return;

  const byPath = {};
  nodes.forEach(n => {
    byPath[clean_(n.path)] = n;
  });

  nodes.forEach(n => {
    n._childrenRaw = [];
  });

  nodes.forEach(n => {
    const path = clean_(n.path);
    if (!path || path === '0') return;

    const parentPath = path.split('.').slice(0, -1).join('.');
    if (!parentPath) return;

    const parent = byPath[parentPath];
    if (!parent) return;

    parent._childrenRaw.push(n);
  });
}

function extractAnswersWithLinksForQa_(cfgJson) {
  if (!cfgJson) return [];

  try {
    const cfg = JSON.parse(cfgJson);

    const answers =
      (cfg.mechanic && cfg.mechanic.answers) ||
      (cfg.multipleChoice && cfg.multipleChoice.options);

    if (!Array.isArray(answers) || !answers.length) return [];

    return answers.map(function (x, idx) {
      const rawText = x && (x.text || x.content || x.title || x.id || '');
      const text = stripAnswerTextForQa_(rawText);
      const mark = x && x.isCorrect ? '✓' : '✗';

      let url = '';

      if (x && x.image && x.image.url) {
        url = normalizeImageUrl_(x.image.url);
      } else if (x && Array.isArray(x.images) && x.images.length && x.images[0] && x.images[0].url) {
        url = normalizeImageUrl_(x.images[0].url);
      }

      return {
        label: (idx + 1) + '. ' + mark + ' ' + text,
        url: url
      };
    });
  } catch (e) {
    return [];
  }
}

function addAnswerLinksRow_(rows, label, ruCfgJson, azCfgJson) {
  const ruAnswers = extractAnswersWithLinksForQa_(ruCfgJson);
  const azAnswers = extractAnswersWithLinksForQa_(azCfgJson);

  if (!ruAnswers.length && !azAnswers.length) return;

  const ruText = ruAnswers.map(x => x.label).join('\n') || '—';
  const azText = azAnswers.map(x => x.label).join('\n') || '—';

  rows.push({
    cells: [label, '', ruText, azText],
    fmt: 'field',
    ruAnswerLinks: ruAnswers,
    azAnswerLinks: azAnswers
  });
}

function hasNamedLinks_(json) {
  try {
    const arr = JSON.parse(json || '[]');
    return Array.isArray(arr) && arr.length > 0;
  } catch (e) {
    return false;
  }
}

function parseNamedLinks_(json) {
  if (!json || json === '[]') return [];
  try {
    const arr = JSON.parse(json);
    if (!Array.isArray(arr)) return [];
    return arr.map(function (x, idx) {
      const url = normalizeUrlForSheetFormula_(clean_(x && x.url));
      if (!url) return null;
      return {
        name: clean_(x && x.name) || ('Ссылка ' + (idx + 1)),
        url: url
      };
    }).filter(Boolean);
  } catch (e) {
    return [];
  }
}

function addNamedLinksRow_(rows, label, ruJson, azJson) {
  const ruLinks = parseNamedLinks_(ruJson);
  const azLinks = parseNamedLinks_(azJson);

  const ruText = ruLinks.map(i => i.name).join('\n') || '—';
  const azText = azLinks.map(i => i.name).join('\n') || '—';

  const ruUrls = ruLinks.map(i => i.url).join('\n');
  const azUrls = azLinks.map(i => i.url).join('\n');

  const sameCount = ruLinks.length === azLinks.length;
  const sameUrls = ruUrls === azUrls;

  let icon = '🔎';
  let fmt = 'embed';

  if (!sameCount || !sameUrls) {
    icon = '🟠';
    fmt = 'embed_diff';
  }

  rows.push({
    cells: [label, icon, ruText, azText],
    fmt: fmt,
    ruEmbeds: ruLinks,
    azEmbeds: azLinks
  });
}

function addImageRowFromItems_(rows, label, ruImgs, azImgs) {
  ruImgs = Array.isArray(ruImgs) ? ruImgs : [];
  azImgs = Array.isArray(azImgs) ? azImgs : [];

  const maxLen = Math.max(ruImgs.length, azImgs.length);
  if (maxLen === 0) return;

  const sameCount = ruImgs.length === azImgs.length;

  for (let i = 0; i < maxLen; i++) {
    const r = ruImgs[i] || null;
    const a = azImgs[i] || null;

    const ruName = r ? formatImageLabelForQa_(r) : '—';
    const azName = a ? formatImageLabelForQa_(a) : '—';

    const ruUrl = r ? r.url : '';
    const azUrl = a ? a.url : '';

    let icon = '';
    let fmt = 'img';

    if (!sameCount) {
      icon = '🔴';
      fmt = 'img_diff';
    } else if (ruUrl !== azUrl) {
      icon = '🟠';
      fmt = 'img_diff';
    }

    rows.push({
      cells: [
        i === 0 ? label : '  ',
        i === 0 ? icon : '',
        ruName,
        azName
      ],
      fmt: fmt,
      ruImgs: r ? [r] : [],
      azImgs: a ? [a] : []
    });
  }
}

function addNamedImageRow_(rows, label, ruJson, azJson) {
  const ruImgs = parseNamedLinks_(ruJson);
  const azImgs = parseNamedLinks_(azJson);

  const maxLen = Math.max(ruImgs.length, azImgs.length);
  if (maxLen === 0) return;

  const sameCount = ruImgs.length === azImgs.length;

  for (let i = 0; i < maxLen; i++) {
    const r = ruImgs[i] || null;
    const a = azImgs[i] || null;

    const ruName = r ? r.name : '—';
    const azName = a ? a.name : '—';

    const ruUrl = r ? r.url : '';
    const azUrl = a ? a.url : '';

    let icon = '';
    let fmt = 'img';

    if (!sameCount) {
      icon = '🔴';
      fmt = 'img_diff';
    } else if (ruUrl !== azUrl) {
      icon = '🟠';
      fmt = 'img_diff';
    }

    rows.push({
      cells: [
        i === 0 ? label : '  ',
        i === 0 ? icon : '',
        ruName,
        azName
      ],
      fmt: fmt,
      ruImgs: r ? [r] : [],
      azImgs: a ? [a] : []
    });
  }
}

function formatDisplayValueForQa_(v) {
  const s = clean_(v);
  if (!s) return '—';

  const upper = s.toUpperCase();
  if (upper === 'TRUE') return '✅';
  if (upper === 'FALSE') return '❌';

  return s;
}

function isImageUrlForQa_(url) {
  const s = clean_(url).toLowerCase();
  if (!s) return false;

  return /\.(png|jpg|jpeg|gif|webp|bmp|svg)(\?|#|$)/i.test(s);
}

function splitNamedLinksAndImagesForQa_(items) {
  const out = {
    links: [],
    images: []
  };

  (items || []).forEach(function (item) {
    const url = clean_(item && item.url);
    if (!url) return;

    if (isImageUrlForQa_(url)) {
      out.images.push(item);
    } else {
      out.links.push(item);
    }
  });

  return out;
}

function addNamedLinksRowFromItems_(rows, label, ruLinks, azLinks) {
  ruLinks = Array.isArray(ruLinks) ? ruLinks : [];
  azLinks = Array.isArray(azLinks) ? azLinks : [];

  const ruText = ruLinks.map(i => i.name).join('\n') || '—';
  const azText = azLinks.map(i => i.name).join('\n') || '—';

  const ruUrls = ruLinks.map(i => i.url).join('\n');
  const azUrls = azLinks.map(i => i.url).join('\n');

  const sameCount = ruLinks.length === azLinks.length;
  const sameUrls = ruUrls === azUrls;

  let icon = '🔎';
  let fmt = 'embed';

  if (!sameCount || !sameUrls) {
    icon = '🟠';
    fmt = 'embed_diff';
  }

  rows.push({
    cells: [label, icon, ruText, azText],
    fmt: fmt,
    ruEmbeds: ruLinks,
    azEmbeds: azLinks
  });
}

function addNamedImageRowFromItems_(rows, label, ruImgs, azImgs) {
  ruImgs = Array.isArray(ruImgs) ? ruImgs : [];
  azImgs = Array.isArray(azImgs) ? azImgs : [];

  const maxLen = Math.max(ruImgs.length, azImgs.length);
  if (maxLen === 0) return;

  const sameCount = ruImgs.length === azImgs.length;

  for (let i = 0; i < maxLen; i++) {
    const r = ruImgs[i] || null;
    const a = azImgs[i] || null;

    const ruName = r ? r.name : '—';
    const azName = a ? a.name : '—';

    const ruUrl = r ? r.url : '';
    const azUrl = a ? a.url : '';

    let icon = '';
    let fmt = 'img';

    if (!sameCount) {
      icon = '🔴';
      fmt = 'img_diff';
    } else if (ruUrl !== azUrl) {
      icon = '🟠';
      fmt = 'img_diff';
    }

    rows.push({
      cells: [
        i === 0 ? label : '  ',
        i === 0 ? icon : '',
        ruName,
        azName
      ],
      fmt: fmt,
      ruImgs: r ? [r] : [],
      azImgs: a ? [a] : []
    });
  }
}

function getLevelHeaderStatusIconByRows_(rows, startIndex, endIndex, fallbackIcon) {
  let hasRed = false;
  let hasYellow = false;

  for (let i = startIndex; i <= endIndex; i++) {
    const row = rows[i];
    if (!row) continue;

    const fmt = String(row.fmt || '').trim();
    const icon = String((row.cells && row.cells[1]) || '').trim();

    if (
      fmt === 'critical' ||
      fmt === 'tl_bad' ||
      icon === '🔴'
    ) {
      hasRed = true;
      break;
    }

    if (
      fmt === 'check' ||
      fmt === 'info' ||
      fmt === 'tl_warn' ||
      fmt === 'tl_embed' ||
      fmt === 'img_diff' ||
      fmt === 'embed' ||
      fmt === 'embed_diff' ||
      icon === '🟠' ||
      icon === '🟡' ||
      icon === '🔎'
    ) {
      hasYellow = true;
    }
  }

  if (hasRed) return '🔴';
  if (hasYellow) return '🟡';
  return fallbackIcon || '🟢';
}

function collectLessonCourseEntriesBySide_(lmsLevelsData, pairsData) {
  const ruIds = {};
  const azIds = {};

  (pairsData || []).forEach(function (pair) {
    const ruId = clean_(pair.ru_mainLevelId);
    const azId = clean_(pair.az_mainLevelId);

    if (ruId) ruIds[ruId] = true;
    if (azId) azIds[azId] = true;
  });

  const ru = [];
  const az = [];
  const seenRu = {};
  const seenAz = {};

  (lmsLevelsData || []).forEach(function (row) {
    const mainLevelId = clean_(row.mainLevelId);
    if (!mainLevelId) return;

    const entry = {
      courseTitle: clean_(row.courseTitle),
      courseUrl: clean_(row.courseUrl),
      courseLocale: clean_(row.courseLocale || row.courseLanguage),
      lessonPositionInCourse: clean_(row.lessonPositionInCourse),
      lessonTotalInCourse: clean_(row.lessonTotalInCourse)
    };

    if (!entry.courseTitle && !entry.courseUrl) return;

    const key = [
      entry.courseTitle,
      entry.courseUrl,
      entry.courseLocale,
      entry.lessonPositionInCourse,
      entry.lessonTotalInCourse
    ].join('||');

    if (ruIds[mainLevelId]) {
      if (!seenRu[key]) {
        seenRu[key] = true;
        ru.push(entry);
      }
    }

    if (azIds[mainLevelId]) {
      if (!seenAz[key]) {
        seenAz[key] = true;
        az.push(entry);
      }
    }
  });

  const sorter = function (a, b) {
    const ca = clean_(a.courseTitle);
    const cb = clean_(b.courseTitle);
    if (ca !== cb) return ca.localeCompare(cb);

    const pa = parseInt(a.lessonPositionInCourse || '0', 10) || 0;
    const pb = parseInt(b.lessonPositionInCourse || '0', 10) || 0;
    return pa - pb;
  };

  ru.sort(sorter);
  az.sort(sorter);

  return { ru: ru, az: az };
}

function addLessonCoursesBlockBySide_(rows, ruEntries, azEntries) {
  const ru = Array.isArray(ruEntries) ? ruEntries : [];
  const az = Array.isArray(azEntries) ? azEntries : [];

  if (!ru.length && !az.length) return;

  rows.push({
    cells: ['Курсы, где лежит урок', '', '', ''],
    fmt: 'taskheader'
  });

  const maxLen = Math.max(ru.length, az.length);

  for (let i = 0; i < maxLen; i++) {
    const r = ru[i] || null;
    const a = az[i] || null;

    const ruTitle = formatCourseEntryTitleForQa_(r);
    const azTitle = formatCourseEntryTitleForQa_(a);

    rows.push({
      cells: ['Курс ' + (i + 1), '', ruTitle || '—', azTitle || '—'],
      fmt: 'field'
    });

    rows.push({
      cells: [
        'Позиция в курсе ' + (i + 1),
        '',
        r ? formatLessonPosition_(r.lessonPositionInCourse, r.lessonTotalInCourse) : '—',
        a ? formatLessonPosition_(a.lessonPositionInCourse, a.lessonTotalInCourse) : '—'
      ],
      fmt: 'field'
    });

    rows.push({
      cells: [
        'Ссылка на курс ' + (i + 1),
        '',
        r ? formatDisplayValueForQa_(r.courseUrl) : '—',
        a ? formatDisplayValueForQa_(a.courseUrl) : '—'
      ],
      fmt: 'field'
    });
  }

  rows.push({ cells: ['', '', '', ''], fmt: 'sep' });
}

function detectTaskTypeForQa_(levelMeta) {
  const meta = levelMeta || {};

  const isBonus = clean_(meta.isBonus).toUpperCase();
  const isTheory = clean_(meta.isTheory).toUpperCase();
  const isQuiz = clean_(meta.isQuiz).toUpperCase();

  if (isBonus === 'TRUE') return 'Бонусная';
  if (isQuiz === 'TRUE') return 'Викторина';
  if (isTheory === 'TRUE') return 'Теоретическое';

  return 'Основная';
}

function formatCourseEntryTitleForQa_(entry) {
  if (!entry) return '';

  const title = clean_(entry.courseTitle);
  const locale = clean_(entry.courseLocale);

  if (title && locale) return title + ' [' + locale + ']';
  return title || locale || '';
}

function isKeyboardTrainerNodeForQa_(node, cfg) {
  const nodeType = clean_(node && node.type);
  const cfgType = clean_(node && node.config_type);
  const mechanicType = clean_(node && node.mechanic_type);

  const parsedCfg = cfg || safeParse_(node && node.config_json) || {};

  const keyboard =
    getPath_(parsedCfg, 'keyboard') !== undefined
      ? getPath_(parsedCfg, 'keyboard')
      : clean_(node && node.keyboard);

  const keyboardLang =
    getPath_(parsedCfg, 'keyboard_lang') !== undefined
      ? getPath_(parsedCfg, 'keyboard_lang')
      : clean_(node && node.keyboard_lang);

  const languageCode =
    getPath_(parsedCfg, 'language_code') !== undefined
      ? getPath_(parsedCfg, 'language_code')
      : clean_(node && node.language_code);

  const pattern =
    getPath_(parsedCfg, 'pattern') !== undefined
      ? getPath_(parsedCfg, 'pattern')
      : clean_(node && node.pattern);

  if (/typing/i.test(nodeType)) return true;
  if (/typing/i.test(cfgType)) return true;
  if (/typing/i.test(mechanicType)) return true;

  if (clean_(keyboard) || clean_(keyboardLang) || clean_(languageCode) || clean_(pattern)) {
    return true;
  }

  return false;
}

function buildLessonBriefInfo_(ruMeta, azMeta) {
  const issues = [];

  function pushIssue(label, ruVal, azVal, severity) {
    const rv = clean_(ruVal);
    const av = clean_(azVal);
    if (rv === av) return;

    issues.push({
      label: label,
      icon: severity === 'critical' ? '🔴' : '🟠',
      comment: QA_BASE_LABEL_ + '[' + shortenForBrief_(rv, 120) + '] / ' + QA_CMP_LABEL_ + '[' + shortenForBrief_(av, 120) + ']',
      fmt: severity === 'critical' ? 'critical' : 'check'
    });
  }

  pushIssue('Позиция в курсе', ruMeta.lessonPositionInCourse, azMeta.lessonPositionInCourse, 'critical');
  pushIssue('МСО', ruMeta.msoStatus, azMeta.msoStatus, 'check');
  pushIssue('Публичное имя', ruMeta.publicName, azMeta.publicName, 'check');
  pushIssue('Есть публичное имя', normalizeBoolText_(ruMeta.hasPublicName), normalizeBoolText_(azMeta.hasPublicName), 'check');
  pushIssue('Статус урока', ruMeta.lessonStatus, azMeta.lessonStatus, 'check');

  const ruNote = clean_(ruMeta.lessonNote);
  const azNote = clean_(azMeta.lessonNote);
  const ruNums = (ruNote.match(/\d+/g) || []).join(',');
  const azNums = (azNote.match(/\d+/g) || []).join(',');
  if (ruNote || azNote) {
    if (ruNote !== azNote) {
      issues.push({
        label: 'Заметка урока',
        icon: '🟠',
        comment: QA_BASE_LABEL_ + '[' + shortenForBrief_(ruNote, 120) + '] / ' + QA_CMP_LABEL_ + '[' + shortenForBrief_(azNote, 120) + ']',
        fmt: 'check'
      });
    } else if (ruNums !== azNums) {
      issues.push({
        label: 'Заметка урока',
        icon: '🟠',
        comment: 'Числа различаются: ' + QA_BASE_LABEL_ + '[' + ruNums + '] / ' + QA_CMP_LABEL_ + '[' + azNums + ']',
        fmt: 'check'
      });
    }
  }

  return {
    statusIcon: issues.length ? (issues.some(x => x.fmt === 'critical') ? '🔴' : '🟡') : '🟢',
    issues: issues
  };
}

function getConfigFlagValueForQa_(cfg, node, path) {
  const fromCfg = getPath_(cfg, path);
  if (fromCfg !== undefined) return fromCfg;

  const n = node || {};

  switch (path) {
    case 'force_enabled':
      return clean_(n.force_enabled);

    case 'isAllowFail':
      return clean_(n.isAllowFail);

    case 'language_code':
      return clean_(n.language_code);

    case 'bigSize':
      return clean_(n.bigSize);

    case 'keyboard':
      return clean_(n.keyboard);

    case 'keyboard_lang':
      return clean_(n.keyboard_lang);

    case 'pattern':
      return clean_(n.pattern);

    case 'mechanic.testInput':
      return extractPythonCodeTestsForQa_(n.config_json).testInput || '';

    case 'mechanic.testOutput':
      return extractPythonCodeTestsForQa_(n.config_json).testOutput || '';

    default:
      return undefined;
  }
}

function findFirstNodeWithConfigFlagForQa_(nodes, path) {
  const list = Array.isArray(nodes) ? nodes : [];

  for (let i = 0; i < list.length; i++) {
    const node = list[i];
    const cfg = safeParse_(node && node.config_json);
    const val = getConfigFlagValueForQa_(cfg, node, path);

    if (val !== undefined && val !== null && String(val).trim() !== '') {
      return {
        node: node,
        value: val
      };
    }
  }

  return null;
}
