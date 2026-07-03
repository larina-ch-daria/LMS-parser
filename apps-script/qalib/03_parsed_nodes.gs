// ============================================================
// STAGE 4 — PARSE NODES
// Читает RAW_API, рекурсивно разбирает дерево узлов каждого
// уровня в плоскую таблицу RAW_PARSED_NODES: механики, тексты,
// изображения, ссылки, конфиг-флаги, задачник (uploader).
// ============================================================

function buildRawParsedNodes(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('RAW_API');
  if (!rawSheet) throw new Error('Лист RAW_API не найден');

  const outSheet = ss.getSheetByName('RAW_PARSED_NODES') || ss.insertSheet('RAW_PARSED_NODES');

  const values = rawSheet.getDataRange().getValues();
  if (values.length < 2) throw new Error('RAW_API пуст');

  const headers = values[0].map(h => String(h || '').trim());
  const rows = values.slice(1);

  const idxLevelId = headers.indexOf('level_id');
  const idxData = headers.indexOf('data');

  if (idxLevelId === -1 || idxData === -1) {
    throw new Error('В RAW_API нужны колонки level_id и data');
  }

  // RAW_API накапливает строки между прогонами и не чистится кнопкой
  // "Очистить LMS_LEVELS". Оставляем только самую свежую запись по level_id.
  const latestRowByLevel = {};
  rows.forEach(row => {
    const lid = String(row[idxLevelId] || '').trim();
    if (!lid) return;
    latestRowByLevel[lid] = row;
  });
  const latestRows = Object.keys(latestRowByLevel).map(k => latestRowByLevel[k]);

  const outRows = [];
  let visitIndex = 0;

  latestRows.forEach(row => {
    const mainLevelId = String(row[idxLevelId] || '').trim();
    const raw = row[idxData];
    if (!mainLevelId || !raw) return;
    let payload;
    try {
      payload = typeof raw === 'string' ? JSON.parse(raw) : raw;
    } catch (e) {
      Logger.log('Не удалось распарсить JSON для level_id=' + mainLevelId + ': ' + e);
      return;
    }

    const data = payload && payload.data ? payload.data : payload;
    if (!data || typeof data !== 'object') return;

    if (data.__TRUNCATED__) {
      const markerRow = makeTruncatedMarkerRow_(mainLevelId, data.__original_length__);
      outRows.push(markerRow);
      visitIndex++;
      return;
    }

    visitIndex = walkNode_(data, mainLevelId, 0, '0', outRows, visitIndex);
  });

  const dedupedRows = dedupeParsedNodeRows_(outRows);

  outSheet.clearContents();

  const headerRow = [
    'mainLevelId',
    'nodeId',
    'depth',
    'path',
    'visitIndex',
    'type',
    'config_type',
    'mechanic_type',
    'levelScore',
    'isAutocheck',
    'isAllowFail',
    'force_enabled',
    'isMulti',
    'bigSize',
    'keyboard',
    'keyboard_lang',
    'language_code',
    'timerType',
    'timerSeconds',
    'treePosition',
    'children_count',
    'description',
    'note',
    'text',
    'template',
    'checkExpression',
    'verificationType',
    'pattern',
    'message_title',
    'message_description',
    'message_question',
    'message_hint',
    'message_congrat',
    'message_fail',
    'variables_json',
    'inputs_json',
    'config_json',
    'images_json',
    'embedded_links_json',
    'hints_count',
    'tasklist_exists',
    'tasklist_visible',
    'tasklist_check_type',
    'tasklist_text',
    'tasklist_links_json',
    'tasklist_images_json',
    'uploader_description',
    'level_links_json',
    'level_materials_json'
  ];

  outSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);

  if (dedupedRows.length) {
    outSheet.getRange(2, 1, dedupedRows.length, headerRow.length).setValues(dedupedRows);
  }

  formatParsedNodesSheet_(outSheet, dedupedRows.length + 1);

  SpreadsheetApp.flush();
  Logger.log('RAW_PARSED_NODES: записано строк ' + dedupedRows.length);
}

function walkNode_(node, mainLevelId, depth, path, outRows, visitIndex) {
  if (!node || typeof node !== 'object') return visitIndex;

  const cfg = node.config || {};
  const mech = cfg.mechanic || {};
  const quiz3 = node.quiz3Data || {};
  const messages = node.messages || {};

  const problemChildren = (node.children || []).filter(c => c && c.type === 'problem');
  const problemChild = problemChildren.length ? problemChildren[0] : null;
  const pCfg = problemChild ? (problemChild.config || {}) : {};
  const pMech = pCfg.mechanic || {};

  const pythonAssets = extractPythonAssetsForNode_(node, cfg, mech, pMech);

  const baseDescription =
    firstNonEmpty_(
      mech.description,
      pMech.description,
      messages.description,
      quiz3.data && quiz3.data.description,
      ''
    );

  const description = mergeTextBlocksForNode_(
    baseDescription,
    pythonAssets.descriptionText
  );

  const note =
    firstNonEmpty_(
      mech.note,
      pMech.note,
      ''
    );

  const problemsCombinedText = extractCombinedProblemTextForNode_(problemChildren);

  const text =
    firstNonEmpty_(
      problemsCombinedText,
      pMech.text,
      mech.text,
      mech.template,
      pMech.template,
      quiz3.data && quiz3.data.template,
      quiz3.template,
      ''
    );

  const template =
    firstNonEmpty_(
      mech.template,
      pMech.template,
      ''
    );

  const checkExpression =
    firstNonEmpty_(
      pMech.checkExpression,
      mech.checkExpression,
      ''
    );

  const verificationType =
    firstNonEmpty_(
      mech.verificationType,
      pMech.verificationType,
      cfg.verificationType,
      ''
    );

  const pattern =
    firstNonEmpty_(
      cfg.pattern,
      mech.pattern,
      pCfg.pattern,
      pMech.pattern,
      ''
    );

  const variables = firstDefined_(
    mech.variables,
    mech.vars,
    pMech.variables,
    pMech.vars,
    null
  );
  const variables_json = variables ? safeJsonStringify_(variables) : '';

  const inputs = firstDefined_(
    quiz3.data && quiz3.data.inputs,
    cfg.inputs,
    null
  );
  const inputs_json = inputs ? safeJsonStringify_(inputs) : '';

  const images = extractImagesFromNode_(node, cfg, mech, pMech, pythonAssets);
  const images_json = images.length ? safeJsonStringify_(images) : '[]';

  const embeddedLinks = extractEmbeddedLinksMergedForQa_(node, cfg, pythonAssets);
  const embedded_links_json = embeddedLinks.length ? safeJsonStringify_(embeddedLinks) : '[]';

  const config_json = (cfg && Object.keys(cfg).length) ? safeJsonStringify_(cfg) : '';

  const mechanicType = firstNonEmpty_(
    mech.type,
    mech.problemType,
    pMech.type,
    pMech.problemType,
    cfg.type,
    ''
  );

  const rawIsAllowFail = firstDefined_(cfg.isAllowFail, mech.isAllowFail, null);
  const rawForceEnabled = firstDefined_(cfg.force_enabled, null);
  const rawIsMulti = firstDefined_(node.isMulti, cfg.isMulti, null);
  const rawBigSize = firstDefined_(cfg.bigSize, null);
  const rawKeyboard = firstDefined_(cfg.keyboard, mech.keyboard, pCfg.keyboard, pMech.keyboard, null);
  const rawKeyboardLang = firstNonEmpty_(
    cfg.keyboard_lang,
    mech.keyboard_lang,
    pCfg.keyboard_lang,
    pMech.keyboard_lang,
    ''
  );
  const rawLanguageCode = firstNonEmpty_(
    cfg.language_code,
    mech.language_code,
    pCfg.language_code,
    pMech.language_code,
    ''
  );
  const rawTimerType = firstNonEmpty_(node.timerType, cfg.timerType, '');
  const rawTimerSeconds = firstDefined_(node.timerSeconds, cfg.timerSeconds, null);
  const rawTreePosition = firstDefined_(node.treePosition, cfg.treePosition, null);

  const isAllowFail =
    rawIsAllowFail === true ? 'TRUE' :
    rawIsAllowFail === false ? 'FALSE' : '';

  const forceEnabled =
    rawForceEnabled === true ? 'TRUE' :
    rawForceEnabled === false ? 'FALSE' : '';

  const isMulti =
    rawIsMulti === true ? 'TRUE' :
    rawIsMulti === false ? 'FALSE' : '';

  const bigSize =
    rawBigSize === true ? 'TRUE' :
    rawBigSize === false ? 'FALSE' :
    (rawBigSize != null ? String(rawBigSize) : '');

  const keyboard =
    rawKeyboard === true ? 'TRUE' :
    rawKeyboard === false ? 'FALSE' :
    (rawKeyboard != null ? String(rawKeyboard) : '');

  const uploaderInfo = extractUploaderTaskListInfo_(node, cfg, mech);

  outRows.push([
    String(mainLevelId || ''),
    String(node.id || ''),
    Number(depth || 0),
    String(path || ''),
    Number(visitIndex || 0),
    String(node.type || ''),
    String(cfg.type || ''),
    String(mechanicType || ''),
    node.levelScore != null ? node.levelScore : '',
    node.isAutocheck != null ? String(node.isAutocheck) : '',
    String(isAllowFail || ''),
    String(forceEnabled || ''),
    String(isMulti || ''),
    String(bigSize || ''),
    String(keyboard || ''),
    String(rawKeyboardLang || ''),
    String(rawLanguageCode || ''),
    String(rawTimerType || ''),
    rawTimerSeconds != null ? String(rawTimerSeconds) : '',
    rawTreePosition != null ? String(rawTreePosition) : '',
    String((node.children || []).length),

    String(description || ''),
    String(note || ''),
    String(text || ''),
    String(template || ''),
    String(checkExpression || ''),
    String(verificationType || ''),
    String(pattern || ''),

    String(messages.title || ''),
    String(messages.description || ''),
    String(messages.question || ''),
    String(messages.hint || ''),
    String(messages.congrat || ''),
    String(messages.fail || ''),

    String(variables_json || ''),
    String(inputs_json || ''),
    String(config_json || ''),
    String(images_json || '[]'),
    String(embedded_links_json || '[]'),
    String((node.hints || []).length),

    String(uploaderInfo.tasklist_exists || ''),
    String(uploaderInfo.tasklist_visible || ''),
    String(uploaderInfo.tasklist_check_type || ''),
    String(uploaderInfo.tasklist_text || ''),
    String(uploaderInfo.tasklist_links_json || '[]'),
    String(uploaderInfo.tasklist_images_json || '[]'),
    String(uploaderInfo.uploader_description || ''),
    String(uploaderInfo.level_links_json || '[]'),
    String(uploaderInfo.level_materials_json || '[]')
  ]);

  visitIndex++;

  const nodeType = clean_(node && node.type);
  const cfgType = clean_(cfg && cfg.type);

  // Контейнерные уровни не раскладываем на внутренние children.
  if (
    depth === 0 && (
      nodeType === 'uploader' || cfgType === 'uploader' ||
      nodeType === 'pdf' || cfgType === 'pdf' ||
      nodeType === 'presentation' || cfgType === 'presentation'
    )
  ) {
    return visitIndex;
  }

  const children = node.children || [];
  children.forEach((child, idx) => {
    if (child) {
      visitIndex = walkNode_(child, mainLevelId, depth + 1, path + '.' + idx, outRows, visitIndex);
    }
  });

  return visitIndex;
}

// ─── uploader / tasklist ─────────────────────────────────────

function extractUploaderTaskListInfo_(node, cfg, mech) {
  const nodeType = clean_(node && node.type);
  const cfgType = clean_(cfg && cfg.type);

  if (nodeType !== 'uploader' && cfgType !== 'uploader') {
    return {
      tasklist_exists: '',
      tasklist_visible: '',
      tasklist_check_type: '',
      tasklist_text: '',
      tasklist_links_json: '[]',
      tasklist_images_json: '[]',
      uploader_description: '',
      level_links_json: '[]',
      level_materials_json: '[]'
    };
  }

  const hints = Array.isArray(node && node.hints) ? node.hints : [];
  const isVisible = cfg && cfg.isTaskListVisible === true;
  const exists = isVisible || hints.length > 0;

  let checkType = '';
  if (exists) {
    checkType = 'Учителем';
  }

  const tasklistText = hints.map(function (h, i) {
    const title = clean_(h && h.title);
    const body = clean_(h && h.text) || clean_(h && h.description);
    if (!title && !body) return '';
    const head = (i + 1) + '. ' + (title || '');
    return body ? (head + '\n' + body) : head;
  }).filter(Boolean).join('\n\n');

  const tasklistLinks = [];
  const tasklistImages = [];

  hints.forEach(function (h, i) {
    const idx = i + 1;

    const url = clean_(h && h.url);
    if (url) {
      tasklistLinks.push({
        name: 'Пункт задачника ' + idx,
        url: normalizeLinkUrl_(url)
      });
    }

    const imgUrl =
      clean_(h && h.imageFull && h.imageFull.url) ||
      clean_(h && h.image && h.image.url) ||
      '';

    if (imgUrl) {
      tasklistImages.push({
        name: 'Пункт задачника ' + idx,
        url: normalizeImageUrl_(imgUrl)
      });
    }
  });

  const levelLinks = [];
  const levelMaterials = [];

  const mechanicImages = Array.isArray(mech && mech.images) ? mech.images : [];
  mechanicImages.forEach(function (img, i) {
    const url = normalizeImageUrl_(img && img.url);
    if (!url) return;
    levelMaterials.push({
      name: clean_(img && img.name) || ('Изображение уровня ' + (i + 1)),
      url: url
    });
  });

  const contentLinks = extractEmbeddedLinksNormalizedForQa_(cfg && cfg.content || '');
  contentLinks.forEach(function (url, i) {
    levelLinks.push({
      name: buildEmbeddedLinkLabel_(url, i + 1),
      url: url
    });
  });

  if (Array.isArray(node && node.files)) {
    node.files.forEach(function (f, i) {
      const url = clean_(f && (f.url || f.path));
      if (!url) return;
      levelMaterials.push({
        name: clean_(f && f.name) || ('Файл ' + (i + 1)),
        url: normalizeLinkUrl_(url)
      });
    });
  }

  if (node && node.pdfData && typeof node.pdfData === 'object') {
    const pdfUrl = clean_(node.pdfData.url || node.pdfData.path);
    if (pdfUrl) {
      levelMaterials.push({
        name: clean_(node.pdfData.name) || 'PDF',
        url: normalizeLinkUrl_(pdfUrl)
      });
    }
  }

  if (node && node.presentation && typeof node.presentation === 'object') {
    const presUrl = clean_(node.presentation.url || node.presentation.path);
    if (presUrl) {
      levelMaterials.push({
        name: clean_(node.presentation.name) || 'Презентация',
        url: normalizeLinkUrl_(presUrl)
      });
    }
  }

  return {
    tasklist_exists: exists ? 'TRUE' : 'FALSE',
    tasklist_visible: isVisible ? 'TRUE' : 'FALSE',
    tasklist_check_type: checkType,
    tasklist_text: tasklistText,
    tasklist_links_json: tasklistLinks.length ? safeJsonStringify_(tasklistLinks) : '[]',
    tasklist_images_json: tasklistImages.length ? safeJsonStringify_(tasklistImages) : '[]',
    uploader_description: clean_(mech && mech.description),
    level_links_json: levelLinks.length ? safeJsonStringify_(levelLinks) : '[]',
    level_materials_json: levelMaterials.length ? safeJsonStringify_(levelMaterials) : '[]'
  };
}

function uniqueStrings_(arr) {
  const out = [];
  const seen = {};
  (arr || []).forEach(v => {
    const s = clean_(v);
    if (!s || seen[s]) return;
    seen[s] = true;
    out.push(s);
  });
  return out;
}

// ─── картинки ────────────────────────────────────────────────

function extractImagesFromNode_(node, cfg, mech, pMech, pythonAssets) {
  const all = [];

  collectImageObjectsForNode_(all, mech && mech.images, 'mechanic.images');
  collectImageObjectsForNode_(all, pMech && pMech.images, 'problem.mechanic.images');
  collectImageObjectsForNode_(all, node && node.comics, 'comics');
  collectImageObjectsForNode_(all, cfg && cfg.images, 'config.images');

  [mech && mech.answers, pMech && pMech.answers].forEach((arr, arrIdx) => {
    if (!Array.isArray(arr)) return;

    arr.forEach(ans => {
      if (!ans || typeof ans !== 'object') return;

      if (ans.image && typeof ans.image === 'object') {
        all.push({
          id: ans.image.id || '',
          url: ans.image.url || '',
          name: ans.image.name || '',
          source: arrIdx === 0 ? 'mechanic.answers.image' : 'problem.answers.image'
        });
      }

      if (Array.isArray(ans.images)) {
        ans.images.forEach(img => {
          if (img && typeof img === 'object') {
            all.push({
              id: img.id || '',
              url: img.url || '',
              name: img.name || '',
              source: arrIdx === 0 ? 'mechanic.answers.images' : 'problem.answers.images'
            });
          }
        });
      }
    });
  });

  [mech && mech.problems, pMech && pMech.problems].forEach((arr, arrIdx) => {
    if (!Array.isArray(arr)) return;

    arr.forEach(problem => {
      if (!problem || typeof problem !== 'object') return;

      if (problem.image && typeof problem.image === 'object') {
        all.push({
          id: problem.image.id || '',
          url: problem.image.url || '',
          name: problem.image.name || '',
          source: arrIdx === 0 ? 'mechanic.problems.image' : 'problem.problems.image'
        });
      }

      if (Array.isArray(problem.images)) {
        problem.images.forEach(img => {
          if (img && typeof img === 'object') {
            all.push({
              id: img.id || '',
              url: img.url || '',
              name: img.name || '',
              source: arrIdx === 0 ? 'mechanic.problems.images' : 'problem.problems.images'
            });
          }
        });
      }
    });
  });

  if (pythonAssets && Array.isArray(pythonAssets.images)) {
    pythonAssets.images.forEach(img => {
      if (img && typeof img === 'object') {
        all.push({
          id: img.id || '',
          url: img.url || '',
          name: img.name || '',
          source: img.source || 'pythonAssets'
        });
      }
    });
  }

  const htmlFields = [
    { html: mech && mech.note,        source: 'mechanic.note' },
    { html: pMech && pMech.note,      source: 'problem.note' },
    { html: mech && mech.description, source: 'mechanic.description' },
    { html: pMech && pMech.description, source: 'problem.description' },
    { html: mech && mech.text,        source: 'mechanic.text' },
    { html: pMech && pMech.text,      source: 'problem.text' },
    { html: mech && mech.template,    source: 'mechanic.template' },
    { html: pMech && pMech.template,  source: 'problem.template' },
    { html: cfg && cfg.content,       source: 'config.content' }
  ];

  htmlFields.forEach(function (item) {
    extractImageObjectsFromHtmlForQa_(item.html, item.source).forEach(function (img) {
      all.push(img);
    });
  });

  const uniq = [];
  const seen = {};

  all.forEach(img => {
    if (!img || typeof img !== 'object') return;

    const url = normalizeImageUrl_(img.url || '');
    if (!url) return;

    const normalized = {
      id: img.id || '',
      url: url,
      name: clean_(img.name) || extractFileNameFromNodeUrl_(url) || 'image',
      source: clean_(img.source)
    };

    const key = normalized.url;

    if (seen[key]) return;
    seen[key] = true;
    uniq.push(normalized);
  });

  return uniq;
}

function extractImageObjectsFromHtmlForQa_(html, source) {
  const s = String(html || '');
  if (!s) return [];

  const out = [];
  const seen = {};
  const re = /<img[^>]+src\s*=\s*["']([^"']+)["'][^>]*>/gi;

  let m;
  while ((m = re.exec(s)) !== null) {
    const rawUrl = clean_(m[1]);
    if (!rawUrl) continue;

    const url = normalizeImageUrl_(rawUrl);
    if (!url || seen[url]) continue;
    seen[url] = true;

    out.push({
      id: '',
      url: url,
      name: extractFileNameFromNodeUrl_(url) || 'image',
      source: source || 'html'
    });
  }

  return out;
}

function extractEmbeddedLinksFromHtmlFieldsForQa_(fields) {
  const out = [];
  const seen = {};

  (fields || []).forEach(function (html) {
    extractEmbeddedLinksNormalizedForQa_(html).forEach(function (url) {
      const s = clean_(url);
      if (!s || seen[s]) return;
      seen[s] = true;
      out.push(s);
    });
  });

  return out;
}

function collectImageObjectsForNode_(out, value, source) {
  if (!Array.isArray(value)) return;

  value.forEach(item => {
    if (!item || typeof item !== 'object') return;

    if (item.url) {
      out.push({
        id: item.id || '',
        url: item.url || '',
        name: item.name || '',
        source: source || ''
      });
    }

    if (item.image && typeof item.image === 'object') {
      out.push({
        id: item.image.id || '',
        url: item.image.url || '',
        name: item.image.name || '',
        source: source || ''
      });
    }

    if (Array.isArray(item.images)) {
      item.images.forEach(img => {
        if (img && typeof img === 'object') {
          out.push({
            id: img.id || '',
            url: img.url || '',
            name: img.name || '',
            source: source || ''
          });
        }
      });
    }
  });
}

function extractPythonAssetsForNode_(node, cfg, mech, pMech) {
  const out = {
    images: [],
    links: [],
    descriptionText: ''
  };

  const textParts = [];

  if (!node || typeof node !== 'object') {
    return out;
  }

  if (Array.isArray(node.hints)) {
    node.hints.forEach((hint, idx) => {
      if (!hint || typeof hint !== 'object') return;

      const title = clean_(hint.title);
      const text = clean_(hint.text);

      if (title) textParts.push(title);
      if (text) textParts.push(text);

      const hintUrl = clean_(hint.url);
      if (hintUrl) {
        if (isImageUrlForNode_(hintUrl)) {
          out.images.push({
            id: hint.id || '',
            url: hintUrl,
            name: title || ('Hint image ' + (idx + 1))
          });
        } else {
          out.links.push(normalizeLinkUrl_(hintUrl));
        }
      }

      if (hint.image && hint.image.url) {
        out.images.push({
          id: hint.image.id || hint.id || '',
          url: hint.image.url,
          name: clean_(hint.image.name) || title || ('Hint image ' + (idx + 1))
        });
      }

      if (hint.imageFull && hint.imageFull.url) {
        out.images.push({
          id: hint.imageFull.id || hint.id || '',
          url: hint.imageFull.url,
          name: clean_(hint.imageFull.name) || title || ('Hint image ' + (idx + 1))
        });
      }
    });
  }

  if (Array.isArray(node.tips)) {
    node.tips.forEach((tipWrap, idx) => {
      const title = clean_(tipWrap.title);
      const description = clean_(tipWrap.description);

      if (title) textParts.push(title);
      if (description) textParts.push(description);

      const videoUrl = clean_(tipWrap.videoUrl);
      if (videoUrl) out.links.push(normalizeLinkUrl_(videoUrl));

      const img =
        (tipWrap.image && typeof tipWrap.image === 'object') ? tipWrap.image : null;

      if (img && img.url) {
        out.images.push({
          id: img.id || '',
          url: img.url,
          name: clean_(img.name) || title || ('Tip image ' + (idx + 1))
        });
      }
    });
  }

  if (Array.isArray(node.files)) {
    node.files.forEach(f => {
      const url = clean_(f && (f.url || f.path));
      if (!url) return;
      out.links.push(normalizeLinkUrl_(url));
    });
  }

  if (node.pdfData && typeof node.pdfData === 'object') {
    const pdfUrl = clean_(node.pdfData.url || node.pdfData.path);
    if (pdfUrl) out.links.push(normalizeLinkUrl_(pdfUrl));
  }

  if (node.presentation && typeof node.presentation === 'object') {
    const presUrl = clean_(node.presentation.url || node.presentation.path);
    if (presUrl) out.links.push(normalizeLinkUrl_(presUrl));
  }

  const quizData = node.quizData || {};

  if (quizData.questionImage && typeof quizData.questionImage === 'object' && quizData.questionImage.url) {
    out.images.push({
      id: quizData.questionImage.id || '',
      url: quizData.questionImage.url,
      name: clean_(quizData.questionImage.name) || 'Question image'
    });
  }

  const quizCandidates = []
    .concat(Array.isArray(quizData.items) ? quizData.items : [])
    .concat(Array.isArray(quizData.choices) ? quizData.choices : [])
    .concat(Array.isArray(quizData.categories) ? quizData.categories : []);

  quizCandidates.forEach((item, idx) => {
    if (!item || typeof item !== 'object') return;

    const img = item.image;
    if (img && typeof img === 'object' && img.url) {
      out.images.push({
        id: img.id || '',
        url: img.url,
        name: clean_(img.name) || ('Quiz image ' + (idx + 1))
      });
    }

    const text = clean_(item.text);
    if (text) textParts.push(text);
  });

  const messages = node.messages || {};
  [
    messages.title,
    messages.description,
    messages.question,
    messages.hint
  ].forEach(v => {
    const s = clean_(v);
    if (s) textParts.push(s);
  });

  out.links = uniqueStrings_(out.links);

  const dedupedImages = [];
  const seenImageUrls = {};

  out.images.forEach(img => {
    if (!img || typeof img !== 'object') return;

    const url = normalizeImageUrl_(img.url || '');
    if (!url || seenImageUrls[url]) return;

    seenImageUrls[url] = true;
    dedupedImages.push({
      id: img.id || '',
      url: url,
      name: clean_(img.name) || extractFileNameFromNodeUrl_(url) || 'image'
    });
  });

  out.images = dedupedImages;
  out.descriptionText = uniqueStrings_(textParts).join('\n\n');

  return out;
}

function extractEmbeddedLinksMergedForQa_(node, cfg, pythonAssets) {
  const out = [];

  extractEmbeddedLinksNormalizedForQa_(cfg && cfg.content || '').forEach(url => out.push(url));

  const mech = (cfg && cfg.mechanic) || {};
  const htmlFieldLinks = extractEmbeddedLinksFromHtmlFieldsForQa_([
    mech.note,
    mech.description,
    mech.text,
    mech.template
  ]);

  htmlFieldLinks.forEach(function (url) {
    out.push(url);
  });

  if (pythonAssets && Array.isArray(pythonAssets.links)) {
    pythonAssets.links.forEach(url => {
      const s = clean_(url);
      if (s) out.push(s);
    });
  }

  return uniqueStrings_(out);
}

function mergeTextBlocksForNode_(baseText, extraText) {
  const parts = [];
  const a = clean_(baseText);
  const b = clean_(extraText);

  if (a) parts.push(a);
  if (b) parts.push(b);

  return uniqueStrings_(parts).join('\n\n');
}

function isImageUrlForNode_(url) {
  const s = clean_(url).toLowerCase();
  if (!s) return false;
  return /\.(png|jpg|jpeg|gif|webp|bmp|svg)(\?|#|$)/i.test(s);
}

function extractFileNameFromNodeUrl_(url) {
  const s = clean_(url);
  if (!s) return '';
  const cleaned = s.split('?')[0].split('#')[0];
  const parts = cleaned.split('/');
  return parts.length ? parts[parts.length - 1] : '';
}

function extractEmbeddedLinksForQa_(html) {
  const s = String(html || '');
  if (!s) return [];

  const links = [];
  const seen = {};

  const patterns = [
    /\bsrc\s*=\s*["']([^"']+)["']/gi,
    /\bhref\s*=\s*["']([^"']+)["']/gi
  ];

  patterns.forEach(re => {
    let m;
    while ((m = re.exec(s)) !== null) {
      const url = String(m[1] || '').trim();
      if (!url) continue;
      if (seen[url]) continue;
      seen[url] = true;
      links.push(url);
    }
  });

  return links;
}

function normalizeKnownEmbedLinkForQa_(url) {
  const s = String(url || '').trim();
  if (!s) return '';

  let m = s.match(/docs\.google\.com\/presentation\/d\/([^\/\?'"#]+)/i);
  if (m) {
    return 'https://docs.google.com/presentation/d/' + m[1];
  }

  m = s.match(/docs\.google\.com\/document\/d\/([^\/\?'"#]+)/i);
  if (m) {
    return 'https://docs.google.com/document/d/' + m[1];
  }

  m = s.match(/docs\.google\.com\/spreadsheets\/d\/([^\/\?'"#]+)/i);
  if (m) {
    return 'https://docs.google.com/spreadsheets/d/' + m[1];
  }

  m = s.match(/docs\.google\.com\/forms\/d\/([^\/\?'"#]+)/i);
  if (m) {
    return 'https://docs.google.com/forms/d/' + m[1];
  }

  return s;
}

function extractEmbeddedLinksNormalizedForQa_(html) {
  const raw = extractEmbeddedLinksForQa_(html);
  const out = [];
  const seen = {};

  raw.forEach(url => {
    const normalized = normalizeKnownEmbedLinkForQa_(url);
    if (!normalized || seen[normalized]) return;
    seen[normalized] = true;
    out.push(normalized);
  });

  return out;
}

function dedupeParsedNodeRows_(rows) {
  const out = [];
  const seen = {};

  rows.forEach(r => {
    const key = [
      r[0], // mainLevelId
      r[1], // nodeId
      r[3], // path
      r[5], // type
      r[36] // config_json
    ].join('||');

    if (seen[key]) return;
    seen[key] = true;
    out.push(r);
  });

  return out;
}

function formatParsedNodesSheet_(sheet, rowCount) {
  if (!sheet) return;

  const widths = [
    120, 100, 60, 100, 80,
    140, 140, 160, 80, 90, 90, 90,
    80, 80, 80, 120, 120, 110, 110, 100, 90,
    300, 250, 220, 220, 160, 140, 220,
    180, 220, 220, 220, 220, 220,
    220, 220, 450, 250, 320, 90,
    110, 110, 150, 420, 260, 260, 320, 260, 260
  ];

  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  if (rowCount > 0) {
    sheet.getRange(1, 1, rowCount, widths.length)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
      .setVerticalAlignment('top');

    sheet.getRange(1, 1, 1, widths.length)
      .setFontWeight('bold')
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff');
  }

  sheet.setFrozenRows(1);
}

function firstDefined_() {
  for (let i = 0; i < arguments.length; i++) {
    if (arguments[i] !== undefined && arguments[i] !== null) return arguments[i];
  }
  return null;
}

function extractCombinedProblemTextForNode_(problemChildren) {
  if (!Array.isArray(problemChildren) || !problemChildren.length) return '';

  const parts = [];

  problemChildren.forEach(function (child) {
    if (!child || typeof child !== 'object') return;

    const cfg = child.config || {};
    const mech = cfg.mechanic || {};

    const text = clean_(mech.text);
    const template = clean_(mech.template);

    if (text) parts.push(text);
    else if (template) parts.push(template);
  });

  return parts.join('\n');
}

function makeTruncatedMarkerRow_(mainLevelId, originalLength) {
  const row = new Array(49).fill('');
  row[0] = String(mainLevelId || '');
  row[1] = 'TRUNCATED';
  row[2] = 0;
  row[3] = '0';
  row[4] = 0;
  row[5] = '__TOO_BIG__';
  row[20] = '0';
  row[21] = 'Уровень слишком большой (>50k символов), не загружен полностью. Проверить вручную. Исходный размер: ' + (originalLength || '?');
  row[37] = '[]';
  row[38] = '[]';
  row[40] = '';
  row[44] = '[]';
  row[45] = '[]';
  row[47] = '[]';
  row[48] = '[]';
  return row;
}
