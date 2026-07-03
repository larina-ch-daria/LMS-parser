// ============================================================
// QA SHEET WRITE
// Запись данных в ячейки листа: изображения через =IMAGE(),
// гиперссылки через RichText и HYPERLINK(),
// нормализация URL для формул, именование листов.
// ============================================================

function normalizeUrlForSheetFormula_(url) {
  let s = String(url || '').trim();
  if (!s) return '';
  s = s.replace(/%25([0-9A-Fa-f]{2})/g, '%$1');
  s = s.replace(/ /g, '%20');
  return s;
}

function escapeForFormula_(s) {
  return String(s || '').replace(/"/g, '""');
}

function makeUniqueSheetName_(ss, baseName) {
  let name = baseName;
  let i = 2;
  if (name.length > 100) name = name.substring(0, 100);
  while (ss.getSheetByName(name)) {
    const suffix = '_' + i;
    const maxBaseLen = 100 - suffix.length;
    name = baseName.substring(0, maxBaseLen) + suffix;
    i++;
  }
  return name;
}

function sanitizeSheetName_(name) {
  let s = String(name || '').trim();
  s = s.replace(/[\[\]\*\?\/\\:]/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();
  if (!s) s = 'Без_названия';
  if (s.length > 70) s = s.substring(0, 70).trim();
  s = s.replace(/\s+/g, '_');
  return s;
}

function getLessonName_(pairsData, nodesData, issuesData, lmsLevelsData) {
  let lessonNote = '';
  if (lmsLevelsData && lmsLevelsData.length) {
    const rowWithNote = lmsLevelsData.find(r => clean_(r.lessonNote));
    if (rowWithNote) lessonNote = clean_(rowWithNote.lessonNote);
  }
  if (lessonNote) return sanitizeSheetName_(lessonNote);

  const sources = [];
  if (lmsLevelsData && lmsLevelsData.length) {
    sources.push(lmsLevelsData[0].lessonTitle);
    sources.push(lmsLevelsData[0].pageTitle);
    sources.push(lmsLevelsData[0].publicName);
  }
  if (pairsData && pairsData.length) {
    sources.push(pairsData[0].ru_lessonTitle);
    sources.push(pairsData[0].az_lessonTitle);
    sources.push(pairsData[0].ru_taskTitle);
    sources.push(pairsData[0].az_taskTitle);
    sources.push(pairsData[0].ru_levelTitle);
  }
  if (nodesData && nodesData.length) {
    const root = nodesData.find(n => String(n.depth || '') === '0') || nodesData[0];
    if (root) {
      sources.push(root.lessonTitle);
      sources.push(root.lessonName);
      sources.push(root.title);
      sources.push(root.mainLevelTitle);
      sources.push(root.description);
    }
  }
  if (issuesData && issuesData.length) {
    sources.push(issuesData[0].lessonTitle);
    sources.push(issuesData[0].lessonName);
  }
  const raw = sources.find(v => clean_(v));
  if (!raw) return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  return sanitizeSheetName_(raw);
}

function parseMaterialsToLinks_(text) {
  if (!text) return [];
  const lines = text.split('\n').map(function(s) { return s.trim(); }).filter(Boolean);
  const links = [];
  lines.forEach(function(line) {
    if (/^https?:\/\//i.test(line)) links.push(line);
  });
  return links;
}

function getFormulaSeparator_(ss) {
  const locale = (ss || SpreadsheetApp.getActiveSpreadsheet())
    .getSpreadsheetLocale();
  // EN-локали используют запятую, остальные — точку с запятой
  return /^en/i.test(locale) ? ',' : ';';
}

function setImageLinks_(sheet, row, col, imgs) {
  if (!imgs || imgs.length === 0) return;
  const img = imgs[0];
  if (!img || !img.url) return;
  const sep = getFormulaSeparator_(sheet.getParent());
  const safeUrl = normalizeUrlForSheetFormula_(img.url);
  sheet.getRange(row, col).setFormula(
    '=IMAGE("' + escapeForFormula_(safeUrl) + '"' + sep + '1)'
  );
}

function setEmbeddedLinks_(sheet, row, col, links) {
  if (!links || links.length === 0) return;
  const cell = sheet.getRange(row, col);
  const sep = getFormulaSeparator_(sheet.getParent());

  if (links.length === 1) {
    cell.setFormula(
      '=HYPERLINK("' +
      escapeForFormula_(normalizeUrlForSheetFormula_(links[0].url)) +
      '"' + sep + '"' +
      escapeForFormula_(links[0].name) +
      '")'
    );
    return;
  }

  // Несколько ссылок — через RichText, формулы не нужны
  const rtb = SpreadsheetApp.newRichTextValue();
  let text = '';
  const ranges = [];
  links.forEach((item, i) => {
    const start = text.length;
    text += item.name;
    ranges.push({ start, end: text.length, url: normalizeUrlForSheetFormula_(item.url) });
    if (i < links.length - 1) text += '\n';
  });
  rtb.setText(text);
  ranges.forEach(r => rtb.setLinkUrl(r.start, r.end, r.url));
  cell.setRichTextValue(rtb.build());
}

function setInlineLinks_(sheet, row, col, cellText, links) {
  if (!links || links.length === 0) return;
  const text = String(cellText || '');
  if (!text) return;
  const rtb = SpreadsheetApp.newRichTextValue();
  rtb.setText(text);
  links.forEach(function(l) {
    if (l.url && l.end <= text.length) {
      rtb.setLinkUrl(l.start, l.end, normalizeUrlForSheetFormula_(l.url));
    }
  });
  sheet.getRange(row, col).setRichTextValue(rtb.build());
}

function setMaterialLinks_(sheet, row, col, text) {
  if (!text || text === '—') return;

  const lines = text.split('\n').map(function(s) { return s.trim(); }).filter(Boolean);
  const rtb = SpreadsheetApp.newRichTextValue();
  let fullText = '';
  const links = [];
  let pendingLabel = '';

  lines.forEach(function(line) {
    if (/^https?:\/\//i.test(line)) {
      const label = pendingLabel || line;
      pendingLabel = '';
      const s = fullText.length;
      fullText += label;
      links.push({ s: s, e: fullText.length, url: normalizeUrlForSheetFormula_(line) });
      fullText += '\n';
    } else {
      pendingLabel = line;
    }
  });

  fullText = fullText.replace(/\n$/, '');
  if (!fullText) return;

  rtb.setText(fullText);
  links.forEach(function(l) { rtb.setLinkUrl(l.s, l.e, l.url); });
  sheet.getRange(row, col).setRichTextValue(rtb.build());
}

function setAnswerLinks_(sheet, row, col, items) {
  if (!items || !items.length) return;
  const text = items.map(x => x.label).join('\n');
  if (!text) return;
  const rtb = SpreadsheetApp.newRichTextValue();
  rtb.setText(text);
  let cursor = 0;
  items.forEach((item, idx) => {
    const start = cursor;
    const end = start + item.label.length;
    if (item.url) rtb.setLinkUrl(start, end, normalizeUrlForSheetFormula_(item.url));
    cursor = end + (idx < items.length - 1 ? 1 : 0);
  });
  sheet.getRange(row, col).setRichTextValue(rtb.build());
}
