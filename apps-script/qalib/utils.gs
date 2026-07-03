// ============================================================
// UTILS
// Общие утилитки пайплайна: строковые хелперы, работа с JSON,
// обход объектов по пути, нормализация URL.
// ============================================================

function clean_(v) {
  return v == null ? '' : String(v).trim();
}

function firstNonEmpty_() {
  var items = (arguments.length === 1 && Array.isArray(arguments[0]))
    ? arguments[0]
    : Array.prototype.slice.call(arguments);
  for (var i = 0; i < items.length; i++) {
    var v = items[i];
    if (v != null && String(v).trim() !== '') return v;
  }
  return '';
}

function safeArray_(v) {
  return Array.isArray(v) ? v : [];
}

function safeJsonStringify_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    try { return JSON.stringify(v); } catch (e) { return String(v); }
  }
  return String(v);
}

function getPath_(obj, path) {
  return String(path || '').split('.').reduce(function(acc, key) {
    if (acc == null) return undefined;
    if (/^\d+$/.test(key)) return Array.isArray(acc) ? acc[Number(key)] : undefined;
    return acc[key];
  }, obj);
}

function buildEmbeddedLinkLabel_(url, index) {
  const s = String(url || '').trim();

  if (/docs\.google\.com\/presentation\/d\//i.test(s)) {
    return 'Google Slides ' + index;
  }
  if (/docs\.google\.com\/document\/d\//i.test(s)) {
    return 'Google Doc ' + index;
  }
  if (/docs\.google\.com\/spreadsheets\/d\//i.test(s)) {
    return 'Google Sheet ' + index;
  }
  if (/docs\.google\.com\/forms\/d\//i.test(s)) {
    return 'Google Form ' + index;
  }
  return 'Ссылка ' + index;
}

function normalizeImageUrl_(url) {
  let s = String(url || '').trim();
  if (!s) return '';
  s = s.replace(/%25([0-9A-Fa-f]{2})/g, '%$1');
  if (/^https?:\/\//i.test(s)) {
    return normalizeUrlForSheetFormula_(s);
  }
  return normalizeUrlForSheetFormula_(IMG_BASE + s);
}

function normalizeLinkUrl_(url) {
  const s = String(url || '').trim();
  if (!s) return '';
  if (/^https?:\/\//i.test(s)) return s;
  if (/^\//.test(s)) return 'https://lms.alg.academy' + s;
  return s;
}

function buildRawApiMapForPairs_(rawSheet) {
  const values = rawSheet.getDataRange().getValues();
  if (values.length < 2) return {};

  const headers = values[0].map(h => String(h || '').trim());
  const rows = values.slice(1);

  const idxLevelId = headers.indexOf('level_id');
  const idxData = headers.indexOf('data');

  if (idxLevelId === -1 || idxData === -1) {
    throw new Error('В RAW_API должны быть колонки level_id и data');
  }

  const map = {};

  rows.forEach(row => {
    const levelId = String(row[idxLevelId] || '').trim();
    const raw = row[idxData];
    if (!levelId || !raw) return;

    try {
      const payload = typeof raw === 'string' ? JSON.parse(raw) : raw;
      map[levelId] = payload && payload.data ? payload.data : payload;
    } catch (e) {}
  });

  return map;
}

var QA_BASE_LABEL_ = 'РУ';
var QA_CMP_LABEL_ = 'ЛОК';

function localeLabelForQa_(code) {
  const s = String(code || '').trim().toUpperCase();
  const map = { RU: 'РУ', EN: 'EN', AZ: 'AZ', ID: 'ID', ES: 'ES', PL: 'PL', TT: 'TT', EL: 'EL', HE: 'HE' };
  return map[s] || s || 'ЛОК';
}

function setQaLocaleLabelsFromPairs_(pairsData) {
  const first = (pairsData && pairsData[0]) || {};
  QA_BASE_LABEL_ = localeLabelForQa_(first.base_locale) || 'РУ';
  QA_CMP_LABEL_ = localeLabelForQa_(first.compare_locale) || 'ЛОК';
}
