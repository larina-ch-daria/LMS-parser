// ============================================================
// QA IMAGES
// Парсинг, нормализация и дедупликация изображений из JSON
// и HTML-полей уровней. Извлечение изображений из колонок
// matching-заданий и из ответов с картинками.
// ============================================================

function getFileNameFromUrl_(url) {
  const s = String(url || '').trim();
  if (!s) return '';
  const cleanUrl = s.split('?')[0].split('#')[0];
  const parts = cleanUrl.split('/');
  return parts.length ? parts[parts.length - 1] : '';
}

function parseImages_(json) {
  if (!json || json === '[]') return [];
  try {
    return JSON.parse(json).map(function (img) {
      return {
        name: clean_(img && img.name) || '?',
        url: normalizeImageUrl_(img && img.url || ''),
        source: clean_(img && img.source)
      };
    }).filter(function (img) { return !!img.url; });
  } catch (e) {
    return [];
  }
}

function dedupeImagesForQa_(items) {
  const out = [];
  const seen = {};
  (items || []).forEach(function (item) {
    const url = normalizeImageUrl_(item && item.url);
    if (!url || seen[url]) return;
    seen[url] = true;
    out.push({
      name: clean_(item && item.name) || getFileNameFromUrl_(url) || 'image',
      url: url,
      source: clean_(item && item.source)
    });
  });
  return out;
}

function extractImageUrlsFromHtml_(html, source) {
  const out = [];
  const seen = {};
  const s = String(html || '');
  if (!s) return out;
  const re = /<img[^>]+src\s*=\s*["']([^"']+)["'][^>]*>/gi;
  let m;
  while ((m = re.exec(s)) !== null) {
    let url = clean_(m[1]);
    if (!url) continue;
    url = normalizeImageUrl_(url);
    if (!url || seen[url]) continue;
    seen[url] = true;
    out.push({ name: getFileNameFromUrl_(url) || 'image', url, source: source || '' });
  }
  return out;
}

function collectNodeImagesForQa_(node) {
  const out = [];
  const seen = {};

  function pushItems(items) {
    (items || []).forEach(function (item) {
      const url = normalizeImageUrl_(item && item.url);
      if (!url || seen[url]) return;
      seen[url] = true;
      out.push({
        name: clean_(item && item.name) || getFileNameFromUrl_(url) || 'image',
        url: url,
        source: clean_(item && item.source)
      });
    });
  }

  if (!node) return out;
  pushItems(parseImages_(node.images_json));
  pushItems(extractImageUrlsFromHtml_(node.note, 'config'));
  pushItems(extractImageUrlsFromHtml_(node.description, 'config'));
  pushItems(extractImageUrlsFromHtml_(node.text, 'config'));
  pushItems(extractImageUrlsFromHtml_(node.content, 'config'));
  return out;
}

function extractMatchingColumnImagesForQa_(cfgJson) {
  const out = [];
  const seen = {};
  const cfg = safeParse_(cfgJson);
  const cols = getPath_(cfg, 'mechanic.columns');
  if (!Array.isArray(cols) || !cols.length) return out;

  cols.forEach(function (col, colIdx) {
    const type = clean_(col && col.type);
    if (type !== 'image') return;
    const items = Array.isArray(col && col.items) ? col.items : [];
    items.forEach(function (item, itemIdx) {
      const content = item && item.content;
      if (!content || typeof content !== 'object') return;
      const url = normalizeImageUrl_(content.url);
      if (!url || seen[url]) return;
      seen[url] = true;
      out.push({
        name: clean_(content.name) || ('matching_col_' + (colIdx + 1) + '_' + (itemIdx + 1)),
        url,
        source: 'config.columns.image'
      });
    });
  });
  return out;
}

function normalizeImageSourceLabelForQa_(source) {
  const s = clean_(source).toLowerCase();
  if (!s) return '';
  if (s === 'comics') return 'comics';
  if (s.indexOf('answers') >= 0) return 'config';
  if (s.indexOf('problems') >= 0) return 'config';
  if (
    s.indexOf('config') === 0 ||
    s.indexOf('mechanic') === 0 ||
    s.indexOf('problem') === 0
  ) return 'config';
  return source;
}

function formatImageLabelForQa_(img) {
  if (!img) return '—';
  const name = clean_(img.name) || getFileNameFromUrl_(img.url) || 'image';
  const source = clean_(img.source);
  if (!source) return name;
  return name + ' [' + normalizeImageSourceLabelForQa_(source) + ']';
}

function buildImageRowLabelForQa_(baseLabel, ruImgs, azImgs) {
  const sources = {};
  (ruImgs || []).forEach(function (img) {
    const s = normalizeImageSourceLabelForQa_(img && img.source);
    if (s) sources[s] = true;
  });
  (azImgs || []).forEach(function (img) {
    const s = normalizeImageSourceLabelForQa_(img && img.source);
    if (s) sources[s] = true;
  });
  const sourceList = Object.keys(sources);
  if (!sourceList.length) return baseLabel;
  return baseLabel + ' [' + sourceList.join(', ') + ']';
}
