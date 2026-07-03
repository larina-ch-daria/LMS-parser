// ============================================================
// QA HTML
// Парсинг и очистка HTML из полей уровней.
// stripHtml_ — вырезает теги, сохраняя структуру текста.
// stripHtmlWithLinks_ — дополнительно извлекает <a href>.
// looksLikeHtmlCodeExample_ — определяет учебный HTML-код.
// ============================================================

function decodeHtmlEntities_(s) {
  return String(s || '')
    .replace(/&quot;/g, '"')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ');
}

function looksLikeHtmlCodeExample_(s) {
  const text = String(s || '').trim().toLowerCase();
  if (!text) return false;

  // Ссылки и картинки — это КОНТЕНТ, а не пример HTML-кода, их не засчитываем.
  const codePatterns = [
    /<html\b/, /<head\b/, /<body\b/, /<table\b/, /<tr\b/, /<td\b/, /<th\b/,
    /<div\b/, /<span\b/, /<ul\b/, /<ol\b/, /<li\b/,
    /\bclass\s*=/, /\bid\s*=/, /\bstyle\s*=/,
    /\bborder\s*=/, /\bpadding\s*:/, /\bmargin\s*:/
  ];

  let hits = 0;
  codePatterns.forEach(re => { if (re.test(text)) hits++; });
  return hits >= 2;
}

function stripHtml_(h) {
  if (!h) return '';
  const s = String(h);
  if (looksLikeHtmlCodeExample_(s)) return decodeHtmlEntities_(s).trim();

  return s
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<sub>/gi, '↓').replace(/<\/sub>/gi, '')
    .replace(/<sup>/gi, '↑').replace(/<\/sup>/gi, '')
    .replace(/<b>/gi, '**').replace(/<\/b>/gi, '**')
    .replace(/<h1[^>]*>/gi, '\n').replace(/<\/h1>/gi, '\n')
    .replace(/<li[^>]*>/gi, '• ').replace(/<\/li>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&quot;/g, '"').replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function stripAnswerTextForQa_(raw) {
  const original = String(raw || '');
  if (!original.trim()) return '';
  const stripped = stripHtml_(original);
  if (stripped && stripped.trim()) return stripped;
  return decodeHtmlEntities_(original).trim();
}

function stripImagesFromHtml_(html) {
  return String(html || '')
    .replace(/<img[^>]*>/gi, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function stripHtmlWithLinks_(html) {
  if (!html) return { text: '', links: [] };

  const s = String(html);
  const links = [];
  let plain = '';
  const aTagRe = /<a\s[^>]*href\s*=\s*["']([^"']*)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let last = 0;

  let match;
  while ((match = aTagRe.exec(s)) !== null) {
    const before = stripHtml_(s.slice(last, match.index));
    plain += before;

    const anchorClean = stripHtml_(match[2]);
    const url = match[1];

    if (anchorClean && url) {
      const needSpaceBefore = plain.length > 0
        && plain[plain.length - 1] !== ' '
        && plain[plain.length - 1] !== '\n';
      if (needSpaceBefore) plain += ' ';
      links.push({ start: plain.length, end: plain.length + anchorClean.length, url });
      plain += anchorClean + ' ';
    } else {
      plain += anchorClean;
    }

    last = match.index + match[0].length;
  }

  plain += stripHtml_(s.slice(last));
  plain = plain.replace(/\n{3,}/g, '\n\n').trim();

  return { text: plain, links };
}
