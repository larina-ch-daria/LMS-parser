// ============================================================
// QA STYLES / MSO
// Цветовые стили ячеек QA-листов и проверка разбалловки МСО.
// ============================================================

function getQaStyles_() {
  return {
    header:    { bg: '#2d3436', fg: '#ffffff', bold: true,  size: 11 },
    taskheader:{ bg: '#0d1b2a', fg: '#e0e1dd', bold: true,  size: 12 },
    subheader: { bg: '#636e72', fg: '#ffffff', bold: true,  size: 9 },
    field:     { bg: '#ffffff', fg: '#2d3436', bold: false, size: 9 },
    critical:  { bg: '#d63031', fg: '#ffffff', bold: true,  size: 9 },
    check:     { bg: '#e17055', fg: '#ffffff', bold: false, size: 9 },
    info:      { bg: '#fdcb6e', fg: '#2d3436', bold: false, size: 9 },
    sep:       { bg: '#dfe6e9', fg: '#dfe6e9', bold: false, size: 4 },
    colheader: { bg: '#1a1a2e', fg: '#ffffff', bold: true,  size: 10 },
    tl_ok:     { bg: '#dfe6e9', fg: '#636e72', bold: false, size: 9 },
    tl_warn:   { bg: '#ffeaa7', fg: '#2d3436', bold: false, size: 9 },
    tl_bad:    { bg: '#fab1a0', fg: '#2d3436', bold: true,  size: 9 },
    img:       { bg: '#ffffff', fg: '#0984e3', bold: false, size: 9 },
    img_diff:  { bg: '#e17055', fg: '#ffffff', bold: false, size: 9 },
    embed:     { bg: '#fff3cd', fg: '#7a4b00', bold: true,  size: 9 },
    embed_diff:{ bg: '#ffb366', fg: '#2d3436', bold: true,  size: 9 },
    tl_embed:  { bg: '#fff3cd', fg: '#7a4b00', bold: false, size: 9 },
  };
}

var MSO_TEMPLATES_ = [
  { levels: 18, scores: [4,4,4,4,4, 5,5,5,5,5,5, 6,6,6,6,6, 10,10] },
  { levels: 14, scores: [4,4,4,4,4, 6,6,6,6,6, 10,10,10, 20] },
  { levels: 13, scores: [4,4,4,4,4, 6,6,6,6,6, 15,15, 20] },
  { levels: 13, scores: [5,5,5,5, 6,6,6,6,6, 10,10,10, 20] },
  { levels: 10, scores: [5,5,5,5, 10,10,10, 15,15, 20] }
];
var MSO_ALLOWED_SCORES_ = [4, 5, 6, 10, 15, 20];

function collectMsoScoringData_(lmsLevelsData, nodesByLevel, sideMainIds) {
  const scoreByLevelId = {};
  Object.keys(nodesByLevel).forEach(function(lid) {
    const root = (nodesByLevel[lid] || []).find(function(n){ return String(n.depth||'0')==='0'; })
              || (nodesByLevel[lid] || [])[0];
    if (root) scoreByLevelId[lid] = root.levelScore;
  });

  const seen = {};
  const mains = [];
  const subsByParent = {};

  lmsLevelsData.forEach(function(r) {
    const mid = clean_(r.mainLevelId);
    if (!mid || seen[mid]) return;
    if (sideMainIds && !sideMainIds[mid]) return;
    seen[mid] = true;

    const kind = clean_(r.levelKind);
    const order = clean_(r.orderInTask);
    const parent = clean_(r.parentMainLevelId);

    if (kind === 'multi') {
      if (!subsByParent[parent]) subsByParent[parent] = [];
      subsByParent[parent].push({ order: order, mainLevelId: mid });
    } else {
      mains.push({ order: order, mainLevelId: mid });
    }
  });

  return { mains: mains, subsByParent: subsByParent, scoreByLevelId: scoreByLevelId };
}

function buildMsoScoringReport_(mains, subsByParent, scoreByLevelId) {
  const lines = [];
  let worst = 'ok';

  const mainScores = mains.map(function(m) {
    return parseInt(scoreByLevelId[m.mainLevelId], 10) || 0;
  });
  const total = mainScores.reduce(function(a, b) { return a + b; }, 0);
  const nLevels = mains.length;

  if (total !== 100) {
    lines.push('🔴 Сумма баллов основных уровней = ' + total + ' (ожидается 100)');
    worst = 'bad';
  }

  const sortedScores = mainScores.slice().sort(function(a, b) { return a - b; });
  const matchedTemplate = MSO_TEMPLATES_.find(function(t) {
    return t.levels === nLevels &&
      t.scores.slice().sort(function(a, b) { return a - b; }).join(',') === sortedScores.join(',');
  });
  if (matchedTemplate) {
    lines.push('🟢 Шаблон: ' + nLevels + ' уровней, раскладка совпадает');
  } else if (total === 100) {
    lines.push('🟠 Сумма 100, но раскладка не из стандартных шаблонов — проверь');
    if (worst === 'ok') worst = 'warn';
  }

  const badNominals = mainScores.filter(function(s) {
    return MSO_ALLOWED_SCORES_.indexOf(s) === -1;
  });
  if (badNominals.length) {
    lines.push('🟠 Нетипичные баллы: ' + badNominals.join(', ') + ' (норма: ' + MSO_ALLOWED_SCORES_.join('/') + ')');
    if (worst === 'ok') worst = 'warn';
  }

  mains.forEach(function(m) {
    const mainScore = parseInt(scoreByLevelId[m.mainLevelId], 10) || 0;
    const subs = subsByParent[m.mainLevelId] || [];
    const wrong = subs.filter(function(s) {
      return (parseInt(scoreByLevelId[s.mainLevelId], 10) || 0) !== mainScore;
    });
    if (wrong.length) {
      const detail = wrong.map(function(s) {
        return s.order + '=' + (scoreByLevelId[s.mainLevelId] || '?');
      }).join(', ');
      lines.push('🔴 Ур.' + m.order + ' (=' + mainScore + '): подуровни не совпадают: ' + detail);
      worst = 'bad';
    }
  });

  const icon = worst === 'bad' ? '🔴' : worst === 'warn' ? '🟠' : '🟢';
  return { icon: icon, lines: lines, total: total, nLevels: nLevels };
}

function addMsoScoringBlock_(rows, ruMeta, lmsLevelsData, nodesByLevel, ruMainIds) {
  if (clean_(ruMeta.msoEnabled).toUpperCase() !== 'TRUE') return;

  const data = collectMsoScoringData_(lmsLevelsData, nodesByLevel, ruMainIds);
  if (!data.mains.length) return;

  const rep = buildMsoScoringReport_(data.mains, data.subsByParent, data.scoreByLevelId);

  rows.push({
    cells: ['📊 Разбалловка МСО', rep.icon, 'Уровней: ' + rep.nLevels + ', сумма: ' + rep.total, ''],
    fmt: rep.icon === '🔴' ? 'critical' : rep.icon === '🟠' ? 'check' : 'taskheader'
  });

  rep.lines.forEach(function(line) {
    rows.push({
      cells: ['  ' + line, '', '', ''],
      fmt: line.indexOf('🔴') === 0 ? 'critical' : line.indexOf('🟠') === 0 ? 'check' : 'field'
    });
  });

  rows.push({ cells: ['', '', '', ''], fmt: 'sep' });
}
