// ==UserScript==
// @name         LMS export levels
// @namespace    lms-export-levels
// @version      1.16
// @match        *://lms.alg.academy/lesson/view/*
// @match        *://lms.alg.academy/course/view/*
// @match        *://lms.algoritmika.az/lesson/view/*
// @match        *://lms.algoritmika.az/course/view/*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  // ─── WEBHOOKS ────────────────────────────────────────────────────────────────
  // Каждый участник → свой Apps Script Web App (личный вебхук) и своя Google-таблица.
  // Значения ниже — плейсхолдеры. Подставь реальные URL в своей копии скрипта,
  // в репозиторий их коммитить не нужно.
  //   url       — деплой вебхука:  https://script.google.com/macros/s/<DEPLOY_ID>/exec
  //   sheetUrl  — личная таблица:  https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit
  const WEBHOOK_GROUPS = [
    {
      label: 'Тестеры',
      members: [
        { name: 'Тестер 1', url: 'https://script.google.com/macros/s/YOUR_WEBHOOK_DEPLOY_ID/exec', sheetUrl: 'https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit' },
        { name: 'Тестер 2', url: 'https://script.google.com/macros/s/YOUR_WEBHOOK_DEPLOY_ID/exec', sheetUrl: 'https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit' },
      ]
    },
    {
      label: 'Локализаторы',
      members: [
        { name: 'Локализатор 1', url: 'https://script.google.com/macros/s/YOUR_WEBHOOK_DEPLOY_ID/exec', sheetUrl: 'https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit' },
        { name: 'Локализатор 2', url: 'https://script.google.com/macros/s/YOUR_WEBHOOK_DEPLOY_ID/exec', sheetUrl: 'https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit' },
      ]
    }
  ];

  const WEBHOOKS = Object.fromEntries(
    WEBHOOK_GROUPS.flatMap(g => g.members.map(m => [m.name, m.url]))
  );

  const SHEET_URLS = Object.fromEntries(
    WEBHOOK_GROUPS.flatMap(g => g.members.map(m => [m.name, m.sheetUrl || '']))
  );

  function getSheetUrl() {
    return SHEET_URLS[getSelectedWebhookName()] || '';
  }

  const WEBHOOK_STORAGE_KEY = 'LMS_SELECTED_WEBHOOK_NAME';
  const courseLanguageCache = {};

  // Единый веб-апп проверки перелинковки (общий для всех, не персональный вебхук).
  const RELINK_URL = 'https://script.google.com/macros/s/YOUR_RELINK_WEBAPP_ID/exec';

  // ─── УТИЛИТЫ ─────────────────────────────────────────────────────────────────
  function cleanText(el) {
    return (el?.textContent || '').replace(/\s+/g, ' ').trim();
  }

  function normalizeText(text) {
    return String(text || '').replace(/\s+/g, ' ').trim();
  }

  function getSelectedWebhookName() {
    return localStorage.getItem(WEBHOOK_STORAGE_KEY) || '';
  }

  function setSelectedWebhookName(name) {
    localStorage.setItem(WEBHOOK_STORAGE_KEY, name);
  }

  function getWebAppUrl() {
    const selectedName = getSelectedWebhookName();
    return WEBHOOKS[selectedName] || '';
  }

  function updateWebhookBadge() {
    const badge = document.getElementById('tm-webhook-badge');
    if (!badge) return;
    const selectedName = getSelectedWebhookName();
    badge.textContent = selectedName ? `Webhook: ${selectedName}` : 'Webhook: не выбран';
    const wbtn = document.getElementById('tm-webhook-btn');
    if (wbtn) wbtn.style.display = selectedName ? 'none' : '';
    updateSheetBtn();
  }

  function updateSheetBtn() {
    const btn = document.getElementById('tm-sheet-btn');
    if (!btn) return;
    const url = getSheetUrl();
    if (url) {
      btn.style.display = '';
      btn.onclick = () => window.open(url, '_blank');
    } else {
      btn.style.display = 'none';
    }
  }

  // ─── ТОСТЫ ───────────────────────────────────────────────────────────────────
  function ensureToastContainer() {
    let c = document.getElementById('tm-toast-container');
    if (c) return c;
    c = document.createElement('div');
    c.id = 'tm-toast-container';
    Object.assign(c.style, {
      position: 'fixed', top: '60px', left: '50%', transform: 'translateX(-50%)',
      zIndex: '2147483647', display: 'flex', flexDirection: 'column', gap: '8px',
      alignItems: 'center', pointerEvents: 'none', maxWidth: 'min(360px, calc(100vw - 32px))'
    });
    document.body.appendChild(c);
    return c;
  }

  function showToast(message, type = 'success', ms = 3500) {
    const colors = {
      success: { bar: '#2e7d32', icon: '✓' },
      error:   { bar: '#c62828', icon: '✕' },
      info:    { bar: '#1565c0', icon: 'ℹ' }
    };
    const c = colors[type] || colors.info;
    const container = ensureToastContainer();

    const toast = document.createElement('div');
    Object.assign(toast.style, {
      pointerEvents: 'auto', cursor: 'pointer',
      background: '#1e1e1e', color: '#f0f0f0',
      borderLeft: `3px solid ${c.bar}`, borderRadius: '8px',
      padding: '10px 14px', boxShadow: '0 8px 28px rgba(0,0,0,0.4)',
      fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
      fontSize: '13px', lineHeight: '1.4', whiteSpace: 'pre-line',
      display: 'flex', gap: '9px', alignItems: 'flex-start',
      opacity: '0', transform: 'translateY(-8px)',
      transition: 'opacity 0.18s ease, transform 0.18s ease'
    });

    const ic = document.createElement('div');
    ic.textContent = c.icon;
    Object.assign(ic.style, { color: c.bar, fontWeight: '700', flex: '0 0 auto', marginTop: '1px' });

    const txt = document.createElement('div');
    txt.textContent = message;

    toast.appendChild(ic);
    toast.appendChild(txt);
    container.appendChild(toast);

    requestAnimationFrame(() => { toast.style.opacity = '1'; toast.style.transform = 'translateY(0)'; });

    let timer = null;
    function dismiss() {
      if (timer) clearTimeout(timer);
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(-8px)';
      setTimeout(() => { try { toast.remove(); } catch (e) {} }, 200);
    }
    toast.addEventListener('click', dismiss);
    if (ms > 0) timer = setTimeout(dismiss, ms);
    return dismiss;
  }

  // ─── МОДАЛЬНОЕ ОКНО ВЫБОРА ВЕБХУКА ───────────────────────────────────────────
  function chooseWebhook() {
    return new Promise((resolve) => {
      document.getElementById('tm-webhook-modal-overlay')?.remove();

      const current = getSelectedWebhookName();

      const styleEl = document.createElement('style');
      styleEl.id = 'tm-webhook-modal-style';
      styleEl.textContent = `
        #tm-webhook-modal-overlay {
          position: fixed; inset: 0; z-index: 2147483646;
          background: rgba(0,0,0,0.55);
          display: flex; align-items: center; justify-content: center;
          backdrop-filter: blur(3px);
          animation: tmFadeIn 0.15s ease;
        }
        @keyframes tmFadeIn { from { opacity: 0 } to { opacity: 1 } }
        @keyframes tmSlideUp { from { opacity: 0; transform: translateY(12px) } to { opacity: 1; transform: translateY(0) } }

        #tm-webhook-modal {
          background: #1e1e1e;
          border: 1px solid #333;
          border-radius: 12px;
          padding: 20px;
          width: 360px;
          max-width: calc(100vw - 32px);
          box-shadow: 0 20px 60px rgba(0,0,0,0.6);
          animation: tmSlideUp 0.18s ease;
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        }

        #tm-webhook-modal h2 {
          margin: 0 0 4px;
          font-size: 15px;
          font-weight: 600;
          color: #f0f0f0;
        }

        #tm-webhook-modal .tm-modal-subtitle {
          font-size: 12px;
          color: #888;
          margin: 0 0 14px;
        }

        #tm-webhook-modal .tm-group-label {
          font-size: 10px;
          font-weight: 700;
          letter-spacing: 0.08em;
          text-transform: uppercase;
          color: #666;
          margin: 12px 0 6px;
          padding: 0 2px;
        }

        #tm-webhook-modal .tm-group-label:first-of-type {
          margin-top: 0;
        }

        #tm-webhook-modal .tm-members {
          display: flex;
          flex-wrap: wrap;
          gap: 6px;
        }

        #tm-webhook-modal .tm-member-btn {
          padding: 6px 12px;
          border-radius: 20px;
          border: 1.5px solid #3a3a3a;
          background: #2a2a2a;
          color: #d0d0d0;
          font-size: 13px;
          cursor: pointer;
          transition: border-color 0.12s, background 0.12s, color 0.12s;
          line-height: 1.2;
          white-space: nowrap;
        }

        #tm-webhook-modal .tm-member-btn:hover {
          border-color: #555;
          background: #333;
          color: #fff;
        }

        #tm-webhook-modal .tm-member-btn.tm-active {
          border-color: #4a9eff;
          background: #1a3a5c;
          color: #7ec8ff;
        }

        #tm-webhook-modal .tm-divider {
          height: 1px;
          background: #2a2a2a;
          margin: 14px 0;
        }

        #tm-webhook-modal .tm-footer {
          display: flex;
          justify-content: flex-end;
          gap: 8px;
          margin-top: 16px;
        }

        #tm-webhook-modal .tm-btn-cancel {
          padding: 7px 14px;
          border-radius: 7px;
          border: 1px solid #3a3a3a;
          background: transparent;
          color: #888;
          font-size: 13px;
          cursor: pointer;
          transition: color 0.12s, border-color 0.12s;
        }
        #tm-webhook-modal .tm-btn-cancel:hover { color: #ccc; border-color: #555; }

        #tm-webhook-modal .tm-btn-confirm {
          padding: 7px 16px;
          border-radius: 7px;
          border: none;
          background: #2e7d32;
          color: #fff;
          font-size: 13px;
          font-weight: 600;
          cursor: pointer;
          transition: background 0.12s;
        }
        #tm-webhook-modal .tm-btn-confirm:hover { background: #388e3c; }
        #tm-webhook-modal .tm-btn-confirm:disabled {
          background: #2a2a2a;
          color: #555;
          cursor: default;
        }
      `;
      document.head.appendChild(styleEl);

      const overlay = document.createElement('div');
      overlay.id = 'tm-webhook-modal-overlay';

      const modal = document.createElement('div');
      modal.id = 'tm-webhook-modal';

      const title = document.createElement('h2');
      title.textContent = 'Выбери webhook';

      const subtitle = document.createElement('div');
      subtitle.className = 'tm-modal-subtitle';
      subtitle.textContent = current ? `Сейчас выбран: ${current}` : 'Webhook ещё не выбран';

      modal.appendChild(title);
      modal.appendChild(subtitle);

      let pendingName = current;

      WEBHOOK_GROUPS.forEach((group, gi) => {
        if (gi > 0) {
          const divider = document.createElement('div');
          divider.className = 'tm-divider';
          modal.appendChild(divider);
        }

        const groupLabel = document.createElement('div');
        groupLabel.className = 'tm-group-label';
        groupLabel.textContent = group.label;
        modal.appendChild(groupLabel);

        const membersRow = document.createElement('div');
        membersRow.className = 'tm-members';

        group.members.forEach(member => {
          const btn = document.createElement('button');
          btn.type = 'button';
          btn.className = 'tm-member-btn' + (member.name === current ? ' tm-active' : '');
          btn.textContent = member.name;
          btn.dataset.name = member.name;

          btn.addEventListener('click', () => {
            modal.querySelectorAll('.tm-member-btn').forEach(b => b.classList.remove('tm-active'));
            btn.classList.add('tm-active');
            pendingName = member.name;
            confirmBtn.disabled = false;
          });

          membersRow.appendChild(btn);
        });

        modal.appendChild(membersRow);
      });

      const footer = document.createElement('div');
      footer.className = 'tm-footer';

      const cancelBtn = document.createElement('button');
      cancelBtn.type = 'button';
      cancelBtn.className = 'tm-btn-cancel';
      cancelBtn.textContent = 'Отмена';

      const confirmBtn = document.createElement('button');
      confirmBtn.type = 'button';
      confirmBtn.className = 'tm-btn-confirm';
      confirmBtn.textContent = 'Выбрать';
      confirmBtn.disabled = !current;

      footer.appendChild(cancelBtn);
      footer.appendChild(confirmBtn);
      modal.appendChild(footer);

      overlay.appendChild(modal);
      document.body.appendChild(overlay);

      function close(confirmed) {
        overlay.remove();
        styleEl.remove();

        if (confirmed && pendingName) {
          setSelectedWebhookName(pendingName);
          updateWebhookBadge();
          resolve(WEBHOOKS[pendingName]);
        } else {
          resolve(getWebAppUrl());
        }
      }

      confirmBtn.addEventListener('click', () => close(true));
      cancelBtn.addEventListener('click', () => close(false));
      overlay.addEventListener('click', (e) => { if (e.target === overlay) close(false); });
      document.addEventListener('keydown', function onKey(e) {
        if (e.key === 'Escape') { close(false); document.removeEventListener('keydown', onKey); }
        if (e.key === 'Enter' && !confirmBtn.disabled) { close(true); document.removeEventListener('keydown', onKey); }
      });
    });
  }

  async function ensureWebAppUrl() {
    let url = getWebAppUrl();
    if (!url) {
      url = await chooseWebhook();
    }
    if (!url) {
      throw new Error('Не выбран WEB_APP_URL');
    }
    return url;
  }

  // ─── DOM-ПАРСЕРЫ ─────────────────────────────────────────────────────────────
  function extractLevelUuidFromHref(href) {
    const match = (href || '').match(/\/level\/update\/([^/?#]+)/);
    return match ? match[1] : '';
  }

  function extractLevelPreviewId(li) {
    const previewLink = li.querySelector('a[href*="/level-preview/"]');
    const href = previewLink?.getAttribute('href') || '';
    const match = href.match(/\/level-preview\/(\d+)/);
    return match ? match[1] : '';
  }

  function extractCourseUuidFromHref(href) {
    const match = (href || '').match(/\/course\/view\/([^/?#]+)/);
    return match ? match[1] : '';
  }

  function getLessonGuidFromEditor(lessonEditor) {
    if (!lessonEditor) return '';
    const copyLinks = [...lessonEditor.querySelectorAll('.note a.copyTextToClipboard')];
    for (const link of copyLinks) {
      const text = cleanText(link);
      if (/^[0-9a-f-]{20,}$/i.test(text)) return text;
    }
    return '';
  }

  function getLessonNote() {
    const lessonEditor = document.querySelector('#lesson-editor');
    if (!lessonEditor) return '';
    const noteEls = [...lessonEditor.querySelectorAll('.box-header .note, .note')];
    for (const el of noteEls) {
      if (el.querySelector('.copyTextToClipboard')) continue;
      const text = normalizeText(el.textContent || '');
      if (!text) continue;
      return text;
    }
    return '';
  }

  function readDropdownStatus(dropdownEl) {
    if (!dropdownEl) return { code: '', label: '' };
    const label = cleanText(dropdownEl.querySelector('.btn-text'));
    const selected = dropdownEl.querySelector('ul.dropdown-menu a.selected, ul.dropdown-menu li a.selected');
    const code = selected ? (selected.getAttribute('data-status') || '') : '';
    return { code, label };
  }

  function getValueParagraphAfterLabel(labelEl) {
    if (!labelEl || !labelEl.parentElement) return null;
    const siblings = [...labelEl.parentElement.children];
    const startIdx = siblings.indexOf(labelEl);
    if (startIdx < 0) return null;
    for (let i = startIdx + 1; i < siblings.length; i++) {
      const el = siblings[i];
      if (el.tagName === 'LABEL') return null;
      if (el.tagName === 'P') return el;
    }
    return null;
  }

  function isEmptyValueParagraph(pEl) {
    if (!pEl) return true;
    if (pEl.querySelector('a[href]')) return false;
    const text = normalizeText(pEl.textContent || '');
    return !text;
  }

  function findEditableFields(container) {
    if (!container) return [];
    const labels = [...container.querySelectorAll('label.control-label')];
    const result = [];
    for (const label of labels) {
      const valueP = getValueParagraphAfterLabel(label);
      if (valueP) result.push({ label, valueP });
    }
    return result;
  }

  function findLessonPanels() {
    const all = [...document.querySelectorAll('.panel .panel-body')];
    return all.filter(pb => {
      if (pb.querySelector('ol.task-levels')) return false;
      if (pb.closest('li[data-task]')) return false;
      if (pb.closest('.row.js-task-level')) return false;
      return true;
    });
  }

  function findTextByLabel(root, label) {
    const text = root?.innerText || '';
    const re = new RegExp(label + '\\s*:\\s*([^\\n]+)');
    const match = text.match(re);
    return match ? match[1].trim() : '';
  }

  function getPublicNameMeta() {
    const panels = findLessonPanels();
    let publicName = '';

    for (const pb of panels) {
      if (pb.querySelector('media-uploader, [media-type], .uploadmanager')) continue;
      if (pb.querySelector('ol.task-levels')) continue;

      const fields = findEditableFields(pb);
      if (fields.length === 0) continue;

      const first = fields[0];
      if (fields.length < 2) continue;

      publicName = isEmptyValueParagraph(first.valueP)
        ? ''
        : normalizeText(first.valueP.textContent || '');
      break;
    }

    const hasPublicName = publicName ? 'TRUE' : 'FALSE';
    return { publicName, hasPublicName };
  }

  function getMsoStatus(lessonId) {
    const dd = document.getElementById(`mso-dropdown-${lessonId}`);
    const { code, label } = readDropdownStatus(dd);
    let enabled = '';
    if (code === '1') enabled = 'TRUE';
    else if (code === '0') enabled = 'FALSE';
    return { msoStatus: label, msoEnabled: enabled };
  }

  function getCourseMeta() {
    const h4s = [...document.querySelectorAll('h4')];
    let container = null;
    for (const h4 of h4s) {
      const parent = h4.parentElement;
      if (!parent) continue;
      const links = parent.querySelectorAll('a[href*="/course/view/"]');
      if (links.length > 0) {
        container = parent;
        break;
      }
    }
    if (!container) return [];
    const courseLinks = [...container.querySelectorAll('a[href*="/course/view/"]')];
    return courseLinks.map(link => {
      const courseUuid = extractCourseUuidFromHref(link.getAttribute('href') || '');
      return { courseTitle: cleanText(link), courseUrl: buildCourseUrl(courseUuid), courseUuid };
    });
  }

  function buildCourseUrl(courseUuid) {
    if (!courseUuid) return '';
    return `${location.origin}/course/view/${courseUuid}#course-general`;
  }

  function getLessonMaterials() {
    const panels = findLessonPanels();
    for (const pb of panels) {
      if (pb.querySelector('media-uploader, [media-type], .uploadmanager')) continue;
      if (pb.querySelector('ol.task-levels')) continue;

      const fields = findEditableFields(pb);
      if (fields.length < 2) continue;

      const second = fields[1];
      const links = [...second.valueP.querySelectorAll('a[href]')];
      if (!links.length) return '';
      return links
        .map(a => {
          const title = normalizeText(a.textContent) || a.getAttribute('href');
          const url = a.getAttribute('href') || '';
          return `${title}\n${url}`;
        })
        .join('\n');
    }
    return '';
  }

  function getLessonVideoUrl() {
    const panels = findLessonPanels();
    for (const pb of panels) {
      if (pb.querySelector('media-uploader, [media-type], .uploadmanager')) continue;
      if (pb.querySelector('ol.task-levels')) continue;

      const fields = findEditableFields(pb);
      if (fields.length !== 1) continue;

      const { valueP } = fields[0];
      const a = valueP.querySelector('a[href]');
      if (a) return a.getAttribute('href') || '';

      const iframe = valueP.querySelector('iframe[src]');
      if (iframe) {
        const src = iframe.getAttribute('src') || '';
        const ytMatch = src.match(/youtube\.com\/embed\/([^?&]+)/);
        if (ytMatch) return `https://www.youtube.com/watch?v=${ytMatch[1]}`;
        return src;
      }

      const text = normalizeText(valueP.textContent);
      if (text && text.length > 5) return text;
      return '';
    }
    return '';
  }

  function getLessonMeta() {
    const lessonEditor = document.querySelector('#lesson-editor');
    const lessonId = lessonEditor?.getAttribute('data-lesson') || '';
    const lessonTitle =
      cleanText(document.querySelector('#lesson-editor .box-title')) ||
      cleanText(document.querySelector('h3.box-title')) ||
      document.title.trim() || '';
    const lessonGuid = getLessonGuidFromEditor(lessonEditor);

    const publishDd = document.getElementById(`publish-dropdown-${lessonId}`);
    const lessonStatus = readDropdownStatus(publishDd).label || '';

    const lessonNote = getLessonNote();
    const courses = getCourseMeta();
    const msoMeta = getMsoStatus(lessonId);
    const publicMeta = getPublicNameMeta();
    const lessonMaterials = getLessonMaterials();
    const lessonVideoUrl = getLessonVideoUrl();
    return {
      lessonId, lessonTitle, lessonGuid, lessonStatus, lessonNote,
      pageTitle: document.title.trim(), pageUrl: location.href, courses,
      msoStatus: msoMeta.msoStatus, msoEnabled: msoMeta.msoEnabled,
      publicName: publicMeta.publicName, hasPublicName: publicMeta.hasPublicName,
      lessonMaterials, lessonVideoUrl
    };
  }

  function detectLocaleByLanguageName(languageName) {
    const map = {
      'English': 'en-US', 'Русский': 'ru-RU', 'Polski': 'pl-PL',
      'עִבְרִית (Иврит)': 'he-HE', 'Azərbaycan': 'az-AZ',
      'Tatar': 'tt-RU', 'ελληνικά (Греческий)': 'el-GR'
    };
    return map[languageName] || '';
  }

  function buildCoursePageUrl(courseUuid) {
    if (!courseUuid) return '';
    return `${location.origin}/course/view/${courseUuid}`;
  }

  async function getLessonPositionInCourse(courseUuid, lessonGuid) {
    if (!courseUuid || !lessonGuid) return { position: '', total: '' };
    const url = buildCoursePageUrl(courseUuid);
    try {
      const resp = await fetch(url, { credentials: 'same-origin' });
      const html = await resp.text();
      const doc = new DOMParser().parseFromString(html, 'text/html');
      const lessons = [...doc.querySelectorAll('#course-lessons-container > li[data-item-type="lesson"]')];
      let position = '';
      const total = lessons.length;
      for (let i = 0; i < lessons.length; i++) {
        const link = lessons[i].querySelector('a[data-qa-id="lesson-title"]');
        const href = link?.getAttribute('href') || '';
        if (href.includes(lessonGuid)) { position = i + 1; break; }
      }
      return { position: position || '', total };
    } catch (err) {
      console.error('[LMS EXPORT] getLessonPositionInCourse error:', err);
      return { position: '', total: '' };
    }
  }

  function getCourseLanguage(courseUrl) {
    if (!courseUrl) return Promise.resolve({ courseLanguage: '', courseLocale: '' });
    const url = courseUrl.includes('#') ? courseUrl : `${courseUrl}#course-general`;
    if (courseLanguageCache[url]) return Promise.resolve(courseLanguageCache[url]);

    return new Promise((resolve) => {
      const iframe = document.createElement('iframe');
      Object.assign(iframe.style, {
        position: 'fixed', left: '-99999px', top: '0',
        width: '1200px', height: '800px', opacity: '0', pointerEvents: 'none'
      });
      iframe.setAttribute('aria-hidden', 'true');
      iframe.src = url;
      let finished = false;

      function done(result) {
        if (finished) return;
        finished = true;
        courseLanguageCache[url] = result;
        try { iframe.remove(); } catch (e) {}
        resolve(result);
      }

      const timeoutId = setTimeout(() => done({ courseLanguage: '', courseLocale: '' }), 12000);

      iframe.onload = () => {
        const startedAt = Date.now();
        function tryReadLanguage() {
          try {
            const doc = iframe.contentDocument || iframe.contentWindow?.document;
            if (!doc) {
              if (Date.now() - startedAt < 10000) setTimeout(tryReadLanguage, 400);
              else { clearTimeout(timeoutId); done({ courseLanguage: '', courseLocale: '' }); }
              return;
            }
            const languageSelect =
              doc.querySelector('#course-language_id') ||
              doc.querySelector('select[name="Course[language_id]"]');
            if (!languageSelect) {
              if (Date.now() - startedAt < 10000) setTimeout(tryReadLanguage, 400);
              else { clearTimeout(timeoutId); done({ courseLanguage: '', courseLocale: '' }); }
              return;
            }
            const selectedOption =
              languageSelect.querySelector('option[selected]') ||
              languageSelect.options[languageSelect.selectedIndex];
            const courseLanguage = normalizeText(selectedOption?.textContent || '');
            const selectedValue = String(selectedOption?.value || '').trim();
            let courseLocale = '';
            if (selectedValue === '1') courseLocale = 'en-US';
            else if (selectedValue === '2') courseLocale = 'ru-RU';
            else if (selectedValue === '3') courseLocale = 'pl-PL';
            else if (selectedValue === '4') courseLocale = 'he-HE';
            else if (selectedValue === '5') courseLocale = 'az-AZ';
            else if (selectedValue === '12') courseLocale = 'tt-RU';
            else if (selectedValue === '13') courseLocale = 'el-GR';
            else courseLocale = detectLocaleByLanguageName(courseLanguage);
            clearTimeout(timeoutId);
            done({ courseLanguage, courseLocale });
          } catch (err) {
            if (Date.now() - startedAt < 10000) setTimeout(tryReadLanguage, 400);
            else { clearTimeout(timeoutId); done({ courseLanguage: '', courseLocale: '' }); }
          }
        }
        setTimeout(tryReadLanguage, 1200);
      };

      document.body.appendChild(iframe);
    });
  }

  // ─── СБОР ДАННЫХ ─────────────────────────────────────────────────────────────
  async function collectRows() {
    const meta = getLessonMeta();
    let coursesEnriched = [];

    if (meta.courses.length > 0) {
      coursesEnriched = await Promise.all(meta.courses.map(async (c) => {
        const [langMeta, posMeta] = await Promise.all([
          getCourseLanguage(c.courseUrl),
          getLessonPositionInCourse(c.courseUuid, meta.lessonGuid)
        ]);
        return { ...c, courseLanguage: langMeta.courseLanguage, courseLocale: langMeta.courseLocale,
          lessonPositionInCourse: posMeta.position, courseLessonsTotal: posMeta.total };
      }));
    } else {
      coursesEnriched = [{
        courseTitle: '', courseUrl: '', courseUuid: '',
        courseLanguage: '', courseLocale: '',
        lessonPositionInCourse: '', courseLessonsTotal: ''
      }];
    }

    const result = [];
    const taskRows = [...document.querySelectorAll('.row.js-task-level')];

    for (const row of taskRows) {
      const rowLessonId = row.getAttribute('data-lesson-id') || meta.lessonId || '';
      const taskLi = row.closest('li[data-task]');

      const bonusCheckbox = taskLi ? taskLi.querySelector('.checkbox-lesson-task-type[name="lessonTaskType"]') : null;
      const isBonus = bonusCheckbox && bonusCheckbox.checked ? 'TRUE' : 'FALSE';

      const theoryCheckbox = taskLi ? taskLi.querySelector('.checkbox-task-display-type[name="taskDisplayType"][value="theory"]') : null;
      const isTheory = theoryCheckbox && theoryCheckbox.checked ? 'TRUE' : 'FALSE';

      const quizCheckbox = taskLi ? taskLi.querySelector('.toggle-kahoot[name="toggleKahoot"]') : null;
      const isQuiz = quizCheckbox && quizCheckbox.checked ? 'TRUE' : 'FALSE';

      const trackCols = [...row.querySelectorAll('.track')];
      for (const col of trackCols) {
        const box = col.querySelector('.box');
        if (!box) continue;
        const trackButton = box.querySelector('.box-header a.btn.btn-sm.btn-primary');
        const taskTitle = cleanText(trackButton);
        const list = box.querySelector('ol.task-levels.sortable');
        if (!list) continue;
        const taskId = list.getAttribute('data-task') || '';
        const track = list.getAttribute('data-track') || '';
        const levels = [...list.querySelectorAll(':scope > li.task-level')];

        levels.forEach((li, index) => {
          const rawTaskLevel = li.getAttribute('data-level') || '';
          const taskLevelId = rawTaskLevel.replace(/^tasklevel_/, '');
          const mainLevelId = li.getAttribute('data-main-level-id') || '';
          const mainLevelLink = li.querySelector(':scope > a[href*="/level/update/"]');
          const levelTitle = cleanText(mainLevelLink);
          const levelHref = mainLevelLink?.getAttribute('href') || '';
          const levelUuid = extractLevelUuidFromHref(levelHref);

          for (const course of coursesEnriched) {
            const baseRow = {
              lessonId: rowLessonId, lessonTitle: meta.lessonTitle,
              lessonGuid: meta.lessonGuid, lessonStatus: meta.lessonStatus, lessonNote: meta.lessonNote,
              lessonMaterials: meta.lessonMaterials, lessonVideoUrl: meta.lessonVideoUrl,
              courseTitle: course.courseTitle, courseUrl: course.courseUrl, courseUuid: course.courseUuid,
              courseLanguage: course.courseLanguage, courseLocale: course.courseLocale,
              lessonPositionInCourse: course.lessonPositionInCourse, courseLessonsTotal: course.courseLessonsTotal,
              msoStatus: meta.msoStatus, msoEnabled: meta.msoEnabled,
              publicName: meta.publicName, hasPublicName: meta.hasPublicName,
              pageTitle: meta.pageTitle, pageUrl: meta.pageUrl,
              taskId, track, taskTitle, isBonus, isTheory, isQuiz,
              levelKind: 'main', orderInTask: index + 1,
              taskLevelId, mainLevelId, multiLevelId: '',
              parentTaskLevelId: '', parentMainLevelId: '', parentLevelTitle: '',
              levelUuid, levelTitle
            };
            result.push(baseRow);

            const multiItems = [...li.querySelectorAll(':scope > .multi-level-holder ol > li')];
            multiItems.forEach((multiLi, multiIndex) => {
              const multiLink = multiLi.querySelector('a[href*="/level/update/"]');
              const multiTitle = cleanText(multiLink);
              const multiHref = multiLink?.getAttribute('href') || '';
              const multiLevelUuid = extractLevelUuidFromHref(multiHref);
              const multiLevelId = extractLevelPreviewId(multiLi);
              result.push({
                lessonId: rowLessonId, lessonTitle: meta.lessonTitle,
                lessonGuid: meta.lessonGuid, lessonStatus: meta.lessonStatus, lessonNote: meta.lessonNote,
                lessonMaterials: meta.lessonMaterials, lessonVideoUrl: meta.lessonVideoUrl,
                courseTitle: course.courseTitle, courseUrl: course.courseUrl, courseUuid: course.courseUuid,
                courseLanguage: course.courseLanguage, courseLocale: course.courseLocale,
                lessonPositionInCourse: course.lessonPositionInCourse, courseLessonsTotal: course.courseLessonsTotal,
                msoStatus: meta.msoStatus, msoEnabled: meta.msoEnabled,
                publicName: meta.publicName, hasPublicName: meta.hasPublicName,
                pageTitle: meta.pageTitle, pageUrl: meta.pageUrl,
                taskId, track, taskTitle, isBonus, isTheory, isQuiz,
                levelKind: 'multi', orderInTask: `${index + 1}.${multiIndex + 1}`,
                taskLevelId: '', mainLevelId: multiLevelId, multiLevelId: multiLevelId,
                parentTaskLevelId: taskLevelId, parentMainLevelId: mainLevelId, parentLevelTitle: levelTitle,
                levelUuid: multiLevelUuid, levelTitle: multiTitle
              });
            });
          }
        });
      }
    }

    return result;
  }

  // ─── ЭКСПОРТ / ОЧИСТКА ───────────────────────────────────────────────────────
  async function exportRows() {
    try {
      const rows = await collectRows();
      if (!rows.length) { showToast('Ничего не найдено', 'error'); return; }
      const response = await fetch(await ensureWebAppUrl(), {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'saveLessonLevels', rows })
      });
      await response.text();
      showToast(`Отправлено строк: ${rows.length}`, 'success');
    } catch (err) {
      console.error('[LMS EXPORT] error:', err);
      showToast('Ошибка: ' + err.message, 'error', 6000);
    }
  }

  async function clearRows() {
    try {
      const ok = confirm('Очистить лист LMS_LEVELS?');
      if (!ok) return;
      const response = await fetch(await ensureWebAppUrl(), {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ action: 'clearLessonLevels' })
      });
      await response.text();
      showToast('Лист очищен', 'success');
    } catch (err) {
      console.error('[LMS CLEAR] error:', err);
      showToast('Ошибка очистки: ' + err.message, 'error', 6000);
    }
  }

  // ─── МАТЕРИАЛЫ УРОКА ─────────────────────────────────────────────────────────
  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, c =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function collectMaterialLinks() {
    const anchors = [...document.querySelectorAll('.uploadmanager a.media-list__link[href]')]
      .filter(a => !a.closest('li[data-task]'));
    const seen = new Set();
    const out = [];
    for (const a of anchors) {
      const url = a.getAttribute('href') || '';
      if (!url || seen.has(url)) continue;
      seen.add(url);
      out.push({ name: normalizeText(a.textContent) || url, url });
    }
    return out;
  }

  async function copyToClipboard(plain, html) {
    try {
      if (navigator.clipboard && window.ClipboardItem) {
        await navigator.clipboard.write([new ClipboardItem({
          'text/plain': new Blob([plain], { type: 'text/plain' }),
          'text/html':  new Blob([html],  { type: 'text/html' })
        })]);
        return true;
      }
    } catch (e) { console.warn('[LMS MAT] clipboard.write failed', e); }
    try {
      if (navigator.clipboard?.writeText) { await navigator.clipboard.writeText(plain); return true; }
    } catch (e) { console.warn('[LMS MAT] writeText failed', e); }
    try {
      const ta = document.createElement('textarea');
      ta.value = plain;
      ta.style.position = 'fixed'; ta.style.left = '-9999px';
      document.body.appendChild(ta); ta.select();
      const ok = document.execCommand('copy'); ta.remove();
      return ok;
    } catch (e) { console.warn('[LMS MAT] execCommand failed', e); return false; }
  }

  async function copyMaterials() {
    const items = collectMaterialLinks();
    if (!items.length) { showToast('Методические материалы не найдены на странице', 'error'); return; }
    const plain = items.map(i => `${i.name}\t${i.url}`).join('\n');
    const html = '<ul>' + items.map(i =>
      `<li><a href="${escapeHtml(i.url)}">${escapeHtml(i.name)}</a></li>`).join('') + '</ul>';
    const ok = await copyToClipboard(plain, html);
    if (ok) {
      showToast(`Скопировано материалов: ${items.length}`, 'success');
    } else {
      showToast('Не удалось скопировать. Ссылки выведены в консоль', 'error', 6000);
      console.log('[LMS MAT]\n' + plain);
    }
  }

  // ─── СКАЧИВАНИЕ МАТЕРИАЛОВ ────────────────────────────────────────────────────
  const EXPORT_PROFILES = {
    office: { document: 'docx', presentation: 'pptx', spreadsheets: 'xlsx',
              label: 'Office', hint: 'docx · pptx · xlsx — чистый текст для проверки' },
    pdf:    { document: 'pdf',  presentation: 'pdf',  spreadsheets: 'pdf',
              label: 'PDF',    hint: 'всё в PDF — для проверки вёрстки и картинок' }
  };

  function chooseExportProfile() {
    return new Promise((resolve) => {
      const overlay = document.createElement('div');
      Object.assign(overlay.style, {
        position: 'fixed', inset: '0', zIndex: '2147483646',
        background: 'rgba(0,0,0,0.55)', backdropFilter: 'blur(3px)',
        display: 'flex', alignItems: 'center', justifyContent: 'center'
      });

      const modal = document.createElement('div');
      Object.assign(modal.style, {
        background: '#1e1e1e', border: '1px solid #333', borderRadius: '12px',
        padding: '20px', width: '340px', maxWidth: 'calc(100vw - 32px)',
        boxShadow: '0 20px 60px rgba(0,0,0,0.6)',
        fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif", color: '#f0f0f0'
      });
      modal.innerHTML = '<h2 style="margin:0 0 14px;font-size:15px;font-weight:600;">В каком формате скачать?</h2>';

      function done(val) { overlay.remove(); resolve(val); }

      Object.entries(EXPORT_PROFILES).forEach(([key, p]) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        Object.assign(btn.style, {
          display: 'block', width: '100%', textAlign: 'left', margin: '0 0 8px',
          padding: '10px 12px', borderRadius: '8px', border: '1.5px solid #3a3a3a',
          background: '#2a2a2a', color: '#d0d0d0', cursor: 'pointer', lineHeight: '1.3'
        });
        btn.innerHTML = `<div style="font-size:13px;font-weight:600;color:#fff;">${p.label}</div>` +
                        `<div style="font-size:11px;color:#888;margin-top:2px;">${p.hint}</div>`;
        btn.onmouseenter = () => { btn.style.borderColor = '#4a9eff'; btn.style.background = '#1a3a5c'; };
        btn.onmouseleave = () => { btn.style.borderColor = '#3a3a3a'; btn.style.background = '#2a2a2a'; };
        btn.onclick = () => done(key);
        modal.appendChild(btn);
      });

      const cancel = document.createElement('button');
      cancel.type = 'button';
      cancel.textContent = 'Отмена';
      Object.assign(cancel.style, {
        marginTop: '6px', padding: '7px 14px', borderRadius: '7px',
        border: '1px solid #3a3a3a', background: 'transparent', color: '#888',
        fontSize: '13px', cursor: 'pointer', float: 'right'
      });
      cancel.onclick = () => done(null);
      modal.appendChild(cancel);

      overlay.appendChild(modal);
      overlay.addEventListener('click', (e) => { if (e.target === overlay) done(null); });
      document.addEventListener('keydown', function onKey(e) {
        if (e.key === 'Escape') { done(null); document.removeEventListener('keydown', onKey); }
      });
      document.body.appendChild(overlay);
    });
  }

  function parseGoogleDoc(url) {
    const m = (url || '').match(/\/(document|presentation|spreadsheets)\/d\/([^/?#]+)/);
    return m ? { kind: m[1], id: m[2] } : null;
  }

  function buildExportUrl(kind, id, fmt) {
    if (kind === 'presentation') return `https://docs.google.com/presentation/d/${id}/export/${fmt}`;
    return `https://docs.google.com/${kind}/d/${id}/export?format=${fmt}`;
  }

  function triggerIframeDownload(url) {
    const iframe = document.createElement('iframe');
    iframe.style.display = 'none';
    iframe.src = url;
    document.body.appendChild(iframe);
    setTimeout(() => { try { iframe.remove(); } catch (e) {} }, 60000);
  }

  async function downloadMaterials() {
    const items = collectMaterialLinks();
    if (!items.length) { showToast('Методические материалы не найдены на странице', 'error'); return; }

    const profileKey = await chooseExportProfile();
    if (!profileKey) return;
    const profile = EXPORT_PROFILES[profileKey];

    const downloadable = [];
    const skipped = [];
    for (const it of items) {
      const g = parseGoogleDoc(it.url);
      if (!g) { skipped.push(it); continue; }
      const fmt = profile[g.kind] || 'pdf';
      downloadable.push({ ...it, exportUrl: buildExportUrl(g.kind, g.id, fmt), fmt });
    }

    if (!downloadable.length) { showToast('Среди материалов нет гугл-документов для экспорта', 'error'); return; }

    const lines = downloadable.map(d => `• ${d.name} → .${d.fmt}`).join('\n');
    const skipNote = skipped.length
      ? `\n\nПропущены (не гугл-файлы): ${skipped.map(s => s.name).join(', ')}`
      : '';
    const go = confirm(
      `Скачать материалов: ${downloadable.length} (${profile.label})\n${lines}${skipNote}\n\n` +
      `Нужно быть залогиненным в гугл с доступом к файлам. ` +
      `Если браузер заблокирует множественную загрузку — разреши её для этого сайта.`
    );
    if (!go) return;

    downloadable.forEach((d, i) => setTimeout(() => triggerIframeDownload(d.exportUrl), i * 1200));
  }

  // ─── ПЕРЕЛИНКОВКА ────────────────────────────────────────────────────────────
  const ROLE_KEYWORDS = {
    presentation: ['презентац', 'təqdimat', 'tequdimat'],
    methodology:  ['методическ', 'metodik'],
    workbook:     ['задачник', 'tapşırıq kitab'],
    homework:     ['домашн', 'ev tapşır', 'ev iş'],
    razbor:       ['разбор', 'izah'],
    errors:       ['работа над ошибками', 'səhvlər']
  };
  const ROLE_LABELS = {
    presentation: 'Презентация', methodology: 'Методичка', workbook: 'Задачник',
    homework: 'Домашка', razbor: 'Разбор', errors: 'Работа над ошибками', unknown: '—'
  };

  function detectRole(name) {
    const n = (name || '').toLowerCase();
    for (const [role, kws] of Object.entries(ROLE_KEYWORDS)) {
      if (kws.some(k => n.includes(k))) return role;
    }
    return 'unknown';
  }

  function collectBundle() {
    return collectMaterialLinks().map(it => {
      const g = parseGoogleDoc(it.url);
      return {
        name: it.name, url: it.url,
        kind: g ? g.kind : '', id: g ? g.id : '',
        role: detectRole(it.name), isEmbed: it.url.includes('/embed')
      };
    });
  }

  function collectLessonLevelIds() {
    const ids = new Set();
    document.querySelectorAll('ol.task-levels.sortable > li.task-level').forEach(li => {
      const mainId = li.getAttribute('data-main-level-id');
      if (mainId) ids.add(String(mainId));
      const link = li.querySelector(':scope > a[href*="/level/update/"]');
      const uuid = extractLevelUuidFromHref(link?.getAttribute('href') || '');
      if (uuid) ids.add(uuid);
      li.querySelectorAll(':scope > .multi-level-holder ol > li').forEach(m => {
        const mu = extractLevelUuidFromHref(m.querySelector('a[href*="/level/update/"]')?.getAttribute('href') || '');
        if (mu) ids.add(mu);
        const mp = extractLevelPreviewId(m);
        if (mp) ids.add(String(mp));
      });
    });
    return [...ids];
  }

  function parseLessonContext() {
    const meta = getLessonMeta();
    const hay = `${meta.lessonTitle} ${meta.pageTitle} ${meta.publicName}`;
    const numMatch = hay.match(/(?:Урок|Dərs|Ders)\s+(\d+)\.(\d+)/i);
    const lessonNumber = numMatch ? `${numMatch[1]}.${numMatch[2]}` : '';
    let classNumber = numMatch ? numMatch[1] : '';
    let year = '';
    for (const c of meta.courses) {
      const ym = (c.courseTitle || '').match(/(\d{2})\s*[\/_]\s*(\d{2})/);
      if (ym) { year = `${ym[1]}/${ym[2]}`; break; }
      if (!classNumber) {
        const cm = (c.courseTitle || '').match(/(\d+)\s*(?:класс|sinif|sinf)/i);
        if (cm) classNumber = cm[1];
      }
    }
    if (!classNumber) {
      for (const c of meta.courses) {
        const cm = (c.courseTitle || '').match(/(\d+)\s*(?:класс|sinif|sinf)/i);
        if (cm) { classNumber = cm[1]; break; }
      }
    }
    return { lessonNumber, classNumber, year };
  }

  function checkEmbed(bundle) {
    const pres = bundle.find(b => b.role === 'presentation');
    if (!pres) return { status: 'orange', text: 'Презентация среди материалов не найдена' };
    if (pres.isEmbed) return { status: 'green', text: 'Презентация встроена через embed' };
    return { status: 'red', text: 'Презентация приложена НЕ через embed-ссылку' };
  }

  function relinkDot(status) {
    return { green: '🟢', orange: '🟠', red: '🔴', grey: '⚪' }[status] || '⚪';
  }

  function openRelinkPanel() {
    document.getElementById('tm-relink-panel')?.remove();
    const panel = document.createElement('div');
    panel.id = 'tm-relink-panel';
    Object.assign(panel.style, {
      position: 'fixed', top: '60px', left: '50%', transform: 'translateX(-50%)',
      zIndex: '2147483646', width: '480px', maxWidth: 'calc(100vw - 32px)',
      maxHeight: 'calc(100vh - 80px)', overflowY: 'auto',
      background: '#1e1e1e', border: '1px solid #333', borderRadius: '12px',
      boxShadow: '0 20px 60px rgba(0,0,0,0.6)', padding: '16px 18px',
      fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
      color: '#f0f0f0', fontSize: '13px', lineHeight: '1.45'
    });
    document.body.appendChild(panel);
    return panel;
  }

  function renderRelink(panel, s) {
    const esc = escapeHtml;
    const section = t => `<div style="font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#777;margin:14px 0 6px;">${esc(t)}</div>`;
    const line = html => `<div style="margin-bottom:3px;">${html}</div>`;
    const dotLine = (st, t) => `<div style="margin-bottom:5px;">${relinkDot(st)} ${esc(t)}</div>`;

    let h = '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">'
      + '<div style="font-size:15px;font-weight:600;">Проверка перелинковки</div>'
      + '<div id="tm-relink-close" style="cursor:pointer;color:#888;font-size:18px;line-height:1;">✕</div></div>';

    if (s.error) {
      h += `<div style="color:#ef9a9a;margin-top:8px;">${esc(s.error)}</div>`;
      panel.innerHTML = h;
      panel.querySelector('#tm-relink-close').onclick = () => panel.remove();
      return;
    }

    h += section('Урок');
    h += line(`Номер: <b>${esc(s.ctx.lessonNumber || '—')}</b> · Класс: ${esc(s.ctx.classNumber || '—')} · Год: ${esc(s.ctx.year || '—')}`);
    if (!s.ctx.lessonNumber) h += dotLine('orange', 'Номер урока не распознан из заголовка — проверки папок на сервере могут не сработать');

    h += section('Материалы');
    for (const b of s.bundle) {
      const tag = b.isEmbed ? ' <span style="color:#7ec8ff;">embed</span>' : '';
      h += line(`<span style="color:#aaa;">${esc(ROLE_LABELS[b.role] || '—')}:</span> ${esc(b.name)}${tag}`);
    }

    h += section('Браузер');
    h += dotLine(s.embed.status, s.embed.text);

    h += section('Сервер');
    if (s.loadingServer) {
      h += `<div style="color:#888;">Считаю на сервере (папки, слайды, ссылки)…</div>`;
    } else if (s.serverError) {
      h += dotLine('grey', 'Сервер недоступен: ' + s.serverError);
    } else if (s.server && !s.server.ok) {
      h += dotLine('red', s.server.error || 'Ошибка сервера');
    } else if (s.server) {
      for (const c of (s.server.checks || [])) {
        h += `<div style="margin:8px 0 2px;font-weight:600;">${relinkDot(c.status)} ${esc(c.title)}</div>`;
        for (const d of (c.details || [])) {
          h += `<div style="margin-left:18px;color:#cfcfcf;">${relinkDot(d.status)} ${esc(d.text)}</div>`;
        }
      }
    }

    panel.innerHTML = h;
    panel.querySelector('#tm-relink-close').onclick = () => panel.remove();
  }

  async function checkRelinking() {
    const panel = openRelinkPanel();
    const bundle = collectBundle();
    if (!bundle.length) { renderRelink(panel, { error: 'Методические материалы на странице не найдены' }); return; }

    const ctx = parseLessonContext();
    const levelIds = collectLessonLevelIds();
    const embed = checkEmbed(bundle);

    renderRelink(panel, { bundle, ctx, embed, loadingServer: true });

    if (!RELINK_URL || RELINK_URL.indexOf('YOUR_') !== -1) {
      renderRelink(panel, { bundle, ctx, embed, serverError: 'не задан RELINK_URL в скрипте' });
      return;
    }

    try {
      const response = await fetch(RELINK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'checkRelinking',
          lessonNumber: ctx.lessonNumber, classNumber: ctx.classNumber, year: ctx.year,
          materials: bundle, levelIds
        })
      });
      const text = await response.text();
      let data;
      try { data = JSON.parse(text); }
      catch (e) { renderRelink(panel, { bundle, ctx, embed, serverError: 'ответ не JSON (бэк не обновлён?)' }); return; }
      renderRelink(panel, { bundle, ctx, embed, server: data });
    } catch (err) {
      renderRelink(panel, { bundle, ctx, embed, serverError: err.message });
    }
  }

  // ─── ТУЛБАР ──────────────────────────────────────────────────────────────────
  function createToolbar() {
    if (document.getElementById('tm-toolbar')) return;

    const bar = document.createElement('div');
    bar.id = 'tm-toolbar';

    Object.assign(bar.style, {
      position: 'fixed', top: '0', left: '0', width: '100%',
      zIndex: '2147483647', display: 'flex', justifyContent: 'flex-end',
      gap: '10px', alignItems: 'center', padding: '8px 12px',
      background: 'rgba(33,33,33,0.95)', boxShadow: '0 2px 8px rgba(0,0,0,0.25)',
      backdropFilter: 'blur(4px)', transition: 'transform 0.25s ease', boxSizing: 'border-box'
    });

    function makeBtn(text, bg, onClick) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = text;
      Object.assign(btn.style, {
        padding: '8px 12px', border: 'none', borderRadius: '6px',
        background: bg, color: 'white', cursor: 'pointer',
        fontSize: '13px', whiteSpace: 'nowrap'
      });
      btn.onclick = onClick;
      return btn;
    }

    function makeDropdown(label, bg, items) {
      const wrap = document.createElement('div');
      wrap.style.position = 'relative';
      const btn = makeBtn(label + ' ▾', bg, () => toggle());
      const menu = document.createElement('div');
      Object.assign(menu.style, {
        position: 'absolute', top: 'calc(100% + 6px)', right: '0',
        minWidth: '210px', background: '#1e1e1e', border: '1px solid #333',
        borderRadius: '8px', boxShadow: '0 12px 32px rgba(0,0,0,0.5)',
        padding: '6px', display: 'none', flexDirection: 'column', gap: '4px',
        zIndex: '2147483647'
      });
      items.forEach(it => {
        const mi = document.createElement('button');
        mi.type = 'button';
        mi.textContent = it.label;
        Object.assign(mi.style, {
          padding: '8px 12px', border: 'none', borderRadius: '6px',
          background: 'transparent', color: '#e8e8e8', cursor: 'pointer',
          fontSize: '13px', textAlign: 'left', whiteSpace: 'nowrap'
        });
        mi.onmouseenter = () => { mi.style.background = '#333'; };
        mi.onmouseleave = () => { mi.style.background = 'transparent'; };
        mi.onclick = () => { close(); it.onClick(); };
        menu.appendChild(mi);
      });
      function onDoc(e) { if (!wrap.contains(e.target)) close(); }
      function open() { menu.style.display = 'flex'; setTimeout(() => document.addEventListener('click', onDoc), 0); }
      function close() { menu.style.display = 'none'; document.removeEventListener('click', onDoc); }
      function toggle() { menu.style.display === 'flex' ? close() : open(); }
      wrap.appendChild(btn);
      wrap.appendChild(menu);
      return wrap;
    }

    const PROJECTS = [
      {
        label: '🧮 Математика АЗ',
        actions: [
          { label: '🔗 Перелинковка', onClick: checkRelinking },
          { label: '📋 Ссылки на материалы', onClick: copyMaterials },
          { label: '⬇️ Скачать материалы', onClick: downloadMaterials }
        ]
      }
    ];

    const projectDropdowns = PROJECTS.map(p => makeDropdown(p.label, '#00897b', p.actions));

    const exportBtn = makeBtn('Выгрузить уровни', '#2e7d32', exportRows);
    const clearBtn = makeBtn('Очистить LMS_LEVELS', '#c62828', clearRows);
    const webhookBtn = makeBtn('Выбрать webhook', '#1565c0', () => chooseWebhook());
    webhookBtn.id = 'tm-webhook-btn';
    const sheetBtn = makeBtn('📊 Моя таблица', '#6a1e8a', () => {});
    sheetBtn.id = 'tm-sheet-btn';
    sheetBtn.style.display = 'none';

    const badge = document.createElement('div');
    badge.id = 'tm-webhook-badge';
    badge.title = 'Сменить webhook';
    badge.onclick = () => chooseWebhook();
    Object.assign(badge.style, {
      padding: '6px 10px', background: '#424242', borderRadius: '6px',
      fontSize: '12px', color: 'white', whiteSpace: 'nowrap', cursor: 'pointer',
      maxWidth: '260px', overflow: 'hidden', textOverflow: 'ellipsis'
    });

    projectDropdowns.forEach(dd => bar.appendChild(dd));
    bar.appendChild(exportBtn);
    bar.appendChild(clearBtn);
    bar.appendChild(webhookBtn);
    bar.appendChild(sheetBtn);
    bar.appendChild(badge);

    document.body.appendChild(bar);
    updateWebhookBadge();
    document.body.style.paddingTop = '52px';

    let lastScroll = window.scrollY;
    let isHover = false;

    function showBar() { bar.style.transform = 'translateY(0)'; }
    function hideBar() { if (isHover) return; bar.style.transform = 'translateY(-100%)'; }

    window.addEventListener('scroll', () => {
      const current = window.scrollY;
      if (current > lastScroll + 5) hideBar();
      else if (current < lastScroll - 5) showBar();
      lastScroll = current;
    }, { passive: true });

    bar.addEventListener('mouseenter', () => { isHover = true; showBar(); });
    bar.addEventListener('mouseleave', () => { isHover = false; });
    document.addEventListener('mousemove', (e) => { if (e.clientY < 20) showBar(); });
  }

  setTimeout(createToolbar, 1000);
})();
