/**
 * Shared utilities used across Domino and Solitaire.
 * No D3 dependency. No app-specific logic.
 */

// ─── String / value helpers ────────────────────────────────────────────────

/**
 * Parse §[...] metadata appended by Jira automation to summary strings.
 * e.g. "Title§[TopPriority P2/Critical,RoiUSCA]" → { clean: "Title", roi: "USCA", topPriority: "P2/Critical" }
 */
export function parseSectionMeta(raw) {
    if (!raw) return { clean: raw || '', roi: null, topPriority: null };
    const match = raw.match(/§\[([^\]]*)\]/);
    if (!match) return { clean: raw, roi: null, topPriority: null };
    const after = raw.slice(match.index + match[0].length).trim();
    const clean = (raw.slice(0, match.index).trim() + (after ? ' ' + after : '')).trim();
    const tokens = match[1].split(',').map(t => t.trim());
    let roi = null, topPriority = null;
    for (const t of tokens) {
        const roiM = t.match(/^Roi(.+)$/i);
        if (roiM) { roi = roiM[1]; continue; }
        const tpM = t.match(/^TopPriority\s+(.+)$/i);
        if (tpM) { topPriority = tpM[1]; }
    }
    return { clean, roi, topPriority };
}

const fullNormalizeWs = (s) => (s ?? '')
    .toString()
    .replace(/\s+/g, ' ')
    .trim();

const FIELDS_WITH_WS_NORMALIZATION = new Set(['Name', 'User', 'Email']);

export const normalizeWs = (value, fieldName) => {
    const raw = (value ?? '').toString();
    return !fieldName || FIELDS_WITH_WS_NORMALIZATION.has(fieldName) ? fullNormalizeWs(raw) : raw.trim();
};

export function normalizeKey(s) {
    return (s ?? '')
        .toString()
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '_')
        .replace(/[^a-z0-9_-]/g, '');
}

export function truncateString(str, maxLength = 25) {
    if (str.length <= maxLength) return str;
    return str.slice(0, maxLength) + '...';
}

const URL_RE = /^https?:\/\/\S+$/i;
export function isUrl(v) {
    return URL_RE.test(v);
}

export function splitValues(raw) {
    if (!raw) return [];
    return raw
        .toString()
        .split(/\s*\|\|\s*|\n|,/)
        .map(s => s.trim())
        .filter(Boolean);
}

export function splitNarrativeValues(raw) {
    if (!raw) return [];
    const s = raw
        .toString()
        .replace(/\r?\n/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    return s
        .split(/\s*\|\|\s*/)
        .map(x => x.trim())
        .filter(Boolean);
}

// ─── Date helpers ──────────────────────────────────────────────────────────

export function isDateTimeValue(value) {
    if (typeof value !== 'string') return false;
    const isoRegex = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})$/;
    if (!isoRegex.test(value)) return false;
    const d = new Date(value);
    return !isNaN(d.getTime());
}

export function formatDateTimeLocal(value) {
    const d = new Date(value);
    if (isNaN(d.getTime())) return value;
    return new Intl.DateTimeFormat(undefined, {
        year: 'numeric',
        month: 'short',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        timeZoneName: 'short'
    }).format(d);
}

export function getFormattedDate(isoDate, locale = 'it-IT', timeZone = 'Europe/Rome') {
    const date = new Date(isoDate);
    return date.toLocaleString(locale, {
        timeZone,
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
    });
}

export function formatMonthYear(value) {
    const d = new Date(value);
    if (isNaN(d)) return value;
    return d.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
}

// ─── URL / query params ────────────────────────────────────────────────────

export function getQueryParam(param) {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(param);
}

export function setQueryParam(param, value) {
    const url = new URL(window.location);
    if (value === undefined || value === null) return;
    url.searchParams.set(param, value);
    window.history.pushState({}, '', url);
}

export function removeQueryParam(param) {
    const url = new URL(window.location);
    url.searchParams.delete(param);
    window.history.pushState({}, '', url);
}

export function setSearchQuery(search) {
    setQueryParam('search', search);
}

// ─── Platform detection ────────────────────────────────────────────────────

export function isMobileDevice() {
    try {
        if (navigator.userAgentData && typeof navigator.userAgentData.mobile === 'boolean') {
            return navigator.userAgentData.mobile;
        }
    } catch (_) {}
    const ua = (navigator.userAgent || navigator.vendor || window.opera || '').toLowerCase();
    const uaIsMobile =
        /android|iphone|ipod|ipad|iemobile|mobile|blackberry|opera mini|opera mobi|silk/.test(ua) ||
        ((/macintosh/.test(ua) || /mac os x/.test(ua)) && 'ontouchend' in document);
    const smallViewport = Math.min(window.screen.width, window.screen.height) <= 820;
    return uaIsMobile || smallViewport;
}

// ─── DOM element builders ──────────────────────────────────────────────────

export function createHrefElement(cleanUrl, textContent) {
    const a = document.createElement('a');
    a.href = cleanUrl;
    a.textContent = textContent ?? '🔗Link';
    a.target = '_blank';
    a.rel = 'noopener noreferrer';
    a.style.color = '#0078d4';
    a.style.textDecoration = 'underline';
    return a;
}

export function addTagToElement(element, number, tag = 'br') {
    element.insertAdjacentHTML('beforeend', `<${tag}>`.repeat(number));
}

const allowedAttributesByTag = {
    'a': new Set(['href', 'title', 'target', 'rel']),
};

function sanitizeUrl(url) {
    if (typeof url !== 'string') return '';
    const trimmed = url.trim();
    const lower = trimmed.toLowerCase();
    const forbiddenSchemes = ['javascript:', 'vbscript:'];
    if (forbiddenSchemes.some(s => lower.startsWith(s))) return '';
    try {
        const u = new URL(trimmed, window.location.origin);
        const allowed = ['http:', 'https:', 'mailto:', 'tel:', 'ftp:'];
        if (!allowed.includes(u.protocol) && !trimmed.startsWith('/')) return '';
    } catch (_) {}
    return trimmed;
}

function copyAllowedAttributes(srcElem, dstElem, allowedAttrsByTag) {
    const tag = srcElem.tagName.toLowerCase();
    const allowedAttrs = allowedAttrsByTag[tag];
    if (!allowedAttrs) return;
    for (const attr of srcElem.attributes) {
        const name = attr.name.toLowerCase();
        if (!allowedAttrs.has(name)) continue;
        let value = attr.value;
        if (tag === 'a') {
            if (name === 'href') {
                value = sanitizeUrl(value);
                if (!value) continue;
            }
            if (name === 'target') {
                const allowedTargets = new Set(['_blank', '_self']);
                if (!allowedTargets.has(value)) value = '_blank';
            }
            if (name === 'rel') {
                const parts = new Set(value.split(/\s+/).filter(Boolean).map(v => v.toLowerCase()));
                parts.add('noopener');
                parts.add('noreferrer');
                value = Array.from(parts).join(' ');
            }
        }
        dstElem.setAttribute(name, value);
    }
    if (tag === 'a' && dstElem.hasAttribute('href')) {
        if (!dstElem.hasAttribute('rel')) dstElem.setAttribute('rel', 'noopener noreferrer');
        if (!dstElem.hasAttribute('target')) dstElem.setAttribute('target', '_blank');
    }
}

function textNodeWithLinksToNodes(text) {
    const nodes = [];
    const urlRe = /(https?:\/\/[^\s<>"')\]]+)/g;
    let lastIndex = 0;
    let match;
    while ((match = urlRe.exec(text)) !== null) {
        const before = text.slice(lastIndex, match.index);
        if (before) nodes.push(document.createTextNode(before));
        let url = match[1];
        const trailingPunct = /[.,;:!?)+\]]+$/;
        let punct = '';
        const m2 = url.match(trailingPunct);
        if (m2) { punct = m2[0]; url = url.slice(0, -punct.length); }
        nodes.push(createHrefElement(url));
        if (punct) nodes.push(document.createTextNode(punct));
        lastIndex = urlRe.lastIndex;
    }
    const rest = text.slice(lastIndex);
    if (rest) nodes.push(document.createTextNode(rest));
    return nodes;
}

function sanitizeAndTransformNode(node, allowedTags) {
    if (node.nodeType === Node.TEXT_NODE) {
        return textNodeWithLinksToNodes(node.nodeValue || '');
    }
    if (node.nodeType === Node.ELEMENT_NODE) {
        const normalizedTag = node.tagName.toLowerCase();
        if (allowedTags.has(normalizedTag)) {
            const clone = document.createElement(normalizedTag);
            copyAllowedAttributes(node, clone, allowedAttributesByTag);
            node.childNodes.forEach(child => {
                const childParts = sanitizeAndTransformNode(child, allowedTags);
                childParts.forEach(p => clone.appendChild(p));
            });
            if (normalizedTag === 'a' && !clone.getAttribute('href')) {
                const fragmentNodes = [];
                clone.childNodes.forEach(c => fragmentNodes.push(c));
                return fragmentNodes;
            }
            return [clone];
        }
        const fragmentNodes = [];
        node.childNodes.forEach(child => {
            const childParts = sanitizeAndTransformNode(child, allowedTags);
            childParts.forEach(p => fragmentNodes.push(p));
        });
        return fragmentNodes;
    }
    return [];
}

export function createFormattedElementsFrom(lines) {
    const elementsToAppend = [];
    const allowedTags = new Set(['b', 'i', 'ul', 'li', 'a']);
    lines.forEach((line, index) => {
        const template = document.createElement('template');
        template.innerHTML = line;
        Array.from(template.content.childNodes).forEach(node => {
            const parts = sanitizeAndTransformNode(node, allowedTags);
            parts.forEach(p => elementsToAppend.push(p));
        });
        if (index < lines.length - 1) {
            elementsToAppend.push(document.createElement('br'));
        }
    });
    return elementsToAppend;
}

export function createFormattedLongTextElementsFrom(longText) {
    if (!longText) return [];
    const normalized = longText.replace(/\s*\|\|\s*/g, '\n\n');
    const rawLines = normalized
        .split(/\r?\n/)
        .reduce((acc, line) => {
            const isEmpty = !line.trim();
            const prevEmpty = acc.length ? !acc[acc.length - 1].trim() : false;
            if (isEmpty && prevEmpty) return acc;
            acc.push(line);
            return acc;
        }, []);

    // Group consecutive lines into typed lists or paragraphs
    const result = [];
    let currentList = null;
    let currentListType = null; // 'bullet' | 'task'

    const flushList = () => {
        if (currentList) { result.push(currentList); currentList = null; currentListType = null; }
    };

    for (const line of rawLines) {
        const taskMatch = line.match(/^\[( |x)\]\s+(.*)$/i);
        const bulletMatch = !taskMatch && line.match(/^[-–•]\s+(.*)$/);

        if (taskMatch) {
            if (currentListType !== 'task') {
                flushList();
                currentList = document.createElement('ul');
                currentList.className = 'jenga-task-list';
                currentListType = 'task';
            }
            const li = document.createElement('li');
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            if (taskMatch[1].toLowerCase() === 'x') cb.checked = true;
            const span = document.createElement('span');
            const tpl = document.createElement('template');
            tpl.innerHTML = taskMatch[2];
            span.append(...Array.from(tpl.content.childNodes));
            li.append(cb, span);
            currentList.appendChild(li);
        } else if (bulletMatch) {
            if (currentListType !== 'bullet') {
                flushList();
                currentList = document.createElement('ul');
                currentListType = 'bullet';
            }
            const li = document.createElement('li');
            const tpl = document.createElement('template');
            tpl.innerHTML = bulletMatch[1];
            li.append(...Array.from(tpl.content.childNodes));
            currentList.appendChild(li);
        } else if (line.trimStart().startsWith('<table')) {
            flushList();
            const tpl = document.createElement('template');
            tpl.innerHTML = line;
            const srcTable = tpl.content.querySelector('table');
            if (srcTable) {
                const table = document.createElement('table');
                if (srcTable.className) table.className = srcTable.className;
                for (const srcRow of srcTable.querySelectorAll('tr')) {
                    const tr = document.createElement('tr');
                    for (const srcCell of srcRow.children) {
                        const isHeader = srcCell.tagName === 'TH';
                        const cell = document.createElement(isHeader ? 'th' : 'td');
                        createFormattedElementsFrom([srcCell.innerHTML]).forEach(el => cell.appendChild(el));
                        tr.appendChild(cell);
                    }
                    table.appendChild(tr);
                }
                result.push(table);
            }
        } else {
            flushList();
            if (line.trim()) {
                // Wrap in <p> so each non-empty line renders as a distinct paragraph block
                const p = document.createElement('p');
                createFormattedElementsFrom([line]).forEach(el => p.appendChild(el));
                result.push(p);
            }
            // Empty lines are intentional separators — handled by the gap between <p> blocks
        }
    }
    flushList();
    return result;
}

// ─── URL link / cell rendering helpers ────────────────────────────────────

export function formatUrlLink(url) {
    let clean = url.replace(/^https?:\/\//, '').split(/[?#]/)[0];
    const segments = clean.split('/').filter(Boolean);
    const segment = segments.length > 0 ? segments[segments.length - 1] || segments[segments.length - 2] || '' : '';
    const display = segment.length > 55 ? '...' + segment.slice(-55) : segment;
    return `<a href="${url}" target="_blank">${display}</a>`;
}

export function renderUrlPartsIntoCell(parts, td) {
    if (parts.length > 1) {
        const ul = document.createElement('ul');
        parts.forEach(p => {
            const li = document.createElement('li');
            li.innerHTML = isUrl(p) ? formatUrlLink(p) : p;
            ul.appendChild(li);
        });
        td.appendChild(ul);
    } else {
        td.innerHTML = formatUrlLink(parts[0]);
    }
}

// ─── Modal ─────────────────────────────────────────────────────────────────

export function createModal({ title, html, buttons }) {
    return new Promise(resolve => {
        const overlay = document.createElement('div');
        overlay.className = 'simple-modal__overlay';
        overlay.setAttribute('role', 'dialog');
        overlay.setAttribute('aria-modal', 'true');

        const modal = document.createElement('div');
        modal.className = 'simple-modal';

        const btnHtml = buttons.map(btn =>
            `<button type="button" class="simple-modal__btn ${btn.primary ? 'simple-modal__btn--primary' : ''}" data-action="${btn.id}">
        ${btn.label}
      </button>`
        ).join('');

        modal.innerHTML = `
      <h3>${title}</h3>
      <p>${html}</p>
      <div class="simple-modal__buttons">${btnHtml}</div>
    `;

        function close(val) {
            overlay.remove();
            resolve(val);
        }

        overlay.addEventListener('click', e => {
            if (e.target === overlay) close(null);
        });

        buttons.forEach(btn => {
            modal.querySelector(`[data-action="${btn.id}"]`)
                ?.addEventListener('click', () => close(btn.id));
        });

        const escHandler = (e) => {
            if (e.key === 'Escape') {
                document.removeEventListener('keydown', escHandler);
                close(null);
            }
        };
        document.addEventListener('keydown', escHandler);

        overlay.appendChild(modal);
        document.body.appendChild(overlay);
        modal.querySelector(`[data-action="${buttons.find(b => b.primary)?.id}"]`)?.focus();
    });
}

// ─── Email / Outlook ───────────────────────────────────────────────────────

export function createOutlookUrl(to, cc = [], subject = '', body = '') {
    const toParam = to.length ? encodeURIComponent(to.join(';')) : '';
    const ccParam = cc.length ? encodeURIComponent(cc.join(';')) : '';
    const subjectParam = encodeURIComponent(subject);
    const bodyParam = encodeURIComponent(body);
    let url = `https://outlook.office.com/mail/deeplink/compose?subject=${subjectParam}&body=${bodyParam}`;
    if (toParam) url += `&to=${toParam}`;
    if (ccParam) url += `&cc=${ccParam}`;
    return url;
}

export function openOutlookWebCompose({ to = [], cc = [], bcc = [], subject = '', body = '' }) {
    const url = createOutlookUrl(to, cc, subject, body);
    window.open(url, '_blank', 'noopener');
}

export function buildFallbackMailToLink(peopleDBUpdateRecipients, subjectParam, bodyParam) {
    window.location.href = `mailto:${peopleDBUpdateRecipients.join(',')}?subject=${encodeURIComponent(subjectParam)}&body=${encodeURIComponent(bodyParam)}`;
}

// ─── Side drawer ───────────────────────────────────────────────────────────

export function closeSideDrawer() {
    const drawer = document.getElementById('side-drawer');
    const overlay = document.getElementById('side-overlay');
    if (!drawer) return;
    drawer.classList.remove('open');
    overlay?.classList.remove('visible');
    document.body.classList.remove('side-drawer-open');
    drawer.setAttribute('aria-hidden', 'true');
    // Keep sub-open during the slide-off animation so sub content travels left with the drawer;
    // clean up after the transition completes
    if (drawer.classList.contains('sub-open')) {
        const onEnd = (e) => {
            if (e.propertyName !== 'transform') return;
            drawer.removeEventListener('transitionend', onEnd);
            drawer.classList.remove('sub-open');
        };
        drawer.addEventListener('transitionend', onEnd);
    }
}

export function openSideDrawer() {
    const drawer = document.getElementById('side-drawer');
    const overlay = document.getElementById('side-overlay');
    if (!drawer) return;
    drawer.classList.add('open');
    overlay?.classList.add('visible');
    document.body.classList.add('side-drawer-open');
    drawer.setAttribute('aria-hidden', 'false');
    document.getElementById('act-upload')?.focus();
}

export function initCommonActions() {
    const overlay = document.getElementById('side-overlay');
    const closeBtn = document.getElementById('side-close');
    // Overlay and close button always close everything immediately
    overlay?.addEventListener('click', closeSideDrawer);
    closeBtn?.addEventListener('click', closeSideDrawer);
    // ESC is layered: closes sub-panel first, then main drawer.
    // If the autocomplete dropdown is open, let AutocompleteEngine consume Escape first.
    window.addEventListener('keydown', (e) => {
        if (e.key !== 'Escape') return;
        if (document.getElementById('ac-dropdown')?.classList.contains('ac-open')) return;
        const drawer = document.getElementById('side-drawer');
        if (drawer?.classList.contains('sub-open')) {
            drawer.classList.remove('sub-open');
        } else if (drawer?.classList.contains('open')) {
            closeSideDrawer();
        }
    });
    const toggleCta = document.getElementById('toggle-cta');
    toggleCta?.addEventListener('click', (e) => {
        e.preventDefault();
        openSideDrawer();
    });
    document.getElementById('act-upload')?.addEventListener('click', () => {
        document.getElementById('fileInput')?.click();
        closeSideDrawer();
    });
}

export function toggleClearButton(buttonId, value) {
    const el = document.getElementById(buttonId);
    if (!el) return;
    el.classList.toggle('hidden', !value);
}

// ─── Keyboard shortcuts ────────────────────────────────────────────────────

export function enableGlobalFindShortcut({ inputSelector, onFocus, selectText = true } = {}) {
    if (!inputSelector) {
        console.warn('[enableGlobalFindShortcut] inputSelector is required');
        return;
    }
    window.addEventListener('keydown', (e) => {
        const isMac = navigator.platform.toUpperCase().includes('MAC');
        const isFindShortcut =
            (isMac && e.metaKey && e.key.toLowerCase() === 'f') ||
            (!isMac && e.ctrlKey && e.key.toLowerCase() === 'f');
        if (!isFindShortcut) return;
        const activeTag = document.activeElement?.tagName;
        const isTyping =
            activeTag === 'INPUT' || activeTag === 'TEXTAREA' || document.activeElement?.isContentEditable;
        if (isTyping) return;
        const input = document.querySelector(inputSelector);
        if (!input) return;
        e.preventDefault();
        e.stopPropagation();
        input.focus({ preventScroll: false });
        if (selectText && typeof input.select === 'function') {
            try { input.select(); } catch {}
        }
        if (typeof onFocus === 'function') onFocus(input);
    }, true);
}

// ─── CSV parser ────────────────────────────────────────────────────────────

export function parseCSV(text) {
    const rows = [];
    let current = [];
    let inQuotes = false;
    let value = '';
    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        if (char === '"') {
            if (inQuotes && text[i + 1] === '"') { value += '"'; i++; }
            else { inQuotes = !inQuotes; }
        } else if (char === ',' && !inQuotes) {
            current.push(value); value = '';
        } else if ((char === '\n' || char === '\r') && !inQuotes) {
            if (value || current.length > 0) { current.push(value); rows.push(current); current = []; value = ''; }
            if (char === '\r' && text[i + 1] === '\n') i++;
        } else {
            value += char;
        }
    }
    if (value || current.length > 0) { current.push(value); rows.push(current); }
    return rows;
}

// ─── Theme (dark / light mode) ────────────────────────────────────────────────

const THEME_KEY = 'dsm-theme-v1';

export function applyTheme(mode) {
    if (mode === 'dark') {
        document.documentElement.setAttribute('data-theme', 'dark');
    } else {
        document.documentElement.removeAttribute('data-theme');
    }
    localStorage.setItem(THEME_KEY, mode);
}

export function loadSavedTheme() {
    const saved = localStorage.getItem(THEME_KEY);
    const preferred = window.matchMedia?.('(prefers-color-scheme: dark)')?.matches ? 'dark' : 'light';
    applyTheme(saved ?? preferred);
    return saved ?? preferred;
}

// Shows a full-screen spinner overlay during a theme switch.
// Returns the overlay element; caller must remove() it when done.
export function showThemeSwitchSpinner() {
    if (!document.getElementById('_dsm-spin-kf')) {
        const s = document.createElement('style');
        s.id = '_dsm-spin-kf';
        s.textContent = '@keyframes _dsm-spin{to{transform:rotate(360deg)}}';
        document.head.appendChild(s);
    }
    const el = document.createElement('div');
    el.style.cssText = 'position:fixed;inset:0;display:flex;align-items:center;justify-content:center;z-index:99999;pointer-events:none;';
    const ring = document.createElement('div');
    ring.style.cssText = 'width:36px;height:36px;border:3px solid rgba(128,128,128,.2);border-top-color:#888;border-radius:50%;animation:_dsm-spin .7s linear infinite;';
    el.appendChild(ring);
    document.body.appendChild(el);
    return el;
}

// ─── Legend drag ──────────────────────────────────────────────────────────────

/**
 * Makes a legend element draggable with pointer events, viewport clamping,
 * and localStorage position persistence.
 *
 * @param {HTMLElement} root           - the legend container element
 * @param {object}      opts
 * @param {string}      opts.handleSelector - CSS selector for drag handle inside root
 * @param {string}      opts.storageKey     - localStorage key for position persistence
 * @param {boolean}    [opts.cornerAnchor]  - if true, snaps to viewport corners on drop
 */
export function makeLegendDraggable(root, { handleSelector, storageKey, cornerAnchor = false } = {}) {
    const clamp = (v, min, max) => Math.max(min, Math.min(max, v));
    const getViewportSize = () => ({ w: document.documentElement.clientWidth, h: document.documentElement.clientHeight });
    const getRootRect = () => root.getBoundingClientRect();

    const ANCHOR_CLASSES = ['legend--anchor-tl', 'legend--anchor-tr', 'legend--anchor-bl', 'legend--anchor-br'];

    const applyCorner = (corner) => {
        ANCHOR_CLASSES.forEach(c => root.classList.remove(c));
        root.classList.add(`legend--anchor-${corner}`);
        ['left', 'top', 'right', 'bottom'].forEach(p => root.style.removeProperty(p));
    };

    const applyFree = (x, y) => {
        ANCHOR_CLASSES.forEach(c => root.classList.remove(c));
        root.style.left = `${x}px`;
        root.style.top = `${y}px`;
        root.style.right = 'auto';
        root.style.bottom = 'auto';
    };

    const isCornerAnchored = () => ANCHOR_CLASSES.some(c => root.classList.contains(c));

    const cornerForQuadrant = (rect) => {
        const { w, h } = getViewportSize();
        const SNAP_PX = 60;
        const nearLeft  = rect.left < SNAP_PX;
        const nearRight = (rect.left + rect.width) > (w - SNAP_PX);
        const nearTop   = rect.top < SNAP_PX;
        const nearBot   = (rect.top + rect.height) > (h - SNAP_PX);
        if (nearLeft && nearTop)  return 'tl';
        if (nearRight && nearTop) return 'tr';
        if (nearLeft && nearBot)  return 'bl';
        if (nearRight && nearBot) return 'br';
        return null;
    };

    const save = (state) => {
        try { localStorage.setItem(storageKey, JSON.stringify(state)); } catch {}
    };

    const restore = () => {
        try {
            const saved = JSON.parse(localStorage.getItem(storageKey) || '{}');
            if (cornerAnchor && saved.type === 'corner' && saved.corner) {
                applyCorner(saved.corner); return true;
            }
            if (typeof saved.x === 'number' && typeof saved.y === 'number') {
                applyFree(saved.x, saved.y); return true;
            }
        } catch {}
        if (cornerAnchor) { applyCorner('bl'); return true; }
        return false;
    };

    restore();

    const reclamp = () => {
        if (cornerAnchor && isCornerAnchored()) return;
        if (cornerAnchor) {
            try {
                const saved = JSON.parse(localStorage.getItem(storageKey) || '{}');
                if (saved.type === 'corner' && saved.corner) { applyCorner(saved.corner); return; }
            } catch {}
        }
        const rect = getRootRect();
        const { w, h } = getViewportSize();
        const nx = clamp(rect.left, 0, w - rect.width);
        const ny = clamp(rect.top, 0, h - rect.height);
        applyFree(nx, ny);
        save(cornerAnchor ? { type: 'free', x: nx, y: ny } : { x: nx, y: ny });
    };
    requestAnimationFrame(reclamp);
    window.addEventListener('resize', reclamp);

    const THRESHOLD = 4;
    let startX = 0, startY = 0, startLeft = 0, startTop = 0;
    let dragging = false, pointerId = null;

    const handle = handleSelector ? root.querySelector(handleSelector) : root;
    const dragClassEl = handle || root;
    if (!handle) return;

    const onPointerDown = (e) => {
        if (e.button !== 0) return;
        pointerId = e.pointerId;
        const rect = getRootRect();
        const cs = window.getComputedStyle(root);
        startLeft = isCornerAnchored() ? rect.left : (parseFloat(cs.left) || rect.left);
        startTop  = isCornerAnchored() ? rect.top  : (parseFloat(cs.top)  || rect.top);
        startX = e.clientX; startY = e.clientY;
        window.addEventListener('pointermove', onPointerMove, { passive: true });
        window.addEventListener('pointerup', onPointerUp, { passive: true });
    };

    const onPointerMove = (e) => {
        const dx = e.clientX - startX;
        const dy = e.clientY - startY;
        if (!dragging) {
            if (Math.abs(dx) <= THRESHOLD && Math.abs(dy) <= THRESHOLD) return;
            dragging = true;
            dragClassEl.classList.add('is-dragging');
            try { dragClassEl.setPointerCapture?.(pointerId); } catch {}
        }
        const { w, h } = getViewportSize();
        const r = getRootRect();
        applyFree(clamp(startLeft + dx, 0, w - r.width), clamp(startTop + dy, 0, h - r.height));
    };

    const onPointerUp = () => {
        window.removeEventListener('pointermove', onPointerMove);
        window.removeEventListener('pointerup', onPointerUp);
        if (!dragging) return;
        dragClassEl.classList.remove('is-dragging');
        dragging = false; pointerId = null;
        const rect = getRootRect();
        const { w, h } = getViewportSize();
        const fx = clamp(rect.left, 0, w - rect.width);
        const fy = clamp(rect.top,  0, h - rect.height);
        if (cornerAnchor) {
            const corner = cornerForQuadrant({ left: fx, top: fy, width: rect.width, height: rect.height });
            if (corner) {
                applyCorner(corner);
                save({ type: 'corner', corner });
                return;
            }
        }
        applyFree(fx, fy);
        save(cornerAnchor ? { type: 'free', x: fx, y: fy } : { x: fx, y: fy });
    };

    handle.style.touchAction = 'none';
    handle.addEventListener('pointerdown', onPointerDown);
}
