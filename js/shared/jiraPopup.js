// ── Shared Jira open-items popup ─────────────────────────────────────────────
// Singleton DOM popup reused by GraphRenderer, StatusView and JengaStatusView.
//
// Usage:
//   import { openJiraPopup, closeJiraPopup } from '../shared/jiraPopup.js';
//   openJiraPopup({ id: 'My Service' }, 'incident', cards, clickEvent, {
//       onServiceClick: serviceId => { /* navigate to service */ }
//   });

const POPUP_ID   = 'domino-jira-cards-popup';
const STYLES_ID  = 'domino-jira-popup-styles';

let _onServiceClickCb = null;

// ── CSS (injected once) ───────────────────────────────────────────────────────

function _ensureStyles() {
    if (document.getElementById(STYLES_ID)) return;
    const style = document.createElement('style');
    style.id = STYLES_ID;
    style.textContent = `
        .domino-jira-popup {
            position: absolute;
            z-index: 1200;
            display: none;
            width: min(440px, calc(100vw - 32px));
            max-height: min(520px, calc(100vh - 48px));
            color: #f5f5f7;
            background: rgba(24,24,28,.96);
            border: 1px solid rgba(255,255,255,.14);
            border-radius: 14px;
            box-shadow: 0 22px 70px rgba(0,0,0,.42);
            backdrop-filter: blur(14px);
            overflow: hidden;
        }
        .domino-jira-popup.open { display: flex; flex-direction: column; }
        .domino-jira-popup__header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
            padding: 12px 14px;
            border-bottom: 1px solid rgba(255,255,255,.10);
            cursor: move;
            user-select: none;
        }
        .domino-jira-popup__eyebrow {
            margin-bottom: 3px;
            color: rgba(255,255,255,.48);
            font-size: 10px;
            font-weight: 800;
            letter-spacing: .14em;
            text-transform: uppercase;
        }
        .domino-jira-popup__title {
            color: #fff;
            font-size: 13px;
            font-weight: 800;
            line-height: 1.25;
        }
        .domino-jira-popup__close {
            width: 28px;
            height: 28px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            border: 1px solid rgba(255,255,255,.14);
            border-radius: 8px;
            color: #fff;
            background: rgba(255,255,255,.07);
            cursor: pointer;
            font-size: 18px;
            line-height: 1;
        }
        .domino-jira-popup__close:hover { background: rgba(255,255,255,.14); }
        .domino-jira-popup__body {
            padding: 12px;
            overflow: auto;
        }
        .domino-jira-card {
            padding: 10px 10px 11px;
            border: 1px solid rgba(255,255,255,.10);
            border-radius: 10px;
            background: rgba(255,255,255,.055);
        }
        .domino-jira-card + .domino-jira-card { margin-top: 8px; }
        .domino-jira-card__topline {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 10px;
            margin-bottom: 6px;
        }
        .domino-jira-card__key-group {
            display: flex;
            align-items: center;
            gap: 6px;
            min-width: 0;
        }
        .domino-jira-card__key,
        .domino-jira-card__key a {
            color: #93c5fd;
            font-size: 12px;
            font-weight: 850;
            text-decoration: none;
        }
        .domino-jira-card__key a:hover { text-decoration: underline; }
        .domino-jira-card__status {
            max-width: 160px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
            color: rgba(255,255,255,.64);
            font-size: 11px;
        }
        .domino-jira-card__summary {
            color: #fff;
            font-size: 12px;
            line-height: 1.35;
        }
        .domino-jira-card__roi {
            display: inline-block;
            font-size: 10px;
            font-weight: 600;
            padding: 1px 5px;
            border-radius: 3px;
            background: rgba(37,99,235,0.25);
            color: #93c5fd;
            margin-top: 4px;
        }
        html:not([data-theme="dark"]) .domino-jira-card__roi {
            background: #eff6ff;
            color: #1d4ed8;
        }
        .domino-jira-card__services {
            margin-top: 7px;
            color: rgba(255,255,255,.62);
            font-size: 11px;
            line-height: 1.35;
        }
        .domino-jira-popup__empty {
            padding: 22px;
            color: rgba(255,255,255,.62);
            text-align: center;
        }
        .domino-jira-popup--incident .domino-jira-popup__header { border-top: 3px solid #ef4444; }
        .domino-jira-popup--request .domino-jira-popup__header { border-top: 3px solid #f59e0b; }
        html:not([data-theme="dark"]) .domino-jira-popup {
            color: #111827;
            background: rgba(255,255,255,.98);
            border-color: rgba(0,0,0,.12);
            box-shadow: 0 22px 70px rgba(0,0,0,.18);
        }
        html:not([data-theme="dark"]) .domino-jira-popup__header { border-bottom-color: rgba(0,0,0,.10); }
        html:not([data-theme="dark"]) .domino-jira-popup__eyebrow { color: rgba(0,0,0,.45); }
        html:not([data-theme="dark"]) .domino-jira-popup__title,
        html:not([data-theme="dark"]) .domino-jira-card__summary { color: #111827; }
        html:not([data-theme="dark"]) .domino-jira-popup__close {
            color: #111827;
            border-color: rgba(0,0,0,.10);
            background: rgba(0,0,0,.04);
        }
        html:not([data-theme="dark"]) .domino-jira-card {
            border-color: rgba(0,0,0,.10);
            background: rgba(0,0,0,.035);
        }
        html:not([data-theme="dark"]) .domino-jira-card__services,
        html:not([data-theme="dark"]) .domino-jira-card__status { color: rgba(0,0,0,.58); }
        body.domino-jira-popup-dragging { user-select: none; }
    `;
    document.head.appendChild(style);
}

// ── DOM helpers ───────────────────────────────────────────────────────────────

function _makePopupDraggable(popup, handle) {
    if (!popup || !handle || handle.dataset.dragBound === '1') return;
    handle.dataset.dragBound = '1';

    let startX = 0, startY = 0, originX = 0, originY = 0, dragging = false;

    const onMove = (event) => {
        if (!dragging) return;
        popup.style.left = `${originX + event.clientX - startX}px`;
        popup.style.top  = `${originY + event.clientY - startY}px`;
    };

    const onUp = () => {
        dragging = false;
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        document.body.classList.remove('domino-jira-popup-dragging');
    };

    handle.addEventListener('mousedown', (event) => {
        if (event.target.closest('button, a')) return;
        dragging = true;
        const rect = popup.getBoundingClientRect();
        startX = event.clientX;
        startY = event.clientY;
        originX = rect.left + window.scrollX;
        originY = rect.top  + window.scrollY;
        document.body.classList.add('domino-jira-popup-dragging');
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
        event.preventDefault();
    });
}

function _ensurePopup() {
    const existing = document.getElementById(POPUP_ID);
    if (existing) return existing;

    _ensureStyles();

    const popup = document.createElement('div');
    popup.id = POPUP_ID;
    popup.className = 'domino-jira-popup';
    popup.setAttribute('aria-hidden', 'true');
    popup.innerHTML = `
        <div class="domino-jira-popup__header">
            <div>
                <div class="domino-jira-popup__eyebrow">Jira open items</div>
                <div class="domino-jira-popup__title">Open items</div>
            </div>
            <button class="domino-jira-popup__close" type="button" aria-label="Close">×</button>
        </div>
        <div class="domino-jira-popup__body"></div>
    `;
    document.body.appendChild(popup);

    popup.querySelector('.domino-jira-popup__close')
        ?.addEventListener('click', () => closeJiraPopup());

    popup.querySelector('.domino-jira-popup__body')
        ?.addEventListener('click', (event) => {
            const trigger = event.target.closest('.domino-jira-card__service-link');
            if (!trigger) return;
            event.preventDefault();
            event.stopPropagation();
            const serviceId = trigger.dataset.serviceId || trigger.textContent || '';
            if (_onServiceClickCb) _onServiceClickCb(serviceId);
        });

    _makePopupDraggable(popup, popup.querySelector('.domino-jira-popup__header'));
    return popup;
}

function _jiraCardHtml(card) {
    const esc = v => String(v ?? '')
        .replace(/&/g, '&amp;').replace(/</g, '&lt;')
        .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');

    const key = esc(card.key || 'No key');
    const summary = esc(card.summary || 'No summary');
    const status  = esc(card.status || '');
    const roiBadge = card.roi ? `<span class="domino-jira-card__roi">${esc(card.roi)}</span>` : '';
    const priorityBadge = card.priority
        ? `<span class="domino-jira-card__priority domino-jira-card__priority--${card.priority.toLowerCase()}">${esc(card.priority)}</span>`
        : '';
    const internalUrl = String(card.jiraUrl || '').trim();
    const url = internalUrl.replace('browse', 'servicedesk/customer/portal/1');
    const keyHtml = url
        ? `<a href="${esc(url)}" target="_blank" rel="noopener noreferrer">${key}</a>`
        : key;

    const impactedServicesHtml = (card.impactedServices || [])
        .map(s => String(s || '').trim())
        .filter(Boolean)
        .map(s => {
            const safe = esc(s);
            return `<button type="button" class="domino-jira-card__service-link" data-service-id="${safe}" title="Search id:&quot;${safe}&quot;">${safe}</button>`;
        })
        .join('');

    return `
        <article class="domino-jira-card">
            <div class="domino-jira-card__topline">
                <span class="domino-jira-card__key-group">
                    <span class="domino-jira-card__key">${keyHtml}</span>
                    ${priorityBadge}
                </span>
                ${status ? `<span class="domino-jira-card__status">${status}</span>` : ''}
            </div>
            <div class="domino-jira-card__summary">${summary}</div>
            ${roiBadge}
            <div class="domino-jira-card__services">
                <strong>Impacted services:</strong>
                ${impactedServicesHtml || '<span class="domino-jira-card__no-services">—</span>'}
            </div>
        </article>
    `;
}

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Open the Jira open-items popup.
 *
 * @param {{ id: string }} node   - service node; only `.id` is used for the title
 * @param {'incident'|'request'} type
 * @param {object[]} cards        - card objects with key, summary, status, priority, jiraUrl, impactedServices
 * @param {MouseEvent|null} event - click event for positioning
 * @param {{ onServiceClick?: (serviceId: string) => void }} [opts]
 */
export function openJiraPopup(node, type, cards = [], event = null, { onServiceClick = null } = {}) {
    const popup = _ensurePopup();
    _onServiceClickCb = onServiceClick;

    const title = popup.querySelector('.domino-jira-popup__title');
    const body  = popup.querySelector('.domino-jira-popup__body');
    const isIncident = type === 'incident';
    const safeCards  = Array.isArray(cards) ? cards : [];

    popup.classList.toggle('domino-jira-popup--incident', isIncident);
    popup.classList.toggle('domino-jira-popup--request', !isIncident);
    title.textContent = `${safeCards.length} ${isIncident ? 'open incident' : 'open service request'}${safeCards.length === 1 ? '' : 's'} · ${node.id}`;

    const PRIORITY_ORDER = { P1: 0, P2: 1, P3: 2, P4: 3 };
    const sortedCards = [...safeCards].sort((a, b) =>
        (PRIORITY_ORDER[a.priority] ?? 99) - (PRIORITY_ORDER[b.priority] ?? 99)
    );
    body.innerHTML = sortedCards.length
        ? sortedCards.map(_jiraCardHtml).join('')
        : '<div class="domino-jira-popup__empty">No open items found for this service.</div>';

    const x = Math.min((event?.pageX ?? 40) + 16, window.scrollX + window.innerWidth - 460);
    const y = Math.min((event?.pageY ?? 80) + 16, window.scrollY + window.innerHeight - 360);
    popup.style.left = `${Math.max(window.scrollX + 16, x)}px`;
    popup.style.top  = `${Math.max(window.scrollY + 16, y)}px`;
    popup.classList.add('open');
    popup.setAttribute('aria-hidden', 'false');
}

export function closeJiraPopup() {
    const popup = document.getElementById(POPUP_ID);
    popup?.classList.remove('open');
    popup?.setAttribute('aria-hidden', 'true');
}
