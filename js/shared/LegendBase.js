import { makeLegendDraggable } from './utils.js';

export class LegendBase {
    constructor() {
        this._dragAttached = false;
        this._resizeAttached = false;
    }

    _getOrCreateRoot(id) {
        let root = document.getElementById(id);
        if (!root) {
            root = document.createElement('div');
            root.id = id;
            document.body.appendChild(root);
        }
        return root;
    }

    _buildShell(root, title) {
        this._dragAttached = false;
        this._resizeAttached = false;
        root.className = 'legend legend--generic';
        root.innerHTML = `
  <div class="legend__header" aria-label="Legend header">
    <div class="legend__title" role="heading" aria-level="2"></div>
    <div class="legend__search-wrap" hidden>
      <input class="legend__search-input" type="text" autocomplete="off" spellcheck="false" aria-label="Search legend" />
      <span class="legend__search-count" aria-live="polite" hidden></span>
    </div>
    <div class="legend__header-actions">
      <button class="legend__search-toggle" type="button" aria-label="Search in legend" aria-pressed="false">
        <svg width="12" height="12" viewBox="0 0 14 14" fill="none" aria-hidden="true">
          <circle cx="6" cy="6" r="4.5" stroke="currentColor" stroke-width="1.5"></circle>
          <line x1="9.5" y1="9.5" x2="13" y2="13" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"></line>
        </svg>
      </button>
      <button class="legend__search-close" type="button" aria-label="Close search" hidden>
        <svg width="12" height="12" viewBox="0 0 14 14" fill="none" aria-hidden="true">
          <line x1="2" y1="2" x2="12" y2="12" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"></line>
          <line x1="12" y1="2" x2="2" y2="12" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"></line>
        </svg>
      </button>
      <button class="legend__filter-toggle" type="button" aria-label="Filter chart by legend search" aria-pressed="false">
        <svg width="12" height="12" viewBox="0 0 12 12" fill="none" aria-hidden="true">
          <path d="M1 1.5h10L7 6v4.5L5 10.5V6L1 1.5z" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/>
        </svg>
      </button>
      <button class="legend__collapse" type="button" aria-label="Toggle legend" aria-expanded="true">
        <span class="chevron" aria-hidden="true"></span>
      </button>
    </div>
  </div>
  <div class="legend__list" aria-label="Legend list"></div>
  <div class="legend__resize-handle legend__resize-handle--nw" aria-hidden="true"></div>
  <div class="legend__resize-handle legend__resize-handle--ne" aria-hidden="true"></div>
  <div class="legend__resize-handle legend__resize-handle--sw" aria-hidden="true"></div>
  <div class="legend__resize-handle legend__resize-handle--se" aria-hidden="true"></div>
`;
        root.querySelector('.legend__title').textContent = title;
    }

    _wireCollapse(root, lsKey) {
        const list = root.querySelector('.legend__list');
        const btn  = root.querySelector('.legend__collapse');
        const applyCollapsed = (collapsed) => {
            root.classList.toggle('legend--collapsed', collapsed);
            list.hidden = collapsed;
            btn.setAttribute('aria-expanded', String(!collapsed));
            btn.setAttribute('aria-label', collapsed ? 'Expand legend' : 'Collapse legend');
            try { localStorage.setItem(lsKey, collapsed ? '1' : '0'); } catch {}
        };
        let initial = false;
        try { initial = localStorage.getItem(lsKey) === '1'; } catch {}
        applyCollapsed(initial);
        btn.addEventListener('click', (e) => {
            e.preventDefault(); e.stopPropagation();
            applyCollapsed(!root.classList.contains('legend--collapsed'));
        });
        btn.addEventListener('pointerdown', (e) => e.stopPropagation());
        return { list, btn };
    }

    _wireListEvents(list, activateFn) {
        list.addEventListener('click', (e) => {
            const el = e.target.closest('.legend__item');
            if (!el) return;
            activateFn(el);
        });
        list.addEventListener('keydown', (e) => {
            if (e.key !== 'Enter' && e.key !== ' ') return;
            const el = e.target.closest('.legend__item');
            if (!el) return;
            e.preventDefault();
            activateFn(el);
        });
    }

    _isFilterActive(root) {
        return root?.querySelector('.legend__filter-toggle')?.getAttribute('aria-pressed') === 'true';
    }

    _wireFilterToggle(root, lsKey) {
        const btn = root?.querySelector('.legend__filter-toggle');
        if (!btn) return;
        let active = false;
        try { active = localStorage.getItem(lsKey) === '1'; } catch {}
        btn.setAttribute('aria-pressed', String(active));
        btn.dataset.legendFilterKey = lsKey;

        // Open or close the legend's search-wrap to match filter state.
        // Deferred so _wireSearch event handlers are already registered.
        const syncSearchWrap = (on) => {
            const wrap       = root.querySelector('.legend__search-wrap');
            const searchToggle = root.querySelector('.legend__search-toggle');
            const closeBtn   = root.querySelector('.legend__search-close');
            if (on) {
                if (wrap?.hidden) {
                    // Directly open without auto-focusing (less jarring on page load)
                    const title = root.querySelector('.legend__title');
                    if (title) title.hidden = true;
                    wrap.hidden = false;
                    if (closeBtn) closeBtn.hidden = false;
                    if (searchToggle) searchToggle.setAttribute('aria-pressed', 'true');
                }
            } else {
                if (!wrap?.hidden) {
                    // Use the X button so _wireSearch's close() runs and cleans up
                    if (closeBtn) closeBtn.click();
                }
            }
        };

        if (active) requestAnimationFrame(() => syncSearchWrap(true));

        btn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            active = !active;
            btn.setAttribute('aria-pressed', String(active));
            try { localStorage.setItem(lsKey, active ? '1' : '0'); } catch {}
            syncSearchWrap(active);
        });
        btn.addEventListener('pointerdown', (e) => e.stopPropagation());
    }

    _wireSearch(root) {
        const title    = root.querySelector('.legend__title');
        const wrap     = root.querySelector('.legend__search-wrap');
        const input    = root.querySelector('.legend__search-input');
        const counter  = root.querySelector('.legend__search-count');
        const toggle   = root.querySelector('.legend__search-toggle');
        const closeBtn = root.querySelector('.legend__search-close');
        const list     = root.querySelector('.legend__list');
        if (!toggle || !input) return;

        let outsideClickFn = null;
        let matchIndex = 0;

        const updateFilter = () => {
            matchIndex = 0;
            const q = input.value.trim().toLowerCase();
            let matches = 0;
            list.querySelectorAll('.legend__item').forEach(item => {
                const label = (item.querySelector('.legend__label')?.textContent || '').toLowerCase();
                const visible = !q || label.includes(q);
                item.hidden = !visible;
                if (visible) matches++;
            });
            counter.textContent = q ? String(matches) : '';
            counter.hidden = !q;
        };

        const scrollToNextMatch = () => {
            const items = [...list.querySelectorAll('.legend__item:not([hidden])')];
            if (!items.length) return;
            if (matchIndex >= items.length) matchIndex = 0;
            const target = items[matchIndex];
            target.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            target.classList.add('legend__item--highlight');
            target.addEventListener('animationend', () =>
                target.classList.remove('legend__item--highlight'), { once: true });
            matchIndex = (matchIndex + 1) % items.length;
        };

        const turnOffFilter = () => {
            const fb = root.querySelector('.legend__filter-toggle');
            if (!fb || fb.getAttribute('aria-pressed') !== 'true') return;
            fb.setAttribute('aria-pressed', 'false');
            const k = fb.dataset.legendFilterKey;
            if (k) try { localStorage.setItem(k, '0'); } catch {}
        };

        const close = () => {
            matchIndex = 0;
            input.value = '';
            title.hidden = false;
            wrap.hidden = true;
            if (closeBtn) closeBtn.hidden = true;
            toggle.setAttribute('aria-pressed', 'false');
            list.querySelectorAll('.legend__item').forEach(item => { item.hidden = false; });
            counter.textContent = '';
            counter.hidden = true;
            if (outsideClickFn) {
                document.removeEventListener('click', outsideClickFn);
                outsideClickFn = null;
            }
        };

        const open = () => {
            title.hidden = true;
            wrap.hidden = false;
            if (closeBtn) closeBtn.hidden = false;
            toggle.setAttribute('aria-pressed', 'true');
            input.focus();
            updateFilter();
            // Outside click closes only when the filter toggle is NOT keeping the search open
            outsideClickFn = (e) => {
                if (!root.contains(e.target) && !this._isFilterActive(root)) close();
            };
            document.addEventListener('click', outsideClickFn);
        };

        toggle.addEventListener('click', (e) => {
            e.stopPropagation();
            if (wrap.hidden) {
                open();
            } else if (input.value.trim()) {
                scrollToNextMatch();
            } else {
                close();
            }
        });
        toggle.addEventListener('pointerdown', (e) => e.stopPropagation());

        if (closeBtn) {
            closeBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                turnOffFilter();
                close();
            });
            closeBtn.addEventListener('pointerdown', (e) => e.stopPropagation());
        }

        input.addEventListener('input', updateFilter);
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') { e.stopPropagation(); turnOffFilter(); close(); }
            if (e.key === 'Enter')  { e.preventDefault();  scrollToNextMatch(); }
        });
    }

    _enableDrag(root, opts = {}) {
        if (!root || this._dragAttached) return;
        this._dragAttached = true;
        makeLegendDraggable(root, opts);
    }

    _enableResize(root, storageKey = 'legend-size-v1') {
        if (!root || this._resizeAttached) return;
        this._resizeAttached = true;

        const list = root.querySelector('.legend__list');
        if (!list) return;

        try {
            const saved = JSON.parse(localStorage.getItem(storageKey) || 'null');
            if (saved?.width)         root.style.width     = saved.width;
            if (saved?.maxListHeight) list.style.maxHeight  = saved.maxListHeight;
        } catch {}

        const corners = [
            { cls: 'nw', wSign: -1, hSign: -1 },
            { cls: 'ne', wSign:  1, hSign: -1 },
            { cls: 'sw', wSign: -1, hSign:  1 },
            { cls: 'se', wSign:  1, hSign:  1 },
        ];

        corners.forEach(({ cls, wSign, hSign }) => {
            const handle = root.querySelector(`.legend__resize-handle--${cls}`);
            if (!handle) return;

            handle.addEventListener('pointerdown', (e) => {
                if (e.button !== 0) return;
                e.preventDefault();
                e.stopPropagation();
                try { handle.setPointerCapture(e.pointerId); } catch {}

                const r      = root.getBoundingClientRect();
                const startX = e.clientX;
                const startY = e.clientY;
                const startW = r.width;
                const startH = list.getBoundingClientRect().height;

                const onMove = (me) => {
                    const rawDx = me.clientX - startX;
                    const rawDy = me.clientY - startY;
                    root.style.width     = `${Math.max(200, Math.min(600, startW + rawDx * wSign))}px`;
                    list.style.maxHeight = `${Math.max(120, startH + rawDy * hSign)}px`;
                };

                const onUp = () => {
                    handle.removeEventListener('pointermove', onMove);
                    handle.removeEventListener('pointerup',   onUp);
                    try {
                        localStorage.setItem(storageKey, JSON.stringify({
                            width:         root.style.width,
                            maxListHeight: list.style.maxHeight,
                        }));
                    } catch {}
                };

                handle.addEventListener('pointermove', onMove);
                handle.addEventListener('pointerup',   onUp);
            });
        });
    }
}
