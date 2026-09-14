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
      <button class="legend__search-clear" type="button" aria-label="Clear search" hidden>&#xD7;</button>
      <span class="legend__search-count" aria-live="polite" hidden></span>
    </div>
    <div class="legend__header-actions">
      <button class="legend__search-toggle" type="button" aria-label="Search in legend" aria-pressed="false">
        <svg width="12" height="12" viewBox="0 0 14 14" fill="none" aria-hidden="true">
          <circle cx="6" cy="6" r="4.5" stroke="currentColor" stroke-width="1.5"></circle>
          <line x1="9.5" y1="9.5" x2="13" y2="13" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"></line>
        </svg>
      </button>
      <button class="legend__filter-toggle" type="button" aria-label="Filter chart by legend search" aria-pressed="false" hidden>
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

        // Open search-wrap if not already open (used on page-load restore and filter-ON).
        // Turning filter OFF does NOT close search — it just re-applies the filter (shows all).
        const syncSearchWrap = (on) => {
            if (!on) return;
            const wrap         = root.querySelector('.legend__search-wrap');
            const searchToggle = root.querySelector('.legend__search-toggle');
            const filterBtnEl  = root.querySelector('.legend__filter-toggle');
            if (wrap?.hidden) {
                // Open without auto-focusing (less jarring on page load)
                const title = root.querySelector('.legend__title');
                if (title) title.hidden = true;
                wrap.hidden = false;
                if (filterBtnEl) filterBtnEl.hidden = false;
                if (searchToggle) searchToggle.setAttribute('aria-pressed', 'true');
            }
        };

        if (active) requestAnimationFrame(() => syncSearchWrap(true));

        btn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            active = !active;
            btn.setAttribute('aria-pressed', String(active));
            try { localStorage.setItem(lsKey, active ? '1' : '0'); } catch {}
            if (active) syncSearchWrap(true);
            // Re-apply filter: when ON hides non-matches, when OFF shows all items again
            const si = root.querySelector('.legend__search-input');
            if (si) si.dispatchEvent(new Event('input'));
        });
        btn.addEventListener('pointerdown', (e) => e.stopPropagation());
    }

    _wireSearch(root) {
        const title    = root.querySelector('.legend__title');
        const wrap     = root.querySelector('.legend__search-wrap');
        const input    = root.querySelector('.legend__search-input');
        const clearBtn = root.querySelector('.legend__search-clear');
        const counter  = root.querySelector('.legend__search-count');
        const toggle   = root.querySelector('.legend__search-toggle');
        const list     = root.querySelector('.legend__list');
        if (!toggle || !input) return;

        let outsideClickFn = null;
        let matchIndex = 0;

        const updateFilter = () => {
            matchIndex = 0;
            const q = input.value.trim().toLowerCase();
            const filterOn = this._isFilterActive(root);
            let matches = 0;
            list.querySelectorAll('.legend__item').forEach(item => {
                const label = (item.querySelector('.legend__label')?.textContent || '').toLowerCase();
                const matched = !q || label.includes(q);
                item.classList.toggle('legend__item--filtered-out', filterOn && !matched);
                item.dataset.matched = (q && matched) ? 'true' : '';
                if (matched) matches++;
            });
            counter.textContent = q ? String(matches) : '';
            counter.hidden = !q;
        };

        const scrollToNextMatch = () => {
            const filterOn = this._isFilterActive(root);
            const items = [...list.querySelectorAll(
                filterOn ? '.legend__item:not(.legend__item--filtered-out)' : '.legend__item[data-matched="true"]'
            )];
            if (!items.length) return;
            if (matchIndex >= items.length) matchIndex = 0;
            const target = items[matchIndex];
            target.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            target.classList.add('legend__item--highlight');
            target.addEventListener('animationend', () =>
                target.classList.remove('legend__item--highlight'), { once: true });
            matchIndex = (matchIndex + 1) % items.length;
        };

        const close = () => {
            matchIndex = 0;
            input.value = '';
            title.hidden = false;
            wrap.hidden = true;
            if (clearBtn) clearBtn.hidden = true;
            const filterBtn = root.querySelector('.legend__filter-toggle');
            if (filterBtn) filterBtn.hidden = true;
            toggle.setAttribute('aria-pressed', 'false');
            list.querySelectorAll('.legend__item').forEach(item => {
                item.classList.remove('legend__item--filtered-out');
                item.dataset.matched = '';
            });
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
            const filterBtn = root.querySelector('.legend__filter-toggle');
            if (filterBtn) {
                filterBtn.classList.add('legend__filter-toggle--entering');
                filterBtn.hidden = false;
                requestAnimationFrame(() => requestAnimationFrame(() =>
                    filterBtn.classList.remove('legend__filter-toggle--entering')
                ));
            }
            toggle.setAttribute('aria-pressed', 'true');
            input.focus();
            updateFilter();
            // Outside click closes only when filter is NOT pinning the search open
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

        if (clearBtn) {
            clearBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                input.value = '';
                clearBtn.hidden = true;
                updateFilter();
                input.focus();
            });
            clearBtn.addEventListener('pointerdown', (e) => e.stopPropagation());
        }

        input.addEventListener('input', () => {
            updateFilter();
            if (clearBtn) clearBtn.hidden = !input.value;
        });
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') { e.stopPropagation(); if (!this._isFilterActive(root)) close(); }
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
