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
      <input class="legend__search-input" type="search" autocomplete="off" spellcheck="false" aria-label="Search legend" />
      <span class="legend__search-count" aria-live="polite" hidden></span>
    </div>
    <div class="legend__header-actions">
      <button class="legend__search-toggle" type="button" aria-label="Search in legend" aria-pressed="false">
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
          <circle cx="6" cy="6" r="4.5" stroke="currentColor" stroke-width="1.5"></circle>
          <line x1="9.5" y1="9.5" x2="13" y2="13" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"></line>
        </svg>
      </button>
      <button class="legend__search-close" type="button" aria-label="Close search" hidden>
        <svg width="14" height="14" viewBox="0 0 14 14" fill="none" aria-hidden="true">
          <line x1="2" y1="2" x2="12" y2="12" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"></line>
          <line x1="12" y1="2" x2="2" y2="12" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"></line>
        </svg>
      </button>
      <button class="legend__collapse" type="button" aria-label="Toggle legend" aria-expanded="true">
        <span class="chevron" aria-hidden="true"></span>
      </button>
    </div>
  </div>
  <div class="legend__list" aria-label="Legend list"></div>
  <div class="legend__resize-handle" aria-hidden="true"></div>
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
            outsideClickFn = (e) => { if (!root.contains(e.target)) close(); };
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
            closeBtn.addEventListener('click', (e) => { e.stopPropagation(); close(); });
            closeBtn.addEventListener('pointerdown', (e) => e.stopPropagation());
        }

        input.addEventListener('input', updateFilter);
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') { e.stopPropagation(); close(); }
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

        const handle = root.querySelector('.legend__resize-handle');
        const list   = root.querySelector('.legend__list');
        if (!handle || !list) return;

        const rightAnchored = root.classList.contains('legend--anchor-br')
                           || root.classList.contains('legend--anchor-tr');
        handle.style.cursor = rightAnchored ? 'nesw-resize' : 'nwse-resize';

        try {
            const saved = JSON.parse(localStorage.getItem(storageKey) || 'null');
            if (saved?.width)         root.style.width        = saved.width;
            if (saved?.maxListHeight) list.style.maxHeight     = saved.maxListHeight;
        } catch {}

        handle.addEventListener('pointerdown', (e) => {
            if (e.button !== 0) return;
            e.preventDefault();
            e.stopPropagation();
            try { handle.setPointerCapture(e.pointerId); } catch {}

            const startX = e.clientX;
            const startY = e.clientY;
            const startW = root.getBoundingClientRect().width;
            const startH = list.getBoundingClientRect().height;

            const onMove = (me) => {
                const dx = rightAnchored ? startX - me.clientX : me.clientX - startX;
                const dy = me.clientY - startY;
                root.style.width     = `${Math.max(200, Math.min(600, startW + dx))}px`;
                list.style.maxHeight = `${Math.max(120, startH + dy)}px`;
            };

            const onUp = () => {
                handle.removeEventListener('pointermove', onMove);
                handle.removeEventListener('pointerup', onUp);
                try {
                    localStorage.setItem(storageKey, JSON.stringify({
                        width: root.style.width,
                        maxListHeight: list.style.maxHeight,
                    }));
                } catch {}
            };

            handle.addEventListener('pointermove', onMove);
            handle.addEventListener('pointerup', onUp);
        });
    }
}
