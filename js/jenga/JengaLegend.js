import { LegendBase } from '../shared/LegendBase.js';

// ─── Colors ───────────────────────────────────────────────────────────────────

const CHIP_LIGHT = {
    SERVICE_OP:     { bg: '#dbeafe', border: '#3b82f6' },
    BUSINESS_EVENT: { bg: '#fce7f3', border: '#db2777' },
    HYBRIS:         { bg: '#e0fdf4', border: '#14b8a6' },
    MILESTONE:      { bg: '#ede9fe', border: '#7c3aed' },
};
const CHIP_DARK = {
    SERVICE_OP:     { bg: '#172233', border: '#5e8fc7' },
    BUSINESS_EVENT: { bg: '#3a2633', border: '#b66a8b' },
    HYBRIS:         { bg: '#17312f', border: '#5bb8a9' },
    MILESTONE:      { bg: '#1e1433', border: '#a78bfa' },
};
const PSPW_SWATCH = {
    RED:   { bg: 'rgba(239,68,68,0.15)',  border: '#ef4444', label: 'Peak — Red' },
    AMBER: { bg: 'rgba(234,179,8,0.16)',  border: '#eab308', label: 'Peak — Amber' },
    BLUE:  { bg: 'rgba(59,130,246,0.15)', border: '#3b82f6', label: 'Peak — Blue' },
};

function _isDark() {
    return document.documentElement.getAttribute('data-theme') === 'dark';
}

export class JengaLegend extends LegendBase {
    constructor() {
        super();
        this._root = null;
        this._list = null;
    }

    render() {
        const root = this._getOrCreateRoot('jenga-legend-root');
        this._root = root;
        this._buildShell(root, 'Legend');
        const { list } = this._wireCollapse(root, 'jenga-legend-collapsed-v1');
        this._list = list;
        this._enableDrag(root, {
            handleSelector: '.legend__header',
            storageKey:     'jenga-legend-pos-v1',
            cornerAnchor:   true,
        });
        root.style.display = 'none';
    }

    /** Rebuild legend items based on the events currently visible in the calendar. */
    update(events) {
        if (!this._list || !this._root) return;

        const colors = _isDark() ? CHIP_DARK : CHIP_LIGHT;
        const types  = new Set(events.map(e => e.type));

        const hasPeak = events.some(
            e => e.type === 'BUSINESS_EVENT' && Array.isArray(e.peakDates) && e.peakDates.length > 0
        );

        const pspwPresent = new Set(
            events
                .filter(e => e.type === 'PEAK_SEASON_PROTECTION_WINDOW')
                .map(e => e.protectionColor)
                .filter(Boolean)
        );

        const items = [];
        if (types.has('SERVICE_OP'))
            items.push({ label: 'Service Operations', ...colors.SERVICE_OP });
        if (types.has('BUSINESS_EVENT'))
            items.push({ label: 'Business Event',     ...colors.BUSINESS_EVENT });
        if (hasPeak)
            items.push({ label: 'High Traffic Day', flame: true });
        if (types.has('HYBRIS'))
            items.push({ label: 'Hybris Release',     ...colors.HYBRIS });
        if (types.has('MILESTONE'))
            items.push({ label: 'Milestone',           ...colors.MILESTONE });
        for (const c of ['RED', 'AMBER', 'BLUE']) {
            if (pspwPresent.has(c)) {
                const s = PSPW_SWATCH[c];
                items.push({ label: s.label, bg: s.bg, border: s.border });
            }
        }

        this._list.innerHTML = '';
        items.forEach(item => {
            const el = document.createElement('div');
            el.className = 'legend__item';

            const sw = document.createElement('span');
            if (item.flame) {
                sw.className = 'legend__swatch legend__swatch--icon';
                sw.textContent = '🔥';
            } else {
                sw.className = 'legend__swatch legend__swatch--chip';
                sw.style.cssText = `border-left: 3px solid ${item.border}; background: ${item.bg};`;
            }

            const lbl = document.createElement('span');
            lbl.className = 'legend__label';
            lbl.textContent = item.label;

            el.append(sw, lbl);
            this._list.appendChild(el);
        });

        this._root.style.display = items.length ? '' : 'none';
    }

    hide() {
        if (this._root) this._root.style.display = 'none';
    }
}
