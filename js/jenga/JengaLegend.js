import { LegendBase } from '../shared/LegendBase.js';
import { HYBRIS_SVG } from './CalendarRenderer.js';

// ─── Colors ───────────────────────────────────────────────────────────────────

const CHIP_LIGHT = {
    SERVICE_OP:     { bg: '#dbeafe', border: '#3b82f6' },
    BUSINESS_EVENT: { bg: '#fce7f3', border: '#db2777' },
    MILESTONE:      { bg: '#ede9fe', border: '#7c3aed' },
};
const CHIP_DARK = {
    SERVICE_OP:     { bg: '#172233', border: '#5e8fc7' },
    BUSINESS_EVENT: { bg: '#3a2633', border: '#b66a8b' },
    MILESTONE:      { bg: '#1e1433', border: '#a78bfa' },
};
const PSPW_SWATCH = {
    RED:   { bg: 'rgba(239,68,68,0.15)',  border: '#ef4444', label: 'Emergency deploys only' },
    AMBER: { bg: 'rgba(234,179,8,0.16)',  border: '#eab308', label: 'Allowed applications only' },
    BLUE:  { bg: 'rgba(59,130,246,0.15)', border: '#3b82f6', label: 'Scheduled approved deploys' },
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
        root.querySelector('.legend__search-toggle')?.remove();
        const { list } = this._wireCollapse(root, 'jenga-legend-collapsed-v1');
        this._list = list;
        this._enableDrag(root, {
            handleSelector: '.legend__header',
            storageKey:     'jenga-legend-pos-v1',
            cornerAnchor:   true,
            minTop: () => parseFloat(
                getComputedStyle(document.documentElement).getPropertyValue('--jenga-topbar-height') || '0'
            ) + 8,
        });
        this._enableResize(root, 'jenga-legend-size-v1', 'jenga-legend-pos-v1');
        root.style.display = 'none';
    }

    /** Rebuild legend items based on the events currently visible in the calendar. */
    update(events, onItemClick) {
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
            items.push({ label: 'Service Operations', filterName: 'Service Operations', ...colors.SERVICE_OP });
        if (types.has('BUSINESS_EVENT'))
            items.push({ label: 'Business Event',     filterName: 'Business Events',    icon: '👜' });
        if (hasPeak)
            items.push({ label: 'High Traffic Day',   filterName: 'Show only calendar', icon: '🔥' });
        if (types.has('HYBRIS'))
            items.push({ label: 'Hybris Release',     filterName: 'Hybris Releases',    hybris: true });
        if (types.has('MILESTONE'))
            items.push({ label: 'Milestone',           filterName: 'Milestones',         icon: '🚩' });
        for (const c of ['RED', 'AMBER', 'BLUE']) {
            if (pspwPresent.has(c)) {
                const s = PSPW_SWATCH[c];
                items.push({ label: s.label, filterName: 'Show only calendar', bg: s.bg, border: s.border, pspw: true });
            }
        }

        this._list.innerHTML = '';
        items.forEach(item => {
            const el = document.createElement('div');
            el.className = 'legend__item';
            if (onItemClick) {
                el.style.cursor = 'pointer';
                el.addEventListener('click', () => onItemClick(item.filterName));
            }

            const sw = document.createElement('span');
            if (item.icon) {
                sw.className = 'legend__swatch legend__swatch--icon';
                sw.textContent = item.icon;
            } else if (item.hybris) {
                sw.className = 'legend__swatch legend__swatch--hybris';
                sw.innerHTML = HYBRIS_SVG;
            } else if (item.pspw) {
                sw.className = 'legend__swatch legend__swatch--pspw';
                sw.style.cssText = `border: 2px solid ${item.border}; background: ${item.bg};`;
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
