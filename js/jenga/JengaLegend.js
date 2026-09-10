import { LegendBase } from '../shared/LegendBase.js';

// Static legend items for Jenga — no interactivity needed, just reference
const LEGEND_ITEMS = [
    // Deployment dots
    { label: 'PRD Deploy',     type: 'dot',  color: '#ef4444' },
    { label: 'QA Deploy',      type: 'dot',  color: '#eab308' },
    { label: 'QAE Deploy',     type: 'dot',  color: '#f59e0b' },
    { label: 'Staging Deploy', type: 'dot',  color: '#6b7280' },
    // Event type chips
    { label: 'Delivery',       type: 'chip', border: '#22c55e', bg: '#dcfce7' },
    { label: 'Business Event', type: 'chip', border: '#db2777', bg: '#fce7f3' },
    { label: 'Hybris Release', type: 'chip', border: '#a855f7', bg: '#f3e8ff' },
    { label: 'Milestone',      type: 'chip', border: '#7c3aed', bg: '#ede9fe' },
];

export class JengaLegend extends LegendBase {
    constructor() {
        super();
    }

    render() {
        const root = this._getOrCreateRoot('jenga-legend-root');
        this._buildShell(root, 'Legend');
        const { list } = this._wireCollapse(root, 'jenga-legend-collapsed-v1');

        LEGEND_ITEMS.forEach(item => {
            const el = document.createElement('div');
            el.className = 'legend__item';

            const sw = document.createElement('span');
            if (item.type === 'dot') {
                sw.className = 'legend__swatch legend__swatch--dot';
                sw.style.backgroundColor = item.color;
            } else {
                sw.className = 'legend__swatch legend__swatch--chip';
                sw.style.cssText = `border-left: 3px solid ${item.border}; background: ${item.bg};`;
            }

            const label = document.createElement('span');
            label.className = 'legend__label';
            label.textContent = item.label;

            el.append(sw, label);
            list.appendChild(el);
        });

        this._enableDrag(root, {
            handleSelector: '.legend__header',
            storageKey: 'jenga-legend-pos-v1',
            cornerAnchor: true,
        });
    }
}
