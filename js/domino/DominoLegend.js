import { LegendBase } from '../shared/LegendBase.js';

export class DominoLegend extends LegendBase {
    constructor(app) {
        super();
        this.app = app;
    }

    render(colorScale) {
        const root = this._getOrCreateRoot('legend');
        this._buildShell(root, 'Service Type');
        const { list } = this._wireCollapse(root, 'domino-legend-collapsed-v1');

        colorScale.domain().forEach(type => {
            const color = colorScale(type);
            const item = document.createElement('div');
            item.className = 'legend__item';
            item.setAttribute('data-value', type);
            item.setAttribute('role', 'button');
            item.setAttribute('tabindex', '0');
            item.setAttribute('aria-label', `Filter by ${type}`);

            const sw = document.createElement('span');
            sw.className = 'legend__swatch';
            sw.style.backgroundColor = color;

            const label = document.createElement('span');
            label.className = 'legend__label';
            label.textContent = type;

            item.append(sw, label);
            list.appendChild(item);
        });

        this._wireListEvents(list, (el) => {
            this.app.search.handleQuery(`Type:"${el.getAttribute('data-value')}"`, false);
        });
        this._wireSearch(root);
        this._enableDrag(root, { handleSelector: '.legend__header', storageKey: 'domino-legend-pos-v1', cornerAnchor: true });
        this._enableResize(root, 'domino-legend-size-v1');
    }

    reset() {
        const el = document.getElementById('legend');
        if (el) el.innerHTML = '';
    }
}
