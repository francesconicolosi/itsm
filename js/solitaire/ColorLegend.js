import * as d3 from 'd3';
import { normalizeWs } from '../shared/utils.js';
import { LegendBase } from '../shared/LegendBase.js';
import {
    ROLE_FIELD_WITH_MAPPING,
    COMPANY_FIELD,
    LOCATION_FIELD,
    BUSINESS_FUNCTION_FIELD,
    NEUTRAL_COLOR,
    TEAM_MEMBER_LEGENDA_LABEL,
} from './constants.js';
import {
    filterOrganizationByStreams,
    getAllowedStreamsSet,
} from './orgUtils.js';
import {
    makeKeyColorScale,
    getLegendTitleFor,
    computeKeysAndCountsFromVisibleOrg,
} from './legend.js';

const UNKNOWN_MATCHER = /^(unknown|n\/?a|not\s*(set|available)|-|—|none)$/i;

export class ColorLegend extends LegendBase {
    constructor(app) {
        super();
        this.app = app;
        this.colorBy = ROLE_FIELD_WITH_MAPPING;
        this.colorScale = null;
        this.colorKeyMappings = new Map();
        this.seenLegendClickKeys = new Set();
        this.lastLegendClickAt = 0;
    }

    isUnknownKey(v) {
        const s = (v ?? '').toString().trim();
        return !s || UNKNOWN_MATCHER.test(s);
    }

    getCardFill(g) {
        if (typeof this.colorScale !== 'function') return NEUTRAL_COLOR;

        let colorKey;
        if (this.colorBy === ROLE_FIELD_WITH_MAPPING) {
            colorKey = normalizeWs(g.attr('data-role')) || TEAM_MEMBER_LEGENDA_LABEL;
        } else if (this.colorBy === COMPANY_FIELD) {
            colorKey = (g.attr('data-company') || 'Unknown');
        } else if (this.colorBy === BUSINESS_FUNCTION_FIELD) {
            colorKey = (g.attr('data-function') || 'Unknown');
        } else {
            colorKey = (g.attr('data-location') || 'Unknown');
        }

        const finalColor = this.colorScale(colorKey);
        return (typeof finalColor === 'string' && finalColor) ? finalColor : NEUTRAL_COLOR;
    }

    renderAll({ title, fieldName = LOCATION_FIELD, keys, counts, topKey, colorOf, maxVisible = 11 }) {
        const { app } = this;
        const root = this._getOrCreateRoot('legend-root');
        this._buildShell(root, `${keys.length} ${title}`);
        const collapsedKey = `legend-collapsed-v1::${String(this.colorBy || fieldName || 'legend').toLowerCase()}`;
        const { list } = this._wireCollapse(root, collapsedKey);

        keys.forEach((key) => {
            const item = document.createElement('div');
            item.className = 'legend__item';
            item.setAttribute('data-value', key);
            item.setAttribute('data-field', this.colorBy);

            const disabled = this.isUnknownKey(key);
            if (disabled) {
                item.classList.add('legend__item--disabled');
                item.setAttribute('aria-disabled', 'true');
            } else {
                item.setAttribute('role', 'button');
                item.setAttribute('tabindex', '0');
                item.setAttribute('aria-label', `Filter by ${key}`);
            }

            const sw = document.createElement('span');
            sw.className = 'legend__swatch';
            const color = colorOf.colorOf(key);
            sw.style.backgroundColor = color;
            if ((color || '').toLowerCase() === '#ffffff' || color === 'white' || color === NEUTRAL_COLOR) {
                sw.classList.add('legend__swatch--white');
            }

            const label = document.createElement('span');
            label.className = 'legend__label';
            label.textContent = key;

            const count = document.createElement('span');
            count.className = 'legend__count';
            count.textContent = counts.get(key) ?? '';

            item.append(sw, label, count);
            list.appendChild(item);
        });

        const isUnknownKey = (el) => {
            const v = (el.getAttribute('data-value') || '').trim();
            return this.isUnknownKey(v);
        };

        const activate = (el) => {
            const value = el.getAttribute('data-value') ?? '';
            const field = (el.getAttribute('data-field') || '').trim();
            const missing = isUnknownKey(el);

            const normalizedField = normalizeWs(field).toLowerCase();

            // Legend always triggers an exact search using the canonical field:"value" syntax.
            // Unknown entries use field:"Unknown" which _parseQuery resolves to a missing-value search.
            const displayValue = missing ? 'Unknown' : value;
            const queryStr = `${normalizedField}:"${displayValue}"`;

            let noZoom = false;
            if (!missing) {
                const now = Date.now();
                const elapsed = now - (this.lastLegendClickAt || 0);
                if (elapsed > 1000) {
                    noZoom = true;
                }
                this.lastLegendClickAt = now;
            }

            app.search.search(queryStr, { noZoom });
        };

        this._wireListEvents(list, (el) => activate(el));
        this._wireSearch(root);
        this._wireFilterToggle(root, 'legend-filter-v1');

        list.style.setProperty('--legend-row', '24px');
        list.style.maxHeight = `calc(${maxVisible} * var(--legend-row))`;
        this.enableDrag({ handleSelector: '.legend__header' });
    }

    recolor(field) {
        this.colorBy = field;

        const allowedStreams = getAllowedStreamsSet?.() ?? null;
        const orgForLegend = filterOrganizationByStreams(this.app.visibleOrg, allowedStreams);

        const fieldName = this.colorBy;
        const { keys, counts, topKey } = computeKeysAndCountsFromVisibleOrg(orgForLegend, fieldName);

        this.colorScale = makeKeyColorScale(keys, topKey);
        this.renderAll({
            title: getLegendTitleFor(fieldName),
            keys,
            counts,
            topKey,
            colorOf: this.colorScale,
            maxVisible: 11
        });

        const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
        d3.selectAll('g[data-key^="card::"]').each((d, i, nodes) => {
            const g = d3.select(nodes[i]);
            const fill = this.getCardFill(g) || NEUTRAL_COLOR;
            const isNeutral = fill === NEUTRAL_COLOR;
            const effectiveFill = (isNeutral && isDark) ? '#1c1c1e' : fill;
            const rect = g.select('rect.profile-box');
            rect.classed('profile-box--neutral', isNeutral);
            rect.transition().duration(200).attr('fill', effectiveFill);
            if ((fill || '').toLowerCase() === '#ffffff' || fill === 'white') {
                rect.attr('stroke', '#b8b8b8').attr('stroke-width', 1);
            } else {
                rect.attr('stroke', null).attr('stroke-width', null);
            }
        });
    }

    setMode(mode) {
        const roleEl = document.getElementById('toggle-color-role');
        const compEl = document.getElementById('toggle-color-company');
        const locEl  = document.getElementById('toggle-color-location');
        const bfEl   = document.getElementById('toggle-color-function');

        if (!roleEl || !compEl || !locEl) return;

        if (mode === ROLE_FIELD_WITH_MAPPING) {
            roleEl.checked = true;
            compEl.checked = false;
            locEl.checked  = false;
            if (bfEl) bfEl.checked = false;
        } else if (mode === COMPANY_FIELD) {
            roleEl.checked = false;
            compEl.checked = true;
            locEl.checked  = false;
            if (bfEl) bfEl.checked = false;
        } else if (mode === LOCATION_FIELD) {
            roleEl.checked = false;
            compEl.checked = false;
            locEl.checked  = true;
            if (bfEl) bfEl.checked = false;
        } else if (mode === BUSINESS_FUNCTION_FIELD) {
            roleEl.checked = false;
            compEl.checked = false;
            locEl.checked  = false;
            if (bfEl) bfEl.checked = true;
        }

        this.recolor(mode);
    }

    enableDrag({ handleSelector = null } = {}) {
        const root = document.getElementById('legend-root');
        this._enableDrag(root, { handleSelector, storageKey: 'legend-pos-v1', cornerAnchor: true });
        this._enableResize(root, 'legend-size-v1', 'legend-pos-v1');
    }
}
