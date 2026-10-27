/* global BRAND */
import * as d3 from 'd3';
import { openJiraPopup, closeJiraPopup } from '../shared/jiraPopup.js';
import { StatusViewBase } from '../shared/StatusViewBase.js';

// ── JengaStatusView ───────────────────────────────────────────────────────────
export class JengaStatusView extends StatusViewBase {
    constructor(app) {
        super();
        this.app = app;
        this._nodes    = null;
        this._cards    = null;
        this._allCards = null;
    }

    // ── Data ──────────────────────────────────────────────────────────────────

    async _ensureData() {
        if (this._cards !== null) return;
        const stripComments = s => s.replace(/^#[^\n]*\n/gm, '');
        const fetchCsv = async (url) => {
            try {
                const text = await fetch(url).then(r => r.ok ? r.text() : '').catch(() => '');
                return text ? d3.csvParse(stripComments(text)) : [];
            } catch { return []; }
        };

        const catalogUrl = (typeof BRAND !== 'undefined' && BRAND.csv?.domino)
            ? BRAND.csv.domino : './service-catalog.csv';
        const currentUrl = (typeof BRAND !== 'undefined' && BRAND.csv?.jiraCardsCurrent)
            ? BRAND.csv.jiraCardsCurrent : './jira-cards-current.csv';
        const allUrl = (typeof BRAND !== 'undefined' && BRAND.csv?.jiraCards)
            ? BRAND.csv.jiraCards : './jira-cards.csv';

        const [catalog, current, all] = await Promise.all([
            fetchCsv(catalogUrl),
            fetchCsv(currentUrl),
            fetchCsv(allUrl),
        ]);

        const DEAD = new Set(['Stopped', 'Decommissioned']);
        this._nodes = catalog
            .filter(d => !DEAD.has(d['Status']) && !d['Decommission Date'])
            .map(d => ({ id: d['Service Name'], ...d }));

        this._cards    = current;
        this._allCards = all;
    }

    // ── Main render ───────────────────────────────────────────────────────────

    async render(containerEl, { tv = false, section = 'full', incidentOffset = 0, incidentLimit = Infinity, columnFill = undefined, zoom = undefined } = {}) {
        await this._ensureData();

        const healthMap = this._buildHealthMap();

        const activeServices = this.app?.search?.activeServices;
        const brandHideTypes = this.app?.brandHideTypes;
        const nodes = (this._nodes || []).filter(n => this._filterNode(n, brandHideTypes, activeServices));

        const { streamData, maxMinorCount } = this._buildStreamData(nodes, healthMap);

        const globalHealth = this._worstHealth(streamData.map(s => s.health));
        const operational  = nodes.filter(n => !healthMap.get(n.id)?.incidents?.length).length;
        const affected     = nodes.length - operational;
        const minorTotal   = nodes.reduce((s, n) => s + (healthMap.get(n.id)?.minorIncidents?.length ?? 0), 0);

        const CLOSED_STATUSES = new Set(['Closed', 'Done', 'Resolved', 'Cancelled', 'Canceled']);
        const majorCards = (this._allCards || []).filter(c =>
            !CLOSED_STATUSES.has(c.Status) &&
            (c.IssueType || '').toLowerCase().includes('incident') &&
            (c.Priority === 'P1/Blocker' || c.Priority === 'P2/Critical')
        );

        const cols = (tv && section !== 'incidents')
            ? Math.min(4, Math.max(2, Math.ceil(streamData.length / 3)))
            : 1;

        const inner = document.createElement('div');
        inner.className = `jenga-status-inner${tv ? ' jenga-status-tv' : ''}`;
        inner.style.setProperty('--status-cols', String(cols));
        if (columnFill) inner.style.columnFill = columnFill;
        if (zoom !== undefined && zoom !== 1) inner.style.zoom = String(zoom);
        const displayCards = isFinite(incidentLimit) || incidentOffset > 0
            ? majorCards.slice(incidentOffset, incidentOffset + incidentLimit)
            : majorCards;
        // TV mode: thermometer is NOT in the column flow (added as absolute overlay below)
        inner.innerHTML =
            this._renderBanner(globalHealth, operational, affected, minorTotal, tv) +
            (section !== 'streams' && displayCards.length ? this._renderMajorIncidents(displayCards, majorCards) : '') +
            (section !== 'incidents' ? streamData.map(s => this._renderStream(s, healthMap, maxMinorCount, tv)).join('') : '') +
            (!tv && section !== 'incidents' ? this._renderStabilityThermometer(false) : '');

        containerEl.innerHTML = '';
        containerEl.appendChild(inner);

        this._bindInteractions(healthMap, inner);

        // TV: overlay thermometer absolutely in the rightmost column's empty space.
        // Keeping it out of the column flow prevents it overflowing to a clipped Col 3.
        if (tv && section !== 'incidents') {
            requestAnimationFrame(() => {
                if (!inner.isConnected) return;
                const innerRect = inner.getBoundingClientRect();
                if (innerRect.width <= 0 || innerRect.height <= 0) return;

                const zoomFactor = parseFloat(inner.style.zoom) || 1;
                const colW       = innerRect.width / cols;
                const rightColX  = innerRect.left + (cols - 1) * colW;

                // Content start = bottom of column-spanning elements (banner, incidents panel)
                let contentBottom = innerRect.top;
                inner.querySelectorAll('.status-global-banner, .status-major-panel').forEach(el => {
                    contentBottom = Math.max(contentBottom, el.getBoundingClientRect().bottom);
                });

                // Advance to the bottom of any streams in the rightmost column
                inner.querySelectorAll('.status-stream').forEach(el => {
                    const r = el.getBoundingClientRect();
                    if (r.left >= rightColX - 20) {
                        contentBottom = Math.max(contentBottom, r.bottom);
                    }
                });

                const availH = innerRect.bottom - contentBottom - 12;
                if (availH < 60) return;

                const tmpDiv = document.createElement('div');
                tmpDiv.innerHTML = this._renderStabilityThermometer(true);
                const thermoEl = tmpDiv.firstElementChild;
                if (!thermoEl) return;

                const topLogical    = Math.round((contentBottom - innerRect.top + 12) / zoomFactor);
                const widthLogical  = Math.round(colW / zoomFactor);
                const heightLogical = Math.round((availH - 8) / zoomFactor);

                thermoEl.style.position   = 'absolute';
                thermoEl.style.top        = `${topLogical}px`;
                thermoEl.style.right      = '0';
                thermoEl.style.width      = `${widthLogical}px`;
                thermoEl.style.height     = `${heightLogical}px`;
                thermoEl.style.marginTop  = '0';
                thermoEl.style.boxSizing  = 'border-box';

                const chartEl = thermoEl.querySelector('.status-thermometer-chart');
                if (chartEl) {
                    const chartH = heightLogical - 30 - 32; // 30 = title, 32 = padding
                    if (chartH > 50) {
                        chartEl.style.height    = `${chartH}px`;
                        chartEl.style.maxHeight = 'none';
                        chartEl.style.minHeight = '0';
                    }
                }

                inner.style.position = 'relative';
                inner.appendChild(thermoEl);
            });
        }
    }

    // ── Interactions ──────────────────────────────────────────────────────────

    _bindInteractions(healthMap, contentEl) {
        this._bindCommonInteractions(healthMap, contentEl);

        contentEl.querySelectorAll('.status-major-svc-tag[data-svc-id]').forEach(tag => {
            tag.style.cursor = 'pointer';
            tag.addEventListener('click', () => window.location.href = `./domino.html?search=${encodeURIComponent(`id:"${tag.dataset.svcId}"`)}`);
            tag.addEventListener('keydown', e => {
                if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); tag.click(); }
            });
        });

        contentEl.querySelectorAll('.status-svc-row').forEach(row => {
            row.style.cursor = 'pointer';
            row.addEventListener('click', e => {
                if (e.target.closest('.status-major-badge, .status-minor-badge')) return;
                const svcId = row.dataset.svcId;
                if (svcId) window.location.href = `./domino.html?search=${encodeURIComponent(`id:"${svcId}"`)}`;
            });
        });

        const mapCards = cards => cards.map(inc => ({
            key: inc.Key, summary: inc.Summary, status: inc.Status,
            priority: (inc.Priority || '').split('/')[0].trim(),
            jiraUrl: inc.JiraUrl,
            impactedServices: (inc.AffectedServices || '').split('||').map(s => s.trim()).filter(Boolean),
        }));

        contentEl.querySelectorAll('.status-major-badge[data-svc-id]').forEach(badge => {
            badge.style.cursor = 'pointer';
            badge.addEventListener('click', e => {
                e.stopPropagation();
                const id = badge.dataset.svcId;
                const entry = healthMap.get(id);
                if (!entry?.incidents?.length) return;
                openJiraPopup({ id }, 'incident', mapCards(entry.incidents), e, {
                    onServiceClick: svcId => { closeJiraPopup(); window.location.href = `./domino.html?search=${encodeURIComponent(`id:"${svcId}"`)}` },
                });
            });
        });

        contentEl.querySelectorAll('.status-minor-badge[data-svc-id]').forEach(badge => {
            badge.style.cursor = 'pointer';
            badge.addEventListener('click', e => {
                e.stopPropagation();
                const id = badge.dataset.svcId;
                const entry = healthMap.get(id);
                if (!entry?.minorIncidents?.length) return;
                openJiraPopup({ id }, 'incident', mapCards(entry.minorIncidents), e, {
                    onServiceClick: svcId => { closeJiraPopup(); window.location.href = `./domino.html?search=${encodeURIComponent(`id:"${svcId}"`)}` },
                });
            });
        });
    }

    invalidate() {
        this._nodes    = null;
        this._cards    = null;
        this._allCards = null;
    }
}
