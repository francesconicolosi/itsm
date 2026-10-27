/* global BRAND */
import * as d3 from 'd3';
import { toFriendlyName } from '../shared/statusUtils.js';
import { openJiraPopup, closeJiraPopup } from '../shared/jiraPopup.js';
import { StatusViewBase, ALL_RANK, HEALTH_ORDER } from '../shared/StatusViewBase.js';

// ── StatusView ────────────────────────────────────────────────────────────────
export class StatusView extends StatusViewBase {
    constructor(app) {
        super();
        this.app = app;
        this.statusViewEl = null;
        this.btnStatus    = null;
        this.btnList      = null;
        this.btnGraph     = null;
        this.btnExport    = null;
        this.mapEl        = null;
        this.legendEl     = null;
        this.listViewEl   = null;
        this._cards    = null;
        this._allCards = null;
        this._boundSearchHandler = null;
        this._boundGoHandler     = null;
        this._boundClearHandler  = null;
        this._preNavSearchTerm   = undefined;
    }

    initDOM() {
        this.statusViewEl = document.getElementById('status-view');
        this.btnStatus    = document.getElementById('view-status');
        this.btnList      = document.getElementById('view-list');
        this.btnGraph     = document.getElementById('view-graph');
        this.btnExport    = document.getElementById('act-export-excel');
        this.mapEl        = document.getElementById('map');
        this.legendEl     = document.getElementById('legend');
        this.listViewEl   = document.getElementById('list-view');
        this.btnStatus?.addEventListener('click', () => this.toStatusView());
    }

    _setActiveButton(active) {
        [this.btnList, this.btnGraph, this.btnStatus].forEach(b => {
            b?.classList.toggle('is-active', b === active);
        });
    }

    // ── View transitions ──────────────────────────────────────────────────────

    async toStatusView() {
        if (this.mapEl)      this.mapEl.style.display = 'none';
        if (this.legendEl)   this.legendEl.style.display = 'none';
        if (this.listViewEl) this.listViewEl.style.display = 'none';
        if (this.statusViewEl) this.statusViewEl.style.display = 'block';

        this._setActiveButton(this.btnStatus);

        const surl = new URL(window.location.href);
        surl.searchParams.set('view', 'status');
        surl.searchParams.delete('listView');
        surl.searchParams.delete('sort');
        window.history.replaceState({}, '', surl.toString());

        await this._ensureData();

        if (!this.statusViewEl.querySelector('.status-content')) {
            this.statusViewEl.innerHTML =
                '<div class="status-view-inner"><div class="status-content"></div></div>';
        }

        if (this._preNavSearchTerm !== undefined) {
            const restore = this._preNavSearchTerm;
            this._preNavSearchTerm = undefined;
            this.app.search.handleQuery(restore, false);
            const inp = document.getElementById('drawer-search-input');
            if (inp) inp.value = restore;
        }

        if (!this._boundSearchHandler) {
            this._boundSearchHandler = (e) => {
                if (e.type === 'keydown') {
                    if (e.key !== 'Enter') return;
                    if (this.app.autocomplete?.hasPendingSelection()) return;
                }
                this._renderContent();
            };
            this._boundGoHandler = this._boundSearchHandler;
            this._boundClearHandler = () => this._renderContent();
        }
        const mainInput = document.getElementById('drawer-search-input');
        mainInput?.addEventListener('keydown', this._boundSearchHandler);
        document.getElementById('drawer-search-go')?.addEventListener('click', this._boundGoHandler);
        document.getElementById('act-clear')?.addEventListener('click', this._boundClearHandler);

        this._renderContent();
    }

    toOtherView() {
        if (this.statusViewEl) this.statusViewEl.style.display = 'none';
        const mainInput = document.getElementById('drawer-search-input');
        mainInput?.removeEventListener('keydown', this._boundSearchHandler);
        document.getElementById('drawer-search-go')?.removeEventListener('click', this._boundGoHandler);
        document.getElementById('act-clear')?.removeEventListener('click', this._boundClearHandler);
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
        const currentUrl = (typeof BRAND !== 'undefined' && BRAND.csv?.jiraCardsCurrent)
            ? BRAND.csv.jiraCardsCurrent : './jira-cards-current.csv';
        const allUrl = (typeof BRAND !== 'undefined' && BRAND.csv?.jiraCards)
            ? BRAND.csv.jiraCards : './jira-cards.csv';
        const [current, all] = await Promise.all([fetchCsv(currentUrl), fetchCsv(allUrl)]);
        this._cards    = current;
        this._allCards = all;
    }

    // ── Render ────────────────────────────────────────────────────────────────

    _renderContent() {
        const contentEl = this.statusViewEl?.querySelector('.status-content');
        if (!contentEl) return;

        const { store } = this.app;
        const healthMap = this._buildHealthMap();

        const DEAD = new Set(['Stopped', 'Decommissioned']);
        const brandHideTypes = this.app?.brandHideTypes;
        const nodes = store.nodes.filter(n => {
            if (DEAD.has(n.Status) || n['Decommission Date']) return false;
            return this._filterNode(n, brandHideTypes, null);
        });

        const { streamData, maxMinorCount } = this._buildStreamData(nodes, healthMap);

        const { search } = this.app;
        const rawText    = (document.getElementById('drawer-search-input')?.value ?? '').toLowerCase().trim();
        const matchedIds = (search.searchTerm && search.currentSearchedNodes)
            ? search.currentSearchedNodes : null;
        const hasFilter  = Boolean(matchedIds) || Boolean(rawText);
        const filtered   = hasFilter ? this._filterStreamData(streamData, matchedIds, rawText) : streamData;

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

        contentEl.innerHTML =
            this._renderBanner(globalHealth, operational, affected, minorTotal) +
            (majorCards.length ? this._renderMajorIncidents(majorCards) : '') +
            (hasFilter && !filtered.length ? '<p class="status-no-results">No matching services or areas found.</p>' : '') +
            filtered.map(s => this._renderStream(s, healthMap, maxMinorCount)).join('');

        this._bindInteractions(healthMap, contentEl);
    }

    _filterStreamData(streamData, matchedIds, rawText) {
        return streamData
            .map(stream => {
                if (rawText && stream.name.toLowerCase().includes(rawText)) return stream;
                const themes = stream.themes
                    .map(theme => {
                        if (rawText && theme.name.toLowerCase().includes(rawText)) return theme;
                        const clusters    = theme.clusters.filter(c =>
                            (rawText && c.label.toLowerCase().includes(rawText)) ||
                            c.services.some(n =>
                                matchedIds ? matchedIds.has(n.id) : (rawText && this._nodeMatchesQuery(n, rawText))
                            )
                        );
                        const individuals = theme.individuals.filter(n =>
                            matchedIds ? matchedIds.has(n.id) : (rawText && this._nodeMatchesQuery(n, rawText))
                        );
                        if (!clusters.length && !individuals.length) return null;
                        return { ...theme, clusters, individuals };
                    })
                    .filter(Boolean);
                if (!themes.length) return null;
                return { ...stream, themes };
            })
            .filter(Boolean);
    }

    _nodeMatchesQuery(node, q) {
        return node.id.toLowerCase().includes(q) ||
               toFriendlyName(node.id).toLowerCase().includes(q) ||
               (node.Description || '').toLowerCase().includes(q);
    }

    // ── Overrides ─────────────────────────────────────────────────────────────

    _renderServiceRow(node, entry) {
        const health     = entry?.health ?? 'operational';
        const hasInc     = Boolean(entry?.incidents?.length);
        const hasMinor   = Boolean(entry?.minorIncidents?.length);
        const friendly   = toFriendlyName(node.id);
        const isSlug     = /[-_]/.test(node.id) || (!/\s/.test(node.id) && node.id === node.id.toLowerCase());
        const code       = isSlug ? node.id : null;
        const badge      = hasInc   ? `<span class="status-badge">${entry.incidents.length} open</span>` : '';
        const minorBadge = hasMinor ? `<span class="status-minor-badge" data-svc-id="${this._esc(node.id)}">▲ ${entry.minorIncidents.length}</span>` : '';
        const expand     = hasInc   ? `<span class="status-expand-icon" aria-hidden="true">›</span>` : '';

        return `
            <div class="status-svc-row${hasInc ? ' has-incidents' : ''}"
                 data-svc-id="${this._esc(node.id)}" data-health="${health}">
                <span class="status-dot status-dot--${health}"></span>
                <span class="status-svc-names">
                    <span class="status-svc-primary">${this._esc(friendly)}</span>
                    ${code ? `<span class="status-svc-code">${this._esc(code)}</span>` : ''}
                </span>
                ${badge}${minorBadge}${expand}
            </div>`;
    }

    // ── Interactions ──────────────────────────────────────────────────────────

    _bindInteractions(healthMap, contentEl) {
        this._bindCommonInteractions(healthMap, contentEl);

        contentEl.querySelectorAll('.status-major-svc-tag[data-svc-id]').forEach(tag => {
            tag.style.cursor = 'pointer';
            tag.addEventListener('click', () => this._navigateToService(tag.dataset.svcId));
            tag.addEventListener('keydown', e => {
                if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this._navigateToService(tag.dataset.svcId); }
            });
        });

        contentEl.querySelectorAll('.status-minor-badge[data-svc-id]').forEach(badge => {
            badge.style.cursor = 'pointer';
            badge.addEventListener('click', e => {
                e.stopPropagation();
                const id = badge.dataset.svcId;
                if (!id) return;
                const entry = healthMap.get(id);
                if (!entry?.minorIncidents?.length) return;
                const cards = entry.minorIncidents.map(inc => ({
                    key:              inc.Key,
                    summary:          inc.Summary,
                    status:           inc.Status,
                    priority:         (inc.Priority || '').split('/')[0].trim(),
                    jiraUrl:          inc.JiraUrl,
                    impactedServices: (inc.AffectedServices || '').split('||').map(s => s.trim()).filter(Boolean),
                }));
                openJiraPopup({ id }, 'incident', cards, e, {
                    onServiceClick: svcId => {
                        closeJiraPopup();
                        this._navigateToService(svcId);
                    },
                });
            });
        });

        contentEl.querySelectorAll('.status-svc-row').forEach(row => {
            row.addEventListener('click', e => {
                if (e.target.closest('.status-minor-badge')) return;
                if (e.target.closest('.status-expand-icon') && row.classList.contains('has-incidents')) {
                    this._toggleIncidents(row, healthMap);
                } else {
                    this._navigateToService(row.dataset.svcId);
                }
            });
        });
    }

    _onSearchUpdate() {
        if (this.statusViewEl?.style.display === 'block') this._renderContent();
    }

    _navigateToService(serviceId) {
        if (!serviceId) return;
        this._preNavSearchTerm = this.app.search.searchTerm ?? '';
        const query = `id:"${serviceId}"`;
        const input = document.getElementById('drawer-search-input');
        if (input) input.value = query;
        this.app.search.handleQuery(query, false);
        this.app.listView.toGraphView();
    }

    _toggleIncidents(row, healthMap) {
        const existing = row.nextElementSibling;
        if (existing?.classList.contains('status-incidents-pane')) {
            existing.remove();
            row.classList.remove('is-open');
            const icon = row.querySelector('.status-expand-icon');
            if (icon) icon.style.transform = '';
            return;
        }
        const svcId = row.dataset.svcId;
        const entry = healthMap.get(svcId);
        if (!entry?.incidents?.length) return;

        const pane = document.createElement('div');
        pane.className = 'status-incidents-pane';
        pane.innerHTML = entry.incidents
            .slice()
            .sort((a, b) => (ALL_RANK[a.Priority] ?? 99) - (ALL_RANK[b.Priority] ?? 99))
            .map(inc => {
                const p       = (inc.Priority || '').toLowerCase().replace('/', '').split(' ')[0];
                const summary = this._esc(this._cleanSummary(inc.Summary || ''));
                const url     = inc.JiraUrl || '';
                return `
                    <div class="status-inc-row">
                        ${url
                            ? `<a class="status-inc-key" href="${url}" target="_blank" rel="noopener">${this._esc(inc.Key)}</a>`
                            : `<span class="status-inc-key">${this._esc(inc.Key)}</span>`}
                        <div class="status-inc-meta">
                            <span class="status-inc-summary">${summary}</span>
                            <span class="domino-jira-card__priority--${p} status-inc-prio">${this._esc(inc.Priority)}</span>
                            <span class="status-inc-status">${this._esc(inc.Status)}</span>
                        </div>
                    </div>`;
            })
            .join('');

        row.insertAdjacentElement('afterend', pane);
        row.classList.add('is-open');
        const icon = row.querySelector('.status-expand-icon');
        if (icon) icon.style.transform = 'rotate(90deg)';
    }
}
