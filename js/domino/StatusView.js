/* global BRAND */
import * as d3 from 'd3';
import { toFriendlyName, CLUSTERS, CLUSTER_BY_ID, resolveCluster } from '../shared/statusUtils.js';
import { openJiraPopup, closeJiraPopup } from '../shared/jiraPopup.js';

// ── Exclude 3rd-party and vendor services from the status page ────────────────
const HIDE_TYPES = new Set(['SaaS', 'Service Provider']);


// ── Health helpers ────────────────────────────────────────────────────────────
const ACTIVE_STATUSES = new Set([
    'Sent to Internal Team', 'Pending', 'UAT',
    'Work in progress', 'Work In Progress', 'Scheduled',
    'Investigating', 'In Progress', 'Open',
]);

// Only P1/P2 drive the health dot/pill; P3/P4 → amber minor badge
const P1P2_RANK     = { 'P1/Blocker': 0, 'P2/Critical': 1 };
// Used only for sorting incident lists in the detail pane
const ALL_RANK      = { 'P1/Blocker': 0, 'P2/Critical': 1, 'P3/Major': 2, 'P4/Minor': 3 };
const HEALTH_FROM_RANK = ['outage', 'critical'];
const HEALTH_ORDER  = { outage: 0, critical: 1, degraded: 2, minor: 3, operational: 4 };

const HEALTH_LABEL  = {
    outage: 'Major Outage', critical: 'Critical',
    degraded: 'Degraded',  minor: 'Minor Issues', operational: 'Operational',
};
const GLOBAL_MSG    = {
    outage:      'Major outage in progress',
    critical:    'Critical incidents active',
    degraded:    'Some services degraded',
    minor:       'Minor issues detected',
    operational: 'All systems operational',
};

// ── StatusView class ──────────────────────────────────────────────────────────
export class StatusView {
    constructor(app) {
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

        // Build content container exactly once
        if (!this.statusViewEl.querySelector('.status-content')) {
            this.statusViewEl.innerHTML =
                '<div class="status-view-inner"><div class="status-content"></div></div>';
        }

        // Restore search term when returning from a status-page service navigation click
        if (this._preNavSearchTerm !== undefined) {
            const restore = this._preNavSearchTerm;
            this._preNavSearchTerm = undefined;
            this.app.search.handleQuery(restore, false);
            const inp = document.getElementById('drawer-search-input');
            if (inp) inp.value = restore;
        }

        // Attach Enter-key and go-button listeners while status view is active
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
        this._cards    = current;  // current-month slice → health map (P3/P4 badges)
        this._allCards = all;      // full history → P1/P2 major incidents panel
    }

    _buildHealthMap() {
        // Map<serviceName, { health, rank, incidents[], minorIncidents[] }>
        // P1/P2 sourced from _allCards so health dots reflect active major incidents
        // even when they are absent from jira-cards-current.csv.
        // P3/P4 sourced from _cards (current export) → amber minor badge only.
        const healthMap = new Map();
        const init = svc => {
            if (!healthMap.has(svc)) {
                healthMap.set(svc, { health: 'operational', rank: 99, incidents: [], minorIncidents: [] });
            }
            return healthMap.get(svc);
        };

        for (const card of (this._allCards || [])) {
            if (!ACTIVE_STATUSES.has(card.Status)) continue;
            if (card.Priority !== 'P1/Blocker' && card.Priority !== 'P2/Critical') continue;
            const svcs = (card.AffectedServices || '').split('||').map(s => s.trim()).filter(Boolean);
            for (const svc of svcs) {
                const entry = init(svc);
                const rank = P1P2_RANK[card.Priority];
                if (rank < entry.rank) { entry.rank = rank; entry.health = HEALTH_FROM_RANK[rank]; }
                entry.incidents.push(card);
            }
        }

        for (const card of (this._cards || [])) {
            if (!ACTIVE_STATUSES.has(card.Status)) continue;
            if (card.Priority !== 'P3/Major' && card.Priority !== 'P4/Minor') continue;
            const svcs = (card.AffectedServices || '').split('||').map(s => s.trim()).filter(Boolean);
            for (const svc of svcs) {
                const entry = init(svc);
                entry.minorIncidents.push(card);
                if (entry.health === 'operational') entry.health = 'minor';
            }
        }

        return healthMap;
    }

    // ── Render ────────────────────────────────────────────────────────────────

    _renderContent() {
        const contentEl = this.statusViewEl?.querySelector('.status-content');
        if (!contentEl) return;

        const { store } = this.app;
        const healthMap = this._buildHealthMap();

        // Exclude decommissioned, 3rd-party types, and unmapped (no stream/theme) services
        const DEAD = new Set(['Stopped', 'Decommissioned']);
        const brandHideTypes = this.app?.brandHideTypes;
        const nodes = store.nodes.filter(n => {
            if (DEAD.has(n.Status) || n['Decommission Date']) return false;
            if (HIDE_TYPES.has(n['Type']) || brandHideTypes?.has(n['Type'])) return false;
            return Boolean((n['Stream'] || '').trim()) && Boolean((n['Theme'] || '').trim());
        });

        // Group by stream → theme
        const streamMap = new Map();
        for (const node of nodes) {
            const stream = node['Stream'].trim();
            const theme  = node['Theme'].trim();
            if (!streamMap.has(stream)) streamMap.set(stream, new Map());
            const themeMap = streamMap.get(stream);
            if (!themeMap.has(theme)) themeMap.set(theme, []);
            themeMap.get(theme).push(node);
        }

        const worstHealth = healths => {
            let best = 4;
            for (const h of healths) { const r = HEALTH_ORDER[h] ?? 4; if (r < best) best = r; }
            return Object.keys(HEALTH_ORDER)[best] ?? 'operational';
        };
        const sumMinor = ns => ns.reduce((s, n) => s + (healthMap.get(n.id)?.minorIncidents?.length ?? 0), 0);

        // Build hierarchical data structure
        const streamData = [];
        for (const [streamName, themeMap] of streamMap) {
            const themes = [];
            for (const [themeName, svcs] of themeMap) {
                const clusterGroups = new Map();
                const individuals = [];
                for (const node of svcs) {
                    const cid = resolveCluster(node, CLUSTERS);
                    if (cid) {
                        if (!clusterGroups.has(cid)) clusterGroups.set(cid, []);
                        clusterGroups.get(cid).push(node);
                    } else {
                        individuals.push(node);
                    }
                }

                const clusters = [];
                for (const [cid, clusterSvcs] of clusterGroups) {
                    const cHealth = worstHealth(clusterSvcs.map(n => healthMap.get(n.id)?.health ?? 'operational'));
                    clusters.push({
                        clusterId: cid,
                        label:     CLUSTER_BY_ID[cid]?.label ?? cid,
                        health:    cHealth,
                        services:  clusterSvcs,
                        totalIncidents: clusterSvcs.reduce((s, n) => s + (healthMap.get(n.id)?.incidents?.length ?? 0), 0),
                        totalMinor:     sumMinor(clusterSvcs),
                    });
                }
                clusters.sort((a, b) => HEALTH_ORDER[a.health] - HEALTH_ORDER[b.health] || a.label.localeCompare(b.label));
                individuals.sort((a, b) => {
                    const ha = HEALTH_ORDER[healthMap.get(a.id)?.health ?? 'operational'];
                    const hb = HEALTH_ORDER[healthMap.get(b.id)?.health ?? 'operational'];
                    return ha - hb || a.id.localeCompare(b.id);
                });

                const themeMinor = clusters.reduce((s, c) => s + c.totalMinor, 0) +
                    individuals.reduce((s, n) => s + (healthMap.get(n.id)?.minorIncidents?.length ?? 0), 0);
                themes.push({
                    name:       themeName,
                    health:     worstHealth([...clusters.map(c => c.health), ...individuals.map(n => healthMap.get(n.id)?.health ?? 'operational')]),
                    clusters,
                    individuals,
                    minorCount: themeMinor,
                });
            }
            themes.sort((a, b) => HEALTH_ORDER[a.health] - HEALTH_ORDER[b.health] || a.name.localeCompare(b.name));
            const streamMinor    = themes.reduce((s, t) => s + t.minorCount, 0);
            const streamSvcCount = themes.reduce((s, t) =>
                s + t.clusters.reduce((cs, c) => cs + c.services.length, 0) + t.individuals.length, 0);
            const healthBreakdown = { outage: 0, critical: 0, degraded: 0, minor: 0, clean: 0 };
            for (const theme of themes) {
                for (const cluster of theme.clusters) {
                    if (cluster.health !== 'operational') {
                        healthBreakdown[cluster.health] = (healthBreakdown[cluster.health] || 0) + cluster.services.length;
                    } else if (cluster.totalMinor > 0) {
                        healthBreakdown.minor += cluster.services.length;
                    } else {
                        healthBreakdown.clean += cluster.services.length;
                    }
                }
                for (const node of theme.individuals) {
                    const e = healthMap.get(node.id);
                    const h = e?.health ?? 'operational';
                    if (h !== 'operational') {
                        healthBreakdown[h] = (healthBreakdown[h] || 0) + 1;
                    } else if (e?.minorIncidents?.length) {
                        healthBreakdown.minor++;
                    } else {
                        healthBreakdown.clean++;
                    }
                }
            }
            streamData.push({
                name:           streamName,
                health:         worstHealth(themes.map(t => t.health)),
                themes,
                minorCount:     streamMinor,
                svcCount:       streamSvcCount,
                healthBreakdown,
            });
        }
        streamData.sort((a, b) => b.svcCount - a.svcCount || a.name.localeCompare(b.name));
        const maxMinorCount = streamData.reduce((max, s) => Math.max(max, s.minorCount), 0);

        // Apply shared search filter — uses graph search results + raw text for stream/theme names
        const { search } = this.app;
        const rawText    = (document.getElementById('drawer-search-input')?.value ?? '').toLowerCase().trim();
        const matchedIds = (search.searchTerm && search.currentSearchedNodes)
            ? search.currentSearchedNodes : null;
        const hasFilter  = Boolean(matchedIds) || Boolean(rawText);
        const filtered   = hasFilter ? this._filterStreamData(streamData, matchedIds, rawText) : streamData;

        // Global stats from full unfiltered data
        const globalHealth = worstHealth(streamData.map(s => s.health));
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

    // ── HTML helpers ──────────────────────────────────────────────────────────

    _renderMajorIncidents(cards) {
        const sorted = [...cards].sort((a, b) => {
            const ra = ALL_RANK[a.Priority] ?? 9;
            const rb = ALL_RANK[b.Priority] ?? 9;
            return ra !== rb ? ra - rb : (a.Created ?? '').localeCompare(b.Created ?? '');
        });

        const cardHtml = sorted.map(c => {
            const health    = HEALTH_FROM_RANK[P1P2_RANK[c.Priority]] ?? 'critical';
            const priority  = (c.Priority || '').split('/')[0].trim();
            const services  = (c.AffectedServices || '').split('||').map(s => s.trim()).filter(Boolean);
            const created   = c.Created ? this._formatIncidentDate(c.Created) : '';
            const summary   = this._cleanSummary(c.Summary || '');
            const desc      = (c.Description || '').trim();
            const latest    = (c.LatestComment || '').trim();
            const jiraKey   = c.Key || '';
            const jiraUrl   = c.JiraUrl || '';
            const reporter  = (c.Reporter || '').trim();

            return `
            <div class="status-major-card status-major-card--${health}">
                <div class="status-major-card-header">
                    <span class="status-pill status-pill--${health}">${this._esc(priority)}</span>
                    <span class="status-major-status">${this._esc(c.Status || '')}</span>
                    ${created ? `<span class="status-major-created">${created}</span>` : ''}
                    ${jiraUrl ? `<a class="status-major-jira-link" href="${this._esc(jiraUrl)}" target="_blank" rel="noopener">${this._esc(jiraKey)} ↗</a>` : ''}
                </div>
                <div class="status-major-summary">${this._esc(summary)}</div>
                ${desc ? `<div class="status-major-description">${this._esc(desc)}</div>` : ''}
                ${latest ? `
                <div class="status-major-latest">
                    <span class="status-major-label">Latest update</span>
                    <span class="status-major-latest-text">${this._esc(latest)}</span>
                </div>` : ''}
                <div class="status-major-footer">
                    ${services.length ? `
                    <div class="status-major-services">
                        <span class="status-major-label">Affected</span>
                        ${services.map(s => `<span class="status-major-svc-tag" data-svc-id="${this._esc(s)}" role="button" tabindex="0">${this._esc(s)}</span>`).join('')}
                    </div>` : ''}
                    ${reporter ? `<span class="status-major-reporter">Reported by ${this._esc(reporter)}</span>` : ''}
                </div>
            </div>`;
        }).join('');

        const worstHealth = cards.some(c => c.Priority === 'P1/Blocker') ? 'outage' : 'critical';
        return `
        <div class="status-major-panel">
            <div class="status-major-panel-title">
                <span class="status-dot status-dot--${worstHealth}"></span>
                Active Major Incidents
                <span class="status-major-count">${cards.length}</span>
            </div>
            ${cardHtml}
        </div>`;
    }

    _formatIncidentDate(iso) {
        try {
            const d = new Date(iso);
            if (isNaN(d)) return iso;
            return d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' })
                + ' · ' + d.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });
        } catch { return iso; }
    }

    _renderBanner(health, operational, affected, minorTotal) {
        const minorTag = minorTotal > 0
            ? `<span class="status-global-minor">▲ ${minorTotal} minor (P3/P4)</span>` : '';
        return `
            <div class="status-global-banner status-global-banner--${health}">
                <span class="status-dot status-dot--${health}"></span>
                <span class="status-global-msg">${GLOBAL_MSG[health]}</span>
                <span class="status-global-stats">
                    <span class="status-stat status-stat--op">${operational}</span> operational &nbsp;·&nbsp;
                    <span class="status-stat status-stat--aff">${affected}</span> affected
                    ${minorTag}
                </span>
            </div>`;
    }

    _renderStream(streamData, healthMap, maxMinorCount = 0, tv = false) {
        const { name, health, themes, minorCount = 0, svcCount = 0, healthBreakdown = {} } = streamData;
        const total    = svcCount || 1;
        const affCount = (healthBreakdown.outage || 0) + (healthBreakdown.critical || 0) + (healthBreakdown.degraded || 0);
        const affPct   = Math.round(affCount / total * 100);
        const minorPct = maxMinorCount > 0 ? Math.round(minorCount / maxMinorCount * 45) : 0;
        const displayHealth = health === 'minor' ? 'operational' : health;
        return `
            <div class="status-stream" data-health="${health}">
                <div class="status-stream-header">
                    <span class="status-dot status-dot--${displayHealth}"></span>
                    <a class="status-stream-name status-solitaire-link" href="solitaire.html?search=${encodeURIComponent(`stream:"${name}"`)}" title="View in Solitaire">${name}</a>
                    ${this._pill(displayHealth)}
                    ${minorCount > 0 ? `<span class="status-minor-badge">▲ ${minorCount}</span>` : ''}
                </div>
                <div class="status-health-bar">
                    ${affPct > 0 ? `<div class="status-hb-seg--affected" style="width:${affPct}%"></div>` : ''}
                    ${minorPct > 0 ? `<div class="status-hb-seg--minor" style="width:${minorPct}%"></div>` : ''}
                    <div class="status-hb-seg--clean" style="flex:1"></div>
                </div>
                <div class="status-stream-body">
                    ${themes.map(t => this._renderTheme(t, healthMap, tv)).join('')}
                </div>
            </div>`;
    }

    _renderTheme(themeData, healthMap, tv = false) {
        const { name, health, clusters, individuals, minorCount = 0 } = themeData;
        const displayHealth = health === 'minor' ? 'operational' : health;

        const affectedClusters = clusters.filter(c => c.health !== 'operational' && c.health !== 'minor');
        const minorClusters    = clusters.filter(c => c.health === 'minor' || (c.health === 'operational' && c.totalMinor > 0));
        const cleanClusters    = clusters.filter(c => c.health === 'operational' && !c.totalMinor);

        const affectedIndivs = individuals.filter(n => { const h = healthMap.get(n.id)?.health ?? 'operational'; return h !== 'operational' && h !== 'minor'; });
        const minorIndivs    = individuals.filter(n => { const e = healthMap.get(n.id); const h = e?.health ?? 'operational'; return h === 'minor' || (h === 'operational' && (e?.minorIncidents?.length ?? 0) > 0); });
        const cleanIndivs    = individuals.filter(n => {
            const e = healthMap.get(n.id);
            return (e?.health ?? 'operational') === 'operational' && !(e?.minorIncidents?.length);
        });

        const opCount  = cleanClusters.length + cleanIndivs.length;
        const hasMinor = minorClusters.length + minorIndivs.length > 0;
        const minorItemCount = minorClusters.length + minorIndivs.length;

        const MAX_PREVIEW = 5;
        const previewNames = [
            ...cleanClusters.map(c => c.label),
            ...cleanIndivs.map(n => toFriendlyName(n.id)),
        ];
        const previewStr = previewNames.slice(0, MAX_PREVIEW).join(', ') +
            (previewNames.length > MAX_PREVIEW ? '…' : '');

        const minorPreviewNames = [
            ...minorClusters.map(c => c.label),
            ...minorIndivs.map(n => toFriendlyName(n.id)),
        ];
        const minorPreviewStr = minorPreviewNames.slice(0, MAX_PREVIEW).join(', ') +
            (minorPreviewNames.length > MAX_PREVIEW ? '…' : '');

        return `
            <div class="status-theme" data-health="${health}">
                <div class="status-theme-header">
                    <a class="status-theme-name status-solitaire-link" href="solitaire.html?search=${encodeURIComponent(`theme:"${name}"`)}" title="View in Solitaire">${name}</a>
                    ${this._pill(displayHealth)}
                    ${minorCount > 0 ? `<span class="status-minor-badge">▲ ${minorCount}</span>` : ''}
                </div>
                <div class="status-theme-body">
                    ${affectedClusters.map(c => this._renderCluster(c, healthMap, true, tv)).join('')}
                    ${affectedIndivs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                    ${hasMinor ? (tv
                        ? `<div class="status-op-toggle status-minor-toggle" tabindex="0" role="button" aria-expanded="false" data-count="${minorItemCount}" data-preview="${this._esc(minorPreviewStr)}">
                                <span class="status-dot status-dot--minor"></span>
                                <span class="status-op-label">Show ${minorItemCount} with minor issues: ${this._esc(minorPreviewStr)} ▼</span>
                            </div>
                            <div class="status-op-list" hidden>
                                ${minorClusters.map(c => this._renderCluster(c, healthMap, false, tv)).join('')}
                                ${minorIndivs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                            </div>`
                        : `<div class="status-minor-section">
                                ${minorClusters.map(c => this._renderCluster(c, healthMap, false, tv)).join('')}
                                ${minorIndivs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                            </div>`)
                    : ''}
                    ${opCount > 0 ? `
                        <div class="status-op-toggle" tabindex="0" role="button" aria-expanded="false" data-count="${opCount}" data-preview="${this._esc(previewStr)}">
                            <span class="status-dot status-dot--operational"></span>
                            <span class="status-op-label">Show ${opCount} fully operational: ${this._esc(previewStr)} ▼</span>
                        </div>
                        <div class="status-op-list" hidden>
                            ${cleanClusters.map(c => this._renderCluster(c, healthMap, false, tv)).join('')}
                            ${cleanIndivs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                        </div>` : ''}
                </div>
            </div>`;
    }

    _renderCluster(clusterData, healthMap, defaultExpanded, tv = false) {
        const { clusterId, label, health, services, totalIncidents, totalMinor } = clusterData;
        const incBadge   = totalIncidents > 0 ? `<span class="status-badge">${totalIncidents} open</span>` : '';
        const mainSvc    = totalMinor > 0
            ? services.reduce((best, n) => {
                const a = healthMap.get(n.id)?.minorIncidents?.length ?? 0;
                const b = healthMap.get(best?.id)?.minorIncidents?.length ?? 0;
                return a >= b ? n : best;
            }, services[0])
            : null;
        const minorBadge = totalMinor > 0
            ? `<span class="status-minor-badge" data-svc-id="${this._esc(mainSvc?.id ?? '')}">▲ ${totalMinor}</span>`
            : '';
        const countTxt   = `${services.length} service${services.length !== 1 ? 's' : ''}`;

        // Split services: clean (operational, no minor incidents) vs non-clean
        const nonCleanSvcs = services.filter(n => {
            const e = healthMap.get(n.id);
            return (e?.health ?? 'operational') !== 'operational' || (e?.minorIncidents?.length ?? 0) > 0;
        });
        const cleanSvcs = services.filter(n => {
            const e = healthMap.get(n.id);
            return (e?.health ?? 'operational') === 'operational' && !(e?.minorIncidents?.length);
        });

        const trulyAffSvcs  = nonCleanSvcs.filter(n => { const h = healthMap.get(n.id)?.health ?? 'operational'; return h !== 'operational' && h !== 'minor'; });
        const minorOnlySvcs = nonCleanSvcs.filter(n => { const e = healthMap.get(n.id); const h = e?.health ?? 'operational'; return h === 'minor' || (h === 'operational' && (e?.minorIncidents?.length ?? 0) > 0); });
        const collapsibleSvcs = [...minorOnlySvcs, ...cleanSvcs];

        let serviceHtml;
        if (collapsibleSvcs.length > 0 && (trulyAffSvcs.length > 0 || collapsibleSvcs.length < services.length)) {
            const MAX_PREVIEW = 5;
            const collNames = collapsibleSvcs.map(n => toFriendlyName(n.id));
            const collPreview = collNames.slice(0, MAX_PREVIEW).join(', ') + (collNames.length > MAX_PREVIEW ? '…' : '');
            const hasMinorInColl = minorOnlySvcs.length > 0;
            const dotCls = (hasMinorInColl && cleanSvcs.length === 0) ? 'status-dot--minor' : 'status-dot--operational';
            const label = hasMinorInColl
                ? `Show ${collapsibleSvcs.length} others: ${this._esc(collPreview)} ▼`
                : `Show ${collapsibleSvcs.length} fully operational: ${this._esc(collPreview)} ▼`;
            serviceHtml = `
                ${trulyAffSvcs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                <div class="status-op-toggle" tabindex="0" role="button" aria-expanded="false" data-count="${collapsibleSvcs.length}" data-preview="${this._esc(collPreview)}">
                    <span class="status-dot ${dotCls}"></span>
                    <span class="status-op-label">${label}</span>
                </div>
                <div class="status-op-list" hidden>
                    ${collapsibleSvcs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                </div>`;
        } else if (cleanSvcs.length > 0 && nonCleanSvcs.length > 0) {
            const MAX_PREVIEW = 5;
            const previewNames = cleanSvcs.map(n => toFriendlyName(n.id));
            const previewStr = previewNames.slice(0, MAX_PREVIEW).join(', ') +
                (previewNames.length > MAX_PREVIEW ? '…' : '');
            serviceHtml = `
                ${nonCleanSvcs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                <div class="status-op-toggle" tabindex="0" role="button" aria-expanded="false" data-count="${cleanSvcs.length}" data-preview="${this._esc(previewStr)}">
                    <span class="status-dot status-dot--operational"></span>
                    <span class="status-op-label">Show ${cleanSvcs.length} fully operational: ${this._esc(previewStr)} ▼</span>
                </div>
                <div class="status-op-list" hidden>
                    ${cleanSvcs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                </div>`;
        } else {
            serviceHtml = services.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('');
        }

        return `
            <div class="status-cluster" data-cluster-id="${this._esc(clusterId)}" data-health="${health}"${defaultExpanded ? ' data-expanded="true"' : ''}>
                <div class="status-cluster-header" tabindex="0" role="button" aria-expanded="${defaultExpanded}">
                    <span class="status-dot status-dot--${health}"></span>
                    <span class="status-cluster-label">${this._esc(label)}</span>
                    ${this._pill(health)}
                    ${incBadge}${minorBadge}
                    <span class="status-cluster-meta">${countTxt}</span>
                    <span class="status-expand-icon" aria-hidden="true">›</span>
                </div>
                <div class="status-cluster-services" ${defaultExpanded ? '' : 'hidden'}>
                    ${serviceHtml}
                </div>
            </div>`;
    }

    _renderServiceRow(node, entry) {
        const health     = entry?.health ?? 'operational';
        const hasInc     = Boolean(entry?.incidents?.length);
        const hasMinor   = Boolean(entry?.minorIncidents?.length);
        const friendly = toFriendlyName(node.id);
        const primary  = friendly;
        // Show the raw technical id as a subtitle only for slug-style identifiers
        const isSlug   = /[-_]/.test(node.id) || (!/\s/.test(node.id) && node.id === node.id.toLowerCase());
        const code     = isSlug ? node.id : null;
        const badge      = hasInc   ? `<span class="status-badge">${entry.incidents.length} open</span>` : '';
        const minorBadge = hasMinor ? `<span class="status-minor-badge" data-svc-id="${this._esc(node.id)}">▲ ${entry.minorIncidents.length}</span>` : '';
        const expand     = hasInc   ? `<span class="status-expand-icon" aria-hidden="true">›</span>` : '';

        return `
            <div class="status-svc-row${hasInc ? ' has-incidents' : ''}"
                 data-svc-id="${this._esc(node.id)}" data-health="${health}">
                <span class="status-dot status-dot--${health}"></span>
                <span class="status-svc-names">
                    <span class="status-svc-primary">${this._esc(primary)}</span>
                    ${code ? `<span class="status-svc-code">${this._esc(code)}</span>` : ''}
                </span>
                ${badge}${minorBadge}${expand}
            </div>`;
    }

    _pill(health) {
        return `<span class="status-pill status-pill--${health}">${HEALTH_LABEL[health] ?? health}</span>`;
    }

    _esc(s) {
        return String(s)
            .replace(/&/g,'&amp;').replace(/</g,'&lt;')
            .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }

    _cleanSummary(s) {
        return String(s)
            .replace(/§\[Roi\w+(?:,Roi\w+)*\]/g, '')
            .replace(/^\[GUCCI PRODUCTION\]\s*-\s*/i, '')
            .trim();
    }

    // ── Interactions ──────────────────────────────────────────────────────────

    _bindInteractions(healthMap, contentEl) {
        contentEl.querySelectorAll('.status-cluster-header').forEach(header => {
            const cluster = header.closest('.status-cluster');
            const body    = cluster?.querySelector('.status-cluster-services');
            if (!body) return;
            header.addEventListener('click', () => {
                const expanded = !body.hidden;
                body.hidden = expanded;
                header.setAttribute('aria-expanded', String(!expanded));
                const icon = header.querySelector('.status-expand-icon');
                if (icon) icon.style.transform = expanded ? '' : 'rotate(90deg)';
            });
            header.addEventListener('keydown', e => {
                if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); header.click(); }
            });
            if (cluster.dataset.expanded === 'true') {
                const icon = header.querySelector('.status-expand-icon');
                if (icon) icon.style.transform = 'rotate(90deg)';
            }
        });

        contentEl.querySelectorAll('.status-major-svc-tag[data-svc-id]').forEach(tag => {
            tag.style.cursor = 'pointer';
            tag.addEventListener('click', () => this._navigateToService(tag.dataset.svcId));
            tag.addEventListener('keydown', e => {
                if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this._navigateToService(tag.dataset.svcId); }
            });
        });

        // Amber badge click → open Jira open-items popup (same as graph view)
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
                if (e.target.closest('.status-minor-badge')) return; // handled above
                if (e.target.closest('.status-expand-icon') && row.classList.contains('has-incidents')) {
                    this._toggleIncidents(row, healthMap);
                } else {
                    this._navigateToService(row.dataset.svcId);
                }
            });
        });

        contentEl.querySelectorAll('.status-op-toggle').forEach(btn => {
            btn.addEventListener('click', () => {
                const list = btn.nextElementSibling;
                if (!list) return;
                const expanded = !list.hidden;
                list.hidden = expanded;
                btn.setAttribute('aria-expanded', String(!expanded));
                const label = btn.querySelector('.status-op-label');
                const count   = btn.dataset.count   || '';
                const preview = btn.dataset.preview || '';
                if (label) {
                    label.textContent = expanded
                        ? `Show ${count} fully operational: ${preview} ▼`
                        : `Hide ${count} fully operational ▲`;
                }
            });
            btn.addEventListener('keydown', e => {
                if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); btn.click(); }
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
