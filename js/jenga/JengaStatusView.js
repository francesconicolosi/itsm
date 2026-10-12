/* global BRAND */
import * as d3 from 'd3';
import { toFriendlyName, CLUSTERS, CLUSTER_BY_ID, resolveCluster } from '../shared/statusUtils.js';

// ── Exclude 3rd-party and vendor services from the status page ────────────────
const HIDE_TYPES = new Set(['SaaS', 'Service Provider']);


// ── Health helpers ────────────────────────────────────────────────────────────
const ACTIVE_STATUSES = new Set([
    'Sent to Internal Team', 'Pending', 'UAT',
    'Work in progress', 'Work In Progress', 'Scheduled',
    'Investigating', 'In Progress', 'Open',
]);

const P1P2_RANK     = { 'P1/Blocker': 0, 'P2/Critical': 1 };
const ALL_RANK      = { 'P1/Blocker': 0, 'P2/Critical': 1, 'P3/Major': 2, 'P4/Minor': 3 };
const HEALTH_FROM_RANK = ['outage', 'critical'];
const HEALTH_ORDER  = { outage: 0, critical: 1, degraded: 2, minor: 3, operational: 4 };

const HEALTH_LABEL  = {
    outage: 'Major Outage', critical: 'Critical',
    degraded: 'Degraded',  minor: 'Minor Issue',  operational: 'Operational',
};
const GLOBAL_MSG    = {
    outage:      'Major outage in progress',
    critical:    'Critical incidents active',
    degraded:    'Some services degraded',
    minor:       'Minor issues detected',
    operational: 'All systems operational',
};

// ── JengaStatusView ───────────────────────────────────────────────────────────
export class JengaStatusView {
    constructor(app) {
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

        // Build nodes keeping all raw CSV columns (same shape ServiceCatalogStore produces)
        const DEAD = new Set(['Stopped', 'Decommissioned']);
        this._nodes = catalog
            .filter(d => !DEAD.has(d['Status']) && !d['Decommission Date'])
            .map(d => ({ id: d['Service Name'], ...d }));

        this._cards    = current;
        this._allCards = all;
    }

    _buildHealthMap() {
        const healthMap = new Map();
        for (const card of this._cards) {
            if (!ACTIVE_STATUSES.has(card.Status)) continue;
            const isP1P2 = card.Priority === 'P1/Blocker' || card.Priority === 'P2/Critical';
            const isP3P4 = card.Priority === 'P3/Major'   || card.Priority === 'P4/Minor';
            if (!isP1P2 && !isP3P4) continue;

            const services = (card.AffectedServices || '').split('||').map(s => s.trim()).filter(Boolean);
            for (const svc of services) {
                if (!healthMap.has(svc)) {
                    healthMap.set(svc, { health: 'operational', rank: 99, incidents: [], minorIncidents: [] });
                }
                const entry = healthMap.get(svc);
                if (isP1P2) {
                    const rank = P1P2_RANK[card.Priority];
                    if (rank < entry.rank) { entry.rank = rank; entry.health = HEALTH_FROM_RANK[rank]; }
                    entry.incidents.push(card);
                } else {
                    entry.minorIncidents.push(card);
                }
            }
        }
        return healthMap;
    }

    // ── Main render ───────────────────────────────────────────────────────────

    async render(containerEl, { tv = false } = {}) {
        await this._ensureData();

        const healthMap = this._buildHealthMap();

        const activeServices = this.app?.search?.activeServices;
        const nodes = (this._nodes || []).filter(n => {
            if (HIDE_TYPES.has(n['Type'])) return false;
            if (!Boolean((n['Stream'] || '').trim()) || !Boolean((n['Theme'] || '').trim())) return false;
            if (activeServices?.size > 0 && !activeServices.has(n.id)) return false;
            return true;
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

        // Global stats
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

        // Compute dynamic column count for TV/multi-column layout
        const cols = tv
            ? Math.min(4, Math.max(2, Math.ceil(streamData.length / 3)))
            : 1;

        const inner = document.createElement('div');
        inner.className = `jenga-status-inner${tv ? ' jenga-status-tv' : ''}`;
        inner.style.setProperty('--status-cols', String(cols));
        inner.innerHTML =
            this._renderBanner(globalHealth, operational, affected, minorTotal) +
            (majorCards.length ? this._renderMajorIncidents(majorCards) : '') +
            streamData.map(s => this._renderStream(s, healthMap, maxMinorCount)).join('');

        containerEl.innerHTML = '';
        containerEl.appendChild(inner);

        this._bindInteractions(healthMap, inner);
    }

    // ── HTML helpers (copied verbatim from StatusView.js) ─────────────────────

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
                        ${services.map(s => `<span class="status-major-svc-tag">${this._esc(s)}</span>`).join('')}
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

    _renderStream(streamData, healthMap, maxMinorCount = 0) {
        const { name, health, themes, minorCount = 0, svcCount = 0, healthBreakdown = {} } = streamData;
        const total    = svcCount || 1;
        const affCount = (healthBreakdown.outage || 0) + (healthBreakdown.critical || 0) + (healthBreakdown.degraded || 0);
        const affPct   = Math.round(affCount / total * 100);
        const minorPct = maxMinorCount > 0 ? Math.round(minorCount / maxMinorCount * 45) : 0;
        return `
            <div class="status-stream" data-health="${health}">
                <div class="status-stream-header">
                    <span class="status-dot status-dot--${health}"></span>
                    <span class="status-stream-name">${name}</span>
                    ${this._pill(health)}
                    ${minorCount > 0 ? `<span class="status-minor-badge">▲ ${minorCount}</span>` : ''}
                </div>
                <div class="status-health-bar">
                    ${affPct > 0 ? `<div class="status-hb-seg--affected" style="width:${affPct}%"></div>` : ''}
                    ${minorPct > 0 ? `<div class="status-hb-seg--minor" style="width:${minorPct}%"></div>` : ''}
                    <div class="status-hb-seg--clean" style="flex:1"></div>
                </div>
                <div class="status-stream-body">
                    ${themes.map(t => this._renderTheme(t, healthMap)).join('')}
                </div>
            </div>`;
    }

    _renderTheme(themeData, healthMap) {
        const { name, health, clusters, individuals, minorCount = 0 } = themeData;

        const affectedClusters = clusters.filter(c => c.health !== 'operational');
        const minorClusters    = clusters.filter(c => c.health === 'operational' && c.totalMinor > 0);
        const cleanClusters    = clusters.filter(c => c.health === 'operational' && !c.totalMinor);

        const affectedIndivs = individuals.filter(n => (healthMap.get(n.id)?.health ?? 'operational') !== 'operational');
        const minorIndivs    = individuals.filter(n => {
            const e = healthMap.get(n.id);
            return (e?.health ?? 'operational') === 'operational' && (e?.minorIncidents?.length ?? 0) > 0;
        });
        const cleanIndivs    = individuals.filter(n => {
            const e = healthMap.get(n.id);
            return (e?.health ?? 'operational') === 'operational' && !(e?.minorIncidents?.length);
        });

        const opCount  = cleanClusters.length + cleanIndivs.length;
        const hasMinor = minorClusters.length + minorIndivs.length > 0;

        const MAX_PREVIEW = 5;
        const previewNames = [
            ...cleanClusters.map(c => c.label),
            ...cleanIndivs.map(n => toFriendlyName(n.id)),
        ];
        const previewStr = previewNames.slice(0, MAX_PREVIEW).join(', ') +
            (previewNames.length > MAX_PREVIEW ? '…' : '');

        return `
            <div class="status-theme" data-health="${health}">
                <div class="status-theme-header">
                    <span class="status-theme-name">${name}</span>
                    ${this._pill(health)}
                    ${minorCount > 0 ? `<span class="status-minor-badge">▲ ${minorCount}</span>` : ''}
                </div>
                <div class="status-theme-body">
                    ${affectedClusters.map(c => this._renderCluster(c, healthMap, true)).join('')}
                    ${affectedIndivs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                    ${hasMinor ? `
                        <div class="status-minor-section">
                            ${minorClusters.map(c => this._renderCluster(c, healthMap, false)).join('')}
                            ${minorIndivs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                        </div>` : ''}
                    ${opCount > 0 ? `
                        <div class="status-op-toggle" tabindex="0" role="button" aria-expanded="false" data-count="${opCount}" data-preview="${this._esc(previewStr)}">
                            <span class="status-dot status-dot--operational"></span>
                            <span class="status-op-label">Show ${opCount} fully operational: ${this._esc(previewStr)} ▼</span>
                        </div>
                        <div class="status-op-list" hidden>
                            ${cleanClusters.map(c => this._renderCluster(c, healthMap, false)).join('')}
                            ${cleanIndivs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                        </div>` : ''}
                </div>
            </div>`;
    }

    _renderCluster(clusterData, healthMap, defaultExpanded) {
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
            ? `<span class="status-minor-badge">▲ ${totalMinor}</span>`
            : '';
        const countTxt   = `${services.length} service${services.length !== 1 ? 's' : ''}`;

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
                    ${services.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                </div>
            </div>`;
    }

    _renderServiceRow(node, entry) {
        const health     = entry?.health ?? 'operational';
        const hasMinor   = Boolean(entry?.minorIncidents?.length);
        const friendly = toFriendlyName(node.id);
        const primary  = friendly;
        const isSlug   = /[-_]/.test(node.id) || (!/\s/.test(node.id) && node.id === node.id.toLowerCase());
        const code     = isSlug ? node.id : null;
        const minorBadge = hasMinor ? `<span class="status-minor-badge">▲ ${entry.minorIncidents.length}</span>` : '';

        return `
            <div class="status-svc-row" data-svc-id="${this._esc(node.id)}" data-health="${health}">
                <span class="status-dot status-dot--${health}"></span>
                <span class="status-svc-names">
                    <span class="status-svc-primary">${this._esc(primary)}</span>
                    ${code ? `<span class="status-svc-code">${this._esc(code)}</span>` : ''}
                </span>
                ${minorBadge}
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

    invalidate() {
        this._nodes    = null;
        this._cards    = null;
        this._allCards = null;
    }
}
