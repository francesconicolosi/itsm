import { toFriendlyName, CLUSTERS, CLUSTER_BY_ID, resolveCluster } from './statusUtils.js';

// ── Constants ─────────────────────────────────────────────────────────────────

const HIDE_TYPES = new Set(['SaaS', 'Service Provider']);

const ACTIVE_STATUSES = new Set([
    'Sent to Internal Team', 'Pending', 'UAT',
    'Work in progress', 'Work In Progress', 'Scheduled',
    'Investigating', 'In Progress', 'Open',
]);

const P1P2_RANK     = { 'P1/Blocker': 0, 'P2/Critical': 1 };
export const ALL_RANK = { 'P1/Blocker': 0, 'P2/Critical': 1, 'P3/Major': 2, 'P4/Minor': 3 };
const HEALTH_FROM_RANK = ['outage', 'critical'];
export const HEALTH_ORDER = { outage: 0, critical: 1, degraded: 2, minor: 3, operational: 4 };

const HEALTH_LABEL = {
    outage: 'Major Outage', critical: 'Critical',
    degraded: 'Degraded',  minor: 'Minor Issues', operational: 'Operational',
};
const GLOBAL_MSG = {
    outage:      'Major outage in progress',
    critical:    'Critical incidents active',
    degraded:    'Some services degraded',
    minor:       'Minor issues detected',
    operational: 'All systems operational',
};

// ── StatusViewBase ────────────────────────────────────────────────────────────

export class StatusViewBase {

    // ── Node filtering ────────────────────────────────────────────────────────

    _filterNode(n, brandHideTypes, activeServices) {
        if (HIDE_TYPES.has(n['Type']) || brandHideTypes?.has(n['Type'])) return false;
        if ((n['Hide from Status Page'] || '').trim().toLowerCase() === 'true') return false;
        if (!Boolean((n['Stream'] || '').trim()) || !Boolean((n['Theme'] || '').trim())) return false;
        if (activeServices?.size > 0 && !activeServices.has(n.id)) return false;
        return true;
    }

    // ── Health ────────────────────────────────────────────────────────────────

    _worstHealth(healths) {
        let best = 4;
        for (const h of healths) { const r = HEALTH_ORDER[h] ?? 4; if (r < best) best = r; }
        return Object.keys(HEALTH_ORDER)[best] ?? 'operational';
    }

    _buildHealthMap() {
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

    // ── Stream data ───────────────────────────────────────────────────────────

    _buildStreamData(nodes, healthMap) {
        const sumMinor = ns => ns.reduce((s, n) => s + (healthMap.get(n.id)?.minorIncidents?.length ?? 0), 0);

        const streamMap = new Map();
        for (const node of nodes) {
            const stream = node['Stream'].trim();
            const theme  = node['Theme'].trim();
            if (!streamMap.has(stream)) streamMap.set(stream, new Map());
            const themeMap = streamMap.get(stream);
            if (!themeMap.has(theme)) themeMap.set(theme, []);
            themeMap.get(theme).push(node);
        }

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
                    const cHealth = this._worstHealth(clusterSvcs.map(n => healthMap.get(n.id)?.health ?? 'operational'));
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

                // Skip themes that have no services after all filters
                if (clusters.length === 0 && individuals.length === 0) continue;

                const themeMinor = clusters.reduce((s, c) => s + c.totalMinor, 0) +
                    individuals.reduce((s, n) => s + (healthMap.get(n.id)?.minorIncidents?.length ?? 0), 0);
                themes.push({
                    name:       themeName,
                    health:     this._worstHealth([...clusters.map(c => c.health), ...individuals.map(n => healthMap.get(n.id)?.health ?? 'operational')]),
                    clusters,
                    individuals,
                    minorCount: themeMinor,
                });
            }

            // Skip streams that have no visible themes
            if (themes.length === 0) continue;

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
                health:         this._worstHealth(themes.map(t => t.health)),
                themes,
                minorCount:     streamMinor,
                svcCount:       streamSvcCount,
                healthBreakdown,
            });
        }
        streamData.sort((a, b) => b.svcCount - a.svcCount || a.name.localeCompare(b.name));
        const maxMinorCount = streamData.reduce((max, s) => Math.max(max, s.minorCount), 0);

        return { streamData, maxMinorCount };
    }

    // ── HTML helpers ──────────────────────────────────────────────────────────

    _renderMajorIncidents(cards, allCards = null) {
        const effectiveAll = allCards ?? cards;
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

        const worstHealth = effectiveAll.some(c => c.Priority === 'P1/Blocker') ? 'outage' : 'critical';
        return `
        <div class="status-major-panel">
            <div class="status-major-panel-title">
                <span class="status-dot status-dot--${worstHealth}"></span>
                Active Major Incidents
                <span class="status-major-count">${effectiveAll.length}</span>
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

    _tvOperationalLabel() { return GLOBAL_MSG['operational']; }

    _renderBanner(health, operational, affected, minorTotal, tv = false) {
        const displayHealth = (tv && health === 'minor') ? 'operational' : health;
        const globalMsg = (tv && health === 'minor') ? this._tvOperationalLabel() : GLOBAL_MSG[displayHealth];
        const minorTag = minorTotal > 0
            ? `<span class="status-global-minor">▲ ${minorTotal} minor (P3/P4)</span>` : '';
        return `
            <div class="status-global-banner status-global-banner--${displayHealth}">
                <span class="status-dot status-dot--${displayHealth}"></span>
                <span class="status-global-msg">${globalMsg}</span>
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
                    <a class="status-stream-name status-solitaire-link" href="solitaire.html?search=${encodeURIComponent(`stream:"${name}"`)}" target="_blank" rel="noopener" title="View in Solitaire">${name}</a>
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
                    <a class="status-theme-name status-solitaire-link" href="solitaire.html?search=${encodeURIComponent(`theme:"${name}"`)}" target="_blank" rel="noopener" title="View in Solitaire">${name}</a>
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
            const collLabel = hasMinorInColl
                ? `Show ${collapsibleSvcs.length} others: ${this._esc(collPreview)} ▼`
                : `Show ${collapsibleSvcs.length} fully operational: ${this._esc(collPreview)} ▼`;
            serviceHtml = `
                ${trulyAffSvcs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                <div class="status-op-toggle" tabindex="0" role="button" aria-expanded="false" data-count="${collapsibleSvcs.length}" data-preview="${this._esc(collPreview)}">
                    <span class="status-dot ${dotCls}"></span>
                    <span class="status-op-label">${collLabel}</span>
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
        const hasMinor   = Boolean(entry?.minorIncidents?.length);
        const friendly   = toFriendlyName(node.id);
        const isSlug     = /[-_]/.test(node.id) || (!/\s/.test(node.id) && node.id === node.id.toLowerCase());
        const code       = isSlug ? node.id : null;
        const hasMajor   = Boolean(entry?.incidents?.length);
        const majorBadge = hasMajor ? `<span class="status-major-badge" data-svc-id="${this._esc(node.id)}">▲ ${entry.incidents.length}</span>` : '';
        const minorBadge = hasMinor ? `<span class="status-minor-badge" data-svc-id="${this._esc(node.id)}">▲ ${entry.minorIncidents.length}</span>` : '';

        return `
            <div class="status-svc-row" data-svc-id="${this._esc(node.id)}" data-health="${health}">
                <span class="status-dot status-dot--${health}"></span>
                <span class="status-svc-names">
                    <span class="status-svc-primary">${this._esc(friendly)}</span>
                    ${code ? `<span class="status-svc-code">${this._esc(code)}</span>` : ''}
                </span>
                ${majorBadge}${minorBadge}
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
        return String(s).replace(/§\[Roi\w+(?:,Roi\w+)*\]/g, '').trim();
    }

    // ── Stability Thermometer ─────────────────────────────────────────────────

    _renderStabilityThermometer(tv = false) {
        const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        const now = new Date();

        // Build last 7 calendar months
        const months = [];
        for (let i = 6; i >= 0; i--) {
            const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
            months.push({ year: d.getFullYear(), month: d.getMonth(), p1: 0, p2: 0 });
        }

        // Count all P1/P2 incidents (including closed) grouped by Created month
        for (const card of (this._allCards || [])) {
            if (!card.Created) continue;
            if (card.Priority !== 'P1/Blocker' && card.Priority !== 'P2/Critical') continue;
            if (!(card.IssueType || '').toLowerCase().includes('incident')) continue;
            const d = new Date(card.Created);
            if (isNaN(d)) continue;
            const mi = months.findIndex(m => m.year === d.getFullYear() && m.month === d.getMonth());
            if (mi === -1) continue;
            if (card.Priority === 'P1/Blocker') months[mi].p1++;
            else months[mi].p2++;
        }

        // SVG canvas
        const W = 560, H = 220;
        const mt = 20, mr = 20, mb = 55, ml = 42;
        const cw = W - ml - mr;
        const ch = H - mt - mb;

        // Y scale – integer steps, min range 3
        const maxTotal = Math.max(...months.map(m => m.p1 + m.p2), 0);
        const rawMax   = Math.ceil(maxTotal) + 1;
        const step     = rawMax <= 11 ? 1 : rawMax <= 22 ? 2 : rawMax <= 55 ? 5 : 10;
        const yMax     = Math.max(Math.ceil(rawMax / step) * step, 3);

        const yPos = v => mt + ch * (1 - v / yMax);
        const barW = Math.round(cw / months.length * 0.55);
        const gapW = cw / months.length;

        // Y grid + tick labels
        const yLines = [];
        for (let t = 0; t <= yMax; t += step) {
            const y = yPos(t).toFixed(1);
            yLines.push(`<line x1="${ml}" y1="${y}" x2="${ml + cw}" y2="${y}" stroke="currentColor" stroke-opacity="0.1" stroke-width="1"/>`);
            yLines.push(`<text x="${ml - 6}" y="${(+y + 4).toFixed(1)}" text-anchor="end" font-size="11" fill="currentColor" fill-opacity="0.55">${t}</text>`);
        }

        // Bars + x-axis labels
        const bars = months.map((m, i) => {
            const cx = ml + (i + 0.5) * gapW;
            const bx = (cx - barW / 2).toFixed(1);
            const parts = [];
            if (m.p1 > 0) {
                const h = (m.p1 / yMax * ch).toFixed(1);
                parts.push(`<rect x="${bx}" y="${yPos(m.p1).toFixed(1)}" width="${barW}" height="${h}" fill="#991b1b" rx="2"/>`);
            }
            if (m.p2 > 0) {
                const h = (m.p2 / yMax * ch).toFixed(1);
                parts.push(`<rect x="${bx}" y="${yPos(m.p1 + m.p2).toFixed(1)}" width="${barW}" height="${h}" fill="#d97706" rx="2"/>`);
            }
            const lx = cx.toFixed(1);
            parts.push(`<text x="${lx}" y="${(H - mb + 16).toFixed(1)}" text-anchor="middle" font-size="11" fill="currentColor" fill-opacity="0.6">${MONTH_NAMES[m.month]}</text>`);
            parts.push(`<text x="${lx}" y="${(H - mb + 30).toFixed(1)}" text-anchor="middle" font-size="10" fill="currentColor" fill-opacity="0.45">${m.year}</text>`);
            return parts.join('');
        }).join('');

        // Legend
        const legX = Math.round(W / 2 - 55);
        const legY = H - 6;
        const legend = `
            <rect x="${legX}" y="${legY - 10}" width="11" height="11" fill="#991b1b" rx="2"/>
            <text x="${legX + 16}" y="${legY}" font-size="11" fill="currentColor" fill-opacity="0.7">P1</text>
            <rect x="${legX + 50}" y="${legY - 10}" width="11" height="11" fill="#d97706" rx="2"/>
            <text x="${legX + 66}" y="${legY}" font-size="11" fill="currentColor" fill-opacity="0.7">P2</text>`;

        return `
            <div class="status-thermometer${tv ? ' status-thermometer--tv' : ''}">
                <div class="status-thermometer-title">Stability Thermometer</div>
                ${tv ? '<div class="status-thermometer-chart-wrap">' : ''}
                <svg viewBox="0 0 ${W} ${H}" class="status-thermometer-chart" role="img" aria-label="P1 and P2 incidents per month">
                    <line x1="${ml}" y1="${mt}" x2="${ml}" y2="${mt + ch}" stroke="currentColor" stroke-opacity="0.25" stroke-width="1"/>
                    <line x1="${ml}" y1="${mt + ch}" x2="${ml + cw}" y2="${mt + ch}" stroke="currentColor" stroke-opacity="0.25" stroke-width="1"/>
                    ${yLines.join('')}
                    ${bars}
                    ${legend}
                </svg>
                ${tv ? '</div>' : ''}
            </div>`;
    }

    // ── Shared interactions ───────────────────────────────────────────────────

    _bindCommonInteractions(healthMap, contentEl) {
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
}
