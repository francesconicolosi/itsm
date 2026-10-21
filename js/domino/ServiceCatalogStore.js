import * as d3 from 'd3';
import { getFormattedDate, splitValues, parseSectionMeta } from '../shared/utils.js';

export class ServiceCatalogStore {
    constructor() {
        this.nodes = [];
        this.links = [];
        this.activeServiceNodes = [];
        this.activeServiceNodeIds = new Set();
        this.jiraCards = [];
        this.jiraCardsByService = new Map();
        this.hasLoaded = false;
    }

    processData(data) {
        const requiredColumns = ['Service Name', 'Description', 'Type', 'Depends on', 'Status', 'Decommission Date'];
        const missingColumns = requiredColumns.filter(col => !data.columns.includes(col));
        if (missingColumns.length > 0) {
            alert(`Missing mandatory columns: ${missingColumns.join(', ')}`);
            return null;
        }

        if (data.columns.includes('Updated')) {
            const validDates = data
                .map(d => new Date(d['Updated']))
                .filter(date => !isNaN(date.getTime()));
            if (validDates.length > 0) {
                const lastUpdateEl = document.getElementById('side-last-update');
                if (lastUpdateEl) {
                    lastUpdateEl.textContent = `Last Update: ${getFormattedDate(new Date(Math.max(...validDates.map(d => d.getTime()))).toISOString())}`;
                }
            }
        }

        const colorScale = d3.scaleOrdinal(d3.schemeCategory10);

        this.nodes = data.map(d => {
            const color = d['Type'] ? colorScale(d['Type']) : '#888888';
            const node = { id: d['Service Name'], color };
            for (const key in d) node[key] = d[key];
            return node;
        });

        const nodeIds = new Set(this.nodes.map(d => d.id));
        const usedByMap = new Map();

        this.links = data.flatMap(d => {
            const src = d['Service Name'];
            const deps = splitValues(d['Depends on']);
            return deps.map(dep => {
                const depTrim = dep.trim();
                if (!nodeIds.has(depTrim)) return null;
                if (!usedByMap.has(depTrim)) usedByMap.set(depTrim, new Set());
                usedByMap.get(depTrim).add(src);
                return { source: src, target: depTrim };
            });
        }).filter(Boolean);

        const hasUsedByColumn = data.columns.includes('Used by');
        if (!hasUsedByColumn) {
            this.nodes.forEach(n => {
                const users = Array.from(usedByMap.get(n.id) || []);
                n['Used by'] = users.join('||');
            });
            data.columns = [...data.columns, 'Used by'];
        }

        this._applyJiraCardsToNodes();

        this.activeServiceNodes = this.nodes.filter(d =>
            d.Status !== 'Stopped' && d.Status !== 'Decommissioned' && !d['Decommission Date']
        );
        this.activeServiceNodeIds = new Set(this.activeServiceNodes.map(d => d.id));

        return colorScale;
    }


    processJiraCards(data = []) {
        this.jiraCards = Array.isArray(data) ? data : [];
        this._buildJiraCardIndex();
        this._applyJiraCardsToNodes();
    }

    _buildJiraCardIndex() {
        this.jiraCardsByService = new Map();

        const splitImpactedServices = (value) => splitValues(value)
            .flatMap(v => String(v || '').split(/[,;\n]/g))
            .map(v => v.trim())
            .filter(Boolean);

        this.jiraCards.forEach(row => {
            if (!this._isOpenJiraCard(row)) return;

            const issueKind = this._getJiraCardKind(row);
            if (!issueKind) return;

            const impactedServices = splitImpactedServices(row.AffectedServices || row['Affected Services'] || row.Service || row.Services);
            if (!impactedServices.length) return;

            const { clean: summary, roi, topPriority: metaPriority } = parseSectionMeta(row.Summary || '');
            const card = {
                key: row.Key || row.key || '',
                issueType: row.IssueType || row['Issue Type'] || '',
                requestType: row.RequestType || row['Request Type'] || '',
                created: row.Created || '',
                summary,
                roi,
                metaPriority,
                priority: (row.Priority || '').split('/')[0].trim(),
                status: row.Status || '',
                impactedServices,
                jiraUrl: row.JiraUrl || row.URL || row.Url || '',
                kind: issueKind,
            };

            impactedServices.forEach(serviceName => {
                const normalized = this._normalizeServiceName(serviceName);
                if (!normalized) return;
                if (!this.jiraCardsByService.has(normalized)) {
                    this.jiraCardsByService.set(normalized, { incidents: [], requests: [] });
                }
                const bucket = this.jiraCardsByService.get(normalized);
                if (issueKind === 'incident') bucket.incidents.push(card);
                if (issueKind === 'request') bucket.requests.push(card);
            });
        });
    }

    _applyJiraCardsToNodes() {
        if (!Array.isArray(this.nodes) || !this.nodes.length) return;

        this.nodes.forEach(node => {
            const bucket = this._findJiraCardsForNode(node);
            const incidents = bucket?.incidents || [];
            const requests = bucket?.requests || [];

            node.__jiraCards = { incidents, requests };
            node.__openIncidentCount = incidents.length;
            node.__openServiceRequestCount = requests.length;
        });
    }

    _findJiraCardsForNode(node) {
        const candidates = [
            node?.id,
            node?.['Service Name'],
            node?.Name,
            node?.Key,
        ].filter(Boolean);

        for (const candidate of candidates) {
            const normalized = this._normalizeServiceName(candidate);
            if (this.jiraCardsByService.has(normalized)) return this.jiraCardsByService.get(normalized);
        }

        return { incidents: [], requests: [] };
    }

    _getJiraCardKind(row) {
        const issueType = String(row?.IssueType || row?.['Issue Type'] || '').toLowerCase();
        const requestType = String(row?.RequestType || row?.['Request Type'] || '').toLowerCase();

        if (issueType.includes('incident') || requestType.includes('incident')) return 'incident';
        if (issueType.includes('service request') || requestType.includes('service request') || requestType.includes('generic service request')) return 'request';
        return null;
    }

    _isOpenJiraCard(row) {
        const status = String(row?.Status || '').trim().toLowerCase();
        if (!status) return true;
        return !/^(closed|resolved|done|cancelled|canceled|rejected)$/i.test(status);
    }

    _normalizeServiceName(value) {
        return String(value || '')
            .trim()
            .toLowerCase()
            .replace(/[\u2018\u2019]/g, "'")
            .replace(/[\u201c\u201d]/g, '"')
            .replace(/\s+/g, ' ');
    }

    allColumnKeys() {
        if (!this.nodes?.length) return [];
        return Object.keys(this.nodes[0]).filter(
            k => !k.startsWith('__') && k !== 'color' && k !== 'index'
        );
    }

    reset() {
        this.nodes = [];
        this.links = [];
        this.activeServiceNodes = [];
        this.activeServiceNodeIds = new Set();
        this.jiraCards = [];
        this.jiraCardsByService = new Map();
        this.hasLoaded = false;
    }
}
