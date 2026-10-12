import { splitValues, setSearchQuery } from '../shared/utils.js';
import { ChipBar } from '../shared/ChipBar.js';

const DEFAULT_SEARCH_KEY = 'id';
const LS_RELAXED_SEARCH = 'solitaire_relaxed_search';
const LS_SHOW_CONNECTIONS = 'domino_show_connections';

export class SearchEngine {
    constructor(app) {
        this.app = app;
        this.searchTerm = '';
        this.hideStoppedServices = true;
        this.showConnections = true;
        this.currentSearchedNodes = new Set();
        this.currentNodes = [];
        this._chipBar = null;
    }

    // ─── Chip bar ─────────────────────────────────────────────────────────────

    initChipBar() {
        const field = document.getElementById('search-field');
        const input = document.getElementById('drawer-search-input');
        if (!field || !input) return;
        this._chipBar = new ChipBar(field, input, (term) => this.updateSearchAndRefresh(term, false));
    }

    _refreshChips() {
        this._chipBar?.render(
            this.searchTerm,
            (t) => this.parseActiveKeyValueSearch(t),
            (k, v, q) => this.buildKeyValueSearch(k, v, q),
            (v) => this.normalizeForCompare(v),
            (t) => this.parseCompoundKeyValueSearch(t)
        );
    }

    getTermToCompare(term) {
        return term.replaceAll('\n', '').replaceAll(' ', '').toLowerCase();
    }
    normalizeForCompare(v) {
        return (v ?? '').toString().replaceAll('\n', '').replaceAll(' ', '').toLowerCase();
    }

    splitSearchClauses(term) {
        const raw = (term ?? '').toString();
        const clauses = [];
        let buf = '';
        let inQuote = false;

        for (let i = 0; i < raw.length; i++) {
            const ch = raw[i];
            if (ch === '"') inQuote = !inQuote;
            if (ch === '&' && !inQuote) {
                if (buf.trim()) clauses.push(buf.trim());
                buf = '';
                continue;
            }
            buf += ch;
        }
        if (buf.trim()) clauses.push(buf.trim());
        return clauses;
    }

    parseKeyValueClause(rawClause) {
        if (!rawClause || !rawClause.includes(':')) return null;
        const raw = rawClause.trim();
        const isNegation = raw.startsWith('!');
        const clause = isNegation ? raw.slice(1).trim() : raw;
        const idx = clause.indexOf(':');
        if (idx === -1) return null;

        const key = clause.slice(0, idx).trim();
        if (!key) return null;

        const valuePart = clause.slice(idx + 1).trim();
        const quoted = valuePart.includes('"');
        const clean = valuePart.replaceAll('"', '');
        const values = splitValues(clean).map(v => v.trim()).filter(Boolean);
        return { key, values, quoted, isNegation, raw: rawClause };
    }

    parseCompoundKeyValueSearch(term) {
        if (!term || !term.includes(':')) return [];
        return this.splitSearchClauses(term)
            .map(clause => this.parseKeyValueClause(clause))
            .filter(Boolean);
    }

    isEffectivelyEmpty() {
        if (this.searchTerm === '') return true;
        if (!this.searchTerm.includes(':')) return false;
        const clauses = this.parseCompoundKeyValueSearch(this.searchTerm);
        return clauses.length > 0 && clauses.every(c => c.values.length === 0);
    }

    parseActiveKeyValueSearch(term) {
        const clauses = this.parseCompoundKeyValueSearch(term);
        // Keep the legacy method non-disruptive: consumers that expect a single
        // active key/value search still receive one only when the query is unambiguous.
        if (clauses.length !== 1) return null;
        if (clauses[0].isNegation) return null;
        const { key, values, quoted } = clauses[0];
        return { key, values, quoted };
    }

    buildKeyValueSearch(key, values, quoted) {
        if (!key || !values || !values.length) return '';
        const body = quoted ? values.map(v => `"${v}"`).join(',') : values.join(',');
        return `${key}:${body}`;
    }

    _matchesKeyValueClause(node, clause) {
        const { key, values, quoted, isNegation } = clause;
        if (!Object.keys(node).includes(key)) return isNegation;

        if (isNegation && values.length === 0) return (node[key] ?? '').trim() !== '';
        if (!values.length) return true;

        const expectedValues = values.map(v => this.getTermToCompare(v));
        const nodeParts = splitValues(node[key] ?? '');
        const matches = expectedValues.some(ev =>
            nodeParts.some(p =>
                quoted
                    ? this.getTermToCompare(p) === ev
                    : this.getTermToCompare(p).includes(ev)
            )
        );
        return isNegation ? !matches : matches;
    }

    isSearchResultWithKeyValue(node) {
        if (!this.searchTerm.includes(':')) return false;
        const clauses = this.parseCompoundKeyValueSearch(this.searchTerm);
        if (!clauses.length) return false;
        // Field clauses are combined with AND. Values inside the same field keep
        // the existing OR semantics via comma-separated multi values.
        return clauses.every(clause => this._matchesKeyValueClause(node, clause));
    }

    isSearchResultValueOnly(d) {
        if (this.searchTerm === '' || this.searchTerm.includes(':')) return false;
        const terms = this.searchTerm.toLowerCase().split(',').map(t => t.trim());
        return Object.values(d).some(value =>
            typeof value === 'string' && terms.some(t => value.toLowerCase().includes(t))
        );
    }
    prepareSearchTerm() {
        const relaxed = document.getElementById('relaxed-search');
        const relaxedEnabled = relaxed?.checked ?? false;
        if (this.searchTerm !== '' && !this.searchTerm.includes(':') && !this.searchTerm.includes(',') && !relaxedEnabled) {
            this.searchTerm = `${DEFAULT_SEARCH_KEY}:"${this.searchTerm}"`;
            const input = document.getElementById('drawer-search-input');
            if (input) input.value = this.searchTerm;
            setSearchQuery(this.searchTerm);
        }
    }

    updateSearchAndRefresh(q, showDrawer = true) {
        this.searchTerm = q || '';
        const displayTerm = this.isEffectivelyEmpty() ? '' : this.searchTerm;
        const input = document.getElementById('drawer-search-input');
        if (input) input.value = displayTerm;
        setSearchQuery(displayTerm);
        this.app.graph.updateVisualization(showDrawer);
        const nodeToRefresh = this.app.graph.clickedNode || this.app.drawer.currentNode;
        if (showDrawer && nodeToRefresh) {
            this.app.drawer.showNodeDetails(nodeToRefresh, true);
        }
        this._refreshChips();
        this.app.statusView?._onSearchUpdate?.();
    }

    handleQuery(q, showDrawer = true) {
        this.app.graph.clickedNode = null;
        this.searchTerm = q;
        const input = document.getElementById('drawer-search-input');
        if (input) input.value = q;
        setSearchQuery(q);
        this.app.graph.updateVisualization(showDrawer);
        window.scrollTo({ top: 0, behavior: 'smooth' });
        this._refreshChips();
        this.app.statusView?._onSearchUpdate?.();
    }

    initRelaxedSearchPersistence() {
        const relaxed = document.getElementById('relaxed-search');
        if (!relaxed) return;
        const saved = localStorage.getItem(LS_RELAXED_SEARCH);
        if (saved !== null) relaxed.checked = saved === '1' || saved === 'true';
        if (!relaxed.dataset.boundRelaxedPersist) {
            relaxed.dataset.boundRelaxedPersist = '1';
            relaxed.addEventListener('change', () => {
                localStorage.setItem(LS_RELAXED_SEARCH, relaxed.checked ? '1' : '0');
                this.app.graph.updateVisualization();
                this._refreshChips();
            });
        }
    }

    initShowConnectionsPersistence() {
        const el = document.getElementById('toggle-show-connections');
        if (!el) return;
        const saved = localStorage.getItem(LS_SHOW_CONNECTIONS);
        if (saved !== null) {
            this.showConnections = saved !== '0';
            el.checked = this.showConnections;
        }
        if (!el.dataset.boundShowConn) {
            el.dataset.boundShowConn = '1';
            el.addEventListener('change', () => {
                this.showConnections = el.checked;
                localStorage.setItem(LS_SHOW_CONNECTIONS, el.checked ? '1' : '0');
                this.app.graph.updateVisualization(false);
            });
        }
    }
}
