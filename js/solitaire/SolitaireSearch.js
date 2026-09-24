import { normalizeWs, setSearchQuery, removeQueryParam } from '../shared/utils.js';
import { ChipBar } from '../shared/ChipBar.js';
import {
    applySearchDimmingForMatches,
    clearFieldHighlights,
    clearSearchDimming,
    highlightGroup as highlightGroupUtils,
} from './search.js';
import { getNameFromTitleEl } from './orgUtils.js';

const UNKNOWN_MATCHER = /^(unknown|n\/?a|not\s*(set|available)|-|—|none)$/i;
const KNOWN_FIELDS = ['name', 'role', 'company', 'location', 'function', 'room', 'service', 'stream', 'theme', 'team'];

export class SolitaireSearch {
    constructor(app) {
        this.app = app;
        this.lastSearch = '';
        this.currentIndex = 0;
        // Collapsed-stream management for search
        this._savedCollapsedKeys = null;   // user's original collapsed state at search start
        this._allMatchDescs = null;        // descriptors for ALL matches across all streams
        this._currentVisibleStream = null; // which originally-collapsed stream is currently expanded
        this._domDirty = false;            // true when _findAllMatchDescs left stream positions stale
        this._chipBar = null;
    }

    // ─── Chip bar ─────────────────────────────────────────────────────────────

    initChipBar() {
        const field = document.getElementById('search-field');
        const input = document.getElementById('drawer-search-input');
        if (!field || !input) return;
        this._chipBar = new ChipBar(field, input, (term) => {
            if (!term) {
                this.clear();
            } else {
                this.search(term);
            }
        });
    }

    _refreshChips(searchTerm) {
        this._chipBar?.render(
            searchTerm || '',
            (t) => this._parseActiveKV(t),
            (k, v, q) => this._buildKV(k, v, q),
            (v) => this._normalizeForCompare(v),
            null   // Solitaire uses single-clause queries; no compound parser needed
        );
    }

    // Minimal key-value helpers for ChipBar (mirrors Domino SearchEngine pattern)
    _normalizeForCompare(v) {
        return (v ?? '').toString().replaceAll('\n', '').replaceAll(' ', '').toLowerCase();
    }

    _parseActiveKV(term) {
        if (!term || !term.includes(':')) return null;
        const colonIdx = term.indexOf(':');
        const key = term.slice(0, colonIdx).trim();
        if (!key) return null;
        const valuePart = term.slice(colonIdx + 1).trim();
        const quoted = valuePart.startsWith('"') && valuePart.endsWith('"') && valuePart.length >= 2;
        const clean = quoted ? valuePart.slice(1, -1) : valuePart;
        const values = clean.split(',').map(v => v.trim()).filter(Boolean);
        return { key, values, quoted };
    }

    _buildKV(key, values, quoted) {
        if (!key || !values?.length) return '';
        const body = quoted ? values.map(v => `"${v}"`).join(',') : values.join(',');
        return `${key}:${body}`;
    }

    // ─── Descriptor helpers ───────────────────────────────────────────────────

    _matchDesc(el) {
        const isGroup = typeof el.matches === 'function' && el.matches('g[data-key]');
        const g = isGroup ? el : el.closest?.('g[data-key]');
        if (!g) return null;
        const groupKey = g.getAttribute('data-key');
        const parts = groupKey.split('::');
        if (parts.length < 2) return null;
        const streamKey = `stream::${parts[1]}`;
        const cls = isGroup ? null : (el.classList?.[0] ?? null);
        return { groupKey, selector: cls ? `.${cls}` : null, streamKey };
    }

    _elFromDesc({ groupKey, selector }) {
        const g = document.querySelector(`g[data-key="${groupKey}"]`);
        if (!g) return null;
        return selector ? (g.querySelector(selector) ?? g) : g;
    }

    // ─── Query parser ─────────────────────────────────────────────────────────

    // Parses "field:value", "field:"exact"", "field:"Unknown"" and bare queries.
    // Returns { raw, q, field, exact, missing, noZoom }.
    _parseQuery(rawInput, extraOpts = {}) {
        const raw = (rawInput ?? '').toString().trim();
        const noZoom = !!extraOpts.noZoom;

        if (!raw.includes(':')) {
            return { raw, q: raw.toLowerCase(), field: '', exact: false, missing: false, noZoom };
        }

        const colonIdx = raw.indexOf(':');
        const fieldToken = raw.slice(0, colonIdx).trim().toLowerCase();
        let valuePart = raw.slice(colonIdx + 1).trim();

        const quoted = valuePart.startsWith('"') && valuePart.endsWith('"') && valuePart.length >= 2;
        if (quoted) valuePart = valuePart.slice(1, -1);

        const knownField = KNOWN_FIELDS.includes(fieldToken) ? fieldToken : '';
        const isMissing = quoted && UNKNOWN_MATCHER.test(valuePart.trim());

        return {
            raw,
            q: valuePart.trim().toLowerCase(),
            field: knownField,
            exact: quoted,
            missing: isMissing,
            noZoom,
        };
    }

    // ─── Query helpers ────────────────────────────────────────────────────────

    _runQuery(q, missing, normalizedField, exact = false) {
        const FIELD_SELECTORS = {
            name: '.profile-name',
            role: '.role-field',
            company: '.company-field',
            location: '.location-field',
            'function': '.function-field',
            room: '.room-field',
            service: '.team-title[data-services]',
            stream: '.stream-title[data-full-name]',
            theme: '.theme-title[data-full-name]',
            team: '.team-title[data-full-name]',
        };

        // Service field — handled separately (searches data-services on team titles)
        if (normalizedField === 'service') {
            return Array.from(document.querySelectorAll('.team-title[data-services]')).filter(n => {
                const services = (n.getAttribute('data-services') || '')
                    .split(',').map(s => s.trim().toLowerCase());
                return exact ? services.includes(q) : services.some(s => s.includes(q));
            });
        }

        if (missing && normalizedField) {
            const attrName =
                normalizedField === 'role' ? 'data-role' :
                normalizedField === 'company' ? 'data-company' :
                normalizedField === 'function' ? 'data-function' : 'data-location';
            return Array.from(document.querySelectorAll('g[data-key^="card::"]')).filter(n => {
                const norm = normalizeWs(n.getAttribute(attrName) || '').trim().toLowerCase();
                return !norm || UNKNOWN_MATCHER.test(norm);
            });
        }

        // For exact field searches, query card groups directly via data attributes
        // (field elements contain "Key: Value" in textContent, not the bare value)
        if (exact && normalizedField) {
            const FIELD_ATTR = {
                role: 'data-role',
                company: 'data-company',
                location: 'data-location',
                'function': 'data-function',
                room: 'data-room',
            };
            const attrName = FIELD_ATTR[normalizedField];
            if (attrName) {
                return Array.from(document.querySelectorAll('g[data-key^="card::"]')).filter(n => {
                    const val = normalizeWs(n.getAttribute(attrName) || '').trim().toLowerCase();
                    return val === q;
                });
            }
        }

        const sel = (normalizedField && FIELD_SELECTORS[normalizedField])
            || '.profile-name, .team-title, .theme-title, .stream-title, .role-field, .company-field, .location-field, .function-field, [data-services]';
        return Array.from(document.querySelectorAll(sel)).filter(n => {
            const txt = (n.getAttribute?.('data-full-name') || n.textContent || '').trim().toLowerCase();
            if (exact) {
                const exactServiceMatch = (n.getAttribute?.('data-services') || '')
                    .split(',').some(s => s.trim().toLowerCase() === q);
                return txt === q || exactServiceMatch;
            }
            const attrMatch = (n.getAttribute?.('data-services') || '').toLowerCase().includes(q);
            return txt.includes(q) || attrMatch;
        });
    }

    // Temporarily expand all collapsed streams (in-place, no localStorage write)
    // to find matches, then collapse them back. Returns descriptors.
    _findAllMatchDescs(q, missing, normalizedField, exact = false) {
        const renderer = this.app.renderer;
        const collapsed = renderer._getCollapsedKeys?.() ?? new Set();
        collapsed.forEach(k => renderer._expandInPlace?.(k));
        const matches = this._runQuery(q, missing, normalizedField, exact);
        collapsed.forEach(k => renderer._collapseInPlace?.(k));
        return matches.map(el => this._matchDesc(el)).filter(Boolean);
    }

    // ─── Result display ───────────────────────────────────────────────────────

    // Expands exactly the stream needed for the current result (if it was originally
    // collapsed), collapses any previously-expanded stream, and re-renders via
    // loadAndRender so neighbouring streams reposition correctly (no overlap).
    _showResult(q, missing, noZoom) {
        if (!this._allMatchDescs?.length) return;
        const { app } = this;
        const renderer = app.renderer;

        const desc = this._allMatchDescs[this.currentIndex];
        const inOriginallyCollapsed = this._savedCollapsedKeys?.has(desc.streamKey) ?? false;
        const nextVisibleStream = inOriginallyCollapsed ? desc.streamKey : null;

        if (nextVisibleStream !== this._currentVisibleStream || this._domDirty) {
            // Build the collapsed set: original, minus the one stream we need open
            const newCollapsed = new Set(this._savedCollapsedKeys ?? []);
            if (nextVisibleStream) newCollapsed.delete(nextVisibleStream);

            // Re-render if localStorage differs from desired state OR if _findAllMatchDescs
            // left stream positions stale (collapsed a stream in-place without repositioning).
            const currentStored = renderer._getCollapsedKeys?.() ?? new Set();
            const stateChanged =
                this._domDirty ||
                newCollapsed.size !== currentStored.size ||
                [...newCollapsed].some(k => !currentStored.has(k));

            this._currentVisibleStream = nextVisibleStream;
            this._domDirty = false;

            if (stateChanged) {
                localStorage.setItem('dsm-collapsed-v1', JSON.stringify([...newCollapsed]));
                app.loadAndRender(app.db.cachedCsvText);
            }
        }

        const target = this._elFromDesc(desc);
        if (!target) return;

        // Collect fresh versions of ALL match descriptors for dimming
        // (elements in collapsed streams are still in the DOM, just hidden)
        const allFresh = this._allMatchDescs.map(d => this._elFromDesc(d)).filter(Boolean);
        applySearchDimmingForMatches(allFresh);

        if (!missing && !noZoom) {
            const zoomTarget = target.matches?.('g[data-key^="card::"]')
                ? (target.querySelector('.profile-box') || target)
                : target;
            const containerGroup = target.closest?.(
                'g[data-key^="stream::"], g[data-key^="theme::"], g[data-key^="team::"]'
            );
            if (containerGroup) {
                renderer.fitElementToView(containerGroup, 750);
            } else {
                renderer.zoomToElement(zoomTarget, 1, 750);
            }
        }
    }

    // ─── Public API ───────────────────────────────────────────────────────────

    clear() {
        this.lastSearch = '';
        this.currentIndex = 0;
        const { app } = this;
        // Restore the user's original collapsed state and re-render if we changed it
        if (this._savedCollapsedKeys !== null) {
            const currentStored = app.renderer._getCollapsedKeys?.() ?? new Set();
            const savedArr  = [...this._savedCollapsedKeys].sort();
            const storedArr = [...currentStored].sort();
            if (JSON.stringify(savedArr) !== JSON.stringify(storedArr)) {
                localStorage.setItem('dsm-collapsed-v1', JSON.stringify([...this._savedCollapsedKeys]));
                app.loadAndRender(app.db.cachedCsvText);
            }
            this._savedCollapsedKeys = null;
            this._allMatchDescs = null;
            this._currentVisibleStream = null;
            this._domDirty = false;
        }
        const output = document.getElementById('output');
        if (output) output.textContent = '';
        app.searchParam = '';
        const searchInput = document.getElementById('drawer-search-input');
        if (searchInput) searchInput.value = '';
        setSearchQuery('');
        removeQueryParam('missing');
        this._refreshChips('');
        clearSearchDimming();
        clearFieldHighlights();
        app.renderer.fitToContent(0.9);
        app.drawer.close();
    }

    search(rawQuery, opts = {}) {
        const { app } = this;

        // Legacy opt-based calls (e.g. from old URL restore code or direct callers):
        // search('', { missing: true, field: 'Role' }) → synthesize canonical query string
        let effectiveRaw = (rawQuery ?? '').toString().trim();
        if (opts.missing && !effectiveRaw.includes(':')) {
            const fieldToken = (opts.field || '').toLowerCase().split(' ')[0] || 'role';
            effectiveRaw = `${fieldToken}:"Unknown"`;
        }

        // Parse the unified query string (field:"value", field:value, or bare text)
        const parsed = this._parseQuery(effectiveRaw, opts);
        const { raw, q, field: normalizedField, exact, missing } = parsed;
        const noZoom = parsed.noZoom;

        if (!q && !missing) {
            this.clear();
            return;
        }

        // Keep search input in sync — show the full canonical query string
        const searchInput = document.getElementById('drawer-search-input');
        if (searchInput && searchInput.value !== raw) {
            searchInput.value = raw;
        }

        const renderer = app.renderer;

        // Capture the user's collapsed state on the first search after a clear
        if (this._savedCollapsedKeys === null) {
            this._savedCollapsedKeys = renderer._getCollapsedKeys?.() ?? new Set();
        }

        const isNewQuery = raw !== this.lastSearch || missing;

        if (isNewQuery) {
            // Restore localStorage to the user's original state before finding matches
            localStorage.setItem('dsm-collapsed-v1', JSON.stringify([...this._savedCollapsedKeys]));
            this._currentVisibleStream = null;

            // Detect if _findAllMatchDescs will dirty the DOM: a stream is dirty when it is
            // in the collapsed set (should be collapsed) but is currently expanded in the DOM
            // because a previous search opened it via loadAndRender.
            this._domDirty = [...this._savedCollapsedKeys].some(k => {
                const el = document.querySelector(`g[data-key="${k}"]`);
                return el && !el.classList.contains('stream-collapsed');
            });

            // Find ALL matches across all streams (in-place expand → query → collapse back)
            this._allMatchDescs = this._findAllMatchDescs(q, missing, normalizedField, exact);

            if (this._allMatchDescs.length === 0) {
                clearSearchDimming();
                setSearchQuery(raw);
                removeQueryParam('missing');
                app.showToast(missing ? 'No result found for Unknown' : `No result found for "${q}"`);
                return;
            }

            this.lastSearch = raw;
            this.currentIndex = 0;
        } else {
            if (!this._allMatchDescs?.length) return;
            this.currentIndex = (this.currentIndex + 1) % this._allMatchDescs.length;
        }

        clearFieldHighlights();
        if (!opts?.keepDrawer) app.drawer.close();

        this._showResult(q, missing, noZoom);

        const count = this._allMatchDescs.length;
        app.showToast(
            missing
                ? `Found ${count} result(s).`
                : `Found ${count} result(s). Showing ${this.currentIndex + 1}/${count}.`
        );

        // Unified URL: always use a single ?search= param, never ?missing=
        setSearchQuery(raw);
        removeQueryParam('missing');
        this._refreshChips(missing ? '' : raw);

        // Field highlight on current result
        if (!missing) {
            const target = this._elFromDesc(this._allMatchDescs[this.currentIndex]);
            if (target) {
                const FIELD_CLASSES = ['role-field', 'company-field', 'location-field'];
                const hitClass = FIELD_CLASSES.find(c => target.classList?.contains(c));
                if (hitClass) {
                    target.classList.add('field-hit-highlight');
                } else {
                    const group = target.closest?.('g[data-key^="card::"]');
                    if (group) {
                        FIELD_CLASSES.forEach(cls => {
                            const el = group.querySelector('.' + cls);
                            if (!el) return;
                            const elTxt = (el.textContent || '').toLowerCase();
                            if (exact ? elTxt === q : elTxt.includes(q))
                                el.classList.add('field-hit-highlight');
                        });
                    }
                }
            }
        }

        // Role / Function drawer
        if (!missing && (normalizedField === 'role' || normalizedField === 'function')) {
            const mapping = normalizedField === 'role'
                ? app.db.roleDetailsMapping
                : app.db.functionDetailsMapping;
            // Keys are original-case; q is lowercased — do a case-insensitive lookup
            const key = mapping
                ? [...mapping.keys()].find(k => k.toLowerCase() === q)
                : null;
            const entry = key ? mapping.get(key) : null;
            if (entry?.description) {
                app.drawer.open({
                    name: key,
                    description: entry.description,
                    elements: entry.grants
                        ? { items: entry.grants.split(',').map(s => s.trim()).filter(Boolean) }
                        : undefined,
                    elementsTitle: normalizedField === 'role' ? 'Role Grants' : undefined,
                    _permalinkSearch: `${normalizedField}:${key}`,
                });
            }
        }

        // Service drawer (triggered when target is a team-title with matching service)
        // Skip when a role/function drawer was (or could be) opened above
        if (!missing && normalizedField !== 'role' && normalizedField !== 'function') {
            try {
                const target = this._elFromDesc(this._allMatchDescs[this.currentIndex]);
                if (!target) return;
                const teamTitleEl = target.matches?.('text.team-title') ? target
                    : target.closest?.('g')?.querySelector('text.team-title');
                if (!teamTitleEl) return;

                const rawServices = (teamTitleEl.getAttribute('data-services') || '')
                    .split(',').map(s => s.trim()).filter(Boolean);
                if (rawServices.length === 0) return;

                const norm = v => (v || '').toString().trim().toLowerCase();
                const normalized = rawServices.map(s => ({ raw: s, norm: norm(s) }));
                const hit = exact
                    ? normalized.find(svc => svc.norm === q)
                    : normalized.find(svc => svc.norm.includes(q));
                if (!hit) return;

                const teamName = teamTitleEl.getAttribute('data-team-name') || getNameFromTitleEl(teamTitleEl);
                const email = teamTitleEl.getAttribute('data-team-email') || '';
                const channels = (() => {
                    try { return JSON.parse(teamTitleEl.getAttribute('data-team-channels') || '[]'); }
                    catch { return []; }
                })();
                const description = teamTitleEl.getAttribute('data-team-description') || '';

                app.drawer.open({
                    name: teamName,
                    description,
                    elements: { items: rawServices },
                    channels,
                    email,
                    highlightService: hit.raw,
                    highlightQuery: q,
                    elementsBaseUrl: (s) => `domino.html?search=id%3A"${encodeURIComponent(s)}"`
                });
            } catch (e) {
                console.warn('Drawer open/highlight skipped:', e);
            }
        }
    }
}
