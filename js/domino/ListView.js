import { isUrl, createFormattedLongTextElementsFrom, splitValues, formatUrlLink, renderUrlPartsIntoCell, createOutlookUrl, truncateString } from '../shared/utils.js';
import { labelForKey, isListViewVisible, refreshDrawerColumnIcons, LABEL_FOR_KEY, getCellValue, descriptionFields } from './columns.js';
import { computeJiraIssuesValue } from './jira.js';

const DEFAULT_COLUMN_KEYS = ['id', 'Description', 'Type', 'Depends on', 'Status', 'Decommission Date'];

export class ListView {
    constructor(app) {
        this.app = app;
        this.sortKey = null;
        this.sortDir = 'asc';
        this.columnKeys = [...DEFAULT_COLUMN_KEYS];
        this.mapEl = null;
        this.listViewEl = null;
        this.legendEl = null;
        this.btnList = null;
        this.btnGraph = null;
        this.btnExport = null;
        this.btnStatus = null;
        this._filterDropdown = null;
    }

    initDOM() {
        this.mapEl = document.getElementById('map');
        this.listViewEl = document.getElementById('list-view');
        this.legendEl = document.getElementById('legend');
        this.btnList = document.getElementById('view-list');
        this.btnGraph = document.getElementById('view-graph');
        this.btnExport = document.getElementById('act-export-excel');
        this.btnStatus = document.getElementById('view-status');

        window.currentColumnKeys = this.columnKeys;

        this.btnList?.addEventListener('click', () => this.toListView());
        this.btnGraph?.addEventListener('click', () => this.toGraphView());
        this.btnExport?.addEventListener('click', () => {
            if (this.listViewEl?.style.display !== 'block') this.toListView();
            this._openExportModal();
        });

        document.getElementById('export-visible-cols')?.addEventListener('click', () => {
            this._closeExportModal(); this.exportToExcel(false);
        });
        document.getElementById('export-all-cols')?.addEventListener('click', () => {
            this._closeExportModal(); this.exportToExcel(true);
        });
        document.getElementById('export-cancel')?.addEventListener('click', () => this._closeExportModal());
        document.getElementById('export-modal')?.addEventListener('click', (e) => {
            if (e.target === e.currentTarget) this._closeExportModal();
        });
    }

    normalizeColumnToken(token) {
        if (!token) return null;
        const t = token.trim();
        if (!t) return null;
        if (/^(id|ID|Service Name)$/i.test(t)) return 'id';
        return t;
    }

    serializeColumnsToParam(keys) {
        return keys.map(k => (k === 'id' ? LABEL_FOR_KEY.id : k)).join(',');
    }

    parseListViewParam(param) {
        if (!param) return null;
        return param
            .split(',')
            .map(s => this.normalizeColumnToken(decodeURIComponent(s)))
            .filter(Boolean);
    }

    parseSortParam(param) {
        if (!param) return null;
        const [rawLabel, rawDir] = param.split(':');
        const key = this.normalizeColumnToken(decodeURIComponent(rawLabel || ''));
        if (!key) return null;
        const dir = (rawDir || 'asc').toLowerCase() === 'desc' ? 'desc' : 'asc';
        return { key, dir };
    }

    syncListViewParamInUrl() {
        const url = new URL(window.location.href);
        if (!this.columnKeys.length) {
            url.searchParams.delete('listView');
        } else {
            url.searchParams.set('listView', this.serializeColumnsToParam(this.columnKeys));
        }
        url.searchParams.delete('view');
        window.history.replaceState({}, '', url.toString());
    }

    _setActiveButton(active) {
        [this.btnList, this.btnGraph, this.btnStatus].forEach(b => {
            b?.classList.toggle('is-active', b === active);
        });
    }

    syncSortParamInUrl() {
        const url = new URL(window.location.href);
        if (!isListViewVisible() || !this.sortKey) {
            url.searchParams.delete('sort');
        } else {
            url.searchParams.set('sort', `${encodeURIComponent(labelForKey(this.sortKey))}:${this.sortDir}`);
        }
        window.history.replaceState({}, '', url.toString());
    }

    toggleColumn(key) {
        const idx = this.columnKeys.indexOf(key);
        if (idx >= 0) {
            if (this.columnKeys.length === 1) return;
            this.columnKeys.splice(idx, 1);
        } else {
            this.columnKeys.push(key);
        }
        window.currentColumnKeys = this.columnKeys;
        this.syncListViewParamInUrl();
        if (document.getElementById('list-view')?.style.display === 'block') {
            this.renderListFromSearch();
        }
        refreshDrawerColumnIcons();
    }

    toListView() {
        this.mapEl.style.display = 'none';
        if (this.legendEl) this.legendEl.style.display = 'none';
        this.listViewEl.style.display = 'block';
        this.app.statusView?.toOtherView();
        this._setActiveButton(this.btnList);
        this.syncListViewParamInUrl();
        this.syncSortParamInUrl();
        this.renderListFromSearch();
    }

    toGraphView() {
        this.mapEl.style.display = 'block';
        if (this.legendEl) this.legendEl.style.display = '';
        this.listViewEl.style.display = 'none';
        this.app.statusView?.toOtherView();
        this._setActiveButton(this.btnGraph);
        const url = new URL(window.location.href);
        url.searchParams.delete('listView');
        url.searchParams.delete('sort');
        url.searchParams.delete('view');
        window.history.replaceState({}, '', url.toString());
    }

    getComparableValue(n, key) {
        if (key === 'id') return (n.id ?? '').toString().toLowerCase();
        const raw = n?.[key] ?? '';
        if (key === 'Decommission Date') {
            const t = Date.parse(raw);
            return isNaN(t) ? Number.NEGATIVE_INFINITY : t;
        }
        return String(raw).toLowerCase();
    }

    getSortIndicator(key) {
        if (key !== this.sortKey) return '';
        return this.sortDir === 'asc' ? ' ↑' : ' ↓';
    }

    renderListFromSearch() {
        const { search, drawer } = this.app;
        const listViewEl = this.listViewEl;

        if (!search.currentNodes) {
            listViewEl.innerHTML = `<p class="empty-state">No data available.</p>`;
            return;
        }

        let results = search.currentNodes.filter(n => search.currentSearchedNodes?.has?.(n.id));
        const noSearch = search.searchTerm === '' || !search.searchTerm;
        if (isListViewVisible() && noSearch && results.length === 0) {
            results = [...search.currentNodes];
        }

        listViewEl.innerHTML = '';

        if (this.sortKey) {
            results = results.slice().sort((a, b) => {
                const va = this.getComparableValue(a, this.sortKey);
                const vb = this.getComparableValue(b, this.sortKey);
                let cmp = 0;
                if (this.sortKey === 'Decommission Date' && typeof va === 'number' && typeof vb === 'number') {
                    cmp = va === vb ? 0 : va < vb ? -1 : 1;
                } else {
                    cmp = String(va).localeCompare(String(vb), undefined, { sensitivity: 'base', numeric: true });
                }
                return this.sortDir === 'asc' ? cmp : -cmp;
            });
        }

        if (!results.length) {
            const empty = document.createElement('p');
            empty.className = 'empty-state';
            empty.textContent = 'No results after your filtered search.';
            listViewEl.appendChild(empty);
            return;
        }

        const table = document.createElement('table');
        table.className = 'result-table';
        table.style.setProperty('--cols', String(this.columnKeys.length));

        const thead = document.createElement('thead');
        thead.addEventListener('click', (e) => {
            const btn = e.target.closest('.col-op');
            if (btn) {
                e.preventDefault();
                e.stopPropagation();
                this.toggleColumn(decodeURIComponent(btn.getAttribute('data-col')));
                return;
            }
            const title = e.target.closest('.th-title');
            if (title) {
                const col = title.closest('th')?.getAttribute('data-col');
                if (!col) return;
                if (this.sortKey === col) {
                    this.sortDir = this.sortDir === 'asc' ? 'desc' : 'asc';
                } else {
                    this.sortKey = col;
                    this.sortDir = 'asc';
                }
                this.syncSortParamInUrl();
                this.renderListFromSearch();
            }
        });

        const trh = document.createElement('tr');
        this.columnKeys.forEach(key => {
            const th = document.createElement('th');
            th.setAttribute('data-col', key);
            const cellWrap = document.createElement('div');
            cellWrap.className = 'th-cell';
            const title = document.createElement('button');
            title.className = 'th-title fade-link';
            title.type = 'button';
            title.textContent = `${labelForKey(key)}${this.getSortIndicator(key)}`;
            const filterBtn = document.createElement('button');
            filterBtn.className = 'list-filter-btn' + (this._isColFilterActive(key) ? ' list-filter-btn--active' : '');
            filterBtn.type = 'button';
            filterBtn.setAttribute('aria-label', `Filter by ${labelForKey(key)}`);
            filterBtn.innerHTML = '<svg width="10" height="10" viewBox="0 0 24 24" fill="currentColor"><path d="M10 18h4v-2h-4v2zM3 6v2h18V6H3zm3 7h12v-2H6v2z"/></svg>';
            filterBtn.addEventListener('click', (e) => { e.stopPropagation(); this._openFilterDropdown(key, filterBtn); });
            const removeBtn = document.createElement('button');
            removeBtn.className = 'col-op fade-link';
            removeBtn.type = 'button';
            removeBtn.textContent = '−';
            removeBtn.setAttribute('data-col', encodeURIComponent(key));
            removeBtn.setAttribute('aria-label', `Remove "${labelForKey(key)}" from list view`);
            cellWrap.appendChild(title);
            cellWrap.appendChild(filterBtn);
            cellWrap.appendChild(removeBtn);
            th.appendChild(cellWrap);
            trh.appendChild(th);
        });
        thead.appendChild(trh);
        table.appendChild(thead);

        const tbody = document.createElement('tbody');
        results.forEach(n => {
            const tr = document.createElement('tr');
            tr.setAttribute('role', 'button');
            tr.tabIndex = 0;
            const openDetails = () => {
                this.app.graph.clickedNode = n;
                if (this.app.graph.labels) {
                    this.app.graph.labels.classed('highlight', d => d.id === n.id);
                }
                drawer.showNodeDetails(n, true);
            };
            tr.addEventListener('click', openDetails);
            tr.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openDetails(); }
            });

            this.columnKeys.forEach(key => {
                const td = document.createElement('td');
                let raw = (key === 'id') ? (n.id ?? '') : (n[key] ?? '');
                if (key === 'Jira Issues' && !raw) {
                    const computed = computeJiraIssuesValue(n);
                    if (computed) raw = computed;
                }
                if (typeof raw === 'string' && raw) {
                    const parts = splitValues(raw);
                    if (key === 'Team Email') {
                        const wrap = document.createElement('div');
                        wrap.style.cssText = 'display:flex;flex-direction:column;gap:4px;';
                        parts.forEach(email => {
                            const row = document.createElement('div');
                            row.style.cssText = 'display:flex;align-items:center;gap:4px;white-space:nowrap;';
                            const a = document.createElement('a');
                            a.href = createOutlookUrl([email]);
                            a.textContent = truncateString(email, 28);
                            a.title = email;
                            a.target = '_blank';
                            a.rel = 'noopener noreferrer';
                            a.addEventListener('click', e => e.stopPropagation());
                            row.appendChild(a);
                            const btn = document.createElement('button');
                            btn.className = 'drawer-icon-btn';
                            btn.setAttribute('aria-label', 'Copy email address');
                            btn.setAttribute('data-tooltip', 'Copy email');
                            btn.setAttribute('data-tooltip-placement', 'top');
                            btn.title = '';
                            btn.innerHTML = '<svg width="1em" height="1em" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>';
                            btn.addEventListener('click', async (e) => {
                                e.stopPropagation();
                                try { await navigator.clipboard.writeText(email); }
                                catch {
                                    const ta = document.createElement('textarea');
                                    ta.value = email;
                                    ta.style.cssText = 'position:fixed;top:-9999px;left:-9999px';
                                    document.body.appendChild(ta); ta.select();
                                    document.execCommand('copy'); ta.remove();
                                }
                                this.app.showToast('Email copied to clipboard');
                            });
                            row.appendChild(btn);
                            wrap.appendChild(row);
                        });
                        td.appendChild(wrap);
                    } else if (parts.some(p => isUrl(p))) {
                        renderUrlPartsIntoCell(parts, td);
                    } else if (descriptionFields.includes(key)) {
                        td.innerHTML = '';
                        createFormattedLongTextElementsFrom(raw).forEach(el => td.appendChild(el));
                    } else {
                        td.textContent = splitValues(raw).join(', ');
                    }
                } else {
                    td.textContent = getCellValue(n, key);
                }
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });

        table.appendChild(tbody);
        listViewEl.appendChild(table);
    }

    // ─── Export ──────────────────────────────────────────────────────────────

    _openExportModal()  { document.getElementById('export-modal').style.display = ''; }
    _closeExportModal() { document.getElementById('export-modal').style.display = 'none'; }

    _getExportRows() {
        const { search } = this.app;
        if (!search.currentNodes) return [];
        let results = search.currentNodes.filter(n => search.currentSearchedNodes?.has?.(n.id));
        const noSearch = !search.searchTerm;
        if (noSearch && results.length === 0) results = [...search.currentNodes];
        if (this.sortKey) {
            results = results.slice().sort((a, b) => {
                const va = this.getComparableValue(a, this.sortKey);
                const vb = this.getComparableValue(b, this.sortKey);
                let cmp = 0;
                if (this.sortKey === 'Decommission Date' && typeof va === 'number' && typeof vb === 'number') {
                    cmp = va === vb ? 0 : va < vb ? -1 : 1;
                } else {
                    cmp = String(va).localeCompare(String(vb), undefined, { sensitivity: 'base', numeric: true });
                }
                return this.sortDir === 'asc' ? cmp : -cmp;
            });
        }
        return results;
    }

    // ── Column filters ────────────────────────────────────────────────────────

    _collectColumnValues(key) {
        const all = this.app.search.currentNodes || [];
        const seen = new Set();
        const values = [];
        for (const node of all) {
            const raw = key === 'id' ? (node.id ?? '') : (node[key] ?? '');
            if (typeof raw !== 'string' || !raw) continue;
            for (const v of splitValues(raw)) {
                const t = v.trim();
                if (t && !seen.has(t)) { seen.add(t); values.push(t); }
            }
        }
        return values.sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true }));
    }

    _getFilterSelectedValues(key) {
        const engine = this.app.search;
        const clauses = engine.parseCompoundKeyValueSearch(engine.searchTerm || '');
        const clause = clauses.find(c => c.key === key && !c.isNegation);
        if (!clause || !clause.values.length) return null;
        return new Set(clause.values);
    }

    _isColFilterActive(key) {
        return this._getFilterSelectedValues(key) !== null;
    }

    _applyColumnFilter(key, selectedValues, allValues) {
        const engine = this.app.search;
        const remaining = engine.splitSearchClauses(engine.searchTerm || '')
            .filter(c => { const p = engine.parseKeyValueClause(c); return !p || p.key !== key || p.isNegation; });
        const isFiltered = selectedValues && selectedValues.size > 0 && selectedValues.size < allValues.length;
        if (isFiltered) remaining.push(engine.buildKeyValueSearch(key, [...selectedValues], true));
        engine.updateSearchAndRefresh(remaining.filter(Boolean).join(' & '), false);
        this.renderListFromSearch();
    }

    _closeFilterDropdown() {
        if (this._filterDropdown) { this._filterDropdown.remove(); this._filterDropdown = null; }
    }

    _openFilterDropdown(key, anchorEl) {
        if (this._filterDropdown) { this._closeFilterDropdown(); return; }

        const allValues = this._collectColumnValues(key);
        if (!allValues.length) return;

        const dropdown = document.createElement('div');
        dropdown.className = 'list-filter-dropdown';
        document.body.appendChild(dropdown);
        this._filterDropdown = dropdown;

        const rect = anchorEl.getBoundingClientRect();
        dropdown.style.top = `${rect.bottom + 2}px`;
        dropdown.style.left = `${Math.min(rect.left, window.innerWidth - 250)}px`;

        const searchWrap = document.createElement('div');
        searchWrap.className = 'list-filter-search-wrap';
        const searchInput = document.createElement('input');
        searchInput.type = 'text';
        searchInput.className = 'list-filter-search';
        searchInput.placeholder = 'Search values...';
        searchWrap.appendChild(searchInput);
        dropdown.appendChild(searchWrap);

        const scroll = document.createElement('div');
        scroll.className = 'list-filter-scroll';
        dropdown.appendChild(scroll);

        const rebuild = (filterText = '') => {
            scroll.innerHTML = '';
            const lc = filterText.toLowerCase();
            const visible = filterText ? allValues.filter(v => v.toLowerCase().includes(lc)) : allValues;
            const selectedSet = this._getFilterSelectedValues(key);

            const allChecked = selectedSet === null;
            const someChecked = selectedSet !== null && selectedSet.size > 0 && selectedSet.size < allValues.length;

            const allLabel = document.createElement('label');
            allLabel.className = 'list-filter-item list-filter-item--selectall';
            const allCheck = document.createElement('input');
            allCheck.type = 'checkbox';
            allCheck.checked = allChecked;
            allCheck.indeterminate = someChecked;
            allCheck.addEventListener('change', (e) => {
                e.stopPropagation();
                this._applyColumnFilter(key, null, allValues);
                rebuild(filterText);
            });
            allLabel.appendChild(allCheck);
            const allSpan = document.createElement('span');
            allSpan.textContent = '(Select All)';
            allLabel.appendChild(allSpan);
            scroll.appendChild(allLabel);

            visible.forEach(v => {
                const isChecked = selectedSet === null || selectedSet.has(v);
                const label = document.createElement('label');
                label.className = 'list-filter-item';
                const cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.checked = isChecked;
                cb.addEventListener('change', (e) => {
                    e.stopPropagation();
                    const current = this._getFilterSelectedValues(key);
                    const newSet = current ? new Set(current) : new Set(allValues);
                    if (cb.checked) newSet.add(v); else newSet.delete(v);
                    this._applyColumnFilter(key, newSet.size >= allValues.length ? null : newSet, allValues);
                    rebuild(filterText);
                });
                const span = document.createElement('span');
                span.textContent = v;
                label.appendChild(cb);
                label.appendChild(span);
                scroll.appendChild(label);
            });
        };

        rebuild();
        searchInput.addEventListener('input', () => rebuild(searchInput.value));
        requestAnimationFrame(() => searchInput.focus());

        const onOutside = (e) => {
            if (!dropdown.contains(e.target)) {
                this._closeFilterDropdown();
                document.removeEventListener('mousedown', onOutside);
            }
        };
        setTimeout(() => document.addEventListener('mousedown', onOutside), 0);
    }

    exportToExcel(allColumns) {
        const keys = allColumns ? this.app.store.allColumnKeys() : this.columnKeys;
        const rows = this._getExportRows();
        const data = rows.map(n =>
            Object.fromEntries(keys.map(k => [k, getCellValue(n, k) ?? '']))
        );
        /* global XLSX */
        const ws = XLSX.utils.json_to_sheet(data, { header: keys });
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Services');
        const date = new Date().toISOString().slice(0, 10);
        XLSX.writeFile(wb, `domino-export-${date}.xlsx`);
    }
}
