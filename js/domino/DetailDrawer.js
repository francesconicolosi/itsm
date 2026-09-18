import { isDateTimeValue, formatDateTimeLocal, createFormattedLongTextElementsFrom, splitValues, isUrl, setSearchQuery, formatUrlLink, renderUrlPartsIntoCell } from '../shared/utils.js';
import { labelForKey, isListViewVisible, refreshDrawerColumnIcons, descriptionFields } from './columns.js';
import { computeJiraIssuesValue } from './jira.js';
import { BRAND } from '../../brand-specific/brand.js';

const SEARCHABLE_ATTRS_ON_PEOPLE_DB = ['Theme', 'Stream', 'Owner', 'Service Manager', 'Responsible Teams', 'Accounts administered by', 'Accounts approved by', 'Accessed by'];
const PRIORITY_KEYS = ['Key', 'id', 'Description', 'Depends on', 'Used by'];

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

const JENGA_ICON_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" fill="none" class="app-link-icon" aria-hidden="true"><rect x="4" y="6" width="24" height="4" rx="2" fill="#3b82f6"/><rect x="4" y="14" width="18" height="4" rx="2" fill="#22c55e"/><rect x="4" y="22" width="24" height="4" rx="2" fill="#f97316"/><rect x="24" y="12" width="4" height="8" rx="1" fill="#a855f7" opacity=".7"/></svg>`;
const SOLITAIRE_ICON_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" class="app-link-icon" aria-hidden="true"><rect width="24" height="24" rx="4" fill="#111"/><path fill="white" d="M12 3C9 5.5 5 8.5 5 13c0 2.5 2 4 4 3.5C7.5 18.5 7 20 7 21H17c0-1-.5-2.5-2-4.5 2 .5 4-1 4-3.5C20 8.5 15 5.5 12 3Z"/></svg>`;

export class DetailDrawer {
    constructor(app) {
        this.app = app;
        this.currentNode = null;
        this._sortOrder = localStorage.getItem('domino_drawer_attr_sort') || 'original';
        this.onClose = null;
        this.jengaEvents = [];
    }

    _normalizeServiceName(name) {
        return (name || '').toLowerCase().trim().replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '');
    }

    _findJengaReleasesMonths(node) {
        if (!this.jengaEvents || !this.jengaEvents.length) return [];
        const nodeId = this._normalizeServiceName(node.id || '');
        const nodeName = this._normalizeServiceName(node['Service Name'] || '');

        const now = new Date();
        const curYear = now.getFullYear();
        const curMonth = now.getMonth(); // 0-indexed

        // Build set of (year, month) for past 5 calendar months (inclusive of current)
        const windowMonths = new Set();
        for (let i = 0; i < 5; i++) {
            let m = curMonth - i;
            let y = curYear;
            if (m < 0) { m += 12; y--; }
            windowMonths.add(`${y}-${m}`);
        }

        // Group matching events by year-month
        const byMonth = new Map();
        for (const ev of this.jengaEvents) {
            const svc = ev['Service'] || '';
            if (!svc) continue;
            const svcNorm = this._normalizeServiceName(svc);
            if (svcNorm !== nodeId && svcNorm !== nodeName) continue;

            const dateStr = ev['DueDate'] || ev['StartDate'] || '';
            if (!dateStr) continue;
            const [yearStr, monthStr] = dateStr.split('-');
            const year = parseInt(yearStr, 10);
            const month = parseInt(monthStr, 10) - 1; // 0-indexed
            if (isNaN(year) || isNaN(month)) continue;

            const key = `${year}-${month}`;
            if (!windowMonths.has(key)) continue;
            if (!byMonth.has(key)) byMonth.set(key, { year, month, service: svc });
        }

        return [...byMonth.values()]
            .sort((a, b) => a.year !== b.year ? a.year - b.year : a.month - b.month);
    }

    _updateSortButtons() {
        ['az', 'za'].forEach(order => {
            document.getElementById(`drawerSort-${order}`)
                ?.classList.toggle('drawer-sort-btn--active', this._sortOrder === order);
        });
    }

    initDOM() {
        document.getElementById('closeDrawer')?.addEventListener('click', () => this.closeDrawer());

        const drawerHeader = document.querySelector('#drawer .drawer-header');
        const closeBtn = document.getElementById('closeDrawer');
        if (drawerHeader && closeBtn) {
            ['az', 'za'].forEach(order => {
                const btn = document.createElement('button');
                btn.id = `drawerSort-${order}`;
                btn.className = 'drawer-sort-btn';
                btn.setAttribute('data-tooltip', order === 'az' ? 'Sort A → Z' : 'Sort Z → A');
                btn.textContent = order === 'az' ? '↑A' : '↓Z';
                btn.addEventListener('click', () => {
                    const wasActive = this._sortOrder === order;
                    this._sortOrder = wasActive ? 'original' : order;
                    localStorage.setItem('domino_drawer_attr_sort', this._sortOrder);
                    this._updateSortButtons();
                    if (this.currentNode) this.showNodeDetails(this.currentNode, false);
                    this.app.showToast(wasActive
                        ? 'Sort reset to original order'
                        : (order === 'az' ? 'Attributes sorted A → Z' : 'Attributes sorted Z → A'));
                });
                drawerHeader.insertBefore(btn, closeBtn);
            });

            const linkBtn = document.createElement('button');
            linkBtn.id = 'drawerCopyLink';
            linkBtn.className = 'drawer-sort-btn';
            linkBtn.setAttribute('data-tooltip', 'Copy link');
            linkBtn.textContent = '🔗';
            linkBtn.addEventListener('click', () => {
                if (!this.currentNode) return;
                const url = new URL(window.location.href);
                url.searchParams.set('search', `id:"${this.currentNode.id}"`);
                url.searchParams.set('openDrawer', 'true');
                navigator.clipboard?.writeText(url.toString());
                linkBtn.textContent = '✔';
                setTimeout(() => { linkBtn.textContent = '🔗'; }, 1500);
                this.app.showToast('Link to this service copied to clipboard');
            });
            drawerHeader.insertBefore(linkBtn, closeBtn);

            this._updateSortButtons();
        }
        document.getElementById('overlay')?.addEventListener('click', () => this.closeDrawer());

        document.addEventListener('keydown', (e) => {
            if (e.key !== 'Escape') return;
            // If the autocomplete dropdown is open, let AutocompleteEngine consume Escape first.
            if (document.getElementById('ac-dropdown')?.classList.contains('ac-open')) return;
            const drawer = document.getElementById('drawer');
            if (drawer?.classList.contains('open')) {
                this.closeDrawer();
                e.preventDefault();
                return;
            }
            const search = this.app.search;
            if (typeof search.searchTerm === 'string' && search.searchTerm.trim() !== '') {
                search.searchTerm = '';
                if (typeof search._refreshChips === 'function') search._refreshChips();
                const input = document.getElementById('drawer-search-input');
                if (input) input.value = '';
                setSearchQuery('');
                this.app.graph.updateVisualization();
                this.app.graph.fitGraphToViewport(0.9);
                e.preventDefault();
            }
        });
    }

    closeDrawer() {
        document.getElementById('drawer')?.classList.remove('open');
        document.getElementById('overlay')?.classList.remove('open');
        this.onClose?.();
    }

    getPeopleDbLink(value, fieldKey = null) {
        const query = fieldKey
            ? `${fieldKey}:"${value}"`
            : value.toLowerCase();
        const encoded = fieldKey
            ? encodeURIComponent(query)
            : encodeURIComponent(query).replace(/%20/g, '+');
        return `<a href="solitaire.html?search=${encoded}" target="_blank">${value}</a>`;
    }

    renderValueCell(key, raw, searchTerm) {
        const { search } = this.app;
        const td = document.createElement('td');
        if (typeof raw !== 'string') return td;

        if (isDateTimeValue(raw)) {
            td.textContent = formatDateTimeLocal(raw);
            td.title = raw;
            return td;
        }
        if (descriptionFields.includes(key)) {
            createFormattedLongTextElementsFrom(raw).forEach(el => td.appendChild(el));
            return td;
        }
        const parts = splitValues(raw);
        if (parts.some(isUrl)) {
            renderUrlPartsIntoCell(parts, td);
            return td;
        }
        const active = search.parseActiveKeyValueSearch(searchTerm);
        const isSameKey = !!active && active.key === key;
        const activeVals = new Set((active?.values || []).map(v => search.normalizeForCompare(v)));

        const makeToggleBtn = (v) => {
            if (!isSameKey) return '';
            const inSearch = activeVals.has(search.normalizeForCompare(v));
            const cls = inSearch ? 'search-remove' : 'search-add';
            const sym = inSearch ? '−' : '+';
            return ` <a class="fade-link search-toggle ${cls}" data-key="${encodeURIComponent(key)}" data-value="${encodeURIComponent(v)}" href="#">${sym}</a>`;
        };

        if (SEARCHABLE_ATTRS_ON_PEOPLE_DB.includes(key)) {
            const SOLITAIRE_FIELD_MAP = { 'Theme': 'theme', 'Stream': 'stream', 'Responsible Teams': 'team' };
            const solitaireField = SOLITAIRE_FIELD_MAP[key] ?? 'name';
            if (parts.length > 1) {
                const ul = document.createElement('ul');
                parts.forEach(v => {
                    const li = document.createElement('li');
                    li.innerHTML = `<i>${this.getPeopleDbLink(v, solitaireField)} <a class="fade-link search-trigger" data-key="${encodeURIComponent(key)}" data-value="${encodeURIComponent(v)}" href="#">⌞ ⌝</a>${makeToggleBtn(v)}</i>`;
                    ul.appendChild(li);
                });
                td.appendChild(ul);
            } else {
                const v = parts[0] || '';
                td.innerHTML = `<i>${this.getPeopleDbLink(v, solitaireField)} <a class="fade-link search-trigger" data-key="${encodeURIComponent(key)}" data-value="${encodeURIComponent(v)}" href="#">⌞ ⌝</a>${makeToggleBtn(v)}</i>`;
            }
            return td;
        }

        if (parts.length > 1) {
            const ul = document.createElement('ul');
            parts.forEach(v => {
                const li = document.createElement('li');
                li.innerHTML = `<i>${v} <a class="fade-link search-trigger" data-key="${encodeURIComponent(key)}" data-value="${encodeURIComponent(v)}" href="#">⌞ ⌝</a>${makeToggleBtn(v)}</i>`;
                ul.appendChild(li);
            });
            td.appendChild(ul);
        } else {
            const v = parts[0] || '';
            td.innerHTML = `<i>${v} <a class="fade-link search-trigger" data-key="${encodeURIComponent(key)}" data-value="${encodeURIComponent(v)}" href="#">⌞ ⌝</a>${makeToggleBtn(v)}</i>`;
        }
        return td;
    }

    renderKeyCell(key, { appIcon } = {}) {
        const { listView } = this.app;
        const td = document.createElement('td');
        const colKey = key === 'Service Name' ? 'id' : key;
        const keyLabel = document.createElement('span');
        keyLabel.textContent = key;
        if (appIcon) {
            const wrap = document.createElement('span');
            wrap.className = 'key-label-wrap';
            wrap.insertAdjacentHTML('afterbegin', appIcon);
            wrap.appendChild(keyLabel);
            td.appendChild(wrap);
        } else {
            td.appendChild(keyLabel);
        }
        if (isListViewVisible()) {
            const selected = listView.columnKeys.includes(colKey);
            const btn = document.createElement('button');
            btn.className = 'col-op fade-link';
            btn.type = 'button';
            btn.setAttribute('data-col', encodeURIComponent(colKey));
            btn.setAttribute('aria-label',
                selected ? `Remove "${labelForKey(colKey)}" from list view` : `Add "${labelForKey(colKey)}" to list view`);
            btn.textContent = selected ? '−' : '+';
            td.appendChild(btn);
        }
        return td;
    }

    showNodeDetails(node, openDrawer = true) {
        this._setDrawerActions(true);
        this.currentNode = node;
        const { search, listView } = this.app;
        const keyRaw = String(node['Key'] ?? '').trim();
        const serviceRaw = String(node['Service Name'] ?? '').trim();
        const idRaw = String(node.id ?? '').trim();
        const keyFromCsv = keyRaw !== '';
        const keyValue = keyFromCsv ? keyRaw : (serviceRaw || idRaw);

        node['Key'] = keyValue;
        const keyNorm = keyValue.toLowerCase();
        const idNorm = idRaw.toLowerCase();
        const keyEqualsId = keyNorm && keyNorm === idNorm;

        const priorityKeys = [...(keyEqualsId ? ['id'] : ['Key', 'id']), 'Description', 'Depends on', 'Used by'];

        const drawer = document.getElementById('drawer');
        const overlay = document.getElementById('overlay');
        const drawerContent = document.getElementById('drawerContent');
        const title = drawer.querySelector('.drawer-header h2');
        title.textContent = node['Service Name'] || 'Service Information';
        drawerContent.innerHTML = '';

        if (!keyFromCsv) node['Key'] = String(node['Service Name'] ?? node.id ?? '').trim();

        const excluded = new Set([
            'index', 'x', 'y', 'vy', 'vx', 'fx', 'fy', 'color',
            'Service Name',
            ...(keyEqualsId ? ['Key'] : [])
        ]);

        const table = document.createElement('table');
        const renderedKeys = new Set();

        const renderRow = (key, value, opts = {}) => {
            if (renderedKeys.has(key)) return;
            if (excluded.has(key)) return;
            if (typeof value !== 'string' || !value) return;
            const tr = document.createElement('tr');
            const appIcon = SEARCHABLE_ATTRS_ON_PEOPLE_DB.includes(key) ? SOLITAIRE_ICON_SVG : (opts.appIcon || null);
            tr.appendChild(this.renderKeyCell(key, { appIcon }));
            tr.appendChild(this.renderValueCell(key, value, search.searchTerm));
            table.appendChild(tr);
            renderedKeys.add(key);
        };

        this._updateSortButtons();
        const jiraUrl = computeJiraIssuesValue(node);
        const jengaMonths = this._findJengaReleasesMonths(node);
        const nodeKeys = Object.keys(node);
        const orderedKeys = priorityKeys.map(pk => nodeKeys.find(k => k === pk)).filter(Boolean);
        const remainingKeys = nodeKeys.filter(k => !orderedKeys.includes(k));
        if (jiraUrl) remainingKeys.push('Jira Issues');
        if (jengaMonths.length) remainingKeys.push('Releases');
        if (this._sortOrder === 'az') remainingKeys.sort((a, b) => a.localeCompare(b));
        else if (this._sortOrder === 'za') remainingKeys.sort((a, b) => b.localeCompare(a));
        [...orderedKeys, ...remainingKeys].forEach(key => {
            if (key === 'Releases') {
                const tr = document.createElement('tr');
                tr.appendChild(this.renderKeyCell('Releases', { appIcon: JENGA_ICON_SVG }));
                const td = document.createElement('td');
                const ul = document.createElement('ul');
                ul.className = 'jenga-releases-list';
                jengaMonths.forEach(({ year, month, service }) => {
                    const li = document.createElement('li');
                    const url = `jenga.html?view=timeline&year=${year}&month=${month + 1}&services=${encodeURIComponent(service)}`;
                    li.innerHTML = `<a href="${url}" target="_blank">${MONTH_NAMES[month]}</a>`;
                    ul.appendChild(li);
                });
                td.appendChild(ul);
                tr.appendChild(td);
                table.appendChild(tr);
                return;
            }
            renderRow(key, key === 'Jira Issues' ? jiraUrl : node[key]);
        });

        table.addEventListener('click', (e) => {
            const btn = e.target.closest('button.col-op');
            if (!btn) return;
            e.stopPropagation();
            listView.toggleColumn(decodeURIComponent(btn.getAttribute('data-col')));
            refreshDrawerColumnIcons();
        });

        drawerContent.appendChild(table);
        refreshDrawerColumnIcons();
        if (openDrawer) {
            drawer.classList.add('open');
            overlay.classList.add('open');
        }
    }

    _setDrawerActions(visible) {
        ['drawerSort-az', 'drawerSort-za', 'drawerCopyLink'].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.style.display = visible ? '' : 'none';
        });
    }

    showAbout() {
        this._setDrawerActions(false);
        const drawer = document.getElementById('drawer');
        const overlay = document.getElementById('overlay');
        const drawerContent = document.getElementById('drawerContent');
        const title = drawer?.querySelector('.drawer-header h2');
        if (title) title.textContent = 'About Domino ⚅';
        if (drawerContent) {
            drawerContent.innerHTML = `
<div class="about-content">
  <p class="about-intro">
    <strong>Domino</strong> is the Digital Service Catalog — a real-time dependency mapping engine
    built on top of the CMDB. It renders service relationships as an interactive force-directed graph,
    making it easy to understand how services connect, who owns them, and what would be affected
    by a change.
  </p>
  <div class="about-section">
    <h4>What it shows</h4>
    <ul>
      <li>Services and their upstream &amp; downstream dependencies across the CMDB</li>
      <li>Service ownership, lifecycle status, and full metadata from Jira Assets</li>
      <li>Dependency chains and blast-radius paths at a glance</li>
      <li>Cross-team service boundaries and shared responsibilities</li>
    </ul>
  </div>
  <div class="about-section">
    <h4>How to use it</h4>
    <ul>
      <li>Search by name, field, or value using the top search bar</li>
      <li>Use <code>field:"value"</code> syntax for precise filtering</li>
      <li>Click any node to see full service details in this panel</li>
      <li>Toggle the list view (&#x1F4CB;) for a tabular perspective</li>
      <li>Use <em>Show Decommissioned</em> to include retired services in the graph</li>
    </ul>
  </div>
  <p class="about-footer">
    More info on <a href="${BRAND.urls.servicePortal}" target="_blank">the ${BRAND.name} Service Portal</a>.
  </p>
</div>`;
        }
        drawer?.classList.add('open');
        overlay?.classList.add('open');
    }
}
