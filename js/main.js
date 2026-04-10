import * as d3 from 'd3';

import {
    getQueryParam,
    setSearchQuery,
    closeSideDrawer,
    initCommonActions,
    getFormattedDate,
    createFormattedLongTextElementsFrom,
    getCellValue,
    refreshDrawerColumnIcons,
    splitValues,
    isUrl,
    labelForKey,
    enableGlobalFindShortcut,
    LABEL_FOR_KEY, isListViewVisible, isDateTimeValue, formatDateTimeLocal
} from './utils.js';


let nodes = [];
let links = [];
let sortKey = null;
let sortDir = 'asc';

const uniqueIds = ["id", "Name"];

function parseSortParam(param) {
    if (!param) return null;
    const [rawLabel, rawDir] = param.split(':');
    const key = normalizeColumnToken(decodeURIComponent(rawLabel || ''));
    if (!key) return null;
    const dir = (rawDir || 'asc').toLowerCase() === 'desc' ? 'desc' : 'asc';
    return {key, dir};
}

function syncSortParamInUrl() {
    const url = new URL(window.location.href);
    const isListVisible = isListViewVisible();
    if (!isListVisible || !sortKey) {
        url.searchParams.delete('sort');
    } else {
        url.searchParams.set('sort', `${encodeURIComponent(labelForKey(sortKey))}:${sortDir}`);
    }
    window.history.replaceState({}, '', url.toString());
}

function getComparableValue(n, key) {
    if (key === 'id') return (n.id ?? '').toString().toLowerCase();

    const raw = n?.[key] ?? '';
    if (key === 'Decommission Date') {
        const t = Date.parse(raw);
        return isNaN(t) ? Number.NEGATIVE_INFINITY : t; // numerico per confronti
    }
    return String(raw).toLowerCase();
}

function getSortIndicator(key) {
    if (key !== sortKey) return '';
    return sortDir === 'asc' ? ' ↑' : ' ↓';
}

const descriptionFields = ['Contingency and Recovery Planning', 'Description'];

function normalizeColumnToken(token) {
    if (!token) return null;
    const t = token.trim();
    if (!t) return null;

    if (/^(id|ID|Service Name)$/i.test(t)) return 'id';

    return t;
}

function serializeColumnsToParam(keys) {
    const tokens = keys.map(k => (k === 'id' ? LABEL_FOR_KEY.id : k));
    return tokens.join(',');
}

function parseListViewParam(param) {
    if (!param) return null;
    return param
        .split(',')
        .map(s => normalizeColumnToken(decodeURIComponent(s)))
        .filter(Boolean);
}

const DEFAULT_COLUMN_KEYS = [
    'id',
    'Description',
    'Type',
    'Depends on',
    'Status',
    'Decommission Date'
];

window.currentColumnKeys = [...DEFAULT_COLUMN_KEYS];

function syncListViewParamInUrl() {
    const url = new URL(window.location.href);
    if (!window.currentColumnKeys.length) {
        url.searchParams.delete('listView');
    } else {
        url.searchParams.set('listView', serializeColumnsToParam(window.currentColumnKeys));
    }
    window.history.replaceState({}, '', url.toString());
}

function toggleColumn(key) {
    const idx = window.currentColumnKeys.indexOf(key);
    if (idx >= 0) {
        if (window.currentColumnKeys.length === 1) return;
        window.currentColumnKeys.splice(idx, 1);
    } else {
        window.currentColumnKeys.push(key);
    }
    syncListViewParamInUrl();

    if (document.getElementById('list-view')?.style.display === 'block') {
        renderListFromSearch();
    }
    refreshDrawerColumnIcons();
}


let hideStoppedServices = !(document.getElementById('toggle-decommissioned').checked);
let searchTerm = "";
let activeServiceNodes;
let activeServiceNodeIds;
let linkGraph;
let nodeGraph;
let labels;
let currentSearchedNodes = new Set();
let currentNodes = [];
let simulation;
let g;
let zoom;
let zoomIdentity;
let svg;
let clickedNode;
let hasLoaded = false;
const width = document.getElementById('map').clientWidth;
const height = document.getElementById('map').clientHeight;

const searchableAttributesOnPeopleDb = ["Product Theme", "Owner"];
const defaultSearchKey = "id";

const serviceInfoEnhancers = [
    function generateIssueTrackingTool(node) {
        const url = computeTrackingSoftwareValue(node);
        if (!url) return null;
        return {key: "Tracking Issues", value: url};
    }
];


function centerAndZoomOnNode(node) {
    const scale = 1;
    const x = -node.x * scale + width / 2;
    const y = -node.y * scale + height / 2;


    const transform = zoomIdentity
        .translate(x, y)
        .scale(scale)
        .translate(-0, -0);

    svg.transition().duration(750).call(
        zoom.transform,
        transform
    );
}

function resetVisualization() {
    d3.select('#map').selectAll('*').remove();
    d3.select('#tooltip').style('opacity', 0);
    d3.select('#legend').selectAll('*').remove();
    d3.select('#serviceDetails').innerHTML = '';
    nodes = [];
    links = [];
    linkGraph = null;
    nodeGraph = null;
    labels = [];
    hideStoppedServices = true;
    searchTerm = "";
    activeServiceNodes = [];
    activeServiceNodeIds = [];
}

function fitGraphToViewport(paddingRatio = 0.90) {
    if (!svg || !g) return;
    const bbox = g.node()?.getBBox();
    if (!bbox || !isFinite(bbox.width) || !isFinite(bbox.height) || bbox.width === 0 || bbox.height === 0) {
        // reset zoom se bbox non valido
        svg.call(zoom.transform, d3.zoomIdentity);
        return;
    }
    const w = width;
    const h = height;
    const scale = Math.min(w / bbox.width, h / bbox.height) * paddingRatio;

    const tx = w / 2 - (bbox.x + bbox.width / 2) * scale;
    const ty = h / 2 - (bbox.y + bbox.height / 2) * scale;

    const t = d3.zoomIdentity.translate(tx, ty).scale(scale);
    svg.transition().duration(400).call(zoom.transform, t);
}

function handleQuery(q, showDrawer = true) {
    clickedNode = null;
    searchTerm = q;
    const searchInput = document.getElementById('drawer-search-input');
    if (searchInput) searchInput.value = q;
    setSearchQuery(q);
    updateVisualization(nodeGraph, linkGraph, labels, showDrawer);
    window.scrollTo({top: 0, behavior: 'smooth'});
}

function initSideDrawerEvents() {
    initCommonActions();

    document.getElementById('act-clear')?.addEventListener('click', () => {
        clickedNode = null;
        searchTerm = '';
        const searchInput = document.getElementById('drawer-search-input');
        if (searchInput) searchInput.value = '';
        setSearchQuery('');
        updateVisualization(nodeGraph, linkGraph, labels);
        fitGraphToViewport(0.9);
        closeSideDrawer();
    });

    document.getElementById('toggle-decommissioned')?.addEventListener('change', (e) => {
        clickedNode = null;
        hideStoppedServices = !(e.target.checked);
        updateVisualization(nodeGraph, linkGraph, labels);
        //closeSideDrawer();
    });

    document.getElementById('act-fit')?.addEventListener('click', () => {
        fitGraphToViewport(0.9);
        closeSideDrawer();
    });

    document.getElementById('drawer-search-go')?.addEventListener('click', () => {
        const q = e.target.value ? e.target.value.trim() : "";
        handleQuery(q, false);
        //closeSideDrawer();
    });

    document.getElementById('drawer-search-input')?.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            const q = e.target.value?.trim();
            if (q) {
                handleQuery(q, false);
            }
            e.preventDefault();
            //closeSideDrawer();
        }
    });
}

window.addEventListener('DOMContentLoaded', initSideDrawerEvents);


window.addEventListener('DOMContentLoaded', () => {
    enableGlobalFindShortcut({
        inputSelector: '#drawer-search-input',
        onFocus: (input) => {
            // opzionale: svuota highlight precedenti o apri drawer
            // openDrawerIfClosed();
        }
    });
});

document.getElementById('closeDrawer').addEventListener('click', closeDrawer);
document.getElementById('overlay').addEventListener('click', closeDrawer);

function closeDrawer() {
    document.getElementById('drawer').classList.remove('open');
    document.getElementById('overlay').classList.remove('open');
}


document.getElementById('fileInput').addEventListener('change', function (event) {
    resetVisualization();
    const file = event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (e) {
        const csvData = e.target.result;
        const data = d3.csvParse(csvData);
        processData(data);
        updateVisualization(nodeGraph, linkGraph, labels);
    };
    reader.readAsText(file);
});

function activateInitialListViewIfNeeded() {
    const listViewParam = getQueryParam('listView');
    const sortParam = getQueryParam('sort');
    const parsedSort = parseSortParam(sortParam);
    if (parsedSort && (!window.currentColumnKeys.length || window.urrentColumnKeys.includes(parsedSort.key))) {
        sortKey = parsedSort.key;
        sortDir = parsedSort.dir;
    }
    if (listViewParam) {
        toListView();
    }
}

window.addEventListener('load', function () {
    let searchParam = null;
    const searchInput = document.getElementById('drawer-search-input');

    const listViewParam = getQueryParam('listView');
    const parsedCols = parseListViewParam(listViewParam);
    if (parsedCols && parsedCols.length) {
        window.currentColumnKeys = parsedCols;
    } else {
        window.currentColumnKeys = [...DEFAULT_COLUMN_KEYS];
    }

    fetch('https://francesconicolosi.github.io/itsm/sample_services.csv')
        .then(response => {
            searchParam = getQueryParam('search')
            if (searchParam) {
                searchTerm = searchParam;
                if (searchInput) searchInput.value = searchParam;
            }
            return response.text();
        })
        .then(csvData => {
            const data = d3.csvParse(csvData);
            processData(data);
            const afterInit = () => {
                const wantListView = Boolean(listViewParam);
                if (wantListView) {
                    toListView();
                    syncListViewParamInUrl();
                    syncSortParamInUrl();
                }
                const showDrawer = typeof searchParam === 'string' && uniqueIds.includes(searchParam.split(':')[0]);
                updateVisualization(nodeGraph, linkGraph, labels, showDrawer);

                if (wantListView && showDrawer) {
                    const id = searchParam.split(':')[1]?.replace(/"/g, '');
                    const node = nodes.find(n => n.id === id);
                    if (node) showNodeDetails(node, true);
                }
            };
            if (searchParam) {
                simulation.on('end', () => {
                    if (!hasLoaded) {
                        hasLoaded = true;
                        afterInit();
                    }
                });
            } else {
                afterInit();
            }
        })
        .catch(error => console.error('Error loading the CSV file:', error));
});

function processData(data) {
    const requiredColumns = ['Service Name', 'Description', 'Type', 'Depends on', 'Status', 'Decommission Date'];
    const missingColumns = requiredColumns.filter(col => !data.columns.includes(col));

    if (missingColumns.length > 0) {
        alert(`Missing mandatory columns: ${missingColumns.join(', ')}`);
        return;
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
    nodes = data.map(d => {
        const node = {id: d['Service Name'], color: colorScale(d['Type'])};
        for (const key in d) node[key] = d[key];
        return node;
    });
    const nodeIds = new Set(nodes.map(d => d.id));
    const usedByMap = new Map(); // dep -> Set(of services that depend on it)

    links = data.flatMap(d => {
        const src = d['Service Name'];
        const deps = splitValues(d['Depends on']); // <-- IMPORTANT: supports ||, \\n, ,
        return deps.map(dep => {
            const depTrim = dep.trim();
            if (!nodeIds.has(depTrim)) return null;

            // reverse mapping: dep is used by src
            if (!usedByMap.has(depTrim)) usedByMap.set(depTrim, new Set());
            usedByMap.get(depTrim).add(src);

            return {source: src, target: depTrim};
        });
    }).filter(Boolean);

    // ✅ 3) If "Used by" column is missing in input CSV, compute it
    const hasUsedByColumn = data.columns.includes('Used by');

    if (!hasUsedByColumn) {
        nodes.forEach(n => {
            const users = Array.from(usedByMap.get(n.id) || []);
            // output as || so it stays consistent with your new CSV standard
            n['Used by'] = users.join('||');
        });

        // opzionale: fai comparire la colonna anche nelle UI che leggono data.columns
        // (non sempre serve, ma è utile per coerenza e per list-view selector)
        data.columns = [...data.columns, 'Used by'];
    }

    activeServiceNodes = nodes.filter(d =>
        (d.Status !== 'Stopped' && d.Status !== 'Decommissioned' && !d['Decommission Date'])
    );
    activeServiceNodeIds = new Set(activeServiceNodes.map(d => d.id));

    createMap();
    createLegend(colorScale);
}

function getTermToCompare(term) {
    return term.replaceAll('\n', '').replaceAll(' ', '').toLowerCase();
}

function parseActiveKeyValueSearch(term) {
    if (!term || !term.includes(':')) return null;
    const raw = term.trim();
    const negated = raw.startsWith('!');
    if (negated) return null; // niente +/- su negazioni

    const idx = raw.indexOf(':');
    const key = raw.slice(0, idx).trim();
    const valuePart = raw.slice(idx + 1).trim();

    const quoted = valuePart.includes('"');
    const clean = valuePart.replaceAll('"', '');
    const values = splitValues(clean).map(v => v.trim()).filter(Boolean);

    return {key, values, quoted};
}

function normalizeForCompare(v) {
    return (v ?? '').toString().replaceAll('\n', '').replaceAll(' ', '').toLowerCase();
}

function buildKeyValueSearch(key, values, quoted) {
    if (!key || !values || !values.length) return '';
    const body = quoted
        ? values.map(v => `"${v}"`).join(',')
        : values.join(',');
    return `${key}:${body}`;
}

function updateSearchAndRefresh(q) {
    searchTerm = q || '';
    const searchInput = document.getElementById('drawer-search-input');
    if (searchInput) searchInput.value = searchTerm;
    setSearchQuery(searchTerm);
    updateVisualization(nodeGraph, linkGraph, labels, true);

    if (clickedNode) showNodeDetails(clickedNode, true);
}

function isSearchResultWithKeyValue(node) {
    if (!searchTerm.includes(":")) return false;
    const isNegation = searchTerm.trim().startsWith("!");
    const term = isNegation ? searchTerm.trim().slice(1) : searchTerm.trim();

    const isAccurateSearch = term.includes('"');
    const termClean = isAccurateSearch ? term.replaceAll('"', '') : term;

    const parts = termClean.split(':');
    if (parts.length !== 2) return false;

    const key = parts[0];
    if (!Object.keys(node).includes(key)) return false;

    const rawValue = parts[1].trim();

    if (isNegation && rawValue === "") {
        return (node[key] ?? "").trim() !== "";
    }

    const expectedValues = splitValues(rawValue)
        .map(v => getTermToCompare(v));

    const nodeParts = splitValues(node[key] ?? "");
    const matches = expectedValues.some(ev =>
        nodeParts.some(p =>
            isAccurateSearch
                ? getTermToCompare(p) === ev
                : getTermToCompare(p).includes(ev)
        )
    );
    return isNegation ? !matches : matches;
}


function isSearchResultValueOnly(d) {
    if (searchTerm === "" || searchTerm.includes(":")) return false;

    const terms = searchTerm.toLowerCase().split(',').map(term => term.trim());

    return Object.values(d).some(value =>
        typeof value === 'string' &&
        terms.some(term => value.toLowerCase().includes(term))
    );
}

const mapEl = document.getElementById('map');
const listViewEl = document.getElementById('list-view');
const legendEl = document.getElementById('legend');
const btnList = document.getElementById('view-list');
const btnGraph = document.getElementById('view-graph');

function toListView() {
    mapEl.style.display = 'none';
    if (legendEl) legendEl.style.display = 'none';
    listViewEl.style.display = 'block';

    btnList.style.display = 'none';
    btnGraph.style.display = 'inline-block';

    syncListViewParamInUrl();
    syncSortParamInUrl();

    renderListFromSearch();
}

function toGraphView() {
    mapEl.style.display = 'block';
    if (legendEl) legendEl.style.display = '';
    listViewEl.style.display = 'none';

    btnGraph.style.display = 'none';
    btnList.style.display = 'inline-block';

    const url = new URL(window.location.href);
    url.searchParams.delete('listView');
    url.searchParams.delete('sort');
    window.history.replaceState({}, '', url.toString());
}

btnList.addEventListener('click', toListView);
btnGraph.addEventListener('click', toGraphView);

function manageMultiPart(parts, td) {
    if (parts.length > 1) {
        const ul = document.createElement('ul');
        parts.forEach(p => {
            const li = document.createElement('li');
            li.innerHTML = isUrl(p) ? getLink(p) : p;
            ul.appendChild(li);
        });
        td.appendChild(ul);
    } else {
        td.innerHTML = getLink(parts[0]);
    }
}

function renderListFromSearch() {
    if (!currentNodes) {
        listViewEl.innerHTML = `<p class="empty-state">No data available.</p>`;
        return;
    }

    let results = currentNodes.filter(n => currentSearchedNodes?.has?.(n.id));

    const isListVisible = isListViewVisible();
    const noSearch = (searchTerm === "" || !searchTerm);
    if (isListVisible && noSearch && results.length === 0) {
        results = [...currentNodes];
    }

    listViewEl.innerHTML = '';

    if (sortKey) {
        results = results.slice().sort((a, b) => {
            const va = getComparableValue(a, sortKey);
            const vb = getComparableValue(b, sortKey);
            let cmp = 0;
            if (sortKey === 'Decommission Date' && typeof va === 'number' && typeof vb === 'number') {
                cmp = (va === vb ? 0 : (va < vb ? -1 : 1));
            } else {
                cmp = String(va).localeCompare(String(vb), undefined, {sensitivity: 'base', numeric: true});
            }
            return sortDir === 'asc' ? cmp : -cmp;
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
    table.style.setProperty('--cols', String(window.currentColumnKeys.length));

    const thead = document.createElement('thead');

    thead.addEventListener('click', (e) => {
        const btn = e.target.closest('.col-op');
        if (btn) {
            e.preventDefault();
            e.stopPropagation();
            const col = decodeURIComponent(btn.getAttribute('data-col'));
            toggleColumn(col);
            return;
        }
        const title = e.target.closest('.th-title');
        if (title) {
            const th = title.closest('th');
            const col = th?.getAttribute('data-col');
            if (!col) return;
            if (sortKey === col) {
                sortDir = (sortDir === 'asc' ? 'desc' : 'asc');
            } else {
                sortKey = col;
                sortDir = 'asc';
            }
            syncSortParamInUrl();
            renderListFromSearch();
        }
    });

    const trh = document.createElement('tr');

    window.currentColumnKeys.forEach(key => {
        const th = document.createElement('th');
        th.setAttribute('data-col', key);

        const cellWrap = document.createElement('div');
        cellWrap.className = 'th-cell';

        const title = document.createElement('button');  // focusable
        title.className = 'th-title fade-link';
        title.type = 'button';
        title.textContent = `${labelForKey(key)}${getSortIndicator(key)}`;

        const removeBtn = document.createElement('button');
        removeBtn.className = 'col-op fade-link';
        removeBtn.type = 'button';
        removeBtn.textContent = '−';
        removeBtn.setAttribute('data-col', encodeURIComponent(key));
        removeBtn.setAttribute('aria-label', `Remove "${labelForKey(key)}" from list view`);

        cellWrap.appendChild(title);
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
            clickedNode = n;
            if (window.d3 && window.labels) {
                labels.classed('highlight', d => d.id === n.id);
            }
            showNodeDetails(n, true);
        };

        tr.addEventListener('click', openDetails);
        tr.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                openDetails();
            }
        });

        window.currentColumnKeys.forEach(key => {
            const td = document.createElement('td');

            let raw = (key === 'id') ? (n.id ?? '') : (n[key] ?? '');
            if (key === 'Tracked Issues' && !raw) {
                const computed = computeTrackingSoftwareValue(n);
                if (computed) raw = computed;
            }

            if (typeof raw === 'string' && raw) {
                const parts = splitValues(raw);

                if (parts.some(p => isUrl(p))) {
                    manageMultiPart(parts, td);
                } else if (descriptionFields.includes(key)) {
                    td.innerHTML = "";
                    createFormattedLongTextElementsFrom(raw)
                        .forEach(el => td.appendChild(el));
                } else {
                    const uniqueParts = splitValues(raw);
                    td.textContent = uniqueParts.join(', ');
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

function focusNodeOnGraph(nodeId) {
    if (window.d3 && window.labels) {
        labels.classed('highlight', d => d.id === nodeId);
        // centerOnNode(nodeId);
    }
}


function updateVisualization(node, link, labels, showDrawer = true) {
    let relaxedSearchEnabled = document.getElementById('relaxed-search').checked;
    if (searchTerm !== "" && !searchTerm.includes(":") && !searchTerm.includes(",") && !relaxedSearchEnabled) {
        searchTerm = `${defaultSearchKey}:"${searchTerm}"`;
        const searchInput = document.getElementById('drawer-search-input');
        if (searchInput) searchInput.value = searchTerm;
        setSearchQuery(searchTerm);
    }

    const filteredLinks = links.filter(link => activeServiceNodeIds.has(link.source.id) && activeServiceNodeIds.has(link.target.id));
    const relatedNodes = new Set();
    const searchedNodes = new Set();
    const relatedLinks = links.filter(link => {
        let isLinkStatusOk = !hideStoppedServices || (filteredLinks.includes(link));
        let isSearchedLink = searchTerm === "";
        if (isSearchResultValueOnly(link.source) || isSearchResultWithKeyValue(link.source)) {
            isSearchedLink = isSearchedLink || true;
            searchedNodes.add(link.source.id);
        }

        if (isSearchResultValueOnly(link.target) || isSearchResultWithKeyValue(link.target)) {
            isSearchedLink = isSearchedLink || true;
            searchedNodes.add(link.target.id);
        }

        if (isLinkStatusOk && isSearchedLink) {
            relatedNodes.add(link.source.id);
            relatedNodes.add(link.target.id);
            return true;
        }
        return false;
    });

    let nodeToZoom;

    node.each(d => {
        const byKey = isSearchResultWithKeyValue(d);
        const byValue = relaxedSearchEnabled && isSearchResultValueOnly(d);
        if (byKey || byValue) {
            nodeToZoom = nodeToZoom || d;
            relatedNodes.add(d.id);
            searchedNodes.add(d.id);
        }
    });

    node.style('display', d => (searchTerm === "" && !hideStoppedServices) || (searchTerm === "" && hideStoppedServices && activeServiceNodeIds.has(d.id)) || relatedNodes.has(d.id) && (!hideStoppedServices || activeServiceNodeIds.has(d.id)) ? 'block' : 'none');
    link.style('display', d => (searchTerm === "" && !hideStoppedServices) || (searchTerm === "" && hideStoppedServices && activeServiceNodeIds.has(d.source.id) && activeServiceNodeIds.has(d.target.id)) || relatedLinks.includes(d) ? 'block' : 'none');
    labels.style('display', d => (searchTerm === "" && !hideStoppedServices) || (searchTerm === "" && hideStoppedServices && activeServiceNodeIds.has(d.id)) || relatedNodes.has(d.id) && (!hideStoppedServices || activeServiceNodeIds.has(d.id)) ? 'block' : 'none');
    labels.style('text-decoration', d => searchedNodes.has(d.id) ? 'underline' : 'none');
    currentNodes = nodes;
    currentSearchedNodes = searchedNodes;


    if (document.getElementById('list-view')?.style.display === 'block') {
        renderListFromSearch();
    } else if (!clickedNode && nodeToZoom && (!hideStoppedServices || activeServiceNodeIds.has(nodeToZoom.id))) {
        centerAndZoomOnNode(nodeToZoom);
        showNodeDetails(nodeToZoom, showDrawer);
    }
}

function zoomed({transform}) {
    g.attr("transform", transform);
}

function getPeopleDbLink(value) {
    return `<a href="solitaire.html?search=${encodeURIComponent(value.toLowerCase()).replace(/%20/g, '+')}" target = "_blank" >${value}</a>`;
}


function getLink(value) {
    let cleanValue = value.replace(/^https?:\/\//, '');
    cleanValue = cleanValue.split(/[?#]/)[0];
    const segments = cleanValue.split('/').filter(Boolean);
    const segment = segments.length > 0
        ? segments[segments.length - 1] || segments[segments.length - 2] || ''
        : '';
    const fixedLength = 55;
    const formattedValue = segment.length > fixedLength
        ? '...' + segment.slice(-fixedLength)
        : segment;
    return `<a href="${value}" target="_blank">${formattedValue}</a>`;
}

function computeTrackingSoftwareValue(node) {
    if (!node?.id) return '';
    const rawId = (node.id ?? '').toLowerCase().trim();
    const noSpaces = rawId.replace(/\s+/g, '');
    const noPunct = noSpaces.replace(/[^\w]/g, '');
    const keepHyphen = noSpaces.replace(/[^\w-]/g, '');
    const hyphenToUnderscore = keepHyphen
        .replace(/-/g, '_')
        .replace(/_+/g, '_')
        .replace(/^_+|_+$/g, '');

    const values = Array.from(new Set([noPunct, keepHyphen, hyphenToUnderscore].filter(Boolean)));
    const inList = values.map(v => `"${v}"`).join(', ');

    const jql = `
   (
     project = "Company Managed Services Support" AND statusCategory in (EMPTY, "To Do", "In Progress")
     OR
     project = GDT AND statusCategory in (EMPTY, "To Do", "In Progress")
     AND labels in (bug-from-incident, from_l1_portal) AND issuetype = Bug
   )
   AND "Theme[Checkboxes]" in (App, "Brand & Content", Krypto, Content, Cross, Omni, "Product Discovery", Purchase, Loyalty, "IT 4 IT")
   AND cf[14139] in (${inList})
   ORDER BY created ASC
 `.replace(/\s+/g, ' ').trim();
    return `https://nycosoft.trackingsowftare.net/issues/?jql=${encodeURIComponent(jql)}`;
}


function showNodeDetails(node, openDrawer = true) {

// ---------- Key detection + normalization ----------
    const keyRaw = String(node['Key'] ?? '').trim();
    const serviceRaw = String(node['Service Name'] ?? '').trim();
    const idRaw = String(node.id ?? '').trim();

    const keyFromCsv = keyRaw !== '';

    const keyValue = keyFromCsv
        ? keyRaw
        : (serviceRaw || idRaw);

// garantisci Key sempre valorizzato
    node['Key'] = keyValue;

// normalizzazione per confronti
    const norm = v => v.toLowerCase();
    const keyNorm = norm(keyValue);
    const idNorm = norm(idRaw);
    const serviceNorm = norm(serviceRaw);

    const keyEqualsId = keyNorm && keyNorm === idNorm;

    const PRIORITY_KEYS = [
        'Key',
        ...(!keyEqualsId ? ['id'] : []),
        'Description',
        'Depends on',
        'Used by'
    ];


    const drawer = document.getElementById('drawer');
    const overlay = document.getElementById('overlay');
    const drawerContent = document.getElementById('drawerContent');


    const title = drawer.querySelector('.drawer-header h2');
    title.textContent = node['Service Name'] || 'Service Information';

    drawerContent.innerHTML = '';

    if (!keyFromCsv) {
        // If Key column is missing/empty, use Service Name or internal id
        node['Key'] = String(node['Service Name'] ?? node.id ?? '').trim();
    }

    const excluded = new Set([
        'index','x','y','vy','vx','fx','fy','color',

        // Service Name mai mostrato come riga
        'Service Name',

        // se Key == id → nascondi id
        ...(keyEqualsId ? ['id'] : [])
    ]);

    const isListVisible = isListViewVisible();
    const table = document.createElement('table');
    const renderedKeys = new Set();

    const renderValueCell = (key, raw) => {
        const td = document.createElement('td');
        if (typeof raw !== 'string') return td;

        if (isDateTimeValue(raw)) {
            td.textContent = formatDateTimeLocal(raw);
            td.title = raw; // tooltip con valore UTC originale
            return td;
        }

        // long text
        if (descriptionFields.includes(key)) {
            createFormattedLongTextElementsFrom(raw).forEach(el => td.appendChild(el));
            return td;
        }

        const parts = splitValues(raw);

        // urls
        if (parts.some(isUrl)) {
            manageMultiPart(parts, td);
            return td;
        }

        // people db links
        if (searchableAttributesOnPeopleDb.includes(key)) {
            td.innerHTML = `<i>${
                parts.length > 1
                    ? `<ul>${parts.map(v => `<li>${getPeopleDbLink(v)}</li>`).join('')}</ul>`
                    : getPeopleDbLink(parts[0] || '')
            }</i>`;
            return td;
        }

        // ===== toggle (+ / −) logic only if there is an active key:value search =====
        const active = parseActiveKeyValueSearch(searchTerm); // {key, values, quoted} | null
        const isSameKey = !!active && active.key === key;
        const activeVals = new Set((active?.values || []).map(normalizeForCompare));

        const makeToggleBtn = (v) => {
            if (!isSameKey) return '';
            const inSearch = activeVals.has(normalizeForCompare(v));
            const cls = inSearch ? 'search-remove' : 'search-add';
            const sym = inSearch ? '−' : '+';
            return ` <a class="fade-link search-toggle ${cls}"
                data-key="${encodeURIComponent(key)}"
                data-value="${encodeURIComponent(v)}"
                href="#">${sym}</a>`;
        };

        // ===== render values =====
        if (parts.length > 1) {
            const ul = document.createElement('ul');
            parts.forEach(v => {
                const li = document.createElement('li');
                li.innerHTML = `<i>${v}
        <a class="fade-link search-trigger"
           data-key="${encodeURIComponent(key)}"
           data-value="${encodeURIComponent(v)}"
           href="#">⌞ ⌝</a>${makeToggleBtn(v)}
      </i>`;
                ul.appendChild(li);
            });
            td.appendChild(ul);
        } else {
            const v = parts[0] || '';
            td.innerHTML = `<i>${v}
      <a class="fade-link search-trigger"
         data-key="${encodeURIComponent(key)}"
         data-value="${encodeURIComponent(v)}"
         href="#">⌞ ⌝</a>${makeToggleBtn(v)}
    </i>`;
        }

        return td;
    };
    const renderKeyCell = (key) => {
        const td = document.createElement('td');
        const colKey = key === 'Service Name' ? 'id' : key;
        const keyLabel = document.createElement('span');
        keyLabel.textContent = key;
        td.innerHTML = '';
        if (isListVisible) {
            td.appendChild(keyLabel);
            const selected = window.currentColumnKeys.includes(colKey);
            const btn = document.createElement('button');
            btn.className = 'col-op fade-link';
            btn.type = 'button';
            btn.setAttribute('data-col', encodeURIComponent(colKey));
            btn.setAttribute('aria-label',
                selected ? `Remove "${labelForKey(colKey)}" from list view`
                    : `Add "${labelForKey(colKey)}" to list view`);
            btn.textContent = selected ? '−' : '+';
            td.appendChild(btn);
        } else {
            td.appendChild(keyLabel);
        }
        return td;
    };

    const renderRow = (key, value) => {
        if (renderedKeys.has(key)) return;           // dedup
        if (excluded.has(key)) return;
        if (typeof value !== 'string' || !value) return;
        const tr = document.createElement('tr');
        tr.appendChild(renderKeyCell(key));
        tr.appendChild(renderValueCell(key, value));
        table.appendChild(tr);
        renderedKeys.add(key);
    };
    const nodeKeys = Object.keys(node);

    const normalizeKeyForOrder = (k) =>
        k === 'Service Name' ? 'id' : k;

    const orderedKeys = PRIORITY_KEYS
        .map(pk => nodeKeys.find(k => k === pk))
        .filter(Boolean);
    const remainingKeys = nodeKeys.filter(
        k => !orderedKeys.includes(k)
    );

    [...orderedKeys, ...remainingKeys].forEach(key => {
        renderRow(key, node[key]);
    });
    serviceInfoEnhancers
        .map(fn => fn(node))
        .filter(r => r && r.key && r.value && !renderedKeys.has(r.key))
        .forEach(r => renderRow(r.key, r.value));
    table.addEventListener('click', (e) => {
        const btn = e.target.closest('button.col-op');
        if (!btn) return;
        e.stopPropagation();
        const col = decodeURIComponent(btn.getAttribute('data-col'));
        toggleColumn(col);
        refreshDrawerColumnIcons();
    });

    drawerContent.appendChild(table);
    refreshDrawerColumnIcons();

    if (openDrawer) {
        drawer.classList.add('open');
        overlay.classList.add('open');
    }
}


function createMap() {

    zoom = d3.zoom()
        .scaleExtent([0.1, 3])
        .on("zoom", zoomed);

    svg = d3.select('#map').append('svg')
        .attr('width', width)
        .attr('height', height)
        .call(zoom);
    g = svg
        .append('g');

    g.append('defs').append('marker')
        .attr('id', 'arrow')
        .attr('viewBox', '0 -5 10 10')
        .attr('refX', 15)
        .attr('refY', 0)
        .attr('markerWidth', 10)
        .attr('markerHeight', 10)
        .attr('orient', 'auto')
        .append('path')
        .attr('d', 'M0,-5L10,0L0,5')
        .attr('fill', '#999');

    simulation = d3.forceSimulation(nodes)
        .force('link', d3.forceLink(links).id(d => d.id).distance(200))
        .force('charge', d3.forceManyBody().strength(-300))
        .force('center', d3.forceCenter(width / 2, height / 2));

    if (getQueryParam('search'))
        simulation.alphaDecay(0.07);

    linkGraph = g.append('g')
        .selectAll('line')
        .data(links)
        .enter().append('line')
        .attr('marker-end', 'url(#arrow)');

    nodeGraph = g.append('g')
        .selectAll('circle')
        .data(nodes)
        .enter().append('circle')
        .attr('r', 20)
        .attr('fill', d => d.color)
        .call(d3.drag()
            .on('start', dragstarted)
            .on('drag', dragged)
            .on('end', dragended))
        .on('mouseover', function (event, d) {
            const tooltip = d3.select('#tooltip');
            tooltip.transition().duration(200).style('opacity', .9);
            tooltip.html(d['Description'] || 'No description available')
                .style('left', (event.pageX + 5) + 'px')
                .style('top', (event.pageY - 28) + 'px');
        })
        .on('mouseout', mouseout)
        .on('click', function (event, d) {
            clickedNode = d;
            showNodeDetails(d);
        });
    labels = g.append('g')
        .selectAll('text')
        .data(nodes)
        .enter().append('text')
        .attr('dy', -2)
        .attr('text-anchor', 'middle')
        .text(d => d.id);

    simulation.on('tick', () => {
        linkGraph
            .attr('x1', d => d.source.x)
            .attr('y1', d => d.source.y)
            .attr('x2', d => {
                const dx = d.target.x - d.source.x;
                const dy = d.target.y - d.source.y;
                const dist = Math.sqrt(dx * dx + dy * dy);
                const offsetX = (dx / dist) * 5;
                return d.target.x - offsetX;
            })
            .attr('y2', d => {
                const dx = d.target.x - d.source.x;
                const dy = d.target.y - d.source.y;
                const dist = Math.sqrt(dx * dx + dy * dy);
                const offsetY = (dy / dist) * 5;
                return d.target.y - offsetY;
            });
        nodeGraph
            .attr('cx', d => d.x)
            .attr('cy', d => d.y);
        labels
            .attr('x', d => d.x)
            .attr('y', d => d.y - 30);
    });

    zoomIdentity = d3.zoomIdentity;

    function dragstarted(event, d) {
        if (!event.active) simulation.alphaTarget(0.3).restart();
        d.fx = d.x;
        d.fy = d.y;
    }

    function dragged(event, d) {
        d.fx = event.x;
        d.fy = event.y;
    }

    function dragended(event, d) {
        if (!event.active) simulation.alphaTarget(0);
        d.fx = d.x;
        d.fy = d.y;
    }

    function mouseout() {
        const tooltip = d3.select('#tooltip');
        tooltip.transition().duration(500).style('opacity', 0);
    }

    document.addEventListener('click', function (e) {
        const trigger = e.target.closest('.search-trigger');
        const addBtn  = e.target.closest('.search-add');
        const remBtn  = e.target.closest('.search-remove');

        // 1) Click su ⌞ ⌝ = comportamento attuale
        if (trigger) {
            clickedNode = null;
            e.preventDefault();

            const key = decodeURIComponent(trigger.getAttribute('data-key'));
            const isAccurateSearch = key === "Depends on" || key === "Used by" || key === "id";
            const mappedKey = isAccurateSearch ? "id" : key;

            const value = isAccurateSearch
                ? `"${decodeURIComponent(trigger.getAttribute('data-value'))}"`
                : `${decodeURIComponent(trigger.getAttribute('data-value'))}`;

            const combinedSearchTerm = `${mappedKey}:${value}`;
            updateSearchAndRefresh(combinedSearchTerm);
            window.scrollTo({top: 0, behavior: 'smooth'});
            return;
        }

        // 2) Click su + / −
        if (addBtn || remBtn) {
            e.preventDefault();
            e.stopPropagation();

            const btn = addBtn || remBtn;
            const key = decodeURIComponent(btn.getAttribute('data-key'));
            const value = decodeURIComponent(btn.getAttribute('data-value'));

            const active = parseActiveKeyValueSearch(searchTerm);
            if (!active || active.key !== key) return; // solo se search key:value attiva e stessa key

            const values = [...active.values];
            const needle = normalizeForCompare(value);

            const idx = values.findIndex(v => normalizeForCompare(v) === needle);

            if (addBtn) {
                if (idx === -1) values.push(value);
            } else {
                if (idx !== -1) values.splice(idx, 1);
            }

            const next = buildKeyValueSearch(active.key, values, active.quoted);
            updateSearchAndRefresh(next);
            return;
        }
    });
}

function createLegend(colorScale) {
    const types = colorScale.domain();
    const legend = d3.select('#legend');
    types.forEach(type => {
        const color = colorScale(type);
        const legendItem = legend.append('div').attr('class', 'legend-item');
        legendItem.append('div').attr('class', 'legend-swatch').style('background-color', color);
        legendItem.append('span').text(type);
    });
}
