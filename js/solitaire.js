import * as d3 from 'd3';

import {
    applyStreamVisibility,
    askHideStreamModal,
    applySearchDimmingForMatches,
    buildCompositeKey,
    buildFallbackMailToLink,
    buildLegendaColorScale,
    clearFieldHighlights,
    clearSearchDimming,
    closeSideDrawer,
    collectMembersFromOrganization,
    filterOrganizationByStreams,
    createFormattedLongTextElementsFrom,
    createHrefElement,
    createOutlookUrl,
    formatMonthYear,
    getAllowedStreamsSet,
    getFormattedDate,
    getNameFromTitleEl,
    getQueryParam,
    highlightGroup as highlightGroupUtils,
    initCommonActions,
    normalizeKey,
    openPersonReportCompose,
    parseCSV,
    SECOND_LEVEL_LABEL_EXTRA,
    setSearchQuery,
    TEAM_MEMBER_LEGENDA_LABEL,
    truncateString,
    countTeamsForMemberInOrg,
    firstOrgLevel,
    secondOrgLevel,
    thirdOrgLevel,
    firstLevelNA,
    secondLevelNA,
    thirdLevelNA,
    normalizeWs,
    makeKeyColorScale,
    getLegendTitleFor,
    computeKeysAndCountsFromVisibleOrg,
    computeStreamBoxWidthByCapacity,
    MAX_TEAMS_PER_ROW, splitValues, NEUTRAL_COLOR,
    ROLE_FIELD_WITH_MAPPING,
    COMPANY_FIELD, LOCATION_FIELD, emailField
} from './utils.js';

let lastSearch = '';
let currentIndex = 0;
let logoLayer;

let people = [];
let colorScale = null;
let cachedCsvText = null;

function resetStreamVisibility() {
    setStreamFilter(null);
}


const secondLevelRowPadY = 60;

const UNKNOWN_MATCHER = /^(unknown|n\/?a|not\s*(set|available)|-|—|none)$/i;
function isUnknownLegendKey(v) {
    const s = (v ?? '').toString().trim();
    return !s || UNKNOWN_MATCHER.test(s);
}

let visibleOrganizationWithManagers = null;

const STREAM_ORDER = [
    'Digital Design',
    'Digital Enablers',
    'Upper Funnel',
    'AI Enablement',
    'E-Commerce Core',
    'Lower Funnel',
    'Order Management System'
].map(s => normalizeKey(s));

let colorBy = ROLE_FIELD_WITH_MAPPING;

const guestRolesMap = new Map([
    ["Team Product Manager", ["Product Manager"]],
    ["Team Delivery Manager", ["Delivery Manager"]],
    ["Team Scrum Master", ["Agile Coach/Scrum Master"]],
    ["Team Solution Architect", ["Solution Architect"]],
    ["Team Development Manager", ["Development Manager"]],
    ["Team Security Champion", ["Security Champion"]]
]);

const guestRoleColumns = Array.from(guestRolesMap.keys());

let colorKeyMappings = new Map();

const peopleDBUpdateRecipients = [
    'teams@share.software.net'
];

const portfolioDBUpdateRecipients = ['portfolio@nycosoft.com', 'bleiz.jonas@nycosoft.com'];

let roleDetailsMapping;
const seenLegendClickKeys = new Set();
let lastLegendClickAt = 0;

function getCardFill(g) {
    if (typeof colorScale !== 'function') return NEUTRAL_COLOR;

    let colorKey;
    if (colorBy === ROLE_FIELD_WITH_MAPPING) {
        colorKey = normalizeWs(g.attr('data-role')) || TEAM_MEMBER_LEGENDA_LABEL;
    } else if (colorBy === COMPANY_FIELD) {
        colorKey = (g.attr('data-company') || 'Unknown');
    } else {
        colorKey = (g.attr('data-location') || 'Unknown');
    }

    const finalColor = colorScale(colorKey);
    return (typeof finalColor === 'string' && finalColor) ? finalColor : NEUTRAL_COLOR;
}

function renderLegendAll({ title, fieldName = LOCATION_FIELD, keys, counts, topKey, colorOf, maxVisible = 11 }) {
    let root = document.getElementById('legend-root');
    if (!root) {
        root = document.createElement('div');
        root.id = 'legend-root';
        document.body.appendChild(root);
    }
    root.className = 'legend legend--generic';
    root.innerHTML = `
    <div class="legend__title"></div>
    <div class="legend__list" aria-label="Legend list"></div>
  `;
    root.querySelector('.legend__title').textContent = title;
    const list = root.querySelector('.legend__list');

    keys.forEach((key) => {
        const item = document.createElement('div');
        item.className = 'legend__item';
        item.setAttribute('data-value', key);
        item.setAttribute('data-field', colorBy);

        const disabled = isUnknownLegendKey(key);
        if (disabled) {
            item.classList.add('legend__item--disabled');
            item.setAttribute('aria-disabled', 'true');
        } else {
            item.setAttribute('role', 'button');
            item.setAttribute('tabindex', '0');
            item.setAttribute('aria-label', `Filter by ${key}`);
        }

        const sw = document.createElement('span');
        sw.className = 'legend__swatch';
        const color = colorOf.colorOf(key);
        sw.style.backgroundColor = color;
        if ((color || '').toLowerCase() === '#ffffff' || color === 'white') {
            sw.classList.add('legend__swatch--white');
        }

        const label = document.createElement('span');
        label.className = 'legend__label';
        label.textContent = key;

        const count = document.createElement('span');
        count.className = 'legend__count';
        count.textContent = counts.get(key) ?? '';

        item.append(sw, label, count);
        list.appendChild(item);
    });

    const isUnknownKey = (el) => {
        const v = (el.getAttribute('data-value') || '').trim();
        return isUnknownLegendKey(v);
    };

    const activate = (el) => {
        const value = el.getAttribute('data-value') ?? '';
        const field = (el.getAttribute('data-field') || '').trim(); // 'role' | 'company' | 'location' |
        const missing = isUnknownKey(el);
        const searchInput = document.getElementById('drawer-search-input');
        if (searchInput) searchInput.value = missing ? '' : value;
        const normalizedValue = normalizeWs(value).toLowerCase();
        const normalizedField = normalizeWs(field).toLowerCase();
        const clickKey = `${normalizedField}::${normalizedValue}`;

        let noZoom = false;
        if (!missing) {
            const now = Date.now();
            const elapsed = now - (lastLegendClickAt || 0);

            if (elapsed > 1000) {
                noZoom = true;
            }

            lastLegendClickAt = now;
        }

        searchByQuery(missing ? '' : value, { field, missing, noZoom });
    };

    list.addEventListener('click', (e) => {
        const el = e.target.closest('.legend__item');
        if (!el) return;
        activate(el);
    });

    list.addEventListener('keydown', (e) => {
        if (e.key !== 'Enter' && e.key !== ' ') return;
        const el = e.target.closest('.legend__item');
        if (!el) return;
        e.preventDefault();
        activate(el);
    });

    list.style.setProperty('--legend-row', '24px');
    list.style.maxHeight = `calc(${maxVisible} * var(--legend-row))`;
    enableLegendDrag({ handleSelector: '.legend__title' });
}

function recolorProfileCards(field) {
    colorBy = field;

    const allowedStreams = getAllowedStreamsSet?.() ?? null;
    const orgForLegend = filterOrganizationByStreams(visibleOrganizationWithManagers, allowedStreams);
    const fieldName = colorBy;
    const { keys, counts, topKey } = computeKeysAndCountsFromVisibleOrg(orgForLegend, fieldName);

    colorScale = makeKeyColorScale(keys, topKey);
    renderLegendAll({
        title: getLegendTitleFor(fieldName),
        keys,
        counts,
        topKey,
        colorOf: colorScale,
        maxVisible: 11
    });

    d3.selectAll('g[data-key^="card::"]').each(function () {
        const g = d3.select(this);
        const fill = getCardFill(g) || NEUTRAL_COLOR;
        const rect = g.select('rect.profile-box');
        rect.transition().duration(200).attr('fill', fill);
        if ((fill || '').toLowerCase() === '#ffffff' || fill === 'white') {
            rect.attr('stroke', '#b8b8b8').attr('stroke-width', 1);
        } else {
            rect.attr('stroke', null).attr('stroke-width', null);
        }
    });

}


function setColorMode(mode) {
    const roleEl = document.getElementById('toggle-color-role');
    const compEl = document.getElementById('toggle-color-company');
    const locEl = document.getElementById('toggle-color-location');

    if (!roleEl || !compEl || !locEl) return;


    if (mode === ROLE_FIELD_WITH_MAPPING) {
        roleEl.checked = true;
        compEl.checked = false;
        locEl.checked = false;
    } else if (mode === COMPANY_FIELD) {
        roleEl.checked = false;
        compEl.checked = true;
        locEl.checked = false;
    } else {
        if (mode === LOCATION_FIELD) {
            roleEl.checked = false;
            compEl.checked = false;
            locEl.checked = true;
        }
    }

    recolorProfileCards(mode);
}


const drag = d3.drag()
    .on("start", function () {
        d3.select(this).raise();
    })
    .on("drag", function (event) {
        const transform = d3.select(this).attr("transform");
        const translate = transform.match(/translate\(([^,]+),([^\)]+)\)/);
        const x = parseFloat(translate[1]) + event.dx;
        const y = parseFloat(translate[2]) + event.dy;
        d3.select(this).attr("transform", `translate(${x},${y})`);
    });


let searchParam;

let svg;
let viewport;
let backgroundLayer;
let cardLayer;
let streamLayer;
let themeLayer;
let teamLayer;

let zoom;
let width = 1200;
let height = 800;

function findHeaderIndex(headers, name) {
    const target = (name || '').trim().toLowerCase();
    return headers.findIndex(h => (h || '').trim().toLowerCase() === target);
}

let LS_KEY = 'dsm-layout-v1:default';

function loadLayout() {
    try {
        return JSON.parse(localStorage.getItem(LS_KEY) || '{}');
    } catch {
        return {};
    }
}

function saveLayout(obj) {
    localStorage.setItem(LS_KEY, JSON.stringify(obj));
}

function getItemLayout(key) {
    return loadLayout()[key];
}

function upsertItemLayout(key, patch) {
    const all = loadLayout();
    all[key] = {...(all[key] || {}), ...patch};
    saveLayout(all);
}

function parseTranslate(transform) {
    if (!transform) return {x: 0, y: 0};
    const m = transform.match(/translate\(([^,]+),\s*([^)]+)\)/);
    return m ? {x: +m[1] || 0, y: +m[2] || 0} : {x: 0, y: 0};
}

function restoreGroupPosition(groupSel) {
    const key = groupSel.attr('data-key');
    if (!key) return false;
    const saved = getItemLayout(key);
    if (!saved || !Number.isFinite(saved.x) || !Number.isFinite(saved.y)) return false;
    groupSel.attr('transform', `translate(${saved.x},${saved.y})`);
    return true;
}

function getSavedSizeForGroup(groupSel) {
    const key = groupSel.attr('data-key');
    if (!key) return null;
    const saved = getItemLayout(key);
    if (!saved || !Number.isFinite(saved.width) || !Number.isFinite(saved.height)) return null;
    return {w: saved.width, h: saved.height};
}


function makeResizable(group, rect, opts = {}) {
    const minW = Number(opts.minWidth) || 200;
    const minH = Number(opts.minHeight) || 150;

    const title = group.select('text');

    const savedSize = getSavedSizeForGroup(group);
    let w = (savedSize?.w ?? Number(rect.attr('width'))) || minW;
    let h = (savedSize?.h ?? Number(rect.attr('height'))) || minH;

    const handleSize = 14;
    const hitPad = 10;

    const handles = group.append('g').attr('class', 'resize-handles');
    handles.raise();

    const handleE = handles.append('rect').attr('class', 'resize-handle e');
    const handleS = handles.append('rect').attr('class', 'resize-handle s');
    const handleSE = handles.append('rect').attr('class', 'resize-handle se');

    const hitE = handles.append('rect').attr('class', 'resize-hit e');
    const hitS = handles.append('rect').attr('class', 'resize-hit s');
    const hitSE = handles.append('rect').attr('class', 'resize-hit se');

    function positionHandles() {
        handleE
            .attr('x', w - handleSize / 2)
            .attr('y', h / 2 - handleSize / 2)
            .attr('width', handleSize)
            .attr('height', handleSize);

        handleS
            .attr('x', w / 2 - handleSize / 2)
            .attr('y', h - handleSize / 2)
            .attr('width', handleSize)
            .attr('height', handleSize);

        handleSE
            .attr('x', w - handleSize / 2)
            .attr('y', h - handleSize / 2)
            .attr('width', handleSize)
            .attr('height', handleSize);

        hitE
            .attr('x', w - (handleSize / 2 + hitPad))
            .attr('y', h / 2 - (handleSize / 2 + hitPad))
            .attr('width', handleSize + 2 * hitPad)
            .attr('height', handleSize + 2 * hitPad);

        hitS
            .attr('x', w / 2 - (handleSize / 2 + hitPad))
            .attr('y', h - (handleSize / 2 + hitPad))
            .attr('width', handleSize + 2 * hitPad)
            .attr('height', handleSize + 2 * hitPad);

        hitSE
            .attr('x', w - (handleSize / 2 + hitPad))
            .attr('y', h - (handleSize / 2 + hitPad))
            .attr('width', handleSize + 2 * hitPad)
            .attr('height', handleSize + 2 * hitPad);
    }

    function applySize() {
        rect.attr('width', w).attr('height', h);

        if (!title.empty()) {
            const anchor = title.attr('text-anchor');
            if (anchor === 'middle') {
                title.attr('x', w / 2);
            }
        }

        positionHandles();
        if (typeof opts.onResize === 'function') {
            opts.onResize({width: w, height: h});
        }
    }

    function makeDeltaTracker() {
        let prev = null;
        const getSvgPoint = (event) => {
            const t = d3.zoomTransform(svg.node());
            const [px, py] = d3.pointer(event, svg.node());
            return t.invert([px, py]);
        };
        return {
            start(event) {
                prev = getSvgPoint(event);
            },
            drag(event) {
                const curr = getSvgPoint(event);
                if (!prev) prev = curr;
                const dx = curr[0] - prev[0];
                const dy = curr[1] - prev[1];
                prev = curr;
                return {dx, dy};
            }
        };
    }

    const trackerE = makeDeltaTracker();
    const trackerS = makeDeltaTracker();
    const trackerSE = makeDeltaTracker();

    const dragE = d3.drag()
        .on('start', (event) => {
            event.sourceEvent?.stopPropagation();
            trackerE.start(event);
        })
        .on('drag', (event) => {
            const {dx} = trackerE.drag(event);
            w = Math.max(minW, w + dx);
            applySize();
        });

    const dragS = d3.drag()
        .on('start', (event) => {
            event.sourceEvent?.stopPropagation();
            trackerS.start(event);
        })
        .on('drag', (event) => {
            const {dy} = trackerS.drag(event);
            h = Math.max(minH, h + dy);
            applySize();
        });

    const dragSE = d3.drag()
        .on('start', (event) => {
            event.sourceEvent?.stopPropagation();
            trackerSE.start(event);
        })
        .on('drag', (event) => {
            const {dx, dy} = trackerSE.drag(event);
            w = Math.max(minW, w + dx);
            h = Math.max(minH, h + dy);
            applySize();
        });

    handleE.call(dragE);
    hitE.call(dragE);
    handleS.call(dragS);
    hitS.call(dragS);
    handleSE.call(dragSE);
    hitSE.call(dragSE);

    handles
        .style('display', isDraggable ? null : 'none')
        .style('pointer-events', isDraggable ? 'all' : 'none');

    applySize();
}


function aggregateInfoByHeader(members, headers, headerName = 'Team Managed Services', sortElements = false) {
    const idx = findHeaderIndex(headers, headerName);
    if (idx === -1) {
        return {exists: false, items: []};
    }
    const headerRealName = headers[idx];
    const set = new Set();

    members.forEach(m => {
        const raw = m[headerRealName];
        if (!raw) return;
        splitValues(raw).forEach(v => set.add(v));
    });

    const itemsToReturn = sortElements
        ? [...set].sort((a, b) => a.localeCompare(b, 'it', { sensitivity: 'base' }))
        : [...set];

    return {
        exists: true,
        items: itemsToReturn
    };
}

function clearSearch() {
    const output = document.getElementById('output');
    output.textContent = '';
    searchParam = '';
    const searchInput = document.getElementById('drawer-search-input');
    searchInput.value = searchParam;
    setSearchQuery(searchParam);
    clearSearchDimming();
    clearFieldHighlights();
    fitToContent(0.9);
    closeDrawer();
    //closeSideDrawer();
}

function isInternalCompany(member) {
    return ((member[COMPANY_FIELD] || '').trim().toLowerCase() === 'internal');
}

function initSideDrawerEvents() {
    initCommonActions();

    document.getElementById('act-clear')?.addEventListener('click', () => {
        resetStreamVisibility();
        clearSearch();
    });

    document.getElementById('act-fit')?.addEventListener('click', () => {
        fitToContent(0.9);
        //closeSideDrawer();
    });

    document.getElementById('toggle-color-role')?.addEventListener('change', (e) => {
        if (e.target.checked) setColorMode(ROLE_FIELD_WITH_MAPPING);
    });
    document.getElementById('toggle-color-company')?.addEventListener('change', (e) => {
        if (e.target.checked) setColorMode(COMPANY_FIELD);
    });
    document.getElementById('toggle-color-location')?.addEventListener('change', (e) => {
        if (e.target.checked) setColorMode(LOCATION_FIELD);
    });

    document.getElementById('act-about')?.addEventListener('click', (e) => {
        closeSideDrawer();
        openDrawer({
            name: "About Solitaire ♤", description:
                `Org charts highlight hierarchy—but not how teams actually work. Much of the real collaboration that drives the Company operations happens across functions, services, and roles, yet remains invisible. This reinforces silos and hides the complexity of our shared work. More info on <a href="https://www.gamerdad.cloud/" target="_blank">my personal blog</a>\n` +
                "\n" +
                `<b><i>Our Vision</b></i>\n` +
                "By visualizing how teams operate—the people, services, and responsibilities behind daily activities—we strengthen a culture that is collaborative, transparent, and service‑oriented. Visibility turns shared accountability into a tangible part of our operating model.\n" +
                "\n" +
                `<b><i>What we're building</b></i>\n` +
                "A custom Visual People Database that brings together data from several systems into a single, interactive view.\n" +
                "\n" +
                `<b><i>It provides:</b></i>\n` +
                `<ul>` +
                `<li>A clear map of team members (internal staff and suppliers)</li>` +
                `<li>The services each team manages</li>` +
                `<li>Roles and responsibilities across the organization</li>` +
                `<li>Quick access to Domino Service Catalog</li>` +
                `<li>A built‑in “Request an update” feature to keep information fresh and accurate</li></ul>` +
                "\n" +
                "<b><i>The Benefits</b></i>\n" +
                `<ul><li>Understand who works on what across projects and services</li>` +
                `<li>Make hidden operational networks visible</li>` +
                `<li>Consolidate data not available in systems like the one used by the HR</li>` +
                `<li>Strengthen transparency, alignment, and cross‑team collaboration</li>` +
                `<li>Provide a single source of truth for service ownership and responsibilities</li></ul>`
        });
    });


    document.getElementById('act-report')?.addEventListener('click', async () => {
        try {
            openPersonReportCompose( peopleDBUpdateRecipients, portfolioDBUpdateRecipients).then(r =>  closeSideDrawer());
        } catch (e) {
            console.log(e);
            buildFallbackMailToLink(peopleDBUpdateRecipients, subject, body);
        }
    });

    document.getElementById('drawer-search-go')?.addEventListener('click', () => {
        const q = document.getElementById('drawer-search-input')?.value?.trim().toLowerCase();
        if (q) searchByQuery(q);
        //closeSideDrawer();
    });
}

window.addEventListener('DOMContentLoaded', initSideDrawerEvents);

(function handleAdvancedMode() {
    const params = new URLSearchParams(window.location.search);
    const isAdvanced = params.get("advanced") === "true";

    function show(elId, visible) {
        const el = document.getElementById(elId);
        if (!el) return;
        el.style.display = visible ? "" : "none";
    }

    show("act-upload", isAdvanced);
    show("label-file", isAdvanced);
    show("toggle-draggable", isAdvanced);
    show("act-save", isAdvanced);
    show("act-reset-layout", isAdvanced);
    show("switch-label", isAdvanced);
})();

(function blockDesktopPinch() {
    const isDesktop = window.matchMedia('(hover: hover) and (pointer: fine)').matches;
    const isMac = (navigator.platform || '').toUpperCase().includes('MAC') || /Mac OS X/.test(navigator.userAgent);

    if (!(isDesktop && isMac)) return; //

    window.addEventListener('wheel', (e) => {
        if (e.ctrlKey) {
            e.preventDefault();
        }
    }, {passive: false});

    window.addEventListener('gesturestart', (e) => e.preventDefault(), {passive: false});
    window.addEventListener('gesturechange', (e) => e.preventDefault(), {passive: false});
    window.addEventListener('gestureend', (e) => e.preventDefault(), {passive: false});
})();

function openDrawer({
                        name: title,
                        description,
                        elements,
                        channels,
                        email,
                        highlightService,
                        highlightQuery,
                        elementsTitle = "Managed Services ⚙️",
                        elementsBaseUrl
                    }) {
    console.log('open');

    const drawer = document.getElementById('drawer');
    const overlay = document.getElementById('drawer-overlay');

    if (!drawer) {
        console.warn('[drawer] #drawer non trovato');
        return;
    }

    // --- recupera/crea in modo robusto i sotto-elementi del drawer ---
    // title
    let titleEl = document.getElementById('drawer-title');
    if (!titleEl) {
        titleEl = document.createElement('h2');
        titleEl.id = 'drawer-title';
        drawer.prepend(titleEl);
    }

    // description
    let descEl = document.getElementById('drawer-description');
    if (!descEl) {
        descEl = document.createElement('div');
        descEl.id = 'drawer-description';
        drawer.appendChild(descEl);
    }

    // list
    let listEl = document.getElementById('drawer-list');
    if (!listEl) {
        listEl = document.createElement('ul');
        listEl.id = 'drawer-list';
        drawer.appendChild(listEl);
    }

    // --- set title ---
    titleEl.textContent = `${title ?? ''}`;

    // --- reset contenuti in modo sicuro ---
    descEl.replaceChildren();     // svuota senza toccare il nodo
    listEl.replaceChildren();     // svuota senza toccare il nodo

    // --- accordion container (sezioni a comparsa) ---
    const accordion = document.createElement('div');
    accordion.className = 'drawer-accordion';
    descEl.appendChild(accordion);

// helper per creare una sezione collapsible
    function addDrawerSection(label, fillFn, { open = false, sectionId = '' } = {}) {
        const details = document.createElement('details');
        details.className = 'drawer-section';
        if (open) details.open = true;
        if (sectionId) details.dataset.sectionId = sectionId;

        const summary = document.createElement('summary');
        summary.className = 'drawer-section__summary';
        summary.textContent = label;

        const body = document.createElement('div');
        body.className = 'drawer-section__body';

        details.appendChild(summary);
        details.appendChild(body);
        accordion.appendChild(details);

        if (typeof fillFn === 'function') fillFn(body, details);

        return { details, body };
    }

// (opzionale) comportamento "accordion": ne lascia aperta una sola
    //   accordion.addEventListener('toggle', (e) => {
    //       const t = e.target;
    //       if (!(t instanceof HTMLDetailsElement)) return;
    //       if (!t.open) return;
    //       accordion.querySelectorAll('details.drawer-section').forEach(d => {
    //           if (d !== t) d.open = false;
    //       });
    //   });

// --- Overview / Description ---
    if (description) {
        addDrawerSection('Overview', (body) => {
            createFormattedLongTextElementsFrom(description).forEach(el => body.appendChild(el));
        }, { open: true, sectionId: 'overview' });
    }

    // --- channels ---
    if (channels && channels.length > 0) {
        addDrawerSection('Channels 💬', (body) => {
            const ul = document.createElement('ul');
            channels.forEach(channel => {
                const li = document.createElement('li');
                const channelLink = createHrefElement(
                    channel,
                    channel?.includes("slack.com") ? "Slack Channel" : "Link"
                );
                li.appendChild(channelLink);
                ul.appendChild(li);
            });
            body.appendChild(ul);
        }, { open: false, sectionId: 'channels' });
    }

    // --- email ---
    if (email && email !== "") {
        addDrawerSection('Team Mailbox ✉️', (body) => {
            body.appendChild(
                createHrefElement(createOutlookUrl([email]), `${truncateString(email, 25)}`)
            );
        }, { open: false, sectionId: 'mailbox' });
    }

    // --- Managed Services / Elements ---
    let servicesSection = null;

    if (elements && elements.items && elements.items.length > 0) {
        // apri di default se stai passando highlightService o highlightQuery
        const shouldOpenServices = !!(highlightService || (highlightQuery && highlightQuery.trim()));

        servicesSection = addDrawerSection(elementsTitle, (body) => {
            // riusa la UL già esistente (#drawer-list) per non rompere highlight/scroll
            // (l'hai già creata sopra come listEl)
            const frag = document.createDocumentFragment();

            elements.items.forEach(s => {
                const li = document.createElement('li');
                if (elementsBaseUrl) {
                    const a = document.createElement('a');
                    a.href = elementsBaseUrl(s);
                    a.textContent = s;
                    a.target = '_blank';
                    li.appendChild(a);
                } else {
                    li.textContent = s;
                }
                frag.appendChild(li);
            });

            listEl.replaceChildren(frag);
            body.appendChild(listEl);

            // evidenzia/scroll ai servizi
            (function multiHighlight() {
                const norm = v => (v || '').toString().trim().toLowerCase();

                const anchors = Array.from(listEl.querySelectorAll('li > a'));
                const items = anchors.length ? anchors : Array.from(listEl.querySelectorAll('li'));

                listEl.querySelectorAll('.service-hit-highlight')
                    .forEach(el => el.classList.remove('service-hit-highlight'));

                let firstHighlighted = null;

                const q = (highlightQuery || '').trim();
                if (q) {
                    const qn = normalizeWs(q).toLowerCase();
                    items.forEach(el => {
                        const text = normalizeWs(el.textContent).toLowerCase();
                        if (text.includes(qn)) {
                            el.classList.add('service-hit-highlight');
                            if (!firstHighlighted) firstHighlighted = el;
                        }
                    });
                }

                if (highlightService) {
                    const target = (highlightService || '').toString().trim().toLowerCase();
                    items.forEach(el => {
                        const text = (el.textContent || '').toString().trim().toLowerCase();
                        if (text === target) {
                            el.classList.add('service-hit-highlight');
                            if (!firstHighlighted) firstHighlighted = el;
                        }
                    });
                }

                if (firstHighlighted) {
                    try { firstHighlighted.scrollIntoView({ block: 'center', behavior: 'smooth' }); } catch {}
                }
            })();

        }, { open: shouldOpenServices, sectionId: 'services' });
    }

    // --- stato visivo del drawer ---
    drawer.classList.add('open');
    overlay?.classList.add('visible');
    document.body.classList.add('drawer-open'); // valuta `right-drawer-open` se separi i due drawer
    drawer.setAttribute('aria-hidden', 'false');
    console.log('fine');
}


function closeDrawer() {
    const drawer = document.getElementById('drawer');
    const overlay = document.getElementById('drawer-overlay');
    if (!drawer) return;
    drawer.classList.remove('open');
    overlay?.classList.remove('visible');
    document.body.classList.remove('drawer-open');
    drawer.setAttribute('aria-hidden', 'true');
}

function initDrawerEvents() {
    const overlay = document.getElementById('drawer-overlay');
    const closeBtn = document.getElementById('drawer-close');
    overlay?.addEventListener('click', closeDrawer);
    closeBtn?.addEventListener('click', closeDrawer);
}

window.addEventListener('DOMContentLoaded', initDrawerEvents);

function clearAllStreamsAndSearch() {
    resetStreamVisibility();
    clearSearch();
}

window.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape') return;

    const drawerOpen = document.body.classList.contains('drawer-open');

    if (drawerOpen) {
        closeDrawer()
    } else {
        clearAllStreamsAndSearch();
    }
});

window.addEventListener('load', function () {
    fetch('https://francesconicolosi.github.io/itsm/sample-people-database.csv')
        .then(response => response.text())
        .then(csvData => {
            cachedCsvText = csvData;
            resetVisualization();
            extractData(csvData);
            searchParam = getQueryParam('search');
            if (searchParam) {
                requestAnimationFrame(() => {
                    requestAnimationFrame(() => {
                        searchByQuery(searchParam);
                    });
                });
            }
        })
        .catch(error => console.error('Error loading the CSV file:', error));
});

document.getElementById('act-save')?.addEventListener('click', () => {
    const layout = {};

    d3.selectAll('.draggable').each(function () {
        const sel = d3.select(this);
        const key = sel.attr('data-key');
        if (!key) return;

        const {x, y} = parseTranslate(sel.attr('transform'));
        layout[key] = {x, y};
    });

    d3.selectAll('.draggable').each(function () {
        const sel = d3.select(this);
        const key = sel.attr('data-key');
        if (!key) return;

        const rect = sel.select('rect');
        if (!rect.empty()) {
            const w = +rect.attr('width');
            const h = +rect.attr('height');
            layout[key] = {...(layout[key] || {}), width: w, height: h};
        }
    });

    localStorage.setItem(LS_KEY, JSON.stringify(layout));
    showToast('Scenario successfully saved!');
});


function resetVisualization() {
    const svgEl = document.getElementById('canvas');
    if (!svgEl) {
        console.error('canvas not found');
        return;
    }

    width = svgEl.clientWidth || +svgEl.getAttribute('width') || 1200;
    height = svgEl.clientHeight || +svgEl.getAttribute('height') || 800;

    d3.select(svgEl).selectAll('*').remove();

    svg = d3.select(svgEl)
        .attr('width', width)
        .attr('height', height)
        .attr('cursor', 'grab');

    viewport = svg.append('g').attr('id', 'viewport');
    streamLayer = viewport.append('g').attr('id', 'streamLayer');
    themeLayer = viewport.append('g').attr('id', 'themeLayer');
    teamLayer = viewport.append('g').attr('id', 'teamLayer');
    cardLayer = viewport.append('g').attr('id', 'cardLayer');
    logoLayer = viewport.append('g').attr('id', 'logoLayer');

    zoom = d3.zoom()
        .filter((event) => {
            if (event.type === 'wheel') return !event.ctrlKey;          // ⬅️ ignora pinch su trackpad
            if (event.type === 'mousedown') return event.button === 0;
            if (event.type.startsWith('touch')) return true;
            return !event.ctrlKey;
        })
        .scaleExtent([0.1, 1])
        .on('start', () => svg.attr('cursor', 'grabbing'))
        .on('end', () => svg.attr('cursor', 'grab'))
        .on('zoom', (event) => {
            viewport.attr('transform', event.transform);
        });

    svg.call(zoom);
}

function setStreamFilter(streamKeys /* Set | null */) {
    const params = new URLSearchParams(window.location.search);

    if (!streamKeys || streamKeys.size === 0) {
        params.delete('stream');
    } else {
        params.set('stream', [...streamKeys].join(','));
    }

    const newUrl =
        window.location.pathname +
        (params.toString() ? '?' + params.toString() : '');

    window.history.replaceState({}, '', newUrl);

    resetVisualization();
    extractData(cachedCsvText);
}

function showToast(message, duration = 3000) {
    // --- helper: posiziona il container in base allo stato del drawer ---
    function positionToastContainer(container) {
        const drawer = document.getElementById('drawer');
        const isDrawerOpen = drawer?.classList.contains('open');

        // basso a destra se il drawer è aperto; altrimenti alto a destra
        if (isDrawerOpen) {
            container.style.top = 'unset';
            container.style.bottom = '20px';
            container.style.right = '20px';
        } else {
            container.style.bottom = 'unset';
            container.style.top = '70px';
            container.style.right = '20px';
        }
        // assicurati che stia sopra a drawer (9999) e overlay (9998)
        container.style.zIndex = '10001';
        container.style.position = 'fixed';
        container.style.display = 'flex';
        container.style.flexDirection = 'column';
        container.style.gap = '10px';
    }

    // --- crea o recupera il container ---
    let container = document.querySelector('.toast-container');
    if (!container) {
        container = document.createElement('div');
        container.className = 'toast-container';
        document.body.appendChild(container);
    }

    // --- posiziona subito (nel caso il drawer sia già aperto) ---
    positionToastContainer(container);

    // --- crea il toast e animazione ---
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.textContent = message;
    container.appendChild(toast);

    // fornisci il tempo al browser per applicare gli stili prima della transizione .show
    setTimeout(() => toast.classList.add('show'), 10);

    // --- riposiziona al prossimo frame e dopo un piccolo delay ---
    // (copre il caso in cui openDrawer venga chiamato subito dopo showToast)
    requestAnimationFrame(() => positionToastContainer(container));
    setTimeout(() => positionToastContainer(container), 180);

    // --- attacca un observer UNA SOLA VOLTA per reagire ai cambi dello stato del drawer ---
    if (!window.__toastDrawerObserverAttached) {
        const drawer = document.getElementById('drawer');
        if (drawer) {
            const mo = new MutationObserver(() => positionToastContainer(container));
            mo.observe(drawer, { attributes: true, attributeFilter: ['class'] });
            window.__toastDrawerObserverAttached = true;
        }
    }

    // --- chiudi e rimuovi il toast dopo la durata ---
    setTimeout(() => {
        toast.classList.remove('show');
        setTimeout(() => toast.remove(), 300);
    }, duration);
}

function fitToContent(paddingRatio = 0.9) {
    if (!viewport) return;

    const bbox = viewport.node().getBBox();
    if (!bbox || !isFinite(bbox.width) || !isFinite(bbox.height) || bbox.width === 0 || bbox.height === 0) {
        svg.call(zoom.transform, d3.zoomIdentity);
        return;
    }

    const scale = Math.min(width / bbox.width, height / bbox.height) * paddingRatio;
    const x = width / 2 - (bbox.x + bbox.width / 2) * scale;
    const y = height / 2 - (bbox.y + bbox.height / 2) * scale;

    const t = d3.zoomIdentity.translate(x, y).scale(scale);
    svg.call(zoom.transform, t);
}

function zoomToElement(element, desiredScale = 1.5, duration = 500) {
    if (!element || !svg) return;

    const svgNode = svg.node();
    const t = d3.zoomTransform(svgNode);

    const elRect = element.getBoundingClientRect();
    const svgRect = svgNode.getBoundingClientRect();
    const centerScreenX = elRect.left + elRect.width / 2 - svgRect.left;
    const centerScreenY = elRect.top + elRect.height / 2 - svgRect.top;

    const [cx, cy] = t.invert([centerScreenX, centerScreenY]);

    const k = desiredScale;
    const offsetY = 190;
    const tx = width / 2 - cx * k;
    const ty = height / 2 - cy * k - offsetY;

    const targetTransform = d3.zoomIdentity.translate(tx, ty).scale(k);
    svg.transition().duration(duration).call(zoom.transform, targetTransform);

    const group = element.closest('g');
    if (group) highlightGroupUtils(d3.select(group));
}

const cleanName = (name) => normalizeWs(name);

const findPersonByName = (targetName, result) => {
    const target = normalizeWs(targetName).toLowerCase();

    return Object.values(result).flatMap(stream =>
        Object.values(stream).flatMap(theme =>
            Object.values(theme).flatMap(team => team)
        )
    ).find(person => {
        const pn = normalizeWs(person?.Name).toLowerCase();
        return pn === target;
    }) || null;
};

function buildOrganization(people) {
    roleDetailsMapping = new Map();
    const organization = {};
    for (const person of people) {
        let firstLevelItems = (person[firstOrgLevel] || '').split(/\n|,/).map(s => s.trim()).filter(Boolean);
        if (firstLevelItems.length === 0) firstLevelItems = [firstLevelNA];

        let secondLevelItems = (person[secondOrgLevel] || '').split(/\n|,/).map(t => t.trim()).filter(Boolean);
        if (secondLevelItems.length === 0) secondLevelItems = [secondLevelNA];

        let thirdLevelItems = (person[thirdOrgLevel] || '').split(/\n|,/).map(t => t.trim()).filter(Boolean);
        if (thirdLevelItems.length === 0) thirdLevelItems = [thirdLevelNA];

        for (const firstLevelItem of firstLevelItems) {
            if (!organization[firstLevelItem]) organization[firstLevelItem] = {};
            for (const theme of secondLevelItems) {
                if (!organization[firstLevelItem][theme]) organization[firstLevelItem][theme] = {};
                for (const team of thirdLevelItems) {
                    if (!organization[firstLevelItem][theme][team]) organization[firstLevelItem][theme][team] = [];
                    person.Name = person.Name ? cleanName(person.Name) : person.User;
                    person.Name = cleanName(person.Name || '')
                        || (person.User || '').trim()
                        || (person[emailField] || '').trim()
                        || 'Unknown';

                    if (!roleDetailsMapping.has(person.Role)) {
                        roleDetailsMapping.set(person.Role, {grants: person["Role Grants"], description: person["Role Description"] });
                    }

                    const teamArr = organization[firstLevelItem][theme][team];

                    const existingKeys = new Set(
                        teamArr.map(p => buildCompositeKey(p, emailField)).filter(Boolean)
                    );

                    let compositeKey = buildCompositeKey(person, emailField);

                    const isFullyEmptyKey = !compositeKey;

                    const isDuplicate = compositeKey ? existingKeys.has(compositeKey) : false;

                    if (!isFullyEmptyKey && !isDuplicate) {
                        teamArr.push(person);
                    }
                }
            }
        }
    }
    return organization;
}

const addGuestManagersByRole = (person, guestRole, thirdLevel, organization) => {
    if (!person[guestRole]) return;
    const guestNames = [...new Set(
        splitValues(person[guestRole] || '')
            .flatMap(v => v.split(/\n|,/))
            .map(v => v.trim())
            .filter(Boolean)
    )];

    guestNames.forEach(name => {
        const manager = findPersonByName(name, organization);
        if (!manager) {
            return;
        }
        const alreadyPresent = thirdLevel.some(member => cleanName(member.Name) === cleanName(name));
        if (!alreadyPresent) {
            manager.guestRole = guestRole;
            thirdLevel.push(manager);
        }
    });
};

function addGuestManagersTo(organization) {
    const result = {};
    for (const [firstLevel, secondLevelItems] of Object.entries(organization)) {
        for (const [secondLevel, thirdLevelItems] of Object.entries(secondLevelItems)) {
            for (const [thirdLevel, members] of Object.entries(thirdLevelItems)) {
                if (!result[firstLevel]) result[firstLevel] = {};
                if (!result[firstLevel][secondLevel]) result[firstLevel][secondLevel] = {};
                if (!result[firstLevel][secondLevel][thirdLevel]) result[firstLevel][secondLevel][thirdLevel] = [];

                for (const p of members) {
                    const names = Object.values(result[firstLevel][secondLevel][thirdLevel]).map(entry => entry.Name);
                    if (!names.includes(p.Name)) result[firstLevel][secondLevel][thirdLevel].push(p);
                    guestRoleColumns.forEach(role => addGuestManagersByRole(p, role, result[firstLevel][secondLevel][thirdLevel], organization));
                }
                result[firstLevel][secondLevel][thirdLevel].sort((a, b) => {
                    const aIsGuest = guestRoleColumns.includes(a.guestRole);
                    const bIsGuest = guestRoleColumns.includes(b.guestRole);
                    if (aIsGuest && !bIsGuest) return 1;
                    if (!aIsGuest && bIsGuest) return -1;
                    return 0;
                });
            }
        }
    }
    return result;
}

function getLatestUpdateFromCsv(headers, rows) {
    if (headers.includes("Updated")) {
        const dateIndex = headers.indexOf("Updated");
        const dates = rows.slice(1)
            .map(row => row[dateIndex]?.trim())
            .filter(Boolean)
            .map(d => new Date(d))
            .filter(d => !isNaN(d.getTime()));

        if (dates.length > 0) {
            const lastUpdateEl = document.getElementById('side-last-update');
            if (lastUpdateEl) {
                lastUpdateEl.textContent = `Last Update: ${getFormattedDate(new Date(Math.max(...dates.map(d => d.getTime()))).toISOString())}`;
            }
        }
    }
}

function getContentBBox() {
    const bg = backgroundLayer?.node()?.getBBox();
    const cards = cardLayer?.node()?.getBBox();

    if (!bg && !cards) return null;
    const boxes = [bg, cards].filter(Boolean);
    const x1 = Math.min(...boxes.map(b => b.x));
    const y1 = Math.min(...boxes.map(b => b.y));
    const x2 = Math.max(...boxes.map(b => b.x + b.width));
    const y2 = Math.max(...boxes.map(b => b.y + b.height));
    return {x: x1, y: y1, width: x2 - x1, height: y2 - y1};

}


function placeCompanyLogoUnderDiagram(url = './assets/company-logo.png', maxWidth = 240, textMargin = 40) {
    if (!viewport || !logoLayer) return;

    const bbox = getContentBBox();
    if (!bbox) {
        console.warn('Visual outcome not found');
        return;
    }

    logoLayer.selectAll('*').remove();

    const img = new Image();
    img.onload = () => {
        const aspect = img.height / img.width || 0.35;
        const width = maxWidth;
        const height = Math.round(width * aspect);

        const x = bbox.x + (bbox.width - width) / 2;
        const y = bbox.y + bbox.height + Math.max(300, bbox.height * 0.12);
        logoLayer.append('image')
            .attr('href', url)
            .attr('x', x)
            .attr('y', y)
            .attr('width', width)
            .attr('height', height)
            .attr('preserveAspectRatio', 'xMidYMid meet')
            .style('pointer-events', 'none');

        logoLayer.append('foreignObject')
            .attr('x', x)
            .attr('y', y + height + textMargin)
            .attr('width', width)
            .attr('height', 100)
            .append('xhtml:div')
            .style('font-size', '10px')
            .style('font-family', '"Montserrat"', '"Sans 3", Arial, sans-serif')
            .style('text-align', 'center')
            .style('color', '#333')
            .html('<p>Author: Francesco Nicolosi</p>' +
                '<p>Personal Blog: <a href="https://www.gamerdad.cloud" target="_blank">www.gamerdad.cloud</a></p>' +
                '<p><img src="https://img.shields.io/badge/license-NonCommercial-blue.svg"></p>');

        let notZoommingToShowSearchResults = !getQueryParam("search");
        if (notZoommingToShowSearchResults) {
            fitToContent(0.9);
        }

    };
    img.onerror = () => {
        console.warn('Logo not found:', url);
    };
    img.src = url;
}


let isDraggable = false;

function applyDraggableToggleState() {
    const groups = d3.selectAll('.draggable');
    const handles = d3.selectAll('.resize-handles');
    if (isDraggable) {
        groups.call(drag);
        handles.style('display', null).style('pointer-events', 'all');
    } else {
        groups.on('.drag', null);
        handles.style('display', 'none')
            .style('pointer-events', 'none');
    }
}

function wireFabsInteractions(cardSel) {
    const SHOW_DELAY = 50;
    const HIDE_DELAY = 120;
    let showTimer = null;
    let hideTimer = null;

    const isTouchEnv = () =>
        ('ontouchstart' in window) || navigator.maxTouchPoints > 0;

    const show = () => {
        clearTimeout(hideTimer);
        showTimer = setTimeout(() => {
            cardSel.classed('card--fabs-visible', true);
        }, SHOW_DELAY);
    };

    const hide = () => {
        clearTimeout(showTimer);
        hideTimer = setTimeout(() => {
            cardSel.classed('card--fabs-visible', false);
        }, HIDE_DELAY);
    };

    cardSel
        .on('pointerenter.fabs', () => {
            if (!isTouchEnv()) show();
        })
        .on('pointerleave.fabs', () => {
            if (!isTouchEnv()) hide();
        });

    cardSel.on('click.fabs', (event) => {
        if (!isTouchEnv()) return;
        event.stopPropagation();
        const vis = cardSel.classed('card--fabs-visible');
        d3.selectAll('g[data-key^="card::"]').classed('card--fabs-visible', false);
        cardSel.classed('card--fabs-visible', !vis);
    });

    cardSel.selectAll('.contact-fabs, .contact-fabs-svg, .contact-fab')
        .on('pointerenter.fabs', (e) => e.stopPropagation())
        .on('pointerleave.fabs', (e) => e.stopPropagation())
        .on('pointerdown.fabs', (e) => e.stopPropagation())
        .on('touchstart.fabs', (e) => e.stopPropagation());

    if (!window.__fabsOutsideHandlerAttached) {
        window.__fabsOutsideHandlerAttached = true;
        document.addEventListener('pointerdown', (e) => {
            const svgEl = document.getElementById('canvas');
            if (!svgEl) return;
            if (!svgEl.contains(e.target)) {
                d3.selectAll('g[data-key^="card::"]').classed('card--fabs-visible', false);
            }
        }, { passive: true });
    }
}

document.getElementById('act-reset-layout')?.addEventListener('click', () => {
    localStorage.removeItem(LS_KEY);
    window.location.reload();
});


function getCompanyGroupTalentUrl(member, emailField = 'Email') {
    const email = (member?.[emailField] ?? '').toString().trim();
    const query = encodeURIComponent(email);
    return `https://company.talentsoftware.ai/careerhub/search/people?query=${query}`;
}

document.getElementById('toggle-draggable')?.addEventListener('change', (e) => {
    isDraggable = e.target.checked;
    applyDraggableToggleState();
});

function getThemeTeamsCount(themeObj) {
    return Object.keys(themeObj || {}).length || 0;
}

function getThemeMaxMemberRows(themeObj, inARow) {
    const counts = Object.values(themeObj || {}).map(members => {
        const n = new Set((members || []).map(m => (m?.Name || '').trim()).filter(Boolean)).size;
        return n;
    });
    const maxMembers = Math.max(0, ...counts);
    return Math.max(1, Math.ceil(maxMembers / inARow));
}

function extractData(csvText) {
    if (!csvText) {
        alert('Missing CSV File!');
        return;
    }
    colorKeyMappings = new Map();
    const rows = parseCSV(csvText);
    if (rows.length < 2) return;

    const headers = rows[0].map(h => h.trim());
    people = rows.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => {
            obj[h] = normalizeWs(row[i] || '', h);
        });
        return obj;
    }).filter(p => (p.Status || '').toLowerCase() !== 'inactive');

    let lastUpdateISO = '';
    if (headers.includes('Updated')) {
        const idx = headers.indexOf('Updated');
        const dates = rows.slice(1)
            .map(r => r[idx]?.trim())
            .filter(Boolean)
            .map(d => new Date(d))
            .filter(d => !isNaN(d));
        if (dates.length) {
            const maxTs = Math.max(...dates.map(d => d.getTime()));
            lastUpdateISO = new Date(maxTs).toISOString().slice(0, 10); // yyyy-mm-dd
        }
    }
    const peopleCount = people.length;
    const datasetVersion = `people:${peopleCount}|lu:${lastUpdateISO || 'n/a'}`;
    LS_KEY = `dsm-layout-v1::${datasetVersion}`;


    getLatestUpdateFromCsv(headers, rows);

    const organization = buildOrganization(people);
    const organizationWithManagers = addGuestManagersTo(organization);

    const filteredStreams = getAllowedStreamsSet();

    const allStreamNames = Object.keys(organizationWithManagers || {})
        .filter(s => s && !s.includes(firstLevelNA));

    const visibleStreamNames = (filteredStreams && filteredStreams.size > 0)
        ? allStreamNames.filter(s => filteredStreams.has(s) || filteredStreams.has(normalizeKey(s)))
        : allStreamNames;


    visibleOrganizationWithManagers = filterOrganizationByStreams(organizationWithManagers, filteredStreams);
    const visiblePeopleForLegend = collectMembersFromOrganization(visibleOrganizationWithManagers);


    const inARow = 6;
    const dateValues = ["In team since"];
    const fieldsToShow = [
        "Role", "Company", "Location", "Room Link",
        ...dateValues
    ];

    const nFields = fieldsToShow.length + 0.5;
    const rowHeight = 11;
    const memberWidth = 160, cardPad = 10, cardBaseHeight = nFields * 4 * rowHeight;
    const thirdLevelBoxWidth = inARow * memberWidth + 100, thirdLevelBoxPadX = 24;
    const secondLevelBoxPadX = 60;
    const firstLevelBoxPadY = 100;

    const largestThirdLevelSize = Math.max(
        ...Object.entries(organizationWithManagers)
            .filter(([streamName]) => streamName !== firstLevelNA)
            .flatMap(([, stream]) =>
                Object.entries(stream)
                    .filter(([themeName]) => themeName !== secondLevelNA)
                    .flatMap(([, theme]) =>
                        Object.values(theme).map(team =>
                            new Set(team.map(m => m.Name?.trim()).filter(Boolean)).size
                        )
                    )
            )
    );

    const rowCount = Math.ceil(largestThirdLevelSize / inARow);
    const thirdLevelBoxHeight = rowCount * cardBaseHeight * 1.2 + 80;
    const secondLevelBoxHeight = thirdLevelBoxHeight * 1.2 + 100;

    let streamY = 40;
    let streamX = 40;

    const orderedStreams = Object.entries(organizationWithManagers)
        .sort(([a], [b]) => {
            const na = normalizeKey(a);
            const nb = normalizeKey(b);

            const ia = STREAM_ORDER.indexOf(na);
            const ib = STREAM_ORDER.indexOf(nb);

            if (ia !== -1 && ib !== -1) return ia - ib;

            if (ia !== -1) return -1;

            if (ib !== -1) return 1;

            return a.localeCompare(b, 'en', { sensitivity: 'base' });
        });

    orderedStreams.forEach(([firstLevel, secondLevelItems]) => {
        if (firstLevel.includes(firstLevelNA)) return;

        // Filtra stream visibili
        if (filteredStreams) {
            const firstLevelNormalized = normalizeKey(firstLevel);
            const isAllowed = filteredStreams.has(firstLevel) || filteredStreams.has(firstLevelNormalized);
            if (!isAllowed) return;
        }

        // Descrizione stream
        const firstLevelMembers =
            Object.values(organization[firstLevel] || {})
                .flatMap(themeObj => Object.values(themeObj))
                .flat();
        const firstLevelDescription =
            aggregateInfoByHeader(firstLevelMembers, headers, "Team Stream Description")?.items?.join("") ?? '';

        // LARGHEZZA stream (come già fai)
        const firstLevelBoxWidth = computeStreamBoxWidthByCapacity(
            secondLevelItems,
            secondLevelBoxPadX,
            secondLevelNA,
            thirdLevelBoxPadX,
            thirdLevelBoxWidth,
            SECOND_LEVEL_LABEL_EXTRA
        );

        // Gruppo stream
        const firstLevelGroup = streamLayer.append('g')
            .attr('class', 'draggable')
            .attr('transform', `translate(${streamX},${streamY})`)
            .attr('data-key', `stream::${normalizeKey(firstLevel)}`);
        restoreGroupPosition(firstLevelGroup);

        // ---------- COSTRUZIONE RIGHE DI THEME ----------
        const rows = [];
        let currentRow = { themes: [], used: 0 };

        for (const [secondLevel, thirdLevelItems] of Object.entries(secondLevelItems)) {
            if (secondLevel.includes(secondLevelNA)) continue;

            const nTeams = Object.keys(thirdLevelItems || {}).length || 0;
            // Vai a capo se superi la capacità
            if (currentRow.used > 0 && (currentRow.used + nTeams) > MAX_TEAMS_PER_ROW) {
                rows.push(currentRow);
                currentRow = { themes: [], used: 0 };
            }

            // Max righe di member richieste dal theme
            const memberCounts = Object.values(thirdLevelItems || {}).map(members => {
                return new Set((members || []).map(m => (m?.Name || '').trim()).filter(Boolean)).size;
            });
            const maxMembersInTheme = Math.max(0, ...memberCounts);
            const themeMaxRows = Math.max(1, Math.ceil(maxMembersInTheme / inARow));

            // Larghezza del theme (n° team * larghezza team + gap + etichetta)
            const themeWidth =
                nTeams * thirdLevelBoxWidth +
                Math.max(0, nTeams - 1) * thirdLevelBoxPadX +
                SECOND_LEVEL_LABEL_EXTRA;

            currentRow.themes.push({
                secondLevel,
                thirdLevelItems,
                nTeams,
                themeMaxRows,
                themeWidth
            });
            currentRow.used += nTeams;
        }
        if (currentRow.themes.length) rows.push(currentRow);

        // ---------- ALTEZZE PER RIGA ----------
        rows.forEach(r => {
            r.rowMaxMemberRows = Math.max(1, ...r.themes.map(t => t.themeMaxRows));
            const teamBoxPadding  = r.rowMaxMemberRows > 1 ? 80 : 120; // tuoi valori
            const themeBoxPadding = 100;                                // tuoi valori
            r.teamBoxHeight  = r.rowMaxMemberRows * cardBaseHeight * 1.2 + teamBoxPadding;
            r.themeBoxHeight = r.teamBoxHeight * 1.2 + themeBoxPadding;
        });

        // ---------- ALTEZZA STREAM (somma delle righe) ----------
        const firstLevelBoxHeight =
            rows.reduce((acc, r) => acc + r.themeBoxHeight, 0) +
            (rows.length > 1 ? (rows.length - 1) * secondLevelRowPadY : 0) +
            140;

        // Rettangolo stream
        const streamRect = firstLevelGroup.append('rect')
            .attr('class', 'stream-box')
            .attr('width', firstLevelBoxWidth)
            .attr('height', firstLevelBoxHeight)
            .attr('rx', 40)
            .attr('ry', 40);

        makeResizable(firstLevelGroup, streamRect, { minWidth: 600, minHeight: 300 });

        // Titolo stream
        const titleText = firstLevelGroup.append('text')
            .attr('x', 50)
            .attr('y', 70)
            .attr('text-anchor', 'start')
            .attr('class', 'stream-title');

        titleText.append('tspan')
            .attr('class', 'stream-title')
            .text(firstLevel);
        titleText.append('tspan').attr('dx', 10).text('');

        if (firstLevelDescription !== "") {
            titleText.append('tspan')
                .attr('class', 'stream-icon stream-icon--desc')
                .attr('data-tooltip', 'View stream details')
                .attr('aria-label', 'View stream details')
                .text(' ℹ️')
                .on('click', (e) => {
                    e?.stopPropagation?.();
                    openDrawer({ name: firstLevel, description: firstLevelDescription });
                });
            titleText.append('tspan').attr('dx', 10).text('');
        }

        if (visibleStreamNames.length > 1) {
            titleText.append('tspan')
                .attr('class', 'stream-icon stream-icon--isolate')
                .attr('data-tooltip', 'Show this stream only (ESC to reset)')
                .attr('aria-label', 'Show this stream only (ESC to reset)')
                .text(' 👁️‍🗨️')
                .style('cursor', 'pointer')
                .on('click', (e) => {
                    e.stopPropagation();
                    const key = normalizeKey(firstLevel);

                    setStreamFilter(new Set([key]));
                });


            titleText.append('tspan')
                .attr('class', 'stream-icon stream-icon--hide')
                .attr('data-tooltip', 'Hide this stream (ESC to reset)')
                .attr('aria-label', 'Hide this stream (ESC to reset)')
                .text(' 🙈')
                .style('cursor', 'pointer')
                .on('click', (e) => {
                    e.stopPropagation();

                    const key = normalizeKey(firstLevel);
                    const current = getAllowedStreamsSet();

                    let next;

                    if (!current) {
                        next = new Set(
                            visibleStreamNames.map(s => normalizeKey(s))
                        );
                        next.delete(key);
                    } else {
                        next = new Set(current);
                        next.delete(key);
                    }
                    setStreamFilter(next.size > 0 ? next : null);
                });
        }

        if (firstLevelDescription !== "") {
            firstLevelGroup.select('rect.stream-box')
                .style('cursor', 'pointer')
                .on('click', () => openDrawer({
                    name: firstLevel,
                    description: firstLevelDescription
                }));

            firstLevelGroup.select('text.stream-title')
                .style('cursor', 'pointer')
                .on('click', () => openDrawer({
                    name: firstLevel,
                    description: firstLevelDescription
                }));
        }

        // ---------- RENDER THEME/TEAM CON ALTEZZE DI RIGA ----------
        let secondLevelYBase = streamY + 100;

        rows.forEach((r) => {
            let secondLevelX = 60; // reset a inizio riga
            const themeBoxHeightRow = r.themeBoxHeight;
            const teamBoxHeightRow  = r.teamBoxHeight;

            r.themes.forEach(({ secondLevel, thirdLevelItems, nTeams, themeWidth }) => {
                const secondLevelY = secondLevelYBase;

                const originalThemeMembers = Object.values(organization[firstLevel]?.[secondLevel] || {}).flat();
                const secondLevelDescription = aggregateInfoByHeader(
                    originalThemeMembers, headers, 'Team Theme Description'
                )?.items?.join("") ?? '';

                // Gruppo theme
                const secondLevelGroup = themeLayer.append('g')
                    .attr('class', 'draggable')
                    .attr('transform', `translate(${streamX + secondLevelX},${secondLevelY})`)
                    .attr('data-key', `theme::${normalizeKey(firstLevel)}::${normalizeKey(secondLevel)}`);
                restoreGroupPosition(secondLevelGroup);

                // Box theme (altezza di riga)
                const secondLevelRect = secondLevelGroup.append('rect')
                    .attr('class', 'theme-box')
                    .attr('width', themeWidth)
                    .attr('height', themeBoxHeightRow)
                    .attr('rx', 30)
                    .attr('ry', 30);
                makeResizable(secondLevelGroup, secondLevelRect, { minWidth: 400, minHeight: 250 });

                // Titolo theme
                secondLevelGroup.append('text')
                    .attr('x', themeWidth / 2)
                    .attr('y', 85)
                    .attr('text-anchor', 'middle')
                    .attr('class', 'theme-title')
                    .text(truncateString(secondLevel));

                if (secondLevelDescription !== "") {
                    secondLevelGroup.select('text.theme-title')
                        .append('tspan')
                        .attr('class', 'theme-icon')
                        .attr('dx', 10)
                        .attr('data-tooltip', 'View theme details')
                        .attr('aria-label', 'View theme details')
                        .text(' ℹ️')
                        .on('click', (e) => {
                            e.stopPropagation();
                            openDrawer({ name: secondLevel, description: secondLevelDescription });
                        });

                    secondLevelGroup.select('rect.theme-box')
                        .style('cursor', 'pointer')
                        .on('click', () => openDrawer({ name: secondLevel, description: secondLevelDescription }));

                    secondLevelGroup.select('text.theme-title')
                        .style('cursor', 'pointer')
                        .on('click', () => openDrawer({ name: secondLevel, description: secondLevelDescription }));
                }

                // Team cards nel theme
                Object.entries(thirdLevelItems).forEach(([thirdLevel, members], teamIdx) => {
                    const originalMembers = (organization[firstLevel]?.[secondLevel]?.[thirdLevel]) || [];
                    const services    = aggregateInfoByHeader(originalMembers, headers, 'Team Managed Services', true);
                    const description = aggregateInfoByHeader(originalMembers, headers, 'Team Description')?.items?.join("") ?? '';
                    const channels    = aggregateInfoByHeader(originalMembers, headers, 'Team Channels', true)?.items;
                    const email       = aggregateInfoByHeader(originalMembers, headers, 'Team Email')?.items?.join("") ?? '';

                    const teamLocalX = teamIdx * (thirdLevelBoxWidth + thirdLevelBoxPadX) + 50;
                    const teamLocalY = 130;

                    const thirdLevelGroup = teamLayer.append('g')
                        .attr('class', 'draggable')
                        .attr('transform', `translate(${streamX + secondLevelX + teamLocalX},${secondLevelY + teamLocalY})`)
                        .attr('data-key', `team::${normalizeKey(firstLevel)}::${normalizeKey(secondLevel)}::${normalizeKey(thirdLevel)}`);
                    restoreGroupPosition(thirdLevelGroup);

                    // Box team (altezza di riga)
                    const thirdLevelRect = thirdLevelGroup.append('rect')
                        .attr('class', 'team-box')
                        .attr('width', thirdLevelBoxWidth)
                        .attr('height', teamBoxHeightRow)
                        .attr('rx', 20)
                        .attr('ry', 20);
                    makeResizable(thirdLevelGroup, thirdLevelRect, { minWidth: 360, minHeight: 220 });

                    const serviceCount = services?.items?.length || 0;
                    const titleText = `${truncateString(thirdLevel)} - ⚙️ (${serviceCount})`;

                    thirdLevelGroup.append('text')
                        .attr('x', thirdLevelBoxWidth / 2)
                        .attr('y', 70)
                        .attr('text-anchor', 'middle')
                        .attr('data-services', services?.items?.filter(Boolean).join(', ') || '')
                        .attr('class', 'team-title')
                        .text(titleText);

                    thirdLevelGroup.select('rect.team-box')
                        .style('cursor', 'pointer')
                        .on('click', () => openDrawer({ name: thirdLevel, description, elements: services, channels, email, elementsBaseUrl: (s) => `domino.html?search=id%3A"${encodeURIComponent(s)}"` }));
                    thirdLevelGroup.select('text.team-title')
                        .style('cursor', 'pointer')
                        .on('click', () => openDrawer({ name: thirdLevel, description, elements: services, channels, email, elementsBaseUrl: (s) => `domino.html?search=id%3A"${encodeURIComponent(s)}"` }));

                    // RENDER CARD MEMBRO (lascia il tuo codice esistente)
                    members.forEach((member, mIdx) => {
                        const col = mIdx % inARow;
                        const row = Math.floor(mIdx / inARow);
                        const cardX = 40 + secondLevelX + teamIdx * (thirdLevelBoxWidth + thirdLevelBoxPadX) + 50 + 20 + col * (memberWidth + cardPad);
                        const cardY = secondLevelY + 70 + 45 + row * (cardBaseHeight + 10) + 130;


                        const group = cardLayer.append('g')
                            .attr('data-role', (member[ROLE_FIELD_WITH_MAPPING] || '').toString().trim())
                            .attr('data-company', (member[COMPANY_FIELD] || '').toString().trim())
                            .attr('data-location', (member[LOCATION_FIELD] || '').toString().trim())
                            .attr('class', 'draggable')
                            .attr('transform', `translate(${cardX},${cardY})`)
                            .attr('data-key', `card::${normalizeKey(firstLevel)}::${normalizeKey(secondLevel)}::${normalizeKey(thirdLevel)}::${normalizeKey(member['Name'] || member['User'] || mIdx)}`);

                        const colorKey =
                            colorBy === ROLE_FIELD_WITH_MAPPING ? group.attr('data-role') :
                                colorBy === COMPANY_FIELD ? group.attr('data-company') :
                                    group.attr('data-location');

                        colorKeyMappings.set(
                            colorBy,
                            (colorKeyMappings.get(colorBy) ?? new Set()).add(colorKey)
                        );

                        restoreGroupPosition(group);

                        const memberRect = group.append('rect')
                            .attr('class', 'profile-box')
                            .attr('width', memberWidth)
                            .attr('height', cardBaseHeight)
                            .attr('rx', 14)
                            .attr('ry', 14)
                            .attr('fill', getCardFill(group) ? getCardFill(group) : NEUTRAL_COLOR);


                        {
                            const f = getCardFill(group) || NEUTRAL_COLOR;
                            if ((f || '').toLowerCase() === '#ffffff' || f === 'white') {
                                memberRect.attr('stroke', '#b8b8b8').attr('stroke-width', 1);
                            }
                        }



                        if (member.guestRole) {
                            memberRect.attr('stroke', '#333')
                                .attr('stroke-width', 1.5)
                                .attr('stroke-dasharray', '4 2');
                        }

                        function getPhotoCandidates(email) {
                            const baseName = (email?.split('@')[0] || '').replace('-ext', '').replace('.', '-');

                            const fileName = `./assets/photos/${baseName}`;

                            return [
                                `${fileName}.webp`,
                                // `${fileName}.avif`,
                                `${fileName}.jpg`,
                                `${fileName}.png`,
                                `${fileName}.jpeg`,
                            ];
                        }

                        function resolvePhoto(email, fallback = './assets/user-icon.png', timeoutMs = 4000) {
                            const candidates = getPhotoCandidates(email);

                            const tryWithTimeout = (url) => new Promise((resolve, reject) => {
                                const img = new Image();
                                const timer = setTimeout(() => {
                                    img.onload = img.onerror = null;
                                    reject(new Error('timeout'));
                                }, timeoutMs);

                                img.onload = () => {
                                    clearTimeout(timer);
                                    resolve(url);
                                };
                                img.onerror = () => {
                                    clearTimeout(timer);
                                    reject(new Error('error'));
                                };

                                img.src = url;
                            });

                            return candidates
                                .reduce(
                                    (chain, url) => chain.catch(() => tryWithTimeout(url)),
                                    Promise.reject()
                                )
                                .catch(() => fallback);
                        }


                        resolvePhoto(member[emailField]).then(photoPath => {
                            const photoSize = 60;
                            const photoX = (memberWidth - photoSize) / 2;
                            const photoY = 8;

                            const photoWrapper = group.append('g')
                                .attr('class', 'photo-wrapper');

                            const photoFO = photoWrapper.append('foreignObject')
                                .attr('x', photoX)
                                .attr('y', photoY)
                                .attr('width', photoSize)
                                .attr('height', photoSize)
                                .attr('requiredExtensions', 'http://www.w3.org/1999/xhtml');

                            const photoDiv = photoFO.append('xhtml:div')
                                .style('width',  `${photoSize}px`)
                                .style('height', `${photoSize}px`)
                                .style('border-radius', '50%')
                                .style('overflow', 'hidden');

                            const photoImg = photoDiv.append('xhtml:img')
                                .attr('src', photoPath)
                                .attr('alt', member.Name || 'profile photo')
                                .style('display', 'block')
                                .style('width',  '100%')
                                .style('height', '100%')
                                .style('object-fit', 'cover')
                                .style('pointer-events', 'none');

                            let nTeams = 0;
                            try {
                                nTeams = countTeamsForMemberInOrg(member, visibleOrganizationWithManagers) || 0;
                            } catch {}

                            if (nTeams > 1) {
                                const badgeR = 10;
                                const bx = photoX + photoSize - badgeR - 1;
                                const by = photoY + photoSize - badgeR - 1;

                                const tooltipText = `Focus shared across ${nTeams} teams. Click to browse them`;

                                const badgeG = photoWrapper.append('g')
                                    .attr('class', 'multi-team-badge')
                                    .attr('transform', `translate(${bx},${by})`)
                                    .style('cursor', 'pointer')
                                    .attr('role', 'button')
                                    .attr('tabindex', 0)
                                    .attr('aria-label', tooltipText)
                                    .attr('data-tooltip', tooltipText);

                                badgeG.append('circle')
                                    .attr('r', badgeR)
                                    .attr('fill', '#111')
                                    .attr('stroke', '#fff')
                                    .attr('stroke-width', 1.5);

                                badgeG.append('text')
                                    .attr('text-anchor', 'middle')
                                    .attr('dominant-baseline', 'central')
                                    .attr('fill', '#fff')
                                    .style('font-weight', 600)
                                    .style('font-size', `${badgeR + 2}px`)
                                    .text(nTeams);

                                const triggerSearch = (e) => {
                                    e?.stopPropagation?.();
                                    const q = member.Name?.toLowerCase();
                                    if (q) searchByQuery(q);
                                };
                                badgeG.on('click', triggerSearch);
                                badgeG.on('keydown', (e) => {
                                    if (e.key === 'Enter' || e.key === ' ') triggerSearch(e);
                                });

                                badgeG.raise();
                            }

                        });


                        const nameY = 72;
                        const defaultNameBoxH = 24;

                        const nameFO = group.append('foreignObject')
                            .attr('x', 0)
                            .attr('y', nameY)
                            .attr('width', memberWidth)
                            .attr('height', defaultNameBoxH);
                        const nameDiv = nameFO.append('xhtml:div')
                            .attr('class', 'profile-name')
                            .html(member['Name']);

                        function adjustNameAndInfoHeights() {
                            const measured = nameDiv.node()?.scrollHeight || defaultNameBoxH;

                            const nameBoxH = Math.max(defaultNameBoxH, Math.ceil(measured) + 2);

                            nameFO.attr('height', nameBoxH);

                            const infoStartY = nameY + nameBoxH + 4;

                            const infoFOExisting = group.select('foreignObject .info').node()
                                ? d3.select(group.select('foreignObject .info').node().closest('foreignObject'))
                                : null;

                            if (infoFOExisting) {
                                infoFOExisting.attr('y', infoStartY);
                            }
                        }

                        requestAnimationFrame(() => requestAnimationFrame(adjustNameAndInfoHeights));

                        const infoDivFO_Y = nameY + defaultNameBoxH + 4;
                        const infoDiv = group.append('foreignObject')
                            .attr('x', 8)
                            .attr('y', infoDivFO_Y)
                            .attr('width', memberWidth - 16)
                            .attr('height', Math.max(0, cardBaseHeight - (infoDivFO_Y - 8)))
                            .append('xhtml:div')
                            .attr('class', 'info');

                        const email = member[emailField];

                        const isWebKit = /AppleWebKit/i.test(navigator.userAgent)
                            && /Safari/i.test(navigator.userAgent)
                            && !/(Chrome|Chromium|Edg)/i.test(navigator.userAgent);

                        const useSvgFabs = isWebKit
                            || /iPad|iPhone|iPod/i.test(navigator.userAgent)
                            || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);

                        const photoSize = 60;
                        const photoX = (memberWidth - photoSize) / 2;

                        const photoY = 8;

                        const spacingX = 17;
                        const leftSpacingX = 1;
                        const fabSize = useSvgFabs ? 28 : 24;
                        const gap = useSvgFabs ? 3 : 8;

                        const fabsHeight = (fabSize * 2) + gap;

                        const rightX = Math.round(photoX + photoSize + spacingX);

                        const leftX = Math.round(photoX - spacingX - fabSize - leftSpacingX);

                        const fabsY = Math.round(photoY + Math.round((photoSize - fabsHeight) / 2) - 4);

                        const r = fabSize / 2;
                        const cx = Math.round(rightX + fabSize / 2);
                        const cy = Math.round(fabsY + fabSize / 2);
                        const dy = fabSize + gap;

                        const lc = {
                            cx: Math.round(leftX + fabSize / 2),
                            cy: Math.round(fabsY + fabSize / 2),
                            r: r
                        };

                        const reportClickHandler = (event) => {
                            event?.preventDefault?.();
                            event?.stopPropagation?.();
                            openPersonReportCompose(
                                peopleDBUpdateRecipients,
                                portfolioDBUpdateRecipients,
                                member,
                                {firstLevel, secondLevel, thirdLevel}
                            ).then(() => console.log('report a change started'));
                        };
                        const companyGroupTalentUrl = getCompanyGroupTalentUrl(member);
                        const isInternal = isInternalCompany(member);


                        if (useSvgFabs) {
                            const reportG = group.append('g')
                                .attr('class', 'contact-fabs-svg contact-fabs--left')
                                .attr('transform', `translate(${lc.cx},${lc.cy})`);

                            const reportA = reportG.append('a')
                                .attr('href', '#')
                                .attr('target', '_blank')
                                .attr('rel', 'noopener noreferrer')
                                .attr('class', 'contact-fab report')
                                .attr('data-tooltip', 'Report change')
                                .attr('aria-label', 'Report change');

                            if (isInternal) {
                                const talentA_left = reportG.append('a')
                                    .attr('href', companyGroupTalentUrl)
                                    .attr('target', '_blank')
                                    .attr('rel', 'noopener noreferrer')
                                    .attr('class', 'contact-fab companygroup-talent')
                                    .attr('data-tooltip', 'Company Group Talent')
                                    .attr('aria-label', 'Company Group Talent');

                                const talentG_left = talentA_left.append('g')
                                    .attr('transform', `translate(0, ${dy})`);

                                talentG_left.append('circle').attr('r', lc.r).attr('class', 'fab-circle');
                                talentG_left.append('text')
                                    .attr('class', 'fab-emoji')
                                    .attr('text-anchor', 'middle')
                                    .attr('dominant-baseline', 'central')
                                    .text('👤');

                                talentA_left
                                    .on('pointerdown', (e) => e.stopPropagation())
                                    .on('touchstart', (e) => e.stopPropagation());
                            }

                            const reportBtn = reportA.append('g').attr('transform', 'translate(0,0)');
                            reportBtn.append('circle')
                                .attr('r', lc.r)
                                .attr('class', 'fab-circle');
                            reportBtn.append('text')
                                .attr('class', 'fab-emoji')
                                .attr('text-anchor', 'middle')
                                .attr('dominant-baseline', 'central')
                                .text('📝');

                            reportA
                                .on('pointerdown', (e) => e.stopPropagation())
                                .on('touchstart', (e) => e.stopPropagation())
                                .on('click', reportClickHandler);

                            if (member[emailField]) {
                                const fabsG = group.append('g')
                                    .attr('class', 'contact-fabs-svg contact-fabs--right')
                                    .attr('transform', `translate(${cx},${cy})`);

                                const chatA = fabsG.append('a')
                                    .attr('href', `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(email)}`)
                                    .attr('target', '_blank')
                                    .attr('rel', 'noopener noreferrer')
                                    .attr('class', 'contact-fab chat')
                                    .attr('data-tooltip', 'Chat')
                                    .attr('aria-label', 'Chat');

                                const chatG = chatA.append('g').attr('transform', 'translate(0,0)');
                                chatG.append('circle').attr('r', r).attr('class', 'fab-circle');
                                chatG.append('text')
                                    .attr('class', 'fab-emoji')
                                    .attr('text-anchor', 'middle')
                                    .attr('dominant-baseline', 'central')
                                    .text('💬');
                                const mailA = fabsG.append('a')
                                    .attr('href', createOutlookUrl([email]))
                                    .attr('target', '_blank')
                                    .attr('rel', 'noopener noreferrer')
                                    .attr('class', 'contact-fab mail')
                                    .attr('data-tooltip', 'Send email')
                                    .attr('aria-label', 'Send email');


                                const mailG = mailA.append('g').attr('transform', `translate(0, ${dy})`);
                                mailG.append('circle').attr('r', r).attr('class', 'fab-circle');
                                mailG.append('text')
                                    .attr('class', 'fab-emoji')
                                    .attr('text-anchor', 'middle')
                                    .attr('dominant-baseline', 'central')
                                    .text('✉️');

                                fabsG.selectAll('a.contact-fab')
                                    .on('pointerdown', (event) => event.stopPropagation())
                                    .on('touchstart', (event) => event.stopPropagation());
                            }
                        } else {
                            const leftColumnCount = isInternal ? 2 : 1;
                            const leftFabsHeight = (fabSize * leftColumnCount) + (gap * (leftColumnCount - 1));


                            const fabsLeft = group.append('foreignObject')
                                .attr('x', leftX)
                                .attr('y', fabsY)
                                .attr('width', fabSize)
                                .attr('height', leftFabsHeight)
                                .attr('pointer-events', 'all')
                                .style('overflow', 'visible')
                                .append('xhtml:div')
                                .attr('class', 'contact-fabs contact-fabs--left');

                            fabsLeft.append('a')
                                .attr('class', 'contact-fab report')
                                .attr('href', '#')
                                .attr('data-tooltip', 'Report change')
                                .attr('aria-label', 'Report change')
                                .html(`<span class="icon" aria-hidden="true">📝</span>`)
                                .on('click', reportClickHandler);

                            if (isInternal) {
                                fabsLeft.append('a')
                                    .attr('class', 'contact-fab companygroup-talent')
                                    .attr('href', companyGroupTalentUrl)
                                    .attr('target', '_blank')
                                    .attr('rel', 'noopener noreferrer')
                                    .attr('data-tooltip', 'Company Group Talent')
                                    .attr('aria-label', 'Company Group Talent')
                                    .html(`<span class="icon" aria-hidden="true">👤</span>`);
                            }

                            if (member[emailField]) {
                                const fabs = group.append('foreignObject')
                                    .attr('x', rightX)
                                    .attr('y', fabsY)
                                    .attr('width', fabSize)
                                    .attr('height', fabsHeight)
                                    .attr('pointer-events', 'all')
                                    .style('overflow', 'visible')
                                    .append('xhtml:div')
                                    .attr('class', 'contact-fabs contact-fabs--right');

                                fabs.append('a')
                                    .attr('class', 'contact-fab chat')
                                    .attr('href', `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(email)}`)
                                    .attr('target', '_blank')
                                    .attr('rel', 'noopener noreferrer')
                                    .attr('data-tooltip', 'Chat')
                                    .attr('aria-label', 'Chat')
                                    .html(`<span class="icon" aria-hidden="true">💬</span>`);

                                fabs.append('a')
                                    .attr('class', 'contact-fab mail')
                                    .attr('href', createOutlookUrl([email]))
                                    .attr('target', '_blank')
                                    .attr('rel', 'noopener noreferrer')
                                    .attr('data-tooltip', 'Send email')
                                    .attr('aria-label', 'Send email')
                                    .html(`<span class="icon" aria-hidden="true">✉️</span>`);
                            }

                        }

                        group.classed('card', true);
                        group.selectAll('.contact-fabs-svg, .contact-fabs').each(function () {
                            this.parentNode.appendChild(this);
                        });
                        wireFabsInteractions(group);

                        Object.entries(member).forEach(([key, value]) => {
                            if (fieldsToShow.includes(key) && value) {
                                let finalValue = value;

                                if (dateValues.includes(key)) {
                                    const parsed = new Date(value);
                                    if (!isNaN(parsed)) {
                                        finalValue = formatMonthYear(parsed);
                                    }

                                }

                                infoDiv.append('div')
                                    .attr('class', key.toLowerCase() + '-field')
                                    .html(`<strong>${key}:</strong> ${finalValue}`);
                            }
                        });
                    });
                });

                // Avanza X per il prossimo theme della stessa riga
                secondLevelX += themeWidth + secondLevelBoxPadX;
            });

            // Avanza Y per la riga successiva
            secondLevelYBase += themeBoxHeightRow + secondLevelRowPadY;
        });

        // Avanza Y per lo stream successivo
        streamY += firstLevelBoxHeight + firstLevelBoxPadY;
    });

    requestAnimationFrame(() => {
        placeCompanyLogoUnderDiagram('./assets/company-logo.png', 200, 50);
    });

    fitToContent(0.9);

    applyDraggableToggleState();
    requestAnimationFrame(() => {
        setColorMode(ROLE_FIELD_WITH_MAPPING);
    });
}

document.getElementById('fileInput')?.addEventListener('change', function (e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = function (evt) {
        resetVisualization();
        extractData(evt.target.result);
    };
    reader.readAsText(file, 'UTF-8');
});

(function setupGlobalTooltip() {
    let tipEl = null;
    let showTimer = null;
    let hideTimer = null;
    let currentAnchor = null;

    const SHOW_DELAY = 90;
    const HIDE_DELAY = 140;

    const isMouseLike = window.matchMedia('(hover: hover) and (pointer: fine)').matches;

    function ensureTip() {
        if (!tipEl) {
            tipEl = document.createElement('div');
            tipEl.className = 'solitaire-tooltip';
            document.body.appendChild(tipEl);
        }
        tipEl.style.zIndex = String(2147483647);
        return tipEl;
    }

    function isVisible() {
        return !!(tipEl && tipEl.classList.contains('show'));
    }

    function positionTip(anchor, placement = 'right') {
        const el = ensureTip();
        const rect = anchor.getBoundingClientRect();

        let x = rect.right + 8;
        let y = rect.top + rect.height / 2;
        el.style.transform = 'translate(0, -50%)';

        if (placement === 'top') {
            x = rect.left + rect.width / 2;
            y = rect.top - 8;
            el.style.transform = 'translate(-50%, -8px)';
        } else if (placement === 'bottom') {
            x = rect.left + rect.width / 2;
            y = rect.bottom + 8;
            el.style.transform = 'translate(-50%, 8px)';
        } else if (placement === 'left') {
            x = rect.left - 8;
            y = rect.top + rect.height / 2;
            el.style.transform = 'translate(-100%, -50%)';
        }

        el.style.left = `${Math.round(x)}px`;
        el.style.top = `${Math.round(y)}px`;
    }

    function showTip(text, anchor, placement = 'right') {
        const el = ensureTip();
        el.textContent = text || '';
        el.classList.add('show');
        positionTip(anchor, placement);
    }

    function hideTipNow() {
        if (tipEl) tipEl.classList.remove('show');
    }

    function getFabAnchor(target) {
        return target?.closest?.('[data-tooltip], .contact-fab') || null;
    }

    if (isMouseLike) {
        document.addEventListener('mouseover', (e) => {
            const a = getFabAnchor(e.target);
            if (!a) return;

            const text = a.getAttribute('data-tooltip') || a.getAttribute('aria-label') || '';
            if (!text) return;

            clearTimeout(hideTimer);
            hideTimer = null;

            if (isVisible() && currentAnchor !== a) {
                currentAnchor = a;
                showTip(text, a, 'right');
                return;
            }

            currentAnchor = a;
            clearTimeout(showTimer);
            showTimer = setTimeout(() => showTip(text, a, 'right'), SHOW_DELAY);
        }, true);

        document.addEventListener('mouseout', (e) => {
            const a = getFabAnchor(e.target);
            if (!a) return;

            clearTimeout(showTimer);

            clearTimeout(hideTimer);
            hideTimer = setTimeout(() => {
                hideTipNow();
                currentAnchor = null;
            }, HIDE_DELAY);
        }, true);

        window.addEventListener('scroll', () => {
            if (isVisible()) hideTipNow();
        }, {passive: true});
        window.addEventListener('resize', () => {
            if (isVisible()) hideTipNow();
        });
        window.addEventListener('dsm-canvas-zoom', () => {
            if (isVisible()) hideTipNow();
        });
    } else {
        document.addEventListener('pointerdown', hideTipNow, {passive: true});
    }
})();


document.getElementById('drawer-search-input')?.addEventListener('keydown', function (e) {
    if (e.key === 'Enter') {
        const query = e.target.value.trim().toLowerCase();
        if (query) {
            searchByQuery(query);
        } else {
            clearSearch();
        }
        e.preventDefault();
    }
});

function searchByQuery(query, opts = {}) {
    const q = (query ?? '').toString().trim().toLowerCase();
    const scopeField = (opts.field || '').toLowerCase();
    const missing = !!opts.missing;
    const noZoom = !!opts.noZoom;

    if (!q && !missing) {
        clearSearch();
        return;
    }

    const searchInput = document.getElementById('drawer-search-input');
    if (searchInput && searchInput.value.trim().toLowerCase() !== q) {
        searchInput.value = q;
    }

    const FIELD_SELECTORS = {
        role: '.role-field',
        company: '.company-field',
        location: '.location-field'
    };

    function normalizeFieldName(f) {
        const fLow = (f || '').toLowerCase();
        if (fLow.includes('role')) return 'role';
        if (fLow.includes('company')) return 'company';
        if (fLow.includes('location')) return 'location';
        return '';
    }

    const normalizedField = normalizeFieldName(scopeField);

    let nodes = [];
    let matches = [];

    if (missing && normalizedField) {
        const attrName =
            normalizedField === 'role' ? 'data-role' :
                normalizedField === 'company' ? 'data-company' : 'data-location';

        nodes = Array.from(document.querySelectorAll(`g[data-key^="card::"]`));

        matches = nodes.filter(n => {
            const raw = (n.getAttribute(attrName) || '');
            const norm = normalizeWs(raw).trim().toLowerCase();
            return !norm || UNKNOWN_MATCHER.test(norm);
        });

    } else {
        const nodesSelector = (normalizedField && FIELD_SELECTORS[normalizedField])
            ? FIELD_SELECTORS[normalizedField]
            : '.profile-name, .team-title, .theme-title, .stream-title, .role-field, .company-field, .location-field, [data-services]';

        nodes = Array.from(document.querySelectorAll(nodesSelector));

        matches = nodes.filter(n => {
            const txt = (n.textContent || '').trim().toLowerCase();
            const textMatch = txt.includes(q);
            const attrMatch = (n.getAttribute?.('data-services') || '').toLowerCase().includes(q);
            return textMatch || attrMatch;
        });
    }

    if (matches.length === 0) {
        clearSearchDimming();
        showToast(missing ? 'No result found for Unknown' : `No result found for ${q}`);
        return;
    }

    if (q === lastSearch && !missing) {
        currentIndex = (currentIndex + 1) % matches.length;
    } else {
        lastSearch = q;
        currentIndex = 0;
    }

    const target = matches[currentIndex];
    clearFieldHighlights();
    closeDrawer();

    if (!missing) {
        const zoomTarget = (target.matches?.('g[data-key^="card::"]'))
            ? (target.querySelector('.profile-box') || target)
            : target;

        if (!noZoom) {
            zoomToElement(zoomTarget, 1, 600);
        }
        applySearchDimmingForMatches(matches);
        showToast(`Found ${matches.length} result(s). Showing ${currentIndex + 1}/${matches.length}.`);
        setSearchQuery(q);

        const FIELD_CLASSES = ['role-field', 'company-field', 'location-field'];
        if (zoomTarget.classList) {
            const hitClass = FIELD_CLASSES.find(c => zoomTarget.classList.contains(c));
            if (hitClass) {
                zoomTarget.classList.add('field-hit-highlight');
            } else {
                const group = zoomTarget.closest('g[data-key^="card::"]');
                if (group) {
                    FIELD_CLASSES.forEach(cls => {
                        const el = group.querySelector('.' + cls);
                        if (!el) return;
                        const tn = (el.textContent || '').toLowerCase();
                        if (tn.includes(q)) el.classList.add('field-hit-highlight');
                    });
                }
            }
        }

        const roleMapping = roleDetailsMapping.get(query);
        if (scopeField?.toLowerCase() === "role" && roleMapping) {
            openDrawer({
                name: query,
                description: roleMapping["description"],
                elements: { items: (roleMapping["grants"] || '')
                        .split(',').map(s => s.trim()).filter(Boolean) }, elementsTitle: "Role Grants"
            });
        }
        try {
            const group = zoomTarget.closest('g');
            const teamTitleEl = group ? group.querySelector('text.team-title') : null;
            if (!teamTitleEl) return;

            const rawServices = (teamTitleEl.getAttribute('data-services') || '')
                .split(',').map(s => s.trim()).filter(Boolean);
            if (rawServices.length === 0) return;

            const norm = v => (v || '').toString().trim().toLowerCase();
            const normalized = rawServices.map(s => ({ raw: s, norm: norm(s) }));
            const hit = normalized.find(svc => svc.norm.includes(q));
            if (!hit) return;

            const teamName =
                teamTitleEl.getAttribute('data-team-name') || getNameFromTitleEl(teamTitleEl);
            const email = teamTitleEl.getAttribute('data-team-email') || '';
            const channels = (() => {
                try { return JSON.parse(teamTitleEl.getAttribute('data-team-channels') || '[]'); }
                catch { return []; }
            })();
            const description = teamTitleEl.getAttribute('data-team-description') || '';

            openDrawer({
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
    } else {
        applySearchDimmingForMatches(matches);
        showToast(`Found ${matches.length} result(s).`);
        setSearchQuery(q);
    }
}

(function enableLegendDragOnce() {
    // evita di aggiungere infiniti listener di resize
    let resizeAttached = false;

    window.enableLegendDrag = function enableLegendDrag({ handleSelector = null } = {}) {
        const root = document.getElementById('legend-root');
        if (!root) return;                    // ✅ rimuovi il guard su "attached" per permettere il rebind

        const LS_KEY = 'legend-pos-v1';
        const clamp = (v, min, max) => Math.max(min, Math.min(max, v));

        function getRootRect() { return root.getBoundingClientRect(); }
        function getViewportSize() { return { w: document.documentElement.clientWidth, h: document.documentElement.clientHeight }; }

        // restore() e save() rimangono invariati, tieni il tuo codice qui
        function restore() {
            try {
                const saved = JSON.parse(localStorage.getItem(LS_KEY) || '{}');
                if (typeof saved.x === 'number' && typeof saved.y === 'number') {
                    root.style.left = `${saved.x}px`;
                    root.style.top  = `${saved.y}px`;
                    root.style.right = 'auto';
                    root.style.bottom = 'auto';
                    return true;
                }
            } catch {}
            return false;
        }

        function save(x, y) {
            localStorage.setItem(LS_KEY, JSON.stringify({ x, y }));
        }

        const restored = restore();
        if (!restored) {
            const cs = getComputedStyle(root);
            const hasAnyInlineAnchor = root.style.left || root.style.top || root.style.right || root.style.bottom;
            if (!hasAnyInlineAnchor) {
                // lascia agire il CSS di default
            }
        }

        function reclamp() {
            const rect = getRootRect();
            const { w, h } = getViewportSize();
            let nx = rect.left;
            let ny = rect.top;

            const usingBottom = root.style.bottom && !root.style.top;
            const usingRight  = root.style.right  && !root.style.left;

            if (usingBottom || usingRight) {
                nx = clamp(rect.left, 0, w - rect.width);
                ny = clamp(rect.top,  0, h - rect.height);
                root.style.left = `${nx}px`;
                root.style.top  = `${ny}px`;
                root.style.right = 'auto';
                root.style.bottom = 'auto';
            } else {
                nx = clamp(rect.left, 0, w - rect.width);
                ny = clamp(rect.top,  0, h - rect.height);
                root.style.left = `${nx}px`;
                root.style.top  = `${ny}px`;
            }
            save(nx, ny);
        }
        requestAnimationFrame(reclamp);

        // —— Drag con soglia (come nel tuo codice) ——
        const THRESHOLD = 4;
        let startX = 0, startY = 0;
        let startLeft = 0, startTop = 0;
        let dragging = false;
        let pointerId = null;

        const handle = handleSelector ? root.querySelector(handleSelector) : root;
        const dragClassEl = handleSelector ? handle : root;
        if (!handle) return; // nel raro caso in cui il titolo non esista ancora

        function onPointerDown(e) {
            if (e.button !== 0) return;
            pointerId = e.pointerId;

            const rect = getRootRect();
            const cs = window.getComputedStyle(root);
            startLeft = parseFloat(cs.left) || rect.left;
            startTop  = parseFloat(cs.top)  || rect.top;
            startX = e.clientX;
            startY = e.clientY;

            window.addEventListener('pointermove', onPointerMove, { passive: true });
            window.addEventListener('pointerup', onPointerUp, { passive: true });
        }

        function onPointerMove(e) {
            const dx = e.clientX - startX;
            const dy = e.clientY - startY;

            if (!dragging) {
                if (Math.abs(dx) <= THRESHOLD && Math.abs(dy) <= THRESHOLD) return;
                dragging = true;
                dragClassEl.classList.add('is-dragging');
                try { dragClassEl.setPointerCapture?.(pointerId); } catch {}
            }

            let nextLeft = startLeft + dx;
            let nextTop  = startTop  + dy;

            const { w, h } = getViewportSize();
            const r = getRootRect();
            const bw = r.width, bh = r.height;

            nextLeft = clamp(nextLeft, 0, w - bw);
            nextTop  = clamp(nextTop,  0, h - bh);

            root.style.left = `${nextLeft}px`;
            root.style.top  = `${nextTop}px`;
            root.style.right = 'auto';
            root.style.bottom = 'auto';
        }

        function onPointerUp() {
            window.removeEventListener('pointermove', onPointerMove);
            window.removeEventListener('pointerup', onPointerUp);

            if (!dragging) return;

            dragClassEl.classList.remove('is-dragging');
            const rect = getRootRect();
            const { w, h } = getViewportSize();
            const nx = clamp(rect.left, 0, w - rect.width);
            const ny = clamp(rect.top,  0, h - rect.height);
            save(nx, ny);
            dragging = false;
            pointerId = null;
        }

        handle.style.touchAction = 'none';
        handle.addEventListener('pointerdown', onPointerDown);

        if (!resizeAttached) {
            window.addEventListener('resize', () => reclamp());
            resizeAttached = true;
        }
    };
})();