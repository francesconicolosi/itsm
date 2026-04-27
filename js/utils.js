import * as d3 from 'd3';

export const SECOND_LEVEL_LABEL_EXTRA = 120;
export const TEAM_MEMBER_LEGENDA_LABEL = 'Team Member';

export const firstOrgLevel = 'Team Stream';
export const secondOrgLevel = 'Team Theme';
export const thirdOrgLevel = 'Team member of';
export const firstLevelNA = `No ${firstOrgLevel}`;
export const secondLevelNA = `No ${secondOrgLevel}`;
export const thirdLevelNA = `No ${thirdOrgLevel}`;
export const ROLE_FIELD_WITH_MAPPING = 'Role';
export const LOCATION_FIELD = 'Location';
export const COMPANY_FIELD = 'Company';
export const NEUTRAL_COLOR = '#fcfcfc';
export const MAX_TEAMS_PER_ROW = 5;
export const emailField = "Company email";

export function buildExpandedLayoutMapFromDom() {
    const map = {};

    document.querySelectorAll('g.draggable[data-key]').forEach(el => {
        const key = el.getAttribute('data-key');
        if (!key) return;

        // position
        let x = 0, y = 0;
        const t = el.getAttribute('transform') || '';
        const m = t.match(/translate\(([^,]+),\s*([^)]+)\)/);
        if (m) {
            x = Math.round(parseFloat(m[1]) || 0);
            y = Math.round(parseFloat(m[2]) || 0);
        }

        const entry = { x, y };

        // size (if exists)
        const rect = el.querySelector('rect');
        if (rect) {
            const w = Number(rect.getAttribute('width'));
            const h = Number(rect.getAttribute('height'));
            if (Number.isFinite(w)) entry.width = Math.round(w);
            if (Number.isFinite(h)) entry.height = Math.round(h);
        }

        map[key] = entry;
    });

    return map;
}

export function enableGlobalFindShortcut({
                                             inputSelector,
                                             onFocus,
                                             selectText = true
                                         } = {}) {
    if (!inputSelector) {
        console.warn('[enableGlobalFindShortcut] inputSelector is required');
        return;
    }

    window.addEventListener('keydown', (e) => {
        const isMac = navigator.platform.toUpperCase().includes('MAC');

        const isFindShortcut =
            (isMac && e.metaKey && e.key.toLowerCase() === 'f') ||
            (!isMac && e.ctrlKey && e.key.toLowerCase() === 'f');

        if (!isFindShortcut) return;

        const activeTag = document.activeElement?.tagName;
        const isTyping =
            activeTag === 'INPUT' ||
            activeTag === 'TEXTAREA' ||
            document.activeElement?.isContentEditable;

        // se sto già scrivendo in un input, non forzare nulla
        if (isTyping) return;

        const input = document.querySelector(inputSelector);
        if (!input) return;

        e.preventDefault();
        e.stopPropagation();

        input.focus({ preventScroll: false });

        if (selectText && typeof input.select === 'function') {
            try { input.select(); } catch {}
        }

        if (typeof onFocus === 'function') {
            onFocus(input);
        }
    }, true);
}

export function isDateTimeValue(value) {
    if (typeof value !== 'string') return false;

    const isoRegex =
        /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?(Z|[+-]\d{2}:\d{2})$/;

    if (!isoRegex.test(value)) return false;

    const d = new Date(value);
    return !isNaN(d.getTime());
}

export function formatDateTimeLocal(value) {
    const d = new Date(value);
    if (isNaN(d.getTime())) return value;

    return new Intl.DateTimeFormat(undefined, {
        year: 'numeric',
        month: 'short',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        timeZoneName: 'short'
    }).format(d);
}



export function countRowsByTeamCapacity(secondLevelItems, capacityPerRow) {
    let rows = 1;
    let used = 0;

    for (const [themeName, themeObj] of Object.entries(secondLevelItems)) {
        if (themeName.includes(secondLevelNA)) continue;

        const nTeams = Object.keys(themeObj || {}).length || 0;

        if (used > 0 && (used + nTeams) > capacityPerRow) {
            rows++;
            used = 0;
        }
        used += nTeams;
    }
    return rows;
}

export function computeStreamBoxWidthByCapacity(
    secondLevelItems,
    secondLevelBoxPadX,
    secondLevelNA,
    thirdLevelBoxPadX,
    thirdLevelBoxWidth,
    labelExtra = SECOND_LEVEL_LABEL_EXTRA,
    leftPad = 60,
    rightPad = 60
) {
    let used = 0;
    let rowWidth = leftPad;
    let maxWidth = 0;

    for (const [themeName, themeObj] of Object.entries(secondLevelItems)) {
        if (themeName.includes(secondLevelNA)) continue;
        const nTeams = Object.keys(themeObj || {}).length || 0;

        const themeInnerGaps = Math.max(0, nTeams - 1) * thirdLevelBoxPadX;
        const themeWidth = (nTeams * thirdLevelBoxWidth) + themeInnerGaps + labelExtra;

        if (used > 0 && (used + nTeams) > MAX_TEAMS_PER_ROW) {
            maxWidth = Math.max(maxWidth, rowWidth + rightPad);
            rowWidth = leftPad;
            used = 0;
        }

        if (used > 0) rowWidth += secondLevelBoxPadX;
        rowWidth += themeWidth;
        used += nTeams;
    }

    maxWidth = Math.max(maxWidth, rowWidth + rightPad);
    return maxWidth;
}

export function computeKeysAndCountsFromVisibleOrg(org, fieldName) {
    const counts = new Map();
    const seen = new Set(); // NEW: deduplica per persona
    const allowed = getAllowedStreamsSet(); // NEW: rispetta ?stream=

    const isAllowedStream = (s) => {
        if (!allowed || allowed.size === 0) return true;
        const n1 = (s ?? '').toString().trim();
        const n2 = normalizeKey(n1);
        return allowed.has(n1) || allowed.has(n2);
    };

    const pickKey = (m) => {
        if (!m) return 'Unknown';
        if (fieldName === ROLE_FIELD_WITH_MAPPING)   return normalizeWs(m?.[ROLE_FIELD_WITH_MAPPING]) || 'Unknown';
        if (fieldName === COMPANY_FIELD)             return normalizeWs(m?.[COMPANY_FIELD])           || 'Unknown';
        if (fieldName === LOCATION_FIELD)            return normalizeWs(m?.[LOCATION_FIELD])          || 'Unknown';
        return normalizeWs(m?.[fieldName]) || 'Unknown';
    };

    Object.entries(org || {}).forEach(([first, themes]) => {
        if ((first || '').includes(firstLevelNA)) return;
        if (!isAllowedStream(first)) return; // NEW: esclude stream non visibili

        Object.entries(themes || {}).forEach(([second, teams]) => {
            if ((second || '').includes(secondLevelNA)) return;

            Object.entries(teams || {}).forEach(([third, members]) => {
                if ((third || '').includes(thirdLevelNA)) return;

                (members || []).forEach(m => {
                    // NEW: deduplica persona su tutta la vista
                    const id = buildCompositeKey(m, emailField);
                    if (id && seen.has(id)) return;
                    if (id) seen.add(id);

                    const key = pickKey(m);
                    counts.set(key, (counts.get(key) || 0) + 1);
                });
            });
        });
    });

    const keys = Array.from(counts.keys())
        .sort((a, b) => (counts.get(b) - counts.get(a)) || a.localeCompare(b, 'en', { sensitivity: 'base' }));

    let topKey = null, max = -1;
    for (const [k, c] of counts) {
        if (c > max) { max = c; topKey = k; }
    }

    return { keys, counts, topKey };
}

export function computeKeysAndCounts(members, fieldName) {
    const counts = new Map();
    for (const m of members || []) {
        const k = normalizeWs(m?.[fieldName]);
        const key = k || 'Unknown';
        counts.set(key, (counts.get(key) || 0) + 1);
    }
    const keys = Array.from(counts.keys())
        .sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' }));
    let topKey = null, max = -1;
    for (const [k, c] of counts) {
        if (c > max) { max = c; topKey = k; }
    }
    return { keys, counts, topKey };
}

const MOST_FREQUENT_FIXED_COLOR = '#ffffff';

export function makeKeyColorScale(keys, topKey) {
    const palette = d3.schemeTableau10;
    const map = new Map();
    keys.forEach((k, i) => map.set(k, palette[i % palette.length]));
    if (topKey) map.set(topKey, MOST_FREQUENT_FIXED_COLOR);

    const scale = (k) => map.get(k) || NEUTRAL_COLOR;
    // helper opzionali
    scale.domain = () => keys.slice();
    scale.colorOf = (k) => map.get(k) || NEUTRAL_COLOR;
    return scale;
}

export function getLegendTitleFor(fieldName) {
    if (fieldName === ROLE_FIELD_WITH_MAPPING)   return 'Roles';
    if (fieldName === COMPANY_FIELD)             return 'Companies';
    if (fieldName === LOCATION_FIELD)            return 'Locations';
    return 'Legend';
}


export const LABEL_FOR_KEY = {
    id: 'ID'
};

const fullNormalizeWs = (s) => (s ?? '')
    .toString()
    .replace(/\s+/g, ' ')
    .trim();

export const FIELDS_WITH_WS_NORMALIZATION = new Set([
    'Name',
    'User',
    'Email'
]);

export const normalizeWs = (value, fieldName) => {
    const raw = (value ?? '').toString();
    return !fieldName || FIELDS_WITH_WS_NORMALIZATION.has(fieldName) ? fullNormalizeWs(raw) : raw.trim();
};


export function labelForKey(key) {
    if (key === 'id') return LABEL_FOR_KEY.id;
    return key;
}

export function isListViewVisible() {
    const el = document.getElementById('list-view');
    if (!el) return false;
    const style = window.getComputedStyle(el);
    return style.display !== 'none' && style.visibility !== 'hidden' && el.offsetParent !== null;
}

let searchActive = false;

const URL_RE = /^https?:\/\/\S+$/i;
export function isUrl(v) {
    return URL_RE.test(v);
}

export function splitValues(raw) {
    if (!raw) return [];

    return raw
        .toString()
        .split(/\s*\|\|\s*|\n|,/)
        .map(s => s.trim())
        .filter(Boolean);
}
export function refreshDrawerColumnIcons() {
    const drawerContent = document.getElementById('drawerContent');
    if (!drawerContent) return;
    const isListVisible = isListViewVisible();
    const buttons = drawerContent.querySelectorAll('button.col-op');
    buttons.forEach(btn => {
        const col = decodeURIComponent(btn.getAttribute('data-col'));
        const selected = window.currentColumnKeys.includes(col);
        btn.textContent = selected ? '−' : '+';
        btn.setAttribute('aria-label', selected
            ? `Remove "${labelForKey(col)}" from list view`
            : `Add "${labelForKey(col)}" to list view`
        );
        btn.style.display = isListVisible ? '' : 'none';
    });
}

export function getCellValue(node, key) {
    if (key === 'id') return node?.id ?? '';

    const raw = node?.[key] ?? '';
    if (key === 'Depends on' && typeof raw === 'string') {
        return raw.split('\n').map(s => s.trim()).filter(Boolean).join(', ');
    }
    if (key === 'Decommission Date' && raw) {
        const d = new Date(raw);
        return isNaN(d.getTime()) ? raw : getFormattedDate(d.toISOString());
    }
    return raw;
}


export function clearFieldHighlights() {
    document
        .querySelectorAll('.field-hit-highlight, .role-hit-highlight')
        .forEach(el => el.classList.remove('field-hit-highlight', 'role-hit-highlight'));
}

export function countTeamsForMemberInOrg(member, org, emailField = 'Company email') {
    const targetKey = buildCompositeKey(member, emailField);
    if (!targetKey) return 0;

    let count = 0;

    for (const [streamName, themes] of Object.entries(org || {})) {

        if (streamName.toLowerCase().includes(firstLevelNA.toLowerCase())) continue;
        for (const [themeName, teams] of Object.entries(themes || {})) {
            if (themeName.toLowerCase().includes(secondLevelNA.toLowerCase())) continue;
            for (const members of Object.values(teams || {})) {
                const found = (members || []).some(
                    m => buildCompositeKey(m, emailField) === targetKey
                );
                if (found) count++;
            }
        }
    }

    return count;
}

export function filterOrganizationByStreams(org, allowed) {
    if (!allowed || allowed.size === 0) return org;
    const out = {};
    for (const [stream, themes] of Object.entries(org || {})) {
        const ok = allowed.has(stream) || allowed.has(normalizeKey(stream));
        if (ok) out[stream] = themes;
    }
    return out;
}


export function collectMembersFromOrganization(filteredOrg) {
    const out = [];
    for (const themes of Object.values(filteredOrg)) {
        for (const teams of Object.values(themes)) {
            for (const members of Object.values(teams)) {
                out.push(...members);
            }
        }
    }
    return out;
}

export function getNameFromTitleEl(teamTitleEl) {
    const raw = teamTitleEl?.textContent || '';
    return raw.replace(/\s*-\s*⚙️.*$/, '').trim();
}

function splitFirstLast(fullName = '') {
    const parts = (fullName || '').trim().split(/\s+/).filter(Boolean);
    if (parts.length === 0) return { firstName: '', lastName: '' };
    if (parts.length === 1) return { firstName: parts[0], lastName: '' };
    return { firstName: parts[0], lastName: parts.slice(1).join(' ') };
}

export function buildPersonReportBody(member = {}, ctx) {
    const { firstLevel, secondLevel, thirdLevel } = ctx || {};
    const name = (member['Name'] || '').trim();
    const { firstName, lastName } = splitFirstLast(name);

    const company = member['Company'] || member['Company Name'] || '';
    const role = member['Role'] || member['Role Name'] || member['Role Title'] || '';
    const startDate = member['In team since'] || '';
    const location = member['Location'] || '';
    const room = member['Room Link'] || member['Room'] || '';
    const lineManager = member['Line Manager'] || member['Manager'] || '';
    const team = thirdLevel || member['Team member of'] || [firstLevel, secondLevel, thirdLevel].filter(Boolean).join(' / ');

    return [
        'Hello,',
        '',
        'I would like to report the need for an update to the People Database:',
        '',
        `FIRST NAME: ${firstName}`,
        `LAST NAME: ${lastName}`,
        `COMPANY NAME: ${company}`,
        `TEAM: ${team}`,
        `ROLE: ${role}`,
        `START DATE (for new Joiners or movers): ${startDate}`,
        `END DATE (for leavers): `,
        `LOCATION: ${location}`,
        `ROOM: ${room}`,
        `LINE MANAGER: ${lineManager}`,
        `PHOTO: `,
        '',
        'Regards,'
    ].join('\n');
}

export function createModal({ title, html, buttons }) {
    return new Promise(resolve => {
        const overlay = document.createElement('div');
        overlay.className = 'simple-modal__overlay';
        overlay.setAttribute('role', 'dialog');
        overlay.setAttribute('aria-modal', 'true');

        const modal = document.createElement('div');
        modal.className = 'simple-modal';
        const btnHtml = buttons.map(btn =>
            `<button type="button" class="simple-modal__btn ${btn.primary ? 'simple-modal__btn--primary' : ''}" data-action="${btn.id}">
        ${btn.label}
      </button>`
        ).join('');
        modal.innerHTML = `
            <h3>${title}</h3>
      <p>${html}</p>
      <div class="simple-modal__buttons">${btnHtml}</div>
    `;

        function close(val) {
            overlay.remove();
            resolve(val);
        }

        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) close(null);
        });

        buttons.forEach(btn => {
            modal.querySelector(`[data-action="${btn.id}"]`)
                ?.addEventListener('click', () => close(btn.id));
        });

        const escHandler = (e) => {
            if (e.key === "Escape") {
                document.removeEventListener('keydown', escHandler);
                close(null);
            }
        };
        document.addEventListener('keydown', escHandler);

        overlay.appendChild(modal);
        document.body.appendChild(overlay);

        modal.querySelector(`[data-action="${buttons.find(b => b.primary)?.id}"]`)?.focus();
    });
}

export function askModal() {
    return createModal({
        title: 'Include Portfolio Team?',
        html: `Notify also the Portfolio Team about changes to Team, Start Date, or End Date? If not included, only Service Management will be informed and will handle the update.`,
        buttons: [
            { id: 'cancel', label: 'Cancel' },
            { id: 'skip', label: "Don't include" },
            { id: 'include', label: 'Include', primary: true }
        ]
    }).then(answer => {
        if (answer === 'include') return true;
        if (answer === 'skip') return false;
        return null;
    });
}

export function applyStreamVisibility({ hiddenStreams, isolatedStream }) {
    d3.selectAll('g[data-key^="stream::"]').each(function () {
        const g = d3.select(this);
        const key = g.attr('data-key');

        const isHidden = hiddenStreams.has(key);
        const isIsolated =
            isolatedStream && key !== isolatedStream;

        g.style('display', (isHidden || isIsolated) ? 'none' : null);
    });
}

export function askHideStreamModal(streamName) {
    return createModal({
        title: `Hide stream "${streamName}"?`,
        html: `
      This stream will be temporarily hidden.<br><br>
      The URL in your browser bar will update and can be reused as a permalink
      to load this filtered view.<br><br>
      To restore the full view, click the ❌ next to the search bar.
    `,
        buttons: [
            { id: 'cancel', label: 'Cancel' },
            { id: 'confirm', label: 'Hide stream', primary: true }
        ]
    }).then(answer => answer === 'confirm');
}

export async function openPersonReportCompose(peopleDBUpdateRecipients, portfolioDBUpdateRecipients, member, ctx) {
    const to = [...(Array.isArray(peopleDBUpdateRecipients) ? peopleDBUpdateRecipients : [])];
    const decision = await askModal();

    if (decision === null) {
        closeSideDrawer();
        return;
    }

    if (decision === true) {
        to.push(...portfolioDBUpdateRecipients);
    }

    const subject = `Request for People Database Update - ${member?.Name ?? ''}`;
    const body = buildPersonReportBody(member, ctx);

    try {

        if (isMobileDevice()) {
            // fallback mailto per mobile
            buildFallbackMailToLink(to, subject, body);
        } else {
            openOutlookWebCompose({
                to,
                cc: [],
                bcc: [],
                subject,
                body
            });
        }
    } catch (e) {
        console.warn('openPersonReportCompose error:', e);
        buildFallbackMailToLink(to, subject, body);
    }
}

function normalizeLower(s) {
    return (s || '').toString().trim().toLowerCase();
}

export function getAllowedStreamsSet() {
    const streamFilterParam = getQueryParam('stream');
    if (!streamFilterParam) return null;

    const items = streamFilterParam
        .split(',')
        .map(s => s.trim())
        .filter(Boolean);

    const set = new Set();
    items.forEach(x => {
        set.add(x);
        set.add(normalizeKey(x));
    });
    return set;
}

export function normalizeKey(s) {
    return (s ?? '')
        .toString()
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '_')
        .replace(/[^a-z0-9_-]/g, '');
}

export function getVisiblePeopleForLegend(people, allowedStreams, firstOrgLevel) {
    if (!allowedStreams || allowedStreams.size === 0) return people;

    return people.filter(p => {
        const raw = (p[firstOrgLevel] || '').toString().trim();
        if (!raw) return false;

        const items = raw.split(/\n|,/).map(s => s.trim()).filter(Boolean);
        if (items.length === 0) return false;

        return items.some(item => {
            const n1 = item;
            const n2 = normalizeKey(item);
            return allowedStreams.has(n1) || allowedStreams.has(n2);
        });
    });
}

export function buildCompositeKey(person, emailField) {
    const name = normalizeWs(person?.Name);
    const email = normalizeWs(person?.[emailField]);
    return (email || name) ? `${name}::${email}` : '';
}

export function formatMonthYear(value) {
    const d = new Date(value);
    if (isNaN(d)) return value;
    return d.toLocaleDateString("en-US", { month: "short", year: "numeric" });
}

export function clearSearchDimming() {
    d3.selectAll('.dimmed').classed('dimmed', false);
    d3.selectAll('.highlighted').classed('highlighted', false);
    searchActive = false;
}

function resolveScopeFromTarget(targetEl) {
    const el = targetEl instanceof Element ? targetEl : targetEl?.node?.();
    if (!el) return null;

    const cardG = el.closest('g[data-key^="card::"]');
    if (cardG) {
        const key = cardG.getAttribute('data-key');
        const parts = key.split('::');
        const s = parts[1], t = parts[2], team = parts[3];
        return {
            mode: 'member',
            streamKey: `stream::${s}`,
            themeKey:  `theme::${s}::${t}`,
            teamKey:   `team::${s}::${t}::${team}`,
            cardKey:   key
        };
    }

    const teamG = el.closest('g[data-key^="team::"]');
    if (teamG) {
        const key = teamG.getAttribute('data-key');
        const parts = key.split('::');
        const s = parts[1], t = parts[2];
        return { mode: 'team', streamKey: `stream::${s}`, themeKey: `theme::${s}::${t}`, teamKey: key };
    }

    const themeG = el.closest('g[data-key^="theme::"]');
    if (themeG) {
        const key = themeG.getAttribute('data-key');
        const parts = key.split('::');
        const s = parts[1];
        return { mode: 'theme', streamKey: `stream::${s}`, themeKey: key };
    }

    const streamG = el.closest('g[data-key^="stream::"]');
    if (streamG) {
        const key = streamG.getAttribute('data-key');
        return { mode: 'stream', streamKey: key };
    }
    return null;
}



export function applySearchDimmingForMatches(matchElements) {
    clearSearchDimming();
    if (!matchElements || matchElements.length === 0) return;

    searchActive = true;

    const hit = {
        streams: new Set(), // stream::<s>
        themes : new Set(), // theme::<s>::<t>
        teams  : new Set(), // team::<s>::<t>::<team>
        cards  : new Set()  // card::<s>::<t>::<team>::<member>
    };

    const scopes = matchElements
        .map(el => resolveScopeFromTarget(el))
        .filter(Boolean);

    scopes.forEach(s => {
        switch (s.mode) {
            case 'stream': hit.streams.add(s.streamKey); break;
            case 'theme' : hit.themes.add(s.themeKey);   break;
            case 'team'  : hit.teams.add(s.teamKey);     break;
            case 'member': hit.cards.add(s.cardKey);     break;
        }
    });

    d3.selectAll('#streamLayer > g, #themeLayer > g, #teamLayer > g, #cardLayer > g')
        .classed('dimmed', true)
        .classed('highlighted', false);

    const undimByKey = (key) => d3.select(`g[data-key="${key}"]`).classed('dimmed', false);
    const markByKey  = (key) => d3.select(`g[data-key="${key}"]`).classed('highlighted', true);
    const undimSel   = (sel) => sel.classed('dimmed', false);

    hit.streams.forEach(streamKey => {
        const s = streamKey.split('::')[1];

        undimByKey(streamKey); markByKey(streamKey);
        undimSel(d3.selectAll(`g[data-key^="theme::${s}::"]`));
        undimSel(d3.selectAll(`g[data-key^="team::${s}::"]`));
        undimSel(d3.selectAll(`g[data-key^="card::${s}::"]`));
    });

    hit.themes.forEach(themeKey => {
        const parts  = themeKey.split('::');           // theme::s::t
        const stream = `stream::${parts[1]}`;
        const suffix = parts.slice(1).join('::');      // s::t

        undimByKey(stream); markByKey(stream);

        undimByKey(themeKey); markByKey(themeKey);

        undimSel(d3.selectAll(`g[data-key^="team::${suffix}::"]`));
        undimSel(d3.selectAll(`g[data-key^="card::${suffix}::"]`));
    });

    hit.teams.forEach(teamKey => {
        const parts   = teamKey.split('::');              // team::s::t::team
        const stream  = `stream::${parts[1]}`;
        const theme   = `theme::${parts[1]}::${parts[2]}`;
        const suffix  = parts.slice(1).join('::');        // s::t::team

        undimByKey(stream); markByKey(stream);
        undimByKey(theme);  markByKey(theme);

        undimByKey(teamKey); markByKey(teamKey);

        undimSel(d3.selectAll(`g[data-key^="card::${suffix}::"]`));
    });

    hit.cards.forEach(cardKey => {
        const parts   = cardKey.split('::');              // card::s::t::team::member
        const stream  = `stream::${parts[1]}`;
        const theme   = `theme::${parts[1]}::${parts[2]}`;
        const team    = `team::${parts[1]}::${parts[2]}::${parts[3]}`;

        undimByKey(stream); markByKey(stream);
        undimByKey(theme);  markByKey(theme);
        undimByKey(team);   markByKey(team);

        undimByKey(cardKey); markByKey(cardKey);
    });
}

export function truncateString(str, maxLength = 25) {
    if (str.length <= maxLength) return str;
    return str.slice(0, maxLength) + '...';
}

export function addTagToElement(element, number, tag = 'br') {
    element.insertAdjacentHTML('beforeend', `<${tag}>`.repeat(number));
}

export function buildFallbackMailToLink(peopleDBUpdateRecipients, subjectParam, bodyParam) {
    window.location.href = `mailto:${peopleDBUpdateRecipients.join(",")}?subject=${encodeURIComponent(subjectParam)}&body=${encodeURIComponent(bodyParam)}`;
}

export function createHrefElement(cleanUrl, textContent) {
    const a = document.createElement('a');
    a.href = cleanUrl;
    a.textContent = textContent ?? "🔗External Link";
    a.target = '_blank';
    a.rel = 'noopener noreferrer';
    a.style.color = '#0078d4';
    a.style.textDecoration = 'underline';
    return a;
}

function textNodeWithLinksToNodes(text) {
    const nodes = [];
    const urlRe = /(https?:\/\/[^\s<>"')\]]+)/g;

    let lastIndex = 0;
    let match;
    while ((match = urlRe.exec(text)) !== null) {
        const before = text.slice(lastIndex, match.index);
        if (before) nodes.push(document.createTextNode(before));

        let url = match[1];
        const trailingPunct = /[.,;:!?)+\]]+$/;
        let punct = '';
        const m2 = url.match(trailingPunct);
        if (m2) {
            punct = m2[0];
            url = url.slice(0, -punct.length);
        }

        nodes.push(createHrefElement(url));
        if (punct) nodes.push(document.createTextNode(punct));
        lastIndex = urlRe.lastIndex;
    }

    const rest = text.slice(lastIndex);
    if (rest) nodes.push(document.createTextNode(rest));
    return nodes;
}

const allowedAttributesByTag = {
    'a': new Set(['href', 'title', 'target', 'rel']),
};

function sanitizeUrl(url) {
    if (typeof url !== 'string') return '';
    const trimmed = url.trim();
    const lower = trimmed.toLowerCase();

    const forbiddenSchemes = ['javascript:', 'vbscript:'];
    if (forbiddenSchemes.some(s => lower.startsWith(s))) {
        return '';
    }

    try {
        const u = new URL(trimmed, window.location.origin);
        const allowed = ['http:', 'https:', 'mailto:', 'tel:', 'ftp:'];
        if (!allowed.includes(u.protocol) && !trimmed.startsWith('/')) {
            return '';
        }
    } catch (_) {
    }

    return trimmed;
}

function copyAllowedAttributes(srcElem, dstElem, allowedAttributesByTag) {
    const tag = srcElem.tagName.toLowerCase();
    const allowedAttrs = allowedAttributesByTag[tag];
    if (!allowedAttrs) return;

    for (const attr of srcElem.attributes) {
        const name = attr.name.toLowerCase();
        if (!allowedAttrs.has(name)) continue;

        let value = attr.value;

        if (tag === 'a') {
            if (name === 'href') {
                value = sanitizeUrl(value);
                if (!value) continue;
            }
            if (name === 'target') {
                const allowedTargets = new Set(['_blank', '_self']);
                if (!allowedTargets.has(value)) value = '_blank';
            }
            if (name === 'rel') {
                const parts = new Set(
                    value.split(/\s+/).filter(Boolean).map(v => v.toLowerCase())
                );
                parts.add('noopener');
                parts.add('noreferrer');
                value = Array.from(parts).join(' ');
            }
        }

        dstElem.setAttribute(name, value);
    }

    if (tag === 'a' && dstElem.hasAttribute('href')) {
        if (!dstElem.hasAttribute('rel')) {
            dstElem.setAttribute('rel', 'noopener noreferrer');
        }
        if (!dstElem.hasAttribute('target')) {
            dstElem.setAttribute('target', '_blank');
        }
    }
}


function sanitizeAndTransformNode(node, allowedTags) {
    if (node.nodeType === Node.TEXT_NODE) {
        return textNodeWithLinksToNodes(node.nodeValue || '');
    }

    if (node.nodeType === Node.ELEMENT_NODE) {
        const normalizedTag = node.tagName.toLowerCase();

        if (allowedTags.has(normalizedTag)) {
            const clone = document.createElement(normalizedTag);

            copyAllowedAttributes(node, clone, allowedAttributesByTag);

            node.childNodes.forEach(child => {
                const childParts = sanitizeAndTransformNode(child, allowedTags);
                childParts.forEach(p => clone.appendChild(p));
            });

            if (normalizedTag === 'a' && !clone.getAttribute('href')) {
                const fragmentNodes = [];
                clone.childNodes.forEach(c => fragmentNodes.push(c));
                return fragmentNodes;
            }
            return [clone];
        }

        const fragmentNodes = [];
        node.childNodes.forEach(child => {
            const childParts = sanitizeAndTransformNode(child, allowedTags);
            childParts.forEach(p => fragmentNodes.push(p));
        });
        return fragmentNodes;
    }

    return [];
}

export function createFormattedElementsFrom(lines) {
    const elementsToAppend = [];
    const allowedTags = new Set(['b', 'i', 'ul', 'li', 'a']);

    lines.forEach((line, index) => {

        const template = document.createElement('template');
        template.innerHTML = line;

        Array.from(template.content.childNodes).forEach(node => {
            const parts = sanitizeAndTransformNode(node, allowedTags);
            parts.forEach(p => elementsToAppend.push(p));
        });

        if (index < lines.length - 1) {
            elementsToAppend.push(document.createElement('br'));
        }
    });
    return elementsToAppend;
}

export function createFormattedLongTextElementsFrom(longText) {
    let elementsToAppend = [];
    if (longText) {
        const lines = longText.split('\n');
        elementsToAppend = createFormattedElementsFrom(lines, elementsToAppend);
    }
    return elementsToAppend;
}

function computeThemeWidth(numTeams, thirdLevelBoxWidth, thirdLevelBoxPadX) {
    const n = Number(numTeams) || 0;
    if (n <= 0) {
        return SECOND_LEVEL_LABEL_EXTRA;
    }
    return n * thirdLevelBoxWidth + (n - 1) * thirdLevelBoxPadX + SECOND_LEVEL_LABEL_EXTRA;
}


export function computeStreamBoxWidthWrapped(
    secondLevelItems,
    secondLevelBoxPadX,
    secondLevelNA,
    thirdLevelBoxPadX,
    thirdLevelBoxWidth,
    themesPerRow = 4,
    minWidth = 600,
    firstLevelPad = 80
) {

    const themeEntries = Object.entries(secondLevelItems)
        .filter(([themeKey]) => !themeKey.includes(secondLevelNA));

    const teamsPerThemeInStream = themeEntries.map(([, thirdLevelItems]) =>
        Object.keys(thirdLevelItems).length
    );

    const themeWidths = teamsPerThemeInStream.map(n =>
        computeThemeWidth(n, thirdLevelBoxWidth, thirdLevelBoxPadX)
    );

    if (!themeWidths || themeWidths.length === 0) return minWidth;

    let maxRowWidth = 0;
    for (let i = 0; i < themeWidths.length; i += themesPerRow) {
        const row = themeWidths.slice(i, i + themesPerRow);
        const rowSum = row.reduce((acc, w) => acc + (Number(w) || 0), 0);
        const pads = (row.length - 1) * secondLevelBoxPadX;
        const rowWidth = rowSum + pads + firstLevelPad;

        if (rowWidth > maxRowWidth) maxRowWidth = rowWidth;
    }
    return Math.max(maxRowWidth, minWidth);
}

export function updateLegend(scale, field, d3param) {
    const legend = d3param.select('#legend');
    legend.html('');

    legend.append('div')
        .attr('class', 'legend-title')
        .text(`${field} Legenda`);

    const itemsWrap = legend.append('div').attr('class', 'legend-items');

    const domain = scale.domain();

    domain.forEach(label => {
        const key = label || 'Unknown';
        const row = itemsWrap.append('div').attr('class', 'legend-item');

        row.append('span')
            .attr('class', 'legend-swatch')
            .style('background', scale(key));

        row.append('span')
            .attr('class', 'legend-label')
            .text(`${key}`);
    });
}

export function buildLegendaColorScale(field, items, d3param, palette, neutralColor, specialMappedField, guestValues) {
    if (specialMappedField === undefined || field !== specialMappedField) {
        const domainArr = Array.from(new Set(
            items.map(m => (m?.[field] ?? '').toString().trim() || 'Unknown')
        ));
        return d3param.scaleOrdinal(domainArr, palette);
    }


    const foundGuests = new Set();
    for (const m of items) {
        const raw = (m?.[specialMappedField] ?? '').toString();
        const rawLower = raw.toLowerCase();
        guestValues.forEach(gv => {
            if (gv && rawLower.includes(gv.toLowerCase())) {
                foundGuests.add(gv);
            }
        });
    }

    const domainWithOther = [...foundGuests, TEAM_MEMBER_LEGENDA_LABEL];
    const paletteForSpecialEntries = domainWithOther.map((_, i) =>
        i < foundGuests.size ? palette[i % palette.length] : neutralColor
    );


    const scale = d3param.scaleOrdinal(domainWithOther, paletteForSpecialEntries);

    scale.isGuest = (specificField) => {
        const val = (specificField || '').toString().toLowerCase();
        return guestValues.some(gv => val.includes(gv.toLowerCase()));
    };
    return scale;
}

export function createOutlookUrl(to, cc = [], subject = '', body = '') {
    const toParam = to.length > 1 ? encodeURIComponent(to.join(';')) : '';
    const ccParam = cc.length > 1 ? encodeURIComponent(cc.join(';')) : '';

    const subjectParam = encodeURIComponent(subject);
    const bodyParam = encodeURIComponent(body);

    let url = `https://outlook.office.com/mail/deeplink/compose?subject=${subjectParam}&body=${bodyParam}`;
    if (toParam) url += `&to=${toParam}`;
    if (ccParam) url += `&cc=${ccParam}`;
    return url;
}

export function openOutlookWebCompose({to = [], cc = [], bcc = [], subject = '', body = ''}) {
    window.open(createOutlookUrl(to, cc, subject, body), '_blank', 'noopener');
}

export function isMobileDevice() {
    try {
        if (navigator.userAgentData && typeof navigator.userAgentData.mobile === 'boolean') {
            return navigator.userAgentData.mobile;
        }
    } catch (_) {
    }

    const ua = (navigator.userAgent || navigator.vendor || window.opera || '').toLowerCase();
    const uaIsMobile =
        /android|iphone|ipod|ipad|iemobile|mobile|blackberry|opera mini|opera mobi|silk/.test(ua) ||
        ((/macintosh/.test(ua) || /mac os x/.test(ua)) && 'ontouchend' in document);

    const smallViewport = Math.min(window.screen.width, window.screen.height) <= 820; // tablet/phone

    return uaIsMobile || smallViewport;
}


export function getQueryParam(param) {
    const urlParams = new URLSearchParams(window.location.search);
    return urlParams.get(param);
}

export function setQueryParam(param, value) {
    const url = new URL(window.location);
    if (value === undefined || value === null) return;
    url.searchParams.set(param, value);
    window.history.pushState({}, '', url);
}

export function setSearchQuery(search) {
    setQueryParam('search', search);
}

export function initCommonActions() {
    const overlay = document.getElementById('side-overlay');
    const closeBtn = document.getElementById('side-close');

    overlay?.addEventListener('click', closeSideDrawer);
    closeBtn?.addEventListener('click', closeSideDrawer);

    window.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') closeSideDrawer();
    });

    const toggleCta = document.getElementById('toggle-cta');
    toggleCta?.addEventListener('click', (e) => {
        e.preventDefault();
        openSideDrawer();
    });

    document.getElementById('act-upload')?.addEventListener('click', () => {
        document.getElementById('fileInput')?.click();
        closeSideDrawer();
    });
}

export function openSideDrawer() {
    const drawer = document.getElementById('side-drawer');
    const overlay = document.getElementById('side-overlay');
    if (!drawer) return;

    drawer.classList.add('open');
    overlay?.classList.add('visible');
    document.body.classList.add('side-drawer-open');
    drawer.setAttribute('aria-hidden', 'false');

    document.getElementById('act-upload')?.focus();
}

export function closeSideDrawer() {
    const drawer = document.getElementById('side-drawer');
    const overlay = document.getElementById('side-overlay');
    if (!drawer) return;
    drawer.classList.remove('open');
    overlay?.classList.remove('visible');
    document.body.classList.remove('side-drawer-open');
    drawer.setAttribute('aria-hidden', 'true');
}

export function toggleClearButton(buttonId, value) {
    const el = document.getElementById(buttonId);
    if (!el) return;
    el.classList.toggle('hidden', !value);
}

export function getFormattedDate(isoDate, locale = 'it-IT', timeZone = 'Europe/Rome') {
    const date = new Date(isoDate);
    return date.toLocaleString(locale, {
        timeZone,
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
    });
}

export function parseCSV(text) {
    const rows = [];
    let current = [];
    let inQuotes = false;
    let value = '';
    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        if (char === '"') {
            if (inQuotes && text[i + 1] === '"') {
                value += '"';
                i++;
            } else {
                inQuotes = !inQuotes;
            }
        } else if (char === ',' && !inQuotes) {
            current.push(value);
            value = '';
        } else if ((char === '\n' || char === '\r') && !inQuotes) {
            if (value || current.length > 0) {
                current.push(value);
                rows.push(current);
                current = [];
                value = '';
            }
            if (char === '\r' && text[i + 1] === '\n') i++;
        } else {
            value += char;
        }
    }
    if (value || current.length > 0) {
        current.push(value);
        rows.push(current);
    }
    return rows;
}

export function clearHighlights(viewport) {
    viewport.selectAll('rect').attr('stroke', null).attr('stroke-width', null);
}

export function highlightGroup(groupSel) {
    clearHighlights(groupSel);
    const rect = groupSel.select('rect.profile-box').node()
        ? groupSel.select('rect.profile-box')
        : groupSel.select('rect');
    if (rect.node()) rect.attr('stroke', '#ff9900').attr('stroke-width', 3);
}
