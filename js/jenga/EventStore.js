/**
 * EventStore — loads and queries Jenga calendar events from CSV.
 *
 * CSV columns:
 *   Key,Summary,Type,Environment,Service,Operation,DueDate,StartDate,
 *   Status,Priority,Labels,Assignee,Stream,Theme,JiraUrl
 */

import { parseSectionMeta } from '../shared/utils.js';

// ─── Time slot parser ─────────────────────────────────────────────────────────

const RE_TIME_SLOT = /^(\d{1,2}:\d{2}\s*[AP]M)\s+to\s+(\d{1,2}:\d{2}\s*[AP]M)$/i;

/** Convert "H:MM AM/PM" → minutes since midnight. 12:xx AM = 0, 12:xx PM = 720. */
function timeToMinutes(timeStr) {
    const m = timeStr.trim().match(/^(\d{1,2}):(\d{2})\s*([AP]M)$/i);
    if (!m) return null;
    let h = parseInt(m[1], 10);
    const min = parseInt(m[2], 10);
    const period = m[3].toUpperCase();
    if (period === 'AM') {
        if (h === 12) h = 0;
    } else {
        if (h !== 12) h += 12;
    }
    return h * 60 + min;
}

export function parseProtectionColor(summary) {
    if (!summary) return null;
    const first = summary.split('|')[0]?.trim().toUpperCase();

    if (first === 'RED') return 'RED';
    if (first === 'AMBER') return 'AMBER';
    if (first === 'BLUE') return 'BLUE';

    return null;
}

export function parseProtectionCab(summary) {
    if (!summary) return false;
    const first = summary.split('|')[0]?.trim().toUpperCase();
    return first === 'CAB';
}

/**
 * Extract time slot from the last pipe-delimited segment of a summary string.
 * Returns { timeStart, timeEnd, startMin, endMin } or null if not found.
 */
export function parseTimeSlot(summary) {
    if (!summary) return null;
    const parts = summary.split('|');
    const last = parts[parts.length - 1].trim();
    const match = last.match(RE_TIME_SLOT);
    if (!match) return null;
    const startMin = timeToMinutes(match[1]);
    let endMin   = timeToMinutes(match[2]);
    if (startMin === null || endMin === null) return null;
    // When endMin ≤ startMin the end-time AM/PM is likely a noon-boundary typo
    // (Jira automation writes "12:14 AM" when it should be "12:14 PM").
    // Fall back to a 60-minute window so column layout and block height remain correct.
    if (endMin <= startMin) endMin = startMin + 60;
    return { timeStart: match[1].trim(), timeEnd: match[2].trim(), startMin, endMin };
}

/**
 * Extract optional context segment from SERVICE_OP summary.
 * Format: env | service | operation | context | timeslot
 * Returns the trimmed context string when parts[3] exists and parts[4] is the timeslot,
 * otherwise null.
 */
export function parseContext(summary) {
    if (!summary) return null;
    const parts = summary.split('|').map(p => p.trim());
    if (parts.length < 5) return null;
    const ctx = parts[3];
    return ctx || null;
}

// ─── Peak date parser ─────────────────────────────────────────────────────────

// Matches DD/MM/YYYY or DD-MM-YYYY (day and month 1-2 digits, year 2 or 4 digits)
const RE_DATE_PART = /(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/;
const RE_DATE_G    = /(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/g;

function parseCalDate(d, m, y) {
    const day   = parseInt(d, 10);
    const month = parseInt(m, 10) - 1; // 0-indexed
    const year  = parseInt(y, 10) < 100 ? 2000 + parseInt(y, 10) : parseInt(y, 10);
    const date  = new Date(year, month, day);
    return isNaN(date.getTime()) ? null : date;
}

/**
 * Extract peak date ranges from a free-text description.
 * Returns an array of { start: Date, end: Date } (end === start for single dates).
 * Returns [] if nothing found or input is empty/null.
 */
export function parsePeakDates(description) {
    if (!description) return [];

    // Split on newlines and sentence-ending periods so multi-sentence single-line
    // descriptions are each evaluated for peak/spike keywords independently.
    const sentences = description
        .split(/\n|\.\s+/)
        .filter(s => /\b(peak|spike)\b/i.test(s));

    if (sentences.length === 0) return [];

    const results = [];

    for (const sentence of sentences) {
        // "from DATE to DATE"
        const fromTo = sentence.match(/from\s+(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\s+to\s+(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/i);
        if (fromTo) {
            const start = parseCalDate(fromTo[1], fromTo[2], fromTo[3]);
            const end   = parseCalDate(fromTo[4], fromTo[5], fromTo[6]);
            if (start && end) { results.push({ start, end }); continue; }
        }

        // "DATE - DATE"
        const rangeDash = sentence.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})\s+-\s+(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
        if (rangeDash) {
            const start = parseCalDate(rangeDash[1], rangeDash[2], rangeDash[3]);
            const end   = parseCalDate(rangeDash[4], rangeDash[5], rangeDash[6]);
            if (start && end) { results.push({ start, end }); continue; }
        }

        // "on DATE"
        const onDate = sentence.match(/\bon\s+(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/i);
        if (onDate) {
            const d = parseCalDate(onDate[1], onDate[2], onDate[3]);
            if (d) { results.push({ start: d, end: d }); continue; }
        }

        // Fallback: any date in this sentence
        const cleaned = sentence.replace(/\d{1,2}:\d{2}/g, '');
        let m;
        RE_DATE_G.lastIndex = 0;
        while ((m = RE_DATE_G.exec(cleaned)) !== null) {
            const d = parseCalDate(m[1], m[2], m[3]);
            if (d) results.push({ start: d, end: d });
        }
    }

    return results;
}

// ─── Classification helpers (also exported for unit testing) ──────────────────

const RE_SERVICE_OP = /^(prd|prod|qa\d*|qaesales|staging|devops|qasales)\s*\|/i;
const RE_HYBRIS     = /^release\//i;
const RE_BUSINESS   = /^Event:/i;

export function classifyEvent(summary, labels = []) {
    const s = (summary || '').trim();
    const lbls = new Set((labels || []).map(l => l.toLowerCase()));

    // Label-based signals take priority over summary patterns
    if (lbls.has('hybrisreleasenote')) {
        return { type: 'HYBRIS', environment: '', service: s, operation: '' };
    }
    if (lbls.has('event_launch')) {
        return { type: 'BUSINESS_EVENT', environment: '', service: '', operation: '' };
    }

    if (RE_SERVICE_OP.test(s)) {
        const parts = s.split('|').map(p => p.trim());
        return {
            type:        'SERVICE_OP',
            environment: parts[0] || '',
            service:     parts[1] || '',
            operation:   parts[2] || '',
        };
    }
    if (RE_HYBRIS.test(s)) {
        return { type: 'HYBRIS', environment: '', service: s, operation: '' };
    }
    if (RE_BUSINESS.test(s)) {
        return { type: 'BUSINESS_EVENT', environment: '', service: '', operation: '' };
    }
    return { type: 'DELIVERY', environment: '', service: '', operation: '' };
}

// ─── CSV parser ───────────────────────────────────────────────────────────────

function parseCSVRow(line) {
    return parseCSVRecords(line)[0] || [];
}

function parseCSVRecords(csvText) {
    const rows = [];
    let row = [];
    let field = '';
    let inQuotes = false;
    const text = String(csvText || '').replace(/^\uFEFF/, '');

    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        const next = text[i + 1];
        if (ch === '"') {
            if (inQuotes && next === '"') { field += '"'; i += 1; }
            else { inQuotes = !inQuotes; }
            continue;
        }
        if (ch === ',' && !inQuotes) { row.push(field); field = ''; continue; }
        if ((ch === '\n' || ch === '\r') && !inQuotes) {
            if (ch === '\r' && next === '\n') i += 1;
            row.push(field); field = '';
            if (row.some(v => String(v).trim() !== '')) rows.push(row);
            row = [];
            continue;
        }
        field += ch;
    }
    row.push(field);
    if (row.some(v => String(v).trim() !== '')) rows.push(row);
    return rows;
}

function normalizeJiraUrl(url, key) {
    const cleanUrl = String(url || '').trim();
    if (/^https?:\/\//i.test(cleanUrl)) return cleanUrl;
    const cleanKey = String(key || '').trim();
    if (/^[A-Z][A-Z0-9]+-\d+$/.test(cleanKey)) {
        return `https://brand.atlassian.net/servicedesk/customer/portal/1/${cleanKey}`;
    }
    return '';
}

/** Format a local Date as YYYY-MM-DD, consistent with the CSV values. */
function localDateKey(d) {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function parseDate(s) {
    if (!s || !s.trim()) return null;
    // Parse YYYY-MM-DD as local midnight so it aligns with the local-midnight
    // boundary comparisons used throughout (new Date(year, month, day)).
    // new Date('YYYY-MM-DD') without a time component is parsed as UTC midnight
    // by the JS spec, which shifts dates by ±1 day in non-UTC timezones.
    const iso = s.trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (iso) return new Date(+iso[1], +iso[2] - 1, +iso[3]);
    const d = new Date(s.trim());
    return isNaN(d.getTime()) ? null : d;
}

// ─── EventStore ───────────────────────────────────────────────────────────────

export class EventStore {
    constructor() {
        this.events = [];
        this.cards  = [];
        this.rawCsv = '';
    }

    load(csvText) {
        this.rawCsv = csvText || '';
        this.events = [];
        if (!csvText || !csvText.trim()) return;

        const records = parseCSVRecords(csvText);
        if (records.length < 2) return;

        const headers = records[0].map(h => String(h).trim());
        const idx = (name) => headers.indexOf(name);

        for (let i = 1; i < records.length; i++) {
            const row = records[i];
            if (!row || !row.some(v => String(v).trim())) continue;
            const get = (col) => {
                const colIdx = idx(col);
                return colIdx >= 0 ? String(row[colIdx] || '').trim() : '';
            };

            const key       = get('Key');
            if (!key) continue;

            const { clean: summary, roi, topPriority: metaPriority } = parseSectionMeta(get('Summary'));
            const dueDate   = parseDate(get('DueDate'));
            const startDate = parseDate(get('StartDate'));

            // Use pre-classified Type from CSV if present, else re-classify
            const csvType = get('Type');
            const parsed  = csvType ? {
                type:        csvType,
                environment: get('Environment'),
                service:     get('Service'),
                operation:   get('Operation'),
            } : classifyEvent(summary);

            this.events.push({
                key,
                summary,
                roi,
                metaPriority,
                protectionColor: parseProtectionColor(summary),
                isCab:         parseProtectionCab(summary),
                type:          parsed.type,
                environment:   parsed.environment,
                service:       parsed.service,
                operation:     parsed.operation,
                context:       parsed.type === 'SERVICE_OP' ? parseContext(summary) : null,
                dueDate,
                startDate,
                status:        get('Status'),
                priority:      get('Priority'),
                labels:        get('Labels'),
                assignee:      get('Assignee'),
                reporter:      get('Reporter'),
                description:   get('Description'),
                latestComment: get('LatestComment'),
                stream:        get('Stream'),
                theme:         get('Theme'),
                jiraUrl:       normalizeJiraUrl(get('JiraUrl'), key),
                timeSlot:      parseTimeSlot(summary),
                peakDates:     parsePeakDates(get('Description')),
            });
        }
    }

    /** Events that occur on the given day (local date).
     *  Genuinely multi-day events (startDate strictly before dueDate) use span overlap.
     *  Point-in-time events (no startDate, or startDate === dueDate) match exactly on dueDate. */
    getEventsForDay(date) {
        if (!date) return [];
        const y = date.getFullYear(), m = date.getMonth(), d = date.getDate();
        const dayStart = new Date(y, m, d);
        const dayEnd   = new Date(y, m, d, 23, 59, 59);
        return this.events.filter(e => {
            const due   = e.dueDate;
            if (!due) return false;
            const start = e.startDate;
            // Multi-day span: startDate strictly before dueDate, and only for
            // event types that genuinely span days (BUSINESS_EVENT, DELIVERY).
            // MILESTONE and SERVICE_OP always use exact date match on dueDate.

            const canSpan =
                e.type === 'BUSINESS_EVENT' ||
                e.type === 'DELIVERY' ||
                e.type === 'PEAK_SEASON_PROTECTION_WINDOW';

            if (canSpan && start && due) {
                let evStart = start;
                let evEnd = due;

                if (evStart > evEnd) {
                    const tmp = evStart;
                    evStart = evEnd;
                    evEnd = tmp;
                }

                if (evStart < evEnd) {
                    return evStart <= dayEnd && evEnd >= dayStart;
                }
            }
            // Point-in-time: exact date match on dueDate
            return due.getFullYear() === y && due.getMonth() === m && due.getDate() === d;
        });
    }

    /** Events visible in the given year/month (0-indexed month).
     *  An event is included if its dueDate is in the month OR (for multi-day
     *  events) its span overlaps with the month at all. */
    getEventsForMonth(year, month) {
        const monthStart = new Date(year, month, 1);
        const monthEnd   = new Date(year, month + 1, 0, 23, 59, 59);
        return this.events.filter(e => {
            const due   = e.dueDate;
            const start = e.startDate || due;
            if (!due && !start) return false;
            // Event overlaps month: normalize start/end because some CSV rows may
            // contain StartDate after DueDate.
            const rawStart = start || due;
            const rawEnd   = due   || start;
            const evStart  = rawStart <= rawEnd ? rawStart : rawEnd;
            const evEnd    = rawStart <= rawEnd ? rawEnd   : rawStart;
            return evStart <= monthEnd && evEnd >= monthStart;
        });
    }

    /** Load ITSM cards CSV (incidents + service requests from jira-cards.csv). */
    loadCards(csvText) {
        this.cards = [];
        this.cardsGeneratedAt = null;
        if (!csvText || !csvText.trim()) return;
        // Strip optional leading comment line "# generated: TIMESTAMP"
        let body = csvText;
        const firstLine = csvText.slice(0, csvText.indexOf('\n'));
        const genMatch = firstLine.match(/^#\s*generated:\s*(\S+)/i);
        if (genMatch) {
            this.cardsGeneratedAt = genMatch[1];
            body = csvText.slice(csvText.indexOf('\n') + 1);
        }
        const records = parseCSVRecords(body);
        if (records.length < 2) return;
        const headers = records[0].map(h => String(h).trim());
        const idx = (name) => headers.indexOf(name);
        for (let i = 1; i < records.length; i++) {
            const row = records[i];
            if (!row || !row.some(v => String(v).trim())) continue;
            const get = (col) => {
                const colIdx = idx(col);
                return colIdx >= 0 ? String(row[colIdx] || '').trim() : '';
            };
            const key = get('Key');
            if (!key) continue;
            const created = parseDate(get('Created'));
            if (!created) continue;
            const { clean: summary, roi, topPriority: metaPriority } = parseSectionMeta(get('Summary'));
            this.cards.push({
                key,
                issueType:        get('IssueType'),
                requestType:      get('RequestType'),
                created,
                summary,
                roi,
                metaPriority,
                affectedServices: get('AffectedServices'),
                status:           get('Status'),
                priority:         get('Priority'),
                reporter:         get('Reporter'),
                jiraUrl:          normalizeJiraUrl(get('JiraUrl'), key),
            });
        }
    }

    /** Count cards per day for a given month, optionally filtered by issueType and service query.
     *  Returns Map<dateStr YYYY-MM-DD, count>. */
    getCardCountsForMonth(year, month, { issueType, serviceQuery, services } = {}) {
        const monthStart = new Date(year, month, 1);
        const monthEnd   = new Date(year, month + 1, 0, 23, 59, 59);
        const counts = new Map();
        (this.cards || []).forEach(c => {
            if (!c.created) return;
            if (c.created < monthStart || c.created > monthEnd) return;
            if (issueType && !c.issueType.toLowerCase().includes(issueType.toLowerCase())) return;
            if (serviceQuery) {
                const q = serviceQuery.toLowerCase();
                const hay = `${c.summary} ${c.affectedServices} ${c.requestType}`.toLowerCase();
                if (!hay.includes(q)) return;
            }
            // Service Set filter: card passes if any of its affectedServices is in the set
            if (services && services.size > 0) {
                const cardServices = (c.affectedServices || '')
                    .split('||')
                    .map(s => s.trim())
                    .filter(Boolean);
                const matches = cardServices.some(s => services.has(s));
                if (!matches) return;
            }
            const k = localDateKey(c.created);
            counts.set(k, (counts.get(k) || 0) + 1);
        });
        return counts;
    }

    /** Return the Set of date keys (YYYY-MM-DD) that have at least one P1 or P2 card. */
    getCriticalDaysForMonth(year, month, { issueType, services } = {}) {
        const monthStart = new Date(year, month, 1);
        const monthEnd   = new Date(year, month + 1, 0, 23, 59, 59);
        const critical = new Set();
        (this.cards || []).forEach(c => {
            if (!c.created || c.created < monthStart || c.created > monthEnd) return;
            if (issueType && !(c.issueType || '').toLowerCase().includes(issueType.toLowerCase())) return;
            if (services && services.size > 0) {
                const cardSvcs = (c.affectedServices || '').split('||').map(s => s.trim()).filter(Boolean);
                if (!cardSvcs.some(s => services.has(s))) return;
            }
            const p = (c.priority || '').trim().toUpperCase();
            if (p.startsWith('P1') || p.startsWith('P2')) critical.add(localDateKey(c.created));
        });
        return critical;
    }

    /** Return cards created on a specific day, filtered by issueType label fragment
     *  and optionally by a Set of service names (matched against affectedServices).
     *
     *  Uses the same local date key as getCardCountsForMonth so the drawer
     *  shows exactly the cards that were counted for that dot. */
    getCardsForDay(date, issueTypeLabelFragment, services) {
        if (!date) return [];
        const dayKey = localDateKey(date);
        return (this.cards || []).filter(c => {
            if (!c.created) return false;
            if (localDateKey(c.created) !== dayKey) return false;
            if (issueTypeLabelFragment) {
                const frag = issueTypeLabelFragment.toLowerCase();
                if (!c.issueType.toLowerCase().includes(frag)) return false;
            }
            if (services && services.size > 0) {
                const cardServices = (c.affectedServices || '')
                    .split('||').map(s => s.trim()).filter(Boolean);
                if (!cardServices.some(s => services.has(s))) return false;
            }
            return true;
        });
    }

    /**
     * Filter events by optional criteria.
     * @param {Object} opts
     * @param {Set<string>} [opts.types]    — allowed event types
     * @param {Set<string>} [opts.envs]     — allowed environments
     * @param {string}      [opts.query]    — free-text search on summary/service
     */
    filter({ types, envs, query } = {}) {
        return this.events.filter(e => {
            if (types && types.size > 0 && !types.has(e.type)) return false;
            if (envs  && envs.size  > 0 && e.environment &&
                !envs.has(e.environment.toLowerCase())) return false;
            if (query) {
                const q = query.toLowerCase();
                const hay = `${e.summary} ${e.service} ${e.environment} ${e.operation}`.toLowerCase();
                if (!hay.includes(q)) return false;
            }
            return true;
        });
    }

    /** Return all open P1/P2 incidents (status != 'Closed', issueType contains 'incident'). */
    getMajorIncidents() {
        return (this.cards || []).filter(c =>
            c.status !== 'Closed' &&
            c.issueType.toLowerCase().includes('incident') &&
            (c.priority.includes('P1') || c.priority.includes('P2'))
        );
    }

    /**
     * Compute a per-day-of-month average curve over the N months preceding (year, month).
     * For each day position 1–daysInMonth, the value is the mean daily count across
     * all sampled months that had that day (shorter months skip day 29/30/31).
     * Returns [{date: Date(year, month, d), value: number}] aligned to the target month.
     */
    getAverageCurveForMonth(issueType, year, month, lookbackMonths, opts = {}) {
        const buckets = new Map(); // dayOfMonth → { sum, count }
        for (let i = 1; i <= lookbackMonths; i++) {
            let y = year, m = month - i;
            while (m < 0) { m += 12; y--; }
            const counts = this.getCardCountsForMonth(y, m, { issueType, services: opts.services });
            const daysInM = new Date(y, m + 1, 0).getDate();
            for (let d = 1; d <= daysInM; d++) {
                const key = `${y}-${String(m + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
                const val = counts.get(key) || 0;
                const b = buckets.get(d) || { sum: 0, count: 0 };
                b.sum += val;
                b.count++;
                buckets.set(d, b);
            }
        }
        const daysInMonth = new Date(year, month + 1, 0).getDate();
        const result = [];
        for (let d = 1; d <= daysInMonth; d++) {
            const b = buckets.get(d);
            result.push({
                date:  new Date(year, month, d),
                value: b && b.count > 0 ? b.sum / b.count : 0,
            });
        }
        return result;
    }

    /**
     * Get daily card counts for a source month, re-mapped to day-of-month positions
     * in the target month so they can be overlaid on the same D3 x-axis.
     * Returns [{date: Date(targetYear, targetMonth, d), value: number}] for all days
     * in the source month (zeros included, so the line spans the full month).
     */
    getComparisonMonthData(issueType, srcYear, srcMonth, targetYear, targetMonth, opts = {}) {
        const counts       = this.getCardCountsForMonth(srcYear, srcMonth, { issueType, services: opts.services });
        const criticalDays = this.getCriticalDaysForMonth(srcYear, srcMonth, { issueType, services: opts.services });
        const daysInSrc = new Date(srcYear, srcMonth + 1, 0).getDate();
        const result = [];
        for (let d = 1; d <= daysInSrc; d++) {
            const key = `${srcYear}-${String(srcMonth + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
            result.push({
                date:        new Date(targetYear, targetMonth, d),
                value:       counts.get(key) || 0,
                hasCritical: criticalDays.has(key),
            });
        }
        return result;
    }
}
