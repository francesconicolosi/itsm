/**
 * CalendarRenderer — monthly calendar grid with event chips.
 *
 * Multi-day events (business events, deliveries) span across day columns
 * as a single continuous bar using an absolute-positioned overlay layer.
 */

const TYPE_COLORS = {
    SERVICE_OP:     { bg: '#dbeafe', border: '#3b82f6', text: '#1d4ed8' },
    DELIVERY:       { bg: '#dcfce7', border: '#22c55e', text: '#15803d' },
    BUSINESS_EVENT: { bg: '#fce7f3', border: '#db2777', text: '#9d174d' },
    HYBRIS:         { bg: '#e0fdf4', border: '#14b8a6', text: '#0f766e' },
    MILESTONE:      { bg: '#ede9fe', border: '#7c3aed', text: '#5b21b6' },
    OTHER:          { bg: '#f1f5f9', border: '#94a3b8', text: '#475569' },
};

const ENV_META = {
    prd:      { color: '#ef4444', label: 'PRD' },
    qa1:      { color: '#eab308', label: 'QA1' },
    qa2:      { color: '#eab308', label: 'QA2' },
    qa3:      { color: '#eab308', label: 'QA3' },
    qa4:      { color: '#eab308', label: 'QA4' },
    qa5:      { color: '#eab308', label: 'QA5' },
    qa:       { color: '#eab308', label: 'QA'  },
    staging:  { color: '#6b7280', label: 'STG' },
    preprod:  { color: '#6b7280', label: 'PRE' },
    devops:   { color: '#3b82f6', label: 'DEV' },
    qasales:  { color: '#f59e0b', label: 'QAS' },
    qaesales: { color: '#f59e0b', label: 'QAE' },
    shadow:   { color: '#6b7280', label: 'SHW' },
    prod:     { color: '#ef4444', label: 'PRD' },
};

// Hybris [Y] icon — inline SVG matching the SAP Hybris bracket logo
export const HYBRIS_SVG = `<svg width="13" height="13" viewBox="0 0 20 20" fill="currentColor" style="flex-shrink:0;display:inline-block;vertical-align:middle" aria-hidden="true">
  <rect x="1"   y="1"    width="2.5" height="18"/>
  <rect x="1"   y="1"    width="6"   height="2.5"/>
  <rect x="1"   y="16.5" width="6"   height="2.5"/>
  <rect x="16.5" y="1"   width="2.5" height="18"/>
  <rect x="13"  y="1"    width="6"   height="2.5"/>
  <rect x="13"  y="16.5" width="6"   height="2.5"/>
  <polygon points="6.5,1 10,8.5 13.5,1 15.5,1 11,10 11,19 9,19 9,10 4.5,1"/>
</svg>`;

const MAX_CHIPS = 5;

export function getTypeColor(type) {
    return (TYPE_COLORS[type] || TYPE_COLORS.OTHER).border;
}

export function getEnvColor(env) {
    return (ENV_META[(env || '').toLowerCase()] || { color: '#6b7280' }).color;
}

export function getEnvLabel(env) {
    return (ENV_META[(env || '').toLowerCase()] || { label: (env || '').toUpperCase().slice(0, 3) }).label;
}

function getDayProtectionColor(date, events) {
    if (!date || !events?.length) return null;

    const y = date.getFullYear();
    const m = date.getMonth();
    const d = date.getDate();

    const dayStart = new Date(y, m, d);
    const dayEnd = new Date(y, m, d, 23, 59, 59);

    const colors = events
        .filter(ev => ev.type === 'PEAK_SEASON_PROTECTION_WINDOW')
        .filter(ev => {
            let start = ev.startDate || ev.dueDate;
            let end = ev.dueDate || ev.startDate;

            if (!start || !end) return false;

            // Safety: handle inverted ranges
            if (start > end) {
                const tmp = start;
                start = end;
                end = tmp;
            }

            return start <= dayEnd && end >= dayStart;
        })
        .map(ev => ev.protectionColor)
        .filter(Boolean);

    if (!colors.length) return null;

    // Priority: RED wins over AMBER, AMBER wins over BLUE
    if (colors.includes('RED')) return 'RED';
    if (colors.includes('AMBER')) return 'AMBER';
    if (colors.includes('BLUE')) return 'BLUE';

    return null;
}


function isCabProtectionEvent(ev) {
    if (!ev || ev.type !== 'PEAK_SEASON_PROTECTION_WINDOW') return false;
    return Boolean(ev.isCab) || /^CAB\s*[|]/i.test(ev.summary || '');
}

function isEventOnDate(ev, date) {
    if (!ev || !date) return false;
    const y = date.getFullYear(), m = date.getMonth(), d = date.getDate();
    const dayStart = new Date(y, m, d);
    const dayEnd   = new Date(y, m, d, 23, 59, 59);

    let start = ev.startDate || ev.dueDate;
    let end   = ev.dueDate   || ev.startDate;
    if (!start || !end) return false;
    if (start > end) [start, end] = [end, start];
    return start <= dayEnd && end >= dayStart;
}

function getDayCabEvent(date, events) {
    if (!date || !events?.length) return null;
    return events.find(ev => isCabProtectionEvent(ev) && isEventOnDate(ev, date)) || null;
}

// ─── Grouping (single-day SERVICE_OP) ─────────────────────────────────────────

function groupSingleDayEvents(events) {
    const groups = new Map();
    const out = [];
    events.forEach(ev => {
        if (ev.type !== 'SERVICE_OP') { out.push({ grouped: false, event: ev }); return; }
        const key = `${(ev.environment||'').toLowerCase()}::${(ev.service||'').toLowerCase()}`;
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push(ev);
    });
    groups.forEach(evs => out.push({ grouped: evs.length > 1, count: evs.length, events: evs, event: evs[0] }));
    const typeOrder = { HYBRIS: 0, SERVICE_OP: 1, MILESTONE: 2, DELIVERY: 3, BUSINESS_EVENT: 4, OTHER: 5 };
    const envOrder  = { prd: 0, prod: 0, qa1: 1, qa2: 1, qa3: 1, qa4: 1, qa5: 1, qa: 1, staging: 2, devops: 3, qasales: 4, qaesales: 4, shadow: 5 };
    out.sort((a, b) => {
        const ta = typeOrder[a.event.type] ?? 5, tb = typeOrder[b.event.type] ?? 5;
        if (ta !== tb) return ta - tb;
        return (envOrder[(a.event.environment||'').toLowerCase()] ?? 9) -
            (envOrder[(b.event.environment||'').toLowerCase()] ?? 9);
    });
    return out;
}

// ─── Shared peak-date helper (also used by TimelineRenderer) ─────────────────

export function isDateInPeakRange(date, peakDates) {
    if (!date || !peakDates || !peakDates.length) return false;
    const y = date.getFullYear(), m = date.getMonth(), d = date.getDate();
    const dayStart = new Date(y, m, d);
    const dayEnd   = new Date(y, m, d, 23, 59, 59);
    return peakDates.some(({ start, end }) =>
        start <= dayEnd && end >= dayStart
    );
}

// ─── ITSM retention cutoff ────────────────────────────────────────────────────
// __ITSM_RETENTION_DAYS__ is injected at build time by webpack DefinePlugin.
const _ITSM_RETENTION_DAYS = typeof __ITSM_RETENTION_DAYS__ !== 'undefined' ? __ITSM_RETENTION_DAYS__ : 100;

function _itsmCutoffDate() {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() - _ITSM_RETENTION_DAYS);
    return d;
}

// ─── CalendarRenderer ─────────────────────────────────────────────────────────

export class CalendarRenderer {
    constructor(app) {
        this.app = app;
        this._onEventClick        = null;
        this._onDayClick          = null;
        this._popover             = null;
        this._popoverAnchor       = null;
        this._scrollHandler       = null;
        this._outsideClickHandler = null;
        this._hoverCard           = null;
        this._initBizTooltip();
    }

    onEventClick(fn) { this._onEventClick = fn; }
    onDayClick(fn)   { this._onDayClick = fn; }

    render(containerEl, year, month, events) {
        if (!containerEl) return;
        containerEl.innerHTML = '';
        this._closePopover();

        const weeks = this._buildWeeks(new Date(year, month, 1), new Date(year, month + 1, 0));

        // Separate multi-day from single-day events
        const multiDay = [];
        const singleByDay = new Map();

        events.forEach(e => {
            if (!e.dueDate) return;
            // BUSINESS_EVENT always uses the spanning bar (even single-day).
            // DELIVERY only spans when startDate < dueDate.
            const isMulti = e.type === 'BUSINESS_EVENT' ||
                (e.startDate && e.startDate < e.dueDate && e.type === 'DELIVERY');
            if (isMulti) {
                multiDay.push(e);
            } else {
                const k = this._dayKey(e.dueDate);
                if (!singleByDay.has(k)) singleByDay.set(k, []);
                singleByDay.get(k).push(e);
            }
        });

        // Pre-compute peak days for the gradient-wash cell highlight
        const peakDaySet = new Set();
        for (const ev of multiDay) {
            if (ev.type === 'BUSINESS_EVENT' && ev.peakDates && ev.peakDates.length > 0) {
                for (const week of weeks) {
                    for (const cellDate of week) {
                        if (this._isDateInPeakRange(cellDate, ev.peakDates)) {
                            peakDaySet.add(this._dayKey(cellDate));
                        }
                    }
                }
            }
        }

        const wrap = document.createElement('div');
        wrap.className = 'jenga-cal';

        // Day-of-week header
        const headerRow = document.createElement('div');
        headerRow.className = 'jenga-cal__header';
        ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'].forEach((d, i) => {
            const cell = document.createElement('div');
            cell.className = 'jenga-cal__dayname' + (i >= 5 ? ' jenga-cal__dayname--weekend' : '');
            cell.textContent = d;
            headerRow.appendChild(cell);
        });
        wrap.appendChild(headerRow);

        // Week rows — each week is a relative-positioned container for spanning bars
        weeks.forEach((week, weekIdx) => {
            const row = document.createElement('div');
            row.className = 'jenga-cal__week';
            row.dataset.weekIdx = weekIdx;

            const allCalendarEvents = this._getAllCalendarEvents(events);

            week.forEach(date => {
                row.appendChild(this._makeCell(date, month, singleByDay, allCalendarEvents, peakDaySet));
            });

            // Multi-day spanning bars for this week
            this._appendMultiDayBars(row, week, multiDay, month);

            wrap.appendChild(row);
        });

        containerEl.appendChild(wrap);
        // Use a persistent handler (not once:true) so re-opened popovers also close.
        // Capturing phase so we can consume the click before calendar cells react.
        if (!this._outsideClickHandler) {
            this._outsideClickHandler = (e) => {
                if (!this._popover) return;
                if (this._popover.contains(e.target) || this._popoverAnchor?.contains(e.target)) return;
                this._closePopover();
                if (e.target.closest('.jenga-calendar-wrap, .jenga-timeline-wrap')) {
                    e.stopPropagation();
                }
            };
            document.addEventListener('click', this._outsideClickHandler, true);
        }
    }

    // ── Single-day cell ───────────────────────────────────────────────────────


    _getAllCalendarEvents(fallbackEvents) {
        return (
            this.app?.eventStore?.events ||
            this.app?.store?.events ||
            this.app?.events ||
            fallbackEvents ||
            []
        );
    }

    _makeCell(date, month, singleByDay, allEvents, peakDaySet) {
        const cell = document.createElement('div');
        cell.className = 'jenga-cal__cell';
        cell.dataset.date = this._dayKey(date);

        const protectionColor = getDayProtectionColor(date, allEvents);

        if (protectionColor) {
            cell.classList.add(`jenga-cal__cell--pspw-${protectionColor.toLowerCase()}`);
            cell.addEventListener('mouseover', (e) => {
                if (e.target.closest('.jenga-chip, .jenga-cal__cab-badge, .jenga-multiday-bar')) return;
                this._showHoverCard(e.clientX, e.clientY, this._pspwTooltipFields(protectionColor));
            });
            cell.addEventListener('mouseout', (e) => {
                if (e.relatedTarget && cell.contains(e.relatedTarget)) return;
                this._hideHoverCard();
            });
        }

        if (peakDaySet && peakDaySet.has(this._dayKey(date))) cell.classList.add('jenga-cal__cell--peak');
        if (this._isToday(date)) cell.classList.add('jenga-cal__cell--today');
        if (date.getDay() === 0 || date.getDay() === 6) cell.classList.add('jenga-cal__cell--weekend');
        if (date.getMonth() !== month) cell.classList.add('jenga-cal__cell--other-month');

        // Days before the ITSM retention cutoff have no incident/SR data.
        // Only mark in-month cells to avoid double-dimming the out-of-month fringe.
        if (date.getMonth() === month && date < _itsmCutoffDate()) {
            cell.classList.add('jenga-cal__cell--no-itsm');
            const noItsmMsg = `No incident or service-request data — records older than ${_ITSM_RETENTION_DAYS} days are outside the retention window`;
            cell.addEventListener('mouseover', (e) => {
                if (e.target.closest('.jenga-chip, .jenga-cal__cab-badge, .jenga-multiday-bar')) return;
                if (protectionColor) return; // PSPW tooltip already shown
                this._showHoverCard(e.clientX, e.clientY, [
                    { label: 'No ITSM data', value: noItsmMsg },
                ]);
            });
            cell.addEventListener('mouseout', (e) => {
                if (e.relatedTarget && cell.contains(e.relatedTarget)) return;
                this._hideHoverCard();
            });
        }

        // Day number
        const dn = document.createElement('div');
        dn.className = 'jenga-cal__daynum';
        if (this._isToday(date)) {
            const b = document.createElement('span');
            b.className = 'jenga-cal__today-badge';
            b.textContent = date.getDate();
            dn.appendChild(b);
        } else { dn.textContent = date.getDate(); }
        cell.appendChild(dn);

        const cabEvent = getDayCabEvent(date, allEvents);
        if (cabEvent) {
            const cab = document.createElement('button');
            cab.type = 'button';
            cab.className = 'jenga-cal__cab-badge';
            cab.dataset.key = cabEvent.key;
            cab.textContent = '📢 Change Advisory board';
            cab.addEventListener('mouseenter', (e) => this._showHoverCard(e.clientX, e.clientY, this._cabTooltipFields(cabEvent)));
            cab.addEventListener('mouseleave', () => this._hideHoverCard());
            cab.addEventListener('click', (e) => {
                e.stopPropagation();
                this._hideHoverCard();
                this._onEventClick?.(cabEvent);
            });
            cell.appendChild(cab);
        }

        // Multi-day placeholder row (pushes single chips down)
        const placeholder = document.createElement('div');
        placeholder.className = 'jenga-cal__multiday-placeholder';
        cell.appendChild(placeholder);

        // Single-day chips
        const key = this._dayKey(date);
        const dayEvts = (singleByDay.get(key) || []).filter(ev => !isCabProtectionEvent(ev));
        const chipsArea = document.createElement('div');
        chipsArea.className = 'jenga-cal__chips';

        const items = groupSingleDayEvents(dayEvts);
        const visible = items.slice(0, MAX_CHIPS);
        const overflow = items.length - visible.length;

        visible.forEach(item => chipsArea.appendChild(this._makeChip(item)));

        if (overflow > 0) {
            const overflowItems = items.slice(visible.length);
            const btn = document.createElement('button');
            btn.className = 'jenga-chip jenga-chip--more';

            const countSpan = document.createElement('span');
            countSpan.textContent = `+${overflow} more`;
            btn.appendChild(countSpan);

            // Unique mini-icons for the overflow events
            const seenEnv  = new Set();
            const seenType = new Set();
            overflowItems.forEach(item => {
                const ev = item.event;
                if (ev.type === 'SERVICE_OP') {
                    // Deduplicate by *label* (e.g. both "prd" and "prod" → "PRD")
                    const label = getEnvLabel(ev.environment);
                    if (!seenEnv.has(label)) {
                        seenEnv.add(label);
                        const badge = document.createElement('span');
                        badge.className = 'jenga-chip__env-badge jenga-chip__env-badge--mini';
                        badge.textContent = label;
                        badge.style.background = getEnvColor(ev.environment);
                        badge.style.color = getEnvColor(ev.environment) === '#eab308' ? '#1a1a1a' : '#fff';
                        btn.appendChild(badge);
                    }
                } else {
                    if (!seenType.has(ev.type)) {
                        seenType.add(ev.type);
                        const icon = document.createElement('span');
                        icon.className = 'jenga-chip__more-icon';
                        if (ev.type === 'HYBRIS')   { icon.innerHTML = HYBRIS_SVG; }
                        else if (ev.type === 'MILESTONE')      { icon.textContent = '🏁'; }
                        else if (ev.type === 'BUSINESS_EVENT') { icon.textContent = '👜'; }
                        else                                    { icon.textContent = '📦'; }
                        btn.appendChild(icon);
                    }
                }
            });

            btn.addEventListener('click', e => { e.stopPropagation(); this._showPopover(date, items, btn); });
            chipsArea.appendChild(btn);
        }

        cell.appendChild(chipsArea);

        // Mobile summary: compact env-count badges (shown on small screens via CSS)
        cell.appendChild(this._makeMobileSummary(dayEvts));

        // Clicking the cell (including badges on mobile) opens day detail
        cell.addEventListener('click', (e) => {
            // On desktop skip if a chip or popover is clicked — open the chip action instead
            if (!e.target.closest('.jenga-cal__mobile-summary') &&
                (e.target.closest('.jenga-chip') || e.target.closest('.jenga-multiday-bar'))) return;
            this._onDayClick?.(date);
        });

        return cell;
    }

    // ── Mobile summary badges ─────────────────────────────────────────────────

    /** Build compact per-env count badges shown on mobile instead of full chips.
     *  One badge per env group (PRD, QA, STG, DEV), one [Y] for Hybris,
     *  one 🏁(N) for milestones. Entire cell is clickable → day drawer. */
    _makeMobileSummary(dayEvts) {
        const wrap = document.createElement('div');
        wrap.className = 'jenga-cal__mobile-summary';

        // Count SERVICE_OP by normalised env label
        const envCounts = new Map(); // label → { count, color }
        let hybrisCount = 0;
        let milestoneCount = 0;

        dayEvts.forEach(ev => {
            if (ev.type === 'SERVICE_OP') {
                const label = getEnvLabel(ev.environment);
                const color = getEnvColor(ev.environment);
                if (!envCounts.has(label)) envCounts.set(label, { count: 0, color });
                envCounts.get(label).count++;
            } else if (ev.type === 'HYBRIS') {
                hybrisCount++;
            } else if (ev.type === 'MILESTONE') {
                milestoneCount++;
            }
        });

        // Short single/double-char labels for compact mobile display
        const MOBILE_LABEL = {
            PRD: 'P', QA: 'Q', QA1: 'Q1', QA2: 'Q2', QA3: 'Q3',
            QAS: 'QS', QAE: 'QE', STG: 'S', DEV: 'D',
        };

        // Render env badges in priority order
        const ENV_ORDER = ['PRD', 'QA', 'QA1', 'QA2', 'QA3', 'QAS', 'QAE', 'STG', 'DEV'];
        const rendered = new Set();
        ENV_ORDER.forEach(lbl => {
            if (rendered.has(lbl) || !envCounts.has(lbl)) return;
            rendered.add(lbl);
            const { count, color } = envCounts.get(lbl);
            const badge = document.createElement('span');
            badge.className = 'jenga-cal__mb-env';
            badge.style.background = color;
            badge.style.color = color === '#eab308' ? '#1a1a1a' : '#fff';
            const shortLbl = MOBILE_LABEL[lbl] || lbl.slice(0, 1);
            badge.textContent = count > 1 ? `${shortLbl}(${count})` : shortLbl;
            wrap.appendChild(badge);
        });
        // Any env not in the priority list
        envCounts.forEach(({ count, color }, lbl) => {
            if (rendered.has(lbl)) return;
            const badge = document.createElement('span');
            badge.className = 'jenga-cal__mb-env';
            badge.style.background = color;
            badge.style.color = '#fff';
            const shortLbl = MOBILE_LABEL[lbl] || lbl.slice(0, 1);
            badge.textContent = count > 1 ? `${shortLbl}(${count})` : shortLbl;
            wrap.appendChild(badge);
        });

        // Hybris badge
        if (hybrisCount > 0) {
            const badge = document.createElement('span');
            badge.className = 'jenga-cal__mb-hybris';
            badge.innerHTML = HYBRIS_SVG + (hybrisCount > 1 ? `(${hybrisCount})` : '');
            wrap.appendChild(badge);
        }

        // Milestone badge
        if (milestoneCount > 0) {
            const badge = document.createElement('span');
            badge.className = 'jenga-cal__mb-milestone';
            badge.textContent = milestoneCount > 1 ? `🏁(${milestoneCount})` : '🏁';
            wrap.appendChild(badge);
        }

        return wrap;
    }

    // ── Multi-day spanning bars ───────────────────────────────────────────────

    _appendMultiDayBars(rowEl, week, multiDayEvents, month) {
        // Find which multi-day events overlap this week
        const weekStart = week[0];
        const weekEnd   = week[6];

        // Track vertical "lanes" to avoid overlap
        const lanes = []; // each lane is the last end date used

        multiDayEvents.forEach(ev => {
            // For single-day business events, start === end (both = dueDate)
            const evStart = ev.startDate || ev.dueDate;
            const evEnd   = ev.dueDate || evStart;

            // Does this event overlap the week?
            if (evEnd < weekStart || evStart > weekEnd) return;

            // Clamp to week boundaries
            const visStart = evStart < weekStart ? weekStart : evStart;
            const visEnd   = evEnd   > weekEnd   ? weekEnd   : evEnd;

            const startDow = (visStart.getDay() + 6) % 7; // Mon=0
            const endDow   = (visEnd.getDay()   + 6) % 7;

            // Find first free lane
            let laneIdx = lanes.findIndex(lEnd => !lEnd || lEnd < visStart);
            if (laneIdx === -1) { laneIdx = lanes.length; }
            lanes[laneIdx] = visEnd;

            const colors = TYPE_COLORS[ev.type] || TYPE_COLORS.OTHER;

            const bar = document.createElement('div');
            bar.className = `jenga-multiday-bar jenga-multiday-bar--${ev.type.toLowerCase()}`;
            bar.dataset.key = ev.key;
            bar.style.cssText = `
                left: calc(${startDow} * (100% / 7) + 2px);
                width: calc(${endDow - startDow + 1} * (100% / 7) - 4px);
                top: calc(26px + ${laneIdx} * 22px);
                background: ${colors.bg};
                border: 1px solid ${colors.border};
                color: ${colors.text};
            `;

            // Icon + label
            const icon = document.createElement('span');
            icon.className = 'jenga-multiday-bar__icon';
            if (ev.type === 'BUSINESS_EVENT') {
                icon.textContent = '👜';
            } else {
                icon.textContent = '📦';
            }
            bar.appendChild(icon);

            const label = document.createElement('span');
            label.className = 'jenga-multiday-bar__label';
            const name = this._bizName(ev);
            label.textContent = name;
            bar.appendChild(label);

            // Continuation arrows
            if (evStart < weekStart) {
                const arrow = document.createElement('span');
                arrow.className = 'jenga-multiday-bar__arrow jenga-multiday-bar__arrow--left';
                arrow.textContent = '‹';
                bar.insertBefore(arrow, bar.firstChild);
            }
            if (evEnd > weekEnd) {
                const arrow = document.createElement('span');
                arrow.className = 'jenga-multiday-bar__arrow jenga-multiday-bar__arrow--right';
                arrow.textContent = '›';
                bar.appendChild(arrow);
            }

            const _peakStr = (ev.peakDates || [])
                .map(r => `${this._fmtDate(r.start)}${r.end && r.end > r.start ? ' → ' + this._fmtDate(r.end) : ''}`)
                .join(', ') || null;
            bar.addEventListener('mouseenter', e => this._showHoverCard(e.clientX, e.clientY, [
                { label: 'Event',      value: this._bizName(ev) },
                { label: 'Start',      value: this._fmtDate(ev.startDate || ev.dueDate) },
                { label: 'End',        value: this._fmtDate(ev.dueDate) },
                { label: 'Peak dates', value: _peakStr },
            ]));
            bar.addEventListener('mouseleave', () => this._hideHoverCard());
            bar.addEventListener('click', e => {
                e.stopPropagation();
                this._onEventClick?.(ev);
            });

            rowEl.appendChild(bar);

            // 🔥 Flame icons for peak dates — one per day-column in the peak range
            if (ev.type === 'BUSINESS_EVENT' && ev.peakDates && ev.peakDates.length > 0) {
                const barTop = 26 + laneIdx * 22;
                for (let d = startDow; d <= endDow; d++) {
                    const cellDate = week[d];
                    if (this._isDateInPeakRange(cellDate, ev.peakDates)) {
                        const flame = document.createElement('div');
                        flame.className = 'jenga-peak-flame';
                        flame.textContent = '🔥';
                        flame.style.cssText = `
                            left: calc(${d} * (100% / 7) + 2px);
                            top: ${barTop}px;
                            width: calc(100% / 7 - 4px);
                        `;
                        rowEl.appendChild(flame);
                    }
                }
            }
        });

        // Adjust cell placeholder heights based on lanes used
        if (lanes.length > 0) {
            rowEl.querySelectorAll('.jenga-cal__multiday-placeholder').forEach(p => {
                p.style.height = `${26 + lanes.length * 22 + 2}px`;
            });
        }
    }

    _isDateInPeakRange(date, peakDates) {
        return isDateInPeakRange(date, peakDates);
    }

    _bizName(ev) {
        if (ev.type === 'HYBRIS') {
            const raw = ev.service || ev.summary;
            const m = raw.match(/release\/(.+)/i);
            return m ? 'Hybris ' + m[1] : raw;
        }
        const m = ev.summary.match(/^Event:([^|]+)/i);
        if (m) return m[1].trim();
        const s = ev.summary;
        return s.length > 26 ? s.slice(0, 24) + '…' : s;
    }

    // ── Single-day chip ───────────────────────────────────────────────────────

    _makeChip(item) {
        const ev = item.event;
        const colors = TYPE_COLORS[ev.type] || TYPE_COLORS.OTHER;

        const chip = document.createElement('div');
        chip.className = `jenga-chip jenga-chip--${ev.type.toLowerCase()}`;
        chip.dataset.key = ev.key;
        chip.style.setProperty('--chip-bg',     colors.bg);
        chip.style.setProperty('--chip-border', colors.border);
        chip.style.setProperty('--chip-text',   colors.text);

        if (ev.type === 'SERVICE_OP') {
            chip.append(...this._serviceChipContent(ev, item));
        } else if (ev.type === 'HYBRIS') {
            chip.append(...this._hybrisChipContent(ev));
        } else if (ev.type === 'MILESTONE') {
            chip.append(...this._milestoneChipContent(ev));
        } else {
            chip.append(...this._genericChipContent(ev));
        }

        chip.addEventListener('mouseenter', e => {
            const fields = this._chipTooltipFields(ev, item);
            this._showHoverCard(e.clientX, e.clientY, fields);
        });
        chip.addEventListener('mouseleave', () => this._hideHoverCard());
        chip.addEventListener('click', e => {
            e.stopPropagation();
            this._onEventClick?.(ev, item.grouped ? item.events : null);
        });
        return chip;
    }

    _serviceChipContent(ev, item) {
        const envMeta = ENV_META[(ev.environment||'').toLowerCase()] || { color: '#6b7280', label: (ev.environment||'').slice(0,3).toUpperCase() };
        const badge = document.createElement('span');
        badge.className = 'jenga-chip__env-badge';
        badge.textContent = envMeta.label;
        badge.style.background = envMeta.color;
        badge.style.color = envMeta.color === '#eab308' ? '#1a1a1a' : '#fff';

        const svc = document.createElement('span');
        svc.className = 'jenga-chip__service';
        const n = ev.service || '';
        svc.textContent = n.length > 18 ? n.slice(0,16)+'…' : n;

        const op = document.createElement('span');
        op.className = 'jenga-chip__operation';
        const opText = ev.operation || '';
        op.textContent = opText.length > 12 ? opText.slice(0, 10) + '…' : opText;

        const nodes = [badge, svc, op];
        if (item.grouped && item.count > 1) {
            const cnt = document.createElement('span');
            cnt.className = 'jenga-chip__count';
            cnt.textContent = `×${item.count}`;
            nodes.push(cnt);
        }
        return nodes;
    }

    _hybrisChipContent(ev) {
        const icon = document.createElement('span');
        icon.className = 'jenga-hybris-icon';
        icon.innerHTML = HYBRIS_SVG;

        const label = document.createElement('span');
        label.className = 'jenga-chip__label';
        const raw = ev.service || ev.summary;
        const m = raw.match(/release\/(.+)/i);
        label.textContent = m ? 'Hybris ' + m[1] : raw;
        return [icon, label];
    }

    _milestoneChipContent(ev) {
        const icon = document.createElement('span');
        icon.className = 'jenga-chip__type-icon';
        icon.textContent = '🏁';
        const label = document.createElement('span');
        label.className = 'jenga-chip__label';
        const s = ev.summary || ev.service || '';
        label.textContent = s.length > 22 ? s.slice(0, 20) + '…' : s;
        return [icon, label];
    }

    _genericChipContent(ev) {
        const icon = document.createElement('span');
        icon.className = 'jenga-chip__type-icon';
        icon.textContent = ev.type === 'DELIVERY' ? '📦' : '•';
        const label = document.createElement('span');
        label.className = 'jenga-chip__label';
        const s = ev.summary;
        label.textContent = s.length > 24 ? s.slice(0,22)+'…' : s;
        return [icon, label];
    }

    _chipTooltipFields(ev, item) {
        if (item.grouped) {
            const contexts = item.events.map((e, i) => e.context || `Unknown (${i + 1})`);
            return [
                { label: 'Count',       value: `${item.count} events` },
                { label: 'Service',     value: ev.service },
                { label: 'Environment', value: ev.environment },
                { label: 'Operation',   value: ev.operation },
                { label: 'Due Date',    value: this._fmtDate(ev.dueDate) },
                { label: 'Context',     value: contexts.join(', ') },
            ];
        }
        if (ev.type === 'MILESTONE') {
            return [
                { label: 'Milestone', value: ev.summary },
                { label: 'Date',      value: this._fmtDate(ev.dueDate) },
            ];
        }
        if (ev.type === 'SERVICE_OP') {
            return [
                { label: 'Service',     value: ev.service },
                { label: 'Environment', value: ev.environment },
                { label: 'Operation',   value: ev.operation },
                { label: 'Due Date',    value: this._fmtDate(ev.dueDate) },
                { label: 'Status',      value: ev.status },
                { label: 'Assignee',    value: ev.assignee },
            ];
        }
        if (ev.type === 'HYBRIS') {
            return [
                { label: 'Release',  value: this._bizName(ev) },
                { label: 'Due Date', value: this._fmtDate(ev.dueDate) },
                { label: 'Status',   value: ev.status },
            ];
        }
        // DELIVERY, OTHER
        return [
            { label: 'Summary',  value: ev.summary },
            { label: 'Due Date', value: this._fmtDate(ev.dueDate) },
            { label: 'Status',   value: ev.status },
            { label: 'Assignee', value: ev.assignee },
        ];
    }

    // ── Hover card (shared with TimelineRenderer via same DOM element) ──────────

    _pspwTooltipFields(color) {
        const policies = {
            RED:   'No deployments are allowed on this day, except for critical P1/P2 fixes required to restore production stability.',
            AMBER: 'Only deployments for approved applications are allowed.',
            BLUE:  'Production changes are allowed only through the official process. If this window falls within an AMBER window, applications in the approved list may be deployed without CAB. All others require CAB approval.',
        };
        return [
            { label: 'Peak Season Protection', value: color },
            { label: 'Policy',                 value: policies[color] || color },
        ];
    }

    _cabTooltipFields(ev) {
        const fmt = d => d ? d.toLocaleDateString('en-GB') : null;
        return [
            { label: 'Key',      value: ev.key },
            { label: 'Status',   value: ev.status },
            ev.priority && ev.priority !== 'None' && { label: 'Priority', value: ev.priority },
            { label: 'Due Date', value: fmt(ev.dueDate) },
            ev.startDate && ev.startDate.getTime() !== ev.dueDate?.getTime() &&
                { label: 'Start', value: fmt(ev.startDate) },
            ev.assignee && { label: 'Assignee', value: ev.assignee },
            ev.reporter && { label: 'Reporter', value: ev.reporter },
            ev.stream   && { label: 'Stream',   value: ev.stream },
            ev.theme    && { label: 'Theme',    value: ev.theme },
            { label: 'Summary',  value: ev.summary },
            ev.roi      && { label: 'Region',   value: ev.roi },
        ].filter(Boolean);
    }

    _initBizTooltip() {
        let card = document.getElementById('jtl-hover-card');
        if (!card) {
            card = document.createElement('div');
            card.id = 'jtl-hover-card';
            card.className = 'jtl-hover-card';
            document.body.appendChild(card);
        }
        this._hoverCard = card;
    }

    _showHoverCard(clientX, clientY, fields) {
        const card = this._hoverCard;
        if (!card) return;
        card.innerHTML = fields
            .filter(f => f.value)
            .map(f => `<div class="jtl-hover-card__row"><span class="jtl-hover-card__lbl">${f.label}</span><span class="jtl-hover-card__val">${f.value}</span></div>`)
            .join('');
        card.style.display = 'block';
        card.style.visibility = 'hidden';
        card.style.top = '-9999px';
        requestAnimationFrame(() => {
            const cw = card.offsetWidth;
            const ch = card.offsetHeight;
            let left = clientX + 14;
            let top  = clientY + 18;
            if (left + cw + 8 > window.innerWidth)  left = clientX - cw - 14;
            if (top  + ch + 8 > window.innerHeight) top  = clientY - ch - 18;
            left = Math.max(8, left);
            top  = Math.max(8, top);
            card.style.left = `${left}px`;
            card.style.top  = `${top}px`;
            card.style.visibility = 'visible';
        });
    }

    _hideHoverCard() {
        if (this._hoverCard) this._hoverCard.style.display = 'none';
    }

    _fmtDate(d) {
        return d ? d.toLocaleDateString('it-IT', { day: '2-digit', month: 'short', year: 'numeric' }) : null;
    }

    // ── Popover ───────────────────────────────────────────────────────────────

    _showPopover(date, items, anchor) {
        this._closePopover();
        const pop = document.createElement('div');
        pop.className = 'jenga-popover';
        pop.innerHTML = `<div class="jenga-popover__header">
            <span>${date.toLocaleDateString('en-US',{weekday:'short',month:'short',day:'numeric'})}</span>
            <button class="jenga-popover__close">✕</button>
        </div>`;
        pop.querySelector('.jenga-popover__close').addEventListener('click', e => { e.stopPropagation(); this._closePopover(); });
        const list = document.createElement('div');
        list.className = 'jenga-popover__list';
        items.forEach(item => { const c = this._makeChip(item); c.style.marginBottom='4px'; list.appendChild(c); });
        pop.appendChild(list);
        pop.addEventListener('click', e => e.stopPropagation());

        // Initial position — clamp after append so actual dimensions are known
        pop.style.cssText = 'position:fixed;left:-9999px;top:-9999px;z-index:9999;visibility:hidden';
        document.body.appendChild(pop);
        this._popover     = pop;
        this._popoverAnchor = anchor;
        this._positionPopover();

        // Reposition on scroll so the popover follows the anchor
        this._scrollHandler = () => {
            if (!this._popover) return;
            // Close if the anchor has scrolled out of the visible calendar area
            const calEl = document.getElementById('jenga-calendar');
            const calRect = calEl ? calEl.getBoundingClientRect() : null;
            const aRect = this._popoverAnchor.getBoundingClientRect();
            if (calRect && (aRect.bottom < calRect.top || aRect.top > calRect.bottom)) {
                this._closePopover();
            } else {
                this._positionPopover();
            }
        };
        document.getElementById('jenga-calendar')?.addEventListener('scroll', this._scrollHandler, { passive: true });
        window.addEventListener('scroll', this._scrollHandler, { passive: true });
    }

    _positionPopover() {
        const pop    = this._popover;
        const anchor = this._popoverAnchor;
        if (!pop || !anchor) return;

        const rect   = anchor.getBoundingClientRect();
        const pw     = pop.offsetWidth  || 280;
        const ph     = pop.offsetHeight || 320;
        const margin = 8;
        const vw     = window.innerWidth;
        const vh     = window.innerHeight;

        // Prefer below; flip above if not enough room
        let top = rect.bottom + 4;
        if (top + ph > vh - margin) top = rect.top - ph - 4;
        top = Math.max(margin, Math.min(top, vh - ph - margin));

        // Align left with anchor; shift left if it would overflow right edge
        let left = rect.left;
        if (left + pw > vw - margin) left = vw - pw - margin;
        left = Math.max(margin, left);

        pop.style.left       = `${left}px`;
        pop.style.top        = `${top}px`;
        pop.style.visibility = '';
    }

    _closePopover() {
        if (this._popover) { this._popover.remove(); this._popover = null; }
        if (this._scrollHandler) {
            document.getElementById('jenga-calendar')?.removeEventListener('scroll', this._scrollHandler);
            window.removeEventListener('scroll', this._scrollHandler);
            this._scrollHandler = null;
        }
        this._popoverAnchor = null;
    }

    // ── Utilities ─────────────────────────────────────────────────────────────

    _buildWeeks(firstDay, lastDay) {
        const start = new Date(firstDay);
        start.setDate(start.getDate() - (start.getDay() + 6) % 7);
        const weeks = [];
        const cur = new Date(start);
        while (cur <= lastDay || weeks.length < 4) {
            const week = [];
            for (let d = 0; d < 7; d++) { week.push(new Date(cur)); cur.setDate(cur.getDate()+1); }
            weeks.push(week);
            if (cur > lastDay && weeks.length >= 5) break;
        }
        return weeks;
    }

    setSelectedKey(key) {
        document.getElementById('jenga-calendar')
            ?.querySelectorAll('.jenga--selected')
            .forEach(el => el.classList.remove('jenga--selected'));
        if (!key) return;
        document.getElementById('jenga-calendar')
            ?.querySelector(`[data-key="${CSS.escape(key)}"]`)
            ?.classList.add('jenga--selected');
    }

    setSelectedDate(dateStr) {
        document.getElementById('jenga-calendar')
            ?.querySelectorAll('.jenga--selected')
            .forEach(el => el.classList.remove('jenga--selected'));
        if (!dateStr) return;
        document.getElementById('jenga-calendar')
            ?.querySelector(`[data-date="${dateStr}"]`)
            ?.classList.add('jenga--selected');
    }

    clearSelection() {
        document.getElementById('jenga-calendar')
            ?.querySelectorAll('.jenga--selected')
            .forEach(el => el.classList.remove('jenga--selected'));
    }

    _dayKey(d) {
        return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
    }

    _isToday(d) {
        const t = new Date();
        return d.getFullYear()===t.getFullYear() && d.getMonth()===t.getMonth() && d.getDate()===t.getDate();
    }
}
