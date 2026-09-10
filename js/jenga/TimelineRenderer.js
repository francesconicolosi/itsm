/**
 * TimelineRenderer — multi-panel dashboard timeline for Jenga.
 *
 * Layout (left panel + right sidebar):
 *   ─ Business Events   : horizontal spanning bars
 *   ─ Releases          : D3 line chart per env (PROD/QA/STG/DEV), count per day
 *   ─ Incidents Created : D3 line chart, count per day (from jira-cards)
 *   ─ Service Requests  : D3 line chart, count per day (from jira-cards)
 *
 *  Right sidebar: Environment + Service filters, auto-computed Insights.
 */

import * as d3 from 'd3';
import { getTypeColor, isDateInPeakRange } from './CalendarRenderer.js';

/** Format a local Date as YYYY-MM-DD (matches CSV values and EventStore.localDateKey). */
function localDateKey(d) {
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

// ─── ITSM data-retention cutoff ───────────────────────────────────────────────
// __ITSM_RETENTION_DAYS__ is injected at build time by webpack DefinePlugin.
// The Python export script uses the same value as --lookback-days.
const ITSM_RETENTION_DAYS = typeof __ITSM_RETENTION_DAYS__ !== 'undefined' ? __ITSM_RETENTION_DAYS__ : 100;

/** Midnight of today minus ITSM_RETENTION_DAYS — the earliest date with ITSM data. */
function itsmCutoffDate() {
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    d.setDate(d.getDate() - ITSM_RETENTION_DAYS);
    return d;
}

// ─── Constants ────────────────────────────────────────────────────────────────

const ENV_SERIES = [
    { key: 'prd',     label: 'PROD', color: '#a855f7' },
    { key: 'qa',      label: 'QA',   color: '#3b82f6' },
    { key: 'staging', label: 'STG',  color: '#10b981' },
    { key: 'devops',  label: 'DEV',  color: '#9ca3af' },
];

const ENV_KEYS = {
    prd: ['prd', 'prod'],
    qa:  ['qa', 'qa1', 'qa2', 'qa3', 'qa4', 'qa5', 'qasales', 'qaesales'],
    staging: ['staging', 'stg', 'preprod', 'shadow'],
    devops: ['devops'],
};

const CHART_MARGIN = { top: 20, right: 16, bottom: 30, left: 44 };
const CHART_HEIGHT  = 180;

const PEAK_PROTECTION_COLORS = {
    RED:   '#ef4444',
    AMBER: '#f59e0b',
    BLUE:  '#3b82f6',
};

function isCabProtectionEvent(ev) {
    if (!ev || ev.type !== 'PEAK_SEASON_PROTECTION_WINDOW') return false;
    return Boolean(ev.isCab) || /^CAB\s*[|]/i.test(ev.summary || '');
}

// ─── TimelineRenderer ─────────────────────────────────────────────────────────

// Colors for comparison overlay lines (avoids red #ef4444 and blue #3b82f6 used by main series)
const COMPARE_COLORS = ['#f59e0b', '#8b5cf6', '#10b981', '#ec4899', '#0ea5e9', '#f97316'];

export class TimelineRenderer {
    constructor(app) {
        this.app           = app;
        this._onEventClick = null;
        this._onDotClick   = null;
        this._onProtectionClick = null;
        this._onComparisonChange = null;

        // Timeline-only issue search.
        // Filters Jira cards behind Incidents Created and Service Requests Created by Summary.
        this.issueSummaryQuery = '';
        this._restoreIssueSearchFocus = false;

        // Used when a lane is expanded after being restored as collapsed.
        // D3 charts rendered while hidden can keep a compressed viewBox,
        // so we re-render the timeline once the lane is visible again.
        this._laneExpandRaf1 = null;
        this._laneExpandRaf2 = null;
        this._laneExpandRaf3 = null;

        // ── Comparison feature state (per-lane) ───────────────────────────────
        this._incidentAvgMonths    = 6;
        this._incidentLastYear     = false;
        this._incidentCompareMonths = [];   // [{year, month}] from free picker
        this._srAvgMonths          = 6;
        this._srLastYear           = false;
        this._srCompareMonths      = [];

        // Refs updated each render — used by per-lane re-render
        this._incidentChartEl  = null;
        this._incidentXDomain  = null;
        this._incidentYear     = null;
        this._incidentMonth    = null;
        this._incidentTodayFmt = null;
        this._srChartEl  = null;
        this._srXDomain  = null;
        this._srYear     = null;
        this._srMonth    = null;
        this._srTodayFmt = null;

        this._restoreComparisonState();

        // Close any open comparison dropdown when clicking outside a control wrapper
        this._dropdownCloseHandler = (e) => {
            if (!e.target.closest('.jtl-lane-ctrl-wrap')) {
                document.querySelectorAll('.jtl-compare-dropdown').forEach(d => { d.style.display = 'none'; });
            }
        };
        document.addEventListener('click', this._dropdownCloseHandler, true);
    }

    onEventClick(fn) { this._onEventClick = fn; }
    onDotClick(fn)   { this._onDotClick   = fn; }
    onProtectionClick(fn) { this._onProtectionClick = fn; }
    onComparisonChange(fn) { this._onComparisonChange = fn; }
    getIssueSummaryQuery() { return (this.issueSummaryQuery || '').trim(); }

    // ── Comparison state ──────────────────────────────────────────────────────

    _saveComparisonState() {
        try {
            localStorage.setItem('jenga-comparison-state-v1', JSON.stringify({
                incidentAvgMonths:     this._incidentAvgMonths,
                incidentLastYear:      this._incidentLastYear,
                incidentCompareMonths: this._incidentCompareMonths,
                srAvgMonths:           this._srAvgMonths,
                srLastYear:            this._srLastYear,
                srCompareMonths:       this._srCompareMonths,
            }));
        } catch (_) {}
        this._onComparisonChange?.();
    }

    _restoreComparisonState() {
        try {
            const raw = localStorage.getItem('jenga-comparison-state-v1');
            if (!raw) return;
            const s = JSON.parse(raw);
            if ([3, 6, 12, 24].includes(s.incidentAvgMonths)) this._incidentAvgMonths = s.incidentAvgMonths;
            if ([3, 6, 12, 24].includes(s.srAvgMonths))       this._srAvgMonths       = s.srAvgMonths;
            this._incidentLastYear = Boolean(s.incidentLastYear);
            this._srLastYear       = Boolean(s.srLastYear);
            this._incidentCompareMonths = Array.isArray(s.incidentCompareMonths)
                ? s.incidentCompareMonths.slice(0, 6) : [];
            this._srCompareMonths = Array.isArray(s.srCompareMonths)
                ? s.srCompareMonths.slice(0, 6) : [];
        } catch (_) {}
    }

    /** Apply comparison state from an external object (e.g. restored from URL params). */
    applyComparisonState(state = {}) {
        if ([3, 6, 12, 24].includes(state.incidentAvgMonths)) this._incidentAvgMonths = state.incidentAvgMonths;
        if ([3, 6, 12, 24].includes(state.srAvgMonths))       this._srAvgMonths       = state.srAvgMonths;
        if (typeof state.incidentLastYear === 'boolean') this._incidentLastYear = state.incidentLastYear;
        if (typeof state.srLastYear === 'boolean')       this._srLastYear       = state.srLastYear;
        if (Array.isArray(state.incidentCompareMonths))  this._incidentCompareMonths = state.incidentCompareMonths.slice(0, 6);
        if (Array.isArray(state.srCompareMonths))        this._srCompareMonths       = state.srCompareMonths.slice(0, 6);
    }

    getComparisonState() {
        return {
            incidentAvgMonths:     this._incidentAvgMonths,
            incidentLastYear:      this._incidentLastYear,
            incidentCompareMonths: this._incidentCompareMonths.slice(),
            srAvgMonths:           this._srAvgMonths,
            srLastYear:            this._srLastYear,
            srCompareMonths:       this._srCompareMonths.slice(),
        };
    }

    // ── Comparison data helpers ───────────────────────────────────────────────

    /** Build comparison series for a lane, mapped to target month's x-axis. */
    _getCompareData(laneKey, year, month, issueType) {
        const store     = this.app.store;
        const activeServices = this.app.search?.activeServices;
        const services  = activeServices && activeServices.size > 0 ? activeServices : undefined;
        const lastYear  = laneKey === 'incidents' ? this._incidentLastYear  : this._srLastYear;
        const freeMonths = laneKey === 'incidents' ? this._incidentCompareMonths : this._srCompareMonths;

        const allMonths = [];
        if (lastYear) allMonths.push({ year: year - 1, month });
        freeMonths.forEach(m => allMonths.push({ year: m.year, month: m.month }));

        return allMonths.slice(0, 6).map((m, i) => {
            const data = store.getComparisonMonthData(issueType, m.year, m.month, year, month, { services });
            const label = new Date(m.year, m.month, 1)
                .toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
            return { label, color: COMPARE_COLORS[i % COMPARE_COLORS.length], data, srcYear: m.year, srcMonth: m.month };
        });
    }

    /** Build per-day-of-month average curve for a lane. */
    _getAvgCurve(laneKey, year, month, issueType) {
        const store = this.app.store;
        const avgMonths = laneKey === 'incidents' ? this._incidentAvgMonths : this._srAvgMonths;
        const activeServices = this.app.search?.activeServices;
        const services = activeServices && activeServices.size > 0 ? activeServices : undefined;
        return store.getAverageCurveForMonth(issueType, year, month, avgMonths, { services });
    }

    /** Single O(n) pass over cards to collect which YYYY-MM months have any data. */
    _getMonthsWithCardData(issueType) {
        const months = new Set();
        (this.app.store.cards || []).forEach(c => {
            if (!c.created) return;
            if (issueType && !(c.issueType || '').toLowerCase().includes(issueType.toLowerCase())) return;
            months.add(`${c.created.getFullYear()}-${String(c.created.getMonth() + 1).padStart(2, '0')}`);
        });
        return months;
    }

    // ── Per-lane re-render ────────────────────────────────────────────────────

    _rerenderLineLane(laneKey) {
        const isInc = laneKey === 'incidents';
        const el     = isInc ? this._incidentChartEl  : this._srChartEl;
        const xDom   = isInc ? this._incidentXDomain  : this._srXDomain;
        const year   = isInc ? this._incidentYear     : this._srYear;
        const month  = isInc ? this._incidentMonth    : this._srMonth;
        const todayFmt = isInc ? this._incidentTodayFmt : this._srTodayFmt;
        if (!el || year === null || month === null || !xDom) return;

        const issueType = isInc ? 'incident' : 'service request';
        const color     = isInc ? '#ef4444' : '#3b82f6';
        const label     = isInc ? 'Incidents' : 'Service Requests';
        const dotLabel  = isInc ? 'incident' : 'service request';

        const store = this.app.store;
        const days  = this._buildDays(year, month, new Date(year, month + 1, 0).getDate());
        const mainData    = this._buildCardData(store, year, month, issueType, days, this.getIssueSummaryQuery());
        const avgCurve    = this._getAvgCurve(laneKey, year, month, issueType);
        const compareData = this._getCompareData(laneKey, year, month, issueType);
        const series      = [{ label, color, data: mainData }];

        requestAnimationFrame(() => {
            this._renderD3Line(el, series, xDom, todayFmt, dotLabel, avgCurve, compareData);
        });

        // Update controls row and legend to reflect new state
        const chartCol = el.closest('.jtl-chart-col');
        if (chartCol) {
            const existingControls = chartCol.querySelector('.jtl-lane-controls');
            if (existingControls) existingControls.replaceWith(this._buildComparisonControls(laneKey, year, month));
            const existingLegend = chartCol.querySelector('.jtl-lane-legend');
            if (existingLegend) existingLegend.replaceWith(this._buildLegend(laneKey, year, month, compareData));
        }

        this._saveComparisonState();
    }

    // ── Comparison controls DOM ───────────────────────────────────────────────

    _buildComparisonControls(laneKey, year, month) {
        const isInc      = laneKey === 'incidents';
        const avgMonths  = isInc ? this._incidentAvgMonths  : this._srAvgMonths;
        const lastYear   = isInc ? this._incidentLastYear   : this._srLastYear;
        const cmpMonths  = isInc ? this._incidentCompareMonths : this._srCompareMonths;
        const issueType  = isInc ? 'incident' : 'service request';

        const row = document.createElement('div');
        row.className = 'jtl-lane-controls';

        // ── Avg span dropdown ──────────────────────────────────────────────
        const avgWrap = document.createElement('div');
        avgWrap.className = 'jtl-lane-ctrl-wrap';

        const avgBtn = document.createElement('button');
        avgBtn.type = 'button';
        avgBtn.className = 'jtl-avg-btn';
        avgBtn.textContent = `avg: ${avgMonths}m ▾`;
        avgBtn.title = 'Average lookback span';

        const avgPanel = document.createElement('div');
        avgPanel.className = 'jtl-compare-dropdown';
        avgPanel.style.display = 'none';
        [3, 6, 12, 24].forEach(n => {
            const opt = document.createElement('div');
            opt.className = 'jtl-compare-dropdown__item' + (n === avgMonths ? ' jtl-compare-dropdown__item--active' : '');
            opt.textContent = `${n} months`;
            opt.addEventListener('click', () => {
                if (isInc) this._incidentAvgMonths = n;
                else       this._srAvgMonths       = n;
                avgPanel.style.display = 'none';
                this._rerenderLineLane(laneKey);
            });
            avgPanel.appendChild(opt);
        });

        avgBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            const open = avgPanel.style.display !== 'none';
            document.querySelectorAll('.jtl-compare-dropdown').forEach(d => { d.style.display = 'none'; });
            if (!open) avgPanel.style.display = 'block';
        });

        avgWrap.appendChild(avgBtn);
        avgWrap.appendChild(avgPanel);
        row.appendChild(avgWrap);

        // ── Last-year toggle ───────────────────────────────────────────────
        const lyBtn = document.createElement('button');
        lyBtn.type  = 'button';
        lyBtn.className = 'jtl-compare-btn' + (lastYear ? ' jtl-compare-btn--active' : '');
        lyBtn.textContent = `↩ ${year - 1}`;
        lyBtn.title = `Compare with ${new Date(year - 1, month, 1).toLocaleDateString('en-US', { month: 'long', year: 'numeric' })}`;
        lyBtn.addEventListener('click', () => {
            if (isInc) this._incidentLastYear = !this._incidentLastYear;
            else       this._srLastYear       = !this._srLastYear;
            this._rerenderLineLane(laneKey);
        });
        row.appendChild(lyBtn);

        // ── Free month-picker dropdown ─────────────────────────────────────
        const cmpWrap = document.createElement('div');
        cmpWrap.className = 'jtl-lane-ctrl-wrap';

        const cmpBtn = document.createElement('button');
        cmpBtn.type  = 'button';
        cmpBtn.className = 'jtl-compare-btn';
        cmpBtn.textContent = '+ Compare ▾';
        cmpBtn.title = 'Overlay another month';

        const cmpPanel = document.createElement('div');
        cmpPanel.className = 'jtl-compare-dropdown jtl-compare-dropdown--months';
        cmpPanel.style.display = 'none';

        const monthsWithData = this._getMonthsWithCardData(issueType);

        for (let i = 1; i <= 24; i++) {
            let y = year, m = month - i;
            while (m < 0) { m += 12; y--; }
            // Skip the month already covered by the "last year" button
            if (lastYear && y === year - 1 && m === month) continue;

            const monthKey = `${y}-${String(m + 1).padStart(2, '0')}`;
            const hasData  = monthsWithData.has(monthKey);
            const isSelected = cmpMonths.some(cm => cm.year === y && cm.month === m);
            const label = new Date(y, m, 1).toLocaleDateString('en-US', { month: 'short', year: 'numeric' });

            const opt   = document.createElement('div');
            opt.className = 'jtl-compare-dropdown__item'
                + (isSelected ? ' jtl-compare-dropdown__item--active' : '')
                + (!hasData   ? ' jtl-compare-dropdown__item--empty'  : '');

            const check = document.createElement('span');
            check.className   = 'jtl-compare-dropdown__check';
            check.textContent = isSelected ? '✓' : '';
            opt.appendChild(check);

            const lbl = document.createElement('span');
            lbl.textContent = label;
            opt.appendChild(lbl);

            if (hasData) {
                opt.style.cursor = 'pointer';
                opt.addEventListener('click', () => {
                    const arr = isInc ? this._incidentCompareMonths : this._srCompareMonths;
                    const idx = arr.findIndex(cm => cm.year === y && cm.month === m);
                    if (idx >= 0) {
                        arr.splice(idx, 1);
                    } else {
                        const totalSelected = arr.length + (isInc ? (this._incidentLastYear ? 1 : 0) : (this._srLastYear ? 1 : 0));
                        if (totalSelected >= 6) return; // max 6
                        arr.push({ year: y, month: m });
                    }
                    if (isInc) this._incidentCompareMonths = arr;
                    else       this._srCompareMonths       = arr;
                    this._rerenderLineLane(laneKey);
                });
            }

            cmpPanel.appendChild(opt);
        }

        cmpBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            const open = cmpPanel.style.display !== 'none';
            document.querySelectorAll('.jtl-compare-dropdown').forEach(d => { d.style.display = 'none'; });
            if (!open) cmpPanel.style.display = 'block';
        });

        cmpWrap.appendChild(cmpBtn);
        cmpWrap.appendChild(cmpPanel);
        row.appendChild(cmpWrap);

        return row;
    }

    // ── Per-lane comparison legend ────────────────────────────────────────────

    _buildLegend(laneKey, year, month, compareData = null) {
        const isInc     = laneKey === 'incidents';
        const lastYear  = isInc ? this._incidentLastYear      : this._srLastYear;
        const cmpMonths = isInc ? this._incidentCompareMonths : this._srCompareMonths;

        const legend = document.createElement('div');
        legend.className = 'jtl-lane-legend';

        const allCmp = [];
        if (lastYear) allCmp.push({ year: year - 1, month, isLastYear: true });
        cmpMonths.forEach(m => allCmp.push(m));

        if (allCmp.length === 0) return legend;

        legend.classList.add('jtl-lane-legend--active');

        allCmp.slice(0, 6).forEach((m, i) => {
            const color = COMPARE_COLORS[i % COMPARE_COLORS.length];
            const item  = document.createElement('div');
            item.className   = 'jtl-lane-legend__item';
            item.style.cursor = 'pointer';
            item.title = 'Click to remove';

            const dot = document.createElement('span');
            dot.className        = 'jtl-lane-legend__dot';
            dot.style.background = color;

            const lbl = document.createElement('span');
            lbl.textContent = new Date(m.year, m.month, 1)
                .toLocaleDateString('en-US', { month: 'short', year: 'numeric' });

            item.appendChild(dot);
            item.appendChild(lbl);

            const cmpTotal = compareData?.[i]?.data?.reduce((s, d) => s + d.value, 0);
            if (cmpTotal !== undefined) {
                const tot = document.createElement('span');
                tot.className = 'jtl-lane-legend__total';
                tot.textContent = `${cmpTotal} total`;
                item.appendChild(tot);
            }

            item.addEventListener('click', () => {
                if (m.isLastYear) {
                    if (isInc) this._incidentLastYear = false;
                    else       this._srLastYear       = false;
                } else {
                    const arr = isInc ? this._incidentCompareMonths : this._srCompareMonths;
                    const idx = arr.findIndex(cm => cm.year === m.year && cm.month === m.month);
                    if (idx >= 0) {
                        arr.splice(idx, 1);
                        if (isInc) this._incidentCompareMonths = arr;
                        else       this._srCompareMonths       = arr;
                    }
                }
                this._rerenderLineLane(laneKey);
            });

            legend.appendChild(item);
        });

        return legend;
    }

    render(containerEl, year, month, events) {
        if (!containerEl) return;
        containerEl.innerHTML = '';
        this._insightsPanel?.remove();
        this._insightsPanel = null;
        this._initBizTooltip();

        const store        = this.app.store;
        const daysInMonth  = new Date(year, month + 1, 0).getDate();
        const days         = this._buildDays(year, month, daysInMonth);
        const today        = new Date();
        const todayFmt     = today.getFullYear() === year && today.getMonth() === month
            ? localDateKey(today) : null;

        // JengaApp.refresh() passes current + adjacent month events for the calendar grid.
        // The timeline only shows events that actually overlap this month.
        // Normalize start/due because some CSV rows may contain StartDate after DueDate.
        const monthStart = new Date(year, month, 1);
        const monthEnd   = new Date(year, month + 1, 0, 23, 59, 59);
        events = events.filter(e => {
            const due   = e.dueDate;
            const start = e.startDate || due;
            if (!due && !start) return false;

            const rawStart = start || due;
            const rawEnd   = due   || start;
            const evStart  = rawStart <= rawEnd ? rawStart : rawEnd;
            const evEnd    = rawStart <= rawEnd ? rawEnd   : rawStart;

            return evStart <= monthEnd && evEnd >= monthStart;
        });

        // ── Full-width layout (no sidebar) ───────────────────────────────────
        const main = document.createElement('div');
        main.className = 'jtl-main';
        containerEl.appendChild(main);

        // ── Shared X axis domain ──────────────────────────────────────────────
        const xDomain = [new Date(year, month, 1), new Date(year, month, daysInMonth)];

        // ── Month label + day-number axis (shown above biz/milestone lanes) ───
        const axisHeader = document.createElement('div');
        axisHeader.className = 'jtl-axis-header';
        const monthLabel = document.createElement('div');
        monthLabel.className = 'jtl-month-label';
        monthLabel.textContent = new Date(year, month, 1)
            .toLocaleDateString('en-US', { month: 'long', year: 'numeric' }).toUpperCase();
        axisHeader.appendChild(monthLabel);
        axisHeader.appendChild(this._buildDayAxisRuler(year, month, main));
        main.appendChild(axisHeader);

        // ── Business Events row ───────────────────────────────────────────────
        const bizEvents = events.filter(e => e.type === 'BUSINESS_EVENT');

        // Compute flame days early so _buildMilestonesRow can use them for orange bands.
        this._flameDays = new Set();
        bizEvents.forEach(ev => {
            if (!ev.peakDates || !ev.peakDates.length) return;
            for (let d = 1; d <= daysInMonth; d++) {
                const cellDate = new Date(year, month, d);
                if (isDateInPeakRange(cellDate, ev.peakDates)) {
                    this._flameDays.add(localDateKey(cellDate));
                }
            }
        });

        const bizLane = this._makeCollapsible('biz',
            false,
            this._buildBizEventsRow(bizEvents, year, month, daysInMonth, todayFmt)
        );
        main.appendChild(bizLane);

        // ── Milestones row (collapsed by default) ─────────────────────────────
        // Milestones are point-in-time events: show them only in the month of their Due Date.
        // StartDate must not make a milestone appear in previous months.
        const milestoneEvents = events.filter(e =>
            e.type === 'MILESTONE' &&
            e.dueDate &&
            e.dueDate.getFullYear() === year &&
            e.dueDate.getMonth() === month
        );
        let milestoneLane = null;
        if (milestoneEvents.length) {
            milestoneLane = this._makeCollapsible('milestones',
                true,
                this._buildMilestonesRow(milestoneEvents, year, month, daysInMonth, todayFmt)
            );
            main.appendChild(milestoneLane);
        }

        // Hide the axis header when both biz and milestones lanes are collapsed.
        const _syncAxisHeader = () => {
            const bizCollapsed  = bizLane.classList.contains('jtl-lane--collapsed');
            const msCollapsed   = !milestoneLane || milestoneLane.classList.contains('jtl-lane--collapsed');
            axisHeader.style.display = (bizCollapsed && msCollapsed) ? 'none' : '';
        };
        _syncAxisHeader();
        [bizLane, milestoneLane].forEach(lane => {
            if (!lane) return;
            lane.querySelector('.jtl-lane__toggle')?.addEventListener('click', () =>
                requestAnimationFrame(_syncAxisHeader)
            );
        });

        // ── Peak Season Protection Window row ────────────────────────────────
        const peakProtectionEvents = events.filter(e => e.type === 'PEAK_SEASON_PROTECTION_WINDOW');
        // Used by the D3 line panels below to render faint background protection bands.
        // CAB markers are intentionally excluded: they are point markers, not protection phases.
        this._peakProtectionEvents = peakProtectionEvents.filter(ev => !isCabProtectionEvent(ev));

        if (peakProtectionEvents.length) {
            main.appendChild(this._makeCollapsible('peak-protection',
                false,
                this._buildPeakProtectionRow(peakProtectionEvents, year, month, daysInMonth, todayFmt)
            ));
        }

        // ── Releases chart ────────────────────────────────────────────────────
        const releaseData = this._buildReleaseData(events, days);
        main.appendChild(this._makeCollapsible('releases',
            false,
            this._buildLinePanel({
                title: 'RELEASES',
                subtitle: '(count per day)',
                icon: '🚀',
                iconColor: '#a855f7',
                series: releaseData,
                xDomain,
                todayFmt,
                dotClickLabel: 'release',
            })
        ));

        // ── Incidents chart ───────────────────────────────────────────────────
        const issueSummaryQuery = this.getIssueSummaryQuery();
        const incidentData = this._buildCardData(store, year, month, 'incident', days, issueSummaryQuery);

        // ── Service Requests chart (built here so criticalDays can be merged before rendering) ──
        const srData = this._buildCardData(store, year, month, 'service request', days, issueSummaryQuery);


        const incidentTotal = incidentData.reduce((s, d) => s + d.value, 0);
        main.appendChild(this._makeCollapsible('incidents',
            false,
            this._buildLinePanel({
                title: 'INCIDENTS\nCREATED',
                subtitle: '(count per day)',
                icon: '⚠️',
                iconColor: '#ef4444',
                series: [{ label: 'Incidents', color: '#ef4444', data: incidentData }],
                xDomain,
                todayFmt,
                dotClickLabel: 'incident',
                totalCount: incidentTotal,
                laneKey: 'incidents',
                year,
                month,
            })
        ));

        // ── Service Requests chart ────────────────────────────────────────────
        const srTotal = srData.reduce((s, d) => s + d.value, 0);
        main.appendChild(this._makeCollapsible('sr',
            false,
            this._buildLinePanel({
                title: 'SERVICE\nREQUESTS\nCREATED',
                subtitle: '(count per day)',
                icon: '💬',
                iconColor: '#3b82f6',
                series: [{ label: 'Service Requests', color: '#3b82f6', data: srData }],
                xDomain,
                todayFmt,
                dotClickLabel: 'service request',
                totalCount: srTotal,
                laneKey: 'sr',
                year,
                month,
            })
        ));

        // ── Append insights panel as a row inside .jenga-top-bar ────────────────
        // The top bar is sticky so the panel gets stickiness for free and inherits
        // the top bar's background. Visibility is gated by body.jenga-view--timeline.
        const insightsPanel = this._buildInsightsPanel({ incidentData, srData, releaseData, days, todayFmt, issueSummaryQuery });
        this._insightsPanel = insightsPanel;
        const topBar = document.querySelector('.jenga-top-bar');
        if (topBar) topBar.appendChild(insightsPanel);
        else containerEl.insertAdjacentElement('beforebegin', insightsPanel);

        // Re-measure top bar height now that the insights row has been added,
        // so --jenga-topbar-height reflects the full bar including the insights row.
        requestAnimationFrame(() => this.app?._refreshTopBarPinnedGeometry?.());

    }

    // ── Filter bar ────────────────────────────────────────────────────────────

    _buildFilterBar() {
        const bar = document.createElement('div');
        bar.className = 'jtl-filterbar';

        const series = [
            { key: 'biz',       label: 'Business Events',     color: '#db2777' },
            { key: 'prd',       label: 'Releases (PROD)',      color: '#a855f7' },
            { key: 'qa',        label: 'Releases (QA)',        color: '#3b82f6' },
            { key: 'staging',   label: 'Releases (STG)',       color: '#10b981' },
            { key: 'devops',    label: 'Releases (DEV)',       color: '#9ca3af' },
            { key: 'incidents', label: 'Incidents Created',    color: '#ef4444' },
            { key: 'sr',        label: 'Service Requests Created', color: '#3b82f6' },
        ];

        series.forEach(({ key, label, color }) => {
            const item = document.createElement('label');
            item.className = 'jtl-filterbar__item';

            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.checked = true;
            cb.className = 'jtl-filterbar__cb';
            cb.dataset.key = key;

            const swatch = document.createElement('span');
            swatch.className = 'jtl-filterbar__swatch';
            swatch.style.background = color;

            const text = document.createElement('span');
            text.textContent = label;

            item.appendChild(cb);
            item.appendChild(swatch);
            item.appendChild(text);
            bar.appendChild(item);
        });

        return bar;
    }

    // ── Business Events ───────────────────────────────────────────────────────

    _buildBizEventsRow(events, year, month, daysInMonth, todayFmt) {
        const wrap = document.createElement('div');
        wrap.className = 'jtl-panel jtl-panel--biz';

        const label = document.createElement('div');
        label.className = 'jtl-panel__label';
        label.innerHTML = '<span class="jtl-panel__icon" style="color:#db2777">📅</span>' +
            '<span class="jtl-panel__title" style="color:#db2777">BUSINESS<br>EVENTS</span>';
        wrap.appendChild(label);

        const chartArea = document.createElement('div');
        chartArea.className = 'jtl-panel__chart jtl-panel__chart--biz';
        wrap.appendChild(chartArea);

        if (!events.length) {
            chartArea.innerHTML = '<p class="jtl-empty">No business events this month.</p>';
            return wrap;
        }

        const bars = document.createElement('div');
        bars.className = 'jtl-biz-bars';
        this._alignBarsToD3PlotArea(bars);

        const monthStart = new Date(year, month, 1);
        const monthEnd   = new Date(year, month, daysInMonth);

        // Sort by clamped start day
        const sorted = [...events].sort((a, b) => {
            const as = this._clampToMonth(a.startDate || a.dueDate, monthStart, monthEnd).getDate();
            const bs = this._clampToMonth(b.startDate || b.dueDate, monthStart, monthEnd).getDate();
            return as - bs;
        });

        const laneEnds = [];
        sorted.forEach(ev => {
            const rawStart = ev.startDate || ev.dueDate;
            const rawEnd   = ev.dueDate   || ev.startDate;
            if (!rawStart || !rawEnd) return;

            // Clamp both ends to the current month so bars are always within bounds
            const startDay = this._clampToMonth(rawStart, monthStart, monthEnd).getDate();
            const endDay   = this._clampToMonth(rawEnd,   monthStart, monthEnd).getDate();

            // Business events are all-day spans and must align with the D3 dot scale.
            const leftPct  = this._dayToPct(startDay, daysInMonth);
            const widthPct = this._daySpanFromDotPct(startDay, endDay, daysInMonth);

            let lane = laneEnds.findIndex(e => e < startDay);
            if (lane === -1) lane = laneEnds.length;
            laneEnds[lane] = endDay;

            const bar = document.createElement('div');
            bar.className = 'jtl-biz-bar';
            bar.style.left       = `${leftPct}%`;
            bar.style.width      = `${widthPct}%`;
            bar.style.top        = `${lane * 34}px`;
            bar.style.background = getTypeColor('BUSINESS_EVENT') + '22';
            bar.style.borderLeft = `3px solid ${getTypeColor('BUSINESS_EVENT')}`;
            bar.style.color      = getTypeColor('BUSINESS_EVENT');

            const icon = document.createElement('span');
            icon.textContent = '🎁 ';
            const txt = document.createElement('span');
            txt.textContent = this._bizLabel(ev);
            bar.appendChild(icon);
            bar.appendChild(txt);

            const _fmt = d => d ? d.toLocaleDateString('it-IT', { day: '2-digit', month: 'short', year: 'numeric' }) : null;
            const _peakStr = (ev.peakDates || [])
                .map(r => `${_fmt(r.start)}${r.end && r.end > r.start ? ' → ' + _fmt(r.end) : ''}`)
                .join(', ') || null;
            bar.addEventListener('mouseenter', (e) => this._showHoverCard(e.clientX, e.clientY, [
                { label: 'Event',      value: this._bizLabel(ev) },
                { label: 'Start',      value: _fmt(ev.startDate || ev.dueDate) },
                { label: 'End',        value: _fmt(ev.dueDate) },
                { label: 'Peak dates', value: _peakStr },
            ]));
            bar.addEventListener('mouseleave', () => this._hideHoverCard());
            bar.addEventListener('click', () => this._onEventClick?.(ev));
            bars.appendChild(bar);

            // 🔥 Flame icons for peak days within this bar's span
            if (ev.peakDates && ev.peakDates.length > 0) {
                for (let d = startDay; d <= endDay; d++) {
                    const cellDate = new Date(year, month, d);
                    if (isDateInPeakRange(cellDate, ev.peakDates)) {
                        const flame = document.createElement('div');
                        flame.className = 'jtl-peak-flame';
                        flame.textContent = '🔥';
                        flame.style.left = this._dotPctCss(d, daysInMonth);
                        flame.style.top  = `${lane * 34}px`;
                        bars.appendChild(flame);
                    }
                }
            }
        });

        const lanes = laneEnds.length || 1;
        bars.style.minHeight = `${lanes * 34 + 8}px`;

        if (todayFmt) {
            const todayDay = parseInt(todayFmt.slice(8, 10), 10);
            const todayLine = document.createElement('div');
            todayLine.className = 'jtl-today-line';
            todayLine.style.left = this._dotPctCss(todayDay, daysInMonth);
            bars.appendChild(todayLine);
        }

        this._appendNoItsmBarOverlay(bars, year, month, daysInMonth);
        chartArea.appendChild(bars);
        return wrap;
    }


    // ── Peak Season Protection Window ─────────────────────────────────────────

    _buildPeakProtectionRow(events, year, month, daysInMonth, todayFmt) {
        const wrap = document.createElement('div');
        wrap.className = 'jtl-panel jtl-panel--peak-protection';

        const label = document.createElement('div');
        label.className = 'jtl-panel__label';
        label.innerHTML = '<span class="jtl-panel__icon" style="color:#f59e0b">🛡️</span>' +
            '<span class="jtl-panel__title" style="color:#f59e0b">PEAK<br>PROTECTION</span>';
        wrap.appendChild(label);

        const chartArea = document.createElement('div');
        chartArea.className = 'jtl-panel__chart jtl-panel__chart--biz jtl-panel__chart--peak-protection';
        wrap.appendChild(chartArea);

        if (!events.length) {
            chartArea.innerHTML = '<p class="jtl-empty">No peak protection windows this month.</p>';
            return wrap;
        }

        const bars = document.createElement('div');
        bars.className = 'jtl-biz-bars jtl-peak-protection-bars';
        this._alignBarsToD3PlotArea(bars);

        const monthStart = new Date(year, month, 1);
        const monthEnd   = new Date(year, month, daysInMonth);

        const cabEvents = events.filter(isCabProtectionEvent);
        const barEvents = events.filter(ev => !isCabProtectionEvent(ev));

        const normalized = barEvents
            .map(ev => {
                const rawStart = ev.startDate || ev.dueDate;
                const rawEnd   = ev.dueDate   || ev.startDate;
                if (!rawStart || !rawEnd) return null;
                const evStart = rawStart <= rawEnd ? rawStart : rawEnd;
                const evEnd   = rawStart <= rawEnd ? rawEnd   : rawStart;
                return { ev, evStart, evEnd };
            })
            .filter(Boolean)
            .sort((a, b) => a.evStart - b.evStart || a.evEnd - b.evEnd);

        const laneEnds = [];
        normalized.forEach(({ ev, evStart, evEnd }) => {
            const startDay = this._clampToMonth(evStart, monthStart, monthEnd).getDate();
            const endDay   = this._clampToMonth(evEnd,   monthStart, monthEnd).getDate();

            // PSPW phases must align with the same horizontal scale used by the D3 dots.
            const leftPct  = this._dayToPct(startDay, daysInMonth);
            const widthPct = this._daySpanFromDotPct(startDay, endDay, daysInMonth);

            let lane = laneEnds.findIndex(e => e < startDay);
            if (lane === -1) lane = laneEnds.length;
            laneEnds[lane] = endDay;

            const colorKey = ev.protectionColor || 'AMBER';
            const color = PEAK_PROTECTION_COLORS[colorKey] || PEAK_PROTECTION_COLORS.AMBER;

            const bar = document.createElement('button');
            bar.type = 'button';
            bar.className = `jtl-biz-bar jtl-peak-protection-bar jtl-peak-protection-bar--${colorKey.toLowerCase()}`;
            bar.style.left       = `${leftPct}%`;
            bar.style.width      = `${widthPct}%`;
            bar.style.top        = `${lane * 34}px`;
            bar.style.background = `${color}22`;
            bar.style.borderLeft = `3px solid ${color}`;
            bar.style.color      = color;

            const icon = document.createElement('span');
            icon.textContent = '🛡️ ';

            const txt = document.createElement('span');
            txt.textContent = `${colorKey} Protection Window`;

            bar.appendChild(icon);
            bar.appendChild(txt);

            const _fmtPW = d => d ? d.toLocaleDateString('it-IT', { day: '2-digit', month: 'short', year: 'numeric' }) : null;
            bar.addEventListener('mouseenter', (e) => this._showHoverCard(e.clientX, e.clientY, [
                { label: 'Window',     value: ev.summary || `${colorKey} Protection Window` },
                { label: 'Protection', value: colorKey },
                { label: 'Start',      value: _fmtPW(evStart) },
                { label: 'End',        value: _fmtPW(evEnd) },
            ]));
            bar.addEventListener('mouseleave', () => this._hideHoverCard());
            bar.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                this._hideHoverCard();
                // Open the DayDrawer for the first visible day of the phase.
                // The drawer itself will load the full day and include the protection-window detail.
                this._onProtectionClick?.(new Date(year, month, startDay), ev);
            });

            bars.appendChild(bar);
        });

        const lanes = laneEnds.length || 1;
        const hasCabMarkers = cabEvents.length > 0;
        const cabMarkerTop = lanes * 34 + 2;
        bars.style.minHeight = `${lanes * 34 + (hasCabMarkers ? 34 : 8)}px`;

        cabEvents.forEach(ev => {
            const rawDate = ev.startDate || ev.dueDate;
            if (!rawDate) return;
            const markerDate = this._clampToMonth(rawDate, monthStart, monthEnd);
            const day = markerDate.getDate();

            const marker = document.createElement('button');
            marker.type = 'button';
            marker.className = 'jtl-cab-marker';
            marker.textContent = '📢';
            marker.style.left = this._cabMarkerLeftCss(day, daysInMonth);
            // CAB is a point marker in the PSPW lane: place it below the phase bars,
            // not on top of RED/AMBER/BLUE windows.
            marker.style.top = `${cabMarkerTop}px`;
            const _fmtCAB = d => d ? d.toLocaleDateString('it-IT', { day: '2-digit', month: 'short', year: 'numeric' }) : null;
            marker.addEventListener('mouseenter', (e) => this._showHoverCard(e.clientX, e.clientY, [
                { label: 'CAB',  value: ev.summary || 'Change Advisory Board' },
                { label: 'Date', value: _fmtCAB(ev.startDate || ev.dueDate) },
            ]));
            marker.addEventListener('mouseleave', () => this._hideHoverCard());
            marker.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                this._hideHoverCard();
                // Same behaviour as CalendarRenderer's CAB badge: open the CAB event detail.
                if (typeof this._onEventClick === 'function') this._onEventClick(ev);
                else this._onProtectionClick?.(new Date(year, month, day), ev);
            });
            bars.appendChild(marker);
        });

        if (todayFmt) {
            const todayDay = parseInt(todayFmt.slice(8, 10), 10);
            const todayLine = document.createElement('div');
            todayLine.className = 'jtl-today-line';
            todayLine.style.left = this._dotPctCss(todayDay, daysInMonth);
            bars.appendChild(todayLine);
        }

        this._appendNoItsmBarOverlay(bars, year, month, daysInMonth);
        chartArea.appendChild(bars);
        return wrap;
    }

    // ── Milestones row ────────────────────────────────────────────────────────

    _buildMilestonesRow(events, year, month, daysInMonth, todayFmt) {
        const wrap = document.createElement('div');
        wrap.className = 'jtl-panel jtl-panel--milestones';

        const label = document.createElement('div');
        label.className = 'jtl-panel__label';
        label.innerHTML = '<span class="jtl-panel__icon" style="color:#7c3aed">🏁</span>' +
            '<span class="jtl-panel__title" style="color:#7c3aed">MILESTONES</span>';
        wrap.appendChild(label);

        const chartArea = document.createElement('div');
        chartArea.className = 'jtl-panel__chart jtl-panel__chart--biz';
        wrap.appendChild(chartArea);

        if (!events.length) {
            chartArea.innerHTML = '<p class="jtl-empty">No milestones this month.</p>';
            return wrap;
        }

        const bars = document.createElement('div');
        bars.className = 'jtl-biz-bars';
        // Milestones are point-in-time markers: align their date positions to the
        // same D3 plot area used by PSPW bars and release dots.
        this._alignBarsToD3PlotArea(bars);

        const monthStart = new Date(year, month, 1);
        const monthEnd   = new Date(year, month, daysInMonth);
        const typeColor  = getTypeColor('MILESTONE');

        // Sort by due date
        const sorted = [...events].sort((a, b) => (a.dueDate || 0) - (b.dueDate || 0));

        // Per-lane tracking: store { leftPct, labelEl } for each placed marker so we
        // can do a second pass to set max-width on each label based on real bar width.
        const DIAMOND_GAP_PX = 21; // diamond (14px) + gap (7px)
        const CHAR_PX_EST    = 6.5;
        const CHART_PX_EST   = daysInMonth * 40;
        // laneItems[lane] = array of { leftPct, labelEl } sorted by leftPct ascending
        const laneItems = [];
        // laneRightPct[lane] = estimated right edge of last placed marker (% of chart)
        const laneRightPct = [];

        sorted.forEach(ev => {
            if (!ev.dueDate) return;
            // Milestones are point-in-time: render only on their actual Due Date,
            // and only when that Due Date belongs to the currently rendered month.
            if (ev.dueDate.getFullYear() !== year || ev.dueDate.getMonth() !== month) return;
            const day = ev.dueDate.getDate();
            // Milestones are point-in-time: align to the D3 scale (same formula as _dayToPct)
            const leftPct  = this._dayToPct(day, daysInMonth);
            const labelPx  = (ev.summary.length * CHAR_PX_EST) + DIAMOND_GAP_PX;
            const rightPct = leftPct + (labelPx / CHART_PX_EST) * 100;

            // Find a lane whose last marker's estimated right edge doesn't collide.
            // Use a 0.5% gap buffer (≈ half a day-cell) to give a small breathing room.
            let lane = laneRightPct.findIndex(r => r <= leftPct - 0.5);
            if (lane === -1) lane = laneRightPct.length;
            laneRightPct[lane] = rightPct;

            if (!laneItems[lane]) laneItems[lane] = [];
            const lbl = document.createElement('span');
            lbl.className = 'jtl-milestone-label';
            lbl.textContent = ev.summary;
            laneItems[lane].push({ leftPct, labelEl: lbl });

            const marker = document.createElement('div');
            marker.className = 'jtl-milestone-marker';
            marker.style.left  = `${leftPct}%`;
            marker.style.top   = `${lane * 42}px`;
            marker.style.color = typeColor;

            const stem = document.createElement('div');
            stem.className = 'jtl-milestone-stem';

            const point = document.createElement('div');
            point.className = 'jtl-milestone-point';

            const diamond = document.createElement('div');
            diamond.className = 'jtl-milestone-diamond';

            point.appendChild(diamond);
            point.appendChild(lbl);
            marker.appendChild(stem);
            marker.appendChild(point);
            const _fmtM = d => d ? d.toLocaleDateString('it-IT', { day: '2-digit', month: 'short', year: 'numeric' }) : null;
            marker.addEventListener('mouseenter', (e) => this._showHoverCard(e.clientX, e.clientY, [
                { label: 'Milestone', value: ev.summary },
                { label: 'Date',      value: _fmtM(ev.dueDate) },
            ]));
            marker.addEventListener('mouseleave', () => this._hideHoverCard());
            marker.addEventListener('click', () => this._onEventClick?.(ev));
            bars.appendChild(marker);
        });

        const lanes = laneRightPct.length || 1;
        bars.style.minHeight = `${lanes * 42 + 16}px`;

        // Second pass: clamp each label's max-width to the available space before the
        // next marker in the same lane, measured against the real rendered bar width.
        requestAnimationFrame(() => {
            const barsW = bars.offsetWidth;
            if (!barsW) return;
            laneItems.forEach(items => {
                items.forEach((item, i) => {
                    const nextLeftPct = i + 1 < items.length ? items[i + 1].leftPct : 100;
                    const availPx = ((nextLeftPct - item.leftPct) / 100) * barsW - DIAMOND_GAP_PX - 4; // 4px gap before next marker
                    if (availPx > 0) item.labelEl.style.maxWidth = `${Math.floor(availPx)}px`;
                });
            });
        });

        // Orange flame-day bands: same peak dates used by the D3 timeline lanes
        if (this._flameDays && this._flameDays.size) {
            const bandOpacity = document.documentElement?.dataset?.theme === 'dark' ? 0.18 : 0.13;
            this._flameDays.forEach(dateKey => {
                const d = new Date(dateKey + 'T00:00:00');
                if (d.getFullYear() !== year || d.getMonth() !== month) return;
                const day = d.getDate();
                const band = document.createElement('div');
                band.className = 'jtl-milestone-flame-band';
                band.style.left    = this._dotPctCss(day, daysInMonth);
                band.style.width   = `${(1 / Math.max(daysInMonth - 1, 1)) * 100}%`;
                band.style.opacity = bandOpacity;
                bars.appendChild(band);
            });
        }

        if (todayFmt) {
            const todayDay = parseInt(todayFmt.slice(8, 10), 10);
            const todayLine = document.createElement('div');
            todayLine.className = 'jtl-today-line';
            todayLine.style.left = this._dotPctCss(todayDay, daysInMonth);
            bars.appendChild(todayLine);
        }

        this._appendNoItsmBarOverlay(bars, year, month, daysInMonth);
        chartArea.appendChild(bars);
        return wrap;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    /** Clamp a date to [monthStart, monthEnd]. */
    _clampToMonth(date, monthStart, monthEnd) {
        if (!date) return monthStart;
        if (date < monthStart) return monthStart;
        if (date > monthEnd)   return monthEnd;
        return date;
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

    /** Human-readable label for a BUSINESS_EVENT — strips "Event:" prefix and pipe-segments. */
    _bizLabel(ev) {
        return ev.summary.replace(/^Event:/i, '').split('|')[0].trim() || ev.summary;
    }

    // ── Line chart panel ──────────────────────────────────────────────────────

    _buildLinePanel({ title, subtitle, icon, iconColor, series, xDomain, todayFmt, dotClickLabel, totalCount, laneKey, year, month }) {
        const isItsmLane = laneKey === 'incidents' || laneKey === 'sr';
        const wrap = document.createElement('div');
        wrap.className = 'jtl-panel' + (isItsmLane ? ' jtl-panel--itsm' : '');

        const label = document.createElement('div');
        label.className = 'jtl-panel__label';
        label.innerHTML =
            `<span class="jtl-panel__icon" style="color:${iconColor}">${icon}</span>` +
            `<span class="jtl-panel__title" style="color:${iconColor}">${title.replace(/\n/g, '<br>')}</span>` +
            `<span class="jtl-panel__subtitle">${subtitle}</span>`;

        if (totalCount !== undefined) {
            const tot = document.createElement('span');
            tot.className = 'jtl-panel__total';
            tot.title = 'Total for this month';
            tot.textContent = `${totalCount} total`;
            label.appendChild(tot);
        }

        // Legend inside label
        if (series.length > 1) {
            const leg = document.createElement('ul');
            leg.className = 'jtl-panel__legend';
            series.forEach(s => {
                const li = document.createElement('li');
                li.innerHTML = `<span class="jtl-panel__legend-dot" style="background:${s.color}"></span>${s.label}`;
                leg.appendChild(li);
            });
            label.appendChild(leg);
        }
        wrap.appendChild(label);

        const chartArea = document.createElement('div');
        chartArea.className = 'jtl-panel__chart';

        const issueType   = laneKey === 'incidents' ? 'incident'
                          : laneKey === 'sr'        ? 'service request'
                          : null;
        const avgCurve    = issueType ? this._getAvgCurve(laneKey, year, month, issueType)    : null;
        const compareData = issueType ? this._getCompareData(laneKey, year, month, issueType) : null;

        if (isItsmLane && year !== undefined) {
            // Store lane context for per-lane re-render
            if (laneKey === 'incidents') {
                this._incidentXDomain  = xDomain;
                this._incidentYear     = year;
                this._incidentMonth    = month;
                this._incidentTodayFmt = todayFmt;
            } else {
                this._srXDomain  = xDomain;
                this._srYear     = year;
                this._srMonth    = month;
                this._srTodayFmt = todayFmt;
            }

            const chartCol = document.createElement('div');
            chartCol.className = 'jtl-chart-col';
            chartCol.appendChild(this._buildComparisonControls(laneKey, year, month));
            chartCol.appendChild(chartArea);
            chartCol.appendChild(this._buildLegend(laneKey, year, month, compareData));
            wrap.appendChild(chartCol);

            if (laneKey === 'incidents') this._incidentChartEl = chartArea;
            else                         this._srChartEl       = chartArea;
        } else {
            wrap.appendChild(chartArea);
        }

        // Render D3 chart after DOM insertion (needs dimensions)
        requestAnimationFrame(() => {
            this._renderD3Line(chartArea, series, xDomain, todayFmt, dotClickLabel, avgCurve, compareData);
        });

        return wrap;
    }

    _renderD3Line(el, series, xDomain, todayFmt, dotClickLabel, avgCurve = null, compareData = null) {
        if (!el || !el.isConnected) return;

        // Do not render D3 charts while their lane is collapsed.
        // Rendering inside display:none / compact containers produces a compressed viewBox.
        if (el.closest('.jtl-lane--collapsed')) {
            el.innerHTML = '';
            return;
        }

        const rawWidth = Math.floor(el.getBoundingClientRect?.().width || el.clientWidth || 0);
        const attempt = Number(el.dataset.jtlRenderAttempt || '0');

        // If layout has not settled yet, retry a couple of frames before using a fallback.
        if (rawWidth < 240 && attempt < 3) {
            el.dataset.jtlRenderAttempt = String(attempt + 1);
            requestAnimationFrame(() => this._renderD3Line(el, series, xDomain, todayFmt, dotClickLabel, avgCurve, compareData));
            return;
        }

        delete el.dataset.jtlRenderAttempt;
        el.innerHTML = '';

        const W = Math.max(rawWidth || el.clientWidth || 0, 900);
        const H = CHART_HEIGHT;
        const m = CHART_MARGIN;

        const innerW = W - m.left - m.right;
        const innerH = H - m.top  - m.bottom;

        // Collect all data points — include comparison and average values in y-axis domain
        const allValues = series.flatMap(s => s.data.map(d => d.value));
        if (compareData) compareData.forEach(c => c.data.forEach(d => allValues.push(d.value)));
        if (avgCurve)    avgCurve.forEach(d => allValues.push(d.value));
        const maxVal = Math.max(...allValues, 1);

        const xScale = d3.scaleTime()
            .domain(xDomain)
            .range([0, innerW]);

        const yScale = d3.scaleLinear()
            .domain([0, Math.ceil(maxVal * 1.15)])
            .range([innerH, 0]);

        const svg = d3.select(el)
            .append('svg')
            .attr('width', '100%')
            .attr('height', H)
            .attr('viewBox', `0 0 ${W} ${H}`)
            .attr('preserveAspectRatio', 'xMidYMid meet');

        const g = svg.append('g')
            .attr('transform', `translate(${m.left},${m.top})`);

        this._renderCriticalDayBands(g, innerW, innerH, xDomain);
        this._renderPeakProtectionBands(g, innerW, innerH, xDomain);
        this._renderNoItsmBand(g, innerW, innerH, xDomain);

        // Grid lines
        const gridStroke = document.documentElement?.dataset?.theme === 'dark'
            ? 'rgba(255,255,255,0.10)'
            : 'rgba(0,0,0,0.16)';

        g.append('g')
            .attr('class', 'jtl-grid')
            .call(
                d3.axisLeft(yScale)
                    .ticks(4)
                    .tickSize(-innerW)
                    .tickFormat('')
            )
            .call(sel => sel.select('.domain').remove())
            .selectAll('line')
            .attr('stroke', gridStroke)
            .attr('stroke-dasharray', '3,3');

        // X axis — fewer ticks on narrow charts (mobile) to avoid overlap
        const tickInterval = W < 500 ? d3.timeDay.every(5) : W < 700 ? d3.timeDay.every(3) : d3.timeDay.every(2);
        g.append('g')
            .attr('transform', `translate(0,${innerH})`)
            .call(
                d3.axisBottom(xScale)
                    .ticks(tickInterval)
                    .tickFormat(d => {
                        const day = d.getDate();
                        const mon = d.toLocaleDateString('en-US', { month: 'short' });
                        return day === 1 ? `${mon} ${day}` : `${mon} ${day}`;
                    })
            )
            .call(sel => sel.select('.domain').attr('stroke', '#ddd'))
            .selectAll('text')
            .attr('font-size', '10px')
            .attr('fill', '#888');

        // Y axis
        g.append('g')
            .call(d3.axisLeft(yScale).ticks(4))
            .call(sel => sel.select('.domain').remove())
            .selectAll('text')
            .attr('font-size', '10px')
            .attr('fill', '#888');

        // Today line
        if (todayFmt) {
            const [ty, tm, td] = todayFmt.split('-').map(Number);
            const todayDate = new Date(ty, tm - 1, td, 0, 0, 0);
            const tx = xScale(todayDate);
            g.append('line')
                .attr('x1', tx).attr('x2', tx)
                .attr('y1', -m.top + 4).attr('y2', innerH)
                .attr('stroke', '#3b82f6')
                .attr('stroke-width', 1.5)
                .attr('stroke-dasharray', '4,3');

            g.append('rect')
                .attr('x', tx - 20).attr('y', -m.top)
                .attr('width', 40).attr('height', 16)
                .attr('rx', 3)
                .attr('fill', '#3b82f6');

            g.append('text')
                .attr('x', tx).attr('y', -m.top + 11)
                .attr('text-anchor', 'middle')
                .attr('font-size', '10px')
                .attr('font-weight', '600')
                .attr('fill', '#fff')
                .text('Today');
        }

        // Lines + dots per series
        const line = d3.line()
            .x(d => xScale(d.date))
            .y(d => yScale(d.value))
            .curve(d3.curveMonotoneX);

        const isDark    = document.documentElement?.dataset?.theme === 'dark';
        const triFill   = '#f5c400';
        const triStroke = isDark ? '#ffffff' : '#111111';

        // day-of-month → [{date, label, color, value}] for collision picker
        const dayEntries = new Map();
        series.forEach(s => {
            s.data.forEach(d => {
                if (d.value === 0) return;
                const day = d.date.getDate();
                const monthLabel = d.date.toLocaleDateString('en-US', { month: 'short', year: 'numeric' }) + ' (current)';
                if (!dayEntries.has(day)) dayEntries.set(day, []);
                dayEntries.get(day).push({ date: d.date, label: monthLabel, color: s.color, value: d.value, hasCritical: !!d.hasCritical });
            });
        });
        if (compareData) {
            compareData.forEach(cmp => {
                cmp.data.forEach(d => {
                    if (d.value === 0) return;
                    const day = d.date.getDate();
                    const actualDate = new Date(cmp.srcYear, cmp.srcMonth, day);
                    if (!dayEntries.has(day)) dayEntries.set(day, []);
                    dayEntries.get(day).push({ date: actualDate, label: cmp.label, color: cmp.color, value: d.value, hasCritical: !!d.hasCritical });
                });
            });
        }

        series.forEach(s => {
            if (!s.data.length) return;

            const sg = g.append('g').attr('class', 'jtl-series-group');

            // Wide transparent hit zone — first child so it renders behind visual elements
            sg.append('path')
                .datum(s.data)
                .attr('fill', 'none')
                .attr('stroke', 'transparent')
                .attr('stroke-width', 20)
                .style('pointer-events', 'stroke')
                .attr('d', line);

            sg.append('path')
                .datum(s.data)
                .attr('fill', 'none')
                .attr('stroke', s.color)
                .attr('stroke-width', 2)
                .attr('d', line);

            // Dots — normal days: circles; P1/P2 days: red triangles
            const normalData   = s.data.filter(d => !d.hasCritical);
            const criticalData = s.data.filter(d => d.hasCritical);
            const r = dotClickLabel ? 5 : 3.5;

            const dots = sg.selectAll(null)
                .data(normalData)
                .join('circle')
                .attr('cx', d => xScale(d.date))
                .attr('cy', d => yScale(d.value))
                .attr('r', r)
                .attr('fill', s.color)
                .attr('stroke', '#fff')
                .attr('stroke-width', 1.5);

            const triSize   = dotClickLabel ? 400 : 260;
            const triPath   = d3.symbol().type(d3.symbolTriangle).size(triSize)();
            const critDots = sg.selectAll(null)
                .data(criticalData)
                .join('path')
                .attr('d', triPath)
                .attr('transform', d => `translate(${xScale(d.date)},${yScale(d.value)})`)
                .attr('fill', triFill)
                .attr('stroke', triStroke)
                .attr('stroke-width', 2);

            // "!" label inside each critical triangle
            sg.selectAll(null)
                .data(criticalData)
                .join('text')
                .attr('x', d => xScale(d.date))
                .attr('y', d => yScale(d.value) + 2)
                .attr('text-anchor', 'middle')
                .attr('dominant-baseline', 'middle')
                .attr('fill', '#111')
                .attr('font-size', dotClickLabel ? '9px' : '7px')
                .attr('font-weight', '900')
                .style('pointer-events', 'none')
                .text('!');

            if (dotClickLabel && this._onDotClick) {
                const clickFn = (event, d) => {
                    event.stopPropagation();
                    if (d.value === 0) return;
                    const entries = dayEntries.get(d.date.getDate()) || [];
                    if (entries.length > 1) {
                        this._showDayMonthPicker(event, entries, dotClickLabel);
                    } else {
                        this._onDotClick(d.date, dotClickLabel);
                    }
                };
                dots.style('cursor', 'pointer').on('click', clickFn);
                critDots.style('cursor', 'pointer').on('click', clickFn);
            }

            // Value labels (every other point to avoid crowding)
            sg.selectAll(null)
                .data(s.data.filter((_, i) => i % 2 === 0))
                .join('text')
                .attr('x', d => xScale(d.date))
                .attr('y', d => yScale(d.value) - (d.hasCritical ? 32 : 8))
                .attr('text-anchor', 'middle')
                .attr('font-size', '9px')
                .attr('fill', s.color)
                .attr('font-weight', '600')
                .text(d => d.value || '');
        });

        // ── Comparison overlay lines (solid, behind avg curve) ────────────────
        if (compareData && compareData.length) {
            const cmpLine = d3.line()
                .x(d => xScale(d.date))
                .y(d => yScale(d.value))
                .curve(d3.curveMonotoneX);

            compareData.forEach(cmp => {
                if (!cmp.data.length) return;
                const cg = g.append('g').attr('class', 'jtl-cmp-group');

                cg.append('path')
                    .datum(cmp.data)
                    .attr('fill', 'none')
                    .attr('stroke', 'transparent')
                    .attr('stroke-width', 20)
                    .style('pointer-events', 'stroke')
                    .attr('d', cmpLine);

                cg.append('path')
                    .datum(cmp.data)
                    .attr('fill', 'none')
                    .attr('stroke', cmp.color)
                    .attr('stroke-width', 1.5)
                    .attr('stroke-opacity', 0.65)
                    .attr('class', 'jtl-cmp-path')
                    .style('pointer-events', 'none')
                    .attr('d', cmpLine);

                // Dots on non-zero days: circles for normal, small amber triangles for P1/P2
                const cmpNormal   = cmp.data.filter(d => d.value > 0 && !d.hasCritical);
                const cmpCritical = cmp.data.filter(d => d.value > 0 &&  d.hasCritical);

                const cmpDots = cg.selectAll(null)
                    .data(cmpNormal)
                    .join('circle')
                    .attr('cx', d => xScale(d.date))
                    .attr('cy', d => yScale(d.value))
                    .attr('r', 3)
                    .attr('fill', cmp.color)
                    .attr('fill-opacity', 0.7)
                    .attr('stroke', '#fff')
                    .attr('stroke-width', 1);

                // Comparison triangles — 50% of main triangle size
                const cmpTriSize = dotClickLabel ? 200 : 130;
                const cmpTriPath = d3.symbol().type(d3.symbolTriangle).size(cmpTriSize)();
                const cmpCritDots = cg.selectAll(null)
                    .data(cmpCritical)
                    .join('path')
                    .attr('d', cmpTriPath)
                    .attr('transform', d => `translate(${xScale(d.date)},${yScale(d.value)})`)
                    .attr('fill', triFill)
                    .attr('fill-opacity', 0.20)
                    .attr('stroke', triStroke)
                    .attr('stroke-width', 1.5);

                cmpCritDots
                    .on('mouseenter', function() { d3.select(this).attr('fill-opacity', 1); })
                    .on('mouseleave', function() { d3.select(this).attr('fill-opacity', 0.20); });

                cg.selectAll(null)
                    .data(cmpCritical)
                    .join('text')
                    .attr('x', d => xScale(d.date))
                    .attr('y', d => yScale(d.value) + 1.5)
                    .attr('text-anchor', 'middle')
                    .attr('dominant-baseline', 'middle')
                    .attr('fill', '#111')
                    .attr('font-size', dotClickLabel ? '6px' : '5px')
                    .attr('font-weight', '900')
                    .style('pointer-events', 'none')
                    .text('!');

                if (dotClickLabel && this._onDotClick) {
                    const cmpClickFn = (event, d) => {
                        event.stopPropagation();
                        const day = d.date.getDate();
                        const actualDate = new Date(cmp.srcYear, cmp.srcMonth, day);
                        const entries = dayEntries.get(day) || [];
                        if (entries.length > 1) {
                            this._showDayMonthPicker(event, entries, dotClickLabel);
                        } else {
                            this._onDotClick(actualDate, dotClickLabel);
                        }
                    };
                    cmpDots.style('cursor', 'pointer').on('click', cmpClickFn);
                    cmpCritDots.style('cursor', 'pointer').on('click', cmpClickFn);
                } else {
                    cmpDots.style('pointer-events', 'none');
                    cmpCritDots.style('pointer-events', 'none');
                }
            });
        }

        // ── Line hover-highlight (raise hovered series, dim others) ──────────
        const allSeriesGroups = g.selectAll('.jtl-series-group, .jtl-cmp-group');
        allSeriesGroups
            .on('mouseenter.highlight', function() {
                allSeriesGroups.style('opacity', 0.2);
                d3.select(this).style('opacity', 1).raise();
            })
            .on('mouseleave.highlight', function() {
                allSeriesGroups.style('opacity', null);
            });

        // ── Average curve (dashed, on top) ────────────────────────────────────
        if (avgCurve && avgCurve.length) {
            const avgLine = d3.line()
                .x(d => xScale(d.date))
                .y(d => yScale(d.value))
                .curve(d3.curveMonotoneX);

            g.append('path')
                .datum(avgCurve)
                .attr('fill', 'none')
                .attr('stroke', '#9ca3af')
                .attr('stroke-width', 1.5)
                .attr('stroke-dasharray', '5,4')
                .attr('class', 'jtl-avg-path')
                .style('pointer-events', 'none')
                .attr('d', avgLine);

            // Small end label
            const last = avgCurve[avgCurve.length - 1];
            if (last) {
                g.append('text')
                    .attr('x', Math.min(xScale(last.date) + 4, innerW - 24))
                    .attr('y', yScale(last.value) - 4)
                    .attr('font-size', '9px')
                    .attr('fill', '#9ca3af')
                    .style('pointer-events', 'none')
                    .text('avg');
            }
        }
    }

    // ── Day / month collision picker ──────────────────────────────────────────

    _showDayMonthPicker(event, entries, dotClickLabel) {
        document.getElementById('jtl-month-picker')?.remove();

        const picker = document.createElement('div');
        picker.id = 'jtl-month-picker';
        picker.className = 'jtl-month-picker';
        picker.style.left = `${event.clientX + 10}px`;
        picker.style.top  = `${event.clientY - 10}px`;

        entries.forEach(entry => {
            const btn = document.createElement('button');
            btn.className = 'jtl-month-picker__item';
            btn.innerHTML =
                `<span class="jtl-month-picker__dot" style="background:${entry.color}"></span>` +
                `<span class="jtl-month-picker__label">${entry.label}</span>` +
                (entry.hasCritical ? `<span class="jtl-month-picker__crit" title="Has P1/P2 incidents">▲</span>` : '') +
                `<span class="jtl-month-picker__count">${entry.value}</span>`;
            btn.addEventListener('click', e => {
                e.stopPropagation();
                picker.remove();
                this._onDotClick(entry.date, dotClickLabel);
            });
            picker.appendChild(btn);
        });

        document.body.appendChild(picker);

        // Flip position if near viewport edge
        requestAnimationFrame(() => {
            const rect = picker.getBoundingClientRect();
            if (rect.right  > window.innerWidth  - 8) picker.style.left = `${event.clientX - rect.width  - 10}px`;
            if (rect.bottom > window.innerHeight - 8) picker.style.top  = `${event.clientY - rect.height + 10}px`;
        });

        const close = e => {
            if (!picker.contains(e.target)) {
                picker.remove();
                document.removeEventListener('click', close, true);
            }
        };
        setTimeout(() => document.addEventListener('click', close, true), 0);
    }

    // ── Inline insights panel ─────────────────────────────────────────────────

    _buildInsightsPanel({ incidentData, srData, releaseData, days, todayFmt, issueSummaryQuery = '' }) {
        const wrap = document.createElement('div');
        wrap.className = 'jtl-insights-panel';

        // Timeline-only search box for Jira cards.
        // It intentionally filters only Incident Created and Service Request Created series,
        // using card.summary as requested.
        wrap.appendChild(this._buildIssueSearchInput(issueSummaryQuery));

        const heading = document.createElement('div');
        heading.className = 'jtl-insights-panel__heading';
        heading.textContent = 'INSIGHTS';
        wrap.appendChild(heading);

        const insights = this._computeInsights({ incidentData, srData, releaseData, days, todayFmt });

        if (!insights.length) {
            const empty = document.createElement('p');
            empty.className = 'jtl-insights-panel__empty';
            empty.textContent = 'No notable patterns this month.';
            wrap.appendChild(empty);
            return wrap;
        }

        const list = document.createElement('div');
        list.className = 'jtl-insights-panel__list';
        insights.forEach(ins => {
            const item = document.createElement('div');
            item.className = 'jtl-insights-panel__item';
            item.innerHTML = `
                <span class="jtl-insights-panel__icon">${ins.icon}</span>
                <strong class="jtl-insights-panel__title">${ins.title}</strong>
                <span class="jtl-insights-panel__sub">${ins.sub} — ${ins.desc}</span>`;
            list.appendChild(item);
        });
        wrap.appendChild(list);
        return wrap;
    }

    _buildIssueSearchInput(value = '') {
        const box = document.createElement('div');
        box.className = 'jtl-issue-search';

        const icon = document.createElement('button');
        icon.type = 'button';
        icon.className = 'jtl-issue-search__icon';
        icon.tabIndex = -1;
        icon.setAttribute('aria-hidden', 'true');
        icon.textContent = '🔎';

        const input = document.createElement('input');
        input.id = 'jtl-issue-search-input';
        input.className = 'jtl-issue-search__input';
        input.type = 'search';
        input.placeholder = 'Search Issues';
        input.autocomplete = 'off';
        input.value = value || '';

        const clear = document.createElement('button');
        clear.type = 'button';
        clear.className = 'jtl-issue-search__clear';
        clear.title = 'Clear issue search';
        clear.setAttribute('aria-label', 'Clear issue search');
        clear.textContent = '✕';
        clear.style.display = input.value ? '' : 'none';

        const apply = () => {
            const next = input.value.trim();
            if (next === this.issueSummaryQuery) return;
            this.issueSummaryQuery = next;
            this._restoreIssueSearchFocus = true;
            this.app.refresh();
        };

        input.addEventListener('input', () => {
            clear.style.display = input.value.trim() ? '' : 'none';
            apply();
        });

        input.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                input.value = '';
                clear.style.display = 'none';
                apply();
            }
        });

        clear.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            input.value = '';
            clear.style.display = 'none';
            apply();
        });

        box.append(icon, input, clear);

        if (this._restoreIssueSearchFocus) {
            this._restoreIssueSearchFocus = false;
            requestAnimationFrame(() => {
                const liveInput = document.getElementById('jtl-issue-search-input');
                if (!liveInput) return;
                liveInput.focus();
                const len = liveInput.value.length;
                try { liveInput.setSelectionRange(len, len); } catch (_) {}
            });
        }

        return box;
    }

    _computeInsights({ incidentData, srData, releaseData, days }) {
        const insights = [];

        // Find incident spike: day with > 1.5× average
        if (incidentData.length) {
            const avg = incidentData.reduce((s, d) => s + d.value, 0) / incidentData.length;
            const spike = incidentData.reduce((best, d) => d.value > best.value ? d : best, incidentData[0]);
            if (spike.value > avg * 1.5) {
                const fmt = spike.date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
                insights.push({
                    icon: '📈',
                    title: 'Incident spike',
                    sub: fmt,
                    desc: `${spike.value} incidents (${Math.round((spike.value / avg - 1) * 100)}% above average)`,
                });
            }
        }

        // Find busiest release day
        const prdSeries = releaseData.find(s => s.key === 'prd');
        if (prdSeries && prdSeries.data.length) {
            const peak = prdSeries.data.reduce((best, d) => d.value > best.value ? d : best, prdSeries.data[0]);
            if (peak.value > 0) {
                const fmt = peak.date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
                insights.push({
                    icon: '🚀',
                    title: 'High release activity',
                    sub: fmt,
                    desc: `${peak.value} PROD releases on peak day`,
                });
            }
        }

        // Find SR spike
        if (srData.length) {
            const peak = srData.reduce((best, d) => d.value > best.value ? d : best, srData[0]);
            if (peak.value > 0) {
                const fmt = peak.date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
                insights.push({
                    icon: '💬',
                    title: 'Service Requests spike',
                    sub: fmt,
                    desc: `${peak.value} SRs created`,
                });
            }
        }

        return insights;
    }

    // ── Data helpers ──────────────────────────────────────────────────────────

    _buildDays(year, month, daysInMonth) {
        const days = [];
        for (let d = 1; d <= daysInMonth; d++) {
            // Use local midnight so localDateKey(date) always matches the CSV
            // date strings (parseDate now returns local-midnight dates).
            days.push(new Date(year, month, d));
        }
        return days;
    }

    /** Build release counts per env series per day. */
    _buildReleaseData(events, days) {
        const releaseTypes = new Set(['DELIVERY', 'HYBRIS', 'SERVICE_OP']);
        const releaseEvents = events.filter(e => releaseTypes.has(e.type));

        return ENV_SERIES.map(({ key, label, color }) => {
            const envKeys = ENV_KEYS[key] || [key];
            const data = days.map(date => {
                const dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate());
                const dayEnd   = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59);
                const count = releaseEvents.filter(e => {
                    const env = (e.environment || '').toLowerCase();
                    if (!envKeys.some(k => env === k || env.startsWith(k))) return false;
                    // Env/service filtering is applied upstream by JengaSearch
                    const due = e.dueDate;
                    if (!due) return false;
                    return due >= dayStart && due <= dayEnd;
                }).length;
                return { date: dayStart, value: count };
            });
            return { key, label, color, data };
        });
    }

    /** Build card counts per day (incidents or service requests). */
    _buildCardData(store, year, month, issueTypeFragment, days, summaryQuery = '') {
        const activeServices = this.app.search?.activeServices;
        const services = activeServices && activeServices.size > 0 ? activeServices : undefined;
        const q = (summaryQuery || '').trim().toLowerCase();

        // Fast path: keep the existing EventStore aggregation when the issue search is empty.
        if (!q) {
            const counts = store.getCardCountsForMonth(year, month, {
                issueType: issueTypeFragment,
                services,
            });
            const criticalDays = store.getCriticalDaysForMonth(year, month, {
                issueType: issueTypeFragment,
                services,
            });
            return days.map(date => ({
                date,
                value: counts.get(localDateKey(date)) || 0,
                hasCritical: criticalDays.has(localDateKey(date)),
            }));
        }

        // Search Issues filters ONLY on Jira card Summary, and only affects the
        // Incidents Created / Service Requests Created timeline series.
        const monthStart = new Date(year, month, 1);
        const monthEnd   = new Date(year, month + 1, 0, 23, 59, 59);
        const counts = new Map();
        const criticalDays = new Set();

        (store.cards || []).forEach(c => {
            if (!c.created) return;
            if (c.created < monthStart || c.created > monthEnd) return;
            if (issueTypeFragment && !(c.issueType || '').toLowerCase().includes(issueTypeFragment.toLowerCase())) return;
            if (!(c.summary || '').toLowerCase().includes(q)) return;

            if (services && services.size > 0) {
                const cardServices = (c.affectedServices || '')
                    .split('||')
                    .map(s => s.trim())
                    .filter(Boolean);
                if (!cardServices.some(svc => services.has(svc))) return;
            }

            const key = localDateKey(c.created);
            counts.set(key, (counts.get(key) || 0) + 1);
            const p = (c.priority || '').trim().toUpperCase();
            if (p.startsWith('P1') || p.startsWith('P2')) criticalDays.add(key);
        });

        return days.map(date => ({
            date,
            value: counts.get(localDateKey(date)) || 0,
            hasCritical: criticalDays.has(localDateKey(date)),
        }));
    }

    /**
     * Convert a day number (1-based) to a percentage position matching D3's scaleTime
     * with domain [new Date(year,month,1), new Date(year,month,daysInMonth)].
     * Day 1 → 0%, day daysInMonth → 100%.
     */
    _dayToPct(day, daysInMonth) {
        if (daysInMonth <= 1) return 0;
        return ((day - 1) / (daysInMonth - 1)) * 100;
    }

    /** CSS percentage for elements already scoped to the D3 plot area. */
    _dotPctCss(day, daysInMonth) {
        return `${this._dayToPct(day, daysInMonth)}%`;
    }

    /** Inclusive full-day span width when the left edge is aligned to D3 day dots. */
    _daySpanFromDotPct(startDay, endDay, daysInMonth) {
        if (daysInMonth <= 1) return 100;
        const leftPct = ((startDay - 1) / (daysInMonth - 1)) * 100;
        const rawWidth = ((endDay - startDay + 1) / (daysInMonth - 1)) * 100;
        return Math.min(leftPct + rawWidth, 100) - leftPct;
    }


    _peakProtectionDayMetrics(ev, daysInMonth, innerW, xDomain) {
        const rawStart = ev.startDate || ev.dueDate;
        const rawEnd   = ev.dueDate   || ev.startDate;
        if (!rawStart || !rawEnd) return null;

        const evStart = rawStart <= rawEnd ? rawStart : rawEnd;
        const evEnd   = rawStart <= rawEnd ? rawEnd   : rawStart;

        // xDomain is the currently rendered month. A PSPW phase can start in the
        // previous month or end in the next one: for timeline background bands we
        // must render only the visible overlap with the current month.
        const domainStart = xDomain?.[0];
        const domainEnd   = xDomain?.[1];
        const monthStart = domainStart instanceof Date
            ? new Date(domainStart.getFullYear(), domainStart.getMonth(), 1)
            : new Date(evStart.getFullYear(), evStart.getMonth(), 1);
        const monthEnd = domainEnd instanceof Date
            ? new Date(domainEnd.getFullYear(), domainEnd.getMonth(), daysInMonth, 23, 59, 59)
            : new Date(evEnd.getFullYear(), evEnd.getMonth(), daysInMonth, 23, 59, 59);

        const visibleStart = evStart < monthStart ? monthStart : evStart;
        const visibleEnd   = evEnd   > monthEnd   ? monthEnd   : evEnd;
        if (visibleEnd < monthStart || visibleStart > monthEnd) return null;

        const startDay = Math.max(1, Math.min(daysInMonth, visibleStart.getDate()));
        const endDay   = Math.max(1, Math.min(daysInMonth, visibleEnd.getDate()));
        const dayStep  = daysInMonth <= 1 ? innerW : innerW / (daysInMonth - 1);

        return {
            x: ((startDay - 1) / Math.max(daysInMonth - 1, 1)) * innerW,
            width: Math.max((endDay - startDay + 1) * dayStep, dayStep),
        };
    }

    _renderCriticalDayBands(g, innerW, innerH, xDomain) {
        const criticalDays = this._flameDays;
        if (!criticalDays || !criticalDays.size) return;

        const domainEnd = xDomain?.[1];
        const daysInMonth = domainEnd instanceof Date ? domainEnd.getDate() : 31;
        const bandOpacity = document.documentElement?.dataset?.theme === 'dark' ? 0.18 : 0.13;
        const dayStep = daysInMonth <= 1 ? innerW : innerW / (daysInMonth - 1);

        const bandLayer = g.append('g').attr('class', 'jtl-critical-day-band-layer');

        criticalDays.forEach(dateKey => {
            const d = new Date(dateKey + 'T00:00:00');
            if (d.getMonth() !== (xDomain?.[0]?.getMonth?.() ?? d.getMonth())) return;
            const day = d.getDate();
            const x = ((day - 1) / Math.max(daysInMonth - 1, 1)) * innerW - dayStep * 0.5;
            bandLayer.append('rect')
                .attr('class', 'jtl-critical-day-band')
                .attr('x', x)
                .attr('y', -CHART_MARGIN.top)
                .attr('width', dayStep)
                .attr('height', innerH + CHART_MARGIN.top + CHART_MARGIN.bottom)
                .attr('fill', '#f97316')
                .attr('fill-opacity', bandOpacity)
                .attr('rx', 2);
        });
    }

    _renderPeakProtectionBands(g, innerW, innerH, xDomain) {
        const events = (this._peakProtectionEvents || []).filter(ev => !isCabProtectionEvent(ev));
        if (!events.length) return;

        const domainEnd = xDomain?.[1];
        const daysInMonth = domainEnd instanceof Date ? domainEnd.getDate() : 31;
        const bandOpacity = document.documentElement?.dataset?.theme === 'dark' ? 0.055 : 0.08;

        const bandLayer = g.append('g')
            .attr('class', 'jtl-peak-protection-band-layer');

        events.forEach(ev => {
            const metrics = this._peakProtectionDayMetrics(ev, daysInMonth, innerW, xDomain);
            if (!metrics) return;

            const colorKey = ev.protectionColor || 'AMBER';
            const color = PEAK_PROTECTION_COLORS[colorKey] || PEAK_PROTECTION_COLORS.AMBER;

            bandLayer.append('rect')
                .attr('class', `jtl-peak-protection-band jtl-peak-protection-band--${colorKey.toLowerCase()}`)
                .attr('x', metrics.x)
                .attr('y', -CHART_MARGIN.top)
                .attr('width', metrics.width)
                .attr('height', innerH + CHART_MARGIN.top + CHART_MARGIN.bottom)
                .attr('fill', color)
                .attr('fill-opacity', bandOpacity)
                .attr('rx', 2);
        });
    }

    /**
     * Build a day-number axis ruler that aligns with the D3 plot area.
     * @param {number} year
     * @param {number} month  0-based
     * @param {HTMLElement} main  — the .jtl-main element (used for rAF measurement)
     * @returns {HTMLElement}  the ruler div (ss-day-axis class, positioning set in rAF)
     */
    _buildDayAxisRuler(year, month, main) {
        const daysInMonth = new Date(year, month + 1, 0).getDate();
        const ruler = document.createElement('div');
        ruler.className = 'ss-day-axis';

        for (let d = 1; d <= daysInMonth; d++) {
            const span = document.createElement('span');
            span.className = 'ss-day-axis__day';
            span.textContent = d;
            const pct = daysInMonth > 1 ? ((d - 1) / (daysInMonth - 1)) * 100 : 0;
            span.style.left = `${pct}%`;
            ruler.appendChild(span);
        }

        // After layout settles: offset ruler so its left edge aligns with the D3 plot area,
        // accounting for the chart element's own padding plus the D3 left margin.
        // Skip chart elements inside collapsed lanes — they have zero width and would
        // produce wrong measurements; fall back to any visible chart in the main container.
        requestAnimationFrame(() => {
            const chartEl = main && Array.from(
                main.querySelectorAll('.jtl-panel__chart, .jtl-panel__chart--biz')
            ).find(el => !el.closest('.jtl-lane--collapsed'));
            if (!chartEl) return;
            const chartRect = chartEl.getBoundingClientRect();
            const mainRect  = main.getBoundingClientRect();
            const cs        = getComputedStyle(chartEl);
            const padLeft   = parseFloat(cs.paddingLeft)  || 0;
            const padRight  = parseFloat(cs.paddingRight) || 0;
            const plotLeft  = chartRect.left - mainRect.left + padLeft + CHART_MARGIN.left;
            const plotWidth = chartRect.width - padLeft - padRight - CHART_MARGIN.left - CHART_MARGIN.right;
            ruler.style.marginLeft = `${plotLeft}px`;
            ruler.style.width      = `${plotWidth}px`;
        });

        return ruler;
    }

    /** Make an HTML bars layer use the same horizontal plot area as the D3 charts. */
    _alignBarsToD3PlotArea(bars) {
        if (!bars) return;
        bars.style.marginLeft = `${CHART_MARGIN.left}px`;
        bars.style.width = `calc(100% - ${CHART_MARGIN.left + CHART_MARGIN.right}px)`;
    }

    /** CSS `left` value for the today-line in biz/milestone bar rows.
     *  Accounts for the D3 chart left margin so the line aligns with
     *  the D3 today marker in the line charts below. */


    _cabMarkerLeftCss(day, daysInMonth) {
        // CAB markers live inside the same bars layer as PSPW bars.
        // The bars layer is aligned to the D3 plot area, so use the pure D3 dot percentage.
        if (typeof this._dotPctCss === 'function') return this._dotPctCss(day, daysInMonth);

        // Fallback for older builds where _dotPctCss is not present yet.
        if (typeof this._dayToPct === 'function') return `${this._dayToPct(day, daysInMonth)}%`;

        // Last-resort fallback: avoid runtime failure and roughly position the marker.
        if (daysInMonth <= 1) return '0%';
        return `${((day - 1) / (daysInMonth - 1)) * 100}%`;
    }

    _todayLineCss(day, daysInMonth) {
        const pct = this._dayToPct(day, daysInMonth);
        const ml  = CHART_MARGIN.left;
        const mr  = CHART_MARGIN.right;
        return `calc(${ml}px + ${pct} * (100% - ${ml}px - ${mr}px) / 100)`;
    }

    /**
     * Append a hatched grey HTML overlay to a bar-row `bars` div for the portion
     * that predates the ITSM retention window.  Used for biz-events, milestones,
     * peak-protection, and releases rows (which are HTML, not SVG).
     */
    _appendNoItsmBarOverlay(bars, year, month, daysInMonth) {
        const cutoff = itsmCutoffDate();
        const monthStart = new Date(year, month, 1);
        const monthEnd   = new Date(year, month, daysInMonth, 23, 59, 59);

        if (monthStart >= cutoff) return; // entire month is within retention

        const isDark = document.documentElement?.dataset?.theme === 'dark';

        // Width of the greyed region as a percentage of the bars area.
        // If the cutoff is beyond month-end, grey the whole row.
        let widthPct;
        if (cutoff >= monthEnd) {
            widthPct = 100;
        } else {
            // Align with _dayToPct: cutoffDay is the day-of-month at local midnight.
            const cutoffDay = cutoff.getDate();
            widthPct = this._dayToPct(cutoffDay, daysInMonth);
        }

        const overlay = document.createElement('div');
        overlay.className = 'jtl-no-itsm-bar-overlay';
        overlay.style.cssText = [
            'position:absolute',
            'top:0',
            'left:0',
            `width:${widthPct}%`,
            'height:100%',
            'pointer-events:auto',
            'z-index:2',
            `background:${isDark ? 'rgba(40,40,46,0.55)' : 'rgba(200,200,210,0.28)'}`,
            `background-image:repeating-linear-gradient(135deg,transparent,transparent 6px,${isDark ? 'rgba(255,255,255,0.04)' : 'rgba(0,0,0,0.04)'} 6px,${isDark ? 'rgba(255,255,255,0.04)' : 'rgba(0,0,0,0.04)'} 12px)`,
            'border-radius:2px',
        ].join(';');

        const tooltip = `Data unavailable — ITSM records are retained for ${ITSM_RETENTION_DAYS} days`;
        overlay.addEventListener('mousemove', (e) => this._showHoverCard(e.clientX, e.clientY, [
            { label: 'No data', value: tooltip },
        ]));
        overlay.addEventListener('mouseleave', () => this._hideHoverCard());

        bars.appendChild(overlay);

        // Dashed boundary line at the cutoff edge (only when inside the month)
        if (cutoff > monthStart && cutoff <= monthEnd) {
            const line = document.createElement('div');
            line.style.cssText = [
                'position:absolute',
                'top:0',
                'bottom:0',
                `left:${widthPct}%`,
                'width:0',
                'pointer-events:none',
                'z-index:3',
                `border-left:1px dashed ${isDark ? 'rgba(255,255,255,0.25)' : 'rgba(0,0,0,0.20)'}`,
            ].join(';');
            bars.appendChild(line);
        }
    }

    /**
     * Render a hatched grey overlay over the portion of the chart that predates
     * the ITSM data-retention window.  Called for all D3 line-chart lanes.
     */
    _renderNoItsmBand(g, innerW, innerH, xDomain) {
        const cutoff = itsmCutoffDate();
        const monthStart = xDomain?.[0];
        const monthEnd   = xDomain?.[1];
        if (!monthStart || !monthEnd) return;

        // If the entire visible month is within retention, nothing to grey out.
        if (monthStart >= cutoff) return;

        const isDark = document.documentElement?.dataset?.theme === 'dark';

        // Pixel position of the cutoff (or right edge if it falls after month end)
        const xScale = d3.scaleTime().domain(xDomain).range([0, innerW]);
        const cutoffX = cutoff >= monthEnd
            ? innerW
            : Math.max(0, Math.min(innerW, xScale(cutoff)));

        const bandLayer = g.append('g').attr('class', 'jtl-no-itsm-band-layer');

        // Hatched-fill definition
        const patternId = 'jtl-no-itsm-hatch';
        if (!g.select(`#${patternId}`).node()) {
            const defs = g.append('defs');
            const pattern = defs.append('pattern')
                .attr('id', patternId)
                .attr('patternUnits', 'userSpaceOnUse')
                .attr('width', 8).attr('height', 8)
                .attr('patternTransform', 'rotate(45)');
            pattern.append('line')
                .attr('x1', 0).attr('y1', 0)
                .attr('x2', 0).attr('y2', 8)
                .attr('stroke', isDark ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.07)')
                .attr('stroke-width', 3);
        }

        // Grey fill rect
        bandLayer.append('rect')
            .attr('x', -CHART_MARGIN.left)
            .attr('y', -CHART_MARGIN.top)
            .attr('width', cutoffX + CHART_MARGIN.left)
            .attr('height', innerH + CHART_MARGIN.top + CHART_MARGIN.bottom)
            .attr('fill', isDark ? 'rgba(40,40,46,0.65)' : 'rgba(200,200,210,0.30)');

        // Hatch overlay
        bandLayer.append('rect')
            .attr('x', -CHART_MARGIN.left)
            .attr('y', -CHART_MARGIN.top)
            .attr('width', cutoffX + CHART_MARGIN.left)
            .attr('height', innerH + CHART_MARGIN.top + CHART_MARGIN.bottom)
            .attr('fill', `url(#${patternId})`);

        // Dashed cutoff boundary line (only when cutoff falls inside the month)
        if (cutoff > monthStart && cutoff < monthEnd) {
            bandLayer.append('line')
                .attr('x1', cutoffX).attr('x2', cutoffX)
                .attr('y1', -CHART_MARGIN.top + 4).attr('y2', innerH)
                .attr('stroke', isDark ? 'rgba(255,255,255,0.30)' : 'rgba(0,0,0,0.25)')
                .attr('stroke-width', 1)
                .attr('stroke-dasharray', '4,3');
        }

        // Invisible hover rect over the grey area for the tooltip
        const tooltip = `Data unavailable — ITSM records are retained for ${ITSM_RETENTION_DAYS} days`;
        bandLayer.append('rect')
            .attr('class', 'jtl-no-itsm-hit')
            .attr('x', -CHART_MARGIN.left)
            .attr('y', -CHART_MARGIN.top)
            .attr('width', cutoffX + CHART_MARGIN.left)
            .attr('height', innerH + CHART_MARGIN.top + CHART_MARGIN.bottom)
            .attr('fill', 'transparent')
            .attr('cursor', 'default')
            .on('mousemove', (event) => this._showHoverCard(event.clientX, event.clientY, [
                { label: 'No data', value: tooltip },
            ]))
            .on('mouseleave', () => this._hideHoverCard());
    }

    // ── Collapsible lane wrapper ───────────────────────────────────────────────

    static get _LS_KEY() { return 'jenga-timeline-lanes-v1'; }

    _loadLaneState() {
        try { return JSON.parse(localStorage.getItem(TimelineRenderer._LS_KEY) || '{}'); }
        catch { return {}; }
    }

    _saveLaneState(state) {
        try { localStorage.setItem(TimelineRenderer._LS_KEY, JSON.stringify(state)); }
        catch {}
    }

    /**
     * Wrap a panel in a collapsible container.
     * @param {string} laneKey  — unique key for localStorage
     * @param {boolean} defaultCollapsed — initial state if not saved
     * @param {HTMLElement} panelEl — the jtl-panel element to wrap
     */
    /**
     * Legacy cleanup hook: lane pinning has been disabled.
     * It removes any stale pinned styles/classes left by previous builds.
     */
    refreshStickyLayout() {
        const root = document.getElementById('jenga-timeline') || document;
        root.querySelectorAll?.('.jtl-lane--pinned').forEach(lane => {
            lane.classList.remove('jtl-lane--pinned');
            lane.style.removeProperty('--jtl-placeholder-height');
            lane.style.removeProperty('--jtl-pinned-top');
            lane.style.removeProperty('--jtl-fixed-left');
            lane.style.removeProperty('--jtl-fixed-width');
            lane.style.removeProperty('--jtl-pinned-z');
            lane.style.minHeight = '';
        });
        try { localStorage.removeItem('jenga.timelinePinnedLanes'); } catch (_) {}
    }

    /**
     * Wrap a panel in a collapsible container.
     * Lane pinning is intentionally disabled; only expand/collapse remains.
     * @param {string} laneKey  — unique key for localStorage
     * @param {boolean} defaultCollapsed — initial state if not saved
     * @param {HTMLElement} panelEl — the jtl-panel element to wrap
     */

    _rerenderTimelineAfterLaneExpand(laneEl) {
        const main = laneEl?.closest('.jtl-main');
        const scrollLeft = main?.scrollLeft ?? 0;
        const scrollTop = window.scrollY;
        const activeLane = laneEl?.dataset?.lane || null;

        cancelAnimationFrame(this._laneExpandRaf1);
        cancelAnimationFrame(this._laneExpandRaf2);
        cancelAnimationFrame(this._laneExpandRaf3);

        // Two frames give the browser time to remove the collapsed class,
        // recalculate layout, and expose the correct chart width.
        this._laneExpandRaf1 = requestAnimationFrame(() => {
            this._laneExpandRaf2 = requestAnimationFrame(() => {
                this.app?.refresh?.();

                this._laneExpandRaf3 = requestAnimationFrame(() => {
                    const newMain = document.querySelector('#jenga-timeline .jtl-main');
                    if (newMain) newMain.scrollLeft = scrollLeft;

                    if (activeLane) {
                        document
                            .querySelector(`#jenga-timeline .jtl-lane[data-lane="${CSS.escape(activeLane)}"]`)
                            ?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
                    }

                    window.scrollTo({
                        top: scrollTop,
                        left: window.scrollX,
                        behavior: 'auto',
                    });

                    this.refreshStickyLayout?.();
                });
            });
        });
    }

    _renderVisibleD3ChartsAgain(container = document.getElementById('jenga-timeline')) {
        // Defensive helper for future use: if external code expands lanes without
        // triggering app.refresh(), this can be used to redraw visible charts.
        container
            ?.querySelectorAll('.jtl-panel__chart, .jtl-panel__chart--biz')
            ?.forEach(el => {
                if (!el.closest('.jtl-lane--collapsed')) {
                    el.dispatchEvent(new CustomEvent('jtl:request-redraw', { bubbles: true }));
                }
            });
    }

    _makeCollapsible(laneKey, defaultCollapsed, panelEl) {
        const state     = this._loadLaneState();
        const collapsed = laneKey in state ? state[laneKey] : defaultCollapsed;

        const wrapper = document.createElement('div');
        wrapper.className = 'jtl-lane';
        wrapper.dataset.lane = laneKey;

        const labelEl = panelEl.querySelector('.jtl-panel__label');
        if (labelEl) {
            const toggle = document.createElement('button');
            toggle.className = 'jtl-lane__toggle';
            toggle.setAttribute('aria-label', collapsed ? 'Expand' : 'Collapse');
            toggle.innerHTML = collapsed ? '▶' : '▼';
            labelEl.prepend(toggle);

            toggle.addEventListener('click', (e) => {
                e.stopPropagation();

                const isNowCollapsed = !wrapper.classList.contains('jtl-lane--collapsed');
                wrapper.classList.toggle('jtl-lane--collapsed', isNowCollapsed);
                toggle.innerHTML = isNowCollapsed ? '▶' : '▼';
                toggle.setAttribute('aria-label', isNowCollapsed ? 'Expand' : 'Collapse');

                const s = this._loadLaneState();
                s[laneKey] = isNowCollapsed;
                this._saveLaneState(s);

                // If we are reopening a lane, force a timeline re-render after the browser
                // has applied the expanded layout. This fixes D3 trends drawn with a
                // collapsed/hidden container width after page refresh.
                if (!isNowCollapsed) {
                    this._rerenderTimelineAfterLaneExpand(wrapper);
                }
            });
        }

        if (collapsed) wrapper.classList.add('jtl-lane--collapsed');
        wrapper.appendChild(panelEl);
        return wrapper;
    }

}
