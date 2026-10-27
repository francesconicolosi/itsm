import { BRAND, renderBrandLogo } from '../../brand-specific/brand.js';
import { applyTheme, loadSavedTheme, showThemeSwitchSpinner, isTVBrowser } from '../shared/utils.js';
import { PostItNote } from '../shared/PostItNote.js';
import { AnnouncementBar } from '../shared/AnnouncementBar.js';
import { EventStore } from './EventStore.js';
import { CalendarRenderer } from './CalendarRenderer.js';
import { EventDrawer } from './EventDrawer.js';
import { DayDrawer } from './DayDrawer.js';
import { JengaSearch } from './JengaSearch.js';
import { TimelineRenderer } from './TimelineRenderer.js';
import { CardsDrawer } from './CardsDrawer.js';
import { SlideshowController } from './SlideshowController.js';
import { MajorIncidentAlert } from './MajorIncidentAlert.js';
import { JengaStatusView } from './JengaStatusView.js';

const TOPBAR_PIN_STORAGE_KEY = 'jenga.topBarPinned';
const SPAN_STORAGE_KEY = 'jenga.spanMode';
const VALID_SPANS = new Set(['month', '1w', '2w', '3w', '4w']);
const SPAN_WEEKS  = { '1w': 1, '2w': 2, '3w': 3, '4w': 4 };

function localDateStr(date) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

export class JengaApp {
    constructor() {
        this.store      = new EventStore();
        this.calendar   = new CalendarRenderer(this);
        this.timeline   = new TimelineRenderer(this);
        this.drawer     = new EventDrawer(this);
        this.dayDrawer  = new DayDrawer();
        this.cardsDrawer = new CardsDrawer();
        this.search     = new JengaSearch(this);
        this.slideshow  = new SlideshowController(this);
        this.brandHideTypes = new Set(BRAND.statusHideTypes || []);
        this.statusView = new JengaStatusView(this);
        this.majorIncidentAlert = new MajorIncidentAlert();
        this.postIt = new PostItNote('jenga');
        this.announcementBar = new AnnouncementBar();
        this._dayDrawerOpen = false;

        this.currentYear  = new Date().getFullYear();
        this.currentMonth = new Date().getMonth();
        this.view = 'calendar'; // 'calendar' | 'timeline' | 'status'
        this.topBarPinned = this._loadTopBarPinned();
        this.spanMode      = this._loadSpanMode(); // 'month' | '1w' | '2w' | '3w' | '4w'
        this.spanStartDate = null;                 // Date (local-midnight Monday) in week mode
    }

    init() {
        this.announcementBar.init();
        // Move announcement bar inside the sticky top bar so it's always above
        // the controls and the single sticky container accounts for its height.
        const annBar = document.querySelector('.announcement-bar');
        const topBar = document.querySelector('.jenga-top-bar');
        if (annBar && topBar) topBar.insertBefore(annBar, topBar.firstChild);
        renderBrandLogo?.();
        this._applyStoredTheme();
        this._applyTopBarPinnedState();
        // Keep --jenga-topbar-height in sync whenever top bar height changes
        // (announcement bar dismissed, expanded, etc.)
        if (typeof ResizeObserver !== 'undefined') {
            new ResizeObserver(() => this._refreshTopBarPinnedGeometry()).observe(topBar || document.querySelector('.jenga-top-bar'));
        }
        this._initSideDrawer();
        this.drawer.initEvents();
        this.dayDrawer.initEvents(this.drawer);
        this.cardsDrawer.initEvents();
        this.search.init();
        this._initNavControls();
        this._initSpanSelector();
        this._initViewToggle();
        this._initSlideshow();
        this._initBuildInfo();
        this.postIt.init();
        this.postIt.attachContextMenu(document.body);
        window.addEventListener('resize', () => this._applyTopBarPinnedState());

        this._initUrlSync();

        // Sync URL whenever the user changes comparison settings
        this.timeline.onComparisonChange(() => {
            if (this._dataLoaded && !this._suppressUrlPush) this._pushUrl();
        });

        window.addEventListener('load', () => {
            const isTV = isTVBrowser() || new URLSearchParams(location.search).get('tvmode') === 'true';
            Promise.all([
                fetch(BRAND.csv.jenga).then(r => r.text()),
                fetch(isTV ? BRAND.csv.jiraCardsCurrent : BRAND.csv.jiraCards).then(r => r.text()).catch(() => ''),
                fetch(BRAND.csv.domino).then(r => r.text()).catch(() => ''),
            ]).then(([eventsCsv, cardsCsv, catalogCsv]) => {
                this.store.load(eventsCsv);
                this.store.loadCards(cardsCsv);
                const { names: serviceNames, amberAllowed } = this._parseServiceNames(catalogCsv);
                this.amberAllowedServices = amberAllowed;
                this.dayDrawer.amberAllowedServices = amberAllowed;
                this.search.setServiceOptions(serviceNames);
                this._updateLastUpdate();
                this.timeline._tvMode = isTV;
                // Apply deep-link params before first render
                this._readUrlParams();
                this._dataLoaded = true;
                this._suppressUrlPush = true;  // keep URL intact until drawer params are applied
                this.refresh();
                // Show major incident alert if there are open P1/P2 incidents,
                // but skip it when the slideshow will auto-start (status slide covers it)
                const willAutoSlideshow = isTV ||
                    new URLSearchParams(location.search).get('slideshow') === 'true';
                const majorIncidents = this.store.getMajorIncidents();
                if (majorIncidents.length > 0 && !willAutoSlideshow) {
                    this.majorIncidentAlert.show(majorIncidents);
                }
                // Open drawer/day from URL after render (needs DOM to be ready)
                requestAnimationFrame(() => {
                    this._suppressUrlPush = false;
                    this._applyUrlDrawerParams();
                    if (isTV && !this.slideshow.active &&
                            new URLSearchParams(location.search).get('slideshow') !== 'false') {
                        this.slideshow.start();
                    }
                    this._hideSpinner();
                });
            }).catch(err => {
                console.error('[Jenga] Failed to load CSV:', err);
                this._hideSpinner();
            });
        });
    }

    refresh() {
        if (this._dataLoaded && !this._suppressUrlPush) this._pushUrl();

        let allEvents;
        if (this.spanMode !== 'month' && this.spanStartDate) {
            // Week span: fetch events directly from the date range
            const spanEnd = this._spanEndDate();
            allEvents = this.store.getEventsForRange(this.spanStartDate, spanEnd);
        } else {
            const cur = this.store.getEventsForMonth(this.currentYear, this.currentMonth);

            // Also fetch adjacent months so events on other-month days in the grid are visible
            const [py, pm] = this.currentMonth === 0
                ? [this.currentYear - 1, 11]
                : [this.currentYear, this.currentMonth - 1];
            const [ny, nm] = this.currentMonth === 11
                ? [this.currentYear + 1, 0]
                : [this.currentYear, this.currentMonth + 1];
            const prev = this.store.getEventsForMonth(py, pm);
            const next = this.store.getEventsForMonth(ny, nm);

            // Deduplicate by key — multi-day events may overlap month boundaries
            const seen = new Set();
            allEvents = [...cur, ...prev, ...next].filter(e => {
                if (seen.has(e.key)) return false;
                seen.add(e.key);
                return true;
            });
        }

        const filtered = this.search.getFilteredEvents(allEvents);

        if (this.view === 'status') {
            const el = document.getElementById('jenga-status');
            if (el) this.statusView.render(el);
            return;
        }

        if (this.view === 'calendar') {
            const el = document.getElementById('jenga-calendar');
            if (this.spanMode !== 'month' && this.spanStartDate) {
                this.calendar.renderSpan(el, this.spanStartDate, this._spanNumWeeks(), filtered);
            } else {
                this.calendar.render(el, this.currentYear, this.currentMonth, filtered);
            }
        } else if (this.view === 'timeline') {
            const el = document.getElementById('jenga-timeline');
            // Peak Season Protection Window must remain visible on the timeline even
            // when the Event Types filter does not explicitly include this technical type.
            // This mirrors the DayDrawer behaviour, where PSPW events are always re-added.
            const timelineEvents = [
                ...new Map([
                    ...filtered,
                    ...allEvents.filter(e => e.type === 'PEAK_SEASON_PROTECTION_WINDOW')
                ].map(e => [e.key, e])).values()
            ];
            this.timeline.render(el, this.currentYear, this.currentMonth, timelineEvents, {
                spanMode: this.spanMode, spanStart: this.spanStartDate,
            });
        }
    }

    _hideSpinner() {
        const el = document.getElementById('app-spinner');
        if (!el) return;
        el.classList.add('is-hidden');
        el.addEventListener('transitionend', () => el.remove(), { once: true });
    }

    _applyStoredTheme() {
        const isTV = isTVBrowser() || new URLSearchParams(location.search).get('tvmode') === 'true';
        if (isTV) {
            // TV browsers: always dark, never touch localStorage
            document.documentElement.setAttribute('data-theme', 'dark');
            return;
        }
        const theme = loadSavedTheme?.() || 'light';
        const toggle = document.getElementById('toggle-dark-mode');
        if (toggle) toggle.checked = theme === 'dark';
        document.getElementById('toggle-dark-mode')?.addEventListener('change', (e) => {
            const spinner = showThemeSwitchSpinner();
            applyTheme(e.target.checked ? 'dark' : 'light');
            requestAnimationFrame(() => requestAnimationFrame(() => {
                this.refresh();
                this.search.refreshChips();
                spinner.remove();
            }));
        });
    }

    _initSideDrawer() {
        const overlay = document.getElementById('side-overlay');
        const drawer  = document.getElementById('side-drawer');
        this._ensureTopBarPinSwitch(drawer);

        const openDrawer = () => {
            this._ensureTopBarPinSwitch(drawer);
            drawer?.classList.add('open');
            drawer?.setAttribute('aria-hidden', 'false');
            overlay?.classList.add('visible');
            document.body.classList.add('side-drawer-open');
        };

        const closeDrawer = () => {
            drawer?.classList.remove('open');
            drawer?.setAttribute('aria-hidden', 'true');
            overlay?.classList.remove('visible');
            document.body.classList.remove('side-drawer-open');
        };

        document.getElementById('toggle-cta')?.addEventListener('click', () => {
            drawer?.classList.contains('open') ? closeDrawer() : openDrawer();
        });

        document.getElementById('side-close')?.addEventListener('click', closeDrawer);
        overlay?.addEventListener('click', closeDrawer);

        document.getElementById('act-today')?.addEventListener('click', () => {
            this._goToToday();
        });

        document.getElementById('act-about')?.addEventListener('click', () => {
            closeDrawer();
            this.drawer.showAbout();
        });
    }

    _loadTopBarPinned() {
        try {
            return localStorage.getItem(TOPBAR_PIN_STORAGE_KEY) === 'true';
        } catch (_) {
            return false;
        }
    }

    _saveTopBarPinned(value) {
        try {
            localStorage.setItem(TOPBAR_PIN_STORAGE_KEY, value ? 'true' : 'false');
        } catch (_) {}
    }

    _applyTopBarPinnedState() {
        document.body.classList.toggle('jenga-topbar-pinned', Boolean(this.topBarPinned));
        const toggle = document.getElementById('jenga-pin-topbar-toggle');
        if (toggle) toggle.checked = Boolean(this.topBarPinned);
        this._refreshTopBarPinnedGeometry();
        requestAnimationFrame(() => {
            this._refreshTopBarPinnedGeometry();
            this.timeline?.refreshStickyLayout?.();
        });
    }

    _refreshTopBarPinnedGeometry() {
        const topBar = document.querySelector('.jenga-top-bar');
        const height = Math.ceil(topBar?.getBoundingClientRect().height || 0);
        document.documentElement.style.setProperty('--jenga-topbar-height', `${height}px`);
        document.body.style.setProperty('--jenga-topbar-height', `${height}px`);
    }

    setTopBarPinned(value, { persist = true } = {}) {
        this.topBarPinned = Boolean(value);
        if (persist) this._saveTopBarPinned(this.topBarPinned);
        this._applyTopBarPinnedState();
    }

    _ensureTopBarPinSwitch(drawer = document.getElementById('side-drawer')) {
        if (!drawer) return;

        // Remove stale controls that may have been created outside the current drawer.
        document.querySelectorAll('#jenga-pin-topbar-control').forEach(el => {
            if (!drawer.contains(el)) el.remove();
        });

        let row = drawer.querySelector('#jenga-pin-topbar-control');
        if (!row) {
            row = document.createElement('div');
            row.id = 'jenga-pin-topbar-control';
            row.className = 'jenga-topbar-pin-control side-menu__item';
            row.innerHTML = `
                <label class="jenga-setting-switch" for="jenga-pin-topbar-toggle">
                    <span class="jenga-setting-switch__text">
                        <span class="jenga-setting-switch__title">Fix top bar</span>
                        <span class="jenga-setting-switch__sub">Keep search and filters visible while scrolling</span>
                    </span>
                    <input id="jenga-pin-topbar-toggle" type="checkbox">
                    <span class="jenga-setting-switch__track" aria-hidden="true"></span>
                </label>`;
            row.querySelector('#jenga-pin-topbar-toggle')?.addEventListener('change', (e) => {
                this.setTopBarPinned(e.target.checked);
            });
        }

        const toggle = row.querySelector('#jenga-pin-topbar-toggle');
        if (toggle) toggle.checked = Boolean(this.topBarPinned);

        const target = drawer.querySelector('.side-drawer__content, .side-content, .drawer-content, .side-drawer__body') || drawer;

        // Prefer placing it right after the Dark Mode switch, because that is where
        // users expect view-level preferences to live. Fallback to the top of the drawer.
        const darkToggle = drawer.querySelector('#toggle-dark-mode');
        const darkRow = darkToggle?.closest('label, button, .side-menu__item, .side__row, .drawer-row, div');
        if (darkRow?.parentElement) {
            darkRow.parentElement.insertBefore(row, darkRow.nextSibling);
        } else if (target.firstElementChild) {
            target.insertBefore(row, target.firstElementChild.nextSibling || target.firstElementChild);
        } else {
            target.appendChild(row);
        }
    }

    _initNavControls() {
        this._updateMonthLabel();

        // Direct chip click → event detail without day context
        this.calendar.onEventClick((ev, groupedEvents) => {
            this.drawer.open(ev, false, groupedEvents);
            this.calendar.setSelectedKey(ev.key);
            // For grouped chips, encode the representative key as `group=` so the URL
            // deep-links back to the group drawer. For single events use `event=`.
            if (!groupedEvents) this._pushUrl({ event: ev.key }, true);
            else this._pushUrl({ group: ev.key }, true);
        });
        this.calendar.onDayClick(date => {
            this._openDayForDate(date);
            this.calendar.setSelectedDate(localDateStr(date));
        });
        this.dayDrawer.onClose(() => this.calendar.clearSelection());
        this.cardsDrawer.onClose(() => this.calendar.clearSelection());

        this.timeline.onEventClick(ev => {
            this.drawer.open(ev, false);
            this._pushUrl({ event: ev.key }, true);
        });
        this.timeline.onProtectionClick((date) => {
            this._openDayForDate(date);
        });
        this.timeline.onDotClick((date, issueTypeLabel) => {
            if (issueTypeLabel === 'release') {
                // Same as clicking a calendar day cell — open day timeline drawer
                this._openDayForDate(date);
            } else {
                const activeServices = this.search.activeServices;
                const issueSummaryQuery = this.timeline.getIssueSummaryQuery?.() || '';
                const cards = this._filterCardsByIssueSummary(
                    this.store.getCardsForDay(date, issueTypeLabel, activeServices),
                    issueSummaryQuery
                );
                this.cardsDrawer.open(date, issueTypeLabel, cards);
                this._pushUrl({ day: localDateStr(date), cards: issueTypeLabel }, true);
            }
        });
        // Event click from day drawer → level-2 panel over day timeline
        this.dayDrawer.onEventClick(ev => this.drawer.open(ev, true));

        document.getElementById('jenga-prev')?.addEventListener('click', () => {
            if (this.spanMode !== 'month' && this.spanStartDate) {
                this.spanStartDate = new Date(this.spanStartDate);
                this.spanStartDate.setDate(this.spanStartDate.getDate() - 7);
                this.currentYear  = this.spanStartDate.getFullYear();
                this.currentMonth = this.spanStartDate.getMonth();
            } else {
                this.currentMonth--;
                if (this.currentMonth < 0) { this.currentMonth = 11; this.currentYear--; }
            }
            this._updateMonthLabel();
            this._refreshTopBarPinnedGeometry?.();
            this.refresh();
            this._pushUrl();
        });

        document.getElementById('jenga-next')?.addEventListener('click', () => {
            if (this.spanMode !== 'month' && this.spanStartDate) {
                this.spanStartDate = new Date(this.spanStartDate);
                this.spanStartDate.setDate(this.spanStartDate.getDate() + 7);
                this.currentYear  = this.spanStartDate.getFullYear();
                this.currentMonth = this.spanStartDate.getMonth();
            } else {
                this.currentMonth++;
                if (this.currentMonth > 11) { this.currentMonth = 0; this.currentYear++; }
            }
            this._updateMonthLabel();
            this._refreshTopBarPinnedGeometry?.();
            this.refresh();
            this._pushUrl();
        });

        document.getElementById('jenga-today-btn')?.addEventListener('click', () => {
            this._goToToday();
        });
    }

    _filterCardsByIssueSummary(cards, query) {
        const q = (query || '').trim().toLowerCase();
        if (!q) return cards || [];
        return (cards || []).filter(c => (c.summary || '').toLowerCase().includes(q));
    }

    _openDayForDate(date) {
        const allForDay = this.store.getEventsForDay(date);
        const filtered  = this.search.getFilteredEvents(allForDay);

        // Include PSPW events from neighbouring days so getViolationColor can
        // compute the BLUE→AMBER sandwich upgrade correctly inside the drawer.
        const prevDay = new Date(date.getFullYear(), date.getMonth(), date.getDate() - 1);
        const nextDay = new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1);
        const neighbourPspw = [
            ...this.store.getEventsForDay(prevDay),
            ...this.store.getEventsForDay(nextDay),
        ].filter(e => e.type === 'PEAK_SEASON_PROTECTION_WINDOW');

        const finalEvents = [
            ...new Map([
                ...filtered,
                ...allForDay.filter(e => e.type === 'PEAK_SEASON_PROTECTION_WINDOW'),
                ...neighbourPspw,
            ].map(e => [e.key, e])).values()
        ];

        this._dayDrawerOpen = true;
        this.dayDrawer.open(date, finalEvents);

        this._pushUrl({ day: localDateStr(date) }, true);
    }

    _goToToday() {
        const now = new Date();
        if (this.spanMode !== 'month') {
            this.spanStartDate = this._mondayOf(now);
            this.currentYear   = this.spanStartDate.getFullYear();
            this.currentMonth  = this.spanStartDate.getMonth();
        } else {
            this.currentYear  = now.getFullYear();
            this.currentMonth = now.getMonth();
        }
        this._updateMonthLabel();
        this.refresh();
    }

    _updateMonthLabel() {
        const label = document.getElementById('jenga-month-label');
        if (!label) return;
        if (this.spanMode !== 'month' && this.spanStartDate) {
            label.textContent = this._spanLabel();
        } else {
            const d = new Date(this.currentYear, this.currentMonth, 1);
            label.textContent = d.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
        }
    }

    // ── Span mode helpers ─────────────────────────────────────────────────────

    _loadSpanMode() {
        try {
            const v = localStorage.getItem(SPAN_STORAGE_KEY);
            return VALID_SPANS.has(v) ? v : 'month';
        } catch (_) { return 'month'; }
    }

    _saveSpanMode(value) {
        try { localStorage.setItem(SPAN_STORAGE_KEY, value); } catch (_) {}
    }

    /** Monday (local-midnight) of the week containing `date`. */
    _mondayOf(date) {
        const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
        const day = d.getDay(); // 0=Sun … 6=Sat
        d.setDate(d.getDate() - ((day + 6) % 7));
        return d;
    }

    /** Number of weeks for the current spanMode (null when 'month'). */
    _spanNumWeeks() {
        return SPAN_WEEKS[this.spanMode] ?? null;
    }

    /** Last Date (local-midnight) of the current span. */
    _spanEndDate() {
        if (this.spanMode === 'month' || !this.spanStartDate) {
            return new Date(this.currentYear, this.currentMonth + 1, 0);
        }
        const end = new Date(this.spanStartDate);
        end.setDate(end.getDate() + SPAN_WEEKS[this.spanMode] * 7 - 1);
        return end;
    }

    /** Human-readable label for the current week span (e.g. "13–26 Oct 2026" or "28 Oct – 3 Nov 2026"). */
    _spanLabel() {
        if (!this.spanStartDate) return '';
        const start = this.spanStartDate;
        const end   = this._spanEndDate();
        const sm = start.getMonth();
        const em = end.getMonth();
        const sy = start.getFullYear();
        const ey = end.getFullYear();

        const fmtDay  = d => d.getDate();
        const fmtMon  = d => d.toLocaleDateString('en-GB', { month: 'short' });
        const fmtYear = d => d.getFullYear();

        if (sy !== ey) {
            // Cross-year
            return `${fmtDay(start)} ${fmtMon(start)} ${fmtYear(start)} – ${fmtDay(end)} ${fmtMon(end)} ${fmtYear(end)}`;
        }
        if (sm !== em) {
            // Same year, cross-month
            return `${fmtDay(start)} ${fmtMon(start)} – ${fmtDay(end)} ${fmtMon(end)} ${sy}`;
        }
        // Same month
        return `${fmtDay(start)}–${fmtDay(end)} ${fmtMon(start)} ${sy}`;
    }

    _initSpanSelector() {
        const sel = document.getElementById('jenga-span-select');
        if (!sel) return;

        // Set selector to the current span mode (loaded from localStorage or default)
        sel.value = this.spanMode;

        // If starting in week mode, initialise spanStartDate
        if (this.spanMode !== 'month' && !this.spanStartDate) {
            const now = new Date();
            const firstOfMonth = new Date(this.currentYear, this.currentMonth, 1);
            // If today is in the current month, start from current week's Monday
            const todayInMonth = now.getFullYear() === this.currentYear && now.getMonth() === this.currentMonth;
            this.spanStartDate = this._mondayOf(todayInMonth ? now : firstOfMonth);
        }

        sel.addEventListener('change', () => {
            const prev = this.spanMode;
            this.spanMode = sel.value;
            this._saveSpanMode(this.spanMode);

            if (this.spanMode !== 'month') {
                if (!this.spanStartDate || prev === 'month') {
                    // Entering week mode: pick the right Monday
                    const now = new Date();
                    const firstOfMonth = new Date(this.currentYear, this.currentMonth, 1);
                    const todayInMonth = now.getFullYear() === this.currentYear && now.getMonth() === this.currentMonth;
                    this.spanStartDate = this._mondayOf(todayInMonth ? now : firstOfMonth);
                }
                this.currentYear  = this.spanStartDate.getFullYear();
                this.currentMonth = this.spanStartDate.getMonth();
            } else {
                // Returning to month mode: derive year/month from span start if available
                if (this.spanStartDate) {
                    this.currentYear  = this.spanStartDate.getFullYear();
                    this.currentMonth = this.spanStartDate.getMonth();
                }
                this.spanStartDate = null;
            }

            this._updateMonthLabel();
            this._refreshTopBarPinnedGeometry?.();
            this.refresh();
            this._pushUrl();
        });
    }

    _initViewToggle() {
        const calBtn  = document.getElementById('jenga-view-calendar');
        const timeBtn = document.getElementById('jenga-view-timeline');
        const statBtn = document.getElementById('jenga-view-status');
        const calWrap = document.getElementById('jenga-calendar');
        const timWrap = document.getElementById('jenga-timeline');
        const statWrap = document.getElementById('jenga-status');

        calBtn?.addEventListener('click', () => {
            this.view = 'calendar';
            calBtn.classList.add('jenga-view-btn--active');
            timeBtn?.classList.remove('jenga-view-btn--active');
            statBtn?.classList.remove('jenga-view-btn--active');
            if (calWrap)  calWrap.style.display  = '';
            if (timWrap)  timWrap.style.display   = 'none';
            if (statWrap) statWrap.style.display  = 'none';
            document.body.classList.remove('jenga-view--timeline');
            this._refreshTopBarPinnedGeometry?.();
            this.refresh();
            this._pushUrl();
        });

        timeBtn?.addEventListener('click', () => {
            this.view = 'timeline';
            timeBtn.classList.add('jenga-view-btn--active');
            calBtn?.classList.remove('jenga-view-btn--active');
            statBtn?.classList.remove('jenga-view-btn--active');
            if (calWrap)  calWrap.style.display  = 'none';
            if (timWrap)  timWrap.style.display   = '';
            if (statWrap) statWrap.style.display  = 'none';
            document.body.classList.add('jenga-view--timeline');
            this._refreshTopBarPinnedGeometry?.();
            this.refresh();
            this._pushUrl();
        });

        statBtn?.addEventListener('click', () => {
            this.view = 'status';
            statBtn.classList.add('jenga-view-btn--active');
            calBtn?.classList.remove('jenga-view-btn--active');
            timeBtn?.classList.remove('jenga-view-btn--active');
            if (calWrap)  calWrap.style.display  = 'none';
            if (timWrap)  timWrap.style.display   = 'none';
            if (statWrap) statWrap.style.display  = '';
            document.body.classList.remove('jenga-view--timeline');
            this._refreshTopBarPinnedGeometry?.();
            this.statusView.invalidate();
            this.refresh();
            this._pushUrl();
        });
    }

    _initSlideshow() {
        const ssBtn = document.getElementById('jenga-view-slideshow');
        ssBtn?.addEventListener('click', () => {
            if (this.slideshow.active) {
                this.slideshow.stop();
            } else {
                this.slideshow.start();
            }
        });
    }

    _initBuildInfo() {
        try {
            const el = document.getElementById('build-info');
            if (el) {
                el.textContent = `Build ${__APP_BUILD__} · ${__BUILD_DATE__}`;
                el.style.cursor = 'pointer';
                el.title = 'Click to show changelog';
                el.addEventListener('click', () => this.announcementBar.show());
            }
            const lastUpdateEl = document.getElementById('side-last-update');
            if (lastUpdateEl) {
                lastUpdateEl.style.cursor = 'pointer';
                lastUpdateEl.title = 'Click to show changelog';
                lastUpdateEl.addEventListener('click', () => this.announcementBar.show());
            }
        } catch {}
    }

    _parseServiceNames(csvText) {
        if (!csvText) return { names: [], amberAllowed: new Set() };
        // RFC 4180-compliant tokeniser: handles quoted fields containing commas
        // AND embedded newlines (which break a naive split('\n') approach).
        const rows = this._parseCsvRows(csvText);
        if (rows.length < 2) return { names: [], amberAllowed: new Set() };
        const headers = rows[0].map(h => h.trim());
        const nameIdx     = headers.indexOf('Service Name');
        const typeIdx     = headers.indexOf('Type');
        const techRiskIdx = headers.indexOf('Technology Risk Level');
        if (nameIdx === -1) return { names: [], amberAllowed: new Set() };
        const AMBER_ALLOWED_TYPES     = new Set(['cots application', 'custom backend', 'custom frontend']);
        const AMBER_ALLOWED_TECH_RISK = new Set(['low', 'medium']);
        const names = new Set();
        const amberAllowed = new Set();
        for (let i = 1; i < rows.length; i++) {
            const name = (rows[i][nameIdx] || '').trim();
            if (!name) continue;
            names.add(name);
            const type     = (typeIdx    >= 0 ? rows[i][typeIdx]    || '' : '').trim().toLowerCase();
            const techRisk = (techRiskIdx >= 0 ? rows[i][techRiskIdx] || '' : '').trim().toLowerCase();
            if (AMBER_ALLOWED_TYPES.has(type) && AMBER_ALLOWED_TECH_RISK.has(techRisk)) {
                amberAllowed.add(name.toLowerCase());
            }
        }
        return {
            names: [...names].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' })),
            amberAllowed,
        };
    }

    /** Parse a CSV string into an array of string arrays (rows × fields).
     *  Handles quoted fields with embedded commas and newlines (RFC 4180). */
    _parseCsvRows(text) {
        const rows = [];
        let row = [], field = '', inQ = false;
        for (let i = 0; i < text.length; i++) {
            const ch = text[i];
            if (inQ) {
                if (ch === '"') {
                    // Peek ahead: "" inside quotes = escaped quote
                    if (text[i + 1] === '"') { field += '"'; i++; }
                    else                      { inQ = false; }
                } else {
                    field += ch;
                }
            } else {
                if (ch === '"') {
                    inQ = true;
                } else if (ch === ',') {
                    row.push(field); field = '';
                } else if (ch === '\r' && text[i + 1] === '\n') {
                    row.push(field); field = ''; rows.push(row); row = []; i++;
                } else if (ch === '\n') {
                    row.push(field); field = ''; rows.push(row); row = [];
                } else {
                    field += ch;
                }
            }
        }
        // Last field / row
        if (field || row.length) { row.push(field); rows.push(row); }
        return rows;
    }

    _updateLastUpdate() {
        const el = document.getElementById('side-last-update');
        if (!el) return;

        // Show when jira-cards.csv was last generated by CI
        const gen = this.store.cardsGeneratedAt;
        if (gen) {
            const d = new Date(gen);
            if (!isNaN(d.getTime())) {
                const time = d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
                const date = d.toLocaleDateString([], { day: '2-digit', month: 'short' });
                el.textContent = `Data as of ${date} ${time} UTC`;
                return;
            }
        }

        // Fallback: latest event due date
        const dates = this.store.events
            .map(e => e.dueDate)
            .filter(Boolean)
            .map(d => d.getTime());
        if (dates.length) {
            const latest = new Date(Math.max(...dates));
            el.textContent = `Last event: ${latest.toLocaleDateString()}`;
        }
    }

    // ── URL sync ──────────────────────────────────────────────────────────────

    _initUrlSync() {
        window.addEventListener('popstate', (e) => {
            const state = e.state || {};
            if (state.year  !== undefined) this.currentYear  = state.year;
            if (state.month !== undefined) this.currentMonth = state.month;
            if (state.view) this._applyView(state.view);
            // Restore span state from history entry
            if (state.spanMode && VALID_SPANS.has(state.spanMode)) {
                this.spanMode = state.spanMode;
            } else {
                this.spanMode = 'month';
            }
            if (this.spanMode !== 'month' && state.spanStart) {
                const [sy, sm, sd] = state.spanStart.split('-').map(Number);
                this.spanStartDate = new Date(sy, sm - 1, sd);
            } else {
                this.spanStartDate = null;
            }
            const spanSel = document.getElementById('jenga-span-select');
            if (spanSel) spanSel.value = this.spanMode;
            this._updateMonthLabel();
            this.refresh();
        });
    }

    /** Build the canonical URL params from current app state and any extras.
     *  @param {Object} [extra]  — additional params (e.g. { day, event, q })
     *  @param {boolean} [push]  — true = pushState (adds history entry); false = replaceState */
    _pushUrl(extra = {}, push = false) {
        try {
            const p = new URLSearchParams();
            p.set('view',  this.view);
            p.set('year',  String(this.currentYear));
            p.set('month', String(this.currentMonth + 1)); // human 1-12
            if (this.spanMode !== 'month') {
                p.set('span', this.spanMode);
                if (this.spanStartDate) p.set('spanStart', localDateStr(this.spanStartDate));
            }
            if (extra.day)   p.set('day',   extra.day);
            if (extra.event) p.set('event', extra.event);
            if (extra.group) p.set('group', extra.group);
            if (extra.cards) p.set('cards', extra.cards);
            const q = document.getElementById('jenga-search-input')?.value?.trim();
            if (q) p.set('q', q);
            const filterParams = this.search.getUrlParams?.() || {};
            for (const [k, v] of Object.entries(filterParams)) p.set(k, v);

            // Comparison state (timeline only — omit defaults to keep URLs short)
            const cs = this.timeline.getComparisonState?.();
            if (cs) {
                if (cs.incidentAvgMonths !== 6) p.set('incAvg', String(cs.incidentAvgMonths));
                if (cs.srAvgMonths       !== 6) p.set('srAvg',  String(cs.srAvgMonths));
                if (cs.incidentLastYear)        p.set('incLY',  '1');
                if (cs.srLastYear)              p.set('srLY',   '1');
                const incCmp = cs.incidentCompareMonths
                    .map(m => `${m.year}-${String(m.month + 1).padStart(2, '0')}`).join(',');
                if (incCmp) p.set('incCmp', incCmp);
                const srCmp = cs.srCompareMonths
                    .map(m => `${m.year}-${String(m.month + 1).padStart(2, '0')}`).join(',');
                if (srCmp) p.set('srCmp', srCmp);
            }

            const url = `${location.pathname}?${p.toString()}`;
            const state = {
                year: this.currentYear, month: this.currentMonth, view: this.view,
                spanMode: this.spanMode,
                spanStart: this.spanStartDate ? localDateStr(this.spanStartDate) : null,
            };
            if (push) history.pushState(state, '', url);
            else      history.replaceState(state, '', url);
        } catch (_) {}
    }

    /** Apply navigation params from URL (view, year, month, q) — called after data loads. */
    _readUrlParams() {
        try {
            const p = new URLSearchParams(location.search);
            const view = p.get('view');
            if (view === 'calendar' || view === 'timeline' || view === 'status') this._applyView(view);
            const year  = parseInt(p.get('year')  || '', 10);
            const month = parseInt(p.get('month') || '', 10);
            if (!isNaN(year)  && year  > 2000) this.currentYear  = year;
            if (!isNaN(month) && month >= 1 && month <= 12) this.currentMonth = month - 1;
            const q = p.get('q');
            if (q) {
                const input = document.getElementById('jenga-search-input');
                if (input) { input.value = q; this.search.query = q; }
            }
            this.search.applyUrlParams?.(p);

            // Restore span mode from URL (overrides localStorage)
            const span = p.get('span');
            if (VALID_SPANS.has(span) && span !== 'month') {
                this.spanMode = span;
                const ss = p.get('spanStart');
                if (ss) {
                    const [sy, sm, sd] = ss.split('-').map(Number);
                    if (sy && sm && sd) this.spanStartDate = new Date(sy, sm - 1, sd);
                }
                if (!this.spanStartDate) {
                    this.spanStartDate = this._mondayOf(new Date(this.currentYear, this.currentMonth, 1));
                }
            }
            // Sync the selector element to current spanMode
            const spanSel = document.getElementById('jenga-span-select');
            if (spanSel) spanSel.value = this.spanMode;

            // Restore comparison state (overrides localStorage)
            const compState = {};
            const incAvg = parseInt(p.get('incAvg') || '', 10);
            if ([3, 6, 12, 24].includes(incAvg)) compState.incidentAvgMonths = incAvg;
            const srAvg  = parseInt(p.get('srAvg')  || '', 10);
            if ([3, 6, 12, 24].includes(srAvg))  compState.srAvgMonths       = srAvg;
            if (p.get('incLY') === '1') compState.incidentLastYear = true;
            if (p.get('srLY')  === '1') compState.srLastYear       = true;
            const parseMonths = str => (str || '').split(',')
                .map(s => { const [y, m] = s.split('-').map(Number); return (y && m) ? { year: y, month: m - 1 } : null; })
                .filter(Boolean);
            const incCmp = parseMonths(p.get('incCmp'));
            if (incCmp.length) compState.incidentCompareMonths = incCmp;
            const srCmp  = parseMonths(p.get('srCmp'));
            if (srCmp.length)  compState.srCompareMonths       = srCmp;
            if (Object.keys(compState).length) this.timeline.applyComparisonState?.(compState);

            this._updateMonthLabel();
        } catch (_) {}
    }

    /** Open drawer/day from URL params — called after first render. */
    _applyUrlDrawerParams() {
        try {
            const p = new URLSearchParams(location.search);
            const groupKey = p.get('group');
            if (groupKey) {
                const rep = this.store.events.find(e => e.key === groupKey);
                if (rep) {
                    const envKey = (rep.environment || '').toLowerCase();
                    const svcKey = (rep.service || '').toLowerCase();
                    const due    = rep.dueDate?.getTime();
                    const grouped = this.store.events.filter(e =>
                        e.type === 'SERVICE_OP' &&
                        (e.environment || '').toLowerCase() === envKey &&
                        (e.service || '').toLowerCase() === svcKey &&
                        e.dueDate?.getTime() === due
                    );
                    this.drawer.open(rep, false, grouped.length > 1 ? grouped : null);
                    this.calendar.setSelectedKey(groupKey);
                }
                return;
            }
            const eventKey = p.get('event');
            if (eventKey) {
                const ev = this.store.events.find(e => e.key === eventKey);
                if (ev) {
                    this.drawer.open(ev, false);
                    this.calendar.setSelectedKey(eventKey);
                }
                return;
            }
            if (p.get('slideshow') === 'true') {
                this.slideshow.start();
                return;
            }
            const day   = p.get('day');
            const cards = p.get('cards');
            if (day) {
                const [y, m, d] = day.split('-').map(Number);
                if (y && m && d) {
                    const date = new Date(y, m - 1, d);
                    if (cards && cards !== 'release') {
                        const activeServices    = this.search.activeServices;
                        const issueSummaryQuery = this.timeline.getIssueSummaryQuery?.() || '';
                        const cardsList = this._filterCardsByIssueSummary(
                            this.store.getCardsForDay(date, cards, activeServices),
                            issueSummaryQuery
                        );
                        this.cardsDrawer.open(date, cards, cardsList);
                        this.calendar.setSelectedDate(day);
                    } else {
                        this._openDayForDate(date);
                        this.calendar.setSelectedDate(day);
                    }
                }
            }
        } catch (_) {}
    }

    /** Switch the view without triggering URL push (used by _readUrlParams/_applyView). */
    _applyView(view) {
        const calBtn   = document.getElementById('jenga-view-calendar');
        const timeBtn  = document.getElementById('jenga-view-timeline');
        const statBtn  = document.getElementById('jenga-view-status');
        const calWrap  = document.getElementById('jenga-calendar');
        const timWrap  = document.getElementById('jenga-timeline');
        const statWrap = document.getElementById('jenga-status');
        this.view = view;
        if (view === 'calendar') {
            calBtn?.classList.add('jenga-view-btn--active');
            timeBtn?.classList.remove('jenga-view-btn--active');
            statBtn?.classList.remove('jenga-view-btn--active');
            if (calWrap)  calWrap.style.display  = '';
            if (timWrap)  timWrap.style.display   = 'none';
            if (statWrap) statWrap.style.display  = 'none';
            document.body.classList.remove('jenga-view--timeline');
        } else if (view === 'status') {
            statBtn?.classList.add('jenga-view-btn--active');
            calBtn?.classList.remove('jenga-view-btn--active');
            timeBtn?.classList.remove('jenga-view-btn--active');
            if (calWrap)  calWrap.style.display  = 'none';
            if (timWrap)  timWrap.style.display   = 'none';
            if (statWrap) statWrap.style.display  = '';
            document.body.classList.remove('jenga-view--timeline');
        } else {
            timeBtn?.classList.add('jenga-view-btn--active');
            calBtn?.classList.remove('jenga-view-btn--active');
            statBtn?.classList.remove('jenga-view-btn--active');
            if (calWrap)  calWrap.style.display  = 'none';
            if (timWrap)  timWrap.style.display   = '';
            if (statWrap) statWrap.style.display  = 'none';
            document.body.classList.add('jenga-view--timeline');
        }
    }

    showToast(message, duration = 3000) {
        let container = document.querySelector('.toast-container');
        if (!container) {
            container = document.createElement('div');
            container.className = 'toast-container';
            document.body.appendChild(container);
        }
        const toast = document.createElement('div');
        toast.className = 'toast';
        toast.textContent = message;
        container.appendChild(toast);
        setTimeout(() => toast.classList.add('show'), 10);
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 300);
        }, duration);
    }
}
