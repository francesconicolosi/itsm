import { BRAND } from '../../brand-specific/brand.js';

const SLIDE_COUNT = 3;
const COMMAND_BAR_HIDE_DELAY = 3000;
const SLIDESHOW_RELOAD_MS = 45 * 60 * 1000;

export class SlideshowController {
    constructor(app, { slideDuration } = {}) {
        this.app = app;
        this.active = false;
        this.playing = false;
        this.currentSlide = 1;
        this.slideDuration = slideDuration ?? this._readSlideDuration();
        this._slideCount = SLIDE_COUNT;
        this._timer = null;
        this._reloadTimer = null;
        this._commandBarTimer = null;
        this._mouseMoveHandler = null;
        this._commandBar = null;
        this._outgoing = null;
    }

    _readSlideDuration() {
        try {
            const p = new URLSearchParams(window.location.search);
            const v = parseInt(p.get('slideDuration'), 10);
            if (!isNaN(v) && v > 0) return v * 1000;
        } catch {}
        return 25000;
    }

    start({ slideCount } = {}) {
        this._slideCount = (slideCount != null && slideCount > 0) ? slideCount : SLIDE_COUNT;
        this.active = true;
        this.playing = true;
        this.currentSlide = 1;

        // Hide normal UI
        document.body.classList.remove('jenga-view--timeline');
        document.querySelectorAll('.jenga-top-bar').forEach(el => el.style.display = 'none');
        const calWrap = document.getElementById('jenga-calendar');
        const timWrap = document.getElementById('jenga-timeline');
        if (calWrap) calWrap.style.display = 'none';
        if (timWrap) timWrap.style.display = 'none';

        // Show slideshow container
        const container = document.getElementById('jenga-slideshow-container');
        if (container) container.style.display = '';

        // Mark button active
        document.getElementById('jenga-view-slideshow')?.classList.add('jenga-view-btn--active');
        document.getElementById('jenga-view-calendar')?.classList.remove('jenga-view-btn--active');
        document.getElementById('jenga-view-timeline')?.classList.remove('jenga-view-btn--active');

        // Push slideshow=true to URL
        this._pushSlideshowUrl(true);

        this._renderCurrentSlide();
        this._startTimer();
        this._attachCommandBar();
        this._reloadTimer = setTimeout(() => location.reload(), SLIDESHOW_RELOAD_MS);

        this._escHandler = (e) => { if (e.key === 'Escape') this.stop(); };
        document.addEventListener('keydown', this._escHandler);
    }

    stop() {
        this._slideCount = SLIDE_COUNT;
        this.active = false;
        this.playing = false;
        this._clearTimer();
        if (this._escHandler) {
            document.removeEventListener('keydown', this._escHandler);
            this._escHandler = null;
        }
        this._detachCommandBar();
        if (this._reloadTimer !== null) {
            clearTimeout(this._reloadTimer);
            this._reloadTimer = null;
        }

        this._outgoing = null;
        const container = document.getElementById('jenga-slideshow-container');
        if (container) {
            container.style.display = 'none';
            container.innerHTML = '';
        }

        // Restore button states
        document.getElementById('jenga-view-slideshow')?.classList.remove('jenga-view-btn--active');

        // Remove slideshow=true from URL
        this._pushSlideshowUrl(false);

        // Restore normal UI based on app.view
        document.querySelectorAll('.jenga-top-bar').forEach(el => el.style.display = '');
        const calWrap = document.getElementById('jenga-calendar');
        const timWrap = document.getElementById('jenga-timeline');
        if (this.app.view === 'timeline') {
            if (calWrap) calWrap.style.display = 'none';
            if (timWrap) timWrap.style.display = '';
            document.body.classList.add('jenga-view--timeline');
            document.getElementById('jenga-view-timeline')?.classList.add('jenga-view-btn--active');
        } else {
            if (calWrap) calWrap.style.display = '';
            if (timWrap) timWrap.style.display = 'none';
            document.getElementById('jenga-view-calendar')?.classList.add('jenga-view-btn--active');
        }
    }

    play() {
        if (!this.active) return;
        this.playing = true;
        this._startTimer();
        this._updateCommandBar();
    }

    pause() {
        if (!this.active) return;
        this.playing = false;
        this._clearTimer();
        this._updateCommandBar();
    }

    next() {
        if (!this.active) return;
        this.currentSlide = (this.currentSlide % this._slideCount) + 1;
        this._renderCurrentSlide();
        if (this.playing) this._startTimer();
        this._updateCommandBar();
    }

    prev() {
        if (!this.active) return;
        this.currentSlide = this.currentSlide === 1 ? this._slideCount : this.currentSlide - 1;
        this._renderCurrentSlide();
        if (this.playing) this._startTimer();
        this._updateCommandBar();
    }

    _startTimer() {
        this._clearTimer();
        if (!this.playing) return;
        this._timer = setTimeout(() => {
            this.currentSlide = (this.currentSlide % this._slideCount) + 1;
            this._renderCurrentSlide();
            this._updateCommandBar();
            this._startTimer();
        }, this.slideDuration);
    }

    _clearTimer() {
        if (this._timer !== null) {
            clearTimeout(this._timer);
            this._timer = null;
        }
    }

    _appendBrandWatermark(slide) {
        if (!BRAND?.logo?.svgInline?.d) return;

        const ns = 'http://www.w3.org/2000/svg';

        const logoSvg = document.createElementNS(ns, 'svg');
        logoSvg.setAttribute('class', 'ss-brand-watermark__logo');
        logoSvg.setAttribute('viewBox', '0 0 1312 210');
        logoSvg.setAttribute('preserveAspectRatio', 'xMidYMid meet');
        logoSvg.setAttribute('aria-hidden', 'true');

        const logoPath = document.createElementNS(ns, 'path');
        logoPath.setAttribute('d', BRAND.logo.svgInline.d);
        logoPath.setAttribute('fill', 'currentColor');

        logoSvg.appendChild(logoPath);

        const subtitle = document.createElement('div');
        subtitle.className = 'ss-brand-watermark__subtitle';
        subtitle.textContent = 'Digital Service Management';

        const progressFill = document.createElement('div');
        progressFill.className = 'ss-brand-watermark__progress-fill';
        progressFill.setAttribute('aria-hidden', 'true');

        const wrap = document.createElement('div');
        wrap.className = 'ss-brand-watermark';

        wrap.appendChild(progressFill);
        wrap.appendChild(logoSvg);
        wrap.appendChild(subtitle);

        slide.appendChild(wrap);

        requestAnimationFrame(() => {
            if (!progressFill.isConnected) return;

            progressFill.style.setProperty(
                'transition-duration',
                `${this.slideDuration}ms`
            );

            progressFill.style.setProperty(
                '-webkit-transition-duration',
                `${this.slideDuration}ms`
            );

            requestAnimationFrame(() => {
                if (!progressFill.isConnected) return;

                progressFill.classList.add(
                    'ss-brand-watermark__progress-fill--active'
                );
            });
        });
    }

    _renderCurrentSlide() {
        const container = document.getElementById('jenga-slideshow-container');
        if (!container) return;

        const app = this.app;
        const store = app.store;

        // Build filtered event list (same logic as JengaApp.refresh)
        const cur  = store.getEventsForMonth(app.currentYear, app.currentMonth);
        const [py, pm] = app.currentMonth === 0
            ? [app.currentYear - 1, 11]
            : [app.currentYear, app.currentMonth - 1];
        const [ny, nm] = app.currentMonth === 11
            ? [app.currentYear + 1, 0]
            : [app.currentYear, app.currentMonth + 1];
        const prev = store.getEventsForMonth(py, pm);
        const next = store.getEventsForMonth(ny, nm);
        const seen = new Set();
        const allEvents = [...cur, ...prev, ...next].filter(e => {
            if (seen.has(e.key)) return false;
            seen.add(e.key);
            return true;
        });
        const filtered = app.search.getFilteredEvents(allEvents);

        // Immediately remove any previously-queued outgoing slide to avoid stacking
        // on rapid navigation, then fade out the current visible slide.
        if (this._outgoing) {
            this._outgoing.remove();
            this._outgoing = null;
        }
        const current = container.querySelector('.jenga-slide');
        if (current) {
            this._outgoing = current;
            current.classList.remove('jenga-slide--visible');
            current.addEventListener('transitionend', () => {
                current.remove();
                if (this._outgoing === current) this._outgoing = null;
            }, { once: true });
        }

        const slideWrap = document.createElement('div');
        slideWrap.className = `jenga-slide jenga-slide--${this.currentSlide}`;
        container.appendChild(slideWrap);

        if (this.currentSlide === 1) {
            // Calendar view — render first, then inject month label so it isn't wiped
            // by CalendarRenderer.render() which does containerEl.innerHTML = ''
            app.calendar.render(slideWrap, app.currentYear, app.currentMonth, filtered, { maxChips: 3 });

            const monthLabel = document.createElement('div');
            monthLabel.className = 'ss-month-label';
            monthLabel.textContent = new Date(app.currentYear, app.currentMonth, 1)
                .toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
            slideWrap.appendChild(monthLabel);
            this._appendBrandWatermark(slideWrap);

            requestAnimationFrame(() => {
                this._scaleCalendar(slideWrap, container);
                slideWrap.classList.add('jenga-slide--visible');
            });
        } else {
            // Timeline panels — render directly, then remove unwanted lanes
            const timelineEvents = [
                ...new Map([
                    ...filtered,
                    ...allEvents.filter(e => e.type === 'PEAK_SEASON_PROTECTION_WINDOW')
                ].map(e => [e.key, e])).values()
            ];
            app.timeline.render(slideWrap, app.currentYear, app.currentMonth, timelineEvents);

            // Filter by data-lane attribute (set by _makeCollapsible)
            const keepLanes = this.currentSlide === 2
                ? new Set(['biz', 'milestones'])
                : new Set(['incidents', 'sr', 'releases', 'peak-protection']);

            const main = slideWrap.querySelector('.jtl-main');
            if (main) {
                // Remove insights panel (not needed in slide 2/3)
                main.querySelector('.jtl-insights-panel')?.remove();

                Array.from(main.children).forEach(child => {
                    const laneKey = child.dataset && child.dataset.lane;
                    if (laneKey !== undefined && !keepLanes.has(laneKey)) {
                        child.remove();
                    }
                });

                // Slideshow is display-only; expand all kept lanes regardless of their
                // default-collapsed state (milestones lane defaults to collapsed).
                Array.from(main.children).forEach(child => {
                    child.classList.remove('jtl-lane--collapsed');
                });

                if (this.currentSlide === 2) {
                    // Slide 2 keeps biz + milestones: the jtl-axis-header (month label +
                    // day ruler) rendered by TimelineRenderer is already present and correct.
                    // Promote it to the top of main so it sits above the lanes.
                    const axisHeader = main.querySelector('.jtl-axis-header');
                    if (axisHeader) main.insertBefore(axisHeader, main.firstChild);
                } else {
                    // Slide 3 includes releases (day-aligned chart) — keep axis header, same as slide 2.
                    const axisHeader = main.querySelector('.jtl-axis-header');
                    if (axisHeader) main.insertBefore(axisHeader, main.firstChild);
                }
            }

            this._appendBrandWatermark(slideWrap);
            requestAnimationFrame(() => slideWrap.classList.add('jenga-slide--visible'));
        }
    }

    _attachCommandBar() {
        if (this._commandBar) this._detachCommandBar();

        const bar = document.createElement('div');
        bar.id = 'jenga-slideshow-bar';
        bar.className = 'jenga-slideshow-bar';
        bar.innerHTML = `
            <button id="ss-btn-prev" class="ss-btn" title="Previous slide" aria-label="Previous">‹</button>
            <button id="ss-btn-play" class="ss-btn" title="Play" aria-label="Play">▶</button>
            <button id="ss-btn-pause" class="ss-btn" title="Pause" aria-label="Pause">⏸</button>
            <button id="ss-btn-next" class="ss-btn" title="Next slide" aria-label="Next">›</button>
            <span class="ss-indicator" id="ss-indicator">1 / ${this._slideCount}</span>
            <button id="ss-btn-stop" class="ss-btn ss-btn--stop" title="Stop slideshow" aria-label="Stop">✖ Stop</button>
        `;
        document.body.appendChild(bar);
        this._commandBar = bar;

        bar.querySelector('#ss-btn-prev').addEventListener('click', () => this.prev());
        bar.querySelector('#ss-btn-play').addEventListener('click', () => this.play());
        bar.querySelector('#ss-btn-pause').addEventListener('click', () => this.pause());
        bar.querySelector('#ss-btn-next').addEventListener('click', () => this.next());
        bar.querySelector('#ss-btn-stop').addEventListener('click', () => this.stop());

        this._updateCommandBar();

        // Show on mouse move, hide after inactivity
        this._mouseMoveHandler = () => this._showCommandBar();
        document.addEventListener('mousemove', this._mouseMoveHandler);
        // Initially hidden — will appear on first mouse move
        bar.classList.add('jenga-slideshow-bar--hidden');
    }

    _detachCommandBar() {
        if (this._commandBar) {
            this._commandBar.remove();
            this._commandBar = null;
        }
        if (this._mouseMoveHandler) {
            document.removeEventListener('mousemove', this._mouseMoveHandler);
            this._mouseMoveHandler = null;
        }
        if (this._commandBarTimer) {
            clearTimeout(this._commandBarTimer);
            this._commandBarTimer = null;
        }
    }

    _showCommandBar() {
        if (!this._commandBar) return;
        this._commandBar.classList.remove('jenga-slideshow-bar--hidden');
        if (this._commandBarTimer) clearTimeout(this._commandBarTimer);
        this._commandBarTimer = setTimeout(() => {
            this._commandBar?.classList.add('jenga-slideshow-bar--hidden');
        }, COMMAND_BAR_HIDE_DELAY);
    }

    _updateCommandBar() {
        if (!this._commandBar) return;
        const indicator = this._commandBar.querySelector('#ss-indicator');
        if (indicator) indicator.textContent = `${this.currentSlide} / ${this._slideCount}`;

        const playBtn  = this._commandBar.querySelector('#ss-btn-play');
        const pauseBtn = this._commandBar.querySelector('#ss-btn-pause');
        if (playBtn)  playBtn.style.display  = this.playing ? 'none' : '';
        if (pauseBtn) pauseBtn.style.display = this.playing ? '' : 'none';
    }

    _scaleCalendar(slideWrap, container) {
        const cal = slideWrap.querySelector('.jenga-cal');
        if (!cal) return;
        // Reset so we measure the calendar's natural (unscaled) dimensions
        cal.style.transform = '';
        cal.style.transformOrigin = '';
        const availW = container.clientWidth;
        // Reserve 32px at the top for the month label so it doesn't overlap the calendar
        const LABEL_RESERVE = 32;
        const availH = container.clientHeight - LABEL_RESERVE;
        const calW = cal.scrollWidth;
        const calH = cal.scrollHeight;
        if (calW > 0 && calH > 0) {
            const scale = Math.min(availW / calW, availH / calH);
            // Anchor at top-center so the calendar stays below the label
            cal.style.transformOrigin = 'top center';
            cal.style.transform = `scale(${scale})`;
        }
    }

    _buildDayAxis(year, month, main) {
        // Delegate to TimelineRenderer which owns the CHART_MARGIN constants.
        return this.app.timeline._buildDayAxisRuler(year, month, main);
    }

    _pushSlideshowUrl(active) {
        try {
            const p = new URLSearchParams(location.search);
            if (active) {
                p.set('slideshow', 'true');
            } else {
                p.delete('slideshow');
            }
            const url = `${location.pathname}?${p.toString()}`;
            history.replaceState(history.state, '', url);
        } catch (_) {}
    }
}
