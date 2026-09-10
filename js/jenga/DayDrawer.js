import { getTypeColor, getEnvColor, getEnvLabel, HYBRIS_SVG } from './CalendarRenderer.js';
import { getStatusStyle } from './EventDrawer.js';

// Always show the full 24-hour day
const DEFAULT_START_MIN = 0;          // 12:00 AM
const DEFAULT_END_MIN   = 24 * 60;    // 12:00 AM (next day)
const SCROLL_TO_HOUR    = 9;          // Initial scroll position: 9:00 AM

const DENSE_SLOT_THRESHOLD = 3; // group only when more than 4 releases share the same slot

function isCabProtectionEvent(ev) {
    if (!ev || ev.type !== 'PEAK_SEASON_PROTECTION_WINDOW') return false;
    return Boolean(ev.isCab) || /^CAB\s*[|]/i.test(ev.summary || '');
}

function formatDayTitle(date) {
    return date.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric' });
}

function minutesToLabel(min) {
    const h = Math.floor(min / 60);
    const m = min % 60;
    const period = h >= 12 ? 'PM' : 'AM';
    const hour = h % 12 || 12;
    return `${hour}:${String(m).padStart(2, '0')} ${period}`;
}

export class DayDrawer {
    constructor() {
        this._onEventClick = null;
        this._slotPopover = null;
        this._slotPopoverOutsideHandler = null;
    }

    onEventClick(fn) { this._onEventClick = fn; }

    open(date, events) {
        const drawer  = document.getElementById('drawer');
        const overlay = document.getElementById('overlay');
        const title   = document.getElementById('day-drawer-title');
        const content = document.getElementById('dayDrawerContent');
        if (!drawer || !content) return;

        // Reset strip to L1 immediately (no transition) so a previous L2 state
        // doesn't persist when the drawer re-opens for a new day.
        const panels = this._panels();
        if (panels) {
            panels.style.transition = 'none';
            panels.classList.remove('panels-level2');
            requestAnimationFrame(() => { panels.style.transition = ''; });
        }

        title.textContent = formatDayTitle(date);
        content.innerHTML = '';
        content.appendChild(this._buildContent(date, events));
        this._initTooltip(content);

        drawer.classList.add('open');
        overlay?.classList.add('open');

        // Scroll grid to 9:00 AM on open (after render so scrollHeight is set)
        const grid = content.querySelector('.day-timeline__grid');
        if (grid) {
            const PAD_TOP   = 20;
            const PX_PER_MIN = 1;
            grid.scrollTop = PAD_TOP + SCROLL_TO_HOUR * 60 * PX_PER_MIN;
        }
    }

    close() {
        this.closeAll();
    }

    _panels() {
        return document.getElementById('drawer')?.querySelector('.drawer-panels') || null;
    }

    openLevel2() {
        this._panels()?.classList.add('panels-level2');
    }

    closeLevel2() {
        this._panels()?.classList.remove('panels-level2');
    }

    closeAll() {
        this._hideSlotPopover();
        const drawer  = document.getElementById('drawer');
        const overlay = document.getElementById('overlay');
        const panels  = this._panels();

        // Slide the whole drawer out. Leave panels-level2 / drawer--direct-only
        // in place during the animation so the strip does not animate while the
        // drawer is still visible. Clean everything up once it has left the screen.
        drawer?.classList.remove('open');
        overlay?.classList.remove('open');

        const onEnd = () => {
            // Suppress the strip transition so the reset is invisible off-screen.
            if (panels) {
                panels.style.transition = 'none';
                panels.classList.remove('panels-level2');
                requestAnimationFrame(() => { panels.style.transition = ''; });
            }
            // Remove direct-mode classes now that the drawer is off-screen.
            drawer?.classList.remove('drawer--direct', 'drawer--direct-only');
        };
        const fallback = setTimeout(onEnd, 420);
        drawer?.addEventListener('transitionend', function handler() {
            clearTimeout(fallback);
            drawer.removeEventListener('transitionend', handler);
            onEnd();
        });
    }

    initEvents(eventDrawer) {
        this._eventDrawer = eventDrawer || null;

        // Both X buttons always close the entire drawer
        document.getElementById('closeDrawer')?.addEventListener('click', () => this.closeAll());
        document.getElementById('closeDrawer2')?.addEventListener('click', () => this.closeAll());
        // Back button: slide back to L1 (keep drawer open)
        document.getElementById('drawerBack')?.addEventListener('click', () => this.closeLevel2());
        // Overlay click: close everything
        document.getElementById('overlay')?.addEventListener('click', () => this.closeAll());
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') this.closeAll();
        });
    }

    _buildProtectionInfo(events) {
        const colors = events
            .filter(e => e.type === 'PEAK_SEASON_PROTECTION_WINDOW' && !isCabProtectionEvent(e))
            .map(e => e.protectionColor)
            .filter(Boolean);

        if (!colors.length) return null;

        let color = colors.includes('RED')
            ? 'RED'
            : colors.includes('AMBER')
                ? 'AMBER'
                : 'BLUE';

        const container = document.createElement('div');
        container.className = `pspw-box pspw-${color.toLowerCase()}`;

        let content = '';

        if (color === 'RED') {
            content = `
<strong>Peak Season Protection — RED</strong><br/>
No deployments are allowed on this day, except for critical P1/P2 fixes required to restore production stability.
        `;
        }

        if (color === 'AMBER') {
            content = `
<strong>Peak Season Protection — AMBER</strong><br/>
Only deployments for approved applications are allowed.<br/>
<a href="domino.html?search=Type%3ACOTS+Application%2CCustom+Backend%2CCustom+Frontend%26Technology+Risk+Level%3ALow%2CMedium&listView=ID%2CDescription%2CType%2CDepends+on%2CStatus%2CDecommission+Date" target="_blank">View allowed applications</a>
        `;
        }

        if (color === 'BLUE') {
            content = `
<strong>Peak Season Protection — PLANNED DEPLOY WINDOW</strong><br/>
Production changes are allowed only through the official process.<br/>
<a href="https://brand.atlassian.net/servicedesk/customer/portal/1/article/2672164887?source=topic" target="_blank">View process</a><br/><br/>

If this window falls within an AMBER window, applications listed here:<br/>
<a href="domino.html?search=Type%3ACOTS+Application%2CCustom+Backend%2CCustom+Frontend%26Technology+Risk+Level%3ALow%2CMedium&listView=ID%2CDescription%2CType%2CDepends+on%2CStatus%2CDecommission+Date" target="_blank">Approved applications</a><br/>
may be deployed without CAB. All others require CAB approval.
        `;
        }

        container.innerHTML = content;

        return container;
    }

    _buildContent(date, events) {
        const frag = document.createDocumentFragment();

        // Multi-day BUSINESS_EVENTs always go to all-day even when they carry a timeSlot
        const allDay = events.filter(e => !e.timeSlot || e.type === 'BUSINESS_EVENT');
        const timed  = events.filter(e => e.timeSlot  && e.type !== 'BUSINESS_EVENT');

        // ── All-day section ───────────────────────────────────────────────────
        const allDaySection = document.createElement('div');
        allDaySection.className = 'day-timeline__allday';

        const allDayLabel = document.createElement('div');
        allDayLabel.className = 'day-timeline__allday-label';
        allDayLabel.textContent = 'All day';
        allDaySection.appendChild(allDayLabel);

        if (allDay.length === 0) {
            const empty = document.createElement('div');
            empty.className = 'day-timeline__allday-empty';
            empty.textContent = '—';
            allDaySection.appendChild(empty);
        } else {
            allDay.forEach(ev => {
                const label = this._eventLabel(ev);
                const chip = document.createElement('div');
                chip.className = 'day-timeline__allday-chip';
                chip.dataset.key = ev.key;
                chip.style.borderLeftColor = getTypeColor(ev.type);
                chip.dataset.tooltip = label;
                // Prepend icon then label text
                const icon = this._buildEventIcon(ev);
                if (icon) chip.appendChild(icon);
                chip.appendChild(document.createTextNode(label));
                if (ev.type === 'MILESTONE' && ev.status) {
                    const { bg, text } = getStatusStyle(ev.status);
                    const sb = document.createElement('span');
                    sb.className = 'jenga-status-badge jenga-status-badge--inline';
                    sb.textContent = ev.status;
                    sb.style.background = bg;
                    sb.style.color = text;
                    chip.appendChild(sb);
                }
                chip.addEventListener('click', () => this._onEventClick?.(ev));
                allDaySection.appendChild(chip);
            });
        }
        frag.appendChild(allDaySection);

        const protectionBox = this._buildProtectionInfo(events);
        if (protectionBox) {
            frag.appendChild(protectionBox);
        }

        // ── Timed timeline section — always full 24h day ──────────────────────
        const rangeStart = DEFAULT_START_MIN; // 0
        const rangeEnd   = DEFAULT_END_MIN;   // 24*60
        const rangeSpan  = rangeEnd - rangeStart; // 1440

        // Outer grid: pure scroll viewport — height constrained by flex panel
        const grid = document.createElement('div');
        grid.className = 'day-timeline__grid';

        // Inner canvas: fixed-pixel height so positions are always well-defined
        // and the user can always scroll to the bottom regardless of viewport height.
        const PX_PER_MIN = 1; // 60px per hour — compact but scrollable
        const PAD_TOP    = 20; // must match .day-timeline__canvas padding-top
        const PAD_BOT    = 24;
        const canvasH = PAD_TOP + rangeSpan * PX_PER_MIN + PAD_BOT;
        const canvas = document.createElement('div');
        canvas.className = 'day-timeline__canvas';
        canvas.style.height = `${canvasH}px`;

        // Hour labels + grid lines (inside canvas, offset by PAD_TOP)
        for (let m = rangeStart; m <= rangeEnd; m += 60) {
            const row = document.createElement('div');
            row.className = 'day-timeline__hour';
            row.style.top = `${PAD_TOP + (m - rangeStart) * PX_PER_MIN}px`;

            const label = document.createElement('span');
            label.className = 'day-timeline__hour-label';
            label.textContent = minutesToLabel(m);
            row.appendChild(label);
            canvas.appendChild(row);
        }

        // Dense time-slot grouping:
// if more than DENSE_SLOT_THRESHOLD events share the same exact time slot,
// render one compact stack block and show all releases inside an interactive popup.
        const displayItems = this._buildDenseSlotDisplayItems(timed);

// Assign columns to resolve overlaps, including dense-slot stack blocks
        const colAssignments = this._assignColumns(displayItems);

// Event blocks — column-split overlapping events side-by-side
        const LABEL_W = 60; // px — must match padding-left of .day-timeline__grid
        const RIGHT_P = 12; // px — right padding of the grid

        colAssignments.forEach(({ event: item, colIdx, colCount }) => {
            const { startMin, endMin } = item.timeSlot;
            const topPx    = PAD_TOP + (startMin - rangeStart) * PX_PER_MIN;
            const heightPx = Math.max(18, (endMin - startMin) * PX_PER_MIN);

            // ── Dense slot compact block ───────────────────────────────────────────
            if (item.__isDenseSlot) {
                const groupEvents = item.events || [];
                const mainType = this._dominantType(groupEvents);
                const color = getTypeColor(mainType);

                const block = document.createElement('div');
                block.className = 'day-timeline__event day-timeline__event--stack';
                block.dataset.slotKey = item.slotKey;
                block.dataset.colIdx   = colIdx;
                block.dataset.colCount = colCount;

                block.style.top    = `${topPx}px`;
                block.style.height = `${Math.max(32, heightPx)}px`;

                block.style.setProperty('--col-idx',   String(colIdx));
                block.style.setProperty('--col-count', String(colCount));

                block.style.left  = `calc(${LABEL_W}px + ${colIdx} * ((100% - ${LABEL_W}px - ${RIGHT_P}px) / ${colCount}))`;
                block.style.width = `calc((100% - ${LABEL_W}px - ${RIGHT_P}px) / ${colCount} - 2px)`;
                block.style.right = 'auto';

                block.style.borderLeftColor = color;
                block.style.background = color + '20';

                const inner = document.createElement('div');
                inner.className = 'day-timeline__event-inner day-timeline__event-inner--stack';

                const badge = this._buildDenseSlotBadge(groupEvents);
                if (badge) inner.appendChild(badge);

                const title = document.createElement('span');
                title.className = 'day-timeline__event-label day-timeline__event-label--stack';
                title.textContent = `${groupEvents.length} releases`;

                const count = document.createElement('span');
                count.className = 'day-timeline__event-stack-count';
                count.textContent = `+${groupEvents.length}`;

                inner.appendChild(title);
                inner.appendChild(count);

                block.appendChild(inner);

                block.addEventListener('click', (e) => {
                    e.stopPropagation();
                    this._showSlotPopover(item, block);
                });

                canvas.appendChild(block);
                return;
            }

            // ── Normal single event block ──────────────────────────────────────────
            const ev = item;

            const block = document.createElement('div');
            block.className = 'day-timeline__event';
            block.dataset.key    = ev.key;
            block.dataset.colIdx   = colIdx;
            block.dataset.colCount = colCount;
            block.style.top    = `${topPx}px`;
            block.style.height = `${heightPx}px`;

            block.style.setProperty('--col-idx',   String(colIdx));
            block.style.setProperty('--col-count', String(colCount));

            block.style.left  = `calc(${LABEL_W}px + ${colIdx} * ((100% - ${LABEL_W}px - ${RIGHT_P}px) / ${colCount}))`;
            block.style.width = `calc((100% - ${LABEL_W}px - ${RIGHT_P}px) / ${colCount} - 2px)`;
            block.style.right = 'auto';

            block.style.borderLeftColor = getTypeColor(ev.type);
            block.style.background      = getTypeColor(ev.type) + '18';

            const evLabel = this._eventLabel(ev);
            const timeStr = `${ev.timeSlot.timeStart} – ${ev.timeSlot.timeEnd}`;
            block.dataset.tooltip = `${evLabel} · ${timeStr}`;

            const inner = document.createElement('div');
            inner.className = 'day-timeline__event-inner';

            const icon = this._buildEventIcon(ev);
            if (icon) inner.appendChild(icon);

            const labelEl = document.createElement('span');
            labelEl.className = 'day-timeline__event-label';
            labelEl.textContent = evLabel;

            inner.appendChild(labelEl);
            block.appendChild(inner);

            block.addEventListener('click', () => this._onEventClick?.(ev));

            canvas.appendChild(block);
        });

        grid.appendChild(canvas);

        frag.appendChild(grid);
        return frag;
    }

    _initTooltip(container) {
        let tip = document.getElementById('jenga-day-tip');
        if (!tip) {
            tip = document.createElement('div');
            tip.id = 'jenga-day-tip';
            tip.className = 'jenga-day-tooltip';
            document.body.appendChild(tip);
        }
        const show = (text, x, y) => {
            tip.textContent = text;
            // Clamp so tooltip stays inside viewport
            const tw = Math.min(320, tip.scrollWidth || 200);
            const left = Math.min(x + 12, window.innerWidth - tw - 8);
            const top  = Math.min(y + 4,  window.innerHeight - 32);
            tip.style.left    = `${left}px`;
            tip.style.top     = `${top}px`;
            tip.style.opacity = '1';
        };
        let timer = null;
        let lastX = 0, lastY = 0;
        container.addEventListener('mouseover', (e) => {
            const el = e.target.closest('[data-tooltip]');
            if (!el) { clearTimeout(timer); return; }
            lastX = e.clientX; lastY = e.clientY;
            clearTimeout(timer);
            timer = setTimeout(() => show(el.dataset.tooltip, lastX, lastY), 100);
        });
        container.addEventListener('mousemove', (e) => {
            lastX = e.clientX; lastY = e.clientY;
            if (tip.style.opacity === '1') {
                const tw = Math.min(320, tip.scrollWidth || 200);
                tip.style.left = `${Math.min(lastX + 12, window.innerWidth - tw - 8)}px`;
                tip.style.top  = `${Math.min(lastY + 4,  window.innerHeight - 32)}px`;
            }
        });
        container.addEventListener('mouseout', (e) => {
            const el = e.target.closest('[data-tooltip]');
            if (!el) return;
            clearTimeout(timer);
            tip.style.opacity = '0';
        });
    }

    _buildDenseSlotDisplayItems(timedEvents) {
        if (!timedEvents.length) return [];

        const groups = new Map();

        timedEvents.forEach(ev => {
            const ts = ev.timeSlot;
            const key = [
                ts.startMin,
                ts.endMin,
                ts.timeStart || '',
                ts.timeEnd || ''
            ].join('|');

            if (!groups.has(key)) groups.set(key, []);
            groups.get(key).push(ev);
        });

        const displayItems = [];

        groups.forEach((slotEvents, slotKey) => {
            if (slotEvents.length > DENSE_SLOT_THRESHOLD) {
                const first = slotEvents[0];

                displayItems.push({
                    __isDenseSlot: true,
                    slotKey,
                    timeSlot: first.timeSlot,
                    events: [...slotEvents].sort((a, b) => {
                        const la = this._eventLabel(a).toLowerCase();
                        const lb = this._eventLabel(b).toLowerCase();
                        return la.localeCompare(lb);
                    })
                });
            } else {
                displayItems.push(...slotEvents);
            }
        });

        return displayItems;
    }

    _dominantType(events) {
        if (!events?.length) return 'OTHER';

        const counts = new Map();

        events.forEach(ev => {
            const key = ev.type || 'OTHER';
            counts.set(key, (counts.get(key) || 0) + 1);
        });

        return [...counts.entries()]
            .sort((a, b) => b[1] - a[1])[0][0];
    }

    _showSlotPopover(slotItem, anchorEl) {
        this._hideSlotPopover();

        const events = slotItem.events || [];
        if (!events.length) return;

        const popover = document.createElement('div');
        popover.className = 'day-slot-popover';
        popover.setAttribute('role', 'dialog');

        const header = document.createElement('div');
        header.className = 'day-slot-popover__header';

        const titleWrap = document.createElement('div');
        titleWrap.className = 'day-slot-popover__title-wrap';

        const time = document.createElement('div');
        time.className = 'day-slot-popover__time';
        time.textContent = slotItem.timeSlot?.timeStart || minutesToLabel(slotItem.timeSlot.startMin);

        const title = document.createElement('div');
        title.className = 'day-slot-popover__title';
        title.textContent = `${events.length} releases`;

        titleWrap.appendChild(time);
        titleWrap.appendChild(title);

        const close = document.createElement('button');
        close.type = 'button';
        close.className = 'day-slot-popover__close';
        close.textContent = '×';
        close.addEventListener('click', (e) => {
            e.stopPropagation();
            this._hideSlotPopover();
        });

        header.appendChild(titleWrap);
        header.appendChild(close);

        const list = document.createElement('div');
        list.className = 'day-slot-popover__list';

        // Group events by operation type, sorted alphabetically
        const opGroups = new Map();
        events.forEach(ev => {
            const op = (ev.operation || 'Other').trim();
            if (!opGroups.has(op)) opGroups.set(op, []);
            opGroups.get(op).push(ev);
        });
        const sortedOps = [...opGroups.keys()].sort((a, b) => a.localeCompare(b));
        const hasMultipleOps = sortedOps.length > 1;

        sortedOps.forEach(op => {
            const opEvents = opGroups.get(op);

            if (hasMultipleOps) {
                const header = document.createElement('div');
                header.className = 'day-slot-popover__section-header';
                header.textContent = op;
                list.appendChild(header);
            }

            opEvents.forEach(ev => {
                const row = document.createElement('button');
                row.type = 'button';
                row.className = 'day-slot-popover__item';

                row.style.borderLeftColor = getTypeColor(ev.type);

                const icon = this._buildEventIcon(ev);
                if (icon) {
                    icon.classList.add('day-slot-popover__item-icon');
                    row.appendChild(icon);
                }

                const label = document.createElement('span');
                label.className = 'day-slot-popover__item-label';
                const svcRaw = this._eventLabel(ev);
                if (ev.context) {
                    const svc = svcRaw.length > 18 ? svcRaw.slice(0, 17) + '…' : svcRaw;
                    label.textContent = svc + ' · ' + ev.context;
                } else {
                    label.textContent = svcRaw;
                }

                const chevron = document.createElement('span');
                chevron.className = 'day-slot-popover__item-chevron';
                chevron.textContent = '›';

                row.appendChild(label);
                row.appendChild(chevron);

                row.addEventListener('click', (e) => {
                    e.stopPropagation();
                    this._hideSlotPopover();
                    this._onEventClick?.(ev);
                });

                list.appendChild(row);
            });
        });

        popover.appendChild(header);
        popover.appendChild(list);

        const container = anchorEl.closest('.day-timeline__canvas');
        container.appendChild(popover);
        this._slotPopover = popover;

        this._positionSlotPopover(popover, anchorEl);

        this._slotPopoverOutsideHandler = (e) => {
            if (
                this._slotPopover &&
                !this._slotPopover.contains(e.target) &&
                !anchorEl.contains(e.target)
            ) {
                this._hideSlotPopover();
            }
        };

        requestAnimationFrame(() => {
            document.addEventListener('mousedown', this._slotPopoverOutsideHandler);
        });
    }

    _positionSlotPopover(popover, anchorEl) {
        const container = anchorEl.closest('.day-timeline__canvas');
        const containerRect = container.getBoundingClientRect();
        const anchorRect = anchorEl.getBoundingClientRect();

        const offsetLeft = anchorRect.left - containerRect.left;
        const offsetTop = anchorRect.bottom - containerRect.top;

        popover.style.left = `${offsetLeft}px`;
        popover.style.top = `${offsetTop + 6}px`;
    }

    _hideSlotPopover() {
        if (this._slotPopoverOutsideHandler) {
            document.removeEventListener('mousedown', this._slotPopoverOutsideHandler);
            this._slotPopoverOutsideHandler = null;
        }

        if (this._slotPopover) {
            this._slotPopover.remove();
            this._slotPopover = null;
        }
    }

    _buildDenseSlotBadge(events) {
        if (!events?.length) return null;

        // If all grouped releases have the same SERVICE_OP environment,
        // show the environment badge, e.g. PRD.
        const serviceOps = events.filter(ev => ev.type === 'SERVICE_OP');
        const environments = new Set(serviceOps.map(ev => ev.environment).filter(Boolean));

        if (serviceOps.length === events.length && environments.size === 1) {
            const env = [...environments][0];
            const envLabel = getEnvLabel(env);
            const envColor = getEnvColor(env);

            const badge = document.createElement('span');
            badge.className = 'jenga-chip__env-badge day-timeline__stack-badge';
            badge.textContent = envLabel;
            badge.style.background = envColor;
            badge.style.color = envColor === '#eab308' ? '#1a1a1a' : '#fff';
            return badge;
        }

        // Otherwise show the dominant type.
        const type = this._dominantType(events);
        const badge = document.createElement('span');
        badge.className = 'day-timeline__stack-type-badge';
        badge.textContent = type.replace(/_/g, ' ');
        badge.style.background = getTypeColor(type);
        return badge;
    }

    /**
     * Assign column index and total column count to each timed event so that
     * overlapping events are laid out side-by-side (Outlook-style).
     *
     * Returns an array of { event, colIdx, colCount } objects.
     */
    _assignColumns(timedEvents) {
        if (!timedEvents.length) return [];

        // Sort by start time, then by end time descending (longer events first)
        const sorted = [...timedEvents].sort((a, b) => {
            const ds = a.timeSlot.startMin - b.timeSlot.startMin;
            if (ds !== 0) return ds;
            return b.timeSlot.endMin - a.timeSlot.endMin;
        });

        // Greedy column assignment: each column tracks its last endMin
        const columns = []; // columns[i] = endMin of last event in column i
        const assignments = sorted.map(event => {
            const { startMin } = event.timeSlot;
            // Find first column that has ended by this event's start
            let colIdx = columns.findIndex(endMin => endMin <= startMin);
            if (colIdx === -1) {
                colIdx = columns.length;
                columns.push(event.timeSlot.endMin);
            } else {
                columns[colIdx] = event.timeSlot.endMin;
            }
            return { event, colIdx, colCount: 1 }; // colCount filled in second pass
        });

        // Second pass: for each event, find all events that overlap with it
        // and compute the max column count within that overlap group.
        // We propagate colCount = max(columns needed) to all events in each cluster.
        const totalCols = columns.length;

        // Build overlap groups: events in the same "cluster" share the same colCount
        // A cluster is a contiguous chain where at least one event overlaps each successive event.
        let clusterStart = 0;
        let clusterEndMin = assignments[0].event.timeSlot.endMin;
        let clusterCols = assignments[0].colIdx + 1;

        const setClusterCount = (from, to, count) => {
            for (let i = from; i <= to; i++) assignments[i].colCount = count;
        };

        for (let i = 1; i < assignments.length; i++) {
            const ev = assignments[i].event;
            if (ev.timeSlot.startMin < clusterEndMin) {
                // Overlaps with cluster — extend cluster
                clusterEndMin = Math.max(clusterEndMin, ev.timeSlot.endMin);
                clusterCols   = Math.max(clusterCols, assignments[i].colIdx + 1);
            } else {
                // Gap — close previous cluster and start new one
                setClusterCount(clusterStart, i - 1, clusterCols);
                clusterStart  = i;
                clusterEndMin = ev.timeSlot.endMin;
                clusterCols   = assignments[i].colIdx + 1;
            }
        }
        setClusterCount(clusterStart, assignments.length - 1, clusterCols);

        return assignments;
    }

    _eventLabel(ev) {
        if (isCabProtectionEvent(ev)) {
            return 'Change Advisory Board';
        }
        if (ev.type === 'SERVICE_OP') {
            return (ev.service || ev.summary).trim();
        }
        if (ev.type === 'HYBRIS') {
            return (ev.service || ev.summary).replace(/^release\//, '');
        }
        if (ev.type === 'BUSINESS_EVENT') {
            return ev.summary.replace(/^Event:/, '').split('|')[0].trim();
        }
        return ev.summary.slice(0, 40);
    }

    /** Returns a DOM element representing the icon for the event, or null for no icon. */
    _buildEventIcon(ev) {
        if (isCabProtectionEvent(ev)) {
            const icon = document.createElement('span');
            icon.className = 'day-timeline__event-icon';
            icon.textContent = '📢';
            return icon;
        }
        if (ev.type === 'SERVICE_OP') {
            const envLabel = getEnvLabel(ev.environment);
            const envColor = getEnvColor(ev.environment);
            const badge = document.createElement('span');
            badge.className = 'jenga-chip__env-badge';
            badge.textContent = envLabel;
            badge.style.background = envColor;
            badge.style.color = envColor === '#eab308' ? '#1a1a1a' : '#fff';
            return badge;
        }
        if (ev.type === 'BUSINESS_EVENT') {
            const icon = document.createElement('span');
            icon.className = 'day-timeline__event-icon';
            icon.textContent = '👜';
            return icon;
        }
        if (ev.type === 'HYBRIS') {
            const icon = document.createElement('span');
            icon.className = 'day-timeline__event-icon';
            icon.innerHTML = HYBRIS_SVG;
            return icon;
        }
        if (ev.type === 'MILESTONE') {
            const icon = document.createElement('span');
            icon.className = 'day-timeline__event-icon';
            icon.textContent = '🏁';
            return icon;
        }
        return null;
    }
}
