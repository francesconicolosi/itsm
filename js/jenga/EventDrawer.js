import { getTypeColor, getEnvColor, getEnvLabel } from './CalendarRenderer.js';
import { createFormattedLongTextElementsFrom } from '../shared/utils.js';

const STATUS_COLORS = {
    'Done':         { bg: '#22c55e', text: '#fff' },
    'In Progress':  { bg: '#3b82f6', text: '#fff' },
    'Implementing': { bg: '#8b5cf6', text: '#fff' },
    'Ready':        { bg: '#06b6d4', text: '#fff' },
    'Backlog':      { bg: '#94a3b8', text: '#fff' },
    'Options':      { bg: '#f59e0b', text: '#1a1a1a' },
    "WON'T DO":     { bg: '#ef4444', text: '#fff' },
};

export function getStatusStyle(status) {
    return STATUS_COLORS[status] ?? { bg: '#6b7280', text: '#fff' };
}

const TYPE_LABELS = {
    SERVICE_OP:     'Service Operation',
    DELIVERY:       'Delivery Release',
    BUSINESS_EVENT: 'Business Event',
    HYBRIS:         'Hybris Release',
    MILESTONE:      'Milestone',
    OTHER:          'Other',
};

export class EventDrawer {
    constructor(app) {
        this.app = app;
        this._fromDay = false; // true when opened from DayDrawer (level 2 context)
    }

    /**
     * Open event detail.
     * @param {object} event
     * @param {boolean} [fromDay=false] — when true, show as level-2 panel over DayDrawer
     * @param {object[]|null} [groupedEvents=null] — all events in a grouped chip
     */
    open(event, fromDay = false, groupedEvents = null) {
        this._fromDay = fromDay;
        const title   = document.getElementById('drawer-title');
        const content = document.getElementById('drawerContent');
        if (!content) return;

        const friendlyTitle = this._eventTitle(event);
        title.textContent = friendlyTitle;

        content.innerHTML = '';
        content.appendChild(this._buildContent(event, groupedEvents));

        const drawer  = document.getElementById('drawer');
        const panels  = drawer?.querySelector('.drawer-panels');
        if (fromDay) {
            // Slide the strip left: L1 exits left, L2 enters from right
            drawer?.classList.remove('drawer--direct', 'drawer--direct-only');
            panels?.classList.add('panels-level2');
        } else {
            // Direct open: jump straight to L2 with no transition, hide Back
            drawer?.classList.add('open', 'drawer--direct', 'drawer--direct-only');
            document.getElementById('overlay')?.classList.add('open');
            panels?.classList.add('panels-level2');
        }
    }

    close() {
        const drawer = document.getElementById('drawer');
        const panels = drawer?.querySelector('.drawer-panels');
        if (!this._fromDay) {
            // Direct open: close entire drawer and reset strip position
            drawer?.classList.remove('open', 'drawer--direct', 'drawer--direct-only');
            document.getElementById('overlay')?.classList.remove('open');
            panels?.classList.remove('panels-level2');
        } else {
            // From day drawer: slide strip back right — L1 re-enters, L2 exits
            panels?.classList.remove('panels-level2');
        }
    }

    initEvents() {
        const shareBtn = document.getElementById('drawerShareBtn');
        if (shareBtn) {
            shareBtn.addEventListener('click', () => {
                navigator.clipboard?.writeText(location.href).catch(() => {});
                shareBtn.textContent = '✔';
                setTimeout(() => { shareBtn.textContent = '🔗'; }, 1500);
                this.app.showToast('Link copied to clipboard');
            });
        }
    }

    _buildContent(ev, groupedEvents = null) {
        const isGrouped = groupedEvents && groupedEvents.length > 1;
        const frag = document.createDocumentFragment();

        // Type badge
        const badge = document.createElement('div');
        badge.className = 'jenga-drawer__badge';
        badge.style.background = getTypeColor(ev.type);
        badge.textContent = TYPE_LABELS[ev.type] || ev.type;
        frag.appendChild(badge);

        // Meta rows
        const _fmtDate = d => d ? d.toLocaleDateString('en-GB') : null;
        const peakDatesStr = ev.type === 'BUSINESS_EVENT' && ev.peakDates && ev.peakDates.length
            ? ev.peakDates.map(r => {
                const s = _fmtDate(r.start);
                const e = _fmtDate(r.end);
                return (s && e && s !== e) ? `${s} – ${e}` : (s || e);
            }).filter(Boolean).join(', ')
            : null;

        // For grouped chips, Key and Context are rendered as multi-link fields — skip plain text versions.
        const meta = [
            !isGrouped && ['Key',      ev.key],
            ['Status',   ev.status],
            ev.priority && ev.priority !== 'None' && ['Priority', ev.priority],
            ['Due Date', ev.dueDate ? ev.dueDate.toLocaleDateString('en-GB') : '—'],
            ev.startDate && ev.startDate.getTime() !== ev.dueDate?.getTime() &&
                ['Start', ev.startDate.toLocaleDateString('en-GB')],
            peakDatesStr && ['🔥 Peak Dates', peakDatesStr],
            ev.type === 'SERVICE_OP' && ev.environment && ['Environment', ev.environment],
            ev.type === 'SERVICE_OP' && ev.service     && ['Service',     ev.service],
            ev.type === 'SERVICE_OP' && ev.operation   && ['Operation',   ev.operation],
            !isGrouped && ev.type === 'SERVICE_OP' && ev.context && ['Context', ev.context],
            !isGrouped && ev.assignee  && ['Assignee', ev.assignee],
            !isGrouped && ev.reporter  && ['Reporter', ev.reporter],
            ev.stream    && ['Stream',   ev.stream],
            ev.theme     && ['Theme',    ev.theme],
            !isGrouped   && ['Summary',  ev.summary],
            ev.roi       && ['Region',   ev.roi],
        ].filter(Boolean);

        const table = document.createElement('dl');
        table.className = 'jenga-drawer__meta';

        // Grouped chip: render Key and Context rows with per-item Jira links
        if (isGrouped) {
            let unknownCount = 0;

            // Key row — all keys as individual Jira links
            const dtKey = document.createElement('dt');
            dtKey.textContent = 'Key';
            const ddKey = document.createElement('dd');
            ddKey.className = 'jenga-drawer__meta-links';
            groupedEvents.forEach(e => {
                const a = document.createElement('a');
                a.href = e.jiraUrl ? e.jiraUrl.trim() : '#';
                a.target = '_blank';
                a.rel = 'noopener';
                a.className = 'jenga-drawer__meta-link';
                a.textContent = e.key;
                ddKey.appendChild(a);
            });
            table.append(dtKey, ddKey);

            // Context row — buttons that open single-event detail in-place
            const dtCtx = document.createElement('dt');
            dtCtx.textContent = 'Context';
            const ddCtx = document.createElement('dd');
            ddCtx.className = 'jenga-drawer__meta-links';
            groupedEvents.forEach(e => {
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'jenga-drawer__meta-link jenga-drawer__meta-link--btn';
                btn.textContent = e.context || `Unknown_${++unknownCount}`;
                btn.addEventListener('click', () => this._openSingleFromGroup(e));
                ddCtx.appendChild(btn);
            });
            table.append(dtCtx, ddCtx);
        }

        meta.forEach(([label, value]) => {
            if (!value) return;
            const dt = document.createElement('dt');
            dt.textContent = label;
            const dd = document.createElement('dd');
            if (label === 'Environment') {
                const badge = document.createElement('span');
                badge.className = 'jenga-chip__env-badge';
                badge.textContent = getEnvLabel(value);
                badge.style.background = getEnvColor(value);
                badge.style.color = getEnvColor(value) === '#eab308' ? '#1a1a1a' : '#fff';
                dd.appendChild(badge);
                dd.appendChild(document.createTextNode(' ' + value));
            } else if (label === 'Status') {
                const { bg, text } = getStatusStyle(value);
                const sb = document.createElement('span');
                sb.className = 'jenga-status-badge';
                sb.textContent = value;
                sb.style.background = bg;
                sb.style.color = text;
                dd.appendChild(sb);
            } else if (label === 'Summary') {
                dd.className = 'jenga-drawer__meta-summary';
                dd.textContent = value;
            } else {
                dd.textContent = value;
            }
            table.append(dt, dd);
        });
        frag.appendChild(table);

        // Description — omitted for grouped chips
        if (!isGrouped && ev.description) {
            frag.appendChild(this._section('Description', ev.description));
        }

        // Latest comment
        if (ev.latestComment) {
            frag.appendChild(this._commentSection(ev.latestComment));
        }

        // Domino + Solitaire links for SERVICE_OP
        if (ev.type === 'SERVICE_OP' && ev.service) {
            const svc = encodeURIComponent(ev.service);
            const ctaRow = document.createElement('div');
            ctaRow.className = 'jenga-drawer__app-links';

            const dominoLink = document.createElement('a');
            dominoLink.href = `./domino.html?search=id:"${svc}"`;
            dominoLink.target = '_blank';
            dominoLink.rel = 'noopener';
            dominoLink.className = 'jenga-drawer__app-link';
            dominoLink.innerHTML = '⚅ View in Domino';

            const solitaireLink = document.createElement('a');
            solitaireLink.href = `./solitaire.html?search=service:"${svc}"`;
            solitaireLink.target = '_blank';
            solitaireLink.rel = 'noopener';
            solitaireLink.className = 'jenga-drawer__app-link';
            solitaireLink.innerHTML = '♤ Team in Solitaire';

            ctaRow.append(dominoLink, solitaireLink);
            frag.appendChild(ctaRow);
        }

        // Jira link
        if (isGrouped) {
            const keys = groupedEvents.map(e => e.key).join(',');
            const jql = `issue in (${keys})`;
            const link = document.createElement('a');
            link.href = `https://brand.atlassian.net/issues/?jql=${encodeURIComponent(jql)}`;
            link.target = '_blank';
            link.rel = 'noopener';
            link.className = 'jenga-drawer__jira-link';
            link.textContent = '↗ Open all in Jira';
            frag.appendChild(link);
        } else if (ev.jiraUrl) {
            const link = document.createElement('a');
            link.href = ev.jiraUrl.trim();
            link.target = '_blank';
            link.rel = 'noopener';
            link.className = 'jenga-drawer__jira-link';
            link.textContent = '↗ Open in Jira';
            frag.appendChild(link);
        }

        return frag;
    }

    _openSingleFromGroup(ev) {
        const title   = document.getElementById('drawer-title');
        const content = document.getElementById('drawerContent');
        if (!content) return;
        if (title) title.textContent = this._eventTitle(ev);
        content.innerHTML = '';
        content.appendChild(this._buildContent(ev, null));
        this.app?._pushUrl?.({ event: ev.key }, true);
    }

    _section(title, text) {
        const wrap = document.createElement('div');
        wrap.className = 'jenga-drawer__section';
        const heading = document.createElement('div');
        heading.className = 'jenga-drawer__section-title';
        heading.textContent = title;
        const body = document.createElement('div');
        body.className = 'jenga-drawer__section-body';
        createFormattedLongTextElementsFrom(text).forEach(el => body.appendChild(el));
        wrap.append(heading, body);
        return wrap;
    }

    _commentSection(text) {
        const wrap = document.createElement('div');
        wrap.className = 'jenga-drawer__section';
        const heading = document.createElement('div');
        heading.className = 'jenga-drawer__section-title';
        heading.textContent = 'Latest Comment';

        // Format: "Author||body segments…"
        // Split on first || only to isolate the author name from the body.
        const firstSep = text.indexOf('||');
        const hasAuthor = firstSep !== -1;
        const author = hasAuthor ? text.slice(0, firstSep).trim() : '';
        const body   = hasAuthor ? text.slice(firstSep + 2) : text;

        const bodyEl = document.createElement('div');
        bodyEl.className = 'jenga-drawer__section-body';

        if (author) {
            const authorEl = document.createElement('strong');
            authorEl.className = 'jenga-drawer__comment-author';
            authorEl.textContent = author;
            bodyEl.appendChild(authorEl);
        }

        createFormattedLongTextElementsFrom(body).forEach(el => bodyEl.appendChild(el));
        wrap.append(heading, bodyEl);
        return wrap;
    }

    showAbout() {
        const title   = document.getElementById('drawer-title');
        const content = document.getElementById('drawerContent');
        if (!content) return;
        if (title) title.textContent = 'About Jenga 🧱';
        content.innerHTML = `
<div class="about-content">
  <p class="about-intro">
    <strong>Jenga</strong> is the Service Operations Calendar — a unified view of all planned changes,
    releases, milestones, and business events across Nycosoft Digital IT services.
  </p>
  <div class="about-section">
    <h4>What it shows</h4>
    <ul>
      <li>Service Operations (maintenance windows, deployments)</li>
      <li>Delivery Releases and Hybris Releases</li>
      <li>Milestones (iOS, Android, SAP releases, …)</li>
      <li>Business Events impacting the platform</li>
    </ul>
  </div>
  <div class="about-section">
    <h4>How to use it</h4>
    <ul>
      <li>Navigate weeks with the arrow buttons or jump to today</li>
      <li>Filter by Environment, Type, and Service using the filter bar</li>
      <li>Click any event to see full details in this panel</li>
      <li>Click a day cell to open a detailed day timeline view</li>
      <li>Use the search bar to find events by name or service</li>
    </ul>
  </div>
</div>`;
        const drawer  = document.getElementById('drawer');
        const overlay = document.getElementById('overlay');
        drawer?.classList.add('open', 'drawer--direct', 'drawer--direct-only');
        overlay?.classList.add('open');
        drawer?.querySelector('.drawer-panels')?.classList.add('panels-level2');
    }

    /** Human-readable title matching what the calendar chip displays */
    _eventTitle(ev) {
        if (ev.type === 'SERVICE_OP') {
            const env = getEnvLabel(ev.environment);
            const svc = ev.service || '';
            const op  = ev.operation || '';
            return [env, svc, op].filter(Boolean).join(' · ');
        }
        if (ev.type === 'BUSINESS_EVENT') {
            return ev.summary.replace(/^Event:/, '').split('|')[0].trim();
        }
        if (ev.type === 'HYBRIS') {
            return (ev.service || ev.summary).replace(/^release\//, '');
        }
        if (ev.type === 'MILESTONE') {
            return ev.summary;
        }
        // DELIVERY / OTHER: strip trailing pipe segments
        return ev.summary.split('|')[0].trim() || ev.summary;
    }
}
