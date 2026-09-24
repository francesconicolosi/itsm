import {
    createFormattedLongTextElementsFrom,
    createHrefElement,
    createOutlookUrl,
    normalizeWs,
    truncateString,
} from '../shared/utils.js';

function _dominoAccentIconSrc() {
    const t = document.documentElement?.getAttribute('data-theme');
    const dark = t === 'dark' || (!t && window.matchMedia?.('(prefers-color-scheme: dark)')?.matches);
    return dark ? 'assets/domino-accent.svg' : 'assets/domino-accent-light.svg';
}

export class TeamDetailDrawer {
    constructor(app) {
        this.app = app;
        this._currentPermalink = null;
        this._showDetails = false;
    }

    open({
        name: title,
        description,
        elements,
        channels,
        email,
        highlightService,
        highlightQuery,
        elementsTitle = null,
        elementsBaseUrl,
        _permalinkSearch = '',
        _showDetails = false,
        _openServices = false,
    }) {
        if (this.app.interaction.isDraggable) return;

        const drawer = document.getElementById('drawer');
        const overlay = document.getElementById('drawer-overlay');

        if (!drawer) {
            console.warn('[drawer] #drawer not found');
            return;
        }

        let titleEl = document.getElementById('drawer-title');
        if (!titleEl) {
            titleEl = document.createElement('h2');
            titleEl.id = 'drawer-title';
            drawer.prepend(titleEl);
        }

        let descEl = document.getElementById('drawer-description');
        if (!descEl) {
            descEl = document.createElement('div');
            descEl.id = 'drawer-description';
            drawer.appendChild(descEl);
        }

        let listEl = document.getElementById('drawer-list');
        if (!listEl) {
            listEl = document.createElement('ul');
            listEl.id = 'drawer-list';
            drawer.appendChild(listEl);
        }

        titleEl.textContent = `${title ?? ''}`;

        this._currentPermalink = _permalinkSearch || null;
        this._showDetails = _showDetails;
        const copyBtn = document.getElementById('drawerCopyLink');
        if (copyBtn) copyBtn.style.display = _permalinkSearch ? '' : 'none';
        const searchBtn = document.getElementById('drawerSearchBtn');
        if (searchBtn) searchBtn.style.display = _permalinkSearch ? '' : 'none';

        descEl.replaceChildren();
        listEl.replaceChildren();

        const accordion = document.createElement('div');
        accordion.className = 'drawer-accordion';
        descEl.appendChild(accordion);

        const addDrawerSection = (label, fillFn, { open = false, sectionId = '' } = {}) => {
            const details = document.createElement('details');
            details.className = 'drawer-section';
            if (open) details.open = true;
            if (sectionId) details.dataset.sectionId = sectionId;

            const summary = document.createElement('summary');
            summary.className = 'drawer-section__summary';
            summary.innerHTML = label;

            const body = document.createElement('div');
            body.className = 'drawer-section__body';

            details.appendChild(summary);
            details.appendChild(body);
            accordion.appendChild(details);

            if (typeof fillFn === 'function') fillFn(body, details);

            return { details, body };
        };

        if (description) {
            addDrawerSection('Overview', (body) => {
                createFormattedLongTextElementsFrom(description).forEach(el => body.appendChild(el));
            }, { open: true, sectionId: 'overview' });
        }

        if (channels && channels.length > 0) {
            addDrawerSection('Channels <span class="drawer-svc-icon">💬️</span>', (body) => {
                const ul = document.createElement('ul');
                channels.forEach(channel => {
                    const li = document.createElement('li');
                    const channelLink = createHrefElement(
                        channel,
                        channel?.includes('slack.com') ? 'Slack Channel' : 'Link'
                    );
                    li.appendChild(channelLink);
                    ul.appendChild(li);
                });
                body.appendChild(ul);
            }, { open: false, sectionId: 'channels' });
        }

        if (email && email !== '') {
            addDrawerSection('Team Mailbox <span class="drawer-svc-icon">✉️</span>', (body) => {
                body.appendChild(
                    createHrefElement(createOutlookUrl([email]), `${truncateString(email, 25)}`)
                );
            }, { open: false, sectionId: 'mailbox' });
        }

        if (elements && elements.items && elements.items.length > 0) {
            const dominoUrl = `./domino.html?search=${encodeURIComponent(`Responsible Teams:${title}`)}&listView=${encodeURIComponent('ID,Description,Type,Depends on,Status,Decommission Date')}`;
            const dominoIconHtml = `<a href="${dominoUrl}" target="_blank" rel="noopener noreferrer" class="drawer-domino-link" data-tooltip="View all services in Domino" data-tooltip-placement="left" onclick="event.stopPropagation()"><img src="${_dominoAccentIconSrc()}" class="drawer-svc-icon" aria-hidden="true"></a>`;
            const resolvedElementsTitle = elementsTitle ?? `Managed Services ${dominoIconHtml}`;
            const shouldOpenServices = _openServices || !!(highlightService || (highlightQuery && highlightQuery.trim()));

            addDrawerSection(resolvedElementsTitle, (body) => {
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
                    if (this.app.jengaServicesThisMonth?.has(s.toLowerCase())) {
                        const now = new Date();
                        const jengaHref = `jenga.html?view=timeline&year=${now.getFullYear()}&month=${now.getMonth() + 1}&services=${encodeURIComponent(s)}`;
                        const iconLink = document.createElement('a');
                        iconLink.href = jengaHref;
                        iconLink.target = '_blank';
                        iconLink.rel = 'noopener noreferrer';
                        iconLink.className = 'jenga-link-icon';
                        iconLink.setAttribute('aria-label', `View ${s} releases timeline`);
                        iconLink.setAttribute('data-tooltip', 'View releases timeline');
                        iconLink.setAttribute('data-tooltip-placement', 'bottom');
                        const img = document.createElement('img');
                        img.src = './assets/jenga.svg';
                        img.alt = '';
                        img.setAttribute('aria-hidden', 'true');
                        iconLink.appendChild(img);
                        li.appendChild(iconLink);
                    }
                    frag.appendChild(li);
                });

                listEl.replaceChildren(frag);
                body.appendChild(listEl);

                (function multiHighlight() {
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

        drawer.classList.add('open');
        overlay?.classList.add('visible');
        document.body.classList.add('drawer-open');
        drawer.setAttribute('aria-hidden', 'false');
    }

    close() {
        const drawer = document.getElementById('drawer');
        const overlay = document.getElementById('drawer-overlay');
        if (!drawer) return;
        drawer.classList.remove('open');
        overlay?.classList.remove('visible');
        document.body.classList.remove('drawer-open');
        drawer.setAttribute('aria-hidden', 'true');
    }

    initEvents() {
        const overlay = document.getElementById('drawer-overlay');
        const closeBtn = document.getElementById('drawer-close');
        const header = document.querySelector('#drawer header');
        if (header && closeBtn) {
            const searchBtn = document.createElement('button');
            searchBtn.id = 'drawerSearchBtn';
            searchBtn.setAttribute('aria-label', 'Search this team');
            searchBtn.setAttribute('data-tooltip', 'Search this team');
            searchBtn.title = '';
            searchBtn.innerHTML = '<svg width="1em" height="1em" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>';
            searchBtn.style.display = 'none';
            searchBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                const query = this._currentPermalink;
                if (!query) return;
                const inp = document.getElementById('drawer-search-input');
                if (inp) inp.value = query;
                this.app.search._refreshChips(query);
                this.app.search.search(query);
            });
            header.insertBefore(searchBtn, closeBtn);

            const copyBtn = document.createElement('button');
            copyBtn.id = 'drawerCopyLink';
            copyBtn.setAttribute('aria-label', 'Copy link');
            copyBtn.setAttribute('data-tooltip', 'Copy link');
            copyBtn.title = '';
            copyBtn.innerHTML = '<svg width="1em" height="1em" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 7h3a5 5 0 0 1 0 10h-3"/><path d="M9 17H6A5 5 0 0 1 6 7h3"/><line x1="8" y1="12" x2="16" y2="12"/></svg>';
            copyBtn.style.display = 'none';
            copyBtn.addEventListener('click', async (e) => {
                e.stopPropagation();
                if (!this._currentPermalink) return;
                const url = new URL(window.location.href);
                url.searchParams.set('search', this._currentPermalink);
                if (this._showDetails) url.searchParams.set('showDetails', 'true');
                try { await navigator.clipboard.writeText(url.toString()); }
                catch {
                    const ta = document.createElement('textarea');
                    ta.value = url.toString();
                    ta.style.cssText = 'position:fixed;top:-9999px;left:-9999px';
                    document.body.appendChild(ta); ta.select();
                    document.execCommand('copy'); ta.remove();
                }
                this.app.showToast('Link copied to clipboard');
            });
            header.insertBefore(copyBtn, closeBtn);
        }
        overlay?.addEventListener('click', () => this.close());
        closeBtn?.addEventListener('click', () => this.close());
    }
}
