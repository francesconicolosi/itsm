import {
    createFormattedLongTextElementsFrom,
    createHrefElement,
    createOutlookUrl,
    normalizeWs,
    truncateString,
} from '../shared/utils.js';

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
        elementsTitle = 'Managed Services ⚙️',
        elementsBaseUrl,
        _permalinkSearch = '',
        _showDetails = false
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
            summary.textContent = label;

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
            addDrawerSection('Channels 💬', (body) => {
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
            addDrawerSection('Team Mailbox ✉️', (body) => {
                body.appendChild(
                    createHrefElement(createOutlookUrl([email]), `${truncateString(email, 25)}`)
                );
            }, { open: false, sectionId: 'mailbox' });
        }

        if (elements && elements.items && elements.items.length > 0) {
            const shouldOpenServices = !!(highlightService || (highlightQuery && highlightQuery.trim()));

            addDrawerSection(elementsTitle, (body) => {
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
            const copyBtn = document.createElement('button');
            copyBtn.id = 'drawerCopyLink';
            copyBtn.setAttribute('aria-label', 'Copy link');
            copyBtn.setAttribute('data-tooltip', 'Copy link');
            copyBtn.textContent = '🔗';
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
                copyBtn.textContent = '✔';
                setTimeout(() => { copyBtn.textContent = '🔗'; }, 1500);
                this.app.showToast('Link copied to clipboard');
            });
            header.insertBefore(copyBtn, closeBtn);
        }
        overlay?.addEventListener('click', () => this.close());
        closeBtn?.addEventListener('click', () => this.close());
    }
}
