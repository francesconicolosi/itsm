function _priorityClass(priority) {
    const m = (priority || '').match(/^(P[1-4])/i);
    return m ? m[1].toLowerCase() : '';
}

function _priorityOrder(p) {
    const m = (p || '').match(/^P([1-4])/i);
    return m ? parseInt(m[1]) : 99;
}

export class CardsDrawer {
    constructor() {
        this._open = false;
    }

    initEvents() {
        document.getElementById('cards-drawer-close')
            ?.addEventListener('click', () => this.close());
        document.getElementById('cards-overlay')
            ?.addEventListener('click', () => this.close());
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && this._open) this.close();
        });
    }

    open(date, issueTypeLabel, cards) {
        const title   = document.getElementById('cards-drawer-title');
        const content = document.getElementById('cards-drawer-content');
        if (!content) return;

        const dateFmt = date.toLocaleDateString('en-US', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
        const dateEl  = document.getElementById('cards-drawer-date');
        if (dateEl) dateEl.textContent = dateFmt;
        title.textContent = "List of " + issueTypeLabel + "s";

        content.innerHTML = '';

        if (!cards.length) {
            const empty = document.createElement('p');
            empty.className = 'cards-drawer__empty';
            empty.textContent = 'No cards found for this day.';
            content.appendChild(empty);
        } else {
            const list = document.createElement('ul');
            list.className = 'cards-drawer__list';
            const sorted = [...cards].sort((a, b) => _priorityOrder(a.priority) - _priorityOrder(b.priority));
            sorted.forEach(c => list.appendChild(this._buildCard(c)));
            content.appendChild(list);
        }

        const drawer  = document.getElementById('cards-drawer');
        const overlay = document.getElementById('cards-overlay');
        drawer?.classList.add('open');
        drawer?.setAttribute('aria-hidden', 'false');
        overlay?.classList.add('open');
        this._open = true;
    }

    close() {
        const drawer  = document.getElementById('cards-drawer');
        const overlay = document.getElementById('cards-overlay');
        drawer?.classList.remove('open');
        drawer?.setAttribute('aria-hidden', 'true');
        overlay?.classList.remove('open');
        this._open = false;
    }

    _buildCard(card) {
        const li = document.createElement('li');
        li.className = 'cards-drawer__card';
        const pc = _priorityClass(card.priority);
        if (pc === 'p1') li.classList.add('cards-drawer__card--p1');
        else if (pc === 'p2') li.classList.add('cards-drawer__card--p2');

        // Key → Jira link
        const header = document.createElement('div');
        header.className = 'cards-drawer__card-header';
        const portalUrl = card.jiraUrl ? card.jiraUrl.replace("browse", "servicedesk/customer/portal/1") : '#';

        const keyLink = document.createElement('a');
        keyLink.className = 'cards-drawer__key';
        keyLink.href      = portalUrl
        keyLink.target    = '_blank';
        keyLink.rel       = 'noopener';
        keyLink.textContent = card.key;
        header.appendChild(keyLink);

        // Priority badge
        if (card.priority) {
            const pBadge = document.createElement('span');
            pBadge.className = 'cards-drawer__priority';
            const pc = _priorityClass(card.priority);
            if (pc) pBadge.classList.add(`cards-drawer__priority--${pc}`);
            pBadge.textContent = card.priority;
            header.appendChild(pBadge);
        }

        // Status badge
        if (card.status) {
            const badge = document.createElement('span');
            badge.className = 'cards-drawer__status';
            badge.textContent = card.status;
            header.appendChild(badge);
        }

        li.appendChild(header);

        // Summary
        const summary = document.createElement('p');
        summary.className = 'cards-drawer__summary';
        summary.textContent = card.summary;
        li.appendChild(summary);

        // Affected Services + ROI badge on the same row
        const services = (card.affectedServices || '')
            .split('||')
            .map(s => s.trim())
            .filter(Boolean);

        if (services.length || card.roi) {
            const svcRow = document.createElement('div');
            svcRow.className = 'cards-drawer__services';
            if (card.roi) {
                const badge = document.createElement('span');
                badge.className = 'jenga-roi-badge';
                badge.textContent = `ROI: ${card.roi}`;
                svcRow.appendChild(badge);
            }
            services.forEach(svc => {
                const a = document.createElement('a');
                a.className = 'cards-drawer__service-link';
                a.href      = `./domino.html?search=id%3A"${encodeURIComponent(svc)}"`;
                a.target    = '_blank';
                a.rel       = 'noopener';
                a.textContent = svc;
                svcRow.appendChild(a);
            });
            li.appendChild(svcRow);
        }

        return li;
    }
}
