import {
    createFormattedLongTextElementsFrom,
    createHrefElement,
    createOutlookUrl,
    truncateString,
    formatMonthYear,
} from '../shared/utils.js';
import { emailField } from './constants.js';

export class PersonDetailDrawer {
    constructor(app) {
        this.app = app;
        this._currentMember = null;
        this._currentL2Permalink = null;
    }

    open(memberData) {
        if (this.app.interaction.isDraggable) return;

        // Dismiss TeamDetailDrawer if open — only one content drawer at a time
        const existingDrawer = document.getElementById('drawer');
        if (existingDrawer?.classList.contains('open')) {
            this.app.drawer.close();
        }

        this._currentMember = memberData;

        const drawer = document.getElementById('person-drawer');
        const overlay = document.getElementById('person-drawer-overlay');
        if (!drawer) {
            console.warn('[PersonDetailDrawer] #person-drawer not found');
            return;
        }

        // Always start at L1
        drawer.querySelector('.person-drawer-panels')?.classList.remove('panels-level2');

        document.getElementById('person-drawer-title').textContent = memberData['Name'] || '';

        const contentEl = document.getElementById('person-drawer-content');
        contentEl.replaceChildren();
        this._buildPersonContent(memberData, contentEl);

        drawer.classList.add('open');
        overlay?.classList.add('visible');
        document.body.classList.add('person-drawer-open');
        drawer.setAttribute('aria-hidden', 'false');
    }

    close() {
        const drawer = document.getElementById('person-drawer');
        const overlay = document.getElementById('person-drawer-overlay');
        if (!drawer) return;
        drawer.classList.remove('open');
        overlay?.classList.remove('visible');
        document.body.classList.remove('person-drawer-open');
        drawer.setAttribute('aria-hidden', 'true');
        drawer.querySelector('.person-drawer-panels')?.classList.remove('panels-level2');
    }

    openL2(type, payload) {
        const panels = document.querySelector('#person-drawer .person-drawer-panels');
        const l2Content = document.getElementById('person-drawer-l2-content');
        const l2Title = document.getElementById('person-drawer-l2-title');
        if (!panels || !l2Content || !l2Title) return;

        l2Content.replaceChildren();

        const l2SearchBtn = document.getElementById('person-drawer-l2-search');

        if (type === 'team') {
            const { teamName } = payload;
            l2Title.textContent = teamName;
            this._currentL2Permalink = `team:"${teamName}"`;
            if (l2SearchBtn) l2SearchBtn.style.display = '';

            const teamTitleEl = Array.from(document.querySelectorAll('text.team-title'))
                .find(el => el.getAttribute('data-full-name') === teamName);

            if (teamTitleEl) {
                const description = teamTitleEl.getAttribute('data-team-description') || '';
                const email = teamTitleEl.getAttribute('data-team-email') || '';
                const channels = JSON.parse(teamTitleEl.getAttribute('data-team-channels') || '[]');
                const servicesStr = teamTitleEl.getAttribute('data-services') || '';
                const services = servicesStr ? servicesStr.split(', ').filter(Boolean) : [];
                this._buildTeamL2Content(l2Content, { description, email, channels, services });
            } else {
                const p = document.createElement('p');
                p.textContent = 'Team details not available.';
                l2Content.appendChild(p);
            }
        } else if (type === 'role') {
            const { name, description, grants } = payload;
            l2Title.textContent = name;
            this._currentL2Permalink = null;
            if (l2SearchBtn) l2SearchBtn.style.display = 'none';
            this._buildDescriptionL2Content(l2Content, { description, extra: grants ? `Grants: ${grants}` : null });
        } else if (type === 'function') {
            const { name, description } = payload;
            l2Title.textContent = name;
            this._currentL2Permalink = null;
            if (l2SearchBtn) l2SearchBtn.style.display = 'none';
            this._buildDescriptionL2Content(l2Content, { description });
        }

        panels.classList.add('panels-level2');
    }

    closeL2() {
        const drawer = document.getElementById('person-drawer');
        drawer?.querySelector('.person-drawer-panels')?.classList.remove('panels-level2');
        this._currentL2Permalink = null;
        const l2SearchBtn = document.getElementById('person-drawer-l2-search');
        if (l2SearchBtn) l2SearchBtn.style.display = 'none';
    }

    initEvents() {
        document.getElementById('person-drawer-overlay')
            ?.addEventListener('click', () => this.close());
        document.getElementById('person-drawer-close')
            ?.addEventListener('click', () => this.close());
        document.getElementById('person-drawer-l2-close')
            ?.addEventListener('click', () => this.close());
        document.getElementById('person-drawer-back')
            ?.addEventListener('click', () => this.closeL2());
        document.getElementById('person-drawer-link')
            ?.addEventListener('click', () => this.copyLink());
        document.getElementById('person-drawer-l2-search')
            ?.addEventListener('click', (e) => {
                e.stopPropagation();
                const query = this._currentL2Permalink;
                if (!query) return;
                const inp = document.getElementById('drawer-search-input');
                if (inp) inp.value = query;
                this.app.search._refreshChips(query);
                this.app.search.search(query, { keepDrawer: true });
            });
    }

    copyLink() {
        const email = (this._currentMember?.[emailField] || '').trim();
        if (!email) return;
        const url = window.location.origin + window.location.pathname + '?person=' + encodeURIComponent(email);
        navigator.clipboard.writeText(url).then(() => {
            this.app.showToast('Link copied to clipboard');
        }).catch(() => {
            this.app.showToast('Could not copy link');
        });
    }

    _buildPersonContent(memberData, container) {
        this._buildPhotoHeader(memberData, container);

        const accordion = document.createElement('div');
        accordion.className = 'drawer-accordion';
        container.appendChild(accordion);

        const addSection = (label, fillFn, { open = false } = {}) => {
            const details = document.createElement('details');
            details.className = 'drawer-section';
            if (open) details.open = true;
            const summary = document.createElement('summary');
            summary.className = 'drawer-section__summary';
            summary.textContent = label;
            const body = document.createElement('div');
            body.className = 'drawer-section__body';
            details.appendChild(summary);
            details.appendChild(body);
            accordion.appendChild(details);
            if (typeof fillFn === 'function') fillFn(body);
        };

        const roleName = memberData['Role'] || '';
        const functionName = memberData['Function'] || '';
        const company = memberData['Company'] || '';
        const location = memberData['Location'] || '';
        const room = memberData['Room'] || '';
        const email = memberData[emailField] || '';
        const inTeamSinceRaw = memberData['In team since'] || '';
        const inTeamSince = (() => {
            if (!inTeamSinceRaw) return '';
            const d = new Date(inTeamSinceRaw);
            return isNaN(d) ? inTeamSinceRaw : formatMonthYear(d);
        })();

        if (roleName || functionName || company || location || room || email || inTeamSince) {
            addSection('Identity', (body) => {
                if (roleName) {
                    const p = document.createElement('p');
                    const strong = document.createElement('strong');
                    strong.textContent = 'Role: ';
                    p.appendChild(strong);
                    const roleInfo = this.app.db.roleDetailsMapping.get(roleName) || {};
                    if (roleInfo.description) {
                        const link = document.createElement('span');
                        link.className = 'person-drawer-link';
                        link.textContent = roleName;
                        link.addEventListener('click', () => this.openL2('role', {
                            name: roleName,
                            description: roleInfo.description,
                            grants: roleInfo.grants,
                        }));
                        p.appendChild(link);
                    } else {
                        p.appendChild(document.createTextNode(roleName));
                    }
                    body.appendChild(p);
                }

                if (functionName) {
                    const p = document.createElement('p');
                    const strong = document.createElement('strong');
                    strong.textContent = 'Function: ';
                    p.appendChild(strong);
                    const fnInfo = this.app.db.functionDetailsMapping?.get(functionName) || {};
                    if (fnInfo.description) {
                        const link = document.createElement('span');
                        link.className = 'person-drawer-link';
                        link.textContent = functionName;
                        link.addEventListener('click', () => this.openL2('function', {
                            name: functionName,
                            description: fnInfo.description,
                        }));
                        p.appendChild(link);
                    } else {
                        p.appendChild(document.createTextNode(functionName));
                    }
                    body.appendChild(p);
                }

                [['Company', company], ['Location', location], ['Room', room]].forEach(([label, value]) => {
                    if (!value) return;
                    const p = document.createElement('p');
                    const strong = document.createElement('strong');
                    strong.textContent = `${label}: `;
                    p.appendChild(strong);
                    p.appendChild(document.createTextNode(value));
                    body.appendChild(p);
                });

                if (email) {
                    const p = document.createElement('p');
                    const strong = document.createElement('strong');
                    strong.textContent = 'Email: ';
                    p.appendChild(strong);
                    p.appendChild(createHrefElement(createOutlookUrl([email]), email));
                    body.appendChild(p);
                }
                if (inTeamSince) {
                    const p = document.createElement('p');
                    const strong = document.createElement('strong');
                    strong.textContent = 'In team since: ';
                    p.appendChild(strong);
                    p.appendChild(document.createTextNode(inTeamSince));
                    body.appendChild(p);
                }
            }, { open: true });
        }

        const teamName = memberData['Team member of'] || '';
        if (teamName) {
            addSection('Team', (body) => {
                const p = document.createElement('p');
                const strong = document.createElement('strong');
                strong.textContent = 'Team member of: ';
                p.appendChild(strong);
                const link = document.createElement('span');
                link.className = 'person-drawer-link';
                link.textContent = teamName;
                link.addEventListener('click', () => this.openL2('team', { teamName }));
                p.appendChild(link);
                body.appendChild(p);
            }, { open: true });
        }

        const guestAppearances = this._findGuestAppearances(email);
        if (guestAppearances.length > 0) {
            addSection('Other Roles', (body) => {
                guestAppearances.forEach(({ guestRole, teams }) => {
                    const label = this.app.db.guestRolesMap.get(guestRole)?.[0] ?? guestRole;
                    teams.forEach(team => {
                        const p = document.createElement('p');
                        const strong = document.createElement('strong');
                        strong.textContent = `${label} in: `;
                        p.appendChild(strong);
                        const link = document.createElement('span');
                        link.className = 'person-drawer-link';
                        link.textContent = team;
                        link.addEventListener('click', () => this.openL2('team', { teamName: team }));
                        p.appendChild(link);
                        body.appendChild(p);
                    });
                });
            }, { open: true });
        }
    }

    _buildPhotoHeader(memberData, container) {
        const email = memberData[emailField] || '';
        const name = memberData['Name'] || '';
        const baseName = (email.split('@')[0] || '').replace('-ext', '').replace('.', '-');
        const photoBase = `./assets/photos/${baseName}`;
        const candidates = [`${photoBase}.webp`, `${photoBase}.jpg`, `${photoBase}.png`, `${photoBase}.jpeg`];

        const initials = name.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase() || '?';
        const initialsEl = document.createElement('div');
        initialsEl.className = 'person-drawer-photo-initials';
        initialsEl.textContent = initials;

        const wrapper = document.createElement('div');
        wrapper.className = 'person-drawer-photo-header';

        const photoDiv = document.createElement('div');
        photoDiv.className = 'person-drawer-photo';

        if (baseName) {
            const img = document.createElement('img');
            img.alt = name;
            let tryIdx = 0;
            img.onerror = () => {
                tryIdx++;
                if (tryIdx < candidates.length) {
                    img.src = candidates[tryIdx];
                } else {
                    photoDiv.replaceChildren(initialsEl);
                }
            };
            img.src = candidates[0];
            photoDiv.appendChild(img);
        } else {
            photoDiv.appendChild(initialsEl);
        }

        wrapper.appendChild(photoDiv);
        container.appendChild(wrapper);
    }

    _findGuestAppearances(email) {
        if (!email || !this.app.visibleOrg) return [];
        const target = email.toLowerCase();
        const byRole = new Map();

        for (const themes of Object.values(this.app.visibleOrg)) {
            for (const teams of Object.values(themes)) {
                for (const [teamName, members] of Object.entries(teams)) {
                    for (const member of (members || [])) {
                        const memberEmail = (member[emailField] || '').toLowerCase();
                        if (memberEmail === target && member.guestRole) {
                            if (!byRole.has(member.guestRole)) byRole.set(member.guestRole, new Set());
                            byRole.get(member.guestRole).add(teamName);
                        }
                    }
                }
            }
        }

        return [...byRole.entries()].map(([guestRole, teams]) => ({ guestRole, teams: [...teams] }));
    }

    _buildTeamL2Content(container, { description, email, channels, services }) {
        const accordion = document.createElement('div');
        accordion.className = 'drawer-accordion';
        container.appendChild(accordion);

        const addSection = (label, fillFn, { open = false } = {}) => {
            const details = document.createElement('details');
            details.className = 'drawer-section';
            if (open) details.open = true;
            const summary = document.createElement('summary');
            summary.className = 'drawer-section__summary';
            summary.innerHTML = label;
            const body = document.createElement('div');
            body.className = 'drawer-section__body';
            details.appendChild(summary);
            details.appendChild(body);
            accordion.appendChild(details);
            if (typeof fillFn === 'function') fillFn(body);
        };

        if (description) {
            addSection('Overview', (body) => {
                createFormattedLongTextElementsFrom(description).forEach(el => body.appendChild(el));
            }, { open: true });
        }

        if (channels && channels.length > 0) {
            addSection('Channels <span class="drawer-svc-icon">💬️</span>', (body) => {
                const ul = document.createElement('ul');
                channels.forEach(channel => {
                    const li = document.createElement('li');
                    li.appendChild(createHrefElement(
                        channel,
                        channel?.includes('slack.com') ? 'Slack Channel' : 'Link'
                    ));
                    ul.appendChild(li);
                });
                body.appendChild(ul);
            });
        }

        if (email) {
            addSection('Team Mailbox <span class="drawer-svc-icon">✉️</span>', (body) => {
                body.appendChild(createHrefElement(createOutlookUrl([email]), truncateString(email, 25)));
            });
        }

        if (services.length > 0) {
            addSection('Managed Services', (body) => {
                const ul = document.createElement('ul');
                services.forEach(s => {
                    const li = document.createElement('li');
                    const a = document.createElement('a');
                    a.href = `domino.html?search=id%3A"${encodeURIComponent(s)}"`;
                    a.textContent = s;
                    a.target = '_blank';
                    li.appendChild(a);

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
                        const img = document.createElement('img');
                        img.src = './assets/jenga.svg';
                        img.alt = '';
                        img.setAttribute('aria-hidden', 'true');
                        iconLink.appendChild(img);
                        li.appendChild(iconLink);
                    }

                    ul.appendChild(li);
                });
                body.appendChild(ul);
            }, { open: true });
        }
    }

    _buildDescriptionL2Content(container, { description, extra = null }) {
        const accordion = document.createElement('div');
        accordion.className = 'drawer-accordion';
        container.appendChild(accordion);

        if (description) {
            const details = document.createElement('details');
            details.className = 'drawer-section';
            details.open = true;
            const summary = document.createElement('summary');
            summary.className = 'drawer-section__summary';
            summary.textContent = 'Description';
            const body = document.createElement('div');
            body.className = 'drawer-section__body';
            details.appendChild(summary);
            details.appendChild(body);
            accordion.appendChild(details);
            createFormattedLongTextElementsFrom(description).forEach(el => body.appendChild(el));
        }

        if (extra) {
            const details = document.createElement('details');
            details.className = 'drawer-section';
            const summary = document.createElement('summary');
            summary.className = 'drawer-section__summary';
            summary.textContent = 'Grants';
            const body = document.createElement('div');
            body.className = 'drawer-section__body';
            details.appendChild(summary);
            details.appendChild(body);
            accordion.appendChild(details);
            const p = document.createElement('p');
            p.textContent = extra;
            body.appendChild(p);
        }
    }
}
