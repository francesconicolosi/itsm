export class MajorIncidentAlert {
    show(incidents) {
        if (!incidents || incidents.length === 0) return;

        const overlay = document.createElement('div');
        overlay.id = 'major-incident-overlay';
        const maxSeverity = incidents.some(c => (c.priority || '').includes('P1')) ? 'p1' : 'p2';
        overlay.dataset.severity = maxSeverity;

        const card = document.createElement('div');
        card.className = 'mia-card';

        card.innerHTML = `
            <div class="mia-header">
                <span class="mia-dot" aria-hidden="true"></span>
                <span class="mia-title">Major Incident Active</span>
                <button class="mia-close" data-action="close" aria-label="Close">✕</button>
            </div>
            <div class="mia-divider"></div>
            <div class="mia-incidents"></div>
        `;

        const list = card.querySelector('.mia-incidents');
        incidents.forEach(incident => {
            const item = document.createElement('div');
            item.className = 'mia-incident-item';

            const link = incident.jiraUrl
                ? `<a class="mia-incident-link" href="${incident.jiraUrl}" target="_blank" rel="noopener">${incident.key}</a>`
                : `<span class="mia-incident-key">${incident.key}</span>`;

            const pc = (incident.priority || '').match(/^(P[1-4])/i)?.[1].toLowerCase() || '';
            const priorityBadge = pc
                ? `<span class="mia-priority mia-priority--${pc}">${incident.priority}</span>`
                : '';

            const affected = incident.affectedServices
                ? `<div class="mia-affected">Affected: ${incident.affectedServices.replace(/\|\|/g, ', ')}</div>`
                : '';

            const roi = incident.roi
                ? `<div class="mia-roi">Region: ${incident.roi}</div>`
                : '';

            item.innerHTML = `
                <div class="mia-incident-row">${link}${priorityBadge}</div>
                <div class="mia-incident-summary">${incident.summary}</div>
                ${affected}${roi}
            `;
            list.appendChild(item);
        });

        overlay.appendChild(card);
        document.body.appendChild(overlay);

        overlay.querySelector('[data-action="close"]').addEventListener('click', () => this.hide());
    }

    hide() {
        const el = document.getElementById('major-incident-overlay');
        if (el) el.remove();
    }
}
