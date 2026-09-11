import * as d3 from 'd3';
import { getQueryParam, setSearchQuery } from '../shared/utils.js';

export class GraphRenderer {
    constructor(app) {
        this.app = app;
        this.simulation = null;
        this.g = null;
        this.zoom = null;
        this.zoomIdentity = null;
        this.svg = null;
        this.linkGraph = null;
        this.nodeGraph = null;
        this.labels = null;
        this.clickedNode = null;
        this.width = 0;
        this.height = 0;
    }

    initDOM() {
        this._ensureJiraBadgeStyles();
        this._ensureJiraCardsPopup();

        const mapEl = document.getElementById('map');
        if (mapEl) {
            this.width = mapEl.clientWidth;
            this.height = mapEl.clientHeight;
        }
        window.addEventListener('resize', () => {
            const el = document.getElementById('map');
            if (!el || !this.svg) return;
            const w = el.clientWidth;
            const h = el.clientHeight;
            this.svg.attr('width', w).attr('height', h);
            this.svg.select('#map-clip rect').attr('width', w).attr('height', h);
        });
    }

    resetVisualization() {
        d3.select('#map').selectAll('*').remove();
        d3.select('#tooltip').style('opacity', 0);
        this.app.legend?.reset();
        d3.select('#serviceDetails').innerHTML = '';
        this.linkGraph = null;
        this.nodeGraph = null;
        this.labels = null;
        this.clickedNode = null;
    }

    createMap() {
        const { store, search, listView, drawer } = this.app;
        const { nodes, links } = store;

        this.zoom = d3.zoom()
            .scaleExtent([0.1, 3])
            .on('zoom', ({ transform }) => { this.g.attr('transform', transform); });

        this.svg = d3.select('#map').append('svg')
            .attr('width', this.width)
            .attr('height', this.height)
            .call(this.zoom);

        const defs = this.svg.append('defs');
        defs.append('clipPath')
            .attr('id', 'map-clip')
            .append('rect')
            .attr('x', 0).attr('y', 0)
            .attr('width', this.width).attr('height', this.height);

        const viewport = this.svg.append('g').attr('clip-path', 'url(#map-clip)');
        this.g = viewport.append('g');

        this.g.append('defs').append('marker')
            .attr('id', 'arrow')
            .attr('viewBox', '0 -5 10 10')
            .attr('refX', 0).attr('refY', 0)
            .attr('markerWidth', 6).attr('markerHeight', 6)
            .attr('orient', 'auto')
            .append('path')
            .attr('d', 'M0,-5L10,0L0,5')
            .attr('fill', this._arrowColor());

        this.simulation = d3.forceSimulation(nodes)
            .force('link', d3.forceLink(links).id(d => d.id).distance(200))
            .force('charge', d3.forceManyBody().strength(-300))
            .force('center', d3.forceCenter(this.width / 2, this.height / 2));

        if (getQueryParam('search')) this.simulation.alphaDecay(0.07);

        this.linkGraph = this.g.append('g')
            .selectAll('line')
            .data(links)
            .enter().append('line')
            .attr('marker-end', 'url(#arrow)');

        const nodeGroups = this.g.append('g')
            .selectAll('g.node-group')
            .data(nodes)
            .enter().append('g')
            .attr('class', 'node-group')
            .call(d3.drag()
                .filter(event => !event.target.closest('.node-icon'))
                .on('start', (event, d) => {
                    if (!event.active) this.simulation.alphaTarget(0.3).restart();
                    d.fx = d.x; d.fy = d.y;
                })
                .on('drag', (event, d) => { d.fx = event.x; d.fy = event.y; })
                .on('end', (event, d) => {
                    if (!event.active) this.simulation.alphaTarget(0);
                    d.fx = d.x; d.fy = d.y;
                }))
            .on('mouseover', (event, d) => {
                const iconEl = event.target.closest('.node-icon');
                let tooltipText;
                if (iconEl?.classList.contains('node-search-icon')) {
                    tooltipText = 'Search for this service';
                } else if (iconEl?.classList.contains('node-expand-icon')) {
                    const active = search.parseActiveKeyValueSearch(search.searchTerm);
                    if (active?.key === 'id') {
                        const inSearch = active.values.some(v => search.normalizeForCompare(v) === search.normalizeForCompare(d.id));
                        tooltipText = inSearch ? 'Remove from search' : 'Add to search';
                    }
                }
                const tooltip = d3.select('#tooltip');
                tooltip.transition().duration(200).style('opacity', .9);
                tooltip.html(tooltipText || d['Description'] || 'No description available')
                    .style('left', (event.pageX + 5) + 'px')
                    .style('top', (event.pageY - 28) + 'px');
            })
            .on('mouseout', () => {
                d3.select('#tooltip').transition().duration(500).style('opacity', 0);
            })
            .on('click', (event, d) => {
                if (event.target.closest('.node-icon')) return;
                this.clickedNode = d;
                drawer.showNodeDetails(d);
            });

        nodeGroups.append('rect')
            .attr('x', -40).attr('y', -20)
            .attr('width', 22).attr('height', 40)
            .attr('fill', 'transparent');

        nodeGroups.append('circle')
            .attr('r', 20)
            .attr('fill', d => d.color);

        this._renderJiraBadges(nodeGroups);

        // Search icon — top-left, always visible on hover
        const searchIcon = nodeGroups.append('g')
            .attr('class', 'node-icon node-search-icon')
            .attr('transform', 'translate(-38, -14)');
        searchIcon.append('circle')
            .attr('cx', 5).attr('cy', 5).attr('r', 4)
            .attr('fill', 'none');
        searchIcon.append('line')
            .attr('x1', 8).attr('y1', 8).attr('x2', 13).attr('y2', 13);
        searchIcon.append('rect')
            .attr('width', 14).attr('height', 14).attr('fill', 'transparent')
            .on('click', (event, d) => {
                event.stopPropagation();
                search.updateSearchAndRefresh(`id:"${d.id}"`, false);
                window.scrollTo({ top: 0, behavior: 'smooth' });
            });

        // Expand/remove icon — bottom-left, visible only when id: search is active
        const expandIcon = nodeGroups.append('g')
            .attr('class', 'node-icon node-expand-icon')
            .attr('transform', 'translate(-38, 4)');
        expandIcon.append('line').attr('class', 'expand-h').attr('x1', 1).attr('y1', 7).attr('x2', 13).attr('y2', 7);
        expandIcon.append('line').attr('class', 'expand-v').attr('x1', 7).attr('y1', 1).attr('x2', 7).attr('y2', 13);
        expandIcon.append('rect')
            .attr('width', 14).attr('height', 14).attr('fill', 'transparent')
            .on('click', (event, d) => {
                event.stopPropagation();
                const active = search.parseActiveKeyValueSearch(search.searchTerm);
                if (!active || active.key !== 'id') return;
                const values = [...active.values];
                const needle = search.normalizeForCompare(d.id);
                const idx = values.findIndex(v => search.normalizeForCompare(v) === needle);
                if (idx === -1) { values.push(d.id); }
                else { values.splice(idx, 1); }
                const newTerm = values.length
                    ? search.buildKeyValueSearch('id', values, active.quoted)
                    : 'id:';
                search.updateSearchAndRefresh(newTerm, false);
            });

        this.nodeGraph = nodeGroups;

        this.labels = this.g.append('g')
            .selectAll('text')
            .data(nodes)
            .enter().append('text')
            .attr('dy', -2)
            .attr('text-anchor', 'middle')
            .text(d => d.id);

        this.simulation.on('tick', () => {
            this.linkGraph
                .attr('x1', d => d.source.x).attr('y1', d => d.source.y)
                .attr('x2', d => {
                    const dx = d.target.x - d.source.x;
                    const dy = d.target.y - d.source.y;
                    const dist = Math.sqrt(dx * dx + dy * dy);
                    return d.target.x - (dx / dist) * 31;
                })
                .attr('y2', d => {
                    const dx = d.target.x - d.source.x;
                    const dy = d.target.y - d.source.y;
                    const dist = Math.sqrt(dx * dx + dy * dy);
                    return d.target.y - (dy / dist) * 31;
                });
            this.nodeGraph.attr('transform', d => `translate(${d.x},${d.y})`);
            this.labels.attr('x', d => d.x).attr('y', d => d.y - 30);
        });

        this.zoomIdentity = d3.zoomIdentity;

        document.addEventListener('click', (e) => {
            const trigger = e.target.closest('.search-trigger');
            const addBtn = e.target.closest('.search-add');
            const remBtn = e.target.closest('.search-remove');

            if (trigger) {
                this.clickedNode = null;
                e.preventDefault();
                const key = decodeURIComponent(trigger.getAttribute('data-key'));
                const isAccurateSearch = key === 'Depends on' || key === 'Used by' || key === 'id';
                const mappedKey = isAccurateSearch ? 'id' : key;
                const value = isAccurateSearch
                    ? `"${decodeURIComponent(trigger.getAttribute('data-value'))}"`
                    : `${decodeURIComponent(trigger.getAttribute('data-value'))}`;
                search.updateSearchAndRefresh(`${mappedKey}:${value}`);
                window.scrollTo({ top: 0, behavior: 'smooth' });
                return;
            }

            if (addBtn || remBtn) {
                e.preventDefault();
                e.stopPropagation();
                const btn = addBtn || remBtn;
                const key = decodeURIComponent(btn.getAttribute('data-key'));
                const value = decodeURIComponent(btn.getAttribute('data-value'));
                const active = search.parseActiveKeyValueSearch(search.searchTerm);
                if (!active || active.key !== key) return;
                const values = [...active.values];
                const needle = search.normalizeForCompare(value);
                const idx = values.findIndex(v => search.normalizeForCompare(v) === needle);
                if (addBtn) { if (idx === -1) values.push(value); }
                else { if (idx !== -1) values.splice(idx, 1); }
                const newTerm = values.length ? search.buildKeyValueSearch(active.key, values, active.quoted) : `${active.key}:`;
                search.updateSearchAndRefresh(newTerm);
            }
        });
    }


    _renderJiraBadges(nodeGroups) {
        const badgeData = (d) => {
            const incidents = d.__jiraCards?.incidents || [];
            const requests = d.__jiraCards?.requests || [];
            const result = [];
            if (incidents.length) {
                result.push({
                    type: 'incident',
                    label: String(incidents.length),
                    title: 'Open incidents',
                    cards: incidents,
                    fill: '#ef4444',
                    text: '#ffffff',
                    x: 15,
                    y: -20,
                });
            }
            if (requests.length) {
                result.push({
                    type: 'request',
                    label: String(requests.length),
                    title: 'Open service requests',
                    cards: requests,
                    fill: '#f59e0b',
                    text: '#111827',
                    x: 15,
                    y: incidents.length ? -2 : -11,
                });
            }
            return result;
        };

        const badges = nodeGroups
            .selectAll('g.node-jira-badge')
            .data(d => badgeData(d).map(b => ({ ...b, node: d })))
            .enter()
            .append('g')
            .attr('class', d => `node-jira-badge node-jira-badge--${d.type}`)
            .attr('transform', d => `translate(${d.x},${d.y})`)
            .style('cursor', 'pointer')
            .on('mouseover', (event, d) => {
                event.stopPropagation();
                const counts = { P1: 0, P2: 0, P3: 0, P4: 0 };
                d.cards.forEach(c => { if (counts[c.priority] !== undefined) counts[c.priority]++; });
                const parts = ['P1', 'P2', 'P3', 'P4'].filter(p => counts[p] > 0).map(p => `${counts[p]} ${p}`);
                const breakdown = parts.length ? parts.join(', ') : String(d.cards.length);
                const tooltip = d3.select('#tooltip');
                tooltip.transition().duration(120).style('opacity', .9);
                tooltip.html(`<strong>${d.title}</strong><br>${breakdown}`)
                    .style('left', (event.pageX + 5) + 'px')
                    .style('top', (event.pageY - 28) + 'px');
            })
            .on('mouseout', () => {
                d3.select('#tooltip').transition().duration(250).style('opacity', 0);
            })
            .on('click', (event, d) => {
                event.stopPropagation();
                this._openJiraCardsPopup(d.node, d.type, d.cards, event);
            });

        badges.append('rect')
            .attr('x', d => -Math.max(9, 5 + d.label.length * 3.2))
            .attr('y', -8)
            .attr('rx', 7)
            .attr('ry', 7)
            .attr('width', d => Math.max(18, 10 + d.label.length * 6.4))
            .attr('height', 16)
            .attr('fill', d => d.fill)
            .attr('stroke', 'rgba(255,255,255,.9)')
            .attr('stroke-width', 1.2);

        badges.append('text')
            .attr('text-anchor', 'middle')
            .attr('dominant-baseline', 'central')
            .attr('fill', d => d.text)
            .attr('font-size', 10)
            .attr('font-weight', 800)
            .text(d => d.label);
    }

    _ensureJiraCardsPopup() {
        if (document.getElementById('domino-jira-cards-popup')) return;

        const popup = document.createElement('div');
        popup.id = 'domino-jira-cards-popup';
        popup.className = 'domino-jira-popup';
        popup.setAttribute('aria-hidden', 'true');
        popup.innerHTML = `
            <div class="domino-jira-popup__header">
                <div>
                    <div class="domino-jira-popup__eyebrow">Jira open items</div>
                    <div class="domino-jira-popup__title">Open items</div>
                </div>
                <button class="domino-jira-popup__close" type="button" aria-label="Close">×</button>
            </div>
            <div class="domino-jira-popup__body"></div>
        `;
        document.body.appendChild(popup);

        popup.querySelector('.domino-jira-popup__close')?.addEventListener('click', () => this._closeJiraCardsPopup());

        const body = popup.querySelector('.domino-jira-popup__body');
        if (body && body.dataset.serviceSearchBound !== '1') {
            body.dataset.serviceSearchBound = '1';
            body.addEventListener('click', (event) => {
                const trigger = event.target.closest('.domino-jira-card__service-link');
                if (!trigger) return;

                event.preventDefault();
                event.stopPropagation();

                const serviceId = trigger.dataset.serviceId || trigger.textContent || '';
                this._searchServiceFromJiraPopup(serviceId);
            });
        }

        this._makePopupDraggable(popup, popup.querySelector('.domino-jira-popup__header'));
    }

    _openJiraCardsPopup(node, type, cards = [], event = null) {
        this._ensureJiraCardsPopup();

        const popup = document.getElementById('domino-jira-cards-popup');
        const title = popup.querySelector('.domino-jira-popup__title');
        const body = popup.querySelector('.domino-jira-popup__body');
        const isIncident = type === 'incident';
        const safeCards = Array.isArray(cards) ? cards : [];

        popup.classList.toggle('domino-jira-popup--incident', isIncident);
        popup.classList.toggle('domino-jira-popup--request', !isIncident);
        title.textContent = `${safeCards.length} ${isIncident ? 'open incident' : 'open service request'}${safeCards.length === 1 ? '' : 's'} · ${node.id}`;

        const PRIORITY_ORDER = { P1: 0, P2: 1, P3: 2, P4: 3 };
        const sortedCards = [...safeCards].sort((a, b) =>
            (PRIORITY_ORDER[a.priority] ?? 99) - (PRIORITY_ORDER[b.priority] ?? 99)
        );
        body.innerHTML = sortedCards.length
            ? sortedCards.map(card => this._jiraCardHtml(card)).join('')
            : '<div class="domino-jira-popup__empty">No open items found for this service.</div>';

        const x = Math.min((event?.pageX ?? 40) + 16, window.scrollX + window.innerWidth - 460);
        const y = Math.min((event?.pageY ?? 80) + 16, window.scrollY + window.innerHeight - 360);
        popup.style.left = `${Math.max(window.scrollX + 16, x)}px`;
        popup.style.top = `${Math.max(window.scrollY + 16, y)}px`;
        popup.classList.add('open');
        popup.setAttribute('aria-hidden', 'false');
    }

    _closeJiraCardsPopup() {
        const popup = document.getElementById('domino-jira-cards-popup');
        popup?.classList.remove('open');
        popup?.setAttribute('aria-hidden', 'true');
    }

    _jiraCardHtml(card) {
        const escape = (value) => String(value ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');

        const key = escape(card.key || 'No key');
        const summary = escape(card.summary || 'No summary');
        const status = escape(card.status || '');
        const roiBadge = card.roi ? `<span class="domino-jira-card__roi">${escape(card.roi)}</span>` : '';
        const priorityBadge = card.priority
            ? `<span class="domino-jira-card__priority domino-jira-card__priority--${card.priority.toLowerCase()}">${escape(card.priority)}</span>`
            : '';
        const internalUrl = String(card.jiraUrl || '').trim();
        const url = internalUrl.replace("browse", "servicedesk/customer/portal/1");
        const keyHtml = url
            ? `<a href="${escape(url)}" target="_blank" rel="noopener noreferrer">${key}</a>`
            : key;

        const impactedServicesHtml = (card.impactedServices || [])
            .map(service => String(service || '').trim())
            .filter(Boolean)
            .map(service => {
                const safe = escape(service);
                return `<button type="button" class="domino-jira-card__service-link" data-service-id="${safe}" title="Search id:&quot;${safe}&quot;">${safe}</button>`;
            })
            .join('');

        return `
            <article class="domino-jira-card">
                <div class="domino-jira-card__topline">
                    <span class="domino-jira-card__key-group">
                        <span class="domino-jira-card__key">${keyHtml}</span>
                        ${priorityBadge}
                    </span>
                    ${status ? `<span class="domino-jira-card__status">${status}</span>` : ''}
                </div>
                <div class="domino-jira-card__summary">${summary}</div>
                ${roiBadge}
                <div class="domino-jira-card__services">
                    <strong>Impacted services:</strong>
                    ${impactedServicesHtml || '<span class="domino-jira-card__no-services">—</span>'}
                </div>
            </article>
        `;
    }

    _searchServiceFromJiraPopup(serviceId) {
        const value = String(serviceId || '').trim();
        if (!value) return;

        this.clickedNode = null;
        this._closeJiraCardsPopup?.();
        this.app?.search?.updateSearchAndRefresh?.(`id:"${value}"`, false);
        window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    _makePopupDraggable(popup, handle) {
        if (!popup || !handle || handle.dataset.dragBound === '1') return;
        handle.dataset.dragBound = '1';

        let startX = 0;
        let startY = 0;
        let originX = 0;
        let originY = 0;
        let dragging = false;

        const onMove = (event) => {
            if (!dragging) return;
            const dx = event.clientX - startX;
            const dy = event.clientY - startY;
            popup.style.left = `${originX + dx}px`;
            popup.style.top = `${originY + dy}px`;
        };

        const onUp = () => {
            dragging = false;
            document.removeEventListener('mousemove', onMove);
            document.removeEventListener('mouseup', onUp);
            document.body.classList.remove('domino-jira-popup-dragging');
        };

        handle.addEventListener('mousedown', (event) => {
            if (event.target.closest('button, a')) return;
            dragging = true;
            const rect = popup.getBoundingClientRect();
            startX = event.clientX;
            startY = event.clientY;
            originX = rect.left + window.scrollX;
            originY = rect.top + window.scrollY;
            document.body.classList.add('domino-jira-popup-dragging');
            document.addEventListener('mousemove', onMove);
            document.addEventListener('mouseup', onUp);
            event.preventDefault();
        });
    }

    _ensureJiraBadgeStyles() {
        if (document.getElementById('domino-jira-badge-styles')) return;
        const style = document.createElement('style');
        style.id = 'domino-jira-badge-styles';
        style.textContent = `
            .node-jira-badge,
            .node-group .node-jira-badge,
            .node-group:hover .node-jira-badge {
                opacity: 1 !important;
                visibility: visible !important;
                pointer-events: auto !important;
                display: inline !important;
            }
            .node-jira-badge text { pointer-events: none; font-family: inherit; font-weight: 800; }
            .node-jira-badge rect { pointer-events: auto; cursor: pointer; filter: drop-shadow(0 1px 2px rgba(0,0,0,.35)); }
            .domino-jira-popup {
                position: absolute;
                z-index: 1200;
                display: none;
                width: min(440px, calc(100vw - 32px));
                max-height: min(520px, calc(100vh - 48px));
                color: #f5f5f7;
                background: rgba(24,24,28,.96);
                border: 1px solid rgba(255,255,255,.14);
                border-radius: 14px;
                box-shadow: 0 22px 70px rgba(0,0,0,.42);
                backdrop-filter: blur(14px);
                overflow: hidden;
            }
            .domino-jira-popup.open { display: flex; flex-direction: column; }
            .domino-jira-popup__header {
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 16px;
                padding: 12px 14px;
                border-bottom: 1px solid rgba(255,255,255,.10);
                cursor: move;
                user-select: none;
            }
            .domino-jira-popup__eyebrow {
                margin-bottom: 3px;
                color: rgba(255,255,255,.48);
                font-size: 10px;
                font-weight: 800;
                letter-spacing: .14em;
                text-transform: uppercase;
            }
            .domino-jira-popup__title {
                color: #fff;
                font-size: 13px;
                font-weight: 800;
                line-height: 1.25;
            }
            .domino-jira-popup__close {
                width: 28px;
                height: 28px;
                display: inline-flex;
                align-items: center;
                justify-content: center;
                border: 1px solid rgba(255,255,255,.14);
                border-radius: 8px;
                color: #fff;
                background: rgba(255,255,255,.07);
                cursor: pointer;
                font-size: 18px;
                line-height: 1;
            }
            .domino-jira-popup__close:hover { background: rgba(255,255,255,.14); }
            .domino-jira-popup__body {
                padding: 12px;
                overflow: auto;
            }
            .domino-jira-card {
                padding: 10px 10px 11px;
                border: 1px solid rgba(255,255,255,.10);
                border-radius: 10px;
                background: rgba(255,255,255,.055);
            }
            .domino-jira-card + .domino-jira-card { margin-top: 8px; }
            .domino-jira-card__topline {
                display: flex;
                align-items: center;
                justify-content: space-between;
                gap: 10px;
                margin-bottom: 6px;
            }
            .domino-jira-card__key-group {
                display: flex;
                align-items: center;
                gap: 6px;
                min-width: 0;
            }
            .domino-jira-card__key,
            .domino-jira-card__key a {
                color: #93c5fd;
                font-size: 12px;
                font-weight: 850;
                text-decoration: none;
            }
            .domino-jira-card__key a:hover { text-decoration: underline; }
            .domino-jira-card__status {
                max-width: 160px;
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
                color: rgba(255,255,255,.64);
                font-size: 11px;
            }
            .domino-jira-card__summary {
                color: #fff;
                font-size: 12px;
                line-height: 1.35;
            }
            .domino-jira-card__roi {
                display: inline-block;
                font-size: 10px;
                font-weight: 600;
                padding: 1px 5px;
                border-radius: 3px;
                background: rgba(37,99,235,0.25);
                color: #93c5fd;
                margin-top: 4px;
            }
            html:not([data-theme="dark"]) .domino-jira-card__roi {
                background: #eff6ff;
                color: #1d4ed8;
            }
            .domino-jira-card__services {
                margin-top: 7px;
                color: rgba(255,255,255,.62);
                font-size: 11px;
                line-height: 1.35;
            }
            .domino-jira-popup__empty {
                padding: 22px;
                color: rgba(255,255,255,.62);
                text-align: center;
            }
            .domino-jira-popup--incident .domino-jira-popup__header { border-top: 3px solid #ef4444; }
            .domino-jira-popup--request .domino-jira-popup__header { border-top: 3px solid #f59e0b; }
            html:not([data-theme="dark"]) .domino-jira-popup {
                color: #111827;
                background: rgba(255,255,255,.98);
                border-color: rgba(0,0,0,.12);
                box-shadow: 0 22px 70px rgba(0,0,0,.18);
            }
            html:not([data-theme="dark"]) .domino-jira-popup__header { border-bottom-color: rgba(0,0,0,.10); }
            html:not([data-theme="dark"]) .domino-jira-popup__eyebrow { color: rgba(0,0,0,.45); }
            html:not([data-theme="dark"]) .domino-jira-popup__title,
            html:not([data-theme="dark"]) .domino-jira-card__summary { color: #111827; }
            html:not([data-theme="dark"]) .domino-jira-popup__close {
                color: #111827;
                border-color: rgba(0,0,0,.10);
                background: rgba(0,0,0,.04);
            }
            html:not([data-theme="dark"]) .domino-jira-card {
                border-color: rgba(0,0,0,.10);
                background: rgba(0,0,0,.035);
            }
            html:not([data-theme="dark"]) .domino-jira-card__services,
            html:not([data-theme="dark"]) .domino-jira-card__status { color: rgba(0,0,0,.58); }
            body.domino-jira-popup-dragging { user-select: none; }
        `;
        document.head.appendChild(style);
    }

    centerAndZoomOnNode(node) {
        const scale = 1;
        const x = -node.x * scale + this.width / 2;
        const y = -node.y * scale + this.height / 2;
        const transform = this.zoomIdentity.translate(x, y).scale(scale).translate(-0, -0);
        this.svg.transition().duration(750).call(this.zoom.transform, transform);
    }

    fitGraphToViewport(paddingRatio = 0.90) {
        if (!this.svg || !this.g) return;
        const bbox = this.g.node()?.getBBox();
        if (!bbox || !isFinite(bbox.width) || !isFinite(bbox.height) || bbox.width === 0 || bbox.height === 0) {
            this.svg.call(this.zoom.transform, d3.zoomIdentity);
            return;
        }
        const scale = Math.min(this.width / bbox.width, this.height / bbox.height) * paddingRatio;
        const tx = this.width / 2 - (bbox.x + bbox.width / 2) * scale;
        const ty = this.height / 2 - (bbox.y + bbox.height / 2) * scale;
        this.svg.transition().duration(400).call(this.zoom.transform, d3.zoomIdentity.translate(tx, ty).scale(scale));
    }

    updateVisualization(showDrawer = true) {
        const { store, search, listView, drawer } = this.app;
        search.prepareSearchTerm();

        const { nodes, links, activeServiceNodeIds } = store;
        const filteredLinks = links.filter(l =>
            activeServiceNodeIds.has(l.source.id) && activeServiceNodeIds.has(l.target.id)
        );

        const relatedNodes = new Set();
        const searchedNodes = new Set();

        const relatedLinks = links.filter(l => {
            const isLinkStatusOk = !search.hideStoppedServices || filteredLinks.includes(l);
            let isSearchedLink = search.isEffectivelyEmpty();
            if (search.isSearchResultValueOnly(l.source) || search.isSearchResultWithKeyValue(l.source)) {
                isSearchedLink = true;
                searchedNodes.add(l.source.id);
            }
            if (search.isSearchResultValueOnly(l.target) || search.isSearchResultWithKeyValue(l.target)) {
                isSearchedLink = true;
                searchedNodes.add(l.target.id);
            }
            if (isLinkStatusOk && isSearchedLink) {
                relatedNodes.add(l.source.id);
                relatedNodes.add(l.target.id);
                return true;
            }
            return false;
        });

        let nodeToZoom;
        const relaxed = document.getElementById('relaxed-search');
        this.nodeGraph.each(d => {
            const byKey = search.isSearchResultWithKeyValue(d);
            const byValue = relaxed?.checked && search.isSearchResultValueOnly(d);
            if (byKey || byValue) {
                nodeToZoom = nodeToZoom || d;
                relatedNodes.add(d.id);
                searchedNodes.add(d.id);
            }
        });

        const visibleNodes = search.showConnections ? relatedNodes : searchedNodes;
        const visible = (d) =>
            (search.isEffectivelyEmpty() && !search.hideStoppedServices) ||
            (search.isEffectivelyEmpty() && search.hideStoppedServices && activeServiceNodeIds.has(d.id)) ||
            (visibleNodes.has(d.id) && (!search.hideStoppedServices || activeServiceNodeIds.has(d.id)));

        this.nodeGraph.style('display', d => visible(d) ? 'block' : 'none');
        this.linkGraph.style('display', d =>
            (search.isEffectivelyEmpty() && !search.hideStoppedServices) ||
            (search.isEffectivelyEmpty() && search.hideStoppedServices && activeServiceNodeIds.has(d.source.id) && activeServiceNodeIds.has(d.target.id)) ||
            relatedLinks.includes(d) && (search.showConnections || (searchedNodes.has(d.source.id) && searchedNodes.has(d.target.id)))
                ? 'block' : 'none'
        );
        this.labels.style('display', d => visible(d) ? 'block' : 'none');
        this.labels.style('text-decoration', d => searchedNodes.has(d.id) ? 'underline' : 'none');

        const activeIdSearch = search.parseActiveKeyValueSearch(search.searchTerm);
        const hasIdSearch = activeIdSearch?.key === 'id';
        this.nodeGraph?.selectAll('.node-expand-icon').style('display', hasIdSearch ? null : 'none');
        if (hasIdSearch) {
            this.nodeGraph?.selectAll('.node-expand-icon').each(function(d) {
                const inSearch = activeIdSearch.values.some(v => search.normalizeForCompare(v) === search.normalizeForCompare(d.id));
                d3.select(this).select('.expand-v').style('display', inSearch ? 'none' : null);
            });
        }

        search.currentNodes = nodes;
        search.currentSearchedNodes = searchedNodes;

        if (document.getElementById('list-view')?.style.display === 'block') {
            listView.renderListFromSearch();
        } else if (!this.clickedNode && nodeToZoom && (!search.hideStoppedServices || activeServiceNodeIds.has(nodeToZoom.id))) {
            this.centerAndZoomOnNode(nodeToZoom);
            drawer.showNodeDetails(nodeToZoom, showDrawer);
        }
    }

    _arrowColor() {
        return document.documentElement.getAttribute('data-theme') === 'dark' ? '#c8d0e8' : '#999';
    }

    updateGraphTheme() {
        const path = document.querySelector('#arrow path');
        if (path) path.setAttribute('fill', this._arrowColor());
    }
}
