import * as d3 from 'd3';
import {
    enableGlobalFindShortcut,
    closeSideDrawer,
    createModal,
    getQueryParam,
    initCommonActions,
    setSearchQuery,
    splitValues,
    applyTheme,
    loadSavedTheme,
} from '../shared/utils.js';
import { BRAND, renderBrandLogo } from '../../brand-specific/brand.js';
import { ROLE_FIELD_WITH_MAPPING, BUSINESS_FUNCTION_FIELD } from './constants.js';
import { AutocompleteEngine } from '../shared/AutocompleteEngine.js';
import { PeopleDatabase } from './PeopleDatabase.js';
import { OrgChartRenderer } from './OrgChartRenderer.js';
import { InteractionController } from './InteractionController.js';
import { ScenarioManager } from './ScenarioManager.js';
import { SolitaireSearch } from './SolitaireSearch.js';
import { ColorLegend } from './ColorLegend.js';
import { TeamDetailDrawer } from './TeamDetailDrawer.js';
import { PersonDetailDrawer } from './PersonDetailDrawer.js';
import { ContextMenu } from './ContextMenu.js';
import { QuickFilters } from './QuickFilters.js';
import { PostItNote } from '../shared/PostItNote.js';
import { AnnouncementBar } from '../shared/AnnouncementBar.js';

export class SolitaireApp {
    constructor() {
        this.db = new PeopleDatabase(this);
        this.renderer = new OrgChartRenderer(this);
        this.interaction = new InteractionController(this);
        this.scenario = new ScenarioManager(this);
        this.search = new SolitaireSearch(this);
        this.legend = new ColorLegend(this);
        this.drawer = new TeamDetailDrawer(this);
        this.personDrawer = new PersonDetailDrawer(this);
        this.contextMenu = new ContextMenu(this);
        this.postIt = new PostItNote('solitaire');
        this.announcementBar = new AnnouncementBar();

        this.quickFilters = new QuickFilters(this);

        this.autocomplete = null;
        this.visibleOrg = null;
        this.searchParam = null;
        this.jengaServicesThisMonth = new Set();
        this.isAdvanced = (() => {
            const p = getQueryParam('advanced');
            return p ? p === 'true' : false;
        })();
    }

    init() {
        this.announcementBar.init();
        renderBrandLogo();
        const theme = loadSavedTheme();
        const dmToggle = document.getElementById('toggle-dark-mode');
        if (dmToggle) dmToggle.checked = theme === 'dark';

        this._initSideDrawerEvents();
        this.search.initChipBar();
        this.drawer.initEvents();
        this.personDrawer.initEvents();
        this.interaction.setupLongPress();
        this._handleAdvancedMode();
        const buildInfoEl = document.getElementById('build-info');
        if (buildInfoEl) {
            buildInfoEl.textContent = `Build ${__APP_BUILD__} · ${__BUILD_DATE__}`;
            buildInfoEl.style.cursor = 'pointer';
            buildInfoEl.title = 'Click to show changelog';
            buildInfoEl.addEventListener('click', () => this.announcementBar.show());
        }
        const lastUpdateEl = document.getElementById('side-last-update');
        if (lastUpdateEl) {
            lastUpdateEl.style.cursor = 'pointer';
            lastUpdateEl.title = 'Click to show changelog';
            lastUpdateEl.addEventListener('click', () => this.announcementBar.show());
        }
        this._enableAppPinchZoomOnly();
        this._setupGlobalTooltip();
        this._initSearchInput();
        this._initImportScenario();
        this._initFileInput();
        this._initToggleDraggable();
        enableGlobalFindShortcut({ inputSelector: '#drawer-search-input' });
        this.postIt.init();

        window.addEventListener('keydown', (e) => {
            if (e.key !== 'Escape') return;
            // If the autocomplete dropdown is open, let AutocompleteEngine consume Escape first.
            if (document.getElementById('ac-dropdown')?.classList.contains('ac-open')) return;
            const personDrawerOpen = document.body.classList.contains('person-drawer-open');
            const drawerOpen = document.body.classList.contains('drawer-open');
            if (personDrawerOpen) {
                // First Escape closes L2 only; second Escape closes the full drawer
                const panels = document.querySelector('#person-drawer .person-drawer-panels');
                if (panels?.classList.contains('panels-level2')) {
                    this.personDrawer.closeL2();
                } else {
                    this.personDrawer.close();
                }
            } else if (drawerOpen) {
                this.drawer.close();
            } else {
                this.handleClearAction('Escape').then(() => {});
            }
        });

        window.addEventListener('load', () => {
            const peopleFetch = fetch(BRAND.csv.solitaire).then(r => r.text());
            const filtersFetch = BRAND.csv.customFilters
                ? fetch(BRAND.csv.customFilters).then(r => r.text()).catch(() => '')
                : Promise.resolve('');
            const jengaFetch = BRAND.csv.jenga
                ? fetch(BRAND.csv.jenga).then(r => r.text()).catch(() => '')
                : Promise.resolve('');

            Promise.all([peopleFetch, filtersFetch, jengaFetch])
                .then(([csvData, filtersCsv, jengaCsv]) => {
                    if (filtersCsv) {
                        this.quickFilters.load(filtersCsv);
                        this.quickFilters.render();
                        this.quickFilters.initEvents();
                    }
                    if (jengaCsv) this._buildJengaServicesThisMonth(jengaCsv);
                    this.loadAndRender(csvData);
                    requestAnimationFrame(() => this._hideSpinner());
                    this.searchParam = getQueryParam('search');
                    if (this.searchParam) {
                        const inp = document.getElementById('drawer-search-input');
                        if (inp) inp.value = this.searchParam;
                        this.search._refreshChips(this.searchParam);
                        const openDetail = getQueryParam('showDetails') === 'true';
                        requestAnimationFrame(() => {
                            requestAnimationFrame(() => {
                                this.search.search(this.searchParam);
                                if (openDetail) this._openDetailFromSearch(this.searchParam);
                            });
                        });
                    }
                    const personParam = getQueryParam('person');
                    if (personParam) {
                        requestAnimationFrame(() => {
                            requestAnimationFrame(() => {
                                this._openPersonFromParam(personParam);
                            });
                        });
                    }
                })
                .catch(err => console.error('Error loading CSV files:', err));
        });
    }

    _hideSpinner() {
        const el = document.getElementById('app-spinner');
        if (!el) return;
        el.classList.add('is-hidden');
        el.addEventListener('transitionend', () => el.remove(), { once: true });
    }

    loadAndRender(csvText) {
        const data = this.db.load(csvText);
        if (!data) return;
        this.renderer.reset();
        this.renderer.render(data);
        this.interaction.applyDraggableToggleState();
        this._buildAutocompleteIndex();
    }

    _buildJengaServicesThisMonth(csvText) {
        const now = new Date();
        const y = now.getFullYear();
        const m = now.getMonth();
        try {
            const rows = d3.csvParse(csvText);
            for (const row of rows) {
                const svc = (row.Service || '').trim();
                if (!svc) continue;
                const dateStr = row.DueDate || row.StartDate;
                if (!dateStr) continue;
                const d = new Date(dateStr);
                if (d.getFullYear() === y && d.getMonth() === m) {
                    this.jengaServicesThisMonth.add(svc.toLowerCase());
                }
            }
        } catch {}
    }

    _buildAutocompleteIndex() {
        if (!this.db.people || !this.db.people.length) return;

        const keys = ['name', 'role', 'company', 'location', 'function', 'room', 'service'];
        const m = new Map();
        const sortFn = (a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' });

        // Read directly from rendered card DOM — the only source guaranteed to
        // match what is actually visible after all filters are applied.
        const cards = Array.from(document.querySelectorAll('g[data-key^="card::"]'))
            .filter(g => g.style.display !== 'none');

        const fromAttr = (attr) => {
            const set = new Set();
            cards.forEach(g => {
                const raw = g.getAttribute(attr) || '';
                splitValues(raw).map(s => s.trim()).filter(Boolean).forEach(v => set.add(v));
            });
            return [...set].sort(sortFn);
        };

        // Name comes from the .profile-name text content inside each card
        const nameSet = new Set();
        cards.forEach(g => {
            const el = g.querySelector('.profile-name');
            const v = (el?.textContent || '').trim();
            if (v) nameSet.add(v);
        });
        m.set('name', [...nameSet].sort(sortFn));

        m.set('role',     fromAttr('data-role'));
        m.set('company',  fromAttr('data-company'));
        m.set('location', fromAttr('data-location'));
        m.set('function', fromAttr('data-function'));
        m.set('room',     fromAttr('data-room'));

        // Services from visible team-title elements
        const serviceSet = new Set();
        document.querySelectorAll('[data-services]').forEach(el => {
            (el.getAttribute('data-services') || '').split(',').forEach(s => {
                if (s.trim()) serviceSet.add(s.trim());
            });
        });
        m.set('service', [...serviceSet].sort(sortFn));

        // Collect stream, theme, team names from the visible org structure
        const org = this.visibleOrg || {};
        const streamSet = new Set();
        const themeSet  = new Set();
        const teamSet   = new Set();
        for (const [stream, themes] of Object.entries(org)) {
            if (stream) streamSet.add(stream);
            for (const [theme, teams] of Object.entries(themes || {})) {
                if (theme) themeSet.add(theme);
                for (const team of Object.keys(teams || {})) {
                    if (team) teamSet.add(team);
                }
            }
        }
        m.set('stream', [...streamSet].sort(sortFn));
        m.set('theme',  [...themeSet].sort(sortFn));
        m.set('team',   [...teamSet].sort(sortFn));

        this.autocomplete = new AutocompleteEngine([...keys, 'stream', 'theme', 'team'], m, { allowMultiValue: false });
        this.autocomplete.init();
    }

    setStreamFilter(streamKeys) {
        const params = new URLSearchParams(window.location.search);
        if (!streamKeys || streamKeys.size === 0) {
            params.delete('stream');
        } else {
            params.set('stream', [...streamKeys].join(','));
        }
        const newUrl = window.location.pathname + (params.toString() ? '?' + params.toString() : '');
        window.history.replaceState({}, '', newUrl);
        this.renderer.reset();
        this.loadAndRender(this.db.cachedCsvText);
    }

    wireFabsInteractions(cardSel) {
        const SHOW_DELAY = 50;
        const HIDE_DELAY = 120;
        let showTimer = null;
        let hideTimer = null;

        const isTouchEnv = () => ('ontouchstart' in window) || navigator.maxTouchPoints > 0;

        const show = () => {
            clearTimeout(hideTimer);
            showTimer = setTimeout(() => cardSel.classed('card--fabs-visible', true), SHOW_DELAY);
        };
        const hide = () => {
            clearTimeout(showTimer);
            hideTimer = setTimeout(() => cardSel.classed('card--fabs-visible', false), HIDE_DELAY);
        };

        cardSel
            .on('pointerenter.fabs', () => { if (!isTouchEnv()) show(); })
            .on('pointerleave.fabs', () => { if (!isTouchEnv()) hide(); });

        cardSel.on('click.fabs', (event) => {
            if (!isTouchEnv()) return;
            event.stopPropagation();
            const vis = cardSel.classed('card--fabs-visible');
            d3.selectAll('g[data-key^="card::"]').classed('card--fabs-visible', false);
            cardSel.classed('card--fabs-visible', !vis);
        });

        cardSel.selectAll('.contact-fabs, .contact-fabs-svg, .contact-fab')
            .on('pointerenter.fabs', (e) => e.stopPropagation())
            .on('pointerleave.fabs', (e) => e.stopPropagation())
            .on('pointerdown.fabs', (e) => e.stopPropagation())
            .on('touchstart.fabs', (e) => e.stopPropagation());

        if (!window.__fabsOutsideHandlerAttached) {
            window.__fabsOutsideHandlerAttached = true;
            document.addEventListener('pointerdown', (e) => {
                const svgEl = document.getElementById('canvas');
                if (!svgEl) return;
                if (!svgEl.contains(e.target)) {
                    d3.selectAll('g[data-key^="card::"]').classed('card--fabs-visible', false);
                }
            }, { passive: true });
        }
    }

    showToast(message, duration = 3000) {
        const positionContainer = (container) => {
            const isOpen = document.getElementById('drawer')?.classList.contains('open')
                || document.getElementById('person-drawer')?.classList.contains('open');
            if (isOpen) {
                container.style.top = 'unset';
                container.style.bottom = '20px';
                container.style.right = '20px';
            } else {
                container.style.bottom = 'unset';
                container.style.top = '70px';
                container.style.right = '20px';
            }
            container.style.zIndex = '10001';
            container.style.position = 'fixed';
            container.style.display = 'flex';
            container.style.flexDirection = 'column';
            container.style.gap = '10px';
        };

        let container = document.querySelector('.toast-container');
        if (!container) {
            container = document.createElement('div');
            container.className = 'toast-container';
            document.body.appendChild(container);
        }

        positionContainer(container);

        const toast = document.createElement('div');
        toast.className = 'toast';
        toast.textContent = message;
        container.appendChild(toast);

        setTimeout(() => toast.classList.add('show'), 10);
        requestAnimationFrame(() => positionContainer(container));
        setTimeout(() => positionContainer(container), 180);

        if (!window.__toastDrawerObserverAttached) {
            const mo = new MutationObserver(() => positionContainer(container));
            ['drawer', 'person-drawer'].forEach(id => {
                const el = document.getElementById(id);
                if (el) mo.observe(el, { attributes: true, attributeFilter: ['class'] });
            });
            window.__toastDrawerObserverAttached = true;
        }

        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 300);
        }, duration);
    }

    async handleClearAction(source = '') {
        const { hasStream, hasOtherValues } = this._getUrlParamsSnapshot();
        const searchInput = document.getElementById('drawer-search-input');
        const hasActiveSearch = !!(searchInput && searchInput.value && searchInput.value.trim() !== '');

        if (hasOtherValues || hasActiveSearch) {
            this._stripUrlParamsExceptStream();
            this.search.clear();
            return;
        }

        if (hasStream) {
            const result = await createModal({
                title: 'Remove the Stream filter?',
                html: 'This will display all streams from the Visual People DB. You can restore a filtered view at any time using eye (hide and show) actions on the top-right corner of the streams lane.',
                buttons: [
                    { id: 'keep', label: 'Keep' },
                    { id: 'remove', label: 'Remove', primary: true }
                ]
            });
            if (result === 'remove') {
                this.setStreamFilter(null);
                this.search.clear();
                this.showToast('Stream filter removed');
            }
            return;
        }

        this.search.clear();
    }

    _getUrlParamsSnapshot() {
        const params = new URLSearchParams(window.location.search);
        const hasStream = params.has('stream') && String(params.get('stream') ?? '').trim() !== '';
        const IGNORE_KEYS = new Set(['advanced', 'mode', 'view', 'source', 'utm_source', 'utm_medium', 'utm_campaign']);
        const otherKeysWithValue = [];
        for (const [key, value] of params.entries()) {
            if (key === 'stream') continue;
            if (IGNORE_KEYS.has(key)) continue;
            if (String(value ?? '').trim() !== '') otherKeysWithValue.push(key);
        }
        return { hasStream, otherKeysWithValue, hasOtherValues: otherKeysWithValue.length > 0 };
    }

    _stripUrlParamsExceptStream() {
        const params = new URLSearchParams(window.location.search);
        Array.from(params.keys()).forEach(k => { if (k !== 'stream') params.delete(k); });
        const newUrl = window.location.pathname + (params.toString() ? '?' + params.toString() : '');
        window.history.replaceState({}, '', newUrl);
    }

    _handleAdvancedMode() {
        const show = (elId, visible) => {
            const el = document.getElementById(elId);
            if (!el) return;
            el.style.display = visible ? '' : 'none';
        };
        show('act-upload', this.isAdvanced);
        show('label-file', this.isAdvanced);
        show('toggle-draggable', this.isAdvanced);
    }

    _enableAppPinchZoomOnly() {
        const svgEl = document.getElementById('canvas');
        if (!svgEl) return;
        svgEl.addEventListener('touchmove', (e) => {
            if (e.touches.length > 1) e.preventDefault();
        }, { passive: false });
    }

    _setupGlobalTooltip() {
        let tipEl = null;
        let showTimer = null;
        let hideTimer = null;
        let currentAnchor = null;

        const SHOW_DELAY = 90;
        const HIDE_DELAY = 140;
        const isMouseLike = window.matchMedia('(hover: hover) and (pointer: fine)').matches;

        const ensureTip = () => {
            if (!tipEl) {
                tipEl = document.createElement('div');
                tipEl.className = 'app-tooltip';
                document.body.appendChild(tipEl);
            }
            tipEl.style.zIndex = String(2147483647);
            return tipEl;
        };

        const isVisible = () => !!(tipEl && tipEl.classList.contains('show'));

        const positionTip = (anchor, placement = 'right') => {
            const el = ensureTip();
            const rect = anchor.getBoundingClientRect();
            const vw = window.innerWidth;
            const vh = window.innerHeight;
            let x = rect.right + 8;
            let y = rect.top + rect.height / 2;
            el.style.transform = 'translate(0, -50%)';
            if (placement === 'top') {
                x = rect.left + rect.width / 2; y = rect.top - 8;
                el.style.transform = 'translate(-50%, -100%)';
            } else if (placement === 'bottom') {
                x = rect.left + rect.width / 2; y = rect.bottom + 8;
                el.style.transform = 'translate(-50%, 0)';
            } else if (placement === 'left') {
                x = rect.left - 8; y = rect.top + rect.height / 2;
                el.style.transform = 'translate(-100%, -50%)';
            }
            // Clamp so the tooltip never bleeds outside the viewport.
            // getBoundingClientRect on a fixed element needs the tip to be shown first,
            // so we use max-width (220px) as a conservative estimate for clamping.
            const TIP_W = 220;
            const TIP_H = 80; // generous estimate for multi-line
            const MARGIN = 8;
            if (placement === 'bottom' || placement === 'top') {
                // centre-aligned: clamp so tip stays inside left/right edges
                const minX = TIP_W / 2 + MARGIN;
                const maxX = vw - TIP_W / 2 - MARGIN;
                x = Math.max(minX, Math.min(maxX, x));
            }
            if (placement === 'bottom' && y + TIP_H > vh - MARGIN) {
                // flip to top if not enough space below
                y = rect.top - 8;
                el.style.transform = 'translate(-50%, -100%)';
            }
            el.style.left = `${Math.round(x)}px`;
            el.style.top = `${Math.round(y)}px`;
        };

        const showTip = (text, anchor, placement = 'right') => {
            const el = ensureTip();
            el.textContent = text || '';
            el.classList.add('show');
            positionTip(anchor, placement);
        };

        const hideTipNow = () => { if (tipEl) tipEl.classList.remove('show'); };

        const getFabAnchor = (target) => {
            const a = target?.closest?.('[data-tooltip], .contact-fab') || null;
            // Drawer header buttons use CSS ::after tooltips — skip the JS system for them
            if (a && a.closest('#drawer header')) return null;
            return a;
        };

        if (isMouseLike) {
            document.addEventListener('mouseover', (e) => {
                const a = getFabAnchor(e.target);
                if (!a) return;
                const text = a.getAttribute('data-tooltip') || a.getAttribute('aria-label') || '';
                if (!text) return;
                const placement = a.getAttribute('data-tooltip-placement') || 'right';
                clearTimeout(hideTimer);
                hideTimer = null;
                if (isVisible() && currentAnchor !== a) {
                    currentAnchor = a;
                    showTip(text, a, placement);
                    return;
                }
                currentAnchor = a;
                clearTimeout(showTimer);
                showTimer = setTimeout(() => showTip(text, a, placement), SHOW_DELAY);
            }, true);

            document.addEventListener('mouseout', (e) => {
                const a = getFabAnchor(e.target);
                if (!a) return;
                clearTimeout(showTimer);
                clearTimeout(hideTimer);
                hideTimer = setTimeout(() => { hideTipNow(); currentAnchor = null; }, HIDE_DELAY);
            }, true);

            window.addEventListener('scroll', () => { if (isVisible()) hideTipNow(); }, { passive: true });
            window.addEventListener('resize', () => { if (isVisible()) hideTipNow(); });
            this.renderer.onZoom = () => { if (isVisible()) hideTipNow(); };
        } else {
            document.addEventListener('pointerdown', hideTipNow, { passive: true });
        }
    }

    _initSearchInput() {
        document.getElementById('drawer-search-input')?.addEventListener('keydown', (e) => {
            if (e.key !== 'Enter') return;
            if (this.autocomplete?.hasPendingSelection()) return;
            const chipBar = this.search._chipBar;
            const composed = chipBar?.getRevealModeQuery();
            if (composed) {
                chipBar.clearRevealMode();
                this.search.search(composed);
            } else {
                const query = e.target.value.trim();
                if (query) {
                    this.search.search(query);
                } else {
                    this.search.clear();
                }
            }
            e.preventDefault();
        });

        document.getElementById('drawer-search-input')?.addEventListener('ac-confirm', (e) => {
            this.autocomplete?.hideDropdown();
            const chipBar = this.search._chipBar;
            const composed = chipBar?.getRevealModeQuery();
            if (composed) {
                chipBar.clearRevealMode();
                this.search.search(composed);
            } else {
                const query = e.target.value ? e.target.value.trim() : '';
                if (query) this.search.search(query);
            }
        });
    }

    _initImportScenario() {
        document.getElementById('act-import-scenario')?.addEventListener('click', async () => {
            try {
                await this.scenario.handleAction('import');
            } catch (e) {
                console.warn('Import scenario failed:', e);
                this.showToast('Import failed: invalid clipboard scenario', 5000);
            }
        });
    }

    _initFileInput() {
        document.getElementById('fileInput')?.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (evt) => {
                this.renderer.reset();
                this.loadAndRender(evt.target.result);
            };
            reader.readAsText(file, 'UTF-8');
        });
    }


    _initToggleDraggable() {
        document.getElementById('toggle-draggable')?.addEventListener('change', (e) => {
            if (!this.isAdvanced) {
                e.target.checked = false;
                this.interaction.isDraggable = false;
                this.interaction.clearSelection();
                return;
            }
            this.interaction.isDraggable = e.target.checked;
            if (!this.interaction.isDraggable) {
                this.interaction.clearSelection();
            }
            this.interaction.applyDraggableToggleState();
        });
    }

    _initSideDrawerEvents() {
        initCommonActions();
        document.getElementById('act-about')?.addEventListener('click', () => {
            closeSideDrawer();
            this.drawer.open({
                name: 'About Solitaire ♤',
                description:
                    `Org charts highlight hierarchy—but not how teams actually work. Much of the real collaboration that drives ${BRAND.name}'s operations happens across functions, services, and roles, yet remains invisible. This reinforces silos and hides the complexity of our shared work.\n` +
                    '\n' +
                    '<b><i>Our Vision</b></i>\n' +
                    'By visualizing how teams operate—the people, services, and responsibilities behind daily activities—we strengthen a culture that is collaborative, transparent, and service‑oriented. Visibility turns shared accountability into a tangible part of our operating model.\n' +
                    '\n' +
                    '<b><i>What we are building</b></i>\n' +
                    'A custom Visual People Database that brings together data from Confluence and Jira into a single, interactive view.\n' +
                    `More info on <a href='${BRAND.urls.servicePortal}' target='_blank'>the ${BRAND.name} Service Portal.</a>\n` +
                    '\n' +
                    '<b><i>It provides:</b></i>\n' +
                    '<ul>' +
                    '<li>A clear map of team members (internal staff and suppliers)</li>' +
                    '<li>The services each team manages</li>' +
                    '<li>Roles and responsibilities across the organization</li>' +
                    `<li>Quick access to ${BRAND.name}'s Service Catalog</li>` +
                    '<li>A built‑in "Report a change" feature to keep information fresh and accurate</li></ul>' +
                    '\n' +
                    '<b><i>The Benefits</b></i>\n' +
                    '<ul><li>Understand who works on what across projects and services</li>' +
                    '<li>Make hidden operational networks visible</li>' +
                    '<li>Consolidate data not available in systems like Workday</li>' +
                    '<li>Strengthen transparency, alignment, and cross‑team collaboration</li>' +
                    '<li>Provide a single source of truth for service ownership and responsibilities</li></ul>'
            });
        });

        document.getElementById('act-clear')?.addEventListener('click', () => {
            this.handleClearAction('act-clear');
        });

        document.getElementById('act-fit')?.addEventListener('click', () => {
            this.renderer.fitToContent(0.9);
        });

        document.getElementById('act-report')?.addEventListener('click', () => {
            window.open(BRAND.urls.reportChange, '_blank', 'noopener');
        });

        document.getElementById('drawer-search-go')?.addEventListener('click', () => {
            this.autocomplete?.hideDropdown();
            const chipBar = this.search._chipBar;
            const composed = chipBar?.getRevealModeQuery();
            if (composed) {
                chipBar.clearRevealMode();
                this.search.search(composed);
            } else {
                const q = document.getElementById('drawer-search-input')?.value?.trim();
                if (q) this.search.search(q);
            }
        });

        this._initSubDrawerEvents();
    }

    _initSubDrawerEvents() {
        const openSub = (panelId, title) => {
            document.querySelectorAll('#sub-content-legend, #sub-content-display, #sub-content-scenario').forEach(el => { el.style.display = 'none'; });
            const panel = document.getElementById(panelId);
            if (panel) panel.style.display = '';
            const titleEl = document.getElementById('sub-drawer-title');
            if (titleEl) titleEl.textContent = title;
            document.getElementById('side-drawer')?.classList.add('sub-open');
        };
        const closeSub = () => {
            document.getElementById('side-drawer')?.classList.remove('sub-open');
        };

        document.getElementById('act-legend')?.addEventListener('click', () => {
            openSub('sub-content-legend', '🎨 Legenda');
        });
        document.getElementById('act-scenario')?.addEventListener('click', () => {
            openSub('sub-content-display', '🖥️ Display');
        });
        document.getElementById('act-scenario-mgr')?.addEventListener('click', () => {
            openSub('sub-content-scenario', '🗂️ Scenario');
        });
        document.getElementById('sub-back')?.addEventListener('click', closeSub);
        document.getElementById('sub-close')?.addEventListener('click', closeSideDrawer);

        document.getElementById('toggle-color-role')?.addEventListener('change', (e) => {
            if (e.target.checked) this.legend.setMode(ROLE_FIELD_WITH_MAPPING);
        });
        document.getElementById('toggle-color-company')?.addEventListener('change', (e) => {
            if (e.target.checked) this.legend.setMode('Company');
        });
        document.getElementById('toggle-color-location')?.addEventListener('change', (e) => {
            if (e.target.checked) this.legend.setMode('Location');
        });
        document.getElementById('toggle-color-function')?.addEventListener('change', (e) => {
            if (e.target.checked) this.legend.setMode(BUSINESS_FUNCTION_FIELD);
        });

        document.getElementById('toggle-dark-mode')?.addEventListener('change', (e) => {
            applyTheme(e.target.checked ? 'dark' : 'light');
            this.legend.recolor(this.legend.colorBy);
        });

        document.getElementById('act-collapse-all')?.addEventListener('click', () => {
            this.renderer.collapseAll();
        });
        document.getElementById('act-expand-all')?.addEventListener('click', () => {
            this.renderer.expandAll();
        });

        document.getElementById('act-scenario-save')?.addEventListener('click', () => {
            this.scenario.handleAction('save');
        });
        document.getElementById('act-scenario-import')?.addEventListener('click', () => {
            this.scenario.handleAction('import');
        });
        document.getElementById('act-scenario-export')?.addEventListener('click', () => {
            this.scenario.handleAction('export');
        });
        document.getElementById('act-scenario-reset')?.addEventListener('click', () => {
            this.scenario.handleAction('reset');
        });
    }

    _openDetailFromSearch(searchParam) {
        const colonIdx = (searchParam ?? '').indexOf(':');
        if (colonIdx < 0) return;
        const field = searchParam.slice(0, colonIdx).trim().toLowerCase();
        const rawValue = searchParam.slice(colonIdx + 1).replace(/^"|"$/g, '').trim();
        const valueLower = rawValue.toLowerCase();

        if (field === 'stream') {
            const titleEl = Array.from(document.querySelectorAll('text.stream-title[data-full-name]'))
                .find(el => (el.getAttribute('data-full-name') || '').toLowerCase() === valueLower);
            if (!titleEl) return;
            const group = titleEl.closest('g[data-key^="stream::"]');
            const description = group?.getAttribute('data-description') ?? '';
            this.drawer.open({ name: rawValue, description, _permalinkSearch: searchParam, _showDetails: true });

        } else if (field === 'theme') {
            const titleEl = Array.from(document.querySelectorAll('text.theme-title[data-full-name]'))
                .find(el => (el.getAttribute('data-full-name') || '').toLowerCase() === valueLower);
            if (!titleEl) return;
            const group = titleEl.closest('g[data-key^="theme::"]');
            const description = group?.getAttribute('data-description') ?? '';
            this.drawer.open({ name: rawValue, description, _permalinkSearch: searchParam, _showDetails: true });

        } else if (field === 'team') {
            const teamTitleEl = Array.from(document.querySelectorAll('text.team-title[data-full-name]'))
                .find(el => (el.getAttribute('data-full-name') || '').toLowerCase() === valueLower);
            if (!teamTitleEl) return;
            const description = teamTitleEl.getAttribute('data-team-description') || '';
            const email = teamTitleEl.getAttribute('data-team-email') || '';
            const channels = (() => {
                try { return JSON.parse(teamTitleEl.getAttribute('data-team-channels') || '[]'); }
                catch { return []; }
            })();
            const services = (teamTitleEl.getAttribute('data-services') || '')
                .split(',').map(s => s.trim()).filter(Boolean);
            this.drawer.open({
                name: rawValue,
                description,
                elements: services.length ? { items: services } : undefined,
                channels,
                email,
                elementsBaseUrl: (s) => `domino.html?search=id%3A"${encodeURIComponent(s)}"`,
                _permalinkSearch: searchParam,
                _showDetails: true,
            });
        }
    }

    _openPersonFromParam(email) {
        if (!email || !this.personDrawer) return;
        const members = this.db.getMembersByEmail(email);
        if (!members.length) return;
        const member = members[0];
        const cards = Array.from(document.querySelectorAll('g[data-key^="card::"]'));
        const cardEl = cards.find(el => (el.getAttribute('data-email') || '').toLowerCase() === email.toLowerCase());
        if (cardEl) this.renderer.fitElementToView(cardEl, 750);
        this.personDrawer.open(member);
    }
}
