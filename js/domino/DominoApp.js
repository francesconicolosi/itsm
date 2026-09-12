import * as d3 from 'd3';
import { getQueryParam, setSearchQuery, initCommonActions, closeSideDrawer, enableGlobalFindShortcut, applyTheme, loadSavedTheme } from '../shared/utils.js';
import { BRAND, renderBrandLogo } from '../../brand-specific/brand.js';
import { ServiceCatalogStore } from './ServiceCatalogStore.js';
import { SearchEngine } from './SearchEngine.js';
import { AutocompleteEngine } from '../shared/AutocompleteEngine.js';
import { GraphRenderer } from './GraphRenderer.js';
import { ListView } from './ListView.js';
import { DetailDrawer } from './DetailDrawer.js';
import { DominoLegend } from './DominoLegend.js';
import { PostItNote } from '../shared/PostItNote.js';
import { AnnouncementBar } from '../shared/AnnouncementBar.js';

export class DominoApp {
    constructor() {
        this.store = new ServiceCatalogStore();
        this.search = new SearchEngine(this);
        this.autocomplete = new AutocompleteEngine(this);
        this.autocomplete.allowMultiField = true;
        this.graph = new GraphRenderer(this);
        this.legend = new DominoLegend(this);
        this.listView = new ListView(this);
        this.drawer = new DetailDrawer(this);
        this.postIt = new PostItNote('domino');
        this.announcementBar = new AnnouncementBar();
    }

    init() {
        this.announcementBar.init();
        renderBrandLogo();
        this.graph.initDOM();
        this.listView.initDOM();
        this.drawer.initDOM();
        this.drawer.onClose = () => this.graph.clearSelection();

        const toggleDecommissioned = document.getElementById('toggle-decommissioned');
        if (toggleDecommissioned) {
            this.search.hideStoppedServices = !toggleDecommissioned.checked;
        }

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

        const theme = loadSavedTheme();
        const dmToggle = document.getElementById('toggle-dark-mode');
        if (dmToggle) dmToggle.checked = theme === 'dark';

        this._initSideDrawerEvents();
        this.search.initChipBar();
        this._initFileUpload();
        this._initKeyboardShortcut();
        this._initLoadEvent();
        this.postIt.init();
        this.postIt.attachContextMenu(document.body);
    }

    _initSideDrawerEvents() {
        initCommonActions();
        this.search.initRelaxedSearchPersistence();
        this.search.initShowConnectionsPersistence();
        this._ensureUploadCsvAction();

        document.getElementById('act-introduce')?.addEventListener('click', () => {
            window.open(BRAND.urls.jiraNewRequest, '_blank', 'noopener');
        });
        document.getElementById('act-update')?.addEventListener('click', () => {
            window.open(BRAND.urls.jiraUpdateRequest, '_blank', 'noopener');
        });

        document.getElementById('act-about')?.addEventListener('click', () => {
            closeSideDrawer();
            this.drawer.showAbout();
        });

        document.getElementById('act-clear')?.addEventListener('click', () => {
            this.graph.clearSelection();
            this.search.searchTerm = '';
            this.search._refreshChips();
            const input = document.getElementById('drawer-search-input');
            if (input) input.value = '';
            setSearchQuery('');
            this.graph.updateVisualization();
            this.graph.fitGraphToViewport(0.9);
            closeSideDrawer();
        });

        document.getElementById('toggle-decommissioned')?.addEventListener('change', (e) => {
            this.graph.clearSelection();
            this.search.hideStoppedServices = !e.target.checked;
            this.graph.updateVisualization();
        });

        document.getElementById('toggle-dark-mode')?.addEventListener('change', (e) => {
            applyTheme(e.target.checked ? 'dark' : 'light');
            this.graph.updateGraphTheme();
        });

        document.getElementById('act-fit')?.addEventListener('click', () => {
            this.graph.fitGraphToViewport(0.9);
            closeSideDrawer();
        });

        document.getElementById('drawer-search-go')?.addEventListener('click', () => {
            this.autocomplete?.hideDropdown();
            const chipBar = this.search._chipBar;
            const composed = chipBar?.getRevealModeQuery();
            if (composed) {
                chipBar.clearRevealMode();
                this.search.handleQuery(composed, false);
            } else {
                const q = document.getElementById('drawer-search-input')?.value?.trim();
                if (q !== undefined) this.search.handleQuery(q, false);
            }
        });

        document.getElementById('drawer-search-input')?.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                if (this.autocomplete?.hasPendingSelection()) return;
                const chipBar = this.search._chipBar;
                const composed = chipBar?.getRevealModeQuery();
                if (composed) {
                    chipBar.clearRevealMode();
                    this.search.handleQuery(composed, false);
                } else {
                    const q = e.target.value ? e.target.value.trim() : '';
                    this.search.handleQuery(q, false);
                }
                e.preventDefault();
            }
        });

        document.getElementById('drawer-search-input')?.addEventListener('ac-confirm', (e) => {
            this.autocomplete?.hideDropdown();
            const chipBar = this.search._chipBar;
            const composed = chipBar?.getRevealModeQuery();
            if (composed) {
                chipBar.clearRevealMode();
                this.search.handleQuery(composed, false);
            } else {
                const q = e.target.value ? e.target.value.trim() : '';
                if (q !== undefined) this.search.handleQuery(q, false);
            }
        });

        const advToggle =
            document.getElementById('advanced-mode') ||
            document.getElementById('toggle-advanced') ||
            document.getElementById('mode-advanced');
        if (advToggle && !advToggle.dataset.boundUploadCta) {
            advToggle.dataset.boundUploadCta = '1';
            advToggle.addEventListener('change', () => this._ensureUploadCsvAction());
        }
    }

    _isAdvancedMode() {
        try {
            const url = new URL(window.location.href);
            const qp = (url.searchParams.get('mode') || url.searchParams.get('view') || url.searchParams.get('advanced') || '').toString();
            if (qp && /^(advanced|1|true|yes)$/i.test(qp)) return true;
        } catch (_) {}
        if (document.body?.classList?.contains('advanced')) return true;
        const toggle =
            document.getElementById('advanced-mode') ||
            document.getElementById('toggle-advanced') ||
            document.getElementById('mode-advanced');
        if (toggle && 'checked' in toggle) return !!toggle.checked;
        return false;
    }

    _ensureUploadCsvAction() {
        const fileInput = document.getElementById('fileInput');
        if (!fileInput) return;
        if (!fileInput.getAttribute('accept')) fileInput.setAttribute('accept', '.csv,text/csv');
        let btn = document.getElementById('act-upload-csv') || document.getElementById('act-upload');
        if (!btn) {
            const container =
                document.getElementById('drawer-actions') ||
                document.querySelector('.drawer-actions') ||
                document.getElementById('actions') ||
                document.querySelector('.actions') ||
                document.querySelector('#drawer .drawer-actions');
            if (!container) return;
            btn = document.createElement('button');
            btn.id = 'act-upload-csv';
            btn.type = 'button';
            btn.className = 'fade-link';
            btn.textContent = 'Upload CSV';
            container.appendChild(btn);
        }
        if (!btn.dataset.boundUploadCsv) {
            btn.dataset.boundUploadCsv = '1';
            btn.addEventListener('click', () => fileInput.click());
        }
        btn.style.display = this._isAdvancedMode() ? '' : 'none';
    }

    _initFileUpload() {
        document.getElementById('fileInput')?.addEventListener('change', (event) => {
            this.graph.resetVisualization();
            this.store.reset();
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = (e) => {
                const data = d3.csvParse(e.target.result);
                this._processAndRender(data, this.store.jiraCards || []);
                this.graph.updateVisualization();
            };
            reader.readAsText(file);
        });
    }

    _initKeyboardShortcut() {
        enableGlobalFindShortcut({ inputSelector: '#drawer-search-input' });
    }

    showToast(message, duration = 3000) {
        const positionContainer = (container) => {
            const drawer = document.getElementById('drawer');
            const isOpen = drawer?.classList.contains('open');
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
            const drawer = document.getElementById('drawer');
            if (drawer) {
                new MutationObserver(() => positionContainer(container))
                    .observe(drawer, { attributes: true, attributeFilter: ['class'] });
                window.__toastDrawerObserverAttached = true;
            }
        }
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => toast.remove(), 300);
        }, duration);
    }

    _processAndRender(data, jiraCardsData = null) {
        const colorScale = this.store.processData(data);
        if (!colorScale) return;

        // Process Jira cards BEFORE createMap(), because badges are created while SVG nodes are rendered.
        if (Array.isArray(jiraCardsData)) {
            this.store.processJiraCards(jiraCardsData);
        }

        this.graph.createMap();
        this.legend.render(colorScale);
        this.autocomplete.buildIndex();
        this.autocomplete.init();
    }

    _initLoadEvent() {
        window.addEventListener('load', () => {
            const listViewParam = getQueryParam('listView');
            const sortParam = getQueryParam('sort');
            const parsedSort = this.listView.parseSortParam(sortParam);
            if (parsedSort && (!this.listView.columnKeys.length || this.listView.columnKeys.includes(parsedSort.key))) {
                this.listView.sortKey = parsedSort.key;
                this.listView.sortDir = parsedSort.dir;
            }

            const parsedCols = this.listView.parseListViewParam(listViewParam);
            if (parsedCols && parsedCols.length) {
                this.listView.columnKeys = parsedCols;
                window.currentColumnKeys = this.listView.columnKeys;
            }

            const searchParam = getQueryParam('search');
            const searchInput = document.getElementById('drawer-search-input');

            Promise.all([
                fetch(BRAND.csv.domino).then(response => response.text()),
                fetch(BRAND.csv.jiraCards || './jira-cards.csv')
                    .then(response => response.ok ? response.text() : '')
                    .catch(() => ''),
            ])
                .then(([csvData, jiraCardsCsv]) => {
                    if (searchParam) {
                        this.search.searchTerm = searchParam;
                        if (searchInput) searchInput.value = searchParam;
                        this.search._refreshChips();
                    }

                    const stripComments = s => s.replace(/^#[^\n]*\n/gm, '');
                    const data = d3.csvParse(csvData);
                    const jiraCardsData = jiraCardsCsv ? d3.csvParse(stripComments(jiraCardsCsv)) : [];
                    this._processAndRender(data, jiraCardsData);

                    const afterInit = () => {
                        const wantListView = Boolean(listViewParam);
                        if (wantListView) {
                            this.listView.toListView();
                            this.listView.syncListViewParamInUrl();
                            this.listView.syncSortParamInUrl();
                        }
                        const uniqueIds = ['id', 'Name'];
                        const showDrawer = typeof searchParam === 'string' && uniqueIds.includes(searchParam.split(':')[0]);
                        this.graph.updateVisualization(showDrawer);

                        const openDrawerParam = getQueryParam('openDrawer');
                        if ((wantListView || openDrawerParam === 'true') && showDrawer) {
                            const id = searchParam.split(':')[1]?.replace(/"/g, '');
                            const node = this.store.nodes.find(n => n.id === id);
                            if (node) this.drawer.showNodeDetails(node, true);
                        }
                        requestAnimationFrame(() => this._hideSpinner());
                    };

                    if (searchParam) {
                        this.graph.simulation.on('end', () => {
                            if (!this.store.hasLoaded) {
                                this.store.hasLoaded = true;
                                afterInit();
                            }
                        });
                    } else {
                        afterInit();
                    }
                })
                .catch(error => console.error('Error loading the CSV file:', error));
        });
    }

    _hideSpinner() {
        const el = document.getElementById('app-spinner');
        if (!el) return;
        el.classList.add('is-hidden');
        el.addEventListener('transitionend', () => el.remove(), { once: true });
    }
}
