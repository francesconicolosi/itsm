import { JengaMultiSelect } from './JengaMultiSelect.js';

const SEARCH_FILTER_STORAGE_KEY = 'jenga.searchFilters.v2';

const ENV_OPTIONS = ['Production', 'QA', 'Staging', 'Dev'];

// Coloured badge meta for env chips in the trigger
const ENV_CHIP_META = {
    'Production': { label: 'PRD', color: '#ef4444' },
    'QA':         { label: 'QA',  color: '#eab308', textColor: '#1a1a1a' },
    'Staging':    { label: 'STG', color: '#6b7280' },
    'Dev':        { label: 'DEV', color: '#3b82f6' },
};

// Maps display name → env key prefixes used in event data
const ENV_DISPLAY_TO_KEYS = {
    'Production': ['prd', 'prod'],
    'QA':         ['qa', 'qa1', 'qa2', 'qa3', 'qa4', 'qa5', 'qasales', 'qaesales'],
    'Staging':    ['staging', 'stg', 'preprod', 'shadow'],
    'Dev':        ['devops'],
};

const TYPE_OPTIONS = [
    'Service Operations',
    'Business Events',
    'Hybris Releases',
    'Milestones',
];

// Maps display name → internal event type key
const TYPE_DISPLAY_TO_KEY = {
    'Service Operations': 'SERVICE_OP',
    'Business Events':    'BUSINESS_EVENT',
    'Hybris Releases':    'HYBRIS',
    'Milestones':         'MILESTONE',
};

export class JengaSearch {
    constructor(app) {
        this.app = app;

        this._state = this._loadFilterState();

        // Raw UI state. Empty array is intentionally preserved as an explicit user state.
        this._selectedEnvs     = new Set(this._state.envs.selected);
        this._selectedTypes    = new Set(this._state.types.selected);
        this._selectedServices = new Set(this._state.services.selected);

        // Effective filtering state.
        // Empty selection = no restriction, but the raw empty selection is still stored and restored.
        this.activeEnvs     = new Set(this._selectedEnvs);
        this.activeTypes    = this._toEffectiveTypeSet(this._selectedTypes);
        this.activeServices = new Set(this._selectedServices);
        this.query          = this._state.query || '';

        this._envMs  = null; // JengaMultiSelect for environment
        this._typeMs = null; // JengaMultiSelect for event types
        this._svcMs  = null; // JengaMultiSelect for service
        this._serviceOptions = []; // populated after catalog loads
    }

    _defaultFilterState() {
        return {
            version: 2,
            query: '',
            envs:     { selected: [] },
            types:    { selected: [] },
            services: { selected: [] },
        };
    }

    _normalizeSelection(values, validOptions = null) {
        const arr = Array.isArray(values) ? values : [];
        const valid = validOptions ? new Set(validOptions) : null;
        return [...new Set(arr)].filter(v => typeof v === 'string' && (!valid || valid.has(v)));
    }

    _loadFilterState() {
        const defaults = this._defaultFilterState();
        try {
            const raw = JSON.parse(localStorage.getItem(SEARCH_FILTER_STORAGE_KEY) || '{}') || {};

            // Backward compatibility with the first persisted-filters attempt, if present.
            const legacyTypes = Array.isArray(raw.types) ? raw.types : raw.types?.selected;
            const legacyEnvs = Array.isArray(raw.envs) ? raw.envs : raw.envs?.selected;
            const legacyServices = Array.isArray(raw.services) ? raw.services : raw.services?.selected;

            return {
                version: 2,
                query: typeof raw.query === 'string' ? raw.query : defaults.query,
                envs: {
                    selected: this._normalizeSelection(legacyEnvs, ENV_OPTIONS),
                },
                types: {
                    selected: this._normalizeSelection(legacyTypes, TYPE_OPTIONS),
                },
                services: {
                    // Service options are loaded later, so do not validate them here.
                    selected: this._normalizeSelection(legacyServices),
                },
            };
        } catch (_) {
            return defaults;
        }
    }

    _saveFilterState() {
        try {
            const state = {
                version: 2,
                query: this.query || '',
                envs: {
                    selected: [...this._selectedEnvs],
                },
                types: {
                    selected: [...this._selectedTypes],
                },
                services: {
                    selected: [...this._selectedServices],
                },
            };
            localStorage.setItem(SEARCH_FILTER_STORAGE_KEY, JSON.stringify(state));
        } catch (_) {}
    }

    _toEffectiveTypeSet(displaySelection) {
        if (!displaySelection || displaySelection.size === 0) {
            return new Set(Object.values(TYPE_DISPLAY_TO_KEY));
        }
        return new Set([...displaySelection].map(d => TYPE_DISPLAY_TO_KEY[d]).filter(Boolean));
    }

    _resetAllFilters({ persist = true } = {}) {
        this.query = '';
        this._selectedEnvs.clear();
        this._selectedTypes.clear();
        this._selectedServices.clear();

        this.activeEnvs.clear();
        this.activeTypes = new Set(Object.values(TYPE_DISPLAY_TO_KEY));
        this.activeServices.clear();

        this._envMs?.setValue(this._selectedEnvs);
        this._typeMs?.setValue(this._selectedTypes);
        this._svcMs?.setValue(this._selectedServices);

        if (persist) this._saveFilterState();
    }

    /** Called once from JengaApp.init() after data is loaded. */
    init() {
        // Search input
        const input = document.getElementById('jenga-search-input');
        if (input) {
            input.value = this.query || '';

            input.addEventListener('input', (e) => {
                this.query = e.target.value.trim();
                this._saveFilterState();
                this.app.refresh();
            });
            input.addEventListener('keydown', (e) => {
                if (e.key === 'Escape') {
                    this.query = '';
                    input.value = '';
                    this._saveFilterState();
                    this.app.refresh();
                }
            });
        }

        document.getElementById('jenga-search-go')?.addEventListener('click', () => {
            this.query = input?.value?.trim() || '';
            this._saveFilterState();
            this.app.refresh();
        });

        document.getElementById('act-clear')?.addEventListener('click', () => {
            this._resetAllFilters({ persist: true });
            if (input) input.value = '';
            this.app.refresh();
        });

        // Build multi-select dropdowns in top bar
        this._buildTopBarFilters();
    }

    /** Called by JengaApp once the service catalog is loaded. */
    setServiceOptions(serviceNames) {
        this._serviceOptions = serviceNames || [];

        // Drop stale service selections that no longer exist in the current catalog.
        const valid = new Set(this._serviceOptions);
        let changed = false;
        for (const svc of [...this._selectedServices]) {
            if (!valid.has(svc)) {
                this._selectedServices.delete(svc);
                changed = true;
            }
        }
        this.activeServices = new Set(this._selectedServices);
        if (changed) this._saveFilterState();

        if (this._svcMs) {
            this._svcMs.setOptions(this._serviceOptions);
            this._svcMs.setValue(this._selectedServices);
        }
    }

    _buildTopBarFilters() {
        // Env and Type go into the "primary" slot (row 2 on mobile)
        // Service goes into the "service" slot (row 3 on mobile)
        // Fallback to legacy #jenga-filters slot if the split layout isn't present
        const primarySlot = document.getElementById('jenga-filters-primary')
            || document.getElementById('jenga-filters');
        const serviceSlot = document.getElementById('jenga-filters-service')
            || primarySlot;
        if (!primarySlot) return;

        // Environment multi-select — uses coloured badge chips in the trigger
        this._envMs = new JengaMultiSelect({
            id: 'jms-env',
            placeholder: 'Environment',
            options: ENV_OPTIONS,
            initial: new Set(this._selectedEnvs),
            allByDefault: true,
            onChange: (sel) => {
                // Store raw UI state exactly as selected by the user.
                this._selectedEnvs = new Set(sel);
                this.activeEnvs = new Set(sel);
                this._saveFilterState();
                this.app.refresh();
            },
            renderChip: (val) => {
                const meta = ENV_CHIP_META[val] || { label: val, color: '#999' };
                const chip = document.createElement('span');
                chip.className = 'jms__env-chip';
                chip.textContent = meta.label;
                chip.style.background = meta.color;
                chip.style.color = meta.textColor || '#fff';
                return chip;
            }
        });
        primarySlot.appendChild(this._envMs.getEl());

        // Type multi-select
        this._typeMs = new JengaMultiSelect({
            id: 'jms-type',
            placeholder: 'Event Types',
            options: TYPE_OPTIONS,
            initial: new Set(this._selectedTypes),
            allByDefault: true,
            onChange: (sel) => {
                // Store raw display values exactly. Empty selection remains empty in storage,
                // even though it is interpreted as "all types" by the data filter.
                this._selectedTypes = new Set(sel);
                this.activeTypes = this._toEffectiveTypeSet(this._selectedTypes);
                this._saveFilterState();
                this.app.refresh();
            },
        });
        primarySlot.appendChild(this._typeMs.getEl());

        // Service multi-select — separate slot on mobile, same row on desktop
        this._svcMs = new JengaMultiSelect({
            id: 'jms-svc',
            placeholder: 'Service',
            options: this._serviceOptions,
            initial: new Set(this._selectedServices),
            allByDefault: true,
            onChange: (sel) => {
                this._selectedServices = new Set(sel);
                this.activeServices = new Set(sel);
                this._saveFilterState();
                this.app.refresh();
            },
        });
        serviceSlot.appendChild(this._svcMs.getEl());
    }

    getUrlParams() {
        const out = {};
        if (this._selectedEnvs.size)    out.envs     = [...this._selectedEnvs].join(',');
        if (this._selectedTypes.size)   out.types    = [...this._selectedTypes].join(',');
        if (this._selectedServices.size) out.services = [...this._selectedServices].join(',');
        return out;
    }

    applyUrlParams(p) {
        const parse = (key, valid) => {
            const raw = p.get(key);
            if (!raw) return null;
            const vals = raw.split(',').map(v => decodeURIComponent(v.trim())).filter(Boolean);
            return valid ? vals.filter(v => valid.has(v)) : vals;
        };

        const envs  = parse('envs',  new Set(ENV_OPTIONS));
        const types = parse('types', new Set(TYPE_OPTIONS));
        const svcs  = parse('services');

        if (envs !== null) {
            this._selectedEnvs = new Set(envs);
            this.activeEnvs    = new Set(envs);
            this._envMs?.setValue(this._selectedEnvs, { emit: false });
        }
        if (types !== null) {
            this._selectedTypes = new Set(types);
            this.activeTypes    = this._toEffectiveTypeSet(this._selectedTypes);
            this._typeMs?.setValue(this._selectedTypes, { emit: false });
        }
        if (svcs !== null) {
            this._selectedServices = new Set(svcs);
            this.activeServices    = new Set(svcs);
            this._svcMs?.setValue(this._selectedServices, { emit: false });
        }

        this._saveFilterState();
    }

    getFilteredEvents(allEvents) {
        return allEvents.filter(e => {
            // Type filter. Empty raw type selection produces activeTypes = all types.
            if (this.activeTypes.size > 0 && !this.activeTypes.has(e.type)) return false;

            // Business events, milestones, and Hybris releases bypass env/service filters
            const bypassFilters = e.type === 'BUSINESS_EVENT' || e.type === 'MILESTONE' || e.type === 'HYBRIS';

            // Environment multi-select filter (applies to SERVICE_OP only)
            if (!bypassFilters && this.activeEnvs.size > 0 && e.type === 'SERVICE_OP') {
                const env = (e.environment || '').toLowerCase();
                const allowed = [...this.activeEnvs].some(displayName => {
                    const keys = ENV_DISPLAY_TO_KEYS[displayName] || [];
                    return keys.some(k => env === k || env.startsWith(k));
                });
                if (!allowed) return false;
            }

            // Service multi-select filter (skipped for biz events and milestones)
            if (!bypassFilters && this.activeServices.size > 0) {
                const svc = (e.service || '').trim();
                if (!this.activeServices.has(svc)) return false;
            }

            // Text search
            if (this.query) {
                const q   = this.query.toLowerCase();
                const hay = `${e.summary} ${e.service} ${e.environment} ${e.operation} ${e.stream} ${e.theme}`.toLowerCase();
                if (!hay.includes(q)) return false;
            }

            return true;
        });
    }
}
