import { splitValues } from './utils.js';

const DROPDOWN_ID = 'ac-dropdown';
const MAX_KEYS = 50;
const MAX_VALUES_PER_KEY = 2000;
const MAX_OPTIONS = 25;
const TAIL_LENGTH = 44;

export class AutocompleteEngine {
    // Can be constructed two ways:
    //   new AutocompleteEngine(keys, valuesByKey, opts)  — pre-built index (Solitaire, new Domino)
    //   new AutocompleteEngine(app)                      — legacy: call buildIndex() after construction
    // opts.allowMultiValue (default true) — when false, comma-based multi-value suggestions are disabled
    constructor(keysOrApp, valuesByKey, opts = {}) {
        if (keysOrApp && typeof keysOrApp === 'object' && !Array.isArray(keysOrApp) && keysOrApp.store) {
            // Legacy app-based construction
            this._app = keysOrApp;
            this.keys = [];
            this.valuesByKey = new Map();
            this.allowMultiValue = true;
            this.allowMultiField = opts.allowMultiField === true;
        } else {
            this._app = null;
            this.keys = keysOrApp ?? [];
            this.valuesByKey = valuesByKey ?? new Map();
            this.allowMultiValue = opts.allowMultiValue !== false;
            this.allowMultiField = opts.allowMultiField === true;
        }
        this._input = null;
        this._dropdown = null;
        this._activeIdx = -1;
        this._suggestions = [];
        this._docMousedownHandler = null;
        this._docKeydownHandler = null;
    }

    // Legacy method used by Domino tests and old DominoApp usage
    buildIndex() {
        if (!this._app) return;
        const { nodes } = this._app.store;
        if (!nodes || !nodes.length) return;

        const keys = new Set(['id']);
        Object.keys(nodes[0] || {}).forEach(k => {
            if (!k) return;
            if (['index', 'x', 'y', 'vx', 'vy', 'fx', 'fy', 'color'].includes(k)) return;
            keys.add(k === 'Service Name' ? 'id' : k);
        });
        this.keys = Array.from(keys).slice(0, MAX_KEYS);

        const m = new Map();
        this.keys.forEach(k => {
            const set = new Set();
            if (k === 'id') {
                for (const n of nodes) {
                    if (set.size >= MAX_VALUES_PER_KEY) break;
                    const v = (n?.id ?? '').toString().trim();
                    if (v) set.add(v);
                }
            } else {
                for (const n of nodes) {
                    if (set.size >= MAX_VALUES_PER_KEY) break;
                    const raw = n?.[k];
                    if (typeof raw !== 'string' || !raw) continue;
                    for (const p of splitValues(raw).map(s => (s ?? '').toString().trim()).filter(Boolean)) {
                        if (set.size >= MAX_VALUES_PER_KEY) break;
                        set.add(p);
                    }
                }
            }
            m.set(k, Array.from(set).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' })));
        });
        this.valuesByKey = m;
    }

    _splitMultiFieldInput(value) {
        let inQuote = false;
        let lastAmp = -1;
        for (let i = 0; i < value.length; i++) {
            const ch = value[i];
            if (ch === '"') inQuote = !inQuote;
            if (ch === '&' && !inQuote) lastAmp = i;
        }
        if (lastAmp === -1) return { prefix: '', current: value };
        return {
            prefix: value.slice(0, lastAmp + 1),
            current: value.slice(lastAmp + 1),
        };
    }

    _computeSingleClauseSuggestions(raw) {
        const value = (raw ?? '').toString();
        if (!value.includes(':')) {
            const leading = (value.match(/^\s*/)?.[0] ?? '');
            const trimmed = value.trimStart();
            const bang = trimmed.startsWith('!') ? '!' : '';
            const keyPrefix = (bang ? trimmed.slice(1) : trimmed).trim();
            const keys = this.keys.length ? this.keys : ['id'];
            const keySuggestions = keys
                .filter(k => k.toLowerCase().startsWith(keyPrefix.toLowerCase()))
                .map(k => `${leading}${bang}${k}:`);
            const idValues = this.valuesByKey.get('id') || [];
            const nameSuggestions = keyPrefix
                ? idValues
                    .filter(v => v.toLowerCase().startsWith(keyPrefix.toLowerCase()))
                    .map(v => `${leading}${bang}${v}`)
                : [];
            return [...nameSuggestions, ...keySuggestions];
        }

        const colonPos = value.indexOf(':');
        const keyTokenOriginal = value.slice(0, colonPos);
        const key = keyTokenOriginal.trim().replace(/^!/, '');
        const valuePartFull = value.slice(colonPos + 1);
        const vParts = this.allowMultiValue ? valuePartFull.split(',') : [valuePartFull];
        const lastValRaw = (vParts.pop() ?? '');
        const preValsRaw = this.allowMultiValue ? vParts.join(',') : '';
        const lastLeadingWs = (lastValRaw.match(/^\s*/)?.[0] ?? '');
        const quotedMode = valuePartFull.includes('"');
        const lastValClean = lastValRaw.trim().replace(/^"/, '').replace(/"$/, '');
        const values = this.valuesByKey.get(key) || [];
        const filtered = values.filter(v => v.toLowerCase().startsWith(lastValClean.toLowerCase()));
        const prefix = `${keyTokenOriginal}:${preValsRaw}${preValsRaw ? ',' : ''}${lastLeadingWs}`;
        const renderValue = (v) => quotedMode ? `"${v}"` : v;
        const out = filtered.map(v => prefix + renderValue(v));

        if (this.allowMultiValue) {
            if (lastValClean && values.some(v => v.toLowerCase() === lastValClean.toLowerCase()) && !value.trimEnd().endsWith(',')) {
                out.unshift(value.trimEnd() + ',');
            }
            if (value.trimEnd().endsWith(',')) {
                const prefixAfterComma = `${keyTokenOriginal}:${preValsRaw}${preValsRaw ? ',' : ''}`;
                out.unshift(...values.slice(0, MAX_OPTIONS).map(v => prefixAfterComma + renderValue(v)));
            }
        }
        return out;
    }

    computeSuggestions(raw) {
        const value = (raw ?? '').toString();

        // When ChipBar is in per-clause edit mode, only suggest completions for the
        // clause being edited — never prepend "&" or show multi-field suggestions.
        const clauseEditMode = this._input?.dataset?.clauseEditMode === '1';

        if (!this.allowMultiField || clauseEditMode) {
            return this._computeSingleClauseSuggestions(value);
        }

        const { prefix, current } = this._splitMultiFieldInput(value);

        // Read active fields set by ChipBar on the input element
        const activeFieldsAttr = this._input?.dataset?.activeFields ?? '';
        const activeFields = activeFieldsAttr ? new Set(activeFieldsAttr.split(',').map(s => s.trim()).filter(Boolean)) : new Set();

        // When the current clause has no colon yet (user typing a field name),
        // split key suggestions into: active fields first (same-clause continuation with ,)
        // and new fields after (new clause with &).
        if (!current.includes(':') && activeFields.size > 0 && !prefix) {
            const leading = (current.match(/^\s*/)?.[0] ?? '');
            const trimmed = current.trimStart();
            const bang = trimmed.startsWith('!') ? '!' : '';
            const keyPrefix = (bang ? trimmed.slice(1) : trimmed).trim();
            const keys = this.keys.length ? this.keys : ['id'];
            const matchingKeys = keys.filter(k => k.toLowerCase().startsWith(keyPrefix.toLowerCase()));

            const activeKeysSuggestions = matchingKeys
                .filter(k => activeFields.has(k))
                .map(k => `${leading}${bang}${k}:`);
            const newKeysSuggestions = matchingKeys
                .filter(k => !activeFields.has(k))
                .map(k => `& ${leading}${bang}${k}:`);
            const idValues = this.valuesByKey.get('id') || [];
            const nameSuggestions = keyPrefix
                ? idValues
                    .filter(v => v.toLowerCase().startsWith(keyPrefix.toLowerCase()))
                    .map(v => `${leading}${bang}${v}`)
                : [];

            return Array.from(new Set([...nameSuggestions, ...activeKeysSuggestions, ...newKeysSuggestions]));
        }

        const out = [];
        const currentSuggestions = this._computeSingleClauseSuggestions(current)
            .map(s => `${prefix}${s}`);
        out.push(...currentSuggestions);

        const trimmed = value.trimEnd();
        const currentTrimmed = current.trimEnd();
        const hasKeyValue = currentTrimmed.includes(':') && currentTrimmed.slice(currentTrimmed.indexOf(':') + 1).trim().length > 0;
        if (hasKeyValue && !trimmed.endsWith('&')) {
            // Suggest adding another field once the current field has at least one value.
            out.unshift(`${trimmed}&`);
        }

        // Remove duplicates while keeping ranking.
        return Array.from(new Set(out));
    }

    refreshSuggestions(raw) {
        if (!this._dropdown || !this._input) return;

        const rawStr = (raw ?? '').toString();
        this._suggestions = this.computeSuggestions(rawStr)
            .slice(0, MAX_OPTIONS);
        this._activeIdx = -1;
        this._dropdown.innerHTML = '';

        if (!this._suggestions.length) {
            this._hideDropdown();
            return;
        }

        this._suggestions.forEach((s, i) => {
            const li = document.createElement('li');
            li.className = 'ac-item';
            li.title = s;
            li.dataset.value = s;
            li.textContent = s.length > TAIL_LENGTH ? '…' + s.slice(-(TAIL_LENGTH)) : s;
            li.addEventListener('mousedown', (e) => {
                e.preventDefault();
                this._selectSuggestion(i);
            });
            li.addEventListener('mousemove', () => {
                this._setActive(i);
            });
            this._dropdown.appendChild(li);
        });

        this._showDropdown();
        this._repositionDropdown();
    }

    _selectSuggestion(idx) {
        const s = this._suggestions[idx];
        if (!s || !this._input) return;
        const isConfirm = s === this._input.value;
        this._input.value = s;
        if (isConfirm) {
            // Value unchanged — user picked the "confirm" entry. Close dropdown and
            // signal the app to trigger a search via a dedicated custom event.
            this._hideDropdown();
            this._input.dispatchEvent(new CustomEvent('ac-confirm', { bubbles: true }));
        } else {
            this._input.dispatchEvent(new Event('input', { bubbles: true }));
        }
        this._input.focus();
    }

    _setActive(idx) {
        const items = this._dropdown.querySelectorAll('.ac-item');
        items.forEach((el, i) => {
            el.classList.toggle('ac-active', i === idx);
        });
        this._activeIdx = idx;
    }

    _moveActive(dir) {
        const count = this._suggestions.length;
        if (!count) return;
        let next = this._activeIdx + dir;
        if (next < 0) next = count - 1;
        if (next >= count) next = 0;
        this._setActive(next);
        // Scroll active item into view
        const items = this._dropdown.querySelectorAll('.ac-item');
        items[next]?.scrollIntoView({ block: 'nearest' });
    }

    _showDropdown() {
        this._dropdown.classList.add('ac-open');
    }

    _hideDropdown() {
        if (!this._dropdown) return;
        this._dropdown.classList.remove('ac-open');
        this._activeIdx = -1;
        this._dropdown.querySelectorAll('.ac-item').forEach(el => el.classList.remove('ac-active'));
    }

    // Public wrapper for external callers (e.g. after committing a search).
    hideDropdown() {
        this._hideDropdown();
    }

    // Returns true when the dropdown is open and a suggestion is actively highlighted.
    // Used by app-level Enter handlers to skip firing a search while the user is
    // still picking from the suggestion list.
    hasPendingSelection() {
        return this._dropdown?.classList.contains('ac-open') && this._activeIdx >= 0;
    }

    _repositionDropdown() {
        if (!this._input || !this._dropdown) return;
        // Find the search bar container to position below it
        const bar = this._input.closest('.top-search') || this._input;
        const barRect = bar.getBoundingClientRect();
        this._dropdown.style.top = `${barRect.bottom + 2}px`;
        this._dropdown.style.left = `${barRect.left}px`;
        this._dropdown.style.minWidth = `${barRect.width}px`;
    }

    init() {
        const input = document.getElementById('drawer-search-input');
        if (!input) return;
        this._input = input;

        // Remove any legacy datalist
        const oldDl = document.getElementById('search-suggestions');
        if (oldDl) oldDl.remove();
        // Remove old list attribute and disable browser native autocomplete/autofill.
        input.removeAttribute('list');
        input.setAttribute('autocomplete', 'off');

        let dropdown = document.getElementById(DROPDOWN_ID);
        if (!dropdown) {
            dropdown = document.createElement('ul');
            dropdown.id = DROPDOWN_ID;
            dropdown.className = 'ac-dropdown';
            dropdown.setAttribute('role', 'listbox');
            document.body.appendChild(dropdown);
        }
        this._dropdown = dropdown;

        const update = () => this.refreshSuggestions(input.value || '');

        if (!input.dataset.boundAutocomplete) {
            input.dataset.boundAutocomplete = '1';

            input.addEventListener('input', update);
            input.addEventListener('focus', update);
            input.addEventListener('click', update);

            input.addEventListener('blur', () => {
                // Delay so mousedown on a suggestion fires first (150 ms).
                // After the delay, keep the dropdown open if focus stayed within
                // the search bar (e.g. moved to a chip via Alt+Arrow navigation).
                setTimeout(() => {
                    const active = document.activeElement;
                    const bar = input.closest('.top-search') || input;
                    if (bar.contains(active) || this._dropdown?.contains(active)) return;
                    this._hideDropdown();
                }, 150);
            });

            input.addEventListener('keydown', (e) => {
                if (!this._dropdown.classList.contains('ac-open')) return;
                if (e.key === 'ArrowDown') {
                    e.preventDefault();
                    this._moveActive(1);
                } else if (e.key === 'ArrowUp') {
                    e.preventDefault();
                    this._moveActive(-1);
                } else if (e.key === 'Enter') {
                    if (this._activeIdx >= 0) {
                        e.preventDefault();
                        e.stopPropagation();
                        this._selectSuggestion(this._activeIdx);
                    } else {
                        // User pressed Enter without selecting a suggestion — hide the dropdown
                        // so the search can proceed.
                        this._hideDropdown();
                    }
                } else if (e.key === 'Escape') {
                    this._hideDropdown();
                    e.preventDefault();
                    e.stopPropagation();
                }
            });
        }

        // Close dropdown when clicking outside the search bar or dropdown.
        // Remove any previous handler first (e.g. after re-init in Solitaire).
        if (this._docMousedownHandler) {
            document.removeEventListener('mousedown', this._docMousedownHandler);
        }
        this._docMousedownHandler = (e) => {
            if (!this._dropdown.classList.contains('ac-open')) return;
            const bar = input.closest('.top-search') || input;
            // Use composedPath() so the check works even after refreshSuggestions() has
            // replaced the dropdown's innerHTML (detached nodes are no longer reachable
            // via contains(), but the original event path is frozen at dispatch time).
            const path = e.composedPath ? e.composedPath() : [];
            const inBar = path.includes(bar) || bar.contains(e.target);
            const inDropdown = path.includes(this._dropdown) || this._dropdown.contains(e.target);
            if (!inBar && !inDropdown) {
                this._hideDropdown();
            }
        };
        document.addEventListener('mousedown', this._docMousedownHandler);

        // Close dropdown on ESC from anywhere (e.g. when focus is on a chip).
        if (this._docKeydownHandler) {
            document.removeEventListener('keydown', this._docKeydownHandler, true);
        }
        this._docKeydownHandler = (e) => {
            if (e.key === 'Escape' && this._dropdown.classList.contains('ac-open')) {
                this._hideDropdown();
                e.preventDefault();
                e.stopPropagation();
            }
        };
        // Use capture phase so this fires before other document/window bubble-phase
        // Escape handlers (e.g. DetailDrawer, SolitaireApp) that would clear the search.
        document.addEventListener('keydown', this._docKeydownHandler, true);

        // Do NOT call update() here — dropdown should only appear after user interaction.
    }
}
