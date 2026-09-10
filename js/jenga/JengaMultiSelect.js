/**
 * JengaMultiSelect — inline multi-select dropdown with compact chip tags.
 *
 * Renders a pill-shaped trigger that shows selected chips + a text input.
 * When many values are selected, only the latest selected values are shown
 * and the remaining values are collapsed into a "+N" chip.
 *
 * Clicking the trigger opens a dropdown with a "Select all" row and a scrollable
 * checkbox list. Emits an onChange(Set<string>) callback on every change.
 */

export class JengaMultiSelect {
    /**
     * @param {object} opts
     * @param {string}   opts.id              — unique HTML id prefix
     * @param {string}   opts.placeholder     — placeholder text when nothing selected
     * @param {string[]} opts.options         — all available option values
     * @param {Set<string>} [opts.initial]    — pre-selected values
     * @param {boolean}  [opts.allByDefault]  — true = "all" = empty set
     * @param {Function} opts.onChange        — callback(Set<string> selected)
     * @param {Function} [opts.renderChip]    — optional: (value) → HTMLElement
     */
    constructor({ id, placeholder, options, initial, allByDefault = true, onChange, renderChip }) {
        this.id           = id;
        this.placeholder  = placeholder;
        this.options      = options || [];
        this.allByDefault = allByDefault;
        this.onChange     = onChange;
        this.renderChip   = renderChip || null;

        // Empty set = "all selected" when allByDefault is true
        this.selected     = initial ? new Set(initial) : new Set();

        this._open        = false;
        this._el          = null;
        this._dropdown    = null;
        this._input       = null;
        this._chipArea    = null;
        this._clearBtn    = null;
        this._query       = '';
        this._focusedIdx  = -1;

        // UX tuning
        this._maxVisibleChips = 2;
    }

    // ── Public API ─────────────────────────────────────────────────────────────

    getEl() {
        if (!this._el) this._build();
        return this._el;
    }

    /**
     * Replace the option list and re-render the dropdown.
     */
    setOptions(options) {
        this.options = options || [];

        // Remove options that are no longer valid
        for (const s of [...this.selected]) {
            if (!this.options.includes(s)) this.selected.delete(s);
        }

        if (this._dropdown) this._rebuildList();
        this._renderChips();
    }

    /**
     * Return the effective filter set.
     * When allByDefault=true and nothing is selected, returns empty Set (= no filter).
     */
    getValue() {
        return new Set(this.selected);
    }

    /**
     * Programmatically replace the current raw selection.
     *
     * Important: an empty Set is a valid explicit state and must stay empty.
     * It is not converted to "all selected" here. The caller decides how an
     * empty selection is interpreted when filtering data.
     */
    setValue(values, { emit = false } = {}) {
        const next = values instanceof Set
            ? [...values]
            : Array.isArray(values)
                ? values
                : [];

        this.selected = new Set(next.filter(v => this.options.includes(v)));
        this._renderChips();
        if (this._dropdown) this._rebuildList();
        if (emit) this._emit();
    }

    // ── Build ──────────────────────────────────────────────────────────────────

    _build() {
        const wrap = document.createElement('div');
        wrap.className = 'jms';
        wrap.id = this.id;

        // ── Trigger pill ─────────────────────────────────────────────────────
        const trigger = document.createElement('div');
        trigger.className = 'jms__trigger';
        trigger.setAttribute('tabindex', '0');
        trigger.setAttribute('role', 'combobox');
        trigger.setAttribute('aria-expanded', 'false');

        const chipArea = document.createElement('div');
        chipArea.className = 'jms__chips';
        this._chipArea = chipArea;

        const input = document.createElement('input');
        input.className = 'jms__input';
        input.placeholder = this.placeholder;
        input.autocomplete = 'off';
        this._input = input;

        const controls = document.createElement('div');
        controls.className = 'jms__controls';

        const clearBtn = document.createElement('button');
        clearBtn.className = 'jms__clear';
        clearBtn.innerHTML = '✕';
        clearBtn.title = 'Clear all';
        clearBtn.style.display = 'none';
        this._clearBtn = clearBtn;

        const chevron = document.createElement('span');
        chevron.className = 'jms__chevron';
        chevron.innerHTML = '▾';

        controls.append(clearBtn, chevron);
        trigger.append(chipArea, input, controls);

        // ── Dropdown ──────────────────────────────────────────────────────────
        const dropdown = document.createElement('div');
        dropdown.className = 'jms__dropdown';
        dropdown.style.display = 'none';
        this._dropdown = dropdown;

        wrap.append(trigger, dropdown);
        this._el = wrap;

        // ── Events ────────────────────────────────────────────────────────────

        trigger.addEventListener('click', (e) => {
            if (e.target === clearBtn || clearBtn.contains(e.target)) return;
            this._toggleOpen();
        });

        trigger.addEventListener('keydown', (e) => {
            if (e.key === 'ArrowDown' || e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                if (!this._open) {
                    this._open_(true);
                    this._setFocus(0);
                }
            } else if (e.key === 'Escape') {
                this._open_(false);
            }
        });

        clearBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.selected.clear();
            this._renderChips();
            this._rebuildList();
            this._emit();
        });

        input.addEventListener('input', () => {
            this._query = input.value.trim().toLowerCase();
            this._focusedIdx = -1;
            this._rebuildList();

            if (!this._open) this._open_(true);
        });

        input.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                this._open_(false);
                input.blur();
                return;
            }
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                const rows = this._visibleRows();
                if (!rows.length) return;
                this._setFocus(this._focusedIdx < rows.length - 1 ? this._focusedIdx + 1 : this._focusedIdx);
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                this._setFocus(this._focusedIdx > 0 ? this._focusedIdx - 1 : 0);
            } else if (e.key === 'Enter') {
                e.preventDefault();
                const rows = this._visibleRows();
                if (this._focusedIdx >= 0 && rows[this._focusedIdx]) rows[this._focusedIdx].click();
            } else if (e.key === 'Tab') {
                this._open_(false);
            }
        });

        // Close when clicking outside; consume the click if it landed on the calendar
        // so that a day-drawer doesn't open just because a dropdown was being dismissed.
        document.addEventListener('click', (e) => {
            if (!wrap.contains(e.target) && this._open) {
                this._open_(false);
                if (e.target.closest('.jenga-calendar-wrap, .jenga-timeline-wrap')) {
                    e.stopPropagation();
                }
            }
        }, true);

        this._rebuildList();
        this._renderChips();

        return wrap;
    }

    _toggleOpen() {
        this._open_(!this._open);
    }

    _open_(state) {
        this._open = state;

        if (this._dropdown) {
            this._dropdown.style.display = state ? '' : 'none';
        }

        this._el
            ?.querySelector('.jms__trigger')
            ?.setAttribute('aria-expanded', String(state));

        if (state) {
            this._focusedIdx = -1;
            this._input?.focus();
            this._rebuildList();
        } else {
            // Clear search query when dropdown closes
            this._query = '';
            if (this._input) this._input.value = '';
            this._rebuildList();
            this._renderChips();
        }
    }

    // ── Keyboard navigation helpers ────────────────────────────────────────────

    _visibleRows() {
        return Array.from(this._dropdown.querySelectorAll('.jms__select-all, .jms__option'));
    }

    _setFocus(idx) {
        const rows = this._visibleRows();
        const clamped = rows.length ? Math.max(0, Math.min(idx, rows.length - 1)) : -1;
        rows.forEach(r => r.classList.remove('jms__option--focused'));
        if (clamped >= 0 && rows[clamped]) {
            rows[clamped].classList.add('jms__option--focused');
            rows[clamped].scrollIntoView({ block: 'nearest' });
        }
        this._focusedIdx = clamped;
    }

    // ── Dropdown list ──────────────────────────────────────────────────────────

    _rebuildList() {
        if (!this._dropdown) return;

        const prevIdx = this._focusedIdx;
        this._dropdown.innerHTML = '';

        const selectedArr = [...this.selected];

        if (selectedArr.length > 0) {

            // Container header
            const selHeader = document.createElement('div');
            selHeader.className = 'jms__selected-header';
            selHeader.innerHTML = `
        <span>${selectedArr.length} selected</span>
        <button class="jms__clear-selected">Clear all</button>
    `;

            this._dropdown.appendChild(selHeader);

            // Click clear all
            selHeader.querySelector('.jms__clear-selected')
                .addEventListener('click', (e) => {
                    e.stopPropagation();
                    this.selected.clear();
                    this._renderChips();
                    this._rebuildList();
                    this._emit();
                });

            // Container chips
            const selWrap = document.createElement('div');
            selWrap.className = 'jms__selected-wrap';

            selectedArr.forEach(val => {
                const chip = document.createElement('span');
                chip.className = 'jms__selected-chip';

                const label = document.createElement('span');
                label.className = 'jms__selected-label';
                label.textContent = val;

                const rm = document.createElement('button');
                rm.className = 'jms__selected-rm';
                rm.innerHTML = '×';

                rm.addEventListener('click', (e) => {
                    e.stopPropagation();
                    this.selected.delete(val);
                    this._renderChips();
                    this._rebuildList();
                    this._emit();
                });

                chip.append(label, rm);
                selWrap.appendChild(chip);
            });

            this._dropdown.appendChild(selWrap);

            // separator
            const sepTop = document.createElement('div');
            sepTop.className = 'jms__sep';
            this._dropdown.appendChild(sepTop);
        }

        const q = this._query;
        const visible = q
            ? this.options.filter(o => String(o).toLowerCase().includes(q))
            : this.options;

        // ── Select all / Deselect all ─────────────────────────────────────────
        const allRow = document.createElement('div');
        allRow.className = 'jms__select-all';

        const allChecked = visible.length > 0 && visible.every(o => this.selected.has(o));
        const someChecked = visible.some(o => this.selected.has(o));

        allRow.innerHTML = `
            <span class="jms__cb ${allChecked ? 'jms__cb--checked' : someChecked ? 'jms__cb--indeterminate' : ''}"></span>
            <span>Select all</span>
        `;

        allRow.addEventListener('click', () => {
            if (allChecked) {
                visible.forEach(o => this.selected.delete(o));
            } else {
                visible.forEach(o => this.selected.add(o));
            }

            this._renderChips();
            this._rebuildList();
            this._emit();
        });

        this._dropdown.appendChild(allRow);

        const sep = document.createElement('div');
        sep.className = 'jms__sep';
        this._dropdown.appendChild(sep);

        // ── Option rows ───────────────────────────────────────────────────────
        visible.forEach(opt => {
            const row = document.createElement('div');
            row.className = `jms__option${this.selected.has(opt) ? ' jms__option--selected' : ''}`;

            row.innerHTML = `
                <span class="jms__cb ${this.selected.has(opt) ? 'jms__cb--checked' : ''}"></span>
                <span class="jms__opt-label"></span>
            `;

            row.querySelector('.jms__opt-label').textContent = opt;

            row.addEventListener('click', (e) => {
                e.stopPropagation();

                if (this.selected.has(opt)) {
                    this.selected.delete(opt);
                } else {
                    this.selected.add(opt);
                }

                this._query = '';
                if (this._input) this._input.value = '';

                this._renderChips();
                this._rebuildList();
                this._emit();
            });

            this._dropdown.appendChild(row);
        });

        if (!visible.length) {
            const empty = document.createElement('div');
            empty.className = 'jms__empty';
            empty.textContent = 'No results';
            this._dropdown.appendChild(empty);
        }

        // Restore keyboard focus position after DOM rebuild (e.g. after toggling a value).
        // prevIdx is -1 when opening or searching, so this is a no-op in those cases.
        if (prevIdx >= 0) {
            const rows = this._visibleRows();
            const idx = Math.min(prevIdx, rows.length - 1);
            if (idx >= 0) {
                rows[idx].classList.add('jms__option--focused');
                this._focusedIdx = idx;
            }
        }
    }

    // ── Chips rendering ────────────────────────────────────────────────────────

    _renderChips() {
        if (!this._chipArea) return;

        this._chipArea.innerHTML = '';

        const selected = [...this.selected];
        const total = selected.length;

        if (this._input) {
            this._input.placeholder = total ? `${total} selected` : this.placeholder;
        }

        if (this._clearBtn) {
            this._clearBtn.style.display = total ? '' : 'none';
        }

        if (!total) return;

        const max = this._maxVisibleChips;

        // Show latest selected values first — better UX when user is actively selecting
        const visible = selected.slice(-max);
        const hiddenCount = total - visible.length;

        visible.forEach(val => {
            this._chipArea.appendChild(this._createChip(val));
        });

        if (hiddenCount > 0) {
            const more = document.createElement('span');
            more.className = 'jms__chip jms__chip--more';
            more.textContent = `+${hiddenCount}`;

            // Native tooltip with the complete selected list
            more.title = selected.join(', ');

            // Click on "+N" opens the dropdown and focuses search
            more.addEventListener('click', (e) => {
                e.stopPropagation();
                this._open_(true);
                this._input?.focus();
            });

            this._chipArea.appendChild(more);
        }
    }

    _createChip(val) {
        const chip = document.createElement('span');
        chip.className = 'jms__chip';
        chip.title = val;

        // Custom renderer for labels, e.g. coloured env badge
        if (this.renderChip) {
            chip.appendChild(this.renderChip(val));
        } else {
            const label = document.createElement('span');
            label.className = 'jms__chip-label';
            label.textContent = val;
            chip.appendChild(label);
        }

        const rm = document.createElement('button');
        rm.className = 'jms__chip-rm';
        rm.innerHTML = '×';
        rm.title = `Remove ${val}`;

        rm.addEventListener('click', (e) => {
            e.stopPropagation();
            this.selected.delete(val);
            this._renderChips();
            this._rebuildList();
            this._emit();
        });

        chip.appendChild(rm);

        return chip;
    }

    _emit() {
        this.onChange?.(this.getValue());
    }
}