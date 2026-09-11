/**
 * ChipBar — renders active search terms as closeable chips inside the search field.
 * Supports both single key:value and compound queries (clause1 & clause2 & …).
 *
 * Usage:
 *   const bar = new ChipBar(fieldEl, inputEl, onUpdate);
 *   bar.render(currentSearchTerm, parseActiveKV, buildKV, normalize, parseClauses);
 *
 * parseClauses(term) → [{key, values, quoted, isNegation, raw}] | null
 * parseActiveKV(term) → {key, values, quoted} | null   (single-clause fallback)
 */
export class ChipBar {
    constructor(fieldEl, inputEl, onUpdate) {
        this._field = fieldEl;
        this._input = inputEl;
        this._onUpdate = onUpdate;
        this._lastArgs = null;
        this._clauses = null;
        this._buildKV = null;
        this._editingClauseIdx = null;
        this._revealModePrefix = null; // set when input is revealed alongside chips for additive search

        // Make the field div focusable so it can receive keyboard events when chips are showing.
        fieldEl.tabIndex = 0;

        // When chips are showing and the field div has focus, forward Enter for next-occurrence
        // cycling and arrow keys for chip navigation.
        fieldEl.addEventListener('keydown', (e) => {
            if (this._input.style.display !== 'none') return; // input visible — let it handle its own keys
            if (e.key === 'Enter' && this._lastArgs?.searchTerm) {
                this._onUpdate(this._lastArgs.searchTerm);
                e.preventDefault();
            }
            if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
                if (e.key === 'ArrowRight') this._focusNextChip(document.activeElement);
                else this._focusPrevChip(document.activeElement);
                e.preventDefault();
            }
        });

        // Restore chip view when focus leaves the input without submitting.
        // Deferred so that a simultaneous click on the "go" button or a chip-close
        // can complete first (they call render() and hide the input before the
        // setTimeout fires, making the condition below false).
        inputEl.addEventListener('blur', () => {
            setTimeout(() => {
                if (this._input.style.display === 'none') return;
                // If focus moved to a chip within this field (e.g. Alt+Arrow navigation),
                // don't tear down the chip view — the user is navigating, not leaving.
                if (this._field.contains(document.activeElement)) return;
                if (this._editingClauseIdx !== null && this._clauses) {
                    // Per-clause edit: merge the edited clause back into the full query.
                    const idx = this._editingClauseIdx;
                    this._editingClauseIdx = null;
                    delete this._input.dataset.clauseEditMode;
                    const editedText = this._input.value.trim().replace(/,+$/, '');
                    const newClauses = this._clauses
                        .map((c, i) => {
                            if (i !== idx) return this._buildKV(c.key, c.values, c.quoted);
                            return editedText;
                        })
                        .filter(Boolean);
                    this._onUpdate(newClauses.join(' & '));
                } else if (this._lastArgs?.searchTerm) {
                    const { searchTerm, parseActiveKV, buildKV, normalize, parseClauses } = this._lastArgs;
                    this._input.value = searchTerm;
                    this.render(searchTerm, parseActiveKV, buildKV, normalize, parseClauses);
                }
            }, 0);
        });

        // Fix 3: When editing a specific clause and the user presses Enter, merge all
        // clauses back (keeping the other chips) instead of submitting just the edited text.
        // Use capture phase so this fires before DominoApp's keydown listener.
        // Guard: if the autocomplete dropdown has an active (highlighted) suggestion,
        // let AutocompleteEngine's handler run first to insert that suggestion.
        inputEl.addEventListener('keydown', (e) => {
            if (e.key !== 'Enter') return;
            if (this._editingClauseIdx === null || !this._clauses) return;
            const dropdown = document.getElementById('ac-dropdown');
            if (dropdown?.classList.contains('ac-open') && dropdown.querySelector('.ac-active')) return;
            const idx = this._editingClauseIdx;
            this._editingClauseIdx = null;
            delete this._input.dataset.clauseEditMode;
            const editedText = this._input.value.trim().replace(/,+$/, '');
            const newClauses = this._clauses
                .map((c, i) => i !== idx ? this._buildKV(c.key, c.values, c.quoted) : editedText)
                .filter(Boolean);
            this._onUpdate(newClauses.join(' & '));
            e.preventDefault();
            e.stopImmediatePropagation();
        }, true);

        // Fix 4a: Clicking blank space inside the field (not on a chip) reveals the input + suggestions.
        // Uses _revealInput() so chips stay in the DOM for immediate Alt+Arrow navigation.
        fieldEl.addEventListener('mousedown', (e) => {
            const CHIP_SELECTORS = '.search-chip, .search-chip-key, .search-chip-sep, .search-chip-close, .search-chip-overflow';
            if (e.target.closest(CHIP_SELECTORS)) return;
            if (this._input.style.display !== 'none') return; // input already visible
            e.preventDefault();
            this._revealInput();
        });

        // Alt+Arrow from the input (any cursor position): navigate chips.
        // Alt+ArrowLeft → focus last chip; Alt+ArrowRight → focus first chip.
        inputEl.addEventListener('keydown', (e) => {
            if (!e.altKey) return;
            if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
            const chips = this._focusableChips();
            if (!chips.length) return;
            e.preventDefault();
            if (e.key === 'ArrowLeft') {
                chips[chips.length - 1]?.focus();
            } else {
                chips[0]?.focus();
            }
        });

        // Plain ArrowLeft at cursor position 0 focuses the last chip;
        // Backspace/Delete with empty input removes the last value/clause.
        inputEl.addEventListener('keydown', (e) => {
            if (e.altKey) return; // handled above
            if (this._input.selectionStart !== 0 || this._input.selectionEnd !== 0) return;
            if (this._input.value !== '' && e.key !== 'ArrowLeft') return;

            const chips = this._focusableChips();
            if (!chips.length) return;

            if (e.key === 'ArrowLeft') {
                chips[chips.length - 1]?.focus();
                e.preventDefault();
            } else if ((e.key === 'Backspace' || e.key === 'Delete') && this._input.value === '') {
                this._removeLastChipValue();
                e.preventDefault();
            }
        });
    }

    render(searchTerm, parseActiveKV, buildKV, normalize, parseClauses) {
        this._lastArgs = { searchTerm, parseActiveKV, buildKV, normalize, parseClauses };
        this._editingClauseIdx = null;
        delete this._input.dataset.clauseEditMode;
        this._revealModePrefix = null;
        this._clearChips();
        if (!searchTerm) {
            this._input.dataset.activeFields = '';
            this._showInput(false);
            return;
        }

        const clauses = parseClauses ? parseClauses(searchTerm) : null;
        this._clauses = clauses;
        this._buildKV = buildKV;

        // Fix 5: track active field keys so AutocompleteEngine can order suggestions.
        const activeFieldKeys = clauses ? clauses.map(c => c.key).filter(Boolean) : [];
        this._input.dataset.activeFields = activeFieldKeys.join(',');

        if (clauses && clauses.length > 1) {
            // Compound query: one group of chips per clause, separated by " & "
            clauses.forEach((clause, idx) => {
                if (idx > 0) this._addSeparator(idx - 1);
                this._renderClause(clause, clauses, idx, buildKV, normalize);
            });
        } else {
            // Single clause or plain text
            const parsed = parseActiveKV(searchTerm);
            if (!parsed) {
                this._addChip(searchTerm, () => this._onUpdate(''), 0);
            } else {
                this._addKeyLabel(parsed.key, 0);
                parsed.values.forEach(v => {
                    this._addChip(v, () => {
                        const rest = parsed.values.filter(x => normalize(x) !== normalize(v));
                        this._onUpdate(rest.length ? buildKV(parsed.key, rest, parsed.quoted) : '');
                    }, 0);
                });
            }
        }

        this._input.style.display = 'none';
        this._updateOverflow();
        // Move focus to the field only if the user was already interacting with the search area,
        // so Enter for next-occurrence works without requiring an extra click.
        const active = document.activeElement;
        if (active && (active === this._input || this._field.contains(active))) {
            this._field.focus();
        }
    }

    _updateOverflow() {
        // Remove any existing overflow pill first
        this._field.querySelectorAll('.search-chip-overflow').forEach(el => el.remove());

        // Make all chips visible so we can measure them
        const chips = Array.from(this._field.querySelectorAll(
            '.search-chip, .search-chip-key, .search-chip-sep'
        ));
        chips.forEach(c => (c.style.display = ''));

        if (!chips.length) return;

        // One rAF so the browser has laid out the visible chips
        requestAnimationFrame(() => {
            const maxRight = this._field.getBoundingClientRect().right;
            let hidden = 0;
            let firstHiddenEl = null;

            for (const chip of chips) {
                const r = chip.getBoundingClientRect();
                if (r.right > maxRight) {
                    chip.style.display = 'none';
                    if (hidden === 0) firstHiddenEl = chip;
                    hidden++;
                }
            }

            if (!hidden) return;

            // Make room for the pill: keep hiding chips from the end until the
            // pill would fit, or we've hidden everything
            const pill = document.createElement('span');
            pill.className = 'search-chip-overflow';
            pill.textContent = `+${hidden}`;
            pill.addEventListener('click', (e) => { e.stopPropagation(); this._showInput(); });
            this._field.insertBefore(pill, this._input);

            // If the pill itself overflows, hide one more chip from the visible set
            requestAnimationFrame(() => {
                const pillR = pill.getBoundingClientRect();
                if (pillR.right > maxRight) {
                    // Find last visible chip and hide it
                    const visible = Array.from(this._field.querySelectorAll(
                        '.search-chip:not([style*="display: none"]), .search-chip-key:not([style*="display: none"]), .search-chip-sep:not([style*="display: none"])'
                    ));
                    if (visible.length) {
                        visible[visible.length - 1].style.display = 'none';
                        hidden++;
                        pill.textContent = `+${hidden}`;
                    }
                }
            });
        });
    }

    // ─── Internals ─────────────────────────────────────────────────────────────

    _renderClause(clause, allClauses, clauseIdx, buildKV, normalize) {
        this._addKeyLabel(clause.key, clauseIdx);
        if (clause.values.length === 0) {
            // Empty-value clause (e.g. negation with no value): show key as a single chip
            this._addChip('', () => {
                const newClauses = allClauses.filter((_, i) => i !== clauseIdx);
                this._onUpdate(this._joinClauses(newClauses, buildKV));
            }, clauseIdx);
        } else {
            clause.values.forEach(v => {
                this._addChip(v, () => {
                    const rest = clause.values.filter(x => normalize(x) !== normalize(v));
                    const newClauses = allClauses.map((c, i) => {
                        if (i !== clauseIdx) return c;
                        return { ...c, values: rest };
                    }).filter(c => c.values.length > 0);
                    this._onUpdate(this._joinClauses(newClauses, buildKV));
                }, clauseIdx);
            });
        }
    }

    _joinClauses(clauses, buildKV) {
        return clauses.map(c => buildKV(c.key, c.values, c.quoted)).filter(Boolean).join(' & ');
    }

    // Returns the fully-composed query when the input is in reveal (additive) mode.
    // If the input value starts with "& " the engine already prepended the combinator;
    // use it as-is with the prefix stripped from the suggestion.
    // Otherwise, join prefix & typed value with " & ".
    getRevealModeQuery() {
        if (!this._revealModePrefix) return null;
        const typed = this._input.value.trim();
        // Empty input in reveal mode → re-fire the existing search as-is.
        if (!typed) return this._revealModePrefix;
        // Suggestion already starts with "& " → prepend only the existing prefix
        const clausePart = typed.startsWith('& ') ? typed.slice(2).trim() : typed;
        return `${this._revealModePrefix} & ${clausePart}`;
    }

    // Clear reveal mode (called after the composed query has been committed).
    clearRevealMode() {
        this._revealModePrefix = null;
    }

    _showInput(focus = true) {
        this._clearChips();
        this._revealModePrefix = null;
        this._input.style.display = '';
        if (focus) {
            this._input.focus();
            // Trigger autocomplete suggestions immediately on reveal.
            this._input.dispatchEvent(new Event('input', { bubbles: true }));
        }
    }

    // Reveal the input alongside existing chips (no chip removal).
    // Used when clicking blank space in the bar: chips stay in the DOM so
    // Alt+Arrow navigation can find them immediately. The input is cleared
    // so the autocomplete shows fresh field-name suggestions. The existing
    // search term is saved as _revealModePrefix so that when the user commits
    // a new clause (Enter / suggestion click) it is prepended automatically.
    _revealInput() {
        this._revealModePrefix = this._lastArgs?.searchTerm || null;
        this._input.style.display = '';
        this._input.value = '';
        this._input.focus();
        this._input.dispatchEvent(new Event('input', { bubbles: true }));
    }

    _clearChips() {
        this._field.querySelectorAll('.search-chip, .search-chip-key, .search-chip-sep').forEach(el => el.remove());
        this._input.style.display = '';
    }

    _addKeyLabel(key, clauseIdx) {
        const span = document.createElement('span');
        span.className = 'search-chip-key';
        span.textContent = key;
        span.tabIndex = 0;
        span.dataset.clauseIdx = String(clauseIdx ?? 0);
        span.addEventListener('click', (e) => {
            e.stopPropagation();
            this._editClause(clauseIdx ?? 0);
        });
        span.addEventListener('keydown', (e) => {
            this._handleChipKeydown(e, span, () => {
                // Deleting the key label removes the whole clause
                if (this._clauses && clauseIdx < this._clauses.length) {
                    const newClauses = this._clauses.filter((_, i) => i !== clauseIdx);
                    this._onUpdate(this._joinClauses(newClauses, this._buildKV));
                } else {
                    this._onUpdate('');
                }
            });
        });
        this._field.insertBefore(span, this._input);
    }

    _addSeparator(afterClauseIdx) {
        const span = document.createElement('span');
        span.className = 'search-chip-sep';
        span.textContent = '&';
        span.dataset.sepAfter = String(afterClauseIdx);
        this._field.insertBefore(span, this._input);
    }

    _addChip(value, onClose, clauseIdx) {
        const chip = document.createElement('span');
        chip.className = 'search-chip';
        chip.tabIndex = 0;
        chip.dataset.clauseIdx = String(clauseIdx ?? 0);

        const text = document.createElement('span');
        text.className = 'search-chip-text';
        text.textContent = value;

        const x = document.createElement('span');
        x.className = 'search-chip-close';
        x.textContent = '×';
        x.setAttribute('aria-label', `Remove ${value}`);
        x.addEventListener('click', (e) => { e.stopPropagation(); this._closeDropdown(); onClose(); });

        // Clicking the chip body (not ×) enters per-clause edit mode.
        chip.addEventListener('click', (e) => {
            if (e.target === x) return;
            e.stopPropagation();
            this._editClause(clauseIdx ?? 0);
        });

        chip.addEventListener('keydown', (e) => {
            this._handleChipKeydown(e, chip, onClose);
        });

        chip.append(text, x);
        this._field.insertBefore(chip, this._input);
    }

    _closeDropdown() {
        document.getElementById('ac-dropdown')?.classList.remove('ac-open');
    }

    _handleChipKeydown(e, el, onClose) {
        if (e.key === 'Delete' || e.key === 'Backspace') {
            this._closeDropdown();
            onClose();
            e.preventDefault();
        } else if (e.key === 'Escape') {
            // Return focus to the input without editing.
            this._input.style.display = '';
            this._input.focus();
            e.preventDefault();
        } else if (e.key === 'ArrowRight' || e.key === 'ArrowLeft') {
            if (e.altKey) {
                // Alt+Arrow: navigate to the next/previous chip.
                if (e.key === 'ArrowRight') this._focusNextChip(el);
                else this._focusPrevChip(el);
            } else {
                // Plain Arrow: enter edit mode for this chip's clause.
                el.click();
            }
            e.preventDefault();
        } else if (e.key === 'Enter' || e.key === ' ') {
            el.click();
            e.preventDefault();
        }
    }

    _focusableChips() {
        return Array.from(this._field.querySelectorAll('.search-chip[tabindex], .search-chip-key[tabindex]'));
    }

    _focusNextChip(current) {
        const chips = this._focusableChips();
        const idx = chips.indexOf(current);
        if (idx >= 0 && idx < chips.length - 1) {
            chips[idx + 1].focus();
        } else if (idx === chips.length - 1) {
            // Wrap from last chip to the input
            this._input.style.display = '';
            this._input.focus();
        }
    }

    _focusPrevChip(current) {
        const chips = this._focusableChips();
        const idx = chips.indexOf(current);
        if (idx > 0) {
            chips[idx - 1].focus();
        } else if (idx === 0) {
            // At the first chip: move focus back to the input.
            this._input.style.display = '';
            this._input.focus();
        }
    }

    // Fix 4 helper: remove the last value/clause from the query.
    _removeLastChipValue() {
        if (!this._lastArgs) return;
        const { searchTerm, parseClauses, buildKV, normalize } = this._lastArgs;
        if (!searchTerm) return;
        this._closeDropdown();

        if (parseClauses && this._clauses && this._clauses.length > 0) {
            const lastClause = this._clauses[this._clauses.length - 1];
            if (lastClause.values.length > 1) {
                // Remove last value of last clause
                const newValues = lastClause.values.slice(0, -1);
                const newClauses = this._clauses.map((c, i) =>
                    i !== this._clauses.length - 1 ? c : { ...c, values: newValues }
                );
                this._onUpdate(this._joinClauses(newClauses, buildKV));
            } else {
                // Remove the entire last clause
                const newClauses = this._clauses.slice(0, -1);
                this._onUpdate(this._joinClauses(newClauses, buildKV));
            }
        } else {
            // Single clause or plain text: clear entirely
            this._onUpdate('');
        }
    }

    _editClause(clauseIdx) {
        const clauses = this._clauses;
        if (!clauses || clauses.length <= 1) {
            // Single-clause or plain text: full edit mode
            this._editingClauseIdx = null;
            this._showInput();
            if (this._lastArgs?.searchTerm) {
                this._input.value = this._lastArgs.searchTerm;
            }
            return;
        }

        // Compound query: edit only this clause inline, leaving the rest as chips.
        const clause = clauses[clauseIdx];
        if (!clause) return;

        const clauseRaw = this._buildKV(clause.key, clause.values, clause.quoted) ?? '';

        // Remove elements belonging to this clause and their adjacent separator.
        this._field.querySelectorAll(`[data-clause-idx="${clauseIdx}"]`).forEach(el => el.remove());
        // Remove the separator between this clause and its neighbour (either side).
        this._field.querySelectorAll(`.search-chip-sep[data-sep-after="${clauseIdx}"]`).forEach(el => el.remove());
        this._field.querySelectorAll(`.search-chip-sep[data-sep-after="${clauseIdx - 1}"]`).forEach(el => el.remove());

        // Fix 5: update active fields to reflect remaining (non-edited) clauses only.
        const remainingKeys = clauses.filter((_, i) => i !== clauseIdx).map(c => c.key).filter(Boolean);
        this._input.dataset.activeFields = remainingKeys.join(',');

        this._editingClauseIdx = clauseIdx;
        this._input.dataset.clauseEditMode = '1';
        this._input.value = clauseRaw;
        this._input.style.display = '';
        this._input.focus();
        this._input.dispatchEvent(new Event('input', { bubbles: true }));
    }
}
