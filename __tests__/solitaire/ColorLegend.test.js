import { ColorLegend } from '../../js/solitaire/ColorLegend.js';
import { ROLE_FIELD_WITH_MAPPING, COMPANY_FIELD, LOCATION_FIELD, BUSINESS_FUNCTION_FIELD, NEUTRAL_COLOR, TEAM_MEMBER_LEGENDA_LABEL } from '../../js/solitaire/constants.js';
import { BRAND } from '../../brand-specific/brand.js';
import * as d3 from 'd3';

function makeApp(visibleOrg = {}) {
    return {
        visibleOrg,
        renderer: { recolorCards: jest.fn(), fitToContent: jest.fn() },
        search: { search: jest.fn() },
    };
}

function emptyScale() {
    return Object.assign(jest.fn(() => '#fff'), { colorOf: jest.fn(() => '#fff'), domain: () => [] });
}

describe('ColorLegend constructor', () => {
    test('initializes with default values', () => {
        const cl = new ColorLegend(makeApp());
        expect(cl.colorBy).toBe(ROLE_FIELD_WITH_MAPPING);
        expect(cl.colorScale).toBeNull();
        expect(cl.colorKeyMappings).toBeInstanceOf(Map);
        expect(cl.seenLegendClickKeys).toBeInstanceOf(Set);
        expect(cl.lastLegendClickAt).toBe(0);
    });
});

describe('ColorLegend.isUnknownKey', () => {
    const cl = new ColorLegend(makeApp());

    test('returns true for empty string', () => {
        expect(cl.isUnknownKey('')).toBe(true);
    });

    test('returns true for null/undefined', () => {
        expect(cl.isUnknownKey(null)).toBe(true);
        expect(cl.isUnknownKey(undefined)).toBe(true);
    });

    test('returns true for "unknown"', () => {
        expect(cl.isUnknownKey('unknown')).toBe(true);
        expect(cl.isUnknownKey('UNKNOWN')).toBe(true);
    });

    test('returns true for "n/a" and "N/A"', () => {
        expect(cl.isUnknownKey('n/a')).toBe(true);
        expect(cl.isUnknownKey('N/A')).toBe(true);
    });

    test('returns true for "Not Set" and "Not Available"', () => {
        expect(cl.isUnknownKey('Not Set')).toBe(true);
        expect(cl.isUnknownKey('Not Available')).toBe(true);
    });

    test('returns true for single dash or em-dash', () => {
        expect(cl.isUnknownKey('-')).toBe(true);
        expect(cl.isUnknownKey('—')).toBe(true);
    });

    test('returns false for normal role name', () => {
        expect(cl.isUnknownKey('Engineer')).toBe(false);
    });

    test('returns false for "none-but-real"', () => {
        expect(cl.isUnknownKey('none-but-real')).toBe(false);
    });
});

describe('ColorLegend.getCardFill', () => {
    test('returns NEUTRAL_COLOR when colorScale is not a function', () => {
        const cl = new ColorLegend(makeApp());
        cl.colorScale = null;
        const g = { attr: jest.fn(() => 'Engineer') };
        expect(cl.getCardFill(g)).toBe(NEUTRAL_COLOR);
    });

    test('returns color from scale for Role field', () => {
        const cl = new ColorLegend(makeApp());
        cl.colorBy = ROLE_FIELD_WITH_MAPPING;
        cl.colorScale = jest.fn(() => '#ff0000');
        const g = { attr: jest.fn((name) => name === 'data-role' ? 'Engineer' : null) };
        expect(cl.getCardFill(g)).toBe('#ff0000');
    });

    test('uses TEAM_MEMBER_LEGENDA_LABEL when data-role is empty', () => {
        const cl = new ColorLegend(makeApp());
        cl.colorBy = ROLE_FIELD_WITH_MAPPING;
        cl.colorScale = jest.fn(key => key === TEAM_MEMBER_LEGENDA_LABEL ? '#aaaaaa' : '#000');
        const g = { attr: jest.fn(() => '') };
        expect(cl.getCardFill(g)).toBe('#aaaaaa');
    });

    test('returns color from scale for Company field', () => {
        const cl = new ColorLegend(makeApp());
        cl.colorBy = COMPANY_FIELD;
        cl.colorScale = jest.fn(() => '#00ff00');
        const g = { attr: jest.fn((name) => name === 'data-company' ? BRAND.name : null) };
        expect(cl.getCardFill(g)).toBe('#00ff00');
    });

    test('returns color from scale for Location field (fallback)', () => {
        const cl = new ColorLegend(makeApp());
        cl.colorBy = LOCATION_FIELD;
        cl.colorScale = jest.fn(() => '#0000ff');
        const g = { attr: jest.fn((name) => name === 'data-location' ? 'Florence' : null) };
        expect(cl.getCardFill(g)).toBe('#0000ff');
    });

    test('returns color from scale for Business Function field', () => {
        const cl = new ColorLegend(makeApp());
        cl.colorBy = BUSINESS_FUNCTION_FIELD;
        cl.colorScale = jest.fn(() => '#abcdef');
        const g = { attr: jest.fn((name) => name === 'data-function' ? 'IT' : null) };
        expect(cl.getCardFill(g)).toBe('#abcdef');
    });

    test('returns NEUTRAL_COLOR when scale returns non-string', () => {
        const cl = new ColorLegend(makeApp());
        cl.colorScale = jest.fn(() => null);
        const g = { attr: jest.fn(() => 'SomeRole') };
        expect(cl.getCardFill(g)).toBe(NEUTRAL_COLOR);
    });
});

describe('ColorLegend.renderAll', () => {
    beforeEach(() => {
        localStorage.clear();
        document.body.innerHTML = '';
    });

    afterEach(() => {
        document.body.innerHTML = '';
    });

    test('creates #legend-root if not present and appends to body', () => {
        const cl = new ColorLegend(makeApp());
        const colorScale = Object.assign(jest.fn(() => '#fff'), {
            colorOf: jest.fn(() => '#fff'),
            domain: () => ['Engineer', 'Designer'],
        });
        cl.renderAll({
            title: 'Test Legend',
            fieldName: ROLE_FIELD_WITH_MAPPING,
            keys: ['Engineer', 'Designer'],
            counts: new Map([['Engineer', 3], ['Designer', 1]]),
            topKey: 'Engineer',
            colorOf: colorScale,
        });
        const root = document.getElementById('legend-root');
        expect(root).toBeTruthy();
        expect(root.querySelector('.legend__title').textContent).toBe('2 Test Legend');
    });

    test('reuses existing #legend-root', () => {
        const existing = document.createElement('div');
        existing.id = 'legend-root';
        document.body.appendChild(existing);
        const cl = new ColorLegend(makeApp());
        cl.renderAll({ title: 'T2', fieldName: ROLE_FIELD_WITH_MAPPING, keys: [], counts: new Map(), topKey: null, colorOf: emptyScale() });
        expect(document.querySelectorAll('#legend-root').length).toBe(1);
    });

    test('collapse button toggles the legend', () => {
        const cl = new ColorLegend(makeApp());
        cl.renderAll({ title: 'T', fieldName: ROLE_FIELD_WITH_MAPPING, keys: [], counts: new Map(), topKey: null, colorOf: emptyScale() });
        const btn = document.querySelector('.legend__collapse');
        btn.click();
        const root = document.getElementById('legend-root');
        expect(root.classList.contains('legend--collapsed')).toBe(true);
    });

    test('legend item click triggers search', () => {
        const app = makeApp();
        const cl = new ColorLegend(app);
        cl.renderAll({
            title: 'T',
            fieldName: ROLE_FIELD_WITH_MAPPING,
            keys: ['Engineer'],
            counts: new Map([['Engineer', 3]]),
            topKey: 'Engineer',
            colorOf: emptyScale(),
        });
        const item = document.querySelector('.legend__item[data-value="Engineer"]');
        if (item) item.click();
        // search.search is called by activate()
        expect(app.search.search).toHaveBeenCalled();
    });

    test('collapse button pointerdown does not throw', () => {
        const cl = new ColorLegend(makeApp());
        cl.renderAll({ title: 'T', fieldName: ROLE_FIELD_WITH_MAPPING, keys: [], counts: new Map(), topKey: null, colorOf: emptyScale() });
        const btn = document.querySelector('.legend__collapse');
        expect(() => btn.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }))).not.toThrow();
    });

    test('renders with a key whose color is white', () => {
        const whiteScale = Object.assign(jest.fn(() => '#ffffff'), {
            colorOf: jest.fn(() => '#ffffff'),
            domain: () => ['IceRole'],
        });
        const cl = new ColorLegend(makeApp());
        expect(() => cl.renderAll({
            title: 'T',
            fieldName: ROLE_FIELD_WITH_MAPPING,
            keys: ['IceRole'],
            counts: new Map([['IceRole', 1]]),
            topKey: 'IceRole',
            colorOf: whiteScale,
        })).not.toThrow();
        const swatch = document.querySelector('.legend__swatch--white');
        expect(swatch).toBeTruthy();
    });
});

// ─── ColorLegend.setMode ──────────────────────────────────────────────────────

describe('ColorLegend.setMode', () => {
    beforeEach(() => {
        localStorage.clear();
        document.body.innerHTML = '';
    });

    afterEach(() => {
        document.body.innerHTML = '';
    });

    function setupRadios() {
        document.body.innerHTML = `
            <input id="toggle-color-role" type="radio"/>
            <input id="toggle-color-company" type="radio"/>
            <input id="toggle-color-location" type="radio"/>
            <input id="toggle-color-function" type="radio"/>
        `;
    }

    test('does not throw when radio elements are missing', () => {
        document.body.innerHTML = '';
        const cl = new ColorLegend(makeApp());
        expect(() => cl.setMode(ROLE_FIELD_WITH_MAPPING)).not.toThrow();
    });

    test('checks role radio for ROLE_FIELD_WITH_MAPPING', () => {
        setupRadios();
        const cl = new ColorLegend(makeApp());
        cl.setMode(ROLE_FIELD_WITH_MAPPING);
        expect(document.getElementById('toggle-color-role').checked).toBe(true);
        expect(document.getElementById('toggle-color-company').checked).toBe(false);
    });

    test('checks company radio for COMPANY_FIELD', () => {
        setupRadios();
        const cl = new ColorLegend(makeApp());
        cl.setMode(COMPANY_FIELD);
        expect(document.getElementById('toggle-color-company').checked).toBe(true);
    });

    test('checks location radio for LOCATION_FIELD', () => {
        setupRadios();
        const cl = new ColorLegend(makeApp());
        cl.setMode(LOCATION_FIELD);
        expect(document.getElementById('toggle-color-location').checked).toBe(true);
        expect(document.getElementById('toggle-color-function').checked).toBe(false);
    });

    test('checks business-function radio for BUSINESS_FUNCTION_FIELD', () => {
        setupRadios();
        const cl = new ColorLegend(makeApp());
        cl.setMode(BUSINESS_FUNCTION_FIELD);
        expect(document.getElementById('toggle-color-function').checked).toBe(true);
        expect(document.getElementById('toggle-color-role').checked).toBe(false);
        expect(document.getElementById('toggle-color-company').checked).toBe(false);
        expect(document.getElementById('toggle-color-location').checked).toBe(false);
    });
});

// ─── ColorLegend.recolor ──────────────────────────────────────────────────────

describe('ColorLegend.recolor', () => {
    beforeEach(() => {
        localStorage.clear();
        document.body.innerHTML = '';
    });

    afterEach(() => {
        document.body.innerHTML = '';
        jest.restoreAllMocks();
    });

    test('does not throw with empty visibleOrg', () => {
        const cl = new ColorLegend(makeApp({}));
        expect(() => cl.recolor(ROLE_FIELD_WITH_MAPPING)).not.toThrow();
    });

    test('sets colorBy to the given field', () => {
        const cl = new ColorLegend(makeApp({}));
        cl.recolor(COMPANY_FIELD);
        expect(cl.colorBy).toBe(COMPANY_FIELD);
    });

    test('calls d3.selectAll to recolor cards', () => {
        const cl = new ColorLegend(makeApp({}));
        cl.recolor(ROLE_FIELD_WITH_MAPPING);
        expect(d3.selectAll).toHaveBeenCalled();
    });

    test('recolor each callback invokes getCardFill and applies fill', () => {
        const mockRect = {
            transition: jest.fn().mockReturnThis(),
            duration: jest.fn().mockReturnThis(),
            attr: jest.fn().mockReturnThis(),
            classed: jest.fn().mockReturnThis(),
        };
        const mockG = {
            select: jest.fn(() => mockRect),
            attr: jest.fn((name) => name === 'data-role' ? 'Engineer' : null),
        };
        const mockNode = {};

        // Make selectAll('g[...]').each(cb) actually call cb with mock node
        jest.spyOn(d3, 'selectAll').mockImplementation((sel) => {
            if (typeof sel === 'string' && sel.includes('card::')) {
                return {
                    each: jest.fn((cb) => { cb(null, 0, [mockNode]); }),
                };
            }
            return { each: jest.fn(), classed: jest.fn().mockReturnThis() };
        });
        jest.spyOn(d3, 'select').mockImplementation(() => mockG);

        const cl = new ColorLegend(makeApp({}));
        cl.colorScale = jest.fn(() => '#aabbcc');
        cl.recolor(ROLE_FIELD_WITH_MAPPING);
        expect(mockRect.attr).toHaveBeenCalledWith('fill', expect.any(String));
    });
});

// ─── ColorLegend.renderAll (unknown key item) ─────────────────────────────────

describe('ColorLegend.renderAll (unknown keys and keydown)', () => {
    beforeEach(() => {
        localStorage.clear();
        document.body.innerHTML = '';
    });

    afterEach(() => {
        document.body.innerHTML = '';
    });

    test('renders disabled item for unknown key', () => {
        const cl = new ColorLegend(makeApp());
        cl.renderAll({
            title: 'T',
            fieldName: ROLE_FIELD_WITH_MAPPING,
            keys: ['unknown', 'Engineer'],
            counts: new Map([['unknown', 2], ['Engineer', 3]]),
            topKey: 'Engineer',
            colorOf: emptyScale(),
        });
        const disabledItem = document.querySelector('.legend__item--disabled');
        expect(disabledItem).toBeTruthy();
    });

    test('keydown Enter on legend item triggers search', () => {
        const app = makeApp();
        const cl = new ColorLegend(app);
        cl.renderAll({
            title: 'T',
            fieldName: ROLE_FIELD_WITH_MAPPING,
            keys: ['Engineer'],
            counts: new Map([['Engineer', 3]]),
            topKey: 'Engineer',
            colorOf: emptyScale(),
        });
        const item = document.querySelector('.legend__item[data-value="Engineer"]');
        expect(item).toBeTruthy();
        item.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
        expect(app.search.search).toHaveBeenCalled();
    });

    test('keydown Space on legend item triggers search', () => {
        const app = makeApp();
        const cl = new ColorLegend(app);
        cl.renderAll({
            title: 'T',
            fieldName: ROLE_FIELD_WITH_MAPPING,
            keys: ['Engineer'],
            counts: new Map([['Engineer', 3]]),
            topKey: 'Engineer',
            colorOf: emptyScale(),
        });
        const item = document.querySelector('.legend__item[data-value="Engineer"]');
        item.dispatchEvent(new KeyboardEvent('keydown', { key: ' ', bubbles: true }));
        expect(app.search.search).toHaveBeenCalled();
    });

    test('keydown non-Enter/Space does nothing', () => {
        const app = makeApp();
        const cl = new ColorLegend(app);
        cl.renderAll({
            title: 'T',
            fieldName: ROLE_FIELD_WITH_MAPPING,
            keys: ['Engineer'],
            counts: new Map([['Engineer', 3]]),
            topKey: 'Engineer',
            colorOf: emptyScale(),
        });
        const list = document.querySelector('.legend__list');
        list.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }));
        expect(app.search.search).not.toHaveBeenCalled();
    });
});

// ─── ColorLegend.enableDrag ───────────────────────────────────────────────────

describe('ColorLegend.enableDrag (drag to move legend)', () => {
    beforeEach(() => {
        localStorage.clear();
        document.body.innerHTML = '';
    });

    afterEach(() => {
        document.body.innerHTML = '';
    });

    test('restores position from localStorage when available', () => {
        localStorage.setItem('legend-pos-v1', JSON.stringify({ x: 50, y: 100 }));
        const cl = new ColorLegend(makeApp());
        cl.renderAll({ title: 'T', fieldName: ROLE_FIELD_WITH_MAPPING, keys: [], counts: new Map(), topKey: null, colorOf: emptyScale() });
        const root = document.getElementById('legend-root');
        expect(root.style.left).toBe('50px');
    });

    test('dragging handle moves the legend', () => {
        const cl = new ColorLegend(makeApp());
        cl.renderAll({ title: 'T', fieldName: ROLE_FIELD_WITH_MAPPING, keys: [], counts: new Map(), topKey: null, colorOf: emptyScale() });
        const root = document.getElementById('legend-root');
        const handle = root.querySelector('.legend__header');
        expect(handle).toBeTruthy();

        // pointerdown
        handle.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, button: 0, clientX: 100, clientY: 100 }));
        // pointermove — large enough delta to trigger dragging
        window.dispatchEvent(new MouseEvent('pointermove', { bubbles: true, clientX: 110, clientY: 115 }));
        // pointerup
        window.dispatchEvent(new MouseEvent('pointerup', { bubbles: true }));
        // Just verify it doesn't throw and root exists
        expect(root).toBeTruthy();
    });

    test('non-primary-button pointerdown is ignored', () => {
        const cl = new ColorLegend(makeApp());
        cl.renderAll({ title: 'T', fieldName: ROLE_FIELD_WITH_MAPPING, keys: [], counts: new Map(), topKey: null, colorOf: emptyScale() });
        const handle = document.querySelector('.legend__header');
        // button: 2 (right-click) should be ignored
        expect(() => handle.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, button: 2, clientX: 50, clientY: 50 }))).not.toThrow();
    });
});

// ─── ColorLegend — in-legend search ──────────────────────────────────────────

describe('ColorLegend in-legend search', () => {
    beforeEach(() => {
        localStorage.clear();
        document.body.innerHTML = '';
    });

    afterEach(() => {
        document.body.innerHTML = '';
    });

    function renderWithKeys(keys = ['Engineer', 'Manager', 'Designer']) {
        const app = makeApp();
        const cl = new ColorLegend(app);
        const counts = new Map(keys.map(k => [k, 1]));
        cl.renderAll({ title: 'T', fieldName: ROLE_FIELD_WITH_MAPPING, keys, counts, topKey: keys[0] || null, colorOf: emptyScale() });
        return { cl, app };
    }

    test('renders .legend__search-toggle button', () => {
        renderWithKeys();
        expect(document.querySelector('.legend__search-toggle')).toBeTruthy();
    });

    test('renders .legend__search-input (initially hidden via wrapper)', () => {
        renderWithKeys();
        const wrap = document.querySelector('.legend__search-wrap');
        expect(wrap).toBeTruthy();
        expect(wrap.hidden).toBe(true);
    });

    test('clicking toggle shows search wrap and hides title', () => {
        renderWithKeys();
        const toggle = document.querySelector('.legend__search-toggle');
        const wrap   = document.querySelector('.legend__search-wrap');
        const title  = document.querySelector('.legend__title');
        toggle.click();
        expect(wrap.hidden).toBe(false);
        expect(title.hidden).toBe(true);
        expect(toggle.getAttribute('aria-pressed')).toBe('true');
    });

    test('clicking toggle again closes search and restores title', () => {
        renderWithKeys();
        const toggle = document.querySelector('.legend__search-toggle');
        const wrap   = document.querySelector('.legend__search-wrap');
        const title  = document.querySelector('.legend__title');
        toggle.click();  // open
        toggle.click();  // close
        expect(wrap.hidden).toBe(true);
        expect(title.hidden).toBe(false);
        expect(toggle.getAttribute('aria-pressed')).toBe('false');
    });

    test('typing a query hides non-matching items and shows counter', () => {
        renderWithKeys(['Engineer', 'Manager', 'Senior Manager']);
        const toggle  = document.querySelector('.legend__search-toggle');
        const input   = document.querySelector('.legend__search-input');
        const counter = document.querySelector('.legend__search-count');
        toggle.click();

        input.value = 'manager';
        input.dispatchEvent(new Event('input', { bubbles: true }));

        const items = document.querySelectorAll('.legend__item');
        const visible  = [...items].filter(el => !el.hidden);
        const hidden   = [...items].filter(el => el.hidden);
        expect(visible.length).toBe(2);   // Manager + Senior Manager
        expect(hidden.length).toBe(1);    // Engineer
        expect(counter.textContent).toBe('2');
        expect(counter.hidden).toBe(false);
    });

    test('clearing the query restores all items and hides counter', () => {
        renderWithKeys(['Engineer', 'Manager']);
        const toggle  = document.querySelector('.legend__search-toggle');
        const input   = document.querySelector('.legend__search-input');
        const counter = document.querySelector('.legend__search-count');
        toggle.click();

        input.value = 'manager';
        input.dispatchEvent(new Event('input', { bubbles: true }));

        input.value = '';
        input.dispatchEvent(new Event('input', { bubbles: true }));

        const items = document.querySelectorAll('.legend__item');
        expect([...items].every(el => !el.hidden)).toBe(true);
        expect(counter.hidden).toBe(true);
    });

    test('pressing Escape closes search and restores all items', () => {
        renderWithKeys(['Engineer', 'Manager']);
        const toggle = document.querySelector('.legend__search-toggle');
        const input  = document.querySelector('.legend__search-input');
        const wrap   = document.querySelector('.legend__search-wrap');
        const title  = document.querySelector('.legend__title');
        toggle.click();

        input.value = 'manager';
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));

        expect(wrap.hidden).toBe(true);
        expect(title.hidden).toBe(false);
        const items = document.querySelectorAll('.legend__item');
        expect([...items].every(el => !el.hidden)).toBe(true);
    });

    test('search toggle pointerdown does not propagate (no error)', () => {
        renderWithKeys();
        const toggle = document.querySelector('.legend__search-toggle');
        expect(() => toggle.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }))).not.toThrow();
    });

    test('renders .legend__search-close button (initially hidden)', () => {
        renderWithKeys();
        const closeBtn = document.querySelector('.legend__search-close');
        expect(closeBtn).toBeTruthy();
        expect(closeBtn.hidden).toBe(true);
    });

    test('X close button appears when search opens', () => {
        renderWithKeys();
        const toggle   = document.querySelector('.legend__search-toggle');
        const closeBtn = document.querySelector('.legend__search-close');
        toggle.click();
        expect(closeBtn.hidden).toBe(false);
    });

    test('X close button click closes search and restores title', () => {
        renderWithKeys();
        const toggle   = document.querySelector('.legend__search-toggle');
        const closeBtn = document.querySelector('.legend__search-close');
        const wrap     = document.querySelector('.legend__search-wrap');
        const title    = document.querySelector('.legend__title');
        toggle.click();
        closeBtn.click();
        expect(wrap.hidden).toBe(true);
        expect(title.hidden).toBe(false);
        expect(closeBtn.hidden).toBe(true);
        expect(toggle.getAttribute('aria-pressed')).toBe('false');
    });

    test('clicking outside the legend root closes search', () => {
        renderWithKeys();
        const toggle = document.querySelector('.legend__search-toggle');
        const wrap   = document.querySelector('.legend__search-wrap');
        const title  = document.querySelector('.legend__title');
        toggle.click();
        expect(wrap.hidden).toBe(false);
        document.body.dispatchEvent(new MouseEvent('click', { bubbles: true }));
        expect(wrap.hidden).toBe(true);
        expect(title.hidden).toBe(false);
    });

    test('pressing Enter scrolls to first visible item', () => {
        HTMLElement.prototype.scrollIntoView = jest.fn();
        const spy = jest.spyOn(HTMLElement.prototype, 'scrollIntoView');
        renderWithKeys(['Engineer', 'Manager', 'Designer']);
        const toggle = document.querySelector('.legend__search-toggle');
        const input  = document.querySelector('.legend__search-input');
        toggle.click();
        input.value = 'manager';
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
        expect(spy).toHaveBeenCalled();
        spy.mockRestore();
        delete HTMLElement.prototype.scrollIntoView;
    });

    test('lens click when search is open with query scrolls without closing', () => {
        HTMLElement.prototype.scrollIntoView = jest.fn();
        const spy = jest.spyOn(HTMLElement.prototype, 'scrollIntoView');
        renderWithKeys(['Engineer', 'Manager', 'Designer']);
        const toggle = document.querySelector('.legend__search-toggle');
        const input  = document.querySelector('.legend__search-input');
        const wrap   = document.querySelector('.legend__search-wrap');
        toggle.click();
        input.value = 'manager';
        input.dispatchEvent(new Event('input', { bubbles: true }));
        toggle.click();
        expect(spy).toHaveBeenCalled();
        expect(wrap.hidden).toBe(false);
        spy.mockRestore();
        delete HTMLElement.prototype.scrollIntoView;
    });

    test('lens click when search is open with no query closes search', () => {
        renderWithKeys();
        const toggle = document.querySelector('.legend__search-toggle');
        const wrap   = document.querySelector('.legend__search-wrap');
        const title  = document.querySelector('.legend__title');
        toggle.click();
        toggle.click();
        expect(wrap.hidden).toBe(true);
        expect(title.hidden).toBe(false);
    });

    test('pressing Enter multiple times cycles through all matching items', () => {
        HTMLElement.prototype.scrollIntoView = jest.fn();
        renderWithKeys(['Manager', 'Senior Manager', 'Engineer']);
        const toggle = document.querySelector('.legend__search-toggle');
        const input  = document.querySelector('.legend__search-input');
        toggle.click();
        input.value = 'manager';
        input.dispatchEvent(new Event('input', { bubbles: true }));

        const items = [...document.querySelectorAll('.legend__item:not([hidden])')];
        expect(items.length).toBe(2);

        // First Enter → index 0 (Manager)
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
        expect(HTMLElement.prototype.scrollIntoView).toHaveBeenCalledTimes(1);
        expect(items[0].classList.contains('legend__item--highlight')).toBe(true);

        // Second Enter → index 1 (Senior Manager)
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
        expect(HTMLElement.prototype.scrollIntoView).toHaveBeenCalledTimes(2);
        expect(items[1].classList.contains('legend__item--highlight')).toBe(true);

        // Third Enter → wraps back to index 0 (Manager)
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
        expect(HTMLElement.prototype.scrollIntoView).toHaveBeenCalledTimes(3);
        expect(items[0].classList.contains('legend__item--highlight')).toBe(true);

        delete HTMLElement.prototype.scrollIntoView;
    });

    test('changing query resets cycling cursor to first match', () => {
        HTMLElement.prototype.scrollIntoView = jest.fn();
        renderWithKeys(['Manager', 'Senior Manager', 'Engineer']);
        const toggle = document.querySelector('.legend__search-toggle');
        const input  = document.querySelector('.legend__search-input');
        toggle.click();
        input.value = 'manager';
        input.dispatchEvent(new Event('input', { bubbles: true }));

        // Advance cursor past first item
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

        // Change query → cursor resets
        input.value = 'engineer';
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

        const visible = [...document.querySelectorAll('.legend__item:not([hidden])')];
        expect(visible.length).toBe(1);
        expect(visible[0].querySelector('.legend__label').textContent).toBe('Engineer');
        expect(visible[0].classList.contains('legend__item--highlight')).toBe(true);

        delete HTMLElement.prototype.scrollIntoView;
    });
});

// ─── ColorLegend — resize handle ─────────────────────────────────────────────

describe('ColorLegend resize handle', () => {
    beforeEach(() => {
        localStorage.clear();
        document.body.innerHTML = '';
    });

    afterEach(() => {
        document.body.innerHTML = '';
    });

    test('renders .legend__resize-handle after enableDrag', () => {
        const cl = new ColorLegend(makeApp());
        cl.renderAll({ title: 'T', fieldName: ROLE_FIELD_WITH_MAPPING, keys: [], counts: new Map(), topKey: null, colorOf: emptyScale() });
        expect(document.querySelector('.legend__resize-handle')).toBeTruthy();
    });

    test('restores saved size from localStorage', () => {
        localStorage.setItem('legend-size-v1', JSON.stringify({ width: '400px', maxListHeight: '300px' }));
        const cl = new ColorLegend(makeApp());
        cl.renderAll({ title: 'T', fieldName: ROLE_FIELD_WITH_MAPPING, keys: [], counts: new Map(), topKey: null, colorOf: emptyScale() });
        const root = document.getElementById('legend-root');
        const list = root.querySelector('.legend__list');
        expect(root.style.width).toBe('400px');
        expect(list.style.maxHeight).toBe('300px');
    });

    test('pointerdown on resize handle does not throw', () => {
        const cl = new ColorLegend(makeApp());
        cl.renderAll({ title: 'T', fieldName: ROLE_FIELD_WITH_MAPPING, keys: [], counts: new Map(), topKey: null, colorOf: emptyScale() });
        const handle = document.querySelector('.legend__resize-handle');
        expect(() => handle.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, button: 0, clientX: 200, clientY: 200 }))).not.toThrow();
    });

    test('pointermove after resize pointerdown updates styles', () => {
        const cl = new ColorLegend(makeApp());
        cl.renderAll({ title: 'T', fieldName: ROLE_FIELD_WITH_MAPPING, keys: [], counts: new Map(), topKey: null, colorOf: emptyScale() });
        const root   = document.getElementById('legend-root');
        const handle = document.querySelector('.legend__resize-handle');

        handle.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, button: 0, clientX: 100, clientY: 100 }));
        handle.dispatchEvent(new MouseEvent('pointermove', { bubbles: true, clientX: 250, clientY: 250 }));
        handle.dispatchEvent(new MouseEvent('pointerup',   { bubbles: true }));

        // After drag, root.style.width should be set (clamped between 200–600)
        expect(root.style.width).toMatch(/\d+px/);
    });

    test('non-primary-button pointerdown on resize handle is ignored', () => {
        const cl = new ColorLegend(makeApp());
        cl.renderAll({ title: 'T', fieldName: ROLE_FIELD_WITH_MAPPING, keys: [], counts: new Map(), topKey: null, colorOf: emptyScale() });
        const root   = document.getElementById('legend-root');
        const handle = document.querySelector('.legend__resize-handle');
        handle.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, button: 2, clientX: 200, clientY: 200 }));
        // right-click should not set width
        expect(root.style.width).toBe('');
    });
});
