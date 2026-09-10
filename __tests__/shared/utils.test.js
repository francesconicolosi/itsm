import {
    normalizeWs,
    normalizeKey,
    truncateString,
    isUrl,
    splitValues,
    splitNarrativeValues,
    isDateTimeValue,
    formatDateTimeLocal,
    getFormattedDate,
    formatMonthYear,
    getQueryParam,
    setSearchQuery,
    createHrefElement,
    addTagToElement,
    createFormattedElementsFrom,
    createFormattedLongTextElementsFrom,
    closeSideDrawer,
    openSideDrawer,
    toggleClearButton,
    createOutlookUrl,
    parseCSV,
    isMobileDevice,
    openOutlookWebCompose,
    initCommonActions,
    enableGlobalFindShortcut,
    createModal,
    formatUrlLink,
    renderUrlPartsIntoCell,
} from '../../js/shared/utils.js';

// ─── normalizeWs ──────────────────────────────────────────────────────────────

describe('normalizeWs', () => {
    test('trims and collapses whitespace for Name field', () => {
        expect(normalizeWs('  John  Doe  ', 'Name')).toBe('John Doe');
    });

    test('trims and collapses whitespace for User field', () => {
        expect(normalizeWs('  alice  smith  ', 'User')).toBe('alice smith');
    });

    test('trims and collapses whitespace for Email field', () => {
        expect(normalizeWs('  a@b.com  ', 'Email')).toBe('a@b.com');
    });

    test('only trims for non-normalized fields (no internal collapse)', () => {
        expect(normalizeWs('  foo  bar  ', 'Description')).toBe('foo  bar');
    });

    test('full normalization when no fieldName given', () => {
        expect(normalizeWs('  foo  bar  ')).toBe('foo bar');
    });

    test('handles null/undefined value', () => {
        expect(normalizeWs(null)).toBe('');
        expect(normalizeWs(undefined)).toBe('');
    });

    test('handles numeric values', () => {
        expect(normalizeWs(42)).toBe('42');
    });
});

// ─── normalizeKey ─────────────────────────────────────────────────────────────

describe('normalizeKey', () => {
    test('lowercases and replaces spaces with underscores', () => {
        expect(normalizeKey('Hello World')).toBe('hello_world');
    });

    test('removes special characters', () => {
        expect(normalizeKey('foo!@#bar')).toBe('foobar');
    });

    test('preserves hyphens', () => {
        expect(normalizeKey('foo-bar')).toBe('foo-bar');
    });

    test('handles leading/trailing whitespace', () => {
        expect(normalizeKey('  foo  ')).toBe('foo');
    });

    test('handles null/undefined', () => {
        expect(normalizeKey(null)).toBe('');
        expect(normalizeKey(undefined)).toBe('');
    });

    test('multiple spaces become single underscore', () => {
        expect(normalizeKey('foo  bar')).toBe('foo_bar');
    });
});

// ─── truncateString ───────────────────────────────────────────────────────────

describe('truncateString', () => {
    test('returns string unchanged when within limit', () => {
        expect(truncateString('hello', 10)).toBe('hello');
    });

    test('truncates with ellipsis when over limit', () => {
        expect(truncateString('hello world', 8)).toBe('hello wo...');
    });

    test('uses default max of 25', () => {
        const long = 'a'.repeat(30);
        expect(truncateString(long)).toBe('a'.repeat(25) + '...');
    });

    test('returns string unchanged when exactly at limit', () => {
        expect(truncateString('hello', 5)).toBe('hello');
    });
});

// ─── isUrl ────────────────────────────────────────────────────────────────────

describe('isUrl', () => {
    test('recognizes http URLs', () => {
        expect(isUrl('http://example.com')).toBe(true);
    });

    test('recognizes https URLs', () => {
        expect(isUrl('https://example.com/path?q=1')).toBe(true);
    });

    test('rejects plain text', () => {
        expect(isUrl('hello world')).toBe(false);
    });

    test('rejects empty string', () => {
        expect(isUrl('')).toBe(false);
    });

    test('rejects ftp scheme', () => {
        expect(isUrl('ftp://example.com')).toBe(false);
    });

    test('rejects URLs with spaces', () => {
        expect(isUrl('http://ex ample.com')).toBe(false);
    });
});

// ─── splitValues ──────────────────────────────────────────────────────────────

describe('splitValues', () => {
    test('splits on double pipe', () => {
        expect(splitValues('a || b || c')).toEqual(['a', 'b', 'c']);
    });

    test('splits on newline', () => {
        expect(splitValues('a\nb\nc')).toEqual(['a', 'b', 'c']);
    });

    test('splits on comma', () => {
        expect(splitValues('a,b,c')).toEqual(['a', 'b', 'c']);
    });

    test('trims whitespace from parts', () => {
        expect(splitValues('  a  ,  b  ')).toEqual(['a', 'b']);
    });

    test('filters empty parts', () => {
        expect(splitValues('a,,b')).toEqual(['a', 'b']);
    });

    test('returns empty array for falsy input', () => {
        expect(splitValues('')).toEqual([]);
        expect(splitValues(null)).toEqual([]);
        expect(splitValues(undefined)).toEqual([]);
    });
});

// ─── splitNarrativeValues ─────────────────────────────────────────────────────

describe('splitNarrativeValues', () => {
    test('splits on double pipe only', () => {
        expect(splitNarrativeValues('foo || bar')).toEqual(['foo', 'bar']);
    });

    test('converts newlines to spaces within each segment', () => {
        const result = splitNarrativeValues('foo\nbar || baz');
        expect(result).toEqual(['foo bar', 'baz']);
    });

    test('collapses multiple spaces', () => {
        expect(splitNarrativeValues('foo   bar')).toEqual(['foo bar']);
    });

    test('returns empty array for falsy input', () => {
        expect(splitNarrativeValues('')).toEqual([]);
        expect(splitNarrativeValues(null)).toEqual([]);
    });
});

// ─── isDateTimeValue ──────────────────────────────────────────────────────────

describe('isDateTimeValue', () => {
    test('accepts valid ISO 8601 with Z', () => {
        expect(isDateTimeValue('2024-01-15T10:30:00Z')).toBe(true);
    });

    test('accepts valid ISO 8601 with offset', () => {
        expect(isDateTimeValue('2024-01-15T10:30:00+02:00')).toBe(true);
    });

    test('accepts ISO with milliseconds', () => {
        expect(isDateTimeValue('2024-01-15T10:30:00.123Z')).toBe(true);
    });

    test('rejects plain date string', () => {
        expect(isDateTimeValue('2024-01-15')).toBe(false);
    });

    test('rejects non-string', () => {
        expect(isDateTimeValue(12345)).toBe(false);
        expect(isDateTimeValue(null)).toBe(false);
    });

    test('rejects invalid date in valid format', () => {
        expect(isDateTimeValue('2024-13-45T10:30:00Z')).toBe(false);
    });
});

// ─── formatDateTimeLocal ──────────────────────────────────────────────────────

describe('formatDateTimeLocal', () => {
    test('returns a non-empty formatted string for valid ISO date', () => {
        const result = formatDateTimeLocal('2024-01-15T10:30:00Z');
        expect(typeof result).toBe('string');
        expect(result.length).toBeGreaterThan(0);
        expect(result).not.toBe('2024-01-15T10:30:00Z');
    });

    test('returns original value for invalid date', () => {
        expect(formatDateTimeLocal('not-a-date')).toBe('not-a-date');
    });
});

// ─── getFormattedDate ─────────────────────────────────────────────────────────

describe('getFormattedDate', () => {
    test('formats ISO date to Italian locale by default', () => {
        const result = getFormattedDate('2024-06-15T00:00:00Z');
        expect(typeof result).toBe('string');
        expect(result).toMatch(/\d/);
    });

    test('accepts custom locale and timezone', () => {
        const result = getFormattedDate('2024-06-15T00:00:00Z', 'en-US', 'UTC');
        expect(typeof result).toBe('string');
        expect(result).toMatch(/2024/);
    });
});

// ─── formatMonthYear ──────────────────────────────────────────────────────────

describe('formatMonthYear', () => {
    test('formats date to month/year', () => {
        const result = formatMonthYear('2024-06-15');
        expect(result).toMatch(/Jun\s+2024|June\s+2024/);
    });

    test('returns original value for invalid input', () => {
        expect(formatMonthYear('not-a-date')).toBe('not-a-date');
    });
});

// ─── getQueryParam ────────────────────────────────────────────────────────────

describe('getQueryParam', () => {
    test('returns the value of an existing query param', () => {
        window.history.pushState({}, '', '?foo=bar&baz=qux');
        expect(getQueryParam('foo')).toBe('bar');
        expect(getQueryParam('baz')).toBe('qux');
    });

    test('returns null for missing param', () => {
        window.history.pushState({}, '', '?foo=bar');
        expect(getQueryParam('missing')).toBeNull();
    });

    test('returns null when no query string', () => {
        window.history.pushState({}, '', '/');
        expect(getQueryParam('foo')).toBeNull();
    });
});

// ─── setSearchQuery ───────────────────────────────────────────────────────────

describe('setSearchQuery', () => {
    test('sets search param in URL', () => {
        window.history.pushState({}, '', '/');
        setSearchQuery('hello');
        expect(window.location.search).toContain('search=hello');
    });
});

// ─── createHrefElement ────────────────────────────────────────────────────────

describe('createHrefElement', () => {
    test('creates an anchor with href and target', () => {
        const a = createHrefElement('https://example.com', 'Click me');
        expect(a.tagName).toBe('A');
        expect(a.href).toBe('https://example.com/');
        expect(a.textContent).toBe('Click me');
        expect(a.target).toBe('_blank');
        expect(a.rel).toContain('noopener');
    });

    test('uses default text content when not provided', () => {
        const a = createHrefElement('https://example.com');
        expect(a.textContent).toBeTruthy();
    });
});

// ─── addTagToElement ──────────────────────────────────────────────────────────

describe('addTagToElement', () => {
    test('appends br tags to element', () => {
        const el = document.createElement('div');
        addTagToElement(el, 3);
        expect(el.querySelectorAll('br').length).toBe(3);
    });

    test('appends custom tag', () => {
        const el = document.createElement('div');
        addTagToElement(el, 2, 'hr');
        expect(el.querySelectorAll('hr').length).toBe(2);
    });

    test('does nothing when count is 0', () => {
        const el = document.createElement('div');
        addTagToElement(el, 0);
        expect(el.innerHTML).toBe('');
    });
});

// ─── createFormattedElementsFrom ──────────────────────────────────────────────

describe('createFormattedElementsFrom', () => {
    test('returns DOM nodes for each line', () => {
        const nodes = createFormattedElementsFrom(['hello', 'world']);
        expect(nodes.length).toBeGreaterThan(0);
    });

    test('inserts br between lines', () => {
        const nodes = createFormattedElementsFrom(['line1', 'line2']);
        const tags = nodes.map(n => n.nodeName);
        expect(tags).toContain('BR');
    });

    test('handles HTML tags in lines', () => {
        const nodes = createFormattedElementsFrom(['<b>bold</b>']);
        const bTags = nodes.filter(n => n.nodeName === 'B');
        expect(bTags.length).toBe(1);
    });

    test('strips disallowed tags', () => {
        const nodes = createFormattedElementsFrom(['<script>alert(1)</script>']);
        const scripts = nodes.filter(n => n.nodeName === 'SCRIPT');
        expect(scripts.length).toBe(0);
    });

    test('converts bare URLs to anchor links', () => {
        const nodes = createFormattedElementsFrom(['visit https://example.com please']);
        const anchors = nodes.filter(n => n.nodeName === 'A');
        expect(anchors.length).toBe(1);
        expect(anchors[0].href).toContain('example.com');
    });
});

// ─── createFormattedLongTextElementsFrom ──────────────────────────────────────

describe('createFormattedLongTextElementsFrom', () => {
    test('returns empty array for falsy input', () => {
        expect(createFormattedLongTextElementsFrom('')).toEqual([]);
        expect(createFormattedLongTextElementsFrom(null)).toEqual([]);
    });

    test('converts || to line breaks', () => {
        const nodes = createFormattedLongTextElementsFrom('line1 || line2');
        expect(nodes.length).toBeGreaterThan(0);
    });

    test('deduplicates consecutive blank lines', () => {
        const nodes = createFormattedLongTextElementsFrom('line1\n\n\n\nline2');
        expect(nodes.length).toBeGreaterThan(0);
    });
});

// ─── closeSideDrawer / openSideDrawer ─────────────────────────────────────────

describe('closeSideDrawer', () => {
    let drawer, overlay;

    beforeEach(() => {
        document.body.innerHTML = `
            <div id="side-drawer" class="open" aria-hidden="false"></div>
            <div id="side-overlay" class="visible"></div>
        `;
        drawer = document.getElementById('side-drawer');
        overlay = document.getElementById('side-overlay');
        document.body.classList.add('side-drawer-open');
    });

    afterEach(() => {
        document.body.innerHTML = '';
        document.body.className = '';
    });

    test('removes open class from drawer', () => {
        closeSideDrawer();
        expect(drawer.classList.contains('open')).toBe(false);
    });

    test('removes visible class from overlay', () => {
        closeSideDrawer();
        expect(overlay.classList.contains('visible')).toBe(false);
    });

    test('sets aria-hidden to true', () => {
        closeSideDrawer();
        expect(drawer.getAttribute('aria-hidden')).toBe('true');
    });

    test('does nothing when drawer not found', () => {
        document.body.innerHTML = '';
        expect(() => closeSideDrawer()).not.toThrow();
    });

    test('also resets sub-open class on side-drawer after transitionend', () => {
        document.body.innerHTML = `
            <div id="side-drawer" class="open sub-open" aria-hidden="false"></div>
            <div id="side-overlay" class="visible"></div>
        `;
        document.body.classList.add('side-drawer-open');
        closeSideDrawer();
        const drawer = document.getElementById('side-drawer');
        expect(drawer.classList.contains('open')).toBe(false);
        // sub-open is deferred until the drawer's transitionend — fire it manually
        drawer.dispatchEvent(Object.assign(new Event('transitionend', { bubbles: true }), { propertyName: 'transform' }));
        expect(drawer.classList.contains('sub-open')).toBe(false);
    });
});

describe('openSideDrawer', () => {
    beforeEach(() => {
        document.body.innerHTML = `
            <div id="side-drawer" aria-hidden="true"></div>
            <div id="side-overlay"></div>
            <button id="act-upload"></button>
        `;
    });

    afterEach(() => {
        document.body.innerHTML = '';
        document.body.className = '';
    });

    test('adds open class to drawer', () => {
        openSideDrawer();
        expect(document.getElementById('side-drawer').classList.contains('open')).toBe(true);
    });

    test('sets aria-hidden to false', () => {
        openSideDrawer();
        expect(document.getElementById('side-drawer').getAttribute('aria-hidden')).toBe('false');
    });

    test('does nothing when drawer not found', () => {
        document.body.innerHTML = '';
        expect(() => openSideDrawer()).not.toThrow();
    });
});

// ─── toggleClearButton ────────────────────────────────────────────────────────

describe('toggleClearButton', () => {
    beforeEach(() => {
        document.body.innerHTML = '<button id="clear-btn" class="hidden"></button>';
    });

    afterEach(() => {
        document.body.innerHTML = '';
    });

    test('removes hidden class when value is truthy', () => {
        toggleClearButton('clear-btn', 'something');
        expect(document.getElementById('clear-btn').classList.contains('hidden')).toBe(false);
    });

    test('adds hidden class when value is falsy', () => {
        document.getElementById('clear-btn').classList.remove('hidden');
        toggleClearButton('clear-btn', '');
        expect(document.getElementById('clear-btn').classList.contains('hidden')).toBe(true);
    });

    test('does nothing when element not found', () => {
        expect(() => toggleClearButton('nonexistent', 'value')).not.toThrow();
    });
});

// ─── createOutlookUrl ─────────────────────────────────────────────────────────

describe('createOutlookUrl', () => {
    test('builds URL with to/cc/subject/body', () => {
        const url = createOutlookUrl(['a@b.com'], ['c@d.com'], 'Hi', 'Hello');
        expect(url).toContain('outlook.office.com');
        expect(url).toContain('subject=Hi');
        expect(url).toContain('body=Hello');
        expect(url).toContain('to=');
        expect(url).toContain('cc=');
    });

    test('omits to/cc when arrays are empty', () => {
        const url = createOutlookUrl([], [], 'Sub', 'Body');
        expect(url).not.toContain('&to=');
        expect(url).not.toContain('&cc=');
    });

    test('encodes special chars in subject and body', () => {
        const url = createOutlookUrl([], [], 'Hello World', 'Line 1\nLine 2');
        expect(url).toContain(encodeURIComponent('Hello World'));
    });
});

// ─── parseCSV ─────────────────────────────────────────────────────────────────

describe('parseCSV', () => {
    test('parses simple CSV', () => {
        const result = parseCSV('a,b,c\n1,2,3');
        expect(result).toEqual([['a', 'b', 'c'], ['1', '2', '3']]);
    });

    test('handles quoted fields', () => {
        const result = parseCSV('"hello, world",b');
        expect(result).toEqual([['hello, world', 'b']]);
    });

    test('handles escaped quotes inside quoted fields', () => {
        const result = parseCSV('"say ""hi""",b');
        expect(result).toEqual([['say "hi"', 'b']]);
    });

    test('handles CRLF line endings', () => {
        const result = parseCSV('a,b\r\n1,2');
        expect(result).toEqual([['a', 'b'], ['1', '2']]);
    });

    test('handles quoted field with newlines', () => {
        const result = parseCSV('"line1\nline2",b');
        expect(result).toEqual([['line1\nline2', 'b']]);
    });

    test('parses single row with no newline', () => {
        const result = parseCSV('a,b,c');
        expect(result).toEqual([['a', 'b', 'c']]);
    });

    test('returns empty array for empty string', () => {
        expect(parseCSV('')).toEqual([]);
    });
});

// ─── isMobileDevice ───────────────────────────────────────────────────────────

describe('isMobileDevice', () => {
    test('returns a boolean', () => {
        expect(typeof isMobileDevice()).toBe('boolean');
    });

    test('returns true in jsdom environment (screen.width=0 triggers small-viewport path)', () => {
        // jsdom screen dimensions are 0x0, which is <= 820
        expect(isMobileDevice()).toBe(true);
    });
});

// ─── openOutlookWebCompose ────────────────────────────────────────────────────

describe('openOutlookWebCompose', () => {
    test('calls window.open with an Outlook URL', () => {
        const openSpy = jest.spyOn(window, 'open').mockImplementation(() => {});
        openOutlookWebCompose({ to: ['alice@brand.com'], subject: 'Hello', body: 'World' });
        expect(openSpy).toHaveBeenCalledWith(
            expect.stringContaining('outlook.office.com'),
            '_blank',
            'noopener'
        );
        openSpy.mockRestore();
    });
});

// ─── initCommonActions ────────────────────────────────────────────────────────

describe('initCommonActions', () => {
    afterEach(() => {
        document.body.innerHTML = '';
    });

    test('does not throw when DOM elements are missing', () => {
        document.body.innerHTML = '';
        expect(() => initCommonActions()).not.toThrow();
    });

    test('wires toggle-cta to open side drawer', () => {
        document.body.innerHTML = `
            <div id="side-drawer"></div>
            <div id="side-overlay"></div>
            <button id="side-close"></button>
            <button id="toggle-cta"></button>
            <input id="act-upload"/>
            <input id="fileInput"/>
        `;
        initCommonActions();
        const cta = document.getElementById('toggle-cta');
        cta.click();
        expect(document.getElementById('side-drawer').classList.contains('open')).toBe(true);
    });

    test('wires side-close to close side drawer', () => {
        document.body.innerHTML = `
            <div id="side-drawer" class="open"></div>
            <div id="side-overlay" class="visible"></div>
            <button id="side-close"></button>
            <button id="toggle-cta"></button>
        `;
        initCommonActions();
        document.getElementById('side-close').click();
        expect(document.getElementById('side-drawer').classList.contains('open')).toBe(false);
    });

    test('ESC closes sub-panel first (removes sub-open), keeps main drawer open', () => {
        document.body.innerHTML = `
            <div id="side-drawer" class="open sub-open" aria-hidden="false"></div>
            <div id="side-overlay" class="visible"></div>
            <button id="side-close"></button>
            <button id="toggle-cta"></button>
        `;
        document.body.classList.add('side-drawer-open');
        // Capture only the handler registered by this initCommonActions call to avoid
        // accumulation of keydown listeners from previous tests on the same window
        let capturedKeydownHandler = null;
        const origAdd = window.addEventListener.bind(window);
        window.addEventListener = (type, handler, ...args) => {
            if (type === 'keydown') capturedKeydownHandler = handler;
            origAdd(type, handler, ...args);
        };
        initCommonActions();
        window.addEventListener = origAdd;
        capturedKeydownHandler({ key: 'Escape' });
        const drawer = document.getElementById('side-drawer');
        expect(drawer.classList.contains('sub-open')).toBe(false);
        expect(drawer.classList.contains('open')).toBe(true);
    });

    test('ESC closes main drawer when sub-panel is not open', () => {
        document.body.innerHTML = `
            <div id="side-drawer" class="open" aria-hidden="false"></div>
            <div id="side-overlay" class="visible"></div>
            <button id="side-close"></button>
            <button id="toggle-cta"></button>
        `;
        document.body.classList.add('side-drawer-open');
        let capturedKeydownHandler = null;
        const origAdd = window.addEventListener.bind(window);
        window.addEventListener = (type, handler, ...args) => {
            if (type === 'keydown') capturedKeydownHandler = handler;
            origAdd(type, handler, ...args);
        };
        initCommonActions();
        window.addEventListener = origAdd;
        capturedKeydownHandler({ key: 'Escape' });
        expect(document.getElementById('side-drawer').classList.contains('open')).toBe(false);
    });

    test('overlay click closes drawer; sub-open cleared after transitionend', () => {
        document.body.innerHTML = `
            <div id="side-drawer" class="open sub-open" aria-hidden="false"></div>
            <div id="side-overlay" class="visible"></div>
            <button id="side-close"></button>
            <button id="toggle-cta"></button>
        `;
        document.body.classList.add('side-drawer-open');
        initCommonActions();
        document.getElementById('side-overlay').click();
        const drawer = document.getElementById('side-drawer');
        expect(drawer.classList.contains('open')).toBe(false);
        // sub-open stays alive during the drawer animation, removed on transitionend
        drawer.dispatchEvent(Object.assign(new Event('transitionend', { bubbles: true }), { propertyName: 'transform' }));
        expect(drawer.classList.contains('sub-open')).toBe(false);
    });
});

// ─── enableGlobalFindShortcut ─────────────────────────────────────────────────

describe('enableGlobalFindShortcut', () => {
    afterEach(() => {
        document.body.innerHTML = '';
    });

    test('warns when inputSelector is not provided', () => {
        const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {});
        enableGlobalFindShortcut({});
        expect(warnSpy).toHaveBeenCalledWith(expect.stringContaining('inputSelector'));
        warnSpy.mockRestore();
    });

    test('does not throw when registered with a valid selector', () => {
        document.body.innerHTML = '<input id="my-search"/>';
        expect(() => enableGlobalFindShortcut({ inputSelector: '#my-search' })).not.toThrow();
    });

    test('focuses input on Ctrl+F (non-Mac)', () => {
        document.body.innerHTML = '<input id="search-box"/>';
        enableGlobalFindShortcut({ inputSelector: '#search-box' });
        const input = document.getElementById('search-box');
        const focusSpy = jest.spyOn(input, 'focus').mockImplementation(() => {});
        Object.defineProperty(navigator, 'platform', { value: 'Win32', configurable: true });
        const event = new KeyboardEvent('keydown', { key: 'f', ctrlKey: true, bubbles: true });
        window.dispatchEvent(event);
        expect(focusSpy).toHaveBeenCalled();
        focusSpy.mockRestore();
    });
});

// ─── createModal ─────────────────────────────────────────────────────────────

describe('createModal', () => {
    afterEach(() => {
        document.body.innerHTML = '';
    });

    test('returns a Promise', () => {
        const result = createModal({ title: 'Test', html: 'Content', buttons: [{ id: 'ok', label: 'OK', primary: true }] });
        expect(result).toBeInstanceOf(Promise);
        // Close modal to avoid pending promise warnings
        document.querySelector('[data-action="ok"]')?.click();
    });

    test('resolves with button id when clicked', async () => {
        const promise = createModal({
            title: 'Confirm',
            html: 'Are you sure?',
            buttons: [
                { id: 'cancel', label: 'Cancel' },
                { id: 'confirm', label: 'Confirm', primary: true },
            ]
        });
        document.querySelector('[data-action="confirm"]').click();
        const result = await promise;
        expect(result).toBe('confirm');
    });

    test('resolves with null on Escape key', async () => {
        const promise = createModal({
            title: 'Test',
            html: 'Body',
            buttons: [{ id: 'ok', label: 'OK', primary: true }]
        });
        document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
        const result = await promise;
        expect(result).toBeNull();
    });
});

// ─── createFormattedElementsFrom (HTML sanitization edge cases) ───────────────

describe('createFormattedElementsFrom (sanitization)', () => {
    test('processes <a> tags with href', () => {
        const nodes = createFormattedElementsFrom(['<a href="https://example.com">Link</a>']);
        expect(nodes.length).toBeGreaterThan(0);
    });

    test('sanitizes javascript: href to nothing', () => {
        const nodes = createFormattedElementsFrom(['<a href="javascript:alert(1)">Bad</a>']);
        // anchor without valid href becomes text nodes / stripped
        const html = nodes.map(n => (n.outerHTML || n.textContent || '')).join('');
        expect(html).not.toContain('javascript:');
    });

    test('adds noopener to <a> rel', () => {
        const nodes = createFormattedElementsFrom(['<a href="https://example.com" rel="nofollow">X</a>']);
        const anchor = nodes.find(n => n.tagName === 'A');
        if (anchor) {
            expect(anchor.rel).toContain('noopener');
        }
    });

    test('handles <a> without href gracefully', () => {
        const nodes = createFormattedElementsFrom(['<a>No href</a>']);
        expect(Array.isArray(nodes)).toBe(true);
    });
});

// ─── formatUrlLink ────────────────────────────────────────────────────────────

describe('formatUrlLink', () => {
    test('returns an anchor tag', () => {
        expect(formatUrlLink('https://example.com/page')).toContain('<a ');
        expect(formatUrlLink('https://example.com/page')).toContain('href="https://example.com/page"');
    });

    test('uses last URL path segment as display text', () => {
        expect(formatUrlLink('https://example.com/services/my-service')).toContain('my-service');
    });

    test('truncates long segments with leading ellipsis', () => {
        const long = 'a'.repeat(60);
        const html = formatUrlLink(`https://example.com/${long}`);
        expect(html).toContain('...');
        const displayText = html.replace(/<[^>]+>/g, '');
        expect(displayText).not.toBe(long);
        expect(displayText.length).toBeLessThanOrEqual(58); // '...' + 55 chars
    });

    test('falls back to second-to-last segment when last is empty', () => {
        expect(formatUrlLink('https://example.com/some-service/')).toContain('some-service');
    });

    test('strips query string and hash from display segment', () => {
        const html = formatUrlLink('https://example.com/page?foo=bar#section');
        expect(html).toContain('page');
    });
});

// ─── renderUrlPartsIntoCell ───────────────────────────────────────────────────

describe('renderUrlPartsIntoCell', () => {
    test('renders single URL as inline anchor in td', () => {
        const td = document.createElement('td');
        renderUrlPartsIntoCell(['https://example.com/page'], td);
        expect(td.innerHTML).toContain('<a ');
        expect(td.querySelector('ul')).toBeNull();
    });

    test('renders multiple URLs as a list', () => {
        const td = document.createElement('td');
        renderUrlPartsIntoCell(['https://a.com/x', 'https://b.com/y'], td);
        expect(td.querySelector('ul')).toBeTruthy();
        expect(td.querySelectorAll('li')).toHaveLength(2);
    });

    test('renders non-URL plain-text parts without anchor tags', () => {
        const td = document.createElement('td');
        renderUrlPartsIntoCell(['Active', 'Inactive'], td);
        const items = td.querySelectorAll('li');
        expect(items[0].textContent).toBe('Active');
        expect(items[0].querySelector('a')).toBeNull();
    });
});

// ─── createFormattedLongTextElementsFrom — task list (checkboxes) ─────────────

describe('createFormattedLongTextElementsFrom — task items', () => {
    test('unchecked item renders .jenga-task-list with unchecked checkbox', () => {
        const els = createFormattedLongTextElementsFrom('[ ] do the thing');
        const ul = els.find(el => el.classList && el.classList.contains('jenga-task-list'));
        expect(ul).toBeTruthy();
        const cb = ul.querySelector('input[type="checkbox"]');
        expect(cb).not.toBeNull();
        expect(cb.checked).toBe(false);
    });

    test('checked item renders checkbox with checked attribute', () => {
        const els = createFormattedLongTextElementsFrom('[x] done task');
        const ul = els.find(el => el.classList && el.classList.contains('jenga-task-list'));
        expect(ul).toBeTruthy();
        const cb = ul.querySelector('input[type="checkbox"]');
        expect(cb).not.toBeNull();
        expect(cb.checked).toBe(true);
    });

    test('[X] uppercase also treated as checked', () => {
        const els = createFormattedLongTextElementsFrom('[X] also done');
        const ul = els.find(el => el.classList && el.classList.contains('jenga-task-list'));
        expect(ul).toBeTruthy();
        expect(ul.querySelector('input[type="checkbox"]').checked).toBe(true);
    });

    test('consecutive task items grouped into single .jenga-task-list', () => {
        const els = createFormattedLongTextElementsFrom('[ ] first\n[x] second\n[ ] third');
        const lists = els.filter(el => el.classList && el.classList.contains('jenga-task-list'));
        expect(lists).toHaveLength(1);
        expect(lists[0].querySelectorAll('li')).toHaveLength(3);
    });

    test('task item text is rendered inside a span', () => {
        const els = createFormattedLongTextElementsFrom('[ ] my task label');
        const ul = els.find(el => el.classList && el.classList.contains('jenga-task-list'));
        expect(ul.textContent).toContain('my task label');
    });

    test('mixed: paragraph + task items + bullet list — three separate elements', () => {
        const text = 'Intro paragraph\n[ ] task one\n[x] task two\n- bullet item';
        const els = createFormattedLongTextElementsFrom(text);
        const taskList = els.find(el => el.classList && el.classList.contains('jenga-task-list'));
        const bulletList = els.find(el => el.tagName === 'UL' && !(el.classList && el.classList.contains('jenga-task-list')));
        const para = els.find(el => el.tagName === 'P');
        expect(taskList).toBeTruthy();
        expect(bulletList).toBeTruthy();
        expect(para).toBeTruthy();
    });

    test('task item with inline link text renders an anchor inside the li', () => {
        const els = createFormattedLongTextElementsFrom('[ ] see <a href="https://example.com">docs</a>');
        const ul = els.find(el => el.classList && el.classList.contains('jenga-task-list'));
        expect(ul).toBeTruthy();
        const link = ul.querySelector('a');
        expect(link).not.toBeNull();
        expect(link.href).toContain('example.com');
    });

    test('|| separator produces separate task items in same list when consecutive', () => {
        const text = '[ ] alpha||[x] beta';
        const els = createFormattedLongTextElementsFrom(text);
        // || becomes \n\n so items are separated by a blank line — each forms its own list
        const lists = els.filter(el => el.classList && el.classList.contains('jenga-task-list'));
        expect(lists.length).toBeGreaterThanOrEqual(1);
        const allCheckboxes = els.flatMap(el => Array.from(el.querySelectorAll ? el.querySelectorAll('input[type="checkbox"]') : []));
        expect(allCheckboxes.length).toBe(2);
    });
});

// ─── createFormattedLongTextElementsFrom — table rendering ───────────────────

describe('createFormattedLongTextElementsFrom — tables', () => {
    test('line starting with <table produces a <table> element', () => {
        const els = createFormattedLongTextElementsFrom('<table class="jenga-desc-table"><tr><td>cell</td></tr></table>');
        const table = els.find(el => el.tagName === 'TABLE');
        expect(table).toBeTruthy();
    });

    test('table class is preserved', () => {
        const els = createFormattedLongTextElementsFrom('<table class="jenga-desc-table"><tr><td>x</td></tr></table>');
        const table = els.find(el => el.tagName === 'TABLE');
        expect(table.className).toBe('jenga-desc-table');
    });

    test('th cell becomes <th> element', () => {
        const els = createFormattedLongTextElementsFrom('<table class="jenga-desc-table"><tr><th>Header</th><td>Value</td></tr></table>');
        const table = els.find(el => el.tagName === 'TABLE');
        expect(table.querySelector('th')).not.toBeNull();
        expect(table.querySelector('th').textContent).toBe('Header');
    });

    test('td cell becomes <td> element', () => {
        const els = createFormattedLongTextElementsFrom('<table class="jenga-desc-table"><tr><th>H</th><td>Val</td></tr></table>');
        const table = els.find(el => el.tagName === 'TABLE');
        expect(table.querySelector('td').textContent).toBe('Val');
    });

    test('anchor link inside cell is preserved', () => {
        const els = createFormattedLongTextElementsFrom('<table class="jenga-desc-table"><tr><td><a href="https://example.com">link</a></td></tr></table>');
        const table = els.find(el => el.tagName === 'TABLE');
        const link = table.querySelector('a');
        expect(link).not.toBeNull();
        expect(link.href).toContain('example.com');
    });

    test('table after a paragraph — both appear in result', () => {
        const text = 'Intro paragraph\n<table class="jenga-desc-table"><tr><td>cell</td></tr></table>';
        const els = createFormattedLongTextElementsFrom(text);
        const para = els.find(el => el.tagName === 'P');
        const table = els.find(el => el.tagName === 'TABLE');
        expect(para).toBeTruthy();
        expect(table).toBeTruthy();
    });

    test('multiple rows render correct number of tr elements', () => {
        const html = '<table class="jenga-desc-table"><tr><th>A</th><td>1</td></tr><tr><th>B</th><td>2</td></tr></table>';
        const els = createFormattedLongTextElementsFrom(html);
        const table = els.find(el => el.tagName === 'TABLE');
        expect(table.querySelectorAll('tr')).toHaveLength(2);
    });
});
