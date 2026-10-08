import { PeopleDatabase } from '../../js/solitaire/PeopleDatabase.js';
import { BRAND } from '../../brand-specific/brand.js';

function makeApp() {
    return {
        legend: { colorKeyMappings: new Map() },
        scenario: { lsKey: 'dsm-layout-v1:default' },
        visibleOrg: null,
    };
}

const MINIMAL_CSV = [
    'Name,Email,Status,Team Stream,Team Theme,Team member of,Role',
    'Alice,alice@ex.com,Active,StreamA,ThemeX,Team1,Engineer',
    'Bob,bob@ex.com,Active,StreamA,ThemeX,Team1,Designer',
    'Inactive,x@x.com,Inactive,StreamA,ThemeX,Team1,Engineer',
].join('\n');

describe('PeopleDatabase constructor', () => {
    test('initializes with empty people array', () => {
        const db = new PeopleDatabase(makeApp());
        expect(db.people).toEqual([]);
    });

    test('has guestRoleColumns list', () => {
        const db = new PeopleDatabase(makeApp());
        expect(Array.isArray(db.guestRoleColumns)).toBe(true);
        expect(db.guestRoleColumns.length).toBeGreaterThan(0);
    });
});

describe('PeopleDatabase.load', () => {
    let db;

    beforeEach(() => {
        global.alert = jest.fn();
        window.history.pushState({}, '', '/');
        document.body.innerHTML = '<span id="side-last-update"></span>';
        db = new PeopleDatabase(makeApp());
    });

    afterEach(() => {
        document.body.innerHTML = '';
    });

    test('returns null and calls alert when csvText is missing', () => {
        const result = db.load('');
        expect(result).toBeNull();
        expect(global.alert).toHaveBeenCalledWith('Missing CSV File!');
    });

    test('returns null for CSV with fewer than 2 rows', () => {
        expect(db.load('Name,Email')).toBeNull();
    });

    test('parses active people and filters out inactive', () => {
        const result = db.load(MINIMAL_CSV);
        expect(result).not.toBeNull();
        expect(result.people).toHaveLength(2);
        const names = result.people.map(p => p.Name);
        expect(names).toContain('Alice');
        expect(names).toContain('Bob');
        expect(names).not.toContain('Inactive');
    });

    test('returns headers array', () => {
        const result = db.load(MINIMAL_CSV);
        expect(Array.isArray(result.headers)).toBe(true);
        expect(result.headers).toContain('Name');
        expect(result.headers).toContain('Email');
    });

    test('builds organization structure', () => {
        const result = db.load(MINIMAL_CSV);
        expect(result.organization).toBeTruthy();
        expect(result.organization['StreamA']).toBeTruthy();
        expect(result.organization['StreamA']['ThemeX']).toBeTruthy();
        expect(result.organization['StreamA']['ThemeX']['Team1']).toHaveLength(2);
    });

    test('returns organizationWithManagers', () => {
        const result = db.load(MINIMAL_CSV);
        expect(result.organizationWithManagers).toBeTruthy();
    });

    test('sets lsKey on scenario', () => {
        db.load(MINIMAL_CSV);
        expect(db.app.scenario.lsKey).toContain('dsm-layout-v1::people:2|lu:');
    });

    test('team name containing a comma is treated as a single team', () => {
        const csv = [
            'Name,Email,Status,Team Stream,Team Theme,Team member of',
            'Alice,alice@ex.com,Active,StreamA,ThemeX,"4. Operating Roles - Squads, Tribes, Chapters"',
        ].join('\n');
        const result = db.load(csv);
        const teams = Object.keys(result.organization['StreamA']['ThemeX']);
        expect(teams).toHaveLength(1);
        expect(teams[0]).toBe('4. Operating Roles - Squads, Tribes, Chapters');
    });

    test('handles person with no team assignment → falls into NA buckets', () => {
        const csv = 'Name,Email,Status\nDave,d@d.com,Active';
        const result = db.load(csv);
        expect(result).not.toBeNull();
        expect(result.people).toHaveLength(1);
    });

    test('deduplicates same person in same team', () => {
        const csv = [
            'Name,Email,Status,Team Stream,Team Theme,Team member of',
            'Alice,alice@ex.com,Active,StreamA,ThemeX,Team1',
            'Alice,alice@ex.com,Active,StreamA,ThemeX,Team1',
        ].join('\n');
        const result = db.load(csv);
        expect(result.organization['StreamA']['ThemeX']['Team1']).toHaveLength(1);
    });

    test('updates last-update label when Updated column is present', () => {
        const csv = [
            'Name,Email,Status,Updated',
            'Alice,a@a.com,Active,2024-06-15T00:00:00Z',
        ].join('\n');
        db.load(csv);
        const el = document.getElementById('side-last-update');
        expect(el.textContent).toContain('Last Update:');
    });
});

describe('PeopleDatabase.isInternalCompany', () => {
    const db = new PeopleDatabase(makeApp());

    test('returns true for internal company name (case-insensitive)', () => {
        expect(db.isInternalCompany({ Company: BRAND.name })).toBe(true);
        expect(db.isInternalCompany({ Company: BRAND.name.toUpperCase() })).toBe(true);
        expect(db.isInternalCompany({ Company: BRAND.internalCompanyName })).toBe(true);
    });

    test('returns false for other companies', () => {
        expect(db.isInternalCompany({ Company: 'Accenture' })).toBe(false);
        expect(db.isInternalCompany({ Company: '' })).toBe(false);
    });

    test('handles missing Company field', () => {
        expect(db.isInternalCompany({})).toBe(false);
    });
});

describe('PeopleDatabase.aggregateInfoByHeader', () => {
    const db = new PeopleDatabase(makeApp());
    const members = [
        { Role: 'Engineer' },
        { Role: 'Designer' },
        { Role: 'Engineer' },
    ];
    const headers = ['Name', 'Role'];

    test('returns exists:false when header not found', () => {
        const result = db.aggregateInfoByHeader(members, headers, 'Missing');
        expect(result).toEqual({ exists: false, items: [] });
    });

    test('returns unique values for a field', () => {
        const result = db.aggregateInfoByHeader(members, headers, 'Role');
        expect(result.exists).toBe(true);
        expect(result.items).toHaveLength(2);
        expect(result.items).toContain('Engineer');
        expect(result.items).toContain('Designer');
    });

    test('sorts items when sortElements is true', () => {
        const result = db.aggregateInfoByHeader(members, headers, 'Role', true);
        expect(result.items[0]).toBe('Designer');
        expect(result.items[1]).toBe('Engineer');
    });
});

describe('PeopleDatabase.load — SOL-2 Visual Boost maps', () => {
    let db;

    beforeEach(() => {
        global.alert = jest.fn();
        window.history.pushState({}, '', '/');
        document.body.innerHTML = '<span id="side-last-update"></span>';
        db = new PeopleDatabase(makeApp());
    });

    afterEach(() => {
        document.body.innerHTML = '';
    });

    test('load() return value includes streamBoosts, themeBoosts, teamBoosts keys', () => {
        const result = db.load(MINIMAL_CSV);
        expect(result).toHaveProperty('streamBoosts');
        expect(result).toHaveProperty('themeBoosts');
        expect(result).toHaveProperty('teamBoosts');
    });

    test('all three maps are empty when boost columns are absent', () => {
        const result = db.load(MINIMAL_CSV);
        expect(result.streamBoosts).toEqual({});
        expect(result.themeBoosts).toEqual({});
        expect(result.teamBoosts).toEqual({});
    });

    test('stream boost is populated from Team Stream Visual Boost column', () => {
        const csv = [
            'Name,Email,Status,Team Stream,Team Theme,Team member of,Team Stream Visual Boost',
            'Alice,alice@ex.com,Active,StreamA,ThemeX,Team1,80',
        ].join('\n');
        const result = db.load(csv);
        expect(result.streamBoosts['StreamA']).toBe(80);
    });

    test('theme boost is keyed as stream::theme', () => {
        const csv = [
            'Name,Email,Status,Team Stream,Team Theme,Team member of,Team Theme Visual Boost',
            'Alice,alice@ex.com,Active,StreamA,ThemeX,Team1,50',
        ].join('\n');
        const result = db.load(csv);
        expect(result.themeBoosts['StreamA::ThemeX']).toBe(50);
    });

    test('team boost is keyed as stream::theme::team', () => {
        const csv = [
            'Name,Email,Status,Team Stream,Team Theme,Team member of,Team Visual Boost',
            'Alice,alice@ex.com,Active,StreamA,ThemeX,Team1,30',
        ].join('\n');
        const result = db.load(csv);
        expect(result.teamBoosts['StreamA::ThemeX::Team1']).toBe(30);
    });

    test('boost value 0 is stored (treated as valid, not null)', () => {
        const csv = [
            'Name,Email,Status,Team Stream,Team Theme,Team member of,Team Stream Visual Boost',
            'Alice,alice@ex.com,Active,StreamA,ThemeX,Team1,0',
        ].join('\n');
        const result = db.load(csv);
        expect(result.streamBoosts['StreamA']).toBe(0);
    });

    test('first non-null boost wins — later rows do not overwrite', () => {
        const csv = [
            'Name,Email,Status,Team Stream,Team Theme,Team member of,Team Stream Visual Boost',
            'Alice,alice@ex.com,Active,StreamA,ThemeX,Team1,80',
            'Bob,bob@ex.com,Active,StreamA,ThemeX,Team1,10',
        ].join('\n');
        const result = db.load(csv);
        expect(result.streamBoosts['StreamA']).toBe(80);
    });

    test('multi-value pipe-separated boost uses first token', () => {
        const csv = [
            'Name,Email,Status,Team Stream,Team Theme,Team member of,Team Stream Visual Boost',
            'Alice,alice@ex.com,Active,StreamA,ThemeX,Team1,70||40',
        ].join('\n');
        const result = db.load(csv);
        expect(result.streamBoosts['StreamA']).toBe(70);
    });

    test('non-numeric boost is ignored (empty maps for that entry)', () => {
        const csv = [
            'Name,Email,Status,Team Stream,Team Theme,Team member of,Team Stream Visual Boost',
            'Alice,alice@ex.com,Active,StreamA,ThemeX,Team1,high',
        ].join('\n');
        const result = db.load(csv);
        expect(result.streamBoosts['StreamA']).toBeUndefined();
    });

    test('multiple streams receive their own boost values', () => {
        const csv = [
            'Name,Email,Status,Team Stream,Team Theme,Team member of,Team Stream Visual Boost',
            'Alice,alice@ex.com,Active,StreamA,ThemeX,Team1,80',
            'Bob,bob@ex.com,Active,StreamB,ThemeY,Team2,40',
        ].join('\n');
        const result = db.load(csv);
        expect(result.streamBoosts['StreamA']).toBe(80);
        expect(result.streamBoosts['StreamB']).toBe(40);
    });
});

describe('PeopleDatabase.truncate', () => {
    const db = new PeopleDatabase(makeApp());

    test('truncates long strings', () => {
        const long = 'a'.repeat(30);
        expect(db.truncate(long)).toMatch(/\.\.\.$/);
    });

    test('returns short strings unchanged', () => {
        expect(db.truncate('hello')).toBe('hello');
    });
});
