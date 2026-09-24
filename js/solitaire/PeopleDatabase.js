import {
    getFormattedDate,
    normalizeKey,
    normalizeWs,
    parseCSV,
    splitValues,
    truncateString,
} from '../shared/utils.js';
import { BRAND } from '../../brand-specific/brand.js';

import {
    emailField,
    firstOrgLevel,
    firstLevelNA,
    secondOrgLevel,
    secondLevelNA,
    thirdOrgLevel,
    thirdLevelNA,
    COMPANY_FIELD,
} from './constants.js';

import {
    buildCompositeKey,
    filterDismissedTeams,
    filterOrganizationByStreams,
    filterOrganizationByQuickFilter,
    getAllowedStreamsSet,
    isOnlyContributorsRow,
} from './orgUtils.js';

const GUEST_ROLES_MAP = new Map([
    ['Team Product Manager',      ['Product Manager']],
    ['Team Delivery Manager',     ['Delivery Manager']],
    ['Team Scrum Master',         ['Coach']],
    ['Team Solution Architect',   ['Solution Architect']],
    ['Team Development Manager',  ['Development Manager']],
    ['Team Service Manager',      ['Service Manager']],
    ['Team Contributors',         ['Contributor']],
    ['Team Security Champion',    ['Security Champion']],
]);

export class PeopleDatabase {
    constructor(app) {
        this.app = app;
        this.people = [];
        this.cachedCsvText = null;
        this.roleDetailsMapping = new Map();
        this.functionDetailsMapping = new Map();
        this.guestRolesMap = GUEST_ROLES_MAP;
        this.guestRoleColumns = Array.from(GUEST_ROLES_MAP.keys());
    }

    load(csvText) {
        if (!csvText) {
            alert('Missing CSV File!');
            return null;
        }

        this.app.legend.colorKeyMappings = new Map();
        this.roleDetailsMapping = new Map();
        this.functionDetailsMapping = new Map();
        this.cachedCsvText = csvText;

        const csvRows = parseCSV(csvText);
        if (csvRows.length < 2) return null;

        const headers = csvRows[0].map(h => h.trim());

        this.people = csvRows.slice(1).map(row => {
            const obj = {};
            headers.forEach((h, i) => {
                obj[h] = normalizeWs(row[i] || '', h);
            });
            return obj;
        }).filter(p => (p.Status || '').toLowerCase() !== 'inactive');

        let lastUpdateISO = '';
        if (headers.includes('Updated')) {
            const idx = headers.indexOf('Updated');
            const dates = csvRows.slice(1)
                .map(r => r[idx]?.trim())
                .filter(Boolean)
                .map(d => new Date(d))
                .filter(d => !isNaN(d));
            if (dates.length) {
                const maxTs = Math.max(...dates.map(d => d.getTime()));
                lastUpdateISO = new Date(maxTs).toISOString().slice(0, 10);
            }
        }

        const peopleCount = this.people.length;
        const datasetVersion = `people:${peopleCount}|lu:${lastUpdateISO || 'n/a'}`;
        this.app.scenario.lsKey = `dsm-layout-v1::${datasetVersion}`;

        this._updateLastUpdateLabel(headers, csvRows);

        const { organization, streamBoosts, themeBoosts, teamBoosts } = this._buildOrganization(this.people);
        const organizationWithManagers = this._addGuestManagersTo(organization);
        const organizationAfterDismissed = filterDismissedTeams(organizationWithManagers);
        const filteredStreams = getAllowedStreamsSet();

        const allStreamNames = Object.keys(organizationAfterDismissed || {})
            .filter(s => s && !s.includes(firstLevelNA));

        const visibleStreamNames = (filteredStreams && filteredStreams.size > 0)
            ? allStreamNames.filter(s => filteredStreams.has(s) || filteredStreams.has(normalizeKey(s)))
            : allStreamNames;

        const streamFilteredOrg = filterOrganizationByStreams(organizationAfterDismissed, filteredStreams);
        const qfConstraints = this.app.quickFilters?.getConstraints?.() ?? null;
        const visibleOrg = qfConstraints
            ? filterOrganizationByQuickFilter(streamFilteredOrg, qfConstraints, organization)
            : streamFilteredOrg;
        this.app.visibleOrg = visibleOrg;

        return {
            people: this.people,
            organization,
            organizationWithManagers: visibleOrg,
            filteredStreams: qfConstraints ? null : filteredStreams,
            visibleStreamNames,
            headers,
            visibleOrg,
            streamBoosts,
            themeBoosts,
            teamBoosts,
        };
    }

    _updateLastUpdateLabel(headers, csvRows) {
        if (!headers.includes('Updated')) return;
        const dateIndex = headers.indexOf('Updated');
        const dates = csvRows.slice(1)
            .map(row => row[dateIndex]?.trim())
            .filter(Boolean)
            .map(d => new Date(d))
            .filter(d => !isNaN(d.getTime()));

        if (dates.length > 0) {
            const lastUpdateEl = document.getElementById('side-last-update');
            if (lastUpdateEl) {
                lastUpdateEl.textContent = `Last Update: ${getFormattedDate(new Date(Math.max(...dates.map(d => d.getTime()))).toISOString())}`;
            }
        }
    }

    _buildOrganization(people) {
        const organization = {};
        const streamBoosts = {};
        const themeBoosts  = {};
        const teamBoosts   = {};
        const parseBoost = v => { const n = parseInt((v || '').split('||')[0], 10); return isNaN(n) ? null : n; };
        for (const person of people) {
            let firstLevelItems = (person[firstOrgLevel] || '').split(/\n|,/).map(s => s.trim()).filter(Boolean);
            if (firstLevelItems.length === 0) firstLevelItems = [firstLevelNA];

            let secondLevelItems = (person[secondOrgLevel] || '').split(/\n|,/).map(t => t.trim()).filter(Boolean);
            if (secondLevelItems.length === 0) secondLevelItems = [secondLevelNA];

            let thirdLevelItems = (person[thirdOrgLevel] || '').split(/\n|,/).map(t => t.trim()).filter(Boolean);
            if (thirdLevelItems.length === 0) thirdLevelItems = [thirdLevelNA];

            const sb = parseBoost(person['Team Stream Visual Boost']);
            const tb = parseBoost(person['Team Theme Visual Boost']);
            const mb = parseBoost(person['Team Visual Boost']);
            for (const s of firstLevelItems) {
                if (sb !== null && !(s in streamBoosts)) streamBoosts[s] = sb;
                for (const t of secondLevelItems) {
                    const sk = `${s}::${t}`;
                    if (tb !== null && !(sk in themeBoosts)) themeBoosts[sk] = tb;
                    for (const m of thirdLevelItems) {
                        const mk = `${s}::${t}::${m}`;
                        if (mb !== null && !(mk in teamBoosts)) teamBoosts[mk] = mb;
                    }
                }
            }

            for (const firstLevelItem of firstLevelItems) {
                if (!organization[firstLevelItem]) organization[firstLevelItem] = {};

                for (const theme of secondLevelItems) {
                    if (!organization[firstLevelItem][theme]) organization[firstLevelItem][theme] = {};

                    for (const team of thirdLevelItems) {
                        if (!organization[firstLevelItem][theme][team]) organization[firstLevelItem][theme][team] = [];

                        person.Name = person.Name ? this._cleanName(person.Name) : person.User;
                        person.Name = this._cleanName(person.Name || '')
                            || (person.User || '').trim()
                            || (person[emailField] || '').trim()
                            || 'Unknown';

                        if (!this.roleDetailsMapping.has(person.Role)) {
                            this.roleDetailsMapping.set(person.Role, {
                                grants: person['Role Grants'],
                                description: person['Role Description'],
                            });
                        }

                        const fn = (person['Function'] || '').trim();
                        if (fn && !this.functionDetailsMapping.has(fn)) {
                            this.functionDetailsMapping.set(fn, {
                                description: person['Function Description'] || '',
                            });
                        }

                        const teamArr = organization[firstLevelItem][theme][team];
                        const existingKeys = new Set(
                            teamArr.map(p => buildCompositeKey(p, emailField)).filter(Boolean)
                        );

                        const compositeKey = buildCompositeKey(person, emailField);
                        const isFullyEmptyKey = !compositeKey;
                        const isDuplicate = compositeKey ? existingKeys.has(compositeKey) : false;

                        if (!isFullyEmptyKey && !isDuplicate) {
                            teamArr.push(person);
                        }
                    }
                }
            }
        }
        return { organization, streamBoosts, themeBoosts, teamBoosts };
    }

    _addGuestManagersTo(organization) {
        const result = {};

        for (const [firstLevel, secondLevelItems] of Object.entries(organization)) {
            for (const [secondLevel, thirdLevelItems] of Object.entries(secondLevelItems)) {
                for (const [thirdLevel, members] of Object.entries(thirdLevelItems)) {
                    if (!result[firstLevel]) result[firstLevel] = {};
                    if (!result[firstLevel][secondLevel]) result[firstLevel][secondLevel] = {};
                    if (!result[firstLevel][secondLevel][thirdLevel]) result[firstLevel][secondLevel][thirdLevel] = [];

                    members.forEach(p => {
                        if (isOnlyContributorsRow(p)) return;
                        if (!result[firstLevel][secondLevel][thirdLevel].some(m => m.Name === p.Name)) {
                            result[firstLevel][secondLevel][thirdLevel].push(p);
                        }
                    });

                    members.forEach(p => {
                        this.guestRoleColumns.forEach(role =>
                            this._addGuestManagersByRole(
                                p,
                                role,
                                result[firstLevel][secondLevel][thirdLevel],
                                organization
                            )
                        );
                    });
                }
            }
        }

        return result;
    }

    _addGuestManagersByRole(person, guestRole, thirdLevel, organization) {
        if (!person[guestRole]) return;

        const guestNames = [...new Set(
            splitValues(person[guestRole] || '')
                .flatMap(v => v.split(/\n|,/))
                .map(v => v.trim())
                .filter(Boolean)
        )];

        guestNames.forEach(name => {
            const manager = this._findPersonByName(name, organization);
            if (!manager) return;
            const alreadyPresent = thirdLevel.some(member => this._cleanName(member.Name) === this._cleanName(name));
            if (!alreadyPresent) {
                thirdLevel.push({ ...manager, guestRole });
            }
        });
    }

    _findPersonByName(targetName, result) {
        const target = normalizeWs(targetName).toLowerCase();
        return Object.values(result).flatMap(stream =>
            Object.values(stream).flatMap(theme =>
                Object.values(theme).flatMap(team => team)
            )
        ).find(person => {
            const pn = normalizeWs(person?.Name).toLowerCase();
            return pn === target;
        }) || null;
    }

    _cleanName(name) {
        return normalizeWs(name);
    }

    _findHeaderIndex(headers, name) {
        const target = (name || '').trim().toLowerCase();
        return headers.findIndex(h => (h || '').trim().toLowerCase() === target);
    }

    isInternalCompany(member) {
        return ((member[COMPANY_FIELD] || '').trim().toLowerCase() === BRAND.internalCompanyName);
    }

    aggregateInfoByHeader(members, headers, headerName, sortElements = false, splitter = splitValues) {
        const idx = this._findHeaderIndex(headers, headerName);
        if (idx === -1) return { exists: false, items: [] };

        const headerRealName = headers[idx];
        const set = new Set();

        members.forEach(m => {
            const raw = m[headerRealName];
            if (!raw) return;
            splitter(raw).forEach(v => set.add(v));
        });

        const itemsToReturn = sortElements
            ? [...set].sort((a, b) => a.localeCompare(b, 'it', { sensitivity: 'base' }))
            : [...set];

        return { exists: true, items: itemsToReturn };
    }

    truncate(s) {
        return truncateString(s);
    }

    getMembersByEmail(email) {
        if (!email) return [];
        const target = email.trim().toLowerCase();
        return this.people.filter(p => (p[emailField] || '').trim().toLowerCase() === target);
    }
}
