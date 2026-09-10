/**
 * brand-specific/brand.js
 *
 */

export const BRAND = {
    /** Display name of the company (used in About text, tooltips, etc.) */
    name: 'Nycosoft',

    /** Footer tagline rendered in the SVG canvas beneath the logo */
    tagline: 'Digital Service Management',

    /** Logo asset paths (relative to the HTML page, i.e. inside dist/) */
    logo: {
        png: './assets/brand-logo.png',

        /** Inline SVG path used to render the logo watermark on the Solitaire canvas */
        svgInline: {
            d: `
M 5.1 53.553 L 0 53.553 L 1.875 0.003 L 2.775 0.003 L 40.5 45.603 L 40.95 45.603 L 39.3 1.053 L 44.4 1.053 L 42.525 54.603 L 41.625 54.603 L 3.9 9.003 L 3.45 9.003 L 5.1 53.553 Z
M 147.9 47.103 L 148.35 47.103 L 147.9 51.678 Q 145.8 52.578 142.8 53.253 Q 139.8 53.928 136.725 54.266 Q 133.65 54.603 131.25 54.603 Q 125.175 54.603 120 52.541 Q 114.825 50.478 110.963 46.803 Q 107.1 43.128 104.925 38.141 Q 102.75 33.153 102.75 27.303 Q 102.75 21.453 104.925 16.466 Q 107.1 11.478 110.963 7.803 Q 114.825 4.128 120 2.066 Q 125.175 0.003 131.25 0.003 Q 133.725 0.003 136.8 0.378 Q 139.875 0.753 142.8 1.391 Q 145.725 2.028 147.75 2.853 L 147.375 7.053 L 146.925 7.053 Q 144.15 4.503 139.95 3.153 Q 135.75 1.803 131.25 1.803 Q 124.5 1.803 119.363 4.991 Q 114.225 8.178 111.338 13.916 Q 108.45 19.653 108.45 27.303 Q 108.45 35.028 111.338 40.728 Q 114.225 46.428 119.325 49.616 Q 124.425 52.803 131.25 52.803 Q 136.35 52.803 140.738 51.153 Q 145.125 49.503 147.9 47.103 Z
M 223.875 51.903 L 224.325 47.403 L 224.775 47.403 Q 227.7 49.953 231.375 51.378 Q 235.05 52.803 238.65 52.803 Q 243.75 52.803 246.9 49.991 Q 250.05 47.178 250.05 42.453 Q 250.05 39.828 248.888 37.578 Q 247.725 35.328 245.063 33.003 Q 242.4 30.678 237.75 27.978 Q 234.525 26.103 231.6 24.041 Q 228.675 21.978 226.8 19.241 Q 224.925 16.503 224.925 12.678 Q 224.925 6.978 229.05 3.491 Q 233.175 0.003 239.925 0.003 Q 242.925 0.003 246.375 0.641 Q 249.825 1.278 251.7 2.253 L 251.325 6.453 L 250.875 6.453 Q 246.6 1.803 239.925 1.803 Q 235.125 1.803 232.2 4.241 Q 229.275 6.678 229.275 10.803 Q 229.275 13.878 230.888 16.091 Q 232.5 18.303 235.05 20.066 Q 237.6 21.828 240.525 23.478 Q 243.9 25.428 247.087 27.603 Q 250.275 29.778 252.337 32.816 Q 254.4 35.853 254.4 40.428 Q 254.4 44.628 252.45 47.816 Q 250.5 51.003 246.938 52.803 Q 243.375 54.603 238.65 54.603 Q 235.65 54.603 231.6 53.853 Q 227.55 53.103 223.875 51.903 Z
M 389.55 53.553 L 384.45 53.553 L 384.45 2.553 Q 383.325 2.553 379.8 3.003 L 365.1 4.878 L 365.1 1.053 L 408.9 1.053 L 408.9 4.878 L 394.2 3.003 Q 390.675 2.553 389.55 2.553 L 389.55 53.553 Z
M 79.425 53.553 L 74.025 53.553 L 74.025 31.353 L 55.425 1.053 L 61.125 1.053 L 77.85 28.503 L 78.3 28.503 L 91.65 1.053 L 97.35 1.053 L 79.425 30.303 L 79.425 53.553 Z
M 340.875 53.553 L 335.775 53.553 L 335.775 1.053 L 357.825 1.053 L 357.825 4.728 L 345.375 3.003 Q 342 2.553 340.875 2.553 L 340.875 26.028 Q 342.075 26.028 345.525 25.803 L 357 25.053 L 357 28.503 L 345.525 27.753 Q 342.075 27.528 340.875 27.528 L 340.875 53.553 Z
M 158.1 27.303 Q 158.1 21.528 160.163 16.541 Q 162.225 11.553 165.938 7.841 Q 169.65 4.128 174.638 2.066 Q 179.625 0.003 185.4 0.003 Q 191.175 0.003 196.163 2.066 Q 201.15 4.128 204.862 7.841 Q 208.575 11.553 210.638 16.541 Q 212.7 21.528 212.7 27.303 Q 212.7 33.078 210.638 38.066 Q 208.575 43.053 204.862 46.766 Q 201.15 50.478 196.163 52.541 Q 191.175 54.603 185.4 54.603 Q 179.625 54.603 174.638 52.541 Q 169.65 50.478 165.938 46.766 Q 162.225 43.053 160.163 38.066 Q 158.1 33.078 158.1 27.303 Z
M 163.8 27.303 Q 163.8 34.953 166.5 40.691 Q 169.2 46.428 174.075 49.616 Q 178.95 52.803 185.4 52.803 Q 191.85 52.803 196.725 49.616 Q 201.6 46.428 204.3 40.653 Q 207 34.878 207 27.303 Q 207 19.728 204.3 13.991 Q 201.6 8.253 196.725 5.028 Q 191.85 1.803 185.4 1.803 Q 178.95 1.803 174.112 4.991 Q 169.275 8.178 166.538 13.916 Q 163.8 19.653 163.8 27.303 Z
M 265.425 27.303 Q 265.425 21.528 267.488 16.541 Q 269.55 11.553 273.263 7.841 Q 276.975 4.128 281.963 2.066 Q 286.95 0.003 292.725 0.003 Q 298.5 0.003 303.488 2.066 Q 308.475 4.128 312.188 7.841 Q 315.9 11.553 317.963 16.541 Q 320.025 21.528 320.025 27.303 Q 320.025 33.078 317.963 38.066 Q 315.9 43.053 312.188 46.766 Q 308.475 50.478 303.488 52.541 Q 298.5 54.603 292.725 54.603 Q 286.95 54.603 281.963 52.541 Q 276.975 50.478 273.263 46.766 Q 269.55 43.053 267.488 38.066 Q 265.425 33.078 265.425 27.303 Z
M 271.125 27.303 Q 271.125 34.953 273.825 40.691 Q 276.525 46.428 281.4 49.616 Q 286.275 52.803 292.725 52.803 Q 299.175 52.803 304.05 49.616 Q 308.925 46.428 311.625 40.653 Q 314.325 34.878 314.325 27.303 Q 314.325 19.728 311.625 13.991 Q 308.925 8.253 304.05 5.028 Q 299.175 1.803 292.725 1.803 Q 286.275 1.803 281.438 4.991 Q 276.6 8.178 273.863 13.916 Q 271.125 19.653 271.125 27.303 Z
        `,
            viewBoxX: -2,
            viewBoxY: -2,
            viewBoxW: 412.9,
            viewBoxH: 58.607,
            fillRule: 'nonzero',
            strokeWidth: 0.5,
            strokeLineCap: 'round',
        },
    },


    /**
     * Company identifier used to distinguish internal staff from suppliers.
     * Compared case-insensitively against the "Company" CSV field.
     */
    internalCompanyName: 'Nycosoft',

    /** External URLs */
    urls: {
        /** ServiceNow "Report a change" form (used in team cards and the side drawer) */
        reportChange: 'https://itsm-now.com/esc?id=sc_cat_item&table=sc_cat_item&sys_id=45e3f6ac4793f25023ef7f61e36d4345',

        /** Confluence/Jira article linked in the Solitaire About panel */
        servicePortal: 'https://itsmdigital.atlassian.net/servicedesk/customer/portal/1/article/3750527008?source=topic',

        /** Domino side-drawer CTA: open a new service introduction request */
        jiraNewRequest: 'https://itsmdigital.atlas.net/servicedesk/customer/portal/1/create/157',

        /** Domino side-drawer CTA: open a service info update request */
        jiraUpdateRequest: 'https://itsmdigital.atlas.net/servicedesk/customer/portal/1/create/174',
    },

    /** CSV data sources fetched at startup */
    csv: {
        /** Solitaire org-chart data (relative to the HTML page, i.e. inside dist/) */
        solitaire: 'people-database.csv',

        /** Domino service-catalog data (relative to the HTML page, i.e. inside dist/) */
        domino: 'service-catalog.csv',

        /** Jira cards CSV for Domino badge overlay (optional — set null to disable) */
        jiraCards: './jira-cards.csv',

        /** Jenga change-calendar events CSV */
        jenga: './jenga-events.csv',

        /** Quick-filter presets CSV for Solitaire (optional — set null to disable) */
        customFilters: './custom-filters.csv',
    },

    /** Jira-specific config (used by js/domino/jira.js) */
    jira: {
        siteUrl: 'https://itsmdigital.atlas.net',
        managedServicesProject: 'Nycosoft Managed Services Support',
    },
};

/**
 * Hydrates the <svg id="brand-logo"> placeholder in the current page with the
 * brand logo path from BRAND.logo.svgInline. Call once on DOMContentLoaded.
 */
export function renderBrandLogo() {
    const el = document.getElementById('brand-logo');
    if (!el) return;

    const {
        d,
        viewBoxX = 0,
        viewBoxY = 0,
        viewBoxW,
        viewBoxH,
        fillRule = 'nonzero',
        strokeWidth = 0,
        strokeLineCap = 'round',
    } = BRAND.logo.svgInline;

    const normalizedD = d.replace(/\s+/g, ' ').trim();

    el.setAttribute('viewBox', `${viewBoxX} ${viewBoxY} ${viewBoxW} ${viewBoxH}`);
    el.setAttribute('aria-label', `${BRAND.name} logo`);
    el.setAttribute('fill', 'none');
    el.setAttribute('xmlns', 'http://www.w3.org/2000/svg');

    el.innerHTML = `
        <path
            d="${normalizedD}"
            fill="var(--_g-logo-fill-color)"
            stroke="var(--_g-logo-fill-color)"
            stroke-width="${strokeWidth}"
            stroke-linecap="${strokeLineCap}"
            fill-rule="${fillRule}"
        ></path>
    `;
}
``


