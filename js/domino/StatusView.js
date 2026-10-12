/* global BRAND */
import * as d3 from 'd3';

// ── Exclude 3rd-party and vendor services from the status page ────────────────
const HIDE_TYPES = new Set(['SaaS', 'Service Provider']);

// ── Friendly names for technical service identifiers ─────────────────────────
const FRIENDLY_NAMES = {
    'pdp': 'Product Detail Page',
    'plp': 'Product Listing Page',
    'sign-in': 'Sign-In',
    'sign-in-spa': 'Sign-In SPA',
    'sign-in-fragments': 'Sign-In Fragments',
    'checkout-service': 'Checkout Service',
    'checkout-app': 'Checkout App',
    'checkout-spa': 'Checkout SPA',
    'commerce-payment-gateway': 'Payment Gateway',
    'order-fulfillment-service': 'Order Fulfillment',
    'order-confirmation-service': 'Order Confirmation',
    'order-api': 'Order API',
    'order-validation-service': 'Order Validation',
    'order-number-generator-service': 'Order Number Generator',
    'orders-cronjob-service': 'Orders Cron Job',
    'price-service-api': 'Price Service',
    'price-ingester': 'Price Ingester',
    'price-adapter': 'Price Adapter',
    'price-data-integrity-checker-workflow': 'Price Integrity Checker',
    'catalog-api': 'Catalog API',
    'catalog-api-product-worker': 'Catalog Product Worker',
    'catalog': 'Catalog Service',
    'analytics-sdk': 'Analytics SDK',
    'web-tracking-compliance-sdk': 'Tracking Compliance SDK',
    'baa-spa': 'Book-an-Appointment SPA',
    'pimgester': 'PIM Ingester',
    'pim-publisher-app': 'PIM Publisher',
    'editorial-hub': 'Editorial Hub',
    'image-server': 'Image Server',
    'image-server-postgres': 'Image Server DB',
    'img-viewer-dashboard': 'Image Viewer Dashboard',
    'idp-user-service': 'Identity Service',
    'gucci-visualmerchandizetool': 'Visual Merch Tool',
    'product-management-service': 'Product Management',
    'product-service-api': 'Product Service API',
    'product-personalization-cfg': 'Personalisation Config',
    'product-personalization-view': 'Personalisation View',
    'orchestrator': 'Content Orchestrator',
    'esm-proxy': 'ESM Proxy',
    'feature-flags': 'Feature Flags',
    'feature-toggle': 'Feature Toggle',
    'beacon': 'Beacon Tracker',
    'buy-area': 'Buy Area Widget',
    'solr': 'Solr Search',
    'solr-esales': 'Solr eSales',
    'hybris-api-proxy': 'Hybris API Proxy',
    'fs-microsites': 'Flagship Microsites',
    'gucci-store': 'Gucci Store (web)',
    'gucci-app-android': 'Gucci App (Android)',
    'gucci-app-ios': 'Gucci App (iOS)',
    'gucci-app-support': 'App Support Service',
    'hello-world-demo': 'Demo Service',
    'iframe-add-to-cart': 'Add-to-Cart iFrame',
    'reco': 'Recommendations',
    'yasa': 'YASA Search',
    'in-store-purchase-api': 'In-Store Purchase API',
    'my-account (order history)': 'My Account – Orders',
    'my-account-spa (settings)': 'My Account – Settings',
    'My-Gucci': 'My Gucci',
    'email-service': 'Email Service',
    'client-services': 'Client Services',
    'cust-anonymization-service': 'Customer Anonymisation',
    'store-queues-svc': 'Store Queues',
    'store-queues-svc-db': 'Store Queues DB',
    'distance-sales-db': 'Distance Sales DB',
    'distance-sales-service': 'Distance Sales',
    'payments-mao-adapter': 'Payments MAO Adapter',
    'payments-service': 'Payments Service',
    'payments-widget': 'Payments Widget',
    'shipping-widget': 'Shipping Widget',
    'stock-management-service': 'Stock Management',
    'stock-service-api': 'Stock Service API',
    'stock-integrity-checker': 'Stock Integrity Checker',
    'tax-calculator-service': 'Tax Calculator',
    'thank-you-spa': 'Thank-You Page SPA',
    'apple-pay-widget': 'Apple Pay Widget',
    'service-pages': 'Service Pages',
    'services-and-gifting': 'Services & Gifting',
    'address-validation-service': 'Address Validation',
    'Geo-Markup': 'Geo Markup',
};

// ── Service clustering rules ──────────────────────────────────────────────────
const CLUSTERS = [
    {
        id: 'hybris',
        label: 'SAP Hybris Platform',
        names: new Set(['SAP Hybris GCOM', 'SAP Hybris Villetta', 'mySQL GCOM', 'mySQL Villetta', 'hybris-api-proxy', 'solr', 'solr-esales']),
        namePatterns: [/hybris/i, /\bsap\b/i],
        descPatterns:  [/hybris/i, /\bsap commerce\b/i],
    },
    {
        id: 'sign-in',
        label: 'Sign-In & Authentication',
        names: new Set(['sign-in', 'sign-in-spa', 'sign-in-fragments', 'idp-user-service', 'SPiD', 'Security Email']),
        namePatterns: [/sign.?in/i, /\bidp[\b._-]/i, /\bspid\b/i],
        descPatterns:  [/authentication/i, /identity.?provider/i, /single.?sign.?on/i, /sso/i],
    },
    {
        id: 'customer-access',
        label: 'Customer Access (Okta)',
        names: new Set(['Okta (Customer Access)', 'Okta user proxy', 'Kering Okta', 'Okta']),
        namePatterns: [/okta/i],
        descPatterns:  [/okta/i, /customer.?identity/i, /workforce.?identity/i],
    },
    {
        id: 'checkout',
        label: 'Checkout',
        names: new Set(['Checkout', 'checkout-service', 'checkout-app', 'checkout-spa']),
        namePatterns: [/checkout/i],
        descPatterns:  [/checkout/i, /order.?completion/i, /purchase.?flow/i],
    },
    {
        id: 'cart',
        label: 'Cart & Buy Area',
        names: new Set(['Cart/Minicart', 'buy-area']),
        namePatterns: [/\bcart\b/i, /minicart/i, /buy.?area/i],
        descPatterns:  [/shopping.?cart/i, /add to cart/i],
    },
    {
        id: 'payments',
        label: 'Payments',
        names: new Set([
            'commerce-payment-gateway', 'payments-service', 'payments-widget',
            'payments-mao-adapter', 'Paypal', 'Apple Pay', 'Apple Pay Express',
            'apple-pay-widget', 'Adyen', 'Klarna', 'Toss', 'Kakao Pay',
            'Cybersource', 'Affirm', 'Amazon Pay', 'Pay By Link', 'Riskified',
            'CommerceLayer', 'MAO Proxy', 'Vertex', 'tax-calculator-service',
            'distance-sales-service', 'distance-sales-db', 'address-validation-service',
        ]),
        namePatterns: [/payment/i, /\bpay(?!pal)\b/i, /adyen/i, /paypal/i, /klarna/i, /cybersource/i, /affirm/i, /commerce.?layer/i, /riskified/i, /vertex/i],
        descPatterns:  [/payment gateway/i, /\bpsp\b/i, /payment.?processing/i, /billing/i],
    },
    {
        id: 'orders',
        label: 'Orders',
        names: new Set([
            'order-api', 'order-confirmation-service', 'order-fulfillment-service',
            'order-number-generator-service', 'order-validation-service',
            'orders-cronjob-service', 'Order Confirmation Email',
            'Order Confirmation Email (New Checkout)', 'Thank you page (New Checkout)',
            'thank-you-spa', 'Collect In-Store', 'shipping-widget',
            'iPad In-Store', 'in-store-purchase-api', 'services-and-gifting',
        ]),
        namePatterns: [/^order/i, /order.?(api|confirm|fulfil|valid|number|cronjob)/i, /fulfil/i, /thank.?you/i, /collect.?in.?store/i, /in.?store.?purchase/i],
        descPatterns:  [/order.?management/i, /post-purchase/i, /order.?lifecycle/i],
    },
    {
        id: 'pricing',
        label: 'Pricing',
        names: new Set(['price-service-api', 'price-ingester', 'price-adapter', 'price-data-integrity-checker-workflow']),
        namePatterns: [/^price/i],
        descPatterns:  [/pricing/i, /price.?management/i, /price.?engine/i],
    },
    {
        id: 'stock',
        label: 'Stock & Inventory',
        names: new Set(['stock-management-service', 'stock-service-api', 'stock-integrity-checker', 'Relaxed Stock Service', 'SAP (Stock Management)']),
        namePatterns: [/stock/i, /inventory/i],
        descPatterns:  [/stock.?management/i, /inventory/i, /warehouse/i],
    },
    {
        id: 'pim',
        label: 'Product Catalog & PIM',
        names: new Set(['Akeneo (PIM)', 'pimgester', 'pim-publisher-app', 'catalog-api', 'catalog-api-product-worker', 'product-service-api', 'Pdp Material Composition Flow', 'Concession']),
        namePatterns: [/\bpim\b/i, /\bakeneo\b/i, /^catalog/i, /catalog.?(api|worker)/i],
        descPatterns:  [/product.?information.?management/i, /\bpim\b/i, /catalog.?management/i],
    },
    {
        id: 'product-pages',
        label: 'Product Pages (PDP / PLP)',
        names: new Set(['pdp', 'plp', 'iframe-add-to-cart', 'product-management-service']),
        namePatterns: [/^pdp$/i, /^plp$/i],
        descPatterns:  [/product.?detail/i, /product.?listing/i, /pdp/i, /plp/i],
    },
    {
        id: 'search',
        label: 'Search & Discovery',
        names: new Set(['Algolia', 'New Search', 'Search', 'yasa', 'Cataloger', 'gucci-visualmerchandizetool', 'Find In Store', 'Recommendation Engine', 'reco', 'product-personalization-cfg', 'product-personalization-view', 'feature-toggle', 'catalog']),
        namePatterns: [/\bsearch\b/i, /\balgolia\b/i, /\byasa\b/i, /cataloger/i, /visual.?merch/i, /recommend/i, /personaliz/i],
        descPatterns:  [/\bsearch\b/i, /product.?discovery/i, /merchandis/i, /personaliz/i, /recommendation/i],
    },
    {
        id: 'my-account',
        label: 'My Account & Customer Profile',
        names: new Set(['My Account', 'my-account (order history)', 'my-account-spa (settings)', 'My-Gucci', 'Customer Profiles', 'CRM Customer DB', 'Kering Customer Interaction DB', 'Wishlist', 'cust-anonymization-service', 'Dynamo DB', 'Kering Kafka', 'Kanonymous', 'client-services']),
        namePatterns: [/my.?account/i, /my.?gucci/i, /customer.?profile/i, /crm.?customer/i, /cust.?anon/i, /wishlist/i],
        descPatterns:  [/customer.?account/i, /user.?profile/i, /loyalty/i, /crm/i],
    },
    {
        id: 'store-locator',
        label: 'Store Locator & Appointments',
        names: new Set(['Store Locator', 'Book an appointment', 'Book An Appointment Confirmation Email', 'baa-spa', 'store-queues-svc', 'store-queues-svc-db', 'Store DB (Kering)', 'Google Maps', 'Powerfront']),
        namePatterns: [/store.?locator/i, /book.?appoint/i, /\bbaa\b/i, /store.?queue/i],
        descPatterns:  [/store.?location/i, /appointment/i, /clienteling/i],
    },
    {
        id: 'email',
        label: 'Email & Notifications',
        names: new Set(['email-service', 'Account Cancellation Email', 'Account Creation Email', 'OTP Email', 'Order Confirmation Email', 'Newsletter']),
        namePatterns: [/email/i, /\bnewsletter\b/i, /\botp\b/i],
        descPatterns:  [/transactional.?email/i, /email.?notification/i, /notification.?service/i],
    },
    {
        id: 'content',
        label: 'Content & CMS',
        names: new Set(['Contentful', 'editorial-hub', 'Bubbles', 'Collection/Capsule Pages', 'BrightCove', 'Media Service', 'Optimizely', 'fs-microsites', 'gucci-store', 'GDE_Activator']),
        namePatterns: [/contentful/i, /editorial/i, /brightcove/i, /optimizely/i, /microsite/i],
        descPatterns:  [/content.?management/i, /\bcms\b/i, /editorial/i, /digital.?content/i],
    },
    {
        id: 'images',
        label: 'Image & Media Serving',
        names: new Set(['image-server', 'image-server-postgres', 'img-viewer-dashboard']),
        namePatterns: [/image.?server/i, /img.?viewer/i],
        descPatterns:  [/image.?serve/i, /media.?cdn/i, /image.?optim/i],
    },
    {
        id: 'analytics',
        label: 'Analytics & Tracking',
        names: new Set(['Google Analytics', 'Google Tag Manager', 'analytics-sdk', 'web-tracking-compliance-sdk', 'One Trust', 'Wunderkind', 'MyFeelBack', 'beacon']),
        namePatterns: [/analytics/i, /tag.?manager/i, /wunderkind/i, /myfeelback/i, /one.?trust/i, /\bbeacon\b/i, /tracking/i],
        descPatterns:  [/analytics/i, /tracking/i, /tag.?management/i, /consent/i],
    },
    {
        id: 'platform-infra',
        label: 'Platform Infrastructure',
        names: new Set(['Akamai', 'Bifrost', 'esm-proxy', 'orchestrator', 'feature-flags', 'service-pages', 'Geo-Markup', 'Gucci Backoffice', 'Static Choose Your Country Page On Akamai', 'backoffice']),
        namePatterns: [/akamai/i, /bifrost/i, /esm.?proxy/i, /orchestrat/i, /feature.?flag/i, /backoffice/i, /geo.?markup/i],
        descPatterns:  [/\bcdn\b/i, /infrastructure/i, /proxy/i, /feature.?toggle/i, /server.?side/i],
    },
    {
        id: 'contact-support',
        label: 'Contact & Support',
        names: new Set(['Contact Us']),
        namePatterns: [/contact.?us/i, /live.?chat/i, /support.?widget/i],
        descPatterns:  [/customer.?support/i, /contact.?form/i, /live.?chat/i],
    },
];

function buildClusterIndex(clusters) {
    const index = new Map();
    for (const cluster of clusters) {
        for (const name of cluster.names) {
            if (!index.has(name)) index.set(name, cluster.id);
        }
    }
    return index;
}

function resolveCluster(node, index, clusters) {
    if (index.has(node.id)) return index.get(node.id);
    const nameLower = node.id.toLowerCase();
    const descLower = (node.Description || '').toLowerCase();
    for (const cluster of clusters) {
        for (const p of cluster.namePatterns) { if (p.test(nameLower)) return cluster.id; }
    }
    for (const cluster of clusters) {
        for (const p of cluster.descPatterns) { if (p.test(descLower)) return cluster.id; }
    }
    return null;
}

const CLUSTER_BY_ID = Object.fromEntries(CLUSTERS.map(c => [c.id, c]));

// ── Health helpers ────────────────────────────────────────────────────────────
const ACTIVE_STATUSES = new Set([
    'Sent to Internal Team', 'Pending', 'UAT',
    'Work in progress', 'Work In Progress', 'Scheduled',
]);

// Only P1/P2 drive the health dot/pill; P3/P4 → amber minor badge
const P1P2_RANK     = { 'P1/Blocker': 0, 'P2/Critical': 1 };
// Used only for sorting incident lists in the detail pane
const ALL_RANK      = { 'P1/Blocker': 0, 'P2/Critical': 1, 'P3/Major': 2, 'P4/Minor': 3 };
const HEALTH_FROM_RANK = ['outage', 'critical'];
const HEALTH_ORDER  = { outage: 0, critical: 1, degraded: 2, minor: 3, operational: 4 };

const HEALTH_LABEL  = {
    outage: 'Major Outage', critical: 'Critical',
    degraded: 'Degraded',  minor: 'Minor Issue',  operational: 'Operational',
};
const GLOBAL_MSG    = {
    outage:      'Major outage in progress',
    critical:    'Critical incidents active',
    degraded:    'Some services degraded',
    minor:       'Minor issues detected',
    operational: 'All systems operational',
};

// ── StatusView class ──────────────────────────────────────────────────────────
export class StatusView {
    constructor(app) {
        this.app = app;
        this.statusViewEl = null;
        this.btnStatus    = null;
        this.btnList      = null;
        this.btnGraph     = null;
        this.btnExport    = null;
        this.mapEl        = null;
        this.legendEl     = null;
        this.listViewEl   = null;
        this._cards = null;
        this._boundSearchHandler = null;
        this._boundGoHandler     = null;
        this._boundClearHandler  = null;
        this._preNavSearchTerm   = undefined;
        this._clusterIndex = buildClusterIndex(CLUSTERS);
    }

    initDOM() {
        this.statusViewEl = document.getElementById('status-view');
        this.btnStatus    = document.getElementById('view-status');
        this.btnList      = document.getElementById('view-list');
        this.btnGraph     = document.getElementById('view-graph');
        this.btnExport    = document.getElementById('act-export-excel');
        this.mapEl        = document.getElementById('map');
        this.legendEl     = document.getElementById('legend');
        this.listViewEl   = document.getElementById('list-view');
        this.btnStatus?.addEventListener('click', () => this.toStatusView());
    }

    _setActiveButton(active) {
        [this.btnList, this.btnGraph, this.btnStatus].forEach(b => {
            b?.classList.toggle('is-active', b === active);
        });
    }

    // ── View transitions ──────────────────────────────────────────────────────

    async toStatusView() {
        if (this.mapEl)      this.mapEl.style.display = 'none';
        if (this.legendEl)   this.legendEl.style.display = 'none';
        if (this.listViewEl) this.listViewEl.style.display = 'none';
        if (this.statusViewEl) this.statusViewEl.style.display = 'block';

        this._setActiveButton(this.btnStatus);

        const surl = new URL(window.location.href);
        surl.searchParams.set('view', 'status');
        surl.searchParams.delete('listView');
        surl.searchParams.delete('sort');
        window.history.replaceState({}, '', surl.toString());

        await this._ensureData();

        // Build content container exactly once
        if (!this.statusViewEl.querySelector('.status-content')) {
            this.statusViewEl.innerHTML =
                '<div class="status-view-inner"><div class="status-content"></div></div>';
        }

        // Restore search term when returning from a status-page service navigation click
        if (this._preNavSearchTerm !== undefined) {
            const restore = this._preNavSearchTerm;
            this._preNavSearchTerm = undefined;
            this.app.search.handleQuery(restore, false);
            const inp = document.getElementById('drawer-search-input');
            if (inp) inp.value = restore;
        }

        // Attach Enter-key and go-button listeners while status view is active
        if (!this._boundSearchHandler) {
            this._boundSearchHandler = (e) => {
                if (e.type === 'keydown') {
                    if (e.key !== 'Enter') return;
                    if (this.app.autocomplete?.hasPendingSelection()) return;
                }
                this._renderContent();
            };
            this._boundGoHandler = this._boundSearchHandler;
            this._boundClearHandler = () => this._renderContent();
        }
        const mainInput = document.getElementById('drawer-search-input');
        mainInput?.addEventListener('keydown', this._boundSearchHandler);
        document.getElementById('drawer-search-go')?.addEventListener('click', this._boundGoHandler);
        document.getElementById('act-clear')?.addEventListener('click', this._boundClearHandler);

        this._renderContent();
    }

    toOtherView() {
        if (this.statusViewEl) this.statusViewEl.style.display = 'none';
        const mainInput = document.getElementById('drawer-search-input');
        mainInput?.removeEventListener('keydown', this._boundSearchHandler);
        document.getElementById('drawer-search-go')?.removeEventListener('click', this._boundGoHandler);
        document.getElementById('act-clear')?.removeEventListener('click', this._boundClearHandler);
    }

    // ── Data ──────────────────────────────────────────────────────────────────

    async _ensureData() {
        if (this._cards !== null) return;
        const url = (typeof BRAND !== 'undefined' && BRAND.csv?.jiraCardsCurrent)
            ? BRAND.csv.jiraCardsCurrent
            : './jira-cards-current.csv';
        try {
            const text = await fetch(url).then(r => r.ok ? r.text() : '').catch(() => '');
            const stripComments = s => s.replace(/^#[^\n]*\n/gm, '');
            this._cards = text ? d3.csvParse(stripComments(text)) : [];
        } catch {
            this._cards = [];
        }
    }

    _buildHealthMap() {
        // Map<serviceName, { health, rank, incidents[], minorIncidents[] }>
        // P1/P2 set `health`; P3/P4 go into `minorIncidents` (amber badge, no health impact).
        const healthMap = new Map();
        for (const card of this._cards) {
            if (!ACTIVE_STATUSES.has(card.Status)) continue;
            const isP1P2 = card.Priority === 'P1/Blocker' || card.Priority === 'P2/Critical';
            const isP3P4 = card.Priority === 'P3/Major'   || card.Priority === 'P4/Minor';
            if (!isP1P2 && !isP3P4) continue;

            const services = (card.AffectedServices || '').split('||').map(s => s.trim()).filter(Boolean);
            for (const svc of services) {
                if (!healthMap.has(svc)) {
                    healthMap.set(svc, { health: 'operational', rank: 99, incidents: [], minorIncidents: [] });
                }
                const entry = healthMap.get(svc);
                if (isP1P2) {
                    const rank = P1P2_RANK[card.Priority];
                    if (rank < entry.rank) { entry.rank = rank; entry.health = HEALTH_FROM_RANK[rank]; }
                    entry.incidents.push(card);
                } else {
                    entry.minorIncidents.push(card);
                }
            }
        }
        return healthMap;
    }

    // ── Render ────────────────────────────────────────────────────────────────

    _renderContent() {
        const contentEl = this.statusViewEl?.querySelector('.status-content');
        if (!contentEl) return;

        const { store } = this.app;
        const healthMap = this._buildHealthMap();

        // Exclude decommissioned, 3rd-party types, and unmapped (no stream/theme) services
        const DEAD = new Set(['Stopped', 'Decommissioned']);
        const nodes = store.nodes.filter(n => {
            if (DEAD.has(n.Status) || n['Decommission Date']) return false;
            if (HIDE_TYPES.has(n['Type'])) return false;
            return Boolean((n['Stream'] || '').trim()) && Boolean((n['Theme'] || '').trim());
        });

        // Group by stream → theme
        const streamMap = new Map();
        for (const node of nodes) {
            const stream = node['Stream'].trim();
            const theme  = node['Theme'].trim();
            if (!streamMap.has(stream)) streamMap.set(stream, new Map());
            const themeMap = streamMap.get(stream);
            if (!themeMap.has(theme)) themeMap.set(theme, []);
            themeMap.get(theme).push(node);
        }

        const worstHealth = healths => {
            let best = 4;
            for (const h of healths) { const r = HEALTH_ORDER[h] ?? 4; if (r < best) best = r; }
            return Object.keys(HEALTH_ORDER)[best] ?? 'operational';
        };
        const sumMinor = ns => ns.reduce((s, n) => s + (healthMap.get(n.id)?.minorIncidents?.length ?? 0), 0);

        // Build hierarchical data structure
        const streamData = [];
        for (const [streamName, themeMap] of streamMap) {
            const themes = [];
            for (const [themeName, svcs] of themeMap) {
                const clusterGroups = new Map();
                const individuals = [];
                for (const node of svcs) {
                    const cid = resolveCluster(node, this._clusterIndex, CLUSTERS);
                    if (cid) {
                        if (!clusterGroups.has(cid)) clusterGroups.set(cid, []);
                        clusterGroups.get(cid).push(node);
                    } else {
                        individuals.push(node);
                    }
                }

                const clusters = [];
                for (const [cid, clusterSvcs] of clusterGroups) {
                    const cHealth = worstHealth(clusterSvcs.map(n => healthMap.get(n.id)?.health ?? 'operational'));
                    clusters.push({
                        clusterId: cid,
                        label:     CLUSTER_BY_ID[cid]?.label ?? cid,
                        health:    cHealth,
                        services:  clusterSvcs,
                        totalIncidents: clusterSvcs.reduce((s, n) => s + (healthMap.get(n.id)?.incidents?.length ?? 0), 0),
                        totalMinor:     sumMinor(clusterSvcs),
                    });
                }
                clusters.sort((a, b) => HEALTH_ORDER[a.health] - HEALTH_ORDER[b.health] || a.label.localeCompare(b.label));
                individuals.sort((a, b) => {
                    const ha = HEALTH_ORDER[healthMap.get(a.id)?.health ?? 'operational'];
                    const hb = HEALTH_ORDER[healthMap.get(b.id)?.health ?? 'operational'];
                    return ha - hb || a.id.localeCompare(b.id);
                });

                const themeMinor = clusters.reduce((s, c) => s + c.totalMinor, 0) +
                    individuals.reduce((s, n) => s + (healthMap.get(n.id)?.minorIncidents?.length ?? 0), 0);
                themes.push({
                    name:       themeName,
                    health:     worstHealth([...clusters.map(c => c.health), ...individuals.map(n => healthMap.get(n.id)?.health ?? 'operational')]),
                    clusters,
                    individuals,
                    minorCount: themeMinor,
                });
            }
            themes.sort((a, b) => HEALTH_ORDER[a.health] - HEALTH_ORDER[b.health] || a.name.localeCompare(b.name));
            const streamMinor    = themes.reduce((s, t) => s + t.minorCount, 0);
            const streamSvcCount = themes.reduce((s, t) =>
                s + t.clusters.reduce((cs, c) => cs + c.services.length, 0) + t.individuals.length, 0);
            const healthBreakdown = { outage: 0, critical: 0, degraded: 0, minor: 0, clean: 0 };
            for (const theme of themes) {
                for (const cluster of theme.clusters) {
                    if (cluster.health !== 'operational') {
                        healthBreakdown[cluster.health] = (healthBreakdown[cluster.health] || 0) + cluster.services.length;
                    } else if (cluster.totalMinor > 0) {
                        healthBreakdown.minor += cluster.services.length;
                    } else {
                        healthBreakdown.clean += cluster.services.length;
                    }
                }
                for (const node of theme.individuals) {
                    const e = healthMap.get(node.id);
                    const h = e?.health ?? 'operational';
                    if (h !== 'operational') {
                        healthBreakdown[h] = (healthBreakdown[h] || 0) + 1;
                    } else if (e?.minorIncidents?.length) {
                        healthBreakdown.minor++;
                    } else {
                        healthBreakdown.clean++;
                    }
                }
            }
            streamData.push({
                name:           streamName,
                health:         worstHealth(themes.map(t => t.health)),
                themes,
                minorCount:     streamMinor,
                svcCount:       streamSvcCount,
                healthBreakdown,
            });
        }
        streamData.sort((a, b) => b.svcCount - a.svcCount || a.name.localeCompare(b.name));
        const maxMinorCount = streamData.reduce((max, s) => Math.max(max, s.minorCount), 0);

        // Apply shared search filter — uses graph search results + raw text for stream/theme names
        const { search } = this.app;
        const rawText    = (document.getElementById('drawer-search-input')?.value ?? '').toLowerCase().trim();
        const matchedIds = (search.searchTerm && search.currentSearchedNodes)
            ? search.currentSearchedNodes : null;
        const hasFilter  = Boolean(matchedIds) || Boolean(rawText);
        const filtered   = hasFilter ? this._filterStreamData(streamData, matchedIds, rawText) : streamData;

        // Global stats from full unfiltered data
        const globalHealth = worstHealth(streamData.map(s => s.health));
        const operational  = nodes.filter(n => !healthMap.get(n.id)?.incidents?.length).length;
        const affected     = nodes.length - operational;
        const minorTotal   = nodes.reduce((s, n) => s + (healthMap.get(n.id)?.minorIncidents?.length ?? 0), 0);

        contentEl.innerHTML =
            this._renderBanner(globalHealth, operational, affected, minorTotal) +
            (hasFilter && !filtered.length ? '<p class="status-no-results">No matching services or areas found.</p>' : '') +
            filtered.map(s => this._renderStream(s, healthMap, maxMinorCount)).join('');

        this._bindInteractions(healthMap, contentEl);
    }

    _filterStreamData(streamData, matchedIds, rawText) {
        return streamData
            .map(stream => {
                if (rawText && stream.name.toLowerCase().includes(rawText)) return stream;
                const themes = stream.themes
                    .map(theme => {
                        if (rawText && theme.name.toLowerCase().includes(rawText)) return theme;
                        const clusters    = theme.clusters.filter(c =>
                            (rawText && c.label.toLowerCase().includes(rawText)) ||
                            c.services.some(n =>
                                matchedIds ? matchedIds.has(n.id) : (rawText && this._nodeMatchesQuery(n, rawText))
                            )
                        );
                        const individuals = theme.individuals.filter(n =>
                            matchedIds ? matchedIds.has(n.id) : (rawText && this._nodeMatchesQuery(n, rawText))
                        );
                        if (!clusters.length && !individuals.length) return null;
                        return { ...theme, clusters, individuals };
                    })
                    .filter(Boolean);
                if (!themes.length) return null;
                return { ...stream, themes };
            })
            .filter(Boolean);
    }

    _nodeMatchesQuery(node, q) {
        return node.id.toLowerCase().includes(q) ||
               (FRIENDLY_NAMES[node.id] || '').toLowerCase().includes(q) ||
               (node.Description || '').toLowerCase().includes(q);
    }

    // ── HTML helpers ──────────────────────────────────────────────────────────

    _renderBanner(health, operational, affected, minorTotal) {
        const minorTag = minorTotal > 0
            ? `<span class="status-global-minor">▲ ${minorTotal} minor (P3/P4)</span>` : '';
        return `
            <div class="status-global-banner status-global-banner--${health}">
                <span class="status-dot status-dot--${health}"></span>
                <span class="status-global-msg">${GLOBAL_MSG[health]}</span>
                <span class="status-global-stats">
                    <span class="status-stat status-stat--op">${operational}</span> operational &nbsp;·&nbsp;
                    <span class="status-stat status-stat--aff">${affected}</span> affected
                    ${minorTag}
                </span>
            </div>`;
    }

    _renderStream(streamData, healthMap, maxMinorCount = 0) {
        const { name, health, themes, minorCount = 0, svcCount = 0, healthBreakdown = {} } = streamData;
        const total    = svcCount || 1;
        const affCount = (healthBreakdown.outage || 0) + (healthBreakdown.critical || 0) + (healthBreakdown.degraded || 0);
        const affPct   = Math.round(affCount / total * 100);
        const minorPct = maxMinorCount > 0 ? Math.round(minorCount / maxMinorCount * 45) : 0;
        return `
            <div class="status-stream" data-health="${health}">
                <div class="status-stream-header">
                    <span class="status-dot status-dot--${health}"></span>
                    <span class="status-stream-name">${name}</span>
                    ${this._pill(health)}
                    ${minorCount > 0 ? `<span class="status-minor-badge">▲ ${minorCount}</span>` : ''}
                </div>
                <div class="status-health-bar">
                    ${affPct > 0 ? `<div class="status-hb-seg--affected" style="width:${affPct}%"></div>` : ''}
                    ${minorPct > 0 ? `<div class="status-hb-seg--minor" style="width:${minorPct}%"></div>` : ''}
                    <div class="status-hb-seg--clean" style="flex:1"></div>
                </div>
                <div class="status-stream-body">
                    ${themes.map(t => this._renderTheme(t, healthMap)).join('')}
                </div>
            </div>`;
    }

    _renderTheme(themeData, healthMap) {
        const { name, health, clusters, individuals, minorCount = 0 } = themeData;

        const affectedClusters = clusters.filter(c => c.health !== 'operational');
        const minorClusters    = clusters.filter(c => c.health === 'operational' && c.totalMinor > 0);
        const cleanClusters    = clusters.filter(c => c.health === 'operational' && !c.totalMinor);

        const affectedIndivs = individuals.filter(n => (healthMap.get(n.id)?.health ?? 'operational') !== 'operational');
        const minorIndivs    = individuals.filter(n => {
            const e = healthMap.get(n.id);
            return (e?.health ?? 'operational') === 'operational' && (e?.minorIncidents?.length ?? 0) > 0;
        });
        const cleanIndivs    = individuals.filter(n => {
            const e = healthMap.get(n.id);
            return (e?.health ?? 'operational') === 'operational' && !(e?.minorIncidents?.length);
        });

        const opCount  = cleanClusters.length + cleanIndivs.length;
        const hasMinor = minorClusters.length + minorIndivs.length > 0;

        return `
            <div class="status-theme" data-health="${health}">
                <div class="status-theme-header">
                    <span class="status-theme-name">${name}</span>
                    ${this._pill(health)}
                    ${minorCount > 0 ? `<span class="status-minor-badge">▲ ${minorCount}</span>` : ''}
                </div>
                <div class="status-theme-body">
                    ${affectedClusters.map(c => this._renderCluster(c, healthMap, true)).join('')}
                    ${affectedIndivs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                    ${hasMinor ? `
                        <div class="status-minor-section">
                            ${minorClusters.map(c => this._renderCluster(c, healthMap, false)).join('')}
                            ${minorIndivs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                        </div>` : ''}
                    ${opCount > 0 ? `
                        <div class="status-op-toggle" tabindex="0" role="button" aria-expanded="false" data-count="${opCount}">
                            <span class="status-dot status-dot--operational"></span>
                            <span class="status-op-label">Show ${opCount} fully operational ▼</span>
                        </div>
                        <div class="status-op-list" hidden>
                            ${cleanClusters.map(c => this._renderCluster(c, healthMap, false)).join('')}
                            ${cleanIndivs.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                        </div>` : ''}
                </div>
            </div>`;
    }

    _renderCluster(clusterData, healthMap, defaultExpanded) {
        const { clusterId, label, health, services, totalIncidents, totalMinor } = clusterData;
        const incBadge   = totalIncidents > 0 ? `<span class="status-badge">${totalIncidents} open</span>` : '';
        const mainSvc    = totalMinor > 0
            ? services.reduce((best, n) => {
                const a = healthMap.get(n.id)?.minorIncidents?.length ?? 0;
                const b = healthMap.get(best?.id)?.minorIncidents?.length ?? 0;
                return a >= b ? n : best;
            }, services[0])
            : null;
        const minorBadge = totalMinor > 0
            ? `<span class="status-minor-badge" data-svc-id="${this._esc(mainSvc?.id ?? '')}">▲ ${totalMinor}</span>`
            : '';
        const countTxt   = `${services.length} service${services.length !== 1 ? 's' : ''}`;

        return `
            <div class="status-cluster" data-cluster-id="${this._esc(clusterId)}" data-health="${health}"${defaultExpanded ? ' data-expanded="true"' : ''}>
                <div class="status-cluster-header" tabindex="0" role="button" aria-expanded="${defaultExpanded}">
                    <span class="status-dot status-dot--${health}"></span>
                    <span class="status-cluster-label">${this._esc(label)}</span>
                    ${this._pill(health)}
                    ${incBadge}${minorBadge}
                    <span class="status-cluster-meta">${countTxt}</span>
                    <span class="status-expand-icon" aria-hidden="true">›</span>
                </div>
                <div class="status-cluster-services" ${defaultExpanded ? '' : 'hidden'}>
                    ${services.map(n => this._renderServiceRow(n, healthMap.get(n.id))).join('')}
                </div>
            </div>`;
    }

    _renderServiceRow(node, entry) {
        const health     = entry?.health ?? 'operational';
        const hasInc     = Boolean(entry?.incidents?.length);
        const hasMinor   = Boolean(entry?.minorIncidents?.length);
        const friendly   = FRIENDLY_NAMES[node.id];
        const primary    = friendly ?? node.id;
        const code       = friendly ? node.id : null;
        const badge      = hasInc   ? `<span class="status-badge">${entry.incidents.length} open</span>` : '';
        const minorBadge = hasMinor ? `<span class="status-minor-badge" data-svc-id="${this._esc(node.id)}">▲ ${entry.minorIncidents.length}</span>` : '';
        const expand     = hasInc   ? `<span class="status-expand-icon" aria-hidden="true">›</span>` : '';

        return `
            <div class="status-svc-row${hasInc ? ' has-incidents' : ''}"
                 data-svc-id="${this._esc(node.id)}" data-health="${health}">
                <span class="status-dot status-dot--${health}"></span>
                <span class="status-svc-names">
                    <span class="status-svc-primary">${this._esc(primary)}</span>
                    ${code ? `<span class="status-svc-code">${this._esc(code)}</span>` : ''}
                </span>
                ${badge}${minorBadge}${expand}
            </div>`;
    }

    _pill(health) {
        return `<span class="status-pill status-pill--${health}">${HEALTH_LABEL[health] ?? health}</span>`;
    }

    _esc(s) {
        return String(s)
            .replace(/&/g,'&amp;').replace(/</g,'&lt;')
            .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }

    _cleanSummary(s) {
        return String(s)
            .replace(/§\[Roi\w+(?:,Roi\w+)*\]/g, '')
            .replace(/^\[GUCCI PRODUCTION\]\s*-\s*/i, '')
            .trim();
    }

    // ── Interactions ──────────────────────────────────────────────────────────

    _bindInteractions(healthMap, contentEl) {
        contentEl.querySelectorAll('.status-cluster-header').forEach(header => {
            const cluster = header.closest('.status-cluster');
            const body    = cluster?.querySelector('.status-cluster-services');
            if (!body) return;
            header.addEventListener('click', () => {
                const expanded = !body.hidden;
                body.hidden = expanded;
                header.setAttribute('aria-expanded', String(!expanded));
                const icon = header.querySelector('.status-expand-icon');
                if (icon) icon.style.transform = expanded ? '' : 'rotate(90deg)';
            });
            header.addEventListener('keydown', e => {
                if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); header.click(); }
            });
            if (cluster.dataset.expanded === 'true') {
                const icon = header.querySelector('.status-expand-icon');
                if (icon) icon.style.transform = 'rotate(90deg)';
            }
        });

        // Amber badge click → open Jira open-items popup (same as graph view)
        contentEl.querySelectorAll('.status-minor-badge[data-svc-id]').forEach(badge => {
            badge.style.cursor = 'pointer';
            badge.addEventListener('click', e => {
                e.stopPropagation();
                const id = badge.dataset.svcId;
                if (!id) return;
                const entry = healthMap.get(id);
                if (!entry?.minorIncidents?.length) return;
                const cards = entry.minorIncidents.map(inc => ({
                    key:              inc.Key,
                    summary:          inc.Summary,
                    status:           inc.Status,
                    priority:         (inc.Priority || '').split('/')[0].trim(),
                    jiraUrl:          inc.JiraUrl,
                    impactedServices: (inc.AffectedServices || '').split('||').map(s => s.trim()).filter(Boolean),
                }));
                this.app.graph._openJiraCardsPopup({ id }, 'incident', cards, e);
            });
        });

        contentEl.querySelectorAll('.status-svc-row').forEach(row => {
            row.addEventListener('click', e => {
                if (e.target.closest('.status-minor-badge')) return; // handled above
                if (e.target.closest('.status-expand-icon') && row.classList.contains('has-incidents')) {
                    this._toggleIncidents(row, healthMap);
                } else {
                    this._navigateToService(row.dataset.svcId);
                }
            });
        });

        contentEl.querySelectorAll('.status-op-toggle').forEach(btn => {
            btn.addEventListener('click', () => {
                const list = btn.nextElementSibling;
                if (!list) return;
                const expanded = !list.hidden;
                list.hidden = expanded;
                btn.setAttribute('aria-expanded', String(!expanded));
                const label = btn.querySelector('.status-op-label');
                const count = btn.dataset.count || '';
                if (label) {
                    label.textContent = expanded
                        ? `Show ${count} fully operational ▼`
                        : `Hide ${count} fully operational ▲`;
                }
            });
            btn.addEventListener('keydown', e => {
                if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); btn.click(); }
            });
        });
    }

    _onSearchUpdate() {
        if (this.statusViewEl?.style.display === 'block') this._renderContent();
    }

    _navigateToService(serviceId) {
        if (!serviceId) return;
        this._preNavSearchTerm = this.app.search.searchTerm ?? '';
        const query = `id:"${serviceId}"`;
        const input = document.getElementById('drawer-search-input');
        if (input) input.value = query;
        this.app.search.handleQuery(query, false);
        this.app.listView.toGraphView();
    }

    _toggleIncidents(row, healthMap) {
        const existing = row.nextElementSibling;
        if (existing?.classList.contains('status-incidents-pane')) {
            existing.remove();
            row.classList.remove('is-open');
            const icon = row.querySelector('.status-expand-icon');
            if (icon) icon.style.transform = '';
            return;
        }
        const svcId = row.dataset.svcId;
        const entry = healthMap.get(svcId);
        if (!entry?.incidents?.length) return;

        const pane = document.createElement('div');
        pane.className = 'status-incidents-pane';
        pane.innerHTML = entry.incidents
            .slice()
            .sort((a, b) => (ALL_RANK[a.Priority] ?? 99) - (ALL_RANK[b.Priority] ?? 99))
            .map(inc => {
                const p       = (inc.Priority || '').toLowerCase().replace('/', '').split(' ')[0];
                const summary = this._esc(this._cleanSummary(inc.Summary || ''));
                const url     = inc.JiraUrl || '';
                return `
                    <div class="status-inc-row">
                        ${url
                            ? `<a class="status-inc-key" href="${url}" target="_blank" rel="noopener">${this._esc(inc.Key)}</a>`
                            : `<span class="status-inc-key">${this._esc(inc.Key)}</span>`}
                        <div class="status-inc-meta">
                            <span class="status-inc-summary">${summary}</span>
                            <span class="domino-jira-card__priority--${p} status-inc-prio">${this._esc(inc.Priority)}</span>
                            <span class="status-inc-status">${this._esc(inc.Status)}</span>
                        </div>
                    </div>`;
            })
            .join('');

        row.insertAdjacentElement('afterend', pane);
        row.classList.add('is-open');
        const icon = row.querySelector('.status-expand-icon');
        if (icon) icon.style.transform = 'rotate(90deg)';
    }
}
