/**
 * Shared utilities for StatusView (Domino) and JengaStatusView (Jenga).
 *
 * Two algorithms replace former hard-coded tables:
 *
 *   toFriendlyName(rawId)        — derives a human label from a technical slug
 *   resolveCluster(node, CLUSTERS) — assigns a service to a cluster by patterns
 *
 * Both are intentionally brand-agnostic; service-name overrides live in
 * CLUSTERS.exactNames (a small set reserved for names no pattern can catch).
 */

// ── Friendly-name algorithm ───────────────────────────────────────────────────

// Short tokens that are always rendered in ALL CAPS
const ACRONYMS = new Set([
    'api', 'spa', 'sdk', 'db', 'pdp', 'plp', 'pim', 'crm', 'idp',
    'sso', 'cms', 'otp', 'cdn', 'ui', 'ux', 'mao', 'esm', 'sap',
    'erp', 'baa', 'gde', 'b2b', 'b2c', 'jwt', 'xml', 'html', 'css',
    'url', 'uri', 'aws', 'gcp',
]);

// Trailing-only tokens that add no meaning when other words are present
const TRAILING_STRIP = new Set(['service', 'svc']);

/**
 * Convert a technical service identifier to a human-readable label.
 *
 * Examples:
 *   game-service            → Game
 *   order-fulfillment       → Order Fulfillment
 *   checkout-spa            → Checkout SPA
 *   catalog-api             → Catalog API
 *   pdp                     → PDP
 *   my-account (orders)     → My Account (Orders)
 *   Google Analytics        → Google Analytics   (already human-readable, left alone)
 *
 * The function detects whether the input is already human-readable (contains a
 * space, or uses mixed/upper case without hyphens/underscores) and skips
 * transformation in that case, so display names in the CSV pass through intact.
 */
export function toFriendlyName(rawId) {
    if (!rawId || typeof rawId !== 'string') return rawId ?? '';

    // Already human-readable: has spaces, or mixed/upper case with no slug separators
    const hasSlugSep = rawId.includes('-') || rawId.includes('_');
    const allLower   = rawId === rawId.toLowerCase();
    if (!hasSlugSep && !allLower) return rawId;

    // Handle parenthesised qualifiers: "my-account (order history)"
    const parenIdx = rawId.indexOf('(');
    if (parenIdx !== -1) {
        const main = toFriendlyName(rawId.slice(0, parenIdx).trim());
        const inner = rawId.slice(parenIdx).match(/^\(([^)]+)\)(.*)/);
        if (inner) {
            const qualifier = toFriendlyName(inner[1].trim());
            const tail = inner[2].trim();
            return tail
                ? `${main} (${qualifier}) ${toFriendlyName(tail)}`
                : `${main} (${qualifier})`;
        }
    }

    let parts = rawId.split(/[-_\s/]+/).filter(Boolean);

    // Strip trailing noise when there are still other words left
    if (parts.length > 1 && TRAILING_STRIP.has(parts[parts.length - 1].toLowerCase())) {
        parts = parts.slice(0, -1);
    }

    return parts
        .map(p => {
            const lp = p.toLowerCase();
            return ACRONYMS.has(lp)
                ? lp.toUpperCase()
                : p.charAt(0).toUpperCase() + p.slice(1).toLowerCase();
        })
        .join(' ');
}

// ── Cluster definitions ───────────────────────────────────────────────────────
//
// Each cluster uses only namePatterns and descPatterns for assignment.
// exactNames is a last-resort override for names that no pattern can catch.
// The order of clusters matters: first match wins on namePatterns, then descPatterns.

export const CLUSTERS = [
    {
        id: 'hybris',
        label: 'SAP Hybris Platform',
        namePatterns: [/hybris/i, /\bsap\b/i, /\bmysql\b/i, /\bsolr\b/i],
        descPatterns:  [/hybris/i, /\bsap commerce\b/i, /\bsap.?ecommerce\b/i],
    },
    {
        id: 'sign-in',
        label: 'Sign-In & Authentication',
        namePatterns: [/sign.?in/i, /\bidp\b/i, /\bspid\b/i, /security.?email/i],
        descPatterns:  [/authentication/i, /identity.?provider/i, /single.?sign.?on/i, /\bsso\b/i],
    },
    {
        id: 'customer-access',
        label: 'Customer Access (Okta)',
        namePatterns: [/okta/i],
        descPatterns:  [/okta/i, /customer.?identity/i, /workforce.?identity/i],
    },
    {
        id: 'checkout',
        label: 'Checkout',
        namePatterns: [/checkout/i],
        descPatterns:  [/checkout/i, /order.?completion/i, /purchase.?flow/i],
    },
    {
        id: 'cart',
        label: 'Cart & Buy Area',
        namePatterns: [/\bcart\b/i, /minicart/i, /buy.?area/i],
        descPatterns:  [/shopping.?cart/i, /add to cart/i],
    },
    {
        id: 'payments',
        label: 'Payments',
        namePatterns: [
            /payment/i, /\bpay\b/i, /adyen/i, /paypal/i, /klarna/i,
            /cybersource/i, /affirm/i, /commerce.?layer/i, /riskified/i,
            /\bvertex\b/i, /\btoss\b/i, /\btax\b/i,
            /distance.?sales/i, /address.?valid/i,
        ],
        descPatterns: [/payment.?gateway/i, /\bpsp\b/i, /payment.?processing/i, /billing/i],
    },
    {
        id: 'orders',
        label: 'Orders',
        namePatterns: [
            /^order/i,
            /order.?(api|confirm|fulfil|valid|number|cronjob)/i,
            /fulfil/i, /thank.?you/i,
            /collect.?in.?store/i, /in.?store.?purchase/i,
            /gifting/i, /\bshipping\b/i,
        ],
        descPatterns: [/order.?management/i, /post.?purchase/i, /order.?lifecycle/i],
    },
    {
        id: 'pricing',
        label: 'Pricing',
        namePatterns: [/^price/i],
        descPatterns:  [/pricing/i, /price.?management/i, /price.?engine/i],
    },
    {
        id: 'stock',
        label: 'Stock & Inventory',
        namePatterns: [/stock/i, /inventory/i],
        descPatterns:  [/stock.?management/i, /inventory/i, /warehouse/i],
    },
    {
        id: 'pim',
        label: 'Product Catalog & PIM',
        namePatterns: [
            /\bpim\b/i, /^pim/i, /\bakeneo\b/i,
            /^catalog/i, /catalog.?(api|worker)/i,
            /product.?service/i, /material.?compos/i,
        ],
        descPatterns: [/product.?information.?management/i, /\bpim\b/i, /catalog.?management/i],
    },
    {
        id: 'product-pages',
        label: 'Product Pages (PDP / PLP)',
        namePatterns: [/^pdp$/i, /^plp$/i, /add.?to.?cart/i, /iframe.*cart/i, /product.?manag/i],
        descPatterns:  [/product.?detail/i, /product.?listing/i, /\bpdp\b/i, /\bplp\b/i],
    },
    {
        id: 'search',
        label: 'Search & Discovery',
        namePatterns: [
            /\bsearch\b/i, /\balgolia\b/i, /\byasa\b/i, /cataloger/i,
            /visual.?merch/i, /recommend/i, /\breco\b/i, /personaliz/i,
            /find.?in.?store/i, /feature.?toggle/i,
        ],
        descPatterns: [/\bsearch\b/i, /product.?discovery/i, /merchandis/i, /personaliz/i, /recommendation/i],
    },
    {
        id: 'my-account',
        label: 'My Account & Customer Profile',
        namePatterns: [
            /my.?account/i, /my.?gucci/i,
            /customer.?profile/i, /cust.?anon/i,
            /wishlist/i, /\bcrm\b/i,
            /dynamo.?db/i, /kanonymous/i, /client.?service/i,
        ],
        descPatterns: [
            /customer.?account/i, /user.?profile/i, /loyalty/i, /\bcrm\b/i,
            /customer.?interaction/i,
        ],
    },
    {
        id: 'store-locator',
        label: 'Store Locator & Appointments',
        namePatterns: [
            /store.?locator/i, /book.?appoint/i, /\bbaa\b/i,
            /store.?queue/i, /store.?db/i,
            /google.?maps/i, /powerfront/i,
        ],
        descPatterns: [/store.?location/i, /appointment/i, /clienteling/i],
    },
    {
        id: 'email',
        label: 'Email & Notifications',
        namePatterns: [/email/i, /\bnewsletter\b/i, /\botp\b/i],
        descPatterns:  [/transactional.?email/i, /email.?notification/i, /notification.?service/i],
    },
    {
        id: 'content',
        label: 'Content & CMS',
        namePatterns: [
            /contentful/i, /editorial/i, /brightcove/i, /optimizely/i,
            /microsite/i, /collection.*page/i, /capsule.*page/i,
        ],
        descPatterns: [/content.?management/i, /\bcms\b/i, /editorial/i, /digital.?content/i],
    },
    {
        id: 'images',
        label: 'Image & Media Serving',
        namePatterns: [/image.?server/i, /img.?viewer/i],
        descPatterns:  [/image.?serv/i, /media.?cdn/i, /image.?optim/i],
    },
    {
        id: 'analytics',
        label: 'Analytics & Tracking',
        namePatterns: [/analytics/i, /tag.?manager/i, /wunderkind/i, /myfeelback/i, /one.?trust/i, /\bbeacon\b/i, /tracking/i],
        descPatterns:  [/analytics/i, /tracking/i, /tag.?management/i, /consent/i],
    },
    {
        id: 'platform-infra',
        label: 'Platform Infrastructure',
        namePatterns: [
            /akamai/i, /bifrost/i, /esm.?proxy/i, /orchestrat/i,
            /feature.?flag/i, /backoffice/i, /geo.?markup/i,
            /service.?pages/i, /choose.?your.?country/i,
        ],
        descPatterns: [/\bcdn\b/i, /infrastructure/i, /proxy/i, /feature.?toggle/i, /server.?side/i],
    },
    {
        id: 'contact-support',
        label: 'Contact & Support',
        namePatterns: [/contact.?us/i, /live.?chat/i, /support.?widget/i],
        descPatterns:  [/customer.?support/i, /contact.?form/i, /live.?chat/i],
    },
];

export const CLUSTER_BY_ID = Object.fromEntries(CLUSTERS.map(c => [c.id, c]));

// ── Cluster resolution ────────────────────────────────────────────────────────

/**
 * Return the cluster id that best fits `node`, or null if none match.
 *
 * Resolution order:
 *   1. exactNames (small per-cluster override set, only for unguessable names)
 *   2. namePatterns  tested against node.id (lowercased)
 *   3. descPatterns  tested against node.Description (lowercased)
 *
 * The first matching cluster wins — order in CLUSTERS matters.
 */
export function resolveCluster(node, clusters) {
    const id       = node.id || node['Service Name'] || '';
    const nameLower = id.toLowerCase();
    const descLower = (node.Description || '').toLowerCase();

    for (const c of clusters) {
        if (c.exactNames?.has(id)) return c.id;
    }
    for (const c of clusters) {
        for (const p of (c.namePatterns ?? [])) {
            if (p.test(nameLower)) return c.id;
        }
    }
    for (const c of clusters) {
        for (const p of (c.descPatterns ?? [])) {
            if (p.test(descLower)) return c.id;
        }
    }
    return null;
}

/**
 * @deprecated No longer needed — resolveCluster uses patterns directly.
 * Kept for any external callers; returns an empty Map.
 */
export function buildClusterIndex(_clusters) {
    return new Map();
}
