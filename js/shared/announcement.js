// type: 'changelog' (default) shows "What's new: [title]" with 🚀 and an expandable item list.
// type: 'announcement' shows "📢 [title]" with no expand button — for plain one-liner notices.
// items is optional; omit it (or leave empty) for a title-only announcement.
export const ANNOUNCEMENTS = [
    {
        id: '2026-08-31',
        type: 'changelog',
        title: 'Post-it notes, search & Monopoli',
        items: [
            'Post-it notes on all apps — multiple notes, editable titles, real-time URL links',
            'Custom autocomplete dropdown with compound & multi-field queries (Domino)',
            'Per-clause chip editing and full keyboard navigation in search bars',
            'Domino list view: Excel export with visible-columns / all-columns choice',
            'Jenga: major incident EVA alert popup, ITSM retention band on timeline',
            'Monopoli: new Jira ticket viewer with timeline and dependency graph',
        ],
    },
];
