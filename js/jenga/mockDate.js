const _raw = typeof window !== 'undefined'
    ? new URLSearchParams(window.location.search).get('today')
    : null;
const _mock = _raw && /^\d{4}-\d{2}-\d{2}$/.test(_raw)
    ? new Date(_raw + 'T00:00:00')
    : null;

/** Returns a Date for "today" (midnight). Honours ?today=YYYY-MM-DD for testing. */
export function getToday() {
    if (_mock) return new Date(_mock);
    const d = new Date();
    d.setHours(0, 0, 0, 0);
    return d;
}
