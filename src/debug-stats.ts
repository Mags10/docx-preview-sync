export const debugStats = (function () {
    const counters: Record<string, number> = {};

    function inc(name: string, by: number = 1) {
        counters[name] = (counters[name] || 0) + by;
        // keep non-enumerable console helper for quick debugging
        return counters[name];
    }

    function dec(name: string, by: number = 1) {
        counters[name] = (counters[name] || 0) - by;
        if (counters[name] <= 0) counters[name] = 0;
        return counters[name];
    }

    function get(name: string) {
        return counters[name] || 0;
    }

    function snapshot() {
        return { ...counters };
    }

    function reset() {
        Object.keys(counters).forEach(k => counters[k] = 0);
    }

    return {
        inc,
        dec,
        get,
        snapshot,
        reset,
    } as const;
})();

export default debugStats;
