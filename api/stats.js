const { fetchSupabase, ENTRY_OPEN_DATE, ENTRY_DEADLINE } = require('./_utils');

module.exports = async (req, res) => {
    if (req.method !== 'GET') {
        res.writeHead(405, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({ error: 'Method not allowed' }));
    }

    try {
        const resp = await fetchSupabase('entries?select=*');
        const entries = resp.ok ? await resp.json() : [];

        const u14Count = entries.filter(e => e.category === 'Under 14').length;
        const u16EventCounts = {};
        entries.filter(e => e.category === 'Under 16').forEach(e => {
            const key = `${e.gender || 'Unknown'}|${(e.events && e.events[0]) || ''}`;
            u16EventCounts[key] = (u16EventCounts[key] || 0) + 1;
        });

        const stats = {
            totalAthletes: entries.length,
            maxAthletes: 13,
            u14Count,
            maxU14: 4,
            u14SpotsRemaining: Math.max(0, 4 - u14Count),
            totalSpotsRemaining: Math.max(0, 13 - entries.length),
            u16EventCounts,
            entryOpenDate: ENTRY_OPEN_DATE,
            entryDeadline: ENTRY_DEADLINE
        };

        res.writeHead(200, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify(stats));
    } catch (err) {
        console.error('Stats error:', err);
        res.writeHead(500, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({ error: 'Failed to retrieve stats' }));
    }
};
