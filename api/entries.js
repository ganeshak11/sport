const { fetchSupabase, validateAthleteEntry } = require('./_utils');

module.exports = async (req, res) => {
    if (req.method !== 'POST') {
        res.writeHead(405, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({ error: 'Method not allowed' }));
    }

    try {
        let body = req.body;
        if (typeof body === 'string') {
            body = JSON.parse(body);
        } else if (!body) {
            let raw = '';
            for await (const chunk of req) raw += chunk;
            body = JSON.parse(raw || '{}');
        }

        const data = body || {};

        // Fetch current entries for real-time quota verification
        const listResp = await fetchSupabase('entries?select=*');
        const currentEntries = listResp.ok ? await listResp.json() : [];

        // Strict validation against AFI Circular No. 88
        const validation = validateAthleteEntry(data, currentEntries);
        if (!validation.valid) {
            res.writeHead(400, { 'Content-Type': 'application/json' });
            return res.end(JSON.stringify({ error: validation.error }));
        }

        // Prepare record matching database schema
        const afiTrimmed = (data.afi_uid || '').trim().toUpperCase();
        const afiFormatted = afiTrimmed
            ? `AFI UID: ${afiTrimmed}`
            : 'AFI UID: None (District Meet Only - Not Eligible for NIDJAM)';

        const coachFormatted = [
            afiFormatted,
            data.coach ? `Coach: ${data.coach.trim()}` : null,
            data.declaration ? `Parent Undertaking: Verified` : null
        ].filter(Boolean).join(' | ');

        const record = {
            name: data.name.trim(),
            dob: data.dob,
            father: (data.father || '').trim(),
            mother: (data.mother || '').trim(),
            aadhaar: (data.aadhaar || '').replace(/\s+/g, ''),
            school: (data.school || '').trim(),
            phone: (data.phone || '').trim(),
            gender: data.gender,
            address: (data.address || '').trim(),
            category: data.category,
            coach: coachFormatted,
            events: data.events
        };

        const insertResp = await fetchSupabase('entries', {
            method: 'POST',
            headers: { 'Prefer': 'return=representation' },
            body: JSON.stringify(record)
        });

        if (!insertResp.ok) {
            const errData = await insertResp.json().catch(() => ({}));
            console.error('Supabase insert failed:', insertResp.status, errData);
            res.writeHead(500, { 'Content-Type': 'application/json' });
            return res.end(JSON.stringify({
                error: errData.message || 'Database error occurred while recording entry.'
            }));
        }

        const inserted = await insertResp.json();
        res.writeHead(201, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({
            success: true,
            message: 'Athlete entry registered successfully for Mysore District Athletic Association District Selection Meet 2026.',
            entry: inserted[0]
        }));
    } catch (err) {
        console.error('API submission error:', err);
        res.writeHead(400, { 'Content-Type': 'application/json' });
        return res.end(JSON.stringify({ error: err.message || 'Invalid request body' }));
    }
};
