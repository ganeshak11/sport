require('dotenv').config();

const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');

const PORT = process.env.PORT || 3000;
const SUPABASE_URL = process.env.SUPABASE_URL || '';
const SUPABASE_ANON_KEY = process.env.SUPABASE_ANON_KEY || '';
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env['supabase service role key'] || '';
const ADMIN_EMAIL = process.env.ADMIN_EMAIL || 'admin@mysuru-athletics.com';
const ENTRY_OPEN_DATE = process.env.ENTRY_OPEN_DATE || '2026-10-15T00:00:00+05:30';
const ENTRY_DEADLINE = process.env.ENTRY_DEADLINE || '2026-10-30T23:59:59+05:30';

if (!SUPABASE_URL || !SUPABASE_ANON_KEY) {
    console.error('❌ Missing SUPABASE_URL or SUPABASE_ANON_KEY in environment');
    process.exit(1);
}

const MIME_TYPES = {
    '.html': 'text/html; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.svg': 'image/svg+xml',
    '.ico': 'image/x-icon',
    '.txt': 'text/plain; charset=utf-8'
};

// Official Rules Constants (Circular No. 88/09-2026)
const U14_MIN_DOB = new Date('2012-11-17T00:00:00+05:30');
const U14_MAX_DOB = new Date('2014-11-16T23:59:59+05:30');
const U16_MIN_DOB = new Date('2010-11-17T00:00:00+05:30');
const U16_MAX_DOB = new Date('2012-11-16T23:59:59+05:30');

const U14_TRIATHLONS = [
    'Tri A',
    'Tri B',
    'Tri C',
    'TriA',
    'TriB',
    'TriC',
    'Triathlon A',
    'Triathlon B',
    'Triathlon C'
];

const U16_ALLOWED_EVENTS = {
    Boy: [
        '60m',
        '80m Hurdles (0.838m)',
        '600m',
        'High Jump (Scissor Only)',
        'Long Jump (5m Approach)',
        '4Kg Shot Put (Standing)',
        '600gms Javelin Throw (10m Approach)',
        'Pentathlon (60m · 80mH · LJ · Shot Put · 600m)'
    ],
    Girl: [
        '60m',
        '80m Hurdles (0.762m)',
        '600m',
        'High Jump (Scissor Only)',
        'Long Jump (5m Approach)',
        '3Kg Shot Put (Standing)',
        '500gms Javelin Throw (10m Approach)',
        'Pentathlon (60m · 80mH · LJ · Shot Put · 600m)'
    ]
};

// Helper: Supabase REST with service role key
async function fetchSupabase(endpoint, options = {}) {
    const key = SUPABASE_SERVICE_ROLE_KEY || SUPABASE_ANON_KEY;
    const headers = {
        'apikey': key,
        'Authorization': `Bearer ${key}`,
        'Content-Type': 'application/json',
        ...(options.headers || {})
    };
    return fetch(`${SUPABASE_URL}/rest/v1/${endpoint}`, {
        ...options,
        headers
    });
}

// Strict Validator for Circular No. 88/09-2026
function validateAthleteEntry(data, currentEntries = []) {
    const now = new Date();
    const deadline = new Date(ENTRY_DEADLINE);

    if (now > deadline) {
        return { valid: false, error: 'Online entries closed on 30 October 2026 as per AFI Circular No. 88/09-2026.' };
    }

    if (!data.name || !data.name.trim()) {
        return { valid: false, error: 'Athlete full name is mandatory (must match Certificate / Official ID).' };
    }

    // AFI UID Policy per MDAA:
    // AFI UID is compulsory for NIDJAM selection.
    // Athletes without AFI UID can participate in this District Meet and win Certificate & Medal,
    // but will NOT be allowed to take part in NIDJAM Meet.
    const rawAfi = (data.afi_uid || '').trim();
    if (rawAfi) {
        const cleanAfiUid = rawAfi.toUpperCase();
        const isDuplicateAfi = currentEntries.some(e => {
            const coachStr = e.coach || '';
            return coachStr.toUpperCase().includes(`AFI UID: ${cleanAfiUid}`) ||
                   coachStr.toUpperCase().includes(cleanAfiUid);
        });
        if (isDuplicateAfi) {
            return { valid: false, error: 'An entry with this AFI UID has already been registered.' };
        }
    }

    if (!data.dob) {
        return { valid: false, error: 'Date of birth is mandatory.' };
    }

    const dob = new Date(data.dob + 'T00:00:00+05:30');
    if (isNaN(dob.getTime())) {
        return { valid: false, error: 'Invalid date of birth format.' };
    }

    if (!['Under 14', 'Under 16'].includes(data.category)) {
        return { valid: false, error: 'Valid age category (Under 14 or Under 16) is required.' };
    }

    if (!['Boy', 'Girl'].includes(data.gender)) {
        return { valid: false, error: 'Valid gender (Boy or Girl) is required.' };
    }

    // Age Cut-off Enforcement
    if (data.category === 'Under 14') {
        if (dob < U14_MIN_DOB || dob > U14_MAX_DOB) {
            return {
                valid: false,
                error: 'U-14 athletes must be born between 17/11/2012 & 16/11/2014 (minimum 12 years as on 16 Nov 2026).'
            };
        }
    } else if (data.category === 'Under 16') {
        if (dob < U16_MIN_DOB || dob > U16_MAX_DOB) {
            return {
                valid: false,
                error: 'U-16 athletes must be born between 17/11/2010 & 16/11/2012 as per AFI Circular No. 88/09-2026.'
            };
        }
    }

    // Aadhaar Verification (12 digits)
    const cleanAadhaar = (data.aadhaar || '').replace(/\s+/g, '');
    if (!/^\d{12}$/.test(cleanAadhaar)) {
        return { valid: false, error: 'Aadhaar must be a valid 12-digit numeric number without spaces.' };
    }

    // Phone Verification (10 digits)
    const cleanPhone = (data.phone || '').replace(/\D+/g, '');
    if (cleanPhone.length < 10) {
        return { valid: false, error: 'Parent / guardian contact number must be at least 10 digits.' };
    }

    // Mandatory Parent Declaration
    if (!data.declaration) {
        return { valid: false, error: 'Parent/Guardian undertaking and declaration consent is mandatory.' };
    }

    // Events Validation
    const events = Array.isArray(data.events) ? data.events : [];
    if (!events.length) {
        return { valid: false, error: 'Event selection is required.' };
    }

    if (data.category === 'Under 14') {
        const triathlon = events.find(e => U14_TRIATHLONS.includes(e));
        const hasKidsJavelin = events.includes('Kids Javelin Throw');
        if (!triathlon) {
            return { valid: false, error: 'U-14 athletes must choose exactly one Triathlon group (Tri A, Tri B, or Tri C).' };
        }
        if (!hasKidsJavelin) {
            return { valid: false, error: 'Kids Javelin Throw is mandatory for all U-14 athletes.' };
        }
        if (events.length !== 2) {
            return { valid: false, error: 'U-14 athletes may only enter one Triathlon group (Tri A, Tri B, or Tri C) and Kids Javelin Throw.' };
        }
    } else if (data.category === 'Under 16') {
        if (events.length !== 1) {
            return { valid: false, error: 'U-16 athletes may participate in only ONE event as per Circular No. 88.' };
        }
        const allowedList = U16_ALLOWED_EVENTS[data.gender] || [];
        if (!allowedList.includes(events[0])) {
            return { valid: false, error: `Invalid event "${events[0]}" for U-16 ${data.gender}.` };
        }
    }

    // Duplicate Check by Aadhaar (same athlete cannot register twice)
    const isDuplicateAadhaar = currentEntries.some(e => e.aadhaar === cleanAadhaar);
    if (isDuplicateAadhaar) {
        return { valid: false, error: 'An entry with this Aadhaar number has already been registered.' };
    }

    return { valid: true };
}

const server = http.createServer(async (req, res) => {
    const parsedUrl = url.parse(req.url, true);
    const pathname = parsedUrl.pathname;

    // 1. Dynamic /config.js
    if (pathname === '/config.js') {
        res.writeHead(200, {
            'Content-Type': 'text/javascript; charset=utf-8',
            'Cache-Control': 'no-store'
        });
        return res.end(
            `window.__ENV__ = {\n` +
            `  SUPABASE_URL: ${JSON.stringify(SUPABASE_URL)},\n` +
            `  SUPABASE_ANON_KEY: ${JSON.stringify(SUPABASE_ANON_KEY)},\n` +
            `  ADMIN_EMAIL: ${JSON.stringify(ADMIN_EMAIL)},\n` +
            `  ENTRY_OPEN_DATE: ${JSON.stringify(ENTRY_OPEN_DATE)},\n` +
            `  ENTRY_DEADLINE: ${JSON.stringify(ENTRY_DEADLINE)}\n` +
            `};\n`
        );
    }

    // 2. GET /api/stats - Live Quota Tracker
    if (pathname === '/api/stats' && req.method === 'GET') {
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
    }

    // 3. POST /api/entries - Strict Entry Registration Endpoint
    if (pathname === '/api/entries' && req.method === 'POST') {
        let body = '';
        req.on('data', chunk => { body += chunk; });
        req.on('end', async () => {
            try {
                const data = JSON.parse(body || '{}');

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
        });
        return;
    }

    // 4. Block access to sensitive files
    const blocked = ['.env', '.env.example', '.gitignore', 'server.js', 'package.json', 'package-lock.json'];
    if (blocked.includes(path.basename(pathname))) {
        res.writeHead(403, { 'Content-Type': 'text/plain' });
        return res.end('403 Forbidden');
    }

    // 5. Static file serving
    let safePath = pathname === '/' ? '/index.html' : pathname;
    const filePath = path.join(__dirname, safePath);

    if (!filePath.startsWith(__dirname + path.sep)) {
        res.writeHead(403, { 'Content-Type': 'text/plain' });
        return res.end('403 Forbidden');
    }

    fs.stat(filePath, (err, stats) => {
        if (err || !stats.isFile()) {
            res.writeHead(404, { 'Content-Type': 'text/plain' });
            return res.end('404 Not Found');
        }
        const ext = path.extname(filePath).toLowerCase();
        const contentType = MIME_TYPES[ext] || 'application/octet-stream';
        res.writeHead(200, { 'Content-Type': contentType });
        fs.createReadStream(filePath).pipe(res);
    });
});

server.listen(PORT, () => {
    console.log(`✅ Mysuru Athletics Entry Portal running at http://localhost:${PORT}`);
    console.log(`   Official Meet: 20th NIDJAM 2026 (Circular No. 88/09-2026)`);
    console.log(`   Entry Window: ${ENTRY_OPEN_DATE} to ${ENTRY_DEADLINE}`);
});