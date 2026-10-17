const SUPABASE_URL = process.env.SUPABASE_URL || '';
const SUPABASE_ANON_KEY = process.env.SUPABASE_ANON_KEY || '';
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env['supabase service role key'] || '';
const ENTRY_OPEN_DATE = process.env.ENTRY_OPEN_DATE || '2026-10-15T00:00:00+05:30';
const ENTRY_DEADLINE = process.env.ENTRY_DEADLINE || '2026-10-30T23:59:59+05:30';

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

module.exports = {
    SUPABASE_URL,
    SUPABASE_ANON_KEY,
    ENTRY_OPEN_DATE,
    ENTRY_DEADLINE,
    fetchSupabase,
    validateAthleteEntry
};
