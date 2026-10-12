module.exports = (req, res) => {
    const SUPABASE_URL = process.env.SUPABASE_URL || '';
    const SUPABASE_ANON_KEY = process.env.SUPABASE_ANON_KEY || '';
    const ADMIN_EMAIL = process.env.ADMIN_EMAIL || 'admin@mysuru-athletics.com';
    const ENTRY_OPEN_DATE = process.env.ENTRY_OPEN_DATE || '2026-10-15T00:00:00+05:30';
    const ENTRY_DEADLINE = process.env.ENTRY_DEADLINE || '2026-10-30T23:59:59+05:30';

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
};
