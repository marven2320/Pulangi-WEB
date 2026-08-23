#!/usr/bin/env node
// Standalone CLI: fetches the exact same generation data viewdata.html
// displays (Power/Frequency per unit, over a date/time range), by querying
// the DB directly - no HTTP server involved. Useful for:
//   - verifying DB connectivity and the `pulangi` table's schema on its own
//   - spot-checking what viewdata.html's dual-axis charts will show for a
//     given date/time range, without going through a browser
//   - debugging data issues independent of server.js/the web layer
//
// Reuses the exact same connection config and query as server.js's
// GET /api/generation-data endpoint.
//
// Usage:
//   node fetch-generation-data.js --date 2025-01-15 --start 09:00:00 --end 10:00:00
//   node fetch-generation-data.js --date 2025-01-15 --start 09:00:00 --end 10:00:00 --out today.json
//   node fetch-generation-data.js                     # defaults to today, full day
//
// Flags:
//   --date   YYYY-MM-DD   (default: today)
//   --start  HH:MM:SS     (default: 00:00:00)
//   --end    HH:MM:SS     (default: 23:59:59)
//   --out    <path>       write JSON to this file instead of only printing it

const fs = require('fs');
const mysql = require('mysql2');

// Same query logic viewdata.html actually uses live, via server.js's
// GET /api/generation-data - see generation-data-query.js.
const { getGenerationData } = require('./generation-data-query');

// --- Same DB connection info as server.js ---
const pool = mysql.createPool({
    host: 'localhost',
    user: 'root',
    password: 'KaRMSys2025!',
    database: 'pulangi_data',
    connectionLimit: 10
});

// --- CLI argument parsing (--flag value) ---
function parseArgs(argv) {
    const args = {};
    for (let i = 0; i < argv.length; i++) {
        const a = argv[i];
        if (a.startsWith('--')) {
            const key = a.slice(2);
            const val = argv[i + 1] && !argv[i + 1].startsWith('--') ? argv[++i] : true;
            args[key] = val;
        }
    }
    return args;
}

function todayDateStr() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

const args = parseArgs(process.argv.slice(2));
const date = args.date || todayDateStr();
const start = args.start || '00:00:00';
const end = args.end || '23:59:59';
const outPath = args.out || null;

console.log(`Querying pulangi_data.pulangi for date=${date}, time ${start}..${end} ...`);

getGenerationData(pool, { date, start, end }, (err, rows) => {
    if (err) {
        if (/date must be/.test(err.message)) {
            console.error('Usage: node fetch-generation-data.js --date YYYY-MM-DD --start HH:MM:SS --end HH:MM:SS [--out file.json]');
            console.error(`Got: date=${date} start=${start} end=${end}`);
        } else {
            console.error('Database query failed:', err.message);
            console.error('If this mentions an unknown column (mw1/freq1/etc.), the `pulangi` table\'s');
            console.error('real column names differ from what server.js/this script assume - check the');
            console.error('table schema (e.g. `DESCRIBE pulangi;`) and update both files to match.');
        }
        pool.end();
        process.exit(1);
    }

    console.log(`Found ${rows.length} row(s).`);

    const json = JSON.stringify({ rows }, null, 2);

    if (outPath) {
        fs.writeFileSync(outPath, json);
        console.log(`Wrote result to ${outPath}`);
    } else {
        console.log(json);
    }

    pool.end();
});
