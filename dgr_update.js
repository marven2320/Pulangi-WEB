const mysql = require('mysql2/promise');
const cron = require('node-cron');
const path = require('path');
const fs = require('fs');

const dbConfig = {
    host: 'localhost',
    user: 'root',
    password: 'KaRMSys2025!',
    database: 'pulangi_data',
    connectionLimit: 10
};

const OUTPUT_DIR = path.join(__dirname, 'rawdata');
// Unique buffer file for Shift Data to avoid conflict with Hourly Data
const BUFFER_FILE = path.join(OUTPUT_DIR, 'data_buffer_shift.json');

if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR);

// --- HELPER FUNCTIONS ---

const formatSQLDate = (date) => date.toLocaleDateString('en-CA');
const formatSQLTime = (date) => {
    const hours = date.getHours();
    return (hours < 12) ? "00:00:00" : "12:00:00";
};

const toInteger = (val) => {
    if (val === null || val === undefined || val === '') return 0;
    return Math.round(Number(val));
};

function getSheetContext(date) {
    const day = date.getDate();
    const hour = date.getHours();
    const month = date.getMonth();
    const year = date.getFullYear();

    // Cycle switches to next month on the 26th at 12:00 PM
    const isNextCycle = (day > 26) || (day === 26 && hour >= 12);

    if (isNextCycle) {
        if (month === 11) return { sheetIndex: 0, targetYear: year + 1 };
        return { sheetIndex: month + 1, targetYear: year };
    } else {
        return { sheetIndex: month, targetYear: year };
    }
}

// --- BUFFER MANAGEMENT ---

function loadBuffer() {
    if (fs.existsSync(BUFFER_FILE)) {
        try {
            return JSON.parse(fs.readFileSync(BUFFER_FILE));
        } catch (e) {
            console.error("[Logger] Error reading buffer, starting fresh.", e);
            return [];
        }
    }
    return [];
}

function saveBuffer(data) {
    fs.writeFileSync(BUFFER_FILE, JSON.stringify(data, null, 2));
}

// --- THE LOGGING TASK ---

async function logShiftData() {
    const now = new Date();
    // Ensure we are targeting the exact hour (0 or 12)
    // In production cron, this runs at 00:01 or 12:01.
    // If running manually at 12:05, this snaps it back to 12:00.
    if (now.getHours() < 12) now.setHours(0, 0, 0, 0);
    else now.setHours(12, 0, 0, 0);

    const sqlDateStr = formatSQLDate(now);
    const sqlTimeStr = formatSQLTime(now);

    console.log(`\n[Shift Logger] Fetching DB Data for: ${sqlDateStr} ${sqlTimeStr}`);

    let connection;

    try {
        // 1. Fetch from DB
        const query = `
            SELECT 
                energy1, energy2, energy3
            FROM pulangi 
            WHERE date = ? AND time = ? 
            LIMIT 1
        `;

        connection = await mysql.createConnection(dbConfig);
        const [rows] = await connection.execute(query, [sqlDateStr, sqlTimeStr]);

        const rawValues = rows.length > 0 ? [
            rows[0].energy1, rows[0].energy2, rows[0].energy3
        ] : new Array(3).fill('');

        const formattedValues = rawValues.map(val => toInteger(val));

        // 2. Calculate Destination Context
        const { sheetIndex, targetYear } = getSheetContext(now);
        const fileName = `Pulangi IV HEP - Generation Data - RAW_${targetYear}.xlsx`;
        const filePath = path.join(OUTPUT_DIR, fileName);

        // Calculate Row
        const cycleStartDate = new Date(targetYear, sheetIndex - 1, 26, 12, 0, 0);
        const diffMs = now - cycleStartDate;
        const diffShifts = Math.floor(diffMs / (1000 * 60 * 60 * 12));

        if (diffShifts < 0) {
            console.warn("[Shift Logger] Time is before cycle start. Skipping.");
            return;
        }

        // Row 5 is the starting row for 1st shift
        const targetRow = 5 + diffShifts;

        // 3. Save to Buffer
        const buffer = loadBuffer();

        const newEntry = {
            id: Date.now(), // Unique ID
            targetYear,
            filePath,
            sheetIndex,
            targetRow,
            values: formattedValues,
            timestamp: new Date().toISOString()
        };

        // Update existing entry if same slot, otherwise push new
        const uniqueBuffer = buffer.filter(item =>
            !(item.filePath === filePath && item.sheetIndex === sheetIndex && item.targetRow === targetRow)
        );

        uniqueBuffer.push(newEntry);
        saveBuffer(uniqueBuffer);

        console.log(`[Shift Logger] Data buffered. Total pending entries: ${uniqueBuffer.length}`);

    } catch (error) {
        console.error('[Shift Logger] Error:', error);
    } finally {
        if (connection) await connection.end();
    }
}

// Run immediately for testing/startup
logShiftData();

// Schedule: 00:01 and 12:01 daily
cron.schedule('1 0,12 * * *', () => logShiftData());
