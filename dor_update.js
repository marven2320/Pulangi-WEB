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
const BUFFER_FILE = path.join(OUTPUT_DIR, 'data_buffer.json');

// Ensure output directory exists
if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR);
}

// --- HELPER FUNCTIONS ---

const formatDate = (date) => date.toLocaleDateString('en-CA');
const formatQueryTime = (date) => {
    const hh = String(date.getHours()).padStart(2, '0');
    return `${hh}:00:00`;
};

const toFixedNumber = (val) => {
    if (val === null || val === undefined || val === '') return '';
    return Number(Number(val).toFixed(2));
};

function getSheetContext(date) {
    const day = date.getDate();
    const month = date.getMonth();
    const year = date.getFullYear();

    if (day >= 26) {
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
            console.error("Error reading buffer, starting fresh.", e);
            return [];
        }
    }
    return [];
}

function saveBuffer(data) {
    fs.writeFileSync(BUFFER_FILE, JSON.stringify(data, null, 2));
}

// --- THE LOGGING TASK ---

async function logToBuffer() {
    const now = new Date();
    const currentDateStr = formatDate(now);
    const currentTimeStr = formatQueryTime(now);

    console.log(`\n[Logger] Fetching DB Data for: ${currentDateStr} ${currentTimeStr}`);

    let connection;

    try {
        // 1. Fetch from DB
        const query = `
            SELECT 
                mw1, (vab1+vbc1+vca1)/3 as v1, (pfa1+pfb1+pfc1)/3 as pf1, mvar1, 
                mw2, (vab2+vbc2+vca2)/3 as v2, (pfa2+pfb2+pfc2)/3 as pf2, mvar2, 
                mw3, (vab3+vbc3+vca3)/3 as v3, (pfa3+pfb3+pfc3)/3 as pf3, mvar3,
                mw, mvar
            FROM pulangi 
            WHERE date = ? AND time = ? 
            LIMIT 1
        `;

        connection = await mysql.createConnection(dbConfig);
        const [rows] = await connection.execute(query, [currentDateStr, currentTimeStr]);

        const rawValues = rows.length > 0 ? [
            rows[0].mw1, rows[0].v1, rows[0].pf1, rows[0].mvar1,
            rows[0].mw2, rows[0].v2, rows[0].pf2, rows[0].mvar2,
            rows[0].mw3, rows[0].v3, rows[0].pf3, rows[0].mvar3,
            rows[0].mw, rows[0].mvar
        ] : new Array(14).fill('');

        const formattedValues = rawValues.map(val => toFixedNumber(val));

        // 2. Calculate Destination Context
        const { sheetIndex, targetYear } = getSheetContext(now);
        const fileName = `Pulangi IV HEP - Operational Highlights - RAW_${targetYear}.xlsx`;
        const filePath = path.join(OUTPUT_DIR, fileName);

        const cycleStartDate = new Date(targetYear, sheetIndex - 1, 26, 0, 0, 0);
        const diffMs = now - cycleStartDate;
        const diffHours = Math.floor(diffMs / (1000 * 60 * 60));

        if (diffHours < 0) {
            console.warn("[Logger] Time is before cycle start. Skipping.");
            return;
        }
        const targetRow = 4 + diffHours;

        // 3. Save to Buffer
        const buffer = loadBuffer();

        const newEntry = {
            id: Date.now(), // Unique ID for tracking in merge script
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

        console.log(`[Logger] Data buffered. Total pending entries: ${uniqueBuffer.length}`);

    } catch (error) {
        console.error('[Logger] Error:', error);
    } finally {
        if (connection) await connection.end();
    }
}

// Run immediately on start
logToBuffer();

// Schedule: Minute 1 of every hour
cron.schedule('1 * * * *', () => logToBuffer());
