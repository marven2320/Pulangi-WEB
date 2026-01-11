const mysql = require('mysql2/promise');
const cron = require('node-cron');
const path = require('path');
const fs = require('fs');
const XlsxPopulate = require('xlsx-populate');

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

// Helper to get time string "HH:MM:SS"
const formatTimeWithSeconds = (date, secondsOffset = 0) => {
    const d = new Date(date);
    d.setSeconds(d.getSeconds() + secondsOffset);
    return d.toTimeString().split(' ')[0]; // Returns "HH:MM:SS"
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

// --- DB FETCHING HELPER (With Retry Logic) ---
// Fetches a single hour of data. If null, retries by incrementing seconds.
async function fetchDataForDate(connection, dateObj) {
    const dateStr = formatDate(dateObj);

    // We try from 0 to 59 seconds
    const MAX_RETRIES = 60;

    for (let sec = 0; sec < MAX_RETRIES; sec++) {
        const timeStr = formatTimeWithSeconds(dateObj, sec);

        // Only log the retry if it's not the first attempt
        if (sec > 0) process.stdout.write(`.`);

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

        const [rows] = await connection.execute(query, [dateStr, timeStr]);

        if (rows.length > 0) {
            if (sec > 0) console.log(` Found at ${timeStr}`); // Log success after retry

            const rawValues = [
                rows[0].mw1, rows[0].v1, rows[0].pf1, rows[0].mvar1,
                rows[0].mw2, rows[0].v2, rows[0].pf2, rows[0].mvar2,
                rows[0].mw3, rows[0].v3, rows[0].pf3, rows[0].mvar3,
                rows[0].mw, rows[0].mvar
            ];
            return rawValues.map(val => toFixedNumber(val));
        }
    }

    // If loop finishes without finding data
    if (MAX_RETRIES > 1) console.log(" No data found within search window.");
    return null;
}

// --- THE LOGGING TASK ---

async function logToBuffer() {
    const now = new Date();
    // Round down to current hour 00:00:00
    now.setMinutes(0, 0, 0);

    const currentDateStr = formatDate(now);
    // Initial display only shows the hour base
    const displayTime = formatTimeWithSeconds(now, 0);

    console.log(`\n[Logger] Started. Current Target: ${currentDateStr} ${displayTime}`);

    let connection;

    try {
        connection = await mysql.createConnection(dbConfig);

        // --- PART 1: FETCH CURRENT HOUR (Standard) ---
        const currentData = await fetchDataForDate(connection, now);

        const { sheetIndex, targetYear } = getSheetContext(now);
        const fileName = `Pulangi IV HEP - Operational Highlights - RAW_${targetYear}.xlsx`;
        const filePath = path.join(OUTPUT_DIR, fileName);

        const cycleStartDate = new Date(targetYear, sheetIndex - 1, 26, 0, 0, 0);
        const diffMs = now - cycleStartDate;
        const diffHours = Math.floor(diffMs / (1000 * 60 * 60));

        let buffer = loadBuffer();

        if (diffHours >= 0 && currentData) {
            const targetRow = 4 + diffHours;

            // Add current hour to buffer
            const newEntry = {
                id: `${targetYear}-${sheetIndex}-${targetRow}`, // ID to prevent dupes in buffer
                targetYear, filePath, sheetIndex, targetRow,
                values: currentData,
                timestamp: new Date().toISOString()
            };

            // Remove if exists, then push
            buffer = buffer.filter(i => i.id !== newEntry.id);
            buffer.push(newEntry);
            console.log(`[Logger] Current hour buffered.`);
        }

        // --- PART 2: BACKFILL MISSING DATA (Scan Excel) ---
        // We only attempt this if the file exists.
        if (fs.existsSync(filePath)) {
            console.log(`[Backfill] Scanning ${fileName} for blanks...`);

            try {
                // Attempt to read Excel (This might fail if locked, which is fine)
                const workbook = await XlsxPopulate.fromFileAsync(filePath);
                const sheet = workbook.sheet(sheetIndex);

                if (sheet) {
                    const currentTargetRow = 4 + diffHours;

                    // Scan from Row 4 up to Current Row
                    for (let r = 4; r <= currentTargetRow; r++) {

                        // Check if cell C (Column 3) is empty. (MW1 column)
                        const cellVal = sheet.row(r).cell(3).value();

                        // Check if it's "blank" (undefined, null, or empty string)
                        if (cellVal === undefined || cellVal === null || cellVal === '') {

                            // Reconstruct Date for this row
                            // Row 4 = CycleStart + 0 hours
                            // Row r = CycleStart + (r-4) hours
                            const rowDate = new Date(cycleStartDate);
                            rowDate.setHours(cycleStartDate.getHours() + (r - 4));

                            const rowId = `${targetYear}-${sheetIndex}-${r}`;

                            // Check 1: Is it already in the buffer?
                            const inBuffer = buffer.some(i => i.id === rowId);

                            if (!inBuffer) {
                                console.log(`[Backfill] Found gap at Row ${r} (${rowDate.toLocaleString()}). Searching DB...`);

                                // Check 2: Fetch from DB (With Retry Loop)
                                const missedData = await fetchDataForDate(connection, rowDate);

                                if (missedData) {
                                    const backfillEntry = {
                                        id: rowId,
                                        targetYear, filePath, sheetIndex,
                                        targetRow: r,
                                        values: missedData,
                                        timestamp: new Date().toISOString()
                                    };
                                    buffer.push(backfillEntry);
                                    console.log(`[Backfill] Recovered data for Row ${r}.`);
                                } else {
                                    // Logged inside fetchDataForDate
                                }
                            }
                        }
                    }
                }
            } catch (fileErr) {
                if (fileErr.code === 'EBUSY' || fileErr.code === 'EPERM' || fileErr.code === 'EACCES') {
                    console.warn(`[Backfill] Skipped scan: Excel file is OPEN/LOCKED.`);
                } else {
                    console.error(`[Backfill] Error reading Excel:`, fileErr.message);
                }
            }
        }

        // --- SAVE BUFFER ---
        saveBuffer(buffer);
        console.log(`[Logger] Complete. Buffer size: ${buffer.length}`);

    } catch (error) {
        console.error('[Logger] Critical Error:', error);
    } finally {
        if (connection) await connection.end();
    }
}

// Run immediately on start
logToBuffer();

// Schedule: Minute 1 of every hour
cron.schedule('1 * * * *', () => logToBuffer());
