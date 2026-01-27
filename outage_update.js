const mysql = require('mysql2/promise');
const XlsxPopulate = require('xlsx-populate');
const cron = require('node-cron');
const path = require('path');
const fs = require('fs');

// --- CONFIGURATION ---
const OUTPUT_DIR = path.join(__dirname, 'reports');
const TEMPLATE_PATH = path.join(__dirname, '/templates/Pulangi IV HEP - Outage Report Template.xlsx');
const DAILY_LOG_FILE = path.join(OUTPUT_DIR, 'daily_outage_log.json');

const UNITS = [
    { name: 1, col: 'mw1', ref: 'freq1' },
    { name: 2, col: 'mw2', ref: 'freq2' },
    { name: 3, col: 'mw3', ref: 'freq3' }
];

const DATE_COL = 'date';
const TIME_COL = 'time';
const NOISE_THRESHOLD_HOURS = 0.15; // Ignore outages <= 9 mins

if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR);

const dbConfig = {
    host: 'localhost',
    user: 'root',
    password: 'KaRMSys2025!',
    database: 'pulangi_data',
    connectionLimit: 10
};

// --- HELPER FUNCTIONS ---

// Keep this as YYYY-MM-DD for Database and Logic (Sorting/Filtering)
const formatDateString = (date) => date.toLocaleDateString('en-CA');

// ** NEW: Specific format for Excel Output (dd-MMM-yy) **
const formatExcelDate = (date) => {
    return date.toLocaleDateString('en-GB', {
        day: '2-digit',
        month: 'short',
        year: '2-digit'
    }).replace(/ /g, '-');
};

const formatTimeString = (date) => {
    return date.toLocaleTimeString('en-US', {
        hour: '2-digit',
        minute: '2-digit',
        hour12: true
    });
};

const formatTime24 = (date) => {
    return date.toTimeString().split(' ')[0]; // HH:MM:SS
};

function getSheetContext(cycleStartDate) {
    const startMonth = cycleStartDate.getMonth();
    const startYear = cycleStartDate.getFullYear();

    if (startMonth === 11) {
        return { sheetIndex: 0, targetYear: startYear + 1 };
    } else {
        return { sheetIndex: startMonth + 1, targetYear: startYear };
    }
}

// Helper to read existing JSON log safely
function readJsonLog() {
    if (!fs.existsSync(DAILY_LOG_FILE)) return [];
    try {
        return JSON.parse(fs.readFileSync(DAILY_LOG_FILE));
    } catch (e) {
        return [];
    }
}

// --- CORE PROCESSING FUNCTION ---
async function processCycle(startDate, endDate, label, forceClose = false, generateDailyJson = false) {
    console.log(`\n[Processing ${label}]`);

    const { sheetIndex, targetYear } = getSheetContext(startDate);
    const startDateStr = formatDateString(startDate);
    const startTimeStr = "00:00:00";
    const endDateStr = formatDateString(endDate);

    // Captured logs for this specific cycle run
    const cycleDailyLogs = [];

    let connection;
    try {
        connection = await mysql.createConnection(dbConfig);
        const targetCols = UNITS.map(u => u.col).join(', ');
        const refCols = UNITS.map(u => u.ref).join(', ');

        const allEvents = [];
        const ongoingEvents = {};
        UNITS.forEach(u => ongoingEvents[u.col] = null);

        // --- STEP A: CHECK PREVIOUS STATE ---
        const prevQuery = `
            SELECT ${targetCols}, ${refCols}
            FROM pulangi 
            WHERE ${DATE_COL} < ? 
               OR (${DATE_COL} = ? AND ${TIME_COL} < ?)
            ORDER BY ${DATE_COL} DESC, ${TIME_COL} DESC
            LIMIT 1
        `;
        const [prevRows] = await connection.execute(prevQuery, [startDateStr, startDateStr, startTimeStr]);

        if (prevRows.length > 0) {
            const prevRow = prevRows[0];
            for (const unit of UNITS) {
                const val = Number(prevRow[unit.col]);
                const refVal = Number(prevRow[unit.ref]);
                const wasDown = ((val === 0 || val < 1) && refVal !== 0);

                if (wasDown) {
                    ongoingEvents[unit.col] = {
                        unitName: unit.name,
                        start: startDate,
                        end: null,
                        isOngoing: false
                    };
                }
            }
        }

        // --- STEP B: FETCH CURRENT CYCLE DATA ---
        const query = `
            SELECT 
                DATE_FORMAT(${DATE_COL}, '%Y-%m-%d') as date_str, 
                ${TIME_COL} as time_str,
                ${targetCols}, 
                ${refCols}
            FROM pulangi 
            WHERE ${DATE_COL} >= ? AND ${DATE_COL} <= ?
            ORDER BY ${DATE_COL} ASC, ${TIME_COL} ASC
        `;

        const [rows] = await connection.execute(query, [startDateStr, endDateStr]);

        // --- STEP C: PROCESS LOGS ---
        for (const row of rows) {
            const dateTimeString = `${row.date_str} ${row.time_str}`;
            const timestamp = new Date(dateTimeString);

            if (timestamp < startDate || timestamp >= endDate) continue;

            for (const unit of UNITS) {
                const val = Number(row[unit.col]);
                const refVal = Number(row[unit.ref]);
                const isDown = ((val === 0 || val < 1) && refVal !== 0);

                if (isDown) {
                    if (!ongoingEvents[unit.col]) {
                        ongoingEvents[unit.col] = {
                            unitName: unit.name,
                            start: timestamp,
                            end: null,
                            isOngoing: false
                        };
                    }
                } else {
                    if (ongoingEvents[unit.col] && !!val) {
                        const event = ongoingEvents[unit.col];
                        event.end = timestamp;
                        event.isOngoing = false;

                        const diffMs = event.end - event.start;
                        const durationHrs = diffMs / (1000 * 60 * 60);

                        if (durationHrs > NOISE_THRESHOLD_HOURS) {
                            event.duration = durationHrs;
                            allEvents.push(event);
                        }
                        ongoingEvents[unit.col] = null;
                    }
                }
            }
        }

        // --- STEP D: HANDLE UNFINISHED EVENTS ---
        UNITS.forEach(unit => {
            if (ongoingEvents[unit.col]) {
                const event = ongoingEvents[unit.col];
                event.end = endDate;
                event.isOngoing = !forceClose;

                const diffMs = event.end - event.start;
                const durationHrs = diffMs / (1000 * 60 * 60);

                if (durationHrs > NOISE_THRESHOLD_HOURS) {
                    event.duration = durationHrs;
                    allEvents.push(event);
                }
            }
        });

        allEvents.sort((a, b) => {
            if (a.unitName !== b.unitName) return a.unitName - b.unitName;
            return a.start - b.start;
        });

        // --- STEP E: COLLECT DAILY JSON LOGS (IN MEMORY) ---
        if (generateDailyJson) {
            const todayStr = formatDateString(new Date());

            allEvents.forEach(event => {
                const evtStartStr = formatDateString(event.start);
                const evtEndStr = event.end ? formatDateString(event.end) : null;

                if (evtStartStr === todayStr) {
                    cycleDailyLogs.push({
                        unit: event.unitName,
                        type: 'OUT',
                        date: evtStartStr,
                        time: formatTime24(event.start)
                    });
                }

                if (evtEndStr === todayStr && !event.isOngoing) {
                    cycleDailyLogs.push({
                        unit: event.unitName,
                        type: 'IN',
                        date: evtEndStr,
                        time: formatTime24(event.end)
                    });
                }
            });
        }

        // --- STEP F: UPDATE EXCEL ---
        const fileName = `Pulangi IV HEP - Outage Report_${targetYear}.xlsx`;
        const filePath = path.join(OUTPUT_DIR, fileName);

        let workbook;
        if (fs.existsSync(filePath)) {
            workbook = await XlsxPopulate.fromFileAsync(filePath);
        } else {
            console.log(`   Creating new file from template...`);
            workbook = await XlsxPopulate.fromFileAsync(TEMPLATE_PATH);
        }

        const sheet = workbook.sheet(sheetIndex);
        if (sheet) {
            let currentRow = 10;
            for (const event of allEvents) {
                sheet.row(currentRow).cell(1).value(event.unitName);
                // USE formatExcelDate for DD-MMM-YY format
                sheet.row(currentRow).cell(3).value(formatExcelDate(event.start));
                sheet.row(currentRow).cell(4).value(formatTimeString(event.start));

                if (!event.isOngoing) {
                    // USE formatExcelDate for DD-MMM-YY format
                    sheet.row(currentRow).cell(5).value(formatExcelDate(event.end));
                    sheet.row(currentRow).cell(6).value(formatTimeString(event.end));
                } else {
                    sheet.row(currentRow).cell(5).value(null);
                    sheet.row(currentRow).cell(6).value(null);
                }
                currentRow = currentRow + 2;
            }
            await workbook.toFileAsync(filePath);
            console.log(`   [Success] Updated Excel: ${fileName}`);
        }

    } catch (error) {
        console.error('   [Error]', error);
    } finally {
        if (connection) await connection.end();
    }

    return cycleDailyLogs;
}

// --- MAIN EXECUTION ---

async function runDualUpdate() {
    const today = new Date();
    const todayStr = formatDateString(today);

    // Calculate Report Date (Yesterday)
    const reportDate = new Date(today);
    reportDate.setDate(today.getDate() - 1);
    const reportDateStr = formatDateString(reportDate);

    console.log(`[Job Start] Today: ${todayStr}, Report Date: ${reportDateStr}`);

    // --- 1. MANAGE EXISTING JSON LOGS ---
    let existingLogs = readJsonLog();

    let filteredLogs = existingLogs.filter(log => {
        return (log.date >= reportDateStr) && (log.date !== todayStr);
    });

    console.log(`[JSON Buffer] Retained ${filteredLogs.length} entries from previous days.`);

    // --- 2. DETERMINE CYCLES ---
    const day = today.getDate();
    const month = today.getMonth();
    const year = today.getFullYear();
    let prevCycleStart, prevCycleEnd, currCycleStart, currCycleEnd;

    if (day < 26) {
        prevCycleStart = new Date(year, month - 2, 26, 0, 0, 0);
        prevCycleEnd = new Date(year, month - 1, 26, 0, 0, 0);
        currCycleStart = new Date(year, month - 1, 26, 0, 0, 0);
        currCycleEnd = today;
    } else {
        prevCycleStart = new Date(year, month - 1, 26, 0, 0, 0);
        prevCycleEnd = new Date(year, month, 26, 0, 0, 0);
        currCycleStart = new Date(year, month, 26, 0, 0, 0);
        currCycleEnd = today;
    }

    // --- 3. RUN CYCLES & COLLECT NEW LOGS ---
    const logs1 = await processCycle(prevCycleStart, prevCycleEnd, "Previous Cycle", true, true);
    const logs2 = await processCycle(currCycleStart, currCycleEnd, "Current Active Cycle", false, true);

    // --- 4. MERGE & SAVE ---
    const finalLogs = [...filteredLogs, ...logs1, ...logs2];

    finalLogs.sort((a, b) => {
        if (a.date !== b.date) return a.date.localeCompare(b.date);
        return a.time.localeCompare(b.time);
    });

    fs.writeFileSync(DAILY_LOG_FILE, JSON.stringify(finalLogs, null, 2));
    console.log(`[JSON Buffer] Updated file with ${finalLogs.length} entries.`);

    console.log(`\n[Job Complete]`);
}

// --- SCHEDULE ---
cron.schedule('*/5 * * * *', () => {
    runDualUpdate();
});

// Run
runDualUpdate();
