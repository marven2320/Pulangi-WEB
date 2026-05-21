const mysql = require('mysql2/promise');
const XlsxPopulate = require('xlsx-populate');
const cron = require('node-cron');
const path = require('path');
const fs = require('fs').promises;
const fsSync = require('fs');

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

if (!fsSync.existsSync(OUTPUT_DIR)) fsSync.mkdirSync(OUTPUT_DIR);

// Use a POOL for concurrent queries and connection stability
const dbPool = mysql.createPool({
    host: 'localhost',
    user: 'root',
    password: 'KaRMSys2025!',
    database: 'pulangi_data',
    connectionLimit: 5,
    enableKeepAlive: true,
    keepAliveInitialDelay: 10000
});

// --- HELPER FUNCTIONS ---

const formatDateString = (date) => date.toLocaleDateString('en-CA'); // YYYY-MM-DD

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

const formatTime24 = (date) => date.toTimeString().split(' ')[0]; // HH:MM:SS

function getSheetContext(cycleStartDate) {
    const startMonth = cycleStartDate.getMonth();
    const startYear = cycleStartDate.getFullYear();

    if (startMonth === 11) {
        return { sheetIndex: 0, targetYear: startYear + 1 };
    } else {
        return { sheetIndex: startMonth + 1, targetYear: startYear };
    }
}

async function readJsonLogAsync() {
    try {
        const data = await fs.readFile(DAILY_LOG_FILE, 'utf-8');
        return JSON.parse(data);
    } catch (e) {
        return []; // Return empty array if missing
    }
}

// --- CORE DATA PROCESSING ---
async function fetchCycleData(startDate, endDate, label, forceClose = false) {
    console.log(`[DB Query] Fetching data for ${label}...`);

    const { sheetIndex, targetYear } = getSheetContext(startDate);
    const startDateStr = formatDateString(startDate);
    const startTimeStr = "00:00:00";
    const endDateStr = formatDateString(endDate);

    const targetCols = UNITS.map(u => u.col).join(', ');
    const refCols = UNITS.map(u => u.ref).join(', ');

    const allEvents = [];
    const ongoingEvents = {};
    UNITS.forEach(u => ongoingEvents[u.col] = null);

    const prevQuery = `
        SELECT ${targetCols}, ${refCols}
        FROM pulangi 
        WHERE ${DATE_COL} < ? OR (${DATE_COL} = ? AND ${TIME_COL} < ?)
        ORDER BY ${DATE_COL} DESC, ${TIME_COL} DESC LIMIT 1
    `;

    const mainQuery = `
        SELECT DATE_FORMAT(${DATE_COL}, '%Y-%m-%d') as date_str, ${TIME_COL} as time_str, ${targetCols}, ${refCols}
        FROM pulangi 
        WHERE ${DATE_COL} >= ? AND ${DATE_COL} <= ?
        ORDER BY ${DATE_COL} ASC, ${TIME_COL} ASC
    `;

    const [[prevRows], [rows]] = await Promise.all([
        dbPool.execute(prevQuery, [startDateStr, startDateStr, startTimeStr]),
        dbPool.execute(mainQuery, [startDateStr, endDateStr])
    ]);

    if (prevRows.length > 0) {
        const prevRow = prevRows[0];
        for (const unit of UNITS) {
            const val = Number(prevRow[unit.col]);
            const refVal = Number(prevRow[unit.ref]);
            if (((val === 0 || val < 1) && refVal !== 0)) {
                ongoingEvents[unit.col] = { unitName: unit.name, start: startDate, end: null, isOngoing: false };
            }
        }
    }

    for (const row of rows) {
        const timestamp = new Date(`${row.date_str} ${row.time_str}`);
        if (timestamp < startDate || timestamp >= endDate) continue;

        for (const unit of UNITS) {
            const val = Number(row[unit.col]);
            const refVal = Number(row[unit.ref]);
            const isDown = ((val === 0 || val < 1) && refVal !== 0);

            if (isDown) {
                if (!ongoingEvents[unit.col]) {
                    ongoingEvents[unit.col] = { unitName: unit.name, start: timestamp, end: null, isOngoing: false };
                }
            } else if (ongoingEvents[unit.col] && !!val) {
                const event = ongoingEvents[unit.col];
                event.end = timestamp;
                event.isOngoing = false;

                const durationHrs = (event.end - event.start) / (1000 * 60 * 60);
                if (durationHrs > NOISE_THRESHOLD_HOURS) {
                    event.duration = durationHrs;
                    allEvents.push(event);
                }
                ongoingEvents[unit.col] = null;
            }
        }
    }

    UNITS.forEach(unit => {
        if (ongoingEvents[unit.col]) {
            const event = ongoingEvents[unit.col];
            event.end = endDate;
            event.isOngoing = !forceClose;
            const durationHrs = (event.end - event.start) / (1000 * 60 * 60);
            if (durationHrs > NOISE_THRESHOLD_HOURS) {
                event.duration = durationHrs;
                allEvents.push(event);
            }
        }
    });

    allEvents.sort((a, b) => (a.unitName !== b.unitName ? a.unitName - b.unitName : a.start - b.start));

    // --- JSON LOGGING (Self-Healing) ---
    const cycleDailyLogs = [];
    const today = new Date();
    const todayStr = formatDateString(today);

    // Calculate Yesterday
    const yesterday = new Date(today);
    yesterday.setDate(today.getDate() - 1);
    const yesterdayStr = formatDateString(yesterday);

    allEvents.forEach(event => {
        const evtStartStr = formatDateString(event.start);
        const evtEndStr = event.end ? formatDateString(event.end) : null;

        // Keep records for BOTH Today and Yesterday to ensure the Daily Report never misses data
        if (evtStartStr === todayStr || evtStartStr === yesterdayStr) {
            cycleDailyLogs.push({ unit: event.unitName, type: 'OUT', date: evtStartStr, time: formatTime24(event.start) });
        }
        if ((evtEndStr === todayStr || evtEndStr === yesterdayStr) && !event.isOngoing) {
            cycleDailyLogs.push({ unit: event.unitName, type: 'IN', date: evtEndStr, time: formatTime24(event.end) });
        }
    });

    return { label, targetYear, sheetIndex, allEvents, cycleDailyLogs };
}

// --- BATCHED EXCEL WRITER ---
async function writeBatchedExcel(cycleResults) {
    const filesToUpdate = {};
    for (const result of cycleResults) {
        if (!filesToUpdate[result.targetYear]) filesToUpdate[result.targetYear] = [];
        filesToUpdate[result.targetYear].push(result);
    }

    for (const [year, cycles] of Object.entries(filesToUpdate)) {
        const fileName = `Pulangi IV HEP - Outage Report_${year}.xlsx`;
        const filePath = path.join(OUTPUT_DIR, fileName);

        console.log(`[Excel] Opening workbook: ${fileName}`);
        let workbook = null;

        try {
            try {
                await fs.access(filePath);
                workbook = await XlsxPopulate.fromFileAsync(filePath);
            } catch {
                console.log(`[Excel] File not found, generating from template...`);
                workbook = await XlsxPopulate.fromFileAsync(TEMPLATE_PATH);
            }

            for (const { label, sheetIndex, allEvents } of cycles) {
                const sheet = workbook.sheet(sheetIndex);
                if (sheet) {
                    let currentRow = 10;
                    for (let wipeRow = 10; wipeRow <= 60; wipeRow++) sheet.row(wipeRow).cell(1).value(null);

                    for (const event of allEvents) {
                        sheet.row(currentRow).cell(1).value(event.unitName);
                        sheet.row(currentRow).cell(3).value(formatExcelDate(event.start));
                        sheet.row(currentRow).cell(4).value(formatTimeString(event.start));

                        if (!event.isOngoing) {
                            sheet.row(currentRow).cell(5).value(formatExcelDate(event.end));
                            sheet.row(currentRow).cell(6).value(formatTimeString(event.end));
                        } else {
                            sheet.row(currentRow).cell(5).value(null);
                            sheet.row(currentRow).cell(6).value(null);
                        }
                        currentRow += 2;
                    }
                    console.log(`[Excel] Applied ${allEvents.length} events for ${label} (Sheet Index ${sheetIndex})`);
                }
            }

            await workbook.toFileAsync(filePath);
            console.log(`[Excel] Saved Changes to: ${fileName}\n`);

        } catch (fileErr) {
            if (fileErr.code === 'EBUSY' || fileErr.code === 'EPERM' || fileErr.code === 'EACCES') {
                console.warn(`\n[WARNING] Could not save Excel file. It is currently OPEN or LOCKED by another program.`);
                console.warn(`[WARNING] Changes will be applied on the next 5-minute cycle once the file is closed.\n`);
            } else {
                console.error(`[Excel Error] Failed to process ${fileName}:`, fileErr);
            }
        } finally {
            workbook = null; // Free Memory
        }
    }
}

// --- MAIN EXECUTION ---
let isJobRunning = false;

async function runUpdate() {
    if (isJobRunning) {
        console.warn(`[Overlap Prevented] Previous cycle is still processing. Skipping this run.`);
        return;
    }

    isJobRunning = true;

    const runId = Date.now();
    const timerLabel = `JobExecutionTime_${runId}`;
    console.time(timerLabel);

    const today = new Date();
    const todayStr = formatDateString(today);

    const yesterday = new Date(today);
    yesterday.setDate(today.getDate() - 1);
    const yesterdayStr = formatDateString(yesterday);

    console.log(`\n[Job Start] Target: ${todayStr} | Threshold (Cleanup): < ${yesterdayStr}`);

    try {
        const day = today.getDate();
        const month = today.getMonth();
        const year = today.getFullYear();
        let currCycleStart, currCycleEnd;

        // ONLY target the Current Cycle
        if (day < 26) {
            currCycleStart = new Date(year, month - 1, 26, 0, 0, 0);
            currCycleEnd = today;
        } else {
            currCycleStart = new Date(year, month, 26, 0, 0, 0);
            currCycleEnd = today;
        }

        const [currResult, rawExistingLogs] = await Promise.all([
            fetchCycleData(currCycleStart, currCycleEnd, "Current Active Cycle", false),
            readJsonLogAsync()
        ]);

        // Process Excel for Current Cycle Only
        await writeBatchedExcel([currResult]);

        // --- JSON CLEANUP & APPEND LOGIC ---
        // Purge any old logs (Older than Yesterday)
        const existingLogs = rawExistingLogs.filter(log => log.date >= yesterdayStr);
        let cleanedCount = rawExistingLogs.length - existingLogs.length;
        if (cleanedCount > 0) console.log(`[JSON Buffer] Cleaned up ${cleanedCount} old entries (Older than ${yesterdayStr}).`);

        const newGeneratedLogs = currResult.cycleDailyLogs;
        const existingSignatures = new Set(existingLogs.map(l => `${l.unit}_${l.type}_${l.date}_${l.time}`));
        let isUpdated = cleanedCount > 0; // If we deleted old rows, we must save the file

        newGeneratedLogs.forEach(newLog => {
            const sig = `${newLog.unit}_${newLog.type}_${newLog.date}_${newLog.time}`;
            if (!existingSignatures.has(sig)) {
                existingLogs.push(newLog);
                existingSignatures.add(sig);
                isUpdated = true;
            }
        });

        if (isUpdated) {
            existingLogs.sort((a, b) => (a.date !== b.date ? a.date.localeCompare(b.date) : a.time.localeCompare(b.time)));
            await fs.writeFile(DAILY_LOG_FILE, JSON.stringify(existingLogs, null, 2));
            console.log(`[JSON Buffer] Saved file. Total entries retained: ${existingLogs.length}.`);
        } else {
            console.log(`[JSON Buffer] No new outages detected. Log file was not modified.`);
        }

    } catch (err) {
        console.error(`[Error in Job]:`, err);
    } finally {
        console.timeEnd(timerLabel);
        isJobRunning = false;
    }
}

// --- SCHEDULE ---
cron.schedule('*/2 * * * *', () => {
    runUpdate();
});

// Run immediately on boot
runUpdate();
