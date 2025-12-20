const mysql = require('mysql2/promise');
const XlsxPopulate = require('xlsx-populate');
const cron = require('node-cron');
const path = require('path');
const fs = require('fs');

// --- CONFIGURATION ---
const OUTPUT_DIR = path.join(__dirname, 'reports');
const TEMPLATE_PATH = path.join(__dirname, '/reports/Pulangi IV HEP - Outage Report Template.xlsx');

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

const formatDateString = (date) => date.toLocaleDateString('en-CA');

const formatTimeString = (date) => {
    return date.toLocaleTimeString('en-US', {
        hour: '2-digit',
        minute: '2-digit',
        hour12: true
    });
};

// Map a Cycle Start Date to the correct Excel Sheet (Month)
function getSheetContext(cycleStartDate) {
    const startMonth = cycleStartDate.getMonth();
    const startYear = cycleStartDate.getFullYear();

    // If cycle starts in Dec (11), it belongs to next year's Jan (0) sheet
    if (startMonth === 11) {
        return { sheetIndex: 0, targetYear: startYear + 1 };
    } else {
        return { sheetIndex: startMonth + 1, targetYear: startYear };
    }
}

// --- CORE PROCESSING FUNCTION ---
// This function handles the logic for ONE specific date range.
async function processCycle(startDate, endDate, label) {
    console.log(`\n[Processing ${label}]`);

    // 1. Identify Target Excel File/Sheet
    const { sheetIndex, targetYear } = getSheetContext(startDate);

    const startDateStr = formatDateString(startDate);
    const endDateStr = formatDateString(endDate);
    const endTimeStr = formatTimeString(endDate); // Useful log for "Current" cycle

    console.log(`   Range: ${startDateStr} to ${endDateStr} ${endTimeStr}`);
    console.log(`   Target: Year ${targetYear}, Sheet Index ${sheetIndex}`);

    let connection;
    try {
        // 2. Fetch Data
        const targetCols = UNITS.map(u => u.col).join(', ');
        const refCols = UNITS.map(u => u.ref).join(', ');

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

        connection = await mysql.createConnection(dbConfig);
        const [rows] = await connection.execute(query, [startDateStr, endDateStr]);

        // 3. Identify Events
        const allEvents = [];
        const ongoingEvents = {};
        UNITS.forEach(u => ongoingEvents[u.col] = null);

        for (const row of rows) {
            const dateTimeString = `${row.date_str} ${row.time_str}`;
            const timestamp = new Date(dateTimeString);

            // Strict Time Filtering
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
                            end: null
                        };
                    }
                } else {
                    if (ongoingEvents[unit.col] && !!val) {
                        const event = ongoingEvents[unit.col];
                        event.end = timestamp;

                        const diffMs = event.end - event.start;
                        const durationHrs = diffMs / (1000 * 60 * 60);

                        // Noise Filter
                        if (durationHrs > NOISE_THRESHOLD_HOURS) {
                            event.duration = durationHrs;
                            allEvents.push(event);
                        }
                        ongoingEvents[unit.col] = null;
                    }
                }
            }
        }

        // Close ongoing events at the exact End Date (or "Now")
        UNITS.forEach(unit => {
            if (ongoingEvents[unit.col]) {
                const event = ongoingEvents[unit.col];
                event.end = endDate;

                const diffMs = event.end - event.start;
                const durationHrs = diffMs / (1000 * 60 * 60);

                if (durationHrs > NOISE_THRESHOLD_HOURS) {
                    event.duration = durationHrs;
                    allEvents.push(event);
                }
            }
        });

        console.log(`   Found ${allEvents.length} events.`);

        // 4. Update Excel
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
        if (!sheet) {
            console.error(`   [Error] Sheet index ${sheetIndex} missing.`);
            return;
        }

        // Clear Old Data (Columns A-F only)
        // We scan 500 rows to ensure we remove any data from previous runs that might be obsolete
        //sheet.range("A2:F500").value(null);

        // Write New Data
        let currentRow = 10; // Start data row
        for (const event of allEvents) {
            sheet.row(currentRow).cell(1).value(event.unitName);
            sheet.row(currentRow).cell(3).value(formatDateString(event.start));
            sheet.row(currentRow).cell(4).value(formatTimeString(event.start));
            sheet.row(currentRow).cell(5).value(formatDateString(event.end));
            sheet.row(currentRow).cell(6).value(formatTimeString(event.end));
            //sheet.row(currentRow).cell(6).value(event.duration).style("numberFormat", "0.00");
            currentRow = currentRow + 2;
        }

        await workbook.toFileAsync(filePath);
        console.log(`   [Success] Updated ${fileName} (Sheet ${sheetIndex})`);

    } catch (error) {
        console.error('   [Error]', error);
    } finally {
        if (connection) await connection.end();
    }
}

// --- MAIN EXECUTION ---

async function runDualUpdate() {
    const today = new Date();
    const day = today.getDate();
    const month = today.getMonth();
    const year = today.getFullYear();

    console.log(`[Job Start] Current Date: ${formatDateString(today)}`);

    let prevCycleStart, prevCycleEnd;
    let currCycleStart, currCycleEnd;

    // LOGIC: Determine the two cycles

    if (day < 26) {
        // SCENARIO: Today is Dec 20
        // Previous Cycle: Oct 26 - Nov 26
        // Current Cycle: Nov 26 - Dec 20 (Now)

        prevCycleStart = new Date(year, month - 2, 26, 0, 0, 0);
        prevCycleEnd = new Date(year, month - 1, 26, 0, 0, 0);

        currCycleStart = new Date(year, month - 1, 26, 0, 0, 0);
        currCycleEnd = today; // Up to NOW
    } else {
        // SCENARIO: Today is Dec 27
        // Previous Cycle: Nov 26 - Dec 26
        // Current Cycle: Dec 26 - Dec 27 (Now)

        prevCycleStart = new Date(year, month - 1, 26, 0, 0, 0);
        prevCycleEnd = new Date(year, month, 26, 0, 0, 0);

        currCycleStart = new Date(year, month, 26, 0, 0, 0);
        currCycleEnd = today; // Up to NOW
    }

    // 1. Process Previous Cycle (Full Month)
    await processCycle(prevCycleStart, prevCycleEnd, "Previous Cycle");

    // 2. Process Current Cycle (Partial Month)
    await processCycle(currCycleStart, currCycleEnd, "Current Active Cycle");

    console.log(`\n[Job Complete]`);
}

// --- SCHEDULE ---
// Run at 00:00 (Midnight) and 12:00 (Noon)
cron.schedule('1 0,12 * * *', () => {
    runDualUpdate();
});

// Run
runDualUpdate();