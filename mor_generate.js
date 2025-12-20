const XlsxPopulate = require('xlsx-populate');
const path = require('path');
const fs = require('fs');
const cron = require('node-cron');

// --- CONFIGURATION ---
const OUTPUT_DIR = path.join(__dirname, 'reports');
const SUMMARY_TEMPLATE = path.join(__dirname, '/reports/Pulangi IV HEP - Monthly Operations Report Template.xlsx');

if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR);

// ==========================================
// 1. MONTHLY LOG 
// ==========================================
const getMonthlyLogFilename = (year) => `Pulangi IV HEP - Operational Highlights - RAW_${year}.xlsx`;
const MONTHLY_START_ROW = 4;

// A. Standard Points (C, G, K)
const MONTHLY_DATA_POINTS = [
    { colLetter: "C", destCellMax: "J14", destCellMin: "J15" },
    { colLetter: "G", destCellMax: "K14", destCellMin: "K15" },
    { colLetter: "K", destCellMax: "L14", destCellMin: "L15" }
];

// B. Special Point Q
const MONTHLY_COL_Q = {
    colLetter: "Q",
    destCellMax: "L43", // Max Value (Last Row of Data)
    destCellMin: "L40"  // Min Value (Start Row of Data)
};

// ==========================================
// 2. SHIFT LOG (Updated Cycle & Columns)
// ==========================================
const getShiftLogFilename = (year) => `Pulangi IV HEP - Generation Data - RAW_${year}.xlsx`;
// Assuming Shift Data starts at Row 4 (Change if headers are different)
const SHIFT_START_ROW = 4;

// Configuration for Columns C, E, G
const SHIFT_DATA_POINTS = [
    { colLetter: "C", destCell: "J16" }, // Overall 1
    { colLetter: "E", destCell: "K16" }, // Overall 2
    { colLetter: "G", destCell: "L16" }  // Overall 3
];

// ==========================================
// 3. DOWNTIME LOG
// ==========================================
const getDowntimeLogFilename = (year) => `Pulangi IV HEP - Outage Report_${year}.xlsx`;

const DOWNTIME_METRICS = [
    { name: "RS", sourceCol: "B", destRow: 27 },
    { name: "PO", sourceCol: "C", destRow: 29 },
    { name: "MO", sourceCol: "D", destRow: 30 },
    { name: "FO", sourceCol: "E", destRow: 31 },
    { name: "OM", sourceCol: "F", destRow: 32 },
    { name: "EOH", sourceCol: "G", destRow: 33 }
];

const DOWNTIME_UNITS = [
    { name: "Unit 1", sourceFixedRow: 3, destCol: "J" },
    { name: "Unit 2", sourceFixedRow: 4, destCol: "K" },
    { name: "Unit 3", sourceFixedRow: 5, destCol: "L" }
];

// ==========================================

// --- HELPER FUNCTIONS ---

function getCycleContext(date) {
    const day = date.getDate();
    const month = date.getMonth();
    const year = date.getFullYear();

    let cycleStartYear = year;
    let cycleStartMonth = month;

    if (day < 26) {
        cycleStartMonth = month - 1;
        if (cycleStartMonth < 0) {
            cycleStartMonth = 11;
            cycleStartYear = year - 1;
        }
    }

    let sheetIndex, fileYear;
    if (cycleStartMonth === 11) {
        sheetIndex = 0;
        fileYear = cycleStartYear + 1;
    } else {
        sheetIndex = cycleStartMonth + 1;
        fileYear = cycleStartYear;
    }

    // Standard Hourly Cycle (26th 00:00 to 26th 00:00)
    const monthlyStartDate = new Date(cycleStartYear, cycleStartMonth, 26, 0, 0, 0);
    const monthlyEndDate = new Date(cycleStartYear, cycleStartMonth + 1, 26, 0, 0, 0);

    // Shift Log Cycle (26th 12:00 PM to 26th 00:00 AM next month)
    const shiftStartDate = new Date(cycleStartYear, cycleStartMonth, 26, 12, 0, 0); // 12:00 NN
    const shiftEndDate = new Date(cycleStartYear, cycleStartMonth + 1, 26, 0, 0, 0);   // 12:00 AM (Midnight)

    return {
        sheetIndex, fileYear,
        monthlyStartDate, monthlyEndDate,
        shiftStartDate, shiftEndDate
    };
}

// --- CORE TASK ---

async function generateMonthlySummary() {
    console.log(`\n[Monthly Summary] Starting generation...`);

    const today = new Date();
    const {
        sheetIndex, fileYear,
        monthlyStartDate, monthlyEndDate,
        shiftStartDate, shiftEndDate
    } = getCycleContext(today);

    // File Names
    const monthlyFilename = getMonthlyLogFilename(fileYear);
    const shiftFilename = getShiftLogFilename(fileYear);
    const downtimeFilename = getDowntimeLogFilename(fileYear);

    // Full Paths
    const monthlyPath = path.join(OUTPUT_DIR, monthlyFilename);
    const shiftPath = path.join(OUTPUT_DIR, shiftFilename);
    const downtimePath = path.join(OUTPUT_DIR, downtimeFilename);

    const summaryFilename = `Pulangi IV HEP - Monthly Operations Report_${fileYear}.xlsx`;
    const summaryPath = path.join(OUTPUT_DIR, summaryFilename);

    let monthlyFormulas = [];
    let shiftFormulas = [];
    let downtimeFormulas = [];

    try {
        // --- STEP A: MONTHLY LOG ---
        if (fs.existsSync(monthlyPath)) {
            console.log(`   [Processing] Monthly Log References...`);
            const wb = await XlsxPopulate.fromFileAsync(monthlyPath);
            const sheet = wb.sheet(sheetIndex);

            if (sheet) {
                const sheetName = sheet.name();

                // Calculate Hourly Rows
                const diffMs = monthlyEndDate - monthlyStartDate;
                const totalHours = Math.round(diffMs / (1000 * 60 * 60));

                const dataEndRow = MONTHLY_START_ROW + totalHours;
                const summaryMaxRow = dataEndRow + 1;
                const summaryMinRow = dataEndRow + 2;

                // 1. Standard Columns
                MONTHLY_DATA_POINTS.forEach(point => {
                    const fMax = `'[${monthlyFilename}]${sheetName}'!${point.colLetter}${summaryMaxRow}`;
                    const fMin = `'[${monthlyFilename}]${sheetName}'!${point.colLetter}${summaryMinRow}`;
                    monthlyFormulas.push({ cell: point.destCellMax, formula: fMax });
                    monthlyFormulas.push({ cell: point.destCellMin, formula: fMin });
                });

                // 2. Special Column Q
                const fQMin = `'[${monthlyFilename}]${sheetName}'!${MONTHLY_COL_Q.colLetter}${MONTHLY_START_ROW}`;
                const fQMax = `'[${monthlyFilename}]${sheetName}'!${MONTHLY_COL_Q.colLetter}${dataEndRow}`;
                monthlyFormulas.push({ cell: MONTHLY_COL_Q.destCellMin, formula: fQMin });
                monthlyFormulas.push({ cell: MONTHLY_COL_Q.destCellMax, formula: fQMax });
            }
        }

        // --- STEP B: SHIFT LOG (Twice Daily Calculation) ---
        if (fs.existsSync(shiftPath)) {
            console.log(`   [Processing] Shift Log References...`);
            const wb = await XlsxPopulate.fromFileAsync(shiftPath);
            const sheet = wb.sheet(sheetIndex);
            if (sheet) {
                const sheetName = sheet.name();

                // 1. Calculate Shift Rows
                // Difference between Start (Noon) and End (Midnight)
                const diffMs = shiftEndDate - shiftStartDate;
                const totalHours = diffMs / (1000 * 60 * 60);

                // 2. Calculate Number of Logs (Every 12 hours)
                // We assume there is a row for the start time.
                // Number of intervals = Total Hours / 12
                // Rows = Intervals + 1 (Start Row) or Intervals? 
                // Usually logs include the start time, so we add 1? 
                // Let's assume strict interval count mapping to rows:
                const numberOfLogs = Math.round(totalHours / 12);

                // 3. Determine Summary Row
                // "The row is directly after the last data log"
                // Last Data Row = Start Row + numberOfLogs
                // Summary Row   = Last Data Row + 1
                const lastDataRow = SHIFT_START_ROW + numberOfLogs;
                const summaryRow = lastDataRow + 1;

                console.log(`     Shift Start: ${shiftStartDate.toLocaleString()} | End: ${shiftEndDate.toLocaleString()}`);
                console.log(`     Shift Hours: ${totalHours} | Logs: ${numberOfLogs}`);
                console.log(`     Shift Summary Row: ${summaryRow}`);

                // 4. Map Columns C, E, G
                SHIFT_DATA_POINTS.forEach(point => {
                    const formula = `'[${shiftFilename}]${sheetName}'!${point.colLetter}${summaryRow}`;
                    shiftFormulas.push({ cell: point.destCell, formula: formula });
                });
            }
        }

        // --- STEP C: DOWNTIME LOG ---
        if (fs.existsSync(downtimePath)) {
            console.log(`   [Processing] Downtime Log References...`);
            const wb = await XlsxPopulate.fromFileAsync(downtimePath);
            const sheet = wb.sheet(sheetIndex);

            if (sheet) {
                const sheetName = sheet.name();
                DOWNTIME_UNITS.forEach(unit => {
                    const sourceRow = unit.sourceFixedRow;
                    DOWNTIME_METRICS.forEach(metric => {
                        const formula = `'[${downtimeFilename}]${sheetName}'!${metric.sourceCol}${sourceRow}`;
                        const destCell = `${unit.destCol}${metric.destRow}`;
                        downtimeFormulas.push({ cell: destCell, formula: formula });
                    });
                });
            }
        }

        // --- STEP D: WRITE SUMMARY REPORT ---
        console.log(`   [Generating] Summary Report with Formulas...`);

        let destWb;
        if (fs.existsSync(SUMMARY_TEMPLATE)) {
            destWb = await XlsxPopulate.fromFileAsync(SUMMARY_TEMPLATE);
        } else {
            console.error(`   [Error] Template file missing: ${SUMMARY_TEMPLATE}`);
            return;
        }

        const destSheet = destWb.sheet(sheetIndex);
        if (!destSheet) {
            console.error(`   [Error] Sheet index ${sheetIndex} does not exist in the template.`);
            return;
        }

        console.log(`   Writing to Destination Sheet: Index ${sheetIndex} (${destSheet.name()})`);
        // Apply Monthly Formulas
        monthlyFormulas.forEach(item => destSheet.cell(item.cell).formula(item.formula));

        // Apply Shift Formulas
        shiftFormulas.forEach(item => destSheet.cell(item.cell).formula(item.formula));

        // Apply Downtime Formulas
        downtimeFormulas.forEach(item => {
            destSheet.cell(item.cell).formula(item.formula).style("numberFormat", "0.00");
        });
        destSheet.cell("H7").value(monthlyStartDate.toLocaleDateString());
        destSheet.cell("K7").value(monthlyEndDate.toLocaleDateString());

        await destWb.toFileAsync(summaryPath);
        console.log(`   [Success] Report saved to ${summaryPath}`);

    } catch (error) {
        console.error(`   [Error]`, error);
    }
}

// --- SCHEDULE ---
generateMonthlySummary();

cron.schedule('1 12 26 * *', () => {
    generateMonthlySummary();
});

console.log("Monthly Summary Service Started.");