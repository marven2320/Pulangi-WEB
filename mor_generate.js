const XlsxPopulate = require('xlsx-populate');
const path = require('path');
const fs = require('fs');
const cron = require('node-cron');

// --- CONFIGURATION ---
const REPORTS_DIR = path.join(__dirname, 'reports');
const RAW_DATA_DIR = path.join(__dirname, 'rawdata');

// Ensure this matches your template location
const SUMMARY_TEMPLATE = path.join(__dirname, 'templates', 'Pulangi IV HEP - Monthly Operations Report Template.xlsx');

if (!fs.existsSync(REPORTS_DIR)) fs.mkdirSync(REPORTS_DIR);
if (!fs.existsSync(RAW_DATA_DIR)) fs.mkdirSync(RAW_DATA_DIR);

// ==========================================
// 1. MONTHLY LOG (Source: rawdata)
// ==========================================
const getMonthlyLogFilename = (year) => `Pulangi IV HEP - Operational Highlights - RAW_${year}.xlsx`;
const MONTHLY_START_ROW = 4;

// A. Standard Points (C, G, K)
const MONTHLY_DATA_POINTS = [
    { colLetter: "C", destCellMax: "J14", destCellMin: "J15" },
    { colLetter: "G", destCellMax: "K14", destCellMin: "K15" },
    { colLetter: "K", destCellMax: "L14", destCellMin: "L15" }
];

// B. Special Point Q (Min at Start, Max at End)
const MONTHLY_COL_Q = {
    colLetter: "Q",
    destCellMax: "L43",
    destCellMin: "L40"
};

// C. Spillage Data (Column AW)
const MONTHLY_COL_AW = {
    colLetter: "AW",
    destCellAvg: "L46",
    destCellTotal: "L47"
};

// ==========================================
// 2. SHIFT LOG (Source: rawdata)
// ==========================================
const getShiftLogFilename = (year) => `Pulangi IV HEP - Generation Data - RAW_${year}.xlsx`;
const SHIFT_START_ROW = 4;

const SHIFT_DATA_POINTS = [
    { colLetter: "I", destCell: "J16" },
    { colLetter: "J", destCell: "J18" },
    { colLetter: "L", destCell: "K16" },
    { colLetter: "M", destCell: "K18" },
    { colLetter: "O", destCell: "L16" },
    { colLetter: "P", destCell: "L18" }
];

// ==========================================
// 3. DOWNTIME LOG (Source: reports)
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

    const monthlyStartDate = new Date(cycleStartYear, cycleStartMonth, 26, 0, 0, 0);
    const monthlyEndDate = new Date(cycleStartYear, cycleStartMonth + 1, 26, 0, 0, 0);
    const monthlyEndDate_report = new Date(cycleStartYear, cycleStartMonth + 1, 25, 0, 0, 0);

    const shiftStartDate = new Date(cycleStartYear, cycleStartMonth, 26, 12, 0, 0);
    const shiftEndDate = new Date(cycleStartYear, cycleStartMonth + 1, 26, 0, 0, 0);

    return {
        sheetIndex, fileYear,
        monthlyStartDate, monthlyEndDate, monthlyEndDate_report,
        shiftStartDate, shiftEndDate
    };
}

// ** NEW HELPER **: Creates formula with full absolute path based on directory
function createAbsFormula(directory, filename, sheetName, cellRef) {
    return `'${directory}/[${filename}]${sheetName}'!${cellRef}`;
}

// --- CORE TASK ---

async function generateMonthlySummary() {
    console.log(`\n[Monthly Summary] Starting generation...`);

    const today = new Date();
    const {
        sheetIndex, fileYear,
        monthlyStartDate, monthlyEndDate, monthlyEndDate_report,
        shiftStartDate, shiftEndDate
    } = getCycleContext(today);

    // File Names
    const monthlyFilename = getMonthlyLogFilename(fileYear);
    const shiftFilename = getShiftLogFilename(fileYear);
    const downtimeFilename = getDowntimeLogFilename(fileYear);

    // Full Paths
    // Monthly & Shift logs are in RAW_DATA_DIR
    const monthlyPath = path.join(RAW_DATA_DIR, monthlyFilename);
    const shiftPath = path.join(RAW_DATA_DIR, shiftFilename);
    // Downtime log is in REPORTS_DIR
    const downtimePath = path.join(REPORTS_DIR, downtimeFilename);

    const summaryFilename = `Pulangi IV HEP - Monthly Operations Report_${fileYear}.xlsx`;
    const summaryPath = path.join(REPORTS_DIR, summaryFilename);

    let monthlyFormulas = [];
    let shiftFormulas = [];
    let downtimeFormulas = [];

    try {
        // --- STEP A: MONTHLY LOG (From rawdata) ---
        if (fs.existsSync(monthlyPath)) {
            console.log(`   [Processing] Monthly Log References...`);
            const wb = await XlsxPopulate.fromFileAsync(monthlyPath);
            const sheet = wb.sheet(sheetIndex);

            if (sheet) {
                const sheetName = sheet.name();
                const diffMs = monthlyEndDate - monthlyStartDate;
                const totalHours = Math.round(diffMs / (1000 * 60 * 60));
                const dataEndRow = MONTHLY_START_ROW + totalHours;

                const summaryMaxRow = dataEndRow + 1;
                const summaryMinRow = dataEndRow + 3; // +3 as per your logic

                // 1. Standard Columns
                MONTHLY_DATA_POINTS.forEach(point => {
                    const fMax = createAbsFormula(RAW_DATA_DIR, monthlyFilename, sheetName, `${point.colLetter}${summaryMaxRow}`);
                    const fMin = createAbsFormula(RAW_DATA_DIR, monthlyFilename, sheetName, `${point.colLetter}${summaryMinRow}`);
                    monthlyFormulas.push({ cell: point.destCellMax, formula: fMax });
                    monthlyFormulas.push({ cell: point.destCellMin, formula: fMin });
                });

                // 2. Special Column Q
                const fQMin = createAbsFormula(RAW_DATA_DIR, monthlyFilename, sheetName, `${MONTHLY_COL_Q.colLetter}${MONTHLY_START_ROW}`);
                const fQMax = createAbsFormula(RAW_DATA_DIR, monthlyFilename, sheetName, `${MONTHLY_COL_Q.colLetter}${dataEndRow}`);
                monthlyFormulas.push({ cell: MONTHLY_COL_Q.destCellMin, formula: fQMin });
                monthlyFormulas.push({ cell: MONTHLY_COL_Q.destCellMax, formula: fQMax });

                // 3. Spillage Column AW
                const rowAvg = dataEndRow + 2;
                const rowTotal = dataEndRow + 4;
                const fAwAvg = createAbsFormula(RAW_DATA_DIR, monthlyFilename, sheetName, `${MONTHLY_COL_AW.colLetter}${rowAvg}`) + "/3600";
                const fAwTotal = createAbsFormula(RAW_DATA_DIR, monthlyFilename, sheetName, `${MONTHLY_COL_AW.colLetter}${rowTotal}`) + "/1000000";

                monthlyFormulas.push({ cell: MONTHLY_COL_AW.destCellAvg, formula: fAwAvg });
                monthlyFormulas.push({ cell: MONTHLY_COL_AW.destCellTotal, formula: fAwTotal });
            }
        } else {
            console.warn(`   [Warning] Monthly Log not found at: ${monthlyPath}`);
        }

        // --- STEP B: SHIFT LOG (From rawdata) ---
        if (fs.existsSync(shiftPath)) {
            console.log(`   [Processing] Shift Log References...`);
            const wb = await XlsxPopulate.fromFileAsync(shiftPath);
            const sheet = wb.sheet(sheetIndex);
            if (sheet) {
                const sheetName = sheet.name();
                const diffMs = shiftEndDate - shiftStartDate;
                const totalHours = diffMs / (1000 * 60 * 60);
                const numberOfLogs = Math.round(totalHours / 12);
                const lastDataRow = SHIFT_START_ROW + numberOfLogs;
                const summaryRow = lastDataRow + 1;

                console.log(`     Shift Summary Row: ${summaryRow}`);

                SHIFT_DATA_POINTS.forEach(point => {
                    const formula = createAbsFormula(RAW_DATA_DIR, shiftFilename, sheetName, `${point.colLetter}${summaryRow}`);
                    shiftFormulas.push({ cell: point.destCell, formula: formula });
                });
            }
        } else {
            console.warn(`   [Warning] Shift Log not found at: ${shiftPath}`);
        }

        // --- STEP C: DOWNTIME LOG (From reports) ---
        if (fs.existsSync(downtimePath)) {
            console.log(`   [Processing] Downtime Log References...`);
            const wb = await XlsxPopulate.fromFileAsync(downtimePath);
            const sheet = wb.sheet(sheetIndex);

            if (sheet) {
                const sheetName = sheet.name();
                DOWNTIME_UNITS.forEach(unit => {
                    const sourceRow = unit.sourceFixedRow;
                    DOWNTIME_METRICS.forEach(metric => {
                        // ** USING REPORTS_DIR HERE **
                        const formula = createAbsFormula(REPORTS_DIR, downtimeFilename, sheetName, `${metric.sourceCol}${sourceRow}`);
                        const destCell = `${unit.destCol}${metric.destRow}`;
                        downtimeFormulas.push({ cell: destCell, formula: formula });
                    });
                });
            }
        } else {
            console.warn(`   [Warning] Downtime Log not found at: ${downtimePath}`);
        }

        // --- STEP D: WRITE SUMMARY REPORT ---
        console.log(`   [Generating] Summary Report...`);

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

        monthlyFormulas.forEach(item => destSheet.cell(item.cell).formula(item.formula));
        shiftFormulas.forEach(item => destSheet.cell(item.cell).formula(item.formula));
        downtimeFormulas.forEach(item => {
            destSheet.cell(item.cell).formula(item.formula).style("numberFormat", "0.00");
        });

        destSheet.cell("H7").value(monthlyStartDate.toLocaleDateString());
        destSheet.cell("K7").value(monthlyEndDate_report.toLocaleDateString());

        //Forebay lookup
        destSheet.cell("L41").formula(`VLOOKUP(L40, '${RAW_DATA_DIR}/[FOREBAY.xlsx]ELEV'!$B$7:$C$137,2,FALSE)`);
        destSheet.cell("L44").formula(`VLOOKUP(L43, '${RAW_DATA_DIR}/[FOREBAY.xlsx]ELEV'!$B$7:$C$137,2,FALSE)`);

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
