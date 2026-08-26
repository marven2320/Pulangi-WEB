const XlsxPopulate = require('xlsx-populate');
const path = require('path');
const fs = require('fs');

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

const MONTHLY_DATA_POINTS = [
    { colLetter: "C", destCellMax: "J14", destCellMin: "J15" },
    { colLetter: "G", destCellMax: "K14", destCellMin: "K15" },
    { colLetter: "K", destCellMax: "L14", destCellMin: "L15" }
];

const MONTHLY_COL_Q = {
    colLetter: "Q",
    destCellMax: "L43",
    destCellMin: "L40"
};

const MONTHLY_COL_AW = {
    colLetter: "AW",
    destCellAvg: "L46",
    destCellTotal: "L47"
};

// ==========================================
// 2. SHIFT LOG (Source: rawdata)
// ==========================================
const getShiftLogFilename = (year) => `Pulangi IV HEP - Generation Data - RAW_${year}.xlsx`;
const SHIFT_START_ROW = 5;

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

// Generates context based on an EXPLICIT start year and month
function createCycleContext(startYear, startMonth) {
    let sheetIndex, fileYear;
    
    // If start month is December (11), it maps to the January (0) sheet of the NEXT year
    if (startMonth === 11) {
        sheetIndex = 0; 
        fileYear = startYear + 1;
    } else {
        sheetIndex = startMonth + 1; // e.g., Jan (0) maps to Sheet index 1
        fileYear = startYear;
    }

    // Start: 26th of Start Month
    const monthlyStartDate = new Date(startYear, startMonth, 26, 0, 0, 0);
    // End: 26th of Next Month
    const monthlyEndDate = new Date(startYear, startMonth + 1, 26, 0, 0, 0);
    // Report End: 25th of Next Month
    const monthlyEndDate_report = new Date(startYear, startMonth + 1, 25, 0, 0, 0);

    const shiftStartDate = new Date(startYear, startMonth, 26, 12, 0, 0);
    const shiftEndDate = new Date(startYear, startMonth + 1, 26, 0, 0, 0);

    console.log(`\n   [Context] Target Cycle: ${monthlyStartDate.toDateString()} to ${monthlyEndDate.toDateString()}`);
    console.log(`   [Context] Target File Year: ${fileYear} | Sheet Index: ${sheetIndex}`);

    return {
        sheetIndex, fileYear,
        monthlyStartDate, monthlyEndDate, monthlyEndDate_report,
        shiftStartDate, shiftEndDate
    };
}

function createAbsFormula(directory, filename, sheetName, cellRef) {
    return `'${directory}/[${filename}]${sheetName}'!${cellRef}`;
}

// --- CORE TASK ---

async function generateMonthlySummary(context) {
    const {
        sheetIndex, fileYear,
        monthlyStartDate, monthlyEndDate, monthlyEndDate_report,
        shiftStartDate, shiftEndDate
    } = context;

    const monthlyFilename = getMonthlyLogFilename(fileYear);
    const shiftFilename = getShiftLogFilename(fileYear);
    const downtimeFilename = getDowntimeLogFilename(fileYear);

    const monthlyPath = path.join(RAW_DATA_DIR, monthlyFilename);
    const shiftPath = path.join(RAW_DATA_DIR, shiftFilename);
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
                const summaryMinRow = dataEndRow + 3;

                MONTHLY_DATA_POINTS.forEach(point => {
                    const fMax = createAbsFormula(RAW_DATA_DIR, monthlyFilename, sheetName, `${point.colLetter}${summaryMaxRow}`);
                    const fMin = createAbsFormula(RAW_DATA_DIR, monthlyFilename, sheetName, `${point.colLetter}${summaryMinRow}`);
                    monthlyFormulas.push({ cell: point.destCellMax, formula: fMax });
                    monthlyFormulas.push({ cell: point.destCellMin, formula: fMin });
                });

                const fQMin = createAbsFormula(RAW_DATA_DIR, monthlyFilename, sheetName, `${MONTHLY_COL_Q.colLetter}${MONTHLY_START_ROW + 1}`);
                const fQMax = createAbsFormula(RAW_DATA_DIR, monthlyFilename, sheetName, `${MONTHLY_COL_Q.colLetter}${dataEndRow}`);
                monthlyFormulas.push({ cell: MONTHLY_COL_Q.destCellMin, formula: fQMin });
                monthlyFormulas.push({ cell: MONTHLY_COL_Q.destCellMax, formula: fQMax });

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
                const summaryRow = lastDataRow + 2;

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
        if (fs.existsSync(summaryPath)) {
            console.log(`   [Info] Output file exists. Updating: ${summaryFilename}`);
            destWb = await XlsxPopulate.fromFileAsync(summaryPath);
        } else if (fs.existsSync(SUMMARY_TEMPLATE)) {
            console.log(`   [Info] Creating new file from template: ${summaryFilename}`);
            destWb = await XlsxPopulate.fromFileAsync(SUMMARY_TEMPLATE);
        } else {
            console.error(`   [Error] Template file missing: ${SUMMARY_TEMPLATE}`);
            return;
        }

        const destSheet = destWb.sheet(sheetIndex);
        if (!destSheet) {
            console.error(`   [Error] Sheet index ${sheetIndex} does not exist in the workbook.`);
            return;
        }

        console.log(`   Writing to Destination Sheet: Index ${sheetIndex} (${destSheet.name()})`);

        monthlyFormulas.forEach(item => destSheet.cell(item.cell).formula(item.formula));
        shiftFormulas.forEach(item => destSheet.cell(item.cell).formula(item.formula));
        downtimeFormulas.forEach(item => {
            destSheet.cell(item.cell).formula(item.formula).style("numberFormat", "0.00");
        });

        // Write Dates
        destSheet.cell("H7").value(monthlyStartDate.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: '2-digit' }).replace(/ /g, '-'));
        destSheet.cell("K7").value(monthlyEndDate_report.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: '2-digit' }).replace(/ /g, '-'));

        //Forebay lookup
        destSheet.cell("L41").formula(`VLOOKUP(L40, '${RAW_DATA_DIR}/[FOREBAY.xlsx]ELEV'!$B$7:$C$137,2,FALSE)`);
        destSheet.cell("L44").formula(`VLOOKUP(L43, '${RAW_DATA_DIR}/[FOREBAY.xlsx]ELEV'!$B$7:$C$137,2,FALSE)`);

        await destWb.toFileAsync(summaryPath);
        console.log(`   [Success] Report saved to ${summaryPath}`);

    } catch (error) {
        console.error(`   [Error]`, error);
    }
}

// --- RECOVERY EXECUTION ---
async function recoverData() {
    console.log("==========================================");
    console.log(" STARTING DATA RECOVERY");
    console.log("==========================================");

    // 1. Establish the starting point (December 2025)
    let currentLoopYear = 2025;
    let currentLoopMonth = 11; 

    // 2. Determine the target "Current Cycle" based on today's date
    const today = new Date();
    let targetYear = today.getFullYear();
    let targetMonth = today.getMonth();

    if (today.getDate() < 26) {
        targetMonth--;
        if (targetMonth < 0) {
            targetMonth = 11;
            targetYear--;
        }
    }

    let cycleNum = 1;

    // 3. Loop from start date until we hit the target date
    while (currentLoopYear < targetYear || (currentLoopYear === targetYear && currentLoopMonth <= targetMonth)) {
        console.log(`\n>>> RUNNING RECOVERY FOR CYCLE ${cycleNum}...`);
        
        const context = createCycleContext(currentLoopYear, currentLoopMonth);
        await generateMonthlySummary(context);
        
        // Increment month and wrap around the year
        currentLoopMonth++;
        if (currentLoopMonth > 11) {
            currentLoopMonth = 0;
            currentLoopYear++;
        }
        
        cycleNum++;
    }

    console.log("\n==========================================");
    console.log(" DATA RECOVERY COMPLETE");
    console.log("==========================================");
}

// Execute immediately
recoverData();
