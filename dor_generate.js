const XlsxPopulate = require('xlsx-populate');
const path = require('path');
const fs = require('fs');
const cron = require('node-cron');
const { checkIfFileIsOpen, isFileLockError } = require('./lib/file-lock');
const jobStatus = require('./lib/job-status');

const JOB = 'dor_generate';

// Set while a freshly-built report couldn't be saved because the previous
// version of the file is open in Excel. The retry cron below keeps
// re-running the full generation (cheap - it's all formulas, no manual
// edits to lose) until the file is free, so the report still updates
// instead of silently going stale.
let pendingRegeneration = false;

// --- CONFIGURATION ---
const REPORTS_DIR = path.join(__dirname, 'reports');
// Source files are in 'rawdata'
const MONTHLY_LOG_DIR = path.join(__dirname, 'rawdata');
const DAILY_TEMPLATE_PATH = path.join(__dirname, 'templates', 'Pulangi IV HEP - Daily Operations Report Template.xlsx');
// Path to the Daily Outage Log JSON
const DAILY_OUTAGE_LOG_PATH = path.join(REPORTS_DIR, 'daily_outage_log.json');

// Create output directory if it doesn't exist
if (!fs.existsSync(REPORTS_DIR)) fs.mkdirSync(REPORTS_DIR);

// --- HELPER FUNCTIONS ---

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

// Creates formula with full absolute path
// Returns: 'C:\Path\To\rawdata\[File.xlsx]Sheet'!Cell
function createAbsFormula(directory, filename, sheetName, cellRef) {
    return `'${directory}/[${filename}]${sheetName}'!${cellRef}`;
}

// Formats date to dd-Mmm-yyyy (e.g., 26-Jan-2026)
function formatExcelDate(date) {
    return date.toLocaleDateString('en-GB', {
        day: '2-digit',
        month: 'short',
        year: 'numeric'
    }).replace(/ /g, '-');
}

// Convert "14:30:00" to "02:30 PM"
function formatTimeAMPM(time24) {
    if (!time24) return null;
    const [hours, minutes] = time24.split(':');
    let h = parseInt(hours, 10);
    const m = minutes;
    const ampm = h >= 12 ? 'PM' : 'AM';
    h = h % 12;
    h = h ? h : 12; // the hour '0' should be '12'
    return `${h.toString().padStart(2, '0')}:${m} ${ampm}`;
}

// Read Outage Logs from JSON and filter by Date
function getLatestOutageTimes(targetDateStr) {
    if (!fs.existsSync(DAILY_OUTAGE_LOG_PATH)) return {};

    try {
        const rawData = fs.readFileSync(DAILY_OUTAGE_LOG_PATH);
        const logs = JSON.parse(rawData);

        const unitTimes = {
            1: { out: null, in: null },
            2: { out: null, in: null },
            3: { out: null, in: null }
        };

        // Filter for events matching the TARGET DATE
        logs.forEach(log => {
            if (log.date === targetDateStr && unitTimes[log.unit]) {
                if (log.type === 'OUT') {
                    unitTimes[log.unit].out = formatTimeAMPM(log.time);
                } else if (log.type === 'IN') {
                    unitTimes[log.unit].in = formatTimeAMPM(log.time);
                }
            }
        });

        return unitTimes;
    } catch (e) {
        console.error("[Error] Failed to read outage log:", e.message);
        return {};
    }
}

// --- THE TASK ---

async function generateDailyReport() {
    console.log(`[Daily Report] Starting generation...`);
    jobStatus.recordEvent(JOB, 'run-start', 'Building daily report', {});

    const today = new Date();
    const targetDate = new Date(today);
    targetDate.setDate(today.getDate() - 1);

    // Keep en-CA (YYYY-MM-DD) strictly for querying the JSON and generating the File Name
    const dateStr = targetDate.toLocaleDateString('en-CA');
    const outputFilename = `Pulangi IV HEP - Daily Operations Report_${dateStr}.xlsx`;
    const outputPath = path.join(REPORTS_DIR, outputFilename);

    console.log(`[Daily Report] Linking to Monthly Log for Date: ${dateStr}`);

    try {
        // 1. Identify Source File Context
        const { sheetIndex, targetYear } = getSheetContext(targetDate);
        const sourceFilename = `Pulangi IV HEP - Operational Highlights - RAW_${targetYear}.xlsx`;
        // Reads from rawdata folder
        const sourcePath = path.join(MONTHLY_LOG_DIR, sourceFilename);

        if (!fs.existsSync(sourcePath)) {
            console.error(`[Error] Source file missing: ${sourcePath}`);
            return;
        }

        // 2. Open Source (Just to get the Sheet Name)
        const sourceWorkbook = await XlsxPopulate.fromFileAsync(sourcePath);
        const sourceSheet = sourceWorkbook.sheet(sheetIndex);

        if (!sourceSheet) {
            console.error(`[Error] Sheet index ${sheetIndex} missing in source.`);
            return;
        }

        const sourceSheetName = sourceSheet.name();

        // 3. Open Destination Template
        if (!fs.existsSync(DAILY_TEMPLATE_PATH)) {
            console.error(`[Error] Template file missing: ${DAILY_TEMPLATE_PATH}`);
            return;
        }
        const destWorkbook = await XlsxPopulate.fromFileAsync(DAILY_TEMPLATE_PATH);
        const destSheet = destWorkbook.sheet(0);

        // 4. Calculate Source Row Offset
        const cycleStartDate = new Date(targetYear, sheetIndex - 1, 26, 0, 0, 0);
        const reportStartTime = new Date(targetDate);
        reportStartTime.setHours(1, 0, 0, 0);

        const diffMs = reportStartTime - cycleStartDate;
        const diffHours = Math.round(diffMs / (1000 * 60 * 60));

        let sourceRow = 4 + diffHours; // Row in Monthly Log
        let sourceRowSpillage = 4 + diffHours; //Row in Monthly Log
        let destRow = 17;              // Row in Daily Report

        // 5. Generate Formulas Loop (24 Hours)
        for (let i = 0; i < 24; i++) {

            // --- Construct Formulas with FULL PATH ---
            const formulaB = createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `C${sourceRow}`);
            const formulaC = createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `G${sourceRow}`);
            const formulaD = createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `K${sourceRow}`);
            const formulaF = createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `Q${sourceRow}`);
            const formulaG = createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `S${sourceRow}`);
            const formulaH = createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `V${sourceRow}`);
            const formulaI = createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `Y${sourceRow}`);
            const formulaJ = createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `AB${sourceRow}`);
            const formulaK = createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `AE${sourceRow}`);
            const formulaL = createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `AH${sourceRow}`);
            const formulaM = createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `AK${sourceRow}`);
            const formulaN = createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `AN${sourceRow}`);
            const formulaO = createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `AQ${sourceRow}`);
            const formulaP = createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `AT${sourceRow}`);

            // Column E (Daily) is the SUM of the local columns B, C, D
            const formulaSum = `SUM(B${destRow}:D${destRow})`;

            // --- Apply Formulas & Styles ---

            destSheet.row(destRow).cell(2).formula(formulaB).style("numberFormat", "0.00");
            destSheet.row(destRow).cell(3).formula(formulaC).style("numberFormat", "0.00");
            destSheet.row(destRow).cell(4).formula(formulaD).style("numberFormat", "0.00");
            destSheet.row(destRow).cell(5).formula(formulaSum).style("numberFormat", "0.00");
            destSheet.row(destRow).cell(6).formula(formulaF).style("numberFormat", "0.00");
            destSheet.row(destRow).cell(7).formula(formulaG).style("numberFormat", "0.00");
            destSheet.row(destRow).cell(8).formula(formulaH).style("numberFormat", "0.00");
            destSheet.row(destRow).cell(9).formula(formulaI).style("numberFormat", "0.00");
            destSheet.row(destRow).cell(10).formula(formulaJ).style("numberFormat", "0.00");
            destSheet.row(destRow).cell(11).formula(formulaK).style("numberFormat", "0.00");
            destSheet.row(destRow).cell(12).formula(formulaL).style("numberFormat", "0.00");
            destSheet.row(destRow).cell(13).formula(formulaM).style("numberFormat", "0.00");
            destSheet.row(destRow).cell(14).formula(formulaN).style("numberFormat", "0.00");
            destSheet.row(destRow).cell(15).formula(formulaO).style("numberFormat", "0.00");
            destSheet.row(destRow).cell(16).formula(formulaP).style("numberFormat", "0.00");

            sourceRow++;
            destRow++;
        }
        //H2O SPILLAGE
        sourceRow--;
        const cellSpillage = `AW${sourceRowSpillage}:AW${sourceRow}`;
        destSheet.cell("Q47").formula(`SUM('${MONTHLY_LOG_DIR}/[${sourceFilename}]${sourceSheetName}'!${cellSpillage})`);

        // ============================================================
        //  NEW SECTION: OUTAGE LOGS (IN/OUT) for PREVIOUS DAY
        // ============================================================
        console.log(`[Daily Report] Processing Outage Logs for ${dateStr}...`);

        // Read JSON and match against today's target date string
        const outageTimes = getLatestOutageTimes(dateStr);

        // Unit 1
        if (outageTimes[1]) {
            if (outageTimes[1].out) destSheet.cell("H44").value(outageTimes[1].out);
            if (outageTimes[1].in) destSheet.cell("J44").value(outageTimes[1].in);
        }
        // Unit 2
        if (outageTimes[2]) {
            if (outageTimes[2].out) destSheet.cell("H45").value(outageTimes[2].out);
            if (outageTimes[2].in) destSheet.cell("J45").value(outageTimes[2].in);
        }
        // Unit 3
        if (outageTimes[3]) {
            if (outageTimes[3].out) destSheet.cell("H46").value(outageTimes[3].out);
            if (outageTimes[3].in) destSheet.cell("J46").value(outageTimes[3].in);
        }
        console.log(`[Success] Outage times populated from JSON.`);

        // ============================================================
        //  NEW SECTION: REFERENCE DATA FROM SHIFT LOGGER (12AM & 12NN)
        // ============================================================

        const shiftLogFilename = `Pulangi IV HEP - Generation Data - RAW_${targetYear}.xlsx`;
        // Reads from rawdata folder
        const shiftLogPath = path.join(MONTHLY_LOG_DIR, shiftLogFilename);

        if (fs.existsSync(shiftLogPath)) {
            console.log(`[Daily Report] Processing Shift Log References...`);

            // Open Shift Log to get the specific Sheet Name
            const shiftWorkbook = await XlsxPopulate.fromFileAsync(shiftLogPath);
            const shiftSheet = shiftWorkbook.sheet(sheetIndex);

            if (shiftSheet) {
                const shiftSheetName = shiftSheet.name();

                // 1. Determine Cycle Start for Shift Log
                const shiftCycleStart = new Date(targetYear, sheetIndex - 1, 26, 12, 0, 0);

                // 2. Calculate Row for 12:00 AM (Midnight)
                const midnightDate = new Date(targetDate);
                midnightDate.setDate(midnightDate.getDate() + 1)
                midnightDate.setHours(0, 0, 0, 0);

                const diffMsMidnight = midnightDate - shiftCycleStart;
                const diffShiftsMidnight = Math.round(diffMsMidnight / (1000 * 60 * 60 * 12));
                const midnightRow = 5 + diffShiftsMidnight;

                // 3. Calculate Row for 12:00 PM (Noon)
                const noonDate = new Date(targetDate);
                noonDate.setHours(12, 0, 0, 0);

                const diffMsNoon = noonDate - shiftCycleStart;
                const diffShiftsNoon = Math.round(diffMsNoon / (1000 * 60 * 60 * 12));
                const noonRow = 5 + diffShiftsNoon;

                // 4. Inject Formulas with FULL PATH

                // -- 12:00 AM DATA --
                destSheet.cell("D44").formula(createAbsFormula(MONTHLY_LOG_DIR, shiftLogFilename, shiftSheetName, `I${midnightRow}`));
                destSheet.cell("D45").formula(createAbsFormula(MONTHLY_LOG_DIR, shiftLogFilename, shiftSheetName, `L${midnightRow}`));
                destSheet.cell("D46").formula(createAbsFormula(MONTHLY_LOG_DIR, shiftLogFilename, shiftSheetName, `O${midnightRow}`));
                destSheet.cell("B50").formula(createAbsFormula(MONTHLY_LOG_DIR, shiftLogFilename, shiftSheetName, `J${midnightRow}`));
                destSheet.cell("C50").formula(createAbsFormula(MONTHLY_LOG_DIR, shiftLogFilename, shiftSheetName, `M${midnightRow}`));
                destSheet.cell("D50").formula(createAbsFormula(MONTHLY_LOG_DIR, shiftLogFilename, shiftSheetName, `P${midnightRow}`));

                // -- 12:00 PM DATA --
                destSheet.cell("E44").formula(createAbsFormula(MONTHLY_LOG_DIR, shiftLogFilename, shiftSheetName, `I${noonRow}`));
                destSheet.cell("E45").formula(createAbsFormula(MONTHLY_LOG_DIR, shiftLogFilename, shiftSheetName, `L${noonRow}`));
                destSheet.cell("E46").formula(createAbsFormula(MONTHLY_LOG_DIR, shiftLogFilename, shiftSheetName, `O${noonRow}`));

                // ==============================
                // Write formatted dates to cells
                // ==============================
                const reportDateFormatted = formatExcelDate(targetDate);
                const todayFormatted = formatExcelDate(today);

                destSheet.cell("Q5").value(todayFormatted);
                destSheet.cell("H13").value(reportDateFormatted);
                destSheet.cell("G54").value(todayFormatted);

                console.log(`[Success] Added Shift References for Rows ${midnightRow} (12AM) and ${noonRow} (12NN)`);
            } else {
                console.warn(`[Warning] Sheet index ${sheetIndex} missing in Shift Log.`);
            }
        } else {
            console.warn(`[Warning] Shift Log file not found: ${shiftLogPath}`);
        }

        // ============================================================
        //  END NEW SECTION
        // ============================================================

        // ============================================================
        //  NEW SECTION: CURRENT DAY 7:00 AM REFERENCE
        // ============================================================

        console.log(`[Daily Report] Processing Current Day 7AM Reference...`);

        // 1. Calculate Target: TODAY at 7:00 AM
        const current7AM = new Date(today);
        current7AM.setHours(7, 0, 0, 0);

        // 2. Calculate Row in the MONTHLY LOG (Hourly)
        const diffMs7AM = current7AM - cycleStartDate;
        const diffHours7AM = Math.round(diffMs7AM / (1000 * 60 * 60));
        const row7AM = 4 + diffHours7AM;

        // 3. Inject Formula with FULL PATH
        destSheet.cell("F60").formula(createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `C${row7AM}`));
        destSheet.cell("F61").formula(createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `G${row7AM}`));
        destSheet.cell("F62").formula(createAbsFormula(MONTHLY_LOG_DIR, sourceFilename, sourceSheetName, `K${row7AM}`));

        console.log(`[Success] Added 7AM Reference pointing to Row ${row7AM}`);

        // ============================================================
        //  END NEW SECTION
        // ============================================================

        // --- SAVE (won't clobber a copy the user has open) ---
        // Check the lock first so we don't even attempt the write while
        // someone has the report open; also catch EBUSY/EPERM/EACCES from
        // toFileAsync itself in case it got opened in the gap between the
        // check and the write. Either way we don't lose the freshly built
        // report - we just flag it for the retry cron to try again once
        // the file is free.
        const isLocked = await checkIfFileIsOpen(outputPath);

        if (isLocked) {
            pendingRegeneration = true;
            console.warn(`[Skip] ${outputFilename} is OPEN by user. Will retry.`);
            jobStatus.recordEvent(JOB, 'skip-locked', `${outputFilename} is open - retry queued`, { outputFilename });
            return;
        }

        try {
            await destWorkbook.toFileAsync(outputPath);
        } catch (writeErr) {
            if (isFileLockError(writeErr)) {
                pendingRegeneration = true;
                console.warn(`[Skip] ${outputFilename} became locked mid-write. Will retry.`);
                jobStatus.recordEvent(JOB, 'skip-locked', `${outputFilename} locked mid-write - retry queued`, { outputFilename });
                return;
            }
            throw writeErr;
        }

        pendingRegeneration = false;
        console.log(`[Success] Generated with Formulas: ${outputFilename}`);
        jobStatus.recordEvent(JOB, 'success', `Generated ${outputFilename}`, { outputFilename, targetDate: dateStr });

    } catch (error) {
        console.error('[Error] Daily Report failed:', error);
        jobStatus.recordEvent(JOB, 'error', error.message, {});
    }
}

// --- SCHEDULE ---
cron.schedule('1 0,7 * * *', () => {
    generateDailyReport();
});

// Retry loop: as long as the last attempt was blocked by the report being
// open, keep re-generating every 2 minutes so the report catches up the
// moment the user closes it, instead of waiting for the next 00:01/07:01
// run.
cron.schedule('*/2 * * * *', () => {
    if (pendingRegeneration) generateDailyReport();
});

generateDailyReport();
