const XlsxPopulate = require('xlsx-populate');
const cron = require('node-cron');
const path = require('path');
const fs = require('fs');
const opened = require('@ronomon/opened');

const OUTPUT_DIR = path.join(__dirname, 'rawdata');
const TEMPLATE_PATH = path.join(__dirname, '/templates/Pulangi IV HEP - Generation Data - RAW Template.xlsx');
const BUFFER_FILE = path.join(OUTPUT_DIR, 'data_buffer_shift.json');

// --- HELPER FUNCTIONS ---

function loadBuffer() {
    if (fs.existsSync(BUFFER_FILE)) {
        try {
            return JSON.parse(fs.readFileSync(BUFFER_FILE));
        } catch (e) {
            return [];
        }
    }
    return [];
}

function saveBuffer(data) {
    fs.writeFileSync(BUFFER_FILE, JSON.stringify(data, null, 2));
}

// Promise wrapper for @ronomon/opened
function checkIfFileIsOpen(filePath) {
    return new Promise((resolve, reject) => {
        if (!fs.existsSync(filePath)) return resolve(false);

        opened.files([filePath], (error, hashTable) => {
            if (error) {
                console.error("[CheckOpen] Error:", error);
                return resolve(true); // Assume locked on error
            }
            resolve(hashTable[filePath] === true);
        });
    });
}

// --- THE MERGE TASK ---

async function processShiftBuffer() {
    let buffer = loadBuffer();

    if (buffer.length === 0) {
        return;
    }

    console.log(`\n[Shift Merger] Found ${buffer.length} pending entries.`);

    // Group entries by File Path
    const filesToProcess = {};
    buffer.forEach(entry => {
        if (!filesToProcess[entry.filePath]) filesToProcess[entry.filePath] = [];
        filesToProcess[entry.filePath].push(entry);
    });

    const successfulEntryIds = [];

    for (const [fPath, entries] of Object.entries(filesToProcess)) {

        // 1. Check Lock Status
        const isLocked = await checkIfFileIsOpen(fPath);

        if (isLocked) {
            console.warn(`[Skip] ${path.basename(fPath)} is OPEN by user. Keeping data in buffer.`);
            continue;
        }

        // 2. Open & Write
        try {
            let workbook;

            // ** LOGIC TO CREATE NEW FILE **
            if (fs.existsSync(fPath)) {
                workbook = await XlsxPopulate.fromFileAsync(fPath);
            } else {
                console.log(`[Info] File not found: ${path.basename(fPath)}`);
                console.log(`[Info] Creating new file from template...`);

                if (fs.existsSync(TEMPLATE_PATH)) {
                    workbook = await XlsxPopulate.fromFileAsync(TEMPLATE_PATH);

                    // --- INIT NEW FILE ---
                    if (entries.length > 0) {
                        const targetYear = entries[0].targetYear;
                        // Date: Dec 26, (Year - 1)
                        const startDate = new Date(targetYear - 1, 11, 26);
                        const shiftLogFilename = `Pulangi IV HEP - Generation Data - RAW_${startDate.getFullYear()}.xlsx`;
                        const sheetHeader = `PULANGI IV HE PLANT - ${targetYear} DAILY INDEX READING AND GENERATION`;

                        for (let i = 0; i < 12; i++) {
                            workbook.sheet(i).cell("A1").value(sheetHeader);
                        }

                        // Write formulas to sheet 0 (Jan)
                        //Gross
                        workbook.sheet(0).cell("I5").formula(`C5-'${OUTPUT_DIR}/[${shiftLogFilename}]Dec'!C63`);
                        workbook.sheet(0).cell("I6").formula(`C6-'${OUTPUT_DIR}/[${shiftLogFilename}]Dec'!C64`);

                        workbook.sheet(0).cell("L5").formula(`E5-'${OUTPUT_DIR}/[${shiftLogFilename}]Dec'!E63`);
                        workbook.sheet(0).cell("L6").formula(`E6-'${OUTPUT_DIR}/[${shiftLogFilename}]Dec'!E64`);

                        workbook.sheet(0).cell("O5").formula(`G5-'${OUTPUT_DIR}/[${shiftLogFilename}]Dec'!G63`);
                        workbook.sheet(0).cell("O6").formula(`G6-'${OUTPUT_DIR}/[${shiftLogFilename}]Dec'!G64`);
                        //Stn
                        workbook.sheet(0).cell("J5").formula(`ROUND((D5-'${OUTPUT_DIR}/[${shiftLogFilename}]Dec'!D63)/1000,3)`);
                        workbook.sheet(0).cell("J6").formula(`ROUND((D6-'${OUTPUT_DIR}/[${shiftLogFilename}]Dec'!D64)/1000,3)`);

                        workbook.sheet(0).cell("M5").formula(`ROUND((F5-'${OUTPUT_DIR}/[${shiftLogFilename}]Dec'!F63)/1000,3)`);
                        workbook.sheet(0).cell("M6").formula(`ROUND((F6-'${OUTPUT_DIR}/[${shiftLogFilename}]Dec'!F64)/1000,3)`);

                        workbook.sheet(0).cell("P5").formula(`ROUND((H5-'${OUTPUT_DIR}/[${shiftLogFilename}]Dec'!H63)/1000,3)`);
                        workbook.sheet(0).cell("P6").formula(`ROUND((H6-'${OUTPUT_DIR}/[${shiftLogFilename}]Dec'!H64)/1000,3)`);
                    }

                } else {
                    console.error(`[Error] Template missing at ${TEMPLATE_PATH}`);
                    continue;
                }
            }

            entries.forEach(entry => {
                const sheet = workbook.sheet(entry.sheetIndex);
                if (sheet) {
                    entry.values.forEach((val, i) => {
                        // Formula: Col C (3), E (5), G (7) -> 2 * (i+1) + 1
                        const targetCol = 2 * (i + 1) + 1;
                        sheet.row(entry.targetRow).cell(targetCol).value(val);
                    });
                    // Mark as successful
                    successfulEntryIds.push(entry.id);
                } else {
                    console.error(`[Error] Sheet index ${entry.sheetIndex} missing in ${path.basename(fPath)}`);
                }
            });

            await workbook.toFileAsync(fPath);
            console.log(`[Success] Written ${entries.length} rows to ${path.basename(fPath)}`);

        } catch (err) {
            console.error(`[Error] Failed writing to ${path.basename(fPath)}:`, err.message);
        }
    }

    // 3. Cleanup Buffer
    const currentBuffer = loadBuffer();
    const remainingBuffer = currentBuffer.filter(item => !successfulEntryIds.includes(item.id));

    if (remainingBuffer.length !== currentBuffer.length) {
        saveBuffer(remainingBuffer);
        console.log(`[Clean] Buffer updated. Remaining entries: ${remainingBuffer.length}`);
    }
}

// Run immediately
processShiftBuffer();

// Schedule: Run every minute
cron.schedule('*/2 * * * *', () => processShiftBuffer());
