const XlsxPopulate = require('xlsx-populate');
const cron = require('node-cron');
const path = require('path');
const fs = require('fs');
const opened = require('@ronomon/opened');

const OUTPUT_DIR = path.join(__dirname, 'rawdata');
const TEMPLATE_PATH = path.join(__dirname, '/templates/Pulangi IV HEP - Operational Highlights - RAW Template.xlsx');
const BUFFER_FILE = path.join(OUTPUT_DIR, 'data_buffer.json');

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
        // If file doesn't exist, it cannot be open/locked by user
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

async function processBuffer() {
    let buffer = loadBuffer();

    if (buffer.length === 0) {
        return;
    }

    console.log(`\n[Merger] Found ${buffer.length} pending entries.`);

    // Group entries by File Path to process one file at a time
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
            continue; // Skip this file, try others
        }

        // 2. Open & Write
        try {
            let workbook;

            // ** LOGIC TO CREATE NEW FILE **
            if (fs.existsSync(fPath)) {
                // Load existing file
                workbook = await XlsxPopulate.fromFileAsync(fPath);
            } else {
                // File does not exist -> Create from Template
                console.log(`[Info] File not found: ${path.basename(fPath)}`);
                console.log(`[Info] Creating new file from template...`);

                if (fs.existsSync(TEMPLATE_PATH)) {
                    workbook = await XlsxPopulate.fromFileAsync(TEMPLATE_PATH);

                    // --- NEW FEATURE: Write Start Date to A4 ---
                    // Since we are grouping by file, all entries here share the same targetYear.
                    // We take the year from the first entry.
                    if (entries.length > 0) {
                        const targetYear = entries[0].targetYear;
                        // Date: Dec 26, (Year - 1)
                        const startDate = new Date(targetYear - 1, 11, 26);

                        // Get the first sheet (January) and write to A4
                        // Format: "26-Dec-YY" (d-mmm-yy)
                        workbook.sheet(0).cell("A4")
                            .value(startDate)
                            .style("numberFormat", "d-mmm-yy");

                        console.log(`[Init] Set Cell A4 to ${startDate.toDateString()} (Format: 26-Dec-YY)`);
                    }

                } else {
                    console.error(`[Error] Template missing at ${TEMPLATE_PATH}`);
                    continue; // Skip if template is missing
                }
            }

            entries.forEach(entry => {
                const sheet = workbook.sheet(entry.sheetIndex);
                if (sheet) {
                    entry.values.forEach((val, i) => {
                        const cell = sheet.row(entry.targetRow).cell(i + 3);
                        cell.value(val);
                        if (val !== '') cell.style("numberFormat", "0.00");
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
            // Do NOT mark as successful, so they stay in buffer
        }
    }

    // 3. Cleanup Buffer
    // Reload buffer to be safe (in case Logger wrote new data while we were merging)
    const currentBuffer = loadBuffer();

    // Filter out ONLY the entries we successfully wrote (using the unique ID)
    const remainingBuffer = currentBuffer.filter(item => !successfulEntryIds.includes(item.id));

    if (remainingBuffer.length !== currentBuffer.length) {
        saveBuffer(remainingBuffer);
        console.log(`[Clean] Buffer updated. Remaining entries: ${remainingBuffer.length}`);
    }
}

// Run immediately on start (to clear backlog)
processBuffer();

// Schedule: Run every 2nd minute
cron.schedule('*/2 * * * *', () => processBuffer());
