const express = require('express');
const mysql = require('mysql2/promise');
const XlsxPopulate = require('xlsx-populate');
const cron = require('node-cron');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Serve the 'reports' folder so users can download files
// Users can access: http://localhost:8000/reports/Pulangi IV HEP - Operational Highlights - RAW_2025.xlsx
// or http://localhost:8000/reports/Pulangi IV HEP - Operational Highlights - RAW_2026.xlsx
app.use('/reports', express.static(path.join(__dirname, 'reports')));

const dbConfig = {
    host: 'localhost',
    user: 'root',
    password: 'KaRMSys2025!',
    database: 'pulangi_data',
    connectionLimit: 10
};

const OUTPUT_DIR = path.join(__dirname, 'reports');
const TEMPLATE_PATH = path.join(__dirname, '/reports/Pulangi IV HEP - Operational Highlights - RAW Template.xlsx');

// Ensure output directory exists
if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR);
}

// --- HELPER FUNCTIONS ---

const formatDate = (date) => date.toLocaleDateString('en-CA');
const formatTime = (date) => date.toTimeString().split(' ')[0].substring(0, 5);
const createLookupKey = (date) => `${formatDate(date)} ${formatTime(date)}`;
// Helper to get "HH:00:00" for the database query
const formatQueryTime = (date) => {
    const hh = String(date.getHours()).padStart(2, '0');
    return `${hh}:00:00`;
};
// Helper for the Excel file display (HH:mm)
const formatDisplayTime = (date) => {
    return date.toTimeString().split(' ')[0].substring(0, 5);
};

// Helper: Convert value to Number and fix to 2 decimal places
// Returns a Number type (e.g. 10.50) so Excel treats it as a number, not text.
const toFixedNumber = (val) => {
    if (val === null || val === undefined || val === '') return '';
    return Number(Number(val).toFixed(2));
};

// Logic: If date >= 26th, it belongs to the NEXT month's sheet.
// If Dec 26, 2025 -> It belongs to Jan 2026 Cycle -> Target Year becomes 2026
function getSheetContext(date) {
    const day = date.getDate();
    const month = date.getMonth(); // 0 = Jan, 11 = Dec
    const year = date.getFullYear();

    if (day >= 26) {
        // If Dec 26, it is Sheet 0 (Jan) of Next Year
        if (month === 11) return { sheetIndex: 0, targetYear: year + 1 };
        return { sheetIndex: month + 1, targetYear: year };
    } else {
        return { sheetIndex: month, targetYear: year };
    }
}

// --- THE UPDATE TASK ---

async function appendCurrentHourData() {
    const now = new Date();

    // 1. Prepare Query Parameters
    const currentDateStr = formatDate(now);       // e.g. "2025-12-19"
    const currentTimeStr = formatQueryTime(now);  // e.g. "14:00:00"

    console.log(`[Update] Querying DB for Date: ${currentDateStr}, Time: ${currentTimeStr}`);
    let connection;

    try {
        // 1. Get File Context
        const { sheetIndex, targetYear } = getSheetContext(now);
        const fileName = `Pulangi IV HEP - Operational Highlights - RAW_${targetYear}.xlsx`;
        const filePath = path.join(OUTPUT_DIR, fileName);

        // 2. Calculate Target Row
        const cycleStartDate = new Date(targetYear, sheetIndex - 1, 26, 0, 0, 0);
        const diffMs = now - cycleStartDate;
        const diffHours = Math.round(diffMs / (1000 * 60 * 60));

        // Safety check: ensure positive index
        if (diffHours < 0) {
            console.warn("Calculated time is before cycle start. Skipping.");
            return;
        }

        const targetRow = 4 + diffHours;

        // 3. The Query (Separated Columns)
        // Adjust 'date_column' and 'time_column' to your actual DB column names
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

        connection = await mysql.createConnection(dbConfig);

        // We pass the Date String and Time String separately
        const [rows] = await connection.execute(query, [currentDateStr, currentTimeStr]);

        const rawValues = rows.length > 0 ? [
            rows[0].mw1, rows[0].v1, rows[0].pf1, rows[0].mvar1,
            rows[0].mw2, rows[0].v2, rows[0].pf2, rows[0].mvar2,
            rows[0].mw3, rows[0].v3, rows[0].pf3, rows[0].mvar3,
            rows[0].mw, rows[0].mvar
        ] : new Array(14).fill('');

        const formattedValues = rawValues.map(val => toFixedNumber(val));

        // 4. Load & Write to File
        let workbook;
        if (fs.existsSync(filePath)) {
            workbook = await XlsxPopulate.fromFileAsync(filePath);
        } else {
            console.log(`Creating new file for ${targetYear}...`);
            workbook = await XlsxPopulate.fromFileAsync(TEMPLATE_PATH);
        }

        const sheet = workbook.sheet(sheetIndex);
        if (!sheet) {
            console.error(`Sheet index ${sheetIndex} missing.`);
            return;
        }

        // 5. Write ONLY the Data (Columns C - P)
        // We skip Column 1 (Date) and Column 2 (Time) entirely.

        // 6. Write Data (Columns C - P)
        formattedValues.forEach((val, i) => {
            const cell = sheet.row(targetRow).cell(i + 3);

            // Write the value
            cell.value(val);

            // OPTIONAL: Force Excel Display Format (0.00) 
            // This ensures 10 is displayed as "10.00"
            if (val !== '') {
                cell.style("numberFormat", "0.00");
            }
        });

        await workbook.toFileAsync(filePath);
        console.log(`[Success] Updated Sheet ${sheetIndex}, Row ${targetRow}`);

    } catch (error) {
        console.error('[Error] Update failed:', error);
    } finally {
        if (connection) await connection.end();
    }
}

// Run every hour at minute 1
cron.schedule('1 * * * *', () => appendCurrentHourData());

// Run once on start
appendCurrentHourData();

app.listen(PORT, () => {
    console.log(`Server running.`);

});
