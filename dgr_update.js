const express = require('express');
const mysql = require('mysql2/promise');
const XlsxPopulate = require('xlsx-populate');
const cron = require('node-cron');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3001;

// --- CONFIGURATION ---
const OUTPUT_DIR = path.join(__dirname, 'reports');
const TEMPLATE_PATH = path.join(__dirname, '/reports/Pulangi IV HEP - Generation Data - RAW Template.xlsx');
const FILE_PREFIX = 'Pulangi IV HEP - Generation Data - RAW_';

if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR);

const dbConfig = {
    host: 'localhost',
    user: 'root',
    password: 'KaRMSys2025!',
    database: 'pulangi_data',
    connectionLimit: 10
};
// --- HELPER FUNCTIONS ---

const formatSQLDate = (date) => date.toLocaleDateString('en-CA');
const formatSQLTime = (date) => {
    // Returns "00:00:00" (Midnight) or "12:00:00" (Noon)
    const hours = date.getHours();

    if (hours < 12) {
        return "00:00:00";
    } else {
        return "12:00:00";
    }
};

const toInteger = (val) => {
    if (val === null || val === undefined || val === '') return 0;
    return Math.round(Number(val));
};

function getSheetContext(date) {
    const day = date.getDate();
    const hour = date.getHours();
    const month = date.getMonth();
    const year = date.getFullYear();

    // ** UPDATED LOGIC **
    // The cycle switches to the next month on the 26th at 12:00 PM (Hour >= 12).
    // Before 12:00 PM on the 26th, it still belongs to the previous month.
    const isNextCycle = (day > 26) || (day === 26 && hour >= 12);

    if (isNextCycle) {
        if (month === 11) return { sheetIndex: 0, targetYear: year + 1 };
        return { sheetIndex: month + 1, targetYear: year };
    } else {
        return { sheetIndex: month, targetYear: year };
    }
}

// --- THE TASK ---

async function updateTwiceDaily() {
    const now = new Date();
    now.setMinutes(0, 0, 0);

    // 1. Prepare SQL Query Parameters
    const sqlDateStr = formatSQLDate(now);
    const sqlTimeStr = formatSQLTime(now);

    console.log(`[Shift Update] Targeting: ${sqlDateStr} ${sqlTimeStr}`);

    let connection;

    try {
        const { sheetIndex, targetYear } = getSheetContext(now);
        const fileName = `${FILE_PREFIX}${targetYear}.xlsx`;
        const filePath = path.join(OUTPUT_DIR, fileName);

        // 2. Calculate Target Row
        // ** UPDATED START TIME **
        // Cycle starts on the 26th at 12:00 PM (12, 0, 0)
        const cycleStartDate = new Date(targetYear, sheetIndex - 1, 26, 12, 0, 0);

        const diffMs = now - cycleStartDate;

        // Count 12-hour shifts elapsed since 26th 12:00 PM
        const diffShifts = Math.floor(diffMs / (1000 * 60 * 60 * 12));

        if (diffShifts < 0) {
            console.warn("Calculated time is before cycle start. Skipping.");
            return;
        }

        // Row 5 = 1st Shift (26th 12:00 PM)
        // Row 6 = 2nd Shift (27th 12:00 AM)
        const targetRow = 5 + diffShifts;

        // 3. Query Database
        const query = `
            SELECT 
                energy1, energy2, energy3
            FROM pulangi 
            WHERE date = ? AND time = ? 
            LIMIT 1
        `;

        connection = await mysql.createConnection(dbConfig);
        const [rows] = await connection.execute(query, [sqlDateStr, sqlTimeStr]);

        const rawValues = rows.length > 0 ? [
            rows[0].energy1, rows[0].energy2, rows[0].energy3
        ] : new Array(3).fill('');

        const formattedValues = rawValues.map(val => toInteger(val));

        // 4. Load Excel
        let workbook;
        if (fs.existsSync(filePath)) {
            workbook = await XlsxPopulate.fromFileAsync(filePath);
        } else {
            console.log(`Creating new Shift Log for ${targetYear}...`);
            workbook = await XlsxPopulate.fromFileAsync(TEMPLATE_PATH);
        }

        const sheet = workbook.sheet(sheetIndex);
        if (!sheet) {
            console.error(`Sheet index ${sheetIndex} missing.`);
            return;
        }

        // 5. Write Data
        formattedValues.forEach((val, i) => {
            sheet.row(targetRow).cell(2 * (i + 1) + 1).value(val);
        });

        await workbook.toFileAsync(filePath);
        console.log(`[Success] Updated Shift Log: Row ${targetRow} (${sqlTimeStr})`);

    } catch (error) {
        console.error('[Error] Shift Update failed:', error);
    } finally {
        if (connection) await connection.end();
    }
}

// --- SCHEDULE ---
// Run at 00:00 (Midnight) and 12:00 (Noon)
cron.schedule('1 0,12 * * *', () => {
    updateTwiceDaily();
});

// Run immediately for testing
updateTwiceDaily();

app.listen(PORT, () => {
    console.log(`Shift Logger running on Port ${PORT}.`);
});