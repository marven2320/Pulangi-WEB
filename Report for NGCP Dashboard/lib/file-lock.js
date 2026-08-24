// Shared "is this workbook open in Excel right now?" helper, used by every
// script that writes into a generated .xlsx file (the mergers and the
// report generators). Centralized here so all 6 jobs treat a locked file
// the same way instead of each re-implementing the check.

const fs = require('fs');
const opened = require('@ronomon/opened');

// Resolves true if filePath exists and is currently held open by another
// process (e.g. a user has it open in Excel). A non-existent file is never
// "locked" - there's nothing to conflict with.
function checkIfFileIsOpen(filePath) {
    return new Promise((resolve) => {
        if (!fs.existsSync(filePath)) return resolve(false);

        opened.files([filePath], (error, hashTable) => {
            if (error) {
                console.error('[CheckOpen] Error:', error);
                return resolve(true); // Assume locked on error - safer to retry later than to lose data
            }
            resolve(hashTable[filePath] === true);
        });
    });
}

// Node surfaces a locked/permission-denied file write as one of these
// codes. Used as a second line of defense when a write is attempted after
// checkIfFileIsOpen() already said "free" but the file was opened in the
// gap between the check and the write.
function isFileLockError(err) {
    return !!err && (err.code === 'EBUSY' || err.code === 'EPERM' || err.code === 'EACCES');
}

module.exports = { checkIfFileIsOpen, isFileLockError };
