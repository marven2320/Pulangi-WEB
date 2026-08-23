// Shared query logic for fetching Power (mw1-3) / Frequency (freq1-3)
// generation data from the `pulangi` table over a date/time range.
//
// Single source of truth used by both:
//   - server.js's GET /api/generation-data endpoint (what viewdata.html's
//     dual-axis charts actually call live)
//   - fetch-generation-data.js, the standalone CLI script
//
// Keeping the validation + SQL in one place means the live web data and the
// CLI tool's output can never drift apart.

const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const TIME_RE = /^\d{2}:\d{2}:\d{2}$/;

function isValidParams(date, start, end) {
    return DATE_RE.test(date || '') && TIME_RE.test(start || '') && TIME_RE.test(end || '');
}

// callback(err, rows)
function getGenerationData(pool, { date, start, end }, callback) {
    if (!isValidParams(date, start, end)) {
        callback(new Error('date must be YYYY-MM-DD and start/end must be HH:MM:SS'));
        return;
    }

    const sql = 'SELECT `time`, `mw1`, `mw2`, `mw3`, `freq1`, `freq2`, `freq3` ' +
                'FROM `pulangi` WHERE `date` = ? AND `time` BETWEEN ? AND ? ORDER BY `time` ASC';

    pool.query(sql, [date, start, end], (err, rows) => {
        if (err) { callback(err); return; }
        callback(null, rows);
    });
}

module.exports = { getGenerationData, isValidParams };
