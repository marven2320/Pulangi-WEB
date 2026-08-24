// Shared status/telemetry recorder for the 6 backend data/report jobs
// (dgr_update, dgr_merger, dor_update, dor_merger, dor_generate, mor_generate).
//
// Each job calls recordEvent(...) at key points (run start, success, skipped
// because a file is open, error). This module persists two small JSON files
// under rawdata/ that the web dashboard (dashboard.html + the /api/jobs*
// routes in server.js) polls to render live status cards and history charts:
//
//   rawdata/job_status.json  - one row per job, current/last-known state
//   rawdata/job_history.json - rolling log of the most recent events (capped)
//
// Each of the 6 scripts runs as its own long-lived node-cron process, so
// writes to these shared files can race between processes. That's the same
// tradeoff the existing data_buffer*.json files already make (no locking),
// so this module follows suit: best-effort read-modify-write, with a
// write-to-temp-then-rename so a reader never sees a half-written file.

const fs = require('fs');
const path = require('path');

const OUTPUT_DIR = path.join(__dirname, '..', 'rawdata');
const STATUS_FILE = path.join(OUTPUT_DIR, 'job_status.json');
const HISTORY_FILE = path.join(OUTPUT_DIR, 'job_history.json');
const MAX_HISTORY = 500;

// Static metadata about each job, used to label the dashboard even before
// a job has ever reported in (e.g. right after a fresh deploy).
const JOB_META = {
    dgr_update: {
        label: 'Generation Data - Shift Logger',
        description: 'Reads shift totals (00:00/12:00) from MySQL and buffers them for the merger.',
        schedule: '1 0,12 * * *',
        group: 'generation'
    },
    dgr_merger: {
        label: 'Generation Data - Shift Merger',
        description: 'Writes buffered shift entries into the Generation Data RAW workbook, retrying while it is open.',
        schedule: '*/2 * * * *',
        group: 'generation'
    },
    dor_update: {
        label: 'Operational Highlights - Hourly Logger',
        description: 'Reads hourly readings from MySQL, backfills gaps, and buffers them for the merger.',
        schedule: '1 * * * *',
        group: 'operational'
    },
    dor_merger: {
        label: 'Operational Highlights - Merger',
        description: 'Writes buffered hourly entries into the Operational Highlights RAW workbook, retrying while it is open.',
        schedule: '*/2 * * * *',
        group: 'operational'
    },
    dor_generate: {
        label: 'Daily Operations Report Generator',
        description: 'Builds the previous day\'s Daily Operations Report from the RAW workbooks and outage log.',
        schedule: '1 0,7 * * *',
        group: 'reports'
    },
    mor_generate: {
        label: 'Monthly Operations Report Generator',
        description: 'Builds the previous cycle\'s Monthly Operations Report from the monthly, shift, and downtime logs.',
        schedule: '1 0 26 * *',
        group: 'reports'
    }
};

function ensureDir() {
    if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

function readJSON(file, fallback) {
    try {
        if (!fs.existsSync(file)) return fallback;
        return JSON.parse(fs.readFileSync(file, 'utf8'));
    } catch (e) {
        return fallback;
    }
}

function writeJSON(file, data) {
    ensureDir();
    const tmp = `${file}.tmp.${process.pid}`;
    fs.writeFileSync(tmp, JSON.stringify(data, null, 2));
    fs.renameSync(tmp, file);
}

// type: 'run-start' | 'success' | 'skip-locked' | 'error' | 'info'
function recordEvent(job, type, message, meta) {
    const now = new Date().toISOString();
    const jobMeta = JOB_META[job] || {};

    const status = readJSON(STATUS_FILE, {});
    const prev = status[job] || {};

    status[job] = {
        job,
        label: jobMeta.label || job,
        description: jobMeta.description || null,
        schedule: jobMeta.schedule || null,
        group: jobMeta.group || 'other',
        status: type === 'run-start' ? 'running' : type,
        lastEventAt: now,
        lastRunAt: type === 'run-start' ? now : (prev.lastRunAt || null),
        lastSuccessAt: type === 'success' ? now : (prev.lastSuccessAt || null),
        lastErrorAt: type === 'error' ? now : (prev.lastErrorAt || null),
        message: message || null,
        meta: meta || {},
        pid: process.pid
    };
    writeJSON(STATUS_FILE, status);

    const history = readJSON(HISTORY_FILE, []);
    history.push({ ts: now, job, type, message: message || null, meta: meta || {} });
    while (history.length > MAX_HISTORY) history.shift();
    writeJSON(HISTORY_FILE, history);

    return status[job];
}

function getStatus() {
    return readJSON(STATUS_FILE, {});
}

function getHistory() {
    return readJSON(HISTORY_FILE, []);
}

module.exports = {
    recordEvent,
    getStatus,
    getHistory,
    JOB_META,
    STATUS_FILE,
    HISTORY_FILE
};
