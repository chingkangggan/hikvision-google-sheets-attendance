const express = require('express');
const multer = require('multer');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

const app = express();

// Standard memory storage - simple, reliable, and handles all multipart fields/files
const upload = multer({ storage: multer.memoryStorage() });

const PORT = 3000;
const BUFFER = new Set();
const APPS_SCRIPT_URL = "APPS_SCRIPT_URL";

const LOG_FILE = path.join(__dirname, 'attendance.log');
const BUFFER_FILE = path.join(__dirname, 'buffer.json');

function log(message, type = 'INFO') {
    const timestamp = new Date().toLocaleString('en-US', { timeZone: 'Asia/Kuala_Lumpur' });
    const logLine = `[${timestamp}] [${type}] ${message}\n`;
    fs.appendFile(LOG_FILE, logLine, (err) => { if (err) console.error('Log write failed:', err); });
    console.log(`[${type}] ${message}`);
}

// ============== Buffer File Persistence ==============
function loadBuffer() {
    try {
        if (fs.existsSync(BUFFER_FILE)) {
            const rawData = fs.readFileSync(BUFFER_FILE, 'utf8');
            const data = JSON.parse(rawData);
            if (Array.isArray(data)) {
                data.forEach(id => BUFFER.add(String(id)));
                log(`Loaded ${BUFFER.size} pending records from buffer.json`, 'START');
            }
        }
    } catch (err) {
        log(`Error loading buffer.json: ${err.message}`, 'ERROR');
    }
}

function saveBuffer() {
    try {
        const data = Array.from(BUFFER);
        fs.writeFileSync(BUFFER_FILE, JSON.stringify(data), 'utf8');
    } catch (err) {
        log(`Error saving buffer.json: ${err.message}`, 'ERROR');
    }
}

loadBuffer();

// ============== Middleware ==============
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));
app.use(express.text({ type: ['*/xml', 'text/xml', 'application/xml'], limit: '10mb' }));

// ============== Hikvision Event Endpoint ==============
app.post('/device/events', upload.any(), async (req, res) => {
    const ip = req.ip || req.socket.remoteAddress;
    log(`Event received from IP: ${ip}`, 'RECEIVED');

    try {
        let rawEvent = null;
	
        // 1. Check if the event data was caught as a multipart "file" because of the text/json header
        if (req.files && req.files.length > 0) {
            const eventFile = req.files.find(f => f.fieldname === 'AccessControllerEvent');
            if (eventFile) {
                rawEvent = eventFile.buffer.toString('utf8');
            }
        }

        // 2. Fallback to standard request body extraction options
        if (!rawEvent) {
            rawEvent = req.body.AccessControllerEvent || 
                       req.body.EventNotificationAlert ||
                       req.body.event || 
                       req.body;
        }

        // 3. Convert stringified JSON payloads into an object
        if (typeof rawEvent === 'string') {
            const trimmed = rawEvent.trim();
            if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
                try {
                    rawEvent = JSON.parse(trimmed);
                } catch (e) {
                    // Fail gracefully, keep as string for potential XML fallback
                }
            }
        }
	if (rawEvent.eventType == 'heartBeat'){
		return res.status(200).send("OK");
	}
        let student_id = '';
        // 4. Extract based on final resolved format (XML or JSON Object)
        if (typeof rawEvent === 'string') {
            if (rawEvent.includes('<')) {
                const matchString = rawEvent.match(/<(?:[a-zA-Z0-9_\-]+:)?(employeeNoString)[^>]*>([\s\S]*?)<\/(?:[a-zA-Z0-9_\-]+:)?\1>/);
                if (matchString && matchString[2] && matchString[2].trim() !== '') { // <-- FIXED [2]
                    student_id = matchString[2].trim(); // <-- FIXED [2]
                } else {
                    const matchInt = rawEvent.match(/<(?:[a-zA-Z0-9_\-]+:)?(employeeNo)[^>]*>([\s\S]*?)<\/(?:[a-zA-Z0-9_\-]+:)?\1>/);
                    if (matchInt && matchInt[2]) { // <-- FIXED [2]
                        student_id = matchInt[2].trim(); // <-- FIXED [2]
                    }
                }
            }
        } else if (typeof rawEvent === 'object' && rawEvent !== null) {
            const acs = rawEvent?.AccessControllerEvent || 
                        rawEvent?.accessControllerEvent || 
                        rawEvent;
            student_id = String(acs?.employeeNoString || acs?.employeeNo || '').trim();
        }

        // 5. Buffer the outcome
        if (student_id) {
            BUFFER.add(student_id);
            saveBuffer(); 
            log(`Buffered: ${student_id} | Total in buffer: ${BUFFER.size}`, 'BUFFER');
        } else {
            //log('Event received but no employeeNoString/employeeNo found', 'WARN');
        }

    } catch (err) {
        log(`Error processing event: ${err.message}`, 'ERROR');
    }

    res.status(200).send("OK");
});

// ============== Main Flush Function ==============
async function flushBuffer() {
    if (BUFFER.size === 0) return;

    const snapshot = Array.from(BUFFER);
    BUFFER.clear();
    saveBuffer();

    try {
        const payload = { student_ids: snapshot };
        await axios.post(APPS_SCRIPT_URL, payload, { timeout: 15000 });
        log(`Successfully flushed ${snapshot.length} records to Google Sheets`, 'SUCCESS');
        
        if (fs.existsSync(BUFFER_FILE)) {
            fs.unlinkSync(BUFFER_FILE);
        }
    } catch (err) {
        log(`Flush to Apps Script failed: ${err.message}. Rebuffering.`, 'ERROR');
        snapshot.forEach(id => BUFFER.add(id));
        saveBuffer();
    }
}

setInterval(flushBuffer, 1 * 60 * 1000);

// ============== Graceful Shutdown Handler ==============
async function handleShutdown(signal) {
    log(`Received ${signal}. Attempting clean shutdown...`, 'SHUTDOWN');
    
    if (BUFFER.size > 0) {
        log(`Flushing remaining ${BUFFER.size} records before exit...`, 'SHUTDOWN');
        const snapshot = Array.from(BUFFER);
        BUFFER.clear();
        saveBuffer();

        try {
            const payload = { student_ids: snapshot };
            await axios.post(APPS_SCRIPT_URL, payload, { timeout: 5000 });
            log(`Shutdown flush completed successfully.`, 'SUCCESS');
            if (fs.existsSync(BUFFER_FILE)) {
                fs.unlinkSync(BUFFER_FILE);
            }
        } catch (err) {
            log(`Shutdown flush failed: ${err.message}. Records remain in buffer.json.`, 'ERROR');
            snapshot.forEach(id => BUFFER.add(id));
            saveBuffer();
        }
    } else {
        log('No pending buffer records to flush.', 'SHUTDOWN');
    }
    
    process.exit(0);
}

process.on('SIGINT', () => handleShutdown('SIGINT'));
process.on('SIGTERM', () => handleShutdown('SIGTERM'));

app.listen(PORT, '0.0.0.0', () => {
    log(`Hikvision Listener started on port ${PORT}`, 'START');
});
