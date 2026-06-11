const express = require('express');
const exceljs = require('exceljs');
const qrcode = require('qrcode');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_FILE = path.join(__dirname, 'attendance.xlsx');

let currentSession = null;

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.use((req, res, next) => {
  console.log(`[${new Date().toLocaleTimeString()}] ${req.method} ${req.url} from ${req.ip}`);
  next();
});

function getSheetName(sessionName) {
  const now = new Date();
  const date = `${String(now.getDate()).padStart(2,'0')}/${String(now.getMonth()+1).padStart(2,'0')}`;
  const raw = `${date} ${sessionName}`;
  return raw.replace(/[:\\\/\?\*\[\]]/g, '').substring(0, 31);
}

async function ensureExcelFile() {
  if (fs.existsSync(EXCEL_FILE)) return;
  const wb = new exceljs.Workbook();
  await wb.xlsx.writeFile(EXCEL_FILE);
  console.log('>>> Created attendance.xlsx');
}

async function addSheetForSession(sessionName) {
  await ensureExcelFile();
  const wb = new exceljs.Workbook();
  await wb.xlsx.readFile(EXCEL_FILE);
  const sheetName = getSheetName(sessionName);
  let ws = wb.getWorksheet(sheetName);
  if (!ws) {
    ws = wb.addWorksheet(sheetName);
    ws.addRow(['Name', 'Session', 'Date', 'Time']);
    ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E5A9C' } };
    ws.getRow(1).alignment = { horizontal: 'center' };
    ws.columns = [
      { key: 'name', width: 30 }, { key: 'session', width: 25 },
      { key: 'date', width: 15 }, { key: 'time', width: 15 }
    ];
  }
  await wb.xlsx.writeFile(EXCEL_FILE);
  return sheetName;
}

app.post('/session/start', async (req, res) => {
  const { name } = req.body;
  const sessionName = name || `Session ${new Date().toLocaleDateString('en-GB')}`;
  const sheetName = await addSheetForSession(sessionName);
  currentSession = {
    id: crypto.randomBytes(16).toString('hex'),
    name: sessionName,
    sheetName,
    startedAt: new Date()
  };
  console.log(`>>> SESSION STARTED: ${currentSession.name}`);
  res.json({ success: true, session: currentSession });
});

app.post('/session/end', (req, res) => {
  currentSession = null;
  res.json({ success: true });
});

app.get('/session/status', async (req, res) => {
  if (!currentSession) return res.json({ active: false });
  const protocol = req.headers['x-forwarded-proto'] || 'http';
  const host = req.headers['x-forwarded-host'] || req.headers.host;
  const checkInUrl = `${protocol}://${host}/checkin?sid=${currentSession.id}`;
  const qrDataUrl = await qrcode.toDataURL(checkInUrl, { width: 300, margin: 2 });
  res.json({ active: true, session: currentSession, qr: qrDataUrl, url: checkInUrl });
});

app.post('/submit', async (req, res) => {
  const { name, sid } = req.body;
  if (!currentSession)
    return res.status(403).json({ error: 'No active session. Please wait for your trainer to start the session.' });
  if (sid !== currentSession.id)
    return res.status(403).json({ error: 'This QR code has expired. Please scan the new QR code.' });
  if (!name || !name.trim())
    return res.status(400).json({ error: 'Name is required' });

  try {
    const now = new Date();
    const uaeTime = new Date(now.getTime() + (4 * 60 * 60 * 1000));
    const date = uaeTime.toLocaleDateString('en-GB');
    const time = uaeTime.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const wb = new exceljs.Workbook();
    await wb.xlsx.readFile(EXCEL_FILE);
    const ws = wb.getWorksheet(currentSession.sheetName);
    ws.addRow([name.trim(), currentSession.name, date, time]);
    await wb.xlsx.writeFile(EXCEL_FILE);
    console.log(`>>> SAVED: ${name.trim()} | ${currentSession.sheetName} | ${date} ${time}`);
    res.json({ success: true, name: name.trim(), session: currentSession.name, date, time });
  } catch (err) {
    console.error('>>> ERROR:', err.message);
    res.status(500).json({ error: 'Failed to save attendance' });
  }
});

app.get('/checkin', (req, res) => res.sendFile(path.join(__dirname, 'public', 'checkin.html')));
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'admin.html')));

const API_KEY = process.env.ATTENDANCE_API_KEY;

app.get('/api/attendance', async (req, res) => {
  if (!API_KEY || req.headers['x-api-key'] !== API_KEY) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  try {
    if (!fs.existsSync(EXCEL_FILE)) return res.json({ rows: [], columns: [] });
    const wb = new exceljs.Workbook();
    await wb.xlsx.readFile(EXCEL_FILE);
    const rows = [];
    wb.eachSheet((ws) => {
      ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber === 1) return; // skip header
        const [name, session, date, time] = row.values.slice(1);
        if (!name) return;
        rows.push({
          Name: String(name ?? ''),
          Session: String(session ?? ''),
          Date: String(date ?? ''),
          Time: String(time ?? ''),
          Sheet: ws.name,
        });
      });
    });
    res.json({ rows, columns: ['Name', 'Session', 'Date', 'Time', 'Sheet'] });
  } catch (err) {
    console.error('>>> /api/attendance error:', err.message);
    res.status(500).json({ error: 'Failed to read attendance' });
  }
});



ensureExcelFile().then(() => {
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`\n✅ Attendance app running on port ${PORT}`);
  });
});
