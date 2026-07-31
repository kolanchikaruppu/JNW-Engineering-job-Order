const http = require('http');
const fs = require('fs/promises');
const path = require('path');
const os = require('os');
const { execFile } = require('child_process');

const PORT = Number(process.env.PORT || 4322);
const HOST = process.env.HOST || '0.0.0.0';
const ROOT = __dirname;
const REPORTS_DIR = path.join(ROOT, 'Reports');
const SEQUENCE_FILE = path.join(REPORTS_DIR, '.sequence.json');
const LOG_FILE = path.join(REPORTS_DIR, 'Job_Order_Log.csv');
const PYTHON = process.env.PYTHON || path.join(os.homedir(), '.cache', 'codex-runtimes', 'codex-primary-runtime', 'dependencies', 'python', 'python.exe');
const PDF_SCRIPT = path.join(ROOT, 'generate_pdf.py');

const mimeTypes = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.json': 'application/json; charset=utf-8',
  '.csv': 'text/csv; charset=utf-8',
};

const json = (response, status, body) => {
  response.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8' });
  response.end(JSON.stringify(body));
};

const readBody = (request) => new Promise((resolve, reject) => {
  let body = '';
  request.on('data', (chunk) => {
    body += chunk;
    if (body.length > 20 * 1024 * 1024) {
      reject(new Error('Submission is too large.'));
      request.destroy();
    }
  });
  request.on('end', () => resolve(body));
  request.on('error', reject);
});

const ensureReportsDir = async () => {
  await fs.mkdir(REPORTS_DIR, { recursive: true });
};

const readSequence = async () => {
  await ensureReportsDir();
  try {
    const sequence = JSON.parse(await fs.readFile(SEQUENCE_FILE, 'utf8'));
    return Number.isFinite(sequence.nextJobOrder) ? sequence.nextJobOrder : 1;
  } catch {
    return 1;
  }
};

const writeSequence = async (nextJobOrder) => {
  await fs.writeFile(SEQUENCE_FILE, JSON.stringify({ nextJobOrder }, null, 2));
};

const createPdf = (jsonPath, pdfPath) => new Promise((resolve, reject) => {
  execFile(PYTHON, [PDF_SCRIPT, jsonPath, pdfPath], { cwd: ROOT }, (error, stdout, stderr) => {
    if (error) {
      reject(new Error(stderr || stdout || error.message));
      return;
    }
    resolve();
  });
});

const formatJobOrder = (number) => String(Math.max(1, Number(number) || 1)).padStart(2, '0');

const escapeHtml = (value) => String(value || '')
  .replaceAll('&', '&amp;')
  .replaceAll('<', '&lt;')
  .replaceAll('>', '&gt;')
  .replaceAll('"', '&quot;');

const csvCell = (value) => `"${String(value || '').replaceAll('"', '""')}"`;

const cleanFilePart = (value) => String(value || '')
  .trim()
  .replace(/[<>:"/\\|?*\x00-\x1F]/g, '')
  .replace(/\s+/g, ' ')
  .slice(0, 60);

const normaliseParts = (parts) => Array.isArray(parts)
  ? parts.map((part, index) => ({
      serial: part.serial || String(index + 1),
      description: part.description || '',
      quantity: part.quantity || part.charge || '',
    })).filter((part) => part.description || part.quantity)
  : [];

const makeReportHtml = (data) => {
  const parts = normaliseParts(data.parts);
  const partRows = parts.length
    ? parts.map((part) => `
      <tr>
        <td>${escapeHtml(part.serial)}</td>
        <td>${escapeHtml(part.description)}</td>
        <td>${escapeHtml(part.quantity)}</td>
      </tr>`).join('')
    : '<tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>';

  const signature = (title, value) => value
    ? `<section><h3>${title}</h3><img class="signature" src="${escapeHtml(value)}" alt="${title}" /></section>`
    : `<section><h3>${title}</h3><div class="signature blank"></div></section>`;

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>JNW Job Order ${escapeHtml(data.jobOrder)}</title>
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; padding: 24px; color: #111; font-family: "Segoe UI", Arial, sans-serif; }
    .page { max-width: 900px; margin: 0 auto; }
    header { display: flex; justify-content: space-between; gap: 24px; align-items: center; border-bottom: 3px solid #0f304d; padding-bottom: 18px; margin-bottom: 22px; }
    .logo { width: 290px; max-width: 55%; }
    .job h1 { margin: 0 0 8px; color: #0f304d; font-size: 18px; letter-spacing: 0.08em; text-transform: uppercase; }
    .job strong { color: #c82f2f; font-size: 34px; letter-spacing: 0.1em; }
    h2 { margin: 24px 0 8px; color: #0f304d; font-size: 18px; text-transform: uppercase; }
    h3, label { margin: 0 0 6px; color: #606975; font-size: 13px; text-transform: uppercase; }
    .grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }
    .three { grid-template-columns: repeat(3, minmax(0, 1fr)); }
    .field, textarea, table, .signature { border: 1px solid #d8dee6; border-radius: 4px; }
    .field { min-height: 46px; padding: 12px; font-size: 18px; }
    .text { min-height: 105px; white-space: pre-wrap; }
    table { width: 100%; border-collapse: collapse; overflow: hidden; }
    th, td { border-bottom: 1px solid #d8dee6; padding: 10px; text-align: left; vertical-align: top; }
    th { color: #606975; font-size: 13px; text-transform: uppercase; }
    th:first-child, td:first-child { width: 80px; }
    th:last-child, td:last-child { width: 150px; }
    .signature-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; }
    .signature { width: 100%; height: 170px; object-fit: contain; display: block; background: #fff; }
    .blank { background: linear-gradient(#fff 92%, #e6ebf0 92%); }
    @media print { body { padding: 10mm; } .page { max-width: none; } }
  </style>
</head>
<body>
  <main class="page">
    <header>
      <img class="logo" src="../JNW.png" alt="JNW Engineering Pte Ltd" />
      <div class="job"><h1>Job Order</h1><strong>${escapeHtml(data.jobOrder)}</strong></div>
    </header>

    <h2>Customer Details</h2>
    <div class="grid">
      <section><label>Company</label><div class="field">${escapeHtml(data.company)}</div></section>
      <section><label>Date</label><div class="field">${escapeHtml(data.date)}</div></section>
      <section><label>Requested By</label><div class="field">${escapeHtml(data.requestedBy)}</div></section>
    </div>

    <h2>Work Report</h2>
    <section><label>Complaint</label><div class="field text">${escapeHtml(data.complaint)}</div></section>
    <section><label>Action Taken</label><div class="field text">${escapeHtml(data.actionTaken)}</div></section>

    <h2>Materials and Parts Used</h2>
    <table>
      <thead><tr><th>S/No</th><th>Materials and Parts Used</th><th>Quantity</th></tr></thead>
      <tbody>${partRows}</tbody>
    </table>

    <h2>Labour</h2>
    <div class="grid three">
      <section><label>Labour Description</label><div class="field">${escapeHtml(data.labourDescription)}</div></section>
      <section><label>Man</label><div class="field">${escapeHtml(data.labourMan)}</div></section>
      <section><label>Hours</label><div class="field">${escapeHtml(data.labourHours)}</div></section>
    </div>

    <h2>Remarks</h2>
    <div class="field text">${escapeHtml(data.remarks)}</div>

    <h2>Signatures</h2>
    <div class="signature-grid">
      ${signature('Customer Signature and Chop', data.signatures?.customerSignature)}
      ${signature('Technician Signature', data.signatures?.technicianSignature)}
    </div>
  </main>
</body>
</html>`;
};

const appendCsv = async (data, htmlFile, jsonFile) => {
  const exists = await fs.access(LOG_FILE).then(() => true).catch(() => false);
  if (!exists) {
    await fs.writeFile(LOG_FILE, [
      'Submitted At',
      'Job Order',
      'Company',
      'Date',
      'Requested By',
      'Complaint',
      'Action Taken',
      'Materials',
      'Labour Description',
      'Man',
      'Hours',
      'Remarks',
      'HTML File',
      'JSON File',
    ].map(csvCell).join(',') + os.EOL);
  }

  const materials = normaliseParts(data.parts)
    .map((part) => `${part.serial}. ${part.description} x ${part.quantity}`)
    .join(' | ');

  await fs.appendFile(LOG_FILE, [
    data.submittedAt,
    data.jobOrder,
    data.company,
    data.date,
    data.requestedBy,
    data.complaint,
    data.actionTaken,
    materials,
    data.labourDescription,
    data.labourMan,
    data.labourHours,
    data.remarks,
    htmlFile,
    jsonFile,
  ].map(csvCell).join(',') + os.EOL);
};

const saveSubmission = async (payload) => {
  await ensureReportsDir();
  const nextNumber = await readSequence();
  const jobOrder = formatJobOrder(nextNumber);
  const nextJobOrder = nextNumber + 1;
  const submittedAt = new Date().toISOString();
  const companyPart = cleanFilePart(payload.company) || 'No Company';
  const datePart = submittedAt.replace(/[:.]/g, '-');
  const baseName = `${jobOrder}_${datePart}_${companyPart}`;
  const htmlFile = `${baseName}.html`;
  const jsonFile = `${baseName}.json`;
  const pdfFile = `${baseName}.pdf`;
  const data = { ...payload, jobOrder, submittedAt };
  const jsonPath = path.join(REPORTS_DIR, jsonFile);
  const pdfPath = path.join(REPORTS_DIR, pdfFile);

  await fs.writeFile(jsonPath, JSON.stringify(data, null, 2), 'utf8');
  await createPdf(jsonPath, pdfPath);
  await writeSequence(nextJobOrder);

  return { jobOrder, nextJobOrder: formatJobOrder(nextJobOrder), pdfFile, jsonFile };
};

const serveStatic = async (request, response) => {
  const url = new URL(request.url, `http://${request.headers.host}`);
  const requestPath = url.pathname === '/' ? '/index.html' : decodeURIComponent(url.pathname);
  const filePath = path.normalize(path.join(ROOT, requestPath));

  if (!filePath.startsWith(ROOT)) {
    response.writeHead(403);
    response.end('Forbidden');
    return;
  }

  try {
    const file = await fs.readFile(filePath);
    response.writeHead(200, { 'Content-Type': mimeTypes[path.extname(filePath).toLowerCase()] || 'application/octet-stream' });
    response.end(file);
  } catch {
    response.writeHead(404);
    response.end('Not found');
  }
};

const server = http.createServer(async (request, response) => {
  try {
    if (request.method === 'GET' && request.url.startsWith('/api/next-job-order')) {
      json(response, 200, { nextJobOrder: formatJobOrder(await readSequence()) });
      return;
    }

    if (request.method === 'POST' && request.url.startsWith('/api/submit')) {
      const payload = JSON.parse(await readBody(request));
      json(response, 200, await saveSubmission(payload));
      return;
    }

    if (request.method === 'GET') {
      await serveStatic(request, response);
      return;
    }

    response.writeHead(405);
    response.end('Method not allowed');
  } catch (error) {
    json(response, 500, { error: error.message || 'Server error.' });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`JNW Job Order server running at http://localhost:${PORT}/`);
  console.log(`Reports save to: ${REPORTS_DIR}`);
});
