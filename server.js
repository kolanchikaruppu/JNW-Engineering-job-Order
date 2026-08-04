const http = require('http');
const fs = require('fs/promises');
const path = require('path');
const os = require('os');
const PDFDocument = require('pdfkit');
const { DefaultAzureCredential } = require('@azure/identity');

const PORT = Number(process.env.PORT || 4322);
const HOST = process.env.HOST || '0.0.0.0';
const ROOT = __dirname;
const REPORTS_DIR = path.join(ROOT, 'Reports');
const SEQUENCE_FILE = path.join(REPORTS_DIR, '.sequence.json');
const LOG_FILE = path.join(REPORTS_DIR, 'Job_Order_Log.csv');

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const GRAPH_SITE_HOSTNAME = process.env.GRAPH_SITE_HOSTNAME || 'jnwengineering.sharepoint.com';
const GRAPH_SITE_PATH = process.env.GRAPH_SITE_PATH || '/';
const GRAPH_DOCUMENT_LIBRARY = process.env.GRAPH_DOCUMENT_LIBRARY || 'Documents';
const GRAPH_FOLDER_PATH = process.env.GRAPH_FOLDER_PATH || 'Job Order Reports';

let graphTokenCache = null;
let graphDriveCache = null;

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

const formatJobOrder = (number) => String(Math.max(1, Number(number) || 1)).padStart(2, '0');

const csvCell = (value) => `"${String(value || '').replaceAll('"', '""')}"`;

const cleanFilePart = (value) => String(value || '')
  .trim()
  .replace(/[<>:"/\\|?*\x00-\x1F]/g, '')
  .replace(/\s+/g, ' ')
  .slice(0, 60);

const formatDateForFile = (value, fallbackDate = new Date()) => {
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) return `${match[3]}-${match[2]}-${match[1]}`;

  const parsed = value ? new Date(value) : fallbackDate;
  if (Number.isNaN(parsed.getTime())) return fallbackDate.toISOString().slice(0, 10);
  const day = String(parsed.getDate()).padStart(2, '0');
  const month = String(parsed.getMonth() + 1).padStart(2, '0');
  const year = parsed.getFullYear();
  return `${day}-${month}-${year}`;
};

const normaliseParts = (parts) => Array.isArray(parts)
  ? parts.map((part, index) => ({
      serial: part.serial || String(index + 1),
      description: part.description || '',
      quantity: part.quantity || part.charge || '',
    })).filter((part) => part.description || part.quantity)
  : [];

const addText = (doc, text, x, y, options = {}) => {
  doc.text(String(text || ''), x, y, options);
};

const addBox = (doc, label, value, x, y, width, height = 34) => {
  doc.font('Helvetica-Bold').fontSize(9).fillColor('#4f5d6d').text(label, x, y);
  doc.roundedRect(x, y + 13, width, height, 3).strokeColor('#d4dce5').stroke();
  doc.font('Helvetica').fontSize(12).fillColor('#111111').text(String(value || ''), x + 8, y + 23, {
    width: width - 16,
    height: height - 10,
  });
};

const addMultilineBox = (doc, label, value, x, y, width, height) => {
  doc.font('Helvetica-Bold').fontSize(9).fillColor('#4f5d6d').text(label, x, y);
  doc.roundedRect(x, y + 13, width, height, 3).strokeColor('#d4dce5').stroke();
  doc.font('Helvetica').fontSize(11).fillColor('#111111').text(String(value || ''), x + 8, y + 23, {
    width: width - 16,
    height: height - 10,
  });
};

const drawSignature = (doc, label, dataUrl, x, y, width, height) => {
  doc.font('Helvetica-Bold').fontSize(9).fillColor('#4f5d6d').text(label, x, y);
  doc.roundedRect(x, y + 13, width, height, 3).strokeColor('#d4dce5').stroke();
  doc.rect(x, y + 13 + height - 10, width, 10).fillColor('#e8edf2').fill();
  if (!dataUrl) return;

  const base64 = String(dataUrl).split(',')[1];
  if (!base64) return;
  try {
    const image = Buffer.from(base64, 'base64');
    doc.image(image, x + 8, y + 21, { fit: [width - 16, height - 22], align: 'center', valign: 'center' });
  } catch {
    // Keep the PDF valid even if a phone submits a damaged signature image.
  }
};

const createPdfBuffer = async (data) => new Promise(async (resolve, reject) => {
  const chunks = [];
  const doc = new PDFDocument({ size: 'A4', margin: 28 });
  doc.on('data', (chunk) => chunks.push(chunk));
  doc.on('end', () => resolve(Buffer.concat(chunks)));
  doc.on('error', reject);

  const pageWidth = doc.page.width - 56;
  const logoPath = path.join(ROOT, 'JNW.png');

  try {
    await fs.access(logoPath);
    doc.image(logoPath, 28, 25, { width: 190 });
  } catch {
    doc.font('Helvetica-Bold').fontSize(18).fillColor('#0f304d').text('JNW ENGINEERING PTE LTD', 28, 42);
  }

  doc.font('Helvetica-Bold').fontSize(11).fillColor('#0f304d').text('JOB ORDER', 390, 32);
  doc.font('Helvetica-Bold').fontSize(24).fillColor('#c92d2d').text(data.jobOrder, 390, 52, { characterSpacing: 2 });
  doc.moveTo(28, 88).lineTo(doc.page.width - 28, 88).lineWidth(1.6).strokeColor('#0f304d').stroke();

  const col = (pageWidth - 18) / 2;
  addBox(doc, 'COMPANY', data.company, 28, 108, col);
  addBox(doc, 'DATE', data.date, 28 + col + 18, 108, col);
  addBox(doc, 'REQUESTED BY', data.requestedBy, 28, 158, col);
  addMultilineBox(doc, 'COMPLAINT', data.complaint, 28, 218, pageWidth, 68);
  addMultilineBox(doc, 'ACTION TAKEN', data.actionTaken, 28, 306, pageWidth, 100);

  doc.font('Helvetica-Bold').fontSize(12).fillColor('#0f304d').text('MATERIALS AND PARTS USED', 28, 432);
  const parts = normaliseParts(data.parts);
  const tableY = 452;
  const serialW = 48;
  const qtyW = 86;
  const descW = pageWidth - serialW - qtyW;
  doc.font('Helvetica-Bold').fontSize(9).fillColor('#4f5d6d');
  addText(doc, 'S/NO', 28, tableY);
  addText(doc, 'MATERIALS AND PARTS USED', 28 + serialW, tableY);
  addText(doc, 'QUANTITY', 28 + serialW + descW, tableY);
  let rowY = tableY + 16;
  const rows = parts.length ? parts : [{ serial: '1', description: '', quantity: '' }];
  rows.slice(0, 8).forEach((part) => {
    doc.roundedRect(28, rowY, serialW - 4, 28, 3).strokeColor('#d4dce5').stroke();
    doc.roundedRect(28 + serialW, rowY, descW - 8, 28, 3).stroke();
    doc.roundedRect(28 + serialW + descW, rowY, qtyW, 28, 3).stroke();
    doc.font('Helvetica-Bold').fontSize(11).fillColor('#111111').text(String(part.serial || ''), 38, rowY + 8);
    doc.font('Helvetica').fontSize(11).text(String(part.description || ''), 28 + serialW + 8, rowY + 8, { width: descW - 24 });
    doc.text(String(part.quantity || ''), 28 + serialW + descW + 8, rowY + 8, { width: qtyW - 16 });
    rowY += 34;
  });

  const labourY = rowY + 16;
  doc.font('Helvetica-Bold').fontSize(12).fillColor('#0f304d').text('LABOUR', 28, labourY);
  const third = (pageWidth - 24) / 3;
  addBox(doc, 'LABOUR DESCRIPTION', data.labourDescription, 28, labourY + 22, third);
  addBox(doc, 'MAN', data.labourMan, 28 + third + 12, labourY + 22, third);
  addBox(doc, 'HOURS', data.labourHours, 28 + (third + 12) * 2, labourY + 22, third);
  addMultilineBox(doc, 'REMARKS', data.remarks, 28, labourY + 82, pageWidth, 70);

  const sigY = labourY + 178;
  if (sigY > 620) doc.addPage();
  const finalSigY = sigY > 620 ? 28 : sigY;
  doc.font('Helvetica-Bold').fontSize(12).fillColor('#0f304d').text('SIGNATURES', 28, finalSigY);
  drawSignature(doc, 'CUSTOMER SIGNATURE AND CHOP', data.signatures?.customerSignature, 28, finalSigY + 22, col, 115);
  drawSignature(doc, 'TECHNICIAN SIGNATURE', data.signatures?.technicianSignature, 28 + col + 18, finalSigY + 22, col, 115);

  doc.end();
});

const getGraphToken = async () => {
  if (graphTokenCache && graphTokenCache.expiresOnTimestamp > Date.now() + 120000) {
    return graphTokenCache.token;
  }
  const credential = new DefaultAzureCredential();
  const token = await credential.getToken(GRAPH_SCOPE);
  graphTokenCache = token;
  return token.token;
};

const graphRequest = async (url, options = {}) => {
  const token = await getGraphToken();
  const response = await fetch(`${GRAPH_BASE_URL}${url}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(options.headers || {}),
    },
  });

  if (response.ok) {
    if (response.status === 204) return null;
    return response.json();
  }

  const detail = await response.text();
  throw new Error(`OneDrive upload failed (${response.status}). ${detail}`);
};

const encodeSharePointPath = (value) => String(value)
  .split('/')
  .filter(Boolean)
  .map(encodeURIComponent)
  .join('/');

const resolveDrive = async () => {
  if (graphDriveCache) return graphDriveCache;

  const sitePath = GRAPH_SITE_PATH === '/' ? '' : `:${GRAPH_SITE_PATH}`;
  const site = await graphRequest(`/sites/${GRAPH_SITE_HOSTNAME}${sitePath}`);
  const drives = await graphRequest(`/sites/${site.id}/drives`);
  const drive = (drives.value || []).find((item) => item.name === GRAPH_DOCUMENT_LIBRARY)
    || (drives.value || [])[0];

  if (!drive) {
    throw new Error(`Could not find SharePoint document library "${GRAPH_DOCUMENT_LIBRARY}".`);
  }

  graphDriveCache = { siteId: site.id, driveId: drive.id };
  return graphDriveCache;
};

const ensureFolderPath = async (driveId, folderPath) => {
  const parts = String(folderPath || '').split('/').map((part) => part.trim()).filter(Boolean);
  let currentPath = '';

  for (const part of parts) {
    const parentPath = currentPath ? `root:/${encodeSharePointPath(currentPath)}:` : 'root';
    try {
      await graphRequest(`/drives/${driveId}/${parentPath}/children`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name: part,
          folder: {},
          '@microsoft.graph.conflictBehavior': 'fail',
        }),
      });
    } catch (error) {
      if (!String(error.message).includes('nameAlreadyExists')) throw error;
    }
    currentPath = currentPath ? `${currentPath}/${part}` : part;
  }
};

const uploadPdfToOneDrive = async (fileName, pdfBuffer) => {
  const { driveId } = await resolveDrive();
  await ensureFolderPath(driveId, GRAPH_FOLDER_PATH);
  const uploadPath = encodeSharePointPath(`${GRAPH_FOLDER_PATH}/${fileName}`);
  const file = await graphRequest(`/drives/${driveId}/root:/${uploadPath}:/content`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/pdf' },
    body: pdfBuffer,
  });
  return file;
};

const appendCsv = async (data, pdfFile, jsonFile) => {
  await ensureReportsDir();
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
      'PDF File',
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
    pdfFile,
    jsonFile,
  ].map(csvCell).join(',') + os.EOL);
};

const saveLocalBackup = async (jsonFile, data, pdfFile, pdfBuffer) => {
  await ensureReportsDir();
  await fs.writeFile(path.join(REPORTS_DIR, jsonFile), JSON.stringify(data, null, 2), 'utf8');
  await fs.writeFile(path.join(REPORTS_DIR, pdfFile), pdfBuffer);
};

const saveSubmission = async (payload) => {
  const nextNumber = await readSequence();
  const jobOrder = formatJobOrder(nextNumber);
  const nextJobOrder = nextNumber + 1;
  const submittedAt = new Date().toISOString();
  const companyPart = cleanFilePart(payload.company) || 'No Company';
  const datePart = formatDateForFile(payload.date, new Date(submittedAt));
  const baseName = `JO ${jobOrder} _ ${companyPart} _ ${datePart}`;
  const jsonFile = `${baseName}.json`;
  const pdfFile = `${baseName}.pdf`;
  const data = { ...payload, jobOrder, submittedAt };

  const pdfBuffer = await createPdfBuffer(data);
  const uploadedFile = await uploadPdfToOneDrive(pdfFile, pdfBuffer);
  await saveLocalBackup(jsonFile, data, pdfFile, pdfBuffer);
  await appendCsv(data, pdfFile, jsonFile);
  await writeSequence(nextJobOrder);

  return {
    jobOrder,
    nextJobOrder: formatJobOrder(nextJobOrder),
    pdfFile,
    oneDriveFolder: `${GRAPH_DOCUMENT_LIBRARY}/${GRAPH_FOLDER_PATH}`,
    webUrl: uploadedFile.webUrl,
  };
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
    response.writeHead(200, {
      'Content-Type': mimeTypes[path.extname(filePath).toLowerCase()] || 'application/octet-stream',
      'Cache-Control': 'no-store',
    });
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
  console.log(`Reports upload to: ${GRAPH_DOCUMENT_LIBRARY}/${GRAPH_FOLDER_PATH}`);
});
