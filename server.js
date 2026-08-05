const http = require('http');
const https = require('https');
const fs = require('fs/promises');
const path = require('path');
const os = require('os');

const PORT = Number(process.env.PORT || 4322);
const HOST = process.env.HOST || '0.0.0.0';
const ROOT = __dirname;
const REPORTS_DIR = path.join(ROOT, 'Reports');
const SEQUENCE_FILE = path.join(REPORTS_DIR, '.sequence.json');
const LOG_FILE = path.join(REPORTS_DIR, 'Job_Order_Log.csv');

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
const GRAPH_SITE_HOSTNAME = process.env.GRAPH_SITE_HOSTNAME || 'jnwengineering.sharepoint.com';
const GRAPH_SITE_PATH = process.env.GRAPH_SITE_PATH || '/';
const GRAPH_DOCUMENT_LIBRARY = process.env.GRAPH_DOCUMENT_LIBRARY || 'Documents';
const GRAPH_FOLDER_PATH = process.env.GRAPH_FOLDER_PATH || 'Mobile Job Order Reports/{year}';
const GRAPH_DRIVE_ID = process.env.GRAPH_DRIVE_ID || '';

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

const getReportYear = (value, fallbackDate = new Date()) => {
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) return match[1];

  const parsed = value ? new Date(value) : fallbackDate;
  const date = Number.isNaN(parsed.getTime()) ? fallbackDate : parsed;
  return String(date.getFullYear());
};

const buildReportFolderPath = (data) => GRAPH_FOLDER_PATH
  .replaceAll('{year}', getReportYear(data.date, new Date(data.submittedAt)))
  .replaceAll('{yyyy}', getReportYear(data.date, new Date(data.submittedAt)));

const normaliseParts = (parts) => Array.isArray(parts)
  ? parts.map((part, index) => ({
      serial: part.serial || String(index + 1),
      description: part.description || '',
      quantity: part.quantity || part.charge || '',
    })).filter((part) => part.description || part.quantity)
  : [];

const webRequest = (url, options = {}) => new Promise((resolve, reject) => {
  const target = new URL(url);
  const transport = target.protocol === 'http:' ? http : https;
  const body = options.body
    ? (Buffer.isBuffer(options.body) ? options.body : Buffer.from(String(options.body)))
    : null;

  const request = transport.request(target, {
    method: options.method || 'GET',
    headers: {
      ...(options.headers || {}),
      ...(body ? { 'Content-Length': body.length } : {}),
    },
  }, (response) => {
    const chunks = [];
    response.on('data', (chunk) => chunks.push(chunk));
    response.on('end', () => {
      const buffer = Buffer.concat(chunks);
      resolve({
        ok: response.statusCode >= 200 && response.statusCode < 300,
        status: response.statusCode,
        text: () => Promise.resolve(buffer.toString('utf8')),
        json: () => Promise.resolve(JSON.parse(buffer.toString('utf8'))),
      });
    });
  });

  request.on('error', reject);
  if (body) request.write(body);
  request.end();
});

const pdfEscape = (value) => String(value || '')
  .replace(/\\/g, '\\\\')
  .replace(/\(/g, '\\(')
  .replace(/\)/g, '\\)')
  .replace(/\r?\n/g, ' ');

const wrapText = (value, maxChars = 88) => {
  const words = String(value || '').replace(/\r?\n/g, ' ').split(/\s+/).filter(Boolean);
  const lines = [];
  let line = '';
  words.forEach((word) => {
    if ((line + ' ' + word).trim().length > maxChars) {
      if (line) lines.push(line);
      line = word;
    } else {
      line = `${line} ${word}`.trim();
    }
  });
  if (line) lines.push(line);
  return lines.length ? lines : [''];
};

const createPdfBuffer = (data) => {
  const lines = [];
  const text = (x, y, size, value, font = 'F1') => {
    lines.push(`BT /${font} ${size} Tf ${x} ${y} Td (${pdfEscape(value)}) Tj ET`);
  };
  const rect = (x, y, w, h) => {
    lines.push(`${x} ${y} ${w} ${h} re S`);
  };
  const label = (x, y, value) => text(x, y, 9, value, 'F2');
  const box = (x, y, w, h, labelText, value, maxChars = 40) => {
    label(x, y + h + 7, labelText);
    rect(x, y, w, h);
    wrapText(value, maxChars).slice(0, Math.max(1, Math.floor(h / 13))).forEach((line, index) => {
      text(x + 7, y + h - 16 - (index * 13), 11, line);
    });
  };

  lines.push('0.5 w');
  text(28, 785, 18, 'JNW ENGINEERING PTE LTD', 'F2');
  text(390, 790, 11, 'JOB ORDER', 'F2');
  text(390, 765, 22, data.jobOrder, 'F2');
  lines.push('28 742 m 567 742 l S');

  box(28, 692, 250, 34, 'COMPANY', data.company);
  box(310, 692, 250, 34, 'DATE', data.date);
  box(28, 640, 250, 34, 'REQUESTED BY', data.requestedBy);
  box(28, 532, 532, 78, 'COMPLAINT', data.complaint, 92);
  box(28, 397, 532, 105, 'ACTION TAKEN', data.actionTaken, 92);

  text(28, 370, 12, 'MATERIALS AND PARTS USED', 'F2');
  label(28, 350, 'S/NO');
  label(85, 350, 'MATERIALS AND PARTS USED');
  label(475, 350, 'QUANTITY');
  let rowY = 316;
  const parts = normaliseParts(data.parts);
  (parts.length ? parts : [{ serial: '1', description: '', quantity: '' }]).slice(0, 6).forEach((part) => {
    box(28, rowY, 45, 26, '', part.serial, 5);
    box(85, rowY, 375, 26, '', part.description, 62);
    box(475, rowY, 85, 26, '', part.quantity, 12);
    rowY -= 32;
  });

  text(28, rowY + 7, 12, 'LABOUR', 'F2');
  box(28, rowY - 45, 165, 32, 'LABOUR DESCRIPTION', data.labourDescription, 25);
  box(210, rowY - 45, 165, 32, 'MAN', data.labourMan, 20);
  box(395, rowY - 45, 165, 32, 'HOURS', data.labourHours, 20);
  box(28, rowY - 145, 532, 68, 'REMARKS', data.remarks, 92);

  text(28, rowY - 174, 12, 'SIGNATURES', 'F2');
  box(28, rowY - 305, 250, 105, 'CUSTOMER SIGNATURE AND CHOP', '', 40);
  box(310, rowY - 305, 250, 105, 'TECHNICIAN SIGNATURE', '', 40);

  const content = lines.join('\n');
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R /F2 5 0 R >> >> /Contents 6 0 R >>',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>',
    `<< /Length ${Buffer.byteLength(content, 'utf8')} >>\nstream\n${content}\nendstream`,
  ];

  let pdf = '%PDF-1.4\n';
  const offsets = [0];
  objects.forEach((object, index) => {
    offsets.push(Buffer.byteLength(pdf, 'utf8'));
    pdf += `${index + 1} 0 obj\n${object}\nendobj\n`;
  });
  const xref = Buffer.byteLength(pdf, 'utf8');
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  offsets.slice(1).forEach((offset) => {
    pdf += `${String(offset).padStart(10, '0')} 00000 n \n`;
  });
  pdf += `trailer << /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`;
  return Buffer.from(pdf, 'utf8');
};

const getGraphToken = async () => {
  if (graphTokenCache && graphTokenCache.expiresOn > Date.now() + 120000) {
    return graphTokenCache.token;
  }

  const resource = encodeURIComponent('https://graph.microsoft.com/');
  const identityEndpoint = process.env.IDENTITY_ENDPOINT || process.env.MSI_ENDPOINT;
  const identityHeader = process.env.IDENTITY_HEADER || process.env.MSI_SECRET;
  const tokenUrl = identityEndpoint
    ? `${identityEndpoint}?api-version=2019-08-01&resource=${resource}`
    : `http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=${resource}`;
  const response = await webRequest(tokenUrl, {
    headers: identityEndpoint
      ? { 'X-IDENTITY-HEADER': identityHeader }
      : { Metadata: 'true' },
  });

  if (!response.ok) {
    throw new Error(`Could not get Azure identity token (${response.status}).`);
  }

  const token = await response.json();
  graphTokenCache = {
    token: token.access_token,
    expiresOn: Number(token.expires_on) * 1000,
  };
  return graphTokenCache.token;
};

const graphRequest = async (url, options = {}) => {
  const token = await getGraphToken();
  const response = await webRequest(`${GRAPH_BASE_URL}${url}`, {
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

  if (GRAPH_DRIVE_ID) {
    graphDriveCache = { driveId: GRAPH_DRIVE_ID, libraryName: 'OneDrive' };
    return graphDriveCache;
  }

  const sitePath = GRAPH_SITE_PATH === '/' ? '' : `:${GRAPH_SITE_PATH}`;
  const site = await graphRequest(`/sites/${GRAPH_SITE_HOSTNAME}${sitePath}`);
  const drives = await graphRequest(`/sites/${site.id}/drives`);
  const drive = (drives.value || []).find((item) => item.name === GRAPH_DOCUMENT_LIBRARY)
    || (drives.value || [])[0];

  if (!drive) {
    throw new Error(`Could not find SharePoint document library "${GRAPH_DOCUMENT_LIBRARY}".`);
  }

  graphDriveCache = { siteId: site.id, driveId: drive.id, libraryName: GRAPH_DOCUMENT_LIBRARY };
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

const uploadPdfToOneDrive = async (fileName, pdfBuffer, folderPath) => {
  const { driveId } = await resolveDrive();
  await ensureFolderPath(driveId, folderPath);
  const uploadPath = encodeSharePointPath(`${folderPath}/${fileName}`);
  return graphRequest(`/drives/${driveId}/root:/${uploadPath}:/content`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/pdf' },
    body: pdfBuffer,
  });
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

  const pdfBuffer = createPdfBuffer(data);
  const reportFolderPath = buildReportFolderPath(data);
  const uploadedFile = await uploadPdfToOneDrive(pdfFile, pdfBuffer, reportFolderPath);
  await saveLocalBackup(jsonFile, data, pdfFile, pdfBuffer);
  await appendCsv(data, pdfFile, jsonFile);
  await writeSequence(nextJobOrder);

  return {
    jobOrder,
    nextJobOrder: formatJobOrder(nextJobOrder),
    pdfFile,
    oneDriveFolder: `${graphDriveCache?.libraryName || GRAPH_DOCUMENT_LIBRARY}/${reportFolderPath}`,
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
