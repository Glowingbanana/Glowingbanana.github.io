/* ========= Config ========= */
const CONFIG = {
  NORMALIZE_CURRENCY_TO: null,
  DATE_FORMAT: 'dd/mm/yyyy',
  EXCLUDE_DRAFT_DEFAULT: false,
  MIN_DIGITAL_TEXT_LEN: 20,
  OCR_SCALE: 2,
  NUMBER_TOLERANCE: 0.02
};

/* ========= Globals ========= */
let selectedFile = null;
let extractedDataRaw = '';
let exportRows = [];
let ocrReady = false;

/* ========= PDF.js Worker ========= */
pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

/* ========= DOM ========= */
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const forceOCRCb = document.getElementById('forceOCR');
const excludeDraftCb = document.getElementById('excludeDraft');
const statusEl = document.getElementById('status');
const fileInfo = document.getElementById('fileInfo');
const fileNameEl = document.getElementById('fileName');
const fileSizeEl = document.getElementById('fileSize');
const preview = document.getElementById('preview');
const previewContent = document.getElementById('previewContent');
const btnConvert = document.getElementById('btnConvert');
const btnDownload = document.getElementById('btnDownload');

/* ========= OCR ========= */
const ocrWorker = Tesseract.createWorker({
  workerPath: 'https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/worker.min.js',
  corePath: 'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.0.0/tesseract-core.wasm.js',
  langPath: 'https://tessdata.projectnaptha.com/4.0.0',
  logger: m => {
    if (m?.status && typeof m.progress === 'number') {
      showStatus(`${m.status} ${(m.progress * 100).toFixed(0)}%`, 'loading');
    }
  }
});

async function ensureOCR() {
  if (ocrReady) return;
  await ocrWorker.load();
  await ocrWorker.loadLanguage('eng');
  await ocrWorker.initialize('eng');
  ocrReady = true;
}

/* ========= UI ========= */
uploadArea.addEventListener('click', () => fileInput.click());
uploadArea.addEventListener('dragover', e => {
  e.preventDefault();
  uploadArea.classList.add('dragover');
});
uploadArea.addEventListener('dragleave', () =>
  uploadArea.classList.remove('dragover')
);
uploadArea.addEventListener('drop', e => {
  e.preventDefault();
  uploadArea.classList.remove('dragover');
  const f = e.dataTransfer.files?.[0];
  if (f) handleFile(f);
});
fileInput.addEventListener('change', e => {
  const f = e.target.files?.[0];
  if (f) handleFile(f);
});

function handleFile(file) {
  if (!(file.type === 'application/pdf' || /\.pdf$/i.test(file.name))) {
    return showStatus('Please select a PDF file', 'error');
  }
  selectedFile = file;
  fileNameEl.textContent = `File: ${file.name}`;
  fileSizeEl.textContent = `Size: ${(file.size / 1024 / 1024).toFixed(2)} MB`;
  fileInfo.classList.add('show');
  btnConvert.disabled = false;
  btnDownload.style.display = 'none';
  preview.classList.remove('show');
  extractedDataRaw = '';
  exportRows = [];
  excludeDraftCb.checked = CONFIG.EXCLUDE_DRAFT_DEFAULT;
  showStatus(`Ready to convert: ${file.name}`, 'success');
}

function showStatus(msg, type = 'loading') {
  statusEl.innerHTML = msg;
  statusEl.className = `status show ${type}`;
}

/* ========= Convert ========= */
async function convertPDF() {
  btnConvert.disabled = true;
  showStatus('Opening PDF…', 'loading');

  try {
    const arrayBuffer = await selectedFile.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    const invoices = new Map();
    let currentInvoiceNo = null;
    extractedDataRaw = '';

    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      let pageText = '';

      if (!forceOCRCb.checked) {
        const tc = await page.getTextContent();
        pageText = tc.items.map(x => x.str).join(' ').replace(/\s+/g, ' ').trim();
      }

      if (
        forceOCRCb.checked ||
        !pageText ||
        pageText.length < CONFIG.MIN_DIGITAL_TEXT_LEN
      ) {
        const vp = page.getViewport({ scale: CONFIG.OCR_SCALE });
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d', { willReadFrequently: true });
        canvas.width = vp.width;
        canvas.height = vp.height;
        await page.render({ canvasContext: ctx, viewport: vp }).promise;
        await ensureOCR();
        const { data } = await ocrWorker.recognize(canvas);
        pageText = data.text.replace(/\s+/g, ' ').trim();
      }

      extractedDataRaw += `\n\n--- Page ${i} ---\n${pageText}`;

      const invNo = findInvoiceNo(pageText);
      if (invNo) currentInvoiceNo = invNo;
      if (!currentInvoiceNo) continue;

      const inv = getOrCreateInvoice(invoices, currentInvoiceNo);
      assignHeader(inv.header, extractHeaderFields(pageText));
      assignTotals(inv.totals, extractTotals(pageText));

      const items = extractLineItems(pageText);
      if (items.length) inv.items.push(...items);
    }

    exportRows = [];
    invoices.forEach((inv, invNo) => {
      if (excludeDraftCb.checked && inv.header.invoiceStatus?.toLowerCase() === 'draft') return;
      computeTotals(inv);

      for (const li of inv.items) {
        exportRows.push([
          inv.header.vendorId || '',
          inv.header.attentionTo || '',
          toExcelDate(inv.header.invoiceDate),
          inv.header.creditTerm || '',
          inv.header.invoiceNo || invNo,
          inv.header.relatedInvoiceNo || '',
          inv.header.invoiceStatus || '',
          inv.header.instructionId || '',
          inv.header.headerDescription || '',
          li.lineNo,
          li.description,
          toNumber(li.quantity),
          toNumber(li.unitPrice),
          toNumber(li.grossEx),
          toNumber(li.gstAmount),
          toNumber(li.grossInc),
          inv.totals.currency || '',
          toNumber(inv.totals.subtotal),
          toNumber(inv.totals.gst)
        ]);
      }
    });

    previewContent.textContent =
      extractedDataRaw.substring(0, 800) +
      (extractedDataRaw.length > 800 ? '…' : '');
    preview.classList.add('show');
    btnDownload.style.display = exportRows.length ? 'block' : 'none';

    showStatus(
      exportRows.length
        ? `Parsed ${exportRows.length} line item row(s).`
        : 'No line items were found.',
      exportRows.length ? 'success' : 'error'
    );
  } catch (err) {
    console.error(err);
    showStatus(`Error: ${err.message}`, 'error');
  } finally {
    btnConvert.disabled = false;
  }
}

/* ========= Download ========= */
function downloadExcel() {
  if (!exportRows.length) return showStatus('No data to download', 'error');

  const headers = [
    'Vendor ID','Attention To','Invoice Date','Credit Term','Invoice No',
    'Related Invoice No','Invoice Status','Invoicing Instruction ID',
    'Description','No.','Description',
    'Quantity','Unit Price','Gross Amt (EX. GST)',
    'GST @ 9%','Gross Amt (Inc. GST)',
    'Currency','Sub Total (Excluding GST)','Total GST Payable'
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...exportRows]);
  XLSX.utils.book_append_sheet(wb, ws, 'Invoice Lines');
  XLSX.writeFile(wb, selectedFile.name.replace(/\.pdf$/i, '') + '.xlsx');
}

/* ========= Helpers ========= */
function normalize(s) {
  return (s || '').replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();
}

function toNumber(v) {
  if (v == null || v === '') return '';
  const n = parseFloat(String(v).replace(/,/g, ''));
  return Number.isFinite(n) ? n : '';
}

function toExcelDate(s) {
  if (!s) return '';
  const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (!m) return '';
  let [, d, mth, y] = m;
  if (y.length === 2) y = '20' + y;
  return new Date(+y, +mth - 1, +d);
}

function getOrCreateInvoice(map, no) {
  if (!map.has(no)) map.set(no, { header: { invoiceNo: no }, items: [], totals: {} });
  return map.get(no);
}

function findInvoiceNo(text) {
  const m = text.match(/Invoice\s*No\s*:\s*([A-Z0-9-]+)/i);
  return m ? m[1] : null;
}

/* ========= Header ========= */
function extractHeaderFields(text) {
  const out = {};
  const grab = (r) => text.match(r)?.[1]?.trim();

  out.vendorId = grab(/Vendor\s*ID\s*:\s*([A-Z0-9]+)/i);
  out.attentionTo = grab(/Attention\s*To\s*:\s*(.+?)(?=\s+Invoice)/i);
  out.invoiceDate = grab(/Invoice\s*Date\s*:\s*([\d\/\-]+)/i);
  out.creditTerm = grab(/Credit\s*Term\s*:\s*(.+?)(?=\s+Invoice)/i);
  out.invoiceNo = grab(/Invoice\s*No\s*:\s*([A-Z0-9-]+)/i);
  out.relatedInvoiceNo = grab(/Related\s*Invoice\s*No\s*:\s*(.+?)(?=\s+Invoice)/i);
  out.invoiceStatus = grab(/Invoice\s*Status\s*:\s*([A-Za-z]+)/i);
  out.instructionId = grab(/Invoicing\s*Instruction\s*ID\s*:\s*([A-Z0-9-]+)/i);
  out.headerDescription = grab(/Description\s*:\s*(.+?)(?=\s+No\.)/i);

  return out;
}

function assignHeader(target, src) {
  for (const k in src) {
    if (src[k] && !target[k]) target[k] = src[k];
  }
}

/* ========= Totals ========= */
function extractTotals(text) {
  return {
    currency: text.match(/Currency\s*:\s*([A-Za-z ]+)/i)?.[1],
    subtotal: text.match(/Sub\s*Total\s*\(Excluding\s*GST\)\s*:\s*([0-9,.]+)/i)?.[1],
    gst: text.match(/Total\s*GST\s*Payable\s*:\s*([0-9,.]+)/i)?.[1]
  };
}

function assignTotals(target, src) {
  for (const k in src) {
    if (src[k] && !target[k]) target[k] = src[k];
  }
}

function computeTotals(inv) {
  let sub = 0, gst = 0;
  inv.items.forEach(li => {
    const ex = toNumber(li.grossEx);
    const inc = toNumber(li.grossInc);
    if (ex) sub += ex;
    if (ex && inc) gst += inc - ex;
  });
  if (!inv.totals.subtotal) inv.totals.subtotal = sub.toFixed(2);
  if (!inv.totals.gst) inv.totals.gst = gst.toFixed(2);
  if (!inv.totals.currency) inv.totals.currency = 'Singapore Dollar';
}

/* ========= ✅ FIXED LINE ITEMS ========= */
function extractLineItems(text) {
  const items = [];
  const rx =
    /(?:^|\n)(\d{1,3})\s+(.+?)\s+([0-9,]+\.\d{2,5})\s+0\s+([0-9]+(?:\.\d+)?)\s+([0-9,]+\.\d{2})\s+([0-9,]+\.\d{2})\s+([0-9,]+\.\d{2})\s*$/gm;

  let m;
  while ((m = rx.exec(text)) !== null) {
    items.push({
      lineNo: parseInt(m[1], 10),
      description: m[2].trim(),
      unitPrice: m[3],
      quantity: m[4],
      grossEx: m[5],
      gstAmount: m[6],
      grossInc: m[7]
    });
  }
  return items;
}
