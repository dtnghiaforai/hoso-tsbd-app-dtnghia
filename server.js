// server.js — robust, compact, compatible with Index Compact
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const expressions = require('angular-expressions');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Multer in-memory with size limit (30 MB)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 30 * 1024 * 1024 }
});

// Basic middleware
app.use(cors());
app.use(express.json({ limit: '1mb' }));
app.use(express.urlencoded({ extended: true }));

// *** START CHANGE: Serve static files from 'public' folder ***
app.use(express.static(path.join(__dirname, 'public')));
// *** END CHANGE ***

// Helpers
const ensureString = v => (v === null || v === undefined || (typeof v === 'object' && Object.keys(v).length === 0)) ? '' : String(v);
const encodeFilename = n => encodeURIComponent(n || 'document.docx');

// angular-expressions filters (useful in templates)
expressions.filters.upper = s => ensureString(s).toUpperCase();
expressions.filters.lower = s => ensureString(s).toLowerCase();
expressions.filters.capitalize = s => { const t = ensureString(s); return t ? t.charAt(0).toUpperCase() + t.slice(1).toLowerCase() : t; };
expressions.filters.number = (val, frac = 0) => {
  const n = parseFloat(val);
  if (isNaN(n)) return ensureString(val);
  return new Intl.NumberFormat('vi-VN', { minimumFractionDigits: +frac, maximumFractionDigits: +frac }).format(n);
};
// *** NEW FILTER START ***
expressions.filters.numflex = (val) => {
    const n = parseFloat(val);
    if (isNaN(n)) return ensureString(val);
    // This format automatically handles integers vs decimals correctly for Vietnamese locale.
    // It will not add trailing zeros for whole numbers, but will keep decimals if they exist.
    return new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 20 }).format(n);
};
// *** NEW FILTER END ***
expressions.filters.currency = v => {
  const n = parseFloat(v);
  if (isNaN(n)) return ensureString(v);
  return new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND', minimumFractionDigits: 0 }).format(n);
};
expressions.filters.date = (input, fmt = 'dd/MM/yyyy') => {
  let s = ensureString(input);
  if (/^\d{1,2}[\/-]\d{1,2}[\/-]\d{4}$/.test(s)) {
    const p = s.split(/[\/-]/); s = `${p[1]}/${p[0]}/${p[2]}`;
  }
  const d = new Date(s);
  if (isNaN(d.getTime())) return ensureString(input);
  const dd = String(d.getDate()).padStart(2, '0'), mm = String(d.getMonth() + 1).padStart(2, '0'), yyyy = d.getFullYear();
  return fmt.replace('dd', dd).replace('MM', mm).replace('yyyy', yyyy);
};
expressions.filters.limitTo = (v, n) => ensureString(v).slice(0, Math.max(0, parseInt(n || 0, 10)));
expressions.filters.json = v => {
  try { return JSON.stringify(v, null, 2); } catch { return ensureString(v); }
};

// Health check
app.get('/health', (req, res) => res.json({ ok: true, ts: new Date().toISOString() }));

// Save data endpoint (return JSON file for download)
app.post('/api/save-data', (req, res) => {
  try {
    const data = req.body.data || {};
    let fileName = req.body.fileName || `DuLieuHoSoTSBD_${new Date().toISOString().slice(0, 10)}.json`;
    if (!fileName.toLowerCase().endsWith('.json')) fileName += '.json';
    const buf = Buffer.from(JSON.stringify(data, null, 2), 'utf8');
    const enc = encodeFilename(fileName);
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="${enc}"; filename*=UTF-8''${enc}`);
    return res.send(buf);
  } catch (err) {
    console.error('[save-data] error', err);
    return res.status(500).json({ message: 'Lỗi khi lưu dữ liệu', details: String(err) });
  }
});

// Generate DOCX (mailmerge)
// Expect FormData: 'template' file + 'data' JSON string + optional 'fileName'
app.post('/generate', upload.single('template'), (req, res) => {
  const start = Date.now();
  if (!req.file) {
    console.warn('[generate] no file');
    return res.status(400).json({ message: 'Không tìm thấy file template.' });
  }

  // Handle client abort
  req.on('aborted', () => {
    console.warn('[generate] client aborted the request');
  });

  // Parse JSON payload safely
  let formData = {};
  if (req.body && req.body.data) {
    try { formData = JSON.parse(req.body.data); } catch (e) {
      console.warn('[generate] invalid JSON in req.body.data, using empty object', e.message);
      formData = {};
    }
  }

  try {
    // Create docxtemplater instance
    const zip = new PizZip(req.file.buffer);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      parser: tag => ({ get: scope => expressions.compile(tag)(scope) })
    });

    doc.setData(formData);

    try {
      doc.render();
    } catch (renderErr) {
      console.error('[generate] render error:', renderErr);
      const details = renderErr.properties && renderErr.properties.errors ? renderErr.properties.errors : (renderErr.message || String(renderErr));
      return res.status(500).json({ message: 'Lỗi khi render template DOCX (kiểm tra placeholders).', details });
    }

    const outputBuffer = doc.getZip().generate({ type: 'nodebuffer' });

    // Prepare output filename
    const templateBase = (req.file.originalname || 'template').replace(/\.docx$/i, '');
    const outName = req.body.fileName && String(req.body.fileName).trim()
      ? (String(req.body.fileName).endsWith('.docx') ? String(req.body.fileName) : `${String(req.body.fileName)}.docx`)
      : `HoSoTSBD_${templateBase}_${new Date().toISOString().slice(0, 10)}.docx`;

    const enc = encodeFilename(outName);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${enc}"; filename*=UTF-8''${enc}`);
    res.send(outputBuffer);

    console.log(`[generate] success: ${outName} (${Date.now() - start}ms)`);
  } catch (err) {
    console.error('[generate] unexpected error:', err);
    return res.status(500).json({ message: 'Lỗi server khi tạo DOCX', details: String(err) });
  }
});

// Multer / general error handler (must be after routes)
app.use((err, req, res, next) => {
  if (err) {
    console.error('[error-middleware]', err);
    // Multer file size limit error
    if (err.code === 'LIMIT_FILE_SIZE') {
      return res.status(413).json({ message: 'File quá lớn. Giới hạn 30MB.', details: err.message });
    }
    // other multer errors
    if (err instanceof multer.MulterError) {
      return res.status(400).json({ message: 'Lỗi upload file', details: err.message });
    }
    return res.status(500).json({ message: 'Lỗi server không xác định', details: String(err) });
  }
  next();
});

// Start server
app.listen(PORT, () => console.log(`Server chạy tại http://localhost:${PORT} (PID ${process.pid})`));
