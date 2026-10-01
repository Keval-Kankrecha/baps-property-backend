const express = require('express');
const multer = require('multer');
const path = require('path');
const crypto = require('crypto');
const fs = require('fs');

const MAX_FILES = 10;

const UPLOAD_DIR = path.join(__dirname, '..', '..', 'uploads');
const ALLOWED = {
  'application/pdf': '.pdf',
  'image/jpeg': '.jpg',
  'image/png': '.png',
  'image/webp': '.webp',
};

const storage = multer.diskStorage({
  destination: UPLOAD_DIR,
  filename: (req, file, cb) => cb(null, crypto.randomUUID() + ALLOWED[file.mimetype]),
});

const upload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) =>
    ALLOWED[file.mimetype] ? cb(null, true) : cb(new Error('ફક્ત PDF, JPG, PNG અથવા WEBP ફાઇલ માન્ય છે')),
});

const router = express.Router();

// POST /api/uploads  (multipart, field name "files", up to 10 files) → array of file records
router.post('/', (req, res) => {
  upload.array('files', MAX_FILES)(req, res, (err) => {
    if (err) {
      // multer may already have written some files of this request — don't leave them behind
      for (const f of req.files || []) fs.unlink(f.path, () => {});
      const message =
        err.code === 'LIMIT_FILE_SIZE' ? 'ફાઇલ 10 MB થી વધુ ન હોવી જોઈએ'
        : err.code === 'LIMIT_UNEXPECTED_FILE' ? `એક સમયે વધુમાં વધુ ${MAX_FILES} ફાઇલ`
        : err.message;
      return res.status(400).json({ message });
    }
    if (!req.files?.length) return res.status(400).json({ message: 'ફાઇલ પસંદ કરો' });
    res.status(201).json(
      req.files.map((f) => ({
        fileName: f.filename,
        originalName: Buffer.from(f.originalname, 'latin1').toString('utf8'),
        mimeType: f.mimetype,
        size: f.size,
      }))
    );
  });
});

module.exports = { router, UPLOAD_DIR };
