// ════════════════════════════════════════════════════════════
//  本地檔案系統 Storage（開發用，保留原行為）
// ════════════════════════════════════════════════════════════
const fs     = require('fs');
const path   = require('path');
const multer = require('multer');
const { v4: uuidv4 } = require('uuid');

const UPLOAD_DIR = path.join(__dirname, '..', 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

function getMulterStorage() {
  return multer.diskStorage({
    destination: (req, file, cb) => cb(null, UPLOAD_DIR),
    filename:    (req, file, cb) => {
      const ext = path.extname(file.originalname);
      cb(null, uuidv4() + ext);
    },
  });
}

function serveFile(req, res, key) {
  const safe = path.basename(key); // 防 path traversal
  const abs  = path.join(UPLOAD_DIR, safe);
  if (!fs.existsSync(abs)) return res.status(404).end();
  res.sendFile(abs);
}

function getPublicUrl(key) {
  return `/uploads/${key}`;
}

async function deleteFile(key) {
  const safe = path.basename(key);
  const abs  = path.join(UPLOAD_DIR, safe);
  if (fs.existsSync(abs)) fs.unlinkSync(abs);
}

module.exports = { getMulterStorage, serveFile, getPublicUrl, deleteFile };
