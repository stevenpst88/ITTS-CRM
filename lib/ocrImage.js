'use strict';
/**
 * 名片 OCR 上傳的圖片檢查（SEC-09）。業務版（POST /api/ai/ocr-card）與管理員版（POST /api/admin/ai-ocr-card）共用。
 *
 * 為什麼：業務版原本的 multer 只有 8MB 上限、沒有任何類型檢查，PDF、文字檔、任何檔案都會被送進 Gemini 並扣 AI 額度；
 * 管理員版只看 MIME 是不是 image/*（MIME 是用戶端自己填的，可以亂寫）。
 *
 * 兩道關卡（都在「扣 AI 額度、呼叫 Gemini」之前）：
 *   1. fileFilter：MIME 與副檔名都必須在白名單內（只接受 jpg／jpeg／png／webp，以及 Gemini 同樣支援、手機相機常見的 heic／heif）
 *   2. 檔頭 magic bytes：JPEG FF D8 FF、PNG 89 50 4E 47 0D 0A 1A 0A、WEBP 'RIFF'????'WEBP'、HEIC／HEIF 'ftyp'＋已知 brand。
 *      送進 Gemini 的 MIME 一律用「檔頭偵測到的類型」，不用用戶端宣稱的。
 * 純函式、沒有依賴，scripts/check-ocr-filter.js 直接測。
 */
const path = require('path');

const MIMES = new Set(['image/jpeg', 'image/jpg', 'image/png', 'image/webp', 'image/heic', 'image/heif']);
const EXTS = new Set(['.jpg', '.jpeg', '.png', '.webp', '.heic', '.heif']);
const BAD_IMAGE_MSG = '只支援 JPG、PNG、WEBP、HEIC 圖片檔，請重新選擇名片照片';
const HEIF_BRANDS = new Set(['heic', 'heix', 'hevc', 'hevx', 'heim', 'heis', 'mif1', 'msf1']);

/** MIME 與副檔名都必須在白名單內（兩者各自檢查；不要求互相對應，因為前端壓縮後可能是 jpeg 內容配原始檔名）。 */
function isAllowedUpload(file) {
  const mime = String((file && file.mimetype) || '').toLowerCase().split(';')[0].trim();
  const ext = path.extname(String((file && file.originalname) || '')).toLowerCase();
  return MIMES.has(mime) && EXTS.has(ext);
}

/** 讀檔頭判斷真實圖片類型；不是支援的圖片回 null。 */
function sniffImageType(buf) {
  if (!Buffer.isBuffer(buf) || buf.length < 12) return null;
  if (buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff) return 'image/jpeg';
  if (buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47 && buf[4] === 0x0d && buf[5] === 0x0a && buf[6] === 0x1a && buf[7] === 0x0a) return 'image/png';
  if (buf.toString('latin1', 0, 4) === 'RIFF' && buf.toString('latin1', 8, 12) === 'WEBP') return 'image/webp';
  if (buf.toString('latin1', 4, 8) === 'ftyp' && HEIF_BRANDS.has(buf.toString('latin1', 8, 12).toLowerCase())) return 'image/heic';
  return null;
}

/** multer 的 fileFilter：MIME／副檔名不合就丟帶 code 的錯誤（由呼叫端轉成 400 JSON）。 */
function fileFilter(req, file, cb) {
  if (isAllowedUpload(file)) return cb(null, true);
  const e = new Error(BAD_IMAGE_MSG);
  e.code = 'OCR_BAD_TYPE';
  cb(e);
}

module.exports = { MIMES, EXTS, BAD_IMAGE_MSG, isAllowedUpload, sniffImageType, fileFilter };
