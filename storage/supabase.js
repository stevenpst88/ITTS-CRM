// ════════════════════════════════════════════════════════════
//  Supabase Storage 實作（Serverless 正式用）
//
//  策略：用 multer memoryStorage 暫存，然後上傳到 Supabase Storage
//  透過 service_role key 做 server-side 上傳，繞過 RLS
// ════════════════════════════════════════════════════════════
const multer = require('multer');
const { v4: uuidv4 } = require('uuid');
const path = require('path');
const { createClient } = require('@supabase/supabase-js');

const SUPABASE_URL     = process.env.SUPABASE_URL;
const SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const BUCKET           = process.env.SUPABASE_BUCKET || 'uploads';

if (!SUPABASE_URL || !SERVICE_ROLE_KEY) {
  throw new Error('[storage/supabase] 缺少 SUPABASE_URL 或 SUPABASE_SERVICE_ROLE_KEY');
}

const supabase = createClient(SUPABASE_URL, SERVICE_ROLE_KEY, {
  auth: { persistSession: false },
});

// ── Multer storage engine：存 memory，_handleFile 內送 Supabase ──
function getMulterStorage() {
  return {
    _handleFile(req, file, cb) {
      const chunks = [];
      file.stream.on('data', (c) => chunks.push(c));
      file.stream.on('error', cb);
      file.stream.on('end', async () => {
        try {
          const buffer = Buffer.concat(chunks);
          const ext    = path.extname(file.originalname).toLowerCase();
          const key    = uuidv4() + ext;

          const { error } = await supabase.storage
            .from(BUCKET)
            .upload(key, buffer, {
              contentType: file.mimetype,
              upsert:      false,
            });
          if (error) return cb(error);

          cb(null, {
            filename: key,
            path:     key,
            size:     buffer.length,
            mimetype: file.mimetype,
          });
        } catch (err) {
          cb(err);
        }
      });
    },
    _removeFile(req, file, cb) {
      supabase.storage.from(BUCKET).remove([file.filename])
        .then(() => cb(null))
        .catch(cb);
    },
  };
}

// ── 伺服器端 proxy 讀檔（給 /uploads/:key 用）──
async function serveFile(req, res, key) {
  const safe = path.basename(key);
  const { data, error } = await supabase.storage.from(BUCKET).download(safe);
  if (error || !data) return res.status(404).end();

  const arrayBuffer = await data.arrayBuffer();
  const ext = path.extname(safe).toLowerCase();
  const mime = {
    '.jpg':'image/jpeg', '.jpeg':'image/jpeg', '.png':'image/png',
    '.gif':'image/gif',  '.webp':'image/webp',
  }[ext] || 'application/octet-stream';

  res.set('Content-Type', mime);
  res.set('Cache-Control', 'private, max-age=3600');
  res.send(Buffer.from(arrayBuffer));
}

function getPublicUrl(key) {
  // 保持跟本地一致的 URL 格式，由 server 代理下載（確保受 requireAuth 保護）
  return `/uploads/${key}`;
}

async function deleteFile(key) {
  const safe = path.basename(key);
  await supabase.storage.from(BUCKET).remove([safe]);
}

module.exports = { getMulterStorage, serveFile, getPublicUrl, deleteFile };
