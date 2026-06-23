// ── secretBox：可逆對稱加密（AES-256-GCM）──────────────────────────
// 用途：把異質系統整合（SAP 等）的機敏連線資料（密碼 / client secret）
//       加密後存進 data（data.json 或 Supabase app_data JSONB），避免明文落地。
//
// 為什麼不用 bcrypt：bcrypt 是單向雜湊，無法還原——但我們之後要拿密碼去打 SAP，
//                  必須能解密，所以改用 AES-256-GCM（附完整性驗證 tag）。
//
// 金鑰來源（依序）：
//   1. 環境變數 SAP_ENCRYPTION_KEY（64 字 hex = 32 bytes，建議正式環境用獨立金鑰）
//      產生：node -e "console.log(require('crypto').randomBytes(32).toString('hex'))"
//   2. 後備：由 JWT_SECRET / SESSION_SECRET 派生（sha256，加命名空間前綴）
//      —— 免額外設定即可運作；缺點是若 JWT_SECRET 輪替，舊密文需重新輸入。
//
// 密文格式：  v1:<base64(iv)>:<base64(authTag)>:<base64(ciphertext)>
const crypto = require('crypto');

function getKey() {
  const hex = process.env.SAP_ENCRYPTION_KEY;
  if (hex && /^[0-9a-fA-F]{64}$/.test(hex.trim())) {
    return Buffer.from(hex.trim(), 'hex');          // 32 bytes
  }
  const base = process.env.JWT_SECRET || process.env.SESSION_SECRET;
  if (!base) {
    throw new Error('[secretBox] 無 SAP_ENCRYPTION_KEY，且缺 JWT_SECRET/SESSION_SECRET 可派生金鑰');
  }
  return crypto.createHash('sha256').update('sap-integration-secretbox:' + base).digest(); // 32 bytes
}

function isEncrypted(v) {
  return typeof v === 'string' && v.startsWith('v1:');
}

// 加密；空值回傳空字串（呼叫端用「空＝不變更」語意）
function encrypt(plaintext) {
  if (plaintext == null || plaintext === '') return '';
  if (isEncrypted(plaintext)) return plaintext;     // 已是密文則原樣保留（避免重複加密）
  const key = getKey();
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
  const ct = Buffer.concat([cipher.update(String(plaintext), 'utf8'), cipher.final()]);
  const tag = cipher.getAuthTag();
  return 'v1:' + [iv, tag, ct].map(b => b.toString('base64')).join(':');
}

// 解密；非 v1 密文（或空）回傳空字串。金鑰不符 / tag 驗證失敗會 throw，由呼叫端 try/catch。
function decrypt(payload) {
  if (!isEncrypted(payload)) return '';
  const parts = payload.split(':');
  if (parts.length !== 4) return '';
  const key = getKey();
  const iv = Buffer.from(parts[1], 'base64');
  const tag = Buffer.from(parts[2], 'base64');
  const ct = Buffer.from(parts[3], 'base64');
  const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv);
  decipher.setAuthTag(tag);
  return Buffer.concat([decipher.update(ct), decipher.final()]).toString('utf8');
}

module.exports = { encrypt, decrypt, isEncrypted };
