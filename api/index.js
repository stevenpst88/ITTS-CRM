// ════════════════════════════════════════════════════════════
//  Vercel Serverless Entry
//
//  把整個 Express app 包成一個 serverless function，
//  Vercel 會把所有 /api/* 以及（透過 vercel.json rewrites）其他路由
//  都導向這裡。
//
//  這個檔不含任何業務邏輯，只負責轉交。
// ════════════════════════════════════════════════════════════
const app = require('../server');
module.exports = app;
