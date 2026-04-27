const { GoogleGenerativeAI } = require('@google/generative-ai');

let _client = null;

function getClient() {
  if (!process.env.GEMINI_API_KEY) return null;
  if (!_client) _client = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
  return _client;
}

function getModel(model = 'gemini-2.0-flash') {
  const client = getClient();
  if (!client) return null;
  return client.getGenerativeModel({ model });
}

// 清除 Gemini 有時包的 markdown fence，回傳 JSON object
function parseJson(text) {
  const clean = text.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/i, '').trim();
  return JSON.parse(clean);
}

module.exports = {
  getModel,
  parseJson,
  isConfigured: () => !!process.env.GEMINI_API_KEY,
};
