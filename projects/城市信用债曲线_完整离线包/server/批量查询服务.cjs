'use strict';
// Node.js 18+，仅依赖标准库；仅监听本机，请求文本不落盘，凭证从本地 Markdown 读取。
const http = require('node:http');
const fs = require('node:fs/promises');
const path = require('node:path');
const { validateItems } = require('../assets/批量查询.js');
const { createSnapshotStore } = require('./快照存储.cjs');
const HOST = '127.0.0.1';
const DEFAULT_PORT = 8765;
const MAX_TEXT = 60000;

function json(res, status, value) {
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
  res.end(JSON.stringify(value));
}
async function readBody(req, limit = 300000) {
  let size = 0;
  const chunks = [];
  for await (const chunk of req) {
    size += chunk.length;
    if (size > limit) { const error = Error('输入文件或文本过大，请减少内容。'); error.status = 413; throw error; }
    chunks.push(chunk);
  }
  try { return JSON.parse(Buffer.concat(chunks).toString('utf8')); }
  catch { const error = Error('请求 JSON 格式无效。'); error.status = 400; throw error; }
}
function createServer({ fetchImpl = globalThis.fetch, env = process.env, root = path.resolve(__dirname, '..'), storageRoot = root } = {}) {
  let busy = false, saving = false;
  const store = createSnapshotStore(storageRoot);
  async function readConfig() {
    let config = {};
    try {
      const markdown = await fs.readFile(path.join(root, 'DeepSeek配置.md'), 'utf8');
      const block = /```json\s*([\s\S]*?)```/.exec(markdown);
      if (!block) throw Error('配置缺少 JSON 段');
      config = JSON.parse(block[1]);
      if (typeof config.apiKey !== 'string' || typeof config.model !== 'string') throw Error('配置字段类型错误');
    } catch (error) {
      if (error.code !== 'ENOENT') throw Error('请检查 DeepSeek配置.md 中的 JSON 格式。');
    }
    return { apiKey: env.DEEPSEEK_API_KEY || config.apiKey || '', model: env.DEEPSEEK_MODEL || config.model || 'deepseek-flash' };
  }
  return http.createServer(async (req, res) => {
    try {
      // 拒绝跨站和 DNS rebinding 请求；不对 file:// 或其他网站开放 CORS。
      const expectedHost = HOST + ':' + req.socket.localPort;
      if (req.headers.host !== expectedHost) return json(res, 403, { error: '请使用本机 127.0.0.1 地址。' });
      if (req.headers.origin && req.headers.origin !== 'http://' + expectedHost) return json(res, 403, { error: '不允许其他网站调用批量服务。' });
      const url = new URL(req.url, 'http://' + expectedHost);
      if (req.method === 'GET' && url.pathname === '/api/snapshot') return json(res, 200, await store.read());
      if (req.method === 'POST' && url.pathname === '/api/snapshot') {
        if (!/^application\/json\b/i.test(req.headers['content-type'] || '')) return json(res, 415, { error: '请使用 JSON 请求。' });
        const body = await readBody(req, 80 * 1024 * 1024);
        if (saving) return json(res, 409, { error: '正在保存另一份数据，请稍后重试。' });
        saving = true;
        try { return json(res, 200, await store.save(body)); }
        catch (error) { return json(res, error.status || 400, { error: error.status ? error.message : '快照未保存：' + error.message }); }
        finally { saving = false; }
      }
      if (req.method === 'GET' && url.pathname === '/api/status') {
        const config = await readConfig();
        return json(res, 200, { configured: Boolean(config.apiKey), model: config.model });
      }
      if (req.method === 'POST' && url.pathname === '/api/parse') {
        if (!/^application\/json\b/i.test(req.headers['content-type'] || '')) return json(res, 415, { error: '请使用 JSON 请求。' });
        if (busy) return json(res, 429, { error: '正在解析上一批文本，请稍后重试。' });
        const body = await readBody(req);
        if (typeof body.text !== 'string' || !body.text.trim() || body.text.length > MAX_TEXT) return json(res, 400, { error: '请粘贴报价文本，每批最多 60,000 个字符。' });
        const config = await readConfig();
        const key = typeof body.apiKey === 'string' && body.apiKey.trim() ? body.apiKey.trim() : config.apiKey;
        if (!key) return json(res, 400, { error: '请在 DeepSeek配置.md 中填写 API Key，或在页面 API 设置中临时填写。' });
        if (busy) return json(res, 429, { error: '正在解析上一批文本，请稍后重试。' });
        busy = true;
        try {
          const prompt = await fs.readFile(path.join(root, '批量查询解析规则.md'), 'utf8');
          const upstream = await fetchImpl('https://api.deepseek.com/chat/completions', {
            method: 'POST',
            headers: { Authorization: 'Bearer ' + key, 'Content-Type': 'application/json' },
            body: JSON.stringify({ model: config.model, messages: [{ role: 'system', content: prompt }, { role: 'user', content: body.text }], response_format: { type: 'json_object' }, thinking: { type: 'disabled' }, max_tokens: 16000, stream: false }),
            signal: AbortSignal.timeout(120000)
          });
          if (!upstream.ok) {
            const messages = { 400: 'DeepSeek 请求参数无效，请检查 DeepSeek配置.md 中的模型名。', 401: 'DeepSeek 密钥无效，请检查 API 设置。', 402: 'DeepSeek 账户余额不足。', 429: 'DeepSeek 请求繁忙或达到限额，请稍后重试。' };
            return json(res, 502, { error: messages[upstream.status] || 'DeepSeek 服务返回错误（HTTP ' + upstream.status + '），请稍后重试。' });
          }
          const data = await upstream.json();
          const choice = data.choices && data.choices[0];
          if (!choice || choice.finish_reason !== 'stop' || !choice.message?.content) return json(res, 502, { error: '模型输出为空或未完整结束，请减少本批条数后重试。' });
          let parsed;
          try { parsed = JSON.parse(choice.message.content); } catch { return json(res, 502, { error: '模型未返回有效 JSON，请重试。' }); }
          try { return json(res, 200, { items: validateItems(parsed, body.text) }); }
          catch (error) { return json(res, 502, { error: '模型结果未通过核对：' + error.message + ' 请检查文本或分批重试。' }); }
        } finally { busy = false; }
      }
      if (req.method !== 'GET') return json(res, 405, { error: '不支持的请求方式。' });
      const files = { '/': ['双击打开.html', 'text/html'], '/双击打开.html': ['双击打开.html', 'text/html'], '/assets/批量查询.js': ['assets/批量查询.js', 'text/javascript'], '/assets/快照数据.js': ['assets/快照数据.js', 'text/javascript'], '/assets/数据更新.js': ['assets/数据更新.js', 'text/javascript'], '/批量查询解析规则.md': ['批量查询解析规则.md', 'text/markdown'] };
      const file = files[decodeURIComponent(url.pathname)];
      if (!file) return json(res, 404, { error: '文件不存在。' });
      const content = await fs.readFile(path.join(root, file[0]));
      res.writeHead(200, { 'Content-Type': file[1] + '; charset=utf-8', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff', 'Referrer-Policy': 'no-referrer', 'Content-Security-Policy': "default-src 'self' data: blob:; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; connect-src 'self'; frame-ancestors 'none'" });
      res.end(content);
    } catch (error) {
      if (!res.headersSent) json(res, error.status || 500, { error: error.status ? error.message : error.name === 'TimeoutError' ? 'DeepSeek 解析超时，请减少本批条数后重试。' : error.message.startsWith('请检查 DeepSeek配置.md') ? error.message : '本地服务处理失败，请检查网络连接或输入后重试。' });
      else res.end();
    }
  });
}
function startServer() {
  const port = Number(process.env.BOND_BATCH_PORT || DEFAULT_PORT);
  const server = createServer();
  server.on('error', error => { console.error(error.code === 'EADDRINUSE' ? '端口已占用。若服务已启动，请打开 http://127.0.0.1:' + port : '启动失败：' + error.message); process.exitCode = 1; });
  server.listen(port, HOST, () => {
    console.log('债券查询已启动：http://' + HOST + ':' + port + '\n请保持本窗口运行。按 Ctrl+C 停止。');
    if (process.argv.includes('--open')) require('node:child_process').spawn('explorer.exe', ['http://' + HOST + ':' + port], { windowsHide: true, stdio: 'ignore' }).on('error', () => {});
  });
}
if (require.main === module) startServer();
module.exports = { createServer, startServer };
