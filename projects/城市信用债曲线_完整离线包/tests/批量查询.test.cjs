'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const fsp = require('node:fs/promises');
const os = require('node:os');
const path = require('node:path');
const vm = require('node:vm');
const B = require('../assets/批量查询.js');
const { createServer } = require('../server/批量查询服务.cjs');
const html = fs.readFileSync(path.join(__dirname, '..', '双击打开.html'), 'utf8');
const scripts = [...html.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
// 仅执行曲线代码和包内既有导出组件，不加载债券明细数据。
const coreScope = { module: { exports: {} } }; vm.runInNewContext(scripts[1], coreScope);
const C = coreScope.module.exports;
const xlsxScope = { module: { exports: {} }, exports: {} }; vm.runInNewContext(scripts[2], xlsxScope);
const X = xlsxScope.exports;
function parsed(overrides = {}) {
  return { id: 1, rawText: '4.77Y AA+/-- 26德财02 520363.SZ --/1.87 --/1000 中债估值1.86 国企 私募 山东省德州市 中债隐含AA(2) YY评级6', name: '26德财02', code: '520363.SZ', termText: '4.77Y', termYears: 4.77, extensionText: null, issuerRating: 'AA+', bondRating: null, bidYield: null, ofrYield: 1.87, valuationType: '估值', valuationYield: 1.86, method: '私募', province: '山东省', city: '德州市', impliedRating: 'AA(2)', yyRating: '6', warnings: [], ...overrides };
}
const bond = (code, term, yieldValue, extra = {}) => ({ code, name: code === '520363.SZ' ? '26德财02' : code, term, yield: yieldValue, province: '山东省', city: '德州市', method: '私募', rating: 'AA+', implied: 'AA(2)', yy: '6', type: '私募公司债', ...extra });
function context() {
  return { core: C, rows: [bond('520363.SZ', 2, 5), bond('a.IB', 1, 1.5), bond('b.IB', 1.2, 1.7), bond('c.IB', 3, 2.5), bond('d.IB', 3.2, 2.7), bond('other.IB', 2, 99, { city: '济南市' })], meta: { date: '2026-09-15', filename: '合成测试数据' }, settings: { filter: { province: '江苏省', city: '南京市', types: ['私募公司债'], yyEnabled: true, yyMin: 4, yyMax: 6, excludeSpecial: true }, termMax: 10, binWidth: 0.5, minBin: 2, excludeTarget: true, quoteDate: '2026-09-15' } };
}
test('解析校验保留正确 OFR，拒绝模型遗漏记录或编造地域', () => {
  const item = parsed(); assert.equal(B.validateItems({ items: [item] }, item.rawText)[0].ofrYield, 1.87);
  assert.throws(() => B.validateItems({ items: [item] }, item.rawText + '\n另一只债券'), /遗漏/);
  assert.throws(() => B.validateItems({ items: [parsed({ city: '南京市' })] }, item.rawText), /city/);
  assert.throws(() => B.validateItems({ items: [parsed({ ofrYield: '1.87' })] }, item.rawText), /数值/);
});
test('OFR 不可用估值或报价量替代；缺失保持 null', () => {
  for (const value of [1.86, 1000]) {
    const item = parsed({ ofrYield: value }); const clean = B.validateItems({ items: [item] }, item.rawText)[0];
    assert.equal(clean.ofrYield, null); assert.match(clean.warnings.join(''), /不一致/);
  }
  const item = parsed({ rawText: parsed().rawText.replace('--/1.87', '--/--'), ofrYield: 1.86 });
  assert.equal(B.validateItems({ items: [item] }, item.rawText)[0].ofrYield, null);
});
test('零收益率保留，YY零评级转缺失，行权保留口径提示', () => {
  const item = parsed({ rawText: parsed().rawText.replace('--/1.87', '--/0'), ofrYield: 0, yyRating: '0', valuationType: '行权', extensionText: 'NY' });
  const clean = B.validateItems({ items: [item] }, item.rawText)[0];
  assert.equal(clean.ofrYield, 0); assert.equal(clean.yyRating, null); assert.match(clean.warnings.join(''), /行权/);
});
test('代码匹配优先，保留代码前导零，不因同名擅自改代码', () => {
  const rows = [bond('032690011.IB', 3, 1.8, { name: '同名债' }), bond('123.SH', 3, 1.8, { name: '同名债' })];
  assert.equal(B.matchBond(rows, { code: '032690011.ib' })[0].code, '032690011.IB');
  assert.equal(B.matchBond(rows, { name: '同名债' }).length, 2);
  assert.equal(B.matchBond(rows, { code: 'missing.IB', name: '同名债' }).length, 0);
});
test('分箱中位数后按期限插值；剔除当前标的，逐券城市独立', () => {
  const ctx = context(), item = parsed({ termYears: 2, ofrYield: 2.3 });
  const r = B.evaluate(ctx, item);
  // 节点(1.1,1.6)、(3.1,2.6)，期限2 -> 2.05%；与单券核心一致。
  assert.equal(r.status, '成功'); assert.ok(Math.abs(r.baseline - 2.05) < 1e-10); assert.ok(Math.abs(r.spread - 25) < 1e-10);
  assert.equal(r.sampleCount, 4); assert.equal(r.filter.city, '德州市'); assert.equal(ctx.settings.filter.city, '南京市');
  const direct = C.interpolate(C.curve(ctx.rows.filter(b => b.city === '德州市' && b.code !== '520363.SZ'), 0.5, 2), 2);
  assert.equal(r.baseline, direct);
});
test('缺OFR、日期不一致、拟合不足、区间外都保留条目且利差空白', () => {
  const item = parsed({ termYears: 2 });
  const missing = B.evaluate(context(), { ...item, ofrYield: null }); assert.equal(missing.spread, null); assert.match(missing.status, /缺少 OFR/);
  const wrongDate = context(); wrongDate.settings.quoteDate = '2026-10-20'; assert.match(B.evaluate(wrongDate, item).status, /日期|快照日/);
  const insufficient = context(); insufficient.settings.minBin = 5; assert.match(B.evaluate(insufficient, item).status, /节点不足/);
  const outside = context(); outside.rows[0].term = 9; assert.match(B.evaluate(outside, item).status, /不外推/);
  assert.match(B.evaluate(context(), { ...item, code: 'absent.IB' }).status, /未匹配/);
});
test('直辖市省份简称和带市后缀等价，真实城市差异仍提示', () => {
  const ctx = context(); ctx.rows = ctx.rows.map(b => ({ ...b, province: '天津', city: '天津市' }));
  const r = B.evaluate(ctx, parsed({ province: '天津市', city: '天津市' }));
  assert.ok(!r.warnings.some(x => x.includes('原文地域')));
  assert.ok(B.evaluate(ctx, parsed({ province: '天津市', city: '德州市' })).warnings.some(x => x.includes('原文地域')));
});
test('Excel保留数值、公式、缺失条目、原文和筛选条件，写出后可读回', () => {
  const ctx = context(), good = parsed({ ofrYield: 2.3, originalOfrYield: 1.87 });
  const results = [B.evaluate(ctx, good), B.evaluate(ctx, parsed({ id: 2, ofrYield: null }))];
  const wb = B.makeWorkbook(X, results, ctx), bytes = X.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const saved = X.read(bytes, { type: 'buffer', cellFormula: true });
  assert.deepEqual(Array.from(saved.SheetNames), ['利差结果', '计算口径', '输入原文']);
  assert.equal(saved.Sheets['利差结果'].I2.f, '(G2-H2)*100');
  assert.ok(Math.abs(saved.Sheets['利差结果'].I2.v - 25) < 1e-10);
  assert.equal(saved.Sheets['利差结果'].I3, undefined); assert.equal(saved.Sheets['利差结果'].M3.v, '缺少 OFR 收益率');
  assert.equal(saved.Sheets['输入原文'].D2.v, 1.87); assert.equal(saved.Sheets['利差结果'].G2.v, 2.3);
  assert.equal(saved.Sheets['利差结果']['!autofilter'].ref, 'A1:R3');
});
async function withServer(options, fn) {
  const server = createServer(options); await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const base = 'http://127.0.0.1:' + server.address().port;
  try { await fn(base); } finally { server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
}
test('服务从Markdown读取配置，但不向浏览器或提示词暴露密钥', async () => {
  const root = await fsp.mkdtemp(path.join(os.tmpdir(), 'bond-batch-test-'));
  const item = parsed(); let calls = 0;
  try {
    await fsp.writeFile(path.join(root, 'DeepSeek配置.md'), '```json\n{"apiKey":"unit-test-secret","model":"configured-model"}\n```');
    await fsp.writeFile(path.join(root, '批量查询解析规则.md'), '# 输出JSON规则');
    await withServer({ root, env: {}, fetchImpl: async (url, options) => {
      calls++; const payload = JSON.parse(options.body);
      assert.equal(url, 'https://api.deepseek.com/chat/completions'); assert.equal(options.headers.Authorization, 'Bearer unit-test-secret');
      assert.equal(payload.model, 'configured-model'); assert.equal(payload.messages[0].content, '# 输出JSON规则'); assert.ok(!options.body.includes('unit-test-secret'));
      return { ok: true, json: async () => ({ choices: [{ finish_reason: 'stop', message: { content: JSON.stringify({ items: [item] }) } }] }) };
    } }, async base => {
      const status = await (await fetch(base + '/api/status')).json(); assert.deepEqual(status, { configured: true, model: 'configured-model' });
      const response = await fetch(base + '/api/parse', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ text: item.rawText }) });
      assert.equal(response.status, 200); assert.equal((await response.json()).items[0].ofrYield, 1.87);
      assert.equal((await fetch(base + '/' + encodeURIComponent('DeepSeek配置.md'))).status, 404);
      assert.equal((await fetch(base + '/api/parse', { method: 'POST', headers: { Origin: 'https://other.example', 'Content-Type': 'application/json' }, body: '{}' })).status, 403);
      assert.equal(calls, 1);
    });
  } finally {
    for (const name of ['DeepSeek配置.md', '批量查询解析规则.md']) await fsp.unlink(path.join(root, name));
    await fsp.rmdir(root);
  }
});
test('上游密钥错误和截断输出返回明确错误，不产生半批结果', async () => {
  const item = parsed();
  for (const upstream of [{ ok: false, status: 401 }, { ok: true, json: async () => ({ choices: [{ finish_reason: 'length', message: { content: '{' } }] }) }]) {
    await withServer({ env: { DEEPSEEK_API_KEY: 'unit-test-key' }, fetchImpl: async () => upstream }, async base => {
      const res = await fetch(base + '/api/parse', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ text: item.rawText }) });
      assert.equal(res.status, 502); const result = await res.json(); assert.equal(result.items, undefined); assert.match(result.error, /密钥|完整/);
    });
  }
});
