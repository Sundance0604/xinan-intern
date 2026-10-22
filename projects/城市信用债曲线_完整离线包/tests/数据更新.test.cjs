'use strict';
const test = require('node:test'), assert = require('node:assert/strict');
const fs = require('node:fs/promises'), os = require('node:os'), path = require('node:path');
const S = require('../assets/快照数据.js');
const { createSnapshotStore } = require('../server/快照存储.cjs');
const { createServer } = require('../server/批量查询服务.cjs');
const { fixture, htmlOf } = require('./fixtures.cjs');
test('从HTML提取JSON，保留字段顺序、代码和null，绝不执行脚本', () => {
  const data = fixture(); data.rows[0][4] = null; data.rows[0][21] = '含大括号 } 与转义 " 以及 </script> 的文本';
  const parsed = S.extractSnapshot(htmlOf(data));
  assert.deepEqual(parsed, data); assert.equal(globalThis.SHOULD_NOT_RUN, undefined);
});
test('拒绝无数据、非JSON表达式、多份数据、非法日期、缺字段和坏行', () => {
  assert.throws(() => S.extractSnapshot('<script>alert(1)</script>'), /未找到/);
  assert.throws(() => S.extractSnapshot('<script>window.CREDIT_DATA={"rows":(()=>{throw 1})()};</script>'), /不可导入/);
  assert.throws(() => S.extractSnapshot(htmlOf(fixture()) + htmlOf(fixture())), /多份/);
  for (const modify of [d => { d.meta.date = '2026-02-30'; }, d => { d.meta.fields[0] = 'other'; }, d => d.rows[0].pop(), d => { d.rows[0][3] = '2年'; }]) {
    const data = fixture(); modify(data); assert.throws(() => S.validateSnapshot(data));
  }
});
test('原始HTML兼容；项目保留相同预置数据、曲线引擎和Excel组件', async t => {
  let original;
  try { original = await fs.readFile(path.resolve(__dirname, '../../../data/城市信用债曲线_完整离线包/双击打开.html'), 'utf8'); }
  catch (error) { if (error.code === 'ENOENT') { t.skip('独立分发包未包含仓库 data 原件，略过原件对比。'); return; } throw error; }
  const current = await fs.readFile(path.resolve(__dirname, '../双击打开.html'), 'utf8');
  const scripts = html => [...html.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
  assert.deepEqual(scripts(current).slice(0, 3), scripts(original).slice(0, 3));
  const result = S.extractSnapshot(original); assert.equal(result.rows.length, 38819); assert.equal(result.meta.date, '2026-09-15');
});
test('快照重启后可读，保存新版本会备份；无效/过期写入不破坏当前数据', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'bond-snapshot-'));
  try {
    const store = createSnapshotStore(root); assert.equal((await store.read()).revision, 'builtin');
    const saved = await store.save({ snapshot: fixture(), sourceName: '新数据.html', expectedRevision: 'builtin' });
    assert.equal((await createSnapshotStore(root).read()).snapshot.meta.date, '2026-10-22');
    await assert.rejects(store.save({ snapshot: fixture('2026-10-23'), expectedRevision: 'builtin' }), /其他页面/);
    await assert.rejects(store.save({ snapshot: { meta: {}, rows: [] }, expectedRevision: saved.revision }));
    assert.equal((await store.read()).revision, saved.revision);
    await store.save({ snapshot: fixture('2026-10-23'), sourceName: '下一版.html', expectedRevision: saved.revision });
    const previous = JSON.parse(await fs.readFile(path.join(root, 'runtime/上一版快照.json'), 'utf8'));
    assert.equal(previous.snapshot.meta.date, '2026-10-22');
    await fs.writeFile(path.join(root, 'runtime/当前快照.json'), 'corrupt');
    await assert.rejects(store.read(), /损坏/);
  } finally {
    for (const name of ['当前快照.json', '上一版快照.json']) await fs.unlink(path.join(root, 'runtime', name)).catch(() => {});
    await fs.rmdir(path.join(root, 'runtime')); await fs.rmdir(root);
  }
});
test('HTTP更新接口可持久保存，校验失败/跨站请求保留旧数据，重启仍可用', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'bond-snapshot-http-'));
  async function session(fn) {
    const server = createServer({ storageRoot: root }); await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
    try { await fn('http://127.0.0.1:' + server.address().port); }
    finally { server.closeAllConnections(); await new Promise(resolve => server.close(resolve)); }
  }
  try {
    await session(async base => {
      assert.equal((await (await fetch(base + '/api/snapshot')).json()).snapshot, null);
      const post = body => fetch(base + '/api/snapshot', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
      assert.equal((await post({ snapshot: fixture(), sourceName: '新数据.html', expectedRevision: 'builtin' })).status, 200);
      assert.equal((await post({ snapshot: fixture(), expectedRevision: 'builtin' })).status, 409);
      assert.equal((await post({ snapshot: {} })).status, 400);
      assert.equal((await fetch(base + '/api/snapshot', { method: 'POST', headers: { 'Content-Type': 'application/json', Origin: 'https://other.example' }, body: '{}' })).status, 403);
      assert.equal((await fetch(base + '/runtime/' + encodeURIComponent('当前快照.json'))).status, 404);
      assert.equal((await fetch(base + '/assets/' + encodeURIComponent('数据更新.js'))).status, 200);
    });
    await session(async base => { const value = await (await fetch(base + '/api/snapshot')).json(); assert.equal(value.snapshot.meta.date, '2026-10-22'); assert.equal(value.snapshot.rows.length, 5); });
  } finally { await fs.unlink(path.join(root, 'runtime/当前快照.json')); await fs.rmdir(path.join(root, 'runtime')); await fs.rmdir(root); }
});
