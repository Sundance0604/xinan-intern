'use strict';
const fs = require('node:fs/promises');
const path = require('node:path');
const crypto = require('node:crypto');
const { validateSnapshot } = require('../assets/快照数据.js');
function createSnapshotStore(root) {
  const directory = path.join(root, 'runtime');
  const file = path.join(directory, '当前快照.json');
  async function read() {
    let raw;
    try { raw = await fs.readFile(file, 'utf8'); }
    catch (error) { if (error.code === 'ENOENT') return { snapshot: null, revision: 'builtin' }; throw error; }
    let envelope;
    try { envelope = JSON.parse(raw); envelope.snapshot = validateSnapshot(envelope.snapshot); }
    catch { const error = Error('已保存的快照文件损坏，请检查 runtime/当前快照.json；上一版备份位于同目录。'); error.status = 503; throw error; }
    return { ...envelope, revision: crypto.createHash('sha256').update(raw).digest('hex') };
  }
  async function atomicWrite(destination, content) {
    const temp = path.join(directory, '.snapshot-' + crypto.randomUUID() + '.tmp');
    try { await fs.writeFile(temp, content, { flag: 'wx' }); await fs.rename(temp, destination); }
    finally { await fs.unlink(temp).catch(error => { if (error.code !== 'ENOENT') throw error; }); }
  }
  async function save({ snapshot, sourceName, expectedRevision }) {
    const validated = validateSnapshot(snapshot), previous = await read();
    if (expectedRevision !== previous.revision) { const error = Error('其他页面已更新数据，请刷新后重新导入，避免覆盖新版本。'); error.status = 409; throw error; }
    await fs.mkdir(directory, { recursive: true });
    if (previous.snapshot) {
      const raw = await fs.readFile(file, 'utf8');
      await atomicWrite(path.join(directory, '上一版快照.json'), raw);
    }
    const envelope = { version: 1, snapshot: validated, sourceName: path.basename(String(sourceName || '导入的HTML')), savedAt: new Date().toISOString() };
    await atomicWrite(file, JSON.stringify(envelope));
    const saved = await read();
    return { revision: saved.revision, sourceName: saved.sourceName, savedAt: saved.savedAt, date: saved.snapshot.meta.date, count: saved.snapshot.rows.length };
  }
  return { read, save };
}
module.exports = { createSnapshotStore };
