(function () {
  'use strict';
  const app = window.CreditApp, S = window.CreditSnapshot, $ = id => document.getElementById(id);
  if (!app || !S) return;
  const local = location.protocol === 'file:';
  let revision = 'builtin', pending = null, sourceName = '', updating = false;
  const style = document.createElement('style');
  style.textContent = '.top{height:auto;min-height:76px;padding-top:12px;padding-bottom:12px}.top-right{flex-wrap:wrap;justify-content:flex-end}.dataset-strip{margin:10px 32px 0;font-size:11px;color:#64788c;line-height:1.8}.dataset-strip.error{color:var(--red)}.dataset-preview{display:grid;grid-template-columns:110px minmax(0,1fr);gap:8px;font-size:12px}.dataset-preview dd,.dataset-preview dt{margin:0;overflow-wrap:anywhere}.dataset-preview dt{color:var(--muted)}.dataset-modal-actions{display:flex;gap:10px;margin-top:18px;flex-wrap:wrap}#datasetModal .modal{width:650px}#datasetModal [role=alert]{white-space:pre-wrap;color:var(--red);font-size:12px}@media(max-width:760px){.top{padding-left:16px;padding-right:16px;flex-wrap:wrap}.top-right{gap:9px}.dataset-strip{margin:10px 16px 0}}';
  document.head.appendChild(style);
  const button = document.createElement('button'); button.className = 'btn'; button.id = 'htmlImportBtn'; button.textContent = '↥ 导入数据 HTML';
  $('importBtn').insertAdjacentElement('beforebegin', button);
  const input = document.createElement('input'); input.type = 'file'; input.accept = '.html,.htm'; input.hidden = true; input.id = 'htmlImportFile'; button.after(input);
  const strip = document.createElement('div'); strip.className = 'dataset-strip'; strip.id = 'datasetStatus'; strip.setAttribute('role', 'status'); strip.setAttribute('aria-live', 'polite'); document.querySelector('.subnav').after(strip);
  const modal = document.createElement('div'); modal.id = 'datasetModal'; modal.className = 'modal-backdrop hidden';
  modal.innerHTML = '<section class="modal" role="dialog" aria-modal="true" aria-labelledby="datasetTitle"><button class="modal-close" id="datasetCancel" aria-label="关闭导入预览">✕</button><h2 id="datasetTitle">确认导入 HTML 数据</h2><dl class="dataset-preview"><dt>选择的 HTML</dt><dd id="datasetSource"></dd><dt>数据来源</dt><dd id="datasetOrigin"></dd><dt>估值快照日</dt><dd id="datasetDate"></dd><dt>债券记录数</dt><dd id="datasetCount"></dd><dt>字段数</dt><dd id="datasetFields"></dd></dl><p id="datasetWarning"></p><p>只提取内嵌的债券数据，不替换当前页面，也不执行新 HTML 的脚本。原有批量查询、曲线和 Excel 导出功能保留。</p><p id="datasetError" role="alert"></p><div class="dataset-modal-actions"><button class="btn primary" id="datasetSave">保存并使用</button><button class="btn" id="datasetBack">取消</button></div></section>';
  document.body.appendChild(modal);
  function setStatus(text, error = false) { strip.textContent = text; strip.className = 'dataset-strip' + (error ? ' error' : ''); }
  function setBusy(value) { updating = value; button.disabled = value; $('datasetSave').disabled = value; $('datasetBack').disabled = value; $('datasetCancel').disabled = value; document.dispatchEvent(new CustomEvent('credit:data-busy', { detail: { busy: value } })); }
  function describe(saved) { return '当前使用已保存数据：' + saved.date + ' · ' + saved.count.toLocaleString('zh-CN') + ' 条 · 导入文件：' + saved.sourceName + '。刷新或重新启动后继续使用。'; }
  function close() { if (updating) return; pending = null; modal.classList.add('hidden'); button.focus(); }
  $('datasetCancel').onclick = $('datasetBack').onclick = close;
  modal.onclick = event => { if (event.target === modal) close(); };
  document.addEventListener('keydown', event => { if (event.key === 'Escape' && !modal.classList.contains('hidden')) close(); });
  button.onclick = () => input.click();
  input.onchange = async () => {
    const file = input.files[0]; if (!file) return;
    setBusy(true);
    try {
      if (file.size > 100 * 1024 * 1024) throw Error('HTML 文件超过 100MB，请检查文件。');
      setStatus('正在提取 HTML 内嵌数据…');
      const candidate = S.extractSnapshot(await file.text());
      pending = candidate; sourceName = file.name;
      $('datasetSource').textContent = file.name; $('datasetOrigin').textContent = candidate.meta.filename;
      $('datasetDate').textContent = candidate.meta.date; $('datasetCount').textContent = candidate.rows.length.toLocaleString('zh-CN'); $('datasetFields').textContent = candidate.meta.fields.length;
      const date = app.state().meta.date;
      $('datasetWarning').textContent = candidate.meta.date < date ? '注意：这份数据日期早于当前的 ' + date + '，确认后将切换到历史快照。' : '当前快照日为 ' + date + '。请确认新数据的估值日期符合需要。';
      $('datasetError').textContent = '';
      $('datasetSave').textContent = local ? '仅本次使用' : '保存并使用';
      if (local) $('datasetWarning').textContent += ' 当前为直接打开文件，无法写入本机服务；若需长期保存，请通过“启动批量查询.cmd”打开。';
      modal.classList.remove('hidden'); $('datasetSave').focus();
      setStatus('已提取数据，等待确认；当前查询数据尚未更改。');
    } catch (error) { pending = null; setStatus('导入失败：' + error.message + ' 当前数据未更改。', true); }
    finally { input.value = ''; setBusy(false); }
  };
  $('datasetSave').onclick = async () => {
    if (!pending || updating) return;
    setBusy(true); $('datasetError').textContent = '';
    try {
      let saved;
      if (!local) {
        const response = await fetch('/api/snapshot', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ snapshot: pending, sourceName, expectedRevision: revision }) });
        saved = await response.json(); if (!response.ok) throw Error(saved.error || '保存失败。');
        revision = saved.revision;
      }
      app.replaceDataset(pending, local ? 'temporary-html' : 'saved-html');
      modal.classList.add('hidden'); pending = null;
      setStatus(local ? '当前使用临时导入的 HTML 数据，刷新会恢复。需要长期保存时请通过 CMD 启动服务。' : describe(saved));
      button.focus();
    } catch (error) { $('datasetError').textContent = error.message + '；当前查询数据未更改。'; setStatus('未完成数据更新，请查看导入窗口中的原因。', true); }
    finally { setBusy(false); }
  };
  document.addEventListener('credit:dataset-changed', event => { if (event.detail.origin === 'excel') setStatus('当前使用临时导入的 Excel 数据；刷新后恢复已保存的 HTML 数据（没有保存记录时恢复预置数据）。'); });
  app.dataReady = (async () => {
    if (local) { setStatus('当前使用 HTML 预置数据。通过“启动批量查询.cmd”打开后，可导入并长期保存新的 HTML 数据。'); return; }
    setBusy(true); setStatus('正在读取已保存的数据…');
    try {
      const response = await fetch('/api/snapshot'), saved = await response.json();
      if (!response.ok) throw Error(saved.error || '读取快照失败。');
      revision = saved.revision;
      if (saved.snapshot) { const dataset = S.validateSnapshot(saved.snapshot); app.replaceDataset(dataset, 'saved-html'); setStatus(describe({ ...saved, date: dataset.meta.date, count: dataset.rows.length })); }
      else setStatus('当前使用预置快照 ' + app.state().meta.date + '。收到新 HTML 后，点击右上角“导入数据 HTML”即可更新并保存。');
    } catch (error) { setStatus('保存数据加载失败：' + error.message + ' 当前显示预置数据；请先处理此问题再进行批量查询。', true); throw error; }
    finally { setBusy(false); }
  })();
  app.dataReady.catch(() => {});
})();
