/* Batch quotes: LLM extraction, local snapshot matching and existing curve calculations. */
(function (root) {
  'use strict';
  const compact = value => String(value ?? '').replace(/\s+/g, '');
  const finite = value => typeof value === 'number' && Number.isFinite(value);
  const normalizeCode = value => String(value || '').trim().toUpperCase();
  const normalizeProvince = value => /^(北京|天津|上海|重庆)市?$/.test(value || '') ? value.replace(/市$/, '') : value;
  const stringFields = ['name', 'code', 'termText', 'extensionText', 'issuerRating', 'bondRating', 'valuationType', 'method', 'province', 'city', 'impliedRating', 'yyRating'];
  const numberFields = ['termYears', 'bidYield', 'ofrYield', 'valuationYield'];
  function validateItems(parsed, original) {
    if (!parsed || !Array.isArray(parsed.items) || !parsed.items.length || parsed.items.length > 200) throw Error('每批须有 1—200 条记录。');
    const source = compact(original);
    let cursor = 0;
    const items = parsed.items.map((item, index) => {
      if (!item || typeof item !== 'object' || item.id !== index + 1) throw Error('记录编号或顺序不完整。');
      if (typeof item.rawText !== 'string' || !item.rawText.trim()) throw Error('第 ' + (index + 1) + ' 条缺原文。');
      const raw = compact(item.rawText);
      if (source.slice(cursor, cursor + raw.length) !== raw) throw Error('第 ' + (index + 1) + ' 条原文不一致或遗漏了记录。');
      cursor += raw.length;
      const clean = { id: index + 1, rawText: item.rawText };
      for (const field of stringFields) {
        if (item[field] !== null && typeof item[field] !== 'string') throw Error('第 ' + (index + 1) + ' 条 ' + field + ' 类型错误。');
        clean[field] = item[field] === null ? null : item[field].trim() || null;
      }
      for (const field of numberFields) {
        if (item[field] !== null && !finite(item[field])) throw Error('第 ' + (index + 1) + ' 条 ' + field + ' 必须是数值或 null。');
        clean[field] = item[field];
      }
      if (!clean.name && !clean.code) throw Error('第 ' + (index + 1) + ' 条缺债券名称和代码。');
      if (!Array.isArray(item.warnings) || item.warnings.some(v => typeof v !== 'string')) throw Error('warnings 格式错误。');
      clean.warnings = item.warnings.slice();
      for (const field of ['name', 'code', 'province', 'city']) {
        if (clean[field] && !raw.toUpperCase().includes(compact(clean[field]).toUpperCase())) throw Error('第 ' + (index + 1) + ' 条 ' + field + ' 不在原文中。');
      }
      if (clean.code && !/^[A-Z0-9]+\.(IB|SH|SZ|BJ)$/i.test(clean.code)) throw Error('债券代码格式错误。');
      clean.code = clean.code ? normalizeCode(clean.code) : null;
      if (clean.termYears !== null && clean.termYears <= 0) throw Error('剩余期限必须大于零。');
      if (clean.ofrYield !== null && (clean.ofrYield <= -100 || clean.ofrYield > 1000)) throw Error('OFR 超出支持的收益率范围。');
      if (clean.method && !['公募', '私募'].includes(clean.method)) throw Error('发行方式无效。');
      if (clean.valuationType && !['估值', '行权'].includes(clean.valuationType)) throw Error('估值类型无效。');
      if (clean.issuerRating === '0') clean.issuerRating = null;
      if (clean.bondRating === '0') clean.bondRating = null;
      if (clean.yyRating === '0') clean.yyRating = null;
      // 只审计已明确符合参考格式的报价对；自由文本仍由 LLM 理解，不用正则替代提取。
      if (clean.code) {
        const tail = item.rawText.toUpperCase().split(clean.code)[1] || '';
        const pair = /^\s+((?:--|[+-]?\d+(?:\.\d+)?))\s*\/\s*((?:--|[+-]?\d+(?:\.\d+)?))(?=\s|$)/.exec(tail);
        if (pair) {
          const expected = pair[2] === '--' ? null : Number(pair[2]);
          if (clean.ofrYield !== expected) {
            clean.ofrYield = null;
            clean.warnings.push('模型 OFR 与原文报价对不一致，已清空，请手工核对。');
          }
        }
      }
      if (clean.valuationType === '行权' || clean.extensionText) clean.warnings.push('原文含行权/延期期限，请核对快照期限与 OFR 的期限口径一致。');
      clean.warnings = [...new Set(clean.warnings)];
      return clean;
    });
    if (cursor !== source.length) throw Error('模型遗漏了部分输入文本。');
    return items;
  }
  function matchBond(rows, item) {
    // 代码存在时必须以代码为准，不能悄悄改成另一个同名市场代码。
    return item.code ? rows.filter(b => normalizeCode(b.code) === normalizeCode(item.code)) : rows.filter(b => compact(b.name) === compact(item.name));
  }
  function evaluate(context, item) {
    const { core: C, rows, meta, settings: s } = context;
    const matches = matchBond(rows, item);
    const result = { item, bond: matches.length === 1 ? matches[0] : null, matches, baseline: null, spread: null, sampleCount: 0, nodeCount: 0, nodes: [], filter: null, warnings: [...item.warnings], status: '' };
    const fail = text => { result.status = text; return result; };
    if (!matches.length) return fail('未匹配快照债券');
    if (matches.length > 1) return fail('匹配多条记录，须选择代码');
    const b = result.bond;
    if (item.name && b.name && compact(item.name) !== compact(b.name)) result.warnings.push('原文简称与快照不同，按代码匹配。');
    if ((item.city && item.city !== b.city) || (item.province && normalizeProvince(item.province) !== normalizeProvince(b.province))) result.warnings.push('原文地域与快照不一致；曲线采用快照地域，请核对。');
    if (finite(item.termYears) && Math.abs(item.termYears - b.term) > 0.02) result.warnings.push('原文期限 ' + item.termYears + ' 年与快照 ' + b.term + ' 年不同，计算采用快照期限。');
    if (s.quoteDate !== meta.date) return fail('报价日与估值快照日不一致');
    if (!b.city) return fail('快照缺城市');
    if (!finite(b.term) || b.term <= 0) return fail('快照缺有效期限');
    const f = result.filter = { ...s.filter, province: b.province || '', city: b.city };
    if (f.yyEnabled && (!Number.isInteger(f.yyMin) || !Number.isInteger(f.yyMax) || f.yyMin < 1 || f.yyMax > 10 || f.yyMin > f.yyMax)) return fail('YY 区间无效');
    if (!C.filter([b], { ...f, excludeSpecial: false }).length) result.warnings.push('标的不满足当前评级/发行等筛选，请核对可比性。');
    if (C.special(b)) result.warnings.push('标的含可识别特殊条款，注意期限及收益率口径。');
    const samples = C.filter(rows, f).filter(x => C.eligible(x) && x.term <= s.termMax && (!s.excludeTarget || x.code !== b.code));
    const nodes = result.nodes = C.curve(samples, s.binWidth, s.minBin);
    result.sampleCount = samples.length; result.nodeCount = nodes.length;
    if (nodes.length >= 2) result.baseline = C.interpolate(nodes, b.term);
    if (item.ofrYield === null) return fail('缺少 OFR 收益率');
    if (!finite(item.ofrYield) || item.ofrYield <= -100 || item.ofrYield > 1000) return fail('OFR 收益率无效');
    if (nodes.length < 2) return fail('拟合节点不足 2 个');
    if (result.baseline === null) return fail('标的期限超出拟合区间，不外推');
    if (nodes.some((n, i) => i && n.x - nodes[i - 1].x > Math.max(2, s.binWidth * 3))) result.warnings.push('部分期限节点间距较大，插值参考性有限。');
    result.spread = (item.ofrYield - result.baseline) * 100;
    return fail('成功');
  }
  function workbookData(results, context) {
    const { meta, settings: s } = context;
    const headers = ['序号', '债券简称', '债券代码', '省份', '城市', '快照期限(年)', '输入OFR(%)', '同期限拟合收益率(%)', '利差(BP)', '单券中债估值(%)', '有效样本数', '拟合节点数', '状态', '提示', '报价日期', '估值快照日', '原文期限', '原文中债类型'];
    const detail = [headers, ...results.map((r, i) => [i + 1, r.bond?.name || r.item.name || '', r.bond?.code || r.item.code || '', r.bond?.province || r.item.province || '', r.bond?.city || r.item.city || '', r.bond?.term ?? null, r.item.ofrYield, r.baseline, r.spread, r.bond?.yield ?? null, r.sampleCount, r.nodeCount, r.status, r.warnings.join('；'), s.quoteDate, meta.date, r.item.termText || '', r.item.valuationType || ''])];
    const f = s.filter;
    const notes = [
      ['项目', '内容'], ['报告', '债券批量 OFR 利差报告'], ['计算日期口径', '报价日与估值快照日必须一致；无实时行情获取'],
      ['数据来源', meta.filename], ['估值快照日', meta.date], ['报价日期', s.quoteDate],
      ['总条数', results.length], ['成功条数', results.filter(r => r.status === '成功').length], ['未完成条数', results.filter(r => r.status !== '成功').length],
      ['计算公式', '利差(BP) = (输入OFR(%) - 同期限拟合收益率(%)) × 100'],
      ['基准', '按分箱取期限及估值收益率中位数，再在标的快照期限处线性插值；不外推'],
      ['地域', '逐券使用数据库快照省份和城市；其他筛选条件共用当前页面设置'],
      ['分箱宽度(年)', s.binWidth], ['每箱最少样本', s.minBin], ['最大剩余期限(年)', s.termMax], ['拟合剔除当前标的', s.excludeTarget ? '是' : '否'],
      ['发行方式', f.method || '全部'], ['主体评级', f.rating || '全部'], ['中债隐含评级', f.implied || '全部'], ['YY区间', f.yyEnabled ? f.yyMin + '—' + f.yyMax + '（含±）' : '全部'], ['债券品种', (f.types || []).join('、')], ['城投范围', f.lgfv || '全部'], ['排除特殊条款', f.excludeSpecial ? '是' : '否'],
      ['来源边界', 'OFR来自用户文本/手工修正；期限、估值和曲线样本来自快照；LLM只提取文本，未参与利差计算'],
      ['未完成项目', '缺报价、未匹配、日期不一致或不可拟合时，利差留空，原因保留在状态列'], ['口径提示', '行权/延期期限等提示保留，须核对报价与快照期限口径。正利差仅表示相对曲线便宜。']
    ];
    const source = [['序号', '输入简称', '原始解析代码', '原始解析OFR(%)', '输入期限', '输入省份', '输入城市', '输入中债类型', '输入中债收益率(%)', '原文', '解析及修正提示'], ...results.map(r => [r.item.id, r.item.name, r.item.originalCode === undefined ? r.item.code : r.item.originalCode, r.item.originalOfrYield === undefined ? r.item.ofrYield : r.item.originalOfrYield, r.item.termText, r.item.province, r.item.city, r.item.valuationType, r.item.valuationYield, r.item.rawText, r.item.warnings.join('；')])];
    return { detail, notes, source };
  }
  function makeWorkbook(X, results, context) {
    const data = workbookData(results, context), workbook = X.utils.book_new();
    const sheet = X.utils.aoa_to_sheet(data.detail);
    sheet['!cols'] = [6, 23, 20, 15, 15, 16, 15, 25, 14, 22, 14, 14, 34, 95, 16, 16, 20, 18].map(wch => ({ wch }));
    sheet['!autofilter'] = { ref: 'A1:R' + data.detail.length };
    for (let row = 2; row <= data.detail.length; row++) {
      for (const col of ['F', 'G', 'H', 'J']) if (sheet[col + row]?.t === 'n') sheet[col + row].z = '0.0000';
      if (results[row - 2].status === '成功') sheet['I' + row] = { t: 'n', v: results[row - 2].spread, f: '(G' + row + '-H' + row + ')*100', z: '0.00' };
    }
    X.utils.book_append_sheet(workbook, sheet, '利差结果');
    const notes = X.utils.aoa_to_sheet(data.notes); notes['!cols'] = [{ wch: 28 }, { wch: 110 }];
    X.utils.book_append_sheet(workbook, notes, '计算口径');
    const source = X.utils.aoa_to_sheet(data.source); source['!cols'] = [6, 23, 20, 16, 20, 16, 16, 18, 24, 150, 95].map(wch => ({ wch })); source['!autofilter'] = { ref: 'A1:K' + data.source.length };
    X.utils.book_append_sheet(workbook, source, '输入原文');
    workbook.Workbook = { CalcPr: { calcMode: 'auto' } };
    return workbook;
  }
  const api = { validateItems, matchBond, evaluate, workbookData, makeWorkbook };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof document === 'undefined' || !root.CreditApp) return;
  root.CreditBatch = api;
  const app = root.CreditApp, $ = id => document.getElementById(id);
  const esc = value => String(value ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const fmt = (value, n = 4) => finite(value) ? value.toFixed(n) : '—';
  const style = document.createElement('style');
  style.textContent = '.batch-panel{max-width:1864px;margin:18px 28px 0;padding:18px 22px}.batch-heading,.batch-actions{display:flex;gap:12px;align-items:center;justify-content:space-between;flex-wrap:wrap}.batch-heading h2{font-size:16px;margin:0}.batch-heading p{color:var(--muted);font-size:11px;margin:7px 0 12px}.batch-panel textarea{min-height:94px;resize:vertical;line-height:1.8;font-size:12px}.batch-actions{justify-content:flex-start;margin-top:10px}.batch-actions label{display:flex;gap:7px;align-items:center;font-size:11px}.batch-actions input{width:145px;padding:7px}.batch-status{font-size:11px;line-height:1.8;color:var(--muted);margin:10px 0 0}.batch-error{color:var(--red)}.batch-tags{display:flex;flex-wrap:wrap;gap:8px;margin-top:13px}.batch-tag{border:1px solid #d7e3f0;background:#f5f8fc;border-radius:18px;color:#365879;padding:7px 12px;font-size:12px}.batch-tag small{font-size:10px;margin-left:7px;color:#75879b}.batch-tag.active{border-color:var(--blue);background:var(--blue-light);color:var(--blue);box-shadow:0 0 0 2px #2069d812}.batch-tag.pending{border-style:dashed}.batch-detail{margin-top:12px;background:#f6f8fb;padding:12px;border-radius:8px;font-size:11px;line-height:1.9;color:#53677c}.batch-detail strong{color:var(--ink)}.batch-detail .batch-metrics{display:flex;gap:20px;flex-wrap:wrap}.batch-detail details{margin-top:7px}.batch-detail pre{white-space:pre-wrap;font:inherit;margin:6px 0;overflow-wrap:anywhere}.batch-settings{margin-top:10px;font-size:11px;color:var(--muted)}.batch-settings>div{display:flex;align-items:center;gap:9px;margin-top:8px;flex-wrap:wrap}.batch-settings input{max-width:340px}.batch-match{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-top:8px}.batch-match select{max-width:440px}.batch-summary-table{max-height:290px;overflow:auto;margin-top:12px}.batch-summary-table th{position:sticky;top:0}.batch-summary-table button{background:none;border:0;color:var(--blue);font:inherit;padding:0}.batch-summary-table td{padding:9px}.batch-summary-table td:last-child{max-width:240px;white-space:normal;text-align:left}.batch-panel input[type=checkbox]{width:14px}.batch-summary-heading{display:flex;justify-content:space-between;align-items:center;margin-top:14px}.batch-summary-heading b{font-size:12px}@media(min-width:1920px){.batch-panel{margin-left:auto;margin-right:auto}}@media(max-width:760px){.batch-panel{margin:12px 16px 0;padding:15px}.batch-heading h2{font-size:15px}.batch-actions{gap:8px}.batch-tag{padding:7px 9px}}';
  document.head.appendChild(style);
  const panel = document.createElement('section');
  panel.className = 'card batch-panel'; panel.id = 'batchPanel';
  panel.innerHTML = '<div class="batch-heading"><div><h2>批量查询</h2><p>直接粘贴报价文本，DeepSeek 识别债券与 OFR；点击债券标签查看城市曲线。</p></div><a class="btn small" href="批量查询解析规则.md" target="_blank" rel="noopener">解析规则 ↗</a></div><label class="hidden" for="batchText">批量债券报价文本</label><textarea id="batchText" maxlength="60000" placeholder="在此粘贴多只债券的报价文本，例如：&#10;4.77Y AA+/-- 26德财02 520363.SZ --/1.87 --/1000 中债估值1.86 国企 私募 山东省德州市 中债隐含AA(2) YY评级6"></textarea><div class="batch-actions"><button id="batchParse" class="btn primary">识别并批量查询</button><button id="batchExport" class="btn" disabled>↓ 利差 Excel</button><button id="batchClear" class="btn small" disabled>清空结果</button><label for="batchQuoteDate">报价日<input id="batchQuoteDate" type="date"></label></div><details class="batch-settings" id="batchSettings"><summary>DeepSeek API 设置</summary><div><label for="batchKey">API Key（仅本次页面使用）</label><input id="batchKey" type="password" autocomplete="off" spellcheck="false" placeholder="填写你的 DeepSeek API Key"><button id="batchForgetKey" class="btn small">清除密钥</button><span id="batchModel"></span></div></details><p class="batch-status" id="batchStatus" role="status" aria-live="polite">解析时仅将粘贴文本发送至 DeepSeek；数据库留在本地。OFR 缺失不会用中债估值替代。</p><div class="batch-tags" id="batchTags" role="group" aria-label="批量债券标签"></div><div id="batchDetail" class="batch-detail hidden"></div><div id="batchSummary" class="hidden"><div class="batch-summary-heading"><b id="batchCount"></b><span class="micro">点击名称查看曲线；Excel 导出时按当前条件重新计算。</span></div><div class="batch-summary-table"><table><thead><tr><th>债券</th><th>城市</th><th>期限(年)</th><th>OFR(%)</th><th>拟合(%)</th><th>利差(BP)</th><th>状态</th></tr></thead><tbody id="batchRows"></tbody></table></div></div>';
  document.querySelector('.top').insertAdjacentElement('afterend', panel);
  let items = [], activeId = null, results = [], loading = false, controller = null, updatingData = false, dataReady = !app.dataReady;
  let activeBondCode = null, datasetGeneration = 0;
  const context = () => app.batchContext();
  $('batchQuoteDate').value = context().settings.quoteDate;
  function status(text, error = false) { $('batchStatus').textContent = text; $('batchStatus').className = 'batch-status' + (error ? ' batch-error' : ''); }
  function compute() { const ctx = context(); results = items.map(item => evaluate(ctx, item)); return ctx; }
  function render() {
    compute();
    $('batchTags').innerHTML = results.map(r => '<button class="batch-tag' + (r.item.id === activeId ? ' active' : '') + (r.status !== '成功' ? ' pending' : '') + '" data-batch-id="' + r.item.id + '" aria-pressed="' + (r.item.id === activeId) + '">' + esc(r.bond?.name || r.item.name || r.item.code) + '<small>' + (r.status === '成功' ? (r.spread >= 0 ? '+' : '') + fmt(r.spread, 2) + ' BP' : esc(r.status)) + '</small></button>').join('');
    $('batchRows').innerHTML = results.map(r => '<tr><td><button data-batch-id="' + r.item.id + '">' + esc(r.bond?.name || r.item.name || r.item.code) + '</button><span class="code">' + esc(r.bond?.code || r.item.code || '') + '</span></td><td>' + esc(r.bond?.city || r.item.city || '—') + '</td><td>' + fmt(r.bond?.term, 2) + '</td><td>' + fmt(r.item.ofrYield) + '</td><td>' + fmt(r.baseline) + '</td><td>' + fmt(r.spread, 2) + '</td><td>' + esc(r.status) + '</td></tr>').join('');
    $('batchSummary').classList.toggle('hidden', !items.length);
    $('batchParse').disabled = loading || updatingData || !dataReady;
    $('batchExport').disabled = !items.length || loading || updatingData || !dataReady;
    $('batchClear').disabled = !items.length || loading;
    $('batchCount').textContent = items.length + ' 只债券 · ' + results.filter(r => r.status === '成功').length + ' 只已计算';
    renderDetail();
  }
  function renderDetail() {
    const r = results.find(x => x.item.id === activeId);
    $('batchDetail').classList.toggle('hidden', !r);
    if (!r) return;
    const warning = r.warnings.length ? '<p style="margin:6px 0;color:var(--orange)">' + r.warnings.map(esc).join('；') + '</p>' : '';
    let match = '';
    if (!r.bond) {
      const candidates = r.matches.length ? r.matches : context().rows.filter(b => compact(b.name) === compact(r.item.name)).slice(0, 30);
      match = '<div class="batch-match"><span>选择本地债券：</span><select id="batchMatch"><option value="">请选择核对后的代码</option>' + candidates.map(b => '<option value="' + esc(b.code) + '">' + esc(b.name + ' · ' + b.code + ' · ' + b.city) + '</option>').join('') + '</select><button id="batchUseSearch" class="btn small">使用右侧已选债券</button></div><p style="margin:6px 0">未匹配的记录不会自动猜测。可在右侧搜索正确债券后点击“使用右侧已选债券”，或导入包含该债券的快照。</p>';
    }
    $('batchDetail').innerHTML = '<strong>' + esc(r.bond?.name || r.item.name || r.item.code) + '</strong> · ' + esc(r.status) + '<div class="batch-metrics"><span>OFR <b>' + fmt(r.item.ofrYield) + '%</b></span><span>同期限拟合 <b>' + fmt(r.baseline) + '%</b></span><span>利差 <b>' + fmt(r.spread, 2) + ' BP</b></span><span>' + r.sampleCount + ' 个样本 / ' + r.nodeCount + ' 个节点</span></div>' + warning + match + '<details><summary>查看解析原文与地域</summary><p>文本地域：' + esc([r.item.province, r.item.city].filter(Boolean).join(' / ') || '未明确') + '；曲线地域：' + esc(r.bond ? [r.bond.province, r.bond.city].filter(Boolean).join(' / ') : '未匹配') + '</p><pre>' + esc(r.item.rawText) + '</pre></details>';
    if ($('batchMatch')) $('batchMatch').onchange = event => { if (event.target.value) rematch(event.target.value); };
    if ($('batchUseSearch')) $('batchUseSearch').onclick = () => { const b = app.state().target; if (b) rematch(b.code); else status('请先在右侧查找并选择正确的债券。', true); };
  }
  function rematch(code) {
    const item = items.find(x => x.id === activeId), b = context().rows.find(x => x.code === code);
    if (!item || !b) return;
    item.warnings.push('用户手工匹配至 ' + b.name + ' ' + b.code + '。'); item.code = b.code;
    openItem(item.id, false);
  }
  function openItem(id, scroll = true) {
    activeId = id;
    const item = items.find(x => x.id === id); if (!item) return;
    const matches = matchBond(context().rows, item);
    activeBondCode = matches.length === 1 ? matches[0].code : null;
    if (matches.length === 1) app.loadBatchQuote(matches[0], item.ofrYield);
    else app.clearBatchQuote();
    render();
    if (scroll) document.querySelector('.chart-card').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  for (const id of ['batchTags', 'batchRows']) $(id).onclick = event => { const button = event.target.closest('[data-batch-id]'); if (button) openItem(Number(button.dataset.batchId)); };
  function syncFromPage(event) {
    if (event.target.id === 'quoteDate') $('batchQuoteDate').value = $('quoteDate').value;
    if (event.target.id === 'offerValue' && activeId && app.state().target?.code === activeBondCode && $('quoteMode').querySelector('button.active')?.dataset.value === 'yield') {
      const item = items.find(x => x.id === activeId);
      if (item) { const value = $('offerValue').value; const next = value === '' ? null : Number(value); if (next !== item.ofrYield) { item.ofrYield = next; if (!item.warnings.includes('用户在页面手工修正过 OFR，原始解析值见输入原文表。')) item.warnings.push('用户在页面手工修正过 OFR，原始解析值见输入原文表。'); } }
    }
    if (items.length) render();
  }
  document.querySelector('.main').addEventListener('change', syncFromPage);
  document.querySelector('.main').addEventListener('input', syncFromPage);
  document.querySelector('.main').addEventListener('click', () => { if (items.length) setTimeout(render, 0); });
  $('batchQuoteDate').onchange = () => { $('quoteDate').value = $('batchQuoteDate').value; $('quoteDate').dispatchEvent(new Event('input', { bubbles: true })); render(); };
  $('batchForgetKey').onclick = () => { $('batchKey').value = ''; status('已清除页面中的密钥。'); };
  $('batchText').addEventListener('input', () => { if (items.length) status('文本已修改；当前标签仍对应上一批结果，请点击“识别并批量查询”更新。'); });
  $('batchClear').onclick = () => { items = []; activeId = null; activeBondCode = null; render(); status('已清空批量结果。'); };
  $('batchParse').onclick = async () => {
    if (updatingData || !dataReady) return;
    const generation = datasetGeneration;
    const text = $('batchText').value.trim();
    if (!text) { status('请先粘贴债券报价文本。', true); $('batchText').focus(); return; }
    if (location.protocol === 'file:') { status('联网解析请先双击包内“启动批量查询.cmd”，在打开的本机网页中粘贴。原离线单券查询仍可直接使用。', true); return; }
    loading = true; controller = new AbortController(); $('batchParse').disabled = true; render();
    status('DeepSeek 正在解析文本并核对字段，请稍候…');
    const timeout = setTimeout(() => controller.abort(), 130000);
    try {
      const response = await fetch('/api/parse', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ text, apiKey: $('batchKey').value.trim() }), signal: controller.signal });
      const payload = await response.json();
      if (generation !== datasetGeneration) return;
      if (!response.ok) throw Error(payload.error || '解析失败，请重试。');
      items = validateItems(payload, text).map(item => ({ ...item, originalCode: item.code, originalOfrYield: item.ofrYield })); activeId = null; activeBondCode = null;
      render();
      status('已识别 ' + items.length + ' 只债券，' + items.filter(x => x.ofrYield === null).length + ' 只缺 OFR。共用当前筛选条件，逐券自动选择城市；点击标签查看，或导出利差 Excel。');
      if (items.length) openItem(items[0].id, false);
    } catch (error) { if (generation === datasetGeneration) status(error.name === 'AbortError' ? '解析超时，请减少条数后重试。' : error.message, true); }
    finally { clearTimeout(timeout); controller = null; loading = false; render(); }
  };
  function createWorkbook() {
    const ctx = compute(), workbook = makeWorkbook(root.XLSX, results, ctx);
    return { workbook, filename: '债券OFR利差报告_' + ctx.meta.date + '.xlsx' };
  }
  $('batchExport').onclick = () => {
    try { const { workbook, filename } = createWorkbook(); root.XLSX.writeFile(workbook, filename); render(); status('已导出 ' + items.length + ' 条记录；利差为数值及 Excel 公式，未完成条目留空并保留原因。'); }
    catch { status('Excel 导出失败，请重试。', true); }
  };
  if (location.protocol !== 'file:') fetch('/api/status').then(r => r.ok ? r.json() : Promise.reject()).then(info => { $('batchModel').textContent = '模型：' + info.model; if (info.configured) $('batchKey').placeholder = '本地服务已配置密钥，可在此临时覆盖'; }).catch(() => status('批量服务未连接，请通过“启动批量查询.cmd”打开页面。', true));
  new MutationObserver(() => { $('batchQuoteDate').value = $('quoteDate').value; if (items.length) { activeBondCode = null; openItem(activeId || items[0].id, false); } }).observe($('topDate'), { childList: true, characterData: true, subtree: true });
  document.addEventListener('credit:data-busy', event => { updatingData = event.detail.busy; render(); });
  document.addEventListener('credit:dataset-changed', () => { datasetGeneration++; if (controller) controller.abort(); items = []; activeId = null; activeBondCode = null; dataReady = true; render(); status('估值数据已更新。请核对报价日期，并重新识别粘贴文本，生成新快照的利差报告。'); });
  if (app.dataReady) { render(); app.dataReady.then(() => { dataReady = true; render(); }).catch(() => { dataReady = false; render(); status('已保存的数据未能加载，批量查询暂不可用；请查看页面数据状态。', true); }); }
  // 只读检查接口：不返回密钥，不改变页面或数据库。
  api.state = () => ({ items, results, activeId });
  api.createWorkbook = createWorkbook;
})(typeof window !== 'undefined' ? window : globalThis);
