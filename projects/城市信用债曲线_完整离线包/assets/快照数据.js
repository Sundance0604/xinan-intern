(function (root) {
  'use strict';
  const required = ['code', 'name', 'issuer', 'term', 'yield', 'rating', 'implied', 'yy', 'method', 'province', 'city', 'type'];
  const numeric = new Set(['term', 'yield', 'row', 'amount', 'face']);
  const textFields = new Set(['code', 'name', 'issuer', 'rating', 'implied', 'yy', 'method', 'province', 'city', 'type', 'sheet', 'couponText', 'start', 'maturity', 'frequency', 'windLGFV', 'yyLGFV', 'fullName']);
  function validateSnapshot(value) {
    if (!value || !value.meta || !Array.isArray(value.rows)) throw Error('未找到 meta 和 rows 数据结构。');
    const m = value.meta, fields = m.fields;
    if (!Array.isArray(fields) || fields.length > 100 || !fields.every(k => typeof k === 'string' && k && !['__proto__', 'constructor', 'prototype'].includes(k)) || new Set(fields).size !== fields.length) throw Error('字段列表无效或存在重复字段。');
    const missing = required.filter(k => !fields.includes(k));
    if (missing.length) throw Error('数据缺少必要字段：' + missing.join('、'));
    if (!/^\d{4}-\d{2}-\d{2}$/.test(m.date || '') || !Number.isFinite(Date.parse(m.date + 'T00:00:00Z')) || new Date(m.date + 'T00:00:00Z').toISOString().slice(0, 10) !== m.date) throw Error('快照日期无效，须为 YYYY-MM-DD。');
    if (typeof m.filename !== 'string' || !m.filename.trim()) throw Error('缺少数据来源文件名。');
    if (!value.rows.length || value.rows.length > 250000) throw Error('记录数须为 1—250,000 条。');
    const codeIndex = fields.indexOf('code');
    for (let i = 0; i < value.rows.length; i++) {
      const row = value.rows[i];
      if (!Array.isArray(row) || row.length !== fields.length) throw Error('第 ' + (i + 1) + ' 条记录与字段数不一致。');
      if (typeof row[codeIndex] !== 'string' || !row[codeIndex].trim()) throw Error('第 ' + (i + 1) + ' 条缺债券代码。');
      for (let j = 0; j < fields.length; j++) {
        const cell = row[j], key = fields[j];
        if (numeric.has(key) && cell !== null && (typeof cell !== 'number' || !Number.isFinite(cell))) throw Error('第 ' + (i + 1) + ' 条 ' + key + ' 须为数值或 null。');
        if (textFields.has(key) && cell !== null && typeof cell !== 'string') throw Error('第 ' + (i + 1) + ' 条 ' + key + ' 须为文本或 null。');
        if (cell !== null && !['string', 'number', 'boolean'].includes(typeof cell)) throw Error('记录中含不支持的嵌套对象。');
      }
    }
    const sheets = Array.isArray(m.sheets) ? m.sheets : [];
    if (!sheets.every(s => s && typeof s.name === 'string' && Number.isInteger(s.count) && s.count >= 0)) throw Error('来源工作表信息无效。');
    return { meta: { ...m, fields: fields.slice(), sheets: sheets.map(s => ({ name: s.name, count: s.count })) }, rows: value.rows };
  }
  function jsonObjectAt(source, start) {
    let depth = 0, quoted = false, escape = false;
    for (let i = start; i < source.length; i++) {
      const c = source[i];
      if (quoted) { if (escape) escape = false; else if (c === '\\') escape = true; else if (c === '"') quoted = false; continue; }
      if (c === '"') quoted = true;
      else if (c === '{') depth++;
      else if (c === '}' && --depth === 0) return source.slice(start, i + 1);
    }
    throw Error('内嵌 JSON 不完整。');
  }
  function extractSnapshot(html) {
    if (typeof html !== 'string' || html.length > 100 * 1024 * 1024) throw Error('HTML 文件为空或超过 100MB。');
    const candidates = [];
    // 仅把 script 正文作为文本读取，绝不执行、eval、创建 iframe 或加载其资源。
    for (const match of html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script\s*>/gi)) {
      const script = match[1], assignment = /\bwindow\s*\.\s*CREDIT_DATA\s*=\s*(?=\{)/g;
      for (const marker of script.matchAll(assignment)) {
        try { candidates.push(validateSnapshot(JSON.parse(jsonObjectAt(script, marker.index + marker[0].length)))); }
        catch (error) { throw Error('CREDIT_DATA 数据不可导入：' + error.message); }
      }
    }
    if (!candidates.length) throw Error('未找到 window.CREDIT_DATA 的 JSON 数据。此 HTML 的数据格式暂不兼容。');
    if (candidates.length !== 1) throw Error('HTML 中有多份 CREDIT_DATA，无法确定应使用哪份。');
    return candidates[0];
  }
  const api = { validateSnapshot, extractSnapshot };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.CreditSnapshot = api;
})(typeof window === 'undefined' ? globalThis : window);
