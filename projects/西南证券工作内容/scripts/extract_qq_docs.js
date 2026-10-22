// Read-only browser export. Document URL and sheet IDs are kept in ignored local configuration.
const puppeteer = require('puppeteer-core');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { parseArgs } = require('node:util');
const { values } = parseArgs({ options: {
  config: { type: 'string', default: path.join(__dirname, '..', 'config.local.json') },
  out: { type: 'string', default: path.join(os.tmpdir(), 'qq_all_sheets.json') },
  chrome: { type: 'string' }, help: { type: 'boolean', default: false }
}});
if (values.help) {
  console.log('node extract_qq_docs.js --config config.local.json --out snapshot.json [--chrome PATH]');
  process.exit(0);
}
if (!fs.existsSync(values.config)) throw new Error('缺少本地配置，请复制config.example.json为config.local.json并填写文档和sheet ID');
const config = JSON.parse(fs.readFileSync(values.config, 'utf8'));
if (!config.url || !Array.isArray(config.sheets) || !config.sheets.length || config.sheets.some(s => !s.id)) throw new Error('配置缺少url或sheet ID');
let activeBrowser;
const profileDir = fs.mkdtempSync(path.join(os.tmpdir(), 'swsc-qq-'));

const CHROME = values.chrome || config.chrome;
const BASE = config.url;
const OUT = path.resolve(values.out);
const SHEETS = config.sheets;

const sleep = (ms) => new Promise(r => setTimeout(r, ms));

(async () => {
  const browser = activeBrowser = await puppeteer.launch({
    executablePath: CHROME,
    headless: 'new',
    args: ['--no-sandbox', '--disable-gpu', '--disable-dev-shm-usage', '--disable-blink-features=AutomationControlled'],
    userDataDir: profileDir,
  });
  const page = await browser.newPage();
  await page.setViewport({ width: 1600, height: 900 });

  const result = { doc: '销售交易部持仓表', url: BASE, sheets: [] };

  for (const sheet of SHEETS) {
    const id = sheet.id;
    console.error('loading sheet', id, sheet.name);
    await page.goto(BASE + '?tab=' + id, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForFunction(
      `!!window.SpreadsheetApp && !!window.SpreadsheetApp.workbook && window.SpreadsheetApp.workbook.worksheetManager.activeSheetId === '${id}'`,
      { timeout: 60000 }
    );
    // 等 chunk 数据加载：轮询前 200 行里出现任意非空单元格
    await page.waitForFunction((id) => {
      try {
        const app = window.SpreadsheetApp;
        const wb = app.workbook;
        const s = wb.worksheetManager.getSheetBySheetId(id);
        const rc = Math.min(s.getRowCount(), 200);
        for (let r = 0; r < rc; r++) {
          for (let c = 0; c < 8; c++) {
            if (app.e2eTools.getCellEditValue(wb, r, c, id) !== '') return true;
          }
        }
        return false;
      } catch (e) { return false; }
    }, { timeout: 60000, polling: 1000 }, id);

    await sleep(2000);

    const sheetData = await page.evaluate((id) => {
      const app = window.SpreadsheetApp;
      const wb = app.workbook;
      const wm = wb.worksheetManager;
      const s = wm.getSheetBySheetId(id);
      const rc = s.getRowCount(), cc = s.getColCount();
      const grid = [];
      const formulas = [];   // 与 grid 同尺寸，公式原文（非公式格为 ""）
      let maxR = -1, maxC = -1;
      for (let r = 0; r < rc; r++) {
        const row = [];
        const frow = [];
        for (let c = 0; c < cc; c++) {
          const edit = app.e2eTools.getCellEditValue(wb, r, c, id);
          let v = edit, f = '';
          if (typeof edit === 'string' && edit.trim().startsWith('=')) {
            f = edit;
            // 取公式的计算结果（腾讯文档已算好），拿不到时保留公式原文
            try {
              const d = s.getCellDataAtPosition(r, c);
              const fr = d && d.getCellFormulaResult && d.getCellFormulaResult();
              const val = fr ? fr.value : null;
              if (val !== undefined && val !== null && typeof val !== 'object') v = val;
            } catch (e) { /* 保留公式原文 */ }
          }
          row.push(v);
          frow.push(f);
          if (v !== '' && v !== null) { if (r > maxR) maxR = r; if (c > maxC) maxC = c; }
        }
        grid.push(row);
        formulas.push(frow);
      }
      const trimmed = maxR < 0 ? [] : grid.slice(0, maxR + 1).map(row => row.slice(0, maxC + 1));
      const trimmedF = maxR < 0 ? [] : formulas.slice(0, maxR + 1).map(row => row.slice(0, maxC + 1));
      return { id, name: wm.getSheetNameBySheetId(id).trim(), rowCount: trimmed.length, colCount: trimmed[0] ? trimmed[0].length : 0, rows: trimmed, formulaRows: trimmedF };
    }, id);

    console.error('  done:', sheetData.rowCount, 'x', sheetData.colCount);
    result.sheets.push(sheetData);
  }

  if (result.sheets.some(s => !s.rowCount || !s.colCount)) throw new Error('工作表为空，拒绝输出不完整快照');
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, JSON.stringify(result), 'utf8');
  console.log('ALL_DONE bytes=', fs.statSync(OUT).size);
  await browser.close();
})().catch(e => { console.error('FATAL', e.message); process.exitCode = 1; })
  .finally(async () => {
    if (activeBrowser) await activeBrowser.close().catch(() => {});
    fs.rmSync(profileDir, { recursive: true, force: true });
  });
