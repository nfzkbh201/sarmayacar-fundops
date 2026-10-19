const fs = require('node:fs');
const path = require('node:path');
const { chromium } = require('playwright');

async function preview(dataPath, outputDir) {
  fs.mkdirSync(outputDir, { recursive: true });
  const browser = await chromium.launch({ channel: process.env.CHROME_CHANNEL || 'chrome', headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1763, height: 1187 }, deviceScaleFactor: 1 });
    await page.setContent('<html><head><style>*{box-sizing:border-box}body{margin:20px;background:#eee}.report{position:relative;width:1723px;height:1147px;background:white;margin-bottom:20px}.node{position:absolute;white-space:pre-wrap;overflow-wrap:anywhere}</style></head><body></body></html>');
    await page.evaluate(() => {
      let id = 0;
      const canvas = document.createElement('canvas').getContext('2d');
      class Node {
        constructor(type) { this.type = type; this.id = String(++id); this.name = ''; this.children = []; this.parent = null; this.x = this.y = 0; this._width = this._height = 100; this.visible = true; }
        appendChild(n) { if (n.parent) n.parent.children = n.parent.children.filter(c => c !== n); n.parent = this; this.children.push(n); }
        resize(w,h) { this.width = w; this._height = h; }
        get width() {
          if (this.autoResize !== 'WIDTH_AND_HEIGHT') return this._width;
          canvas.font = `${this.fontName?.style === 'Bold' ? 'bold' : 'normal'} ${this.fontSize}px "${this.fontName?.family || 'Inter'}"`;
          return canvas.measureText(this.characters || '').width + 1;
        }
        set width(v) { this._width = v; }
        get height() {
          if (!['HEIGHT', 'WIDTH_AND_HEIGHT'].includes(this.autoResize)) return this._height;
          canvas.font = `${this.fontName?.style === 'Bold' ? 'bold' : 'normal'} ${this.fontSize}px "${this.fontName?.family || 'Inter'}"`;
          let lines = 0;
          for (const paragraph of String(this.characters || '').split('\n')) {
            let used = ''; lines++;
            for (const word of paragraph.split(/\s+/)) {
              const candidate = used ? used + ' ' + word : word;
              if (used && canvas.measureText(candidate).width > this.width) { lines++; used = word; } else used = candidate;
            }
          }
          return lines * this.fontSize * (this.lineHeight?.value || 135) / 100;
        }
        set height(v) { this._height = v; }
        set textAutoResize(v) { this.autoResize = v; }
        get textAutoResize() { return this.autoResize; }
        findAll(p) { return this.children.flatMap(n => [...(p(n) ? [n] : []), ...n.findAll(p)]); }
        findOne(p) { return this.findAll(p)[0] || null; }
        remove() { if (this.parent) this.parent.children = this.parent.children.filter(n => n !== this); }
      }
      window.figma = {
        currentPage: new Node('PAGE'), root: { name: 'Preview' }, createFrame: () => new Node('FRAME'),
        createText: () => new Node('TEXT'), createRectangle: () => new Node('RECTANGLE'),
        loadFontAsync: async () => {}, ui: { postMessage() {} }, showUI() {}, on() {},
      };
      window.__html__ = '';
    });
    await page.addScriptTag({ content: fs.readFileSync(path.join(__dirname, '../figma_quarterly_report_plugin/code.js'), 'utf8') });
    const data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
    const result = await page.evaluate(async data => {
      await ensureFonts();
      const warnings = [], output = [];
      for (const company of data.companies) {
        const frame = await createCompanyTemplate(0, 0, data.report);
        await fillCompanyFrame(frame, company, data.report, warnings);
        frame.findOne(n => n.name === 'report.page_number').characters = String(data.companies.indexOf(company) + 6).padStart(2, '0');
        const container = document.createElement('section');
        container.className = 'report'; container.dataset.name = company.name;
        document.body.appendChild(container);
        for (const node of frame.children.filter(n => n.visible)) {
          const el = document.createElement('div'); el.className = 'node'; el.dataset.layer = node.name;
          Object.assign(el.style, { left: node.x + 'px', top: node.y + 'px', width: node.width + 'px', height: node.height + 'px' });
          const rgb = c => `rgb(${Math.round(c.r*255)},${Math.round(c.g*255)},${Math.round(c.b*255)})`;
          if (node.type === 'TEXT') {
            el.textContent = node.characters;
            Object.assign(el.style, { fontFamily: `"${node.fontName.family}",Arial,sans-serif`, fontSize: node.fontSize + 'px', fontWeight: node.fontName.style === 'Bold' ? 700 : 400, lineHeight: (node.lineHeight?.value || 135) + '%', color: rgb(node.fills[0].color), textAlign: node.textAlignHorizontal.toLowerCase() });
          } else {
            if (node.fills?.length) el.style.background = rgb(node.fills[0].color);
            if (node.strokes?.length) el.style.border = `${node.strokeWeight}px solid ${rgb(node.strokes[0].color)}`;
            if (node.cornerRadius) el.style.borderRadius = node.cornerRadius + 'px';
          }
          container.appendChild(el);
        }
        const overflow = [...container.querySelectorAll('.node')].filter(el => el.textContent && (el.scrollHeight > el.clientHeight + 1 || el.scrollWidth > el.clientWidth + 1)).map(el => ({ layer: el.dataset.layer, text: el.textContent, width: el.clientWidth, height: el.clientHeight, scrollHeight: el.scrollHeight, scrollWidth: el.scrollWidth }));
        output.push({ company: company.name, overflow });
      }
      return { output, warnings };
    }, data);
    for (const company of data.companies) {
      await page.locator(`.report[data-name="${company.name}"]`).screenshot({ path: path.join(outputDir, company.slug + '.png') });
    }
    fs.writeFileSync(path.join(outputDir, 'index.html'), await page.content());
    fs.writeFileSync(path.join(outputDir, 'layout-checks.json'), JSON.stringify(result, null, 2));
    console.log(JSON.stringify(result, null, 2));
    if (result.output.some(company => company.overflow.length)) process.exitCode = 1;
  } finally { await browser.close(); }
}

preview(process.argv[2], process.argv[3]).catch(error => { console.error(error); process.exitCode = 1; });
