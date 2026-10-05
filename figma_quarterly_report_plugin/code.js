figma.showUI(__html__, { width: 440, height: 520 });

const PAGE_W = 1723;
const PAGE_H = 1147;
const GREEN = { r: 0 / 255, g: 140 / 255, b: 120 / 255 };
const DARK = { r: 67 / 255, g: 67 / 255, b: 67 / 255 };
const MID = { r: 103 / 255, g: 103 / 255, b: 103 / 255 };
const LIGHT = { r: 243 / 255, g: 245 / 255, b: 247 / 255 };
const LINE = { r: 213 / 255, g: 218 / 255, b: 220 / 255 };
const WHITE = { r: 1, g: 1, b: 1 };

const FIELD_LIMITS = {
  "company.description": 520,
  "notes.disclaimer": 240,
};

function sendStatus(text) {
  figma.ui.postMessage({ type: "status", text });
}

function solid(color) {
  return [{ type: "SOLID", color }];
}

async function ensureFonts() {
  await figma.loadFontAsync({ family: "Inter", style: "Regular" });
  await figma.loadFontAsync({ family: "Inter", style: "Bold" });
}

function append(parent, node) {
  parent.appendChild(node);
  return node;
}

function rect(parent, name, x, y, w, h, fill, radius = 0) {
  const node = figma.createRectangle();
  node.name = name;
  node.x = x;
  node.y = y;
  node.resize(w, h);
  node.fills = solid(fill);
  node.cornerRadius = radius;
  append(parent, node);
  return node;
}

function line(parent, name, x, y, w, color = GREEN) {
  rect(parent, name, x, y, w, 1, color);
}

async function text(parent, name, value, x, y, w, h, size, options = {}) {
  const node = figma.createText();
  node.name = name;
  node.x = x;
  node.y = y;
  node.resize(w, h);
  node.fontName = { family: "Inter", style: options.bold ? "Bold" : "Regular" };
  node.fontSize = size;
  node.lineHeight = { unit: "PERCENT", value: options.lineHeight || 135 };
  node.fills = solid(options.color || DARK);
  node.textAlignHorizontal = options.align || "LEFT";
  node.textAlignVertical = options.valign || "TOP";
  node.characters = value || "";
  append(parent, node);
  return node;
}

function createFrame(name, x, y) {
  const frame = figma.createFrame();
  frame.name = name;
  frame.x = x;
  frame.y = y;
  frame.resize(PAGE_W, PAGE_H);
  frame.fills = solid(WHITE);
  figma.currentPage.appendChild(frame);
  return frame;
}

function logo(parent, x, y, scale = 1) {
  const size = 22 * scale;
  const stroke = 7 * scale;
  const a = rect(parent, "brand.logo_block_a", x, y, size, size, WHITE);
  a.strokes = solid(GREEN);
  a.strokeWeight = stroke;
  const b = rect(parent, "brand.logo_block_b", x + 27 * scale, y + 27 * scale, size, size, WHITE);
  b.strokes = solid(GREEN);
  b.strokeWeight = stroke;
}

async function commonHeader(frame, section, pageNumber) {
  rect(frame, "layout.top_rule", 0, 0, PAGE_W, 6, GREEN);
  await text(frame, "report.header", `Sarmayacar Ventures  |  ${section}  |  March 2026`, 46, 45, 800, 28, 16, { color: { r: 0.58, g: 0.58, b: 0.58 } });
  logo(frame, 1622, 24, 1);
  line(frame, "layout.footer_rule", 46, 1048, 1608, LINE);
  await text(frame, "report.footer", "Strictly Confidential  |  Do Not Duplicate Or Distribute Without Permission", 48, 1072, 520, 22, 13, { color: GREEN });
  rect(frame, "report.page_chip", 1616, 1062, 38, 36, GREEN, 4);
  await text(frame, "report.page_number", String(pageNumber).padStart(2, "0"), 1616, 1072, 38, 16, 12, { bold: true, color: WHITE, align: "CENTER" });
}

async function createCover(x, y) {
  const frame = createFrame("01 Cover", x, y);
  rect(frame, "layout.top_rule", 0, 0, PAGE_W, 6, GREEN);
  logo(frame, 817, 106, 1.25);
  await text(frame, "brand.wordmark", "SARMAYACAR", 705, 198, 320, 48, 36, { bold: true, color: { r: 0.13, g: 0.13, b: 0.13 }, align: "CENTER" });
  rect(frame, "cover.title_panel", 479, 428, 744, 328, GREEN, 88);
  await text(frame, "cover.title", "QUARTERLY\nREPORT", 590, 506, 520, 132, 58, { color: WHITE, align: "CENTER", lineHeight: 108 });
  line(frame, "cover.title_rule", 708, 650, 282, WHITE);
  await text(frame, "report.period", "MARCH 2026", 775, 672, 180, 32, 23, { color: WHITE, align: "CENTER" });
  rect(frame, "layout.bottom_bar", 0, 1041, PAGE_W, 69, GREEN);
  await text(frame, "report.footer", "Strictly Confidential  |  Do Not Duplicate Or Distribute Without Permission", 635, 1072, 460, 18, 13, { color: WHITE, align: "CENTER" });
  return frame;
}

async function createContents(x, y) {
  const frame = createFrame("02 Contents", x, y);
  rect(frame, "layout.left_rule", 0, 0, 6, PAGE_H, GREEN);
  await text(frame, "contents.legal", "Sarmayacar Ventures Cooperatief U.A.\n(the \"Fund\" or \"SV\" or collectively with\nrelated Fund Manager entities\n\"Sarmayacar\" or the \"Firm\")\n\nMonthly Report\nFor the period ending March 31, 2026\n\n\nStrictly Confidential\nDo Not Duplicate Or Distribute Without Permission", 72, 76, 450, 330, 16, { color: MID });
  await text(frame, "contents.title", "CONTENTS", 1070, 106, 460, 80, 58, { bold: true, color: { r: 0.82, g: 0.82, b: 0.82 } });
  await text(frame, "contents.highlights_page", "02", 1078, 203, 70, 26, 22, { bold: true, color: GREEN });
  await text(frame, "contents.highlights", "HIGHLIGHTS", 1078, 244, 250, 34, 24, { color: DARK });
  line(frame, "contents.rule_1", 1078, 277, 424, GREEN);
  await text(frame, "contents.portfolio_page", "03", 1078, 297, 70, 26, 22, { bold: true, color: GREEN });
  await text(frame, "contents.portfolio", "PORTFOLIO", 1078, 338, 250, 34, 24, { color: DARK });
  await text(frame, "contents.company_list", "Simpaisa                         06\nAbhi Finance                  07\nBykea                              08\nOladoc                            08\nTapmad                           09\nRevolving Games          10\nOneLoad                         11\nJiye Technologies       12\nRoomy                            13\nProcheck                        14\nPatari                              15\nDot & Line                     16", 1080, 384, 420, 450, 16, { color: MID, lineHeight: 180 });
  line(frame, "contents.rule_2", 1078, 869, 424, GREEN);
  await text(frame, "contents.operator_page", "21", 1078, 888, 70, 26, 22, { bold: true, color: GREEN });
  await text(frame, "contents.operator", "OPERATOR ANGEL COMPANIES", 1078, 922, 400, 34, 24, { color: DARK });
  rect(frame, "report.page_chip", 1616, 1053, 38, 36, GREEN, 4);
  await text(frame, "report.page_number", "01", 1616, 1064, 38, 16, 12, { bold: true, color: WHITE, align: "CENTER" });
  return frame;
}

async function createHighlights(x, y) {
  const frame = createFrame("03 Highlights", x, y);
  await commonHeader(frame, "Investor Report", "02");
  await text(frame, "highlights.title", "HIGHLIGHTS", 46, 99, 220, 42, 34, { color: GREEN });
  line(frame, "highlights.title_rule", 272, 115, 1382, GREEN);
  await text(frame, "highlights.investment_title", "Investment Activity", 50, 178, 420, 28, 21, { color: DARK });
  await text(frame, "highlights.investment_bullets", "-  The fund has now made commitments of $16.9 million across 26 investments", 46, 215, 1500, 28, 16, { color: MID });
  await text(frame, "highlights.portfolio_title", "Portfolio Highlights", 50, 268, 420, 28, 21, { color: DARK });
  await text(frame, "highlights.portfolio_bullets", "-  Company highlights will be filled from the reviewed commentary sheet.\n-  AI can draft these bullets from QoQ movements, but reviewed text stays the source of truth.\n-  Keep bullets short so they fit the fixed report layout.", 46, 304, 1540, 260, 16, { color: MID, lineHeight: 170 });
  await text(frame, "highlights.other_title", "Other Firm Matters", 50, 664, 420, 28, 21, { color: DARK });
  await text(frame, "highlights.other_bullets", "-  Add firm-level updates in the Quarterly Commentary sheet.\n-  These should stay separate from company operating commentary.", 46, 700, 1540, 120, 16, { color: MID, lineHeight: 170 });
  return frame;
}

async function createPortfolioTable(x, y) {
  const frame = createFrame("04 Portfolio Table", x, y);
  await commonHeader(frame, "Investor Report", "03");
  await text(frame, "portfolio.title", "PORTFOLIO", 46, 99, 200, 42, 34, { color: GREEN });
  line(frame, "portfolio.title_rule", 248, 115, 1406, GREEN);
  const headers = ["Company", "Description", "First Cash\nInjection", "Ownership", "SV Board\nSeats", "Cost ($)", "Carrying\nValue ($)", "Unrealized\nGain/Loss ($)", "Realized\nGain/Loss ($)"];
  const xs = [46, 174, 563, 704, 874, 1038, 1202, 1358, 1509];
  const widths = [124, 386, 136, 164, 158, 160, 153, 147, 141];
  for (let i = 0; i < headers.length; i += 1) {
    rect(frame, `portfolio.header_${i + 1}`, xs[i], 158, widths[i], 62, GREEN, 2);
    await text(frame, `portfolio.header_${i + 1}.text`, headers[i], xs[i] + 8, 172, widths[i] - 16, 40, 16, { color: WHITE, align: i > 1 ? "CENTER" : "LEFT" });
  }
  await text(frame, "portfolio.rows", "Portfolio table rows will be mapped from a separate portfolio input file in Phase 3.", 52, 244, 1500, 60, 18, { color: MID });
  return frame;
}

async function createCompanyTemplate(x, y) {
  const frame = createFrame("Company Page Template", x, y);
  await commonHeader(frame, "Investor Report", "06");
  await text(frame, "company.name", "COMPANY NAME", 46, 99, 360, 42, 34, { bold: true, color: GREEN });
  line(frame, "company.title_rule", 215, 119, 1418, GREEN);
  await text(frame, "company.description_heading", "Company Description", 46, 157, 300, 28, 21, { color: DARK });
  await text(frame, "company.description", "Company description from Quarterly Commentary sheet.", 46, 198, 1580, 88, 16, { color: MID, lineHeight: 170 });
  await text(frame, "company.updates_heading", "Key Updates", 46, 299, 240, 28, 21, { color: DARK });
  for (let i = 1; i <= 4; i += 1) {
    rect(frame, `company.update_${i}.bullet`, 47, 333 + (i - 1) * 31, 9, 9, GREEN, 9);
    await text(frame, `company.update_${i}`, `Update ${i}`, 68, 326 + (i - 1) * 31, 1500, 28, 16, { color: MID });
  }
  await text(frame, "metrics.heading", "Key Financial & Operating Metrics", 48, 439, 400, 28, 21, { color: DARK });
  await text(frame, "metrics.units", "(Numbers in 000's)", 50, 466, 190, 20, 13, { color: MID });
  const startX = 214;
  const startY = 494;
  const monthW = 238;
  for (let m = 1; m <= 6; m += 1) {
    rect(frame, `metric_month_${m}.header_bg`, startX + (m - 1) * monthW, startY, monthW - 5, 42, GREEN);
    await text(frame, `metric_month_${m}.label`, `Month ${m}`, startX + (m - 1) * monthW, startY + 10, monthW - 5, 20, 16, { color: WHITE, align: "CENTER" });
  }
  for (let r = 1; r <= 10; r += 1) {
    const yRow = 550 + (r - 1) * 25;
    await text(frame, `metric_${r}.label`, r <= 7 ? `Metric ${r}` : "", 50, yRow, 160, 20, 14, { color: DARK });
    for (let m = 1; m <= 6; m += 1) {
      await text(frame, `metric_${r}.month_${m}`, "", startX + (m - 1) * monthW + 20, yRow, monthW - 42, 20, 14, { color: DARK, align: "RIGHT" });
    }
  }
  line(frame, "metrics.bottom_rule", 46, 967, 1610, { r: 0.18, g: 0.18, b: 0.18 });
  await text(frame, "source.company", "Source: Company Information", 48, 986, 400, 24, 14, { bold: true, color: { r: 0.56, g: 0.56, b: 0.56 } });
  await text(frame, "notes.disclaimer", "", 46, 1015, 1500, 24, 11, { bold: true, color: { r: 0.56, g: 0.56, b: 0.56 } });
  return frame;
}

async function createOperatorTemplate(x, y) {
  const frame = createFrame("Operator Angel Template", x, y);
  await commonHeader(frame, "Investor Report", "21");
  await text(frame, "operator.section_title", "OPERATOR ANGEL COMPANIES", 46, 99, 480, 42, 34, { color: GREEN });
  line(frame, "operator.section_rule", 542, 121, 1080, GREEN);
  await text(frame, "operator.company_1.name", "Company Name", 46, 185, 360, 40, 32, { bold: true, color: GREEN });
  line(frame, "operator.company_1.rule", 422, 201, 1200, GREEN);
  await text(frame, "operator.company_1.description_heading", "Company Description", 46, 241, 300, 28, 21, { color: DARK });
  await text(frame, "operator.company_1.description", "Operator angel description.", 46, 282, 1500, 70, 16, { color: MID, lineHeight: 170 });
  await text(frame, "operator.company_1.updates_heading", "Key Updates", 46, 361, 240, 28, 21, { color: DARK });
  await text(frame, "operator.company_1.updates", "-  Update bullets", 46, 404, 1500, 110, 16, { color: MID, lineHeight: 170 });
  return frame;
}

async function createFinancialExhibits(x, y) {
  const frame = createFrame("Financial Exhibits Template", x, y);
  await commonHeader(frame, "Financial Exhibits", "26");
  await text(frame, "financial.title", "FINANCIAL EXHIBITS", 100, 111, 400, 42, 34, { color: GREEN });
  line(frame, "financial.title_rule", 460, 138, 1060, GREEN);
  rect(frame, "financial.summary_header", 88, 152, 1505, 52, GREEN, 10);
  await text(frame, "financial.summary_header_text", "Financial Summary", 112, 170, 400, 20, 13, { bold: true, color: WHITE });
  await text(frame, "financial.table_placeholder", "Financial exhibits will be mapped from fund finance inputs in Phase 3.", 112, 286, 1450, 80, 18, { color: MID });
  return frame;
}

async function createTemplateFrames() {
  await ensureFonts();
  const gap = 90;
  const frames = [];
  frames.push(await createCover(0, 0));
  frames.push(await createContents(PAGE_W + gap, 0));
  frames.push(await createHighlights((PAGE_W + gap) * 2, 0));
  frames.push(await createPortfolioTable(0, PAGE_H + gap));
  frames.push(await createCompanyTemplate(PAGE_W + gap, PAGE_H + gap));
  frames.push(await createOperatorTemplate((PAGE_W + gap) * 2, PAGE_H + gap));
  frames.push(await createFinancialExhibits(0, (PAGE_H + gap) * 2));
  figma.currentPage.selection = frames;
  figma.viewport.scrollAndZoomIntoView(frames);
  return frames;
}

function isTextNode(node) {
  return node && node.type === "TEXT";
}

function collectNodesByName(root) {
  const map = new Map();
  function visit(node) {
    if (node.name) map.set(node.name.trim(), node);
    if ("children" in node) {
      for (const child of node.children) visit(child);
    }
  }
  visit(root);
  return map;
}

async function loadTextFont(node) {
  if (!isTextNode(node)) return;
  if (node.fontName === figma.mixed) return;
  await figma.loadFontAsync(node.fontName);
}

async function setText(namedNodes, fieldName, value, warnings) {
  const node = namedNodes.get(fieldName);
  const textValue = value == null ? "" : String(value);
  if (!node) {
    warnings.push(`Missing layer: ${fieldName}`);
    return;
  }
  if (!isTextNode(node)) {
    warnings.push(`Layer is not text: ${fieldName}`);
    return;
  }
  await loadTextFont(node);
  node.characters = textValue;
  const limit = FIELD_LIMITS[fieldName];
  if (limit && textValue.length > limit) {
    warnings.push(`${fieldName} may overflow (${textValue.length}/${limit} chars)`);
  }
}

async function fillCompanyFrame(frame, company, report, warnings) {
  const nodes = collectNodesByName(frame);
  await setText(nodes, "company.name", (company.name || "").toUpperCase(), warnings);
  await setText(nodes, "company.description", company.description || "", warnings);
  await setText(nodes, "notes.disclaimer", company.notes || "", warnings);

  const reportMonths = report.months || [];
  for (let index = 1; index <= 6; index += 1) {
    await setText(nodes, `metric_month_${index}.label`, reportMonths[index - 1] || "", warnings);
  }

  const updates = company.key_updates || [];
  for (let index = 1; index <= 4; index += 1) {
    await setText(nodes, `company.update_${index}`, updates[index - 1] || "", warnings);
  }

  const metrics = company.metrics || [];
  for (let metricIndex = 1; metricIndex <= 10; metricIndex += 1) {
    const metric = metrics[metricIndex - 1] || {};
    await setText(nodes, `metric_${metricIndex}.label`, metric.label || "", warnings);
    const values = metric.values || [];
    for (let monthIndex = 1; monthIndex <= 6; monthIndex += 1) {
      const value = values[monthIndex - 1] || {};
      await setText(nodes, `metric_${metricIndex}.month_${monthIndex}`, value.value || "", warnings);
    }
  }
}

function bulletText(items) {
  return (items || []).filter(Boolean).map((item) => `-  ${item}`).join("\n");
}

async function fillHighlights(data, warnings) {
  const frame = figma.currentPage.findOne((node) => node.name === "03 Highlights");
  if (!frame) return null;
  const nodes = collectNodesByName(frame);
  const highlights = data.highlights || [];
  const byName = new Map(highlights.map((item) => [String(item.name || "").toLowerCase(), item]));
  const investment = byName.get("investment activity") || {};
  const portfolio = byName.get("portfolio highlights") || {};
  const other = byName.get("other firm matters") || {};
  await setText(nodes, "highlights.investment_bullets", bulletText(investment.bullets), warnings);
  await setText(nodes, "highlights.portfolio_bullets", bulletText(portfolio.bullets), warnings);
  await setText(nodes, "highlights.other_bullets", bulletText(other.bullets), warnings);
  return frame;
}

async function createOperatorAngelPages(operatorAngels, startX, startY) {
  const created = [];
  const pageSize = 2;
  for (let page = 0; page < operatorAngels.length; page += pageSize) {
    const frame = createFrame(`Operator Angel Companies ${Math.floor(page / pageSize) + 1}`, startX + created.length * (PAGE_W + 90), startY);
    await commonHeader(frame, "Investor Report", String(21 + created.length).padStart(2, "0"));
    await text(frame, "operator.section_title", "OPERATOR ANGEL COMPANIES", 46, 99, 540, 42, 34, { color: GREEN });
    line(frame, "operator.section_rule", 542, 121, 1080, GREEN);
    const pair = operatorAngels.slice(page, page + pageSize);
    for (let i = 0; i < pair.length; i += 1) {
      const company = pair[i];
      const yBase = 185 + i * 365;
      await text(frame, `operator.company_${i + 1}.name`, company.name || "", 46, yBase, 420, 40, 32, { bold: true, color: GREEN });
      line(frame, `operator.company_${i + 1}.rule`, 422, yBase + 16, 1200, GREEN);
      await text(frame, `operator.company_${i + 1}.description_heading`, "Company Description", 46, yBase + 56, 300, 28, 21, { color: DARK });
      await text(frame, `operator.company_${i + 1}.description`, company.description || "", 46, yBase + 97, 1500, 70, 16, { color: MID, lineHeight: 170 });
      await text(frame, `operator.company_${i + 1}.updates_heading`, "Key Updates", 46, yBase + 176, 240, 28, 21, { color: DARK });
      await text(frame, `operator.company_${i + 1}.updates`, bulletText(company.key_updates), 46, yBase + 219, 1500, 110, 16, { color: MID, lineHeight: 170 });
    }
    created.push(frame);
  }
  return created;
}

async function createFinancialExhibitPages(financialExhibits, startX, startY) {
  if (!financialExhibits || !financialExhibits.sections) return [];
  const created = [];
  const periods = financialExhibits.periods || [];
  for (const section of financialExhibits.sections.slice(0, 8)) {
    const frame = createFrame(`Financial Exhibits - ${section.section || "Section"}`, startX + created.length * (PAGE_W + 90), startY);
    await commonHeader(frame, "Financial Exhibits", String(26 + created.length).padStart(2, "0"));
    await text(frame, "financial.title", "FINANCIAL EXHIBITS", 100, 111, 420, 42, 34, { color: GREEN });
    line(frame, "financial.title_rule", 460, 138, 1060, GREEN);
    rect(frame, "financial.section_header", 88, 152, 1505, 52, GREEN, 10);
    await text(frame, "financial.section_header_text", section.section || "Financial Summary", 112, 170, 900, 20, 13, { bold: true, color: WHITE });
    await text(frame, "financial.units", "In USD $ 000's\nUnaudited figures", 112, 216, 260, 42, 14, { color: { r: 0.58, g: 0.58, b: 0.58 } });
    const labelX = 112;
    const y0 = 286;
    const colW = 132;
    for (let c = 0; c < periods.length; c += 1) {
      await text(frame, `financial.period_${c + 1}`, periods[c], 660 + c * colW, y0, colW - 8, 20, 13, { bold: true, color: MID, align: "RIGHT" });
    }
    const rows = section.rows || [];
    for (let r = 0; r < Math.min(rows.length, 30); r += 1) {
      const row = rows[r];
      const y = y0 + 34 + r * 25;
      await text(frame, `financial.row_${r + 1}.label`, row.label || "", labelX, y, 480, 20, 13, { color: MID });
      const values = row.values || [];
      for (let c = 0; c < Math.min(values.length, periods.length); c += 1) {
        await text(frame, `financial.row_${r + 1}.value_${c + 1}`, values[c].value || "", 660 + c * colW, y, colW - 8, 20, 13, { color: MID, align: "RIGHT" });
      }
    }
    created.push(frame);
  }
  return created;
}

async function findOrCreateCompanyTemplate() {
  const selection = figma.currentPage.selection;
  if (selection.length === 1 && ["FRAME", "COMPONENT", "INSTANCE"].includes(selection[0].type)) {
    return selection[0];
  }
  const existing = figma.currentPage.findOne((node) => node.name === "Company Page Template");
  if (existing && ["FRAME", "COMPONENT", "INSTANCE"].includes(existing.type)) {
    return existing;
  }
  return await createCompanyTemplate(0, 0);
}

async function importCompanies(data) {
  if (!data || data.schema_version !== "sarmayacar.quarterly_report.v1") {
    throw new Error("Unsupported or missing Sarmayacar quarterly report schema.");
  }
  const companies = data.companies || [];
  const operatorAngels = data.operator_angels || [];
  if (!companies.length && !operatorAngels.length && !data.financial_exhibits) {
    throw new Error("No report data found in the data pack.");
  }

  await ensureFonts();
  const template = await findOrCreateCompanyTemplate();
  const report = data.report || {};
  const spacing = template.width + 90;
  const warnings = [];
  const created = [];
  const highlightFrame = await fillHighlights(data, warnings);
  if (highlightFrame) created.push(highlightFrame);

  for (let index = 0; index < companies.length; index += 1) {
    const company = companies[index];
    const clone = template.clone();
    clone.name = `${company.name || "Company"} - ${report.quarter || "Quarterly Report"}`;
    clone.x = template.x + spacing * (index + 1);
    clone.y = template.y;
    template.parent.appendChild(clone);
    await fillCompanyFrame(clone, company, report, warnings);
    created.push(clone);
  }
  const operatorStartX = template.x;
  const operatorStartY = template.y + PAGE_H + 90;
  created.push(...await createOperatorAngelPages(operatorAngels, operatorStartX, operatorStartY));
  created.push(...await createFinancialExhibitPages(data.financial_exhibits, operatorStartX, operatorStartY + PAGE_H + 90));

  figma.currentPage.selection = created;
  figma.viewport.scrollAndZoomIntoView(created);
  return { createdCount: created.length, warnings };
}

figma.ui.onmessage = async (message) => {
  if (message.type === "cancel") {
    figma.closePlugin();
    return;
  }

  try {
    if (message.type === "create-template") {
      sendStatus("Creating Sarmayacar XD-style template frames...");
      await createTemplateFrames();
      sendStatus("Created template frames: cover, contents, highlights, portfolio table, company page, operator angel page, and financial exhibits.");
      figma.notify("Created Sarmayacar quarterly template frames");
      return;
    }

    if (message.type !== "import") return;
    sendStatus("Importing company pages...");
    const result = await importCompanies(message.data);
    const warningText = result.warnings.length
      ? `\n\nWarnings:\n- ${result.warnings.slice(0, 50).join("\n- ")}`
      : "";
    sendStatus(`Created ${result.createdCount} company pages.${warningText}`);
    figma.notify(`Created ${result.createdCount} company pages`);
  } catch (error) {
    const messageText = error && error.message ? error.message : String(error);
    sendStatus(messageText);
    figma.notify(messageText, { error: true });
  }
};
