const sessionPage = figma.currentPage;
const PLUGIN_BUILD = "2026.10.13.3";
let uiSession = "";
let activeReportRoot = null;
let busy = false;

function reportRoot() {
  return activeReportRoot || sessionPage;
}

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

function postUI(message) {
  figma.ui.postMessage({ ...message, uiSession, build: PLUGIN_BUILD });
}

function sendStatus(text) {
  postUI({ type: "status", text });
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
  append(parent, node);
  node.name = name;
  node.x = x;
  node.y = y;
  node.resize(w, h);
  node.fills = solid(fill);
  node.cornerRadius = radius;
  return node;
}

function line(parent, name, x, y, w, color = GREEN) {
  rect(parent, name, x, y, w, 1, color);
}

function estimateTextWidth(value, fontSize, isBold = false) {
  const textValue = String(value || "");
  return textValue.length * fontSize * (isBold ? 0.66 : 0.58);
}

function titleRuleStart(x, value, fontSize, isBold = false, fallback = 215) {
  return Math.min(1500, Math.max(fallback, x + estimateTextWidth(value, fontSize, isBold) + 34));
}

function reportPeriodLabel(report = {}) {
  const months = report.months || [];
  const lastMonth = months.length ? String(months[months.length - 1] || "") : "";
  const match = lastMonth.match(/^([A-Za-z]{3})\s+(\d{2,4})$/);
  if (!match) return "March 2026";
  const monthNames = {
    jan: "January",
    feb: "February",
    mar: "March",
    apr: "April",
    may: "May",
    jun: "June",
    jul: "July",
    aug: "August",
    sep: "September",
    oct: "October",
    nov: "November",
    dec: "December",
  };
  const monthName = monthNames[match[1].toLowerCase()] || match[1];
  const yearText = match[2].length === 2 ? `20${match[2]}` : match[2];
  return `${monthName} ${yearText}`;
}

function reportPeriodUpper(report = {}) {
  return reportPeriodLabel(report).toUpperCase();
}

function reportPeriodEndingText(report = {}) {
  const period = reportPeriodLabel(report);
  const parts = period.split(" ");
  if (parts.length !== 2) return "For the period ending March 31, 2026";
  const monthName = parts[0];
  const year = Number(parts[1]);
  const monthIndex = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ].indexOf(monthName);
  if (monthIndex < 0 || !year) return `For the period ending ${period}`;
  const endDay = new Date(year, monthIndex + 1, 0).getDate();
  return `For the period ending ${monthName} ${endDay}, ${year}`;
}

function cleanMetricLabel(label) {
  const textValue = String(label || "").trim();
  const replacements = new Map([
    ["Revenue (New lines/codes)", "Revenue (New)"],
    ["Revenue (Recurring lines/codes)", "Revenue (Recurring)"],
    ["Software Development Cost", "Software Dev."],
    ["Sales & Marketing", "Sales & Mktg."],
    ["Rent & CAM / Renovation Costs", "Rent / Renovation"],
    ["Admin & General Expenses", "Admin & General"],
    ["Total Operating Expenses", "Total OpEx"],
    ["Operational Rooms", "Op. Rooms"],
    ["Occupancy Rate (%)", "Occupancy"],
    ["Business Development", "Business Dev."],
    ["Professional Fees", "Prof. Fees"],
    ["Total No. of Transactions", "Total Transactions"],
    ["Disbursement Transactions", "Disbursements"],
  ]);
  return replacements.get(textValue) || textValue;
}

function financialDisplayRows(section) {
  const rows = section.rows || [];
  if (String(section.section || "").toLowerCase() === "income statement") {
    const segmentIndex = rows.findIndex((row) => String(row.label || "").trim().toLowerCase() === "segment");
    return segmentIndex >= 0 ? rows.slice(0, segmentIndex) : rows;
  }
  return rows;
}

function cleanFinancialLabel(label) {
  const textValue = String(label || "").trim();
  const replacements = new Map([
    ["Total amount of carried interest earned on realized investments", "Carry earned on realized investments"],
    ["Total amount of carried interest earned accrued on unrealized investments", "Carry accrued on unrealized investments"],
    ["Distribution to paid-in-capital (DPI)", "DPI"],
    ["Residual value to paid-in-capital (RVPI)", "RVPI"],
    ["Total Value to paid-in-capital (TVPI)", "TVPI"],
    ["Capital Contributions Receivables", "Capital contributions receivable"],
    ["Cash and Cash Equivalents", "Cash and cash equivalents"],
    ["Total remaining available drawdown", "Remaining available drawdown"],
    ["Total committed in portfolio companies", "Committed in portfolio companies"],
  ]);
  return replacements.get(textValue) || textValue;
}

function isFinancialSubheader(row) {
  const values = row.values || [];
  return Boolean(row.subsection) || !values.length;
}

function isFinancialTotal(label) {
  const normalised = String(label || "").trim().toLowerCase();
  return normalised.startsWith("total ") || normalised === "net income (loss)";
}

async function text(parent, name, value, x, y, w, h, size, options = {}) {
  const node = figma.createText();
  append(parent, node);
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
  return node;
}

function createFrame(name, x, y) {
  const frame = figma.createFrame();
  reportRoot().appendChild(frame);
  frame.name = name;
  frame.x = x;
  frame.y = y;
  frame.resize(PAGE_W, PAGE_H);
  frame.fills = solid(WHITE);
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

async function commonHeader(frame, section, pageNumber, report = {}) {
  rect(frame, "layout.top_rule", 0, 0, PAGE_W, 6, GREEN);
  await text(frame, "report.header", `Sarmayacar Ventures  |  ${section}  |  ${reportPeriodLabel(report)}`, 46, 45, 800, 28, 16, { color: { r: 0.58, g: 0.58, b: 0.58 } });
  logo(frame, 1622, 24, 1);
  line(frame, "layout.footer_rule", 46, 1048, 1608, LINE);
  await text(frame, "report.footer", "Strictly Confidential  |  Do Not Duplicate Or Distribute Without Permission", 48, 1072, 520, 22, 13, { color: GREEN });
  rect(frame, "report.page_chip", 1616, 1062, 38, 36, GREEN, 4);
  await text(frame, "report.page_number", String(pageNumber).padStart(2, "0"), 1616, 1072, 38, 16, 12, { bold: true, color: WHITE, align: "CENTER" });
}

async function createCover(x, y, report = {}) {
  const frame = createFrame("01 Cover", x, y);
  rect(frame, "layout.top_rule", 0, 0, PAGE_W, 6, GREEN);
  logo(frame, 817, 106, 1.25);
  await text(frame, "brand.wordmark", "SARMAYACAR", 705, 198, 320, 48, 36, { bold: true, color: { r: 0.13, g: 0.13, b: 0.13 }, align: "CENTER" });
  rect(frame, "cover.title_panel", 479, 428, 744, 328, GREEN, 88);
  await text(frame, "cover.title", "QUARTERLY\nREPORT", 590, 506, 520, 132, 58, { color: WHITE, align: "CENTER", lineHeight: 108 });
  line(frame, "cover.title_rule", 708, 650, 282, WHITE);
  await text(frame, "report.period", reportPeriodUpper(report), 738, 672, 255, 32, 23, { color: WHITE, align: "CENTER" });
  rect(frame, "layout.bottom_bar", 0, 1041, PAGE_W, 69, GREEN);
  await text(frame, "report.footer", "Strictly Confidential  |  Do Not Duplicate Or Distribute Without Permission", 635, 1072, 460, 18, 13, { color: WHITE, align: "CENTER" });
  return frame;
}

async function createContents(x, y, report = {}) {
  const frame = createFrame("02 Contents", x, y);
  rect(frame, "layout.left_rule", 0, 0, 6, PAGE_H, GREEN);
  await text(frame, "contents.legal", `Sarmayacar Ventures Cooperatief U.A.\n(the "Fund" or "SV" or collectively with\nrelated Fund Manager entities\n"Sarmayacar" or the "Firm")\n\nQuarterly Report\n${reportPeriodEndingText(report)}\n\n\nStrictly Confidential\nDo Not Duplicate Or Distribute Without Permission`, 72, 76, 450, 330, 16, { color: MID });
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

async function createHighlights(x, y, report = {}) {
  const frame = createFrame("03 Highlights", x, y);
  await commonHeader(frame, "Investor Report", "02", report);
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

// Approved report order and page breaks are independent of workbook row order.
const PORTFOLIO_ORDER = [
  "Trukker", "SimPaisa", "Abhi", "Bykea", "Oladoc", "TapMad", "Revolving Games",
  "OneLoad", "Jiye Technologies", "Roomy", "Procheck", "Patari", "Dot & Line",
  "Dawaai", "Jugnu", "Buyzilla",
];
const PORTFOLIO_ANGEL_ORDER = [
  "Taraki", "Savvy Technologies", "Delsys Technologies", "Startup Early (Holistiq)",
  "Scholar Den", "House Call", "Orko", "L.L.M.Bots", "Aabshar", "Raptr Games",
];
const PORTFOLIO_XS = [46, 174, 563, 704, 874, 1038, 1202, 1358, 1509];
const PORTFOLIO_WIDTHS = [124, 386, 136, 164, 158, 160, 153, 147, 141];
const PORTFOLIO_NOTE = "*Note: Carrying value of each portfolio company is based on the last completed financing round";

function portfolioIdentity(name) {
  const key = String(name || "").toLowerCase().replace(/\((inactive|holistiq)\)/g, "").replace(/[^a-z0-9]/g, "");
  const aliases = {
    abhifinance: "abhi", jiyetech: "jiyetechnologies", procheck1: "procheck",
    buyzilla2: "buyzilla", dotandline: "dotline", bolandi: "taraki",
    bolanditechnologies: "taraki", savvy: "savvytechnologies", holistiq: "startupearly",
    truckistan: "llmbots", raptr: "raptrgames",
  };
  return aliases[key] || key;
}

function splitPortfolioRows(summary = {}) {
  const rows = Array.isArray(summary.rows) ? summary.rows : [];
  const known = new Set([...PORTFOLIO_ORDER, ...PORTFOLIO_ANGEL_ORDER].map(portfolioIdentity));
  const byName = new Map();
  for (const row of rows) {
    const key = portfolioIdentity(row.name);
    if (!known.has(key)) throw new Error(`Portfolio layout has no position for "${row.name}". Update the approved company order before importing.`);
    if (byName.has(key)) throw new Error(`Duplicate portfolio company "${row.name}". Review the data pack before importing.`);
    byName.set(key, row);
  }
  // Reserve missing rows with blank data, never reuse another company's figures.
  const ordered = (names) => names.map((name) => byName.get(portfolioIdentity(name)) || { name });
  const portfolioRows = ordered(PORTFOLIO_ORDER);
  const operatorRows = ordered(PORTFOLIO_ANGEL_ORDER);
  return [
    { title: "04 Portfolio", pageNumber: "03", portfolioRows: portfolioRows.slice(0, 10), operatorRows: [], showNote: true },
    { title: "05 Portfolio Continued", pageNumber: "04", portfolioRows: portfolioRows.slice(10), operatorRows: operatorRows.slice(0, 3), totalKey: "portfolio_totals" },
    { title: "06 Portfolio Operator Angels", pageNumber: "05", portfolioRows: [], operatorRows: operatorRows.slice(3), totalKey: "totals" },
  ];
}

async function renderPortfolioHeaders(frame, y, prefix) {
  const headers = ["Company", "Description", "First Cash\nInjection", "Ownership", "SV Board\nSeats", "Cost ($)", "Carrying\nValue ($)", "Unrealized\nGain/Loss ($)", "Realized\nGain/Loss ($)"];
  const xs = PORTFOLIO_XS;
  const widths = PORTFOLIO_WIDTHS;
  for (let i = 0; i < headers.length; i += 1) {
    rect(frame, `${prefix}.header_${i + 1}`, xs[i], y, widths[i], 62, GREEN, 2);
    await text(frame, `${prefix}.header_${i + 1}.text`, headers[i], xs[i] + 8, y + 9, widths[i] - 16, 46, 16, { color: WHITE, align: i > 1 ? "CENTER" : "LEFT" });
  }
}

function portfolioCellValue(value) {
  return value == null || value === "" ? "-" : String(value);
}

async function renderPortfolioRows(frame, rows, startY, prefix, totals = null) {
  const xs = PORTFOLIO_XS;
  const widths = PORTFOLIO_WIDTHS;
  let y = startY;
  for (let r = 0; r < rows.length; r += 1) {
    const row = rows[r] || {};
    const bg = r % 2 === 0 ? rect(frame, `${prefix}.row_${r + 1}.bg`, 46, y, 1608, 68, { r: 0.96, g: 0.96, b: 0.96 }) : null;
    const company = await text(frame, `${prefix}.row_${r + 1}.company`, row.name || "", xs[0] + 8, y + 6, widths[0] - 16, 24, 14, { bold: true, color: MID, lineHeight: 135 });
    const description = await text(frame, `${prefix}.row_${r + 1}.description`, row.description || "", xs[1] + 12, y + 6, widths[1] - 22, 24, 14, { color: MID, lineHeight: 138 });
    company.textAutoResize = "HEIGHT";
    description.textAutoResize = "HEIGHT";
    const cells = [company, description];
    const fields = [
      ["first_cash", "first_cash_injection"], ["ownership", "ownership"], ["board", "sv_board_seats"],
      ["cost", "cost"], ["carrying", "carrying_value"], ["unrealized", "unrealized_gain_loss"], ["realized", "realized_gain_loss"],
    ];
    for (let i = 0; i < fields.length; i += 1) {
      const [suffix, field] = fields[i];
      const cell = await text(frame, `${prefix}.row_${r + 1}.${suffix}`, portfolioCellValue(row[field]), xs[i + 2], y + 6, widths[i + 2], 24, 14, { color: MID, align: "CENTER" });
      cell.textAutoResize = "HEIGHT";
      cells.push(cell);
    }
    const rowH = Math.max(68, Math.ceil(Math.max(...cells.map((cell) => cell.height))) + 12);
    if (bg) bg.resize(1608, rowH);
    y += rowH;
  }
  const afterRowsY = y;
  if (totals && rows.length) {
    line(frame, `${prefix}.total_rule`, 46, afterRowsY + 3, 1608, DARK);
    const bg = rect(frame, `${prefix}.total_bg`, 46, afterRowsY + 6, 1608, 30, LIGHT);
    await text(frame, `${prefix}.total_label`, "Total", 62, afterRowsY + 10, 140, 24, 16, { bold: true, color: MID });
    const fields = [["cost", "cost"], ["carrying", "carrying_value"], ["unrealized", "unrealized_gain_loss"], ["realized", "realized_gain_loss"]];
    let totalHeight = 30;
    for (let i = 0; i < fields.length; i += 1) {
      const [suffix, field] = fields[i];
      const cell = await text(frame, `${prefix}.total_${suffix}`, portfolioCellValue(totals[field]), xs[i + 5], afterRowsY + 10, widths[i + 5], 24, 15, { bold: true, color: MID, align: "CENTER" });
      cell.textAutoResize = "HEIGHT";
      totalHeight = Math.max(totalHeight, Math.ceil(cell.height) + 8);
    }
    bg.resize(1608, totalHeight);
    return afterRowsY + totalHeight + 14;
  }
  return afterRowsY + 12;
}

async function createPortfolioPage(x, y, report = {}, page = {}, summary = {}) {
  const frame = createFrame(page.title || "04 Portfolio", x, y);
  await commonHeader(frame, "Investor Report", page.pageNumber || "03", report);
  await text(frame, "portfolio.title", "PORTFOLIO", 46, 99, 200, 48, 34, { color: GREEN });
  line(frame, "portfolio.title_rule", 248, 115, 1406, GREEN);
  await renderPortfolioHeaders(frame, 158, "portfolio");
  let nextY = 220;
  const pageTotals = page.totalKey ? summary[page.totalKey] || {} : null;
  if (page.portfolioRows && page.portfolioRows.length) {
    nextY = await renderPortfolioRows(frame, page.portfolioRows, nextY, "portfolio.company", pageTotals);
  }
  if (page.operatorRows && page.operatorRows.length) {
    if ((page.portfolioRows || []).length) {
      await text(frame, "portfolio.operator_title", "OPERATOR ANGEL COMPANIES", 46, nextY + 4, 700, 46, 34, { color: GREEN });
      nextY += 60;
    }
    nextY = await renderPortfolioRows(frame, page.operatorRows, nextY, "portfolio.operator", pageTotals && !(page.portfolioRows || []).length ? pageTotals : null);
  }
  if (page.showNote) {
    line(frame, "portfolio.note_rule", 46, nextY, 1608, MID);
    const note = await text(frame, "portfolio.note", summary.note || PORTFOLIO_NOTE, 52, nextY + 10, 1590, 24, 15, { color: MID });
    note.textAutoResize = "HEIGHT";
    nextY += note.height + 20;
  }
  if (nextY > 1032) {
    throw new Error(`${frame.name}: portfolio content exceeds the footer boundary. Review long descriptions before importing.`);
  }
  return frame;
}

async function createPortfolioPages(x, y, report = {}, summary = {}) {
  const pages = splitPortfolioRows(summary);
  const frames = [];
  for (let i = 0; i < pages.length; i += 1) {
    const frame = await createPortfolioPage(x + i * (PAGE_W + 90), y, report, pages[i], summary);
    frames.push(frame);
  }
  return frames;
}

async function createCompanyTemplate(x, y, report = {}) {
  const frame = createFrame("Company Page Template", x, y);
  await commonHeader(frame, "Investor Report", "06", report);
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
  const startX = 345;
  const startY = 494;
  const monthW = 205;
  for (let m = 1; m <= 6; m += 1) {
    rect(frame, `metric_month_${m}.header_bg`, startX + (m - 1) * monthW, startY, monthW - 5, 42, GREEN);
    await text(frame, `metric_month_${m}.label`, `Month ${m}`, startX + (m - 1) * monthW, startY + 10, monthW - 5, 20, 16, { color: WHITE, align: "CENTER" });
  }
  for (let r = 1; r <= 10; r += 1) {
    const yRow = 550 + (r - 1) * 25;
    await text(frame, `metric_${r}.label`, r <= 7 ? `Metric ${r}` : "", 50, yRow, 280, 20, 13, { color: DARK });
    for (let m = 1; m <= 6; m += 1) {
      await text(frame, `metric_${r}.month_${m}`, "", startX + (m - 1) * monthW + 20, yRow, monthW - 32, 20, 13, { color: DARK, align: "RIGHT" });
    }
  }
  line(frame, "metrics.bottom_rule", 46, 967, 1610, { r: 0.18, g: 0.18, b: 0.18 });
  await text(frame, "source.company", "Source: Company Information", 48, 986, 400, 24, 14, { bold: true, color: { r: 0.56, g: 0.56, b: 0.56 } });
  await text(frame, "notes.disclaimer", "", 46, 1015, 1500, 24, 11, { bold: true, color: { r: 0.56, g: 0.56, b: 0.56 } });
  return frame;
}

async function createOperatorTemplate(x, y, report = {}) {
  const frame = createFrame("Operator Angel Template", x, y);
  await commonHeader(frame, "Investor Report", "21", report);
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

async function createFinancialExhibits(x, y, report = {}) {
  const frame = createFrame("Financial Exhibits Template", x, y);
  await commonHeader(frame, "Financial Exhibits", "26", report);
  await text(frame, "financial.title", "FINANCIAL EXHIBITS", 100, 111, 400, 42, 34, { color: GREEN });
  line(frame, "financial.title_rule", 460, 138, 1060, GREEN);
  rect(frame, "financial.summary_header", 88, 152, 1505, 52, GREEN, 10);
  await text(frame, "financial.summary_header_text", "Financial Summary", 112, 170, 400, 20, 13, { bold: true, color: WHITE });
  await text(frame, "financial.table_placeholder", "Financial exhibits will be mapped from fund finance inputs in Phase 3.", 112, 286, 1450, 80, 18, { color: MID });
  return frame;
}

function findFrameByName(name) {
  return reportRoot().findOne((node) => node.name === name && node.type === "FRAME");
}

function removeFramesByNames(names) {
  const wanted = new Set(names);
  const frames = reportRoot().findAll((node) => node.type === "FRAME" && wanted.has(node.name));
  for (const frame of frames) frame.remove();
}

async function updateReportPeriodText(frame, report, warnings) {
  if (!frame) return;
  const nodes = collectNodesByName(frame);
  if (nodes.has("report.period")) await setText(nodes, "report.period", reportPeriodUpper(report), warnings);
  if (nodes.has("report.header")) {
    const headerNode = nodes.get("report.header");
    const existing = isTextNode(headerNode) ? headerNode.characters : "";
    const section = String(existing).includes("Financial Exhibits") ? "Financial Exhibits" : "Investor Report";
    await setText(nodes, "report.header", `Sarmayacar Ventures  |  ${section}  |  ${reportPeriodLabel(report)}`, warnings);
  }
  if (nodes.has("contents.legal")) {
    await setText(nodes, "contents.legal", `Sarmayacar Ventures Cooperatief U.A.\n(the "Fund" or "SV" or collectively with\nrelated Fund Manager entities\n"Sarmayacar" or the "Firm")\n\nQuarterly Report\n${reportPeriodEndingText(report)}\n\n\nStrictly Confidential\nDo Not Duplicate Or Distribute Without Permission`, warnings);
  }
}

async function createTemplateFrames() {
  await ensureFonts();
  const gap = 90;
  const frames = [];
  frames.push(await createCover(0, 0));
  frames.push(await createContents(PAGE_W + gap, 0));
  frames.push(await createHighlights((PAGE_W + gap) * 2, 0));
  frames.push(...await createPortfolioPages(0, PAGE_H + gap));
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

function removeNodesByPrefixes(root, prefixes) {
  if (!root || !("findAll" in root)) return;
  const nodes = root.findAll((node) => prefixes.some((prefix) => String(node.name || "").startsWith(prefix)));
  for (const node of nodes) node.remove();
}

function metricSections(company) {
  return Array.isArray(company.metric_sections) ? company.metric_sections.filter((section) => Array.isArray(section.rows) && section.rows.length) : [];
}

function hasForecastValues(section) {
  return (section.rows || []).some((row) => (row.values || []).some((value) => value && value.forecast && value.forecast !== "-"));
}

async function drawValueText(frame, name, value, x, y, w, options = {}) {
  await text(frame, name, value || "-", x, y, w, options.h || 18, options.size || 12, {
    color: options.color || MID,
    align: "RIGHT",
    bold: Boolean(options.bold),
  });
}

async function drawMetricSection(frame, section, y, sectionIndex, report) {
  const rows = section.rows || [];
  const reportMonths = report.months || [];
  const maxMonths = Math.min(6, Math.max(reportMonths.length, rows[0] && rows[0].values ? rows[0].values.length : 0));
  if (!rows.length || !maxMonths) return y;

  const labelX = 50;
  const labelW = 285;
  const tableX = 344;
  const tableRight = 1604;
  const monthW = (tableRight - tableX) / maxMonths;
  const useForecast = hasForecastValues(section);
  const rowH = rows.length > 9 ? 21 : 24;
  const valueSize = rows.length > 9 || maxMonths >= 6 ? 11 : 12;

  if (section.title) {
    await text(frame, `metrics.section_${sectionIndex}.heading`, section.title, 48, y, 520, 28, 21, { color: DARK });
    y += 27;
  }
  if (section.units) {
    await text(frame, `metrics.section_${sectionIndex}.units`, section.units, 50, y, 210, 18, 12, { color: MID });
    y += 23;
  }

  if (section.show_header !== false) {
    const headerY = y;
    for (let m = 0; m < maxMonths; m += 1) {
      const monthLabel = (rows[0].values && rows[0].values[m] && rows[0].values[m].month) || reportMonths[m] || "";
      rect(frame, `metrics.section_${sectionIndex}.month_${m + 1}.bg`, tableX + m * monthW, headerY, monthW - 4, 42, GREEN);
      await text(frame, `metrics.section_${sectionIndex}.month_${m + 1}.label`, monthLabel, tableX + m * monthW, headerY + 11, monthW - 4, 18, 14, { color: WHITE, align: "CENTER", bold: true });
      if (useForecast) {
        const subW = (monthW - 4) / 2;
        await text(frame, `metrics.section_${sectionIndex}.month_${m + 1}.actual`, "Actual", tableX + m * monthW, headerY + 50, subW, 18, 12, { color: GREEN, align: "CENTER", bold: true });
        await text(frame, `metrics.section_${sectionIndex}.month_${m + 1}.forecast`, "Forecast", tableX + m * monthW + subW, headerY + 50, subW, 18, 12, { color: { r: 0.58, g: 0.58, b: 0.58 }, align: "CENTER" });
        line(frame, `metrics.section_${sectionIndex}.month_${m + 1}.divider`, tableX + m * monthW + subW, headerY + 42, 1, LINE);
      }
    }
    y += useForecast ? 78 : 54;
  } else {
    y += 3;
  }

  for (let r = 0; r < rows.length; r += 1) {
    const row = rows[r];
    const yRow = y + r * rowH;
    const isTotal = String(row.label || "").toLowerCase().includes("total") || String(row.label || "").toLowerCase().includes("net burn");
    await text(frame, `metrics.section_${sectionIndex}.row_${r + 1}.label`, cleanMetricLabel(row.label), labelX, yRow, labelW, rowH, valueSize, { color: DARK, bold: isTotal });
    const values = row.values || [];
    for (let m = 0; m < maxMonths; m += 1) {
      const value = values[m] || {};
      if (useForecast) {
        const subW = (monthW - 4) / 2;
        await drawValueText(frame, `metrics.section_${sectionIndex}.row_${r + 1}.month_${m + 1}.actual`, value.actual, tableX + m * monthW + 10, yRow, subW - 18, { size: valueSize, bold: isTotal, color: DARK });
        await drawValueText(frame, `metrics.section_${sectionIndex}.row_${r + 1}.month_${m + 1}.forecast`, value.forecast, tableX + m * monthW + subW + 10, yRow, subW - 18, { size: valueSize, bold: isTotal, color: { r: 0.62, g: 0.62, b: 0.62 } });
      } else {
        await drawValueText(frame, `metrics.section_${sectionIndex}.row_${r + 1}.month_${m + 1}`, value.actual || value.value, tableX + m * monthW + 14, yRow, monthW - 28, { size: valueSize, bold: isTotal, color: DARK });
      }
    }
  }

  y += rows.length * rowH + 18;
  line(frame, `metrics.section_${sectionIndex}.bottom_rule`, 46, y, 1608, sectionIndex === 1 ? LINE : { r: 0.32, g: 0.32, b: 0.32 });
  return y + 12;
}

async function drawCompanyMetricSections(frame, company, report) {
  const sections = metricSections(company);
  if (!sections.length) return false;
  removeNodesByPrefixes(frame, ["metric_", "metrics.", "source.company", "notes.disclaimer"]);
  let y = 439;
  let note = "";
  for (let i = 0; i < sections.length; i += 1) {
    y = await drawMetricSection(frame, sections[i], y, i + 1, report);
    if (sections[i].note) note = sections[i].note;
  }
  const sourceY = Math.min(984, Math.max(900, y + 6));
  await text(frame, "source.company", "Source: Company Information", 48, sourceY, 400, 24, 14, { bold: true, color: { r: 0.56, g: 0.56, b: 0.56 } });
  await text(frame, "notes.disclaimer", company.notes || note || "", 46, sourceY + 34, 1500, 30, 11, { bold: true, color: { r: 0.56, g: 0.56, b: 0.56 } });
  return true;
}

async function fillCompanyFrame(frame, company, report, warnings) {
  const nodes = collectNodesByName(frame);
  await setText(nodes, "company.name", (company.name || "").toUpperCase(), warnings);
  const titleRule = nodes.get("company.title_rule");
  if (titleRule && "resize" in titleRule) {
    const start = titleRuleStart(46, (company.name || "").toUpperCase(), 34, true, 215);
    titleRule.x = start;
    titleRule.resize(Math.max(80, 1654 - start), titleRule.height);
  }
  await setText(nodes, "company.description", company.description || "", warnings);

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
    await setText(nodes, `metric_${metricIndex}.label`, cleanMetricLabel(metric.label), warnings);
    const values = metric.values || [];
    for (let monthIndex = 1; monthIndex <= 6; monthIndex += 1) {
      const value = values[monthIndex - 1] || {};
      await setText(nodes, `metric_${metricIndex}.month_${monthIndex}`, value.value || "", warnings);
    }
  }
  const drewDynamicMetrics = await drawCompanyMetricSections(frame, company, report);
  if (!drewDynamicMetrics) {
    await setText(nodes, "notes.disclaimer", company.notes || "", warnings);
  }
}

function bulletText(items) {
  return (items || []).filter(Boolean).map((item) => `-  ${item}`).join("\n");
}

async function fillHighlights(data, warnings) {
  const frame = reportRoot().findOne((node) => node.name === "03 Highlights");
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

async function createOperatorAngelPages(operatorAngels, startX, startY, report = {}) {
  const created = [];
  const pageSize = 2;
  for (let page = 0; page < operatorAngels.length; page += pageSize) {
    const frame = createFrame(`Operator Angel Companies ${Math.floor(page / pageSize) + 1}`, startX + created.length * (PAGE_W + 90), startY);
    await commonHeader(frame, "Investor Report", String(21 + created.length).padStart(2, "0"), report);
    await text(frame, "operator.section_title", "OPERATOR ANGEL COMPANIES", 46, 99, 540, 42, 34, { color: GREEN });
    line(frame, "operator.section_rule", 542, 121, 1080, GREEN);
    const pair = operatorAngels.slice(page, page + pageSize);
    for (let i = 0; i < pair.length; i += 1) {
      const company = pair[i];
      const yBase = 185 + i * 365;
      await text(frame, `operator.company_${i + 1}.name`, company.name || "", 46, yBase, 420, 40, 32, { bold: true, color: GREEN });
      const ruleStart = titleRuleStart(46, company.name || "", 32, true, 422);
      line(frame, `operator.company_${i + 1}.rule`, ruleStart, yBase + 16, Math.max(80, 1622 - ruleStart), GREEN);
      await text(frame, `operator.company_${i + 1}.description_heading`, "Company Description", 46, yBase + 56, 300, 28, 21, { color: DARK });
      await text(frame, `operator.company_${i + 1}.description`, company.description || "", 46, yBase + 97, 1500, 70, 16, { color: MID, lineHeight: 170 });
      await text(frame, `operator.company_${i + 1}.updates_heading`, "Key Updates", 46, yBase + 176, 240, 28, 21, { color: DARK });
      await text(frame, `operator.company_${i + 1}.updates`, bulletText(company.key_updates), 46, yBase + 219, 1500, 110, 16, { color: MID, lineHeight: 170 });
    }
    created.push(frame);
  }
  return created;
}

async function createFinancialExhibitPages(financialExhibits, startX, startY, report = {}) {
  if (!financialExhibits || !financialExhibits.sections) return [];
  const created = [];
  const periods = financialExhibits.periods || [];
  for (const section of financialExhibits.sections.slice(0, 8)) {
    const frame = createFrame(`Financial Exhibits - ${section.section || "Section"}`, startX + created.length * (PAGE_W + 90), startY);
    await commonHeader(frame, "Financial Exhibits", String(26 + created.length).padStart(2, "0"), report);
    await text(frame, "financial.title", "FINANCIAL EXHIBITS", 100, 111, 420, 42, 34, { color: GREEN });
    line(frame, "financial.title_rule", 460, 138, 1060, GREEN);
    rect(frame, "financial.section_header", 88, 152, 1505, 52, GREEN, 10);
    await text(frame, "financial.section_header_text", section.section || "Financial Summary", 112, 170, 900, 20, 13, { bold: true, color: WHITE });
    await text(frame, "financial.units", "In USD $ 000's\nUnaudited figures", 112, 216, 260, 42, 14, { color: { r: 0.58, g: 0.58, b: 0.58 } });
    const labelX = 112;
    const y0 = 298;
    const labelW = 610;
    const maxCols = Math.min(periods.length, 7);
    const colW = 116;
    const colStart = labelX + labelW;
    for (let c = 0; c < maxCols; c += 1) {
      await text(frame, `financial.period_${c + 1}`, periods[c], colStart + c * colW, y0, colW - 8, 20, 13, { bold: true, color: MID, align: "RIGHT" });
    }
    const rows = financialDisplayRows(section);
    const rowLimit = String(section.section || "").toLowerCase() === "financial summary" ? 34 : 31;
    const rowH = rows.length > 28 ? 22 : 25;
    line(frame, "financial.header_rule", labelX, y0 + 31, 1480, LINE);
    for (let r = 0; r < Math.min(rows.length, rowLimit); r += 1) {
      const row = rows[r];
      const y = y0 + 48 + r * rowH;
      const subheader = isFinancialSubheader(row);
      const total = isFinancialTotal(row.label);
      await text(frame, `financial.row_${r + 1}.label`, cleanFinancialLabel(row.label), labelX, y, labelW - 18, 18, subheader ? 13 : 12, { bold: subheader || total, color: subheader ? GREEN : MID });
      const values = row.values || [];
      for (let c = 0; c < Math.min(values.length, maxCols); c += 1) {
        await text(frame, `financial.row_${r + 1}.value_${c + 1}`, values[c].value || "", colStart + c * colW, y, colW - 8, 18, 12, { bold: total, color: MID, align: "RIGHT" });
      }
      if (total) line(frame, `financial.row_${r + 1}.rule`, labelX, y + 22, 1480, LINE);
    }
    created.push(frame);
  }
  return created;
}

async function ensureReportFrontMatter(report, warnings, portfolioSummary = {}) {
  const frames = [];
  let cover = findFrameByName("01 Cover");
  if (!cover) cover = await createCover(0, 0, report);
  await updateReportPeriodText(cover, report, warnings);
  frames.push(cover);

  let contents = findFrameByName("02 Contents");
  if (!contents) contents = await createContents(PAGE_W + 90, 0, report);
  await updateReportPeriodText(contents, report, warnings);
  frames.push(contents);

  let highlights = findFrameByName("03 Highlights");
  if (!highlights) highlights = await createHighlights((PAGE_W + 90) * 2, 0, report);
  await updateReportPeriodText(highlights, report, warnings);
  frames.push(highlights);

  removeFramesByNames(["04 Portfolio Table", "04 Portfolio", "05 Portfolio Continued", "06 Portfolio Operator Angels"]);
  frames.push(...await createPortfolioPages(0, PAGE_H + 90, report, portfolioSummary));

  return frames;
}

async function findOrCreateCompanyTemplate(report = {}) {
  const existing = reportRoot().findOne((node) => node.name === "Company Page Template");
  if (existing && ["FRAME", "COMPONENT", "INSTANCE"].includes(existing.type)) {
    await updateReportPeriodText(existing, report, []);
    existing.visible = false;
    return existing;
  }
  const template = await createCompanyTemplate(PAGE_W + 90, PAGE_H + 90, report);
  template.visible = false;
  return template;
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
  const report = data.report || {};
  const warnings = [];
  const frontMatterFrames = await ensureReportFrontMatter(report, warnings, data.portfolio_summary || {});
  const template = await findOrCreateCompanyTemplate(report);
  const spacing = template.width + 90;
  const created = [...frontMatterFrames];
  const highlightFrame = await fillHighlights(data, warnings);
  if (highlightFrame && !created.includes(highlightFrame)) created.push(highlightFrame);

  for (let index = 0; index < companies.length; index += 1) {
    const company = companies[index];
    const clone = template.clone();
    clone.name = `${company.name || "Company"} - ${report.quarter || "Quarterly Report"}`;
    clone.visible = true;
    clone.x = template.x + spacing * (index + 1);
    clone.y = template.y;
    template.parent.appendChild(clone);
    await fillCompanyFrame(clone, company, report, warnings);
    created.push(clone);
  }
  const operatorStartX = template.x;
  const operatorStartY = template.y + PAGE_H + 90;
  created.push(...await createOperatorAngelPages(operatorAngels, operatorStartX, operatorStartY, report));
  created.push(...await createFinancialExhibitPages(data.financial_exhibits, operatorStartX, operatorStartY + PAGE_H + 90, report));

  template.remove();
  return { createdCount: created.length, warnings, frames: created };
}

async function updateReport(data, jobId = "") {
  if (figma.currentPage.id !== sessionPage.id) throw new Error("Return to the paired Figma page before updating.");
  const quarter = data && data.report && data.report.quarter;
  if (!/^Q[1-4] 20\d{2}$/.test(quarter || "")) throw new Error("The report needs a valid quarter and year.");
  for (const node of sessionPage.children) {
    if (node.getPluginData("sarmayacar_staging") === quarter) node.remove();
  }
  const owned = sessionPage.children.filter((node) => node.getPluginData("sarmayacar_report") === quarter);
  if (jobId && owned.length && owned.every((node) => node.getPluginData("sarmayacar_job") === jobId)
      && owned.length === Number(owned[0].getPluginData("sarmayacar_page_count"))) {
    const warnings = JSON.parse(owned[0].getPluginData("sarmayacar_warnings") || "[]");
    return { createdCount: owned.length, warnings };
  }
  const stage = figma.createFrame();
  stage.name = "Preparing quarterly report";
  stage.setPluginData("sarmayacar_staging", quarter);
  stage.visible = false;
  stage.clipsContent = false;
  stage.fills = [];
  sessionPage.appendChild(stage);
  const offsetX = owned.length ? Math.min(...owned.map((node) => node.x)) : 0;
  const other = sessionPage.children.filter((node) => node.id !== stage.id && !owned.includes(node));
  const offsetY = owned.length ? Math.min(...owned.map((node) => node.y))
    : other.reduce((bottom, node) => Math.max(bottom, node.y + node.height + 90), 0);
  let newFrames = [];
  let committed = false;
  try {
    activeReportRoot = stage;
    const result = await importCompanies(data);
    newFrames = result.frames;
    if (figma.currentPage.id !== sessionPage.id) throw new Error("Page changed during import. The previous report was kept.");
    for (let index = 0; index < newFrames.length; index += 1) {
      const frame = newFrames[index];
      frame.setPluginData("sarmayacar_report", quarter);
      frame.setPluginData("sarmayacar_job", jobId);
      frame.setPluginData("sarmayacar_page_count", String(newFrames.length));
      frame.setPluginData("sarmayacar_warnings", JSON.stringify(result.warnings.slice(0, 50)));
      sessionPage.appendChild(frame);
      frame.x = offsetX + (index % 3) * (PAGE_W + 90);
      frame.y = offsetY + Math.floor(index / 3) * (PAGE_H + 90);
    }
    committed = true;
    for (const old of owned) old.remove();
    sessionPage.selection = newFrames;
    try { figma.viewport.scrollAndZoomIntoView(newFrames); } catch (_) { /* The report is already committed. */ }
    return { createdCount: result.createdCount, warnings: result.warnings };
  } catch (error) {
    if (!committed) for (const frame of newFrames) if (!frame.removed) frame.remove();
    throw error;
  } finally {
    activeReportRoot = null;
    if (!stage.removed) stage.remove();
  }
}

async function connectionContext() {
  const connectionId = sessionPage.getPluginData("sarmayacar_connection");
  const saved = connectionId ? await figma.clientStorage.getAsync(`bridge.${connectionId}`) : null;
  postUI({ type: "context", context: {
    page_id: sessionPage.id, page_name: sessionPage.name, file_name: figma.root.name,
    target_id: sessionPage.getPluginData("sarmayacar_target"),
  }, saved });
}

figma.ui.onmessage = async (message) => {
  if (!message || typeof message !== "object") return;
  if (message.type === "ready" && !uiSession && /^[a-f0-9]{32}$/.test(message.uiSession || "")) {
    uiSession = message.uiSession;
  }
  if (!uiSession || message.uiSession !== uiSession) return;
  let acquired = false;
  if (message.type === "cancel") {
    figma.closePlugin();
    return;
  }

  try {
    if (message.type === "ready") {
      console.info(`[Sarmayacar ${PLUGIN_BUILD}] UI handshake received.`);
      await connectionContext();
      return;
    }
    if (message.type === "save-connection") {
      if (figma.currentPage.id !== sessionPage.id) throw new Error("Return to the page where you opened this plugin.");
      const id = String(message.connection.code || "").split(".")[0];
      if (!id || !message.connection.target_id) throw new Error("Invalid connection.");
      await figma.clientStorage.setAsync(`bridge.${id}`, message.connection);
      sessionPage.setPluginData("sarmayacar_connection", id);
      sessionPage.setPluginData("sarmayacar_target", message.connection.target_id);
      postUI({ type: "connection-saved" });
      return;
    }
    if (message.type === "forget-connection") {
      const id = sessionPage.getPluginData("sarmayacar_connection");
      if (id) await figma.clientStorage.deleteAsync(`bridge.${id}`);
      sessionPage.setPluginData("sarmayacar_connection", "");
      postUI({ type: "connection-forgotten" });
      return;
    }
    if (busy) throw new Error("A report is already importing. Please wait.");
    if (figma.currentPage.id !== sessionPage.id) throw new Error("Return to the page where you opened this plugin.");
    busy = true;
    acquired = true;
    if (message.type === "create-template") {
      sendStatus("Creating Sarmayacar XD-style template frames...");
      await createTemplateFrames();
      sendStatus("Created template frames: cover, contents, highlights, portfolio table, company page, operator angel page, and financial exhibits.");
      figma.notify("Created Sarmayacar quarterly template frames");
      postUI({ type: "template-result" });
      return;
    }

    if (message.type !== "import") return;
    sendStatus("Importing company pages...");
    const result = await updateReport(message.data, message.jobId || "");
    const warningText = result.warnings.length
      ? `\n\nWarnings:\n- ${result.warnings.slice(0, 50).join("\n- ")}`
      : "";
    sendStatus(`Updated ${result.createdCount} report pages.${warningText}`);
    postUI({ type: "import-result", jobId: message.jobId || "", ok: true, ...result });
    figma.notify(`Updated ${result.createdCount} report pages`);
  } catch (error) {
    const messageText = error && error.message ? error.message : String(error);
    sendStatus(messageText);
    if (["ready", "save-connection", "forget-connection"].includes(message.type)) {
      postUI({ type: "connection-error", error: messageText });
    } else {
      postUI({ type: "import-result", jobId: message.jobId || "", ok: false, error: messageText });
    }
    figma.notify(messageText, { error: true });
  } finally {
    if (acquired) busy = false;
  }
};

figma.on("currentpagechange", () => {
  postUI({ type: "page-active", active: figma.currentPage.id === sessionPage.id });
});

figma.showUI(__html__, { width: 460, height: 660 });
console.info(`[Sarmayacar ${PLUGIN_BUILD}] Plugin opened.`);
postUI({ type: "plugin-boot" });
