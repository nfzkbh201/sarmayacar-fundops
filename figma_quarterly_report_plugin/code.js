figma.showUI(__html__, { width: 420, height: 460 });

const FIELD_LIMITS = {
  "company.description": 520,
  "notes.disclaimer": 240,
};

function sendStatus(text) {
  figma.ui.postMessage({ type: "status", text });
}

function isTextNode(node) {
  return node && node.type === "TEXT";
}

function collectNodesByName(root) {
  const map = new Map();

  function visit(node) {
    if (node.name) {
      map.set(node.name.trim(), node);
    }
    if ("children" in node) {
      for (const child of node.children) {
        visit(child);
      }
    }
  }

  visit(root);
  return map;
}

async function loadTextFont(node) {
  if (!isTextNode(node)) return;
  if (node.fontName === figma.mixed) {
    const firstFont = node.getRangeFontName(0, Math.max(1, node.characters.length));
    if (firstFont !== figma.mixed) {
      await figma.loadFontAsync(firstFont);
    }
    return;
  }
  await figma.loadFontAsync(node.fontName);
}

async function setText(namedNodes, fieldName, value, warnings) {
  const node = namedNodes.get(fieldName);
  const text = value == null ? "" : String(value);
  if (!node) {
    warnings.push(`Missing layer: ${fieldName}`);
    return;
  }
  if (!isTextNode(node)) {
    warnings.push(`Layer is not text: ${fieldName}`);
    return;
  }
  await loadTextFont(node);
  node.characters = text;
  const limit = FIELD_LIMITS[fieldName];
  if (limit && text.length > limit) {
    warnings.push(`${fieldName} may overflow (${text.length}/${limit} chars)`);
  }
}

async function fillCompanyFrame(frame, company, report, warnings) {
  const nodes = collectNodesByName(frame);
  await setText(nodes, "report.quarter", report.quarter || "", warnings);
  await setText(nodes, "company.name", company.name || "", warnings);
  await setText(nodes, "company.description", company.description || "", warnings);
  await setText(nodes, "notes.disclaimer", company.notes || "", warnings);

  const updates = company.key_updates || [];
  for (let index = 1; index <= 4; index += 1) {
    await setText(nodes, `company.update_${index}`, updates[index - 1] || "", warnings);
  }

  const metrics = company.metrics || [];
  for (let metricIndex = 1; metricIndex <= 7; metricIndex += 1) {
    const metric = metrics[metricIndex - 1] || {};
    await setText(nodes, `metric_${metricIndex}.label`, metric.label || "", warnings);
    const values = metric.values || [];
    for (let monthIndex = 1; monthIndex <= 3; monthIndex += 1) {
      const value = values[monthIndex - 1] || {};
      await setText(nodes, `metric_${metricIndex}.month_${monthIndex}`, value.value || "", warnings);
    }
  }

  const qoqRows = company.qoq_changes || [];
  for (let qoqIndex = 1; qoqIndex <= 6; qoqIndex += 1) {
    const row = qoqRows[qoqIndex - 1] || {};
    await setText(nodes, `qoq_${qoqIndex}.metric`, row.metric || "", warnings);
    await setText(nodes, `qoq_${qoqIndex}.percent`, row.qoq_percent || "", warnings);
    await setText(nodes, `qoq_${qoqIndex}.change`, row.change || "", warnings);
  }
}

function selectedTemplateFrame() {
  const selection = figma.currentPage.selection;
  if (selection.length !== 1) {
    throw new Error("Select exactly one company-page template frame before importing.");
  }
  const node = selection[0];
  if (node.type !== "FRAME" && node.type !== "COMPONENT" && node.type !== "INSTANCE") {
    throw new Error("Selected template must be a frame, component, or instance.");
  }
  return node;
}

async function importCompanies(data) {
  if (!data || data.schema_version !== "sarmayacar.quarterly_report.v1") {
    throw new Error("Unsupported or missing Sarmayacar quarterly report schema.");
  }
  const companies = data.companies || [];
  if (!companies.length) {
    throw new Error("No companies found in the data pack.");
  }

  const template = selectedTemplateFrame();
  const report = data.report || {};
  const spacing = template.width + 80;
  const warnings = [];
  const created = [];

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

  figma.currentPage.selection = created;
  figma.viewport.scrollAndZoomIntoView(created);
  return {
    createdCount: created.length,
    warnings,
  };
}

figma.ui.onmessage = async (message) => {
  if (message.type === "cancel") {
    figma.closePlugin();
    return;
  }

  if (message.type !== "import") return;

  try {
    sendStatus("Importing company pages...");
    const result = await importCompanies(message.data);
    const warningText = result.warnings.length
      ? `\n\nWarnings:\n- ${result.warnings.slice(0, 40).join("\n- ")}`
      : "";
    sendStatus(`Created ${result.createdCount} company pages.${warningText}`);
    figma.notify(`Created ${result.createdCount} company pages`);
  } catch (error) {
    const messageText = error && error.message ? error.message : String(error);
    sendStatus(messageText);
    figma.notify(messageText, { error: true });
  }
};
