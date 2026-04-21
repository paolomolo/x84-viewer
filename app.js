const state = {
  rawXml: "",
  xmlDoc: null,
  projectInfo: {},
  bidInfo: {},
  positions: [],
  filteredPositions: [],
};

const el = {
  dropZone: document.getElementById("dropZone"),
  fileInput: document.getElementById("fileInput"),
  selectFileBtn: document.getElementById("selectFileBtn"),
  fileMeta: document.getElementById("fileMeta"),
  errorBox: document.getElementById("errorBox"),
  projectInfo: document.getElementById("projectInfo"),
  bidInfo: document.getElementById("bidInfo"),
  tableBody: document.querySelector("#positionsTable tbody"),
  tableMeta: document.getElementById("tableMeta"),
  xmlTree: document.getElementById("xmlTree"),
  rawXml: document.getElementById("rawXml"),
  exportBtn: document.getElementById("exportBtn"),
  searchInput: document.getElementById("searchInput"),
  ozMinInput: document.getElementById("ozMinInput"),
  ozMaxInput: document.getElementById("ozMaxInput"),
  resetFiltersBtn: document.getElementById("resetFiltersBtn"),
  themeToggle: document.getElementById("themeToggle"),
};

function init() {
  wireUpload();
  wireFilters();
  wireTheme();
  wireExport();
}

function wireUpload() {
  el.selectFileBtn.addEventListener("click", () => el.fileInput.click());
  el.fileInput.addEventListener("change", (event) => {
    const [file] = event.target.files || [];
    if (file) {
      handleFile(file);
    }
  });

  ["dragenter", "dragover"].forEach((type) => {
    el.dropZone.addEventListener(type, (event) => {
      event.preventDefault();
      el.dropZone.classList.add("drag-active");
    });
  });

  ["dragleave", "drop"].forEach((type) => {
    el.dropZone.addEventListener(type, (event) => {
      event.preventDefault();
      el.dropZone.classList.remove("drag-active");
    });
  });

  el.dropZone.addEventListener("drop", (event) => {
    const [file] = event.dataTransfer?.files || [];
    if (file) {
      handleFile(file);
    }
  });

  el.dropZone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      el.fileInput.click();
    }
  });
}

function wireFilters() {
  const trigger = () => applyFiltersAndRender();
  el.searchInput.addEventListener("input", trigger);
  el.ozMinInput.addEventListener("input", trigger);
  el.ozMaxInput.addEventListener("input", trigger);
  el.resetFiltersBtn.addEventListener("click", () => {
    el.searchInput.value = "";
    el.ozMinInput.value = "";
    el.ozMaxInput.value = "";
    applyFiltersAndRender();
  });
}

function wireTheme() {
  const saved = localStorage.getItem("gaeb-theme");
  if (saved === "dark") {
    document.body.classList.add("dark");
    el.themeToggle.textContent = "Light Mode";
  }

  el.themeToggle.addEventListener("click", () => {
    const nowDark = document.body.classList.toggle("dark");
    localStorage.setItem("gaeb-theme", nowDark ? "dark" : "light");
    el.themeToggle.textContent = nowDark ? "Light Mode" : "Dark Mode";
  });
}

function wireExport() {
  el.exportBtn.addEventListener("click", () => {
    if (!state.positions.length) {
      return;
    }

    const wb = XLSX.utils.book_new();
    const projectRows = Object.entries(state.projectInfo).map(([k, v]) => ({
      Feld: k,
      Wert: v,
    }));
    const bidRows = Object.entries(state.bidInfo).map(([k, v]) => ({
      Feld: k,
      Wert: v,
    }));
    const posRows = state.filteredPositions.map((p) => ({
      Nr: p.index,
      OZ: p.oz,
      Beschreibung: p.description,
      Menge: p.quantity,
      Einheit: p.unit,
      EP: p.unitPrice,
      GP: p.totalPrice,
      Waehrung: p.currency,
      Bereich: p.scope,
    }));

    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(projectRows), "Projekt");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(bidRows), "Angebot");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(posRows), "Positionen");

    const fileName = `gaeb-export-${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
  });
}

async function handleFile(file) {
  resetError();
  const fileName = file.name || "Unbekannt";
  if (!/\.(x84|xml)$/i.test(fileName)) {
    showError("Bitte eine .x84- oder .xml-Datei auswählen.");
    return;
  }

  try {
    const xmlText = await file.text();
    parseAndRender(xmlText, file);
  } catch (error) {
    showError(`Datei konnte nicht gelesen werden: ${error.message}`);
  }
}

function parseAndRender(xmlText, file) {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlText, "application/xml");
  const parserError = xmlDoc.querySelector("parsererror");
  if (parserError) {
    throw new Error("Ungueltiges XML-Format.");
  }

  const root = xmlDoc.documentElement;
  if (!root || localName(root).toLowerCase() !== "gaeb") {
    showError("Kein GAEB-Root-Element gefunden (<GAEB>). Datei pruefen.");
    return;
  }

  state.rawXml = xmlText;
  state.xmlDoc = xmlDoc;
  state.projectInfo = extractProjectInfo(xmlDoc);
  state.bidInfo = extractBidInfo(xmlDoc);
  state.positions = extractPositions(xmlDoc);

  el.fileMeta.textContent = `Geladen: ${file.name} (${formatBytes(file.size)})`;
  el.rawXml.textContent = xmlText.slice(0, 120000);
  if (xmlText.length > 120000) {
    el.rawXml.textContent += "\n\n... Vorschau gekuerzt ...";
  }

  renderInfoGrid(el.projectInfo, state.projectInfo, "Keine Projektdaten gefunden.");
  renderInfoGrid(el.bidInfo, state.bidInfo, "Keine Angebotsdaten gefunden.");
  renderTree(xmlDoc);
  applyFiltersAndRender();
  el.exportBtn.disabled = false;
}

function extractProjectInfo(xmlDoc) {
  const project = firstByLocalName(xmlDoc, "Project");
  const tender = firstByLocalName(xmlDoc, "Tender");
  return {
    Projektname: deepText(project, ["Name", "Description", "Title"]) || "-",
    Projekt-ID: deepText(project, ["ID", "ProjectNo", "Number"]) || "-",
    LV-Bezeichnung: deepText(tender, ["Name", "Description", "Title"]) || "-",
    Vergabeart: deepText(tender, ["Award", "Type", "Procedure"]) || "-",
    Quelle: xmlDoc.documentElement.getAttribute("xmlns") || "GAEB",
  };
}

function extractBidInfo(xmlDoc) {
  const bid = firstByLocalName(xmlDoc, "Bid");
  if (!bid) {
    return {
      Hinweis: "Kein <Bid>-Element gefunden.",
    };
  }
  return {
    Bieter: deepText(bid, ["Bidder", "Company", "Name"]) || "-",
    Angebotsnummer: deepText(bid, ["ID", "BidNo", "Number"]) || "-",
    Datum: deepText(bid, ["Date", "IssueDate", "Created"]) || "-",
    Waehrung:
      deepText(bid, ["Currency", "CurrencyCode"]) ||
      xmlDoc.documentElement.getAttribute("Cur") ||
      "EUR",
  };
}

function extractPositions(xmlDoc) {
  const all = Array.from(xmlDoc.getElementsByTagName("*"));
  const positionNodes = all.filter((node) => {
    const name = localName(node).toLowerCase();
    return name === "item" || name === "boqitem" || name === "position";
  });

  const positions = positionNodes.map((node, idx) => {
    const oz =
      deepText(node, ["RNoPart", "OZ", "OutlineText", "ItemNo", "Reference"]) ||
      node.getAttribute("RNoPart") ||
      node.getAttribute("ID") ||
      "-";
    const description =
      deepText(node, ["Description", "OutlineText", "Text", "ShortText", "LongText"]) ||
      collectTextSnippet(node, 220) ||
      "-";
    const quantityRaw =
      deepText(node, ["Qty", "Quantity", "QTakeoff"]) ||
      node.getAttribute("Qty") ||
      "";
    const unit =
      deepText(node, ["QU", "Unit", "UoM", "Measure"]) ||
      node.getAttribute("QU") ||
      "-";
    const unitPriceRaw =
      deepText(node, ["UP", "UnitPrice", "EP", "Price"]) ||
      node.getAttribute("UP") ||
      "";
    const totalPriceRaw =
      deepText(node, ["IT", "TotalPrice", "GP", "Amount"]) ||
      node.getAttribute("IT") ||
      "";
    const currency =
      deepText(node, ["Cur", "Currency"]) ||
      node.getAttribute("Cur") ||
      xmlDoc.documentElement.getAttribute("Cur") ||
      "EUR";

    const quantity = toNumber(quantityRaw);
    const unitPrice = toNumber(unitPriceRaw);
    const totalPrice = toNumber(totalPriceRaw) || quantity * unitPrice || 0;

    return {
      index: idx + 1,
      oz: oz.trim(),
      description: description.trim().replace(/\s+/g, " "),
      quantity,
      unit: unit.trim(),
      unitPrice,
      totalPrice,
      currency: currency.trim(),
      scope: inferScope(oz),
    };
  });

  return positions.filter((p) => p.oz !== "-" || p.description !== "-");
}

function applyFiltersAndRender() {
  const search = el.searchInput.value.trim().toLowerCase();
  const ozMin = normalizeOz(el.ozMinInput.value.trim());
  const ozMax = normalizeOz(el.ozMaxInput.value.trim());

  state.filteredPositions = state.positions.filter((p) => {
    const matchSearch =
      !search ||
      p.oz.toLowerCase().includes(search) ||
      p.description.toLowerCase().includes(search);
    const pOz = normalizeOz(p.oz);
    const matchMin = !ozMin || pOz >= ozMin;
    const matchMax = !ozMax || pOz <= ozMax;
    return matchSearch && matchMin && matchMax;
  });

  renderPositionsTable(state.filteredPositions, state.positions.length);
}

function renderPositionsTable(rows, totalCount) {
  el.tableBody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 7;
    td.className = "empty-state";
    td.textContent = totalCount
      ? "Keine Positionen fuer aktuelle Filter gefunden."
      : "Keine Positionen gefunden.";
    tr.appendChild(td);
    el.tableBody.appendChild(tr);
  } else {
    for (const row of rows) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(String(row.index))}</td>
        <td>${escapeHtml(row.oz)}</td>
        <td>${escapeHtml(row.description)}</td>
        <td>${formatNumber(row.quantity)}</td>
        <td>${escapeHtml(row.unit)}</td>
        <td>${formatCurrency(row.unitPrice, row.currency)}</td>
        <td>${formatCurrency(row.totalPrice, row.currency)}</td>
      `;
      el.tableBody.appendChild(tr);
    }
  }

  const shown = rows.length;
  el.tableMeta.textContent = `${shown} von ${totalCount} Positionen sichtbar`;
}

function renderInfoGrid(container, data, emptyText) {
  container.innerHTML = "";
  const entries = Object.entries(data || {});
  if (!entries.length) {
    container.classList.add("empty-state");
    container.textContent = emptyText;
    return;
  }

  container.classList.remove("empty-state");
  for (const [key, value] of entries) {
    const item = document.createElement("article");
    item.className = "kv-item";
    item.innerHTML = `
      <span class="key">${escapeHtml(key)}</span>
      <strong>${escapeHtml(String(value ?? "-"))}</strong>
    `;
    container.appendChild(item);
  }
}

function renderTree(xmlDoc) {
  const root = xmlDoc.documentElement;
  if (!root) {
    el.xmlTree.textContent = "Kein XML-Baum vorhanden.";
    return;
  }

  el.xmlTree.innerHTML = "";
  el.xmlTree.classList.remove("empty-state");
  const rootNode = createTreeNode(root, 0, 5);
  el.xmlTree.appendChild(rootNode);
}

function createTreeNode(node, depth, maxDepth) {
  const wrap = document.createElement("div");
  wrap.className = "tree-node";

  const attrs = Array.from(node.attributes || [])
    .map((a) => `${a.name}="${a.value}"`)
    .join(" ");
  const childElements = Array.from(node.children || []);
  const textContent = node.childElementCount ? "" : (node.textContent || "").trim();
  const textPreview = textContent ? `: ${textContent.slice(0, 80)}` : "";

  const label = document.createElement("div");
  label.innerHTML = `<span class="tree-label">&lt;${escapeHtml(localName(node))}&gt;</span>${
    attrs ? ` <span class="tree-attr">${escapeHtml(attrs)}</span>` : ""
  }${escapeHtml(textPreview)}`;
  wrap.appendChild(label);

  if (depth >= maxDepth) {
    if (childElements.length) {
      const info = document.createElement("div");
      info.className = "muted";
      info.textContent = `... ${childElements.length} weitere Knoten`;
      wrap.appendChild(info);
    }
    return wrap;
  }

  for (const child of childElements.slice(0, 30)) {
    wrap.appendChild(createTreeNode(child, depth + 1, maxDepth));
  }

  if (childElements.length > 30) {
    const info = document.createElement("div");
    info.className = "muted";
    info.textContent = `... ${childElements.length - 30} weitere Kinder`;
    wrap.appendChild(info);
  }

  return wrap;
}

function firstByLocalName(parent, name) {
  if (!parent) {
    return null;
  }
  return (
    Array.from(parent.getElementsByTagName("*")).find((n) => localName(n).toLowerCase() === name.toLowerCase()) ||
    null
  );
}

function deepText(parent, names) {
  if (!parent) {
    return "";
  }
  const lower = names.map((n) => n.toLowerCase());
  for (const node of Array.from(parent.getElementsByTagName("*"))) {
    if (lower.includes(localName(node).toLowerCase())) {
      const value = (node.textContent || "").trim();
      if (value) {
        return value;
      }
    }
  }
  return "";
}

function collectTextSnippet(node, maxLen) {
  const chunks = [];
  for (const child of Array.from(node.getElementsByTagName("*")).slice(0, 25)) {
    const text = (child.textContent || "").trim();
    if (text) {
      chunks.push(text);
    }
    if (chunks.join(" ").length > maxLen) {
      break;
    }
  }
  return chunks.join(" ").slice(0, maxLen);
}

function localName(node) {
  return node.localName || node.nodeName || "";
}

function toNumber(value) {
  if (!value) {
    return 0;
  }
  let raw = String(value).trim().replace(/\s+/g, "");
  raw = raw.replace(/[^\d.,-]/g, "");
  if (!raw) {
    return 0;
  }

  const lastComma = raw.lastIndexOf(",");
  const lastDot = raw.lastIndexOf(".");
  if (lastComma !== -1 && lastDot !== -1) {
    if (lastComma > lastDot) {
      raw = raw.replace(/\./g, "").replace(",", ".");
    } else {
      raw = raw.replace(/,/g, "");
    }
  } else if (lastComma !== -1) {
    raw = raw.replace(",", ".");
  }

  const num = Number(raw);
  return Number.isFinite(num) ? num : 0;
}

function formatNumber(value) {
  return new Intl.NumberFormat("de-DE", { maximumFractionDigits: 3 }).format(value || 0);
}

function formatCurrency(value, currency) {
  return new Intl.NumberFormat("de-DE", {
    style: "currency",
    currency: currency || "EUR",
  }).format(value || 0);
}

function normalizeOz(oz) {
  const digits = oz.replace(/[^\d]/g, "");
  return digits ? digits.padEnd(12, "0") : "";
}

function inferScope(oz) {
  if (!oz) {
    return "-";
  }
  const part = oz.split(".")[0].trim();
  return part || "-";
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes)) {
    return "-";
  }
  if (bytes < 1024) {
    return `${bytes} B`;
  }
  const units = ["KB", "MB", "GB"];
  let value = bytes / 1024;
  let unit = units[0];
  for (let i = 1; i < units.length && value >= 1024; i += 1) {
    value /= 1024;
    unit = units[i];
  }
  return `${value.toFixed(1)} ${unit}`;
}

function showError(message) {
  el.errorBox.classList.remove("hidden");
  el.errorBox.textContent = message;
}

function resetError() {
  el.errorBox.classList.add("hidden");
  el.errorBox.textContent = "";
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

window.addEventListener("DOMContentLoaded", init);
