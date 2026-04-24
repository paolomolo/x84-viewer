const state = {
  rawXml: "",
  xmlDoc: null,
  projectInfo: {},
  bidInfo: {},
  positions: [],
  filteredPositions: [],
  allFields: [],
  viewMode: "grouped",
  collapsedGroupIds: new Set(),
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
  allFieldsTableBody: document.querySelector("#allFieldsTable tbody"),
  allFieldsMeta: document.getElementById("allFieldsMeta"),
  xmlTree: document.getElementById("xmlTree"),
  rawXml: document.getElementById("rawXml"),
  exportBtn: document.getElementById("exportBtn"),
  searchInput: document.getElementById("searchInput"),
  ozMinInput: document.getElementById("ozMinInput"),
  ozMaxInput: document.getElementById("ozMaxInput"),
  resetFiltersBtn: document.getElementById("resetFiltersBtn"),
  groupedViewBtn: document.getElementById("groupedViewBtn"),
  listViewBtn: document.getElementById("listViewBtn"),
  allFieldsPanel: document.getElementById("allFieldsPanel"),
  expertToggle: document.getElementById("expertToggle"),
  themeToggle: document.getElementById("themeToggle"),
};

function init() {
  wireUpload();
  wireFilters();
  wireViewMode();
  wireExpertMode();
  wireTheme();
  wireExport();
}

function wireViewMode() {
  const saved = localStorage.getItem("gaeb-view-mode");
  setViewMode(saved === "list" ? "list" : "grouped");

  el.groupedViewBtn.addEventListener("click", () => setViewMode("grouped"));
  el.listViewBtn.addEventListener("click", () => setViewMode("list"));
}

function setViewMode(mode) {
  state.viewMode = mode === "list" ? "list" : "grouped";
  el.groupedViewBtn.classList.toggle("active", state.viewMode === "grouped");
  el.listViewBtn.classList.toggle("active", state.viewMode === "list");
  localStorage.setItem("gaeb-view-mode", state.viewMode);
  if (state.positions.length) {
    applyFiltersAndRender();
  }
}

function wireExpertMode() {
  const saved = localStorage.getItem("gaeb-expert-mode");
  setExpertMode(saved === "1");

  el.expertToggle.addEventListener("click", () => {
    const enabled = el.allFieldsPanel.classList.contains("hidden");
    setExpertMode(enabled);
  });
}

function setExpertMode(enabled) {
  el.allFieldsPanel.classList.toggle("hidden", !enabled);
  el.expertToggle.textContent = enabled ? "Expertenmodus: An" : "Expertenmodus: Aus";
  localStorage.setItem("gaeb-expert-mode", enabled ? "1" : "0");
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
      Gruppe: p.groupDisplay || "-",
      OZ: p.oz,
      Beschreibung: p.description,
      Menge: p.quantity,
      Einheit: p.unit,
      EP: p.unitPrice,
      GP: p.totalPrice,
      Waehrung: p.currency,
      Bereich: p.scope,
    }));
    const groupModel = buildGroupModel(state.filteredPositions);
    const lvRows = buildLvStructureRows(groupModel);
    const allRows = state.allFields.map((entry) => ({
      Nr: entry.index,
      Pfad: entry.path,
      Typ: entry.type,
      Wert: entry.value,
    }));

    const projectSheet = XLSX.utils.json_to_sheet(projectRows);
    const bidSheet = XLSX.utils.json_to_sheet(bidRows);
    const positionsSheet = XLSX.utils.json_to_sheet(posRows);
    const lvSheet = XLSX.utils.json_to_sheet(lvRows);
    const allFieldsSheet = XLSX.utils.json_to_sheet(allRows);

    styleKeyValueSheet(projectSheet, projectRows);
    styleKeyValueSheet(bidSheet, bidRows);
    stylePositionsSheet(positionsSheet, posRows);
    styleLvStructureSheet(lvSheet, lvRows);
    styleAllFieldsSheet(allFieldsSheet, allRows);

    XLSX.utils.book_append_sheet(wb, projectSheet, "Projekt");
    XLSX.utils.book_append_sheet(wb, bidSheet, "Angebot");
    XLSX.utils.book_append_sheet(wb, positionsSheet, "Positionen");
    XLSX.utils.book_append_sheet(wb, lvSheet, "LV-Struktur");
    XLSX.utils.book_append_sheet(wb, allFieldsSheet, "Alle XML-Felder");

    const fileName = `gaeb-export-${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(wb, fileName);
  });
}

function styleKeyValueSheet(sheet, rows) {
  const keyLength = Math.max("Feld".length, ...rows.map((r) => String(r.Feld || "").length));
  const valueLength = Math.max("Wert".length, ...rows.map((r) => String(r.Wert || "").length));
  sheet["!cols"] = [
    { wch: clampColWidth(keyLength + 2, 18, 32) },
    { wch: clampColWidth(valueLength + 2, 24, 90) },
  ];
}

function stylePositionsSheet(sheet, rows) {
  const headers = ["Nr", "Gruppe", "OZ", "Beschreibung", "Menge", "Einheit", "EP", "GP", "Waehrung", "Bereich"];
  const colWidths = headers.map((header) => {
    const maxLen = Math.max(
      header.length,
      ...rows.map((row) => String(row[header] ?? "").length),
    );
    if (header === "Beschreibung") {
      return { wch: clampColWidth(maxLen + 2, 36, 90) };
    }
    if (header === "Gruppe") {
      return { wch: clampColWidth(maxLen + 2, 16, 42) };
    }
    if (header === "OZ") {
      return { wch: clampColWidth(maxLen + 2, 10, 24) };
    }
    return { wch: clampColWidth(maxLen + 2, 10, 18) };
  });
  sheet["!cols"] = colWidths;
}

function styleLvStructureSheet(sheet, rows) {
  const headers = ["Typ", "Ebene", "Gruppe", "OZ", "Beschreibung", "Menge", "Einheit", "EP", "GP", "Waehrung"];
  const colWidths = headers.map((header) => {
    const maxLen = Math.max(header.length, ...rows.map((row) => String(row[header] ?? "").length));
    if (header === "Gruppe" || header === "Beschreibung") {
      return { wch: clampColWidth(maxLen + 2, 24, 90) };
    }
    if (header === "OZ") {
      return { wch: clampColWidth(maxLen + 2, 10, 24) };
    }
    return { wch: clampColWidth(maxLen + 2, 8, 16) };
  });
  sheet["!cols"] = colWidths;
}

function styleAllFieldsSheet(sheet, rows) {
  const headers = ["Nr", "Pfad", "Typ", "Wert"];
  const colWidths = headers.map((header) => {
    const maxLen = Math.max(
      header.length,
      ...rows.map((row) => String(row[header] ?? "").length),
    );
    if (header === "Pfad") {
      return { wch: clampColWidth(maxLen + 2, 30, 120) };
    }
    if (header === "Wert") {
      return { wch: clampColWidth(maxLen + 2, 24, 120) };
    }
    return { wch: clampColWidth(maxLen + 2, 8, 20) };
  });
  sheet["!cols"] = colWidths;
}

function clampColWidth(value, min, max) {
  return Math.max(min, Math.min(max, value));
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
  state.allFields = extractAllFields(xmlDoc);

  el.fileMeta.textContent = `Geladen: ${file.name} (${formatBytes(file.size)})`;
  el.rawXml.textContent = xmlText.slice(0, 120000);
  if (xmlText.length > 120000) {
    el.rawXml.textContent += "\n\n... Vorschau gekuerzt ...";
  }

  renderInfoGrid(el.projectInfo, state.projectInfo, "Keine Projektdaten gefunden.");
  renderBidInfo(state.bidInfo, "Keine Angebotsdaten gefunden.");
  renderAllFieldsTable(state.allFields);
  renderTree(xmlDoc);
  applyFiltersAndRender();
  el.exportBtn.disabled = false;
}

function extractProjectInfo(xmlDoc) {
  const project = firstByLocalName(xmlDoc, "Project");
  const prjInfo = firstByLocalName(xmlDoc, "PrjInfo");
  const tender = firstByLocalName(xmlDoc, "Tender");
  const award = firstByLocalName(xmlDoc, "Award");
  const boqInfo = firstByLocalName(xmlDoc, "BoQInfo");
  return {
    Projektname:
      deepText(prjInfo, ["LblPrj", "Name", "Description", "Title"]) ||
      deepText(project, ["Name", "Description", "Title"]) ||
      "-",
    "Projekt-ID": deepText(project, ["ID", "ProjectNo", "Number"]) || "-",
    "LV-Bezeichnung":
      deepText(boqInfo, ["Name", "Description", "Title"]) ||
      deepText(tender, ["Name", "Description", "Title"]) ||
      "-",
    Vergabeart: deepText(award, ["DP", "Award", "Type", "Procedure"]) || deepText(tender, ["Award", "Type", "Procedure"]) || "-",
    Quelle: xmlDoc.documentElement.getAttribute("xmlns") || "GAEB",
  };
}

function extractBidInfo(xmlDoc) {
  const bid = firstByLocalName(xmlDoc, "Bid");
  const ctr = firstByLocalName(xmlDoc, "CTR");
  const awardInfo = firstByLocalName(xmlDoc, "AwardInfo");
  const address = firstByLocalName(ctr || xmlDoc, "Address");
  const bidderName =
    [deepText(ctr, ["Name1"]), deepText(ctr, ["Name2"]), deepText(ctr, ["Name"])]
      .filter(Boolean)
      .join(" ")
      .trim() || "-";
  const addressStreet = deepText(address, ["Street"]) || "-";
  const addressPostalCode = deepText(address, ["PCode", "PostalCode", "Zip"]) || "-";
  const addressCity = deepText(address, ["City", "Town"]) || "-";
  const contactPhone = deepText(address, ["Phone", "Telephone", "Tel"]) || "-";
  const contactFax = deepText(address, ["Fax"]) || "-";
  const contactEmail = deepText(address, ["Email", "Mail"]) || "-";
  if (!bid) {
    return {
      Bieter: bidderName,
      Angebotsnummer: deepText(ctr, ["AcctsPayNo"]) || "-",
      Datum: deepText(xmlDoc, ["Date", "IssueDate", "Created"]) || "-",
      Waehrung:
        deepText(awardInfo, ["Cur", "Currency", "CurrencyCode"]) ||
        xmlDoc.documentElement.getAttribute("Cur") ||
        "EUR",
      Strasse: addressStreet,
      PLZ: addressPostalCode,
      Ort: addressCity,
      Telefon: contactPhone,
      Fax: contactFax,
      Email: contactEmail,
    };
  }
  return {
    Bieter: deepText(bid, ["Bidder", "Company", "Name"]) || bidderName,
    Angebotsnummer: deepText(bid, ["ID", "BidNo", "Number"]) || deepText(ctr, ["AcctsPayNo"]) || "-",
    Datum: deepText(bid, ["Date", "IssueDate", "Created"]) || "-",
    Waehrung:
      deepText(bid, ["Currency", "CurrencyCode"]) ||
      deepText(awardInfo, ["Cur", "Currency", "CurrencyCode"]) ||
      xmlDoc.documentElement.getAttribute("Cur") ||
      "EUR",
    Strasse: addressStreet,
    PLZ: addressPostalCode,
    Ort: addressCity,
    Telefon: contactPhone,
    Fax: contactFax,
    Email: contactEmail,
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
    const description = extractDescriptionText(node) || collectTextSnippet(node, 220) || "-";
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
    const groupChain = extractGroupChain(node, oz);

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
      groupChain,
      groupDisplay: groupChain.length ? groupChain[groupChain.length - 1].code : "-",
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
      (p.groupDisplay || "").toLowerCase().includes(search) ||
      p.groupChain.some((g) => g.label.toLowerCase().includes(search)) ||
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
    td.colSpan = 8;
    td.className = "empty-state";
    td.textContent = totalCount
      ? "Keine Positionen fuer aktuelle Filter gefunden."
      : "Keine Positionen gefunden.";
    tr.appendChild(td);
    el.tableBody.appendChild(tr);
  } else if (state.viewMode === "list") {
    renderListRows(rows);
  } else {
    renderGroupedRows(rows);
  }

  const shown = rows.length;
  const groupModel = buildGroupModel(rows);
  el.tableMeta.textContent = `${shown} von ${totalCount} Positionen sichtbar · ${groupModel.roots.length} Hauptgruppen`;
}

function renderListRows(rows) {
  for (const row of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(String(row.index))}</td>
      <td>${escapeHtml(row.groupDisplay || "-")}</td>
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

function renderGroupedRows(rows) {
  const model = buildGroupModel(rows);

  for (const rootKey of model.roots) {
    renderGroupNode(model, rootKey);
  }
}

function renderGroupNode(model, groupKey) {
  const group = model.groups.get(groupKey);
  if (!group) {
    return;
  }
  const isCollapsed = state.collapsedGroupIds.has(group.key);

  const tr = document.createElement("tr");
  tr.className = "group-row";
  tr.innerHTML = `
    <td></td>
    <td class="group-cell" style="padding-left:${Math.max(0, group.level - 1) * 14 + 8}px">
      <button class="group-toggle" type="button" data-group-key="${escapeHtml(group.key)}">${isCollapsed ? "▸" : "▾"}</button>
      <span>${escapeHtml(group.label)}</span>
    </td>
    <td colspan="4"></td>
    <td></td>
    <td>${formatCurrency(group.totalPrice, group.currency)}</td>
  `;
  const toggle = tr.querySelector(".group-toggle");
  toggle?.addEventListener("click", () => {
    if (state.collapsedGroupIds.has(group.key)) {
      state.collapsedGroupIds.delete(group.key);
    } else {
      state.collapsedGroupIds.add(group.key);
    }
    renderGroupedRowsFromState();
  });
  el.tableBody.appendChild(tr);

  if (isCollapsed) {
    return;
  }

  for (const childKey of group.children) {
    renderGroupNode(model, childKey);
  }

  for (const position of group.positions) {
    const rowTr = document.createElement("tr");
    rowTr.innerHTML = `
      <td>${escapeHtml(String(position.index))}</td>
      <td style="padding-left:${Math.max(0, group.level) * 14 + 8}px">${escapeHtml(position.groupDisplay || "-")}</td>
      <td>${escapeHtml(position.oz)}</td>
      <td>${escapeHtml(position.description)}</td>
      <td>${formatNumber(position.quantity)}</td>
      <td>${escapeHtml(position.unit)}</td>
      <td>${formatCurrency(position.unitPrice, position.currency)}</td>
      <td>${formatCurrency(position.totalPrice, position.currency)}</td>
    `;
    el.tableBody.appendChild(rowTr);
  }
}

function renderGroupedRowsFromState() {
  el.tableBody.innerHTML = "";
  renderGroupedRows(state.filteredPositions);
}

function renderBidInfo(data, emptyText) {
  el.bidInfo.innerHTML = "";
  if (!data || !Object.keys(data).length) {
    el.bidInfo.classList.add("empty-state");
    el.bidInfo.textContent = emptyText;
    return;
  }

  el.bidInfo.classList.remove("empty-state");
  const bidder = String(data.Bieter || "-");
  const addressLine = String(data.Strasse || "-");
  const cityLine = [data.PLZ, data.Ort].filter((value) => value && value !== "-").join(" ") || "-";

  const card = document.createElement("article");
  card.className = "bid-card";
  card.innerHTML = `
    <span class="key">Bieter</span>
    <div class="bid-name">${escapeHtml(bidder)}</div>
    <div class="bid-address">${escapeHtml(addressLine)}<br>${escapeHtml(cityLine)}</div>
    <div class="bid-contact">
      <span>Telefon: ${escapeHtml(String(data.Telefon || "-"))}</span>
      <span>Fax: ${escapeHtml(String(data.Fax || "-"))}</span>
      <span>Email: ${escapeHtml(String(data.Email || "-"))}</span>
    </div>
  `;

  el.bidInfo.appendChild(card);

  const metaGrid = document.createElement("div");
  metaGrid.className = "bid-meta-grid";
  const metaEntries = [
    ["Angebotsnummer", data.Angebotsnummer || "-"],
    ["Datum", data.Datum || "-"],
    ["Waehrung", data.Waehrung || "-"],
  ];
  for (const [key, value] of metaEntries) {
    const item = document.createElement("article");
    item.className = "kv-item";
    item.innerHTML = `
      <span class="key">${escapeHtml(key)}</span>
      <strong>${escapeHtml(String(value))}</strong>
    `;
    metaGrid.appendChild(item);
  }
  el.bidInfo.appendChild(metaGrid);
}

function renderAllFieldsTable(rows) {
  el.allFieldsTableBody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 4;
    td.className = "empty-state";
    td.textContent = "Keine XML-Felder gefunden.";
    tr.appendChild(td);
    el.allFieldsTableBody.appendChild(tr);
    el.allFieldsMeta.textContent = "0 Felder";
    return;
  }

  for (const row of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(String(row.index))}</td>
      <td><code>${escapeHtml(row.path)}</code></td>
      <td>${escapeHtml(row.type)}</td>
      <td>${escapeHtml(row.value)}</td>
    `;
    el.allFieldsTableBody.appendChild(tr);
  }

  el.allFieldsMeta.textContent = `${rows.length} Felder sichtbar`;
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
    if (key === "Projektname") {
      item.classList.add("full-width");
    }
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
    const tag = localName(child).toLowerCase();
    if (["up", "it", "qty", "qtakeoff", "qu", "ep", "gp", "price", "total", "totalprice"].includes(tag)) {
      continue;
    }
    const text = (child.textContent || "").trim();
    if (text && !looksLikeNumber(text)) {
      chunks.push(text);
    }
    if (chunks.join(" ").length > maxLen) {
      break;
    }
  }
  return chunks.join(" ").slice(0, maxLen);
}

function extractDescriptionText(node) {
  const descriptionNode = firstByLocalName(node, "Description");
  const source = descriptionNode || node;
  const preferredTags = [
    "OutlineText",
    "OutlTxt",
    "TextOutlTxt",
    "ShortText",
    "LongText",
    "Text",
    "span",
    "p",
  ].map((n) => n.toLowerCase());

  const chunks = [];
  for (const child of Array.from(source.getElementsByTagName("*"))) {
    if (!preferredTags.includes(localName(child).toLowerCase())) {
      continue;
    }
    const text = (child.textContent || "").trim();
    if (text && !looksLikeNumber(text)) {
      chunks.push(text);
    }
  }
  if (chunks.length) {
    return chunks.join(" ").slice(0, 220);
  }

  const plain = (source.textContent || "").trim();
  if (plain && !looksLikeNumber(plain)) {
    return plain.slice(0, 220);
  }
  return "";
}

function extractAllFields(xmlDoc) {
  const root = xmlDoc.documentElement;
  if (!root) {
    return [];
  }

  const rows = [];
  let idx = 1;

  function walk(node, path) {
    const children = Array.from(node.children || []);

    for (const attr of Array.from(node.attributes || [])) {
      rows.push({
        index: idx++,
        path: `${path}/@${attr.name}`,
        type: "Attribut",
        value: attr.value || "-",
      });
    }

    const text = (node.textContent || "").trim();
    if (!children.length && text) {
      rows.push({
        index: idx++,
        path,
        type: "Text",
        value: text,
      });
      return;
    }

    const siblingCountByName = {};
    for (const child of children) {
      const name = localName(child);
      siblingCountByName[name] = (siblingCountByName[name] || 0) + 1;
      const childPath = `${path}/${name}[${siblingCountByName[name]}]`;
      walk(child, childPath);
    }
  }

  walk(root, `/${localName(root)}`);
  return rows;
}

function extractGroupChain(node, oz) {
  const groupNames = ["boqctgy", "ctgy", "category", "group", "section", "chapter"];
  const chain = [];
  let current = node?.parentElement || null;
  while (current) {
    const name = localName(current).toLowerCase();
    if (groupNames.includes(name)) {
      chain.push(current);
    }
    current = current.parentElement;
  }
  chain.reverse();

  const normalizedChain = chain.map((groupNode, idx) => {
    const rawCode =
      (groupNode.getAttribute("RNoPart") || "").trim() ||
      (groupNode.getAttribute("NoPart") || "").trim() ||
      (groupNode.getAttribute("ID") || "").trim();
    if (!rawCode) {
      return null;
    }
    const codeSegment = normalizeGroupCodeSegment(rawCode);
    const parentCode = idx > 0 ? chain[idx - 1].__groupCode || "" : "";
    const code = parentCode ? `${parentCode}${codeSegment}.` : `${codeSegment}.`;
    groupNode.__groupCode = code;
    const title = extractGroupTitle(groupNode);
    const idPart = (groupNode.getAttribute("ID") || "").trim() || codeSegment;
    const key = idx > 0 ? `${chain[idx - 1].__groupKey}/${idPart}` : idPart;
    groupNode.__groupKey = key;
    return {
      key,
      code,
      title,
      label: title ? `${code} ${title}` : code,
      level: idx + 1,
    };
  }).filter(Boolean);

  if (normalizedChain.length) {
    return normalizedChain;
  }
  return [];
}

function normalizeGroupCodeSegment(value) {
  const digits = String(value || "").replace(/[^\d]/g, "");
  if (!digits) {
    return "00";
  }
  if (digits.length === 1) {
    return `0${digits}`;
  }
  return digits;
}

function extractGroupTitle(groupNode) {
  const directChildren = Array.from(groupNode.children || []);
  const preferred = ["Name", "Title", "LblTx", "LblTxt", "LblBoQCtgy", "ShortText"];
  for (const tag of preferred) {
    const direct = directChildren.find((child) => localName(child).toLowerCase() === tag.toLowerCase());
    if (direct) {
      const text = (direct.textContent || "").trim().replace(/\s+/g, " ");
      if (text && !looksLikeNumber(text)) {
        return text;
      }
    }
  }
  return "";
}

function buildGroupModel(rows) {
  const groups = new Map();
  const roots = [];

  const ensureGroup = (group, parentKey) => {
    if (!groups.has(group.key)) {
      groups.set(group.key, {
        key: group.key,
        label: group.label,
        code: group.code,
        title: group.title,
        level: group.level,
        parentKey,
        children: [],
        positions: [],
        totalPrice: 0,
        currency: "EUR",
      });
      if (!parentKey) {
        roots.push(group.key);
      } else {
        const parent = groups.get(parentKey);
        if (parent && !parent.children.includes(group.key)) {
          parent.children.push(group.key);
        }
      }
    }
    return groups.get(group.key);
  };

  for (const row of rows) {
    const chain = row.groupChain?.length ? row.groupChain : [];
    let parentKey = null;
    for (const group of chain) {
      ensureGroup(group, parentKey);
      parentKey = group.key;
    }
    if (!parentKey) {
      parentKey = "ungrouped";
      ensureGroup({ key: "ungrouped", label: "-", code: "-", title: "", level: 1 }, null);
    }
    const leafGroup = groups.get(parentKey);
    leafGroup.positions.push(row);
    leafGroup.totalPrice += row.totalPrice || 0;
    leafGroup.currency = row.currency || "EUR";
  }

  const aggregate = (key) => {
    const group = groups.get(key);
    if (!group) {
      return 0;
    }
    let total = group.totalPrice;
    for (const childKey of group.children) {
      total += aggregate(childKey);
    }
    group.totalPrice = total;
    return total;
  };

  for (const rootKey of roots) {
    aggregate(rootKey);
  }

  return { groups, roots };
}

function buildLvStructureRows(groupModel) {
  const rows = [];
  const walk = (key) => {
    const group = groupModel.groups.get(key);
    if (!group) {
      return;
    }
    rows.push({
      Typ: "Gruppe",
      Ebene: group.level,
      Gruppe: group.label,
      OZ: "",
      Beschreibung: group.title || "-",
      Menge: "",
      Einheit: "",
      EP: "",
      GP: group.totalPrice,
      Waehrung: group.currency || "EUR",
    });
    for (const childKey of group.children) {
      walk(childKey);
    }
    for (const position of group.positions) {
      rows.push({
        Typ: "Position",
        Ebene: group.level + 1,
        Gruppe: group.label,
        OZ: position.oz,
        Beschreibung: position.description,
        Menge: position.quantity,
        Einheit: position.unit,
        EP: position.unitPrice,
        GP: position.totalPrice,
        Waehrung: position.currency,
      });
    }
  };

  for (const rootKey of groupModel.roots) {
    walk(rootKey);
  }
  return rows;
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

function looksLikeNumber(value) {
  const raw = String(value || "")
    .trim()
    .replace(/\s+/g, "");
  if (!raw) {
    return false;
  }
  return /^[-+]?\d{1,3}([.,]\d{3})*([.,]\d+)?$/.test(raw) || /^[-+]?\d+([.,]\d+)?$/.test(raw);
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
