(() => {
  "use strict";

  const STATUS_OPTIONS = ["New", "Working", "Stuck", "On Hold", "Out for Review", "Complete"];
  const ASSIGNED_OPTIONS = ["Todd", "Travis", "Brady"];

  const COLORS = {
    black: "FF000000",
    white: "FFFFFFFF",
    hyperlink: "FF0563C1",
    labelFill: "FFEEE8F7",
    statusTab: "FFFFF2CC",
    updatesTab: "FFE2F0D9",
    attachmentsTab: "FFD9E1F2",
    taskTabs: [
      "FFE4DFEC", "FFDDEBF7", "FFFCE4D6", "FFE2F0D9",
      "FFFFE699", "FFF4CCCC", "FFD9EAD3", "FFD0E0E3"
    ]
  };

  const fileInput = document.getElementById("fileInput");
  const chooseButton = document.getElementById("chooseButton");
  const changeButton = document.getElementById("changeButton");
  const cleanButton = document.getElementById("cleanButton");
  const dropZone = document.getElementById("dropZone");
  const filePanel = document.getElementById("filePanel");
  const fileName = document.getElementById("fileName");
  const fileMeta = document.getElementById("fileMeta");
  const statusPanel = document.getElementById("statusPanel");
  const statusTitle = document.getElementById("statusTitle");
  const statusText = document.getElementById("statusText");
  const spinner = document.getElementById("spinner");
  const downloadButton = document.getElementById("downloadButton");

  let selectedFile = null;
  let currentObjectUrl = null;

  chooseButton.addEventListener("click", (e) => {
    e.stopPropagation();
    fileInput.click();
  });
  changeButton.addEventListener("click", () => fileInput.click());
  dropZone.addEventListener("click", () => fileInput.click());
  dropZone.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") fileInput.click();
  });

  ["dragenter", "dragover"].forEach(type => {
    dropZone.addEventListener(type, (e) => {
      e.preventDefault();
      dropZone.classList.add("dragging");
    });
  });
  ["dragleave", "drop"].forEach(type => {
    dropZone.addEventListener(type, (e) => {
      e.preventDefault();
      dropZone.classList.remove("dragging");
    });
  });
  dropZone.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files?.[0];
    if (file) setFile(file);
  });
  fileInput.addEventListener("change", () => {
    const file = fileInput.files?.[0];
    if (file) setFile(file);
  });
  cleanButton.addEventListener("click", processSelectedFile);

  function setFile(file) {
    resetDownload();
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      showStatus("That isn't an .xlsx file.", "Choose the Excel export from Monday.com.", false, true);
      selectedFile = null;
      cleanButton.disabled = true;
      return;
    }
    selectedFile = file;
    fileName.textContent = file.name;
    fileMeta.textContent = `${formatBytes(file.size)} · Excel workbook`;
    filePanel.classList.remove("hidden");
    cleanButton.disabled = false;
    statusPanel.classList.add("hidden");
  }

  function formatBytes(bytes) {
    if (bytes < 1024) return `${bytes} bytes`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  }

  function showStatus(title, text, busy = false, isError = false) {
    statusPanel.classList.remove("hidden");
    statusPanel.style.background = isError ? "#fff0f0" : "#f6f6f7";
    statusTitle.style.color = isError ? "#9e2525" : "";
    statusTitle.textContent = title;
    statusText.textContent = text;
    spinner.classList.toggle("hidden", !busy);
  }

  function resetDownload() {
    downloadButton.classList.add("hidden");
    downloadButton.removeAttribute("download");
    downloadButton.href = "#";
    if (currentObjectUrl) {
      URL.revokeObjectURL(currentObjectUrl);
      currentObjectUrl = null;
    }
  }

  async function processSelectedFile() {
    if (!selectedFile) return;
    resetDownload();

    if (typeof ExcelJS === "undefined") {
      showStatus(
        "Excel library couldn't load.",
        "The browser could not reach the ExcelJS library. Check your network/security settings and refresh the page.",
        false,
        true
      );
      return;
    }

    cleanButton.disabled = true;
    showStatus("Reading Monday export…", "Your workbook is being processed locally in this browser.", true);

    try {
      // Let the UI paint before CPU-heavy workbook work.
      await new Promise(resolve => setTimeout(resolve, 30));

      const buffer = await selectedFile.arrayBuffer();
      const inputWorkbook = new ExcelJS.Workbook();
      await inputWorkbook.xlsx.load(buffer);

      const creativeSheet = findSheet(inputWorkbook, "creative requests") || inputWorkbook.worksheets[0];
      if (!creativeSheet) throw new Error("No worksheets were found in the Excel file.");

      const requests = parseRequests(creativeSheet);
      if (!requests.length) {
        throw new Error("No creative request rows were found. Make sure this is the Monday Creative Requests export.");
      }

      const updates = parseUpdates(inputWorkbook);
      showStatus(
        "Building clean workbook…",
        `${requests.length} request${requests.length === 1 ? "" : "s"} and ${updates.length} update/repl${updates.length === 1 ? "y" : "ies"} found.`,
        true
      );

      await new Promise(resolve => setTimeout(resolve, 20));

      const outputWorkbook = buildWorkbook(requests, updates);
      const outputBuffer = await outputWorkbook.xlsx.writeBuffer();
      const blob = new Blob(
        [outputBuffer],
        { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
      );

      currentObjectUrl = URL.createObjectURL(blob);
      const outputName = currentOutputName();

      downloadButton.href = currentObjectUrl;
      downloadButton.download = outputName;
      downloadButton.textContent = `Download ${outputName}`;
      downloadButton.classList.remove("hidden");

      const attachmentCount = collectAttachments(requests).length;
      showStatus(
        "Ready!",
        `${requests.length} requests · ${updates.length} updates/replies · ${attachmentCount} attachment references`,
        false
      );
    } catch (err) {
      console.error(err);
      showStatus("Could not clean this workbook.", err?.message || String(err), false, true);
    } finally {
      cleanButton.disabled = !selectedFile;
    }
  }

  function findSheet(workbook, name) {
    const target = name.toLowerCase();
    return workbook.worksheets.find(ws => ws.name.toLowerCase() === target);
  }

  function rawValue(cell) {
    const v = cell?.value;
    if (v == null) return "";
    if (v instanceof Date) return v;
    if (typeof v === "object") {
      if (v.text != null) return v.text;
      if (v.richText) return v.richText.map(x => x.text || "").join("");
      if (v.result != null) return v.result;
      if (v.formula != null) return v.result ?? "";
      if (v.hyperlink != null) return v.text || v.hyperlink;
    }
    return v;
  }

  function cleanText(value) {
    if (value == null) return "";
    if (value instanceof Date) return value;
    return String(value).trim();
  }

  function textValue(cell) {
    const v = rawValue(cell);
    return v instanceof Date ? v : cleanText(v);
  }

  function normalizeStatus(value) {
    const s = String(cleanText(value)).trim();
    const mapping = {
      "working on it": "Working",
      "working": "Working",
      "new": "New",
      "stuck": "Stuck",
      "on hold": "On Hold",
      "out for review": "Out for Review",
      "complete": "Complete",
      "completed": "Complete",
      "approved": "Working"
    };
    return mapping[s.toLowerCase()] || s || "New";
  }

  function ownerToAssigned(owner) {
    const s = String(cleanText(owner)).toLowerCase();
    if (s.includes("todd")) return "Todd";
    if (s.includes("travis")) return "Travis";
    if (s.includes("brady")) return "Brady";
    return "";
  }

  function defaultAction(item) {
    const status = normalizeStatus(item["Status"]);
    if (status === "On Hold") return "No active work unless the request is reactivated.";
    if (status === "Complete") return "Complete — no further action unless revisions are requested.";
    if (String(cleanText(item["Status"])).toLowerCase() === "approved") {
      return "Confirm final production specs and release/deliver the approved creative.";
    }
    const specs = String(cleanText(item["Specs"]));
    const brief = String(cleanText(item["Creative Brief"]));
    if (specs) return `Review the brief and assets, confirm ${specs}, then begin/continue production.`;
    if (brief) return "Review the creative brief and assets, resolve any missing production details, then begin/continue production.";
    return "Review the request, confirm requirements, and begin/continue production.";
  }

  function findHeaderRow(ws) {
    const maxRows = Math.min(ws.rowCount || 0, 30);
    const maxCols = Math.min(ws.columnCount || 0, 40);
    for (let r = 1; r <= maxRows; r++) {
      const values = [];
      for (let c = 1; c <= maxCols; c++) {
        values.push(String(cleanText(textValue(ws.getCell(r, c)))));
      }
      if (values.includes("Name") && values.includes("Due Date") &&
          values.includes("Status") && values.includes("Creative Brief")) {
        const cols = {};
        values.forEach((v, i) => { if (v) cols[v] = i + 1; });
        return { row: r, cols };
      }
    }
    throw new Error("Could not find the Monday.com request header row.");
  }

  function parseRequests(ws) {
    const { row: headerRow, cols } = findHeaderRow(ws);
    const requests = [];
    let currentParent = null;
    let inSubitems = false;

    const val = (row, name) => cols[name] ? textValue(ws.getCell(row, cols[name])) : "";

    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const name = String(cleanText(val(r, "Name")));
      const firstCell = String(cleanText(textValue(ws.getCell(r, 1))));

      if (firstCell === "Subitems" && String(cleanText(textValue(ws.getCell(r, 2)))) === "Name") {
        inSubitems = true;
        continue;
      }

      if (inSubitems) {
        const childName = String(cleanText(textValue(ws.getCell(r, 2))));
        if (!firstCell && childName && currentParent) {
          const files = String(cleanText(textValue(ws.getCell(r, 6))));
          const subStatus = String(cleanText(textValue(ws.getCell(r, 4))));
          const subDate = textValue(ws.getCell(r, 5));
          let detail = childName;
          if (subStatus) detail += ` — ${subStatus}`;
          if (subDate) detail += ` (${displayDate(subDate)})`;
          if (files) detail += ` — ${files}`;

          currentParent._subitems.push(detail);
          currentParent._subitem_records.push({
            "Name": childName,
            "Status": subStatus,
            "Date": subDate,
            "Files": files,
            "Item ID": String(cleanText(textValue(ws.getCell(r, 8))))
          });
          continue;
        } else {
          inSubitems = false;
        }
      }

      if (!name && !firstCell) continue;
      if (!name) continue;

      const due = val(r, "Due Date");
      const status = String(cleanText(val(r, "Status")));
      const brief = String(cleanText(val(r, "Creative Brief")));
      const requester = String(cleanText(val(r, "Requested by")));
      if (!(status || due || brief || requester)) continue;

      const item = {};
      Object.entries(cols).forEach(([key, col]) => {
        item[key] = textValue(ws.getCell(r, col));
      });
      item._source_row = r;
      item._subitems = [];
      item._subitem_records = [];
      requests.push(item);
      currentParent = item;
    }
    return requests;
  }

  function parseUpdates(inputWorkbook) {
    const ws = findSheet(inputWorkbook, "updates");
    if (!ws) return [];

    let headerRow = null;
    const headers = {};

    for (let r = 1; r <= Math.min(ws.rowCount || 0, 15); r++) {
      const vals = [];
      for (let c = 1; c <= ws.columnCount; c++) {
        vals.push(String(cleanText(textValue(ws.getCell(r, c)))));
      }
      if (vals.includes("Item ID") && vals.includes("Item Name") && vals.includes("Update Content")) {
        headerRow = r;
        vals.forEach((v, i) => {
          if (!v) return;
          if (!headers[v]) headers[v] = [];
          headers[v].push(i + 1);
        });
        break;
      }
    }

    if (!headerRow) return [];

    const firstCol = name => headers[name]?.[0] || null;
    const contentCols = headers["Content Type"] || [];
    const rows = [];

    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const get = name => firstCol(name) ? textValue(ws.getCell(r, firstCol(name))) : "";
      const itemId = String(cleanText(get("Item ID")));
      const itemName = String(cleanText(get("Item Name")));
      const content = String(cleanText(get("Update Content")));
      if (!(itemId || itemName || content)) continue;

      let kind = "";
      for (const c of contentCols) {
        const candidate = String(cleanText(textValue(ws.getCell(r, c))));
        if (candidate) {
          kind = candidate;
          break;
        }
      }

      rows.push({
        "Item ID": itemId,
        "Item Name": itemName,
        "Type": kind || "Update",
        "User": String(cleanText(get("User"))),
        "Created At": parseMondayDate(get("Created At")),
        "Content": content,
        "Asset IDs": String(cleanText(get("Asset IDs"))),
        "Post ID": String(cleanText(get("Post ID"))),
        "Parent Post ID": String(cleanText(get("Parent Post ID")))
      });
    }

    rows.sort((a, b) => {
      const aa = a["Created At"] instanceof Date ? a["Created At"].getTime() : Number.MAX_SAFE_INTEGER;
      const bb = b["Created At"] instanceof Date ? b["Created At"].getTime() : Number.MAX_SAFE_INTEGER;
      return aa - bb;
    });
    return rows;
  }

  function parseMondayDate(value) {
    if (value instanceof Date) return value;
    const s = String(cleanText(value));
    if (!s) return "";
    const m = s.match(/^(\d{1,2})\/([A-Za-z]+)\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})\s*([AP]M)$/i);
    if (!m) {
      const parsed = new Date(s);
      return Number.isNaN(parsed.getTime()) ? s : parsed;
    }
    const months = {
      january:0, february:1, march:2, april:3, may:4, june:5,
      july:6, august:7, september:8, october:9, november:10, december:11
    };
    let hour = Number(m[4]);
    const ampm = m[7].toUpperCase();
    if (ampm === "PM" && hour !== 12) hour += 12;
    if (ampm === "AM" && hour === 12) hour = 0;
    return new Date(Number(m[3]), months[m[2].toLowerCase()], Number(m[1]), hour, Number(m[5]), Number(m[6]));
  }

  function extractUrls(value) {
    const s = String(cleanText(value));
    if (!s) return [];
    return s.match(/https?:\/\/[^\s,]+/g) || [];
  }

  function collectAttachments(requests) {
    const records = [];
    for (const item of requests) {
      const parentId = String(cleanText(item["Item ID (auto generated)"]));
      const parentName = String(cleanText(item["Name"]));

      for (const url of extractUrls(item["Assets"])) {
        records.push({
          "Parent Item ID": parentId,
          "Request": parentName,
          "Source": "Request Assets",
          "Subitem": "",
          "URL": url
        });
      }
      for (const sub of item._subitem_records || []) {
        for (const url of extractUrls(sub["Files"])) {
          records.push({
            "Parent Item ID": parentId,
            "Request": parentName,
            "Source": "Subitem Files",
            "Subitem": String(cleanText(sub["Name"])),
            "URL": url
          });
        }
      }
    }
    return records;
  }

  function dueSortValue(value) {
    if (value instanceof Date) return value.getTime();
    const s = String(cleanText(value));
    if (!s) return Number.MAX_SAFE_INTEGER;
    const d = new Date(s);
    return Number.isNaN(d.getTime()) ? Number.MAX_SAFE_INTEGER : d.getTime();
  }

  function safeSheetName(name, used) {
    let clean = String(cleanText(name)).replace(/[\\/*?:[\]]/g, "-").replace(/\s+/g, " ").trim() || "Task";
    const base = clean.slice(0, 31);
    let candidate = base;
    let n = 2;
    while (used.has(candidate.toLowerCase())) {
      const suffix = ` ${n}`;
      candidate = base.slice(0, 31 - suffix.length) + suffix;
      n += 1;
    }
    used.add(candidate.toLowerCase());
    return candidate;
  }

  function displayDate(value) {
    if (!(value instanceof Date)) return String(cleanText(value));
    return `${String(value.getMonth()+1).padStart(2,"0")}/${String(value.getDate()).padStart(2,"0")}/${value.getFullYear()}`;
  }

  function currentOutputName() {
    const d = new Date();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    const yy = String(d.getFullYear()).slice(-2);
    return `${mm}-${dd}-${yy}_CR_ToDos.xlsx`;
  }

  function setBasicCellStyle(cell, size = 8, bold = false) {
    cell.font = { name: "Arial", size, bold, color: { argb: COLORS.black } };
    cell.alignment = { horizontal: "left", vertical: "top", wrapText: true };
  }

  function setHyperlinkStyle(cell, size = 8, centered = false) {
    cell.font = {
      name: "Arial",
      size,
      bold: true,
      color: { argb: COLORS.hyperlink },
      underline: true
    };
    cell.alignment = {
      horizontal: centered ? "center" : "left",
      vertical: "center",
      wrapText: true
    };
  }

  function setSheetTabColor(ws, argb) {
    ws.properties.tabColor = { argb };
  }

  function buildWorkbook(requests, updates) {
    const wb = new ExcelJS.Workbook();
    wb.creator = "TheMondayCleaner";
    wb.created = new Date();
    wb.modified = new Date();
    wb.calcProperties.fullCalcOnLoad = true;
    wb.calcProperties.forceFullCalc = true;
    wb.calcProperties.calcMode = "auto";

    const statusWs = wb.addWorksheet("Status");
    setSheetTabColor(statusWs, COLORS.statusTab);

    const headers = [
      "Priority", "Worksheet", "To-Do", "Status", "Due Date",
      "Action", "Requester", "Assigned", "Monday URL", "SharePoint URL"
    ];
    statusWs.addRow(headers);

    const sortedItems = [...requests].sort((a, b) => {
      const statusA = normalizeStatus(a["Status"]);
      const statusB = normalizeStatus(b["Status"]);
      const exemptA = ["Complete", "On Hold"].includes(statusA) ? 1 : 0;
      const exemptB = ["Complete", "On Hold"].includes(statusB) ? 1 : 0;
      if (exemptA !== exemptB) return exemptA - exemptB;
      const dueA = dueSortValue(a["Due Date"]);
      const dueB = dueSortValue(b["Due Date"]);
      if (dueA !== dueB) return dueA - dueB;
      return String(cleanText(a["Name"])).localeCompare(String(cleanText(b["Name"])));
    });

    sortedItems.forEach((item, i) => {
      const row = statusWs.addRow([
        i + 1,
        "",
        String(cleanText(item["Name"])),
        normalizeStatus(item["Status"]),
        item["Due Date"] || "",
        defaultAction(item),
        String(cleanText(item["Requested by"])) || String(cleanText(item["RMD Owner"])),
        ownerToAssigned(item["Owner"]),
        "",
        ""
      ]);
      row.getCell(4).dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: [`"${STATUS_OPTIONS.join(",")}"`]
      };
      row.getCell(8).dataValidation = {
        type: "list",
        allowBlank: true,
        formulae: [`"${ASSIGNED_OPTIONS.join(",")}"`]
      };
      if (row.getCell(5).value instanceof Date) row.getCell(5).numFmt = "mm/dd/yyyy";
    });

    statusWs.columns = [
      { width: 10 }, { width: 14 }, { width: 42 }, { width: 18 }, { width: 13 },
      { width: 58 }, { width: 22 }, { width: 14 }, { width: 38 }, { width: 38 }
    ];
    statusWs.eachRow(row => row.eachCell({ includeEmpty: true }, cell => setBasicCellStyle(cell)));
    statusWs.getRow(1).eachCell(cell => setBasicCellStyle(cell, 9, true));
    statusWs.views = [{ state: "frozen", ySplit: 1 }];

    // Updates
    const updatesWs = wb.addWorksheet("Updates");
    setSheetTabColor(updatesWs, COLORS.updatesTab);
    updatesWs.addRow([
      "Request / Item", "Parent Item ID", "Item ID", "Date", "By",
      "Type", "Note / Update", "Asset IDs", "Post ID", "Parent Post ID"
    ]);

    const idToParent = new Map();
    const nameToParent = new Map();
    sortedItems.forEach(item => {
      const pid = String(cleanText(item["Item ID (auto generated)"]));
      const pname = String(cleanText(item["Name"]));
      if (pid) idToParent.set(pid, [pname, pid]);
      if (pname) nameToParent.set(pname.toLowerCase(), [pname, pid]);
      (item._subitem_records || []).forEach(sub => {
        const sid = String(cleanText(sub["Item ID"]));
        const sname = String(cleanText(sub["Name"]));
        if (sid) idToParent.set(sid, [pname, pid]);
        if (sname && !nameToParent.has(sname.toLowerCase())) nameToParent.set(sname.toLowerCase(), [pname, pid]);
      });
    });

    updates.forEach(upd => {
      const itemId = String(cleanText(upd["Item ID"]));
      const itemName = String(cleanText(upd["Item Name"]));
      let parent = idToParent.get(itemId);
      if (!parent) parent = nameToParent.get(itemName.toLowerCase()) || ["", ""];
      const parentName = parent[0] || "";
      const displayName =
        itemName && parentName && itemName.toLowerCase() !== parentName.toLowerCase()
          ? `${parentName} → ${itemName}`
          : (parentName || itemName);

      const row = updatesWs.addRow([
        displayName, parent[1], itemId, upd["Created At"], upd["User"],
        upd["Type"], upd["Content"], upd["Asset IDs"], upd["Post ID"], upd["Parent Post ID"]
      ]);
      if (row.getCell(4).value instanceof Date) row.getCell(4).numFmt = "mm/dd/yy h:mm AM/PM";
    });

    const updateWidths = [52, 18, 18, 19, 22, 12, 90, 22, 18, 18];
    updateWidths.forEach((w, i) => updatesWs.getColumn(i + 1).width = w);
    [2, 3, 8, 9, 10].forEach(c => updatesWs.getColumn(c).hidden = true);
    updatesWs.eachRow(row => row.eachCell({ includeEmpty: true }, cell => setBasicCellStyle(cell)));
    updatesWs.getRow(1).eachCell(cell => setBasicCellStyle(cell, 9, true));
    for (let r = 2; r <= updatesWs.rowCount; r++) {
      updatesWs.getCell(r, 1).font = { name: "Arial", size: 8, bold: true, color: { argb: COLORS.black } };
    }
    updatesWs.getCell("L1").value = { text: "← BACK TO STATUS", hyperlink: "#'Status'!A1" };
    setHyperlinkStyle(updatesWs.getCell("L1"), 8, true);
    updatesWs.getColumn(12).width = 20;
    updatesWs.views = [{ state: "frozen", xSplit: 1, ySplit: 1, topLeftCell: "B2" }];
    updatesWs.autoFilter = { from: "A1", to: `J${Math.max(1, updatesWs.rowCount)}` };

    // Attachments
    const attachmentsWs = wb.addWorksheet("Attachments");
    setSheetTabColor(attachmentsWs, COLORS.attachmentsTab);
    attachmentsWs.addRow(["Request", "Parent Item ID", "Source", "Subitem", "File / Link", "Monday URL"]);
    const attachments = collectAttachments(sortedItems);
    attachments.forEach(rec => {
      const filename = rec.URL.includes("/") ? rec.URL.split("/").pop() : rec.URL;
      const row = attachmentsWs.addRow([
        rec.Request, rec["Parent Item ID"], rec.Source, rec.Subitem, filename,
        { text: rec.URL, hyperlink: rec.URL }
      ]);
      setHyperlinkStyle(row.getCell(6), 8, false);
    });
    [38, 18, 18, 24, 45, 90].forEach((w, i) => attachmentsWs.getColumn(i + 1).width = w);
    attachmentsWs.eachRow(row => row.eachCell({ includeEmpty: true }, cell => {
      if (!(cell.value && typeof cell.value === "object" && cell.value.hyperlink)) setBasicCellStyle(cell);
    }));
    attachmentsWs.getRow(1).eachCell(cell => setBasicCellStyle(cell, 9, true));
    for (let r = 2; r <= attachmentsWs.rowCount; r++) {
      attachmentsWs.getCell(r, 1).font = { name: "Arial", size: 8, bold: true, color: { argb: COLORS.black } };
    }
    attachmentsWs.getCell("H1").value = { text: "← BACK TO STATUS", hyperlink: "#'Status'!A1" };
    setHyperlinkStyle(attachmentsWs.getCell("H1"), 8, true);
    attachmentsWs.getColumn(8).width = 20;
    attachmentsWs.views = [{ state: "frozen", xSplit: 1, ySplit: 1, topLeftCell: "B2" }];
    attachmentsWs.autoFilter = { from: "A1", to: `F${Math.max(1, attachmentsWs.rowCount)}` };

    // Individual task sheets
    const usedNames = new Set(["status", "updates", "attachments"]);
    sortedItems.forEach((item, index) => {
      const statusRow = index + 2;
      const name = String(cleanText(item["Name"]));
      const tabName = safeSheetName(name, usedNames);
      const tab = wb.addWorksheet(tabName);
      setSheetTabColor(tab, COLORS.taskTabs[index % COLORS.taskTabs.length]);

      const statusNav = statusWs.getCell(statusRow, 2);
      statusNav.value = { text: "WORKSHEET", hyperlink: `#'${tabName.replace(/'/g, "''")}'!A1` };
      setHyperlinkStyle(statusNav, 6, true);

      const labels = [
        "Priority", "Priority / Urgency", "Status", "Due Date", "To-Do",
        "Clear Next Action", "Requested By", "Target Market", "Specs", "CTA",
        "Phone Number", "Email", "Request Type", "Objective",
        "Creative Brief / Source Notes", "Assets / References", "Checks / Blockers",
        "Assigned", "Monday URL", "SharePoint URL"
      ];

      tab.addRow(["Label", "Contents"]);
      labels.forEach(label => tab.addRow([label, ""]));

      tab.getCell("D1").value = { text: "← BACK TO STATUS", hyperlink: "#'Status'!A1" };
      setHyperlinkStyle(tab.getCell("D1"), 8, true);
      tab.getColumn(4).width = 20;

      const statusValue = normalizeStatus(item["Status"]);
      const dueValue = item["Due Date"] || "";
      const actionValue = defaultAction(item);
      const requesterValue = String(cleanText(item["Requested by"])) || String(cleanText(item["RMD Owner"]));
      const assignedValue = ownerToAssigned(item["Owner"]);

      tab.getCell("B2").value = { formula: `Status!A${statusRow}`, result: index + 1 };
      const urgency = statusValue === "On Hold" ? "HOLD" : statusValue === "Complete" ? "COMPLETE" : "ACTIVE";
      tab.getCell("B3").value = {
        formula: `IF(Status!D${statusRow}="On Hold","HOLD",IF(Status!D${statusRow}="Complete","COMPLETE","ACTIVE"))`,
        result: urgency
      };
      tab.getCell("B4").value = { formula: `Status!D${statusRow}`, result: statusValue };
      tab.getCell("B5").value = { formula: `Status!E${statusRow}`, result: dueValue };
      tab.getCell("B6").value = { formula: `Status!C${statusRow}`, result: name };
      tab.getCell("B7").value = { formula: `Status!F${statusRow}`, result: actionValue };
      tab.getCell("B8").value = { formula: `Status!G${statusRow}`, result: requesterValue };

      tab.getCell("B9").value = String(cleanText(item["Target Market"]));
      tab.getCell("B10").value = String(cleanText(item["Specs"]));
      tab.getCell("B11").value = String(cleanText(item["CTA"]));
      tab.getCell("B12").value = String(cleanText(item["Phone Number"]));
      tab.getCell("B13").value = String(cleanText(item["Email"]));
      tab.getCell("B14").value = String(cleanText(item["Single select"])) || String(cleanText(item["Type"]));
      tab.getCell("B15").value = String(cleanText(item["Multi select"]));

      let brief = String(cleanText(item["Creative Brief"]));
      if ((item._subitems || []).length) {
        brief += (brief ? "\n\nSubitems:\n" : "Subitems:\n") + item._subitems.map(x => `• ${x}`).join("\n");
      }
      tab.getCell("B16").value = brief;
      tab.getCell("B17").value = String(cleanText(item["Assets"]));

      const blockers = [];
      if (!String(cleanText(item["Specs"]))) blockers.push("Confirm final production specs.");
      if (!String(cleanText(item["CTA"]))) blockers.push("Confirm CTA.");
      if (!["Complete", "On Hold"].includes(statusValue) && !item["Due Date"]) blockers.push("Confirm due date.");
      tab.getCell("B18").value = blockers.join(" ");

      tab.getCell("B19").value = { formula: `Status!H${statusRow}`, result: assignedValue };
      tab.getCell("B20").value = { formula: `Status!I${statusRow}`, result: "" };
      tab.getCell("B21").value = { formula: `Status!J${statusRow}`, result: "" };

      tab.getCell("A22").value = "Monday Item ID";
      tab.getCell("B22").value = String(cleanText(item["Item ID (auto generated)"]));

      tab.getCell("A23").value = "Notes / Updates";
      tab.getCell("B23").value = { text: "View Updates/Notes", hyperlink: "#'Updates'!A1" };

      tab.getCell("A24").value = "Attachments";
      tab.getCell("B24").value = { text: "View Attachments", hyperlink: "#'Attachments'!A1" };

      // Styles: apply base first, then hyperlink styling so it stays visible.
      tab.eachRow(row => row.eachCell({ includeEmpty: true }, cell => setBasicCellStyle(cell)));
      tab.getRow(1).eachCell(cell => setBasicCellStyle(cell, 9, true));

      for (let r = 2; r <= 24; r++) {
        tab.getCell(r, 1).font = { name: "Arial", size: 8, bold: true, color: { argb: COLORS.black } };
        tab.getCell(r, 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.labelFill } };
      }
      setHyperlinkStyle(tab.getCell("D1"), 8, true);
      setHyperlinkStyle(tab.getCell("B23"), 8, false);
      setHyperlinkStyle(tab.getCell("B24"), 8, false);

      tab.getColumn(1).width = 26;
      tab.getColumn(2).width = 90;
      tab.views = [{ state: "frozen", ySplit: 1 }];
      tab.getCell("B5").numFmt = "mm/dd/yyyy";
    });

    return wb;
  }
})();
