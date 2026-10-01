const DEPARTMENTS = ["แผนกแรงสูง", "แผนกแรงสูง TAC", "แผนกหม้อแปลง", "แผนกสายส่ง"];
const DEFAULT_DEPARTMENT = DEPARTMENTS[0];

const state = {
  department: DEFAULT_DEPARTMENT,
  activeWorkType: "install",
  resultWorkType: "install",
  sizes: [],
  pages: [[blankRow(), blankRow()]],
  currentPage: 0,
  results: [],
  activeProjectId: "",
  cloudConfigured: false,
  googleCredential: sessionStorage.getItem("material-calculator-google-credential") || "",
  cloudProjects: [],
  cloudFolders: [],
  cloudUser: null,
  adminToken: sessionStorage.getItem("material-calculator-admin-token") || "",
  baseRequestOriginalRows: [],
  baseRequestImage: null,
  baseRequests: [],
  activeRequestStatus: "pending",
  insulatorRates: {},
  headImages: {},
  headDataCache: new Map(),
  departmentWork: {},
};

const SAVED_PROJECTS_KEY = "material-calculator-projects-v1";
const ADMIN_APPROVAL_REFRESH_KEY = "material-calculator-admin-approval-refresh";

const els = {
  status: document.getElementById("status"),
  sizeColumnTitle: document.getElementById("sizeColumnTitle"),
  inputTable: document.getElementById("inputTable"),
  headImageColumnTitle: document.getElementById("headImageColumnTitle"),
  headColumnTitle: document.getElementById("headColumnTitle"),
  countColumnTitle: document.getElementById("countColumnTitle"),
  departmentSelect: document.getElementById("departmentSelect"),
  totalPages: document.getElementById("totalPages"),
  demolitionPages: document.getElementById("demolitionPages"),
  installPageCount: document.getElementById("installPageCount"),
  demolitionPageCount: document.getElementById("demolitionPageCount"),
  applyPages: document.getElementById("applyPages"),
  prevPage: document.getElementById("prevPage"),
  nextPage: document.getElementById("nextPage"),
  pageLabel: document.getElementById("pageLabel"),
  pagePicker: document.getElementById("pagePicker"),
  addRow: document.getElementById("addRow"),
  removeRow: document.getElementById("removeRow"),
  clearPage: document.getElementById("clearPage"),
  inputRows: document.getElementById("inputRows"),
  calculate: document.getElementById("calculate"),
  expandSet: document.getElementById("expandSet"),
  exportExcel: document.getElementById("exportExcel"),
  exportPages: document.getElementById("exportPages"),
  exportPageHardware: document.getElementById("exportPageHardware"),
  exportPageCrossarms: document.getElementById("exportPageCrossarms"),
  baseRequestForm: document.getElementById("baseRequestForm"),
  requesterName: document.getElementById("requesterName"),
  requesterEmployeeId: document.getElementById("requesterEmployeeId"),
  requesterDepartment: document.getElementById("requesterDepartment"),
  requestTargetDepartment: document.getElementById("requestTargetDepartment"),
  requestAction: document.getElementById("requestAction"),
  requestSize: document.getElementById("requestSize"),
  requestHead: document.getElementById("requestHead"),
  requestSizeField: document.getElementById("requestSizeField"),
  requestHeadField: document.getElementById("requestHeadField"),
  requestSizeLabel: document.getElementById("requestSizeLabel"),
  requestHeadLabel: document.getElementById("requestHeadLabel"),
  requestInsulatorUpright: document.getElementById("requestInsulatorUpright"),
  requestInsulatorHorizontal: document.getElementById("requestInsulatorHorizontal"),
  requestImagePanel: document.getElementById("requestImagePanel"),
  requestImage: document.getElementById("requestImage"),
  requestImagePreview: document.getElementById("requestImagePreview"),
  removeRequestImage: document.getElementById("removeRequestImage"),
  requestAddTarget: document.getElementById("requestAddTarget"),
  requestReplaceTarget: document.getElementById("requestReplaceTarget"),
  replaceSize: document.getElementById("replaceSize"),
  replaceHead: document.getElementById("replaceHead"),
  existingDataHint: document.getElementById("existingDataHint"),
  requestMaterialRows: document.getElementById("requestMaterialRows"),
  addRequestMaterial: document.getElementById("addRequestMaterial"),
  requestNote: document.getElementById("requestNote"),
  adminLoginPanel: document.getElementById("adminLoginPanel"),
  adminRequestsPanel: document.getElementById("adminRequestsPanel"),
  adminLoginForm: document.getElementById("adminLoginForm"),
  adminUsername: document.getElementById("adminUsername"),
  adminPassword: document.getElementById("adminPassword"),
  adminLogout: document.getElementById("adminLogout"),
  connectGoogleDrive: document.getElementById("connectGoogleDrive"),
  driveConnectStatus: document.getElementById("driveConnectStatus"),
  baseRequestList: document.getElementById("baseRequestList"),
  pendingRequestCount: document.getElementById("pendingRequestCount"),
  approvedRequestCount: document.getElementById("approvedRequestCount"),
  rejectedRequestCount: document.getElementById("rejectedRequestCount"),
  clearApprovedRequests: document.getElementById("clearApprovedRequests"),
  installResultSection: document.getElementById("installResultSection"),
  demolitionResultSection: document.getElementById("demolitionResultSection"),
  installResultRows: document.getElementById("installResultRows"),
  demolitionResultRows: document.getElementById("demolitionResultRows"),
  installResultCount: document.getElementById("installResultCount"),
  demolitionResultCount: document.getElementById("demolitionResultCount"),
  resultMeta: document.getElementById("resultMeta"),
  projectName: document.getElementById("projectName"),
  planNumber: document.getElementById("planNumber"),
  saveLocalProject: document.getElementById("saveLocalProject"),
  saveCloudProject: document.getElementById("saveCloudProject"),
  cloudSaveDialog: document.getElementById("cloudSaveDialog"),
  cloudSaveForm: document.getElementById("cloudSaveForm"),
  cloudSaveFolder: document.getElementById("cloudSaveFolder"),
  cancelCloudSave: document.getElementById("cancelCloudSave"),
  newProject: document.getElementById("newProject"),
  saveHint: document.getElementById("saveHint"),
  savedCount: document.getElementById("savedCount"),
  localSavedProjectList: document.getElementById("localSavedProjectList"),
  cloudSavedProjectList: document.getElementById("cloudSavedProjectList"),
  localSavedCount: document.getElementById("localSavedCount"),
  cloudSavedCount: document.getElementById("cloudSavedCount"),
  cloudFolderTitle: document.getElementById("cloudFolderTitle"),
  createCloudFolder: document.getElementById("createCloudFolder"),
  savedLocationText: document.getElementById("savedLocationText"),
  googleSignIn: document.getElementById("googleSignIn"),
  googleSignInQuick: document.getElementById("googleSignInQuick"),
  cloudUser: document.getElementById("cloudUser"),
  cloudUserName: document.getElementById("cloudUserName"),
  cloudQuickUser: document.getElementById("cloudQuickUser"),
  googleSignOut: document.getElementById("googleSignOut"),
  cloudNotice: document.getElementById("cloudNotice"),
  headImageDialog: document.getElementById("headImageDialog"),
  headImageDialogTitle: document.getElementById("headImageDialogTitle"),
  headImageDialogImage: document.getElementById("headImageDialogImage"),
  closeHeadImageDialog: document.getElementById("closeHeadImageDialog"),
};

function blankRow(department = DEFAULT_DEPARTMENT, workType = "install") {
  return { department, workType, size: "", head: "", count: "", wire1: "", wire2: "", wire3: "", latWire: "", surgeNgr: false, surgeWithin3km: false, surgeMounting: "crossarm", insulatorUpright: null, insulatorHorizontal: null };
}

function pageWorkType(page) {
  return page?.[0]?.workType === "demolition" ? "demolition" : "install";
}

function workTypePageIndices(workType = state.activeWorkType) {
  return state.pages.map((page, index) => ({ page, index })).filter((entry) => pageWorkType(entry.page) === workType).map((entry) => entry.index);
}

function workTypeLabel(workType) {
  return workType === "demolition" ? "งานรื้อถอน" : "งานติดตั้ง";
}

function insulatorRateKey(department, size, head) {
  return `${department}\u0000${size}\u0000${head}`;
}

function headImageKey(department, size, head) {
  return `${department}\u0000${size}\u0000${head}`;
}

function imageUrl(imageId) {
  return `/api/head-image?id=${encodeURIComponent(imageId)}`;
}

function openHeadImage(image, title) {
  if (!image?.id) return;
  els.headImageDialogTitle.textContent = title || image.name || "รูปประกอบหัวเสา";
  els.headImageDialogImage.src = imageUrl(image.id);
  els.headImageDialog.showModal();
}

function setStatus(message, isError = false) {
  els.status.textContent = message;
  els.status.classList.toggle("error", isError);
}

async function postJson(endpoint, payload = {}) {
  const response = await fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  return readJson(response);
}

async function readJson(response) {
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.error || "เกิดข้อผิดพลาด");
  }
  return data;
}

function saveCurrentPageFromDom() {
  const rows = [...els.inputRows.querySelectorAll("tr.input-row")];
  rows.forEach((tr, index) => {
    const headInput = tr.querySelector(".head");
    const typedHead = headInput.value.trim();
    const validHeads = headInput._allHeadOptions || [];
    const savedHead = !typedHead || validHeads.includes(typedHead)
      ? typedHead
      : state.pages[state.currentPage][index].head;
    Object.assign(state.pages[state.currentPage][index], {
      department: state.department,
      workType: state.activeWorkType,
      size: tr.querySelector(".size").value,
      head: savedHead,
      count: tr.querySelector(".count").value,
    });
  });
}

function renderInputs() {
  updateInputColumnTitles();
  els.inputRows.innerHTML = "";
  const page = state.pages[state.currentPage];
  page.forEach((row, index) => {
    const tr = document.createElement("tr");
    tr.className = "input-row";
    tr.innerHTML = `
      ${state.department === "แผนกแรงสูง TAC" ? '<td class="head-image-cell"></td>' : ""}
      <td class="size-cell"><select class="size"></select></td>
      <td class="head-cell"><div class="head-combobox"><div class="head-combobox-control"><input class="head searchable-dropdown" type="text" autocomplete="off" role="combobox" aria-expanded="false"><button class="head-toggle" type="button" aria-label="เปิดรายการหัวเสา">⌄</button></div><div class="head-options" role="listbox" hidden></div></div></td>
      <td class="count-cell"><input class="count" type="text" inputmode="text" placeholder="เช่น 4+4+5+6"></td>
    `;

    const sizeSelect = tr.querySelector(".size");
    const headSelect = tr.querySelector(".head");
    const headCombobox = tr.querySelector(".head-combobox");
    const headToggle = tr.querySelector(".head-toggle");
    const headOptions = tr.querySelector(".head-options");
    const countInput = tr.querySelector(".count");
    updateHeadImageCell(tr, row);

    fillSelect(sizeSelect, state.sizes, state.department === "แผนกหม้อแปลง" ? "เลือกแผนก" : "เลือกขนาดเสา");
    sizeSelect.value = row.size || "";
    countInput.value = row.count || "";
    headSelect.placeholder = state.department === "แผนกหม้อแปลง" ? "เลือกหรือพิมพ์ค้นหาชนิดหม้อแปลง" : "เลือกหรือพิมพ์ค้นหาหัวเสา";
    headSelect.value = row.head || "";
    const positionHeadOptions = () => {
      if (headOptions.hidden) return;
      const bounds = headCombobox.getBoundingClientRect();
      const spaceBelow = window.innerHeight - bounds.bottom;
      const spaceAbove = bounds.top;
      const naturalHeight = Math.min(headOptions.scrollHeight, 400);
      const showBelow = spaceBelow >= naturalHeight + 4 || spaceBelow >= spaceAbove;
      const availableHeight = Math.max(120, Math.min(400, (showBelow ? spaceBelow : spaceAbove) - 12));
      headOptions.style.maxHeight = `${availableHeight}px`;
      const renderedHeight = Math.min(headOptions.scrollHeight, availableHeight);
      const top = showBelow ? bounds.bottom + 4 : Math.max(8, bounds.top - renderedHeight - 4);
      headOptions.style.left = `${bounds.left}px`;
      headOptions.style.top = `${top}px`;
      headOptions.style.width = `${bounds.width}px`;
    };
    const closeHeadOptions = () => {
      headOptions.hidden = true;
      headSelect.setAttribute("aria-expanded", "false");
      window.removeEventListener("scroll", positionHeadOptions, true);
      window.removeEventListener("resize", positionHeadOptions);
    };
    const openHeadOptions = () => {
      renderHeadOptions(headSelect, headOptions);
      headOptions.hidden = false;
      headSelect.setAttribute("aria-expanded", "true");
      positionHeadOptions();
      window.addEventListener("scroll", positionHeadOptions, true);
      window.addEventListener("resize", positionHeadOptions);
    };

    sizeSelect.addEventListener("change", () => {
      page[index].department = state.department;
      page[index].size = sizeSelect.value;
      page[index].head = "";
      page[index].wire1 = "";
      page[index].wire2 = "";
      page[index].wire3 = "";
      page[index].latWire = "";
      renderInputs();
    });
    const applyHeadSelection = () => {
      if (headSelect.value && !(headSelect._allHeadOptions || []).includes(headSelect.value)) {
        setStatus("กรุณาเลือกหัวเสาจากรายการที่ค้นหา", true);
        headSelect.value = row.head || "";
        return;
      }
      if (page[index].head === headSelect.value) return;
      page[index].head = headSelect.value;
      const rate = state.insulatorRates[insulatorRateKey(state.department, page[index].size, headSelect.value)];
      page[index].insulatorUpright = rate ? Number(rate[0]) : null;
      page[index].insulatorHorizontal = rate ? Number(rate[1]) : null;
      page[index].wire1 = "";
      page[index].wire2 = "";
      page[index].wire3 = "";
      page[index].latWire = "";
      renderInputs();
    };
    headSelect.addEventListener("input", () => {
      openHeadOptions();
      if ((headSelect._allHeadOptions || []).includes(headSelect.value)) applyHeadSelection();
    });
    headSelect.addEventListener("change", applyHeadSelection);
    headSelect.addEventListener("focus", openHeadOptions);
    headToggle.addEventListener("mousedown", (event) => event.preventDefault());
    headToggle.addEventListener("click", () => {
      if (headOptions.hidden) {
        headSelect.focus();
        openHeadOptions();
      } else {
        closeHeadOptions();
      }
    });
    headOptions.addEventListener("mousedown", (event) => {
      const option = event.target.closest("button[data-head]");
      if (!option) return;
      event.preventDefault();
      headSelect.value = option.dataset.head;
      closeHeadOptions();
      applyHeadSelection();
    });
    headCombobox.addEventListener("focusout", () => {
      setTimeout(() => {
        if (!headCombobox.contains(document.activeElement)) closeHeadOptions();
      }, 0);
    });
    countInput.addEventListener("input", () => {
      page[index].count = countInput.value;
      renderInsulators();
    });

    els.inputRows.appendChild(tr);
    const wireKind = classifyWireHead(row.head);
    if (wireKind) {
      els.inputRows.appendChild(createWireDetailsRow(row, index, wireKind));
    }
    if ((row.department || state.department) === "แผนกหม้อแปลง" && row.head) {
      els.inputRows.appendChild(createSurgeDetailsRow(row, index));
    }
    if (row.size) {
      loadHeads(row.size, headSelect, row.head, row.department || state.department, headOptions).then((data) => {
        if (!data || !row.head) return;
        updateHeadImageCell(tr, row);
        const rate = data.insulatorRates?.[row.head];
        if (rate && (row.insulatorUpright === null || row.insulatorUpright === undefined)) {
          row.insulatorUpright = Number(rate[0]);
          row.insulatorHorizontal = Number(rate[1]);
          renderInsulators();
        }
      });
    }
  });
  renderPageControls();
  renderInsulators();
}

function updateHeadImageCell(rowElement, row) {
  const cell = rowElement.querySelector(".head-image-cell");
  if (!cell) return;
  cell.replaceChildren();
  const image = state.headImages[headImageKey(row.department || state.department, row.size, row.head)];
  if (!image?.id) {
    const empty = document.createElement("span");
    empty.className = "head-image-empty";
    empty.textContent = "ไม่มีรูป";
    cell.appendChild(empty);
    return;
  }
  const button = document.createElement("button");
  button.type = "button";
  button.className = "head-image-thumbnail";
  button.title = "คลิกเพื่อดูภาพขยาย";
  const preview = document.createElement("img");
  preview.src = imageUrl(image.id);
  preview.alt = `รูปประกอบ ${row.head}`;
  button.appendChild(preview);
  button.addEventListener("click", () => openHeadImage(image, row.head));
  cell.appendChild(button);
}

function updateInputColumnTitles() {
  const transformer = state.department === "แผนกหม้อแปลง";
  const tac = state.department === "แผนกแรงสูง TAC";
  els.headImageColumnTitle.hidden = !tac;
  els.inputTable.classList.toggle("has-head-images", tac);
  els.sizeColumnTitle.textContent = transformer ? "แผนก" : "ขนาดเสา (m)";
  els.headColumnTitle.textContent = transformer ? "ชนิดหม้อแปลง" : "รหัสหัวเสา";
  els.countColumnTitle.textContent = "จำนวน";
}

function createSurgeDetailsRow(row, index) {
  const detailRow = document.createElement("tr");
  detailRow.className = "surge-details-row";
  const cell = document.createElement("td");
  cell.colSpan = state.department === "แผนกแรงสูง TAC" ? 4 : 3;
  const panel = document.createElement("div");
  panel.className = "surge-details";
  panel.innerHTML = `
    <div class="surge-title"><strong>เลือก Surge Arrester</strong><span class="surge-preview"></span></div>
    <label><span>ระบบ</span><select class="surge-ngr"><option value="normal">ระบบปกติ</option><option value="ngr">ระบบ NGR</option></select></label>
    <label><span>ระยะจากสถานี</span><select class="surge-distance"><option value="outside">นอกระยะ 3 กม.</option><option value="inside">ภายในระยะ 3 กม.</option></select></label>
    <label><span>ตำแหน่งติดตั้ง</span><select class="surge-mounting"><option value="crossarm">ติดตั้งบนคอน</option><option value="tank">ติดตั้งตัวถังหม้อแปลง</option></select></label>
  `;
  const ngrSelect = panel.querySelector(".surge-ngr");
  const distanceSelect = panel.querySelector(".surge-distance");
  const mountingSelect = panel.querySelector(".surge-mounting");
  const preview = panel.querySelector(".surge-preview");
  ngrSelect.value = row.surgeNgr ? "ngr" : "normal";
  distanceSelect.value = row.surgeWithin3km ? "inside" : "outside";
  mountingSelect.value = row.surgeMounting || "crossarm";
  const update = () => {
    row.surgeNgr = ngrSelect.value === "ngr";
    row.surgeWithin3km = distanceSelect.value === "inside";
    if (row.surgeWithin3km) {
      mountingSelect.value = "crossarm";
      mountingSelect.querySelector('option[value="tank"]').disabled = true;
    } else {
      mountingSelect.querySelector('option[value="tank"]').disabled = false;
    }
    row.surgeMounting = mountingSelect.value;
    const crossarmCodes = row.surgeNgr
      ? (row.surgeWithin3km ? "1040000003" : "1040000002")
      : (row.surgeWithin3km ? "1040000001" : "1040000000");
    const tankCode = row.surgeNgr ? "1040000008" : "1040000007";
    preview.textContent = row.surgeMounting === "tank"
      ? `เลือกติดตัวถัง: ${tankCode} · 5 kA WITHOUT BRACKET`
      : `เลือกติดตั้งบนคอน: ${crossarmCodes}`;
  };
  ngrSelect.addEventListener("change", update);
  distanceSelect.addEventListener("change", update);
  mountingSelect.addEventListener("change", update);
  update();
  cell.appendChild(panel);
  detailRow.appendChild(cell);
  return detailRow;
}

function renderInsulators() {
  const totals = summarizeInsulatorsByWorkType(state.pages);
  document.getElementById("installUprightCount").textContent = formatAmount(totals.install.upright);
  document.getElementById("installHorizontalCount").textContent = formatAmount(totals.install.horizontal);
  document.getElementById("demolitionUprightCount").textContent = formatAmount(totals.demolition.upright);
  document.getElementById("demolitionHorizontalCount").textContent = formatAmount(totals.demolition.horizontal);
  const warnings = document.getElementById("insulatorWarnings");
  warnings.hidden = totals.warnings.length === 0;
  warnings.querySelector("summary").textContent = `รายการที่ยังไม่รวมในยอด (${totals.warnings.length} แถว)`;
  const list = document.getElementById("insulatorWarningList");
  list.replaceChildren();
  totals.warnings.forEach(message => {
    const item = document.createElement("li");
    item.textContent = message;
    list.appendChild(item);
  });
}

const wireOptions = [
  "50 PIC",
  "95 PIC",
  "185 PIC",
  "50 SAC",
  "185 SAC",
  "50 A",
  "50 ACSR",
  "185 ACSR",
  "185 A",
];

function classifyWireHead(head) {
  let normalized = String(head || "").trim().toUpperCase();
  const compact = normalized.replace(/\s+/g, "");
  const combinedRules = {
    "DDE,DE(ST.4.5M)": "dde_de",
    "SP,DDE.BLST.4.5M": "dde_bl",
    "DP,DDE.BLST.4.5M": "dde_bl",
    "DP,DDEST.4.5M": "dde",
    "DP,DEST.4.5M": "de",
  };
  if (combinedRules[compact]) return combinedRules[compact];
  if (/^LAT\.SLK(?=$|\s)/.test(normalized)) return "de";
  if (normalized.startsWith("2")) {
    if (normalized.includes("+")) return "";
    normalized = normalized.slice(1);
  }
  if (normalized.startsWith("DDE.BL")) return "dde_bl";
  if (normalized.startsWith("DDE")) return "dde";
  if (normalized.startsWith("DE")) return "de";
  if (normalized.startsWith("BA")) return "ba";
  return "";
}

function createWireDetailsRow(row, index, wireKind) {
  const detailRow = document.createElement("tr");
  detailRow.className = "wire-details-row";
  const cell = document.createElement("td");
  cell.colSpan = state.department === "แผนกแรงสูง TAC" ? 4 : 3;
  const panel = document.createElement("div");
  panel.className = "wire-details";

  const labels = wireKind === "de"
    ? [["สาย Dead End", "wire1"]]
    : wireKind === "ba"
      ? [["Main Line", "wire1"], ["Tap Line", "wire2"]]
      : wireKind === "dde_de"
        ? [["สาย DDE ด้านซ้าย", "wire1"], ["สาย DDE ด้านขวา", "wire2"], ["สาย DE", "wire3"]]
        : [["สายด้านซ้าย", "wire1"], ["สายด้านขวา", "wire2"]];
  if (String(row.head || "").toUpperCase().replace(/\s+/g, "") === "DDE.ST3M,LAT.SLK") {
    labels.push(["สาย LAT.SLK (ยึดแบบ DE)", "latWire"]);
  }

  labels.forEach(([labelText, key]) => {
    const label = document.createElement("label");
    const caption = document.createElement("span");
    caption.textContent = labelText;
    const select = document.createElement("select");
    select.className = "wire-select";
    fillSelect(select, wireOptions, "เลือกชนิดสาย");
    if (wireKind === "ba" && key === "wire1" && !row[key]) {
      row[key] = "185 SAC";
    }
    select.value = row[key] || "";
    select.addEventListener("change", () => {
      state.pages[state.currentPage][index][key] = select.value;
    });
    label.append(caption, select);
    panel.appendChild(label);
  });

  cell.appendChild(panel);
  detailRow.appendChild(cell);
  return detailRow;
}

function fillSelect(select, values, placeholder) {
  select.innerHTML = "";
  const empty = document.createElement("option");
  empty.value = "";
  empty.textContent = placeholder;
  select.appendChild(empty);
  values.forEach((value) => {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = value;
    select.appendChild(opt);
  });
}

function renderHeadOptions(input, optionsList) {
  const query = input.value.trim().toLocaleLowerCase("th");
  const matches = (input._allHeadOptions || []).filter((head) =>
    !query || String(head).toLocaleLowerCase("th").includes(query)
  );
  if (!matches.length) {
    const empty = document.createElement("div");
    empty.className = "head-option-empty";
    empty.textContent = "ไม่พบหัวเสาที่ค้นหา";
    optionsList.replaceChildren(empty);
    return;
  }
  optionsList.replaceChildren(...matches.map((head) => {
    const option = document.createElement("button");
    option.type = "button";
    option.className = "head-option";
    option.dataset.head = head;
    option.setAttribute("role", "option");
    option.textContent = head;
    return option;
  }));
}

async function loadHeads(size, select, selected, department = state.department, optionsList = null) {
  if (!size) {
    select.value = "";
    select._allHeadOptions = [];
    optionsList?.replaceChildren();
    return null;
  }
  try {
    const cacheKey = `${department}\u0000${size}`;
    let cached = state.headDataCache.get(cacheKey);
    if (!cached) {
      cached = fetch(`/api/heads?size=${encodeURIComponent(size)}&department=${encodeURIComponent(department)}`)
        .then(readJson)
        .catch((error) => {
          state.headDataCache.delete(cacheKey);
          throw error;
        });
      state.headDataCache.set(cacheKey, cached);
    }
    const data = await cached;
    state.headDataCache.set(cacheKey, data);
    Object.entries(data.insulatorRates || {}).forEach(([head, rate]) => {
      state.insulatorRates[insulatorRateKey(department, size, head)] = rate;
    });
    Object.entries(data.headImages || {}).forEach(([head, image]) => {
      state.headImages[headImageKey(department, size, head)] = image;
    });
    select._allHeadOptions = data.heads;
    if (optionsList) {
      renderHeadOptions(select, optionsList);
    }
    select.value = selected || "";
    return data;
  } catch (error) {
    setStatus(error.message, true);
  }
}

function renderPageControls() {
  const installIndices = workTypePageIndices("install");
  const demolitionIndices = workTypePageIndices("demolition");
  const activeIndices = state.activeWorkType === "demolition" ? demolitionIndices : installIndices;
  let localPageIndex = activeIndices.indexOf(state.currentPage);
  if (localPageIndex < 0 && activeIndices.length) {
    state.currentPage = activeIndices[0];
    localPageIndex = 0;
  }
  els.totalPages.value = String(installIndices.length);
  els.demolitionPages.value = String(demolitionIndices.length);
  els.installPageCount.textContent = String(installIndices.length);
  els.demolitionPageCount.textContent = String(demolitionIndices.length);
  document.querySelectorAll(".work-type-button").forEach((button) => {
    button.classList.toggle("active", button.dataset.workType === state.activeWorkType);
  });
  els.pageLabel.textContent = `${workTypeLabel(state.activeWorkType)} หน้า ${localPageIndex + 1}/${activeIndices.length}`;
  els.prevPage.disabled = localPageIndex <= 0;
  els.nextPage.disabled = localPageIndex < 0 || localPageIndex >= activeIndices.length - 1;
  els.pagePicker.innerHTML = "";
  activeIndices.forEach((index, localIndex) => {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = String(localIndex + 1);
    button.className = "page-number";
    button.classList.toggle("active", index === state.currentPage);
    button.setAttribute("aria-label", `${workTypeLabel(state.activeWorkType)} หน้าที่ ${localIndex + 1}`);
    button.setAttribute("aria-current", index === state.currentPage ? "page" : "false");
    button.addEventListener("click", () => {
      if (index === state.currentPage) return;
      saveCurrentPageFromDom();
      state.currentPage = index;
      renderInputs();
    });
    els.pagePicker.appendChild(button);
  });
}

function renderResults(items, meta) {
  state.results = items || [];
  const installItems = state.results.filter((item) => Number(item["จำนวนติดตั้ง"] ?? item["จำนวนรวม"] ?? 0) !== 0);
  const demolitionItems = state.results.filter((item) => Number(item["จำนวนรื้อถอน"] ?? 0) !== 0);
  const renderWorkRows = (container, rows, quantityKey, fallbackKey = null) => {
    container.innerHTML = rows.map((item) => `
      <tr>
        <td>${escapeHtml(item["รายการวัสดุ"] || "")}</td>
        <td>${String(item["รหัสพัสดุ"] || "").trim().toLowerCase().startsWith("set") ? `<button type="button" class="set-expand-button" data-set-code="${escapeHtml(item["รหัสพัสดุ"] || "")}" data-set-quantity="${escapeHtml(String(item[quantityKey] ?? (fallbackKey ? item[fallbackKey] : 0)))}" aria-label="ดูไส้ใน ${escapeHtml(item["รหัสพัสดุ"] || "")}">+</button>` : ""}<span class="material-code">${escapeHtml(item["รหัสพัสดุ"] || "")}</span></td>
        <td>${formatAmount(item[quantityKey] ?? (fallbackKey ? item[fallbackKey] : 0))}</td>
      </tr>
    `).join("");
  };
  renderWorkRows(els.installResultRows, installItems, "จำนวนติดตั้ง", "จำนวนรวม");
  renderWorkRows(els.demolitionResultRows, demolitionItems, "จำนวนรื้อถอน");
  els.installResultCount.textContent = `${installItems.length} รายการ`;
  els.demolitionResultCount.textContent = `${demolitionItems.length} รายการ`;
  updateResultTypeView();
  els.resultMeta.textContent = meta || `ทั้งหมด ${state.results.length} รายการ`;
}

async function toggleSetComponents(event) {
  const button = event.target.closest(".set-expand-button");
  if (!button) return;
  const parentRow = button.closest("tr");
  const existing = parentRow.nextElementSibling;
  if (existing?.classList.contains("set-component-row")) {
    existing.remove();
    button.textContent = "+";
    button.setAttribute("aria-expanded", "false");
    return;
  }
  try {
    button.disabled = true;
    button.textContent = "…";
    const data = await postJson("/api/set-components", {
      code: button.dataset.setCode,
      quantity: button.dataset.setQuantity,
    });
    const detailRow = document.createElement("tr");
    detailRow.className = "set-component-row";
    const rows = data.items || [];
    detailRow.innerHTML = `<td colspan="3"><div class="set-component-panel">
      <div class="set-component-title">รายการประมาณการภายใน ${escapeHtml(data.setCode || button.dataset.setCode)}</div>
      ${rows.length ? `<table><thead><tr><th>รายการวัสดุ</th><th>รหัสพัสดุ 10 หลัก</th><th>จำนวน</th></tr></thead><tbody>${rows.map((item) => `<tr><td>${escapeHtml(item.material || "")}</td><td>${escapeHtml(item.code || "")}</td><td>${formatAmount(item.quantity)}</td></tr>`).join("")}</tbody></table>` : '<div class="set-component-empty">ไม่พบรายการรหัส 10 หลักใน SET นี้</div>'}
    </div></td>`;
    parentRow.after(detailRow);
    button.textContent = "−";
    button.setAttribute("aria-expanded", "true");
  } catch (error) {
    button.textContent = "+";
    setStatus(error.message, true);
  } finally {
    button.disabled = false;
  }
}

els.installResultRows.addEventListener("click", toggleSetComponents);
els.demolitionResultRows.addEventListener("click", toggleSetComponents);

function updateResultTypeView() {
  const isInstall = state.resultWorkType === "install";
  els.installResultSection.hidden = !isInstall || els.installResultRows.children.length === 0;
  els.demolitionResultSection.hidden = isInstall || els.demolitionResultRows.children.length === 0;
}

function formatAmount(value) {
  const number = Number(value || 0);
  return Number.isInteger(number) ? String(number) : number.toFixed(3).replace(/0+$/, "").replace(/\.$/, "");
}

function exportFileName(topic) {
  const planNumber = els.planNumber.value.trim();
  const prefix = planNumber ? `${planNumber}_` : "";
  return `${prefix}${topic}.xlsx`.replace(/[\\/:*?"<>|]+/g, "-").replace(/\s+/g, " ").trim();
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function clonePages(pages) {
  return pages.map((page) => {
    const workType = pageWorkType(page);
    return page.map((row) => ({ ...blankRow(state.department, workType), ...row, workType: row.workType === "demolition" ? "demolition" : workType }));
  });
}

function emptyDepartmentWork(department) {
  return { pages: [[blankRow(department, "install"), blankRow(department, "install")]], currentPage: 0, activeWorkType: "install", results: [], resultMeta: "ยังไม่มีผลคำนวณ" };
}

function stashCurrentDepartment() {
  saveCurrentPageFromDom();
  state.departmentWork[state.department] = {
    pages: clonePages(state.pages), currentPage: state.currentPage, activeWorkType: state.activeWorkType,
    results: structuredClone(state.results), resultMeta: els.resultMeta.textContent || "ยังไม่มีผลคำนวณ",
  };
}

function projectDepartmentCount(project) {
  return Object.keys(project.departments || {}).length || 1;
}

function projectPageCount(project) {
  if (project.departments) return Object.values(project.departments).reduce((sum, work) => sum + Number(work.pages?.length || 0), 0);
  return Number(project.pages?.length || 0);
}

function getSavedProjects() {
  try {
    const value = JSON.parse(localStorage.getItem(SAVED_PROJECTS_KEY) || "[]");
    return Array.isArray(value) ? value : [];
  } catch {
    return [];
  }
}

function writeSavedProjects(projects) {
  localStorage.setItem(SAVED_PROJECTS_KEY, JSON.stringify(projects));
  renderSavedProjects();
}

async function cloudRequest(endpoint, options = {}) {
  const headers = { ...(options.headers || {}), Authorization: `Bearer ${state.googleCredential}` };
  const response = await fetch(endpoint, { ...options, headers });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "เชื่อมต่อ Google Sheet ไม่สำเร็จ");
  return data;
}

function formatSavedDate(value) {
  if (!value) return "";
  return new Intl.DateTimeFormat("th-TH", { dateStyle: "medium", timeStyle: "short" }).format(new Date(value));
}

function switchTab(tabName) {
  document.querySelectorAll(".app-tab").forEach((button) => {
    button.classList.toggle("active", button.dataset.tab === tabName);
  });
  document.querySelectorAll(".tab-panel").forEach((panel) => {
    panel.hidden = panel.dataset.panel !== tabName;
  });
  if (tabName === "saved") renderSavedProjects();
  if (tabName === "base-admin") showAdminPanel();
}

function renderSavedProjects() {
  const localProjects = getSavedProjects().sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
  const cloudProjects = [...state.cloudProjects].sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
  els.savedCount.textContent = String(localProjects.length + cloudProjects.length);
  els.localSavedCount.textContent = `${localProjects.length} งาน`;
  els.cloudSavedCount.textContent = state.cloudUser ? `${cloudProjects.length} งาน` : "ยังไม่เข้าสู่ระบบ";
  els.cloudFolderTitle.textContent = state.cloudUser
    ? `โฟลเดอร์ Cloud ของ ${state.cloudUser.email}`
    : "โฟลเดอร์ Cloud ส่วนตัว";
  els.savedLocationText.textContent = "แสดงงานในเครื่องนี้และงานบน Google Sheet แยกจากกัน";
  renderSavedProjectList(els.localSavedProjectList, localProjects, "local");
  if (state.cloudUser) {
    renderCloudProjectFolders(cloudProjects);
  } else {
    els.cloudSavedProjectList.innerHTML = '<div class="empty-saved">เข้าสู่ระบบ Google เพื่อดูและเปิดงานบน Cloud</div>';
  }
}

function savedProjectCardsMarkup(projects, source) {
  return projects.map((project) => `
    <article class="saved-card">
      <div class="saved-info">
        <h3>${escapeHtml(project.name)}</h3>
        <div class="saved-meta">
          <span>${projectDepartmentCount(project)} แผนก</span>
          <span>เลขผัง: ${escapeHtml(project.planNumber || "-")}</span>
          <span>${projectPageCount(project)} หน้ารวม</span>
          <span>แก้ไขล่าสุด ${escapeHtml(formatSavedDate(project.updatedAt))}</span>
        </div>
      </div>
      <div class="saved-actions">
        <span class="saved-source-badge ${source}">${source === "cloud" ? "Cloud" : "เครื่องนี้"}</span>
        ${source === "cloud" ? `<label class="move-project-label">ย้ายไป<select data-action="move" data-source="cloud" data-project-id="${escapeHtml(project.id)}">
          <option value=""${project.folderId ? "" : " selected"}>งานทั่วไป</option>
          ${state.cloudFolders.map((folder) => `<option value="${escapeHtml(folder.id)}"${project.folderId === folder.id ? " selected" : ""}>${escapeHtml(folder.name)}</option>`).join("")}
        </select></label>` : ""}
        <button type="button" data-action="open" data-source="${source}" data-project-id="${escapeHtml(project.id)}">เปิดงาน</button>
        <button type="button" class="danger" data-action="delete" data-source="${source}" data-project-id="${escapeHtml(project.id)}">ลบ</button>
      </div>
    </article>
  `).join("");
}

function renderSavedProjectList(container, projects, source) {
  if (!projects.length) {
    container.innerHTML = '<div class="empty-saved">ยังไม่มีงานที่บันทึกไว้</div>';
    return;
  }
  container.innerHTML = savedProjectCardsMarkup(projects, source);
}

function renderCloudProjectFolders(projects) {
  const folders = [{ id: "", name: "งานทั่วไป" }, ...state.cloudFolders];
  els.cloudSavedProjectList.innerHTML = folders.map((folder) => {
    const folderProjects = projects.filter((project) => String(project.folderId || "") === folder.id);
    return `<details class="cloud-project-folder">
      <summary class="cloud-project-folder-title">
        <div><span class="folder-chevron">›</span><span class="folder-icon">▰</span><strong>${escapeHtml(folder.name)}</strong><small>${folderProjects.length} งาน</small></div>
        ${folder.id ? `<button type="button" class="folder-delete" data-action="delete-folder" data-folder-id="${escapeHtml(folder.id)}">ลบโฟลเดอร์</button>` : ""}
      </summary>
      <div class="cloud-folder-projects">${folderProjects.length ? savedProjectCardsMarkup(folderProjects, "cloud") : '<div class="empty-saved compact">ยังไม่มีงานในโฟลเดอร์นี้</div>'}</div>
    </details>`;
  }).join("");
}

function resetProject() {
  state.activeProjectId = "";
  state.departmentWork = {};
  state.activeWorkType = "install";
  state.resultWorkType = "install";
  state.pages = [[blankRow(state.department, "install"), blankRow(state.department, "install")]];
  state.currentPage = 0;
  state.results = [];
  els.projectName.value = "";
  els.planNumber.value = "";
  els.saveHint.textContent = "ยังไม่ได้บันทึกงานนี้";
  renderInputs();
  renderResults([], "ยังไม่มีข้อมูล");
  switchTab("calculator");
  setStatus("สร้างงานใหม่แล้ว");
}

async function openSavedProject(projectId, source = "local") {
  const projects = source === "cloud" ? state.cloudProjects : getSavedProjects();
  const project = projects.find((item) => item.id === projectId);
  if (!project) {
    setStatus("ไม่พบงานที่บันทึกไว้", true);
    renderSavedProjects();
    return;
  }
  state.activeProjectId = project.id;
  state.department = project.department || project.pages?.[0]?.[0]?.department || DEFAULT_DEPARTMENT;
  state.departmentWork = project.departments ? structuredClone(project.departments) : {
    [state.department]: {
      pages: project.pages?.length ? structuredClone(project.pages) : [[blankRow(state.department), blankRow(state.department)]],
      currentPage: Number(project.currentPage || 0), activeWorkType: project.activeWorkType || "install", results: structuredClone(project.results || []),
      resultMeta: project.resultMeta || "ยังไม่มีผลคำนวณ",
    },
  };
  els.departmentSelect.value = state.department;
  await loadDepartmentSizes(state.department);
  const work = state.departmentWork[state.department] || emptyDepartmentWork(state.department);
  state.pages = clonePages(work.pages);
  state.pages.forEach((page) => page.forEach((row) => { row.department = state.department; }));
  state.currentPage = Math.min(Number(work.currentPage || 0), state.pages.length - 1);
  state.activeWorkType = work.activeWorkType || pageWorkType(state.pages[state.currentPage]);
  state.resultWorkType = state.activeWorkType;
  state.results = structuredClone(work.results || []);
  els.projectName.value = project.name || "";
  els.planNumber.value = project.planNumber || "";
  els.saveHint.textContent = `เปิดงานที่บันทึกเมื่อ ${formatSavedDate(project.updatedAt)}`;
  renderInputs();
  renderResults(state.results, work.resultMeta || "ยังไม่มีผลคำนวณ");
  switchTab("calculator");
  setStatus("เปิดงานเดิมสำเร็จ");
}

async function saveProject(destination, cloudFolderId = null) {
  if (destination === "cloud" && !state.cloudUser) {
    setStatus("กรุณาเข้าสู่ระบบ Google ก่อนบันทึกขึ้น Cloud", true);
    els.googleSignInQuick.scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }
  stashCurrentDepartment();
  const name = els.projectName.value.trim();
  if (!name) {
    els.projectName.focus();
    setStatus("กรุณาใส่ชื่องานก่อนบันทึก", true);
    return;
  }
  const projects = destination === "cloud" ? state.cloudProjects : getSavedProjects();
  const now = new Date().toISOString();
  const id = state.activeProjectId || `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
  const existingProject = projects.find((item) => item.id === id);
  const project = {
    id,
    name,
    planNumber: els.planNumber.value.trim(),
    department: state.department,
    pages: clonePages(state.pages),
    currentPage: state.currentPage,
    activeWorkType: state.activeWorkType,
    results: state.results,
    resultMeta: els.resultMeta.textContent,
    departments: structuredClone(state.departmentWork),
    createdAt: existingProject?.createdAt || now,
    updatedAt: now,
  };
  if (destination === "cloud") project.folderId = cloudFolderId ?? existingProject?.folderId ?? "";
  const index = projects.findIndex((item) => item.id === id);
  state.activeProjectId = id;
  els.saveHint.textContent = `${destination === "cloud" ? "Cloud" : "เครื่องนี้"} · บันทึกล่าสุด ${formatSavedDate(now)}`;
  if (destination === "cloud") {
    try {
      setStatus("กำลังบันทึกลง Google Sheet...");
      const data = await cloudRequest("/api/cloud-projects", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ project }),
      });
      const cloudIndex = state.cloudProjects.findIndex((item) => item.id === id);
      if (cloudIndex >= 0) state.cloudProjects[cloudIndex] = data.project;
      else state.cloudProjects.push(data.project);
      renderSavedProjects();
      setStatus("บันทึกงานลง Google Sheet แล้ว");
      return;
    } catch (error) {
      setStatus(`บันทึกขึ้น Cloud ไม่สำเร็จ: ${error.message}`, true);
      return;
    }
  }
  if (index >= 0) projects[index] = project;
  else projects.push(project);
  writeSavedProjects(projects);
  setStatus("บันทึกงานลงเครื่องนี้แล้ว");
}

function waitForGoogleIdentity(timeoutMs = 10000) {
  return new Promise((resolve, reject) => {
    const started = Date.now();
    const timer = window.setInterval(() => {
      if (window.google?.accounts?.id) {
        window.clearInterval(timer);
        resolve();
      } else if (Date.now() - started > timeoutMs) {
        window.clearInterval(timer);
        reject(new Error("โหลดระบบ Google Sign-in ไม่สำเร็จ"));
      }
    }, 100);
  });
}

async function setupCloudLogin(config) {
  state.cloudConfigured = Boolean(config.configured);
  if (!state.cloudConfigured) {
    els.saveCloudProject.disabled = true;
    els.cloudNotice.hidden = false;
    els.cloudNotice.textContent = "Cloud ยังไม่พร้อมใช้งาน ผู้ดูแลต้องตั้งค่า Google Client ID และ Service Account บน Render";
    renderSavedProjects();
    return;
  }
  els.saveCloudProject.disabled = false;
  try {
    await waitForGoogleIdentity();
    window.google.accounts.id.initialize({
      client_id: config.clientId,
      callback: handleGoogleCredential,
      auto_select: false,
    });
    window.google.accounts.id.renderButton(els.googleSignIn, {
      theme: "outline",
      size: "large",
      text: "signin_with",
      locale: "th",
    });
    window.google.accounts.id.renderButton(els.googleSignInQuick, {
      theme: "outline",
      size: "medium",
      text: "signin_with",
      locale: "th",
    });
    if (state.googleCredential) await loadCloudProjects();
  } catch (error) {
    els.cloudNotice.hidden = false;
    els.cloudNotice.textContent = error.message;
  }
}

async function handleGoogleCredential(response) {
  state.googleCredential = response.credential || "";
  sessionStorage.setItem("material-calculator-google-credential", state.googleCredential);
  await loadCloudProjects();
}

async function loadCloudProjects() {
  try {
    const data = await cloudRequest("/api/cloud-projects");
    state.cloudProjects = data.projects || [];
    state.cloudFolders = data.folders || [];
    state.cloudUser = data.user || null;
    els.googleSignIn.hidden = true;
    els.googleSignInQuick.hidden = true;
    els.cloudUser.hidden = false;
    els.cloudUserName.textContent = state.cloudUser.name || state.cloudUser.email;
    els.cloudQuickUser.hidden = false;
    els.cloudQuickUser.textContent = `Cloud: ${state.cloudUser.name || state.cloudUser.email}`;
    els.cloudNotice.hidden = true;
    renderSavedProjects();
    setStatus("เชื่อมต่อ Google Sheet แล้ว");
  } catch (error) {
    state.googleCredential = "";
    state.cloudUser = null;
    state.cloudProjects = [];
    state.cloudFolders = [];
    sessionStorage.removeItem("material-calculator-google-credential");
    els.googleSignIn.hidden = false;
    els.googleSignInQuick.hidden = false;
    els.cloudUser.hidden = true;
    els.cloudQuickUser.hidden = true;
    els.cloudNotice.hidden = false;
    els.cloudNotice.textContent = error.message;
    renderSavedProjects();
  }
}

document.querySelectorAll(".app-tab").forEach((button) => {
  button.addEventListener("click", () => switchTab(button.dataset.tab));
});

els.saveLocalProject.addEventListener("click", () => saveProject("local"));
function openCloudSaveDialog() {
  if (!state.cloudUser) {
    saveProject("cloud");
    return;
  }
  const existingProject = state.cloudProjects.find((project) => project.id === state.activeProjectId);
  els.cloudSaveFolder.replaceChildren();
  [{ id: "", name: "งานทั่วไป" }, ...state.cloudFolders].forEach((folder) => {
    const option = document.createElement("option");
    option.value = folder.id;
    option.textContent = folder.name;
    option.selected = String(existingProject?.folderId || "") === folder.id;
    els.cloudSaveFolder.appendChild(option);
  });
  els.cloudSaveDialog.showModal();
}
els.saveCloudProject.addEventListener("click", openCloudSaveDialog);
els.cancelCloudSave.addEventListener("click", () => els.cloudSaveDialog.close());
els.cloudSaveForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const folderId = els.cloudSaveFolder.value;
  els.cloudSaveDialog.close();
  saveProject("cloud", folderId);
});
els.createCloudFolder.addEventListener("click", async () => {
  if (!state.cloudUser) {
    setStatus("กรุณาเข้าสู่ระบบ Google ก่อนสร้างโฟลเดอร์", true);
    return;
  }
  const name = window.prompt("ตั้งชื่อโฟลเดอร์ Cloud");
  if (!name?.trim()) return;
  try {
    const data = await cloudRequest("/api/cloud-folders", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: name.trim() }),
    });
    state.cloudFolders.push(data.folder);
    state.cloudFolders.sort((a, b) => a.name.localeCompare(b.name, "th"));
    renderSavedProjects();
    setStatus(`สร้างโฟลเดอร์ “${data.folder.name}” แล้ว`);
  } catch (error) { setStatus(error.message, true); }
});
els.newProject.addEventListener("click", resetProject);
els.googleSignOut.addEventListener("click", () => {
  window.google?.accounts?.id?.disableAutoSelect();
  state.googleCredential = "";
  state.cloudUser = null;
  state.cloudProjects = [];
  state.cloudFolders = [];
  sessionStorage.removeItem("material-calculator-google-credential");
  els.googleSignIn.hidden = false;
  els.googleSignInQuick.hidden = false;
  els.cloudUser.hidden = true;
  els.cloudQuickUser.hidden = true;
  renderSavedProjects();
  setStatus("ออกจากระบบ Google แล้ว");
});
function handleSavedProjectAction(event) {
  const button = event.target.closest("button[data-action]");
  if (!button) return;
  const projectId = button.dataset.projectId;
  const source = button.dataset.source || "local";
  if (button.dataset.action === "open") {
    openSavedProject(projectId, source);
    return;
  }
  const sourceProjects = source === "cloud" ? state.cloudProjects : getSavedProjects();
  const project = sourceProjects.find((item) => item.id === projectId);
  if (!project || !window.confirm(`ลบงาน “${project.name}” ใช่ไหม`)) return;
  if (source === "cloud") {
    cloudRequest("/api/cloud-projects/delete", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ projectId }),
    }).then(() => {
      state.cloudProjects = state.cloudProjects.filter((item) => item.id !== projectId);
      renderSavedProjects();
      setStatus("ลบงานออกจาก Google Sheet แล้ว");
    }).catch((error) => setStatus(error.message, true));
    return;
  }
  writeSavedProjects(getSavedProjects().filter((item) => item.id !== projectId));
  if (state.activeProjectId === projectId) state.activeProjectId = "";
  setStatus("ลบงานที่บันทึกแล้ว");
}
async function handleCloudFolderChange(event) {
  const select = event.target.closest('select[data-action="move"]');
  if (!select) return;
  const projectId = select.dataset.projectId;
  try {
    select.disabled = true;
    const data = await cloudRequest("/api/cloud-projects/move", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ projectId, folderId: select.value }),
    });
    const index = state.cloudProjects.findIndex((project) => project.id === projectId);
    if (index >= 0) state.cloudProjects[index] = data.project;
    renderSavedProjects();
    setStatus("ย้ายงานไปยังโฟลเดอร์แล้ว");
  } catch (error) {
    select.disabled = false;
    setStatus(error.message, true);
  }
}
async function handleCloudFolderAction(event) {
  const button = event.target.closest('button[data-action="delete-folder"]');
  if (!button) return;
  event.preventDefault();
  const folder = state.cloudFolders.find((item) => item.id === button.dataset.folderId);
  if (!folder || !window.confirm(`ลบโฟลเดอร์ “${folder.name}” ใช่ไหม`)) return;
  try {
    await cloudRequest("/api/cloud-folders/delete", {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ folderId: folder.id }),
    });
    state.cloudFolders = state.cloudFolders.filter((item) => item.id !== folder.id);
    renderSavedProjects();
    setStatus("ลบโฟลเดอร์แล้ว");
  } catch (error) { setStatus(error.message, true); }
}
els.localSavedProjectList.addEventListener("click", handleSavedProjectAction);
els.cloudSavedProjectList.addEventListener("click", handleSavedProjectAction);
els.cloudSavedProjectList.addEventListener("change", handleCloudFolderChange);
els.cloudSavedProjectList.addEventListener("click", handleCloudFolderAction);

els.applyPages.addEventListener("click", () => {
  saveCurrentPageFromDom();
  let installTotal = Math.max(0, Number.parseInt(els.totalPages.value || "0", 10));
  const demolitionTotal = Math.max(0, Number.parseInt(els.demolitionPages.value || "0", 10));
  if (installTotal + demolitionTotal === 0) {
    installTotal = 1;
    els.totalPages.value = "1";
  }
  const resizePages = (workType, total) => {
    const pages = state.pages.filter((page) => pageWorkType(page) === workType).slice(0, total);
    while (pages.length < total) pages.push([blankRow(state.department, workType), blankRow(state.department, workType)]);
    pages.forEach((page) => page.forEach((row) => { row.workType = workType; row.department = state.department; }));
    return pages;
  };
  state.pages = [...resizePages("install", installTotal), ...resizePages("demolition", demolitionTotal)];
  if (!workTypePageIndices(state.activeWorkType).length) state.activeWorkType = installTotal ? "install" : "demolition";
  state.resultWorkType = state.activeWorkType;
  state.currentPage = workTypePageIndices(state.activeWorkType)[0];
  renderInputs();
  setStatus(`กำหนดงานติดตั้ง ${installTotal} หน้า และงานรื้อถอน ${demolitionTotal} หน้าแล้ว`);
});

document.querySelectorAll(".work-type-button").forEach((button) => {
  button.addEventListener("click", () => {
    const workType = button.dataset.workType;
    const indices = workTypePageIndices(workType);
    if (!indices.length) {
      setStatus(`ยังไม่มี${workTypeLabel(workType)} กรุณากำหนดจำนวนหน้าก่อน`, true);
      return;
    }
    saveCurrentPageFromDom();
    state.activeWorkType = workType;
    state.resultWorkType = workType;
    state.currentPage = indices[0];
    renderInputs();
    updateResultTypeView();
  });
});

els.prevPage.addEventListener("click", () => {
  saveCurrentPageFromDom();
  const indices = workTypePageIndices();
  const position = indices.indexOf(state.currentPage);
  state.currentPage = indices[Math.max(0, position - 1)];
  renderInputs();
});

els.nextPage.addEventListener("click", () => {
  saveCurrentPageFromDom();
  const indices = workTypePageIndices();
  const position = indices.indexOf(state.currentPage);
  state.currentPage = indices[Math.min(indices.length - 1, position + 1)];
  renderInputs();
});

els.addRow.addEventListener("click", () => {
  saveCurrentPageFromDom();
  state.pages[state.currentPage].push(blankRow(state.department, state.activeWorkType));
  renderInputs();
});

els.removeRow.addEventListener("click", () => {
  saveCurrentPageFromDom();
  if (state.pages[state.currentPage].length > 1) {
    state.pages[state.currentPage].pop();
  }
  renderInputs();
});

els.clearPage.addEventListener("click", () => {
  state.pages[state.currentPage] = [blankRow(state.department, state.activeWorkType), blankRow(state.department, state.activeWorkType)];
  renderInputs();
  setStatus("ล้างข้อมูลหน้านี้แล้ว");
});

els.calculate.addEventListener("click", async () => {
  try {
    saveCurrentPageFromDom();
    setStatus("กำลังแสดง Detail พัสดุ...");
    const data = await postJson("/api/calculate", { pages: state.pages });
    renderResults(data.items, `รวม ${data.summaryRows} รายการ จากข้อมูลที่เลือก ${data.inputRows} แถว`);
    setStatus("แสดง Detail พัสดุสำเร็จ");
  } catch (error) {
    setStatus(error.message, true);
  }
});

els.expandSet.addEventListener("click", async () => {
  try {
    setStatus("กำลังสร้างรายการประมาณการ...");
    const data = await postJson("/api/expand-set");
    let meta = `รวม ${data.summaryRows} รายการ | พบ SET ${data.setFound} รายการ | แตกได้ ${data.expandedLines} แถว`;
    if (data.setMissing.length) {
      meta += ` | ไม่พบ: ${data.setMissing.slice(0, 6).join(", ")}`;
    }
    renderResults(data.items, meta);
    setStatus("สร้างรายการประมาณการสำเร็จ");
  } catch (error) {
    setStatus(error.message, true);
  }
});

els.exportExcel.addEventListener("click", async () => {
  try {
    setStatus("กำลังสร้าง Excel...");
    const response = await fetch("/api/export", { method: "POST" });
    if (!response.ok) {
      const data = await response.json();
      throw new Error(data.error || "Export ไม่สำเร็จ");
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = exportFileName("ผลลัพธ์");
    link.click();
    URL.revokeObjectURL(url);
    setStatus("Export สำเร็จ");
  } catch (error) {
    setStatus(error.message, true);
  }
});

els.exportPages.addEventListener("click", async () => {
  try {
    saveCurrentPageFromDom();
    setStatus("กำลังสร้างสรุปแต่ละหน้า...");
    const response = await fetch("/api/export-pages", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pages: state.pages }),
    });
    if (!response.ok) {
      const data = await response.json();
      throw new Error(data.error || "Export รายการแต่ละหน้าไม่สำเร็จ");
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = exportFileName("รายการหัวเสาแยกหน้า");
    link.click();
    URL.revokeObjectURL(url);
    setStatus("Export รายการแต่ละหน้าสำเร็จ");
  } catch (error) {
    setStatus(error.message, true);
  }
});

els.exportPageHardware.addEventListener("click", async () => {
  try {
    saveCurrentPageFromDom();
    setStatus("กำลังสร้างไฟล์ลูกถ้วย/Preform และสรุปลูกถ้วยแยกหน้า...");
    const response = await fetch("/api/export-page-hardware", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pages: state.pages }),
    });
    if (!response.ok) {
      const data = await response.json();
      throw new Error(data.error || "Export ลูกถ้วยและอุปกรณ์ยึดสายไม่สำเร็จ");
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = exportFileName("ลูกถ้วย-Preformแยกหน้า");
    link.click();
    URL.revokeObjectURL(url);
    setStatus("Export ลูกถ้วย/Preform และสรุปลูกถ้วยแยกหน้าสำเร็จ");
  } catch (error) {
    setStatus(error.message, true);
  }
});

els.exportPageCrossarms.addEventListener("click", async () => {
  try {
    saveCurrentPageFromDom();
    setStatus("กำลัง Export คอนแยกตามหน้าและหัวเสา...");
    const response = await fetch("/api/export-page-crossarms", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pages: state.pages }),
    });
    if (!response.ok) {
      const data = await response.json();
      throw new Error(data.error || "Export คอนแยกหน้าไม่สำเร็จ");
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = exportFileName("คอนแยกหน้า");
    link.click();
    URL.revokeObjectURL(url);
    setStatus("Export คอนแยกหน้าสำเร็จ");
  } catch (error) {
    setStatus(error.message, true);
  }
});

function addRequestMaterialRow(value = {}, isOriginal = false) {
  const row = document.createElement("tr");
  if (isOriginal) row.dataset.original = JSON.stringify({ material: value.material || "", code: value.code || "", quantity: Number(value.quantity) });
  for (const [className, placeholder, type] of [
    ["request-material", "ชื่อรายการวัสดุ", "text"],
    ["request-code", "รหัสพัสดุหรือ Set", "text"],
    ["request-quantity", "จำนวน", "number"],
  ]) {
    const cell = document.createElement("td");
    const input = document.createElement("input");
    input.className = className;
    input.type = type;
    input.placeholder = placeholder;
    input.required = true;
    if (type === "number") input.step = "any";
    input.value = className === "request-material" ? (value.material || "")
      : className === "request-code" ? (value.code || "") : (value.quantity ?? "");
    input.addEventListener("input", () => updateRequestRowStatus(row));
    cell.appendChild(input);
    row.appendChild(cell);
  }
  const statusCell = document.createElement("td");
  const status = document.createElement("span");
  status.className = "change-badge";
  statusCell.appendChild(status);
  row.appendChild(statusCell);
  const actionCell = document.createElement("td");
  const remove = document.createElement("button");
  remove.type = "button";
  remove.className = "danger compact";
  remove.textContent = isOriginal ? "นำออก" : "ลบ";
  remove.addEventListener("click", () => {
    if (row.dataset.original) {
      const removed = row.classList.toggle("row-removed");
      remove.textContent = removed ? "คืนค่า" : "นำออก";
      row.querySelectorAll("input").forEach((input) => { input.disabled = removed; });
      updateRequestRowStatus(row);
    } else if (els.requestMaterialRows.children.length > 1) {
      row.remove();
    }
  });
  actionCell.appendChild(remove);
  row.appendChild(actionCell);
  els.requestMaterialRows.appendChild(row);
  updateRequestRowStatus(row);
}

function updateRequestRowStatus(row) {
  const badge = row.querySelector(".change-badge");
  if (row.classList.contains("row-removed")) {
    badge.textContent = "นำออก"; badge.dataset.kind = "removed"; return;
  }
  if (!row.dataset.original) {
    badge.textContent = "เพิ่มใหม่"; badge.dataset.kind = "added"; return;
  }
  const original = JSON.parse(row.dataset.original);
  const current = {
    material: row.querySelector(".request-material").value.trim(),
    code: row.querySelector(".request-code").value.trim(),
    quantity: Number(row.querySelector(".request-quantity").value),
  };
  const changed = original.material !== current.material || original.code !== current.code || original.quantity !== current.quantity;
  badge.textContent = changed ? "แก้ไข" : "คงเดิม";
  badge.dataset.kind = changed ? "changed" : "same";
}

function requestMaterialValues() {
  return [...els.requestMaterialRows.querySelectorAll("tr:not(.row-removed)")].map((row) => ({
    material: row.querySelector(".request-material").value.trim(),
    code: row.querySelector(".request-code").value.trim(),
    quantity: row.querySelector(".request-quantity").value,
  }));
}

function selectedRequestImage() {
  return els.requestImage.files?.[0] || null;
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("อ่านไฟล์รูปไม่สำเร็จ"));
    reader.readAsDataURL(file);
  });
}

function renderRequestImagePreview(localUrl = "") {
  const image = els.requestImagePreview.querySelector("img");
  const inherited = state.baseRequestImage;
  const url = localUrl || (inherited?.id ? imageUrl(inherited.id) : "");
  els.requestImagePreview.hidden = !url;
  els.removeRequestImage.hidden = !url;
  image.src = url;
}

async function uploadSelectedRequestImage() {
  const file = selectedRequestImage();
  if (!file) return state.baseRequestImage;
  if (file.size > 5 * 1024 * 1024) throw new Error("รูปต้องมีขนาดไม่เกิน 5 MB");
  const data = await postJson("/api/base-images/upload", { name: file.name, data: await fileToDataUrl(file) });
  return data.image;
}

async function submitBaseRequest(event) {
  event.preventDefault();
  try {
    setStatus("กำลังส่งคำขอให้ Admin ตรวจ...");
    const action = els.requestAction.value;
    const usesSource = action !== "add";
    const size = action === "add" ? els.requestSize.value : els.replaceSize.value;
    const head = action === "replace" ? els.replaceHead.value : els.requestHead.value;
    const requestImage = els.requestTargetDepartment.value === "แผนกแรงสูง TAC" ? await uploadSelectedRequestImage() : null;
    const response = await fetch("/api/base-requests", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        submitter_name: els.requesterName.value, employee_id: els.requesterEmployeeId.value,
        department: els.requesterDepartment.value, action,
        target_department: els.requestTargetDepartment.value,
        insulator_upright: els.requestInsulatorUpright.value,
        insulator_horizontal: els.requestInsulatorHorizontal.value,
        size, head, rows: requestMaterialValues(), original_rows: usesSource ? state.baseRequestOriginalRows : [],
        source_size: usesSource ? els.replaceSize.value : "", source_head: usesSource ? els.replaceHead.value : "",
        image_file_id: requestImage?.id || "", image_name: requestImage?.name || "", image_mime_type: requestImage?.mimeType || "",
        note: els.requestNote.value,
      }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "ส่งคำขอไม่สำเร็จ");
    els.requestSize.value = "";
    els.requestHead.value = "";
    els.requestInsulatorUpright.value = "0";
    els.requestInsulatorHorizontal.value = "0";
    els.requestNote.value = "";
    els.requestMaterialRows.innerHTML = "";
    state.baseRequestOriginalRows = [];
    state.baseRequestImage = null;
    els.requestImage.value = "";
    renderRequestImagePreview();
    els.requestAction.value = "add";
    await switchRequestAction();
    setStatus(`ส่งคำขอ ${data.request.id.slice(0, 8)} ให้ Admin แล้ว`);
  } catch (error) {
    setStatus(error.message, true);
  }
}

function setSelectOptions(select, values, placeholder) {
  select.innerHTML = "";
  const empty = document.createElement("option");
  empty.value = ""; empty.textContent = placeholder; select.appendChild(empty);
  values.forEach((value) => {
    const option = document.createElement("option"); option.value = value; option.textContent = value; select.appendChild(option);
  });
}

function populateDepartmentSelectors(departments = DEPARTMENTS) {
  const values = departments.length ? departments : DEPARTMENTS;
  for (const select of [els.departmentSelect, els.requestTargetDepartment]) {
    select.innerHTML = "";
    values.forEach((value) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = value;
      select.appendChild(option);
    });
    select.value = DEFAULT_DEPARTMENT;
  }
}

async function loadDepartmentSizes(department) {
  const response = await fetch(`/api/sizes?department=${encodeURIComponent(department)}`);
  const data = await readJson(response);
  state.sizes = data.sizes || [];
}

async function changeDepartment(department) {
  stashCurrentDepartment();
  state.department = department || DEFAULT_DEPARTMENT;
  const work = state.departmentWork[state.department] || emptyDepartmentWork(state.department);
  state.pages = clonePages(work.pages);
  state.pages.forEach((page) => page.forEach((row) => { row.department = state.department; }));
  state.currentPage = Math.min(Number(work.currentPage || 0), state.pages.length - 1);
  state.activeWorkType = work.activeWorkType || pageWorkType(state.pages[state.currentPage]);
  state.resultWorkType = state.activeWorkType;
  state.results = structuredClone(work.results || []);
  try {
    await loadDepartmentSizes(state.department);
    renderInputs();
    renderResults(state.results, work.resultMeta || "ยังไม่มีผลคำนวณ");
    setStatus(state.sizes.length ? `เปิดข้อมูล ${state.department} แล้ว` : `${state.department} ยังไม่มีข้อมูล เริ่มเพิ่มผ่านเมนูเพิ่มเติม/แก้ไขหัวเสาได้เลย`);
  } catch (error) {
    setStatus(error.message, true);
  }
}

async function switchRequestAction() {
  const action = els.requestAction.value;
  const usesSource = action !== "add";
  els.requestAddTarget.hidden = action === "replace";
  els.requestReplaceTarget.hidden = !usesSource;
  els.existingDataHint.hidden = !usesSource;
  els.requestSizeField.hidden = action !== "add";
  els.requestHeadLabel.textContent = usesSource ? "ชื่อหัวเสาใหม่" : "รหัสหัวเสา / รายการใหม่";
  els.requestMaterialRows.innerHTML = "";
  state.baseRequestOriginalRows = [];
  state.baseRequestImage = null;
  els.requestImage.value = "";
  els.requestImagePanel.hidden = els.requestTargetDepartment.value !== "แผนกแรงสูง TAC";
  renderRequestImagePreview();
  if (!usesSource) {
    addRequestMaterialRow();
    return;
  }
  await loadRequestDepartmentSizes();
  setSelectOptions(els.replaceHead, [], "เลือกหัวเสา");
  els.existingDataHint.textContent = "เลือกขนาดเสาและหัวเสาเพื่อโหลดข้อมูลเดิม";
}

async function loadRequestDepartmentSizes() {
  try {
    const response = await fetch(`/api/sizes?department=${encodeURIComponent(els.requestTargetDepartment.value)}`);
    const data = await readJson(response);
    setSelectOptions(els.replaceSize, data.sizes, data.sizes.length ? "เลือกขนาด/KeyCode" : "แผนกนี้ยังไม่มีข้อมูล");
    setSelectOptions(els.replaceHead, [], "เลือกหัวเสา/รายการ");
  } catch (error) { setStatus(error.message, true); }
}

async function loadReplaceHeads() {
  setSelectOptions(els.replaceHead, [], "กำลังโหลด...");
  els.requestMaterialRows.innerHTML = "";
  state.baseRequestOriginalRows = [];
  if (!els.replaceSize.value) return;
  try {
    const response = await fetch(`/api/heads?size=${encodeURIComponent(els.replaceSize.value)}&department=${encodeURIComponent(els.requestTargetDepartment.value)}`);
    const data = await readJson(response);
    setSelectOptions(els.replaceHead, data.heads, "เลือกหัวเสา");
    els.existingDataHint.textContent = "เลือกหัวเสาเพื่อดูรายการเดิม";
  } catch (error) { setStatus(error.message, true); }
}

async function loadExistingBaseEntry() {
  els.requestMaterialRows.innerHTML = "";
  state.baseRequestOriginalRows = [];
  if (!els.replaceSize.value || !els.replaceHead.value) return;
  try {
    const response = await fetch(`/api/base-entry?size=${encodeURIComponent(els.replaceSize.value)}&head=${encodeURIComponent(els.replaceHead.value)}&department=${encodeURIComponent(els.requestTargetDepartment.value)}`);
    const data = await readJson(response);
    state.baseRequestOriginalRows = data.rows.map((row) => ({ ...row }));
    els.requestInsulatorUpright.value = String(data.insulatorUpright ?? 0);
    els.requestInsulatorHorizontal.value = String(data.insulatorHorizontal ?? 0);
    state.baseRequestImage = data.image?.id ? data.image : null;
    renderRequestImagePreview();
    data.rows.forEach((row) => addRequestMaterialRow(row, true));
    const action = els.requestAction.value;
    if (action === "rename") {
      els.requestHead.value = els.replaceHead.value;
    } else if (action === "copy") {
      els.requestHead.value = `${els.replaceHead.value} COPY`;
    }
    els.existingDataHint.textContent = action === "rename"
      ? `โหลดข้อมูลเดิม ${data.rows.length} รายการแล้ว ใส่ชื่อหัวเสาใหม่ด้านบน`
      : action === "copy"
        ? `คัดลอกข้อมูลเดิม ${data.rows.length} รายการแล้ว เปลี่ยนชื่อและแก้ไส้ในได้ทันที`
        : `โหลดข้อมูลเดิม ${data.rows.length} รายการแล้ว แก้ไข เพิ่ม หรือนำรายการออกได้`;
  } catch (error) { setStatus(error.message, true); }
}

async function adminFetch(url, options = {}) {
  const headers = { ...(options.headers || {}), Authorization: `Bearer ${state.adminToken}` };
  const response = await fetch(url, { ...options, headers });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "ดำเนินการ Admin ไม่สำเร็จ");
  return data;
}

function showAdminPanel() {
  els.adminLoginPanel.hidden = Boolean(state.adminToken);
  els.adminRequestsPanel.hidden = !state.adminToken;
  if (state.adminToken) {
    loadBaseRequests();
    loadDriveOAuthStatus();
  }
}

async function loadDriveOAuthStatus() {
  try {
    const response = await fetch("/api/base-admin/config", { cache: "no-store" });
    const data = await readJson(response);
    els.connectGoogleDrive.hidden = Boolean(data.driveConnected);
    els.connectGoogleDrive.disabled = !data.driveOAuthReady;
    els.driveConnectStatus.textContent = data.driveConnected
      ? "เชื่อมต่อ OAuth แล้ว พร้อมเก็บรูปใน Drive ส่วนตัว"
      : data.driveOAuthReady
        ? "พร้อมเชื่อมต่อ กรุณากดปุ่มและอนุญาตบัญชีเจ้าของ Drive"
        : "กรุณาตั้งค่า GOOGLE_CLIENT_SECRET บน Render ก่อน";
  } catch (error) {
    els.driveConnectStatus.textContent = error.message;
  }
}

els.connectGoogleDrive.addEventListener("click", async () => {
  try {
    els.connectGoogleDrive.disabled = true;
    const data = await adminFetch("/api/google-drive/oauth/start", { method: "POST" });
    window.location.href = data.authorizationUrl;
  } catch (error) {
    setStatus(error.message, true);
    els.connectGoogleDrive.disabled = false;
  }
});

async function loadBaseRequests() {
  try {
    const data = await adminFetch("/api/base-requests/admin");
    state.baseRequests = data.requests || [];
    renderBaseRequests();
  } catch (error) {
    state.adminToken = "";
    sessionStorage.removeItem("material-calculator-admin-token");
    showAdminPanel();
    setStatus(error.message, true);
  }
}

function renderBaseRequests() {
  const counts = { pending: 0, approved: 0, rejected: 0 };
  state.baseRequests.forEach((request) => { if (request.status in counts) counts[request.status] += 1; });
  els.pendingRequestCount.textContent = counts.pending;
  els.approvedRequestCount.textContent = counts.approved;
  els.rejectedRequestCount.textContent = counts.rejected;
  els.clearApprovedRequests.hidden = state.activeRequestStatus !== "approved" || counts.approved === 0;
  document.querySelectorAll(".request-status-tab").forEach((button) => {
    button.classList.toggle("active", button.dataset.requestStatus === state.activeRequestStatus);
  });
  const requests = state.baseRequests.filter((request) => request.status === state.activeRequestStatus);
  els.baseRequestList.innerHTML = "";
  if (!requests.length) {
    const empty = document.createElement("div");
    empty.className = "empty-saved";
    empty.textContent = state.activeRequestStatus === "pending" ? "ไม่มีคำขอที่รอตรวจ" : state.activeRequestStatus === "approved" ? "ยังไม่มีคำขอที่อนุมัติแล้ว" : "ยังไม่มีคำขอที่ปฏิเสธ";
    els.baseRequestList.appendChild(empty);
    return;
  }
  requests.forEach((request) => {
    const card = document.createElement("article");
    card.className = `request-card status-${request.status}`;
    const title = document.createElement("h3");
    const sourceLabel = request.sourceHead ? `${request.sourceSize} · ${request.sourceHead}` : "";
    title.textContent = sourceLabel && request.action !== "replace"
      ? `${sourceLabel} → ${request.size} · ${request.head}`
      : `${request.size} · ${request.head}`;
    const meta = document.createElement("p");
    const actionLabels = { add: "เพิ่มข้อมูล", replace: "แก้ไขไส้ในหัวเดิม", rename: "แก้ไขหัวเสา", copy: "คัดลอกเป็นหัวใหม่" };
    meta.textContent = `${request.targetDepartment || DEFAULT_DEPARTMENT} · ลูกถ้วยตั้ง ${formatAmount(request.insulatorUpright)} / นอน ${formatAmount(request.insulatorHorizontal)} ต่อหัว · ผู้เสนอ ${request.submitterName} · ${request.employeeId} · ${request.department} · ${actionLabels[request.action] || "เพิ่มข้อมูล"}`;
    const table = document.createElement("table");
    const comparesOriginal = request.action !== "add";
    table.innerHTML = comparesOriginal
      ? "<thead><tr><th>สถานะ</th><th>รายการวัสดุ</th><th>รหัส</th><th>เดิม</th><th>ใหม่</th></tr></thead>"
      : "<thead><tr><th>รายการวัสดุ</th><th>รหัส</th><th>จำนวน</th></tr></thead>";
    const body = document.createElement("tbody");
    const displayedRows = comparesOriginal ? compareBaseRows(request.originalRows || [], request.rows) : request.rows;
    displayedRows.forEach((item) => {
      const row = document.createElement("tr");
      const values = comparesOriginal
        ? [item.statusLabel, item.material, item.code, item.oldQuantity ?? "", item.newQuantity ?? ""]
        : [item.material, item.code, item.quantity];
      values.forEach((value, index) => {
        const cell = document.createElement("td"); cell.textContent = value; row.appendChild(cell);
        if (comparesOriginal && index === 0) cell.className = `diff-${item.status}`;
      });
      body.appendChild(row);
    });
    table.appendChild(body);
    const footer = document.createElement("div");
    footer.className = "request-card-footer";
    const status = document.createElement("strong");
    status.textContent = request.status === "pending" ? "รอตรวจ" : request.status === "approved" ? "อนุมัติแล้ว" : "ปฏิเสธแล้ว";
    footer.appendChild(status);
    if (request.status === "pending") {
      for (const [label, approve, className] of [["อนุมัติ", true, "primary"], ["ปฏิเสธ", false, "danger"]]) {
        const button = document.createElement("button");
        button.type = "button"; button.textContent = label; button.className = className;
        button.addEventListener("click", () => reviewBaseRequest(request.id, approve, button));
        footer.appendChild(button);
      }
    }
    card.append(title, meta);
    if (request.note) { const note = document.createElement("p"); note.textContent = `หมายเหตุ: ${request.note}`; card.appendChild(note); }
    if (request.imageFileId) {
      const imageButton = document.createElement("button");
      imageButton.type = "button";
      imageButton.className = "request-admin-image";
      const image = document.createElement("img");
      image.src = imageUrl(request.imageFileId);
      image.alt = `รูปประกอบ ${request.head}`;
      imageButton.append(image, document.createTextNode("คลิกดูรูปประกอบขนาดใหญ่"));
      imageButton.addEventListener("click", () => openHeadImage({ id: request.imageFileId, name: request.imageName }, request.head));
      card.appendChild(imageButton);
    }
    const wrap = document.createElement("div"); wrap.className = "table-wrap"; wrap.appendChild(table); card.append(wrap, footer);
    els.baseRequestList.appendChild(card);
  });
}

function compareBaseRows(originalRows, newRows) {
  const keyOf = (row) => `${String(row.material).trim()}\u0000${String(row.code).trim()}`;
  const aggregate = (rows) => {
    const result = new Map();
    rows.forEach((row) => {
      const key = keyOf(row);
      if (!result.has(key)) result.set(key, { ...row, quantity: 0 });
      result.get(key).quantity += Number(row.quantity || 0);
    });
    return result;
  };
  const original = aggregate(originalRows);
  const proposed = aggregate(newRows);
  const keys = [...new Set([...original.keys(), ...proposed.keys()])];
  return keys.map((key) => {
    const oldRow = original.get(key);
    const newRow = proposed.get(key);
    let status = "same";
    if (!oldRow) status = "added";
    else if (!newRow) status = "removed";
    else if (Number(oldRow.quantity) !== Number(newRow.quantity)) status = "changed";
    return {
      status, statusLabel: { same: "คงเดิม", added: "เพิ่ม", removed: "นำออก", changed: "เปลี่ยนจำนวน" }[status],
      material: (newRow || oldRow).material, code: (newRow || oldRow).code,
      oldQuantity: oldRow?.quantity, newQuantity: newRow?.quantity,
    };
  });
}

async function reviewBaseRequest(requestId, approve, button) {
  const note = window.prompt(approve ? "หมายเหตุการอนุมัติ (เว้นว่างได้)" : "เหตุผลที่ปฏิเสธ");
  if (note === null) return;
  const originalButtonText = button?.textContent || "";
  if (button) {
    button.disabled = true;
    button.textContent = approve ? "กำลังอนุมัติ..." : "กำลังปฏิเสธ...";
  }
  setStatus(approve ? "กำลังอนุมัติและอัปเดต BaseData กรุณารอสักครู่..." : "กำลังปฏิเสธคำขอ...");
  try {
    await adminFetch("/api/base-requests/review", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ requestId, approve, note }),
    });
    if (approve) {
      sessionStorage.setItem(ADMIN_APPROVAL_REFRESH_KEY, "1");
      const refreshedUrl = new URL(window.location.href);
      refreshedUrl.searchParams.set("baseRefresh", Date.now().toString());
      window.location.replace(refreshedUrl.toString());
      return;
    }
    await loadDepartmentSizes(state.department);
    renderInputs();
    if (els.requestAction.value === "replace") await loadRequestDepartmentSizes();
    state.activeRequestStatus = approve ? "approved" : "rejected";
    setStatus(approve ? "อนุมัติและอัปเดต BaseData แล้ว" : "ปฏิเสธคำขอแล้ว");
    await loadBaseRequests();
  } catch (error) {
    setStatus(error.message, true);
    if (button) {
      button.disabled = false;
      button.textContent = originalButtonText;
    }
  }
}

els.addRequestMaterial.addEventListener("click", () => addRequestMaterialRow());
els.requestImage.addEventListener("change", async () => {
  const file = selectedRequestImage();
  if (!file) { renderRequestImagePreview(); return; }
  if (!(["image/jpeg", "image/png", "image/webp"].includes(file.type)) || file.size > 5 * 1024 * 1024) {
    els.requestImage.value = "";
    setStatus("รองรับรูป JPG, PNG หรือ WEBP ขนาดไม่เกิน 5 MB", true);
    renderRequestImagePreview();
    return;
  }
  renderRequestImagePreview(await fileToDataUrl(file));
});
els.removeRequestImage.addEventListener("click", () => {
  state.baseRequestImage = null;
  els.requestImage.value = "";
  renderRequestImagePreview();
});
els.requestImagePreview.addEventListener("click", () => {
  const source = els.requestImagePreview.querySelector("img").src;
  if (!source) return;
  els.headImageDialogTitle.textContent = els.requestHead.value || els.replaceHead.value || "รูปประกอบหัวเสา";
  els.headImageDialogImage.src = source;
  els.headImageDialog.showModal();
});
els.closeHeadImageDialog.addEventListener("click", () => els.headImageDialog.close());
els.headImageDialog.addEventListener("click", (event) => {
  if (event.target === els.headImageDialog) els.headImageDialog.close();
});
els.baseRequestForm.addEventListener("submit", submitBaseRequest);
els.requestAction.addEventListener("change", switchRequestAction);
els.requestTargetDepartment.addEventListener("change", switchRequestAction);
document.querySelectorAll(".request-status-tab").forEach((button) => {
  button.addEventListener("click", () => {
    state.activeRequestStatus = button.dataset.requestStatus;
    renderBaseRequests();
  });
});
els.clearApprovedRequests.addEventListener("click", async () => {
  const count = Number(els.approvedRequestCount.textContent || 0);
  if (!count || !window.confirm(`ล้างประวัติรายการอนุมัติแล้วทั้งหมด ${count} รายการใช่ไหม?\n\nข้อมูลหัวเสาที่อนุมัติและใช้งานอยู่จะไม่ถูกลบ`)) return;
  try {
    els.clearApprovedRequests.disabled = true;
    els.clearApprovedRequests.textContent = "กำลังล้าง...";
    const data = await adminFetch("/api/base-requests/clear-approved", { method: "POST" });
    setStatus(`ล้างประวัติอนุมัติแล้ว ${data.cleared || 0} รายการ โดยไม่กระทบ BaseData`);
    await loadBaseRequests();
  } catch (error) {
    setStatus(error.message, true);
  } finally {
    els.clearApprovedRequests.disabled = false;
    els.clearApprovedRequests.textContent = "ล้างประวัติอนุมัติแล้ว";
  }
});
els.departmentSelect.addEventListener("change", () => changeDepartment(els.departmentSelect.value));
els.replaceSize.addEventListener("change", loadReplaceHeads);
els.replaceHead.addEventListener("change", loadExistingBaseEntry);
els.adminLoginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  try {
    const response = await fetch("/api/base-admin/login", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username: els.adminUsername.value, password: els.adminPassword.value }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "เข้าสู่ระบบไม่สำเร็จ");
    state.adminToken = data.token;
    sessionStorage.setItem("material-calculator-admin-token", data.token);
    els.adminPassword.value = "";
    showAdminPanel();
    setStatus("เข้าสู่ระบบ Admin แล้ว");
  } catch (error) { setStatus(error.message, true); }
});
els.adminLogout.addEventListener("click", () => {
  state.adminToken = "";
  sessionStorage.removeItem("material-calculator-admin-token");
  showAdminPanel();
});

async function initialize() {
  const returnToApprovedRequests = sessionStorage.getItem(ADMIN_APPROVAL_REFRESH_KEY) === "1";
  sessionStorage.removeItem(ADMIN_APPROVAL_REFRESH_KEY);
  populateDepartmentSelectors();
  if (!els.requestMaterialRows.children.length) addRequestMaterialRow();
  renderSavedProjects();
  renderInputs();
  try {
    const [statusResponse, cloudResponse] = await Promise.all([fetch("/api/status"), fetch("/api/cloud-config")]);
    const data = await readJson(statusResponse);
    const cloudConfig = await readJson(cloudResponse);
    if (data.base) {
      populateDepartmentSelectors(data.base.departments || DEPARTMENTS);
      await loadDepartmentSizes(state.department);
    }
    renderInputs();
    setStatus(data.base ? "โหลดฐานข้อมูลเริ่มต้นแล้ว" : "กรุณาโหลด BaseData");
    setupCloudLogin(cloudConfig);
    if (returnToApprovedRequests) {
      state.activeRequestStatus = "pending";
      switchTab("base-admin");
      setStatus("อนุมัติและรีเฟรช BaseData เรียบร้อยแล้ว — กลับสู่รายการรอตรวจ");
    }
  } catch (error) {
    setStatus(`โหลดฐานข้อมูลเริ่มต้นไม่สำเร็จ: ${error.message}`, true);
  }
}

initialize();
