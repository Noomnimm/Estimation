const DEPARTMENTS = ["แผนกแรงสูง", "แผนกแรงสูง TAC", "แผนกหม้อแปลง", "แผนกสายส่ง"];
const DEFAULT_DEPARTMENT = DEPARTMENTS[0];

const state = {
  department: DEFAULT_DEPARTMENT,
  sizes: [],
  pages: [[blankRow(), blankRow()]],
  currentPage: 0,
  results: [],
  activeProjectId: "",
  cloudConfigured: false,
  googleCredential: sessionStorage.getItem("material-calculator-google-credential") || "",
  cloudProjects: [],
  cloudUser: null,
  adminToken: sessionStorage.getItem("material-calculator-admin-token") || "",
  baseRequestOriginalRows: [],
};

const SAVED_PROJECTS_KEY = "material-calculator-projects-v1";

const els = {
  status: document.getElementById("status"),
  departmentSelect: document.getElementById("departmentSelect"),
  totalPages: document.getElementById("totalPages"),
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
  exportPageInsulators: document.getElementById("exportPageInsulators"),
  exportPageCrossarms: document.getElementById("exportPageCrossarms"),
  baseRequestForm: document.getElementById("baseRequestForm"),
  requesterName: document.getElementById("requesterName"),
  requesterEmployeeId: document.getElementById("requesterEmployeeId"),
  requesterDepartment: document.getElementById("requesterDepartment"),
  requestTargetDepartment: document.getElementById("requestTargetDepartment"),
  requestAction: document.getElementById("requestAction"),
  requestSize: document.getElementById("requestSize"),
  requestHead: document.getElementById("requestHead"),
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
  baseRequestList: document.getElementById("baseRequestList"),
  resultRows: document.getElementById("resultRows"),
  resultMeta: document.getElementById("resultMeta"),
  projectName: document.getElementById("projectName"),
  planNumber: document.getElementById("planNumber"),
  saveProject: document.getElementById("saveProject"),
  newProject: document.getElementById("newProject"),
  saveHint: document.getElementById("saveHint"),
  savedCount: document.getElementById("savedCount"),
  savedProjectList: document.getElementById("savedProjectList"),
  savedLocationText: document.getElementById("savedLocationText"),
  googleSignIn: document.getElementById("googleSignIn"),
  cloudUser: document.getElementById("cloudUser"),
  cloudUserName: document.getElementById("cloudUserName"),
  googleSignOut: document.getElementById("googleSignOut"),
  cloudNotice: document.getElementById("cloudNotice"),
};

function blankRow(department = DEFAULT_DEPARTMENT) {
  return { department, size: "", head: "", count: "", wire1: "", wire2: "", latWire: "" };
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
    Object.assign(state.pages[state.currentPage][index], {
      department: state.department,
      size: tr.querySelector(".size").value,
      head: tr.querySelector(".head").value,
      count: tr.querySelector(".count").value,
    });
  });
}

function renderInputs() {
  els.inputRows.innerHTML = "";
  const page = state.pages[state.currentPage];
  page.forEach((row, index) => {
    const tr = document.createElement("tr");
    tr.className = "input-row";
    tr.innerHTML = `
      <td><select class="size"></select></td>
      <td><select class="head"></select></td>
      <td><input class="count" type="text" inputmode="text" placeholder="เช่น 4+4+5+6"></td>
    `;

    const sizeSelect = tr.querySelector(".size");
    const headSelect = tr.querySelector(".head");
    const countInput = tr.querySelector(".count");

    fillSelect(sizeSelect, state.sizes, "เลือกขนาดเสา");
    sizeSelect.value = row.size || "";
    countInput.value = row.count || "";
    fillSelect(headSelect, [], "เลือกรหัสหัวเสา");

    sizeSelect.addEventListener("change", () => {
      page[index].department = state.department;
      page[index].size = sizeSelect.value;
      page[index].head = "";
      page[index].wire1 = "";
      page[index].wire2 = "";
      page[index].latWire = "";
      renderInputs();
    });
    headSelect.addEventListener("change", () => {
      page[index].head = headSelect.value;
      page[index].wire1 = "";
      page[index].wire2 = "";
      page[index].latWire = "";
      renderInputs();
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
    if (row.size) {
      loadHeads(row.size, headSelect, row.head, row.department || state.department);
    }
  });
  renderPageControls();
  renderInsulators();
}

function renderInsulators() {
  const totals = summarizeInsulators(state.pages);
  document.getElementById("uprightCount").textContent = formatAmount(totals.upright);
  document.getElementById("horizontalCount").textContent = formatAmount(totals.horizontal);
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
  cell.colSpan = 3;
  const panel = document.createElement("div");
  panel.className = "wire-details";

  const labels = wireKind === "de"
    ? [["สาย Dead End", "wire1"]]
    : wireKind === "ba"
      ? [["Main Line", "wire1"], ["Tap Line", "wire2"]]
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

async function loadHeads(size, select, selected, department = state.department) {
  if (!size) {
    fillSelect(select, [], "เลือกรหัสหัวเสา");
    return;
  }
  try {
    const response = await fetch(`/api/heads?size=${encodeURIComponent(size)}&department=${encodeURIComponent(department)}`);
    const data = await readJson(response);
    fillSelect(select, data.heads, "เลือกรหัสหัวเสา");
    select.value = selected || "";
  } catch (error) {
    setStatus(error.message, true);
  }
}

function renderPageControls() {
  els.totalPages.value = String(state.pages.length);
  els.pageLabel.textContent = `หน้า ${state.currentPage + 1}/${state.pages.length}`;
  els.prevPage.disabled = state.currentPage === 0;
  els.nextPage.disabled = state.currentPage === state.pages.length - 1;
  els.pagePicker.innerHTML = "";
  state.pages.forEach((_, index) => {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = String(index + 1);
    button.className = "page-number";
    button.classList.toggle("active", index === state.currentPage);
    button.setAttribute("aria-label", `ไปหน้าที่ ${index + 1}`);
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
  els.resultRows.innerHTML = "";
  state.results.forEach((item) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item["รายการวัสดุ"] || "")}</td>
      <td>${escapeHtml(item["รหัสพัสดุ"] || "")}</td>
      <td>${formatAmount(item["จำนวนรวม"])}</td>
    `;
    els.resultRows.appendChild(tr);
  });
  els.resultMeta.textContent = meta || `ทั้งหมด ${state.results.length} รายการ`;
}

function formatAmount(value) {
  const number = Number(value || 0);
  return Number.isInteger(number) ? String(number) : number.toFixed(3).replace(/0+$/, "").replace(/\.$/, "");
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
  return pages.map((page) => page.map((row) => ({ ...blankRow(state.department), ...row })));
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
  const projects = (state.cloudUser ? state.cloudProjects : getSavedProjects())
    .sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
  els.savedCount.textContent = String(projects.length);
  els.savedLocationText.textContent = state.cloudUser
    ? "ข้อมูลจาก Google Sheet กลางของทีม"
    : "ข้อมูลในเบราว์เซอร์เครื่องนี้ - เข้าสู่ระบบเพื่อเปิดงาน Cloud";
  if (!projects.length) {
    els.savedProjectList.innerHTML = '<div class="empty-saved">ยังไม่มีงานที่บันทึกไว้</div>';
    return;
  }
  els.savedProjectList.innerHTML = projects.map((project) => `
    <article class="saved-card">
      <div class="saved-info">
        <h3>${escapeHtml(project.name)}</h3>
        <div class="saved-meta">
          <span>${escapeHtml(project.department || project.pages?.[0]?.[0]?.department || DEFAULT_DEPARTMENT)}</span>
          <span>เลขผัง: ${escapeHtml(project.planNumber || "-")}</span>
          <span>${Number(project.pages?.length || 0)} หน้า</span>
          <span>แก้ไขล่าสุด ${escapeHtml(formatSavedDate(project.updatedAt))}</span>
        </div>
      </div>
      <div class="saved-actions">
        <button type="button" data-action="open" data-source="${state.cloudUser ? "cloud" : "local"}" data-project-id="${escapeHtml(project.id)}">เปิดงาน</button>
        <button type="button" class="danger" data-action="delete" data-source="${state.cloudUser ? "cloud" : "local"}" data-project-id="${escapeHtml(project.id)}">ลบ</button>
      </div>
    </article>
  `).join("");
}

function resetProject() {
  state.activeProjectId = "";
  state.pages = [[blankRow(state.department), blankRow(state.department)]];
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
  els.departmentSelect.value = state.department;
  await loadDepartmentSizes(state.department);
  state.pages = clonePages(project.pages?.length ? project.pages : [[blankRow(), blankRow()]]);
  state.pages.forEach((page) => page.forEach((row) => { row.department = state.department; }));
  state.currentPage = Math.min(Number(project.currentPage || 0), state.pages.length - 1);
  els.projectName.value = project.name || "";
  els.planNumber.value = project.planNumber || "";
  els.saveHint.textContent = `เปิดงานที่บันทึกเมื่อ ${formatSavedDate(project.updatedAt)}`;
  renderInputs();
  renderResults(project.results || [], project.resultMeta || "ยังไม่มีผลคำนวณ");
  switchTab("calculator");
  setStatus("เปิดงานเดิมสำเร็จ");
}

async function saveProject() {
  saveCurrentPageFromDom();
  const name = els.projectName.value.trim();
  if (!name) {
    els.projectName.focus();
    setStatus("กรุณาใส่ชื่องานก่อนบันทึก", true);
    return;
  }
  const projects = getSavedProjects();
  const now = new Date().toISOString();
  const id = state.activeProjectId || `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
  const project = {
    id,
    name,
    planNumber: els.planNumber.value.trim(),
    department: state.department,
    pages: clonePages(state.pages),
    currentPage: state.currentPage,
    results: state.results,
    resultMeta: els.resultMeta.textContent,
    createdAt: projects.find((item) => item.id === id)?.createdAt || now,
    updatedAt: now,
  };
  const index = projects.findIndex((item) => item.id === id);
  if (index >= 0) projects[index] = project;
  else projects.push(project);
  state.activeProjectId = id;
  writeSavedProjects(projects);
  els.saveHint.textContent = `บันทึกล่าสุด ${formatSavedDate(now)}`;
  if (state.cloudUser) {
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
      setStatus(`บันทึกในเครื่องแล้ว แต่ Cloud ไม่สำเร็จ: ${error.message}`, true);
      return;
    }
  }
  setStatus(state.cloudConfigured ? "บันทึกในเครื่องแล้ว - เข้าสู่ระบบเพื่อบันทึก Cloud" : "บันทึกงานแล้ว");
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
    els.cloudNotice.hidden = false;
    els.cloudNotice.textContent = "Cloud ยังไม่พร้อมใช้งาน ผู้ดูแลต้องตั้งค่า Google Client ID และ Service Account บน Render";
    renderSavedProjects();
    return;
  }
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
    state.cloudUser = data.user || null;
    els.googleSignIn.hidden = true;
    els.cloudUser.hidden = false;
    els.cloudUserName.textContent = state.cloudUser.name || state.cloudUser.email;
    els.cloudNotice.hidden = true;
    renderSavedProjects();
    setStatus("เชื่อมต่อ Google Sheet แล้ว");
  } catch (error) {
    state.googleCredential = "";
    state.cloudUser = null;
    state.cloudProjects = [];
    sessionStorage.removeItem("material-calculator-google-credential");
    els.googleSignIn.hidden = false;
    els.cloudUser.hidden = true;
    els.cloudNotice.hidden = false;
    els.cloudNotice.textContent = error.message;
    renderSavedProjects();
  }
}

document.querySelectorAll(".app-tab").forEach((button) => {
  button.addEventListener("click", () => switchTab(button.dataset.tab));
});

els.saveProject.addEventListener("click", saveProject);
els.newProject.addEventListener("click", resetProject);
els.googleSignOut.addEventListener("click", () => {
  window.google?.accounts?.id?.disableAutoSelect();
  state.googleCredential = "";
  state.cloudUser = null;
  state.cloudProjects = [];
  sessionStorage.removeItem("material-calculator-google-credential");
  els.googleSignIn.hidden = false;
  els.cloudUser.hidden = true;
  renderSavedProjects();
  setStatus("ออกจากระบบ Google แล้ว");
});
els.savedProjectList.addEventListener("click", (event) => {
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
});

els.applyPages.addEventListener("click", () => {
  saveCurrentPageFromDom();
  const total = Math.max(1, Number.parseInt(els.totalPages.value || "1", 10));
  while (state.pages.length < total) state.pages.push([blankRow(state.department), blankRow(state.department)]);
  state.pages = state.pages.slice(0, total);
  state.currentPage = Math.min(state.currentPage, state.pages.length - 1);
  renderInputs();
  setStatus("กำหนดจำนวนหน้าแล้ว");
});

els.prevPage.addEventListener("click", () => {
  saveCurrentPageFromDom();
  state.currentPage = Math.max(0, state.currentPage - 1);
  renderInputs();
});

els.nextPage.addEventListener("click", () => {
  saveCurrentPageFromDom();
  state.currentPage = Math.min(state.pages.length - 1, state.currentPage + 1);
  renderInputs();
});

els.addRow.addEventListener("click", () => {
  saveCurrentPageFromDom();
  state.pages[state.currentPage].push(blankRow(state.department));
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
  state.pages[state.currentPage] = [blankRow(state.department), blankRow(state.department)];
  renderInputs();
  setStatus("ล้างข้อมูลหน้านี้แล้ว");
});

els.calculate.addEventListener("click", async () => {
  try {
    saveCurrentPageFromDom();
    setStatus("กำลังคำนวณ...");
    const data = await postJson("/api/calculate", { pages: state.pages });
    renderResults(data.items, `รวม ${data.summaryRows} รายการ จากข้อมูลที่เลือก ${data.inputRows} แถว`);
    setStatus("คำนวณสำเร็จ");
  } catch (error) {
    setStatus(error.message, true);
  }
});

els.expandSet.addEventListener("click", async () => {
  try {
    setStatus("กำลังแตก SET...");
    const data = await postJson("/api/expand-set");
    let meta = `รวม ${data.summaryRows} รายการ | พบ SET ${data.setFound} รายการ | แตกได้ ${data.expandedLines} แถว`;
    if (data.setMissing.length) {
      meta += ` | ไม่พบ: ${data.setMissing.slice(0, 6).join(", ")}`;
    }
    renderResults(data.items, meta);
    setStatus("แตก SET สำเร็จ");
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
    link.download = "material_summary_web.xlsx";
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
    link.download = "page_summary.xlsx";
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
    setStatus("กำลังสร้างสรุปลูกถ้วยและอุปกรณ์ยึดสาย...");
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
    link.download = "page_insulators_hardware.xlsx";
    link.click();
    URL.revokeObjectURL(url);
    setStatus("Export ลูกถ้วยและอุปกรณ์ยึดสายสำเร็จ");
  } catch (error) {
    setStatus(error.message, true);
  }
});

els.exportPageInsulators.addEventListener("click", async () => {
  try {
    saveCurrentPageFromDom();
    setStatus("กำลังสร้างสรุปลูกถ้วยแยกหน้า...");
    const response = await fetch("/api/export-page-insulators", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ pages: state.pages }),
    });
    if (!response.ok) {
      const data = await response.json();
      throw new Error(data.error || "Export ลูกถ้วยแยกหน้าไม่สำเร็จ");
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "page_insulators.xlsx";
    link.click();
    URL.revokeObjectURL(url);
    setStatus("Export ลูกถ้วยแยกหน้าสำเร็จ");
  } catch (error) {
    setStatus(error.message, true);
  }
});

els.exportPageCrossarms.addEventListener("click", async () => {
  try {
    saveCurrentPageFromDom();
    setStatus("กำลังนับคอนแยกตามหน้าและหัวเสา...");
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
    link.download = "page_crossarms.xlsx";
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

async function submitBaseRequest(event) {
  event.preventDefault();
  try {
    setStatus("กำลังส่งคำขอให้ Admin ตรวจ...");
    const replacing = els.requestAction.value === "replace";
    const size = replacing ? els.replaceSize.value : els.requestSize.value;
    const head = replacing ? els.replaceHead.value : els.requestHead.value;
    const response = await fetch("/api/base-requests", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        submitter_name: els.requesterName.value, employee_id: els.requesterEmployeeId.value,
        department: els.requesterDepartment.value, action: els.requestAction.value,
        target_department: els.requestTargetDepartment.value,
        size, head, rows: requestMaterialValues(), original_rows: replacing ? state.baseRequestOriginalRows : [], note: els.requestNote.value,
      }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "ส่งคำขอไม่สำเร็จ");
    els.requestSize.value = "";
    els.requestHead.value = "";
    els.requestNote.value = "";
    els.requestMaterialRows.innerHTML = "";
    state.baseRequestOriginalRows = [];
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
  saveCurrentPageFromDom();
  state.department = department || DEFAULT_DEPARTMENT;
  state.activeProjectId = "";
  state.pages = [[blankRow(state.department), blankRow(state.department)]];
  state.currentPage = 0;
  state.results = [];
  els.projectName.value = "";
  els.planNumber.value = "";
  try {
    await loadDepartmentSizes(state.department);
    renderInputs();
    renderResults([], "ยังไม่มีข้อมูล");
    els.saveHint.textContent = "ยังไม่ได้บันทึกงานนี้";
    setStatus(state.sizes.length ? `เปิดข้อมูล ${state.department} แล้ว` : `${state.department} ยังไม่มีข้อมูล เริ่มเพิ่มผ่านเมนูเสนอหัวเสาได้เลย`);
  } catch (error) {
    setStatus(error.message, true);
  }
}

async function switchRequestAction() {
  const replacing = els.requestAction.value === "replace";
  els.requestAddTarget.hidden = replacing;
  els.requestReplaceTarget.hidden = !replacing;
  els.existingDataHint.hidden = !replacing;
  els.requestMaterialRows.innerHTML = "";
  state.baseRequestOriginalRows = [];
  if (!replacing) {
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
    data.rows.forEach((row) => addRequestMaterialRow(row, true));
    els.existingDataHint.textContent = `โหลดข้อมูลเดิม ${data.rows.length} รายการแล้ว แก้ไข เพิ่ม หรือนำรายการออกได้`;
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
  if (state.adminToken) loadBaseRequests();
}

async function loadBaseRequests() {
  try {
    const data = await adminFetch("/api/base-requests/admin");
    renderBaseRequests(data.requests);
  } catch (error) {
    state.adminToken = "";
    sessionStorage.removeItem("material-calculator-admin-token");
    showAdminPanel();
    setStatus(error.message, true);
  }
}

function renderBaseRequests(requests) {
  els.baseRequestList.innerHTML = "";
  if (!requests.length) {
    const empty = document.createElement("div");
    empty.className = "empty-saved";
    empty.textContent = "ยังไม่มีคำขอแก้ไข BaseData";
    els.baseRequestList.appendChild(empty);
    return;
  }
  requests.forEach((request) => {
    const card = document.createElement("article");
    card.className = `request-card status-${request.status}`;
    const title = document.createElement("h3");
    title.textContent = `${request.size} · ${request.head}`;
    const meta = document.createElement("p");
    meta.textContent = `${request.targetDepartment || DEFAULT_DEPARTMENT} · ผู้เสนอ ${request.submitterName} · ${request.employeeId} · ${request.department} · ${request.action === "replace" ? "แทนที่ข้อมูลเดิม" : "เพิ่มข้อมูล"}`;
    const table = document.createElement("table");
    const isReplace = request.action === "replace";
    table.innerHTML = isReplace
      ? "<thead><tr><th>สถานะ</th><th>รายการวัสดุ</th><th>รหัส</th><th>เดิม</th><th>ใหม่</th></tr></thead>"
      : "<thead><tr><th>รายการวัสดุ</th><th>รหัส</th><th>จำนวน</th></tr></thead>";
    const body = document.createElement("tbody");
    const displayedRows = isReplace ? compareBaseRows(request.originalRows || [], request.rows) : request.rows;
    displayedRows.forEach((item) => {
      const row = document.createElement("tr");
      const values = isReplace
        ? [item.statusLabel, item.material, item.code, item.oldQuantity ?? "", item.newQuantity ?? ""]
        : [item.material, item.code, item.quantity];
      values.forEach((value, index) => {
        const cell = document.createElement("td"); cell.textContent = value; row.appendChild(cell);
        if (isReplace && index === 0) cell.className = `diff-${item.status}`;
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
        button.addEventListener("click", () => reviewBaseRequest(request.id, approve));
        footer.appendChild(button);
      }
    }
    card.append(title, meta);
    if (request.note) { const note = document.createElement("p"); note.textContent = `หมายเหตุ: ${request.note}`; card.appendChild(note); }
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

async function reviewBaseRequest(requestId, approve) {
  const note = window.prompt(approve ? "หมายเหตุการอนุมัติ (เว้นว่างได้)" : "เหตุผลที่ปฏิเสธ") ?? "";
  try {
    await adminFetch("/api/base-requests/review", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ requestId, approve, note }),
    });
    if (approve) {
      await loadDepartmentSizes(state.department);
      renderInputs();
    }
    setStatus(approve ? "อนุมัติและอัปเดต BaseData แล้ว" : "ปฏิเสธคำขอแล้ว");
    await loadBaseRequests();
  } catch (error) {
    setStatus(error.message, true);
  }
}

els.addRequestMaterial.addEventListener("click", () => addRequestMaterialRow());
els.baseRequestForm.addEventListener("submit", submitBaseRequest);
els.requestAction.addEventListener("change", switchRequestAction);
els.requestTargetDepartment.addEventListener("change", switchRequestAction);
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
  } catch (error) {
    setStatus(`โหลดฐานข้อมูลเริ่มต้นไม่สำเร็จ: ${error.message}`, true);
  }
}

initialize();
