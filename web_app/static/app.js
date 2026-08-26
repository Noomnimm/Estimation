const state = {
  sizes: [],
  pages: [[blankRow(), blankRow()]],
  currentPage: 0,
  results: [],
};

const els = {
  status: document.getElementById("status"),
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
  resultRows: document.getElementById("resultRows"),
  resultMeta: document.getElementById("resultMeta"),
};

function blankRow() {
  return { size: "", head: "", count: "", wire1: "", wire2: "" };
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
      page[index].size = sizeSelect.value;
      page[index].head = "";
      page[index].wire1 = "";
      page[index].wire2 = "";
      renderInputs();
    });
    headSelect.addEventListener("change", () => {
      page[index].head = headSelect.value;
      page[index].wire1 = "";
      page[index].wire2 = "";
      renderInputs();
    });
    countInput.addEventListener("input", () => {
      page[index].count = countInput.value;
    });

    els.inputRows.appendChild(tr);
    const wireKind = classifyWireHead(row.head);
    if (wireKind) {
      els.inputRows.appendChild(createWireDetailsRow(row, index, wireKind));
    }
    if (row.size) {
      loadHeads(row.size, headSelect, row.head);
    }
  });
  renderPageControls();
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
  const normalized = String(head || "").trim().toUpperCase();
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

async function loadHeads(size, select, selected) {
  if (!size) {
    fillSelect(select, [], "เลือกรหัสหัวเสา");
    return;
  }
  try {
    const response = await fetch(`/api/heads?size=${encodeURIComponent(size)}`);
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

els.applyPages.addEventListener("click", () => {
  saveCurrentPageFromDom();
  const total = Math.max(1, Number.parseInt(els.totalPages.value || "1", 10));
  while (state.pages.length < total) state.pages.push([blankRow(), blankRow()]);
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
  state.pages[state.currentPage].push(blankRow());
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
  state.pages[state.currentPage] = [blankRow(), blankRow()];
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

async function initialize() {
  renderInputs();
  try {
    const response = await fetch("/api/status");
    const data = await readJson(response);
    if (data.base) {
      state.sizes = data.base.sizes;
    }
    renderInputs();
    setStatus(data.base ? "โหลดฐานข้อมูลเริ่มต้นแล้ว" : "กรุณาโหลด BaseData");
  } catch (error) {
    setStatus(`โหลดฐานข้อมูลเริ่มต้นไม่สำเร็จ: ${error.message}`, true);
  }
}

initialize();
