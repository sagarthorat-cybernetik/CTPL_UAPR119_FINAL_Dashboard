let state = {
  page: 1,
  pageSize: 100,
  totalPages: 1,
  charts: { main: null, ok: null, ng: null },
  lastFilters: {}
};

document.addEventListener("DOMContentLoaded", () => {
  setDefaultDates();
  attachUI();
  loadPage(1);
});

function attachUI() {
  // Example: hook up a form apply button if you have one, else call loadPage on any input change
  const applyBtn = document.getElementById("applyBtn");
  if (applyBtn) applyBtn.addEventListener("click", () => loadPage(1));

  const exportBtn = document.getElementById("exportBtn");
  if (exportBtn) exportBtn.addEventListener("click", startExport);

  // Simple pagination hooks (expects #pagination container)
  document.getElementById("pagination")?.addEventListener("click", (e) => {
    if (e.target.dataset.page) {
      const p = parseInt(e.target.dataset.page, 10);
      if (!isNaN(p) && p >= 1 && p <= state.totalPages) loadPage(p);
    }
  });
}

// Defaults to today's range
function setDefaultDates() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0);
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  document.getElementById("startDateTime").value = formatForInput(start);
  document.getElementById("endDateTime").value = formatForInput(end);
}

function formatForInput(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mm = String(date.getMinutes()).padStart(2, "0");
  const ss = String(date.getSeconds()).padStart(2, "0");
  const ms = String(date.getMilliseconds()).padStart(3, "0");
  return `${y}-${m}-${d} ${hh}:${mm}:${ss}.${ms}`;
}

function getFilters() {
  return {
    start_date: document.getElementById("startDateTime").value,
    end_date: document.getElementById("endDateTime").value,
    barcode: document.getElementById("barcode")?.value?.trim() || "",
    barleyStatus: document.getElementById("barleyStatus")?.value || "",
    capacityStatus: document.getElementById("capacityStatus")?.value || "",
    measurementStatus: document.getElementById("measurementStatus")?.value || "",
    finalStatus: document.getElementById("finalStatus")?.value || "",
    grade: document.getElementById("grade")?.value || ""
  };
}

async function loadPage(page = 1) {
  const f = getFilters();
  state.lastFilters = f;

  const params = new URLSearchParams({
    ...f,
    page: String(page),
    page_size: String(state.pageSize)
  });

  const res = await fetch(`/api/cell_dashboard?${params.toString()}`);
  const payload = await res.json();
  if (!res.ok) {
    console.error("Load failed:", payload);
    return;
  }

  // Update pagination & data
  state.page = payload.page;
  state.totalPages = payload.total_pages || 1;

  renderStats(payload.stats || {});
  renderTable(payload.rows || []);
  renderPagination();

  // charts
  updateCharts(payload.stats || {});
}

function renderPagination() {
  const el = document.getElementById("pagination");
  if (!el) return;

  const p = state.page;
  const tp = state.totalPages;
  const html = [];

  const mkBtn = (label, page, disabled=false, active=false) => {
    return `<button class="page-btn ${active ? "active" : ""}" data-page="${page}" ${disabled ? "disabled" : ""}>${label}</button>`;
  };

  html.push(mkBtn("«", 1, p === 1));
  html.push(mkBtn("‹", Math.max(1, p - 1), p === 1));

  // windowed pages
  const start = Math.max(1, p - 2);
  const end = Math.min(tp, p + 2);
  for (let i = start; i <= end; i++) {
    html.push(mkBtn(String(i), i, false, i === p));
  }

  html.push(mkBtn("›", Math.min(tp, p + 1), p === tp));
  html.push(mkBtn("»", tp, p === tp));

  el.innerHTML = html.join("");
}



function renderStats(s) {
  const byId = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.textContent = val ?? 0;
  };
  byId("totalCells", s.totalCells);
  byId("okCells", s.okCells);
  byId("okCellsG1", s.okCellsG1);
  byId("okCellsG2", s.okCellsG2);
  byId("okCellsG3", s.okCellsG3);
  byId("okCellsG4", s.okCellsG4);
  byId("okCellsG5", s.okCellsG5);
  byId("okCellsG6", s.okCellsG6);
  byId("tngCells", s.tngCells);
  byId("bngCells", s.bngCells);
  byId("vngCells", s.vngCells);
  byId("ingCells", s.ingCells);
  byId("vingCells", s.vingCells);
  byId("cngCells", s.cngCells);
  byId("bpaperngCells", s.bpaperngCells);
  byId("dpngCells", s.dpngCells);
}

function updateCharts(stats) {
  // main pie
  const ctxMain = document.getElementById("mainPieChart")?.getContext("2d");
  if (ctxMain) {
    if (!state.charts.main) {
      state.charts.main = new Chart(ctxMain, {
        type: "pie",
        data: {
          labels: ["OK Cells", "NG Cells"],
          datasets: [{ data: [stats.okCells || 0, stats.tngCells || 0] }]
        },
        options: { responsive: true, plugins: { legend: { position: "top" }, title: { display: true, text: "Overall OK vs NG Cells" } } }
      });
    } else {
      state.charts.main.data.datasets[0].data = [stats.okCells || 0, stats.tngCells || 0];
      state.charts.main.update();
    }
  }

  // ok per grade
  const ctxOk = document.getElementById("okbarChart")?.getContext("2d");
  const okData = [
    stats.okCellsG1 || 0,
    stats.okCellsG2 || 0,
    stats.okCellsG3 || 0,
    stats.okCellsG4 || 0,
    stats.okCellsG5 || 0,
    stats.okCellsG6 || 0
  ];
  if (ctxOk) {
    if (!state.charts.ok) {
      state.charts.ok = new Chart(ctxOk, {
        type: "bar",
        data: { labels: ["Gear1","Gear2","Gear3","Gear4","Gear5","Gear6"], datasets: [{ label: "OK Cells Per Gear", data: okData }] },
        options: { responsive: true, plugins: { legend: { position: "top" }, title: { display: true, text: "OK Cells Per Gear" } } }
      });
    } else {
      state.charts.ok.data.datasets[0].data = okData;
      state.charts.ok.update();
    }
  }

  // ng reasons
  const ctxNg = document.getElementById("ngbarChart")?.getContext("2d");
  const ngData = [
    stats.bngCells || 0,
    stats.vngCells || 0,
    stats.ingCells || 0,
    stats.vingCells || 0,
    stats.cngCells || 0,
    stats.bpaperngCells || 0,
    stats.dpngCells || 0
  ];
  if (ctxNg) {
    if (!state.charts.ng) {
      state.charts.ng = new Chart(ctxNg, {
        type: "bar",
        data: { labels: ["Barcode","Voltage","Resistance","Voltage & Resistance","Capacity","Barley Paper","Duplicate"], datasets: [{ label: "NG Cells", data: ngData }] },
        options: { responsive: true, plugins: { legend: { position: "top" }, title: { display: true, text: "NG Cell Per Reason" } } }
      });
    } else {
      state.charts.ng.data.datasets[0].data = ngData;
      state.charts.ng.update();
    }
  }
}

async function startExport() {
  // Show progress ring overlay
  showLoader(0);

  const body = { ...state.lastFilters };

  const res = await fetch("/api/export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  const data = await res.json();
  
  if (!res.ok) {
    hideLoader();
    alert("Failed to start export");
    return;
  }

  const taskId = data.task_id;
  // poll status
  const timer = setInterval(async () => {
    const sres = await fetch(`/api/export/status?task_id=${taskId}`);
    const sdata = await sres.json();
    
    if (!sres.ok) {
      clearInterval(timer);
      hideLoader();
      alert("Export status failed");
      return;
    }
    const prog = Number(sdata.progress || 0);
    showLoader(prog);

    if (sdata.done) {
      clearInterval(timer);
      if (sdata.error) {
        hideLoader();
        alert("Export failed: " + sdata.error);
        return;
      }
      // trigger download
      window.location = `/api/export/download?task_id=${taskId}`;
      setTimeout(hideLoader, 1200);
    }
  }, 800);
}
function showLoader(percent) {
  const box = document.getElementById("logoProgress");
  if (!box) return;
  box.style.setProperty("--p", `${percent}%`);
  const txt = document.getElementById("progressText");
  if (txt) txt.textContent = `${percent}%`;
  box.style.display = "block";
}
function hideLoader() {
  const box = document.getElementById("logoProgress");
  if (box) box.style.display = "block";
}


// Navigate to zone01 cell dashboard (kept)
function celldashboard() {
  window.location = "/";
}
// Navigate to zone01 model dashboard (kept)
function modeldashboard() {
  window.location = "modeldashboard";
}
// Navigate to zone01 dashboard (kept)
function zone01dashboard() {
  window.location = "/";
}
// Navigate to zone02 dashboard (kept)
function zone02dashboard() {
  window.location = "/zone02_dashboard";
}
// Navigate to model dashboard (kept)
function zone03dashboard() {
  window.location = "zone03_dashboard";
}

document.addEventListener("DOMContentLoaded", () => {
  setDefaultDates();
  attachUI();
  loadPage(1);
});

function attachUI() {
  document.getElementById("searchBtn")?.addEventListener("click", () => loadPage(1));
  document.getElementById("prevPage")?.addEventListener("click", () => {
    if (state.page > 1) loadPage(state.page - 1);
  });
  document.getElementById("nextPage")?.addEventListener("click", () => {
    if (state.page < state.totalPages) loadPage(state.page + 1);
  });
  document.getElementById("exportBtn")?.addEventListener("click", startExport);
}

// Default date range = today
function setDefaultDates() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0, 0);
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
  document.getElementById("startDateTime").value = formatForInput(start);
  document.getElementById("endDateTime").value = formatForInput(end);
}

function formatForInput(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mm = String(date.getMinutes()).padStart(2, "0");
  const ss = String(date.getSeconds()).padStart(2, "0");
  return `${y}-${m}-${d} ${hh}:${mm}:${ss}`;
}

function getFilters() {
  return {
    start_date: document.getElementById("startDateTime").value,
    end_date: document.getElementById("endDateTime").value,
    barcode: document.getElementById("barcode")?.value?.trim() || "",
    barleyStatus: document.getElementById("barleyStatus")?.value || "",
    capacityStatus: document.getElementById("capacityStatus")?.value || "",
    measurementStatus: document.getElementById("measurementStatus")?.value || "",
    finalStatus: document.getElementById("finalStatus")?.value || "",
    grade: document.getElementById("grade")?.value || ""
  };
}

async function loadPage(page = 1) {
  const f = getFilters();
  state.lastFilters = f;
  showLoader();

  try {
    const params = new URLSearchParams({
      ...f,
      page: String(page),
      page_size: String(state.pageSize)
    });

    const res = await fetch(`/api/cell_dashboard?${params.toString()}`);
    const payload = await res.json();
    if (!res.ok) throw new Error(payload.error || "Load failed");

    // Update pagination & data
    state.page = payload.page;
    state.totalPages = payload.total_pages || 1;

    renderStats(payload.stats || {});
    renderTable(payload.rows || []);
    updateCharts(payload.stats || {});
    renderPageInfo();

  } catch (err) {
    console.error("Error loading page:", err);
    alert("Failed: " + err.message);
  } finally {
    hideLoader();
  }
}

function renderPageInfo() {
  const el = document.getElementById("pageInfo");
  if (!el) return;
  el.textContent = `Page ${state.page} of ${state.totalPages}`;
  const percent = Math.round((state.page / state.totalPages) * 100);

}

function renderTable(rows) {

  const tbody = document.querySelector("#dataTable tbody");
  if (!tbody) return;

  const frag = document.createDocumentFragment();

  if (!rows.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 22;
    td.textContent = "No data available";
    tr.appendChild(td);
    frag.appendChild(tr);
  } else {
    rows.forEach((row, idx) => {
      const tr = document.createElement("tr");

      // Date formatting (kept close to your previous format)
      let dt = row.Date_Time ? new Date(row.Date_Time) : null;
      let dtStr = dt ? formatForInput(dt) : "-";

      tr.innerHTML = `
        <td>${(idx + 1) + (state.page - 1) * state.pageSize}</td>
        <td>${dtStr}</td>
        <td>${row.Shift ?? "-"}</td>
        <td>${row.Operator ?? "-"}</td>
        <td>${row.Cell_Position ?? "-"}</td>
        <td>${row.Cell_Barcode ?? "-"}</td>
        <td>${row.Cell_Barley_Paper_Positive ?? "-"}</td>
        <td>${row.Cell_Barley_Paper_Negative ?? "-"}</td>
        <td>${row.Cell_Barley_Paper_Status ?? "-"}</td>
       <td>${
            row.Cell_Capacity_Min_Set_Value != null && !isNaN(row.Cell_Capacity_Min_Set_Value)
                ? Number(row.Cell_Capacity_Min_Set_Value).toFixed(1)
                : "-"
        }</td>

        <td>${
            row.Cell_Capacity_Max_Set_Value != null && !isNaN(row.Cell_Capacity_Max_Set_Value)
                ? Number(row.Cell_Capacity_Max_Set_Value).toFixed(1)
                : "-"
        }</td>

        <td>${
            row.Cell_Capacity_Actual != null && !isNaN(row.Cell_Capacity_Actual)
                ? Number(row.Cell_Capacity_Actual).toFixed(1)
                : "-"
        }</td>
        <td>${row.Cell_Capacity_Status ?? "-"}</td>
        <td>${row.Cell_Voltage_Min_Set_Value ?? "-"}</td>
        <td>${row.Cell_Voltage_Max_Set_Value ?? "-"}</td>
        <td>${row.Cell_Voltage_Actual ?? "-"}</td>
        <td>${row.Cell_Resistance_Min_Set_Value ?? "-"}</td>
        <td>${row.Cell_Resistance_Max_Set_Value ?? "-"}</td>
        <td>${row.Cell_Resistance_Actual ?? "-"}</td>
        <td>${row.Cell_Measurement_Status ?? "-"}</td>
        <td>${row.Cell_Final_Status ?? "-"}</td>
        <td>${row.Cell_Grade ?? "-"}</td>
        <td>${row.Cell_Fail_Reason ?? "-"}</td>
      `;
      frag.appendChild(tr);
    });
  }

  tbody.innerHTML = "";
  tbody.appendChild(frag);
}
/* Loader around logo */
function showLoader() {
  const loader = document.getElementById("logoProgress");
  loader.style.display = "flex";
  // loader.dataset.intervalId = interval;
}

function hideLoader() {
  const loader = document.getElementById("logoProgress");
  clearInterval(loader.dataset.intervalId);
  setTimeout(() => {
    loader.style.display = "none";
  }, 500);
}
