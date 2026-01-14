let state = {
  page: 1,
  pageSize: 100,
  totalPages: 1,
  lastFilters: {}
};

document.addEventListener("DOMContentLoaded", () => {
  setDefaultDates();
  attachUI();
  loadPage(1);
});

function attachUI() {
  // Search button
  document.getElementById("searchBtn")?.addEventListener("click", () => loadPage(1));

  // Export button
  document.getElementById("exportExcel")?.addEventListener("click", startExport);
  
  // Pagination
  document.getElementById("prevPage")?.addEventListener("click", () => {
    if (state.page > 1) loadPage(state.page - 1);
  });
  document.getElementById("nextPage")?.addEventListener("click", () => {
    if (state.page < state.totalPages) loadPage(state.page + 1);
  });
}

// --- Default Dates = today ---
function setDefaultDates() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
  const end = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
  document.getElementById("startDateTime").value = formatForInput(start);
  document.getElementById("endDateTime").value = formatForInput(end);
}

function formatForInput(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mm = String(date.getMinutes()).padStart(2, "0");
  return `${y}-${m}-${d}T${hh}:${mm}`;
}

function getFilters() {
  return {
    start_date: document.getElementById("startDateTime").value,
    end_date: document.getElementById("endDateTime").value,
    moduleid: document.getElementById("moduleid")?.value?.trim() || "",
    grade: document.getElementById("grade")?.value || ""
  };
}


// --- Show/hide loader ---
function showLoader(show = true) {
  const loader = document.getElementById("logoProgress");
  if (loader) loader.style.display = show ? "flex" : "none";
}


// --- Load paged data ---
async function loadPage(page = 1) {
  const f = getFilters();
  state.lastFilters = f;

  try {
        showLoader(true);
    const params = new URLSearchParams({
      ...f,
      page: String(page),
      page_size: String(state.pageSize)
    });

    const res = await fetch(`/api/module_dashboard?${params.toString()}`);
    const payload = await res.json();
    if (!res.ok) throw new Error(payload.error || "Load failed");

    state.page = payload.page;
    state.totalPages = payload.total_pages || 1;
    
    renderTable(payload.rows || []);
    renderPagination();
   rendersummary(payload.total_module || 0, payload.total_ok || 0, payload.total_ng || 0, payload.total_inprogress || 0);

  } catch (err) {
    console.error("Error loading page:", err);
    alert("Failed: " + err.message);
  }
  finally {
    showLoader(false);
  }
}
function rendersummary(total, total_ok, total_ng,total_inprogress) {
  const totalEl = document.getElementById("total");
  const totalOkEl = document.getElementById("totalok");
  const totalNgEl = document.getElementById("totalng");
  const total_Inprogress = document.getElementById("totalinprogress");
  if (totalEl) totalEl.textContent = total;
  if (totalOkEl) totalOkEl.textContent = total_ok;
  if (totalNgEl) totalNgEl.textContent = total_ng;
  if (total_inprogress) total_Inprogress.textContent = total_inprogress;
}
// --- Render table rows ---
function renderTable(rows) {
  const tbody = document.querySelector("#dataTable tbody");
  if (!tbody) return;

  const frag = document.createDocumentFragment();

  if (!rows.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = 14;
    td.textContent = "No data available";
    tr.appendChild(td);
    frag.appendChild(tr);
  } else {
    rows.forEach((row, idx) => {
      const tr = document.createElement("tr");

      let dt = row.Date_Time ? new Date(row.Date_Time) : null;
      let dtStr = dt ? dt.toLocaleString() : "-";

      tr.innerHTML = `
        <td>${(idx + 1) + (state.page - 1) * state.pageSize}</td>
        <td>${dtStr}</td>
        <td>${row.Shift ?? "-"}</td>
        <td>${row.Operator ?? "-"}</td>
        <td>${row.Module_Type ?? "-"}</td>
        <td>${row.Module_Grade ?? "-"}</td>
        <td>${row.Module_ID ?? "-"}</td>
        <td>${row.Cell_ID ?? "-"}</td>
        <td>${row.Cell_Capacity_Actual ?? "-"}</td>
        <td>${row.Cell_Voltage_Actual ?? "-"}</td>
        <td>${row.Cell_Resistance_Actual ?? "-"}</td>
        <td>${(row.Module_Capacity_Max - row.Module_Capacity_Min).toFixed(4)  ?? "-"}</td>
        <td>${((row.Module_Voltage_Max - row.Module_Voltage_Min)*1000).toFixed(4)  ?? "-"}</td>
        <td>${(row.Module_Resistance_Max - row.Module_Resistance_Min).toFixed(4)  ?? "-"}</td>
        <td>${row.Module_Capacity_Range ?? "-"}</td>
        <td>${row.Module_Capacity_Name ?? "-"}</td>
        <td>${row.Status ?? "-"}</td>
      `;
      frag.appendChild(tr);
    });
  }

  tbody.innerHTML = "";
  tbody.appendChild(frag);
}

// --- Render pagination info ---
function renderPagination() {
  const info = document.getElementById("pageInfo");
  if (info) {
    info.textContent = `Page ${state.page} of ${state.totalPages}`;
  }
}
// --- Export ---
async function startExport() {
        showLoader(true);
  const body = { ...state.lastFilters };

  const res = await fetch("/api/module_export", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body)
  });
  const data = await res.json();

  if (!res.ok) {
    alert("Failed to start export");
    return;
  }

  const taskId = data.task_id;

  // Polling
  const timer = setInterval(async () => {
    const sres = await fetch(`/api/module_export/status?task_id=${taskId}`);
    const sdata = await sres.json();

    if (!sres.ok) {
      clearInterval(timer);
      alert("Export status failed");
      return;
    }

    if (sdata.done) {
      clearInterval(timer);
      if (sdata.error) {
        alert("Export failed: " + sdata.error);
        return;
      }
      window.location = `/api/module_export/download?task_id=${taskId}`;
          showLoader(false);
    }
  }, 1000);
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
// Navigate to cell suggestions page
function cellsuggestions() {
  window.location = "/cellsuggestions";
}
// Navigate to combined statistics page
function combinedstatistics() {
  window.location = "/combinedstatistics";
}