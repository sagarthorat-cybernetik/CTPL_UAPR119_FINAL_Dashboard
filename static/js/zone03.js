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
  document.getElementById("searchBtn")?.addEventListener("click", () => loadPage(1));

  document.getElementById("prevPage")?.addEventListener("click", () => {
    if (state.page > 1) loadPage(state.page - 1);
  });

  document.getElementById("nextPage")?.addEventListener("click", () => {
    if (state.page < state.totalPages) loadPage(state.page + 1);
  });

  document.getElementById("exportBtn")?.addEventListener("click", startExport);
}

// === Default date range = today ===
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
  return `${y}-${m}-${d}T${hh}:${mm}`;
}

function getFilters() {
  return {
    start_date: document.getElementById("startDateTime").value,
    end_date: document.getElementById("endDateTime").value,
    barcode: document.getElementById("barcode")?.value?.trim() || "",
    station_name: document.getElementById("stationName")?.value || "",
     shift:document.getElementById("shift")?.value || ""
  };
}

// === Load paginated data ===
async function loadPage(page = 1) {
  const f = getFilters();
  state.lastFilters = f;
  showLoader();

  try {
    const res = await fetch("/fetch_data_zone03", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        ...f,
        page,
        limit: state.pageSize
      })
    });

    if (!res.ok) throw new Error("Failed to fetch");

    // 🔑 handle gzip OR plain JSON
    let result;
    const contentType = res.headers.get("Content-Type") || "";
    if (contentType.includes("application/gzip")) {
      const blob = await res.blob();
      const ds = new DecompressionStream("gzip");
      const decompressed = blob.stream().pipeThrough(ds);
      const text = await new Response(decompressed).text();
      result = JSON.parse(text);
    } else {
      result = await res.json();
    }

    if (result.error) throw new Error(result.error);

    state.page = result.page || 1;
    state.totalPages = result.pages || 1;
    rendersummary(result.total || 0, result.total_ok || 0, result.total_ng || 0);
    renderTable(result.data || [], result.columns || []);
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
  if (el) el.textContent = `Page ${state.page} of ${state.totalPages}`;
}

function rendersummary(total, total_ok, total_ng) {
  const totalEl = document.getElementById("total");
  const totalOkEl = document.getElementById("totalok");
  const totalNgEl = document.getElementById("totalng");
  if (totalEl) totalEl.textContent = total;
  if (totalOkEl) totalOkEl.textContent = total_ok;
  if (totalNgEl) totalNgEl.textContent = total_ng;
}

// === Render Table ===
function renderTable(data, columns) {
  const table = document.getElementById("dataTable");
    table.innerHTML = "";

    // --- Table Head ---
    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    columns.forEach(col => {
        const th = document.createElement("th");
        th.textContent = col;
        headRow.appendChild(th);
    });
    thead.appendChild(headRow);

    // --- Table Body ---
    const tbody = document.createElement("tbody");
    data.forEach(row => {
        const tr = document.createElement("tr");
        columns.forEach(col => {
            const td = document.createElement("td");
            td.textContent = row[col] !== null ? row[col] : "";
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });

    table.appendChild(thead);
    table.appendChild(tbody);
}

// === Export ===
async function startExport() {
  const f = state.lastFilters;
  showLoader();

  try {
    const res = await fetch("/export_excel_zone03", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(f)
    });

    if (!res.ok) throw new Error("Export failed");

    const blob = await res.blob();
    const url = window.URL.createObjectURL(blob);

    const cd = res.headers.get("Content-Disposition");
    let filename = "export.xlsx";
    if (cd && cd.includes("filename=")) {
      filename = cd.split("filename=")[1].replace(/["']/g, "");
    }

    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();

  } catch (err) {
    console.error("Error exporting:", err);
    alert("Export failed: " + err.message);
  } finally {
    hideLoader();
  }
}

// === Loader ===
function showLoader() {
  document.getElementById("logoProgress").style.display = "flex";
}
function hideLoader() {
  document.getElementById("logoProgress").style.display = "none";
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
// Navigate to combined statistics page
function combinedstatistics() {
  window.location = "/combinedstatistics";
}