// Navigation functions
function zone01dashboard() {
    window.location.href = "/";
}

function zone02dashboard() {
    window.location.href = "/zone02_dashboard";
}

function zone03dashboard() {
    window.location.href = "/zone03_dashboard";
}
function combinedstatistics() {
    window.location.href = "/combinedstatistics";
}

document.addEventListener("DOMContentLoaded", () => {
  setDefaultDates();
  fillConfigInputs();

});
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

// Chart instances
let irChartInstance = null;
let voltageChartInstance = null;

async function fillConfigInputs() {
     const response = await fetch('/api/grade_config', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },

    });

    const config = await response.json();
//    console.log("API DATA:", config);

    if (!response.ok) {
        throw new Error(data.error || 'API error');
    }
    document.getElementById('ir_bin_width').value = config.ir_bin_width;
    document.getElementById('ir_underflow').value = config.ir_underflow;
    document.getElementById('ir_overflow').value = config.ir_overflow;

    document.getElementById('voltage_bin_width').value = config.voltage_bin_width;
    document.getElementById('voltage_underflow').value = config.voltage_underflow;
    document.getElementById('voltage_overflow').value = config.voltage_overflow;
}

// ---------------- FETCH DATA ----------------
async function getSuggestions() {
    const startDateTime = document.getElementById('startDateTime').value;
    const endDateTime = document.getElementById('endDateTime').value;

    if (!startDateTime || !endDateTime) {
        showError('Please select both start and end date/time');
        return;
    }

    document.getElementById('loadingIndicator').style.display = 'block';
    document.getElementById('errorMessage').style.display = 'none';
    document.getElementById('resultsContainer').style.display = 'none';

    try {
        // 🔹 Read config inputs
        const payload = {
            start_date: startDateTime,
            end_date: endDateTime,

            ir_bin_width: parseFloat(document.getElementById('ir_bin_width').value),
            ir_underflow: parseFloat(document.getElementById('ir_underflow').value),
            ir_overflow: parseFloat(document.getElementById('ir_overflow').value),

            voltage_bin_width: parseFloat(document.getElementById('voltage_bin_width').value),
            voltage_underflow: parseFloat(document.getElementById('voltage_underflow').value),
            voltage_overflow: parseFloat(document.getElementById('voltage_overflow').value)
        };
        const response = await fetch('/api/grade_suggestions', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        const data = await response.json();
//        console.log("API DATA:", data);

        if (!response.ok) {
            throw new Error(data.error || 'API error');
        }

        document.getElementById('loadingIndicator').style.display = 'none';

        if (data.final_results.total_cells === 0) {
            showError('No rejected cells found');
            return;
        }

        displayResults(data.final_results);

    } catch (err) {
        document.getElementById('loadingIndicator').style.display = 'none';
        showError(err.message);
        console.error(err);
    }
}

// ---------------- DISPLAY RESULTS ----------------
function displayResults(results) {

    // Summary
    document.getElementById('summaryStats').style.display = 'block';
    document.getElementById('totalCells').textContent = results.total_cells;
    document.getElementById('outliersRemoved').textContent = results.ignored_outliers_count;

    // Tables
    populateHistogramTable(
        'irTableBody',
        results.bin_edges_ir,
        results.hist_ir
    );

    populateHistogramTable(
        'voltageTableBody',
        results.bin_edges_voltage,
        results.hist_voltage
    );

    // Charts
    createHistogramChart(
        'irChart',
        results.bin_edges_ir,
        results.hist_ir,
        'IR Histogram',
        'Ω',
        'ir'
    );

    createHistogramChart(
        'voltageChart',
        results.bin_edges_voltage,
        results.hist_voltage,
        'Voltage Histogram',
        'V',
        'voltage'
    );

    document.getElementById('resultsContainer').style.display = 'block';
}

// ---------------- TABLE ----------------
function populateHistogramTable(tableBodyId, binEdges, hist) {
    const tbody = document.getElementById(tableBodyId);
    tbody.innerHTML = '';

    for (let i = 0; i < hist.length; i++) {
//        const rangeLabel = `${binEdges[i]} – ${binEdges[i + 1] ?? ''}`;

        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${binEdges[i]}</td>
            <td>${hist[i]}</td>
        `;
        tbody.appendChild(row);
    }
}

// ---------------- CHART ----------------
function createHistogramChart(canvasId, binEdges, hist, title, unit, type) {
    const ctx = document.getElementById(canvasId).getContext('2d');

    const labels = [];
    for (let i = 0; i < hist.length; i++) {
        labels.push(`${binEdges[i]}`);
    }

    // Destroy previous instance
    if (type === 'ir' && irChartInstance) irChartInstance.destroy();
    if (type === 'voltage' && voltageChartInstance) voltageChartInstance.destroy();

    const chart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: `Cell Count (${unit})`,
                data: hist,
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            plugins: {
                title: {
                    display: true,
                    text: title
                },
                legend: { display: false }
            },
            scales: {
                x: {
                    title: { display: true, text: 'Range' }
                },
                y: {
                    beginAtZero: true,
                    title: { display: true, text: 'Cell Count' }
                }
            }
        }
    });

    if (type === 'ir') irChartInstance = chart;
    if (type === 'voltage') voltageChartInstance = chart;
}

// ---------------- ERROR ----------------
function showError(message) {
    document.getElementById('errorText').textContent = message;
    document.getElementById('errorMessage').style.display = 'block';
}
