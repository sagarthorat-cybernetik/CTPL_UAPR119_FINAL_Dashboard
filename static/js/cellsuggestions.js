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
let equalWidthChartInstance = null;
let kmeansChartInstance = null;

// Get Suggestions function
async function getSuggestions() {
    const startDateTime = document.getElementById('startDateTime').value;
    const endDateTime = document.getElementById('endDateTime').value;

    // Validate date inputs
    if (!startDateTime || !endDateTime) {
        showError('Please select both start and end date/time');
        return;
    }

    // Hide previous results and errors
    document.getElementById('resultsContainer').style.display = 'none';
    document.getElementById('errorMessage').style.display = 'none';
    document.getElementById('summaryStats').style.display = 'none';

    // Show loading indicator
    document.getElementById('loadingIndicator').style.display = 'block';

    try {
        const response = await fetch('/api/grade_suggestions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                start_date: startDateTime,
                end_date: endDateTime
            })
        });

        const data = await response.json();
//        console.log(data)
        if (!response.ok) {
            throw new Error(data.error || 'Failed to fetch suggestions');
        }

        // Hide loading indicator
        document.getElementById('loadingIndicator').style.display = 'none';

        // Check if we have data
        if (data.equal_width.total_cells === 0) {
            showError('No rejected cells found in the selected date range');
            return;
        }

        // Display results
        displayResults(data);

    } catch (error) {
        document.getElementById('loadingIndicator').style.display = 'none';
        showError('Error: ' + error.message);
        console.error('Error fetching suggestions:', error);
    }
}

function displayResults(data) {
    const equalWidth = data.equal_width;
    const kmeans = data.kmeans;

    // Show summary stats
    document.getElementById('summaryStats').style.display = 'block';
    document.getElementById('totalCells').textContent = equalWidth.total_cells;
    document.getElementById('outliersRemoved').textContent = equalWidth.ignored_outliers_count;

    // Display Equal Width results
    document.getElementById('equalAcceptedCount').textContent = equalWidth.accepted_count;
    document.getElementById('equalAcceptedPct').textContent = equalWidth.accepted_pct;
    populateTable('equalWidthTableBody', equalWidth.grades);
    createChart('equalWidthChart', equalWidth.grades, 'Equal Width Binning', 'equalWidth');

    // Display K-Means results
    if (kmeans.error) {
        document.getElementById('kmeansAcceptedCount').textContent = 'N/A';
        document.getElementById('kmeansAcceptedPct').textContent = 'N/A';
        const kmeansBody = document.getElementById('kmeansTableBody');
        kmeansBody.innerHTML = '<tr><td colspan="5" style="text-align: center; color: red;">K-Means unavailable: ' + kmeans.error + '</td></tr>';
    } else {
        document.getElementById('kmeansAcceptedCount').textContent = kmeans.accepted_count;
        document.getElementById('kmeansAcceptedPct').textContent = kmeans.accepted_pct;
        populateTable('kmeansTableBody', kmeans.grades);
        createChart('kmeansChart', kmeans.grades, 'K-Means Clustering', 'kmeans');
    }

    // Show results container
    document.getElementById('resultsContainer').style.display = 'block';
}

function populateTable(tableBodyId, grades) {
    const tbody = document.getElementById(tableBodyId);
    tbody.innerHTML = '';

    grades.forEach(grade => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${grade.grade_name}</td>
            <td>${grade.vmin}</td>
            <td>${grade.vmax}</td>
            <td>${grade.count}</td>
            <td>${grade.pct}%</td>
        `;
        tbody.appendChild(row);
    });
}

function createChart(canvasId, grades, title, chartType) {
    const ctx = document.getElementById(canvasId).getContext('2d');

    // Destroy existing chart if it exists
    if (chartType === 'equalWidth' && equalWidthChartInstance) {
        equalWidthChartInstance.destroy();
    }
    if (chartType === 'kmeans' && kmeansChartInstance) {
        kmeansChartInstance.destroy();
    }

    const labels = grades.map(g => g.grade_name);
    const counts = grades.map(g => g.count);
    const percentages = grades.map(g => g.pct);

    // Color palette
    const colors = [
        'rgba(255, 99, 132, 0.7)',
        'rgba(54, 162, 235, 0.7)',
        'rgba(255, 206, 86, 0.7)',
        'rgba(75, 192, 192, 0.7)',
        'rgba(153, 102, 255, 0.7)',
        'rgba(255, 159, 64, 0.7)'
    ];

    const borderColors = [
        'rgba(255, 99, 132, 1)',
        'rgba(54, 162, 235, 1)',
        'rgba(255, 206, 86, 1)',
        'rgba(75, 192, 192, 1)',
        'rgba(153, 102, 255, 1)',
        'rgba(255, 159, 64, 1)'
    ];

    const chartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Cell Count',
                data: counts,
                backgroundColor: colors.slice(0, grades.length),
                borderColor: borderColors.slice(0, grades.length),
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                title: {
                    display: true,
                    text: title,
                    font: {
                        size: 18
                    }
                },
                legend: {
                    display: true,
                    position: 'top'
                },
                tooltip: {
                    callbacks: {
                        afterLabel: function(context) {
                            const idx = context.dataIndex;
                            return `Percentage: ${percentages[idx]}%\nVoltage Range: ${grades[idx].vmin}V - ${grades[idx].vmax}V`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Number of Cells'
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Grade'
                    }
                }
            }
        }
    });

    // Store chart instance
    if (chartType === 'equalWidth') {
        equalWidthChartInstance = chartInstance;
    } else {
        kmeansChartInstance = chartInstance;
    }
}

function showError(message) {
    document.getElementById('errorText').textContent = message;
    document.getElementById('errorMessage').style.display = 'block';
}

