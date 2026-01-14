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

// Current statistics data
let currentStats = null;
let currentZone = null;

// Set default dates on load
document.addEventListener('DOMContentLoaded', function() {
    setDefaultDates();

    // Add zone change listener
    document.getElementById('zoneSelect').addEventListener('change', function() {
        hideAllContent();
    });
});

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

function hideAllContent() {
    document.getElementById('zone1Content').style.display = 'none';
    document.getElementById('zone2Content').style.display = 'none';
    document.getElementById('zone3Content').style.display = 'none';
    document.getElementById('exportZoneBtn').style.display = 'none';
    document.getElementById('errorMessage').style.display = 'none';
}

async function getStatistics() {
    const zone = document.getElementById('zoneSelect').value;
    const startDateTime = document.getElementById('startDateTime').value;
    const endDateTime = document.getElementById('endDateTime').value;

    // Validate inputs
    if (!startDateTime || !endDateTime) {
        showError('Please select both start and end date/time');
        return;
    }

    // Hide previous content
    hideAllContent();

    // Show loading
    document.getElementById('loadingIndicator').style.display = 'block';

    try {
        const response = await fetch('/api/combined_statistics', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                zone: zone,
                start_date: startDateTime,
                end_date: endDateTime
            })
        });

        const data = await response.json();

        if (!response.ok) {
            throw new Error(data.error || 'Failed to fetch statistics');
        }

        // Store current data
        currentStats = data;
        currentZone = zone;

        // Hide loading
        document.getElementById('loadingIndicator').style.display = 'none';

        // Display based on zone
        if (zone === 'zone1' ) {
            displayZone1Stats(data);
        } else if (zone === 'zone2') {
            displayZone2Stats(data);
        } else if (zone === 'zone3') {
            displayZone3Stats(data);
        }

        // Show export button for current zone
        document.getElementById('exportZoneBtn').style.display = 'inline-block';

    } catch (error) {
        document.getElementById('loadingIndicator').style.display = 'none';
        showError('Error: ' + error.message);
        console.error('Error fetching statistics:', error);
    }
}

function displayZone1Stats(data) {
    try{
        // Display cell stats
        document.getElementById('z1TotalCells').textContent = data.cells.total.toLocaleString();
        document.getElementById('z1OkCells').textContent = data.cells.ok.toLocaleString();
        document.getElementById('z1NgCells').textContent = data.cells.ng.toLocaleString();

        // Display module stats
        document.getElementById('z1TotalModules').textContent = data.modules.total.toLocaleString();
        document.getElementById('z1OkModules').textContent = data.modules.ok.toLocaleString();
        document.getElementById('z1NgModules').textContent = data.modules.ng.toLocaleString();
        document.getElementById('z1InProgressModules').textContent = data.modules.inprogress.toLocaleString();

        // Show zone 1 content
        document.getElementById('zone1Content').style.display = 'block';
    }catch (error) {
        document.getElementById('loadingIndicator').style.display = 'none';
        showError('No Data Found For This Date Range : ' + error.message);
        console.error('No Data Found For This Date Range : ', error);
    }
}

function displayZone2Stats(data) {
try{
    const tbody = document.getElementById('zone2TableBody');
    tbody.innerHTML = '';

    let totalSum = 0, okSum = 0, ngSum = 0, avgcytime=0;

    data.stations.forEach(station => {
        const row = document.createElement('tr');
        const okPercent = station.total > 0 ? ((station.ok / station.total) * 100).toFixed(2) : '0.00';

        row.innerHTML = `
            <td>${station.station.replace(/_/g, ' ')}</td>
            <td>${station.total.toLocaleString()}</td>
            <td style="color: #4BC0C0;">${station.ok.toLocaleString()}</td>
            <td style="color: #FF6384;">${station.ng.toLocaleString()}</td>
            <td style="color: #FF6384;">${station.avgcytime.toLocaleString()}</td>

            <td>${okPercent}%</td>
        `;
        tbody.appendChild(row);

        totalSum += station.total;
        okSum += station.ok;
        ngSum += station.ng;

    });

    // Update totals
    const totalOkPercent = totalSum > 0 ? ((okSum / totalSum) * 100).toFixed(2) : '0.00';
    document.getElementById('z2TotalSum').textContent = totalSum.toLocaleString();
    document.getElementById('z2OkSum').textContent = okSum.toLocaleString();
    document.getElementById('z2NgSum').textContent = ngSum.toLocaleString();
    document.getElementById('z2avgcytime').textContent = "-";
    document.getElementById('z2OkPercent').textContent = totalOkPercent + '%';

    // Show zone 2 content
    document.getElementById('zone2Content').style.display = 'block';
    }catch (error) {
        document.getElementById('loadingIndicator').style.display = 'none';
        showError('No Data Found For This Date Range : ' + error.message);
        console.error('No Data Found For This Date Range : ', error);
    }
}

function displayZone3Stats(data) {
try{
    const tbody = document.getElementById('zone3TableBody');
    tbody.innerHTML = '';

    let totalSum = 0, okSum = 0, ngSum = 0;

    data.stations.forEach(station => {
        const row = document.createElement('tr');
        const okPercent = station.total > 0 ? ((station.ok / station.total) * 100).toFixed(2) : '0.00';

        row.innerHTML = `
            <td>${station.station.replace(/_/g, ' ')}</td>
            <td>${station.total.toLocaleString()}</td>
            <td style="color: #4BC0C0;">${station.ok.toLocaleString()}</td>
            <td style="color: #FF6384;">${station.ng.toLocaleString()}</td>
            <td style="color: #FF6384;">${station.avgcytime.toLocaleString()}</td>

            <td>${okPercent}%</td>
        `;
        tbody.appendChild(row);

        totalSum += station.total;
        okSum += station.ok;
        ngSum += station.ng;
    });

    // Update totals
    const totalOkPercent = totalSum > 0 ? ((okSum / totalSum) * 100).toFixed(2) : '0.00';
    document.getElementById('z3TotalSum').textContent = totalSum.toLocaleString();
    document.getElementById('z3OkSum').textContent = okSum.toLocaleString();
    document.getElementById('z3NgSum').textContent = ngSum.toLocaleString();
    document.getElementById('z3avgcytime').textContent = "-";
    document.getElementById('z3OkPercent').textContent = totalOkPercent + '%';

    // Show zone 3 content
    document.getElementById('zone3Content').style.display = 'block';
    }catch (error) {
        document.getElementById('loadingIndicator').style.display = 'none';
        showError('No Data Found For This Date Range : ' + error.message);
        console.error('No Data Found For This Date Range : ', error);
    }
}

async function exportCurrentZone() {
    if (!currentZone) {
        showError('Please get statistics first');
        return;
    }

    const startDateTime = document.getElementById('startDateTime').value;
    const endDateTime = document.getElementById('endDateTime').value;

    try {
        // Start export
        const startResponse = await fetch('/api/combined_statistics/export', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                zone: currentZone,
                start_date: startDateTime,
                end_date: endDateTime
            })
        });

        if (!startResponse.ok) {
            throw new Error('Failed to start export');
        }

        const startData = await startResponse.json();
        const taskId = startData.task_id;

        // Poll for status
        let progress = 0;
        while (progress < 100) {
            await new Promise(resolve => setTimeout(resolve, 1000)); // Wait 1 second

            const statusResponse = await fetch(`/api/combined_statistics/export/status?task_id=${taskId}`);
            if (!statusResponse.ok) {
                throw new Error('Failed to check export status');
            }

            const statusData = await statusResponse.json();
            progress = statusData.progress;

            if (statusData.error) {
                throw new Error(statusData.error);
            }

            // Update UI with progress if needed
            console.log(`Export progress: ${progress}%`);
        }

        // Download file
        const downloadResponse = await fetch(`/api/combined_statistics/export/download?task_id=${taskId}&zone=${currentZone}`);
        if (!downloadResponse.ok) {
            throw new Error('Failed to download file');
        }

        const blob = await downloadResponse.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `${currentZone}_statistics_${new Date().getTime()}.xlsx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

    } catch (error) {
        showError('Export failed: ' + error.message);
        console.error('Export error:', error);
    }
}

async function exportAllZones() {
    const startDateTime = document.getElementById('startDateTime').value;
    const endDateTime = document.getElementById('endDateTime').value;

    // Validate inputs
    if (!startDateTime || !endDateTime) {
        showError('Please select both start and end date/time');
        return;
    }

    // Show loading
    document.getElementById('loadingIndicator').style.display = 'block';

    try {
        // Start export
        const startResponse = await fetch('/api/combined_statistics/export_all', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                start_date: startDateTime,
                end_date: endDateTime
            })
        });

        if (!startResponse.ok) {
            throw new Error('Failed to start export');
        }

        const { task_id } = await startResponse.json();

        // Poll for status
        const pollStatus = async () => {
            try {
                const statusResponse = await fetch(`/api/combined_statistics/export_all/status?task_id=${task_id}`);
                if (!statusResponse.ok) {
                    throw new Error('Failed to check status');
                }

                const status = await statusResponse.json();

                if (status.error) {
                    throw new Error(status.error);
                }

                if (status.done) {
                    // Download file
                    const downloadResponse = await fetch(`/api/combined_statistics/export_all/download?task_id=${task_id}`);
                    if (!downloadResponse.ok) {
                        throw new Error('Failed to download file');
                    }

                    const blob = await downloadResponse.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `all_zones_statistics_${new Date().getTime()}.xlsx`;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);

                    document.getElementById('loadingIndicator').style.display = 'none';
                } else {
                    // Continue polling
                    setTimeout(pollStatus, 2000);
                }
            } catch (error) {
                document.getElementById('loadingIndicator').style.display = 'none';
                showError('Export failed: ' + error.message);
                console.error('Export error:', error);
            }
        };

        // Start polling
        setTimeout(pollStatus, 2000);

    } catch (error) {
        document.getElementById('loadingIndicator').style.display = 'none';
        showError('Export failed: ' + error.message);
        console.error('Export error:', error);
    }
}

function showError(message) {
    document.getElementById('errorText').textContent = message;
    document.getElementById('errorMessage').style.display = 'block';
}
