let APP_DATA = null;
let APP_SUMMARIES = null;
let charts = {};

// ===== INIT =====
document.addEventListener("DOMContentLoaded", () => {
    setupNav();
    loadFileInfo();
    fetchData();
    loadSlides();
});

function setupNav() {
    document.querySelectorAll(".sidebar-nav a").forEach(link => {
        link.addEventListener("click", e => {
            e.preventDefault();
            document.querySelectorAll(".sidebar-nav a").forEach(l => l.classList.remove("active"));
            link.classList.add("active");
            const sec = link.dataset.section;
            document.querySelectorAll(".section").forEach(s => s.classList.remove("active"));
            document.getElementById("sec-" + sec).classList.add("active");
        });
    });
}

async function fetchData() {
    try {
        const res = await fetch("/api/data");
        const json = await res.json();
        APP_DATA = json.data;
        APP_SUMMARIES = json.summaries;
        renderDashboard();
        renderAllTables();
        populateFilters();
        document.getElementById("loading").style.display = "none";
    } catch (err) {
        document.getElementById("loading").innerHTML = "<p style='color:red'>Error loading data: " + err.message + "</p>";
    }
}

// ===== DASHBOARD =====
function renderDashboard() {
    const s = APP_SUMMARIES;
    const cardsHtml = `
        <div class="summary-card">
            <div class="card-label">Total Teams</div>
            <div class="card-value">${s.total_teams}</div>
            <div class="card-sub">${s.active_teams} active</div>
        </div>
        <div class="summary-card accent">
            <div class="card-label">Total Test Cases</div>
            <div class="card-value">${formatNum(s.total_tcs)}</div>
            <div class="card-sub">${formatNum(s.total_automated)} automated</div>
        </div>
        <div class="summary-card success">
            <div class="card-label">Automation Coverage</div>
            <div class="card-value">${s.automation_coverage}%</div>
            <div class="card-sub">${formatNum(s.total_feasible)} feasible TCs</div>
        </div>
        <div class="summary-card danger">
            <div class="card-label">Total Defects</div>
            <div class="card-value">${formatNum(Object.values(s.defect_by_domain).reduce((a,b)=>a+b,0))}</div>
            <div class="card-sub">across all sprints</div>
        </div>
        <div class="summary-card info">
            <div class="card-label">Daily Pass Rate</div>
            <div class="card-value">${s.daily_pass_rate}%</div>
            <div class="card-sub">${formatNum(s.daily_passed)} passed / ${formatNum(s.daily_failed)} failed</div>
        </div>
        <div class="summary-card">
            <div class="card-label">Not Feasible TCs</div>
            <div class="card-value">${formatNum(s.total_not_feasible)}</div>
            <div class="card-sub">cannot be automated</div>
        </div>
    `;
    document.getElementById("summaryCards").innerHTML = cardsHtml;

    renderCharts();
}

function renderCharts() {
    const s = APP_SUMMARIES;
    Object.values(charts).forEach(c => c.destroy());
    charts = {};

    // Automation Coverage by Domain
    const domainLabels = Object.keys(s.tcs_by_domain);
    const domainAuto = domainLabels.map(d => s.tcs_by_domain[d].automated);
    const domainTotal = domainLabels.map(d => s.tcs_by_domain[d].total);
    const domainManual = domainLabels.map((d, i) => domainTotal[i] - domainAuto[i]);

    charts.autoDomain = new Chart(document.getElementById("chartAutoDomain"), {
        type: "bar",
        data: {
            labels: domainLabels.map(l => truncate(l, 20)),
            datasets: [
                { label: "Automated", data: domainAuto, backgroundColor: "#1a237e" },
                { label: "Not Automated", data: domainManual, backgroundColor: "#e0e0e0" }
            ]
        },
        options: { ...barOpts(), plugins: { ...barOpts().plugins, title: { display: false } } }
    });

    // Defect Severity
    const sevLabels = Object.keys(s.defect_severity);
    const sevData = Object.values(s.defect_severity);
    charts.defectSev = new Chart(document.getElementById("chartDefectSeverity"), {
        type: "doughnut",
        data: {
            labels: sevLabels,
            datasets: [{ data: sevData, backgroundColor: ["#c62828", "#ff6f00", "#f9a825", "#0277bd"] }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "bottom" } } }
    });

    // Defects by Domain
    const defDomLabels = Object.keys(s.defect_by_domain).filter(d => s.defect_by_domain[d] > 0);
    const defDomData = defDomLabels.map(d => s.defect_by_domain[d]);
    charts.defectDomain = new Chart(document.getElementById("chartDefectDomain"), {
        type: "bar",
        data: {
            labels: defDomLabels.map(l => truncate(l, 20)),
            datasets: [{ label: "Defects", data: defDomData, backgroundColor: "#ff6f00" }]
        },
        options: barOpts()
    });

    // Daily Execution
    charts.dailyExec = new Chart(document.getElementById("chartDailyExec"), {
        type: "doughnut",
        data: {
            labels: ["Passed", "Failed"],
            datasets: [{ data: [s.daily_passed, s.daily_failed], backgroundColor: ["#2e7d32", "#c62828"] }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "bottom" } } }
    });

    // RAD Enablement
    const radLabels = Object.keys(s.rad_by_domain);
    const radEnabled = radLabels.map(d => s.rad_by_domain[d].rad_enabled);
    const radTotal = radLabels.map(d => s.rad_by_domain[d].total_builds);
    charts.rad = new Chart(document.getElementById("chartRAD"), {
        type: "bar",
        data: {
            labels: radLabels.map(l => truncate(l, 20)),
            datasets: [
                { label: "RAD Enabled", data: radEnabled, backgroundColor: "#2e7d32" },
                { label: "Total Builds", data: radTotal, backgroundColor: "#e0e0e0" }
            ]
        },
        options: barOpts()
    });

    // Release Day
    const relLabels = Object.keys(s.release_by_domain);
    const relFeasible = relLabels.map(d => s.release_by_domain[d].feasible);
    const relAuto = relLabels.map(d => s.release_by_domain[d].automated);
    charts.release = new Chart(document.getElementById("chartRelease"), {
        type: "bar",
        data: {
            labels: relLabels.map(l => truncate(l, 20)),
            datasets: [
                { label: "Automated", data: relAuto, backgroundColor: "#1a237e" },
                { label: "Feasible", data: relFeasible, backgroundColor: "#bbdefb" }
            ]
        },
        options: barOpts()
    });
}

// ===== TABLES =====
function renderAllTables() {
    renderDataTable("projectTableWrapper", "projectTable", APP_DATA.project_details,
        ["S. No.", "Scrum Team Name", "Project Name", "Team Type", "Domain", "Sub Domain",
         "QE Manager", "Portfolio Lead", "Project Status", "Offshore POC", "Onshore POC",
         "Submitted By", "Submitted On"]);

    renderDataTable("tcsTableWrapper", "tcsTable", APP_DATA.tcs_details,
        ["S. No.", "Scrum Team Name",
         "Total TCs P0", "Total TCs P1", "Total TCs P2", "Total TCs P3", "Total TCs P4",
         "Total TCs Automated P0", "Total TCs Automated P1", "Total TCs Automated P2", "Total TCs Automated P3", "Total TCs Automated P4",
         "Total TCs Feasible P0", "Total TCs Feasible P1", "Total TCs Feasible P2", "Total TCs Feasible P3", "Total TCs Feasible P4",
         "Total TCs Not Feasible P0", "Total TCs Not Feasible P1", "Total TCs Not Feasible P2", "Total TCs Not Feasible P3", "Total TCs Not Feasible P4",
         "Total Tech Debt P0", "Total Tech Debt P1", "Total Tech Debt P2", "Total Tech Debt P3", "Total Tech Debt P4",
         "Automation Coverage", "Submitted By", "Submitted On", "Domain", "Sub Domain"]);

    renderDataTable("insprintTableWrapper", "insprintTable", APP_DATA.insprint_data,
        ["S. No.", "Scrum team", "Sprint",
         "Test Case Designed P0", "Test Case Designed P1", "Test Case Designed P2", "Test Case Designed P3", "Test Case Designed P4",
         "Test case Automated P0", "Test case Automated P1", "Test case Automated P2", "Test case Automated P3", "Test case Automated P4",
         "Not Feasible P0", "Not Feasible P1", "Not Feasible P2", "Not Feasible P3", "Not Feasible P4",
         "Selected Auto tech debt P0", "Selected Auto tech debt P1", "Selected Auto tech debt P2", "Selected Auto tech debt P3", "Selected Auto tech debt P4",
         "Auto tech debt P0", "Auto tech debt P1", "Auto tech debt P2", "Auto tech debt P3", "Auto tech debt P4",
         "Automation coverage", "TCs Maintained",
         "Comment (If Automation Cov is 0%)", "Comment (if have Not Feasible TCs)",
         "Submitted By", "Submitted On", "Domain", "Sub Domain"]);

    renderDataTable("defectTableWrapper", "defectTable", APP_DATA.defect_data,
        ["S. No.", "Sprint", "ScrumTeam", "Type Of Defect",
         "InSprint Fatel", "InSprint Serious", "InSprint Medium", "InSprint Low",
         "Regression Fatel", "Regression Serious", "Regression Medium", "Regression Low",
         "Prod Defect Count", "Prod Defect RCA", "Comments For Not having Any Defect",
         "Submitted By", "Submitted On", "Domain", "Sub Domain"]);

    renderDataTable("dailyTableWrapper", "dailyTable", APP_DATA.daily_execution,
        ["S. No.", "Scrum Team Name", "Automation Execution Frequency",
         "P0", "P1", "P2", "P3", "P4",
         "P0 - P2 Passed TCs", "P3 - P4 Passed TCs", "Total Passed TC #",
         "P0 - P2 Failed TCs", "P3 - P4 Failed TCs", "Total Failed TC #",
         "Unattended Execution Pass %", "Failure Reason", "Comment(If Failure Reason ~ Others)",
         "Submitted By", "Submitted On", "Domain", "Sub Domain"]);

    renderDataTable("branchTableWrapper", "branchTable", APP_DATA.branch_testcase,
        ["S. No.", "Scrum Team Name",
         "Main Branch: Total No. of TCs", "Main Branch: No. of Scripts part of Open PR", "Main Branch:No. of Open Pr Dev To Main",
         "Develop Branch: Total No. of TCs", "Develop Branch: No. of Scripts part of Open PR", "Develop Branch: No. of Open Pr Feature To Dev",
         "Feature Branch: Total No. of TCs",
         "Submitted By", "Submitted On", "Domain", "Sub Domain"]);

    renderDataTable("releaseTableWrapper", "releaseTable", APP_DATA.release_day,
        ["S. No.", "Scrum Team Name", "Release Month",
         "No Of Stories part of the Release", "Release Automation Coverage",
         "No of Automation feasible TCs", "No of TCs Automated", "Comment",
         "Submitted By", "Submitted On", "Domain", "Sub Domain"]);

    renderDataTable("radTableWrapper", "radTable", APP_DATA.rad_enabled,
        ["S. No.", "Scrum Team Name", "Independent Build #", "Shared Build #",
         "# Builds QA Pull Enabled", "# Builds RAD Enabled",
         "RAD Eligible build from Independent Build", "RAD Eligible build from Shared build",
         "Remarks", "Submitted By", "Submitted On", "Domain", "Sub Domain"]);

    renderDataTable("monthlyTableWrapper", "monthlyTable", APP_DATA.monthly_sheet,
        ["S. No.", "Employee Id", "Empolyee Name", "Scrum Team Name", "Month'Year",
         "Feature Branch Count", "Develop Branch Count", "Main Branch Count", "Total Count",
         "DashBoard Count", "Validation of DashBoard Count", "Comment",
         "Submitted By", "Submitted On",
         "Automation Feature Branch Count", "Automation Develop Branch Count", "Automation Main Branch Count", "Automation Total Count",
         "Technology Used", "Domain", "Sub Domain"]);

    // Extra charts for test cases section
    renderTCCharts();
    renderDefectCharts();
}

function renderDataTable(wrapperId, tableId, records, columns) {
    if (!records || records.length === 0) {
        document.getElementById(wrapperId).innerHTML = "<p style='padding:20px;color:#999'>No data available</p>";
        return;
    }

    // Find matching column keys
    const sampleKeys = Object.keys(records[0]);
    const colMap = columns.map(col => {
        const exact = sampleKeys.find(k => k === col);
        if (exact) return { label: col, key: exact };
        const trimmed = sampleKeys.find(k => k.trim() === col.trim());
        if (trimmed) return { label: col, key: trimmed };
        return { label: col, key: col };
    });

    let html = `<table id="${tableId}"><thead><tr>`;
    colMap.forEach(c => { html += `<th>${c.label}</th>`; });
    html += "</tr></thead><tbody>";
    records.forEach(r => {
        html += "<tr>";
        colMap.forEach(c => {
            let val = r[c.key] !== undefined ? r[c.key] : "";
            if (c.label === "Project Status") {
                const cls = String(val).toLowerCase() === "active" ? "badge-active" : "badge-inactive";
                val = `<span class="badge ${cls}">${val}</span>`;
            } else if (c.label === "Automation Coverage" || c.label === "Automation coverage" || c.label === "Release Automation Coverage" || c.label === "Unattended Execution Pass %") {
                const pct = parseFloat(val);
                if (!isNaN(pct)) {
                    const display = pct <= 1 ? (pct * 100).toFixed(1) : pct.toFixed ? pct.toFixed(1) : pct;
                    const color = display >= 70 ? "var(--success)" : display >= 40 ? "var(--warning)" : "var(--danger)";
                    val = `<span style="font-weight:700;color:${color}">${display}%</span>`;
                }
            }
            html += `<td>${val}</td>`;
        });
        html += "</tr>";
    });
    html += "</tbody></table>";
    document.getElementById(wrapperId).innerHTML = html;
}

function renderTCCharts() {
    const data = APP_DATA.tcs_details;
    if (!data || data.length === 0) return;

    // Priority distribution
    const priorities = ["P0", "P1", "P2", "P3", "P4"];
    const priData = priorities.map(p => data.reduce((sum, r) => sum + num(r["Total TCs " + p]), 0));
    charts.tcPriority = new Chart(document.getElementById("chartTCPriority"), {
        type: "bar",
        data: {
            labels: priorities,
            datasets: [{ label: "Test Cases", data: priData, backgroundColor: ["#c62828", "#ff6f00", "#f9a825", "#0277bd", "#757575"] }]
        },
        options: barOpts()
    });

    // Automation by team (top 15)
    const teamData = data.map(r => ({
        name: r["Scrum Team Name"] || "",
        coverage: num(r["Automation Coverage"])
    })).filter(t => t.name).sort((a, b) => b.coverage - a.coverage).slice(0, 15);

    charts.tcAutoTeam = new Chart(document.getElementById("chartTCAutoTeam"), {
        type: "bar",
        data: {
            labels: teamData.map(t => truncate(t.name, 18)),
            datasets: [{
                label: "Coverage",
                data: teamData.map(t => t.coverage <= 1 ? (t.coverage * 100).toFixed(1) : t.coverage.toFixed(1)),
                backgroundColor: teamData.map(t => {
                    const v = t.coverage <= 1 ? t.coverage * 100 : t.coverage;
                    return v >= 70 ? "#2e7d32" : v >= 40 ? "#f9a825" : "#c62828";
                })
            }]
        },
        options: { ...barOpts(), indexAxis: "y" }
    });
}

function renderDefectCharts() {
    const data = APP_DATA.defect_data;
    if (!data || data.length === 0) return;

    // By sprint
    const sprints = [...new Set(data.map(r => r.Sprint))].filter(Boolean);
    const sprintTotals = sprints.map(sp => data.filter(r => r.Sprint === sp).reduce((sum, r) =>
        sum + num(r["InSprint Fatel"]) + num(r["InSprint Serious"]) + num(r["InSprint Medium"]) + num(r["InSprint Low"])
        + num(r["Regression Fatel"]) + num(r["Regression Serious"]) + num(r["Regression Medium"]) + num(r["Regression Low"]), 0));

    charts.defectSprint = new Chart(document.getElementById("chartDefectSprint"), {
        type: "bar",
        data: { labels: sprints, datasets: [{ label: "Defects", data: sprintTotals, backgroundColor: "#ff6f00" }] },
        options: barOpts()
    });

    // Manual vs Automation
    const manual = data.filter(r => r["Type Of Defect"] === "Manual").reduce((sum, r) =>
        sum + num(r["InSprint Fatel"]) + num(r["InSprint Serious"]) + num(r["InSprint Medium"]) + num(r["InSprint Low"])
        + num(r["Regression Fatel"]) + num(r["Regression Serious"]) + num(r["Regression Medium"]) + num(r["Regression Low"]), 0);
    const auto = data.filter(r => r["Type Of Defect"] === "Automation").reduce((sum, r) =>
        sum + num(r["InSprint Fatel"]) + num(r["InSprint Serious"]) + num(r["InSprint Medium"]) + num(r["InSprint Low"])
        + num(r["Regression Fatel"]) + num(r["Regression Serious"]) + num(r["Regression Medium"]) + num(r["Regression Low"]), 0);

    charts.defectType = new Chart(document.getElementById("chartDefectType"), {
        type: "doughnut",
        data: {
            labels: ["Manual", "Automation"],
            datasets: [{ data: [manual, auto], backgroundColor: ["#1a237e", "#ff6f00"] }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "bottom" } } }
    });
}

// ===== FILTERS =====
function populateFilters() {
    // Domains from project details
    const domains = [...new Set(APP_DATA.project_details.map(r => r.Domain))].filter(Boolean).sort();
    populateSelect("filterProjectDomain", domains);
    populateSelect("filterTCDomain", domains);
    populateSelect("filterInSprintDomain", domains);
    populateSelect("filterDefectDomain", domains);
    populateSelect("filterDailyDomain", domains);
    populateSelect("filterBranchDomain", domains);
    populateSelect("filterReleaseDomain", domains);
    populateSelect("filterRADDomain", domains);
    populateSelect("filterMonthlyDomain", domains);

    // Sprints
    const sprints = [...new Set(APP_DATA.defect_data.map(r => r.Sprint))].filter(Boolean).sort();
    populateSelect("filterDefectSprint", sprints);
    const insprintSprints = [...new Set(APP_DATA.insprint_data.map(r => r["Sprint"] || r["Sprint  "]))].filter(Boolean).sort();
    populateSelect("filterInSprintSprint", insprintSprints);

    // Months
    const months = [...new Set(APP_DATA.monthly_sheet.map(r => r["Month'Year"]))].filter(Boolean).sort();
    populateSelect("filterMonthlyMonth", months);
}

function populateSelect(id, options) {
    const sel = document.getElementById(id);
    if (!sel) return;
    options.forEach(opt => {
        const o = document.createElement("option");
        o.value = opt;
        o.textContent = opt;
        sel.appendChild(o);
    });
}

function filterTable(tableId, value, colIndexOrName) {
    const table = document.getElementById(tableId);
    if (!table) return;
    const rows = table.querySelectorAll("tbody tr");
    const headers = [...table.querySelectorAll("thead th")].map(th => th.textContent.trim());

    let colIdx = typeof colIndexOrName === "number" ? colIndexOrName :
        headers.findIndex(h => h === colIndexOrName || h.trim() === String(colIndexOrName).trim());

    rows.forEach(row => {
        if (!value) { row.style.display = ""; return; }
        const cell = row.cells[colIdx];
        row.style.display = cell && cell.textContent.trim().includes(value) ? "" : "none";
    });
}

function searchTable(tableId, query) {
    const table = document.getElementById(tableId);
    if (!table) return;
    const rows = table.querySelectorAll("tbody tr");
    const q = query.toLowerCase();
    rows.forEach(row => {
        row.style.display = row.textContent.toLowerCase().includes(q) ? "" : "none";
    });
}

// ===== FILE UPLOAD =====
async function loadFileInfo() {
    try {
        const res = await fetch("/api/file-info");
        const info = await res.json();
        document.getElementById("currentExcelName").textContent = info.excel_name;
        document.getElementById("currentPptxName").textContent = info.pptx_name;
    } catch (e) { /* ignore */ }
}

async function uploadExcel(input) {
    const file = input.files[0];
    if (!file) return;

    const nameEl = document.getElementById("currentExcelName");
    nameEl.textContent = "Uploading...";

    const formData = new FormData();
    formData.append("file", file);

    try {
        const res = await fetch("/api/upload-excel", { method: "POST", body: formData });
        const data = await res.json();
        if (data.error) {
            nameEl.textContent = "Upload failed";
            alert(data.error);
            return;
        }
        nameEl.innerHTML = data.filename + ' <span class="upload-success">&#10003; Loaded</span>';
        if (data.warning) alert(data.warning);
        // Reload all data
        document.getElementById("loading").style.display = "flex";
        await fetchData();
    } catch (err) {
        nameEl.textContent = "Upload failed";
        alert("Upload failed: " + err.message);
    }
    input.value = "";
}

async function uploadPptx(input) {
    const file = input.files[0];
    if (!file) return;

    const nameEl = document.getElementById("currentPptxName");
    nameEl.textContent = "Uploading...";

    const formData = new FormData();
    formData.append("file", file);

    try {
        const res = await fetch("/api/upload-pptx", { method: "POST", body: formData });
        const data = await res.json();
        if (data.error) {
            nameEl.textContent = "Upload failed";
            alert(data.error);
            return;
        }
        nameEl.innerHTML = data.filename + ' <span class="upload-success">&#10003; ' + data.slide_count + ' slides</span>';
        // Reload slides
        await loadSlides();
    } catch (err) {
        nameEl.textContent = "Upload failed";
        alert("Upload failed: " + err.message);
    }
    input.value = "";
}

// ===== DYNAMIC SLIDE SELECTOR =====
async function loadSlides() {
    try {
        const res = await fetch("/api/pptx-slides");
        const data = await res.json();
        if (data.error) {
            document.getElementById("slideSelector").innerHTML = '<p style="color:var(--danger)">' + data.error + '</p>';
            return;
        }
        renderSlideSelector(data.slides);
    } catch (e) {
        document.getElementById("slideSelector").innerHTML = '<p style="color:var(--danger)">Failed to load slides</p>';
    }
}

function renderSlideSelector(slides) {
    const container = document.getElementById("slideSelector");
    let html = '';

    slides.forEach(s => {
        html += `<label class="slide-option">
            <input type="checkbox" name="slide" value="${s.num}" checked>
            <span class="slide-num">${s.num}</span>
            ${escapeHtml(s.title)}
        </label>`;
    });

    container.innerHTML = '<div class="slide-group"><div class="slide-group-title">All Slides (' + slides.length + ')</div>' + html + '</div>';
    updateSlideCount();
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

function getSelectedSlides() {
    return Array.from(document.querySelectorAll('input[name="slide"]:checked')).map(cb => parseInt(cb.value));
}

function updateSlideCount() {
    const selected = getSelectedSlides().length;
    const total = document.querySelectorAll('input[name="slide"]').length;
    const el = document.getElementById("slideCount");
    if (el) el.textContent = `${selected} of ${total} slides selected`;
}

function toggleAllSlides(checked) {
    document.querySelectorAll('input[name="slide"]').forEach(cb => cb.checked = checked);
    updateSlideCount();
}

document.addEventListener("change", e => {
    if (e.target.name === "slide") updateSlideCount();
});

function navigateTo(section) {
    document.querySelectorAll(".sidebar-nav a").forEach(l => {
        l.classList.toggle("active", l.dataset.section === section);
    });
    document.querySelectorAll(".section").forEach(s => s.classList.remove("active"));
    document.getElementById("sec-" + section).classList.add("active");
    window.scrollTo(0, 0);
}

// ===== PPT GENERATION =====
async function generatePPT() {
    const statusEl = document.getElementById("pptStatus");
    const selected = getSelectedSlides();

    if (selected.length === 0) {
        if (statusEl) statusEl.innerHTML = '<span style="color:var(--danger)">Please select at least one slide.</span>';
        return;
    }

    if (statusEl) statusEl.textContent = `Generating ${selected.length} slides...`;

    try {
        // 1. Download the populated PPT
        const res = await fetch("/api/generate-ppt", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ slides: selected })
        });
        if (!res.ok) throw new Error("Server error");
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "LPL_QE_Connect_Generated.pptx";
        a.click();
        URL.revokeObjectURL(url);

        // 2. Also download the calculation report PDF.
        // Don't block PPT success on PDF failure — best effort.
        try {
            const pdfRes = await fetch("/api/generate-report-pdf", { method: "POST" });
            if (pdfRes.ok) {
                const pdfBlob = await pdfRes.blob();
                const pdfUrl = URL.createObjectURL(pdfBlob);
                const a2 = document.createElement("a");
                a2.href = pdfUrl;
                a2.download = "LPL_QE_Calculation_Report.pdf";
                a2.click();
                URL.revokeObjectURL(pdfUrl);
                if (statusEl) statusEl.innerHTML = `<span style="color:var(--success)">PPT + Calculation Report downloaded.</span>`;
            } else {
                if (statusEl) statusEl.innerHTML = `<span style="color:var(--success)">PPT downloaded.</span> <span style="color:var(--danger)">Report PDF failed.</span>`;
            }
        } catch (pdfErr) {
            if (statusEl) statusEl.innerHTML = `<span style="color:var(--success)">PPT downloaded.</span> <span style="color:var(--danger)">Report PDF error: ${pdfErr.message}</span>`;
        }
    } catch (err) {
        if (statusEl) statusEl.innerHTML = '<span style="color:var(--danger)">Error: ' + err.message + '</span>';
    }
}

// ===== UTILS =====
function formatNum(n) {
    return Number(n).toLocaleString();
}

function num(v) {
    const n = parseFloat(v);
    return isNaN(n) ? 0 : n;
}

function truncate(str, len) {
    return str.length > len ? str.substring(0, len) + "..." : str;
}

function barOpts() {
    return {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
            x: { ticks: { maxRotation: 45, font: { size: 10 } } },
            y: { beginAtZero: true }
        },
        plugins: { legend: { position: "top", labels: { font: { size: 11 } } } }
    };
}
