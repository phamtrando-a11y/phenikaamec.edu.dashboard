let globalData = { sessions: [], students: [] };
let choicesYear, choicesMonth, choicesUnit;
let isUpdatingFilters = false;
let activePage = 'page-overview';

document.addEventListener("DOMContentLoaded", () => {
    fetch('data.json')
        .then(response => response.json())
        .then(data => {
            globalData = data;
            initChoices();
            updateDashboard();
        })
        .catch(err => console.error("Error loading data:", err));
        
    // Close modal when clicking outside
    window.onclick = function(event) {
        let modal = document.getElementById("studentModal");
        if (event.target == modal) {
            closeModal();
        }
    }
});

function switchPage(pageId) {
    activePage = pageId;
    
    // Update navigation UI
    document.querySelectorAll('.nav-btn').forEach(btn => btn.classList.remove('active'));
    event.currentTarget.classList.add('active');
    
    // Update active page content
    document.querySelectorAll('.page-section').forEach(page => {
        page.style.display = 'none';
        page.classList.remove('active');
    });
    let activeElem = document.getElementById(pageId);
    activeElem.style.display = 'block';
    
    // Trigger reflow for animation
    void activeElem.offsetWidth; 
    activeElem.classList.add('active');
    
    // Update header title
    const titles = {
        'page-overview': 'Tổng quan',
        'page-department': 'Hoạt động khoa phòng',
        'page-training': 'Hoạt động đào tạo'
    };
    document.getElementById('page-title').innerText = titles[pageId];
    
    // Manage filters visibility
    if (pageId === 'page-training') {
        document.getElementById('filter-year-col').style.display = 'none';
        document.getElementById('filter-month-col').style.display = 'none';
        document.getElementById('filter-wrapper').style.gridTemplateColumns = '1fr';
    } else {
        document.getElementById('filter-year-col').style.display = 'flex';
        document.getElementById('filter-month-col').style.display = 'flex';
        document.getElementById('filter-wrapper').style.gridTemplateColumns = 'repeat(3, 1fr)';
    }
    
    // Resize charts if overview is shown
    if (pageId === 'page-overview') {
        window.dispatchEvent(new Event('resize'));
    }
    
    updateDashboard(); // Re-render with potentially new filters
}

function initChoices() {
    choicesYear = new Choices('#yearSelect', { removeItemButton: true, searchEnabled: true, placeholderValue: 'Chọn Năm' });
    choicesMonth = new Choices('#monthSelect', { removeItemButton: true, searchEnabled: true, placeholderValue: 'Chọn Tháng' });
    choicesUnit = new Choices('#unitSelect', { removeItemButton: true, searchEnabled: true, placeholderValue: 'Chọn Đơn Vị' });

    document.querySelectorAll('select').forEach(el => {
        el.addEventListener('change', updateDashboard);
    });
}

function getSelected(id) {
    let sel = document.getElementById(id);
    return Array.from(sel.selectedOptions).map(opt => opt.value);
}

function getFilteredValues(excludeField, selectedYears, selectedMonths, selectedUnits) {
    let set = new Set();
    globalData.sessions.forEach(s => {
        let y = String(s["Năm"] || "").replace('.0', '');
        let m = String(s["Tháng"] || "").replace('.0', '');
        let u = s["ĐƠN VỊ"];

        if (excludeField !== 'year' && selectedYears.length > 0 && !selectedYears.includes(y)) return;
        if (excludeField !== 'month' && selectedMonths.length > 0 && !selectedMonths.includes(m)) return;
        if (excludeField !== 'unit' && selectedUnits.length > 0 && !selectedUnits.includes(u)) return;

        if (excludeField === 'year' && y) set.add(y);
        if (excludeField === 'month' && m) set.add(m);
        if (excludeField === 'unit' && u) set.add(u);
    });
    return Array.from(set);
}

function updateFilterOptions(selectedYears, selectedMonths, selectedUnits) {
    const updateChoice = (choiceInstance, excludeField, selectedArr, sortFn) => {
        let available = getFilteredValues(excludeField, selectedYears, selectedMonths, selectedUnits);
        if (sortFn) available.sort(sortFn);
        
        let newChoices = available.map(val => ({
            value: val,
            label: val,
            selected: selectedArr.includes(val)
        }));
        
        selectedArr.forEach(val => {
            if (!available.includes(val)) {
                newChoices.push({ value: val, label: val, selected: true });
            }
        });

        choiceInstance.clearStore();
        choiceInstance.setChoices(newChoices, 'value', 'label', true);
    };

    updateChoice(choicesYear, 'year', selectedYears, (a,b) => b - a);
    updateChoice(choicesMonth, 'month', selectedMonths, (a,b) => a - b);
    updateChoice(choicesUnit, 'unit', selectedUnits);
}

function truncate(str, n=40) {
    if(!str) return "";
    return (str.length > n) ? str.substr(0, n-3) + '...' : str;
}

function safeFloat(val) {
    if (val === "" || val === null || val === undefined) return null;
    return parseFloat(val) || 0;
}

function updateDashboard() {
    if (isUpdatingFilters) return;
    
    // For page-training, we ignore year and month filters completely
    let selectedYears = activePage === 'page-training' ? [] : getSelected('yearSelect');
    let selectedMonths = activePage === 'page-training' ? [] : getSelected('monthSelect');
    const selectedUnits = getSelected('unitSelect');

    isUpdatingFilters = true;
    updateFilterOptions(selectedYears, selectedMonths, selectedUnits);
    isUpdatingFilters = false;

    // Filter sessions (For page-overview and page-department)
    let filteredSessions = [];
    if (activePage === 'page-overview' || activePage === 'page-department') {
        filteredSessions = globalData.sessions.filter(s => {
            let match = true;
            let y = String(s["Năm"] || "").replace('.0', '');
            let m = String(s["Tháng"] || "").replace('.0', '');
            if (selectedYears.length > 0 && !selectedYears.includes(y)) match = false;
            if (selectedMonths.length > 0 && !selectedMonths.includes(m)) match = false;
            if (selectedUnits.length > 0 && !selectedUnits.includes(s["ĐƠN VỊ"])) match = false;
            return match;
        });
    }

    // Filter students (For page-training)
    let filteredStudents = [];
    if (activePage === 'page-training') {
        filteredStudents = globalData.students.filter(st => {
            let match = true;
            if (selectedUnits.length > 0 && !selectedUnits.includes(st["Đơn vị đào tạo"])) match = false;
            return match;
        });
    }

    if (activePage === 'page-overview') {
        renderOverview(filteredSessions);
    } else if (activePage === 'page-department') {
        renderDepartment(filteredSessions);
    } else if (activePage === 'page-training') {
        renderTraining(filteredStudents);
    }
}

function renderOverview(filteredSessions) {
    const totalPrograms = filteredSessions.length;
    let totalParticipants = 0;
    let totalPassed = 0;
    let programsMap = {};
    
    filteredSessions.forEach(s => {
        let attended = safeFloat(s["Số học viên tham gia"]);
        let passed = safeFloat(s["Số Học viên đạt"]);
        totalParticipants += attended;
        totalPassed += passed;
        
        let pName = s["NỘI DUNG/CHƯƠNG TRÌNH"];
        if (pName) {
            if (!programsMap[pName]) programsMap[pName] = { attended: 0, passed: 0 };
            programsMap[pName].attended += attended;
            programsMap[pName].passed += passed;
        }
    });

    document.getElementById("kpi-total-activities").innerText = totalPrograms;
    document.getElementById("kpi-total-participants").innerText = totalParticipants;
    document.getElementById("kpi-total-passed").innerText = totalPassed;

    // Pie Chart
    let failed = totalParticipants - totalPassed;
    if (failed < 0) failed = 0;
    const pieData = [{
        values: [totalPassed, failed],
        labels: ['Đạt', 'Không đạt'],
        type: 'pie',
        hole: 0.4,
        marker: { colors: ['#2ecc71', '#e74c3c'] },
        textinfo: 'percent',
        hoverinfo: 'label+value+percent'
    }];
    Plotly.newPlot('chartPie', pieData, { margin: {t:10, b:10, l:10, r:10} }, {responsive: true});

    // Bar Chart
    let pList = Object.keys(programsMap).map(k => ({
        name: k,
        shortName: truncate(k),
        attended: programsMap[k].attended,
        passed: programsMap[k].passed
    })).sort((a,b) => b.attended - a.attended).slice(0, 10).reverse();

    const trace1 = { x: pList.map(p => p.attended), y: pList.map(p => p.name), name: 'Tham gia', type: 'bar', orientation: 'h', marker: {color: '#3498db'} };
    const trace2 = { x: pList.map(p => p.passed), y: pList.map(p => p.name), name: 'Đạt', type: 'bar', orientation: 'h', marker: {color: '#2ecc71'} };
    
    Plotly.newPlot('chartProgram', [trace1, trace2], {
        barmode: 'group',
        margin: { t: 10, l: 10, r: 20, b: 40 },
        yaxis: { automargin: true, tickmode: 'array', tickvals: pList.map(p => p.name), ticktext: pList.map(p => p.shortName) },
        legend: { orientation: "h", yanchor: "bottom", y: 1.02, xanchor: "right", x: 1 }
    }, {responsive: true});
}

function renderDepartment(filteredSessions) {
    let tbody = document.querySelector("#departmentTable tbody");
    tbody.innerHTML = "";
    
    filteredSessions.forEach((s, i) => {
        let timeStr = "";
        if (s["Ngày"] || s["Tháng"] || s["Năm"]) {
            timeStr = `${String(s["Ngày"]||"").replace('.0','')} / ${String(s["Tháng"]||"").replace('.0','')} / ${String(s["Năm"]||"").replace('.0','')}`;
        }
        
        let tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${i + 1}</td>
            <td>${s["NỘI DUNG/CHƯƠNG TRÌNH"] || ""}</td>
            <td>${timeStr}</td>
            <td>${s["NGƯỜI PHỤ TRÁCH NỘI DUNG"] || ""}</td>
            <td>${safeFloat(s["Số học viên tham gia"])}</td>
            <td>${safeFloat(s["Số Học viên đạt"])}</td>
            <td>${s["Gắn link Điểm danh và post test"] ? "✅" : "❌"}</td>
            <td>${s["Gắn link bảng kiểm đào tạo nội bộ"] ? "✅" : "❌"}</td>
            <td>${s["Gắn link bài báo cáo"] ? "✅" : "❌"}</td>
            <td>${s["Gắn link video buổi đào tạo"] ? "✅" : "❌"}</td>
        `;
        tbody.appendChild(tr);
    });
}

function renderTraining(filteredStudents) {
    let studentMap = {};
    
    filteredStudents.forEach(st => {
        let name = st["Họ và tên"] || "";
        let code = st["Mã nhân sự"] || "";
        let unit = st["Đơn vị đào tạo"] || "";
        
        // Use code as primary key if available, else use name + unit as fallback
        let key = code ? code : `${name}_${unit}`;
        if (!key) return; // Skip empty rows
        
        if (!studentMap[key]) {
            studentMap[key] = {
                name: name,
                code: code,
                unit: unit,
                totalSessions: 0,
                passedSessions: 0,
                details: []
            };
        }
        
        studentMap[key].totalSessions++;
        if (st["Trạng thái"] === "Đạt") {
            studentMap[key].passedSessions++;
        }
        
        studentMap[key].details.push(st);
    });
    
    let sortedStudents = Object.values(studentMap).sort((a,b) => b.totalSessions - a.totalSessions);
    
    let tbody = document.querySelector("#trainingTable tbody");
    tbody.innerHTML = "";
    
    sortedStudents.forEach((st, i) => {
        let tr = document.createElement("tr");
        tr.className = "student-row";
        tr.onclick = () => openModal(st);
        tr.innerHTML = `
            <td>${i + 1}</td>
            <td><strong>${st.name}</strong></td>
            <td>${st.code}</td>
            <td>${st.unit}</td>
            <td>${st.totalSessions}</td>
            <td>${st.passedSessions}</td>
        `;
        tbody.appendChild(tr);
    });
}

function openModal(studentData) {
    document.getElementById("modal-student-name").innerText = studentData.name;
    document.getElementById("modal-student-unit").innerText = studentData.code ? `${studentData.code} - ${studentData.unit}` : studentData.unit;
    
    let tbody = document.querySelector("#modalTable tbody");
    tbody.innerHTML = "";
    
    // Sort details by newest if possible? Just raw order for now
    studentData.details.forEach((st, i) => {
        let timeStr = "";
        if (st["Ngày đào tạo"] || st["Tháng thực hiện"]) {
            timeStr = `${st["Ngày đào tạo"]||""} / ${st["Tháng thực hiện"]||""}`.replace(/^ \/ | \/ $/g, '');
        }
        
        let tr = document.createElement("tr");
        tr.innerHTML = `
            <td>${i + 1}</td>
            <td>${st["Tên chương trình"] || ""}</td>
            <td>${timeStr}</td>
            <td>${st["Điểm_Quy_Đổi"] !== "" && st["Điểm_Quy_Đổi"] != null ? Number(st["Điểm_Quy_Đổi"]).toFixed(1) : ""}</td>
            <td><span class="${st["Trạng thái"] === 'Đạt' ? 'status-dat' : 'status-khongdat'}">${st["Trạng thái"] || ""}</span></td>
        `;
        tbody.appendChild(tr);
    });
    
    document.getElementById("studentModal").style.display = "block";
}

function closeModal() {
    document.getElementById("studentModal").style.display = "none";
}
