
// ====================================================================
// MULTI-EMPLOYEE SELECTION - TWO-STEP WORKFLOW
// ====================================================================
// Workflow: Load ID → Preview Details → Click "Add" → Stack in List

let employeeDataList = []; // Array to store multiple employees
let currentPreviewEmployee = null; // Employee being previewed

// Show employee preview with "Add to List" button
function showEmployeePreview(employeeData) {
    currentPreviewEmployee = employeeData;
    currentEmployeeData = employeeData; // For backward compatibility

    const display = document.getElementById('employee-data');
    if (!display) return;

    const isReg = employeeData.isReg;

    display.innerHTML = `
        <div class="employee-preview-card">
            <h4>📋 Preview Employee Details</h4>
            <div class="employee-card-details">
                <p><strong>Name:</strong> ${employeeData.name}</p>
                <p><strong>ID:</strong> ${employeeData.id}</p>
                <p><strong>Current Position:</strong> ${employeeData.pos}</p>
                <p><strong>Target Position:</strong> ${employeeData.targetPos}</p>
                <p><strong>Department:</strong> ${employeeData.dept}</p>
                <p><strong>Branch:</strong> ${employeeData.branch}</p>
                <p><strong>Mentor:</strong> ${employeeData.mentor}</p>
                <p><strong>Hired Date:</strong> ${employeeData.hired}</p>
                ${isReg && employeeData.tor.school ? `<p><strong>Education:</strong> ${employeeData.tor.school}</p>` : ''}
            </div>
            <button type="button" class="btn-add-employee" onclick="addPreviewedEmployee()">
                ✓ Add to List
            </button>
        </div>
    `;

    updateButtonStatesForPreview();
}

// Add the previewed employee to the list
function addPreviewedEmployee() {
    if (!currentPreviewEmployee) return;

    addEmployeeToList(currentPreviewEmployee);
    currentPreviewEmployee = null;

    // Clear the employee ID input for next entry
    const idInput = document.getElementById('employee-id');
    if (idInput) idInput.value = '';
}

// Update button states for preview mode
function updateButtonStatesForPreview() {
    const summaryBtn = document.getElementById('generate-summary');
    if (!summaryBtn || !currentPreviewEmployee) return;

    const isReg = currentPreviewEmployee.isReg;
    if (isReg) {
        if (globalEduData) {
            summaryBtn.disabled = false;
            summaryBtn.textContent = "📄 Generate Summary (.doc)";
        } else {
            summaryBtn.disabled = true;
            summaryBtn.textContent = "Upload Edu File to Enable";
        }
    } else {
        summaryBtn.disabled = false;
        summaryBtn.textContent = "📄 Generate Summary (.doc)";
    }
}

// Add employee to the list (with duplicate checking)
function addEmployeeToList(employeeData) {
    const existingIndex = employeeDataList.findIndex(emp => emp.id === employeeData.id);

    if (existingIndex !== -1) {
        employeeDataList[existingIndex] = { ...employeeData, selected: true };
        alert(`Employee ${employeeData.name} updated in the list.`);
    } else {
        employeeDataList.push({ ...employeeData, selected: true });
    }

    renderEmployeeList();
    updateButtonStates();
}

// Remove employee from list
function removeEmployeeFromList(employeeId) {
    employeeDataList = employeeDataList.filter(emp => emp.id !== employeeId);

    if (employeeDataList.length === 0) {
        showEmptyState();
    } else {
        renderEmployeeList();
    }

    updateButtonStates();
}

// Clear all employees
function clearAllEmployees() {
    if (employeeDataList.length === 0) return;

    if (confirm('Are you sure you want to clear all loaded employees?')) {
        employeeDataList = [];
        showEmptyState();
        updateButtonStates();
    }
}

// Show empty state
function showEmptyState() {
    const display = document.getElementById('employee-data');
    if (display) {
        display.innerHTML = `
            <p style="text-align: center; color: #999; padding: 20px;">
                No data loaded. Please upload a file and search.
            </p>
        `;
    }

    const clearBtn = document.getElementById('clear-all-btn');
    if (clearBtn) clearBtn.style.display = 'none';
}

// Toggle employee selection
function toggleEmployeeSelection(employeeId) {
    const employee = employeeDataList.find(emp => emp.id === employeeId);
    if (employee) {
        employee.selected = !employee.selected;
    }
    updateSelectionCounter();
    updateButtonStates();
}

// Get selected employees
function getSelectedEmployees() {
    return employeeDataList.filter(emp => emp.selected);
}

// Render employee list
function renderEmployeeList() {
    const display = document.getElementById('employee-data');
    const clearBtn = document.getElementById('clear-all-btn');

    if (!display) return;

    if (employeeDataList.length === 0) {
        showEmptyState();
        return;
    }

    if (clearBtn) clearBtn.style.display = 'block';

    const listHTML = employeeDataList.map(emp => `
        <div class="employee-list-item ${emp.selected ? 'selected' : ''}" data-emp-id="${emp.id}">
            <div class="employee-checkbox-wrapper">
                <input type="checkbox" 
                       id="emp-check-${emp.id}" 
                       ${emp.selected ? 'checked' : ''}
                       onchange="toggleEmployeeSelection('${emp.id}')">
            </div>
            <div class="employee-info-compact">
                <h4>${emp.name} (ID: ${emp.id})</h4>
                <div class="info-row">
                    <div class="info-item"><strong>Position:</strong> ${emp.pos}</div>
                    <div class="info-item"><strong>Department:</strong> ${emp.dept}</div>
                    <div class="info-item"><strong>Branch:</strong> ${emp.branch}</div>
                    <div class="info-item"><strong>Hired:</strong> ${emp.hired}</div>
                </div>
            </div>
            <button class="btn-remove-employee" onclick="removeEmployeeFromList('${emp.id}')">
                ✕
            </button>
        </div>
    `).join('');

    display.innerHTML = `
        <div class="employee-list-container">
            ${listHTML}
        </div>
        <div style="margin-top:15px; padding:10px; background:#e8f5e9; border-radius:5px; color:#2e7d32; font-size:0.9em;">
            ✓ ${employeeDataList.length} employee(s) in list
        </div>
    `;

    updateSelectionCounter();
}

// Update selection counter
function updateSelectionCounter() {
    const counter = document.getElementById('selection-counter');
    if (!counter) return;

    const selectedCount = getSelectedEmployees().length;
    if (selectedCount > 0) {
        counter.textContent = `(${selectedCount} selected)`;
        counter.style.display = 'inline-block';
    } else {
        counter.style.display = 'none';
    }
}

// Update button states
function updateButtonStates() {
    const summaryBtn = document.getElementById('generate-summary');
    const panelBtn = document.getElementById('generate-panel-sheet');

    const selectedCount = getSelectedEmployees().length;

    if (panelBtn) {
        panelBtn.disabled = selectedCount === 0;
    }

    if (summaryBtn && employeeDataList.length > 0) {
        const selected = getSelectedEmployees();
        currentEmployeeData = selected.length > 0 ? selected[0] : employeeDataList[0];

        const isReg = currentEmployeeData.isReg;
        if (isReg) {
            if (globalEduData) {
                summaryBtn.disabled = false;
                summaryBtn.textContent = selectedCount > 1 ? `📄 Generate Summary (${selectedCount} employees)` : "📄 Generate Summary (.doc)";
            } else {
                summaryBtn.disabled = true;
                summaryBtn.textContent = "Upload Edu File to Enable";
            }
        } else {
            summaryBtn.disabled = false;
            summaryBtn.textContent = selectedCount > 1 ? `📄 Generate Summary (${selectedCount} employees)` : "📄 Generate Summary (.doc)";
        }
    } else if (summaryBtn) {
        summaryBtn.disabled = true;
    }
}

// Initialize Clear All button
const clearAllBtn = document.getElementById('clear-all-btn');
if (clearAllBtn) {
    clearAllBtn.addEventListener('click', clearAllEmployees);
}

// Override panel sheet generation for multi-employee
const panelSheetBtn = document.getElementById('generate-panel-sheet');
if (panelSheetBtn) {
    const newPanelBtn = panelSheetBtn.cloneNode(true);
    panelSheetBtn.parentNode.replaceChild(newPanelBtn, panelSheetBtn);

    newPanelBtn.addEventListener('click', () => {
        const selectedEmployees = getSelectedEmployees();
        if (selectedEmployees.length === 0) {
            alert("Please select at least one employee.");
            return;
        }
        generateMultiEmployeePanelSheets(selectedEmployees);
    });
}

// Override the Generate Summary button to handle multiple employees
const originalSummaryBtn = document.getElementById('generate-summary');
if (originalSummaryBtn) {
    // Clone to remove existing event listeners
    const newSummaryBtn = originalSummaryBtn.cloneNode(true);
    originalSummaryBtn.parentNode.replaceChild(newSummaryBtn, originalSummaryBtn);

    newSummaryBtn.addEventListener('click', () => {
        const selectedEmployees = getSelectedEmployees();

        if (selectedEmployees.length === 0) {
            alert("Please select at least one employee.");
            return;
        }

        if (selectedEmployees.length === 1) {
            // Single employee - use regular generation from script.js
            currentEmployeeData = selectedEmployees[0];
            if (currentEmployeeData.isReg) {
                generateRegularizationWordDoc(currentEmployeeData);
            } else {
                generatePromotionWordDoc(currentEmployeeData);
            }
        } else {
            // Multiple employees - generate stacked summaries
            generateStackedSummary(selectedEmployees);
        }
    });
}

// Generate stacked summary for multiple employees (one file, each with full headers)
function generateStackedSummary(employees) {
    if (!employees || employees.length === 0) {
        alert("No employees selected!");
        return;
    }

    const isReg = employees[0].isReg;

    // Generate employee sections - each with full header
    const employeeSections = employees.map((emp, index) => {
        const panelDate = new Date().toLocaleDateString();

        return `
        ${index > 0 ? '<div style="page-break-before: always;"></div>' : ''}
        
        <div class="line-1">ABENSON GROUP OF COMPANIES</div>
        <div class="line-2">Human Resource Department</div>
        <div class="space-17">&nbsp;</div>
        <div class="line-4">${isReg ? 'REGULARIZATION' : 'PROMOTION'} SUMMARY</div>
        <div class="space-20">&nbsp;</div>

        <table class="info-table">
            <tr><td><span class="label-normal">Name:</span> &nbsp;&nbsp; <span class="data-bold">${emp.name}</span></td><td class="right"><span class="label-normal">On-Board Date:</span> &nbsp;&nbsp; <span class="data-bold">${emp.hired}</span></td></tr>
            <tr><td><span class="label-normal">Position:</span> &nbsp;&nbsp; <span class="data-bold">${emp.pos}</span></td><td class="right"><span class="label-normal">Panel Interview:</span> &nbsp;&nbsp; <span class="data-bold">${panelDate}</span></td></tr>
            <tr><td colspan="2"><span class="label-normal">Branch/Department:</span> &nbsp;&nbsp; <span class="data-bold">${emp.branch} - ${emp.dept}</span></td></tr>
            <tr><td colspan="2"><span class="label-normal">Immediate Supervisor:</span> &nbsp;&nbsp; <span class="data-bold">${emp.mentor}</span></td></tr>
        </table>

        <div class="space-10">&nbsp;</div>

        <table class="forms-table" align="center">
            <tr class="forms-header"><td>${isReg ? 'REGULARIZATION FORMS' : 'PROMOTION FORMS'}</td><td>REMARKS</td></tr>
            <tr><td class="col-title">Recommendation & Essay</td><td class="col-remark">See attached files</td></tr>
            <tr><td class="col-title">Individual Score Card</td><td class="col-remark"></td></tr>
            <tr><td class="col-title">Personal Discipline Score [PDS]</td><td class="col-remark"></td></tr>
            <tr><td class="col-title">Internal Customer Rating</td><td class="col-remark"></td></tr>
            ${isReg ? `
            <tr><td class="col-title">Medical Result</td><td class="col-remark">Fit to work</td></tr>
            <tr><td class="col-title">Background Investigation</td><td class="col-remark">Cleared</td></tr>
            <tr>
                <td class="col-title" style="vertical-align: top; padding-top: 10px;">Transcript of Records</td>
                <td class="tor-cell">
                    <div class="tor-content">
                        ${emp.tor.school}<br>${emp.tor.major}<br>${emp.tor.years}
                    </div>
                </td>
            </tr>
            ` : `
            <tr><td class="col-title">BSC KPI Scorecard</td><td class="col-remark"></td></tr>
            <tr><td class="col-title">PDS Attendance</td><td class="col-remark"></td></tr>
            <tr><td class="col-title">Last Promotion Date</td><td class="col-remark">${emp.lastPromoDate}</td></tr>
            `}
        </table>

        <div class="space-16">&nbsp;</div>
        <div class="feedback-title">360 Feedback Summary</div>
        <br>
        <table class="feedback-table" align="center">
            <tr class="feedback-header"><td></td><td>RATING</td></tr>
            <tr><td class="col-competency">Self-Awareness</td><td class="col-rating"></td></tr>
            <tr><td class="col-competency">Drive for results</td><td class="col-rating"></td></tr>
            <tr><td class="col-competency">Communication</td><td class="col-rating"></td></tr>
            <tr><td class="col-competency">Leadership</td><td class="col-rating"></td></tr>
            <tr><td class="col-competency">Teamwork</td><td class="col-rating"></td></tr>
            <tr><td style="font-weight:bold; font-style:italic;">Average score (3.75 passing rate)</td><td class="col-rating" style="font-weight:bold">${emp.feedback.avg}</td></tr>
        </table>

        <br>

        <div style="width: 6.5in; margin: auto;">
            <div class="comments-title">Comments/Remarks</div>
            <div class="na-text">· N/A</div>
            <div class="space-11">&nbsp;</div>
            <div style="font-weight:bold; font-size:12pt;">Prepared by:</div>
            <div class="space-11">&nbsp;</div>
            <div class="space-11">&nbsp;</div>
            <div style="border-bottom: 1px solid black; width: 200px;"></div>
            <div class="comments-title">Julie Mae Sasutona</div>
            <div style="font-weight:bold; margin-top:5px; font-size:12pt;">HR Staff</div>
            <br>
            <br>
            <br>
            <br>
            <br>
            <br>
            <br>
            
        </div>
        `;
    }).join('');

    const content = `
    <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
        <meta charset="utf-8">
        <xml>
            <w:WordDocument>
                <w:View>Print</w:View>
                <w:Zoom>100</w:Zoom>
                <w:DoNotOptimizeForBrowser/>
            </w:WordDocument>
        </xml>
        <style>
            @page { size: 8.5in 14in; margin: 1in; mso-page-orientation: portrait; }
            body { font-family: 'Calibri', sans-serif; }
            .line-1 { font-size: 14pt; font-weight: bold; text-align: center; margin: 0; padding: 0; }
            .line-2 { font-size: 13pt; text-align: center; margin: 0; padding: 0; }
            .line-4 { font-size: 17pt; font-weight: bold; text-align: center; margin: 0; padding: 0; }
            .space-17 { font-size: 17pt; line-height: 17pt; }
            .space-20 { font-size: 20pt; line-height: 20pt; }
            .space-10 { font-size: 10pt; line-height: 10pt; }
            .space-16 { font-size: 16pt; line-height: 16pt; }
            .space-14 { font-size: 14pt; line-height: 14pt; } 
            .space-11 { font-size: 11pt; line-height: 11pt; } 
            .space-13 { font-size: 13pt; line-height: 13pt; } 
            .info-table { width: 100%; border-collapse: separate; border-spacing: 0 10px; font-size: 12pt; }
            .info-table td { padding: 2px 0; vertical-align: top; border: none; }
            .data-bold { font-weight: bold; }
            .label-normal { font-weight: normal; }
            .right { text-align: right; }
            .forms-table { width: 6.1in; border-collapse: collapse; margin-left: auto; margin-right: auto; font-size: 11pt; }
            .forms-table td { border: 1px solid black; padding: 2px; vertical-align: middle; height: 0.25in; }
            .forms-header td { font-weight: bold; text-align: center; background-color: #E7E6E6; padding: 4px; }
            .col-remark { text-align: center; }
            .col-title { font-weight: bold; }
            .tor-cell { padding: 0 !important; height: 100%; text-align: center; }
            .tor-content { padding: 2px; font-size: 11pt; line-height: 1.1; }
            .feedback-title { text-align: center; font-style: italic; font-weight: bold; font-size: 16pt; margin: 0; }
            .feedback-table { width: 3.7in; border-collapse: collapse; margin-left: auto; margin-right: auto; font-size: 11pt; }
            .feedback-table td { border: 1px solid black; padding: 2px; vertical-align: middle; height: 0.2in; }
            .feedback-header td { font-weight: bold; text-align: center; background-color: #E7E6E6; padding: 4px; }
            .col-rating { text-align: center; font-weight: bold; } 
            .col-competency { font-weight: bold; } 
            .comments-title { font-size: 14pt; font-weight: bold; }
            .na-text { font-size: 11pt; font-family: 'Calibri', sans-serif; }
        </style>
    </head>
    <body>
        ${employeeSections}
    </body>
    </html>
    `;

    const filePrefix = isReg ? "Regularization" : "Promotion";
    const dateStr = new Date().toISOString().split('T')[0];
    const filename = `${filePrefix}_Summary_${employees.length}employees_${dateStr}.doc`;

    saveDoc(content, filename);
    alert(`Successfully generated summary for ${employees.length} employee(s)!`);
}

// Generate panel sheets for multiple employees
function generateMultiEmployeePanelSheets(employees) {
    const wb = XLSX.utils.book_new();

    employees.forEach((emp) => {
        const ws = createPanelSheet(emp);
        let sheetName = `${emp.name}`.substring(0, 28);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    const filePrefix = employees[0].isReg ? "Regularization" : "Promotion";
    const dateStr = new Date().toISOString().split('T')[0];
    const filename = filePrefix + "_Panels_" + dateStr + ".xlsx";

    // Explicitly set bookType to ensure proper .xlsx download
    XLSX.writeFile(wb, filename, { bookType: 'xlsx' });

    alert(`Successfully generated ${employees.length} panel sheet(s)!`);
}

// Create a single panel sheet (same as original logic)
function createPanelSheet(emp) {
    const ws = XLSX.utils.aoa_to_sheet([]);

    const borderStyle = { style: "thin", color: { rgb: "000000" } };
    const allBorders = { top: borderStyle, bottom: borderStyle, left: borderStyle, right: borderStyle };
    const noTopBorder = { top: { style: "none" }, bottom: borderStyle, left: borderStyle, right: borderStyle };
    const noBottomBorder = { top: borderStyle, bottom: { style: "none" }, left: borderStyle, right: borderStyle };

    const f18b = { font: { name: "Calibri", sz: 18, bold: true } };
    const f14b = { font: { name: "Calibri", sz: 14, bold: true } };

    const scoreCellStyle = { border: allBorders, font: { name: "Calibri", sz: 18, bold: true }, alignment: { horizontal: "center", vertical: "center" } };
    const headerGray = { fill: { fgColor: { rgb: "E7E6E6" } }, border: allBorders, font: { name: "Calibri", sz: 10, bold: true }, alignment: { horizontal: "center", vertical: "center", wrapText: true } };
    const questionWhite = { border: allBorders, font: { name: "Calibri", sz: 14 }, alignment: { wrapText: true, vertical: "center" } };
    const headerMajorParts = { fill: { fgColor: { rgb: "E7E6E6" } }, border: allBorders, font: { name: "Calibri", sz: 16, bold: true }, alignment: { horizontal: "center", vertical: "center" } };
    const headerCyan = { fill: { fgColor: { rgb: "00FFFF" } }, border: allBorders, font: { name: "Calibri", sz: 14, bold: true }, alignment: { horizontal: "center", vertical: "center" } };
    const headerComments = { fill: { fgColor: { rgb: "E7E6E6" } }, border: allBorders, font: { name: "Calibri", sz: 14, bold: true }, alignment: { horizontal: "center", vertical: "center" } };
    const processActiveNoBorder = { fill: { fgColor: { rgb: "FFFF00" } }, font: { name: "Calibri", sz: 18, bold: true }, alignment: { horizontal: "left", vertical: "center" } };
    const processInactiveNoBorder = { font: { name: "Calibri", sz: 18 }, alignment: { horizontal: "left", vertical: "center" } };
    const recommendActive = { fill: { fgColor: { rgb: "FFFF00" } }, font: { name: "Calibri", sz: 14, bold: true } };
    const recommendInactive = { font: { name: "Calibri", sz: 14, bold: true } };
    const inputCell = { border: allBorders, font: { name: "Calibri", sz: 14 }, alignment: { horizontal: "center", vertical: "center" } };
    const blueInputCell = { border: allBorders, fill: { fgColor: { rgb: "99CCFF" } }, font: { name: "Calibri", sz: 18, bold: true }, alignment: { horizontal: "center", vertical: "center" } };
    const blueInputNoTop = { border: noTopBorder, fill: { fgColor: { rgb: "99CCFF" } }, font: { name: "Calibri", sz: 18, bold: true }, alignment: { horizontal: "center", vertical: "center" } };
    const blueInputNoBottom = { border: noBottomBorder, fill: { fgColor: { rgb: "99CCFF" } }, font: { name: "Calibri", sz: 18, bold: true }, alignment: { horizontal: "center", vertical: "center" } };

    place(ws, "A3", "Employee Name: " + emp.name, f18b);
    place(ws, "A4", "Present Position: " + emp.pos, f18b);

    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s: { r: 3, c: 5 }, e: { r: 3, c: 7 } });

    const regText = emp.isReg ? "☑ REGULARIZATION" : "☐ REGULARIZATION";
    const promoText = emp.isReg ? "☐ PROMOTION" : "☑ PROMOTION";

    place(ws, "F4", regText, emp.isReg ? processActiveNoBorder : processInactiveNoBorder);
    place(ws, "I4", promoText, emp.isReg ? processInactiveNoBorder : processActiveNoBorder);
    styleRange(ws, 3, 3, 5, 7, emp.isReg ? processActiveNoBorder : processInactiveNoBorder);

    place(ws, "A5", "Position Recommended for: " + emp.targetPos, f18b);
    place(ws, "A6", "Department/Branch: " + emp.branch + " - " + emp.dept, f18b);
    place(ws, "E6", "Mentor: " + emp.mentor, f18b);
    place(ws, "A7", "OBD: " + emp.hired, f18b);

    place(ws, "A9", "MAJOR PARTS", headerMajorParts);
    const scale = ["Excellent", "Good", "Average", "Below Average", "Poor"];
    scale.forEach((txt, i) => {
        const cell = XLSX.utils.encode_cell({ r: 8, c: i + 1 });
        place(ws, cell, txt, headerGray);
    });
    place(ws, "G9", "Grade", headerCyan);
    place(ws, "H9", "Ave Score", headerGray);
    place(ws, "I9", "COMMENTS/REMARKS", headerComments);

    const rows = [
        { t: "Part 1. Job Knowledge Competency", type: "title", ref: "p1" },
        { t: "▪ Demonstrates the knowledge, skills and abilities necessary to perform his/her duties and responsibilities", type: "q" },
        { t: "▪ Demonstrate accountability and honesty in the conduct of his/her job", type: "q" },
        { t: "▪ Completes assigned tasks efficiently and effectively", type: "q" },
        { t: "Part 2. Quality Attitude (Values)", type: "title", ref: "p2" },
        { t: "▪ Fit to the organization and culture", type: "q" },
        { t: "▪ Client/Customer Focus", type: "q" },
        { t: "▪ Ability to work and develop harmonious relationship with colleagues", type: "q" },
        { t: "▪ Open to feedback and with positive attitude", type: "q" },
        { t: "Part 3. Career Growth", type: "title", ref: "p3" },
        { t: "▪ Long-term plan with company", type: "q" },
        { t: "▪ Shows concern for career growth", type: "q" },
        { t: "Total", type: "total" }
    ];

    let r = 9;
    rows.forEach((row) => {
        const rowNum = r + 1;
        if (row.type === "title") {
            place(ws, `A${rowNum}`, row.t, { border: allBorders, font: { name: "Calibri", sz: 16, bold: true }, alignment: { wrapText: true, vertical: "center" } });
            const count = row.ref === 'p1' ? 3 : (row.ref === 'p2' ? 4 : 2);
            const range = `G${rowNum + 1}:G${rowNum + count}`;
            place(ws, `H${rowNum}`, { t: 'n', f: `AVERAGE(${range})` }, { ...f14b, border: allBorders, alignment: { horizontal: "center" } });
            place(ws, `G${rowNum}`, "", blueInputCell);
            for (let c = 1; c <= 5; c++) place(ws, XLSX.utils.encode_cell({ r, c }), "", { border: allBorders });
        } else if (row.type === "q") {
            place(ws, `A${rowNum}`, row.t, questionWhite);
            [5, 4, 3, 2, 1].forEach((s, i) => {
                place(ws, XLSX.utils.encode_cell({ r, c: i + 1 }), s, scoreCellStyle);
            });
            let currentGStyle = blueInputCell;
            if (r === 10) currentGStyle = blueInputNoTop;
            if (r === 11) currentGStyle = blueInputNoBottom;
            if (r === 14) currentGStyle = blueInputNoTop;
            if (r === 16) currentGStyle = blueInputNoBottom;
            if (r === 19) currentGStyle = blueInputNoTop;
            place(ws, `G${rowNum}`, "", currentGStyle);
            place(ws, `H${rowNum}`, "", inputCell);
            place(ws, `I${rowNum}`, "", { border: allBorders });
        } else if (row.type === "total") {
            place(ws, `A${rowNum}`, "Total", { border: allBorders, font: { name: "Calibri", sz: 11, bold: true }, fill: { fgColor: { rgb: "E7E6E6" } } });
            place(ws, `G${rowNum}`, "", blueInputCell);
            place(ws, `H${rowNum}`, { t: 'n', f: "SUM(H10,H14,H19)" }, { ...f14b, border: allBorders, color: { rgb: "FF0000" }, alignment: { horizontal: "center" } });
            for (let c = 1; c <= 5; c++) place(ws, XLSX.utils.encode_cell({ r, c }), "", { border: allBorders });
            place(ws, `I${rowNum}`, "", { border: allBorders });
        }
        r++;
    });

    ws['!merges'].push({ s: { r: 9, c: 8 }, e: { r: 12, c: 8 } });
    ws['!merges'].push({ s: { r: 13, c: 8 }, e: { r: 17, c: 8 } });
    ws['!merges'].push({ s: { r: 18, c: 8 }, e: { r: 20, c: 8 } });
    const sCommentBox = { border: allBorders, alignment: { vertical: "top", wrapText: true } };
    styleRange(ws, 9, 12, 8, 8, sCommentBox);
    styleRange(ws, 13, 17, 8, 8, sCommentBox);
    styleRange(ws, 18, 20, 8, 8, sCommentBox);

    const passGradeStyle = { font: { name: "Calibri", sz: 14, bold: true, italic: true }, alignment: { horizontal: "right" } };
    place(ws, "A23", "Passing grade: 12.5 -up", passGradeStyle);
    ws['!merges'].push({ s: { r: 22, c: 0 }, e: { r: 22, c: 7 } });

    place(ws, "A24", "Recommendation:", f14b);
    if (emp.isReg) {
        place(ws, "A25", "□ OK FOR REGULARIZATION", recommendActive);
        place(ws, "A26", "□ NOT FOR REGULARIZATION", recommendInactive);
    } else {
        place(ws, "A25", "□ RECOMMENDED FOR PROMOTION", recommendActive);
        place(ws, "A26", "□ NOT RECOMMENDED", recommendInactive);
    }
    place(ws, "A28", "Other Comments and Justification:", f14b);
    ws['!merges'].push({ s: { r: 28, c: 0 }, e: { r: 32, c: 8 } });
    styleRange(ws, 28, 32, 0, 8, sCommentBox);
    place(ws, "A36", "Panelist's Name:", f14b);
    place(ws, "B36", "Date:", f14b);

    ws['!cols'] = [{ wch: 68.41 }, { wch: 7.11 }, { wch: 7.11 }, { wch: 7.11 }, { wch: 7.11 }, { wch: 7.11 }, { wch: 12.88 }, { wch: 17.71 }, { wch: 56.89 }];

    return ws;
}

function place(ws, ref, val, style) {
    XLSX.utils.sheet_add_aoa(ws, [[val]], { origin: ref });
    if (ws[ref]) ws[ref].s = style;
}

function styleRange(ws, rStart, rEnd, cStart, cEnd, style) {
    for (let r = rStart; r <= rEnd; r++) {
        for (let c = cStart; c <= cEnd; c++) {
            const cellRef = XLSX.utils.encode_cell({ r, c });
            if (!ws[cellRef]) ws[cellRef] = { v: "", t: "s" };
            ws[cellRef].s = style;
        }
    }
}

// Override form submission to show preview instead of adding directly
document.addEventListener('DOMContentLoaded', function () {
    const dataForm = document.getElementById('data-lookup-form');
    if (!dataForm) return;

    dataForm.addEventListener('submit', async function (e) {
        e.preventDefault();
        e.stopImmediatePropagation();

        const isReg = document.title.includes('Regularization');
        const mainFileInput = document.getElementById('excel-upload');
        const eduFileInput = document.getElementById('edu-upload');
        const idInput = document.getElementById('employee-id').value.trim();

        // Load Master Data
        if (mainFileInput.files.length > 0) {
            try {
                globalMasterData = await readExcelFile(mainFileInput.files[0]);
                console.log("Master DB Loaded");
            } catch (err) {
                alert("Error reading Master File.");
                return;
            }
        } else if (!globalMasterData) {
            alert("Please upload the Master Database file first.");
            return;
        }

        // Load Education Data (optional)
        if (eduFileInput && eduFileInput.files.length > 0) {
            try {
                globalEduData = await readExcelFile(eduFileInput.files[0]);
                console.log("Education DB Loaded");
            } catch (err) {
                console.warn("Error reading Edu File.");
            }
        }

        // Search Employee
        if (!idInput) {
            alert("Please enter an Employee ID.");
            return;
        }

        const emp = globalMasterData.find(row => {
            const dbId = String(row['EmployeeNo']).trim();
            const inputId = idInput.replace(/\D/g, '');
            return dbId === idInput || dbId === inputId;
        });

        if (!emp) {
            alert("Employee ID not found in loaded database.");
            return;
        }

        // Search Education Record
        let eduRecord = {};
        if (globalEduData) {
            eduRecord = globalEduData.find(row => {
                const dbId = String(row['EMPNO']).trim().replace(/^0+/, '');
                const targetId = String(emp['EmployeeNo']).trim().replace(/^0+/, '');
                return dbId === targetId;
            }) || {};
        }

        // Map Data
        const employeeData = {
            id: emp['EmployeeNo'],
            name: (emp['EmployeeName'] || "").toUpperCase(),
            pos: (emp['EmployeePosition'] || "").toUpperCase(),
            dept: (emp['Department'] || "").toUpperCase(),
            branch: (emp['WorkLocation'] || "").toUpperCase(),
            mentor: (emp['DirectSupervisor'] || "N/A").toUpperCase(),
            hired: formatDate(emp['Hired Date']),
            targetPos: isReg ? 'REGULARIZATION' : 'SUPERVISOR (PROPOSED)',
            isReg: isReg,
            feedback: {
                avg: emp['Feedback_Avg'] || ''
            },
            tor: {
                school: (eduRecord['SCHOOL'] || "").toUpperCase(),
                major: (eduRecord['MAJOR'] || "").toUpperCase(),
                years: (eduRecord['STARTYEAR'] && eduRecord['ENDYEAR'])
                    ? `${eduRecord['STARTYEAR']} - ${eduRecord['ENDYEAR']}`
                    : ""
            },
            lastPromoDate: "N/A"
        };

        // Show preview with Add button
        showEmployeePreview(employeeData);
    }, true);
});
