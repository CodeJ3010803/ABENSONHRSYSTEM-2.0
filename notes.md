// ====================================================================
// 1. GLOBAL UTILITIES
// ====================================================================
document.addEventListener('DOMContentLoaded', () => {
    // Password Toggle
    document.querySelectorAll('.toggle-password').forEach(btn => {
        btn.addEventListener('click', function() {
            const input = this.parentElement.querySelector('input');
            if (input) {
                const type = input.getAttribute('type') === 'password' ? 'text' : 'password';
                input.setAttribute('type', type);
                this.textContent = type === 'password' ? '👁️' : '🙈';
            }
        });
    });

    // Registration Popup
    const regForm = document.getElementById('register-form');
    if (regForm) {
        regForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const firstName = document.getElementById('first-name').value;
            const lastName = document.getElementById('last-name').value;
            const email = document.getElementById('email').value;
            const password = document.getElementById('password').value;
            const confirm = document.getElementById('confirm-password').value;

            if (password !== confirm) {
                alert("Passwords do not match!");
                return;
            }

            const users = JSON.parse(localStorage.getItem('abenson_hr_users') || '[]');
            if (users.find(u => u.email === email)) {
                alert("This email is already registered.");
                return;
            }

            users.push({ firstName, lastName, email, password });
            localStorage.setItem('abenson_hr_users', JSON.stringify(users));
            document.getElementById('registration-popup').style.display = 'flex';
        });

        document.getElementById('popup-ok-btn').addEventListener('click', () => {
            window.location.href = 'index.html';
        });
    }

    // Login Logic
    const loginForm = document.getElementById('login-form');
    if (loginForm) {
        loginForm.addEventListener('submit', (e) => {
            e.preventDefault();
            const email = document.getElementById('email').value;
            const password = document.getElementById('password').value;
            const users = JSON.parse(localStorage.getItem('abenson_hr_users') || '[]');
            const validUser = users.find(u => u.email === email && u.password === password);

            if (validUser) {
                sessionStorage.setItem('currentUser', JSON.stringify(validUser));
                window.location.href = 'home.html';
            } else {
                alert("Invalid email or password.");
            }
        });
    }

    // Home Page Personalization
    const welcomeMsg = document.getElementById('welcome-message');
    if (welcomeMsg) {
        const currentUser = JSON.parse(sessionStorage.getItem('currentUser'));
        if (currentUser && currentUser.firstName) {
            welcomeMsg.textContent = `Welcome Back, ${currentUser.firstName}!`;
        }
    }
});

// ====================================================================
// 2. DATABASE LOADING & MAPPING
// ====================================================================
let currentEmployeeData = null;

const dataForm = document.getElementById('data-lookup-form');
if (dataForm) {
    dataForm.addEventListener('submit', function(e) {
        e.preventDefault();
        
        const isReg = document.title.includes('Regularization');
        const fileInput = document.getElementById('excel-upload');
        const idInput = document.getElementById('employee-id').value.trim();

        if (!fileInput.files[0] || !idInput) {
            alert("Please upload the database file and enter an ID.");
            return;
        }

        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(sheet);

            const emp = jsonData.find(row => {
                const dbId = String(row['EmployeeNo']).trim();
                const inputId = idInput.replace(/\D/g,''); 
                return dbId === idInput || dbId === inputId;
            });

            if (!emp) {
                alert("Employee ID not found.");
                return;
            }

            currentEmployeeData = {
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
                }
            };

            const display = document.getElementById('employee-data');
            display.innerHTML = `
                <h3>2. Employee Details</h3>
                <div class="employee-card-details">
                    <p><strong>Name:</strong> ${currentEmployeeData.name}</p>
                    <p><strong>ID:</strong> ${currentEmployeeData.id}</p>
                    <p><strong>Current Pos:</strong> ${currentEmployeeData.pos}</p>
                    <p><strong>Target Pos:</strong> ${currentEmployeeData.targetPos}</p>
                    <p><strong>Dept/Branch:</strong> ${currentEmployeeData.branch} - ${currentEmployeeData.dept}</p>
                    <p><strong>Mentor:</strong> ${currentEmployeeData.mentor}</p>
                    <p><strong>Hired Date:</strong> ${currentEmployeeData.hired}</p>
                </div>
                <div style="margin-top:15px; padding:10px; background:#e8f5e9; border-radius:5px; color:#2e7d32; font-size:0.9em;">
                    ✓ Data Loaded Successfully
                </div>
            `;
            
            const summaryBtn = document.getElementById('generate-summary');
            summaryBtn.disabled = false;
            summaryBtn.textContent = "Generate Summary (.doc)"; 
            
            document.getElementById('generate-panel-sheet').disabled = false;
        };
        reader.readAsArrayBuffer(fileInput.files[0]);
    });
}

function formatDate(excelDate) {
    if (!excelDate) return "";
    if (typeof excelDate === 'number') {
        const date = new Date(Math.round((excelDate - 25569) * 86400 * 1000));
        return date.toLocaleDateString();
    }
    return excelDate;
}

// ====================================================================
// 3. WORD DOCUMENT GENERATION (UPDATED LAYOUT)
// ====================================================================

document.getElementById('generate-summary')?.addEventListener('click', () => {
    if (!currentEmployeeData) return;
    
    if (currentEmployeeData.isReg) {
        generateRegularizationWordDoc(currentEmployeeData);
    } else {
        const txt = `PROMOTION SUMMARY\n\nName: ${currentEmployeeData.name}\nPosition: ${currentEmployeeData.pos}`;
        const blob = new Blob([txt], { type: 'text/plain' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = `Promotion_Summary_${currentEmployeeData.id}.txt`;
        link.click();
    }
});

function generateRegularizationWordDoc(emp) {
    const panelDate = new Date().toLocaleDateString();

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
            @page {
                size: 8.5in 14in; /* Legal Size */
                margin: 1in;
                mso-page-orientation: portrait;
            }
            body { 
                font-family: 'Calibri', sans-serif;
                font-size: 11pt;
            }
            
            /* HEADERS */
            .line-1 { font-size: 14pt; font-weight: bold; text-align: center; margin: 0; }
            .line-2 { font-size: 13pt; text-align: center; margin: 0; }
            .line-4 { font-size: 17pt; font-weight: bold; text-align: center; margin: 0; }

            /* EMPLOYEE INFO TABLE (NO BOLD PLACEHOLDERS, BOLD DATA) */
            .info-table {
                width: 100%;
                border-collapse: collapse;
                margin-top: 30px;
                font-size: 12pt; /* Calibri 12 */
            }
            .info-table td {
                padding: 2px;
                vertical-align: top;
                border: none; 
            }
            .data-bold { font-weight: bold; }
            .right { text-align: right; }

            /* SPACER */
            .spacer { height: 30px; }

            /* FORMS TABLE (5 INCHES WIDTH) */
            .forms-table {
                width: 5in; 
                border-collapse: collapse;
                margin-left: auto; 
                margin-right: auto; /* Centered */
                font-size: 11pt;
            }
            .forms-table td {
                border: 1px solid black;
                padding: 5px;
                vertical-align: middle;
            }
            .forms-header td {
                font-weight: bold;
                text-align: center;
                background-color: #E7E6E6;
            }
            .col-remark { text-align: center; } /* Remarks Centered */

            /* 360 FEEDBACK TABLE (3 INCHES WIDTH) */
            .feedback-title {
                text-align: center;
                font-style: italic;
                margin-top: 30px;
                margin-bottom: 10px;
                font-weight: bold;
            }
            .feedback-table {
                width: 3in;
                border-collapse: collapse;
                margin-left: auto;
                margin-right: auto; /* Centered */
                font-size: 11pt;
            }
            .feedback-table td {
                border: 1px solid black;
                padding: 5px;
                vertical-align: middle;
            }
            .feedback-header td {
                font-weight: bold;
                text-align: center;
                background-color: #E7E6E6;
            }
            .col-rating { text-align: center; } /* Rating Centered */

        </style>
    </head>
    <body>
        <div class="line-1">ABENSON GROUP OF COMPANIES</div>
        <div class="line-2">Human Resource Department</div>
        <br>
        <div class="line-4">REGULARIZATION SUMMARY</div>

        <!-- EMPLOYEE INFO -->
        <table class="info-table">
            <tr>
                <td>Name: &nbsp; <span class="data-bold">${emp.name}</span></td>
                <td class="right">On-Board Date: &nbsp; <span class="data-bold">${emp.hired}</span></td>
            </tr>
            <tr>
                <td>Position: &nbsp; <span class="data-bold">${emp.pos}</span></td>
                <td class="right">Panel Interview: &nbsp; <span class="data-bold">${panelDate}</span></td>
            </tr>
            <tr>
                <td colspan="2">Branch/Department: &nbsp; <span class="data-bold">${emp.branch} - ${emp.dept}</span></td>
            </tr>
            <tr>
                <td colspan="2">Immediate Supervisor: &nbsp; <span class="data-bold">${emp.mentor}</span></td>
            </tr>
        </table>

        <div class="spacer"></div>

        <!-- TABLE 1: REGULARIZATION FORMS (5 inches wide) -->
        <table class="forms-table">
            <tr class="forms-header">
                <td>REGULARIZATION FORMS</td>
                <td>REMARKS</td>
            </tr>
            <tr><td>Recommendation & Essay</td><td class="col-remark">See attached files</td></tr>
            <tr><td>Individual Score Card</td><td class="col-remark"></td></tr>
            <tr><td>Personal Discipline Score [PDS]</td><td class="col-remark"></td></tr>
            <tr><td>Internal Customer Rating</td><td class="col-remark"></td></tr>
            <tr><td>Medical Result</td><td class="col-remark">Fit to work</td></tr>
            <tr><td>Background Investigation</td><td class="col-remark">Cleared</td></tr>
            <tr><td>Transcript of Records</td><td class="col-remark"></td></tr>
        </table>

        <!-- TABLE 2: 360 FEEDBACK (3 inches wide) -->
        <div class="feedback-title">360 Feedback Summary</div>
        <table class="feedback-table">
            <tr class="feedback-header">
                <td>COMPETENCY</td>
                <td>RATING</td>
            </tr>
            <tr><td>Self-Awareness</td><td class="col-rating"></td></tr>
            <tr><td>Drive for results</td><td class="col-rating"></td></tr>
            <tr><td>Communication</td><td class="col-rating"></td></tr>
            <tr><td>Leadership</td><td class="col-rating"></td></tr>
            <tr><td>Teamwork</td><td class="col-rating"></td></tr>
            <tr><td style="font-weight:bold">AVERAGE</td><td class="col-rating" style="font-weight:bold">${emp.feedback.avg}</td></tr>
        </table>

        <!-- COMMENTS & PREPARED BY -->
        <br><br>
        <div style="width: 6.5in; margin: auto;">
            <div style="font-weight:bold">Comments/Remarks:</div>
            <div style="height: 50px; border-bottom: 1px solid black;"></div>
            <br><br>
            <div style="font-weight:bold">Prepared by:</div>
            <br>
            <div style="border-bottom: 1px solid black; width: 200px;"></div>
            <div style="font-weight:bold; margin-top:5px;">HR Staff</div>
        </div>

    </body>
    </html>
    `;

    const blob = new Blob(['\ufeff', content], { type: 'application/msword' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `Regularization_Summary_${emp.id}.doc`; 
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// ====================================================================
// 4. EXCEL PANEL SHEET LOGIC (UNCHANGED)
// ====================================================================

document.getElementById('generate-panel-sheet')?.addEventListener('click', () => {
    if (!currentEmployeeData) return;
    const emp = currentEmployeeData;
    const wb = XLSX.utils.book_new();
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

    if(!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s: {r:3, c:5}, e: {r:3, c:7} }); 

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
        const cell = XLSX.utils.encode_cell({r: 8, c: i+1});
        place(ws, cell, txt, headerGray);
    });
    place(ws, "G9", "Grade", headerCyan);
    place(ws, "H9", "Ave Score", headerGray);
    place(ws, "I9", "COMMENTS/REMARKS", headerComments);

    const rows = [
        {t: "Part 1. Job Knowledge Competency", type: "title", ref: "p1"},
        {t: "▪ Demonstrates the knowledge, skills and abilities necessary to perform his/her duties and responsibilities", type: "q"},
        {t: "▪ Demonstrate accountability and honesty in the conduct of his/her job", type: "q"},
        {t: "▪ Completes assigned tasks efficiently and effectively", type: "q"},
        {t: "Part 2. Quality Attitude (Values)", type: "title", ref: "p2"},
        {t: "▪ Fit to the organization and culture", type: "q"},
        {t: "▪ Client/Customer Focus", type: "q"},
        {t: "▪ Ability to work and develop harmonious relationship with colleagues", type: "q"},
        {t: "▪ Open to feedback and with positive attitude", type: "q"},
        {t: "Part 3. Career Growth", type: "title", ref: "p3"},
        {t: "▪ Long-term plan with company", type: "q"},
        {t: "▪ Shows concern for career growth", type: "q"},
        {t: "Total", type: "total"}
    ];

    let r = 9; 
    rows.forEach((row) => {
        const rowNum = r + 1;
        if (row.type === "title") {
             place(ws, `A${rowNum}`, row.t, { border: allBorders, font: { name: "Calibri", sz: 16, bold: true }, alignment: { wrapText: true, vertical: "center" } });
             const count = row.ref === 'p1' ? 3 : (row.ref === 'p2' ? 4 : 2);
             const range = `G${rowNum+1}:G${rowNum+count}`;
             place(ws, `H${rowNum}`, {t:'n', f:`AVERAGE(${range})`}, { ...f14b, border: allBorders, alignment: { horizontal: "center" } });
             place(ws, `G${rowNum}`, "", blueInputCell);
             for(let c=1; c<=5; c++) place(ws, XLSX.utils.encode_cell({r, c}), "", { border: allBorders });
        } else if (row.type === "q") {
             place(ws, `A${rowNum}`, row.t, questionWhite);
             [5,4,3,2,1].forEach((s, i) => {
                 place(ws, XLSX.utils.encode_cell({r, c: i+1}), s, scoreCellStyle);
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
             place(ws, `H${rowNum}`, {t:'n', f:"SUM(H10,H14,H19)"}, { ...f14b, border: allBorders, color: { rgb: "FF0000" }, alignment: { horizontal: "center" } });
             for(let c=1; c<=5; c++) place(ws, XLSX.utils.encode_cell({r, c}), "", { border: allBorders });
             place(ws, `I${rowNum}`, "", { border: allBorders });
        }
        r++;
    });

    ws['!merges'].push({ s: {r:9, c:8}, e: {r:12, c:8} }); 
    ws['!merges'].push({ s: {r:13, c:8}, e: {r:17, c:8} }); 
    ws['!merges'].push({ s: {r:18, c:8}, e: {r:20, c:8} }); 
    const sCommentBox = { border: allBorders, alignment: { vertical: "top", wrapText: true } };
    styleRange(ws, 9, 12, 8, 8, sCommentBox);
    styleRange(ws, 13, 17, 8, 8, sCommentBox);
    styleRange(ws, 18, 20, 8, 8, sCommentBox);

    const passGradeStyle = { font: { name: "Calibri", sz: 14, bold: true, italic: true }, alignment: { horizontal: "right" } };
    place(ws, "A23", "Passing grade: 12.5 -up", passGradeStyle);
    ws['!merges'].push({ s: {r:22, c:0}, e: {r:22, c:7} }); 
    
    place(ws, "A24", "Recommendation:", f14b);
    if (emp.isReg) {
        place(ws, "A25", "□ OK FOR REGULARIZATION", recommendActive);
        place(ws, "A26", "□ NOT FOR REGULARIZATION", recommendInactive);
    } else {
        place(ws, "A25", "□ RECOMMENDED FOR PROMOTION", recommendActive);
        place(ws, "A26", "□ NOT RECOMMENDED", recommendInactive);
    }
    place(ws, "A28", "Other Comments and Justification:", f14b);
    ws['!merges'].push({ s: {r:28, c:0}, e: {r:32, c:8} }); 
    styleRange(ws, 28, 32, 0, 8, sCommentBox); 
    place(ws, "A36", "Panelist's Name:", f14b);
    place(ws, "B36", "Date:", f14b);

    ws['!cols'] = [{wch: 68.41}, {wch: 7.11}, {wch: 7.11}, {wch: 7.11}, {wch: 7.11}, {wch: 7.11}, {wch: 12.88}, {wch: 17.71}, {wch: 56.89}];
    const filePrefix = emp.isReg ? "Regularization" : "Promotion";
    XLSX.utils.book_append_sheet(wb, ws, "Panel Sheet");
    XLSX.writeFile(wb, `${filePrefix}_Panel_${emp.id}.xlsx`);
});

function place(ws, ref, val, style) {
    XLSX.utils.sheet_add_aoa(ws, [[val]], { origin: ref });
    if (ws[ref]) ws[ref].s = style;
}

function styleRange(ws, rStart, rEnd, cStart, cEnd, style) {
    for (let r = rStart; r <= rEnd; r++) {
        for (let c = cStart; c <= cEnd; c++) {
            const cellRef = XLSX.utils.encode_cell({r, c});
            if (!ws[cellRef]) ws[cellRef] = { v: "", t: "s" }; 
            ws[cellRef].s = style;
        }
    }
}