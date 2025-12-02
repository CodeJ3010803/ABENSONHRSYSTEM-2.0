// ====================================================================
// 1. GLOBAL UTILITIES
// ====================================================================

// GLOBAL DATA STORAGE
let globalMasterData = null;
let globalEduData = null;
// GLOBAL CURRENT EMPLOYEES (Array for batch processing)
let currentEmployeesBatch = []; 

document.addEventListener('DOMContentLoaded', () => {
    // ... (Password Toggle & Auth Logic remains same) ...
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
// 2. DATABASE LOADING & MAPPING (UPDATED FOR BATCH)
// ====================================================================
const dataForm = document.getElementById('data-lookup-form');
if (dataForm) {
    dataForm.addEventListener('submit', async function(e) {
        e.preventDefault();
        
        const isReg = document.title.includes('Regularization');
        const mainFileInput = document.getElementById('excel-upload');
        const eduFileInput = document.getElementById('edu-upload');
        // Allow comma separated IDs: "70, 111, 141"
        const idInputRaw = document.getElementById('employee-id').value.trim();
        
        // Parse IDs into an array
        const targetIds = idInputRaw.split(',').map(id => id.trim()).filter(id => id !== "");

        // --- 1. LOAD MASTER DATA ---
        if (mainFileInput.files.length > 0) {
            try {
                globalMasterData = await readExcelFile(mainFileInput.files[0]);
                console.log("Master DB Loaded/Updated");
            } catch (err) {
                alert("Error reading Master File.");
                return;
            }
        } else if (!globalMasterData) {
            alert("Please upload the Master Database file first.");
            return;
        }

        // --- 2. LOAD EDUCATION DATA (OPTIONAL) ---
        if (eduFileInput && eduFileInput.files.length > 0) {
            try {
                globalEduData = await readExcelFile(eduFileInput.files[0]);
                console.log("Education DB Loaded/Updated");
            } catch (err) {
                console.warn("Error reading Edu File, skipping.");
            }
        } 

        // --- 3. SEARCH EMPLOYEES (BATCH) ---
        if (targetIds.length === 0) {
            alert("Please enter at least one Employee ID.");
            return;
        }

        currentEmployeesBatch = []; // Reset batch

        targetIds.forEach(targetId => {
            // Search logic
            const emp = globalMasterData.find(row => {
                const dbId = String(row['EmployeeNo']).trim();
                const inputId = targetId.replace(/\D/g,''); 
                return dbId === targetId || dbId === inputId;
            });

            if (emp) {
                // Find Edu Record
                let eduRecord = {};
                if (globalEduData) {
                    eduRecord = globalEduData.find(row => {
                        const dbId = String(row['EMPNO']).trim().replace(/^0+/, ''); 
                        const empNo = String(emp['EmployeeNo']).trim().replace(/^0+/, '');
                        return dbId === empNo;
                    }) || {};
                }

                // Map Data
                currentEmployeesBatch.push({
                    id: emp['EmployeeNo'],
                    name: (emp['EmployeeName'] || "").toUpperCase(),
                    pos: (emp['EmployeePosition'] || "").toUpperCase(),
                    dept: (emp['Department'] || "").toUpperCase(),
                    branch: (emp['WorkLocation'] || "").toUpperCase(),
                    mentor: (emp['DirectSupervisor'] || "N/A").toUpperCase(),
                    hired: formatDate(emp['Hired Date']),
                    targetPos: isReg ? 'REGULARIZATION' : 'SUPERVISOR (PROPOSED)',
                    isReg: isReg,
                    feedback: { avg: emp['Feedback_Avg'] || '' },
                    tor: {
                        school: (eduRecord['SCHOOL'] || "").toUpperCase(),
                        major: (eduRecord['MAJOR'] || "").toUpperCase(),
                        years: (eduRecord['STARTYEAR'] && eduRecord['ENDYEAR']) 
                            ? `${eduRecord['STARTYEAR']} - ${eduRecord['ENDYEAR']}` 
                            : ""
                    },
                    lastPromoDate: "N/A"
                });
            }
        });

        if (currentEmployeesBatch.length === 0) {
            alert("No matching employees found.");
            return;
        }

        // --- 4. DISPLAY DATA (First Found or Summary) ---
        const display = document.getElementById('employee-data');
        
        if (currentEmployeesBatch.length === 1) {
            // Single Result Display (Same as before)
            const emp = currentEmployeesBatch[0];
            const hasEdu = !!(emp.tor.school); 
            display.innerHTML = `
                <h3>2. Employee Details</h3>
                <div class="employee-card-details">
                    <p><strong>Name:</strong> ${emp.name}</p>
                    <p><strong>ID:</strong> ${emp.id}</p>
                    <p><strong>Current Pos:</strong> ${emp.pos}</p>
                    <p><strong>Target Pos:</strong> ${emp.targetPos}</p>
                    <p><strong>Dept/Branch:</strong> ${emp.branch} - ${emp.dept}</p>
                    <p><strong>Mentor:</strong> ${emp.mentor}</p>
                    <p><strong>Hired Date:</strong> ${emp.hired}</p>
                    ${isReg ? (hasEdu ? `<p><strong>Education:</strong> ${emp.tor.school}</p>` : '<p style="color:orange"><em>Education data not found.</em></p>') : ''}
                </div>
                <div style="margin-top:15px; padding:10px; background:#e8f5e9; border-radius:5px; color:#2e7d32; font-size:0.9em;">
                    ✓ Loaded 1 Employee
                </div>
            `;
        } else {
            // Batch Result Display
            display.innerHTML = `
                <h3>2. Batch Selection</h3>
                <div class="employee-card-details">
                    <p><strong>Employees Found:</strong> ${currentEmployeesBatch.length}</p>
                    <ul style="font-size:0.9em; margin-top:10px; padding-left:20px;">
                        ${currentEmployeesBatch.map(e => `<li>${e.name} (${e.id})</li>`).join('')}
                    </ul>
                </div>
                <div style="margin-top:15px; padding:10px; background:#e8f5e9; border-radius:5px; color:#2e7d32; font-size:0.9em;">
                    ✓ Ready to Generate Batch Files
                </div>
            `;
        }
        
        // --- BUTTON STATE LOGIC ---
        const summaryBtn = document.getElementById('generate-summary');
        const panelBtn = document.getElementById('generate-panel-sheet');

        panelBtn.disabled = false; 
        
        // Update button text to indicate batch
        panelBtn.textContent = currentEmployeesBatch.length > 1 ? `Generate Batch Panel Sheets (${currentEmployeesBatch.length})` : "Generate Panel Sheet (.xlsx)";

        // Check if ALL in batch have edu data (strict) or ANY (loose). Let's go with loose but warn.
        const anyEduMissing = currentEmployeesBatch.some(e => !e.tor.school);
        const hasEduFile = !!globalEduData; // Basic check if file exists

        if (isReg) {
            if (hasEduFile) {
                summaryBtn.disabled = false;
                summaryBtn.textContent = currentEmployeesBatch.length > 1 ? `Generate Batch Summaries (${currentEmployeesBatch.length})` : "Generate Summary (.doc)";
                summaryBtn.classList.remove('btn-disabled-custom');
            } else {
                summaryBtn.disabled = true;
                summaryBtn.textContent = "Upload Edu File to Enable";
                summaryBtn.classList.add('btn-disabled-custom');
            }
        } else {
            summaryBtn.disabled = false;
            summaryBtn.textContent = currentEmployeesBatch.length > 1 ? `Generate Batch Summaries (${currentEmployeesBatch.length})` : "Generate Summary (.doc)";
            summaryBtn.classList.remove('btn-disabled-custom');
        }
    });
}

// Helper: Promisified Excel Reader
function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(sheet);
            resolve(jsonData);
        };
        reader.onerror = (error) => reject(error);
        reader.readAsArrayBuffer(file);
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
// 3. DOCUMENT GENERATION (BATCH SUPPORT)
// ====================================================================

document.getElementById('generate-summary')?.addEventListener('click', () => {
    if (!currentEmployeesBatch || currentEmployeesBatch.length === 0) return;
    
    // Loop through batch and generate for each
    currentEmployeesBatch.forEach(emp => {
        if (emp.isReg) {
            generateRegularizationWordDoc(emp);
        } else {
            generatePromotionWordDoc(emp);
        }
    });
});

document.getElementById('generate-panel-sheet')?.addEventListener('click', () => {
    if (!currentEmployeesBatch || currentEmployeesBatch.length === 0) return;

    const wb = XLSX.utils.book_new(); // Create ONE workbook for ALL employees

    currentEmployeesBatch.forEach(emp => {
        // Create a sheet for this employee
        const ws = createPanelWorksheet(emp);
        
        // Name the sheet (ID - Name) to keep unique
        // Excel sheet names max 31 chars
        let sheetName = `${emp.id} - ${emp.name.split(',')[0]}`; 
        if(sheetName.length > 31) sheetName = sheetName.substring(0, 31);
        
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });

    // Download ONE file containing multiple sheets
    const filePrefix = currentEmployeesBatch[0].isReg ? "Regularization" : "Promotion";
    const fileName = currentEmployeesBatch.length > 1 ? `${filePrefix}_Batch_Panel_Sheets.xlsx` : `${filePrefix}_Panel_${currentEmployeesBatch[0].id}.xlsx`;
    
    XLSX.writeFile(wb, fileName);
});


// --- HELPER: CREATE WORKSHEET FOR A SINGLE EMPLOYEE ---
// This contains all the logic previously in the event listener
function createPanelWorksheet(emp) {
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
    
    return ws;
}

// --- REGULARIZATION TEMPLATE ---
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
        <div class="line-1">ABENSON GROUP OF COMPANIES</div>
        <div class="line-2">Human Resource Department</div>
        <div class="space-17">&nbsp;</div>
        <div class="line-4">REGULARIZATION SUMMARY</div>
        <div class="space-20">&nbsp;</div>

        <table class="info-table">
            <tr><td><span class="label-normal">Name:</span> &nbsp;&nbsp; <span class="data-bold">${emp.name}</span></td><td class="right"><span class="label-normal">On-Board Date:</span> &nbsp;&nbsp; <span class="data-bold">${emp.hired}</span></td></tr>
            <tr><td><span class="label-normal">Position:</span> &nbsp;&nbsp; <span class="data-bold">${emp.pos}</span></td><td class="right"><span class="label-normal">Panel Interview:</span> &nbsp;&nbsp; <span class="data-bold">${panelDate}</span></td></tr>
            <tr><td colspan="2"><span class="label-normal">Branch/Department:</span> &nbsp;&nbsp; <span class="data-bold">${emp.branch} - ${emp.dept}</span></td></tr>
            <tr><td colspan="2"><span class="label-normal">Immediate Supervisor:</span> &nbsp;&nbsp; <span class="data-bold">${emp.mentor}</span></td></tr>
        </table>

        <div class="space-10">&nbsp;</div>

        <table class="forms-table" align="center">
            <tr class="forms-header"><td>REGULARIZATION FORMS</td><td>REMARKS</td></tr>
            <tr><td class="col-title">Recommendation & Essay</td><td class="col-remark">See attached files</td></tr>
            <tr><td class="col-title">Individual Score Card</td><td class="col-remark"></td></tr>
            <tr><td class="col-title">Personal Discipline Score [PDS]</td><td class="col-remark"></td></tr>
            <tr><td class="col-title">Internal Customer Rating</td><td class="col-remark"></td></tr>
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
            <div class="space-14">&nbsp;</div>
            <div class="na-text">· N/A</div>
            <div class="space-11">&nbsp;</div>
            <div class="space-11">&nbsp;</div>
            <div class="space-11">&nbsp;</div>
            <div class="space-13">&nbsp;</div>
            <div style="font-weight:bold; font-size:12pt;">Prepared by:</div>
            <div class="space-11">&nbsp;</div>
            <div class="space-11">&nbsp;</div>
            <div style="border-bottom: 1px solid black; width: 200px;"></div>
            <div style="font-weight:bold; margin-top:5px; font-size:12pt;">HR Staff</div>
        </div>
    </body>
    </html>
    `;

    saveDoc(content, `Regularization_Summary_${emp.id}.doc`);
}

// --- PROMOTION TEMPLATE (NEW) ---
function generatePromotionWordDoc(emp) {
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
            @page { size: 8.5in 14in; margin: 1in; mso-page-orientation: portrait; }
            body { font-family: 'Calibri', sans-serif; }
            
            /* Header styles */
            .header-1 { font-size: 14pt; font-weight: bold; text-align: center; margin: 0; }
            .header-2 { font-size: 13pt; text-align: center; margin: 0; }
            .header-title { font-size: 17pt; font-weight: bold; text-align: center; margin: 0; }

            /* Employee Info Table */
            .info-table { width: 100%; border-collapse: collapse; font-size: 12pt; margin-top: 20px; }
            .info-table td { border: 1px solid black; padding: 5px; vertical-align: middle; } /* Borders ON for Promo Summary */
            .info-label { font-weight: normal; width: 2.5in; }
            .info-data { font-weight: bold; }

            /* 360 Feedback Table */
            .feedback-title { font-size: 14pt; font-weight: bold; text-decoration: underline; margin-top: 20px; margin-bottom: 10px; text-align: center; }
            .feedback-table { width: 4in; border-collapse: collapse; font-size: 11pt; margin-left: auto; margin-right: auto; } /* Centered */
            .feedback-table td { border: 1px solid black; padding: 5px; }
            
            /* Signatures */
            .sig-table { width: 100%; border-collapse: collapse; margin-top: 40px; }
            .sig-table td { vertical-align: top; width: 33%; padding-right: 20px; }
            .sig-line { border-bottom: 1px solid black; margin-bottom: 5px; }
            .sig-label { font-weight: bold; margin-bottom: 40px; }

            .spacer { height: 15px; }
            .spacer-lg { height: 30px; }
        </style>
    </head>
    <body>
        <div class="header-1">ABENSON GROUP OF COMPANIES</div>
        <div class="header-2">Human Resource Department</div>
        <div class="spacer"></div>
        <div class="header-title">PROMOTION SUMMARY</div>
        <div class="spacer"></div>

        <!-- Table 1: Employee Info (Borders visible per template) -->
        <table class="info-table">
            <tr><td class="info-label">Employee Name:</td><td class="info-data">${emp.name}</td></tr>
            <tr><td class="info-label">Current Position:</td><td class="info-data">${emp.pos}</td></tr>
            <tr><td class="info-label">Proposed Position:</td><td class="info-data">${emp.targetPos}</td></tr>
            <tr><td class="info-label">On Board Date:</td><td class="info-data">${emp.hired}</td></tr>
            <tr><td class="info-label">Last Promotion Date:</td><td class="info-data">${emp.lastPromoDate}</td></tr>
            <tr><td class="info-label">Branch/Department:</td><td class="info-data">${emp.branch} - ${emp.dept}</td></tr>
            <tr><td class="info-label">Immediate Supervisor:</td><td class="info-data">${emp.mentor}</td></tr>
            <tr><td class="info-label">BSC KPI Scorecard:</td><td class="info-data"></td></tr>
            <tr><td class="info-label">Internal Customer Rating:</td><td class="info-data"></td></tr>
            <tr><td class="info-label">PDS Attendance:</td><td class="info-data"></td></tr>
        </table>

        <!-- Table 2: 360 Feedback -->
        <div class="spacer-lg"></div>
        <div class="feedback-title">360 Feedback Summary</div>
        <table class="feedback-table" align="center">
            <tr>
                <td style="font-weight:bold;">Average Score</td>
                <td style="font-weight:bold; text-align:center;">${emp.feedback.avg}</td>
            </tr>
        </table>

        <!-- Comments -->
        <br><br>
        <div style="font-weight:bold; font-size:12pt;">Comments/Remarks:</div>
        <div style="height: 40px; border-bottom: 1px solid black; margin-bottom: 5px;"></div>
        <div style="height: 30px; border-bottom: 1px solid black;"></div>
        <div style="height: 30px; border-bottom: 1px solid black;"></div>

        <!-- Signatures -->
        <table class="sig-table">
            <tr>
                <td>
                    <div class="sig-label">Prepared by:</div>
                    <div class="sig-line"></div>
                    <div style="font-weight:bold;">HR Staff</div>
                </td>
                <td>
                    <div class="sig-label">Noted by:</div>
                    <div class="sig-line"></div>
                    <div style="font-weight:bold;">Head - HR</div>
                </td>
                <td>
                    <div class="sig-label">Approved by:</div>
                    <div class="sig-line"></div>
                    <div style="font-weight:bold;">President</div>
                </td>
            </tr>
        </table>

    </body>
    </html>
    `;

    saveDoc(content, `Promotion_Summary_${emp.id}.doc`);
}

function saveDoc(content, filename) {
    const blob = new Blob(['\ufeff', content], { type: 'application/msword' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// ====================================================================
// 4. EXCEL PANEL SHEET LOGIC (UNCHANGED)
// ====================================================================
// ... (Keeping the existing Excel Logic) ...

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