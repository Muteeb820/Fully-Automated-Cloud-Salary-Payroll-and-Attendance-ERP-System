/**
 * SILVER EDGE TECHNOLOGIES - ERP SYSTEM
 * Final Production Build: Aggregation, Batching (No-Timeout), Date Bug Fix, and Editable PDF
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Silver Edge ERP')
    .addItem('1️⃣ Process Attendance & Salary', 'processFullPayroll')
    .addItem('2️⃣ Launch Cloud Dispatcher (Emails)', 'showDispatcher')
    .addSeparator()
    .addItem('💻 Launch Live Dashboard', 'showEnterpriseDashboard')
    .addToUi();
}

function showDispatcher() {
  const html = HtmlService.createHtmlOutputFromFile('Dispatcher')
      .setWidth(450).setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, '📨 Bulk Email Dispatcher');
}

function showEnterpriseDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
      .setWidth(1150).setHeight(850)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, '💻 ERP Simulator');
}

// ==========================================
// MODULE 1: SALARY CALCULATION ENGINE
// ==========================================
function processFullPayroll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt("Silver Edge Payroll", "Kis mahine ki sheet calculate karni hai? (e.g., Jan):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return; 

  const targetSheetName = response.getResponseText().trim();
  const masterSheet = ss.getSheetByName("ID_DATA"); 
  const rawSheet = ss.getSheetByName(targetSheetName);

  if (!masterSheet || !rawSheet) {
    ui.alert("❌ Error: 'ID_DATA' ya '" + targetSheetName + "' sheet nahi mili.");
    return;
  }

  const masterData = masterSheet.getDataRange().getValues();
  const rawData = rawSheet.getDataRange().getValues();

  const employees = {};
  for(let i = 1; i < masterData.length; i++) {
    let id = String(masterData[i][0]).trim();
    if(id && id.toLowerCase() !== "id") {
      employees[id] = { 
        name: String(masterData[i][1]).trim(), 
        email: String(masterData[i][2]).trim(),
        salary: parseFloat(masterData[i][3]) || 0,
        totalLate: 0, totalBonus: 0, totalFine: 0, totalPaid: 0, totalUnpaid: 0, totalSandwich: 0
      };
    }
  }

  for(let i = 1; i < rawData.length; i++) {
    let row = rawData[i];
    let id = String(row[1]).trim(); 
    if(!employees[id]) continue; 

    employees[id].totalLate += parseFloat(row[4] || 0);     
    employees[id].totalBonus += parseFloat(row[5] || 0);    
    employees[id].totalFine += parseFloat(row[6] || 0);     
    employees[id].totalPaid += parseFloat(row[7] || 0);     
    employees[id].totalUnpaid += parseFloat(row[8] || 0);   
    employees[id].totalSandwich += parseFloat(row[9] || 0); 
  }

  const finalOutput = [[
    "ID", "Name", "Email", "Basic Salary", "Total Late Mins", "Paid Leaves", 
    "Unpaid/Absents", "Sandwich Days", "Late Penalty", "Unpaid Penalty", 
    "Sandwich Penalty", "Fine (Jarmana)", "Income Tax", "Fixed (PF+Med)", "Bonus", "NET PAYABLE"
  ]];

  for(let id in employees) {
    let emp = employees[id];
    let sal = emp.salary;
    
    let perDay = sal / 30;
    let perMin = sal / 16200;
    
    let latePen = emp.totalLate * perMin;
    let unpaidPen = emp.totalUnpaid * perDay;
    let sandwichPen = emp.totalSandwich * perDay;
    let tax = getTax(sal);
    let fixed = 1280; 
    
    let totalDeductions = latePen + unpaidPen + sandwichPen + emp.totalFine + tax + fixed;
    let netPay = (sal + emp.totalBonus) - totalDeductions;

    finalOutput.push([
      id, emp.name, emp.email, sal, emp.totalLate, emp.totalPaid, 
      emp.totalUnpaid, emp.totalSandwich, 
      Math.round(latePen), Math.round(unpaidPen), Math.round(sandwichPen), 
      emp.totalFine, tax, fixed, emp.totalBonus, Math.round(netPay)
    ]);
  }

  let lastRow = Math.max(rawSheet.getLastRow(), 1);
  rawSheet.getRange(1, 13, lastRow, 16).clearContent().clearFormat(); 
  
  rawSheet.getRange(1, 13, finalOutput.length, finalOutput[0].length).setValues(finalOutput);
  rawSheet.getRange(1, 13, 1, 16).setBackground("#0f172a").setFontColor("white").setFontWeight("bold");
  rawSheet.autoResizeColumns(13, 16);
  
  ui.alert("✅ Process Complete! Ab menu se Option 2 dabayen PDF emails ke liye.");
}

function getTax(sal) {
  if (sal <= 50000) return 0; if (sal <= 60000) return 100; if (sal <= 70000) return 200;
  if (sal <= 80000) return 300; if (sal <= 90000) return 400; if (sal <= 100000) return 500;
  if (sal <= 110000) return 1600; if (sal <= 120000) return 2700; if (sal <= 130000) return 3800;
  if (sal <= 140000) return 4900; if (sal <= 150000) return 6000; if (sal <= 160000) return 7100;
  if (sal <= 170000) return 8200; if (sal <= 180000) return 9300; if (sal <= 190000) return 11200;
  return 13500;
}

// ==========================================
// MODULE 2: BATCH EMAIL DISPATCHER
// ==========================================
function getEmployeeDataForBatching(monthName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(monthName);
  if(!rawSheet) throw new Error("Sheet '" + monthName + "' nahi mili!");

  const data = rawSheet.getDataRange().getValues();
  let employeesToProcess = [];

  for(let i = 1; i < data.length; i++) {
    let id = data[i][12]; 
    let email = data[i][14]; 
    
    // Check if ID and Email exist
    if(id && String(id).toLowerCase() !== "id" && email && String(email).includes("@")) {
      
      // 🛠️ BUG FIX: Google Frontend strictly blocks Date objects. 
      // Hum Dates ko simple text mein convert kar rahe hain taake 'null' error na aaye.
      let safeRow = [];
      for(let j = 0; j < data[i].length; j++) {
        let cellValue = data[i][j];
        if (cellValue instanceof Date) {
          safeRow.push(cellValue.toString()); // Convert Date to text
        } else {
          safeRow.push(cellValue);
        }
      }
      employeesToProcess.push(safeRow);
    }
  }
  return employeesToProcess;
}

function processSingleBatch(monthName, batchData) {
  const mainFolder = DriveApp.getFoldersByName("Pay slip details").hasNext() ? DriveApp.getFoldersByName("Pay slip details").next() : DriveApp.createFolder("Pay slip details");
  const monthFolder = mainFolder.getFoldersByName(monthName).hasNext() ? mainFolder.getFoldersByName(monthName).next() : mainFolder.createFolder(monthName);

  let successCount = 0;

  for(let i = 0; i < batchData.length; i++) {
    let row = batchData[i];
    let id = row[12]; let name = row[13]; let email = row[14]; let basic = row[15];
    let lateMin = row[16]; let paidL = row[17]; let unpaidL = row[18];
    let sandwich = row[19]; let latePen = row[20]; let unpaidPen = row[21];
    let sandwichPen = row[22]; let fine = row[23]; let tax = row[24];
    let fixed = row[25]; let bonus = row[26]; let netPay = row[27];

    // 🎨 PDF DESIGN TEMPLATE
    let htmlBody = `
      <div style="font-family: 'Helvetica Neue', Arial, sans-serif; border: 2px solid #e2e8f0; border-radius: 12px; padding: 30px; max-width: 700px; margin: auto; background-color: #ffffff;">
        
        <div style="text-align:center; border-bottom: 3px solid #0f172a; padding-bottom: 15px; margin-bottom: 25px;">
            <h1 style="color:#0f172a; margin: 0; font-size: 26px; font-weight: 800; letter-spacing: 1px;">SILVER EDGE TECHNOLOGIES</h1>
            <p style="color:#64748b; margin: 5px 0 0; font-size: 14px; text-transform: uppercase;">Official Salary Slip - ${monthName}</p>
        </div>
        
        <table style="width:100%; margin-bottom: 25px; font-size: 15px;">
            <tr>
                <td style="color:#475569;">Employee Name: <strong style="color:#0f172a;">${name}</strong></td>
                <td style="text-align:right; color:#475569;">Employee ID: <strong style="color:#0f172a;">${id}</strong></td>
            </tr>
        </table>
        
        <table style="width:100%; border-collapse: collapse; margin-bottom: 20px;">
          <thead>
            <tr style="background:#f1f5f9; text-transform: uppercase; font-size: 12px; color: #475569;">
                <th style="padding: 12px; text-align:left; border: 1px solid #e2e8f0;">Earnings & Additions</th>
                <th style="padding: 12px; text-align:right; border: 1px solid #e2e8f0;">Amount (PKR)</th>
            </tr>
          </thead>
          <tbody>
            <tr>
                <td style="padding: 12px; border: 1px solid #e2e8f0; color:#334155;">Basic Salary</td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; text-align:right; font-family: monospace; font-size: 14px; font-weight: bold;">${Number(basic).toLocaleString()}</td>
            </tr>
            <tr style="background:#f0fdf4;">
                <td style="padding: 12px; border: 1px solid #e2e8f0; color:#16a34a; font-weight: bold;">Target Bonus</td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; text-align:right; color:#16a34a; font-family: monospace; font-size: 14px; font-weight: bold;">+ ${Number(bonus).toLocaleString()}</td>
            </tr>
          </tbody>
        </table>

        <table style="width:100%; border-collapse: collapse; margin-bottom: 25px;">
          <thead>
            <tr style="background:#f1f5f9; text-transform: uppercase; font-size: 12px; color: #475569;">
                <th style="padding: 12px; text-align:left; border: 1px solid #e2e8f0;">Deductions & Penalties</th>
                <th style="padding: 12px; text-align:right; border: 1px solid #e2e8f0;">Amount (PKR)</th>
            </tr>
          </thead>
          <tbody>
            <tr>
                <td style="padding: 12px; border: 1px solid #e2e8f0; color:#334155;">Late Penalty (${lateMin} mins)</td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; text-align:right; color:#dc2626; font-family: monospace; font-size: 14px; font-weight: bold;">- ${Number(latePen).toLocaleString()}</td>
            </tr>
            <tr>
                <td style="padding: 12px; border: 1px solid #e2e8f0; color:#334155;">Absent Penalty (${unpaidL} days)</td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; text-align:right; color:#dc2626; font-family: monospace; font-size: 14px; font-weight: bold;">- ${Number(unpaidPen).toLocaleString()}</td>
            </tr>
            <tr>
                <td style="padding: 12px; border: 1px solid #e2e8f0; color:#334155;">Sandwich Penalty (${sandwich} days)</td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; text-align:right; color:#dc2626; font-family: monospace; font-size: 14px; font-weight: bold;">- ${Number(sandwichPen).toLocaleString()}</td>
            </tr>
            <tr>
                <td style="padding: 12px; border: 1px solid #e2e8f0; color:#334155;">Disciplinary Fines</td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; text-align:right; color:#dc2626; font-family: monospace; font-size: 14px; font-weight: bold;">- ${Number(fine).toLocaleString()}</td>
            </tr>
            <tr>
                <td style="padding: 12px; border: 1px solid #e2e8f0; color:#334155;">Income Tax (FBR)</td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; text-align:right; color:#dc2626; font-family: monospace; font-size: 14px; font-weight: bold;">- ${Number(tax).toLocaleString()}</td>
            </tr>
            <tr>
                <td style="padding: 12px; border: 1px solid #e2e8f0; color:#334155;">Fixed Deductions (PF+Med)</td>
                <td style="padding: 12px; border: 1px solid #e2e8f0; text-align:right; color:#dc2626; font-family: monospace; font-size: 14px; font-weight: bold;">- ${Number(fixed).toLocaleString()}</td>
            </tr>
          </tbody>
        </table>

        <div style="background-color: #0f172a; border-radius: 8px; padding: 20px; color: white; display: flex; justify-content: space-between; align-items: center;">
            <span style="font-size: 16px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px;">Final Net Payable</span>
            <strong style="font-size: 26px; color: #10b981;">PKR ${Number(netPay).toLocaleString()}</strong>
        </div>
        
        <p style="text-align:center; color:#94a3b8; font-size: 11px; margin-top: 25px;">This is a computer-generated document by Silver Edge ERP. No signature is required.</p>
      </div>
    `;

    let blob = Utilities.newBlob(htmlBody, MimeType.HTML).getAs(MimeType.PDF).setName(name + "_" + monthName + "_SilverEdge.pdf");
    monthFolder.createFile(blob);

    MailApp.sendEmail({ to: email, subject: "Salary Slip - Silver Edge Technologies", htmlBody: "Attached is your official salary slip.", attachments: [blob] });
    successCount++;
    Utilities.sleep(3000); 
  }
  return successCount;
}