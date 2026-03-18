/**
 * SILVER EDGE TECHNOLOGIES (SET) - ENTERPRISE ERP SYSTEM
 * Modules: V8 Payroll Calculation Engine, Cloud PDF Dispatcher, and Financial Logic
 */

function doGet(e) {
  // Yeh line aapke 'Dashboard.html' ko website ke roop mein open kar degi
  return HtmlService.createHtmlOutputFromFile('Dashboard')
      .setTitle('Silver Edge ERP')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SET ERP System')
    .addItem('Process Attendance & Salary', 'processFullPayroll')
    .addItem('Launch Cloud Dispatcher', 'showDispatcher')
    .addSeparator()
    .addItem('Launch Dashboard', 'showEnterpriseDashboard')
    .addToUi();
}

function showDispatcher() {
  const html = HtmlService.createHtmlOutputFromFile('Dispatcher').setWidth(550).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'SET Cloud Dispatcher');
}

function showEnterpriseDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard').setWidth(3000).setHeight(2000).setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'SET HR Simulator');
}

// =========================================================================
// MODULE 1: V8 PAYROLL CALCULATION ENGINE (WITH ADVANCED SHEET UI)
// =========================================================================
function processFullPayroll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt("SET Payroll System", "Enter the target month sheet name (e.g., Jan):", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return; 

  const targetSheetName = response.getResponseText().trim();
  const masterSheet = ss.getSheetByName("ID_DATA"); 
  const rawSheet = ss.getSheetByName(targetSheetName);

  if (!masterSheet || !rawSheet) {
    ui.alert("System Error: Unable to locate 'ID_DATA' or '" + targetSheetName + "' sheet.");
    return;
  }

  const masterData = masterSheet.getDataRange().getValues();
  const rawData = rawSheet.getDataRange().getValues();

  // Initialize Active Employees
  const employees = {};
  for(let i = 1; i < masterData.length; i++) {
    let status = String(masterData[i][0]).trim().toLowerCase();
    
    if(status === "active") {
      let id = String(masterData[i][1]).trim();
      
      if(id && id.toLowerCase() !== "id") {
        let medTier = parseInt(masterData[i][9]) || 0; 
        let medAmount = 0;
        
        if(medTier === 1) medAmount = 600;
        else if(medTier === 2) medAmount = 1200;
        else if(medTier === 3) medAmount = 1800;
        else if(medTier === 4) medAmount = 2400;
        else if(medTier === 5) medAmount = 3000;

        employees[id] = { 
          designation: String(masterData[i][2]).trim(),
          name: String(masterData[i][3]).trim(),
          email: String(masterData[i][4]).trim(),
          bank: String(masterData[i][5]).trim(),
          accTitle: String(masterData[i][6]).trim(),
          accNo: String(masterData[i][7]).trim(),
          salary: parseFloat(masterData[i][8]) || 0,
          medicalDed: medAmount,
          totalLoan: parseFloat(masterData[i][10]) || 0,
          remainingLoan: parseFloat(masterData[i][11]) || 0,
          loanInstallment: parseFloat(masterData[i][12]) || 0,
          totalLate: 0, totalBonus: 0, totalFine: 0, totalPaid: 0, totalUnpaid: 0, totalSandwich: 0
        };
      }
    }
  }

  // Aggregate Biometric Input
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

  // Define Output Matrix (ADDED EXTRA MESSAGE COLUMN HERE)
  const finalOutput = [[
    "ID", "Name", "Designation", "Bank", "Account Title", "Account No", "Email", "Basic Salary", 
    "Late Mins", "Unpaid/Absents", "Sandwich Days", 
    "Late Penalty", "Unpaid Penalty", "Sandwich Penalty", "Fine", 
    "Income Tax", "PF (680)", "Medical Ded", "Total Loan", "Remaining Loan", "Monthly Installment", "Loan Ded", "Bonus", "NET PAYABLE", "Extra Message"
  ]];

  // Execute Financial Logic
  for(let id in employees) {
    let emp = employees[id];
    let sal = emp.salary;
    
    let perDay = sal / 30;
    let perMin = sal / 16200;
    
    let latePen = emp.totalLate * perMin;
    let unpaidPen = emp.totalUnpaid * perDay;
    let sandwichPen = emp.totalSandwich * perDay;
    let tax = getTax(sal);
    let pfFixed = 680; 
    
    let actualLoanDeduction = 0;
    if(emp.remainingLoan > 0) {
      actualLoanDeduction = emp.loanInstallment;
      if(actualLoanDeduction > emp.remainingLoan) actualLoanDeduction = emp.remainingLoan;
    }
    
    let totalDeductions = latePen + unpaidPen + sandwichPen + emp.totalFine + tax + pfFixed + emp.medicalDed + actualLoanDeduction;
    let netPay = (sal + emp.totalBonus) - totalDeductions;

    // Blank string added at the end for Extra Message manual entry
    finalOutput.push([
      id, emp.name, emp.designation, emp.bank, emp.accTitle, emp.accNo, emp.email, sal, 
      emp.totalLate, emp.totalUnpaid, emp.totalSandwich, 
      Math.round(latePen), Math.round(unpaidPen), Math.round(sandwichPen), emp.totalFine, 
      tax, pfFixed, emp.medicalDed, emp.totalLoan, emp.remainingLoan, emp.loanInstallment, actualLoanDeduction, emp.totalBonus, Math.round(netPay), ""
    ]);
  }

  // =========================================================================
  // ADVANCED SHEET RENDERING & UI FORMATTING
  // =========================================================================
  let lastRow = Math.max(rawSheet.getLastRow(), 1);
  let numRows = finalOutput.length;
  let numCols = finalOutput[0].length;
  
  // Clear old data and formats
  rawSheet.getRange(1, 15, lastRow, numCols).clearContent().clearFormat(); 
  
  // Set new calculated data
  let targetRange = rawSheet.getRange(1, 15, numRows, numCols);
  targetRange.setValues(finalOutput);
  
  // 1. Format Header (Corporate Blue)
  let headerRange = rawSheet.getRange(1, 15, 1, numCols);
  headerRange.setBackground("#1e3a8a")
             .setFontColor("white")
             .setFontWeight("bold")
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle");
             
  // Freeze Header
  rawSheet.setFrozenRows(1);
  
  // 2. Add Professional Borders
  targetRange.setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  
  if (numRows > 1) {
    let dataRange = rawSheet.getRange(2, 15, numRows - 1, numCols);
    
    // Set default alignment
    dataRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
    
    // Left align Text columns (Name, Designation, Email, etc.)
    rawSheet.getRange(2, 16, numRows - 1, 2).setHorizontalAlignment("left"); 
    rawSheet.getRange(2, 18, numRows - 1, 2).setHorizontalAlignment("left"); 
    rawSheet.getRange(2, 21, numRows - 1, 1).setHorizontalAlignment("left"); 
    
    // 3. Create Zebra Striping (Alternating Row Colors)
    let backgrounds = [];
    for (let r = 0; r < numRows - 1; r++) {
      let rowColor = (r % 2 === 0) ? "#f8fafc" : "#ffffff"; 
      let rowBackgrounds = new Array(numCols).fill(rowColor);
      
      // Highlight Net Payable Column with Emerald Green Background
      rowBackgrounds[numCols - 2] = "#dcfce7"; 
      backgrounds.push(rowBackgrounds);
    }
    dataRange.setBackgrounds(backgrounds);
    
    // 4. Highlight Penalties & Deductions Text in Red
    rawSheet.getRange(2, 26, numRows - 1, 7).setFontColor("#ef4444"); // Red Text
    rawSheet.getRange(2, 36, numRows - 1, 1).setFontColor("#ef4444"); // Red Text for Loan Ded
    
    // 5. Highlight Net Payable Text in Dark Green & Bold
    let netPayRange = rawSheet.getRange(2, 38, numRows - 1, 1);
    netPayRange.setFontColor("#166534").setFontWeight("bold");

    // NEW: Highlight Extra Message Column
    rawSheet.getRange(2, 39, numRows - 1, 1).setBackground("#fff8f1").setFontColor("#92400e");
    
    // 6. Number Formatting (Apply commas)
    rawSheet.getRange(2, 22, numRows - 1, 17).setNumberFormat("#,##0");
  }
  
  rawSheet.autoResizeColumns(15, numCols);
  ui.alert("System Notification: Advanced payroll processing completed! Extra Message column generated.");
}

function getTax(sal) {
  if (sal <= 50000) return 0; if (sal <= 60000) return 100; if (sal <= 70000) return 200;
  if (sal <= 80000) return 300; if (sal <= 90000) return 400; if (sal <= 100000) return 500;
  if (sal <= 110000) return 1600; if (sal <= 120000) return 2700; if (sal <= 130000) return 3800;
  if (sal <= 140000) return 4900; if (sal <= 150000) return 6000; if (sal <= 160000) return 7100;
  if (sal <= 170000) return 8200; if (sal <= 180000) return 9300; if (sal <= 190000) return 11200;
  return 13500;
}

// =========================================================================
// MODULE 2: SECURE BATCH DISPATCHER & PDF GENERATOR
// =========================================================================
function getEmployeeDataForBatching(monthName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(monthName);
  if(!rawSheet) throw new Error("System Error: Sheet '" + monthName + "' not found.");

  const data = rawSheet.getDataRange().getValues();
  let employeesToProcess = [];

  for(let i = 1; i < data.length; i++) {
    let id = data[i][14]; 
    let email = data[i][20]; 
    
    if(id && String(id).toLowerCase() !== "id" && email && String(email).includes("@")) {
      let safeRow = [];
      for(let j = 0; j < data[i].length; j++) {
        let cellValue = data[i][j];
        if (cellValue instanceof Date) { safeRow.push(cellValue.toString()); } 
        else { safeRow.push(cellValue); }
      }
      employeesToProcess.push(safeRow);
    }
  }
  return employeesToProcess;
}

function processSingleBatch(monthName, batchData) {
  const mainFolder = DriveApp.getFoldersByName("Pay slip details").hasNext() ? DriveApp.getFoldersByName("Pay slip details").next() : DriveApp.createFolder("Pay slip details");
  const monthFolder = mainFolder.getFoldersByName(monthName).hasNext() ? mainFolder.getFoldersByName(monthName).next() : mainFolder.createFolder(monthName);

  // ==========================================
  // SMART LOGO FETCHER (Directly from Google Drive)
  // ==========================================
  let myLogoBase64 = "";
  try {
    let logoFile = DriveApp.getFileById("1WmvI69t5UfsKfaRx-_X8UPZsyqRej8md"); 
    let blob = logoFile.getBlob();
    myLogoBase64 = "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
  } catch(e) {
    Logger.log("Logo Load Error: " + e);
  }

  let successCount = 0;

  for(let i = 0; i < batchData.length; i++) {
    let row = batchData[i];
    
    let id = row[14]; let name = row[15]; let designation = row[16]; let bank = row[17]; 
    let accTitle = row[18]; let accNo = row[19]; let email = row[20]; let basic = row[21]; 
    let lateMin = row[22]; let unpaidL = row[23]; let sandwich = row[24];
    let latePen = row[25]; let unpaidPen = row[26]; let sandwichPen = row[27]; let fine = row[28];
    let tax = row[29]; let pfFixed = row[30]; let medDed = row[31]; 
    let totalLoan = row[32]; let remainingLoan = row[33]; let loanDed = row[35]; 
    let bonus = row[36]; let netPay = row[37];
    let extraMsg = row[38]; // NEW: EXTRA MESSAGE EXTRACTED HERE

    let totalDed = latePen + unpaidPen + sandwichPen + fine + tax + pfFixed + medDed + loanDed;
    let grossSalary = basic + bonus;

    // ==========================================
    // PDF GENERATION: DETAILED CORPORATE BLUE
    // ==========================================
    let htmlBody = `
      <div style="font-family: 'Segoe UI', Helvetica, Arial, sans-serif; width: 750px; margin: 0 auto; background-color: transparent; position: relative; min-height: 1050px; box-sizing: border-box; border: 1px solid #cbd5e1;">

          <div style="height: 12px; background-color: transparent; width: 100%; border-bottom: 4px solid #1e3a8a;"></div>

          <div style="position: absolute; top: 0; left: 0; right: 0; bottom: 0; display: table; width: 100%; height: 100%; z-index: 0; pointer-events: none;">
              <div style="display: table-cell; vertical-align: middle; text-align: center;">
                  <img src="${myLogoBase64}" width="550" style="opacity: 0.08; filter: alpha(opacity=8);">
              </div>
          </div>

          <div style="position: relative; z-index: 1; padding: 40px;">
              
              <table style="width: 100%; margin-bottom: 40px;">
                  <tr>
                      <td style="width: 50%; vertical-align: middle;">
                          <img src="${myLogoBase64}" alt="Logo" style="max-height: 110px; max-width: 280px;">
                      </td>
                      <td style="width: 50%; text-align: right; vertical-align: middle;">
                          <div style="font-size: 28px; font-weight: 900; color: #1e3a8a; letter-spacing: 1px; text-transform: uppercase;">SALARY STATEMENT</div>
                          <div style="font-size: 15px; font-weight: bold; color: #1e3a8a; margin-top: 5px; text-transform: uppercase; letter-spacing: 1px;">
                              PAYROLL MONTH: ${monthName} 2026
                          </div>
                      </td>
                  </tr>
              </table>

              <table style="width: 100%; margin-bottom: 35px; border-spacing: 0;">
                  <tr>
                      <td style="width: 48%; vertical-align: top; background-color: transparent;">
                          <div style="color: #1e3a8a; padding: 8px 0px; font-size: 14px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; border-bottom: 2px solid #1e3a8a; margin-bottom: 10px;">Company Information</div>
                          <div style="padding: 5px 0px; background-color: transparent; height: 70px;">
                              <div style="font-size: 14px; font-weight: bold; color: #000000; margin-bottom: 4px;">Silver Edge Technologies</div>
                              <div style="font-size: 12px; color: #000000; line-height: 1.6;">DHA Rahbar Phase 1, Lahore<br>Email: hr@silveredge.co</div>
                          </div>
                      </td>
                      <td style="width: 4%;"></td> <td style="width: 48%; vertical-align: top; background-color: transparent;">
                          <div style="color: #1e3a8a; padding: 8px 0px; font-size: 14px; font-weight: bold; text-transform: uppercase; letter-spacing: 1px; border-bottom: 2px solid #1e3a8a; margin-bottom: 10px;">Employee Information</div>
                          <div style="padding: 5px 0px; background-color: transparent; height: 70px;">
                              <table style="width: 100%; font-size: 12px; line-height: 1.6;">
                                  <tr><td style="color: #000000; width: 35%;">Name:</td><td style="font-weight: bold; color: #000000;">${name}</td></tr>
                                  <tr><td style="color: #000000;">Emp ID & Role:</td><td style="font-weight: bold; color: #000000;">${id} - ${designation || '-'}</td></tr>
                                  <tr><td style="color: #000000;">Bank A/C:</td><td style="font-weight: bold; color: #000000;">${bank || '-'} | ${accNo || '-'}</td></tr>
                              </table>
                          </div>
                      </td>
                  </tr>
              </table>

              <table style="width: 100%; border-collapse: collapse; margin-bottom: 25px; background-color: transparent;">
                  <tr>
                      <td style="width: 48%; vertical-align: top; padding: 0; border: 1px solid #1e3a8a; background-color: transparent;">
                          <table style="width: 100%; border-collapse: collapse; background-color: transparent;">
                              <thead>
                                  <tr>
                                      <th colspan="2" style="color: #1e3a8a; text-align: left; padding: 10px 15px; font-size: 14px; text-transform: uppercase; letter-spacing: 1px; border-bottom: 2px solid #1e3a8a; font-weight: bold;">Earnings Description</th>
                                  </tr>
                              </thead>
                              <tbody>
                                  <tr style="background-color: transparent;">
                                      <td style="padding: 12px 15px; border-bottom: 1px solid #cbd5e1; font-size: 13px; color: #000000;">Basic Salary</td>
                                      <td style="padding: 12px 15px; border-bottom: 1px solid #cbd5e1; font-size: 13px; font-weight: bold; text-align: right; color: #000000;">${Number(basic).toLocaleString()}</td>
                                  </tr>
                                  <tr style="background-color: transparent;">
                                      <td style="padding: 12px 15px; border-bottom: 1px solid #cbd5e1; font-size: 13px; color: #000000;">Target Bonus</td>
                                      <td style="padding: 12px 15px; border-bottom: 1px solid #cbd5e1; font-size: 13px; font-weight: bold; text-align: right; color: #000000;">${Number(bonus).toLocaleString()}</td>
                                  </tr>
                                  
                                  <tr><td colspan="2" style="height: 148px; border-bottom: 1px solid #1e3a8a; background-color: transparent;"></td></tr>
                                  
                                  <tr style="background-color: transparent;">
                                      <td style="padding: 14px 15px; font-size: 13px; font-weight: bold; color: #1e3a8a;">Total Gross Earnings</td>
                                      <td style="padding: 14px 15px; font-size: 14px; font-weight: 900; text-align: right; color: #1e3a8a;">${Number(grossSalary).toLocaleString()}</td>
                                  </tr>
                              </tbody>
                          </table>
                      </td>
                      <td style="width: 4%; background-color: transparent;"></td> <td style="width: 48%; vertical-align: top; padding: 0; border: 1px solid #1e3a8a; background-color: transparent;">
                          <table style="width: 100%; border-collapse: collapse; background-color: transparent;">
                              <thead>
                                  <tr>
                                      <th colspan="2" style="color: #1e3a8a; text-align: left; padding: 10px 15px; font-size: 14px; text-transform: uppercase; letter-spacing: 1px; border-bottom: 2px solid #1e3a8a; font-weight: bold;">Deductions Description</th>
                                  </tr>
                              </thead>
                              <tbody>
                                  <tr style="background-color: transparent;">
                                      <td style="padding: 10px 15px; border-bottom: 1px solid #cbd5e1; font-size: 13px; color: #000000;">Late Penalties</td>
                                      <td style="padding: 10px 15px; border-bottom: 1px solid #cbd5e1; font-size: 13px; font-weight: bold; text-align: right; color: #000000;">${Number(latePen).toLocaleString()}</td>
                                  </tr>
                                  <tr style="background-color: transparent;">
                                      <td style="padding: 10px 15px; border-bottom: 1px solid #cbd5e1; font-size: 13px; color: #000000;">Absents/Sandwich</td>
                                      <td style="padding: 10px 15px; border-bottom: 1px solid #cbd5e1; font-size: 13px; font-weight: bold; text-align: right; color: #000000;">${Number(unpaidPen + sandwichPen).toLocaleString()}</td>
                                  </tr>
                                  <tr style="background-color: transparent;">
                                      <td style="padding: 10px 15px; border-bottom: 1px solid #cbd5e1; font-size: 13px; color: #000000;">Income Tax</td>
                                      <td style="padding: 10px 15px; border-bottom: 1px solid #cbd5e1; font-size: 13px; font-weight: bold; text-align: right; color: #000000;">${Number(tax).toLocaleString()}</td>
                                  </tr>
                                  <tr style="background-color: transparent;">
                                      <td style="padding: 10px 15px; border-bottom: 1px solid #cbd5e1; font-size: 13px; color: #000000;">Fixed PF & Medical</td>
                                      <td style="padding: 10px 15px; border-bottom: 1px solid #cbd5e1; font-size: 13px; font-weight: bold; text-align: right; color: #000000;">${Number(pfFixed + medDed).toLocaleString()}</td>
                                  </tr>
                                  <tr style="background-color: transparent;">
                                      <td style="padding: 10px 15px; border-bottom: 1px solid #cbd5e1; font-size: 13px; color: #000000;">Loan Installment</td>
                                      <td style="padding: 10px 15px; border-bottom: 1px solid #cbd5e1; font-size: 13px; font-weight: bold; text-align: right; color: #000000;">${Number(loanDed).toLocaleString()}</td>
                                  </tr>
                                  <tr style="background-color: transparent;">
                                      <td style="padding: 10px 15px; border-bottom: 1px solid #1e3a8a; font-size: 13px; color: #000000;">Disciplinary Fine</td>
                                      <td style="padding: 10px 15px; border-bottom: 1px solid #1e3a8a; font-size: 13px; font-weight: bold; text-align: right; color: #000000;">${Number(fine).toLocaleString()}</td>
                                  </tr>
                                  <tr style="background-color: transparent;">
                                      <td style="padding: 14px 15px; font-size: 13px; font-weight: bold; color: #1e3a8a;">Total Deductions</td>
                                      <td style="padding: 14px 15px; font-size: 14px; font-weight: 900; text-align: right; color: #1e3a8a;">${Number(totalDed).toLocaleString()}</td>
                                  </tr>
                              </tbody>
                          </table>
                      </td>
                  </tr>
              </table>

              <table style="width: 100%; margin-bottom: 20px; border-collapse: collapse; background-color: transparent;">
                  <tr>
                      <td style="text-align: right; font-size: 13px; color: #000000; padding: 5px 0;">
                          Outstanding Loan Balance: <strong style="color: #1e3a8a; border-bottom: 1px solid #1e3a8a;">PKR ${Number(remainingLoan - loanDed).toLocaleString()}</strong>
                      </td>
                  </tr>
              </table>

              <table style="width: 100%; background-color: transparent; border: 2px solid #1e3a8a; border-left: 8px solid #1e3a8a; border-radius: 4px; margin-bottom: 40px; border-collapse: collapse;">
                  <tr>
                      <td style="padding: 20px 25px; width: 60%; background-color: transparent;">
                          <span style="font-size: 14px; font-weight: 900; text-transform: uppercase; color: #1e3a8a; letter-spacing: 1px;">Net Transferable Salary</span><br>
                          <span style="font-size: 12px; color: #000000; font-style: italic;">Final amount to be credited to the employee's account.</span>
                      </td>
                      <td style="padding: 20px 25px; text-align: right; vertical-align: middle; background-color: transparent;">
                          <span style="font-size: 32px; font-weight: 900; color: #1e3a8a; letter-spacing: 1px;">PKR ${Number(netPay).toLocaleString()}</span>
                      </td>
                  </tr>
              </table>

              <table style="width: 100%; border-top: 1px solid #1e3a8a; padding-top: 20px; background-color: transparent;">
                  <tr>
                      <td style="width: 60%; vertical-align: bottom;">
                          <div style="font-size: 11px; color: #000000; line-height: 1.6;">
                              <strong style="color: #1e3a8a;">IMPORTANT NOTE:</strong> This is a secure, system-generated document.<br>
                              It does not require a physical signature for authentication.<br>
                              Generated securely via Silver Edge Technologies ERP.
                          </div>
                      </td>
                      <td style="width: 40%; text-align: right; vertical-align: bottom;">
                          <div style="border-top: 2px solid #1e3a8a; padding-top: 8px; width: 220px; float: right;">
                              <strong style="font-size: 12px; text-transform: uppercase; color: #1e3a8a; letter-spacing: 1px;">Authorized Signatory</strong>
                          </div>
                      </td>
                  </tr>
              </table>

          </div> </div>
    `;

    let blob = Utilities.newBlob(htmlBody, MimeType.HTML).getAs(MimeType.PDF).setName(name + "_" + monthName + "_SalarySlip.pdf");
    monthFolder.createFile(blob);

    // NEW: EXTRA MESSAGE HTML FOR EMAIL INJECTION
    let extraMessageHtml = "";
    if (extraMsg && String(extraMsg).trim() !== "") {
      extraMessageHtml = `
        <div style="background-color: #fffbeb; border: 1px solid #fde68a; border-left: 5px solid #f59e0b; padding: 18px; border-radius: 6px; margin: 25px 0;">
          <h3 style="color: #b45309; margin-top: 0; font-size: 14px; text-transform: uppercase; letter-spacing: 1px;">Important Note from HR</h3>
          <p style="color: #78350f; font-size: 14px; margin-bottom: 0; line-height: 1.6;">${extraMsg}</p>
        </div>
      `;
    }

    // ==========================================
    // PROFESSIONAL EMAIL TEMPLATE INTEGRATION
    // ==========================================
    let emailSubject = `Confidential: Salary Slip for ${monthName} 2026 - Silver Edge Technologies`;

    let emailBody = `
      <div style="font-family: 'Segoe UI', Arial, sans-serif; color: #334155; line-height: 1.6; max-width: 600px; margin: 0 auto; padding: 30px; border: 1px solid #bfdbfe; border-radius: 12px; background-color: #ffffff;">
        <h2 style="color: #1e3a8a; margin-top: 0; border-bottom: 2px solid #bfdbfe; padding-bottom: 15px;">Silver Edge Technologies</h2>
        <p>Dear <strong>${name}</strong>,</p>
        <p>We hope this email finds you well.</p>
        <p>Please find attached your official salary slip for the payroll month of <strong style="color: #1e3a8a;">${monthName} 2026</strong>.</p>
        
        ${extraMessageHtml} <p style="background-color: #eff6ff; padding: 15px; border-left: 4px solid #1e3a8a; border-radius: 4px; font-size: 13px; color: #475569;">
          <strong>Security Note:</strong> This is a confidential, system-generated document. We kindly request you to review the details. If you have any questions or require further clarification regarding your calculations, please reach out to the HR department within the next 3 working days.
        </p>
        <p>Thank you for your continued dedication, hard work, and contribution to the team.</p>
        <br>
        <p style="margin-bottom: 0; color: #64748b;">Best Regards,</p>
        <p style="font-weight: 800; color: #1e3a8a; margin-top: 5px;">
          Human Resources Department<br>
          <span style="font-weight: 500; font-size: 14px; color: #475569;">Silver Edge Technologies (Pvt) Ltd.</span>
        </p>
      </div>
    `;

    MailApp.sendEmail({ 
      to: email, 
      subject: emailSubject, 
      htmlBody: emailBody, 
      attachments: [blob] 
    });
    
    successCount++;
    Utilities.sleep(3000); // Prevents hitting Gmail rate limits
  }
  return successCount;
}
