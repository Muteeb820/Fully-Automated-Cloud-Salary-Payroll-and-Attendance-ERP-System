# 🚀 Company - Enterprise ERP & Payroll System

A complete, serverless Enterprise Resource Planning (ERP) and Payroll automation system built natively on Google Workspace (Google Sheets + Google Apps Script). This system features an advanced calculation engine, secure PDF payslip generation, anti-spam email dispatcher, and an interactive financial simulator dashboard.

## ✨ Key Features

### 1. 🧮 V8 Payroll Calculation Engine
* **Dynamic Base Rates:** Auto-calculates Per-Day (Basic/30) and Per-Minute (Basic/16,200) rates.
* **Penalty Engine:** Accurately deducts Late Minutes, Absents, and applies the Strict Sandwich Rule.
* **Smart Loan Deduction:** Intelligently compares the monthly installment with the remaining loan balance to prevent negative figures.
* **Tax & Tiers:** Automatically applies FBR Income Tax slabs, Fixed Provident Fund (680 PKR), and 5-Tier Medical Deductions.
* **Extra Message Injection:** Auto-generates an "Extra Message" column. HR can write custom notes here that will automatically appear as a highlighted alert box in the employee's email.

### 2. 📄 Secure Cloud Dispatcher
* **Premium PDF Generation:** Creates detailed, transparent-style, corporate blue salary slips natively in the cloud.
* **Drive Integration:** Automatically fetches the company logo from Google Drive and saves generated PDFs in neatly organized month-wise folders.
* **Smart Emailing:** Dispatches highly professional HTML emails with injected custom HR messages and attached PDF slips.
* **Anti-Spam Batching:** Processes emails in safe batches (default: 15) with programmed delays to prevent Google Workspace account suspension.

### 3. 📊 Interactive Employee Dashboard & Simulator
* **Real-time Simulator:** Employees/Admins can input parameters to see live Net Pay calculations.
* **Visual Analytics:** Integrates Chart.js for beautiful "Take Home vs Deducted" doughnut charts.
* **Live Document Preview:** Shows exactly how the generated payslip will look.
* **Logic Transparency:** A dedicated "How We Calculate" modal explaining all mathematical formulas used.
* **Roadmap Module:** A "🚀 Coming Soon" tab showcasing upcoming features (UI/UX updates, Company Letters, etc.).

### 4. 📁 Document Tracking Ledger
* Tracks employee document compliance directly in the master database (`Submitted`, `Not Submitted`, `Returned`).

---

## 🛠️ System Architecture & Files

The system consists of three main files in Google Apps Script:
1. **`Code.gs`**: The main backend server file. Contains the calculation logic, PDF generation HTML, email HTML templates, and the batch processing loop.
2. **`Dashboard.html`**: A heavy, interactive front-end web portal (UI) built with Bootstrap 5, Chart.js, and custom CSS. 
3. **`Dispatcher.html`**: A lightweight popup interface for triggering the email dispatcher directly from the Google Sheet menu.

---

## ⚙️ Initial Setup Guide

### Step 1: Prepare the Google Sheet
1. Create a Master Sheet named exactly **`ID_DATA`**.
2. Create columns for: `Status` (Active/Inactive), `ID`, `Designation`, `Name`, `Email`, `Bank`, `Account Title`, `Account No`, `Basic Salary`, `Medical Tier (0-5)`, `Total Loan`, `Remaining Loan`, `Installment`, and `Document Status`.
3. Create Monthly Sheets (e.g., **`Jan`**, **`Feb`**).
4. In Monthly sheets, ensure biometrics/attendance data (Late Mins, Absents, Sandwich, Bonus, Fine) is aligned properly for the script to read.

### Step 2: Install the Code
1. Open your Google Sheet > Click **Extensions** > **Apps Script**.
2. Create 3 files (`Code.gs`, `Dashboard.html`, `Dispatcher.html`).
3. Paste the respective code into each file and Save (`Ctrl+S`).
4. Reload your Google Sheet. You will see a new custom menu: **"SET ERP System"**.

---

## 🚀 How to Use the System (Standard Workflow)

### Phase 1: Calculation
1. Enter the raw attendance and bonus/fine data into the current month's sheet (e.g., `Jan`).
2. Click **SET ERP System** > **Process Attendance & Salary** from the top menu.
3. Enter the month name (`Jan`) when prompted.
4. The system will calculate everything in 2 seconds, format the sheet beautifully (zebra striping, red for deductions, green for net pay), and freeze the header.
5. **Important:** An `Extra Message` column will be created at the extreme right.

### Phase 2: Maker / Checker & Customization
1. Review the calculated "NET PAYABLE" column.
2. If an employee needs a specific warning or appreciation (e.g., *"Great performance this month!"* or *"Please submit your ID card"*), type it in their respective row under the **Extra Message** column.

### Phase 3: Dispatch & PDF Generation
1. Click **SET ERP System** > **Launch Cloud Dispatcher**.
2. Enter the month name (`Jan`) and click **Initialize Protocol**.
3. A glowing, animated terminal will appear. It will:
   * Query the database.
   * Generate PDFs and save them to Google Drive (`Pay slip details > Jan`).
   * Inject the Extra Messages into the HTML email bodies.
   * Dispatch emails in batches to avoid spam filters.
4. A Green Success screen will appear once 100% of emails are sent.

---

## 🔮 Future Roadmap (Coming Soon)
* **UI/UX Enhancements:** Continuous optimization for smoother navigation.
* **New Joiner Management:** A dedicated onboarding module to seamlessly integrate new employees.
* **Company Letters Generation:** Automated system for Offer Letters, Experience Certificates, and Warning Notices.
* **Standalone Web Portal:** Full migration to a headless React/Laravel architecture.

---
*Developed securely for Company.*
