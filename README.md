# 🚀 Silver Edge Technologies - Cloud ERP & Payroll System

## 📌 System Overview
Silver Edge ERP is a custom-built, serverless payroll and dispatch architecture hosted on Google Apps Script. It automates attendance processing, applies complex financial penalties/bonuses, generates high-fidelity PDF salary slips, and dispatches them via email using an intelligent anti-timeout batching system.

## 🏗️ System Components (The 3 Pillars)

### 1. `Code.gs` (The Backend Engine)
This is the brain of the system. It handles data aggregation, complex financial mathematics, and APIs.
* **Aggregator:** Reads Raw Sheet (Cols A-J) and Master Data (`ID_DATA`).
* **Processor:** Calculates Late Mins, Unpaid Days, Sandwich Rule, FBR Taxes, and Bonuses.
* **Output Engine:** Prints the finalized ledger starting from **Column M**.
* **Cloud API:** Generates HTML-to-PDF blobs, saves to Google Drive, and connects to Gmail App API.

### 2. `Dispatcher.html` (The Cloud Dispatcher)
A frontend UI designed to bypass Google's 6-minute execution limit.
* **Batch Processing:** Sends emails in batches of **15 employees** at a time.
* **Anti-Spam Filter:** Enforces a strict **3-second delay** between emails to prevent Gmail account blocks.
* **Live UI:** Displays a real-time progress bar and success status.

### 3. `Dashboard.html` (The Live Simulator)
A standalone interface for HR to test payroll policies before actual execution.
* **Dynamic Calculator:** Instantly calculates net pay based on input variables.
* **Visualizer:** Features a Chart.js Doughnut chart for Gross vs Deductions.
* **Live Slip Preview:** Renders a real-time preview of the corporate PDF salary slip.

## 🔄 Operational Workflow (How to Use)
1. **Input Data:** HR inputs monthly raw data into Columns A to J.
2. **Process Salary:** Click `🚀 Silver Edge ERP > 1️⃣ Process Attendance & Salary`. The system will populate Columns M to AB with final calculations.
3. **Dispatch Slips:** Click `Option 2 (Cloud Dispatcher)`. Enter the month name and click Start. The system will auto-generate PDFs, save them in Drive (Year/Month folders), and email them.
4. **Simulate (Optional):** Use `Option 3 (Live Dashboard)` anytime to manually test specific employee salary scenarios.

## 🧮 Core Calculation Logic
* **Per Day Rate:** `Basic Salary ÷ 30`
* **Per Minute Rate:** `Basic Salary ÷ 16,200`
* **Total Deductions:** `Late Penalty + Absent Penalty + Sandwich Penalty + Fine + Tax + Fixed(1280)`
* **Net Payable:** `(Basic + Bonus) - Total Deductions`
