/**
 * SAP B1 REQUIREMENT GATHERING — GOOGLE SHEETS BACKEND
 * ----------------------------------------------------
 * SETUP (one time):
 * 1. Create a new Google Sheet (sheets.new).
 * 2. Extensions → Apps Script. Delete any code there and paste this whole file.
 * 3. In the toolbar, select the function "setupSheet" and click Run.
 *    (Authorize when prompted. This creates the "Requirements" tab pre-loaded
 *    with the full SAP B1 questionnaire.)
 * 4. Click Deploy → New deployment → type: Web app.
 *      - Execute as: Me
 *      - Who has access: Anyone   (or "Anyone in your organization")
 * 5. Copy the Web app URL (ends with /exec) and paste it into the webpage.
 *
 * After any code change: Deploy → Manage deployments → Edit → New version.
 */

const SHEET_NAME = 'Requirements';
const HEADERS = ['Req ID','Module','Category','Question / Requirement','Client Response',
  'Priority','SAP B1 Fit','Gap (Y/N)','Proposed Solution','Owner','Status','Notes','Updated At'];

/* ---------------- WEB API ---------------- */

function doGet() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    rows.push({
      reqId: String(values[i][0]), module: values[i][1], category: values[i][2],
      question: values[i][3], response: values[i][4], priority: values[i][5],
      fit: values[i][6], gap: values[i][7], solution: values[i][8],
      owner: values[i][9], status: values[i][10], notes: values[i][11],
      updatedAt: values[i][12] ? String(values[i][12]) : ''
    });
  }
  return json_({ ok: true, rows: rows });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'upsert') return json_(upsert_(body.record));
    if (body.action === 'delete') return json_(remove_(body.reqId));
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function upsert_(r) {
  const sheet = getSheet_();
  const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat().map(String);
  const now = new Date().toISOString().replace('T', ' ').substring(0, 19);
  const rowData = [r.reqId, r.module, r.category, r.question, r.response, r.priority,
    r.fit, r.gap, r.solution, r.owner, r.status, r.notes, now];
  const idx = ids.indexOf(String(r.reqId));
  if (idx >= 0) {
    sheet.getRange(idx + 2, 1, 1, HEADERS.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return { ok: true, reqId: r.reqId, updatedAt: now };
}

function remove_(reqId) {
  const sheet = getSheet_();
  const ids = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues().flat().map(String);
  const idx = ids.indexOf(String(reqId));
  if (idx < 0) return { ok: false, error: 'Req ID not found' };
  sheet.deleteRow(idx + 2);
  return { ok: true };
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); sheet.appendRow(HEADERS); }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- ONE-TIME SETUP: SEEDS ALL QUESTIONS ---------------- */

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(SHEET_NAME, 0);
  sheet.appendRow(HEADERS);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold')
    .setBackground('#1F3864').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);

  const rows = [];
  QUESTION_BANK.forEach(function (mod) {
    mod.items.forEach(function (it, i) {
      rows.push([mod.prefix + '-' + String(i + 1).padStart(3, '0'), mod.module,
        it[0], it[1], '', '', '', '', '', '', 'Open', '', '']);
    });
  });
  sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);

  // Dropdown validations
  const n = rows.length;
  setList_(sheet, 6, n, ['Must Have', 'Should Have', 'Nice to Have']);
  setList_(sheet, 7, n, ['Standard', 'Configuration', 'Customization', 'Add-on', 'Out of Scope']);
  setList_(sheet, 8, n, ['Y', 'N']);
  setList_(sheet, 11, n, ['Open', 'In Discussion', 'Confirmed', 'Deferred', 'Closed']);
  sheet.setColumnWidths(1, 1, 90); sheet.setColumnWidths(2, 2, 160);
  sheet.setColumnWidths(4, 1, 420); sheet.setColumnWidths(5, 1, 320);
  sheet.setColumnWidths(9, 1, 280);
  SpreadsheetApp.flush();
}

function setList_(sheet, col, n, values) {
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(values, true).build();
  sheet.getRange(2, col, n, 1).setDataValidation(rule);
}

const QUESTION_BANK = [
 { module:'1. Company Profile', prefix:'CP', items:[
  ['Legal Entity','Full legal name of the company, registered address, and company registration number?'],
  ['Legal Entity','Number of legal entities / companies to be implemented in SAP B1? Will inter-company transactions be required?'],
  ['Legal Entity','Branches / plants / sales offices and their locations? Should each be tracked separately (branch functionality)?'],
  ['Industry','Primary industry and nature of business (trading, manufacturing, services, distribution, project-based)?'],
  ['Industry','Key products / services sold and key materials purchased?'],
  ['Organization','Organization chart: departments, headcount, and key decision makers for the project?'],
  ['Organization','Named project sponsor, project manager, and key users per module (Finance, Sales, Purchase, Inventory, Production)?'],
  ['Fiscal','Financial year period (e.g., Apr\u2013Mar or Jan\u2013Dec)? Number of posting periods (monthly/quarterly)?'],
  ['Fiscal','Planned go-live date and cut-over date for opening balances?'],
  ['Currency','Local currency, system currency, and foreign currencies used? How are exchange rates updated (manual / automatic)?'],
  ['Tax','Tax regime applicable (e.g., GST in India: CGST/SGST/IGST, TDS/TCS)? GSTIN numbers per state/branch?'],
  ['Compliance','Statutory reporting requirements: e-invoicing, e-way bill, GST returns, audit requirements?'],
  ['Localization','Which SAP B1 localization is required (e.g., India, US, UK)? Any multi-country operations?'],
  ['Language','Working language(s) required for the system and printed documents?'],
  ['Volume','Approximate monthly transaction volumes: sales invoices, purchase invoices, journal entries, deliveries?'],
  ['Volume','Number of items (SKUs), customers, vendors, and employees to be managed?']]},
 { module:'2. IT Infrastructure', prefix:'IT', items:[
  ['Deployment','Preferred deployment: on-premise, private cloud, or hosted? Any existing server infrastructure?'],
  ['Database','Database preference: SAP HANA or Microsoft SQL Server?'],
  ['Licensing','Estimated number of users: Professional, Limited (CRM/Financials/Logistics), and Indirect Access users?'],
  ['Licensing','How many users will use the system simultaneously?'],
  ['Hardware','Existing server specifications (CPU, RAM, storage)? Or should sizing recommendations be provided?'],
  ['Network','Network setup: LAN at HO, connectivity at branches (VPN/MPLS/internet)? Bandwidth availability?'],
  ['Client Access','Access methods needed: Windows client, Web client, mobile app (Sales/Service), remote desktop?'],
  ['Peripherals','Printers (invoice, label, barcode), barcode scanners, weighing scales, or other devices to integrate?'],
  ['Backup','Backup and disaster recovery expectations: RPO/RTO, backup frequency, offsite storage?'],
  ['Security','IT security policies: password policy, domain authentication (SSO), data access restrictions?'],
  ['Environment','Non-production environments required: test/sandbox company database, training database?'],
  ['Email','Email server details for sending documents from SAP B1 (SMTP/Office 365)?']]},
 { module:'3. Financials', prefix:'FIN', items:[
  ['Chart of Accounts','Existing chart of accounts structure \u2014 number of levels, drawer structure? Redesigned or migrated as-is?'],
  ['Segmentation','Are segmented accounts or cost dimensions needed (department, branch, project, product line)?'],
  ['Cost Accounting','Cost centers / profit centers required? Distribution rules for overhead allocation?'],
  ['Projects','Project accounting needed? Should revenues/costs be tracked per project with stages and budgets?'],
  ['Budgets','Budget control required? At GL account level? Warning or blocking when budget exceeded?'],
  ['Journal Entries','Recurring journal entries, posting templates, or reversal transactions used today?'],
  ['Multi-currency','Foreign currency transactions: revaluation frequency, realized/unrealized gain-loss posting requirements?'],
  ['Taxation','Tax codes required (GST rates, exempt, zero-rated)? Withholding tax (TDS/TCS) setup with sections and thresholds?'],
  ['Fixed Assets','Fixed asset register required? Depreciation methods (SLM/WDV), asset classes, Companies Act vs IT Act books?'],
  ['Banking','House banks list, account details, and cheque printing formats required?'],
  ['Banking','Bank reconciliation process: manual or via bank statement import (MT940, CSV, Excel)?'],
  ['Payments','Payment methods: cheque, NEFT/RTGS, wire; is the Payment Wizard needed for batch vendor payments?'],
  ['AR/AP','Customer/vendor aging buckets required (30/60/90/120)? Dunning letters for overdue receivables?'],
  ['Period End','Month-end and year-end closing process today: checklists, accrual entries, period closing controls?'],
  ['Financial Reports','Required statements: Balance Sheet, P&L, Cash Flow, Trial Balance \u2014 formats and comparisons (budget vs actual, YoY)?'],
  ['Statutory','GST returns (GSTR-1, GSTR-3B), TDS returns, e-invoicing (IRP), e-way bill generation requirements?'],
  ['Controls','Approval requirements for journal entries, payments above thresholds, or credit notes?'],
  ['Opening Balances','Source and quality of opening balances: GL, open AR/AP, open POs/SOs, stock as on cut-over date?']]},
 { module:'4. Sales & AR', prefix:'SLS', items:[
  ['Process Flow','Describe the end-to-end sales process: enquiry \u2192 quotation \u2192 order \u2192 delivery \u2192 invoice \u2192 collection.'],
  ['Documents','Sales documents needed: Quotation, Sales Order, Delivery, AR Invoice, Credit Memo, Returns, Down Payment?'],
  ['Customers','Customer master data: groups, territories, payment terms, credit limits, multiple ship-to addresses?'],
  ['Pricing','Pricing structure: price lists, customer-specific prices, period/volume discounts, discount groups?'],
  ['Credit Control','Credit limit checks: warning or block at order/delivery/invoice? Who can override?'],
  ['Sales Team','Sales employees / commission structure to be tracked? Commission calculation rules?'],
  ['Approvals','Approval workflows needed (discount above X%, order above credit limit)?'],
  ['Dispatch','Delivery process: partial deliveries, consolidated deliveries, transporter details, freight charges?'],
  ['Invoicing','Invoicing scenarios: invoice per delivery, consolidated monthly invoice, milestone/progress billing?'],
  ['E-Invoicing','E-invoice (IRN/QR) and e-way bill generation required from within SAP B1?'],
  ['Returns','Sales returns / credit note process and reasons tracking?'],
  ['Print Layouts','Print formats: quotation, order confirmation, delivery challan, tax invoice \u2014 branding and statutory content?'],
  ['Reports','Key sales reports: analysis by customer/item/salesperson/region, order book, pending deliveries?'],
  ['Special','Special scenarios: consignment sales, export sales (LUT/with tax), SEZ customers, sample/FOC issues?']]},
 { module:'5. Purchasing & AP', prefix:'PUR', items:[
  ['Process Flow','Describe the procurement process: purchase request \u2192 RFQ \u2192 PO \u2192 GRN \u2192 invoice \u2192 payment.'],
  ['Documents','Documents needed: Purchase Request, RFQ, Purchase Order, GRPO, AP Invoice, AP Credit Memo, Landed Costs?'],
  ['Vendors','Vendor master data: groups, payment terms, addresses, bank details, MSME classification?'],
  ['Approvals','PO approval matrix: by amount, item category, department? Multi-level approvals?'],
  ['Pricing','Vendor price lists, last purchase price usage, quotation comparison requirements?'],
  ['Imports','Import purchases: landed cost allocation (freight, duty, clearing), customs docs, foreign vendor payments?'],
  ['GRN','Goods receipt: quality check before acceptance, partial receipts, over/under delivery tolerance?'],
  ['3-Way Match','Invoice verification: 3-way matching (PO\u2013GRN\u2013Invoice), tolerance rules, blocking mismatched invoices?'],
  ['Subcontracting','Subcontracting / job work: material sent to vendor, receipt of processed goods, ITC-04 reporting?'],
  ['Services','Service procurement (non-inventory): service POs, contract-based purchasing, recurring services?'],
  ['TDS','TDS deduction on vendor invoices: sections, thresholds, lower deduction certificates?'],
  ['Reports','Key purchase reports: pending POs, GRN vs invoice status, vendor performance, purchase analysis?']]},
 { module:'6. Inventory & Warehouse', prefix:'INV', items:[
  ['Item Master','Item master structure: item groups, properties, UoM groups (purchase/sales/inventory conversions)?'],
  ['Item Types','Item types used: inventory, non-inventory, service, fixed asset items? Any phantom/kit items?'],
  ['Coding','Item coding logic: existing codes retained or new scheme? Manual or auto-generated?'],
  ['Warehouses','Number of warehouses/locations? Bin location management within warehouses?'],
  ['Valuation','Inventory valuation method: Moving Average, FIFO, or Standard Cost? Per item or global?'],
  ['Batch/Serial','Batch or serial managed items? Expiry tracking, batch attributes, recall traceability?'],
  ['Transfers','Inter-warehouse / inter-branch transfers; transfer request, in-transit tracking, e-way bill on transfers?'],
  ['Counting','Physical inventory / cycle counting process and frequency? Stock adjustment approval?'],
  ['Barcode','Barcode usage: labeling, scanning at GRN/dispatch/counting? Formats (EAN/Code128/QR)?'],
  ['Min/Max','Reorder planning: min/max stock levels, reorder quantities per warehouse? MRP requirement?'],
  ['Pick & Pack','Pick and pack process: pick lists, wave picking, packing slips?'],
  ['Special Stock','Consignment (in/out), customer-owned material, demo stock, scrap management?'],
  ['Reports','Stock status, stock aging, slow/non-moving, inventory valuation, batch traceability?']]},
 { module:'7. Production & MRP', prefix:'PRD', items:[
  ['Applicability','Is manufacturing in scope? Type: make-to-stock, make-to-order, assembly, process, job-work based?'],
  ['BOM','Bill of Materials: levels, BOM types (production/sales/assembly/template), by-products/co-products?'],
  ['Routing','Routing/operations and work centers to be tracked (resources)? Machine/labour capacity?'],
  ['Orders','Production order process: planned \u2192 released \u2192 closed; backflush or manual component issue?'],
  ['Costing','Product costing: standard vs actual, overhead absorption, WIP tracking, variance analysis?'],
  ['MRP','MRP: planning horizon, demand sources (sales orders, forecasts, min stock), recommendations to PO/production?'],
  ['Quality','QC steps: incoming, in-process, final inspection? QC results recording (add-on may be needed)?'],
  ['Scrap','Scrap/wastage recording and yield tracking requirements?'],
  ['Subcontract','Outsourced operations within production process (job work stages)?'],
  ['Shop Floor','Shop floor data capture: manual entry or real-time via terminals/scanners?'],
  ['Reports','Order status, WIP, consumption vs standard, production variance, capacity load?']]},
 { module:'8. CRM & Service', prefix:'CRM', items:[
  ['Leads','Lead and opportunity management required? Pipeline stages, win/loss reasons?'],
  ['Activities','Activity management: calls, meetings, follow-ups linked to customers/opportunities? Calendar sync?'],
  ['Campaigns','Marketing campaign tracking and lead generation sources?'],
  ['Service Contracts','After-sales service in scope? Service contracts, warranty templates (customer/serial-based)?'],
  ['Service Calls','Service call process: logging, technician assignment, resolution tracking, SLA response times?'],
  ['Equipment','Customer equipment cards (installed base) with serial numbers?'],
  ['Spare Parts','Spare parts consumption against service calls; billing for out-of-warranty service?'],
  ['Mobile','Field service via mobile app for technicians?'],
  ['Reports','Pipeline, conversion rates, open service calls, SLA compliance, technician productivity?']]},
 { module:'9. Reporting & Analytics', prefix:'REP', items:[
  ['MIS','Top 10 management reports used today (attach samples). Frequency and recipients?'],
  ['Dashboards','Dashboard/KPI requirements: sales, collections, inventory, production KPIs? For which roles?'],
  ['Tools','Reporting tools: Crystal Reports, SAP B1 queries, Excel integration, Pervasive Analytics (HANA)?'],
  ['Scheduling','Report scheduling and email distribution (e.g., daily sales report at 9 AM)?'],
  ['Print Layouts','Complete list of print layouts to develop (invoice, PO, GRN, challan) with sample formats?'],
  ['Alerts','Alerts: credit limit breach, stock below minimum, overdue payments, pending approvals?'],
  ['Excel','XL Reporter / Excel add-in usage for finance team?'],
  ['External BI','External BI tools (Power BI, Tableau) to connect to the SAP B1 database?']]},
 { module:'10. Authorizations', prefix:'AUT', items:[
  ['Roles','User roles and count per role (Finance Manager, Accountant, Sales Executive, Store Keeper)?'],
  ['Access Matrix','Module-wise access per role: full / read-only / no access?'],
  ['Data Ownership','Should sales staff see only their own customers/orders?'],
  ['Field Level','Sensitive field restrictions: cost price visibility, salary data, bank details?'],
  ['Approvals','Consolidated list of approval procedures across modules with approvers and thresholds.'],
  ['Audit','Audit trail requirements: change log access, amendment restrictions after printing/approval?'],
  ['Series','Document numbering series per branch/document type: prefixes and yearly reset rules?']]},
 { module:'11. Data Migration', prefix:'MIG', items:[
  ['Masters','Masters to migrate: chart of accounts, customers, vendors, items, price lists, assets \u2014 record counts?'],
  ['Source','Current system(s) holding data (Tally, Excel, legacy ERP)? Export capability and data quality?'],
  ['Cleansing','Who is responsible for data cleansing and migration templates? Timeline?'],
  ['Open Transactions','Open SOs, POs, AR/AP invoices at cut-over \u2014 migrate or re-enter?'],
  ['Balances','Opening balances: GL trial balance, item-wise stock with batches/serials, asset register?'],
  ['History','Years of history needed in SAP B1 vs archived reference?'],
  ['Validation','Who verifies migrated balances match the legacy system? Sign-off process?'],
  ['Tools','Migration via DTW templates / import wizards \u2014 training needed for client team?']]},
 { module:'12. Integrations & Add-ons', prefix:'INT', items:[
  ['Systems','External systems to integrate: e-commerce, POS, WMS, payroll, banking, portals?'],
  ['E-Invoice/EWB','GSP/ASP provider for e-invoicing and e-way bill integration?'],
  ['Payroll','Payroll in SAP B1 add-on or external? If external, JV import format?'],
  ['EDI/API','EDI or API-based data exchange with customers/vendors (orders, invoices, ASN)?'],
  ['Add-ons','Known add-on needs: QC, WMS/barcode, payroll, industry-specific solutions?'],
  ['Direction','For each integration: direction, frequency (real-time/batch), and data volume?'],
  ['Legacy','Legacy systems to retire vs run in parallel? Duration of parallel run?']]}
];
