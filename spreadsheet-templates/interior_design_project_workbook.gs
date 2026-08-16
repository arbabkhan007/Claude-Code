/**
 * Interior Design Project Management Workbook Builder
 * Google Sheets Apps Script
 *
 * How to use:
 * 1. Create/open a blank Google Sheet.
 * 2. Go to Extensions > Apps Script.
 * 3. Paste this file into Code.gs.
 * 4. Save, then run buildInteriorDesignProjectWorkbook().
 * 5. Authorize the script when prompted.
 *
 * The generated workbook is designed to work in Google Sheets and export cleanly
 * to Excel (.xlsx). Most formulas, dropdowns, filters, and formatting are
 * compatible after export. Google-only dashboard charts may need minor resizing
 * in Excel after download.
 */

const TEMPLATE = {
  rows: 250,
  colors: {
    navy: '#17365D',
    blue: '#1F4E78',
    teal: '#0F766E',
    lightBlue: '#D9EAF7',
    lightTeal: '#DDEFEA',
    green: '#D9EAD3',
    yellow: '#FFF2CC',
    red: '#F4CCCC',
    gray: '#D9D9D9',
    darkGray: '#666666',
    white: '#FFFFFF',
    black: '#000000'
  }
};

const LISTS = {
  projectTypes: ['Residential', 'Commercial', 'Office', 'Retail', 'Hospitality', 'Staging'],
  yesNo: ['Yes', 'No'],
  projectStatuses: ['Not started', 'In progress', 'On hold', 'Complete', 'Delayed', 'Cancelled'],
  billingMethods: ['Flat fee', 'Hourly', 'Percentage', 'Hybrid'],
  designStatuses: ['Not started', 'Concept', 'Sourcing', 'Client review', 'Approved', 'Ordered', 'Installed', 'Complete', 'On hold'],
  priorities: ['Low', 'Medium', 'High', 'Urgent'],
  roomTypes: ['Living room', 'Bedroom', 'Kitchen', 'Bathroom', 'Dining room', 'Entryway', 'Office', 'Nursery', 'Outdoor', 'Commercial space', 'Other'],
  ffeCategories: ['Sofa', 'Chair', 'Table', 'Bed', 'Rug', 'Lighting', 'Artwork', 'Mirror', 'Window treatment', 'Hardware', 'Plumbing fixture', 'Appliance', 'Decor', 'Textile', 'Cabinetry', 'Flooring', 'Wallcovering', 'Other'],
  itemStatuses: ['Sourcing', 'Option selected', 'Client review', 'Approved', 'Ordered', 'Shipped', 'Delivered', 'Installed', 'Cancelled', 'Backordered'],
  orderStatuses: ['Not ordered', 'Ordered', 'Confirmed', 'In production', 'Shipped', 'Delivered', 'Delayed', 'Cancelled', 'Returned'],
  approvalStatuses: ['Not sent', 'Sent', 'Client review', 'Approved', 'Revisions requested', 'Rejected', 'On hold'],
  deliveryStatuses: ['Not ordered', 'Ordered', 'Shipped', 'Delivered', 'Delayed', 'Cancelled', 'Returned', 'Backordered'],
  installStatuses: ['Not scheduled', 'Scheduled', 'Installed', 'Punch list', 'Delayed', 'Cancelled'],
  fixtureTypes: ['Lighting', 'Faucet', 'Sink', 'Shower fixture', 'Tub', 'Toilet', 'Cabinet hardware', 'Door hardware', 'Appliance', 'Electrical', 'Other'],
  finishTypes: ['Paint', 'Wallpaper', 'Flooring', 'Tile', 'Countertop', 'Backsplash', 'Cabinet finish', 'Fabric', 'Trim', 'Stone', 'Wood', 'Metal', 'Other'],
  surfaces: ['Wall', 'Ceiling', 'Floor', 'Cabinet', 'Countertop', 'Shower wall', 'Backsplash', 'Upholstery', 'Drapery'],
  vendorTypes: ['Furniture', 'Lighting', 'Fabric', 'Flooring', 'Tile', 'Paint', 'Wallpaper', 'Decor', 'Hardware', 'Contractor', 'Installer', 'Workroom', 'Custom furniture', 'Other'],
  budgetCategories: ['Furniture', 'Lighting', 'Rugs', 'Decor', 'Artwork', 'Window treatments', 'Flooring', 'Paint/wallcovering', 'Contractor labor', 'Installation', 'Shipping', 'Taxes', 'Design fee', 'Contingency'],
  damageChecks: ['Not checked', 'Good condition', 'Damaged', 'Missing item', 'Wrong item'],
  approvalTypes: ['Furniture', 'Finish', 'Layout', 'Color palette', 'Budget', 'Purchase', 'Vendor', 'Design concept'],
  phases: ['Consultation', 'Measurements', 'Concept design', 'Sourcing', 'Client approval', 'Purchasing', 'Construction', 'Delivery', 'Installation', 'Styling', 'Final reveal', 'Punch list'],
  taskStatuses: ['Not started', 'In progress', 'Waiting on client', 'Waiting on vendor', 'Complete', 'Delayed'],
  measurementTypes: ['Room', 'Wall', 'Window', 'Door', 'Cabinet', 'Furniture space', 'Rug area', 'Lighting location'],
  trades: ['General contractor', 'Painter', 'Electrician', 'Plumber', 'Flooring installer', 'Tile installer', 'Wallpaper installer', 'Carpenter', 'Window treatment installer', 'Furniture installer', 'Cleaner', 'Other'],
  paymentStatuses: ['Draft', 'Sent', 'Paid', 'Partially paid', 'Overdue', 'Cancelled', 'Unpaid'],
  meetingTypes: ['Client consultation', 'Vendor meeting', 'Site visit', 'Installation meeting', 'Contractor meeting', 'Final walkthrough'],
  preferredOptions: ['Option 1', 'Option 2', 'Option 3']
};

const SHEETS = [
  { name: 'Welcome - Instructions', tabColor: '#17365D' },
  { name: 'Project Dashboard', tabColor: '#1F4E78' },
  { name: 'Client & Project Profile', tabColor: '#0F766E' },
  { name: 'Room List', tabColor: '#6AA84F', headers: ['Room name', 'Room type', 'Floor/level', 'Square footage', 'Design status', 'Budget allocation', 'Estimated spend', 'Actual spend', 'Remaining budget', 'Priority level', 'Notes'] },
  { name: 'Master FF&E Schedule', tabColor: '#674EA7', headers: ['Item ID', 'Room', 'Category', 'Item name', 'Description', 'Quantity', 'Dimensions', 'Material', 'Finish', 'Color', 'Vendor', 'Brand', 'Product link', 'SKU/model number', 'Unit cost', 'Total cost', 'Tax', 'Shipping', 'Discount', 'Final cost', 'Lead time', 'Order status', 'Approval status', 'Delivery status', 'Install status', 'Notes', 'Image link'] },
  { name: 'Furniture Schedule', tabColor: '#8E7CC3', headers: ['Room', 'Furniture item', 'Dimensions', 'Quantity', 'Vendor', 'Product link', 'Fabric/material', 'Finish', 'Cost', 'Approval status', 'Order status', 'Delivery date', 'Install date'] },
  { name: 'Fixtures Schedule', tabColor: '#C27BA0', headers: ['Room', 'Fixture type', 'Item name', 'Finish', 'Size', 'Vendor', 'Model/SKU', 'Quantity', 'Unit cost', 'Total cost', 'Required rough-in date', 'Order deadline', 'Delivery date', 'Installer notes'] },
  { name: 'Finishes Schedule', tabColor: '#A64D79', headers: ['Room', 'Surface/application', 'Finish type', 'Product name', 'Color/finish', 'Material', 'Vendor', 'Product link', 'SKU', 'Quantity needed', 'Unit cost', 'Total cost', 'Sample ordered', 'Sample approved', 'Installer notes'] },
  { name: 'Product Sourcing Tracker', tabColor: '#E69138', headers: ['Room', 'Item needed', 'Option 1 link', 'Option 1 price', 'Option 2 link', 'Option 2 price', 'Option 3 link', 'Option 3 price', 'Preferred option', 'Pros', 'Cons', 'Client feedback', 'Final selection'] },
  { name: 'Vendor & Supplier Database', tabColor: '#B45F06', headers: ['Vendor name', 'Vendor type', 'Contact person', 'Email', 'Phone', 'Website', 'Account number', 'Trade discount %', 'Lead time', 'Shipping terms', 'Return policy', 'Notes'] },
  { name: 'Budget Tracker', tabColor: '#38761D', headers: ['Budget category', 'Budgeted amount', 'Estimated amount', 'Actual amount', 'Difference', 'Paid amount', 'Balance due', 'Payment status', 'Notes'] },
  { name: 'Purchase Order Tracker', tabColor: '#134F5C', headers: ['PO number', 'Vendor', 'Room', 'Item', 'Quantity', 'Order date', 'Expected ship date', 'Expected delivery date', 'Actual delivery date', 'Order status', 'Tracking number', 'Carrier', 'Invoice number', 'Amount', 'Paid status', 'Notes'] },
  { name: 'Delivery & Installation', tabColor: '#45818E', headers: ['Item', 'Room', 'Vendor', 'Delivery address', 'Expected delivery date', 'Actual delivery date', 'Delivery status', 'Received by', 'Damage check', 'Damage notes', 'Storage location', 'Install date', 'Installer', 'Installation status', 'Punch list notes'] },
  { name: 'Client Approval Tracker', tabColor: '#3C78D8', headers: ['Item/design decision', 'Room', 'Approval type', 'Sent to client date', 'Client response date', 'Status', 'Approved option', 'Client comments', 'Revision needed', 'Final approval date'] },
  { name: 'Project Timeline', tabColor: '#990000', headers: ['Phase', 'Task', 'Assigned to', 'Start date', 'Due date', 'Completion date', 'Status', 'Priority', 'Dependencies', 'Notes', 'Duration days', 'Days overdue', 'Gantt W1', 'Gantt W2', 'Gantt W3', 'Gantt W4', 'Gantt W5', 'Gantt W6', 'Gantt W7', 'Gantt W8', 'Gantt W9', 'Gantt W10', 'Gantt W11', 'Gantt W12'] },
  { name: 'Task & To-Do List', tabColor: '#CC0000', headers: ['Task', 'Room', 'Category', 'Assigned to', 'Priority', 'Due date', 'Status', 'Notes', 'Days overdue'] },
  { name: 'Measurement Log', tabColor: '#741B47', headers: ['Room', 'Measurement type', 'Width', 'Height', 'Depth', 'Length', 'Ceiling height', 'Window size', 'Door size', 'Wall length', 'Notes'] },
  { name: 'Contractor Installer Tracker', tabColor: '#7F6000', headers: ['Contractor/vendor name', 'Trade', 'Contact', 'Scope', 'Quote amount', 'Start date', 'End date', 'Payment status', 'Insurance received', 'Contract signed', 'Notes'] },
  { name: 'Invoice Payment Tracker', tabColor: '#274E13', headers: ['Invoice number', 'Invoice date', 'Client', 'Description', 'Amount', 'Due date', 'Paid date', 'Payment status', 'Payment method', 'Notes', 'Balance due'] },
  { name: 'Project Meeting Notes', tabColor: '#666666', headers: ['Date', 'Meeting type', 'Attendees', 'Discussion points', 'Decisions made', 'Action items', 'Follow-up date', 'Notes'] },
  { name: 'Client FFE Report', tabColor: '#D9EAD3', headers: ['Room', 'Item name', 'Description', 'Quantity', 'Vendor', 'Product link', 'Final cost', 'Approval status', 'Image link', 'Client notes'] },
  { name: 'Room Budget Report', tabColor: '#D9EAD3', headers: ['Room', 'Budget allocation', 'Estimated spend', 'Actual spend', 'Remaining budget', 'Budget used %', 'Status'] },
  { name: 'Vendor Order Report', tabColor: '#D9EAD3', headers: ['Vendor', 'PO number', 'Room', 'Item', 'Quantity', 'Order status', 'Expected delivery date', 'Amount', 'Paid status', 'Tracking number'] },
  { name: 'Installation Checklist', tabColor: '#D9EAD3', headers: ['Room', 'Item', 'Vendor', 'Delivery status', 'Install date', 'Installer', 'Installation status', 'Damage check', 'Punch list notes', 'Complete?'] },
  { name: 'Final Punch List', tabColor: '#D9EAD3', headers: ['Room', 'Issue/item', 'Responsible party', 'Priority', 'Due date', 'Status', 'Resolution notes', 'Completed date'] }
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Interior Design Template')
    .addItem('Build / Reset Workbook', 'buildInteriorDesignProjectWorkbook')
    .addToUi();
}

function buildInteriorDesignProjectWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Building interior design project workbook...', 'Template Builder', 5);
  resetWorkbook_(ss);
  createDataValidationSheet_(ss);

  SHEETS.forEach(config => {
    const sheet = ss.getSheetByName(config.name) || ss.insertSheet(config.name);
    sheet.setTabColor(config.tabColor);
    if (config.headers) setupTableSheet_(ss, sheet, config.headers, config.name);
  });

  setupWelcome_(ss.getSheetByName('Welcome - Instructions'));
  setupProfile_(ss, ss.getSheetByName('Client & Project Profile'));
  setupBudgetSeed_(ss.getSheetByName('Budget Tracker'));
  applyAllValidations_(ss);
  applyAllFormulas_(ss);
  setupDashboard_(ss, ss.getSheetByName('Project Dashboard'));
  setupPrintableReports_(ss);
  hideHelperSheet_(ss);

  ss.setActiveSheet(ss.getSheetByName('Welcome - Instructions'));
  ss.toast('Done. Your workbook is ready.', 'Template Builder', 5);
}

function resetWorkbook_(ss) {
  const existing = ss.getSheets();
  let scratch = ss.getSheetByName('__Scratch') || ss.insertSheet('__Scratch');
  ss.setActiveSheet(scratch);
  existing.forEach(sh => {
    if (sh.getName() !== '__Scratch') ss.deleteSheet(sh);
  });
  SHEETS.forEach(config => ss.insertSheet(config.name));
  ss.deleteSheet(scratch);
}

function createDataValidationSheet_(ss) {
  const sheet = ss.insertSheet('Lists');
  sheet.setTabColor('#999999');
  const keys = Object.keys(LISTS);
  ensureGrid_(sheet, Math.max(...keys.map(key => LISTS[key].length)) + 1, keys.length);
  keys.forEach((key, i) => {
    const col = i + 1;
    sheet.getRange(1, col).setValue(key).setFontWeight('bold').setBackground(TEMPLATE.colors.gray);
    sheet.getRange(2, col, LISTS[key].length, 1).setValues(LISTS[key].map(v => [v]));
  });
  sheet.autoResizeColumns(1, keys.length);
}

function hideHelperSheet_(ss) {
  const sheet = ss.getSheetByName('Lists');
  if (sheet) sheet.hideSheet();
}

function ensureGrid_(sheet, minRows, minColumns) {
  if (sheet.getMaxRows() < minRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), minRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < minColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), minColumns - sheet.getMaxColumns());
  }
}

function setupTableSheet_(ss, sheet, headers, name) {
  ensureGrid_(sheet, TEMPLATE.rows + 10, headers.length);
  sheet.clear();
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setFontColor(TEMPLATE.colors.white)
    .setBackground(TEMPLATE.colors.blue)
    .setWrap(true)
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 42);
  sheet.getRange(2, 1, TEMPLATE.rows, headers.length).setVerticalAlignment('top');
  sheet.getRange(1, 1, TEMPLATE.rows + 1, headers.length).createFilter();
  sheet.getRange(1, 1, TEMPLATE.rows + 1, headers.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  sheet.autoResizeColumns(1, headers.length);
  sheet.setColumnWidths(1, headers.length, 130);
  if (name.includes('Notes')) sheet.setColumnWidths(4, 5, 220);
  if (name.includes('Master FF&E')) sheet.setColumnWidths(4, 24, 130);
  protectHeader_(sheet);
}

function setupWelcome_(sheet) {
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidths(1, 6, 170);
  sheet.getRange('A1:F1').merge().setValue('Interior Design Project Management Workbook')
    .setFontSize(20).setFontWeight('bold').setFontColor(TEMPLATE.colors.white).setBackground(TEMPLATE.colors.navy)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 48);

  const sections = [
    ['Quick Start Guide', 'Step 1: Add client/project details in Client & Project Profile.\nStep 2: Add rooms in Room List.\nStep 3: Add FF&E items in Master FF&E Schedule.\nStep 4: Track budget, orders, approvals, delivery, and installation.\nStep 5: Use Project Dashboard for a live overview.'],
    ['What Each Tab Is Used For', 'Profile stores client/project details. Room List tracks spaces and budgets. Master FF&E is the core item schedule. Furniture, Fixtures, Finishes, Sourcing, Vendors, Budget, PO, Delivery, Approvals, Timeline, Tasks, Measurements, Contractors, Invoices, and Meeting Notes support full project management. Printable report tabs create client/vendor/install views.'],
    ['How To Enter Data', 'Type into white cells beginning on row 2. Header rows are formatted/protected. Formula columns auto-calculate and should not be overwritten unless you intentionally customize the workbook.'],
    ['Color Coding', 'Green = complete / approved / on budget. Yellow = pending / in progress. Red = overdue / over budget / delayed. Gray = cancelled / on hold. Blue headers indicate table fields.'],
    ['Dropdown Instructions', 'Click cells in status/category/room/vendor/priority columns and choose from the dropdown. Room and vendor dropdowns update from the Room List and Vendor Database where used.'],
    ['Compatibility Note', 'Built in Google Sheets and designed to export to Excel (.xlsx). Formulas, filters, dropdowns, and conditional formatting are broadly compatible. Google chart placement may need minor adjustment after Excel export.'],
    ['Editable / Unlocked Version Note', 'This is an editable template. Header rows may be lightly protected to prevent accidental changes; remove sheet/range protection if you want a completely unlocked version.'],
    ['Support / Contact Note', 'Add your support email, shop URL, license terms, or client handoff notes here before selling or sharing this template.']
  ];

  let row = 3;
  sections.forEach(([title, text]) => {
    sheet.getRange(row, 1, 1, 6).merge().setValue(title)
      .setFontWeight('bold').setFontSize(13).setFontColor(TEMPLATE.colors.white).setBackground(TEMPLATE.colors.teal);
    sheet.getRange(row + 1, 1, 2, 6).merge().setValue(text)
      .setWrap(true).setVerticalAlignment('top').setBackground('#F8FBFD');
    sheet.setRowHeights(row + 1, 2, 42);
    row += 4;
  });
}

function setupProfile_(ss, sheet) {
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidths(1, 4, 210);
  sheet.getRange('A1:D1').merge().setValue('Client & Project Profile')
    .setFontSize(18).setFontWeight('bold').setFontColor(TEMPLATE.colors.white).setBackground(TEMPLATE.colors.teal)
    .setHorizontalAlignment('center');
  const fields = [
    ['Client name', ''], ['Client email', ''], ['Phone number', ''], ['Project name', ''], ['Project address', ''],
    ['Project type', 'Residential'], ['Design style', ''], ['Start date', ''], ['Target completion date', ''], ['Total budget', 0],
    ['Designer name', ''], ['Client notes', ''], ['Scope summary', ''], ['Project status', 'In progress'], ['Contract signed', 'No'],
    ['Deposit paid', 'No'], ['Design package selected', ''], ['Billing method', 'Flat fee']
  ];
  sheet.getRange(3, 1, fields.length, 2).setValues(fields);
  sheet.getRange(3, 1, fields.length, 1).setFontWeight('bold').setBackground(TEMPLATE.colors.lightTeal);
  sheet.getRange(3, 2, fields.length, 1).setBackground(TEMPLATE.colors.white);
  sheet.getRange('B12:B14').setWrap(true);
  setListValidation_(ss, sheet.getRange('B8'), 'projectTypes');
  setListValidation_(ss, sheet.getRange('B16'), 'projectStatuses');
  setListValidation_(ss, sheet.getRange('B17:B18'), 'yesNo');
  setListValidation_(ss, sheet.getRange('B20'), 'billingMethods');
  sheet.getRange('B12').setNumberFormat('$#,##0.00');
  sheet.getRange('B10:B11').setNumberFormat('m/d/yyyy');
}

function setupBudgetSeed_(sheet) {
  const values = LISTS.budgetCategories.map(cat => [cat]);
  sheet.getRange(2, 1, values.length, 1).setValues(values);
}

function applyAllValidations_(ss) {
  const roomRange = ss.getSheetByName('Room List').getRange('A2:A251');
  const vendorRange = ss.getSheetByName('Vendor & Supplier Database').getRange('A2:A251');

  applyValidationMap_(ss, 'Room List', {
    2: 'roomTypes', 5: 'designStatuses', 10: 'priorities'
  });
  applyValidationMap_(ss, 'Master FF&E Schedule', {
    3: 'ffeCategories', 22: 'orderStatuses', 23: 'approvalStatuses', 24: 'deliveryStatuses', 25: 'installStatuses'
  });
  applyRangeValidation_(ss.getSheetByName('Master FF&E Schedule').getRange('B2:B251'), roomRange);
  applyRangeValidation_(ss.getSheetByName('Master FF&E Schedule').getRange('K2:K251'), vendorRange);

  applyValidationMap_(ss, 'Furniture Schedule', { 10: 'approvalStatuses', 11: 'orderStatuses' });
  applyRangeValidation_(ss.getSheetByName('Furniture Schedule').getRange('A2:A251'), roomRange);
  applyRangeValidation_(ss.getSheetByName('Furniture Schedule').getRange('E2:E251'), vendorRange);

  applyValidationMap_(ss, 'Fixtures Schedule', { 2: 'fixtureTypes' });
  applyRangeValidation_(ss.getSheetByName('Fixtures Schedule').getRange('A2:A251'), roomRange);
  applyRangeValidation_(ss.getSheetByName('Fixtures Schedule').getRange('F2:F251'), vendorRange);

  applyValidationMap_(ss, 'Finishes Schedule', { 2: 'surfaces', 3: 'finishTypes', 13: 'yesNo', 14: 'yesNo' });
  applyRangeValidation_(ss.getSheetByName('Finishes Schedule').getRange('A2:A251'), roomRange);
  applyRangeValidation_(ss.getSheetByName('Finishes Schedule').getRange('G2:G251'), vendorRange);

  applyValidationMap_(ss, 'Product Sourcing Tracker', { 9: 'preferredOptions' });
  applyRangeValidation_(ss.getSheetByName('Product Sourcing Tracker').getRange('A2:A251'), roomRange);

  applyValidationMap_(ss, 'Vendor & Supplier Database', { 2: 'vendorTypes' });
  applyValidationMap_(ss, 'Budget Tracker', { 1: 'budgetCategories', 8: 'paymentStatuses' });
  applyValidationMap_(ss, 'Purchase Order Tracker', { 10: 'orderStatuses', 15: 'paymentStatuses' });
  applyRangeValidation_(ss.getSheetByName('Purchase Order Tracker').getRange('B2:B251'), vendorRange);
  applyRangeValidation_(ss.getSheetByName('Purchase Order Tracker').getRange('C2:C251'), roomRange);

  applyValidationMap_(ss, 'Delivery & Installation', { 7: 'deliveryStatuses', 9: 'damageChecks', 14: 'installStatuses' });
  applyRangeValidation_(ss.getSheetByName('Delivery & Installation').getRange('B2:B251'), roomRange);
  applyRangeValidation_(ss.getSheetByName('Delivery & Installation').getRange('C2:C251'), vendorRange);

  applyValidationMap_(ss, 'Client Approval Tracker', { 3: 'approvalTypes', 6: 'approvalStatuses', 9: 'yesNo' });
  applyRangeValidation_(ss.getSheetByName('Client Approval Tracker').getRange('B2:B251'), roomRange);

  applyValidationMap_(ss, 'Project Timeline', { 1: 'phases', 7: 'taskStatuses', 8: 'priorities' });
  applyValidationMap_(ss, 'Task & To-Do List', { 5: 'priorities', 7: 'taskStatuses' });
  applyRangeValidation_(ss.getSheetByName('Task & To-Do List').getRange('B2:B251'), roomRange);

  applyValidationMap_(ss, 'Measurement Log', { 2: 'measurementTypes' });
  applyRangeValidation_(ss.getSheetByName('Measurement Log').getRange('A2:A251'), roomRange);

  applyValidationMap_(ss, 'Contractor Installer Tracker', { 2: 'trades', 8: 'paymentStatuses', 9: 'yesNo', 10: 'yesNo' });
  applyValidationMap_(ss, 'Invoice Payment Tracker', { 8: 'paymentStatuses' });
  applyValidationMap_(ss, 'Project Meeting Notes', { 2: 'meetingTypes' });
  applyValidationMap_(ss, 'Final Punch List', { 4: 'priorities', 6: 'taskStatuses' });
}

function applyValidationMap_(ss, sheetName, map) {
  const sheet = ss.getSheetByName(sheetName);
  Object.keys(map).forEach(col => setListValidation_(ss, sheet.getRange(2, Number(col), TEMPLATE.rows, 1), map[col]));
}

function setListValidation_(ss, range, listKey) {
  const col = Object.keys(LISTS).indexOf(listKey) + 1;
  const source = ss.getSheetByName('Lists').getRange(2, col, LISTS[listKey].length, 1);
  applyRangeValidation_(range, source);
}

function applyRangeValidation_(range, sourceRange) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

function applyAllFormulas_(ss) {
  fillFormula_(ss, 'Room List', 'G2:G251', '=IFERROR(SUMIF(\'Master FF&E Schedule\'!$B:$B,$A2,\'Master FF&E Schedule\'!$P:$P),0)');
  fillFormula_(ss, 'Room List', 'H2:H251', '=IFERROR(SUMIF(\'Master FF&E Schedule\'!$B:$B,$A2,\'Master FF&E Schedule\'!$T:$T),0)');
  fillFormula_(ss, 'Room List', 'I2:I251', '=IF($A2="","",$F2-$H2)');

  fillFormula_(ss, 'Master FF&E Schedule', 'P2:P251', '=IF($F2="","",$F2*$O2)');
  fillFormula_(ss, 'Master FF&E Schedule', 'T2:T251', '=IF($P2="","",$P2+$Q2+$R2-$S2)');

  fillFormula_(ss, 'Fixtures Schedule', 'J2:J251', '=IF($H2="","",$H2*$I2)');
  fillFormula_(ss, 'Finishes Schedule', 'L2:L251', '=IF($J2="","",$J2*$K2)');

  fillFormula_(ss, 'Budget Tracker', 'E2:E251', '=IF($A2="","",$B2-$D2)');
  fillFormula_(ss, 'Budget Tracker', 'G2:G251', '=IF($A2="","",$D2-$F2)');
  addBudgetTotals_(ss.getSheetByName('Budget Tracker'));

  fillFormula_(ss, 'Project Timeline', 'K2:K251', '=IF(OR($D2="",$E2=""),"",$E2-$D2+1)');
  fillFormula_(ss, 'Project Timeline', 'L2:L251', '=IF(AND($E2<TODAY(),$G2<>"Complete"),TODAY()-$E2,0)');
  fillGanttFormulas_(ss.getSheetByName('Project Timeline'));
  fillFormula_(ss, 'Task & To-Do List', 'I2:I251', '=IF(AND($F2<TODAY(),$G2<>"Complete"),TODAY()-$F2,0)');
  fillFormula_(ss, 'Invoice Payment Tracker', 'K2:K251', '=IF($A2="","",IF($H2="Paid",0,$E2))');

  formatCurrencyColumns_(ss);
  formatDateColumns_(ss);
  applyConditionalFormatting_(ss);
}

function fillGanttFormulas_(sheet) {
  const base = 'IF(\'Client & Project Profile\'!$B$10="",TODAY(),\'Client & Project Profile\'!$B$10)';
  for (let col = 13; col <= 24; col++) {
    const letter = columnToLetter_(col);
    const weekOffset = col - 13;
    const formula = `=IF(OR($D2="",$E2=""),"",IF(AND($D2<=${base}+${weekOffset}*7+6,$E2>=${base}+${weekOffset}*7),"■",""))`;
    const range = sheet.getRange(`${letter}2:${letter}251`);
    const formulas = [];
    for (let row = 2; row <= 251; row++) formulas.push([formula.replace(/(\$?[A-Z]{1,3})2(?![0-9])/g, '$1' + row)]);
    range.setFormulas(formulas).setHorizontalAlignment('center').setFontColor(TEMPLATE.colors.teal);
  }
}

function columnToLetter_(column) {
  let temp = '';
  while (column > 0) {
    const remainder = (column - 1) % 26;
    temp = String.fromCharCode(65 + remainder) + temp;
    column = Math.floor((column - remainder) / 26);
  }
  return temp;
}

function fillFormula_(ss, sheetName, a1, formula) {
  const range = ss.getSheetByName(sheetName).getRange(a1);

  // Dynamic-array report formulas should live only in the first cell so the
  // spilled output can populate the report area.
  if (range.getNumColumns() > 1) {
    range.clearContent();
    range.getCell(1, 1).setFormula(formula);
    return;
  }

  if (range.getNumRows() === 1) {
    range.setFormula(formula);
    return;
  }

  // Apps Script applies the exact same A1 formula to every cell in a range.
  // This helper fills row-aware formulas down while preserving absolute columns.
  const startRow = range.getRow();
  const formulas = [];
  for (let i = 0; i < range.getNumRows(); i++) {
    const rowNumber = startRow + i;
    formulas.push([formula.replace(/(\$?[A-Z]{1,3})2(?![0-9])/g, '$1' + rowNumber)]);
  }
  range.setFormulas(formulas);
}

function addBudgetTotals_(sheet) {
  const start = TEMPLATE.rows + 4;
  const totals = [
    ['Total budget', '=SUM(B2:B251)'],
    ['Total estimated', '=SUM(C2:C251)'],
    ['Total actual', '=SUM(D2:D251)'],
    ['Over/under budget', '=B255-B257'],
    ['Total paid', '=SUM(F2:F251)'],
    ['Remaining balance', '=SUM(G2:G251)']
  ];
  sheet.getRange(start, 1, totals.length, 2).setValues(totals);
  sheet.getRange(start, 1, totals.length, 1).setFontWeight('bold').setBackground(TEMPLATE.colors.lightTeal);
  sheet.getRange(start, 2, totals.length, 1).setNumberFormat('$#,##0.00').setFontWeight('bold');
}

function setupDashboard_(ss, sheet) {
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setColumnWidths(1, 12, 120);
  sheet.getRange('A1:L1').merge().setValue('Project Dashboard')
    .setFontSize(22).setFontWeight('bold').setFontColor(TEMPLATE.colors.white).setBackground(TEMPLATE.colors.navy)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 48);

  const kpis = [
    ['Total project budget', '=\'Client & Project Profile\'!B12'],
    ['Total estimated cost', '=SUM(\'Budget Tracker\'!C2:C251)'],
    ['Total actual cost', '=SUM(\'Budget Tracker\'!D2:D251)'],
    ['Remaining budget', '=\'Client & Project Profile\'!B12-SUM(\'Budget Tracker\'!D2:D251)'],
    ['Budget used %', '=IFERROR(SUM(\'Budget Tracker\'!D2:D251)/\'Client & Project Profile\'!B12,0)'],
    ['Number of rooms', '=COUNTA(\'Room List\'!A2:A251)'],
    ['Total FF&E items', '=COUNTA(\'Master FF&E Schedule\'!D2:D251)'],
    ['Items approved', '=COUNTIF(\'Master FF&E Schedule\'!W2:W251,"Approved")'],
    ['Items ordered', '=COUNTIF(\'Master FF&E Schedule\'!V2:V251,"Ordered")+COUNTIF(\'Master FF&E Schedule\'!V2:V251,"Confirmed")+COUNTIF(\'Master FF&E Schedule\'!V2:V251,"In production")+COUNTIF(\'Master FF&E Schedule\'!V2:V251,"Shipped")+COUNTIF(\'Master FF&E Schedule\'!V2:V251,"Delivered")'],
    ['Items delivered', '=COUNTIF(\'Master FF&E Schedule\'!X2:X251,"Delivered")'],
    ['Items installed', '=COUNTIF(\'Master FF&E Schedule\'!Y2:Y251,"Installed")'],
    ['Overdue tasks', '=COUNTIF(\'Task & To-Do List\'!I2:I251,">0")+COUNTIF(\'Project Timeline\'!L2:L251,">0")'],
    ['Pending client approvals', '=COUNTIF(\'Client Approval Tracker\'!F2:F251,"Sent")+COUNTIF(\'Client Approval Tracker\'!F2:F251,"Not sent")+COUNTIF(\'Master FF&E Schedule\'!W2:W251,"Client review")'],
    ['Open vendor orders', '=COUNTIF(\'Purchase Order Tracker\'!J2:J251,"Ordered")+COUNTIF(\'Purchase Order Tracker\'!J2:J251,"Confirmed")+COUNTIF(\'Purchase Order Tracker\'!J2:J251,"In production")+COUNTIF(\'Purchase Order Tracker\'!J2:J251,"Shipped")+COUNTIF(\'Purchase Order Tracker\'!J2:J251,"Delayed")'],
    ['Items not ordered', '=COUNTIF(\'Master FF&E Schedule\'!V2:V251,"Not ordered")+COUNTIFS(\'Master FF&E Schedule\'!D2:D251,"<>",\'Master FF&E Schedule\'!V2:V251,"")'],
    ['Delivered not installed', '=COUNTIFS(\'Master FF&E Schedule\'!X2:X251,"Delivered",\'Master FF&E Schedule\'!Y2:Y251,"<>Installed")']
  ];

  let r = 3;
  kpis.forEach((kpi, i) => {
    const col = i % 4 * 3 + 1;
    const row = r + Math.floor(i / 4) * 3;
    sheet.getRange(row, col, 1, 2).merge().setValue(kpi[0]).setFontWeight('bold').setFontColor(TEMPLATE.colors.white).setBackground(TEMPLATE.colors.blue).setHorizontalAlignment('center');
    sheet.getRange(row + 1, col, 1, 2).merge().setFormula(kpi[1]).setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#F4F8FB');
  });
  sheet.getRangeList(['A4', 'D4', 'G4', 'J4']).setNumberFormat('$#,##0.00');
  sheet.getRange('A7').setNumberFormat('0.0%');

  setupDashboardData_(sheet);
  addCharts_(sheet);
}

function setupDashboardData_(sheet) {
  const c = TEMPLATE.colors;
  sheet.getRange('A16').setValue('Budget by Room').setFontWeight('bold').setBackground(c.lightBlue);
  sheet.getRange('A17:C17').setValues([['Room', 'Estimated', 'Actual']]).setFontWeight('bold');
  sheet.getRange('A18').setFormula('=FILTER({\'Room List\'!A2:A251,\'Room List\'!G2:H251},\'Room List\'!A2:A251<>"")');

  sheet.getRange('E16').setValue('Budget vs Actual').setFontWeight('bold').setBackground(c.lightBlue);
  sheet.getRange('E17:F19').setValues([['Metric', 'Amount'], ['Budget', ''], ['Actual', '']]);
  sheet.getRange('F18').setFormula('=\'Client & Project Profile\'!B12');
  sheet.getRange('F19').setFormula('=SUM(\'Budget Tracker\'!D2:D251)');

  sheet.getRange('H16').setValue('Item Status Breakdown').setFontWeight('bold').setBackground(c.lightBlue);
  sheet.getRange('H17:I27').setValues([['Status', 'Count'], ...LISTS.itemStatuses.map(s => [s, ''])]);
  LISTS.itemStatuses.forEach((s, i) => sheet.getRange(18 + i, 9).setFormula(`=COUNTIF('Master FF&E Schedule'!V2:Y251,"${s}")`));

  sheet.getRange('A35').setValue('Delivery Status Breakdown').setFontWeight('bold').setBackground(c.lightBlue);
  sheet.getRange('A36:B44').setValues([['Status', 'Count'], ...LISTS.deliveryStatuses.map(s => [s, ''])]);
  LISTS.deliveryStatuses.forEach((s, i) => sheet.getRange(37 + i, 2).setFormula(`=COUNTIF('Delivery & Installation'!G2:G251,"${s}")`));

  sheet.getRange('E35').setValue('Category Spend Breakdown').setFontWeight('bold').setBackground(c.lightBlue);
  sheet.getRange('E36:F54').setValues([['Category', 'Final Cost'], ...LISTS.ffeCategories.map(s => [s, ''])]);
  LISTS.ffeCategories.forEach((s, i) => sheet.getRange(37 + i, 6).setFormula(`=SUMIF('Master FF&E Schedule'!C2:C251,"${s}",'Master FF&E Schedule'!T2:T251)`));
  sheet.getRange('B18:C30').setNumberFormat('$#,##0.00');
  sheet.getRange('F18:F19').setNumberFormat('$#,##0.00');
  sheet.getRange('F37:F54').setNumberFormat('$#,##0.00');
}

function addCharts_(sheet) {
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));
  const charts = [
    sheet.newChart().setChartType(Charts.ChartType.COLUMN).addRange(sheet.getRange('A17:C30')).setOption('title', 'Budget by Room').setPosition(16, 1, 0, 0).build(),
    sheet.newChart().setChartType(Charts.ChartType.COLUMN).addRange(sheet.getRange('E17:F19')).setOption('title', 'Budget vs Actual').setPosition(16, 5, 0, 0).build(),
    sheet.newChart().setChartType(Charts.ChartType.PIE).addRange(sheet.getRange('H17:I27')).setOption('title', 'Item Status Breakdown').setPosition(16, 8, 0, 0).build(),
    sheet.newChart().setChartType(Charts.ChartType.PIE).addRange(sheet.getRange('A36:B44')).setOption('title', 'Delivery Status Breakdown').setPosition(35, 1, 0, 0).build(),
    sheet.newChart().setChartType(Charts.ChartType.BAR).addRange(sheet.getRange('E36:F54')).setOption('title', 'Category Spend Breakdown').setPosition(35, 5, 0, 0).build()
  ];
  charts.forEach(chart => sheet.insertChart(chart));
}

function setupPrintableReports_(ss) {
  fillFormula_(ss, 'Client FFE Report', 'A2:I251', '=IFERROR(FILTER({\'Master FF&E Schedule\'!B2:B251,\'Master FF&E Schedule\'!D2:D251,\'Master FF&E Schedule\'!E2:E251,\'Master FF&E Schedule\'!F2:F251,\'Master FF&E Schedule\'!K2:K251,\'Master FF&E Schedule\'!M2:M251,\'Master FF&E Schedule\'!T2:T251,\'Master FF&E Schedule\'!W2:W251,\'Master FF&E Schedule\'!AA2:AA251},\'Master FF&E Schedule\'!D2:D251<>""),"")');
  fillFormula_(ss, 'Room Budget Report', 'A2:G251', '=IFERROR(FILTER({\'Room List\'!A2:A251,\'Room List\'!F2:F251,\'Room List\'!G2:G251,\'Room List\'!H2:H251,\'Room List\'!I2:I251,IFERROR(\'Room List\'!H2:H251/\'Room List\'!F2:F251,0),IF(\'Room List\'!I2:I251>=0,"On budget","Over budget")},\'Room List\'!A2:A251<>""),"")');
  fillFormula_(ss, 'Vendor Order Report', 'A2:J251', '=IFERROR(FILTER({\'Purchase Order Tracker\'!B2:B251,\'Purchase Order Tracker\'!A2:A251,\'Purchase Order Tracker\'!C2:C251,\'Purchase Order Tracker\'!D2:D251,\'Purchase Order Tracker\'!E2:E251,\'Purchase Order Tracker\'!J2:J251,\'Purchase Order Tracker\'!H2:H251,\'Purchase Order Tracker\'!N2:N251,\'Purchase Order Tracker\'!O2:O251,\'Purchase Order Tracker\'!K2:K251},\'Purchase Order Tracker\'!A2:A251<>""),"")');
  fillFormula_(ss, 'Installation Checklist', 'A2:J251', '=IFERROR(FILTER({\'Delivery & Installation\'!B2:B251,\'Delivery & Installation\'!A2:A251,\'Delivery & Installation\'!C2:C251,\'Delivery & Installation\'!G2:G251,\'Delivery & Installation\'!L2:L251,\'Delivery & Installation\'!M2:M251,\'Delivery & Installation\'!N2:N251,\'Delivery & Installation\'!I2:I251,\'Delivery & Installation\'!O2:O251,IF(\'Delivery & Installation\'!N2:N251="Installed","Yes","No")},\'Delivery & Installation\'!A2:A251<>""),"")');

  ['Client FFE Report', 'Room Budget Report', 'Vendor Order Report', 'Installation Checklist', 'Final Punch List'].forEach(name => {
    const sh = ss.getSheetByName(name);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setBackground(TEMPLATE.colors.teal).setFontColor(TEMPLATE.colors.white).setFontWeight('bold');
    sh.setHiddenGridlines(false);
  });
}

function formatCurrencyColumns_(ss) {
  const specs = [
    ['Client & Project Profile', 'B12'], ['Room List', 'F2:I251'], ['Master FF&E Schedule', 'O2:T251'], ['Furniture Schedule', 'I2:I251'],
    ['Fixtures Schedule', 'I2:J251'], ['Finishes Schedule', 'K2:L251'], ['Product Sourcing Tracker', 'D2:D251'], ['Product Sourcing Tracker', 'F2:F251'], ['Product Sourcing Tracker', 'H2:H251'],
    ['Budget Tracker', 'B2:G260'], ['Contractor Installer Tracker', 'E2:E251'], ['Invoice Payment Tracker', 'E2:E251'], ['Invoice Payment Tracker', 'K2:K251'], ['Purchase Order Tracker', 'N2:N251']
  ];
  specs.forEach(([sheet, range]) => ss.getSheetByName(sheet).getRange(range).setNumberFormat('$#,##0.00'));
}

function formatDateColumns_(ss) {
  const specs = [
    ['Furniture Schedule', 'L2:M251'], ['Fixtures Schedule', 'K2:M251'], ['Purchase Order Tracker', 'F2:I251'], ['Delivery & Installation', 'E2:F251'], ['Delivery & Installation', 'L2:L251'],
    ['Client Approval Tracker', 'D2:E251'], ['Client Approval Tracker', 'J2:J251'], ['Project Timeline', 'D2:F251'], ['Task & To-Do List', 'F2:F251'],
    ['Contractor Installer Tracker', 'F2:G251'], ['Invoice Payment Tracker', 'B2:B251'], ['Invoice Payment Tracker', 'F2:G251'], ['Project Meeting Notes', 'A2:A251'], ['Project Meeting Notes', 'G2:G251'], ['Final Punch List', 'E2:E251'], ['Final Punch List', 'H2:H251']
  ];
  specs.forEach(([sheet, range]) => ss.getSheetByName(sheet).getRange(range).setNumberFormat('m/d/yyyy'));
}

function applyConditionalFormatting_(ss) {
  SHEETS.filter(s => s.headers).forEach(config => {
    const sheet = ss.getSheetByName(config.name);
    const range = sheet.getRange(2, 1, TEMPLATE.rows, config.headers.length);
    const rules = [];
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Complete').setBackground(TEMPLATE.colors.green).setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Approved').setBackground(TEMPLATE.colors.green).setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Installed').setBackground(TEMPLATE.colors.green).setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Paid').setBackground(TEMPLATE.colors.green).setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Pending').setBackground(TEMPLATE.colors.yellow).setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('In progress').setBackground(TEMPLATE.colors.yellow).setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Sent').setBackground(TEMPLATE.colors.yellow).setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Delayed').setBackground(TEMPLATE.colors.red).setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Overdue').setBackground(TEMPLATE.colors.red).setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Rejected').setBackground(TEMPLATE.colors.red).setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('Cancelled').setBackground(TEMPLATE.colors.gray).setRanges([range]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('On hold').setBackground(TEMPLATE.colors.gray).setRanges([range]).build());
    sheet.setConditionalFormatRules(rules);
  });

  const budget = ss.getSheetByName('Budget Tracker');
  const existing = budget.getConditionalFormatRules();
  existing.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground(TEMPLATE.colors.red).setRanges([budget.getRange('E2:E251')]).build());
  existing.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0).setBackground(TEMPLATE.colors.green).setRanges([budget.getRange('E2:E251')]).build());
  budget.setConditionalFormatRules(existing);
}

function protectHeader_(sheet) {
  try {
    const protection = sheet.getRange(1, 1, 1, sheet.getLastColumn()).protect();
    protection.setDescription('Header row - protected to prevent accidental edits');
    protection.setWarningOnly(true);
  } catch (err) {
    // Protection can fail in some account/domain contexts; the template still works.
  }
}
