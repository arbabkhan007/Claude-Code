# Interior Design Google Sheets / Excel Workbook Template

This folder contains `interior_design_project_workbook.gs`, a Google Apps Script that builds a premium interior design project management workbook with:

- Welcome / instructions tab
- Project dashboard with KPI cards and charts
- Client/project profile
- Room tracker
- Master FF&E schedule
- Furniture, fixtures, finishes, sourcing, vendors, budget, purchase order, delivery/install, approval, timeline, tasks, measurements, contractors, invoices, meeting notes tabs
- Printable report tabs
- Dropdown menus, filters, formulas, conditional formatting, and protected header warnings

## Use in Google Sheets

1. Create a new blank Google Sheet.
2. Open **Extensions > Apps Script**.
3. Paste the contents of `interior_design_project_workbook.gs` into `Code.gs`.
4. Save and run `buildInteriorDesignProjectWorkbook()`.
5. Authorize the script.

## Use with Excel

After generating the workbook in Google Sheets, choose **File > Download > Microsoft Excel (.xlsx)**. The generated structure, formulas, filters, dropdowns, and formatting are intended to export cleanly. Google chart placement may require small layout adjustments in Excel.
