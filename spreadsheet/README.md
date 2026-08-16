# Interior Design Project Manager Workbook

This folder contains both the source code and the generated Excel/Google Sheets-compatible workbook.

## Deliverables

- `Interior_Design_Project_Manager.xlsx` — ready-to-use spreadsheet
- `create_interior_design_workbook.py` — reproducible workbook generator

## Generate the workbook

```bash
python -m venv .venv
.venv/bin/pip install XlsxWriter
.venv/bin/python spreadsheet/create_interior_design_workbook.py
```

On Windows, activate the virtual environment or run its Python executable directly.

## Use in Google Sheets

1. Open Google Sheets.
2. Choose **File → Import → Upload**.
3. Upload `Interior_Design_Project_Manager.xlsx`.
4. Choose **Create new spreadsheet**.

The workbook is macro-free, unlocked, and editable. Standard formulas, dropdowns, conditional formatting, filters, charts, freeze panes, and print settings are included. Current versions of Excel and Google Sheets support the `FILTER` formula used by the five live printable reports.

## Included tabs

The workbook includes the requested Welcome guide, Dashboard, Profile, Room List, Master FF&E, Furniture, Fixtures, Finishes, Sourcing, Vendors, Budget, Orders, Delivery/Installation, Client Approvals, Timeline/Gantt, Tasks, Measurements, Contractors, Invoices, Notes, and five printable reports. A hidden `Lists` tab powers consistent dropdown values.

## Capacity and customization

Most schedules include 250 prepared data rows; the Timeline includes 100 tasks. Change the `ROWS` constant in the source and regenerate if more rows are needed. Update the support line on the Welcome tab before distributing the template commercially.
