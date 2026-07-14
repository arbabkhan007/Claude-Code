# Interior Design FF&E Project Tracker

Professional multi-tab spreadsheet for interior designers to manage **Furniture, Fixtures & Equipment (FF&E)** projects end-to-end.

**File:** [`Interior_Design_FFE_Project_Tracker.xlsx`](./Interior_Design_FFE_Project_Tracker.xlsx)

**Compatible with:** Microsoft Excel (Windows/Mac), Excel Online, Google Sheets (upload to Drive → Open with Google Sheets), LibreOffice Calc.

---

## What's Included (25 visible tabs + Lists)

| Tab | Purpose |
|-----|---------|
| **01 Welcome** | Quick start, tab guide, color coding, dropdown help, compatibility notes |
| **02 Dashboard** | KPI cards + charts (budget by room, status, category spend, delivery) |
| **03 Client Profile** | Client contact, project scope, budget, contract & billing |
| **04 Room List** | Rooms/areas with budget allocation, spend, remaining, priority |
| **05 Master FF&E** | Full FF&E schedule (core tab) with auto Total/Final cost |
| **06 Furniture** | Furniture-only focused schedule |
| **07 Fixtures** | Lighting, plumbing, hardware, appliances |
| **08 Finishes** | Paint, tile, flooring, wallpaper, countertops, fabrics |
| **09 Sourcing** | Compare up to 3 product options before selecting |
| **10 Vendors** | Supplier database with trade discounts & lead times |
| **11 Budget** | Category budget vs estimated vs actual + summary |
| **12 Purchase Orders** | PO tracker with tracking #, carrier, paid status |
| **13 Delivery & Install** | Delivery logistics, damage checks, installation |
| **14 Client Approvals** | Decision/approval log by type and status |
| **15 Timeline** | Phased tasks + simple 12-week Gantt view |
| **16 To-Do List** | Daily tasks with overdue highlighting |
| **17 Measurements** | Room/window/door/furniture space log |
| **18 Contractors** | Trade coordination, insurance, contracts |
| **19 Invoices** | Client billing & payment tracker |
| **20 Meeting Notes** | Consultations, site visits, vendor meetings |
| **21–25 Reports** | Printable client FF&E, room budget, vendor orders, install checklist, punch list |

A hidden **Lists** sheet stores all dropdown option values for easy customization.

---

## Premium Features

- **Dropdown menus** on room, category, vendor, status, priority, approval, order, delivery, install, payment fields
- **Auto-calculations**
  - Total cost = Quantity × Unit cost
  - Final cost = Total + Tax + Shipping − Discount
  - Room remaining budget, budget variance, paid balance
  - Dashboard KPIs (items approved/ordered/delivered, overdue tasks, open POs, etc.)
- **Conditional formatting**
  - Green = complete / approved / on budget  
  - Yellow = pending / in progress  
  - Red = overdue / over budget / delayed  
  - Gray = cancelled / on hold  
- **Filters + freeze panes** on every major schedule
- **Charts** on the Dashboard (bar + pie)
- **Printable reports** (landscape/portrait ready)
- **Sample project data** (Chen Residence Refresh) so the file demos immediately

---

## Quick Start

1. Open `Interior_Design_FFE_Project_Tracker.xlsx` in Excel or Google Sheets.
2. Go to **03 Client Profile** and replace sample client/project fields (gold cells).
3. Update **04 Room List** with your rooms and budget allocations.
4. Enter items on **05 Master FF&E** (or specialized Furniture / Fixtures / Finishes tabs).
5. Track orders, approvals, deliveries, and budget as the project runs.
6. Use **02 Dashboard** for live overview; print **21–25** for client/site packages.

---

## Regenerate the Workbook

```bash
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python create_ffe_spreadsheet.py
```

This overwrites `Interior_Design_FFE_Project_Tracker.xlsx` with a fresh copy (including sample data).

---

## Color Coding

| Color | Meaning |
|-------|---------|
| Green | Complete / Approved / On budget / Good condition |
| Yellow | Pending / In progress / Awaiting response |
| Red | Overdue / Over budget / Delayed / Damaged |
| Gray | Cancelled / On hold / Not started |
| Navy headers | Column headers |
| Gold cells (Client Profile) | Editable inputs |

---

## Editable / Unlocked

This is the fully unlocked version — all sheets, formulas, and cells can be edited. Keep a clean master copy and **Save As** for each new client project.

---

## Version

Interior Design FF&E Tracker **v1.0**
