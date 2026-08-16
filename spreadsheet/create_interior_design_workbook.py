#!/usr/bin/env python3
"""Create a professional interior-design project workbook.

Run:
    python -m pip install XlsxWriter
    python create_interior_design_workbook.py

The generated .xlsx file can be opened in Excel or imported into Google Sheets.
All input cells are unlocked/editable; no password or macros are used.
"""
from pathlib import Path
from zipfile import ZipFile

try:
    import xlsxwriter
    from xlsxwriter.utility import xl_col_to_name
except ImportError as exc:
    raise SystemExit(
        "Missing dependency: XlsxWriter. Install it with:\n"
        "  python -m pip install -r spreadsheet/requirements.txt"
    ) from exc

FINAL_OUT = Path(__file__).with_name("Interior_Design_Project_Manager.xlsx")
# Build to a temporary file first so a generation error cannot damage an
# existing customer-ready workbook.
OUT = Path(__file__).with_name("Interior_Design_Project_Manager.tmp.xlsx")
if OUT.exists():
    OUT.unlink()
ROWS = 250
NAVY, BLUE, TEAL = "#17324D", "#2F75B5", "#2A9D8F"
GREEN, YELLOW, RED, GRAY = "#C6EFCE", "#FFF2CC", "#FFC7CE", "#D9E1F2"
WHITE, LIGHT, TEXT = "#FFFFFF", "#F4F7FA", "#243447"

wb = xlsxwriter.Workbook(OUT)
wb.set_properties({
    "title": "Interior Design Project Manager",
    "subject": "FF&E, budget, purchasing, delivery, installation, approvals and project management",
    "author": "Interior Design Project Manager",
    "comments": "Excel and Google Sheets compatible; editable and unlocked.",
})

fmt = {
    "title": wb.add_format({"bold": True, "font_size": 20, "font_color": WHITE, "bg_color": NAVY, "align": "left", "valign": "vcenter"}),
    "subtitle": wb.add_format({"font_size": 10, "font_color": "#DCE6F1", "bg_color": NAVY, "align": "left", "valign": "vcenter"}),
    "section": wb.add_format({"bold": True, "font_size": 12, "font_color": WHITE, "bg_color": BLUE, "align": "left", "valign": "vcenter"}),
    "header": wb.add_format({"bold": True, "font_color": WHITE, "bg_color": NAVY, "border": 1, "border_color": WHITE, "align": "center", "valign": "vcenter", "text_wrap": True}),
    "input": wb.add_format({"font_color": TEXT, "bg_color": WHITE, "border": 1, "border_color": "#D5DEE8", "valign": "top"}),
    "input_wrap": wb.add_format({"font_color": TEXT, "bg_color": WHITE, "border": 1, "border_color": "#D5DEE8", "valign": "top", "text_wrap": True}),
    "money": wb.add_format({"font_color": TEXT, "bg_color": WHITE, "border": 1, "border_color": "#D5DEE8", "num_format": '$#,##0.00;[Red]-$#,##0.00'}),
    "percent": wb.add_format({"font_color": TEXT, "bg_color": WHITE, "border": 1, "border_color": "#D5DEE8", "num_format": "0.0%"}),
    "date": wb.add_format({"font_color": TEXT, "bg_color": WHITE, "border": 1, "border_color": "#D5DEE8", "num_format": "mmm d, yyyy"}),
    "int": wb.add_format({"font_color": TEXT, "bg_color": WHITE, "border": 1, "border_color": "#D5DEE8", "num_format": "0"}),
    "decimal": wb.add_format({"font_color": TEXT, "bg_color": WHITE, "border": 1, "border_color": "#D5DEE8", "num_format": "0.00"}),
    "formula": wb.add_format({"font_color": NAVY, "bg_color": "#EAF3F8", "border": 1, "border_color": "#D5DEE8", "num_format": '$#,##0.00;[Red]-$#,##0.00'}),
    "formula_int": wb.add_format({"font_color": NAVY, "bg_color": "#EAF3F8", "border": 1, "border_color": "#D5DEE8", "num_format": "0"}),
    "label": wb.add_format({"bold": True, "font_color": TEXT, "bg_color": "#EAF0F6", "border": 1, "border_color": "#CBD5E1", "valign": "top"}),
    "note": wb.add_format({"font_color": TEXT, "bg_color": LIGHT, "text_wrap": True, "valign": "top"}),
    "small": wb.add_format({"font_size": 9, "font_color": "#5B6770", "text_wrap": True}),
    "link": wb.add_format({"font_color": "#0563C1", "underline": True, "bg_color": WHITE, "border": 1, "border_color": "#D5DEE8"}),
    "card_label": wb.add_format({"bold": True, "font_size": 10, "font_color": "#DCE6F1", "bg_color": NAVY, "align": "center", "valign": "vcenter", "text_wrap": True}),
    "card_money": wb.add_format({"bold": True, "font_size": 18, "font_color": WHITE, "bg_color": BLUE, "align": "center", "valign": "vcenter", "num_format": '$#,##0'}),
    "card_num": wb.add_format({"bold": True, "font_size": 18, "font_color": WHITE, "bg_color": TEAL, "align": "center", "valign": "vcenter", "num_format": "0"}),
    "card_pct": wb.add_format({"bold": True, "font_size": 18, "font_color": WHITE, "bg_color": BLUE, "align": "center", "valign": "vcenter", "num_format": "0.0%"}),
    "report_title": wb.add_format({"bold": True, "font_size": 18, "font_color": WHITE, "bg_color": NAVY, "align": "center", "valign": "vcenter"}),
}

lists = {
    "RoomTypes": ["Living room", "Bedroom", "Kitchen", "Bathroom", "Dining room", "Entryway", "Office", "Nursery", "Outdoor", "Commercial space", "Other"],
    "ProjectTypes": ["Residential", "Commercial", "Office", "Retail", "Hospitality", "Staging"],
    "ProjectStatuses": ["Not started", "Planning", "In progress", "On hold", "Complete", "Cancelled"],
    "BillingMethods": ["Flat fee", "Hourly", "Percentage", "Hybrid"],
    "YesNo": ["Yes", "No"],
    "Priorities": ["Low", "Medium", "High", "Urgent"],
    "DesignStatuses": ["Not started", "Concept", "Sourcing", "Client review", "Approved", "Ordered", "Installed", "Complete", "On hold"],
    "Categories": ["Sofa", "Chair", "Table", "Bed", "Rug", "Lighting", "Artwork", "Mirror", "Window treatment", "Hardware", "Plumbing fixture", "Appliance", "Decor", "Textile", "Cabinetry", "Flooring", "Wallcovering", "Other"],
    "ApprovalStatuses": ["Not sent", "Sent", "Client review", "Approved", "Revisions requested", "Rejected", "On hold"],
    "OrderStatuses": ["Not ordered", "Sourcing", "Option selected", "Client review", "Approved", "Ordered", "Confirmed", "In production", "Shipped", "Delivered", "Delayed", "Cancelled", "Returned", "Backordered"],
    "DeliveryStatuses": ["Not scheduled", "Scheduled", "In transit", "Delivered", "Delayed", "Cancelled", "Returned"],
    "InstallStatuses": ["Not ready", "Ready to install", "Scheduled", "Installed", "Punch list", "Complete", "On hold"],
    "FixtureTypes": ["Lighting", "Faucet", "Sink", "Shower fixture", "Tub", "Toilet", "Cabinet hardware", "Door hardware", "Appliance", "Electrical", "Other"],
    "FinishTypes": ["Paint", "Wallpaper", "Flooring", "Tile", "Countertop", "Backsplash", "Cabinet finish", "Fabric", "Trim", "Stone", "Wood", "Metal", "Other"],
    "Surfaces": ["Wall", "Ceiling", "Floor", "Cabinet", "Countertop", "Shower wall", "Backsplash", "Upholstery", "Drapery", "Other"],
    "VendorTypes": ["Furniture", "Lighting", "Fabric", "Flooring", "Tile", "Paint", "Wallpaper", "Decor", "Hardware", "Contractor", "Installer", "Workroom", "Custom furniture", "Other"],
    "BudgetCategories": ["Furniture", "Lighting", "Rugs", "Decor", "Artwork", "Window treatments", "Flooring", "Paint/wallcovering", "Contractor labor", "Installation", "Shipping", "Taxes", "Design fee", "Contingency"],
    "PaymentStatuses": ["Not paid", "Deposit paid", "Partially paid", "Paid", "Overdue", "Cancelled", "Draft", "Sent"],
    "DamageChecks": ["Not checked", "Good condition", "Damaged", "Missing item", "Wrong item"],
    "ApprovalTypes": ["Furniture", "Finish", "Layout", "Color palette", "Budget", "Purchase", "Vendor", "Design concept"],
    "ApprovalTrackerStatuses": ["Not sent", "Sent", "Approved", "Revisions requested", "Rejected", "On hold"],
    "Phases": ["Consultation", "Measurements", "Concept design", "Sourcing", "Client approval", "Purchasing", "Construction", "Delivery", "Installation", "Styling", "Final reveal", "Punch list"],
    "TaskStatuses": ["Not started", "In progress", "Waiting on client", "Waiting on vendor", "Complete", "Delayed"],
    "MeasurementTypes": ["Room", "Wall", "Window", "Door", "Cabinet", "Furniture space", "Rug area", "Lighting location"],
    "Trades": ["General contractor", "Painter", "Electrician", "Plumber", "Flooring installer", "Tile installer", "Wallpaper installer", "Carpenter", "Window treatment installer", "Furniture installer", "Cleaner", "Other"],
    "InvoiceStatuses": ["Draft", "Sent", "Paid", "Partially paid", "Overdue", "Cancelled"],
    "MeetingTypes": ["Client consultation", "Vendor meeting", "Site visit", "Installation meeting", "Contractor meeting", "Final walkthrough"],
}

# Hidden validation lists.
ls = wb.add_worksheet("Lists")
for c, (name, values) in enumerate(lists.items()):
    ls.write(0, c, name, fmt["header"])
    for r, value in enumerate(values, 1): ls.write(r, c, value)
    col = xl_col_to_name(c)
    wb.define_name(name, f"=Lists!${col}$2:${col}${len(values)+1}")
ls.hide()


def title(ws, text, subtitle="Enter data in white cells; blue cells calculate automatically.", end_col=8):
    ws.merge_range(0, 0, 0, end_col, text, fmt["title"])
    ws.merge_range(1, 0, 1, end_col, subtitle, fmt["subtitle"])
    ws.set_row(0, 30); ws.set_row(1, 20)
    ws.hide_gridlines(2)


def add_status_cf(ws, first_row, last_row, cols):
    green = 'OR(ISNUMBER(SEARCH("approved",A1)),ISNUMBER(SEARCH("complete",A1)),ISNUMBER(SEARCH("delivered",A1)),ISNUMBER(SEARCH("installed",A1)),ISNUMBER(SEARCH("paid",A1)),ISNUMBER(SEARCH("good condition",A1)))'
    yellow = 'OR(ISNUMBER(SEARCH("pending",A1)),ISNUMBER(SEARCH("progress",A1)),ISNUMBER(SEARCH("review",A1)),ISNUMBER(SEARCH("sent",A1)),ISNUMBER(SEARCH("scheduled",A1)),ISNUMBER(SEARCH("waiting",A1)),ISNUMBER(SEARCH("ordered",A1)),ISNUMBER(SEARCH("shipped",A1)))'
    red = 'OR(ISNUMBER(SEARCH("overdue",A1)),ISNUMBER(SEARCH("delayed",A1)),ISNUMBER(SEARCH("rejected",A1)),ISNUMBER(SEARCH("damaged",A1)),ISNUMBER(SEARCH("missing",A1)),ISNUMBER(SEARCH("wrong",A1)),ISNUMBER(SEARCH("backordered",A1)))'
    gray = 'OR(ISNUMBER(SEARCH("cancelled",A1)),ISNUMBER(SEARCH("on hold",A1)),ISNUMBER(SEARCH("returned",A1)))'
    for col in cols:
        rng = f"{xl_col_to_name(col)}{first_row+1}:{xl_col_to_name(col)}{last_row+1}"
        cell = f"{xl_col_to_name(col)}{first_row+1}"
        for formula, color in [(green,GREEN),(yellow,YELLOW),(red,RED),(gray,GRAY)]:
            ws.conditional_format(rng, {"type":"formula", "criteria":f'={formula.replace("A1",cell)}', "format":wb.add_format({"bg_color":color})})


def schedule(name, headers, widths=None, validations=None, formats=None, formulas=None, status_cols=(), rows=ROWS, subtitle=None):
    ws = wb.add_worksheet(name)
    title(ws, name, subtitle or "Use filters in the header row. White cells are inputs; blue cells are formulas.", len(headers)-1)
    for c, h in enumerate(headers): ws.write(2, c, h, fmt["header"])
    ws.set_row(2, 32); ws.freeze_panes(3, 0)
    ws.autofilter(2, 0, rows+2, len(headers)-1)
    ws.set_landscape(); ws.fit_to_pages(1, 0); ws.repeat_rows(0, 2); ws.set_margins(0.25,0.25,0.5,0.5)
    validations = validations or {}; formats = formats or {}; formulas = formulas or {}
    for c, h in enumerate(headers):
        width = (widths[c] if widths and c < len(widths) else max(11, min(22, len(h)+3)))
        ws.set_column(c, c, width)
        cellfmt = fmt[formats.get(c, "input_wrap")]
        for r in range(3, rows+3):
            if c in formulas:
                ws.write_formula(r, c, formulas[c](r+1), cellfmt)
            else: ws.write_blank(r, c, None, cellfmt)
        if c in validations:
            source = validations[c]
            ws.data_validation(3, c, rows+2, c, {"validate":"list", "source":source, "error_title":"Choose a listed value", "error_message":"Please select an option from the dropdown.", "input_title":h, "input_message":"Select from the dropdown list.", "show_input":True})
    add_status_cf(ws, 3, rows+2, status_cols)
    return ws

# Welcome / Instructions
ws = wb.add_worksheet("Welcome")
title(ws, "INTERIOR DESIGN PROJECT MANAGER", "Excel + Google Sheets compatible • Fully editable and unlocked • No macros required", 7)
ws.set_column("A:A", 3); ws.set_column("B:B", 24); ws.set_column("C:H", 18)
ws.merge_range("B4:H4", "QUICK START GUIDE", fmt["section"])
steps = [
    ("Step 1", "Add client/project details", "Complete the Client & Project Profile, including total budget and target dates."),
    ("Step 2", "Add rooms", "List each space in Room List. Room names become your standard reference across schedules."),
    ("Step 3", "Add FF&E items", "Use Master FF&E for furniture, fixtures and equipment. Add product links, quantities and costs."),
    ("Step 4", "Track execution", "Update approvals, orders, delivery, installation, budget and tasks as the project progresses."),
    ("Step 5", "Review the dashboard", "Use Project Dashboard for KPIs, risks and visual project summaries."),
]
for i,(n,h,d) in enumerate(steps,5):
    ws.write(i-1,1,n,fmt["label"]); ws.write(i-1,2,h,fmt["label"]); ws.merge_range(i-1,3,i-1,7,d,fmt["note"]); ws.set_row(i-1,32)
ws.merge_range("B12:H12", "HOW TO USE THIS WORKBOOK", fmt["section"])
notes = [
    ("Data entry", "Enter information in white cells. Blue-tinted cells contain formulas and calculate automatically. Do not paste over formula columns."),
    ("Dropdowns", "Click a dropdown cell and choose a standard value. Dropdowns improve dashboard accuracy. Room and vendor dropdowns use the Room List and Vendor Database."),
    ("Color coding", "Green = complete/approved/on budget • Yellow = pending/in progress • Red = overdue/over budget/delayed • Gray = cancelled/on hold."),
    ("Filters", "Use the arrows in schedule headers to filter by room, vendor, status, approval or date. Clear filters to show all records."),
    ("Compatibility", "Designed for Microsoft Excel (.xlsx) and Google Sheets. To use in Sheets: File → Import → Upload. Some chart styling may vary slightly."),
    ("Editable version", "This workbook is fully editable and unlocked. Duplicate the original before customizing formulas, colors or tab names."),
    ("Support", "Add your business support email, website or product documentation link here before distributing this template."),
]
for i,(h,d) in enumerate(notes,13):
    ws.write(i-1,1,h,fmt["label"]); ws.merge_range(i-1,2,i-1,7,d,fmt["note"]); ws.set_row(i-1,35)
ws.merge_range("B22:H22", "WHAT EACH TAB IS FOR", fmt["section"])
tab_help = [
    "Dashboard — project KPIs and charts", "Profile / Room List — project setup and space budgets", "Master FF&E — central product and cost schedule",
    "Furniture / Fixtures / Finishes — focused specifications", "Sourcing / Vendors — compare products and manage suppliers", "Budget / Orders / Delivery — purchasing controls",
    "Approvals / Timeline / Tasks — project workflow", "Measurements / Contractors / Invoices / Notes — field and admin records", "Reports — printable client, budget, vendor and installation views"
]
for i,t in enumerate(tab_help,23): ws.merge_range(i-1,1,i-1,7,t,fmt["note"])
ws.set_tab_color(TEAL)

# Profile
ws = wb.add_worksheet("Client & Project Profile")
title(ws, "CLIENT & PROJECT PROFILE", "Complete this tab first. Dashboard budget and dates reference these values.", 5)
ws.set_column("A:A", 23); ws.set_column("B:B", 32); ws.set_column("C:C", 3); ws.set_column("D:D", 23); ws.set_column("E:F", 32)
fields = [
    ("Client name",""),("Client email",""),("Phone number",""),("Project name",""),("Project address",""),("Project type","ProjectTypes"),("Design style",""),
    ("Start date","date"),("Target completion date","date"),("Total budget","money"),("Designer name",""),("Project status","ProjectStatuses"),
    ("Contract signed","YesNo"),("Deposit paid","YesNo"),("Design package selected",""),("Billing method","BillingMethods")
]
for r,(label,kind) in enumerate(fields,3):
    ws.write(r,0,label,fmt["label"]); f = fmt[kind] if kind in ("date","money") else fmt["input"]
    ws.merge_range(r,1,r,2,"",f)
    if kind in lists: ws.data_validation(r,1,r,2,{"validate":"list","source":f"={kind}"})
ws.write("D4","Scope summary",fmt["label"]); ws.merge_range("E4:F10","",fmt["input_wrap"])
ws.write("D12","Client notes",fmt["label"]); ws.merge_range("E12:F19","",fmt["input_wrap"])
wb.define_name("ProjectBudget", "='Client & Project Profile'!$B$13")

# Room List
headers = ["Room Name","Room Type","Floor / Level","Square Footage","Design Status","Budget Allocation","Estimated Spend","Actual Spend","Remaining Budget","Priority Level","Notes"]
ws = schedule("Room List",headers,[22,18,14,14,18,16,16,16,17,14,30],{1:"=RoomTypes",4:"=DesignStatuses",9:"=Priorities"},{3:"decimal",5:"money",6:"formula",7:"formula",8:"formula"},{6:lambda r:f'=IF($A{r}="","",SUMIF(\'Master FF&E\'!$B$4:$B$253,$A{r},\'Master FF&E\'!$P$4:$P$253))',7:lambda r:f'=IF($A{r}="","",SUMIF(\'Master FF&E\'!$B$4:$B$253,$A{r},\'Master FF&E\'!$T$4:$T$253))',8:lambda r:f'=IF($A{r}="","",F{r}-H{r})'},[4,9])
wb.define_name("RoomNames", "='Room List'!$A$4:$A$253")
ws.conditional_format("I4:I253",{"type":"cell","criteria":"<","value":0,"format":wb.add_format({"bg_color":RED,"font_color":"#9C0006"})})

# Master FF&E
headers = ["Item ID","Room","Category","Item Name","Description","Quantity","Dimensions","Material","Finish","Color","Vendor","Brand","Product Link","SKU / Model Number","Unit Cost","Total Cost","Tax ($)","Shipping ($)","Discount ($)","Final Cost","Lead Time","Order Status","Approval Status","Delivery Status","Install Status","Notes","Image Link"]
widths=[12,18,18,22,28,10,18,16,16,14,20,16,28,18,14,14,12,14,14,14,14,18,18,18,18,28,28]
vals={1:"=RoomNames",2:"=Categories",10:"=VendorNames",21:"=OrderStatuses",22:"=ApprovalStatuses",23:"=DeliveryStatuses",24:"=InstallStatuses"}
forms={15:lambda r:f'=IF(OR(F{r}="",O{r}=""),"",F{r}*O{r})',19:lambda r:f'=IF(P{r}="","",P{r}+N(Q{r})+N(R{r})-N(S{r}))'}
ws=schedule("Master FF&E",headers,widths,vals,{5:"int",12:"link",14:"money",15:"formula",16:"money",17:"money",18:"money",19:"formula",26:"link"},forms,[21,22,23,24],subtitle="Furniture, Fixtures & Equipment master schedule. Total Cost = Quantity × Unit Cost; Final Cost = Total + Tax + Shipping − Discount.")
ws.set_tab_color(BLUE)

# Focused schedules
schedule("Furniture Schedule",["Room","Furniture Item","Dimensions","Quantity","Vendor","Product Link","Fabric / Material","Finish","Cost","Approval Status","Order Status","Delivery Date","Install Date"],[18,24,18,10,20,28,20,18,14,18,18,16,16],{0:"=RoomNames",4:"=VendorNames",9:"=ApprovalStatuses",10:"=OrderStatuses"},{3:"int",5:"link",8:"money",11:"date",12:"date"},status_cols=[9,10])
schedule("Fixtures Schedule",["Room","Fixture Type","Item Name","Finish","Size","Vendor","Model / SKU","Quantity","Unit Cost","Total Cost","Required Rough-in Date","Order Deadline","Delivery Date","Installer Notes"],[18,18,22,16,14,20,18,10,14,14,18,16,16,30],{0:"=RoomNames",1:"=FixtureTypes",5:"=VendorNames"},{7:"int",8:"money",9:"formula",10:"date",11:"date",12:"date"},{9:lambda r:f'=IF(OR(H{r}="",I{r}=""),"",H{r}*I{r})'})
schedule("Finishes Schedule",["Room","Surface / Application","Finish Type","Product Name","Color / Finish","Material","Vendor","Product Link","SKU","Quantity Needed","Unit Cost","Total Cost","Sample Ordered","Sample Approved","Installer Notes"],[18,20,18,22,18,18,20,28,16,14,14,14,16,16,30],{0:"=RoomNames",1:"=Surfaces",2:"=FinishTypes",6:"=VendorNames",12:"=YesNo",13:"=YesNo"},{7:"link",9:"decimal",10:"money",11:"formula"},{11:lambda r:f'=IF(OR(J{r}="",K{r}=""),"",J{r}*K{r})'},[12,13])
schedule("Product Sourcing",["Room","Item Needed","Option 1 Link","Option 1 Price","Option 2 Link","Option 2 Price","Option 3 Link","Option 3 Price","Preferred Option","Pros","Cons","Client Feedback","Final Selection"],[18,22,28,14,28,14,28,14,18,24,24,30,24],{0:"=RoomNames"},{2:"link",3:"money",4:"link",5:"money",6:"link",7:"money"})

# Vendor database comes after references; named range can refer ahead.
headers=["Vendor Name","Vendor Type","Contact Person","Email","Phone","Website","Account Number","Trade Discount %","Lead Time","Shipping Terms","Return Policy","Notes"]
ws=schedule("Vendor Database",headers,[24,18,20,24,16,26,18,18,14,24,28,30],{1:"=VendorTypes"},{5:"link",7:"percent"})
wb.define_name("VendorNames", "='Vendor Database'!$A$4:$A$253")

# Budget Tracker with category starter rows
headers=["Budget Category","Budgeted Amount","Estimated Amount","Actual Amount","Difference","Paid Amount","Balance Due","Payment Status","Notes"]
ws=schedule("Budget Tracker",headers,[24,18,18,18,18,16,16,18,30],{0:"=BudgetCategories",7:"=PaymentStatuses"},{1:"money",2:"money",3:"money",4:"formula",5:"money",6:"formula"},{4:lambda r:f'=IF(B{r}="","",B{r}-D{r})',6:lambda r:f'=IF(D{r}="","",D{r}-N(F{r}))'},[7])
for i,cat in enumerate(lists["BudgetCategories"],3): ws.write(i,0,cat,fmt["input"])
ws.write(ROWS+4,0,"TOTAL",fmt["header"])
for c in range(1,7): ws.write_formula(ROWS+4,c,f'=SUM({xl_col_to_name(c)}4:{xl_col_to_name(c)}253)',fmt["formula"])
ws.conditional_format("E4:E253",{"type":"cell","criteria":"<","value":0,"format":wb.add_format({"bg_color":RED})})

# Orders, delivery, approval
schedule("Order Tracker",["PO Number","Vendor","Room","Item","Quantity","Order Date","Expected Ship Date","Expected Delivery Date","Actual Delivery Date","Order Status","Tracking Number","Carrier","Invoice Number","Amount","Paid Status","Notes"],[14,20,18,24,10,15,17,19,18,18,20,16,16,14,16,30],{1:"=VendorNames",2:"=RoomNames",9:"=OrderStatuses",14:"=PaymentStatuses"},{4:"int",5:"date",6:"date",7:"date",8:"date",13:"money"},status_cols=[9,14])
schedule("Delivery & Installation",["Item","Room","Vendor","Delivery Address","Expected Delivery Date","Actual Delivery Date","Delivery Status","Received By","Damage Check","Damage Notes","Storage Location","Install Date","Installer","Installation Status","Punch List Notes"],[24,18,20,28,19,18,18,18,18,26,20,16,20,20,30],{1:"=RoomNames",2:"=VendorNames",6:"=DeliveryStatuses",8:"=DamageChecks",13:"=InstallStatuses"},{4:"date",5:"date",11:"date"},status_cols=[6,8,13])
schedule("Client Approvals",["Item / Design Decision","Room","Approval Type","Sent to Client Date","Client Response Date","Status","Approved Option","Client Comments","Revision Needed","Final Approval Date"],[26,18,18,19,20,20,22,30,18,18],{1:"=RoomNames",2:"=ApprovalTypes",5:"=ApprovalTrackerStatuses",8:"=YesNo"},{3:"date",4:"date",9:"date"},status_cols=[5,8])

# Timeline + simple Gantt (26 weeks)
base_headers=["Phase","Task","Assigned To","Start Date","Due Date","Completion Date","Status","Priority","Dependencies","Notes"]
week_headers=[f"Wk {i}" for i in range(1,27)]
ws=schedule("Project Timeline",base_headers+week_headers,[18,26,18,15,15,17,18,12,18,28]+[5]*26,{0:"=Phases",6:"=TaskStatuses",7:"=Priorities"},{3:"date",4:"date",5:"date"},status_cols=[6,7],rows=100,subtitle="Enter task dates, then set the Gantt start date in K2. Weekly bars shade automatically.")
ws.write("K2","Gantt start:",fmt["label"]); ws.write_formula("L2","=TODAY()-WEEKDAY(TODAY(),2)+1",fmt["date"])
for c in range(10,36):
    week=c-9
    ws.write(2,c, f"Wk {week}",fmt["header"])
    rng=f"{xl_col_to_name(c)}4:{xl_col_to_name(c)}103"
    ws.conditional_format(rng,{"type":"formula","criteria":f'=AND({xl_col_to_name(c)}$3>=$D4,{xl_col_to_name(c)}$3<=$E4)',"format":wb.add_format({"bg_color":TEAL})})
    ws.write_formula(2,c,f'=$L$2+{week-1}*7',wb.add_format({"bold":True,"font_color":WHITE,"bg_color":NAVY,"num_format":"mmm d","align":"center","text_wrap":True,"border":1}))

schedule("Task List",["Task","Room","Category","Assigned To","Priority","Due Date","Status","Notes"],[28,18,18,18,14,16,20,32],{1:"=RoomNames",2:"=Categories",4:"=Priorities",6:"=TaskStatuses"},{5:"date"},status_cols=[4,6])
schedule("Measurement Log",["Room","Measurement Type","Width","Height","Depth","Length","Ceiling Height","Window Size","Door Size","Wall Length","Notes"],[18,20,12,12,12,12,16,18,18,14,30],{0:"=RoomNames",1:"=MeasurementTypes"},{2:"decimal",3:"decimal",4:"decimal",5:"decimal",6:"decimal",9:"decimal"})
schedule("Contractor Tracker",["Contractor / Vendor Name","Trade","Contact","Scope","Quote Amount","Start Date","End Date","Payment Status","Insurance Received","Contract Signed","Notes"],[25,24,20,30,16,15,15,18,18,16,30],{1:"=Trades",7:"=PaymentStatuses",8:"=YesNo",9:"=YesNo"},{4:"money",5:"date",6:"date"},status_cols=[7,8,9])
schedule("Invoice & Payment",["Invoice Number","Invoice Date","Client","Description","Amount","Due Date","Paid Date","Payment Status","Payment Method","Notes"],[18,16,22,28,16,16,16,18,18,30],{7:"=InvoiceStatuses"},{1:"date",4:"money",5:"date",6:"date"},status_cols=[7])
schedule("Project Notes",["Date","Meeting Type","Attendees","Discussion Points","Decisions Made","Action Items","Follow-up Date","Notes"],[16,22,24,34,34,32,18,34],{1:"=MeetingTypes"},{0:"date",6:"date"})

# Dashboard
ws=wb.add_worksheet("Project Dashboard"); ws.set_tab_color(TEAL)
title(ws,"PROJECT DASHBOARD","Live overview • Complete the Profile, Room List, FF&E, Orders, Approvals and Tasks tabs to populate metrics.",15)
ws.hide_gridlines(2); ws.set_column("A:P",12); ws.freeze_panes(2,0)
cards=[
("Total Project Budget",'=ProjectBudget',"money"),("Total Estimated Cost",'=SUM(\'Master FF&E\'!$P$4:$P$253)',"money"),("Total Actual Cost",'=SUM(\'Master FF&E\'!$T$4:$T$253)',"money"),("Remaining Budget",'=ProjectBudget-SUM(\'Master FF&E\'!$T$4:$T$253)',"money"),
("Budget Used %",'=IFERROR(SUM(\'Master FF&E\'!$T$4:$T$253)/ProjectBudget,0)',"pct"),("Number of Rooms",'=COUNTIF(\'Room List\'!$A$4:$A$253,"<>")',"num"),("Total FF&E Items",'=COUNTIF(\'Master FF&E\'!$A$4:$A$253,"<>")',"num"),("Items Approved",'=COUNTIF(\'Master FF&E\'!$W$4:$W$253,"Approved")',"num"),
("Items Ordered",'=COUNTIF(\'Master FF&E\'!$V$4:$V$253,"Ordered")+COUNTIF(\'Master FF&E\'!$V$4:$V$253,"Confirmed")',"num"),("Items Delivered",'=COUNTIF(\'Master FF&E\'!$X$4:$X$253,"Delivered")',"num"),("Items Installed",'=COUNTIF(\'Master FF&E\'!$Y$4:$Y$253,"Installed")+COUNTIF(\'Master FF&E\'!$Y$4:$Y$253,"Complete")',"num"),("Overdue Tasks",'=COUNTIFS(\'Task List\'!$F$4:$F$253,"<"&TODAY(),\'Task List\'!$F$4:$F$253,"<>",\'Task List\'!$G$4:$G$253,"<>Complete")',"num"),
("Pending Client Approvals",'=COUNTIF(\'Client Approvals\'!$F$4:$F$253,"Sent")+COUNTIF(\'Client Approvals\'!$F$4:$F$253,"Revisions requested")',"num"),("Open Vendor Orders",'=COUNTIFS(\'Order Tracker\'!$A$4:$A$253,"<>",\'Order Tracker\'!$J$4:$J$253,"<>Delivered",\'Order Tracker\'!$J$4:$J$253,"<>Cancelled",\'Order Tracker\'!$J$4:$J$253,"<>Returned")',"num"),
("Items Not Ordered",'=COUNTIFS(\'Master FF&E\'!$A$4:$A$253,"<>",\'Master FF&E\'!$V$4:$V$253,"<>Ordered",\'Master FF&E\'!$V$4:$V$253,"<>Confirmed",\'Master FF&E\'!$V$4:$V$253,"<>In production",\'Master FF&E\'!$V$4:$V$253,"<>Shipped",\'Master FF&E\'!$V$4:$V$253,"<>Delivered",\'Master FF&E\'!$V$4:$V$253,"<>Cancelled")',"num"),("Delivered, Not Installed",'=COUNTIFS(\'Master FF&E\'!$X$4:$X$253,"Delivered",\'Master FF&E\'!$Y$4:$Y$253,"<>Installed",\'Master FF&E\'!$Y$4:$Y$253,"<>Complete")',"num")]
for i,(label,formula,kind) in enumerate(cards):
    row=3+(i//4)*3; col=(i%4)*4
    ws.merge_range(row,col,row,col+2,label,fmt["card_label"])
    ws.merge_range(row+1,col,row+1,col+2,"",fmt["card_money"] if kind=="money" else fmt["card_pct"] if kind=="pct" else fmt["card_num"])
    ws.write_formula(row+1,col,formula,fmt["card_money"] if kind=="money" else fmt["card_pct"] if kind=="pct" else fmt["card_num"])
    ws.set_row(row,24); ws.set_row(row+1,34)
# Chart data at rows 18 onward / hide columns Q onward
start=18
ws.write_row(start,16,["Room","Budget","Actual"],fmt["header"])
for i in range(15):
    r=start+1+i; ws.write_formula(r,16,f"='Room List'!A{4+i}"); ws.write_formula(r,17,f"='Room List'!F{4+i}"); ws.write_formula(r,18,f"='Room List'!H{4+i}")
ws.write_row(start,20,["Budget Category","Budgeted","Actual"],fmt["header"])
for i,cat in enumerate(lists["BudgetCategories"]):
    r=start+1+i; ws.write(r,20,cat); ws.write_formula(r,21,f'=SUMIF(\'Budget Tracker\'!$A$4:$A$253,U{r+1},\'Budget Tracker\'!$B$4:$B$253)'); ws.write_formula(r,22,f'=SUMIF(\'Budget Tracker\'!$A$4:$A$253,U{r+1},\'Budget Tracker\'!$D$4:$D$253)')
statuses=["Sourcing","Option selected","Client review","Approved","Ordered","Shipped","Delivered","Installed","Cancelled","Backordered"]
ws.write_row(start,24,["Item Status","Count"],fmt["header"])
for i,s in enumerate(statuses): r=start+1+i; ws.write(r,24,s); ws.write_formula(r,25,f'=COUNTIF(\'Master FF&E\'!$V$4:$V$253,Y{r+1})')
delivery=lists["DeliveryStatuses"]
ws.write_row(start,27,["Delivery Status","Count"],fmt["header"])
for i,s in enumerate(delivery): r=start+1+i; ws.write(r,27,s); ws.write_formula(r,28,f'=COUNTIF(\'Master FF&E\'!$X$4:$X$253,AB{r+1})')
ws.write_row(start,30,["Category","Spend"],fmt["header"])
for i,s in enumerate(lists["Categories"]): r=start+1+i; ws.write(r,30,s); ws.write_formula(r,31,f'=SUMIF(\'Master FF&E\'!$C$4:$C$253,AE{r+1},\'Master FF&E\'!$T$4:$T$253)')

def chart(kind,title,cat_col,val_cols,pos):
    ch=wb.add_chart({"type":kind})
    for col,name in val_cols:
        ch.add_series({"name":name,"categories":["Project Dashboard",start+1,cat_col,start+15,cat_col],"values":["Project Dashboard",start+1,col,start+15,col]})
    ch.set_title({"name":title}); ch.set_legend({"position":"bottom"}); ch.set_style(10); ch.set_size({"width":500,"height":280}); ch.set_chartarea({"border":{"none":True}}); ws.insert_chart(pos,ch)
chart("column","Budget by Room",16,[(17,"Budget"),(18,"Actual")],"A16")
chart("column","Budget vs Actual",20,[(21,"Budgeted"),(22,"Actual")],"I16")
chart("doughnut","Item Status Breakdown",24,[(25,"Items")],"A31")
# Correct doughnut category row range still 15; blanks harmless.
chart("doughnut","Delivery Status Breakdown",27,[(28,"Items")],"I31")
chart("bar","Category Spend Breakdown",30,[(31,"Spend")],"A46")
ws.set_column(16,31,None,None,{"hidden":True})

# Printable reports with formulas linked to live schedules (Excel 365 / Google Sheets FILTER).
def report(name,title_text,headers,source_sheet,source_range,filter_col,filter_note):
    ws=wb.add_worksheet(name); ws.hide_gridlines(2); ws.set_landscape(); ws.fit_to_pages(1,0); ws.set_margins(.3,.3,.5,.5)
    ws.merge_range(0,0,0,len(headers)-1,title_text,fmt["report_title"]); ws.set_row(0,30)
    ws.merge_range(1,0,1,len(headers)-1,filter_note,fmt["small"])
    for c,h in enumerate(headers): ws.write(3,c,h,fmt["header"]); ws.set_column(c,c,max(14,min(28,len(h)+4)))
    # FILTER is supported by current Excel and Google Sheets; IFERROR keeps empty reports clean.
    ws.write_dynamic_array_formula(4,0,4,len(headers)-1,f'=IFERROR(FILTER(\'{source_sheet}\'!{source_range},\'{source_sheet}\'!{filter_col}<>""),"")',fmt["input_wrap"])
    ws.repeat_rows(0,3); return ws
report("Client FF&E Report","CLIENT FF&E SPECIFICATION REPORT",["Item ID","Room","Category","Item","Description","Qty","Dimensions","Material","Finish","Color","Vendor","Brand","Product Link","SKU","Unit Cost","Total","Tax","Shipping","Discount","Final Cost","Lead Time","Order","Approval","Delivery","Install","Notes","Image"],"Master FF&E","A4:AA253","A4:A253","Client-facing live report. Use print settings or filter the Master FF&E tab before sharing.")
report("Room Budget Report","ROOM-BY-ROOM BUDGET REPORT",["Room","Type","Floor","Sq Ft","Status","Budget","Estimated","Actual","Remaining","Priority","Notes"],"Room List","A4:K253","A4:A253","Live summary of room allocations and spend.")
report("Vendor Order Report","VENDOR ORDER REPORT",["PO","Vendor","Room","Item","Qty","Order Date","Expected Ship","Expected Delivery","Actual Delivery","Status","Tracking","Carrier","Invoice","Amount","Paid","Notes"],"Order Tracker","A4:P253","A4:A253","Live purchasing report. Filter the Order Tracker by vendor or status for targeted reports.")
report("Installation Checklist","DELIVERY & INSTALLATION CHECKLIST",["Item","Room","Vendor","Address","Expected","Actual","Delivery Status","Received By","Damage Check","Damage Notes","Storage","Install Date","Installer","Install Status","Punch List"],"Delivery & Installation","A4:O253","A4:A253","Print-ready field checklist for receiving and installation days.")
report("Final Punch List","FINAL PUNCH LIST",["Item","Room","Vendor","Address","Expected","Actual","Delivery Status","Received By","Damage Check","Damage Notes","Storage","Install Date","Installer","Install Status","Punch List"],"Delivery & Installation","A4:O253","O4:O253","Shows records with punch-list notes. Keep this report internal until ready to share.")

# Put the beginner guide first, the Dashboard second, and the hidden validation
# sheet last. XlsxWriter keeps worksheets in creation order unless explicitly
# reordered before the package is written.
sheet_order = [
    "Welcome", "Project Dashboard", "Client & Project Profile", "Room List",
    "Master FF&E", "Furniture Schedule", "Fixtures Schedule", "Finishes Schedule",
    "Product Sourcing", "Vendor Database", "Budget Tracker", "Order Tracker",
    "Delivery & Installation", "Client Approvals", "Project Timeline", "Task List",
    "Measurement Log", "Contractor Tracker", "Invoice & Payment", "Project Notes",
    "Client FF&E Report", "Room Budget Report", "Vendor Order Report",
    "Installation Checklist", "Final Punch List", "Lists",
]
by_name = {sheet.get_name(): sheet for sheet in wb.worksheets_objs}
wb.worksheets_objs = [by_name[name] for name in sheet_order]
for index, sheet in enumerate(wb.worksheets_objs):
    sheet.index = index

# Workbook-wide active sheet and formula recalculation.
ws_welcome=wb.get_worksheet_by_name("Welcome"); ws_welcome.activate(); ws_welcome.select()
wb.set_calc_mode("auto")
wb.close()

# Confirm that the generated XLSX package is structurally readable before it
# replaces the previous deliverable.
with ZipFile(OUT) as archive:
    damaged_member = archive.testzip()
    if damaged_member:
        OUT.unlink(missing_ok=True)
        raise SystemExit(f"Workbook validation failed at: {damaged_member}")

try:
    OUT.replace(FINAL_OUT)
except PermissionError as exc:
    raise SystemExit(
        f"Cannot replace {FINAL_OUT.name}. Close it in Excel/another app and rerun.\n"
        f"The validated temporary workbook is available at: {OUT}"
    ) from exc

print(f"Created and validated: {FINAL_OUT}")
