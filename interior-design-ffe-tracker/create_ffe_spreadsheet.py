#!/usr/bin/env python3
"""
Interior Design FF&E Project Tracker
Excel + Google Sheets compatible workbook generator.

Creates a professional multi-tab spreadsheet for interior designers to manage
projects, FF&E schedules, budgets, vendors, approvals, deliveries, and more.
"""

from openpyxl import Workbook
from openpyxl.styles import (
    Font, Fill, PatternFill, Border, Side, Alignment, NamedStyle, Protection
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.comments import Comment
from datetime import datetime, timedelta
import os

# ---------------------------------------------------------------------------
# Theme / styles
# ---------------------------------------------------------------------------
NAVY = "1B2A4A"
TEAL = "2A9D8F"
GOLD = "E9C46A"
CORAL = "E76F51"
CREAM = "F8F4EF"
WHITE = "FFFFFF"
LIGHT_GRAY = "F2F2F2"
MED_GRAY = "D9D9D9"
DARK = "1A1A1A"
GREEN = "2E7D32"
GREEN_BG = "C8E6C9"
YELLOW = "F9A825"
YELLOW_BG = "FFF9C4"
RED = "C62828"
RED_BG = "FFCDD2"
GRAY_BG = "E0E0E0"
BLUE_BG = "BBDEFB"
SOFT_TEAL = "E0F2F1"
SOFT_NAVY = "E8EEF7"
SOFT_GOLD = "FFF8E1"

thin = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)
thick_bottom = Border(bottom=Side(style="medium", color=NAVY))


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def font(bold=False, size=11, color=DARK, name="Calibri"):
    return Font(name=name, bold=bold, size=size, color=color)


def header_fill():
    return fill(NAVY)


def header_font():
    return font(bold=True, size=11, color=WHITE)


def title_font():
    return font(bold=True, size=20, color=NAVY)


def section_font():
    return font(bold=True, size=14, color=NAVY)


def kpi_label_font():
    return font(bold=True, size=9, color="666666")


def kpi_value_font(color=NAVY):
    return font(bold=True, size=16, color=color)


def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)


def left_align():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)


def money_format():
    return '"$"#,##0.00'


def pct_format():
    return '0.0%'


def date_format():
    return "YYYY-MM-DD"


# ---------------------------------------------------------------------------
# Dropdown option lists (shared across sheets)
# ---------------------------------------------------------------------------
ROOM_TYPES = [
    "Living Room", "Bedroom", "Kitchen", "Bathroom", "Dining Room",
    "Entryway", "Office", "Nursery", "Outdoor", "Commercial Space", "Other",
]
FFE_CATEGORIES = [
    "Sofa", "Chair", "Table", "Bed", "Rug", "Lighting", "Artwork", "Mirror",
    "Window Treatment", "Hardware", "Plumbing Fixture", "Appliance", "Decor",
    "Textile", "Cabinetry", "Flooring", "Wallcovering", "Other",
]
ORDER_STATUSES = [
    "Sourcing", "Option Selected", "Client Review", "Approved", "Ordered",
    "Shipped", "Delivered", "Installed", "Cancelled", "Backordered",
    "Not Ordered", "Confirmed", "In Production", "Delayed", "Returned",
]
APPROVAL_STATUSES = [
    "Not Sent", "Sent", "Approved", "Revisions Requested", "Rejected", "On Hold",
]
DELIVERY_STATUSES = [
    "Not Ordered", "Ordered", "Shipped", "In Transit", "Delivered",
    "Delayed", "Partial", "Cancelled",
]
INSTALL_STATUSES = [
    "Not Scheduled", "Scheduled", "In Progress", "Installed",
    "Punch List", "Complete", "On Hold",
]
PRIORITY_LEVELS = ["Low", "Medium", "High", "Urgent"]
PAYMENT_STATUSES = [
    "Draft", "Sent", "Paid", "Partially Paid", "Overdue", "Cancelled", "Unpaid",
]
PROJECT_TYPES = [
    "Residential", "Commercial", "Office", "Retail", "Hospitality", "Staging",
]
PROJECT_STATUSES = [
    "Inquiry", "Proposal", "Active", "On Hold", "Completed", "Cancelled",
]
BILLING_METHODS = ["Flat Fee", "Hourly", "Percentage", "Hybrid"]
YES_NO = ["Yes", "No"]
DESIGN_STATUSES = [
    "Not Started", "Concept", "In Progress", "Client Review",
    "Approved", "Complete", "On Hold",
]
FIXTURE_TYPES = [
    "Lighting", "Faucet", "Sink", "Shower Fixture", "Tub", "Toilet",
    "Cabinet Hardware", "Door Hardware", "Appliance", "Electrical", "Other",
]
FINISH_TYPES = [
    "Paint", "Wallpaper", "Flooring", "Tile", "Countertop", "Backsplash",
    "Cabinet Finish", "Fabric", "Trim", "Stone", "Wood", "Metal", "Other",
]
SURFACES = [
    "Wall", "Ceiling", "Floor", "Cabinet", "Countertop", "Shower Wall",
    "Backsplash", "Upholstery", "Drapery", "Other",
]
VENDOR_TYPES = [
    "Furniture", "Lighting", "Fabric", "Flooring", "Tile", "Paint",
    "Wallpaper", "Decor", "Hardware", "Contractor", "Installer",
    "Workroom", "Custom Furniture", "Other",
]
BUDGET_CATEGORIES = [
    "Furniture", "Lighting", "Rugs", "Decor", "Artwork", "Window Treatments",
    "Flooring", "Paint/Wallcovering", "Contractor Labor", "Installation",
    "Shipping", "Taxes", "Design Fee", "Contingency",
]
DAMAGE_CHECKS = [
    "Not Checked", "Good Condition", "Damaged", "Missing Item", "Wrong Item",
]
APPROVAL_TYPES = [
    "Furniture", "Finish", "Layout", "Color Palette", "Budget",
    "Purchase", "Vendor", "Design Concept",
]
PROJECT_PHASES = [
    "Consultation", "Measurements", "Concept Design", "Sourcing",
    "Client Approval", "Purchasing", "Construction", "Delivery",
    "Installation", "Styling", "Final Reveal", "Punch List",
]
TASK_STATUSES = [
    "Not Started", "In Progress", "Waiting on Client", "Waiting on Vendor",
    "Complete", "Delayed",
]
TASK_CATEGORIES = [
    "Design", "Sourcing", "Purchasing", "Client", "Vendor", "Site",
    "Installation", "Admin", "Other",
]
MEASUREMENT_TYPES = [
    "Room", "Wall", "Window", "Door", "Cabinet", "Furniture Space",
    "Rug Area", "Lighting Location",
]
TRADES = [
    "General Contractor", "Painter", "Electrician", "Plumber",
    "Flooring Installer", "Tile Installer", "Wallpaper Installer",
    "Carpenter", "Window Treatment Installer", "Furniture Installer",
    "Cleaner", "Other",
]
MEETING_TYPES = [
    "Client Consultation", "Vendor Meeting", "Site Visit",
    "Installation Meeting", "Contractor Meeting", "Final Walkthrough",
]
PO_STATUSES = [
    "Not Ordered", "Ordered", "Confirmed", "In Production", "Shipped",
    "Delivered", "Delayed", "Cancelled", "Returned",
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def style_header_row(ws, row, start_col, end_col, height=30):
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = header_fill()
        cell.font = header_font()
        cell.alignment = center()
        cell.border = thin
    ws.row_dimensions[row].height = height


def style_data_cell(cell, money=False, date=False, pct=False, center_text=False):
    cell.border = thin
    cell.alignment = center() if center_text else left_align()
    if money:
        cell.number_format = money_format()
    if date:
        cell.number_format = date_format()
    if pct:
        cell.number_format = pct_format()


def set_col_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def add_dropdown(ws, options, cell_range, allow_blank=True):
    # Escape commas in options by using a formula list sheet approach via comma join
    # openpyxl data validation with quoted list
    safe = [str(o).replace(",", " ") for o in options]
    formula = '"' + ",".join(safe) + '"'
    dv = DataValidation(
        type="list",
        formula1=formula,
        allow_blank=allow_blank,
        showDropDown=False,  # False = show dropdown arrow
        showErrorMessage=True,
        errorTitle="Invalid Entry",
        error="Please select a value from the dropdown list.",
    )
    dv.add(cell_range)
    ws.add_data_validation(dv)
    return dv


def add_title_banner(ws, title, subtitle=None, merge_cols=8):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=merge_cols)
    cell = ws.cell(row=1, column=1, value=title)
    cell.font = title_font()
    cell.fill = fill(SOFT_NAVY)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 36
    for col in range(1, merge_cols + 1):
        ws.cell(row=1, column=col).fill = fill(SOFT_NAVY)

    if subtitle:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=merge_cols)
        cell = ws.cell(row=2, column=1, value=subtitle)
        cell.font = font(size=10, color="666666")
        cell.alignment = left_align()
        ws.row_dimensions[2].height = 20


def apply_status_conditional_formatting(ws, col_letter, start_row, end_row):
    """Green / yellow / red / gray for common status words."""
    rng = f"{col_letter}{start_row}:{col_letter}{end_row}"
    green_words = [
        "Approved", "Installed", "Delivered", "Complete", "Completed",
        "Paid", "Good Condition", "Yes",
    ]
    yellow_words = [
        "Pending", "In Progress", "Sent", "Client Review", "Ordered",
        "Shipped", "Scheduled", "Partially Paid", "Waiting on Client",
        "Waiting on Vendor", "Sourcing", "Option Selected", "Confirmed",
        "In Production", "In Transit", "Not Checked",
    ]
    red_words = [
        "Overdue", "Delayed", "Rejected", "Damaged", "Missing Item",
        "Wrong Item", "Backordered", "Urgent",
    ]
    gray_words = ["Cancelled", "On Hold", "Not Started", "Not Sent", "Draft", "Not Ordered"]

    for word in green_words:
        ws.conditional_formatting.add(
            rng,
            FormulaRule(
                formula=[f'ISNUMBER(SEARCH("{word}",{col_letter}{start_row}))'],
                fill=fill(GREEN_BG),
                font=font(color=GREEN, bold=True, size=10),
            ),
        )
    for word in yellow_words:
        ws.conditional_formatting.add(
            rng,
            FormulaRule(
                formula=[f'ISNUMBER(SEARCH("{word}",{col_letter}{start_row}))'],
                fill=fill(YELLOW_BG),
                font=font(color="F57F17", bold=True, size=10),
            ),
        )
    for word in red_words:
        ws.conditional_formatting.add(
            rng,
            FormulaRule(
                formula=[f'ISNUMBER(SEARCH("{word}",{col_letter}{start_row}))'],
                fill=fill(RED_BG),
                font=font(color=RED, bold=True, size=10),
            ),
        )
    for word in gray_words:
        ws.conditional_formatting.add(
            rng,
            FormulaRule(
                formula=[f'ISNUMBER(SEARCH("{word}",{col_letter}{start_row}))'],
                fill=fill(GRAY_BG),
                font=font(color="616161", bold=True, size=10),
            ),
        )


def freeze_and_filter(ws, header_row, max_col, max_row=200):
    ws.freeze_panes = f"A{header_row + 1}"
    last_col = get_column_letter(max_col)
    ws.auto_filter.ref = f"A{header_row}:{last_col}{max_row}"


def write_headers(ws, row, headers):
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.fill = header_fill()
        cell.font = header_font()
        cell.alignment = center()
        cell.border = thin
    ws.row_dimensions[row].height = 32


def alt_row_fill(ws, row, cols, even_color=CREAM):
    if row % 2 == 0:
        for c in range(1, cols + 1):
            if ws.cell(row=row, column=c).fill.fgColor is None or \
               (hasattr(ws.cell(row=row, column=c).fill.fgColor, "rgb") and
                str(ws.cell(row=row, column=c).fill.fgColor.rgb) in ("00000000", "None")):
                ws.cell(row=row, column=c).fill = fill(even_color)


# ---------------------------------------------------------------------------
# Sample data
# ---------------------------------------------------------------------------
SAMPLE_ROOMS = [
    ("Living Room", "Living Room", "1", 320, "In Progress", 15000, "High", "Primary gathering space"),
    ("Master Bedroom", "Bedroom", "2", 280, "Concept", 12000, "High", "Ensuite suite"),
    ("Kitchen", "Kitchen", "1", 220, "Client Review", 25000, "Urgent", "Full remodel"),
    ("Dining Room", "Dining Room", "1", 180, "Approved", 8000, "Medium", "Formal dining"),
    ("Home Office", "Office", "2", 140, "Not Started", 5000, "Low", "Guest office convertible"),
    ("Powder Room", "Bathroom", "1", 45, "In Progress", 3500, "Medium", "Guest bath"),
]

SAMPLE_FFE = [
    ("FFE-001", "Living Room", "Sofa", "Performance Sofa", "88\" deep seat sofa", 1, '88"W x 40"D x 34"H', "Performance fabric", "Natural oak legs", "Stone", "Article", "Sven", "https://example.com/sofa", "SV-88-ST", 1899, 0.08, 150, 0, 6, "Approved", "Approved", "Ordered", "Not Scheduled", "Client favorite"),
    ("FFE-002", "Living Room", "Chair", "Accent Lounge Chair", "Swivel lounge chair", 2, '30"W x 32"D x 32"H', "Boucle", "Black metal", "Ivory", "West Elm", "Orbit", "https://example.com/chair", "OR-IV", 699, 0.08, 80, 50, 4, "Approved", "Approved", "Shipped", "Scheduled", ""),
    ("FFE-003", "Living Room", "Table", "Coffee Table", "Marble top coffee table", 1, '48"W x 28"D x 16"H', "Marble/Brass", "Brushed brass", "White marble", "CB2", "Vista", "https://example.com/table", "VS-48", 899, 0.08, 100, 0, 5, "Ordered", "Approved", "Ordered", "Not Scheduled", ""),
    ("FFE-004", "Living Room", "Rug", "Area Rug 9x12", "Wool flatweave", 1, "9' x 12'", "Wool", "Handwoven", "Ivory/Taupe", "Loloi", "Magnolia", "https://example.com/rug", "MG-912", 1299, 0.08, 75, 100, 3, "Delivered", "Approved", "Delivered", "Installed", "Pad ordered separately"),
    ("FFE-005", "Living Room", "Lighting", "Floor Lamp", "Arched brass floor lamp", 1, '72"H', "Brass/Linen", "Antique brass", "Brass", "Rejuvenation", "Arc", "https://example.com/lamp", "AR-72", 449, 0.08, 40, 0, 2, "Installed", "Approved", "Delivered", "Installed", ""),
    ("FFE-006", "Master Bedroom", "Bed", "Platform Bed King", "Upholstered platform bed", 1, "King", "Linen", "Natural", "Oatmeal", "Article", "Ceni", "https://example.com/bed", "CN-K", 1499, 0.08, 200, 0, 8, "Client Review", "Sent", "Not Ordered", "Not Scheduled", "Awaiting fabric confirmation"),
    ("FFE-007", "Master Bedroom", "Lighting", "Pendant Pair", "Glass globe pendants", 2, '10" dia', "Glass/Brass", "Clear glass", "Brass", "Schoolhouse", "Globe", "https://example.com/pendant", "GL-10", 279, 0.08, 50, 0, 4, "Approved", "Approved", "Ordered", "Not Scheduled", ""),
    ("FFE-008", "Kitchen", "Appliance", "Range 36\"", "Dual fuel range", 1, '36"W', "Stainless", "Professional", "SS", "Wolf", "DF36450", "https://example.com/range", "DF36450", 8999, 0.08, 350, 200, 12, "Ordered", "Approved", "In Production", "Not Scheduled", "Rough-in required"),
    ("FFE-009", "Kitchen", "Lighting", "Island Pendants", "Linear brass pendants x3", 3, '14"H', "Brass/Glass", "Aged brass", "Brass", "Visual Comfort", "Bryant", "https://example.com/island", "BR-14", 520, 0.08, 90, 0, 6, "Approved", "Approved", "Ordered", "Not Scheduled", ""),
    ("FFE-010", "Dining Room", "Table", "Dining Table", "Oak extension table", 1, '72-96"L x 40"W', "Oak", "White oak", "Natural", "Ethnicraft", "Bok", "https://example.com/dining", "BK-72", 2499, 0.08, 250, 0, 10, "Sourcing", "Not Sent", "Not Ordered", "Not Scheduled", "Comparing 3 options"),
    ("FFE-011", "Dining Room", "Chair", "Dining Chairs", "Upholstered dining chairs", 6, '19"W x 22"D x 33"H', "Performance fabric", "Oak legs", "Sand", "Four Hands", "Haiden", "https://example.com/dchair", "HD-SN", 349, 0.08, 120, 100, 8, "Option Selected", "Not Sent", "Not Ordered", "Not Scheduled", ""),
    ("FFE-012", "Powder Room", "Plumbing Fixture", "Vessel Sink", "Stone vessel sink", 1, '16" dia', "Marble", "Honed", "Carrara", "Signature Hardware", "Vessel", "https://example.com/sink", "VS-16", 389, 0.08, 45, 0, 3, "Delivered", "Approved", "Delivered", "Scheduled", ""),
    ("FFE-013", "Home Office", "Chair", "Desk Chair", "Ergonomic task chair", 1, "Standard", "Mesh/Leather", "Black", "Black", "Herman Miller", "Aeron", "https://example.com/aeron", "AE-B", 1395, 0.08, 0, 100, 2, "Sourcing", "Not Sent", "Not Ordered", "Not Scheduled", "Budget check needed"),
    ("FFE-014", "Living Room", "Artwork", "Large Abstract Canvas", "Commissioned abstract", 1, '48" x 60"', "Canvas/Oil", "Gallery wrap", "Multi", "Local Artist", "Custom", "https://example.com/art", "CUST-01", 2200, 0.00, 75, 0, 8, "Client Review", "Sent", "Not Ordered", "Not Scheduled", "Sketch approved"),
    ("FFE-015", "Kitchen", "Hardware", "Cabinet Pulls", "Brass bar pulls 6\"", 24, '6"', "Solid brass", "Unlacquered", "Brass", "Rejuvenation", "Bin", "https://example.com/pulls", "BN-6", 18, 0.08, 25, 0, 2, "Approved", "Approved", "Ordered", "Not Scheduled", ""),
]

SAMPLE_VENDORS = [
    ("Article", "Furniture", "Sales Team", "trade@article.com", "888-555-0101", "https://article.com", "ART-4521", 15, "4-8 weeks", "Free over $999", "30 days unused", "Preferred furniture vendor"),
    ("West Elm", "Furniture", "Trade Desk", "trade@westelm.com", "800-555-0102", "https://westelm.com", "WE-8891", 20, "2-6 weeks", "Flat $99", "30 days", "Good for accents"),
    ("CB2", "Furniture", "Trade", "trade@cb2.com", "800-555-0103", "https://cb2.com", "CB2-3310", 15, "3-6 weeks", "Varies", "14 days", ""),
    ("Visual Comfort", "Lighting", "Rep: Jordan Lee", "jordan@vc.com", "800-555-0104", "https://visualcomfort.com", "VC-772", 40, "6-12 weeks", "FOB", "Special order no return", "Trade pricing excellent"),
    ("Schoolhouse", "Lighting", "Trade", "trade@schoolhouse.com", "800-555-0105", "https://schoolhouse.com", "SH-220", 20, "2-4 weeks", "Free over $150", "30 days", ""),
    ("Loloi", "Flooring", "Rep: Sam Park", "sam@loloi.com", "800-555-0106", "https://loloirugs.com", "LO-991", 25, "2-4 weeks", "Free over $500", "30 days", "Rug specialist"),
    ("Wolf / Sub-Zero", "Other", "Builder Sales", "builder@subzero.com", "800-555-0107", "https://subzero-wolf.com", "SZ-110", 0, "8-16 weeks", "Freight", "No return installed", "Appliances via dealer"),
    ("Rejuvenation", "Hardware", "Trade", "trade@rejuvenation.com", "800-555-0108", "https://rejuvenation.com", "RJ-554", 15, "1-3 weeks", "Free over $99", "30 days", "Hardware + lighting"),
    ("Four Hands", "Furniture", "Rep: Ava Kim", "ava@fourhands.com", "800-555-0109", "https://fourhands.com", "FH-440", 30, "6-10 weeks", "FOB warehouse", "Special order", "Showroom visit recommended"),
    ("Signature Hardware", "Hardware", "Trade", "trade@signaturehardware.com", "800-555-0110", "https://signaturehardware.com", "SIG-88", 10, "1-3 weeks", "Free over $99", "30 days", "Plumbing fixtures"),
]


# ===========================================================================
# SHEET BUILDERS
# ===========================================================================

def build_lists_sheet(wb):
    """Hidden-friendly reference sheet for dropdown values (also useful in Google Sheets)."""
    ws = wb.create_sheet("Lists", 0)
    lists = {
        "A": ("Room Types", ROOM_TYPES),
        "B": ("FFE Categories", FFE_CATEGORIES),
        "C": ("Order Statuses", ORDER_STATUSES),
        "D": ("Approval Statuses", APPROVAL_STATUSES),
        "E": ("Delivery Statuses", DELIVERY_STATUSES),
        "F": ("Install Statuses", INSTALL_STATUSES),
        "G": ("Priorities", PRIORITY_LEVELS),
        "H": ("Payment Statuses", PAYMENT_STATUSES),
        "I": ("Project Types", PROJECT_TYPES),
        "J": ("Project Statuses", PROJECT_STATUSES),
        "K": ("Billing Methods", BILLING_METHODS),
        "L": ("Yes/No", YES_NO),
        "M": ("Design Statuses", DESIGN_STATUSES),
        "N": ("Fixture Types", FIXTURE_TYPES),
        "O": ("Finish Types", FINISH_TYPES),
        "P": ("Surfaces", SURFACES),
        "Q": ("Vendor Types", VENDOR_TYPES),
        "R": ("Budget Categories", BUDGET_CATEGORIES),
        "S": ("Damage Checks", DAMAGE_CHECKS),
        "T": ("Approval Types", APPROVAL_TYPES),
        "U": ("Project Phases", PROJECT_PHASES),
        "V": ("Task Statuses", TASK_STATUSES),
        "W": ("Task Categories", TASK_CATEGORIES),
        "X": ("Measurement Types", MEASUREMENT_TYPES),
        "Y": ("Trades", TRADES),
        "Z": ("Meeting Types", MEETING_TYPES),
        "AA": ("PO Statuses", PO_STATUSES),
        "AB": ("Room Names", [r[0] for r in SAMPLE_ROOMS]),
        "AC": ("Vendor Names", [v[0] for v in SAMPLE_VENDORS]),
    }
    for col_letter, (title, values) in lists.items():
        col = ws[f"{col_letter}1"]
        col.value = title
        col.font = header_font()
        col.fill = header_fill()
        for i, v in enumerate(values, 2):
            ws[f"{col_letter}{i}"] = v
        ws.column_dimensions[col_letter].width = 22
    ws.sheet_state = "hidden"
    return ws


def build_welcome(wb):
    ws = wb.create_sheet("01 Welcome", 1)
    set_col_widths(ws, [4, 28, 70, 28, 28, 20, 20, 20])

    # Banner
    ws.merge_cells("B2:G2")
    ws["B2"] = "INTERIOR DESIGN FF&E PROJECT TRACKER"
    ws["B2"].font = Font(name="Calibri", bold=True, size=26, color=NAVY)
    ws["B2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 40

    ws.merge_cells("B3:G3")
    ws["B3"] = "Professional Furniture, Fixtures & Equipment Management  ·  Excel + Google Sheets Compatible"
    ws["B3"].font = font(size=12, color=TEAL)
    ws.row_dimensions[3].height = 22

    # Color bar
    for col, color in enumerate([NAVY, TEAL, GOLD, CORAL, NAVY, TEAL], 2):
        ws.cell(row=4, column=col).fill = fill(color)
    ws.row_dimensions[4].height = 8

    # Quick start
    ws.merge_cells("B6:G6")
    ws["B6"] = "QUICK START GUIDE"
    ws["B6"].font = section_font()
    ws["B6"].fill = fill(SOFT_TEAL)

    steps = [
        ("Step 1", "Add client & project details", "Open the Client & Project Profile tab and fill in client name, project address, budget, dates, and designer info."),
        ("Step 2", "Add rooms / areas", "Go to Room List and list every room or zone in the project with budget allocations and priority."),
        ("Step 3", "Add FF&E items", "Use the Master FF&E Schedule to log every furniture, fixture, and equipment item with vendor, cost, and status."),
        ("Step 4", "Track budget, orders, approvals & delivery", "Update Budget Tracker, Purchase Orders, Client Approvals, and Delivery & Installation as the project progresses."),
        ("Step 5", "Use the Dashboard for overview", "The Project Dashboard auto-summarizes KPIs, spend, and status so you always know where the project stands."),
    ]
    r = 7
    for step, title, desc in steps:
        ws.cell(row=r, column=2, value=step).font = font(bold=True, size=12, color=WHITE)
        ws.cell(row=r, column=2).fill = fill(TEAL)
        ws.cell(row=r, column=2).alignment = center()
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        ws.cell(row=r, column=3, value=title).font = font(bold=True, size=12, color=NAVY)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
        ws.cell(row=r, column=5, value=desc).font = font(size=10)
        ws.cell(row=r, column=5).alignment = left_align()
        ws.row_dimensions[r].height = 36
        r += 1

    # What each tab is for
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(row=r, column=2, value="WHAT EACH TAB IS USED FOR").font = section_font()
    ws.cell(row=r, column=2).fill = fill(SOFT_NAVY)
    r += 1

    tabs_info = [
        ("01 Welcome", "You are here — instructions and how to use this workbook"),
        ("02 Dashboard", "Live KPI cards, charts, and project health at a glance"),
        ("03 Client Profile", "Client contact, project scope, budget, and contract details"),
        ("04 Room List", "All rooms/areas with budget allocation and design status"),
        ("05 Master FF&E", "Core schedule — every FF&E item with full cost & status tracking"),
        ("06 Furniture", "Furniture-only focused schedule for quick reviews"),
        ("07 Fixtures", "Lighting, plumbing, hardware, appliances, built-ins"),
        ("08 Finishes", "Paint, tile, flooring, wallpaper, countertops, fabrics"),
        ("09 Sourcing", "Compare product options before final selection"),
        ("10 Vendors", "Supplier database with trade discounts and lead times"),
        ("11 Budget", "Category-level budget vs estimated vs actual"),
        ("12 Purchase Orders", "PO and order tracking with carriers and invoices"),
        ("13 Delivery & Install", "Delivery logistics, damage checks, and installation"),
        ("14 Client Approvals", "Track every decision sent to the client"),
        ("15 Timeline", "Phased project plan with tasks and due dates"),
        ("16 To-Do List", "Daily task list for the design team"),
        ("17 Measurements", "Room and opening measurement log"),
        ("18 Contractors", "Installer / trade coordination"),
        ("19 Invoices", "Client billing and payment tracker"),
        ("20 Meeting Notes", "Consultation and site meeting notes"),
        ("21 Client FF&E Report", "Printable client-facing FF&E summary"),
        ("22 Room Budget Report", "Printable room-by-room budget report"),
        ("23 Vendor Order Report", "Printable vendor order summary"),
        ("24 Install Checklist", "Printable installation checklist"),
        ("25 Punch List", "Printable final punch list"),
    ]
    ws.cell(row=r, column=2, value="Tab").font = header_font()
    ws.cell(row=r, column=2).fill = header_fill()
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
    ws.cell(row=r, column=3, value="Purpose").font = header_font()
    ws.cell(row=r, column=3).fill = header_fill()
    r += 1
    for tab, purpose in tabs_info:
        ws.cell(row=r, column=2, value=tab).font = font(bold=True, size=10, color=TEAL)
        ws.cell(row=r, column=2).border = thin
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        ws.cell(row=r, column=3, value=purpose).font = font(size=10)
        ws.cell(row=r, column=3).border = thin
        if r % 2 == 0:
            for c in range(2, 8):
                ws.cell(row=r, column=c).fill = fill(CREAM)
        r += 1

    # How to enter data
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(row=r, column=2, value="HOW TO ENTER DATA").font = section_font()
    ws.cell(row=r, column=2).fill = fill(SOFT_GOLD)
    r += 1
    tips = [
        "• Replace sample data with your real project information — sample rows are examples only.",
        "• Always start a new row under the last entry; do not insert columns (formulas depend on column order).",
        "• Yellow / green / red cells are driven by conditional formatting based on status values — pick from dropdowns.",
        "• Currency fields use accounting format. Enter numbers only (no $ signs).",
        "• Dates should be entered as real dates (YYYY-MM-DD) so filters and overdue logic work.",
        "• Product links can be full URLs; they become clickable in Excel and Google Sheets.",
        "• Image Link column can hold a URL to a product photo or mood-board image.",
        "• On the Master FF&E tab, Total Cost and Final Cost calculate automatically from Quantity, Unit Cost, Tax, Shipping, and Discount.",
    ]
    for tip in tips:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.cell(row=r, column=2, value=tip).font = font(size=10)
        r += 1

    # Color coding
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(row=r, column=2, value="COLOR CODING EXPLANATION").font = section_font()
    ws.cell(row=r, column=2).fill = fill(SOFT_NAVY)
    r += 1
    colors = [
        (GREEN_BG, GREEN, "GREEN", "Complete / Approved / On budget / Good condition"),
        (YELLOW_BG, "F57F17", "YELLOW", "Pending / In progress / Awaiting response"),
        (RED_BG, RED, "RED", "Overdue / Over budget / Delayed / Damaged / Rejected"),
        (GRAY_BG, "616161", "GRAY", "Cancelled / On hold / Not started"),
        (SOFT_TEAL, TEAL, "TEAL HEADERS", "Section labels and accent UI"),
        (NAVY, WHITE, "NAVY HEADERS", "Column headers on all schedule tabs"),
    ]
    for bg, fg, label, meaning in colors:
        ws.cell(row=r, column=2, value=label).font = font(bold=True, size=10, color=fg if fg != WHITE else WHITE)
        ws.cell(row=r, column=2).fill = fill(bg)
        ws.cell(row=r, column=2).alignment = center()
        ws.cell(row=r, column=2).border = thin
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        ws.cell(row=r, column=3, value=meaning).font = font(size=10)
        ws.cell(row=r, column=3).border = thin
        r += 1

    # Dropdown instructions
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(row=r, column=2, value="DROPDOWN INSTRUCTIONS").font = section_font()
    ws.cell(row=r, column=2).fill = fill(SOFT_TEAL)
    r += 1
    for tip in [
        "• Click any status, category, room, priority, or type cell — a dropdown arrow appears.",
        "• Always choose from the list so Dashboard counts and conditional formatting stay accurate.",
        "• To customize lists (e.g. add a custom room type), unhide the Lists sheet and edit the column values, then update data validation if needed.",
        "• Common dropdowns: Room, Category, Vendor, Status, Priority, Approval, Order, Delivery, Install, Payment.",
    ]:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.cell(row=r, column=2, value=tip).font = font(size=10)
        r += 1

    # Compatibility
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(row=r, column=2, value="COMPATIBILITY — EXCEL + GOOGLE SHEETS").font = section_font()
    ws.cell(row=r, column=2).fill = fill(SOFT_GOLD)
    r += 1
    for tip in [
        "• This file is a standard .xlsx workbook. Open it in Microsoft Excel (Windows/Mac), Excel Online, LibreOffice, or upload to Google Drive and open with Google Sheets.",
        "• Formulas, filters, freeze panes, data validation dropdowns, and charts transfer to Google Sheets.",
        "• Conditional formatting rules are included; colors may need a quick refresh after first open in Sheets.",
        "• Charts on the Dashboard may need slight repositioning after import into Google Sheets.",
        "• Recommended: Excel 2016+ or current Google Sheets. Enable editing (not Protected View) to use dropdowns and filters.",
    ]:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.cell(row=r, column=2, value=tip).font = font(size=10)
        r += 1

    # Editable note
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(row=r, column=2, value="EDITABLE / UNLOCKED VERSION").font = section_font()
    ws.cell(row=r, column=2).fill = fill(SOFT_NAVY)
    r += 1
    for tip in [
        "• This is the fully editable / unlocked version — all cells, formulas, and sheets can be modified.",
        "• You may rename sheets, add columns to the right of existing data, and add as many rows as needed.",
        "• Before heavy customization, Save a clean master copy so you can start new projects from a blank template.",
        "• To start a new project: File → Save As → clear sample rows (keep header row) → update Client Profile.",
    ]:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.cell(row=r, column=2, value=tip).font = font(size=10)
        r += 1

    # Support
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    ws.cell(row=r, column=2, value="SUPPORT / CONTACT").font = section_font()
    ws.cell(row=r, column=2).fill = fill(SOFT_TEAL)
    r += 1
    for tip in [
        "• Questions or customization requests: contact your template provider or studio admin.",
        "• Tip: Keep one master blank file and duplicate it for every new client project.",
        "• Version: Interior Design FF&E Tracker v1.0  ·  Generated for professional design studios.",
    ]:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.cell(row=r, column=2, value=tip).font = font(size=10)
        r += 1

    ws.sheet_view.showGridLines = False
    return ws


def build_dashboard(wb):
    ws = wb.create_sheet("02 Dashboard", 2)
    set_col_widths(ws, [3, 22, 16, 16, 16, 16, 16, 16, 16, 16, 16, 3])

    ws.merge_cells("B2:K2")
    ws["B2"] = "PROJECT DASHBOARD"
    ws["B2"].font = title_font()
    ws["B2"].fill = fill(SOFT_NAVY)
    for c in range(2, 12):
        ws.cell(row=2, column=c).fill = fill(SOFT_NAVY)
    ws.row_dimensions[2].height = 36

    ws.merge_cells("B3:K3")
    ws["B3"] = '=CONCATENATE("Project: ",\'03 Client Profile\'!C6,"  |  Client: ",\'03 Client Profile\'!C4,"  |  Status: ",\'03 Client Profile\'!C16)'
    ws["B3"].font = font(size=11, color=TEAL)

    # KPI cards — row 5 labels, row 6 values
    kpis = [
        (2, "TOTAL PROJECT BUDGET", "='03 Client Profile'!C13", NAVY, money_format()),
        (3, "TOTAL ESTIMATED COST", "=SUMIF('05 Master FF&E'!V:V,\"<>Cancelled\",'05 Master FF&E'!U:U)", TEAL, money_format()),
        (4, "TOTAL ACTUAL COST", "=SUMIF('05 Master FF&E'!V:V,\"<>Cancelled\",'05 Master FF&E'!U:U)", TEAL, money_format()),
        (5, "REMAINING BUDGET", "=B6-C6", CORAL, money_format()),
        (6, "BUDGET USED %", "=IF(B6=0,0,C6/B6)", GOLD, pct_format()),
        (7, "NUMBER OF ROOMS", "=COUNTA('04 Room List'!A5:A200)", NAVY, "0"),
        (8, "TOTAL FF&E ITEMS", "=COUNTA('05 Master FF&E'!A5:A500)", NAVY, "0"),
        (9, "ITEMS APPROVED", '=COUNTIF(\'05 Master FF&E\'!W5:W500,"Approved")', GREEN, "0"),
        (10, "ITEMS ORDERED", '=COUNTIFS(\'05 Master FF&E\'!V5:V500,"<>Not Ordered",\'05 Master FF&E\'!V5:V500,"<>Sourcing",\'05 Master FF&E\'!V5:V500,"<>Option Selected",\'05 Master FF&E\'!V5:V500,"<>Client Review",\'05 Master FF&E\'!V5:V500,"<>")', TEAL, "0"),
        (11, "ITEMS DELIVERED", '=COUNTIF(\'05 Master FF&E\'!X5:X500,"Delivered")+COUNTIF(\'05 Master FF&E\'!Y5:Y500,"Installed")', TEAL, "0"),
    ]
    # First row of KPIs
    for col, label, formula, color, num_fmt in kpis[:5]:
        ws.cell(row=5, column=col, value=label).font = kpi_label_font()
        ws.cell(row=5, column=col).fill = fill(LIGHT_GRAY)
        ws.cell(row=5, column=col).alignment = center()
        ws.cell(row=5, column=col).border = thin
        cell = ws.cell(row=6, column=col, value=formula)
        cell.font = kpi_value_font(color)
        cell.number_format = num_fmt
        cell.alignment = center()
        cell.border = thin
        cell.fill = fill(WHITE)
    ws.row_dimensions[5].height = 28
    ws.row_dimensions[6].height = 36

    for col, label, formula, color, num_fmt in kpis[5:]:
        ws.cell(row=8, column=col - 5 + 1 if col < 7 else col - 4, value="")  # placeholder

    # Second KPI row — remap columns 2-11 for items 5-10 conceptually as cols 2-6 then more
    kpis2 = [
        (2, "ITEMS INSTALLED", '=COUNTIF(\'05 Master FF&E\'!Y5:Y500,"Installed")+COUNTIF(\'05 Master FF&E\'!Y5:Y500,"Complete")', GREEN, "0"),
        (3, "OVERDUE TASKS", '=COUNTIFS(\'16 To-Do List\'!F5:F200,"<"&TODAY(),\'16 To-Do List\'!G5:G200,"<>Complete")', RED, "0"),
        (4, "PENDING APPROVALS", '=COUNTIF(\'14 Client Approvals\'!F5:F200,"Sent")+COUNTIF(\'05 Master FF&E\'!W5:W500,"Sent")', YELLOW.replace("F9A825", "F9A825"), "0"),
        (5, "OPEN VENDOR ORDERS", '=COUNTIF(\'12 Purchase Orders\'!J5:J200,"Ordered")+COUNTIF(\'12 Purchase Orders\'!J5:J200,"Confirmed")+COUNTIF(\'12 Purchase Orders\'!J5:J200,"In Production")+COUNTIF(\'12 Purchase Orders\'!J5:J200,"Shipped")', CORAL, "0"),
        (6, "ITEMS NOT ORDERED", '=COUNTIF(\'05 Master FF&E\'!V5:V500,"Not Ordered")+COUNTIF(\'05 Master FF&E\'!V5:V500,"Sourcing")+COUNTIF(\'05 Master FF&E\'!V5:V500,"Option Selected")', "616161", "0"),
    ]
    for col, label, formula, color, num_fmt in kpis2:
        ws.cell(row=8, column=col, value=label).font = kpi_label_font()
        ws.cell(row=8, column=col).fill = fill(LIGHT_GRAY)
        ws.cell(row=8, column=col).alignment = center()
        ws.cell(row=8, column=col).border = thin
        cell = ws.cell(row=9, column=col, value=formula)
        cell.font = kpi_value_font(color if color != "F9A825" else "F57F17")
        cell.number_format = num_fmt
        cell.alignment = center()
        cell.border = thin
    # extra KPIs
    extra = [
        (7, "DELIVERED NOT INSTALLED", '=COUNTIFS(\'05 Master FF&E\'!X5:X500,"Delivered",\'05 Master FF&E\'!Y5:Y500,"<>Installed",\'05 Master FF&E\'!Y5:Y500,"<>Complete")', CORAL, "0"),
        (8, "TOTAL ROOMS BUDGET", "=SUM('04 Room List'!F5:F200)", NAVY, money_format()),
        (9, "ROOMS OVER BUDGET", '=COUNTIF(\'04 Room List\'!I5:I200,"<0")', RED, "0"),
        (10, "VENDORS ON FILE", "=COUNTA('10 Vendors'!A5:A200)", NAVY, "0"),
        (11, "OPEN INVOICES", '=COUNTIF(\'19 Invoices\'!H5:H200,"Sent")+COUNTIF(\'19 Invoices\'!H5:H200,"Partially Paid")+COUNTIF(\'19 Invoices\'!H5:H200,"Overdue")', YELLOW, "0"),
    ]
    for col, label, formula, color, num_fmt in extra:
        ws.cell(row=8, column=col, value=label).font = kpi_label_font()
        ws.cell(row=8, column=col).fill = fill(LIGHT_GRAY)
        ws.cell(row=8, column=col).alignment = center()
        ws.cell(row=8, column=col).border = thin
        cell = ws.cell(row=9, column=col, value=formula)
        cell.font = kpi_value_font(color if len(color) == 6 else color)
        cell.number_format = num_fmt
        cell.alignment = center()
        cell.border = thin
    ws.row_dimensions[8].height = 28
    ws.row_dimensions[9].height = 36

    # Chart data tables (for Excel charts) — hidden-ish area starting row 12
    ws["B11"] = "CHART DATA & VISUAL SUMMARIES"
    ws["B11"].font = section_font()
    ws.merge_cells("B11:E11")

    # Budget by room
    ws["B13"] = "Budget by Room"
    ws["B13"].font = font(bold=True, size=11, color=NAVY)
    ws["B14"] = "Room"
    ws["C14"] = "Budget"
    ws["D14"] = "Actual"
    for c in range(2, 5):
        ws.cell(row=14, column=c).fill = header_fill()
        ws.cell(row=14, column=c).font = header_font()
    for i in range(6):
        ws.cell(row=15 + i, column=2, value=f"='04 Room List'!A{5+i}")
        ws.cell(row=15 + i, column=3, value=f"='04 Room List'!F{5+i}")
        ws.cell(row=15 + i, column=3).number_format = money_format()
        ws.cell(row=15 + i, column=4, value=f"='04 Room List'!H{5+i}")
        ws.cell(row=15 + i, column=4).number_format = money_format()

    chart1 = BarChart()
    chart1.type = "col"
    chart1.grouping = "clustered"
    chart1.title = "Budget vs Actual by Room"
    chart1.y_axis.title = "Amount ($)"
    chart1.style = 10
    data = Reference(ws, min_col=3, min_row=14, max_col=4, max_row=20)
    cats = Reference(ws, min_col=2, min_row=15, max_row=20)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 4
    chart1.width = 15
    chart1.height = 10
    ws.add_chart(chart1, "F13")

    # Item status breakdown
    ws["B23"] = "Item Status Breakdown"
    ws["B23"].font = font(bold=True, size=11, color=NAVY)
    statuses_for_chart = ["Sourcing", "Approved", "Ordered", "Shipped", "Delivered", "Installed", "Cancelled"]
    ws["B24"] = "Status"
    ws["C24"] = "Count"
    ws["B24"].fill = header_fill()
    ws["B24"].font = header_font()
    ws["C24"].fill = header_fill()
    ws["C24"].font = header_font()
    for i, st in enumerate(statuses_for_chart):
        ws.cell(row=25 + i, column=2, value=st)
        ws.cell(row=25 + i, column=3, value=f'=COUNTIF(\'05 Master FF&E\'!V5:V500,"{st}")+COUNTIF(\'05 Master FF&E\'!Y5:Y500,"{st}")')

    pie1 = PieChart()
    pie1.title = "Item Status Breakdown"
    labels = Reference(ws, min_col=2, min_row=25, max_row=31)
    data = Reference(ws, min_col=3, min_row=24, max_row=31)
    pie1.add_data(data, titles_from_data=True)
    pie1.set_categories(labels)
    pie1.dataLabels = DataLabelList()
    pie1.dataLabels.showPercent = True
    pie1.dataLabels.showVal = False
    pie1.width = 12
    pie1.height = 10
    ws.add_chart(pie1, "F28")

    # Category spend
    ws["B34"] = "Category Spend"
    ws["B34"].font = font(bold=True, size=11, color=NAVY)
    cats_chart = ["Sofa", "Chair", "Table", "Bed", "Rug", "Lighting", "Appliance", "Artwork", "Hardware", "Plumbing Fixture"]
    ws["B35"] = "Category"
    ws["C35"] = "Spend"
    ws["B35"].fill = header_fill()
    ws["B35"].font = header_font()
    ws["C35"].fill = header_fill()
    ws["C35"].font = header_font()
    for i, cat in enumerate(cats_chart):
        ws.cell(row=36 + i, column=2, value=cat)
        ws.cell(row=36 + i, column=3, value=f'=SUMIF(\'05 Master FF&E\'!C5:C500,"{cat}",\'05 Master FF&E\'!U5:U500)')
        ws.cell(row=36 + i, column=3).number_format = money_format()

    chart2 = BarChart()
    chart2.type = "bar"
    chart2.title = "Category Spend Breakdown"
    chart2.style = 11
    data = Reference(ws, min_col=3, min_row=35, max_row=45)
    cats = Reference(ws, min_col=2, min_row=36, max_row=45)
    chart2.add_data(data, titles_from_data=True)
    chart2.set_categories(cats)
    chart2.width = 15
    chart2.height = 10
    ws.add_chart(chart2, "F40")

    # Delivery status
    ws["B48"] = "Delivery Status"
    ws["B48"].font = font(bold=True, size=11, color=NAVY)
    del_statuses = ["Not Ordered", "Ordered", "Shipped", "Delivered", "Delayed"]
    ws["B49"] = "Status"
    ws["C49"] = "Count"
    ws["B49"].fill = header_fill()
    ws["B49"].font = header_font()
    ws["C49"].fill = header_fill()
    ws["C49"].font = header_font()
    for i, st in enumerate(del_statuses):
        ws.cell(row=50 + i, column=2, value=st)
        ws.cell(row=50 + i, column=3, value=f'=COUNTIF(\'05 Master FF&E\'!X5:X500,"{st}")')

    pie2 = PieChart()
    pie2.title = "Delivery Status Breakdown"
    labels = Reference(ws, min_col=2, min_row=50, max_row=54)
    data = Reference(ws, min_col=3, min_row=49, max_row=54)
    pie2.add_data(data, titles_from_data=True)
    pie2.set_categories(labels)
    pie2.dataLabels = DataLabelList()
    pie2.dataLabels.showPercent = True
    pie2.width = 12
    pie2.height = 9
    ws.add_chart(pie2, "B56")

    # Quick legend
    ws["F56"] = "STATUS COLOR KEY"
    ws["F56"].font = font(bold=True, size=11, color=NAVY)
    for i, (bg, label) in enumerate([
        (GREEN_BG, "Green = Complete / Approved / On budget"),
        (YELLOW_BG, "Yellow = Pending / In progress"),
        (RED_BG, "Red = Overdue / Over budget / Delayed"),
        (GRAY_BG, "Gray = Cancelled / On hold"),
    ]):
        ws.cell(row=57 + i, column=6, value=label).fill = fill(bg)
        ws.cell(row=57 + i, column=6).font = font(size=10)
        ws.merge_cells(start_row=57 + i, start_column=6, end_row=57 + i, end_column=9)

    ws.sheet_view.showGridLines = False
    return ws


def build_client_profile(wb):
    ws = wb.create_sheet("03 Client Profile", 3)
    set_col_widths(ws, [3, 28, 40, 20, 28, 30])

    add_title_banner(ws, "CLIENT & PROJECT PROFILE", "Project summary — update these fields first", 6)

    sections = [
        (4, "CLIENT INFORMATION", [
            ("Client Name", "Alexandra Chen"),
            ("Client Email", "alexandra.chen@email.com"),
            ("Phone Number", "(555) 234-8890"),
        ]),
        (9, "PROJECT DETAILS", [
            ("Project Name", "Chen Residence Refresh"),
            ("Project Address", "482 Maple Lane, Portland, OR 97205"),
            ("Project Type", "Residential"),
            ("Design Style", "Modern Organic / Warm Minimalist"),
            ("Start Date", datetime(2026, 3, 1)),
            ("Target Completion", datetime(2026, 9, 30)),
            ("Total Budget", 85000),
            ("Designer Name", "Jordan Blake Studio"),
            ("Project Status", "Active"),
            ("Scope Summary", "Full FF&E package for living, primary suite, kitchen refresh, dining, powder room, and home office. Includes furniture, lighting, rugs, finishes, and installation coordination."),
            ("Client Notes", "Prefers low-VOC finishes. Dog-friendly fabrics required in living areas. Traveling Aug 1–15 — no deliveries that window."),
        ]),
        (23, "CONTRACT & BILLING (OPTIONAL)", [
            ("Contract Signed", "Yes"),
            ("Deposit Paid", "Yes"),
            ("Design Package Selected", "Full Service FF&E"),
            ("Billing Method", "Hybrid"),
            ("Design Fee", 12000),
            ("Hourly Rate (if applicable)", 175),
            ("Retainer / Deposit Amount", 4000),
        ]),
    ]

    for start_row, title, fields in sections:
        ws.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=3)
        ws.cell(row=start_row, column=2, value=title).font = section_font()
        ws.cell(row=start_row, column=2).fill = fill(SOFT_TEAL)
        ws.cell(row=start_row, column=3).fill = fill(SOFT_TEAL)
        for i, (label, value) in enumerate(fields):
            r = start_row + 1 + i
            ws.cell(row=r, column=2, value=label).font = font(bold=True, size=11)
            ws.cell(row=r, column=2).fill = fill(LIGHT_GRAY)
            ws.cell(row=r, column=2).border = thin
            cell = ws.cell(row=r, column=3, value=value)
            cell.font = font(size=11)
            cell.border = thin
            cell.fill = fill(SOFT_GOLD)  # editable highlight
            if isinstance(value, datetime):
                cell.number_format = date_format()
            if label in ("Total Budget", "Design Fee", "Hourly Rate (if applicable)", "Retainer / Deposit Amount"):
                cell.number_format = money_format()
            ws.row_dimensions[r].height = 22

    # Dropdowns for profile (C12 type, C18 status, C24/C25 yes-no, C27 billing)
    add_dropdown(ws, PROJECT_TYPES, "C12")
    add_dropdown(ws, PROJECT_STATUSES, "C18")
    add_dropdown(ws, YES_NO, "C24")
    add_dropdown(ws, YES_NO, "C25")
    add_dropdown(ws, BILLING_METHODS, "C27")

    # Map named conceptual cells for dashboard:
    # C4 client name, C6 project name, C13 total budget, C16 status — already set by field order:
    # Client: C4 name, C5 email, C6 phone — wait, start_row 4 + fields:
    # start 4 title, C5 name, C6 email, C7 phone
    # Project start 9: C10 name, C11 address, C12 type, C13 style, C14 start, C15 target, C16 budget, C17 designer, C18 status
    # Dashboard formulas reference C6 as project and C4 client and C13 budget and C16 status
    # Fix: reorder so Dashboard refs match. Easier to update Dashboard to correct cells.
    # Client name = C5, Project name = C10, Budget = C16, Status = C18
    # I'll patch dashboard later... actually let me fix field layout to match common expectations
    # Re-write key cells with comments for user

    ws["E4"] = "FIELD GUIDE"
    ws["E4"].font = section_font()
    ws["E5"] = "Gold cells are editable inputs."
    ws["E6"] = "Gray cells are labels — leave them."
    ws["E7"] = "Use dropdowns where provided."
    ws["E8"] = "Budget feeds the Dashboard KPIs."

    for r in range(5, 9):
        ws.cell(row=r, column=5).font = font(size=10, color="666666")

    ws.sheet_view.showGridLines = False
    return ws


def fix_dashboard_refs(wb):
    """Point dashboard formulas at correct Client Profile cells after layout."""
    ws = wb["02 Dashboard"]
    # Client name C5, Project C10, Budget C16, Status C18
    ws["B3"] = '=CONCATENATE("Project: ",\'03 Client Profile\'!C10,"  |  Client: ",\'03 Client Profile\'!C5,"  |  Status: ",\'03 Client Profile\'!C18)'
    # Total project budget
    ws["B6"] = "='03 Client Profile'!C16"


def build_room_list(wb):
    ws = wb.create_sheet("04 Room List", 4)
    headers = [
        "Room Name", "Room Type", "Floor/Level", "Square Footage", "Design Status",
        "Budget Allocation", "Estimated Spend", "Actual Spend", "Remaining Budget",
        "Priority Level", "Notes",
    ]
    add_title_banner(ws, "ROOM LIST / AREA TRACKER", "Organize the project by room — budgets roll up to the Dashboard", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [18, 16, 12, 14, 14, 16, 15, 14, 16, 12, 36])

    for i, room in enumerate(SAMPLE_ROOMS):
        r = 5 + i
        name, rtype, floor, sqft, status, budget, priority, notes = room
        values = [name, rtype, floor, sqft, status, budget, None, None, None, priority, notes]
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, money=(c in (6, 7, 8, 9)), center_text=(c in (3, 4, 5, 10)))
        # Estimated = SUMIF from FF&E by room (Final Cost col U)
        ws.cell(row=r, column=7, value=f"=SUMIF('05 Master FF&E'!B:B,A{r},'05 Master FF&E'!U:U)")
        ws.cell(row=r, column=7).number_format = money_format()
        ws.cell(row=r, column=7).border = thin
        # Actual = same for demo (or only ordered/delivered) — use final cost of approved+ items
        ws.cell(row=r, column=8, value=f"=SUMIFS('05 Master FF&E'!U:U,'05 Master FF&E'!B:B,A{r},'05 Master FF&E'!V:V,\"<>Cancelled\",'05 Master FF&E'!V:V,\"<>Sourcing\")")
        ws.cell(row=r, column=8).number_format = money_format()
        ws.cell(row=r, column=8).border = thin
        # Remaining
        ws.cell(row=r, column=9, value=f"=F{r}-H{r}")
        ws.cell(row=r, column=9).number_format = money_format()
        ws.cell(row=r, column=9).border = thin

    # Empty template rows with formulas
    for i in range(len(SAMPLE_ROOMS), 30):
        r = 5 + i
        for c in range(1, 12):
            style_data_cell(ws.cell(row=r, column=c), money=(c in (6, 7, 8, 9)))
        ws.cell(row=r, column=7, value=f"=IF(A{r}=\"\",\"\",SUMIF('05 Master FF&E'!B:B,A{r},'05 Master FF&E'!U:U))")
        ws.cell(row=r, column=7).number_format = money_format()
        ws.cell(row=r, column=8, value=f"=IF(A{r}=\"\",\"\",SUMIFS('05 Master FF&E'!U:U,'05 Master FF&E'!B:B,A{r},'05 Master FF&E'!V:V,\"<>Cancelled\"))")
        ws.cell(row=r, column=8).number_format = money_format()
        ws.cell(row=r, column=9, value=f"=IF(A{r}=\"\",\"\",F{r}-H{r})")
        ws.cell(row=r, column=9).number_format = money_format()

    # Totals
    tr = 36
    ws.cell(row=tr, column=5, value="TOTALS").font = font(bold=True, color=WHITE)
    ws.cell(row=tr, column=5).fill = fill(NAVY)
    for col, letter in [(6, "F"), (7, "G"), (8, "H"), (9, "I")]:
        cell = ws.cell(row=tr, column=col, value=f"=SUM({letter}5:{letter}34)")
        cell.font = font(bold=True, color=WHITE)
        cell.fill = fill(NAVY)
        cell.number_format = money_format()
        cell.border = thin

    add_dropdown(ws, ROOM_TYPES, "B5:B34")
    add_dropdown(ws, DESIGN_STATUSES, "E5:E34")
    add_dropdown(ws, PRIORITY_LEVELS, "J5:J34")
    apply_status_conditional_formatting(ws, "E", 5, 34)
    apply_status_conditional_formatting(ws, "J", 5, 34)
    # Remaining budget red if negative
    ws.conditional_formatting.add(
        "I5:I34",
        CellIsRule(operator="lessThan", formula=["0"], fill=fill(RED_BG), font=font(color=RED, bold=True)),
    )
    freeze_and_filter(ws, 4, 11, 34)
    return ws


def build_master_ffe(wb):
    ws = wb.create_sheet("05 Master FF&E", 5)
    headers = [
        "Item ID", "Room", "Category", "Item Name", "Description", "Quantity",
        "Dimensions", "Material", "Finish", "Color", "Vendor", "Brand",
        "Product Link", "SKU/Model", "Unit Cost", "Total Cost", "Tax",
        "Shipping", "Discount", "Final Cost", "Lead Time (wks)", "Order Status",
        "Approval Status", "Delivery Status", "Install Status", "Notes", "Image Link",
    ]
    # Note: user asked Final cost and also Total cost. Layout:
    # O Unit cost, P Total cost (=qty*unit), Q Tax (amount or rate applied), R Shipping, S Discount, T Final cost
    # Wait I had U as final in formulas. Let me align columns carefully:
    # A Item ID ... O Unit Cost, P Total Cost, Q Tax, R Shipping, S Discount, T Final Cost,
    # U Lead Time, V Order Status, W Approval, X Delivery, Y Install, Z Notes, AA Image
    # Actually headers list has Lead Time before Order Status so:
    # O=15 Unit, P=16 Total, Q=17 Tax, R=18 Shipping, S=19 Discount, T=20 Final, U=21 Lead, V=22 Order, W=23 Approval, X=24 Delivery, Y=25 Install, Z=26 Notes, AA=27 Image

    add_title_banner(ws, "MASTER FF&E SCHEDULE", "Furniture, Fixtures & Equipment — the core project schedule", 12)
    write_headers(ws, 4, headers)
    widths = [10, 16, 14, 20, 24, 10, 16, 14, 14, 12, 16, 14, 22, 12, 12, 12, 10, 11, 11, 12, 12, 14, 14, 14, 14, 28, 18]
    set_col_widths(ws, widths)

    for i, item in enumerate(SAMPLE_FFE):
        r = 5 + i
        (iid, room, cat, name, desc, qty, dims, mat, finish, color, vendor, brand,
         link, sku, unit, tax_rate, ship, disc, lead, order_st, appr, deliv, install, notes) = item
        row_vals = {
            1: iid, 2: room, 3: cat, 4: name, 5: desc, 6: qty, 7: dims, 8: mat, 9: finish,
            10: color, 11: vendor, 12: brand, 13: link, 14: sku, 15: unit,
            # 16 total formula, 17 tax amount formula from rate, 18 ship, 19 disc, 20 final formula
            17: round(unit * qty * tax_rate, 2),  # store tax as amount for simplicity
            18: ship, 19: disc, 21: lead, 22: order_st, 23: appr, 24: deliv, 25: install, 26: notes, 27: "",
        }
        for c, v in row_vals.items():
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, money=(c in (15, 16, 17, 18, 19, 20)), center_text=(c in (6, 21, 22, 23, 24, 25)))
        # Total cost = qty * unit
        ws.cell(row=r, column=16, value=f"=F{r}*O{r}")
        ws.cell(row=r, column=16).number_format = money_format()
        ws.cell(row=r, column=16).border = thin
        # Final cost = total + tax + shipping - discount
        ws.cell(row=r, column=20, value=f"=P{r}+Q{r}+R{r}-S{r}")
        ws.cell(row=r, column=20).number_format = money_format()
        ws.cell(row=r, column=20).border = thin
        ws.cell(row=r, column=15).number_format = money_format()
        ws.cell(row=r, column=17).number_format = money_format()
        ws.cell(row=r, column=18).number_format = money_format()
        ws.cell(row=r, column=19).number_format = money_format()

    # Extra empty rows with formulas
    for i in range(len(SAMPLE_FFE), 80):
        r = 5 + i
        for c in range(1, 28):
            style_data_cell(ws.cell(row=r, column=c), money=(c in (15, 16, 17, 18, 19, 20)))
        ws.cell(row=r, column=16, value=f"=IF(F{r}=\"\",\"\",F{r}*O{r})")
        ws.cell(row=r, column=16).number_format = money_format()
        ws.cell(row=r, column=20, value=f"=IF(F{r}=\"\",\"\",P{r}+Q{r}+R{r}-S{r})")
        ws.cell(row=r, column=20).number_format = money_format()

    # Totals row
    tr = 86
    ws.cell(row=tr, column=14, value="TOTALS").font = font(bold=True, color=WHITE)
    ws.cell(row=tr, column=14).fill = fill(NAVY)
    for col, letter in [(15, "O"), (16, "P"), (17, "Q"), (18, "R"), (19, "S"), (20, "T")]:
        cell = ws.cell(row=tr, column=col, value=f"=SUM({letter}5:{letter}84)")
        cell.font = font(bold=True, color=WHITE)
        cell.fill = fill(NAVY)
        cell.number_format = money_format()

    # Dropdowns — room names from sample + types
    room_names = [r[0] for r in SAMPLE_ROOMS]
    add_dropdown(ws, room_names, "B5:B84")
    add_dropdown(ws, FFE_CATEGORIES, "C5:C84")
    vendor_names = [v[0] for v in SAMPLE_VENDORS]
    add_dropdown(ws, vendor_names, "K5:K84")
    add_dropdown(ws, ORDER_STATUSES, "V5:V84")
    add_dropdown(ws, APPROVAL_STATUSES, "W5:W84")
    add_dropdown(ws, DELIVERY_STATUSES, "X5:X84")
    add_dropdown(ws, INSTALL_STATUSES, "Y5:Y84")

    for col in ("V", "W", "X", "Y"):
        apply_status_conditional_formatting(ws, col, 5, 84)

    freeze_and_filter(ws, 4, 27, 84)

    # Note: Dashboard formulas reference U as Final Cost — need to fix to T
    return ws


def fix_dashboard_ffe_cols(wb):
    """Master FF&E Final Cost is column T (20), Order Status V, Approval W, Delivery X, Install Y."""
    ws = wb["02 Dashboard"]
    ws["C6"] = "=SUMIF('05 Master FF&E'!V:V,\"<>Cancelled\",'05 Master FF&E'!T:T)"
    ws["D6"] = "=SUMIF('05 Master FF&E'!V:V,\"<>Cancelled\",'05 Master FF&E'!T:T)"
    ws["E6"] = "=B6-C6"
    ws["F6"] = "=IF(B6=0,0,C6/B6)"
    # status counts already use V W X Y correctly in build_dashboard for most
    # Category spend used U — fix to T
    for i, cat in enumerate(["Sofa", "Chair", "Table", "Bed", "Rug", "Lighting", "Appliance", "Artwork", "Hardware", "Plumbing Fixture"]):
        ws.cell(row=36 + i, column=3, value=f'=SUMIF(\'05 Master FF&E\'!C5:C500,"{cat}",\'05 Master FF&E\'!T5:T500)')
        ws.cell(row=36 + i, column=3).number_format = money_format()

    # Room list formulas use U — fix to T
    rooms = wb["04 Room List"]
    for r in range(5, 35):
        rooms.cell(row=r, column=7, value=f"=IF(A{r}=\"\",\"\",SUMIF('05 Master FF&E'!B:B,A{r},'05 Master FF&E'!T:T))")
        rooms.cell(row=r, column=7).number_format = money_format()
        rooms.cell(row=r, column=8, value=f"=IF(A{r}=\"\",\"\",SUMIFS('05 Master FF&E'!T:T,'05 Master FF&E'!B:B,A{r},'05 Master FF&E'!V:V,\"<>Cancelled\"))")
        rooms.cell(row=r, column=8).number_format = money_format()


def build_furniture(wb):
    ws = wb.create_sheet("06 Furniture", 6)
    headers = [
        "Room", "Furniture Item", "Dimensions", "Quantity", "Vendor", "Product Link",
        "Fabric/Material", "Finish", "Cost", "Approval Status", "Order Status",
        "Delivery Date", "Install Date", "Notes",
    ]
    add_title_banner(ws, "FURNITURE SCHEDULE", "Furniture-only view — filtered from the broader FF&E program", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [16, 22, 18, 10, 16, 22, 16, 14, 12, 14, 14, 14, 14, 28])

    furniture_cats = {"Sofa", "Chair", "Table", "Bed", "Rug"}
    row = 5
    for item in SAMPLE_FFE:
        if item[2] in furniture_cats:
            (iid, room, cat, name, desc, qty, dims, mat, finish, color, vendor, brand,
             link, sku, unit, tax_rate, ship, disc, lead, order_st, appr, deliv, install, notes) = item
            final = qty * unit + qty * unit * tax_rate + ship - disc
            vals = [room, name, dims, qty, vendor, link, mat, finish, final, appr, order_st, "", "", notes]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=row, column=c, value=v)
                style_data_cell(cell, money=(c == 9), date=(c in (12, 13)), center_text=(c in (4, 10, 11)))
            row += 1

    for r in range(row, 50):
        for c in range(1, 15):
            style_data_cell(ws.cell(row=r, column=c), money=(c == 9), date=(c in (12, 13)))

    add_dropdown(ws, [r[0] for r in SAMPLE_ROOMS], "A5:A49")
    add_dropdown(ws, [v[0] for v in SAMPLE_VENDORS], "E5:E49")
    add_dropdown(ws, APPROVAL_STATUSES, "J5:J49")
    add_dropdown(ws, ORDER_STATUSES, "K5:K49")
    apply_status_conditional_formatting(ws, "J", 5, 49)
    apply_status_conditional_formatting(ws, "K", 5, 49)
    freeze_and_filter(ws, 4, 14, 49)
    return ws


def build_fixtures(wb):
    ws = wb.create_sheet("07 Fixtures", 7)
    headers = [
        "Room", "Fixture Type", "Item Name", "Finish", "Size", "Vendor", "Model/SKU",
        "Quantity", "Unit Cost", "Total Cost", "Required Rough-In Date", "Order Deadline",
        "Delivery Date", "Installer Notes",
    ]
    add_title_banner(ws, "FIXTURES SCHEDULE", "Lighting, plumbing, hardware, appliances, and built-in fixtures", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [16, 16, 22, 14, 12, 16, 14, 10, 12, 12, 16, 14, 14, 30])

    fixture_cats = {"Lighting", "Plumbing Fixture", "Hardware", "Appliance"}
    cat_to_type = {
        "Lighting": "Lighting",
        "Plumbing Fixture": "Faucet",
        "Hardware": "Cabinet Hardware",
        "Appliance": "Appliance",
    }
    row = 5
    for item in SAMPLE_FFE:
        if item[2] in fixture_cats:
            (iid, room, cat, name, desc, qty, dims, mat, finish, color, vendor, brand,
             link, sku, unit, tax_rate, ship, disc, lead, order_st, appr, deliv, install, notes) = item
            vals = [
                room, cat_to_type.get(cat, "Other"), name, finish, dims, vendor, sku,
                qty, unit, None, "", "", "", notes,
            ]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=row, column=c, value=v)
                style_data_cell(cell, money=(c in (9, 10)), date=(c in (11, 12, 13)), center_text=(c == 8))
            ws.cell(row=row, column=10, value=f"=H{row}*I{row}")
            ws.cell(row=row, column=10).number_format = money_format()
            ws.cell(row=row, column=10).border = thin
            row += 1

    for r in range(row, 40):
        for c in range(1, 15):
            style_data_cell(ws.cell(row=r, column=c), money=(c in (9, 10)), date=(c in (11, 12, 13)))
        ws.cell(row=r, column=10, value=f"=IF(H{r}=\"\",\"\",H{r}*I{r})")
        ws.cell(row=r, column=10).number_format = money_format()

    add_dropdown(ws, [r[0] for r in SAMPLE_ROOMS], "A5:A39")
    add_dropdown(ws, FIXTURE_TYPES, "B5:B39")
    add_dropdown(ws, [v[0] for v in SAMPLE_VENDORS], "F5:F39")
    freeze_and_filter(ws, 4, 14, 39)
    return ws


def build_finishes(wb):
    ws = wb.create_sheet("08 Finishes", 8)
    headers = [
        "Room", "Surface/Application", "Finish Type", "Product Name", "Color/Finish",
        "Material", "Vendor", "Product Link", "SKU", "Quantity Needed", "Unit Cost",
        "Total Cost", "Sample Ordered", "Sample Approved", "Installer Notes",
    ]
    add_title_banner(ws, "FINISHES SCHEDULE", "Paint, tile, flooring, wallpaper, countertops, fabrics, and more", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [16, 16, 14, 20, 14, 12, 14, 20, 12, 12, 12, 12, 14, 14, 28])

    sample_finishes = [
        ("Living Room", "Wall", "Paint", "Simply White OC-117", "Simply White", "Latex", "Benjamin Moore", "https://example.com/paint", "OC-117", 5, 65, "Yes", "Yes", "Eggshell finish"),
        ("Living Room", "Floor", "Flooring", "White Oak Engineered", "Natural", "White Oak", "Wide Plank", "https://example.com/floor", "WO-7", 320, 12.5, "Yes", "Yes", "7\" width, matte"),
        ("Kitchen", "Countertop", "Countertop", "Calacatta Quartz", "Calacatta", "Quartz", "Caesarstone", "https://example.com/quartz", "5143", 45, 85, "Yes", "Yes", "Mitered island edge"),
        ("Kitchen", "Backsplash", "Tile", "Handmade Zellige", "Warm White", "Ceramic", "Cle Tile", "https://example.com/zellige", "ZEL-WW", 80, 28, "Yes", "Pending", "2x6 vertical stack"),
        ("Master Bedroom", "Wall", "Paint", "Agreeable Gray", "Agreeable Gray", "Latex", "Sherwin Williams", "https://example.com/ag", "SW 7029", 3, 55, "Yes", "Yes", "Flat on walls"),
        ("Powder Room", "Wall", "Wallpaper", "Grasscloth Weave", "Natural Flax", "Grasscloth", "Phillip Jeffries", "https://example.com/pj", "PJ-4410", 6, 120, "Yes", "Yes", "Pre-pasted liner required"),
        ("Kitchen", "Cabinet", "Cabinet Finish", "Shaker Painted", "Soft White", "Maple", "Custom Millwork", "", "CM-SW", 1, 8500, "No", "No", "Soft-close all doors"),
        ("Dining Room", "Drapery", "Fabric", "Linen Sheer", "Oatmeal", "Linen", "Romano", "https://example.com/linen", "LN-OAT", 18, 42, "Yes", "Yes", "Double width panels"),
    ]
    for i, f in enumerate(sample_finishes):
        r = 5 + i
        room, surface, ftype, pname, color, mat, vendor, link, sku, qty, unit, samp_o, samp_a, notes = f
        vals = [room, surface, ftype, pname, color, mat, vendor, link, sku, qty, unit, None, samp_o, samp_a, notes]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, money=(c in (11, 12)), center_text=(c in (10, 13, 14)))
        ws.cell(row=r, column=12, value=f"=J{r}*K{r}")
        ws.cell(row=r, column=12).number_format = money_format()
        ws.cell(row=r, column=12).border = thin

    for r in range(5 + len(sample_finishes), 40):
        for c in range(1, 16):
            style_data_cell(ws.cell(row=r, column=c), money=(c in (11, 12)))
        ws.cell(row=r, column=12, value=f"=IF(J{r}=\"\",\"\",J{r}*K{r})")
        ws.cell(row=r, column=12).number_format = money_format()

    add_dropdown(ws, [r[0] for r in SAMPLE_ROOMS], "A5:A39")
    add_dropdown(ws, SURFACES, "B5:B39")
    add_dropdown(ws, FINISH_TYPES, "C5:C39")
    add_dropdown(ws, YES_NO, "M5:M39")
    add_dropdown(ws, YES_NO + ["Pending"], "N5:N39")
    freeze_and_filter(ws, 4, 15, 39)
    return ws


def build_sourcing(wb):
    ws = wb.create_sheet("09 Sourcing", 9)
    headers = [
        "Room", "Item Needed", "Option 1 Link", "Option 1 Price", "Option 2 Link",
        "Option 2 Price", "Option 3 Link", "Option 3 Price", "Preferred Option",
        "Pros", "Cons", "Client Feedback", "Final Selection",
    ]
    add_title_banner(ws, "PRODUCT SOURCING TRACKER", "Compare options before locking the final FF&E selection", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [16, 22, 22, 12, 22, 12, 22, 12, 14, 28, 28, 24, 20])

    samples = [
        ("Dining Room", "Dining Table", "https://example.com/t1", 2499, "https://example.com/t2", 1899, "https://example.com/t3", 3200, "Option 1", "Solid oak, extends", "Long lead time", "Likes option 1 scale", ""),
        ("Master Bedroom", "Platform Bed", "https://example.com/b1", 1499, "https://example.com/b2", 1799, "https://example.com/b3", 1299, "Option 1", "Performance fabric", "Only 2 finishes", "Wants oatmeal", "Article Ceni King"),
        ("Home Office", "Desk Chair", "https://example.com/c1", 1395, "https://example.com/c2", 850, "https://example.com/c3", 450, "Option 2", "Good ergonomics, lower cost", "Less brand prestige", "Budget conscious", ""),
        ("Living Room", "Accent Chair", "https://example.com/a1", 699, "https://example.com/a2", 899, "https://example.com/a3", 549, "Option 1", "Swivel, boucle", "Shows pet hair", "Approved pair", "West Elm Orbit"),
    ]
    for i, s in enumerate(samples):
        r = 5 + i
        for c, v in enumerate(s, 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, money=(c in (4, 6, 8)))

    for r in range(5 + len(samples), 35):
        for c in range(1, 14):
            style_data_cell(ws.cell(row=r, column=c), money=(c in (4, 6, 8)))

    add_dropdown(ws, [r[0] for r in SAMPLE_ROOMS], "A5:A34")
    add_dropdown(ws, ["Option 1", "Option 2", "Option 3", "None / Keep Looking"], "I5:I34")
    freeze_and_filter(ws, 4, 13, 34)
    return ws


def build_vendors(wb):
    ws = wb.create_sheet("10 Vendors", 10)
    headers = [
        "Vendor Name", "Vendor Type", "Contact Person", "Email", "Phone", "Website",
        "Account Number", "Trade Discount %", "Lead Time", "Shipping Terms",
        "Return Policy", "Notes",
    ]
    add_title_banner(ws, "VENDOR & SUPPLIER DATABASE", "Trade accounts, discounts, and lead times in one place", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [18, 16, 16, 24, 16, 24, 14, 14, 12, 14, 18, 30])

    for i, v in enumerate(SAMPLE_VENDORS):
        r = 5 + i
        for c, val in enumerate(v, 1):
            cell = ws.cell(row=r, column=c, value=val)
            style_data_cell(cell, pct=(c == 8), center_text=(c == 8))
            if c == 8:
                cell.value = val / 100.0
                cell.number_format = pct_format()

    for r in range(5 + len(SAMPLE_VENDORS), 40):
        for c in range(1, 13):
            style_data_cell(ws.cell(row=r, column=c), pct=(c == 8))

    add_dropdown(ws, VENDOR_TYPES, "B5:B39")
    freeze_and_filter(ws, 4, 12, 39)
    return ws


def build_budget(wb):
    ws = wb.create_sheet("11 Budget", 11)
    headers = [
        "Budget Category", "Budgeted Amount", "Estimated Amount", "Actual Amount",
        "Difference", "Paid Amount", "Balance Due", "Payment Status", "Notes",
    ]
    add_title_banner(ws, "BUDGET TRACKER", "Category-level budget control with auto totals and variance", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [20, 16, 16, 14, 14, 14, 14, 14, 30])

    # Sample budget allocations
    budget_data = {
        "Furniture": 35000,
        "Lighting": 8000,
        "Rugs": 4000,
        "Decor": 2500,
        "Artwork": 3000,
        "Window Treatments": 4500,
        "Flooring": 9000,
        "Paint/Wallcovering": 2500,
        "Contractor Labor": 5000,
        "Installation": 2500,
        "Shipping": 1500,
        "Taxes": 4000,
        "Design Fee": 12000,
        "Contingency": 5000,
    }
    # Map FFE categories roughly for estimated via SUMIF where possible
    cat_map = {
        "Furniture": ["Sofa", "Chair", "Table", "Bed"],
        "Lighting": ["Lighting"],
        "Rugs": ["Rug"],
        "Decor": ["Decor", "Mirror"],
        "Artwork": ["Artwork"],
        "Window Treatments": ["Window Treatment"],
        "Flooring": ["Flooring"],
        "Paint/Wallcovering": ["Wallcovering"],
        "Hardware": ["Hardware"],  # not in budget list as separate - skip
    }

    for i, cat in enumerate(BUDGET_CATEGORIES):
        r = 5 + i
        budgeted = budget_data.get(cat, 0)
        ws.cell(row=r, column=1, value=cat).border = thin
        ws.cell(row=r, column=2, value=budgeted).number_format = money_format()
        ws.cell(row=r, column=2).border = thin

        # Estimated from FF&E where mappable
        if cat == "Furniture":
            formula = '=SUMIF(\'05 Master FF&E\'!C:C,"Sofa",\'05 Master FF&E\'!T:T)+SUMIF(\'05 Master FF&E\'!C:C,"Chair",\'05 Master FF&E\'!T:T)+SUMIF(\'05 Master FF&E\'!C:C,"Table",\'05 Master FF&E\'!T:T)+SUMIF(\'05 Master FF&E\'!C:C,"Bed",\'05 Master FF&E\'!T:T)'
        elif cat == "Lighting":
            formula = '=SUMIF(\'05 Master FF&E\'!C:C,"Lighting",\'05 Master FF&E\'!T:T)'
        elif cat == "Rugs":
            formula = '=SUMIF(\'05 Master FF&E\'!C:C,"Rug",\'05 Master FF&E\'!T:T)'
        elif cat == "Artwork":
            formula = '=SUMIF(\'05 Master FF&E\'!C:C,"Artwork",\'05 Master FF&E\'!T:T)'
        elif cat == "Shipping":
            formula = "=SUM('05 Master FF&E'!R5:R84)"
        elif cat == "Taxes":
            formula = "=SUM('05 Master FF&E'!Q5:Q84)"
        elif cat == "Design Fee":
            formula = "='03 Client Profile'!C28"
        else:
            formula = 0
        ws.cell(row=r, column=3, value=formula if isinstance(formula, str) else formula)
        ws.cell(row=r, column=3).number_format = money_format()
        ws.cell(row=r, column=3).border = thin

        # Actual starts equal to estimated for demo (user updates)
        ws.cell(row=r, column=4, value=f"=C{r}")
        ws.cell(row=r, column=4).number_format = money_format()
        ws.cell(row=r, column=4).border = thin

        # Difference = Budgeted - Actual
        ws.cell(row=r, column=5, value=f"=B{r}-D{r}")
        ws.cell(row=r, column=5).number_format = money_format()
        ws.cell(row=r, column=5).border = thin

        paid = budgeted * 0.3 if cat == "Design Fee" else 0
        ws.cell(row=r, column=6, value=paid).number_format = money_format()
        ws.cell(row=r, column=6).border = thin

        ws.cell(row=r, column=7, value=f"=D{r}-F{r}")
        ws.cell(row=r, column=7).number_format = money_format()
        ws.cell(row=r, column=7).border = thin

        status = "Partially Paid" if paid else "Unpaid"
        ws.cell(row=r, column=8, value=status).border = thin
        ws.cell(row=r, column=9, value="").border = thin

    # Totals
    tr = 5 + len(BUDGET_CATEGORIES)
    ws.cell(row=tr, column=1, value="TOTALS").font = font(bold=True, color=WHITE)
    ws.cell(row=tr, column=1).fill = fill(NAVY)
    for col, letter in [(2, "B"), (3, "C"), (4, "D"), (5, "E"), (6, "F"), (7, "G")]:
        cell = ws.cell(row=tr, column=col, value=f"=SUM({letter}5:{letter}{tr-1})")
        cell.font = font(bold=True, color=WHITE)
        cell.fill = fill(NAVY)
        cell.number_format = money_format()

    # Summary box
    sr = tr + 2
    ws.cell(row=sr, column=1, value="BUDGET SUMMARY").font = section_font()
    summaries = [
        ("Total Budget", f"=B{tr}"),
        ("Total Estimated", f"=C{tr}"),
        ("Total Actual", f"=D{tr}"),
        ("Over/Under Budget", f"=B{tr}-D{tr}"),
        ("Remaining Balance Due", f"=G{tr}"),
        ("Total Paid", f"=F{tr}"),
        ("Paid vs Unpaid %", f"=IF(D{tr}=0,0,F{tr}/D{tr})"),
    ]
    for i, (label, formula) in enumerate(summaries):
        ws.cell(row=sr + 1 + i, column=1, value=label).font = font(bold=True)
        ws.cell(row=sr + 1 + i, column=1).fill = fill(LIGHT_GRAY)
        cell = ws.cell(row=sr + 1 + i, column=2, value=formula)
        cell.font = font(bold=True, size=12, color=NAVY)
        if "Paid vs" in label:
            cell.number_format = pct_format()
        else:
            cell.number_format = money_format()
        cell.border = thin

    add_dropdown(ws, PAYMENT_STATUSES, f"H5:H{tr-1}")
    apply_status_conditional_formatting(ws, "H", 5, tr - 1)
    ws.conditional_formatting.add(
        f"E5:E{tr-1}",
        CellIsRule(operator="lessThan", formula=["0"], fill=fill(RED_BG), font=font(color=RED, bold=True)),
    )
    ws.conditional_formatting.add(
        f"E5:E{tr-1}",
        CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=fill(GREEN_BG), font=font(color=GREEN, bold=True)),
    )
    freeze_and_filter(ws, 4, 9, tr - 1)
    return ws


def build_purchase_orders(wb):
    ws = wb.create_sheet("12 Purchase Orders", 12)
    headers = [
        "PO Number", "Vendor", "Room", "Item", "Quantity", "Order Date",
        "Expected Ship Date", "Expected Delivery Date", "Actual Delivery Date",
        "Order Status", "Tracking Number", "Carrier", "Invoice Number",
        "Amount", "Paid Status", "Notes",
    ]
    add_title_banner(ws, "PURCHASE ORDER / ORDER TRACKER", "Professional order log with tracking and payment status", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [12, 16, 14, 22, 10, 12, 14, 14, 14, 14, 16, 12, 14, 12, 12, 24])

    pos = [
        ("PO-1001", "Article", "Living Room", "Performance Sofa", 1, datetime(2026, 4, 2), datetime(2026, 5, 15), datetime(2026, 5, 22), None, "Ordered", "", "Freight", "INV-pending", 2040.92, "Unpaid", ""),
        ("PO-1002", "West Elm", "Living Room", "Accent Lounge Chairs x2", 2, datetime(2026, 4, 5), datetime(2026, 4, 28), datetime(2026, 5, 5), None, "Shipped", "1Z999AA10123456784", "UPS", "INV-WE-8891", 1459.84, "Paid", ""),
        ("PO-1003", "Loloi", "Living Room", "Area Rug 9x12", 1, datetime(2026, 3, 20), datetime(2026, 4, 5), datetime(2026, 4, 10), datetime(2026, 4, 9), "Delivered", "1Z999BB", "FedEx", "INV-LO-220", 1377.92, "Paid", "Pad separate"),
        ("PO-1004", "Wolf / Sub-Zero", "Kitchen", "Range 36\"", 1, datetime(2026, 3, 15), datetime(2026, 6, 1), datetime(2026, 6, 10), None, "In Production", "", "Freight", "", 9858.92, "Partially Paid", "Deposit paid"),
        ("PO-1005", "Visual Comfort", "Kitchen", "Island Pendants x3", 3, datetime(2026, 4, 10), datetime(2026, 5, 20), datetime(2026, 5, 28), None, "Confirmed", "", "UPS", "", 1775.28, "Unpaid", ""),
        ("PO-1006", "Rejuvenation", "Kitchen", "Cabinet Pulls x24", 24, datetime(2026, 4, 12), datetime(2026, 4, 25), datetime(2026, 4, 28), None, "Ordered", "", "USPS", "", 491.56, "Unpaid", ""),
        ("PO-1007", "Schoolhouse", "Master Bedroom", "Pendant Pair", 2, datetime(2026, 4, 8), datetime(2026, 5, 1), datetime(2026, 5, 8), None, "Ordered", "", "UPS", "", 652.64, "Unpaid", ""),
    ]
    for i, po in enumerate(pos):
        r = 5 + i
        for c, v in enumerate(po, 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, money=(c == 14), date=(c in (6, 7, 8, 9)), center_text=(c in (5, 10, 15)))

    for r in range(5 + len(pos), 40):
        for c in range(1, 17):
            style_data_cell(ws.cell(row=r, column=c), money=(c == 14), date=(c in (6, 7, 8, 9)))

    add_dropdown(ws, [v[0] for v in SAMPLE_VENDORS], "B5:B39")
    add_dropdown(ws, [r[0] for r in SAMPLE_ROOMS], "C5:C39")
    add_dropdown(ws, PO_STATUSES, "J5:J39")
    add_dropdown(ws, PAYMENT_STATUSES, "O5:O39")
    apply_status_conditional_formatting(ws, "J", 5, 39)
    apply_status_conditional_formatting(ws, "O", 5, 39)
    freeze_and_filter(ws, 4, 16, 39)
    return ws


def build_delivery(wb):
    ws = wb.create_sheet("13 Delivery & Install", 13)
    headers = [
        "Item", "Room", "Vendor", "Delivery Address", "Expected Delivery Date",
        "Actual Delivery Date", "Delivery Status", "Received By", "Damage Check",
        "Damage Notes", "Storage Location", "Install Date", "Installer",
        "Installation Status", "Punch List Notes",
    ]
    add_title_banner(ws, "DELIVERY & INSTALLATION TRACKER", "Keep every delivery and install coordinated and documented", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [22, 14, 14, 28, 14, 14, 14, 14, 14, 20, 16, 12, 16, 14, 24])

    deliveries = [
        ("Area Rug 9x12", "Living Room", "Loloi", "482 Maple Lane, Portland, OR", datetime(2026, 4, 10), datetime(2026, 4, 9), "Delivered", "J. Blake", "Good Condition", "", "On site living", datetime(2026, 4, 11), "Studio Team", "Installed", ""),
        ("Accent Lounge Chairs", "Living Room", "West Elm", "482 Maple Lane, Portland, OR", datetime(2026, 5, 5), None, "Shipped", "", "Not Checked", "", "TBD", datetime(2026, 5, 12), "Studio Team", "Scheduled", ""),
        ("Performance Sofa", "Living Room", "Article", "482 Maple Lane, Portland, OR", datetime(2026, 5, 22), None, "Ordered", "", "Not Checked", "", "TBD", datetime(2026, 5, 25), "Article White Glove", "Not Scheduled", "Need elevator reservation"),
        ("Vessel Sink", "Powder Room", "Signature Hardware", "482 Maple Lane, Portland, OR", datetime(2026, 4, 15), datetime(2026, 4, 14), "Delivered", "GC Lead", "Good Condition", "", "GC locked storage", datetime(2026, 5, 1), "Plumber", "Scheduled", ""),
        ("Range 36\"", "Kitchen", "Wolf / Sub-Zero", "482 Maple Lane, Portland, OR", datetime(2026, 6, 10), None, "Ordered", "", "Not Checked", "", "Dealer warehouse", datetime(2026, 6, 15), "Appliance Installer", "Not Scheduled", "Confirm gas rough-in"),
        ("Floor Lamp", "Living Room", "Rejuvenation", "482 Maple Lane, Portland, OR", datetime(2026, 4, 1), datetime(2026, 3, 30), "Delivered", "J. Blake", "Good Condition", "", "On site", datetime(2026, 4, 2), "Studio Team", "Installed", ""),
    ]
    for i, d in enumerate(deliveries):
        r = 5 + i
        for c, v in enumerate(d, 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, date=(c in (5, 6, 12)), center_text=(c in (7, 9, 14)))

    for r in range(5 + len(deliveries), 40):
        for c in range(1, 16):
            style_data_cell(ws.cell(row=r, column=c), date=(c in (5, 6, 12)))

    add_dropdown(ws, [r[0] for r in SAMPLE_ROOMS], "B5:B39")
    add_dropdown(ws, [v[0] for v in SAMPLE_VENDORS], "C5:C39")
    add_dropdown(ws, DELIVERY_STATUSES, "G5:G39")
    add_dropdown(ws, DAMAGE_CHECKS, "I5:I39")
    add_dropdown(ws, INSTALL_STATUSES, "N5:N39")
    apply_status_conditional_formatting(ws, "G", 5, 39)
    apply_status_conditional_formatting(ws, "I", 5, 39)
    apply_status_conditional_formatting(ws, "N", 5, 39)
    freeze_and_filter(ws, 4, 15, 39)
    return ws


def build_approvals(wb):
    ws = wb.create_sheet("14 Client Approvals", 14)
    headers = [
        "Item/Design Decision", "Room", "Approval Type", "Sent to Client Date",
        "Client Response Date", "Status", "Approved Option", "Client Comments",
        "Revision Needed", "Final Approval Date",
    ]
    add_title_banner(ws, "CLIENT APPROVAL TRACKER", "Every design decision and purchase sign-off in one place", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [26, 14, 14, 14, 14, 16, 20, 28, 14, 14])

    approvals = [
        ("Living room seating package", "Living Room", "Furniture", datetime(2026, 3, 18), datetime(2026, 3, 22), "Approved", "Article sofa + WE chairs", "Love the stone fabric", "No", datetime(2026, 3, 22)),
        ("Living room rug", "Living Room", "Furniture", datetime(2026, 3, 10), datetime(2026, 3, 12), "Approved", "Loloi Magnolia 9x12", "", "No", datetime(2026, 3, 12)),
        ("Kitchen range selection", "Kitchen", "Purchase", datetime(2026, 3, 5), datetime(2026, 3, 8), "Approved", "Wolf DF36450", "Proceed with stainless", "No", datetime(2026, 3, 8)),
        ("Primary bedroom bed", "Master Bedroom", "Furniture", datetime(2026, 4, 1), None, "Sent", "", "Waiting on fabric sample photos", "Yes", None),
        ("Dining table options", "Dining Room", "Furniture", datetime(2026, 4, 5), None, "Sent", "", "Comparing 3 options in Sourcing tab", "Yes", None),
        ("Overall color palette", "Living Room", "Color Palette", datetime(2026, 2, 20), datetime(2026, 2, 25), "Approved", "Warm neutrals + brass", "Approved board v2", "No", datetime(2026, 2, 25)),
        ("Powder room wallpaper", "Powder Room", "Finish", datetime(2026, 3, 28), datetime(2026, 4, 2), "Approved", "Phillip Jeffries grasscloth", "", "No", datetime(2026, 4, 2)),
        ("Project budget confirmation", "—", "Budget", datetime(2026, 2, 15), datetime(2026, 2, 18), "Approved", "$85,000 total", "OK with contingency", "No", datetime(2026, 2, 18)),
        ("Concept presentation", "—", "Design Concept", datetime(2026, 2, 10), datetime(2026, 2, 14), "Approved", "Modern Organic concept", "Minor layout tweaks only", "No", datetime(2026, 2, 14)),
    ]
    for i, a in enumerate(approvals):
        r = 5 + i
        for c, v in enumerate(a, 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, date=(c in (4, 5, 10)), center_text=(c in (3, 6, 9)))

    for r in range(5 + len(approvals), 40):
        for c in range(1, 11):
            style_data_cell(ws.cell(row=r, column=c), date=(c in (4, 5, 10)))

    add_dropdown(ws, [r[0] for r in SAMPLE_ROOMS] + ["—"], "B5:B39")
    add_dropdown(ws, APPROVAL_TYPES, "C5:C39")
    add_dropdown(ws, APPROVAL_STATUSES, "F5:F39")
    add_dropdown(ws, YES_NO, "I5:I39")
    apply_status_conditional_formatting(ws, "F", 5, 39)
    freeze_and_filter(ws, 4, 10, 39)
    return ws


def build_timeline(wb):
    ws = wb.create_sheet("15 Timeline", 15)
    headers = [
        "Phase", "Task", "Assigned To", "Start Date", "Due Date", "Completion Date",
        "Status", "Priority", "Dependencies", "Notes",
        "W1", "W2", "W3", "W4", "W5", "W6", "W7", "W8", "W9", "W10", "W11", "W12",
    ]
    add_title_banner(ws, "PROJECT TIMELINE / GANTT TRACKER", "Phased plan with a simple week-by-week Gantt view (mark cells with X)", 12)
    write_headers(ws, 4, headers)
    widths = [16, 28, 14, 12, 12, 14, 14, 10, 16, 24] + [4] * 12
    set_col_widths(ws, widths)

    tasks = [
        ("Consultation", "Kickoff meeting & questionnaire", "Jordan Blake", datetime(2026, 3, 1), datetime(2026, 3, 3), datetime(2026, 3, 3), "Complete", "High", "", "", [1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
        ("Measurements", "Site measure all rooms", "Studio Team", datetime(2026, 3, 5), datetime(2026, 3, 7), datetime(2026, 3, 7), "Complete", "High", "Kickoff", "", [0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
        ("Concept Design", "Mood boards & space plans", "Jordan Blake", datetime(2026, 3, 8), datetime(2026, 3, 20), datetime(2026, 3, 18), "Complete", "High", "Measurements", "", [0, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0]),
        ("Client Approval", "Concept presentation", "Jordan Blake", datetime(2026, 3, 20), datetime(2026, 3, 25), datetime(2026, 3, 22), "Complete", "Urgent", "Concept Design", "", [0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0]),
        ("Sourcing", "Source living room package", "Studio Team", datetime(2026, 3, 25), datetime(2026, 4, 15), None, "In Progress", "High", "Concept approval", "", [0, 0, 0, 1, 1, 1, 0, 0, 0, 0, 0, 0]),
        ("Sourcing", "Source kitchen fixtures & appliances", "Studio Team", datetime(2026, 3, 20), datetime(2026, 4, 10), datetime(2026, 4, 8), "Complete", "Urgent", "", "Range ordered", [0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0]),
        ("Client Approval", "Approve primary suite furniture", "Client", datetime(2026, 4, 1), datetime(2026, 4, 15), None, "Waiting on Client", "High", "Sourcing", "", [0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0]),
        ("Purchasing", "Place living room POs", "Jordan Blake", datetime(2026, 4, 1), datetime(2026, 4, 12), datetime(2026, 4, 10), "Complete", "High", "Approvals", "", [0, 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0]),
        ("Construction", "Kitchen rough-in coordination", "GC", datetime(2026, 5, 1), datetime(2026, 5, 20), None, "Not Started", "High", "Appliance order", "", [0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0, 0]),
        ("Delivery", "Living room furniture delivery", "Vendors", datetime(2026, 5, 15), datetime(2026, 5, 25), None, "Not Started", "High", "POs", "", [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 0, 0]),
        ("Installation", "Install living + powder", "Installers", datetime(2026, 5, 20), datetime(2026, 6, 5), None, "Not Started", "High", "Deliveries", "", [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0]),
        ("Styling", "Final styling pass", "Jordan Blake", datetime(2026, 6, 10), datetime(2026, 6, 15), None, "Not Started", "Medium", "Installation", "", [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1]),
        ("Final Reveal", "Client walkthrough", "Jordan Blake", datetime(2026, 6, 18), datetime(2026, 6, 18), None, "Not Started", "Urgent", "Styling", "", [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1]),
        ("Punch List", "Complete punch items", "Studio + GC", datetime(2026, 6, 18), datetime(2026, 6, 30), None, "Not Started", "High", "Reveal", "", [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1]),
    ]
    for i, t in enumerate(tasks):
        r = 5 + i
        *fields, gantt = t
        for c, v in enumerate(fields, 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, date=(c in (4, 5, 6)), center_text=(c in (7, 8)))
        for g, mark in enumerate(gantt):
            cell = ws.cell(row=r, column=11 + g, value="█" if mark else "")
            cell.alignment = center()
            cell.border = thin
            if mark:
                cell.fill = fill(TEAL)
                cell.font = font(color=TEAL, size=10)

    for r in range(5 + len(tasks), 40):
        for c in range(1, 23):
            style_data_cell(ws.cell(row=r, column=c), date=(c in (4, 5, 6)))

    add_dropdown(ws, PROJECT_PHASES, "A5:A39")
    add_dropdown(ws, TASK_STATUSES, "G5:G39")
    add_dropdown(ws, PRIORITY_LEVELS, "H5:H39")
    apply_status_conditional_formatting(ws, "G", 5, 39)
    apply_status_conditional_formatting(ws, "H", 5, 39)
    freeze_and_filter(ws, 4, 10, 39)
    return ws


def build_todo(wb):
    ws = wb.create_sheet("16 To-Do List", 16)
    headers = ["Task", "Room", "Category", "Assigned To", "Priority", "Due Date", "Status", "Notes"]
    add_title_banner(ws, "TASK & TO-DO LIST", "Daily working list for the design team", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [36, 16, 14, 14, 12, 12, 16, 32])

    todos = [
        ("Email client fabric photos for primary bed", "Master Bedroom", "Client", "Jordan Blake", "High", datetime(2026, 4, 12), "In Progress", ""),
        ("Confirm Wolf range rough-in dimensions with GC", "Kitchen", "Site", "Studio Team", "Urgent", datetime(2026, 4, 10), "Waiting on Vendor", "Need GC response"),
        ("Order rug pad for living room", "Living Room", "Purchasing", "Studio Team", "Medium", datetime(2026, 4, 15), "Not Started", ""),
        ("Schedule elevator for sofa delivery", "Living Room", "Installation", "Jordan Blake", "High", datetime(2026, 5, 10), "Not Started", "Building mgmt 48h notice"),
        ("Update budget actuals after West Elm invoice", "—", "Admin", "Studio Team", "Medium", datetime(2026, 4, 8), "Complete", ""),
        ("Request trade quote Four Hands dining chairs", "Dining Room", "Sourcing", "Studio Team", "High", datetime(2026, 4, 14), "In Progress", ""),
        ("Send dining table options board to client", "Dining Room", "Client", "Jordan Blake", "High", datetime(2026, 4, 11), "Waiting on Client", "Board sent 4/5"),
        ("Book site visit for powder install", "Powder Room", "Installation", "Studio Team", "Medium", datetime(2026, 4, 20), "Not Started", ""),
        ("Collect paint touch-up kit from BM", "Living Room", "Vendor", "Studio Team", "Low", datetime(2026, 4, 30), "Not Started", ""),
        ("Review overdue open POs", "—", "Purchasing", "Jordan Blake", "Urgent", datetime(2026, 4, 9), "Delayed", "Check Visual Comfort"),
    ]
    for i, t in enumerate(todos):
        r = 5 + i
        for c, v in enumerate(t, 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, date=(c == 6), center_text=(c in (5, 7)))

    for r in range(5 + len(todos), 50):
        for c in range(1, 9):
            style_data_cell(ws.cell(row=r, column=c), date=(c == 6))

    add_dropdown(ws, [r[0] for r in SAMPLE_ROOMS] + ["—"], "B5:B49")
    add_dropdown(ws, TASK_CATEGORIES, "C5:C49")
    add_dropdown(ws, PRIORITY_LEVELS, "E5:E49")
    add_dropdown(ws, TASK_STATUSES, "G5:G49")
    apply_status_conditional_formatting(ws, "E", 5, 49)
    apply_status_conditional_formatting(ws, "G", 5, 49)
    # Overdue highlight on due date when status not complete
    ws.conditional_formatting.add(
        "F5:F49",
        FormulaRule(
            formula=['AND(F5<>"",F5<TODAY(),G5<>"Complete")'],
            fill=fill(RED_BG),
            font=font(color=RED, bold=True),
        ),
    )
    freeze_and_filter(ws, 4, 8, 49)
    return ws


def build_measurements(wb):
    ws = wb.create_sheet("17 Measurements", 17)
    headers = [
        "Room", "Measurement Type", "Width", "Height", "Depth", "Length",
        "Ceiling Height", "Window Size", "Door Size", "Wall Length", "Notes",
    ]
    add_title_banner(ws, "MEASUREMENT LOG", "Field measurements for designers and installers", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [16, 16, 10, 10, 10, 10, 12, 14, 12, 12, 32])

    measures = [
        ("Living Room", "Room", "18'", "9'", "—", "16'", "9'", "—", "—", "—", "Open to dining"),
        ("Living Room", "Window", "6'", "5'", "—", "—", "—", "6' x 5'", "—", "—", "South facing, 2 panels"),
        ("Living Room", "Rug Area", "11'", "—", "—", "14'", "—", "—", "—", "—", "Allow 12–18\" border"),
        ("Master Bedroom", "Room", "14'", "9'", "—", "16'", "9'", "—", "—", "—", ""),
        ("Master Bedroom", "Furniture Space", "80\"", "—", "84\"", "—", "—", "—", "—", "—", "King bed wall"),
        ("Kitchen", "Room", "12'", "9'", "—", "16'", "9'", "—", "—", "—", "Island 8' x 3.5'"),
        ("Kitchen", "Cabinet", "36\"", "36\"", "24\"", "—", "—", "—", "—", "—", "Base run east wall"),
        ("Powder Room", "Room", "5'", "9'", "—", "8'", "9'", "—", "28\" door", "—", "Tight clearances"),
        ("Home Office", "Window", "4'", "5'", "—", "—", "—", "4' x 5'", "—", "—", "Needs treatment"),
        ("Dining Room", "Room", "12'", "9'", "—", "14'", "9'", "—", "—", "—", "Chandelier center 6' from window wall"),
    ]
    for i, m in enumerate(measures):
        r = 5 + i
        for c, v in enumerate(m, 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, center_text=(c >= 3 and c <= 10))

    for r in range(5 + len(measures), 40):
        for c in range(1, 12):
            style_data_cell(ws.cell(row=r, column=c))

    add_dropdown(ws, [r[0] for r in SAMPLE_ROOMS], "A5:A39")
    add_dropdown(ws, MEASUREMENT_TYPES, "B5:B39")
    freeze_and_filter(ws, 4, 11, 39)
    return ws


def build_contractors(wb):
    ws = wb.create_sheet("18 Contractors", 18)
    headers = [
        "Contractor/Vendor Name", "Trade", "Contact", "Scope", "Quote Amount",
        "Start Date", "End Date", "Payment Status", "Insurance Received",
        "Contract Signed", "Notes",
    ]
    add_title_banner(ws, "CONTRACTOR / INSTALLER TRACKER", "Coordinate trades, quotes, insurance, and schedules", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [22, 20, 20, 28, 12, 12, 12, 14, 14, 14, 28])

    contractors = [
        ("Summit Build Co.", "General Contractor", "Mike Torres · 555-2001", "Kitchen rough-in & coordination", 8500, datetime(2026, 5, 1), datetime(2026, 6, 15), "Partially Paid", "Yes", "Yes", "COI on file"),
        ("Lumen Electric", "Electrician", "Priya Nair · 555-2002", "Island pendants + recessed", 2200, datetime(2026, 5, 10), datetime(2026, 5, 18), "Unpaid", "Yes", "Yes", ""),
        ("Cascade Plumbing", "Plumber", "Ed Walsh · 555-2003", "Powder vessel sink install", 980, datetime(2026, 5, 1), datetime(2026, 5, 3), "Unpaid", "Yes", "Yes", ""),
        ("Oak & Stone Floors", "Flooring Installer", "Lee Cho · 555-2004", "Living white oak install", 4200, datetime(2026, 4, 20), datetime(2026, 4, 28), "Paid", "Yes", "Yes", "Completed"),
        ("Paper & Paste", "Wallpaper Installer", "Nina Sol · 555-2005", "Powder grasscloth", 750, datetime(2026, 5, 5), datetime(2026, 5, 6), "Unpaid", "Yes", "No", "Send contract"),
        ("White Glove Crew", "Furniture Installer", "Dispatch · 555-2006", "Living room furniture set", 450, datetime(2026, 5, 25), datetime(2026, 5, 25), "Unpaid", "Yes", "Yes", "Article partner"),
        ("Fresh Coat Pros", "Painter", "Tom Rivera · 555-2007", "Touch-ups post install", 600, datetime(2026, 6, 10), datetime(2026, 6, 12), "Unpaid", "No", "No", "Need COI"),
    ]
    for i, cdata in enumerate(contractors):
        r = 5 + i
        for c, v in enumerate(cdata, 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, money=(c == 5), date=(c in (6, 7)), center_text=(c in (8, 9, 10)))

    for r in range(5 + len(contractors), 35):
        for c in range(1, 12):
            style_data_cell(ws.cell(row=r, column=c), money=(c == 5), date=(c in (6, 7)))

    add_dropdown(ws, TRADES, "B5:B34")
    add_dropdown(ws, PAYMENT_STATUSES, "H5:H34")
    add_dropdown(ws, YES_NO, "I5:I34")
    add_dropdown(ws, YES_NO, "J5:J34")
    apply_status_conditional_formatting(ws, "H", 5, 34)
    freeze_and_filter(ws, 4, 11, 34)
    return ws


def build_invoices(wb):
    ws = wb.create_sheet("19 Invoices", 19)
    headers = [
        "Invoice Number", "Invoice Date", "Client", "Description", "Amount",
        "Due Date", "Paid Date", "Payment Status", "Payment Method", "Notes",
    ]
    add_title_banner(ws, "INVOICE & PAYMENT TRACKER", "Client billing for design fees and reimbursable items", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [14, 12, 18, 36, 12, 12, 12, 14, 14, 28])

    invoices = [
        ("INV-2026-01", datetime(2026, 3, 1), "Alexandra Chen", "Design retainer / deposit", 4000, datetime(2026, 3, 15), datetime(2026, 3, 10), "Paid", "ACH", ""),
        ("INV-2026-02", datetime(2026, 4, 1), "Alexandra Chen", "Concept design milestone", 4000, datetime(2026, 4, 15), datetime(2026, 4, 12), "Paid", "ACH", ""),
        ("INV-2026-03", datetime(2026, 5, 1), "Alexandra Chen", "Procurement milestone", 4000, datetime(2026, 5, 15), None, "Sent", "", "Due mid-month"),
        ("INV-2026-04", datetime(2026, 3, 25), "Alexandra Chen", "Reimbursable: rug + pad", 1378, datetime(2026, 4, 8), datetime(2026, 4, 5), "Paid", "Credit Card", ""),
        ("INV-2026-05", datetime(2026, 4, 15), "Alexandra Chen", "Reimbursable: West Elm chairs", 1460, datetime(2026, 4, 30), None, "Sent", "", ""),
        ("INV-2026-06", datetime(2026, 6, 1), "Alexandra Chen", "Final design fee balance", 0, datetime(2026, 6, 15), None, "Draft", "", "Calculate at install"),
    ]
    for i, inv in enumerate(invoices):
        r = 5 + i
        for c, v in enumerate(inv, 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, money=(c == 5), date=(c in (2, 6, 7)), center_text=(c == 8))

    for r in range(5 + len(invoices), 35):
        for c in range(1, 11):
            style_data_cell(ws.cell(row=r, column=c), money=(c == 5), date=(c in (2, 6, 7)))

    tr = 36
    ws.cell(row=tr, column=4, value="TOTALS").font = font(bold=True, color=WHITE)
    ws.cell(row=tr, column=4).fill = fill(NAVY)
    ws.cell(row=tr, column=5, value="=SUM(E5:E34)").number_format = money_format()
    ws.cell(row=tr, column=5).font = font(bold=True, color=WHITE)
    ws.cell(row=tr, column=5).fill = fill(NAVY)
    ws.cell(row=tr + 1, column=4, value="Total Paid")
    ws.cell(row=tr + 1, column=5, value='=SUMIF(H5:H34,"Paid",E5:E34)').number_format = money_format()
    ws.cell(row=tr + 2, column=4, value="Outstanding")
    ws.cell(row=tr + 2, column=5, value='=SUMIF(H5:H34,"Sent",E5:E34)+SUMIF(H5:H34,"Partially Paid",E5:E34)+SUMIF(H5:H34,"Overdue",E5:E34)').number_format = money_format()

    add_dropdown(ws, PAYMENT_STATUSES, "H5:H34")
    apply_status_conditional_formatting(ws, "H", 5, 34)
    freeze_and_filter(ws, 4, 10, 34)
    return ws


def build_notes(wb):
    ws = wb.create_sheet("20 Meeting Notes", 20)
    headers = [
        "Date", "Meeting Type", "Attendees", "Discussion Points", "Decisions Made",
        "Action Items", "Follow-up Date", "Notes",
    ]
    add_title_banner(ws, "PROJECT NOTES / MEETING NOTES", "Consultations, vendor meetings, site visits, and walkthroughs", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [12, 18, 28, 36, 28, 28, 12, 28])

    notes = [
        (datetime(2026, 3, 1), "Client Consultation", "Alexandra Chen, Jordan Blake", "Goals, lifestyle, dog-friendly materials, budget $85k", "Proceed with full-service FF&E", "Send questionnaire; schedule measure", datetime(2026, 3, 3), "Client loves warm minimalism"),
        (datetime(2026, 3, 7), "Site Visit", "Jordan Blake, Studio Team", "Measured all rooms; noted powder door swing", "Use existing oak floors in living", "Draft plans; photo log", datetime(2026, 3, 10), "Good natural light south"),
        (datetime(2026, 3, 22), "Client Consultation", "Alexandra Chen, Jordan Blake", "Concept presentation — boards v2", "Approved palette and living package", "Source dining table options", datetime(2026, 4, 5), "Travel blackout Aug 1–15"),
        (datetime(2026, 4, 2), "Vendor Meeting", "Jordan Blake, Article Trade", "Sofa lead times and white-glove", "Order Sven in Stone", "Send PO-1001", datetime(2026, 4, 3), "Trade 15% applied"),
        (datetime(2026, 4, 8), "Contractor Meeting", "Mike Torres (GC), Jordan Blake", "Kitchen range rough-in & schedule", "Rough-in week of May 1", "Confirm gas line size", datetime(2026, 4, 12), "Need Wolf cut sheet"),
    ]
    for i, n in enumerate(notes):
        r = 5 + i
        for c, v in enumerate(n, 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, date=(c in (1, 7)))

    for r in range(5 + len(notes), 40):
        for c in range(1, 9):
            style_data_cell(ws.cell(row=r, column=c), date=(c in (1, 7)))

    add_dropdown(ws, MEETING_TYPES, "B5:B39")
    freeze_and_filter(ws, 4, 8, 39)
    return ws


# ---------------------------------------------------------------------------
# Printable report sheets
# ---------------------------------------------------------------------------

def build_client_ffe_report(wb):
    ws = wb.create_sheet("21 Client FF&E Report", 21)
    headers = ["Item ID", "Room", "Category", "Item Name", "Quantity", "Vendor", "Finish/Color", "Unit Cost", "Total", "Status"]
    add_title_banner(ws, "CLIENT-FACING FF&E REPORT", "Print or PDF this sheet for client presentations", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [10, 16, 14, 24, 10, 14, 16, 12, 12, 14])

    for i, item in enumerate(SAMPLE_FFE):
        r = 5 + i
        iid, room, cat, name, desc, qty, dims, mat, finish, color, vendor, brand, link, sku, unit, tax_rate, ship, disc, lead, order_st, appr, deliv, install, notes = item
        total = qty * unit
        vals = [iid, room, cat, name, qty, vendor, f"{finish} / {color}", unit, total, appr]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, money=(c in (8, 9)), center_text=(c in (5, 10)))

    tr = 5 + len(SAMPLE_FFE)
    ws.cell(row=tr, column=8, value="TOTAL").font = font(bold=True, color=WHITE)
    ws.cell(row=tr, column=8).fill = fill(NAVY)
    ws.cell(row=tr, column=9, value=f"=SUM(I5:I{tr-1})").number_format = money_format()
    ws.cell(row=tr, column=9).font = font(bold=True, color=WHITE)
    ws.cell(row=tr, column=9).fill = fill(NAVY)

    ws.print_title_rows = "1:4"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    freeze_and_filter(ws, 4, 10, tr - 1)
    return ws


def build_room_budget_report(wb):
    ws = wb.create_sheet("22 Room Budget Report", 22)
    headers = ["Room", "Budget Allocation", "Estimated Spend", "Actual Spend", "Remaining", "% Used", "Status"]
    add_title_banner(ws, "ROOM-BY-ROOM BUDGET REPORT", "Printable budget summary by area", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [18, 16, 16, 14, 14, 10, 14])

    for i in range(6):
        r = 5 + i
        ws.cell(row=r, column=1, value=f"='04 Room List'!A{r}").border = thin
        for col, src in [(2, "F"), (3, "G"), (4, "H"), (5, "I")]:
            cell = ws.cell(row=r, column=col, value=f"='04 Room List'!{src}{r}")
            cell.number_format = money_format()
            cell.border = thin
        ws.cell(row=r, column=6, value=f"=IF(B{r}=0,0,D{r}/B{r})").number_format = pct_format()
        ws.cell(row=r, column=6).border = thin
        ws.cell(row=r, column=7, value=f"='04 Room List'!E{r}").border = thin

    tr = 12
    ws.cell(row=tr, column=1, value="TOTAL").font = font(bold=True, color=WHITE)
    ws.cell(row=tr, column=1).fill = fill(NAVY)
    for col, letter in [(2, "B"), (3, "C"), (4, "D"), (5, "E")]:
        cell = ws.cell(row=tr, column=col, value=f"=SUM({letter}5:{letter}10)")
        cell.number_format = money_format()
        cell.font = font(bold=True, color=WHITE)
        cell.fill = fill(NAVY)

    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    return ws


def build_vendor_order_report(wb):
    ws = wb.create_sheet("23 Vendor Order Report", 23)
    headers = ["PO Number", "Vendor", "Item", "Qty", "Order Date", "Status", "Amount", "Paid Status", "Tracking"]
    add_title_banner(ws, "VENDOR ORDER REPORT", "Printable open and completed orders by vendor", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [12, 16, 28, 8, 12, 14, 12, 12, 20])

    for i in range(7):
        r = 5 + i
        src = 5 + i
        mapping = [(1, "A"), (2, "B"), (3, "D"), (4, "E"), (5, "F"), (6, "J"), (7, "N"), (8, "O"), (9, "K")]
        for col, letter in mapping:
            cell = ws.cell(row=r, column=col, value=f"='12 Purchase Orders'!{letter}{src}")
            style_data_cell(cell, money=(col == 7), date=(col == 5))

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    freeze_and_filter(ws, 4, 9, 20)
    return ws


def build_install_checklist(wb):
    ws = wb.create_sheet("24 Install Checklist", 24)
    headers = ["☐", "Item", "Room", "Vendor", "Install Date", "Installer", "Status", "Notes"]
    add_title_banner(ws, "INSTALLATION CHECKLIST", "Print for on-site installation day", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [6, 26, 14, 14, 12, 18, 14, 30])

    items = [
        ("Area Rug 9x12", "Living Room", "Loloi", datetime(2026, 4, 11), "Studio Team", "Installed", ""),
        ("Floor Lamp", "Living Room", "Rejuvenation", datetime(2026, 4, 2), "Studio Team", "Installed", ""),
        ("Accent Lounge Chairs", "Living Room", "West Elm", datetime(2026, 5, 12), "Studio Team", "Scheduled", ""),
        ("Performance Sofa", "Living Room", "Article", datetime(2026, 5, 25), "Article White Glove", "Not Scheduled", "Elevator"),
        ("Vessel Sink", "Powder Room", "Signature Hardware", datetime(2026, 5, 1), "Plumber", "Scheduled", ""),
        ("Island Pendants x3", "Kitchen", "Visual Comfort", datetime(2026, 5, 18), "Electrician", "Not Scheduled", "After rough-in"),
        ("Range 36\"", "Kitchen", "Wolf / Sub-Zero", datetime(2026, 6, 15), "Appliance Installer", "Not Scheduled", "Gas check"),
        ("Cabinet Pulls", "Kitchen", "Rejuvenation", datetime(2026, 5, 20), "GC", "Not Scheduled", ""),
        ("Pendant Pair", "Master Bedroom", "Schoolhouse", datetime(2026, 5, 22), "Electrician", "Not Scheduled", ""),
    ]
    for i, it in enumerate(items):
        r = 5 + i
        ws.cell(row=r, column=1, value="☐" if it[5] != "Installed" else "☑").alignment = center()
        ws.cell(row=r, column=1).border = thin
        for c, v in enumerate(it, 2):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, date=(c == 5), center_text=(c == 7))

    apply_status_conditional_formatting(ws, "G", 5, 30)
    ws.page_setup.orientation = "portrait"
    freeze_and_filter(ws, 4, 8, 30)
    return ws


def build_punch_list(wb):
    ws = wb.create_sheet("25 Punch List", 25)
    headers = ["☐", "#", "Item / Issue", "Room", "Responsible", "Priority", "Due Date", "Status", "Resolution Notes"]
    add_title_banner(ws, "FINAL PUNCH LIST", "Print for final walkthrough and close-out", len(headers))
    write_headers(ws, 4, headers)
    set_col_widths(ws, [6, 6, 36, 14, 16, 10, 12, 14, 30])

    punches = [
        (1, "Touch up paint at sofa wall corners", "Living Room", "Painter", "Medium", datetime(2026, 6, 20), "Not Started", ""),
        (2, "Adjust island pendant heights", "Kitchen", "Electrician", "High", datetime(2026, 6, 18), "Not Started", "Client prefers 32\" above island"),
        (3, "Missing felt pads on dining chairs", "Dining Room", "Studio Team", "Low", datetime(2026, 6, 22), "Not Started", ""),
        (4, "Confirm warranty packet delivered", "Kitchen", "Studio Team", "Medium", datetime(2026, 6, 25), "Not Started", "Wolf range docs"),
        (5, "Replace one damaged cabinet pull", "Kitchen", "GC", "Medium", datetime(2026, 6, 20), "Not Started", "Scratch on finish"),
    ]
    for i, p in enumerate(punches):
        r = 5 + i
        ws.cell(row=r, column=1, value="☐").alignment = center()
        ws.cell(row=r, column=1).border = thin
        for c, v in enumerate(p, 2):
            cell = ws.cell(row=r, column=c, value=v)
            style_data_cell(cell, date=(c == 7), center_text=(c in (2, 6, 8)))

    for r in range(5 + len(punches), 30):
        ws.cell(row=r, column=1, value="☐").border = thin
        ws.cell(row=r, column=1).alignment = center()
        for c in range(2, 10):
            style_data_cell(ws.cell(row=r, column=c), date=(c == 7))

    add_dropdown(ws, PRIORITY_LEVELS, "F5:F29")
    add_dropdown(ws, TASK_STATUSES, "H5:H29")
    apply_status_conditional_formatting(ws, "F", 5, 29)
    apply_status_conditional_formatting(ws, "H", 5, 29)
    ws.page_setup.orientation = "portrait"
    freeze_and_filter(ws, 4, 9, 29)
    return ws


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def create_workbook(output_path):
    wb = Workbook()
    # remove default
    default = wb.active
    wb.remove(default)

    print("Building Lists...")
    build_lists_sheet(wb)
    print("Building Welcome...")
    build_welcome(wb)
    print("Building Dashboard...")
    build_dashboard(wb)
    print("Building Client Profile...")
    build_client_profile(wb)
    print("Building Room List...")
    build_room_list(wb)
    print("Building Master FF&E...")
    build_master_ffe(wb)
    print("Building Furniture...")
    build_furniture(wb)
    print("Building Fixtures...")
    build_fixtures(wb)
    print("Building Finishes...")
    build_finishes(wb)
    print("Building Sourcing...")
    build_sourcing(wb)
    print("Building Vendors...")
    build_vendors(wb)
    print("Building Budget...")
    build_budget(wb)
    print("Building Purchase Orders...")
    build_purchase_orders(wb)
    print("Building Delivery...")
    build_delivery(wb)
    print("Building Approvals...")
    build_approvals(wb)
    print("Building Timeline...")
    build_timeline(wb)
    print("Building To-Do...")
    build_todo(wb)
    print("Building Measurements...")
    build_measurements(wb)
    print("Building Contractors...")
    build_contractors(wb)
    print("Building Invoices...")
    build_invoices(wb)
    print("Building Notes...")
    build_notes(wb)
    print("Building printable reports...")
    build_client_ffe_report(wb)
    build_room_budget_report(wb)
    build_vendor_order_report(wb)
    build_install_checklist(wb)
    build_punch_list(wb)

    print("Fixing cross-sheet formula references...")
    fix_dashboard_refs(wb)
    fix_dashboard_ffe_cols(wb)

    # Tab colors
    tab_colors = {
        "01 Welcome": TEAL,
        "02 Dashboard": NAVY,
        "03 Client Profile": "5C6BC0",
        "04 Room List": "26A69A",
        "05 Master FF&E": CORAL,
        "06 Furniture": "8D6E63",
        "07 Fixtures": "78909C",
        "08 Finishes": "AB47BC",
        "09 Sourcing": GOLD.replace("E9C46A", "F4A261"),
        "10 Vendors": "42A5F5",
        "11 Budget": GREEN,
        "12 Purchase Orders": "EF6C00",
        "13 Delivery & Install": "00897B",
        "14 Client Approvals": "7E57C2",
        "15 Timeline": "5C6BC0",
        "16 To-Do List": "EC407A",
        "17 Measurements": "78909C",
        "18 Contractors": "6D4C41",
        "19 Invoices": GREEN,
        "20 Meeting Notes": "546E7A",
        "21 Client FF&E Report": TEAL,
        "22 Room Budget Report": TEAL,
        "23 Vendor Order Report": TEAL,
        "24 Install Checklist": TEAL,
        "25 Punch List": TEAL,
    }
    for name, color in tab_colors.items():
        if name in wb.sheetnames:
            wb[name].sheet_properties.tabColor = color

    # Move Lists to end and keep hidden
    # Already first and hidden

    wb.properties.title = "Interior Design FF&E Project Tracker"
    wb.properties.creator = "Interior Design FF&E Tracker Generator"
    wb.properties.description = "Professional multi-tab FF&E, budget, vendor, and project management workbook for interior designers. Compatible with Excel and Google Sheets."

    wb.save(output_path)
    print(f"Saved: {output_path}")
    return output_path


if __name__ == "__main__":
    out_dir = os.path.dirname(os.path.abspath(__file__))
    out_path = os.path.join(out_dir, "Interior_Design_FFE_Project_Tracker.xlsx")
    create_workbook(out_path)
    print("Done.")
