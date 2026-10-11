#!/usr/bin/env python3
"""
Render the 15 Etsy listing images for the workbook.

    python tools/make_listing_images.py                  # -> output/etsy_package/listing_images
    python tools/make_listing_images.py -o some/folder

The images are drawn with Pillow from the same example records the workbook
ships with (``pet_sitting_spreadsheet/sample_data.py``).  They are illustrations
of the workbook's tabs and contents, not screenshots of Excel.  Requires Pillow
(``pip install pillow``).
"""

from __future__ import annotations

import argparse
import os
import sys
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

from PIL import Image, ImageDraw, ImageFont  # noqa: E402

from pet_sitting_spreadsheet import sample_data  # noqa: E402
from pet_sitting_spreadsheet.config import Config  # noqa: E402

W, H = 2000, 1500
MARGIN = 80
TOTAL = 15

TEAL = (31, 111, 107)
TEAL_DARK = (20, 78, 75)
CORAL = (228, 118, 47)
GOLD = (244, 185, 66)
CREAM = (251, 247, 240)
ZEBRA = (243, 238, 230)
WHITE = (255, 255, 255)
INK = (38, 49, 58)
GREY = (120, 128, 136)
LINE = (224, 218, 208)

FONT_DIR = "/usr/share/fonts/truetype/dejavu"
_fonts: dict = {}


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    key = (size, bold)
    if key not in _fonts:
        name = "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf"
        _fonts[key] = ImageFont.truetype(os.path.join(FONT_DIR, name), size)
    return _fonts[key]


# ---------------------------------------------------------------- formatting

def fmt_date(value) -> str:
    return f"{value:%d %b %Y}" if value else ""


def fmt_time(value) -> str:
    return f"{value:%H:%M}" if value else ""


def fmt_money(value, symbol: str) -> str:
    if value is None or value == "":
        return ""
    return f"{symbol}{float(value):,.2f}"


def fit(draw: ImageDraw.ImageDraw, text: str, f, max_w: int) -> str:
    text = "" if text is None else str(text)
    if draw.textlength(text, font=f) <= max_w:
        return text
    while text and draw.textlength(text + "…", font=f) > max_w:
        text = text[:-1]
    return text.rstrip() + "…"


# ---------------------------------------------------------------- drawing kit

def new_canvas(bg=CREAM):
    img = Image.new("RGB", (W, H), bg)
    return img, ImageDraw.Draw(img)


def frame(title: str, subtitle: str, number: int):
    """Standard page: teal header band, cream body, footer note."""
    img, d = new_canvas()
    d.rectangle([0, 0, W, 230], fill=TEAL)
    d.rectangle([0, 230, W, 240], fill=CORAL)
    d.text((MARGIN, 62), title, font=font(64, True), fill=WHITE)
    d.text((MARGIN, 150), subtitle, font=font(32), fill=(214, 236, 233))
    d.text((MARGIN, H - 70),
           "Example records shown (fictional). Pet Sitting & Dog Walking "
           "Business Spreadsheet",
           font=font(24), fill=GREY)
    d.text((W - MARGIN, H - 70), f"{number} / {TOTAL}", font=font(24),
           fill=GREY, anchor="ra")
    return img, d


def card(d, box, fill=WHITE, radius=28, outline=LINE):
    d.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=2)


def pill(d, x, y, text, fill, color=WHITE, size=30):
    f = font(size, True)
    w = int(d.textlength(text, font=f)) + 46
    d.rounded_rectangle([x, y, x + w, y + 64], radius=32, fill=fill)
    d.text((x + 23, y + 31), text, font=f, fill=color, anchor="lm")
    return x + w


def table(d, x, y, widths, headers, rows, aligns=None, row_h=46,
          head_size=26, body_size=24):
    """Draw a simple zebra table; returns the y coordinate under the table."""
    total = sum(widths)
    aligns = aligns or ["l"] * len(widths)
    hf, bf = font(head_size, True), font(body_size)
    d.rounded_rectangle([x, y, x + total, y + row_h], radius=14, fill=TEAL)
    cx = x
    for head, w, al in zip(headers, widths, aligns):
        tx = cx + 18 if al == "l" else (cx + w - 18 if al == "r" else cx + w // 2)
        anchor = {"l": "lm", "r": "rm", "c": "mm"}[al]
        d.text((tx, y + row_h / 2), fit(d, head, hf, w - 30), font=hf,
               fill=WHITE, anchor=anchor)
        cx += w
    y += row_h
    for i, row in enumerate(rows):
        fill = WHITE if i % 2 == 0 else ZEBRA
        d.rectangle([x, y, x + total, y + row_h], fill=fill)
        cx = x
        for cell, w, al in zip(row, widths, aligns):
            tx = cx + 18 if al == "l" else (cx + w - 18 if al == "r" else cx + w // 2)
            anchor = {"l": "lm", "r": "rm", "c": "mm"}[al]
            d.text((tx, y + row_h / 2), fit(d, cell, bf, w - 30), font=bf,
                   fill=INK, anchor=anchor)
            cx += w
        y += row_h
    d.line([x, y, x + total, y], fill=LINE, width=2)
    return y


def kpi(d, box, label, value, color=TEAL):
    x0, y0, x1, y1 = box
    card(d, box)
    d.rectangle([x0, y0 + 26, x0 + 10, y1 - 26], fill=color)
    d.text((x0 + 40, y0 + 36), label, font=font(30), fill=GREY)
    size = 70
    max_w = (x1 - x0) - 80
    while size > 36 and d.textlength(value, font=font(size, True)) > max_w:
        size -= 4
    d.text((x0 + 40, y0 + 100), value, font=font(size, True), fill=color)


def hbar_chart(d, box, title, items, color=TEAL, value_fmt=str):
    """items: list of (label, value). Horizontal bars inside a card."""
    x0, y0, x1, y1 = box
    card(d, box)
    d.text((x0 + 36, y0 + 30), title, font=font(34, True), fill=INK)
    top = y0 + 110
    label_w = 360
    bar_max = (x1 - x0) - label_w - 260
    biggest = max((v for _, v in items), default=1) or 1
    row_h = min(78, (y1 - top - 30) // max(1, len(items)))
    for i, (label, value) in enumerate(items):
        cy = top + i * row_h + row_h // 2
        d.text((x0 + 36, cy), fit(d, label, font(26), label_w - 20), font=font(26),
               fill=INK, anchor="lm")
        length = max(8, int(bar_max * value / biggest))
        bx = x0 + 36 + label_w
        d.rounded_rectangle([bx, cy - 22, bx + length, cy + 22], radius=12, fill=color)
        d.text((bx + length + 18, cy), value_fmt(value), font=font(26, True),
               fill=INK, anchor="lm")


# ---------------------------------------------------------------- data

def load():
    cfg = Config()
    sym = cfg.currency_symbol
    data = sample_data.get(cfg)
    return cfg, sym, data


# ---------------------------------------------------------------- the 15 images

def img01_cover(sym, data, cfg):
    img, d = new_canvas(TEAL)
    d.ellipse([1350, -250, 2250, 650], fill=TEAL_DARK)
    d.ellipse([-300, 1050, 500, 1850], fill=TEAL_DARK)
    d.rectangle([MARGIN, 300, MARGIN + 140, 312], fill=GOLD)
    d.text((MARGIN, 350), "Pet Sitting &", font=font(120, True), fill=WHITE)
    d.text((MARGIN, 490), "Dog Walking", font=font(120, True), fill=WHITE)
    d.text((MARGIN, 630), "Business Spreadsheet", font=font(120, True), fill=GOLD)
    d.text((MARGIN, 830),
           "Clients, pets, bookings, visits, invoices, expenses and reports",
           font=font(42), fill=(214, 236, 233))
    d.text((MARGIN, 885), "in one ready-to-use Excel workbook.",
           font=font(42), fill=(214, 236, 233))
    x = MARGIN
    for label, color in [("27 tabs", CORAL), ("Excel .xlsx", GOLD),
                         ("Sample data included", WHITE),
                         ("Blank version included", WHITE)]:
        text_color = WHITE if color == CORAL else INK
        x = pill(d, x, 1100, label, color, text_color) + 24
    d.text((MARGIN, 1260), "Instant digital download", font=font(40, True),
           fill=WHITE)
    return img


def img02_contents(sym, data, cfg):
    names = ["Start Here", "Dashboard", "Settings", "Lists", "Services", "Clients",
             "Pets", "Bookings", "Visit Log", "Health Log", "Incidents", "Messages",
             "Invoices", "Payments", "Expenses", "Mileage", "Staff", "Inventory",
             "Reminders", "Reports", "Budget", "Print Invoice", "Print Pet Care",
             "Print Visit", "Print Forms", "Help", "Calc"]
    img, d = frame("What's inside", "27 tabs, grouped by job", 2)
    cols, rows_n = 6, 5
    gap = 22
    tile_w = (W - 2 * MARGIN - gap * (cols - 1)) // cols
    tile_h = 178
    for i, name in enumerate(names):
        r, c = divmod(i, cols)
        x = MARGIN + c * (tile_w + gap)
        y = 290 + r * (tile_h + gap)
        hidden = name == "Calc"
        card(d, [x, y, x + tile_w, y + tile_h],
             fill=(235, 242, 241) if hidden else WHITE)
        d.text((x + 28, y + 30), f"{i + 1:02d}", font=font(28, True),
               fill=CORAL)
        d.text((x + 28, y + 88), name, font=font(34, True), fill=INK)
        d.text((x + 28, y + 136),
               "Hidden engine" if hidden else "Visible tab",
               font=font(24), fill=GREY)
    return img


def img03_start_here(sym, data, cfg):
    img, d = frame("Start here in four steps", "From blank file to a working business", 3)
    steps = [
        ("1", "Set up Settings",
         "Business name, currency, distance unit (km or miles) and\nthe rates for your services."),
        ("2", "Add clients and pets",
         "Clients, Pets and Staff tabs, with dropdowns for status,\nkey, contact method and pet details."),
        ("3", "Log bookings and visits",
         "Bookings, Visit Log, Health Log and Incidents keep the\nday-to-day record in one place."),
        ("4", "Track the money",
         "Invoices, Payments, Expenses and Mileage feed the Dashboard,\nReports and Budget tabs."),
    ]
    y = 290
    for num, head, body in steps:
        card(d, [MARGIN, y, W - MARGIN, y + 250])
        d.ellipse([MARGIN + 40, y + 60, MARGIN + 170, y + 190], fill=CORAL)
        d.text((MARGIN + 105, y + 125), num, font=font(72, True), fill=WHITE, anchor="mm")
        d.text((MARGIN + 230, y + 56), head, font=font(50, True), fill=INK)
        for j, line in enumerate(body.split("\n")):
            d.text((MARGIN + 230, y + 130 + j * 44), line, font=font(30), fill=GREY)
        y += 280
    return img


def img04_dashboard(sym, data, cfg):
    img, d = frame("Dashboard at a glance", "Every number here is calculated for you", 4)
    paid = sum(float(p["amount"] or 0) for p in data.payments)
    spent = sum(float(e["amount"] or 0) for e in data.expenses)
    miles = sum((m["end_odo"] or 0) - (m["start_odo"] or 0) for m in data.mileage)
    cards = [
        ("Clients", str(len(data.clients)), TEAL),
        ("Pets", str(len(data.pets)), TEAL),
        ("Bookings", str(len(data.bookings)), TEAL),
        ("Staff", str(len(data.staff)), TEAL),
        ("Payments received", fmt_money(paid, sym), CORAL),
        ("Expenses logged", fmt_money(spent, sym), CORAL),
        ("Mileage logged", f"{miles:,} {cfg.distance_unit}", CORAL),
        ("Services", str(len(data.services)), TEAL),
    ]
    cw, ch, gap = (W - 2 * MARGIN - 3 * 26) // 4, 230, 26
    for i, (label, value, color) in enumerate(cards):
        r, c = divmod(i, 4)
        kpi(d, [MARGIN + c * (cw + 26), 290 + r * (ch + 26),
                MARGIN + c * (cw + 26) + cw, 290 + r * (ch + 26) + ch],
            label, value, color)
    y = 290 + 2 * (ch + 26) + 10
    d.text((MARGIN, y), "Upcoming reminders", font=font(38, True), fill=INK)
    reminders = sorted(data.reminders, key=lambda r: r["due"])
    rows = [[fmt_date(r["due"]), r["task"], r["owner"], r["priority"], r["status"]]
            for r in reminders]
    table(d, MARGIN, y + 70, [260, 720, 330, 220, 310],
          ["Due", "Task", "Owner", "Priority", "Status"], rows, row_h=60)
    return img


def img05_clients(sym, data, cfg):
    img, d = frame("Clients", "Contact details, keys and referral source", 5)
    rows = [[c["name"], c["city"], c["contact_method"], c["status"],
             c["key_status"] or "", c["referral"] or ""] for c in data.clients]
    table(d, MARGIN, 290, [390, 300, 300, 240, 340, 270],
          ["Client", "City", "Contact", "Status", "Key", "Referral"], rows, row_h=86,
          body_size=26)
    return img


def img06_pets(sym, data, cfg):
    img, d = frame("Pets", "Breed, weight, temperament and vet, linked to each owner", 6)
    rows = [[p["name"], p["owner"], p["type"], p["breed"],
             f"{p['weight']} kg" if p["weight"] else "", p["temperament"] or "",
             p["vet_name"] or ""] for p in data.pets]
    table(d, MARGIN, 290, [200, 330, 150, 380, 180, 260, 340],
          ["Pet", "Owner", "Type", "Breed", "Weight", "Temperament", "Vet"],
          rows, row_h=78, body_size=24)
    return img


def img07_bookings(sym, data, cfg):
    shown = data.bookings[:14]
    img, d = frame("Bookings", "Every walk and visit, with status and staff", 7)
    rows = [[fmt_date(b["date"]), b["client"], b["pet"], b["service"],
             b["staff"] or "", f"{fmt_time(b['start'])}-{fmt_time(b['end'])}",
             b["status"]] for b in shown]
    table(d, MARGIN, 290, [200, 330, 330, 360, 220, 190, 210],
          ["Date", "Client", "Pet", "Service", "Staff", "Time", "Status"],
          rows, row_h=66, body_size=23)
    d.text((MARGIN, 290 + 66 * 15 + 30),
           f"Showing {len(shown)} of {len(data.bookings)} example bookings",
           font=font(26), fill=GREY)
    return img


def img08_visits(sym, data, cfg):
    img, d = frame("Visit Log", "What happened on each visit, for the owner update", 8)
    rows = [[v["booking"], fmt_time(v["arrive"]), fmt_time(v["depart"]), v["potty"],
             v["mood"], v["appetite"], v["health"], v["update_sent"], v["status"]]
            for v in data.visits]
    table(d, MARGIN, 290, [200, 170, 170, 190, 190, 210, 240, 230, 230],
          ["Booking", "Arrive", "Depart", "Potty", "Mood", "Appetite",
           "Health", "Update sent", "Status"],
          rows, row_h=66, body_size=24)
    return img


def img09_invoices(sym, data, cfg):
    img, d = frame("Invoices", "Client invoices with quantity, rate and status", 9)
    rows = []
    for inv in data.invoices:
        rows.append([inv["client"], fmt_date(inv["date"]), inv["description"],
                     str(inv["qty"]), fmt_money(inv["rate"], sym), inv["status"]])
    table(d, MARGIN, 290, [330, 220, 630, 130, 260, 270],
          ["Client", "Date", "Description", "Qty", "Rate", "Status"],
          rows, aligns=["l", "l", "l", "c", "r", "l"], row_h=86, body_size=24)
    return img


def img10_payments(sym, data, cfg):
    img, d = frame("Payments", "Money received, how it arrived and whether it is reconciled", 10)
    total = sum(float(p["amount"] or 0) for p in data.payments)
    kpi(d, [MARGIN, 290, MARGIN + 560, 500], "Received in this sample", fmt_money(total, sym),
        TEAL)
    rows = [[fmt_date(p["date"]), p["client"], p["invoice_no"] or "",
             fmt_money(p["amount"], sym), fmt_money(p["tip"], sym) if p["tip"] else "",
             p["method"], p["reconciled"]] for p in data.payments]
    table(d, MARGIN, 560, [220, 330, 240, 240, 180, 300, 330],
          ["Date", "Client", "Invoice", "Amount", "Tip", "Method", "Reconciled"],
          rows, aligns=["l", "l", "l", "r", "r", "l", "l"], row_h=86, body_size=24)
    return img


def img11_expenses(sym, data, cfg):
    shown = data.expenses[:14]
    img, d = frame("Expenses", "Fuel, supplies and costs, with receipts and approval", 11)
    rows = [[fmt_date(e["date"]), e["category"], e["description"], e["supplier"] or "",
             fmt_money(e["amount"], sym), e["method"], e["approval"]] for e in shown]
    table(d, MARGIN, 290, [200, 230, 520, 230, 200, 230, 230],
          ["Date", "Category", "Description", "Supplier", "Amount", "Method", "Approval"],
          rows, aligns=["l", "l", "l", "l", "r", "l", "l"], row_h=66, body_size=23)
    d.text((MARGIN, 290 + 66 * 15 + 30),
           f"Showing {len(shown)} of {len(data.expenses)} example expenses",
           font=font(26), fill=GREY)
    return img


def img12_mileage(sym, data, cfg):
    img, d = frame("Mileage", "Business journeys, with the distance worked out for you", 12)
    rows = []
    for m in data.mileage:
        rows.append([fmt_date(m["date"]), m["staff"], m["client"] or "", m["purpose"],
                     m["start_loc"], m["end_loc"],
                     f"{(m['end_odo'] or 0) - (m['start_odo'] or 0):,} {cfg.distance_unit}"])
    table(d, MARGIN, 290, [200, 230, 290, 230, 300, 380, 210],
          ["Date", "Staff", "Client", "Purpose", "From", "To", "Distance"],
          rows, aligns=["l", "l", "l", "l", "l", "l", "r"], row_h=80, body_size=23)
    return img


def img13_health(sym, data, cfg):
    img, d = frame("Health Log", "Medication given, supplies left and refills due", 13)
    rows = [[fmt_date(h["date"]), h["pet"], h["type"], h["item"], h["dosage"] or "",
             h["given_by"] or "", h["status"], fmt_date(h["refill_due"])]
            for h in data.health]
    table(d, MARGIN, 290, [190, 300, 170, 270, 300, 190, 150, 230],
          ["Date", "Pet", "Type", "Item", "Dosage", "Given by", "Status", "Refill due"],
          rows, row_h=86, body_size=22)
    return img


def img14_reports(sym, data, cfg):
    img, d = frame("Reports", "Bookings per service and expenses per category", 14)
    counts = Counter(b["service"] for b in data.bookings)
    service_items = sorted(counts.items(), key=lambda kv: -kv[1])
    spend = defaultdict(float)
    for e in data.expenses:
        spend[e["category"]] += float(e["amount"] or 0)
    spend_items = sorted(spend.items(), key=lambda kv: -kv[1])
    half = (W - 2 * MARGIN - 40) // 2
    hbar_chart(d, [MARGIN, 290, MARGIN + half, 1340], "Bookings per service",
               service_items, TEAL)
    hbar_chart(d, [MARGIN + half + 40, 290, W - MARGIN, 1340],
               "Expenses per category", [(k, v) for k, v in spend_items],
               CORAL, value_fmt=lambda v: fmt_money(v, sym))
    return img


def img15_features(sym, data, cfg):
    img, d = frame("Why this workbook", "What you get when you download it", 15)
    items = [
        ("27 tabs", "Start Here, Dashboard, clients, pets, bookings, money and printables."),
        ("Dropdowns", "Lists for status, payment method, service and staff are built in."),
        ("Protected sheets", "Formulas are locked so the calculations stay intact."),
        ("Sample and blank", "Use the filled version to learn, then start from the empty one."),
        ("Excel 2010 and newer", "Formulas chosen to work in Excel 2010 and later versions."),
        ("One .xlsx file", "No macros and nothing to install beyond Excel."),
    ]
    cw, ch = (W - 2 * MARGIN - 40) // 2, 300
    for i, (head, body) in enumerate(items):
        r, c = divmod(i, 2)
        x = MARGIN + c * (cw + 40)
        y = 290 + r * (ch + 30)
        card(d, [x, y, x + cw, y + ch])
        d.rounded_rectangle([x + 40, y + 46, x + 110, y + 116], radius=18, fill=CORAL)
        d.text((x + 75, y + 81), "✓", font=font(44, True), fill=WHITE, anchor="mm")
        d.text((x + 140, y + 56), head, font=font(44, True), fill=INK)
        for j, line in enumerate(_wrap(d, body, font(30), cw - 200)):
            d.text((x + 140, y + 130 + j * 46), line, font=font(30), fill=GREY)
    return img


def _wrap(d, text, f, max_w):
    words, lines, cur = text.split(), [], ""
    for w in words:
        trial = (cur + " " + w).strip()
        if d.textlength(trial, font=f) <= max_w:
            cur = trial
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


IMAGES = [
    ("01_cover", img01_cover),
    ("02_whats_inside", img02_contents),
    ("03_start_here", img03_start_here),
    ("04_dashboard", img04_dashboard),
    ("05_clients", img05_clients),
    ("06_pets", img06_pets),
    ("07_bookings", img07_bookings),
    ("08_visit_log", img08_visits),
    ("09_invoices", img09_invoices),
    ("10_payments", img10_payments),
    ("11_expenses", img11_expenses),
    ("12_mileage", img12_mileage),
    ("13_health_log", img13_health),
    ("14_reports", img14_reports),
    ("15_features", img15_features),
]


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    parser.add_argument("-o", "--output",
                        default=os.path.join(ROOT, "output", "etsy_package",
                                             "listing_images"))
    args = parser.parse_args(argv)
    os.makedirs(args.output, exist_ok=True)
    cfg, sym, data = load()
    for name, make in IMAGES:
        path = os.path.join(args.output, name + ".png")
        make(sym, data, cfg).save(path, optimize=True)
        print("wrote", path)
    print(f"{len(IMAGES)} images written to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
