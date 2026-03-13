import random
from pathlib import Path
from openpyxl import Workbook

BASE = Path(__file__).parent
excel_dir = BASE / "input" / "excel"
pdf_dir = BASE / "input" / "pdf_inbox"

excel_dir.mkdir(parents=True, exist_ok=True)
pdf_dir.mkdir(parents=True, exist_ok=True)

SECURITIES = [
    ("US0378331005", "Apple Inc.", "15.02.2024"),
    ("US5949181045", "Microsoft Corp.", "14.03.2024"),
    ("US0231351067", "Amazon.com Inc.", "05.03.2024"),
    ("US02079K3059", "Alphabet Inc.", "18.03.2024"),
    ("US88160R1014", "Tesla Inc.", "22.03.2024"),
    ("DE0008404005", "Allianz SE", "10.05.2024"),
    ("DE0007164600", "SAP SE", "18.05.2024"),
    ("FR0000120271", "TotalEnergies SE", "03.04.2024"),
    ("NL0000009538", "Royal Philips", "11.04.2024"),
]

BENEFICIAL_OWNERS = [
    "BlackRock Institutional",
    "UBS Asset Management",
    "Allianz Global Investors",
    "Vanguard Group",
    "State Street Global Advisors",
    "Norges Bank Investment Management"
]


def create_visible_pdf(path, request_id, isin, bo, payment_date, company):
    text_lines = [
        "TAX DOCUMENT TEST PDF",
        "",
        f"Request ID: {request_id}",
        f"Company: {company}",
        f"ISIN: {isin}",
        f"Beneficial Owner: {bo}",
        f"Payment Date: {payment_date}",
        "",
        "Used for testing sort_rename_merge pipeline"
    ]

    y = 750
    text_commands = []

    for line in text_lines:
        safe_line = line.replace("(", "").replace(")", "")
        text_commands.append(f"BT /F1 14 Tf 50 {y} Td ({safe_line}) Tj ET")
        y -= 30

    content_stream = "\n".join(text_commands)
    content_bytes = content_stream.encode("latin-1")

    pdf = b"%PDF-1.4\n"

    objects = []

    objects.append(b"<< /Type /Catalog /Pages 2 0 R >>")

    objects.append(b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>")

    objects.append(
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 800] "
        b"/Resources << /Font << /F1 4 0 R >> >> "
        b"/Contents 5 0 R >>"
    )

    objects.append(
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
    )

    objects.append(
        f"<< /Length {len(content_bytes)} >>\nstream\n".encode()
        + content_bytes +
        b"\nendstream"
    )

    offsets = []
    pos = len(pdf)

    for i, obj in enumerate(objects, start=1):
        offsets.append(pos)
        obj_bytes = f"{i} 0 obj\n".encode() + obj + b"\nendobj\n"
        pdf += obj_bytes
        pos += len(obj_bytes)

    xref_pos = len(pdf)

    pdf += f"xref\n0 {len(objects)+1}\n".encode()
    pdf += b"0000000000 65535 f \n"

    for off in offsets:
        pdf += f"{off:010d} 00000 n \n".encode()

    pdf += (
        f"trailer\n<< /Size {len(objects)+1} /Root 1 0 R >>\n"
        f"startxref\n{xref_pos}\n%%EOF".encode()
    )

    with open(path, "wb") as f:
        f.write(pdf)


rows = []

for i in range(30):
    isin, company, payment_date = random.choice(SECURITIES)
    bo = random.choice(BENEFICIAL_OWNERS)

    request_id = f"REQ{i+1:05d}"

    rows.append({
        "Corp. Action Ref.": f"CA-{1000+i}",
        "Financial Instrument": isin,
        "Paymt. Date": payment_date,
        "Market": "GLOBAL",
        "Document Category": "Dividend",
        "Safekeeping Account": f"ACC-{random.randint(10000,99999)}",
        "Common Code": random.randint(100000,999999),
        "BO Name": bo,
        "Quantity": random.randint(10,1000),
        "Clearstream Request ID": request_id
    })

    pdf_path = pdf_dir / f"{request_id}.pdf"

    create_visible_pdf(
        pdf_path,
        request_id,
        isin,
        bo,
        payment_date,
        company
    )

wb = Workbook()
ws = wb.active

headers = list(rows[0].keys())
ws.append(headers)

for r in rows:
    ws.append(list(r.values()))

excel_file = excel_dir / "realistic_test_data.xlsx"
wb.save(excel_file)

print("Test dataset generated")
print("Excel:", excel_file)
print("PDFs:", pdf_dir)