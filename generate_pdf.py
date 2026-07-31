import base64
import json
import re
import sys
from io import BytesIO
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    Image,
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parent
PRIMARY = colors.HexColor("#0f304d")
ACCENT = colors.HexColor("#c82f2f")
BORDER = colors.HexColor("#b8c1cc")
MUTED = colors.HexColor("#606975")


def clean(value):
    return str(value or "").replace("\r\n", "\n").replace("\r", "\n")


def para(text, style):
    escaped = (
        clean(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\n", "<br/>")
    )
    return Paragraph(escaped or "&nbsp;", style)


def data_url_image(value, width, height):
    if not value or "," not in value:
        return Spacer(width, height)
    payload = value.split(",", 1)[1]
    try:
        raw = base64.b64decode(payload)
    except Exception:
        return Spacer(width, height)
    return Image(BytesIO(raw), width=width, height=height, kind="proportional")


def field(label, value, styles, height=12 * mm):
    return Table(
        [[Paragraph(label.upper(), styles["label"])], [para(value, styles["value"])]],
        rowHeights=[None, height],
        style=TableStyle(
            [
                ("BOX", (0, 1), (-1, 1), 0.7, BORDER),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        ),
    )


def section(title, styles):
    return Paragraph(title.upper(), styles["section"])


def normalise_parts(parts):
    if not isinstance(parts, list):
        return []
    cleaned = []
    for index, part in enumerate(parts, start=1):
        description = clean(part.get("description", ""))
        quantity = clean(part.get("quantity", part.get("charge", "")))
        serial = clean(part.get("serial", index))
        if description or quantity:
            cleaned.append([serial, description, quantity])
    return cleaned or [["", "", ""]]


def build_pdf(data, output_path):
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=12 * mm,
        leftMargin=12 * mm,
        topMargin=6 * mm,
        bottomMargin=6 * mm,
    )
    width = A4[0] - doc.leftMargin - doc.rightMargin

    styles = {
        "section": ParagraphStyle(
            "Section",
            fontName="Helvetica-Bold",
            fontSize=12,
            leading=14,
            textColor=PRIMARY,
            spaceBefore=6,
            spaceAfter=3,
        ),
        "label": ParagraphStyle(
            "Label",
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=11,
            textColor=MUTED,
            spaceAfter=2,
        ),
        "value": ParagraphStyle(
            "Value",
            fontName="Helvetica",
            fontSize=11,
            leading=14,
            textColor=colors.black,
        ),
        "jobLabel": ParagraphStyle(
            "JobLabel",
            fontName="Helvetica-Bold",
            fontSize=11,
            leading=12,
            textColor=PRIMARY,
        ),
        "jobNumber": ParagraphStyle(
            "JobNumber",
            fontName="Helvetica-Bold",
            fontSize=24,
            leading=27,
            textColor=ACCENT,
        ),
    }

    story = []
    logo_path = ROOT / "JNW.png"
    logo = Image(str(logo_path), width=62 * mm, height=20 * mm, kind="proportional")
    job = Table(
        [
            [Paragraph("JOB ORDER", styles["jobLabel"])],
            [Paragraph(clean(data.get("jobOrder")), styles["jobNumber"])],
        ],
        colWidths=[70 * mm],
    )
    header = Table([[logo, job]], colWidths=[width - 75 * mm, 75 * mm])
    header.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LINEBELOW", (0, 0), (-1, -1), 1.2, PRIMARY),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.extend([header, Spacer(1, 3 * mm)])

    story.append(section("Customer Details", styles))
    story.append(
        Table(
            [
                [
                    field("Company", data.get("company"), styles),
                    field("Date", data.get("date"), styles),
                ],
                [
                    field("Requested By", data.get("requestedBy"), styles),
                    "",
                ],
            ],
            colWidths=[width / 2 - 3 * mm, width / 2 - 3 * mm],
            hAlign="LEFT",
            style=TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            ),
        )
    )

    story.append(section("Work Report", styles))
    story.append(field("Complaint", data.get("complaint"), styles, height=16 * mm))
    story.append(Spacer(1, 2 * mm))
    story.append(field("Action Taken", data.get("actionTaken"), styles, height=24 * mm))

    story.append(section("Materials and Parts Used", styles))
    rows = [["S/NO", "MATERIALS AND PARTS USED", "QUANTITY"]] + normalise_parts(data.get("parts"))
    materials = Table(rows, colWidths=[18 * mm, width - 55 * mm, 32 * mm], repeatRows=1)
    materials.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, -1), 0.7, BORDER),
                ("GRID", (0, 0), (-1, -1), 0.5, BORDER),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f4f6f8")),
                ("TEXTCOLOR", (0, 0), (-1, 0), MUTED),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.append(materials)

    story.append(section("Labour", styles))
    story.append(
        Table(
            [
                [
                    field("Labour Description", data.get("labourDescription"), styles),
                    field("Man", data.get("labourMan"), styles),
                    field("Hours", data.get("labourHours"), styles),
                ]
            ],
            colWidths=[width / 3 - 3 * mm] * 3,
            style=TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 3),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ]
            ),
        )
    )

    story.append(section("Remarks", styles))
    story.append(field("Remarks", data.get("remarks"), styles, height=14 * mm))

    story.append(section("Signatures", styles))
    sig_width = width / 2 - 5 * mm
    sig_height = 18 * mm
    sigs = data.get("signatures") if isinstance(data.get("signatures"), dict) else {}
    sig_table = Table(
        [
            [
                Paragraph("CUSTOMER SIGNATURE AND CHOP", styles["label"]),
                Paragraph("TECHNICIAN SIGNATURE", styles["label"]),
            ],
            [
                data_url_image(sigs.get("customerSignature"), sig_width, sig_height),
                data_url_image(sigs.get("technicianSignature"), sig_width, sig_height),
            ],
        ],
        colWidths=[sig_width, sig_width],
    )
    sig_table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOX", (0, 1), (0, 1), 0.7, BORDER),
                ("BOX", (1, 1), (1, 1), 0.7, BORDER),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 1), (-1, 1), 5),
                ("BOTTOMPADDING", (0, 1), (-1, 1), 5),
            ]
        )
    )
    story.append(sig_table)

    doc.build(story)


def main():
    if len(sys.argv) != 3:
        raise SystemExit("Usage: generate_pdf.py input.json output.pdf")
    data = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    build_pdf(data, sys.argv[2])


if __name__ == "__main__":
    main()
