"""Generacion del catalogo PDF con las propiedades seleccionadas."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Iterable

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

PRIMARY = colors.HexColor("#1e3a5f")
ACCENT = colors.HexColor("#c9a227")
MUTED = colors.HexColor("#64748b")
LIGHT_BG = colors.HexColor("#f1f5f9")

_styles = getSampleStyleSheet()

COVER_TITLE = ParagraphStyle(
    "CoverTitle", parent=_styles["Title"], fontSize=28, textColor=PRIMARY, spaceAfter=6,
)
COVER_SUBTITLE = ParagraphStyle(
    "CoverSubtitle", parent=_styles["Normal"], fontSize=13, textColor=MUTED, spaceAfter=4,
)
PROPERTY_NAME = ParagraphStyle(
    "PropertyName", parent=_styles["Heading2"], fontSize=17, textColor=PRIMARY, spaceAfter=2,
)
PROPERTY_LOCATION = ParagraphStyle(
    "PropertyLocation", parent=_styles["Normal"], fontSize=11, textColor=MUTED, spaceAfter=8,
)
PRICE_STYLE = ParagraphStyle(
    "Price", parent=_styles["Normal"], fontSize=20, textColor=ACCENT, spaceAfter=8, fontName="Helvetica-Bold",
)
LABEL_STYLE = ParagraphStyle("Label", parent=_styles["Normal"], fontSize=9, textColor=MUTED)
VALUE_STYLE = ParagraphStyle("Value", parent=_styles["Normal"], fontSize=12, textColor=colors.black, fontName="Helvetica-Bold")
DESC_STYLE = ParagraphStyle("Desc", parent=_styles["Normal"], fontSize=9.5, textColor=colors.HexColor("#334155"), leading=13)
LINK_STYLE = ParagraphStyle("Link", parent=_styles["Normal"], fontSize=9, textColor=colors.HexColor("#2563eb"))
BADGE_DEV = ParagraphStyle("BadgeDev", parent=_styles["Normal"], fontSize=9, textColor=colors.white, alignment=1)


def _cover_page(count: int) -> list:
    elements = [
        Spacer(1, 4 * cm),
        Paragraph("Catalogo de Inmuebles Seleccionados", COVER_TITLE),
        Paragraph("La Paz, Baja California Sur", COVER_SUBTITLE),
        Spacer(1, 1 * cm),
        Paragraph(f"{count} propiedad{'es' if count != 1 else ''} incluida{'s' if count != 1 else ''}", COVER_SUBTITLE),
        Paragraph(f"Generado el {datetime.now().strftime('%d/%m/%Y %H:%M')}", COVER_SUBTITLE),
        Spacer(1, 6 * cm),
    ]
    return elements


def _spec_cell(label: str, value: str) -> list:
    return [Paragraph(label, LABEL_STYLE), Paragraph(value, VALUE_STYLE)]


def _property_block(prop: dict) -> list:
    image_path = str(prop.get("image_path") or "")
    try:
        # Las fotos ya se cachean recortadas a 3:2 (ver data_loader._crop_to_fill),
        # se respeta esa proporcion aqui para que no se vean estiradas.
        img = Image(image_path, width=8 * cm, height=8 * cm * 2 / 3)
        img.hAlign = "LEFT"
    except Exception:
        img = Spacer(8 * cm, 8 * cm * 2 / 3)

    estado = prop.get("estado", "")
    badge_color = ACCENT if estado == "Desarrollo" else PRIMARY
    header_table = Table(
        [[Paragraph(prop.get("nombre", "Propiedad"), PROPERTY_NAME),
          Table([[Paragraph(estado or "-", BADGE_DEV)]], style=TableStyle([
              ("BACKGROUND", (0, 0), (-1, -1), badge_color),
              ("TOPPADDING", (0, 0), (-1, -1), 4),
              ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
              ("LEFTPADDING", (0, 0), (-1, -1), 10),
              ("RIGHTPADDING", (0, 0), (-1, -1), 10),
              ("ROUNDEDCORNERS", [6, 6, 6, 6]),
          ]))]],
        colWidths=[13 * cm, 4 * cm],
    )
    header_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))

    specs = Table(
        [
            _spec_cell("Recamaras", str(int(prop.get("recamaras") or 0))) + _spec_cell("Banos", str(int(prop.get("banos") or 0))),
            _spec_cell("Estacionamientos", str(int(prop.get("estacionamientos") or 0))) + _spec_cell("m2 Construccion", prop.get("m2_construccion_display", "N/D")),
            _spec_cell("m2 Terreno", prop.get("m2_terreno_display", "N/D")) + _spec_cell("Ubicacion", prop.get("ubicacion", "N/D")),
        ],
        colWidths=[4 * cm, 4 * cm, 4 * cm, 4 * cm],
    )
    specs.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("BACKGROUND", (0, 0), (-1, -1), LIGHT_BG),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.white),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
    ]))

    right_col = [Paragraph(prop.get("precio_display", "Precio no disponible"), PRICE_STYLE), specs]
    if prop.get("descripcion"):
        right_col.append(Spacer(1, 8))
        right_col.append(Paragraph(prop["descripcion"][:400], DESC_STYLE))
    if prop.get("enlace_externo"):
        right_col.append(Spacer(1, 6))
        right_col.append(Paragraph(f'<link href="{prop["enlace_externo"]}">Ver anuncio original -&gt;</link>', LINK_STYLE))

    body = Table([[img, right_col]], colWidths=[8 * cm, 9.5 * cm])
    body.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (1, 0), (1, 0), 12)]))

    divider = Table([[""]], colWidths=[17.5 * cm], rowHeights=[1])
    divider.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 1, colors.HexColor("#e2e8f0"))]))

    return [header_table, Spacer(1, 6), body, Spacer(1, 14), divider]


def generate_catalog_pdf(properties: Iterable[dict], output_path: str | Path) -> Path:
    """Genera el PDF del catalogo. `properties` es una lista de dicts con las
    claves: nombre, precio_display, ubicacion, estado, recamaras, banos,
    estacionamientos, m2_construccion_display, m2_terreno_display,
    descripcion, enlace_externo, image_path."""
    properties = list(properties)
    output_path = Path(output_path)

    doc = SimpleDocTemplate(
        str(output_path), pagesize=letter,
        topMargin=1.8 * cm, bottomMargin=1.5 * cm, leftMargin=1.5 * cm, rightMargin=1.5 * cm,
        title="Catalogo de Inmuebles Seleccionados",
    )

    elements: list = _cover_page(len(properties))
    elements.append(PageBreak())

    for i, prop in enumerate(properties):
        elements.extend(_property_block(prop))
        if i < len(properties) - 1 and (i + 1) % 2 == 0:
            elements.append(PageBreak())

    doc.build(elements)
    return output_path
