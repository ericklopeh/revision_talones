from pathlib import Path
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)

from services.generador_excel import limpiar_nombre_archivo


def formato_moneda_pdf(valor: float) -> str:
    valor = float(valor or 0)
    return f"-${abs(valor):,.2f}" if valor < 0 else f"${valor:,.2f}"


def generar_pdf_revision(
    datos: dict,
    revision: dict,
    mensaje_vendedor: str,
    promotor: str,
    qna: str,
    carpeta_salida: str = "output"
) -> str:
    Path(carpeta_salida).mkdir(exist_ok=True)

    cliente = limpiar_nombre_archivo(datos["nombre"])
    vendedor = limpiar_nombre_archivo(promotor)
    ruta_salida = (
        Path(carpeta_salida)
        / f"REVISION_{cliente}_{vendedor}.pdf"
    )

    documento = SimpleDocTemplate(
        str(ruta_salida),
        pagesize=letter,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm
    )
    estilos = getSampleStyleSheet()
    estilos.add(ParagraphStyle(
        name="TituloRevision",
        parent=estilos["Title"],
        alignment=TA_CENTER,
        textColor=colors.HexColor("#1F4E78"),
        spaceAfter=12
    ))
    estilos.add(ParagraphStyle(
        name="SeccionRevision",
        parent=estilos["Heading2"],
        textColor=colors.HexColor("#1F4E78"),
        spaceBefore=10,
        spaceAfter=6
    ))

    contenido = [
        Paragraph("Revisión de talón", estilos["TituloRevision"]),
        Paragraph(
            f"<b>Cliente:</b> {datos['nombre']}<br/>"
            f"<b>RFC:</b> {datos['rfc']}<br/>"
            f"<b>Promotor:</b> {promotor}<br/>"
            f"<b>QNA revisión:</b> {qna}<br/>"
            f"<b>Fecha de revisión:</b> {datetime.now().strftime('%d/%m/%Y')}",
            estilos["BodyText"]
        ),
        Spacer(1, 8),
        Paragraph("Resultado de revisión", estilos["SeccionRevision"])
    ]

    filas_resumen = [
        ["Concepto", "Importe"],
        ["Liquidez del talón", formato_moneda_pdf(revision["liquidez_talon"])],
        ["Apoyo adicional", formato_moneda_pdf(revision["abono_extra"])],
        ["Saldo liberado", formato_moneda_pdf(revision["total_saldo_liberado"])],
        ["Solo observado", formato_moneda_pdf(revision.get("total_solo_observado", 0))],
        ["Programado", formato_moneda_pdf(revision["programado"])],
        ["Liquidez final", formato_moneda_pdf(revision["liquidez_final"])]
    ]
    tabla_resumen = Table(filas_resumen, colWidths=[110 * mm, 45 * mm])
    tabla_resumen.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("BACKGROUND", (0, 1), (-1, -2), colors.HexColor("#F8FAFC")),
        ("ALIGN", (1, 1), (1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6)
    ]))
    contenido.append(tabla_resumen)

    cuentas = revision.get("cuentas_terminadas", [])
    if cuentas:
        contenido.append(Paragraph(
            "Cuentas que terminan / saldo que se libera",
            estilos["SeccionRevision"]
        ))
        filas_cuentas = [[
            "QNA termina",
            "Saldo liberado",
            "Suma",
            "Observación"
        ]]

        for cuenta in cuentas:
            filas_cuentas.append([
                cuenta.get("qna_termina", ""),
                formato_moneda_pdf(cuenta.get("saldo_liberado", 0)),
                "Sí" if cuenta.get("sumar_a_liquidez", False) else "No",
                Paragraph(
                    str(cuenta.get("observacion", "")),
                    estilos["BodyText"]
                )
            ])

        tabla_cuentas = Table(
            filas_cuentas,
            colWidths=[30 * mm, 35 * mm, 20 * mm, 70 * mm],
            repeatRows=1
        )
        tabla_cuentas.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#334155")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5)
        ]))
        contenido.append(tabla_cuentas)

    contenido.extend([
        Paragraph("Mensaje generado", estilos["SeccionRevision"]),
        Paragraph(
            mensaje_vendedor.replace("\n", "<br/>"),
            estilos["BodyText"]
        )
    ])

    documento.build(contenido)
    return str(ruta_salida)
