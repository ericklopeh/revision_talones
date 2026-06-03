from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook


def limpiar_nombre_archivo(texto: str) -> str:
    texto = texto.upper().strip()
    caracteres_invalidos = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']

    for caracter in caracteres_invalidos:
        texto = texto.replace(caracter, "")

    texto = texto.replace(" ", "_")
    return texto


def escribir_celda(ws, celda: str, valor):
    """
    Escribe en una celda normal o combinada.
    Si la celda pertenece a un rango combinado, escribe en la celda principal.
    """

    for rango in ws.merged_cells.ranges:
        if celda in rango:
            celda_principal = rango.coord.split(":")[0]
            ws[celda_principal] = valor
            return

    ws[celda] = valor


def generar_excel_revision(
    datos: dict,
    revision: dict,
    mensaje_vendedor: str,
    promotor: str,
    qna: str = "09-2026",
    ruta_template: str = "templates/plantilla_revision_talon.xlsx",
    carpeta_salida: str = "output"
) -> str:
    Path(carpeta_salida).mkdir(exist_ok=True)

    wb = load_workbook(ruta_template)
    ws = wb["Hoja1"]

    codigos = revision["codigos_revision"]

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass

    # =========================
    # ENCABEZADO
    # =========================

    escribir_celda(ws, "B3", datos["rfc"])
    escribir_celda(ws, "D3", datos["nombre"])

    escribir_celda(ws, "C4", qna)
    escribir_celda(ws, "E4", datos.get("fecha_pago", ""))

    # =========================
    # INGRESOS
    # Mapeo correcto según tu plantilla.
    # =========================

    escribir_celda(ws, "B6", codigos["E4"])
    escribir_celda(ws, "B7", codigos["E3"])
    escribir_celda(ws, "B8", codigos["Q"])
    escribir_celda(ws, "B9", codigos["CP"])
    escribir_celda(ws, "B10", codigos["7"])
    escribir_celda(ws, "B11", codigos["CT"])
    escribir_celda(ws, "B12", codigos["7B"])
    escribir_celda(ws, "B13", codigos["E9"])
    escribir_celda(ws, "B14", codigos["SG"])
    escribir_celda(ws, "B15", codigos["O1"])

    # Total ingresos
    escribir_celda(ws, "C6", revision["ingresos"])

    # =========================
    # DEDUCCIONES
    # =========================

    escribir_celda(ws, "E6", revision["descuentos"])
    escribir_celda(ws, "E7", revision["saldo_100"])
    escribir_celda(ws, "E8", revision["total_para_venta_70"])
    escribir_celda(ws, "E10", revision["saldo_70"])
    escribir_celda(ws, "D12", revision["saldo_70"])
    escribir_celda(ws, "D15", revision["saldo_100"])

    # DC TALON y DC SISTEMA
    escribir_celda(ws, "B16", codigos["DC"])
    escribir_celda(ws, "D16", codigos["DC"])

    # =========================
    # NO ESCRIBIR VENDEDOR EN A18
    # A18:A28 está combinado verticalmente y por eso se ve letra por letra.
    # Lo dejamos limpio.
    # =========================

    escribir_celda(ws, "A18", "")

    # =========================
    # RECUPERACIÓN / VENTA
    # Por ahora no estamos llenando folios de venta.
    # =========================

    escribir_celda(ws, "C20", 0)
    escribir_celda(ws, "C21", 0)
    escribir_celda(ws, "C22", 0)
    escribir_celda(ws, "C23", 0)

    escribir_celda(ws, "E20", 0)
    escribir_celda(ws, "E21", 0)
    escribir_celda(ws, "E22", 0)
    escribir_celda(ws, "E23", 0)

    escribir_celda(ws, "B25", 0)
    escribir_celda(ws, "D25", 0)

    escribir_celda(ws, "B28", revision["saldo_70"])
    escribir_celda(ws, "D28", revision["saldo_100"])

    # =========================
    # REVISIÓN DEL CLIENTE
    # =========================

    escribir_celda(ws, "B31", promotor)
    escribir_celda(ws, "B32", datos["nombre"])
    escribir_celda(ws, "B33", datos["rfc"])
    escribir_celda(ws, "B34", revision["liquidez_final"])
    escribir_celda(ws, "B35", qna)

    escribir_celda(ws, "B36", mensaje_vendedor)
    escribir_celda(ws, "B40", datetime.now().strftime("%d/%m/%Y"))

    # =========================
    # FORMATO DE MONEDA
    # =========================

    celdas_moneda = [
        "B6", "B7", "B8", "B9", "B10", "B11", "B12", "B13", "B14", "B15",
        "C6",
        "E6", "E7", "E8", "E10", "D12", "D15",
        "B16", "D16",
        "B25", "D25", "B28", "D28",
        "B34"
    ]

    for celda in celdas_moneda:
        escribir_celda(ws, celda, ws[celda].value)
        ws[celda].number_format = '$#,##0.00;-$#,##0.00;$-'

    # =========================
    # GUARDAR ARCHIVO
    # =========================

    nombre_cliente = limpiar_nombre_archivo(datos["nombre"])
    nombre_promotor = limpiar_nombre_archivo(promotor)

    nombre_archivo = f"REVISION_{nombre_cliente}_{nombre_promotor}.xlsx"
    ruta_salida = Path(carpeta_salida) / nombre_archivo

    wb.save(ruta_salida)

    return str(ruta_salida)