import re
from datetime import date, datetime
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


COLUMNAS_MANUALES = [
    "FACT",
    "VTA",
    "SALDO",
    "ABONO",
    "EN COBRO",
    "QNAS TOMADAS A CUENTA"
]

COLUMNAS_CALCULADAS = [
    "PAGADO",
    "ABONO DE QUINCENAS",
    "SALDO PENDIENTE",
    "PORCENTAJE PAGADO",
    "REFINANCIAMIENTO",
    "PUEDE REFINANCIAR"
]

PLAZOS_REFINANCIAMIENTO = [72, 60, 46, 34]


def to_number(valor) -> float:
    if valor is None or pd.isna(valor):
        return 0.0

    if isinstance(valor, str):
        valor = (
            valor.replace("$", "")
            .replace(",", "")
            .replace("%", "")
            .strip()
        )

    try:
        return float(valor)
    except (TypeError, ValueError):
        return 0.0


convertir_numero = to_number


def extraer_fecha_edad_desde_rfc(
    rfc: str,
    forzar_1900: bool = True,
    fecha_referencia: date = None
) -> dict:
    fecha_referencia = fecha_referencia or date.today()
    rfc = str(rfc or "").upper().strip()

    for coincidencia in re.finditer(r"\d{6}", rfc):
        bloque = coincidencia.group(0)
        anio_corto = int(bloque[:2])
        mes = int(bloque[2:4])
        dia = int(bloque[4:6])

        if forzar_1900:
            anio = 1900 + anio_corto
        else:
            anio_actual_corto = fecha_referencia.year % 100
            siglo = 2000 if anio_corto <= anio_actual_corto else 1900
            anio = siglo + anio_corto

        try:
            fecha_nacimiento = date(anio, mes, dia)
        except ValueError:
            continue

        if fecha_nacimiento > fecha_referencia:
            continue

        edad = fecha_referencia.year - fecha_nacimiento.year
        if (
            fecha_referencia.month,
            fecha_referencia.day
        ) < (
            fecha_nacimiento.month,
            fecha_nacimiento.day
        ):
            edad -= 1

        return {
            "fecha_nacimiento": fecha_nacimiento,
            "edad": edad
        }

    return {
        "fecha_nacimiento": None,
        "edad": None
    }


def dataframe_facturas_vacio(filas: int = 5) -> pd.DataFrame:
    return pd.DataFrame([
        {
            "FACT": "",
            "VTA": 0.0,
            "SALDO": 0.0,
            "ABONO": 0.0,
            "EN COBRO": "",
            "QNAS TOMADAS A CUENTA": ""
        }
        for _ in range(filas)
    ])


def calcular_facturas_refinanciamiento(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=COLUMNAS_MANUALES + COLUMNAS_CALCULADAS)

    resultado = df.copy()

    for columna in COLUMNAS_MANUALES:
        if columna not in resultado:
            resultado[columna] = "" if columna in ["FACT", "EN COBRO"] else 0.0

    columnas_numericas = [
        "VTA",
        "SALDO",
        "QNAS TOMADAS A CUENTA",
        "ABONO"
    ]
    for columna in columnas_numericas:
        resultado[columna] = resultado[columna].map(convertir_numero)

    resultado["PAGADO"] = 0.0
    tiene_venta = resultado["VTA"] > 0
    resultado.loc[tiene_venta, "PAGADO"] = (
        resultado.loc[tiene_venta, "VTA"]
        - resultado.loc[tiene_venta, "SALDO"]
    ).round(2)
    resultado["ABONO DE QUINCENAS"] = (
        resultado["QNAS TOMADAS A CUENTA"] * resultado["ABONO"]
    ).round(2)
    resultado["SALDO PENDIENTE"] = (
        resultado["SALDO"] - resultado["ABONO DE QUINCENAS"]
    ).round(2)
    resultado["PORCENTAJE PAGADO"] = 0.0
    resultado.loc[tiene_venta, "PORCENTAJE PAGADO"] = (
        resultado.loc[tiene_venta, "PAGADO"]
        / resultado.loc[tiene_venta, "VTA"]
    )
    resultado["PORCENTAJE PAGADO"] = resultado[
        "PORCENTAJE PAGADO"
    ].round(6)
    resultado["REFINANCIAMIENTO"] = resultado["PAGADO"].round(2)
    fila_activa = (
        resultado["FACT"].astype(str).str.strip().ne("")
        | resultado["EN COBRO"].astype(str).str.strip().ne("")
        | resultado[columnas_numericas].abs().sum(axis=1).gt(0)
    )
    resultado["PUEDE REFINANCIAR"] = ""
    resultado.loc[fila_activa, "PUEDE REFINANCIAR"] = resultado.loc[
        fila_activa,
        "PORCENTAJE PAGADO"
    ].map(lambda porcentaje: "SI" if porcentaje >= 0.39 else "NO")

    columnas_salida = [
        "FACT",
        "VTA",
        "PAGADO",
        "SALDO",
        "QNAS TOMADAS A CUENTA",
        "ABONO",
        "EN COBRO",
        "ABONO DE QUINCENAS",
        "SALDO PENDIENTE",
        "PORCENTAJE PAGADO",
        "REFINANCIAMIENTO",
        "PUEDE REFINANCIAR"
    ]
    return resultado[columnas_salida]


def filtrar_facturas_capturadas(df: pd.DataFrame) -> pd.DataFrame:
    facturas = calcular_facturas_refinanciamiento(df)

    if facturas.empty:
        return facturas

    columnas_numericas = [
        "VTA",
        "SALDO",
        "QNAS TOMADAS A CUENTA",
        "ABONO"
    ]
    fila_activa = (
        facturas["FACT"].astype(str).str.strip().ne("")
        | facturas["EN COBRO"].astype(str).str.strip().ne("")
        | facturas[columnas_numericas].abs().sum(axis=1).gt(0)
    )
    return facturas.loc[fila_activa].reset_index(drop=True)


def calcular_resumen_refinanciamiento(
    df: pd.DataFrame,
    aumento_descuento: float
) -> dict:
    facturas = calcular_facturas_refinanciamiento(df)
    aumento_descuento = convertir_numero(aumento_descuento)

    def total(columna: str) -> float:
        if facturas.empty:
            return 0.0
        return round(float(facturas[columna].sum()), 2)

    facturas_refinanciables = facturas[
        facturas["PUEDE REFINANCIAR"] == "SI"
    ]
    abono_ref = round(float(facturas_refinanciables["ABONO"].sum()), 2)
    abono_antes = total("ABONO")
    total_abono_nuevo = round(abono_ref + aumento_descuento, 2)
    total_saldo_pendiente = total("SALDO PENDIENTE")

    simulacion = {}
    for plazo in PLAZOS_REFINANCIAMIENTO:
        total_saldo_nuevo = round(total_abono_nuevo * plazo, 2)
        venta_posible = round(
            total_saldo_nuevo - total_saldo_pendiente,
            2
        )
        simulacion[plazo] = {
            "SALDO PENDIENTE": total_saldo_pendiente,
            "VENTA POSIBLE": venta_posible,
            "TOTAL SALDO NUEVO": total_saldo_nuevo,
            "DESCUENTO NUEVO": total_abono_nuevo,
            "TOTAL ADEUDO CLIENTE": total_saldo_nuevo
        }

    return {
        "total_vta": total("VTA"),
        "total_pagado": total("PAGADO"),
        "total_saldo": total("SALDO"),
        "total_abono": abono_antes,
        "total_abono_quincenas": total("ABONO DE QUINCENAS"),
        "total_saldo_pendiente": total_saldo_pendiente,
        "total_refinanciamiento": total("REFINANCIAMIENTO"),
        "abono_ref": abono_ref,
        "abono_antes": abono_antes,
        "aumento_descuento": round(aumento_descuento, 2),
        "total_abono_nuevo": total_abono_nuevo,
        "simulacion": simulacion
    }


def dataframe_simulacion(resumen: dict) -> pd.DataFrame:
    filas = [
        "SALDO PENDIENTE",
        "VENTA POSIBLE",
        "TOTAL SALDO NUEVO",
        "DESCUENTO NUEVO",
        "TOTAL ADEUDO CLIENTE"
    ]

    return pd.DataFrame({
        plazo: [
            resumen["simulacion"][plazo][fila]
            for fila in filas
        ]
        for plazo in PLAZOS_REFINANCIAMIENTO
    }, index=filas)


def generar_excel_refinanciamiento(
    facturas: pd.DataFrame,
    resumen: dict,
    datos_cliente: dict
) -> bytes:
    facturas = filtrar_facturas_capturadas(facturas)
    wb = Workbook()
    ws_facturas = wb.active
    ws_facturas.title = "Facturas"

    encabezado_fill = PatternFill("solid", fgColor="1F4E78")
    encabezado_font = Font(color="FFFFFF", bold=True)
    subtitulo_fill = PatternFill("solid", fgColor="D9EAF7")
    moneda = '$ #,##0.00;-$ #,##0.00;$ -'
    porcentaje = "0%"

    columnas_facturas = [
        "FACT",
        "VTA",
        "PAGADO",
        "SALDO",
        "QNAS TOMADAS A CUENTA",
        "ABONO",
        "EN COBRO",
        "ABONO DE QUINCENAS",
        "SALDO PENDIENTE",
        "PORCENTAJE PAGADO",
        "REFINANCIAMIENTO",
        "PUEDE REFINANCIAR"
    ]
    ws_facturas.append(columnas_facturas)

    for celda in ws_facturas[1]:
        celda.fill = encabezado_fill
        celda.font = encabezado_font
        celda.alignment = Alignment(horizontal="center", wrap_text=True)

    for _, fila in facturas.iterrows():
        ws_facturas.append([fila[columna] for columna in columnas_facturas])

    columnas_moneda = {
        "VTA",
        "PAGADO",
        "SALDO",
        "ABONO",
        "ABONO DE QUINCENAS",
        "SALDO PENDIENTE",
        "REFINANCIAMIENTO"
    }
    for indice, columna in enumerate(columnas_facturas, start=1):
        ancho = min(max(len(columna) + 3, 13), 27)
        ws_facturas.column_dimensions[get_column_letter(indice)].width = ancho
        for celda in ws_facturas[get_column_letter(indice)][1:]:
            if columna in columnas_moneda:
                celda.number_format = moneda
            elif columna == "PORCENTAJE PAGADO":
                celda.number_format = porcentaje

    ws_facturas.freeze_panes = "A2"
    ws_facturas.auto_filter.ref = ws_facturas.dimensions

    ws_resumen = wb.create_sheet("Resumen Refinanciamiento")
    ws_resumen["A1"] = "RESUMEN DE REFINANCIAMIENTO"
    ws_resumen["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws_resumen["A1"].fill = encabezado_fill
    ws_resumen.merge_cells("A1:E1")
    ws_resumen["A1"].alignment = Alignment(horizontal="center")

    campos = [
        ("Fecha", datos_cliente.get("fecha")),
        ("Cliente", datos_cliente.get("cliente", "")),
        ("RFC/NAC", datos_cliente.get("rfc_nac", "")),
        ("Fecha nacimiento", datos_cliente.get("fecha_nacimiento")),
        ("Edad", datos_cliente.get("edad")),
        ("Quinquenio", datos_cliente.get("quinquenio", "")),
        ("ABONO REF", resumen["abono_ref"]),
        ("ABONO ANTES", resumen["abono_antes"]),
        ("LIQUIDEZ / AUMENTO EN DESCUENTO", resumen["aumento_descuento"]),
        ("TOTAL ABONO NUEVO", resumen["total_abono_nuevo"]),
        ("total_vta", resumen["total_vta"]),
        ("total_pagado", resumen["total_pagado"]),
        ("total_saldo", resumen["total_saldo"]),
        ("total_abono", resumen["total_abono"]),
        ("total_saldo_pendiente", resumen["total_saldo_pendiente"])
    ]

    for fila, (etiqueta, valor) in enumerate(campos, start=3):
        ws_resumen.cell(fila, 1, etiqueta)
        ws_resumen.cell(fila, 1).font = Font(bold=True)
        ws_resumen.cell(fila, 2, valor)

        if fila >= 9:
            ws_resumen.cell(fila, 2).number_format = moneda

    ws_resumen["B3"].number_format = "dd/mm/yyyy"
    ws_resumen["B6"].number_format = "dd/mm/yyyy"

    inicio_simulacion = 20
    ws_resumen.cell(inicio_simulacion, 1, "PLAZO")
    for columna, plazo in enumerate(PLAZOS_REFINANCIAMIENTO, start=2):
        ws_resumen.cell(inicio_simulacion, columna, plazo)

    filas_simulacion = [
        "SALDO PENDIENTE",
        "VENTA POSIBLE",
        "TOTAL SALDO NUEVO",
        "DESCUENTO NUEVO",
        "TOTAL ADEUDO CLIENTE"
    ]
    for desplazamiento, nombre_fila in enumerate(filas_simulacion, start=1):
        fila_excel = inicio_simulacion + desplazamiento
        ws_resumen.cell(fila_excel, 1, nombre_fila)
        for columna, plazo in enumerate(PLAZOS_REFINANCIAMIENTO, start=2):
            celda = ws_resumen.cell(
                fila_excel,
                columna,
                resumen["simulacion"][plazo][nombre_fila]
            )
            celda.number_format = moneda

    for celda in ws_resumen[inicio_simulacion]:
        celda.fill = encabezado_fill
        celda.font = encabezado_font
        celda.alignment = Alignment(horizontal="center")

    for fila in range(inicio_simulacion + 1, inicio_simulacion + 6):
        ws_resumen.cell(fila, 1).fill = subtitulo_fill
        ws_resumen.cell(fila, 1).font = Font(bold=True)

    ws_resumen.column_dimensions["A"].width = 35
    for columna in "BCDE":
        ws_resumen.column_dimensions[columna].width = 18
    ws_resumen.freeze_panes = "A3"

    salida = BytesIO()
    wb.save(salida)
    return salida.getvalue()
