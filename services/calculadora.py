def obtener_importe(codigos: dict, codigo: str, equivalencias: list[str] = None) -> float:
    """
    Busca un código en el diccionario de códigos extraídos.
    También permite buscar equivalencias.
    Ejemplo:
    Q puede venir como A2
    7 puede venir como 07
    O1 puede venir como 01
    """

    equivalencias = equivalencias or []

    posibles_codigos = [codigo] + equivalencias

    for posible in posibles_codigos:
        if posible in codigos:
            return float(codigos[posible]["importe"])

    return 0.0


def filtrar_codigos_revision(codigos_extraidos: dict) -> dict:
    """
    Devuelve solo los códigos que usa tu formato de revisión.
    """

    codigos_revision = {
        "E4": obtener_importe(codigos_extraidos, "E4"),
        "E3": obtener_importe(codigos_extraidos, "E3"),
        "Q": obtener_importe(codigos_extraidos, "Q", ["A2"]),
        "CP": obtener_importe(codigos_extraidos, "CP"),
        "7": obtener_importe(codigos_extraidos, "7", ["07"]),
        "CT": obtener_importe(codigos_extraidos, "CT"),
        "7B": obtener_importe(codigos_extraidos, "7B"),
        "E9": obtener_importe(codigos_extraidos, "E9"),
        "SG": obtener_importe(codigos_extraidos, "SG"),
        "O1": obtener_importe(codigos_extraidos, "O1", ["01"]),
        "DC": obtener_importe(codigos_extraidos, "DC")
    }

    return codigos_revision


def calcular_ingresos_revision(codigos_revision: dict) -> float:
    """
    Suma solo los códigos que forman los ingresos liquidables.
    Según tu formato:
    E4 + E3 + Q + CP + 7 + CT + 7B + E9 + SG
    """

    codigos_ingresos = [
        "E4",
        "E3",
        "Q",
        "CP",
        "7",
        "CT",
        "7B",
        "E9",
        "SG"
    ]

    total = sum(codigos_revision.get(codigo, 0) for codigo in codigos_ingresos)

    return round(total, 2)


def calcular_revision_talon(
    codigos_extraidos: dict,
    descuentos_talon: float,
    abono_extra: float = 0,
    programado: float = 0
) -> dict:
    """
    Calcula los datos principales de tu formato.
    """

    codigos_revision = filtrar_codigos_revision(codigos_extraidos)

    ingresos = calcular_ingresos_revision(codigos_revision)

    saldo_100 = ingresos - descuentos_talon
    total_para_venta_70 = ingresos * 0.70
    saldo_70 = total_para_venta_70 - descuentos_talon

    saldo_mas_abonos_70 = saldo_70 - programado
    saldo_mas_abono_100 = saldo_100 - programado

    liquidez_final = saldo_100 + abono_extra - programado

    return {
        "codigos_revision": codigos_revision,
        "ingresos": round(ingresos, 2),
        "descuentos": round(descuentos_talon, 2),
        "saldo_100": round(saldo_100, 2),
        "total_para_venta_70": round(total_para_venta_70, 2),
        "saldo_70": round(saldo_70, 2),
        "saldo_mas_abonos_70": round(saldo_mas_abonos_70, 2),
        "saldo_mas_abono_100": round(saldo_mas_abono_100, 2),
        "abono_extra": round(abono_extra, 2),
        "programado": round(programado, 2),
        "liquidez_final": round(liquidez_final, 2)
    }