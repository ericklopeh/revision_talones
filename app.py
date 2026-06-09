import streamlit as st
from pathlib import Path
from datetime import datetime
from typing import Optional

import pandas as pd

from services.extractor_pdf import extraer_datos_talon, extraer_datos_talon_imagen
from services.calculadora import calcular_revision_talon
from services.generador_excel import generar_excel_revision
from services.generador_pdf import generar_pdf_revision
from services.graph_storage import subir_revision_a_graph, GraphStorageError
from utils.refinanciamiento import (
    COLUMNAS_CALCULADAS,
    calcular_facturas_refinanciamiento,
    calcular_resumen_refinanciamiento,
    dataframe_facturas_vacio,
    dataframe_simulacion,
    extraer_fecha_edad_desde_rfc,
    generar_excel_refinanciamiento
)


st.set_page_config(
    page_title="Revisión de Talones",
    page_icon="📄",
    layout="wide"
)

PROMOTORES = [
    "Victor Vega",
    "Juan Manuel",
    "Leonardo Arevalo",
    "Eliezer Chipuli",
    "Gerardo Santana",
    "Sergio Valadez",
    "Sergio Vazquez"
]

CODIGOS_FORMATO = [
    ("E4", ""),
    ("E3", ""),
    ("Q", "A2"),
    ("CP", ""),
    ("7", "07"),
    ("CT", "CT"),
    ("7B", ""),
    ("E9", ""),
    ("SG", "SG"),
    ("O1", "01"),
    ("DC", "DC")
]

Path("uploads").mkdir(exist_ok=True)
Path("output").mkdir(exist_ok=True)


def formato_moneda(valor: float) -> str:
    if valor < 0:
        return f"-${abs(valor):,.2f}"

    return f"${valor:,.2f}"


def nueva_cuenta_terminada(cuenta_id: int) -> dict:
    return {
        "_id": cuenta_id,
        "qna_termina": "",
        "saldo_liberado": 0.0,
        "observacion": "",
        "sumar_a_liquidez": True
    }


def normalizar_cuentas_terminadas(cuentas: list) -> list:
    cuentas = cuentas or []
    siguiente_id = 1

    for cuenta in cuentas:
        if "_id" not in cuenta:
            cuenta["_id"] = siguiente_id

        siguiente_id = max(siguiente_id, int(cuenta["_id"]) + 1)
        cuenta.setdefault("qna_termina", "")
        cuenta.setdefault("saldo_liberado", 0.0)
        cuenta.setdefault("observacion", "")
        cuenta.setdefault("sumar_a_liquidez", True)

    return cuentas


def cuentas_por_tipo(cuentas: list, sumar_a_liquidez: bool) -> list:
    return [
        cuenta for cuenta in cuentas
        if bool(cuenta.get("sumar_a_liquidez", False)) == sumar_a_liquidez
    ]


def lineas_cuentas_para_mensaje(cuentas: list) -> list[str]:
    lineas = []

    for cuenta in cuentas:
        qna = str(cuenta.get("qna_termina", "")).strip() or "Sin QNA"
        monto = float(cuenta.get("saldo_liberado", 0) or 0)
        observacion = str(cuenta.get("observacion", "")).strip()
        linea = f"- QNA {qna}: {formato_moneda(monto)}"

        if observacion:
            linea += f" ({observacion})"

        lineas.append(linea)

    return lineas


def generar_resultado_liquidez(revision: dict, tiene_programado: str) -> str:
    liquidez_final = revision["liquidez_final"]
    programado = revision["programado"]
    cuentas_terminadas = revision.get("cuentas_terminadas", [])

    if cuentas_terminadas:
        if (
            revision.get("liquidez_talon", 0) < 0
            and revision.get("liquidez_antes_liberacion", 0) < 0
            and liquidez_final > 0
            and revision.get("total_saldo_liberado", 0) > 0
        ):
            resultado = (
                "Tenía sobregiro, pero queda con liquidez por cuentas terminadas: "
                f"{formato_moneda(liquidez_final)}."
            )
        elif liquidez_final > 0:
            resultado = (
                f"Tiene liquidez disponible de "
                f"{formato_moneda(liquidez_final)}."
            )
        elif liquidez_final < 0:
            resultado = (
                f"No tiene liquidez. Sobregiro final de "
                f"{formato_moneda(liquidez_final)}."
            )
        else:
            resultado = "Queda sin liquidez disponible ni sobregiro."

        if tiene_programado == "Sí" and programado > 0:
            resultado += f" Tiene un programado por {formato_moneda(programado)}."

        return resultado

    if liquidez_final > 0:
        if tiene_programado == "Sí" and programado > 0:
            return (
                f"Tiene liquidez de {formato_moneda(liquidez_final)}. "
                f"Tiene un programado por {formato_moneda(programado)}."
            )

        return f"Tiene liquidez de {formato_moneda(liquidez_final)}."

    if liquidez_final < 0 and tiene_programado == "Sí" and programado > 0:
        return (
            f"No tiene liquidez. Tiene un sobregiro de: {formato_moneda(liquidez_final)}. "
            f"Tiene un programado por {formato_moneda(programado)}."
        )

    if liquidez_final < 0:
        return f"No tiene liquidez. Tiene un sobregiro de: {formato_moneda(liquidez_final)}."

    return "No tiene liquidez disponible."


def generar_mensaje_vendedor(datos: dict, revision: dict, tiene_programado: str) -> str:
    nombre = datos["nombre"]
    rfc = datos["rfc"]

    resultado_liquidez = generar_resultado_liquidez(
        revision=revision,
        tiene_programado=tiene_programado
    )
    cuentas = revision.get("cuentas_terminadas", [])

    if not cuentas:
        return f"""Se realizó la revisión del talón correspondiente al cliente:

Cliente: {nombre}
RFC: {rfc}

Resultado de la revisión:
{resultado_liquidez}"""

    cuentas_liberadas = cuentas_por_tipo(cuentas, True)
    cuentas_observadas = cuentas_por_tipo(cuentas, False)
    liquidez_talon = revision.get("liquidez_talon", 0)
    estado_talon = (
        f"liquidez de {formato_moneda(liquidez_talon)}"
        if liquidez_talon >= 0
        else f"sobregiro de {formato_moneda(liquidez_talon)}"
    )
    bloques = [
        f"El cliente {nombre} presenta {estado_talon} en talón.",
        f"RFC: {rfc}"
    ]

    if cuentas_liberadas:
        bloques.append(
            "Adicionalmente, cuenta con saldos liberados por cuentas que terminan:\n"
            + "\n".join(lineas_cuentas_para_mensaje(cuentas_liberadas))
            + "\n\nTotal liberado: "
            + formato_moneda(revision.get("total_saldo_liberado", 0))
        )

    if cuentas_observadas:
        bloques.append(
            "Cuentas registradas solo como observación, sin sumar a liquidez:\n"
            + "\n".join(lineas_cuentas_para_mensaje(cuentas_observadas))
            + "\n\nTotal solo observado: "
            + formato_moneda(revision.get("total_solo_observado", 0))
        )

    bloques.append(f"Considerando lo anterior, {resultado_liquidez.lower()}")
    return "\n\n".join(bloques)


def importe_detectado(codigos: dict, cod: str, equiv: str) -> float:
    if cod in codigos:
        return float(codigos[cod]["importe"])

    if equiv and equiv in codigos:
        return float(codigos[equiv]["importe"])

    return 0.0


def construir_codigos_manual(codigos_detectados: dict) -> dict:
    codigos_manual = {}

    for cod, equiv in CODIGOS_FORMATO:
        valor = importe_detectado(codigos_detectados, cod, equiv)
        codigos_manual[cod] = {"descripcion": "manual", "importe": valor}

    return codigos_manual


def tiene_programado_desde_monto(programado: float) -> str:
    return "Sí" if programado > 0 else "No"


def extraer_datos_desde_archivo(ruta: Path) -> dict:
    es_imagen = ruta.suffix.lower() in [".jpg", ".jpeg", ".png"]

    if es_imagen:
        return extraer_datos_talon_imagen(str(ruta))

    return extraer_datos_talon(str(ruta))


def calcular_revision_desde_registro(registro: dict) -> dict:
    codigos_manual = construir_codigos_manual(registro.get("codigos", {}))

    return calcular_revision_talon(
        codigos_extraidos=codigos_manual,
        descuentos_talon=float(registro["descuentos"]),
        abono_extra=float(registro.get("abono_extra", 0)),
        programado=float(registro.get("programado", 0)),
        cuentas_terminadas=(
            registro.get("cuentas_terminadas", [])
            if registro.get("tiene_cuentas_terminadas", False)
            else []
        )
    )


def registros_activos(registros: list) -> list:
    return [r for r in registros if r.get("procesar", True)]


def consolidar_codigos(registros: list) -> dict:
    codigos_sumados = {}

    for registro in registros:
        codigos_manual = construir_codigos_manual(registro.get("codigos", {}))

        for cod, info in codigos_manual.items():
            codigos_sumados[cod] = codigos_sumados.get(cod, 0.0) + float(info["importe"])

    return {
        cod: {"descripcion": "consolidado", "importe": round(valor, 2)}
        for cod, valor in codigos_sumados.items()
    }


def calcular_revision_consolidada(registros: list) -> Optional[dict]:
    activos = registros_activos(registros)

    if not activos:
        return None

    codigos = consolidar_codigos(activos)
    total_descuentos = sum(float(r["descuentos"]) for r in activos)
    total_abono = sum(float(r.get("abono_extra", 0)) for r in activos)
    total_programado = sum(float(r.get("programado", 0)) for r in activos)
    cuentas_terminadas = []

    for registro in activos:
        if registro.get("tiene_cuentas_terminadas", False):
            cuentas_terminadas.extend(registro.get("cuentas_terminadas", []))

    return calcular_revision_talon(
        codigos_extraidos=codigos,
        descuentos_talon=total_descuentos,
        abono_extra=total_abono,
        programado=total_programado,
        cuentas_terminadas=cuentas_terminadas
    )


def generar_mensaje_vendedor_lote(promotor: str, registros: list, revision: dict) -> str:
    activos = registros_activos(registros)
    tiene_prog = tiene_programado_desde_monto(revision["programado"])
    resultado_total = generar_resultado_liquidez(revision, tiene_prog)

    cantidad = len(activos)
    lineas_clientes = []

    for registro in activos:
        revision_individual = calcular_revision_desde_registro(registro)
        nombre = registro.get("nombre", "").strip() or registro.get("archivo", "Sin nombre")
        lineas_clientes.append(
            f"- {nombre} — Liquidez: {formato_moneda(revision_individual['liquidez_final'])}"
        )

    clientes_texto = "\n".join(lineas_clientes)
    cuentas = revision.get("cuentas_terminadas", [])
    cuentas_liberadas = cuentas_por_tipo(cuentas, True)
    cuentas_observadas = cuentas_por_tipo(cuentas, False)
    bloques_cuentas = []

    if cuentas_liberadas:
        bloques_cuentas.append(
            "Saldos liberados:\n"
            + "\n".join(lineas_cuentas_para_mensaje(cuentas_liberadas))
            + f"\nTotal liberado: "
            + formato_moneda(revision.get("total_saldo_liberado", 0))
        )

    if cuentas_observadas:
        bloques_cuentas.append(
            "Solo observadas:\n"
            + "\n".join(lineas_cuentas_para_mensaje(cuentas_observadas))
            + f"\nTotal observado: "
            + formato_moneda(revision.get("total_solo_observado", 0))
        )

    bloque_cuentas = (
        "\n\n" + "\n\n".join(bloques_cuentas)
        if bloques_cuentas
        else ""
    )

    if cantidad == 1:
        intro = f"Se realizó la revisión del talón correspondiente al promotor {promotor}:"
    else:
        intro = (
            f"Se realizó la revisión de {cantidad} talones "
            f"correspondientes al promotor {promotor}:"
        )

    return f"""{intro}

Clientes revisados:
{clientes_texto}{bloque_cuentas}

Resultado total de la revisión:
{resultado_total}"""


def render_resumen_revision(revision: dict, tiene_programado: str, key_prefix: str = ""):
    resultado_liquidez = generar_resultado_liquidez(revision, tiene_programado)

    with st.container(border=True):
        st.caption(
            "Liquidez del talón + apoyo adicional + saldo liberado "
            "- programado = liquidez final"
        )

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Liquidez del talón", formato_moneda(revision["liquidez_talon"]))
        with col2:
            st.metric("Apoyo adicional", formato_moneda(revision["abono_extra"]))
        with col3:
            st.metric(
                "Saldo liberado",
                formato_moneda(revision["total_saldo_liberado"])
            )

        col4, col5, col6 = st.columns(3)
        with col4:
            st.metric("Programado", formato_moneda(revision["programado"]))
        with col5:
            st.metric(
                "Solo observado",
                formato_moneda(revision.get("total_solo_observado", 0))
            )
        with col6:
            st.metric("Liquidez final", formato_moneda(revision["liquidez_final"]))

        if revision["liquidez_final"] > 0:
            st.success(resultado_liquidez)
        elif revision["liquidez_final"] < 0:
            st.error(resultado_liquidez)
        else:
            st.info(resultado_liquidez)

    return resultado_liquidez


def render_mensaje_vendedor(mensaje: str, key_prefix: str = ""):
    st.caption("Usa el botón de copia del bloque para llevar el texto al portapapeles.")
    st.code(mensaje, language="text")
    st.text_area(
        "Editar mensaje antes de enviarlo",
        value=mensaje,
        height=220,
        key=f"{key_prefix}mensaje_area"
    )


def registro_a_datos_excel(registro: dict) -> dict:
    return {
        "nombre": registro["nombre"],
        "rfc": registro["rfc"],
        "fecha_pago": registro.get("fecha_pago", ""),
        "percepciones": float(registro.get("percepciones", 0)),
        "descuentos": float(registro["descuentos"]),
        "liquido": float(registro.get("liquido", 0)),
        "codigos": registro.get("codigos", {}),
        "texto_original": registro.get("texto_original", "")
    }


def procesar_archivo_lote(archivo) -> dict:
    ruta = Path("uploads") / f"lote_{archivo.name}"

    with open(ruta, "wb") as salida:
        salida.write(archivo.getbuffer())

    try:
        datos = extraer_datos_desde_archivo(ruta)
    except Exception as error:
        return {
            "archivo": archivo.name,
            "ruta": str(ruta),
            "nombre": "",
            "rfc": "",
            "fecha_pago": "",
            "percepciones": 0.0,
            "descuentos": 0.0,
            "liquido": 0.0,
            "abono_extra": 0.0,
            "programado": 0.0,
            "cuentas_terminadas": [],
            "tiene_cuentas_terminadas": False,
            "codigos": {},
            "texto_original": "",
            "error": str(error)
        }

    return {
        "archivo": archivo.name,
        "ruta": str(ruta),
        "nombre": datos.get("nombre", ""),
        "rfc": datos.get("rfc", ""),
        "fecha_pago": datos.get("fecha_pago", ""),
        "percepciones": float(datos.get("percepciones", 0)),
        "descuentos": float(datos.get("descuentos", 0)),
        "liquido": float(datos.get("liquido", 0)),
        "abono_extra": 0.0,
        "programado": 0.0,
        "cuentas_terminadas": [],
        "tiene_cuentas_terminadas": False,
        "codigos": datos.get("codigos", {}),
        "texto_original": datos.get("texto_original", ""),
        "error": None
    }


def render_ajustes_revision(key_prefix: str = ""):
    col_v, col_q, col_anio, col_semana = st.columns(4)

    with col_v:
        promotor = st.selectbox(
            "Promotor / vendedor",
            PROMOTORES,
            key=f"{key_prefix}promotor"
        )

    with col_q:
        qna = st.text_input(
            "QNA",
            value="09-2026",
            key=f"{key_prefix}qna"
        )

    with col_anio:
        anio = st.number_input(
            "Año",
            min_value=2024,
            max_value=2035,
            value=datetime.now().year,
            step=1,
            key=f"{key_prefix}anio"
        )

    with col_semana:
        semana = st.number_input(
            "Semana",
            min_value=1,
            max_value=53,
            value=datetime.now().isocalendar().week,
            step=1,
            key=f"{key_prefix}semana"
        )

    return promotor, qna, int(anio), int(semana)


def render_cuentas_terminadas(
    cuentas: list,
    tiene_cuentas: bool,
    key_prefix: str
) -> tuple[list, bool]:
    nivel_titulo = "### 5." if key_prefix == "ind_" else "####"
    st.markdown(
        f"{nivel_titulo} Cuentas que terminan / saldo que se libera"
    )
    st.caption(
        "Registra las cuentas que concluyen y decide cuáles afectan la liquidez."
    )

    respuesta = st.selectbox(
        "¿Tiene cuentas que terminan?",
        ["No", "Sí"],
        index=1 if tiene_cuentas else 0,
        key=f"{key_prefix}tiene_cuentas_terminadas"
    )
    tiene_cuentas = respuesta == "Sí"
    cuentas = normalizar_cuentas_terminadas(cuentas)

    if not tiene_cuentas:
        return cuentas, False

    if not cuentas:
        cuentas.append(nueva_cuenta_terminada(1))

    cuenta_a_eliminar = None

    for indice, cuenta in enumerate(cuentas):
        cuenta_id = cuenta["_id"]
        with st.container(border=True):
            st.markdown(f"**Cuenta terminada {indice + 1}**")
            col_qna, col_monto, col_suma = st.columns([1.2, 1.2, 1])

            with col_qna:
                cuenta["qna_termina"] = st.text_input(
                    "QNA en que termina",
                    value=str(cuenta.get("qna_termina", "")),
                    placeholder="16-2026",
                    key=f"{key_prefix}cuenta_{cuenta_id}_qna"
                )

            with col_monto:
                cuenta["saldo_liberado"] = st.number_input(
                    "Saldo que libera",
                    min_value=0.0,
                    value=float(cuenta.get("saldo_liberado", 0)),
                    step=100.0,
                    format="%.2f",
                    key=f"{key_prefix}cuenta_{cuenta_id}_monto"
                )

            with col_suma:
                cuenta["sumar_a_liquidez"] = st.checkbox(
                    "Sumar a liquidez",
                    value=bool(cuenta.get("sumar_a_liquidez", True)),
                    key=f"{key_prefix}cuenta_{cuenta_id}_sumar"
                )

            col_observacion, col_eliminar = st.columns([4, 1])
            with col_observacion:
                cuenta["observacion"] = st.text_input(
                    "Observación (opcional)",
                    value=str(cuenta.get("observacion", "")),
                    key=f"{key_prefix}cuenta_{cuenta_id}_observacion"
                )

            with col_eliminar:
                st.write("")
                if st.button(
                    "Eliminar",
                    key=f"{key_prefix}cuenta_{cuenta_id}_eliminar"
                ):
                    cuenta_a_eliminar = cuenta_id

    if cuenta_a_eliminar is not None:
        cuentas[:] = [
            cuenta for cuenta in cuentas
            if cuenta["_id"] != cuenta_a_eliminar
        ]
        st.rerun()

    if st.button(
        "+ Agregar otra cuenta",
        key=f"{key_prefix}agregar_cuenta"
    ):
        siguiente_id = max(
            (int(cuenta["_id"]) for cuenta in cuentas),
            default=0
        ) + 1
        cuentas.append(nueva_cuenta_terminada(siguiente_id))
        st.rerun()

    total = sum(
        float(cuenta.get("saldo_liberado", 0) or 0)
        for cuenta in cuentas
        if cuenta.get("sumar_a_liquidez", False)
    )
    total_observado = sum(
        float(cuenta.get("saldo_liberado", 0) or 0)
        for cuenta in cuentas
        if not cuenta.get("sumar_a_liquidez", False)
    )
    col_total, col_observado = st.columns(2)
    with col_total:
        st.success(f"Total que sí se suma: {formato_moneda(total)}")
    with col_observado:
        st.info(f"Total solo observado: {formato_moneda(total_observado)}")

    return cuentas, True


st.markdown(
    """
    <style>
    .block-container {max-width: 1180px; padding-top: 2rem; padding-bottom: 3rem;}
    div[data-testid="stMetric"] {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 0.8rem;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("Sistema de revisión de talones")
st.caption(
    "Lectura del talón, cálculo de liquidez y generación de evidencia de revisión."
)

tab_revision, tab_refinanciamiento = st.tabs([
    "Revisión de Talón",
    "Refinanciamiento"
])

with tab_revision:
    tab_individual, tab_lote = st.tabs([
        "Un talón",
        "Varios talones (lote)"
    ])


# =============================================================================
# TAB: UN TALÓN
# =============================================================================

with tab_individual:
    st.markdown("### 1. Subir talón")
    archivo_pdf = st.file_uploader(
        "Sube un talón en PDF o imagen (JPG/PNG)",
        type=["pdf", "jpg", "jpeg", "png"],
        key="uploader_individual"
    )

    modo_manual = st.checkbox(
        "Captura manual (sin subir archivo o si no se detectó bien el talón)",
        key="modo_manual_individual"
    )

    datos = None
    ruta_pdf = None

    if archivo_pdf:
        ruta_pdf = Path("uploads") / archivo_pdf.name

        with open(ruta_pdf, "wb") as archivo:
            archivo.write(archivo_pdf.getbuffer())

        es_imagen = ruta_pdf.suffix.lower() in [".jpg", ".jpeg", ".png"]

        try:
            if es_imagen:
                datos = extraer_datos_talon_imagen(str(ruta_pdf))
            else:
                datos = extraer_datos_talon(str(ruta_pdf))
        except RuntimeError as error:
            st.error(
                "No se pudo leer la imagen por OCR. "
                "Verifica que Tesseract esté disponible en el entorno."
            )
            st.exception(error)
            st.stop()

        if es_imagen:
            st.success("Imagen leída por OCR. Revisa y corrige los datos antes de generar el Excel.")
        else:
            st.success("Talón leído correctamente. Puedes corregir cualquier dato antes de generar el Excel.")

    elif modo_manual:
        datos = {
            "nombre": "",
            "rfc": "",
            "fecha_pago": "",
            "percepciones": 0.0,
            "descuentos": 0.0,
            "liquido": 0.0,
            "codigos": {},
            "texto_original": ""
        }
        st.info("Captura manual activada. Llena los campos a mano.")

    if datos is not None:
        st.divider()
        st.markdown("### 2. Datos del cliente")
        st.caption(
            "Los valores se precargan con lo detectado. "
            "Corrige lo que haga falta; el cálculo se actualiza al instante."
        )

        col1, col2, col3 = st.columns(3)

        with col1:
            nombre = st.text_input("Nombre", value=datos["nombre"], key="ind_nombre")

        with col2:
            rfc = st.text_input("RFC", value=datos["rfc"], key="ind_rfc")

        with col3:
            fecha_pago = st.text_input(
                "Fecha talón",
                value=datos.get("fecha_pago", ""),
                key="ind_fecha"
            )

        st.markdown("### 3. Importes del talón")
        col4, col5, col6 = st.columns(3)

        with col4:
            percepciones = st.number_input(
                "Percepciones talón",
                value=float(datos["percepciones"]),
                step=100.0,
                format="%.2f",
                key="ind_percepciones"
            )

        with col5:
            descuentos = st.number_input(
                "Descuentos talón",
                value=float(datos["descuentos"]),
                step=100.0,
                format="%.2f",
                key="ind_descuentos"
            )

        with col6:
            liquido = st.number_input(
                "Líquido talón",
                value=float(datos["liquido"]),
                step=100.0,
                format="%.2f",
                key="ind_liquido"
            )

        datos_editados = {
            "nombre": nombre,
            "rfc": rfc,
            "fecha_pago": fecha_pago,
            "percepciones": percepciones,
            "descuentos": descuentos,
            "liquido": liquido,
            "codigos": datos["codigos"],
            "texto_original": datos.get("texto_original", "")
        }

        st.divider()
        st.markdown("### 4. Ajustes para revisión")
        promotor, qna, anio, semana = render_ajustes_revision("ind_")

        col_a, col_b, col_c = st.columns(3)

        with col_a:
            abono_extra = st.number_input(
                "Abono extra / apoyo adicional",
                min_value=0.0,
                value=0.0,
                step=500.0,
                format="%.2f",
                key="ind_abono"
            )

        with col_b:
            tiene_programado = st.selectbox(
                "¿Tiene programado?",
                ["No", "Sí"],
                key="ind_tiene_programado"
            )

        with col_c:
            if tiene_programado == "Sí":
                programado = st.number_input(
                    "Monto programado",
                    min_value=0.0,
                    value=0.0,
                    step=100.0,
                    format="%.2f",
                    key="ind_programado"
                )
            else:
                programado = 0.0
                st.number_input(
                    "Monto programado",
                    min_value=0.0,
                    value=0.0,
                    step=100.0,
                    format="%.2f",
                    disabled=True,
                    key="ind_programado_disabled"
                )

        if "ind_cuentas_terminadas" not in st.session_state:
            st.session_state.ind_cuentas_terminadas = []

        cuentas_terminadas, tiene_cuentas_terminadas = render_cuentas_terminadas(
            cuentas=st.session_state.ind_cuentas_terminadas,
            tiene_cuentas=st.session_state.get(
                "ind_cuentas_habilitadas",
                False
            ),
            key_prefix="ind_"
        )
        st.session_state.ind_cuentas_terminadas = cuentas_terminadas
        st.session_state.ind_cuentas_habilitadas = tiene_cuentas_terminadas

        st.markdown("### 6. Importes por código")
        codigos_manual = {}

        with st.expander("Mostrar / editar"):
            st.caption(
                "Se precargan con lo detectado. Corrige los importes si la lectura "
                "del talón no fue exacta."
            )
            columnas_codigos = st.columns(3)

            for indice, (cod, equiv) in enumerate(CODIGOS_FORMATO):
                with columnas_codigos[indice % 3]:
                    etiqueta = cod + (f" (PDF: {equiv})" if equiv else "")
                    valor = st.number_input(
                        etiqueta,
                        value=importe_detectado(datos["codigos"], cod, equiv),
                        step=100.0,
                        format="%.2f",
                        key=f"ind_codigo_{cod}"
                    )
                    codigos_manual[cod] = {
                        "descripcion": "manual",
                        "importe": valor
                    }

        st.divider()

        revision = calcular_revision_talon(
            codigos_extraidos=codigos_manual,
            descuentos_talon=descuentos,
            abono_extra=abono_extra,
            programado=programado,
            cuentas_terminadas=(
                cuentas_terminadas if tiene_cuentas_terminadas else []
            )
        )

        resultado_liquidez = generar_resultado_liquidez(
            revision=revision,
            tiene_programado=tiene_programado
        )

        mensaje = generar_mensaje_vendedor(
            datos=datos_editados,
            revision=revision,
            tiene_programado=tiene_programado
        )

        codigos_revision = revision["codigos_revision"]

        tabla_codigos = [
            {
                "Código formato": cod,
                "Equivale en PDF": equiv,
                "Importe": codigos_revision[cod]
            }
            for cod, equiv in CODIGOS_FORMATO
        ]

        with st.expander("Ver resumen de códigos utilizados"):
            st.dataframe(tabla_codigos, use_container_width=True)

        st.divider()
        st.markdown("### 7. Resultado de revisión")
        tiene_programado_consolidado = tiene_programado
        render_resumen_revision(revision, tiene_programado_consolidado, "ind_")

        st.divider()
        st.markdown("### 8. Mensaje generado")
        render_mensaje_vendedor(mensaje, "ind_")

        st.divider()
        st.markdown("### 9. Acciones")
        st.caption(
            "Genera los reportes de revisión. Al crear el Excel también se intenta "
            "subir el Excel, el PDF de revisión y el talón original."
        )

        with st.expander("Vista previa del Excel"):
            st.dataframe(
                pd.DataFrame([{
                    "Liquidez del talón": revision["liquidez_talon"],
                    "Apoyo adicional": revision["abono_extra"],
                    "Saldo liberado": revision["total_saldo_liberado"],
                    "Solo observado": revision.get("total_solo_observado", 0),
                    "Programado": revision["programado"],
                    "Liquidez final": revision["liquidez_final"]
                }]),
                use_container_width=True,
                hide_index=True
            )

            if revision.get("cuentas_terminadas"):
                st.dataframe(
                    pd.DataFrame([
                        {
                            "QNA termina": cuenta.get("qna_termina", ""),
                            "Saldo liberado": cuenta.get("saldo_liberado", 0),
                            "Sumar a liquidez": cuenta.get(
                                "sumar_a_liquidez",
                                False
                            ),
                            "Observación": cuenta.get("observacion", "")
                        }
                        for cuenta in revision["cuentas_terminadas"]
                    ]),
                    use_container_width=True,
                    hide_index=True
                )

        col_excel, col_pdf = st.columns(2)
        with col_excel:
            generar_excel = st.button(
                "Generar Excel de revisión",
                type="primary",
                key="ind_generar_excel",
                use_container_width=True
            )
        with col_pdf:
            generar_pdf = st.button(
                "Generar PDF",
                key="ind_generar_pdf",
                use_container_width=True
            )

        if generar_pdf:
            if not datos_editados["nombre"].strip() or not datos_editados["rfc"].strip():
                st.error("Captura al menos el Nombre y el RFC antes de generar el PDF.")
                st.stop()

            try:
                ruta_revision_pdf = generar_pdf_revision(
                    datos=datos_editados,
                    revision=revision,
                    mensaje_vendedor=mensaje,
                    promotor=promotor,
                    qna=qna
                )
                with open(ruta_revision_pdf, "rb") as archivo:
                    st.download_button(
                        label="Descargar PDF generado",
                        data=archivo,
                        file_name=Path(ruta_revision_pdf).name,
                        mime="application/pdf",
                        key="ind_descargar_pdf"
                    )
                st.success(f"PDF generado correctamente: {ruta_revision_pdf}")
            except Exception as error:
                st.error("Ocurrió un error al generar el PDF.")
                st.exception(error)

        if generar_excel:
            if not datos_editados["nombre"].strip() or not datos_editados["rfc"].strip():
                st.error("Captura al menos el Nombre y el RFC antes de generar el Excel.")
                st.stop()

            try:
                ruta_excel = generar_excel_revision(
                    datos=datos_editados,
                    revision=revision,
                    mensaje_vendedor=mensaje,
                    promotor=promotor,
                    qna=qna
                )
                ruta_revision_pdf = generar_pdf_revision(
                    datos=datos_editados,
                    revision=revision,
                    mensaje_vendedor=mensaje,
                    promotor=promotor,
                    qna=qna
                )

                with open(ruta_excel, "rb") as archivo:
                    st.download_button(
                        label="Descargar Excel generado",
                        data=archivo,
                        file_name=Path(ruta_excel).name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="ind_descargar_excel"
                    )

                with open(ruta_revision_pdf, "rb") as archivo:
                    st.download_button(
                        label="Descargar PDF de revisión",
                        data=archivo,
                        file_name=Path(ruta_revision_pdf).name,
                        mime="application/pdf",
                        key="ind_descargar_pdf_excel"
                    )

                st.success("Excel y PDF de revisión generados correctamente.")

                try:
                    resultado_graph = subir_revision_a_graph(
                        ruta_pdf=str(ruta_pdf) if ruta_pdf else None,
                        ruta_excel=ruta_excel,
                        anio=anio,
                        semana=semana,
                        promotor=promotor,
                        nombre_cliente=datos_editados["nombre"],
                        rfc=datos_editados["rfc"],
                        ruta_revision_pdf=ruta_revision_pdf
                    )

                    st.success("Archivos subidos correctamente a OneDrive/SharePoint.")
                    st.write(f"Ruta remota: `{resultado_graph['remote_folder_path']}`")

                    if resultado_graph.get("pdf_web_url"):
                        st.link_button(
                            "Abrir PDF en OneDrive",
                            resultado_graph["pdf_web_url"],
                            key="ind_link_pdf"
                        )

                    if resultado_graph.get("excel_web_url"):
                        st.link_button(
                            "Abrir Excel en OneDrive",
                            resultado_graph["excel_web_url"],
                            key="ind_link_excel"
                        )

                    if resultado_graph.get("revision_pdf_web_url"):
                        st.link_button(
                            "Abrir PDF de revisión en OneDrive",
                            resultado_graph["revision_pdf_web_url"],
                            key="ind_link_revision_pdf"
                        )

                except GraphStorageError as error:
                    st.warning(
                        "Los reportes se generaron localmente, pero no se pudieron "
                        "subir a OneDrive/SharePoint."
                    )
                    st.error(str(error))

                except Exception as error:
                    st.warning(
                        "Los reportes se generaron localmente, pero ocurrió un error inesperado "
                        "al subir a OneDrive/SharePoint."
                    )
                    st.exception(error)

            except FileNotFoundError:
                st.error(
                    "No se encontró la plantilla. Verifica que exista el archivo: "
                    "templates/plantilla_revision_talon.xlsx"
                )

            except Exception as error:
                st.error("Ocurrió un error al generar el Excel.")
                st.exception(error)

        with st.expander("Ver todos los códigos detectados del PDF"):
            tabla_todos = []

            for codigo, info in datos["codigos"].items():
                tabla_todos.append({
                    "Código": codigo,
                    "Descripción": info["descripcion"],
                    "Importe": info["importe"]
                })

            st.dataframe(tabla_todos, use_container_width=True)

        with st.expander("Ver texto extraído del PDF"):
            st.text(datos["texto_original"])

    else:
        st.info("Sube un talón (PDF o imagen) o activa la captura manual para comenzar.")


# =============================================================================
# TAB: VARIOS TALONES (LOTE)
# =============================================================================

with tab_lote:
    st.subheader("Revisión de uno o varios talones del mismo promotor")
    st.caption(
        "Sube 1 o más talones del mismo promotor. La app suma todo y muestra "
        "un resumen consolidado como en la pestaña Un talón."
    )

    promotor, qna, anio, semana = render_ajustes_revision("lote_")

    archivos_lote = st.file_uploader(
        "Sube uno o varios talones (PDF o imagen)",
        type=["pdf", "jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key="uploader_lote"
    )

    if archivos_lote:
        nombres_actuales = tuple(sorted(archivo.name for archivo in archivos_lote))

        if st.session_state.get("lote_nombres") != nombres_actuales:
            st.session_state.lote_nombres = nombres_actuales
            registros = []

            with st.spinner(f"Procesando {len(archivos_lote)} talón(es)..."):
                for archivo in archivos_lote:
                    registros.append(procesar_archivo_lote(archivo))

            st.session_state.lote_registros = registros

        registros = st.session_state.get("lote_registros", [])

        errores_lectura = [r for r in registros if r.get("error")]
        if errores_lectura:
            st.warning(
                f"{len(errores_lectura)} archivo(s) no se pudieron leer correctamente. "
                "Corrígelos manualmente en la tabla."
            )

        df_lote = pd.DataFrame([
            {
                "Procesar": r.get("procesar", True),
                "Archivo": r["archivo"],
                "Nombre": r["nombre"],
                "RFC": r["rfc"],
                "Descuentos": float(r["descuentos"]),
                "Abono extra": float(r.get("abono_extra", 0)),
                "Programado": float(r.get("programado", 0))
            }
            for r in registros
        ])

        st.subheader("Tabla editable")
        st.caption("Corrige nombre, RFC, descuentos, abono extra y programado antes de generar.")

        df_editado = st.data_editor(
            df_lote,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "Procesar": st.column_config.CheckboxColumn(
                    "Procesar",
                    help="Desmarca si no quieres generar este talón",
                    default=True
                ),
                "Archivo": st.column_config.TextColumn("Archivo", disabled=True),
                "Descuentos": st.column_config.NumberColumn("Descuentos", format="%.2f"),
                "Abono extra": st.column_config.NumberColumn("Abono extra", format="%.2f"),
                "Programado": st.column_config.NumberColumn("Programado", format="%.2f")
            },
            key="lote_data_editor"
        )

        for indice, fila in df_editado.iterrows():
            registros[indice]["nombre"] = str(fila["Nombre"]).strip()
            registros[indice]["rfc"] = str(fila["RFC"]).strip()
            registros[indice]["descuentos"] = float(fila["Descuentos"])
            registros[indice]["abono_extra"] = float(fila["Abono extra"])
            registros[indice]["programado"] = float(fila["Programado"])
            registros[indice]["procesar"] = bool(fila["Procesar"])

        st.session_state.lote_registros = registros

        with st.expander("Corregir códigos y cuentas de un talón"):
            opciones_archivo = [r["archivo"] for r in registros]
            archivo_seleccionado = st.selectbox(
                "Selecciona el talón",
                opciones_archivo,
                key="lote_archivo_codigos"
            )

            indice_sel = opciones_archivo.index(archivo_seleccionado)
            registro_sel = registros[indice_sel]

            cols = st.columns(4)
            codigos_actualizados = dict(registro_sel.get("codigos", {}))

            for idx, (cod, equiv) in enumerate(CODIGOS_FORMATO):
                with cols[idx % 4]:
                    valor_actual = importe_detectado(codigos_actualizados, cod, equiv)
                    nuevo_valor = st.number_input(
                        cod + (f" ({equiv})" if equiv else ""),
                        value=float(valor_actual),
                        step=100.0,
                        format="%.2f",
                        key=f"lote_cod_{indice_sel}_{cod}"
                    )
                    codigos_actualizados[cod] = {
                        "descripcion": "manual",
                        "importe": nuevo_valor
                    }

            registros[indice_sel]["codigos"] = codigos_actualizados
            st.divider()

            cuentas_registro, tiene_cuentas_registro = render_cuentas_terminadas(
                cuentas=registro_sel.get("cuentas_terminadas", []),
                tiene_cuentas=registro_sel.get(
                    "tiene_cuentas_terminadas",
                    False
                ),
                key_prefix=f"lote_{indice_sel}_"
            )
            registros[indice_sel]["cuentas_terminadas"] = cuentas_registro
            registros[indice_sel][
                "tiene_cuentas_terminadas"
            ] = tiene_cuentas_registro
            st.session_state.lote_registros = registros

        revision_consolidada = calcular_revision_consolidada(registros)

        if revision_consolidada:
            cantidad_activos = len(registros_activos(registros))
            tiene_prog_total = tiene_programado_desde_monto(revision_consolidada["programado"])
            mensaje_lote = generar_mensaje_vendedor_lote(
                promotor=promotor,
                registros=registros,
                revision=revision_consolidada
            )

            st.divider()
            st.subheader(
                f"Resumen consolidado ({cantidad_activos} talón"
                f"{'es' if cantidad_activos != 1 else ''})"
            )
            st.caption(
                "Suma de todos los talones marcados con Procesar. "
                "El cálculo se actualiza al instante."
            )

            st.subheader("Códigos usados en tu formato (sumados)")
            codigos_revision = revision_consolidada["codigos_revision"]

            tabla_codigos = [
                {
                    "Código formato": cod,
                    "Equivale en PDF": equiv,
                    "Importe": codigos_revision[cod]
                }
                for cod, equiv in CODIGOS_FORMATO
            ]

            st.dataframe(tabla_codigos, use_container_width=True)
            st.divider()

            render_resumen_revision(revision_consolidada, tiene_prog_total, "lote_")

            st.divider()
            render_mensaje_vendedor(mensaje_lote, "lote_")

            with st.expander("Detalle por cliente"):
                resumen_filas = []

                for registro in registros_activos(registros):
                    revision_ind = calcular_revision_desde_registro(registro)
                    tiene_prog = tiene_programado_desde_monto(registro.get("programado", 0))
                    resultado = generar_resultado_liquidez(revision_ind, tiene_prog)

                    resumen_filas.append({
                        "Archivo": registro["archivo"],
                        "Nombre": registro["nombre"],
                        "RFC": registro["rfc"],
                        "Liquidez": formato_moneda(revision_ind["liquidez_final"]),
                        "Resultado": resultado,
                        "Error lectura": registro.get("error") or ""
                    })

                st.dataframe(
                    pd.DataFrame(resumen_filas),
                    use_container_width=True,
                    hide_index=True
                )
        else:
            st.warning("Marca al menos un talón con Procesar para ver el resumen consolidado.")

        st.divider()

        if st.button("Generar y subir todos", type="primary", key="lote_generar_todos"):
            resultados_proceso = []
            procesados = 0
            exitosos = 0

            for registro in registros:
                if not registro.get("procesar", True):
                    continue

                procesados += 1
                nombre_cliente = registro.get("nombre", "").strip()
                rfc_cliente = registro.get("rfc", "").strip()

                if not nombre_cliente or not rfc_cliente:
                    resultados_proceso.append({
                        "Archivo": registro["archivo"],
                        "Estado": "Error",
                        "Detalle": "Falta Nombre o RFC"
                    })
                    continue

                try:
                    revision = calcular_revision_desde_registro(registro)
                    tiene_prog = tiene_programado_desde_monto(registro.get("programado", 0))
                    datos_excel = registro_a_datos_excel(registro)
                    mensaje = generar_mensaje_vendedor(
                        datos=datos_excel,
                        revision=revision,
                        tiene_programado=tiene_prog
                    )

                    ruta_excel = generar_excel_revision(
                        datos=datos_excel,
                        revision=revision,
                        mensaje_vendedor=mensaje,
                        promotor=promotor,
                        qna=qna
                    )
                    ruta_revision_pdf = generar_pdf_revision(
                        datos=datos_excel,
                        revision=revision,
                        mensaje_vendedor=mensaje,
                        promotor=promotor,
                        qna=qna
                    )

                    detalle = (
                        f"Excel: {Path(ruta_excel).name} | "
                        f"PDF: {Path(ruta_revision_pdf).name}"
                    )

                    try:
                        resultado_graph = subir_revision_a_graph(
                            ruta_pdf=registro.get("ruta"),
                            ruta_excel=ruta_excel,
                            anio=anio,
                            semana=semana,
                            promotor=promotor,
                            nombre_cliente=nombre_cliente,
                            rfc=rfc_cliente,
                            ruta_revision_pdf=ruta_revision_pdf
                        )

                        detalle += f" | Remoto: {resultado_graph['remote_folder_path']}"
                        exitosos += 1
                        resultados_proceso.append({
                            "Archivo": registro["archivo"],
                            "Estado": "OK",
                            "Detalle": detalle
                        })

                    except GraphStorageError as error:
                        exitosos += 1
                        resultados_proceso.append({
                            "Archivo": registro["archivo"],
                            "Estado": "Excel OK / Graph falló",
                            "Detalle": str(error)
                        })

                except Exception as error:
                    resultados_proceso.append({
                        "Archivo": registro["archivo"],
                        "Estado": "Error",
                        "Detalle": str(error)
                    })

            if procesados == 0:
                st.warning("No hay talones marcados para procesar.")
            else:
                st.success(f"Procesados: {procesados} | Exitosos: {exitosos}")
                st.dataframe(pd.DataFrame(resultados_proceso), use_container_width=True, hide_index=True)

    else:
        st.info("Sube uno o varios talones del mismo promotor para comenzar.")


with tab_refinanciamiento:
    st.markdown("## Calculadora de refinanciamiento")
    st.caption(
        "Captura las facturas actuales, valida el porcentaje pagado y simula "
        "el nuevo saldo por plazo."
    )

    st.markdown("### 1. Datos del cliente")
    with st.container(border=True):
        col_fecha, col_cliente = st.columns([1, 2])

        with col_fecha:
            fecha_refinanciamiento = st.date_input(
                "Fecha",
                value=datetime.now().date(),
                format="DD/MM/YYYY",
                key="ref_fecha"
            )

        with col_cliente:
            cliente_refinanciamiento = st.text_input(
                "Cliente",
                key="ref_cliente"
            )

        col_rfc, col_quinquenio, col_aumento = st.columns(3)

        with col_rfc:
            rfc_nac = st.text_input(
                "RFC/NAC",
                placeholder="J860901UT6",
                key="ref_rfc_nac"
            )

        with col_quinquenio:
            quinquenio = st.number_input(
                "Quinquenio",
                min_value=0,
                value=0,
                step=1,
                key="ref_quinquenio"
            )

        with col_aumento:
            aumento_descuento = st.number_input(
                "Liquidez / aumento en descuento",
                value=0.0,
                step=100.0,
                format="%.2f",
                key="ref_aumento_descuento"
            )

        forzar_1900 = st.checkbox(
            "Forzar siglo 1900",
            value=True,
            help=(
                "Activado: 01 se interpreta como 1901. "
                "Desactivado: se elige 1900 o 2000 según el año actual."
            ),
            key="ref_forzar_1900"
        )
        nacimiento = extraer_fecha_edad_desde_rfc(
            rfc_nac,
            forzar_1900=forzar_1900
        )

        col_nacimiento, col_edad = st.columns(2)
        with col_nacimiento:
            st.metric(
                "Fecha de nacimiento",
                value=(
                    nacimiento["fecha_nacimiento"].strftime("%d/%m/%Y")
                    if nacimiento["fecha_nacimiento"]
                    else "No disponible"
                )
            )
        with col_edad:
            st.metric(
                "Edad",
                value=(
                    str(nacimiento["edad"])
                    if nacimiento["edad"] is not None
                    else "No disponible"
                )
            )

        if rfc_nac and nacimiento["fecha_nacimiento"] is None:
            st.warning(
                "No se pudo calcular la edad desde el RFC/NAC. Verifica el dato."
            )

    st.markdown("### 2. Facturas actuales")
    st.caption(
        "Agrega o elimina filas directamente en la tabla. "
        "Las columnas sombreadas se calculan automáticamente."
    )

    if "ref_facturas" not in st.session_state:
        st.session_state.ref_facturas = dataframe_facturas_vacio()

    facturas_calculadas = calcular_facturas_refinanciamiento(
        st.session_state.ref_facturas
    )
    facturas_editadas = st.data_editor(
        facturas_calculadas,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        disabled=COLUMNAS_CALCULADAS,
        column_config={
            "FACT": st.column_config.TextColumn("FACT"),
            "VTA": st.column_config.NumberColumn("VTA", format="$ %.2f"),
            "PAGADO": st.column_config.NumberColumn("PAGADO", format="$ %.2f"),
            "SALDO": st.column_config.NumberColumn("SALDO", format="$ %.2f"),
            "QNAS TOMADAS A CUENTA": st.column_config.NumberColumn(
                "QNAS TOMADAS A CUENTA",
                min_value=0,
                step=1,
                format="%d"
            ),
            "ABONO": st.column_config.NumberColumn("ABONO", format="$ %.2f"),
            "ABONO DE QUINCENAS": st.column_config.NumberColumn(
                "ABONO DE QUINCENAS",
                format="$ %.2f"
            ),
            "SALDO PENDIENTE": st.column_config.NumberColumn(
                "SALDO PENDIENTE",
                format="$ %.2f"
            ),
            "PORCENTAJE PAGADO": st.column_config.NumberColumn(
                "PORCENTAJE PAGADO",
                format="percent"
            ),
            "REFINANCIAMIENTO": st.column_config.NumberColumn(
                "REFINANCIAMIENTO",
                format="$ %.2f"
            ),
            "PUEDE REFINANCIAR": st.column_config.TextColumn(
                "PUEDE REFINANCIAR"
            )
        },
        key="ref_data_editor"
    )
    st.session_state.ref_facturas = facturas_editadas[
        [
            "FACT",
            "VTA",
            "PAGADO",
            "SALDO",
            "QNAS TOMADAS A CUENTA",
            "ABONO",
            "EN COBRO"
        ]
    ].copy()

    facturas_resultado = calcular_facturas_refinanciamiento(facturas_editadas)
    resumen_refinanciamiento = calcular_resumen_refinanciamiento(
        facturas_resultado,
        aumento_descuento
    )

    st.markdown("### 3. Totales principales")
    col_total_vta, col_total_pagado, col_total_saldo = st.columns(3)
    with col_total_vta:
        st.metric(
            "Total VTA",
            formato_moneda(resumen_refinanciamiento["total_vta"])
        )
    with col_total_pagado:
        st.metric(
            "Total pagado",
            formato_moneda(resumen_refinanciamiento["total_pagado"])
        )
    with col_total_saldo:
        st.metric(
            "Total saldo",
            formato_moneda(resumen_refinanciamiento["total_saldo"])
        )

    col_total_abono, col_abono_qnas, col_saldo_pendiente, col_total_ref = (
        st.columns(4)
    )
    with col_total_abono:
        st.metric(
            "Total abono",
            formato_moneda(resumen_refinanciamiento["total_abono"])
        )
    with col_abono_qnas:
        st.metric(
            "Abono de quincenas",
            formato_moneda(
                resumen_refinanciamiento["total_abono_quincenas"]
            )
        )
    with col_saldo_pendiente:
        st.metric(
            "Saldo pendiente",
            formato_moneda(
                resumen_refinanciamiento["total_saldo_pendiente"]
            )
        )
    with col_total_ref:
        st.metric(
            "Refinanciamiento",
            formato_moneda(
                resumen_refinanciamiento["total_refinanciamiento"]
            )
        )

    st.markdown("### 4. Resumen de refinanciamiento")
    with st.container(border=True):
        col_abono_ref, col_abono_antes, col_abono_nuevo = st.columns(3)
        with col_abono_ref:
            st.metric(
                "ABONO REF",
                formato_moneda(resumen_refinanciamiento["abono_ref"])
            )
        with col_abono_antes:
            st.metric(
                "ABONO ANTES",
                formato_moneda(resumen_refinanciamiento["abono_antes"])
            )
        with col_abono_nuevo:
            st.metric(
                "TOTAL ABONO NUEVO",
                formato_moneda(
                    resumen_refinanciamiento["total_abono_nuevo"]
                )
            )

        simulacion_df = dataframe_simulacion(resumen_refinanciamiento)
        st.dataframe(
            simulacion_df.style.format("${:,.2f}"),
            use_container_width=True
        )

    st.markdown("### 5. Exportar")
    datos_cliente_refinanciamiento = {
        "fecha": fecha_refinanciamiento,
        "cliente": cliente_refinanciamiento,
        "rfc_nac": rfc_nac,
        "fecha_nacimiento": nacimiento["fecha_nacimiento"],
        "edad": nacimiento["edad"],
        "quinquenio": quinquenio
    }
    excel_refinanciamiento = generar_excel_refinanciamiento(
        facturas=facturas_resultado,
        resumen=resumen_refinanciamiento,
        datos_cliente=datos_cliente_refinanciamiento
    )
    nombre_cliente_ref = (
        cliente_refinanciamiento.strip().upper().replace(" ", "_")
        or "CLIENTE"
    )
    st.download_button(
        "Exportar refinanciamiento a Excel",
        data=excel_refinanciamiento,
        file_name=f"REFINANCIAMIENTO_{nombre_cliente_ref}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        key="ref_descargar_excel"
    )
