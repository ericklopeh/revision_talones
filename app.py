import streamlit as st
from pathlib import Path
from datetime import datetime

from services.extractor_pdf import extraer_datos_talon
from services.calculadora import calcular_revision_talon
from services.generador_excel import generar_excel_revision
from services.graph_storage import subir_revision_a_graph, GraphStorageError


st.set_page_config(
    page_title="Revisión de Talones",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Revisión de Talones")
st.caption("Sistema para leer talón, calcular liquidez, generar mensaje y crear Excel de revisión.")

archivo_pdf = st.file_uploader(
    "Sube un talón en PDF",
    type=["pdf"]
)


def formato_moneda(valor: float) -> str:
    if valor < 0:
        return f"-${abs(valor):,.2f}"

    return f"${valor:,.2f}"


def generar_resultado_liquidez(revision: dict, tiene_programado: str) -> str:
    liquidez_final = revision["liquidez_final"]
    programado = revision["programado"]

    if liquidez_final > 0:
        return f"Tiene liquidez de {formato_moneda(liquidez_final)}."

    if liquidez_final < 0 and tiene_programado == "Sí" and programado > 0:
        return (
            f"Tiene un sobregiro de {formato_moneda(liquidez_final)}. "
            f"Tiene un programado por {formato_moneda(programado)}."
        )

    if liquidez_final < 0:
        return f"Tiene un sobregiro de {formato_moneda(liquidez_final)}."

    return "No tiene liquidez disponible."


def generar_mensaje_vendedor(datos: dict, revision: dict, tiene_programado: str) -> str:
    nombre = datos["nombre"]
    rfc = datos["rfc"]

    resultado_liquidez = generar_resultado_liquidez(
        revision=revision,
        tiene_programado=tiene_programado
    )

    texto = f"""Se realizó la revisión del talón correspondiente al cliente:

Cliente: {nombre}
RFC: {rfc}

Resultado de la revisión:
{resultado_liquidez}"""

    return texto


if archivo_pdf:
    Path("uploads").mkdir(exist_ok=True)
    Path("output").mkdir(exist_ok=True)

    ruta_pdf = Path("uploads") / archivo_pdf.name

    with open(ruta_pdf, "wb") as archivo:
        archivo.write(archivo_pdf.getbuffer())

    datos = extraer_datos_talon(str(ruta_pdf))

    st.success("Talón leído correctamente.")

    # =========================
    # DATOS DEL TALÓN
    # =========================

    st.subheader("Datos del talón")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Nombre", datos["nombre"])

    with col2:
        st.metric("RFC", datos["rfc"])

    with col3:
        st.metric("Líquido talón", formato_moneda(datos["liquido"]))

    col4, col5, col6 = st.columns(3)

    with col4:
        st.metric("Percepciones talón", formato_moneda(datos["percepciones"]))

    with col5:
        st.metric("Descuentos talón", formato_moneda(datos["descuentos"]))

    with col6:
        st.metric("Fecha talón", datos.get("fecha_pago", ""))

    st.divider()

    # =========================
    # AJUSTES PARA REVISIÓN
    # =========================

    st.subheader("Ajustes para revisión")

    col_v, col_q, col_anio, col_semana = st.columns(4)

    with col_v:
        promotor = st.selectbox(
            "Promotor / vendedor",
            [
                "Victor Vega",
                "Juan Manuel",
                "Leonardo Arevalo",
                "Eliezer Chipuli",
                "Gerardo Santana",
                "Sergio Valadez",
                "Sergio Vazquez"
            ]
        )

    with col_q:
        qna = st.text_input(
            "QNA",
            value="09-2026"
        )

    with col_anio:
        anio = st.number_input(
            "Año",
            min_value=2024,
            max_value=2035,
            value=datetime.now().year,
            step=1
        )

    with col_semana:
        semana = st.number_input(
            "Semana",
            min_value=1,
            max_value=53,
            value=datetime.now().isocalendar().week,
            step=1
        )

    col_a, col_b, col_c = st.columns(3)

    with col_a:
        abono_extra = st.number_input(
            "Abono extra / apoyo adicional",
            min_value=0.0,
            value=0.0,
            step=500.0,
            format="%.2f"
        )

    with col_b:
        tiene_programado = st.selectbox(
            "¿Tiene programado?",
            ["No", "Sí"]
        )

    with col_c:
        if tiene_programado == "Sí":
            programado = st.number_input(
                "Monto programado",
                min_value=0.0,
                value=0.0,
                step=100.0,
                format="%.2f"
            )
        else:
            programado = 0.0
            st.number_input(
                "Monto programado",
                min_value=0.0,
                value=0.0,
                step=100.0,
                format="%.2f",
                disabled=True
            )

    # =========================
    # CÁLCULO
    # =========================

    revision = calcular_revision_talon(
        codigos_extraidos=datos["codigos"],
        descuentos_talon=datos["descuentos"],
        abono_extra=abono_extra,
        programado=programado
    )

    resultado_liquidez = generar_resultado_liquidez(
        revision=revision,
        tiene_programado=tiene_programado
    )

    mensaje = generar_mensaje_vendedor(
        datos=datos,
        revision=revision,
        tiene_programado=tiene_programado
    )

    st.divider()

    # =========================
    # CÓDIGOS USADOS
    # =========================

    st.subheader("Códigos usados en tu formato")

    codigos_revision = revision["codigos_revision"]

    tabla_codigos = [
        {
            "Código formato": "E4",
            "Equivale en PDF": "",
            "Importe": codigos_revision["E4"]
        },
        {
            "Código formato": "E3",
            "Equivale en PDF": "",
            "Importe": codigos_revision["E3"]
        },
        {
            "Código formato": "Q",
            "Equivale en PDF": "A2",
            "Importe": codigos_revision["Q"]
        },
        {
            "Código formato": "CP",
            "Equivale en PDF": "",
            "Importe": codigos_revision["CP"]
        },
        {
            "Código formato": "7",
            "Equivale en PDF": "07",
            "Importe": codigos_revision["7"]
        },
        {
            "Código formato": "CT",
            "Equivale en PDF": "CT",
            "Importe": codigos_revision["CT"]
        },
        {
            "Código formato": "7B",
            "Equivale en PDF": "",
            "Importe": codigos_revision["7B"]
        },
        {
            "Código formato": "E9",
            "Equivale en PDF": "",
            "Importe": codigos_revision["E9"]
        },
        {
            "Código formato": "SG",
            "Equivale en PDF": "SG",
            "Importe": codigos_revision["SG"]
        },
        {
            "Código formato": "O1",
            "Equivale en PDF": "01",
            "Importe": codigos_revision["O1"]
        },
        {
            "Código formato": "DC",
            "Equivale en PDF": "DC",
            "Importe": codigos_revision["DC"]
        }
    ]

    st.dataframe(tabla_codigos, use_container_width=True)

    st.divider()

    # =========================
    # CÁLCULO DE REVISIÓN
    # =========================

    st.subheader("Cálculo de revisión")

    col7, col8, col9 = st.columns(3)

    with col7:
        st.metric("Ingresos revisión", formato_moneda(revision["ingresos"]))

    with col8:
        st.metric("Descuentos", formato_moneda(revision["descuentos"]))

    with col9:
        st.metric("Saldo al 100", formato_moneda(revision["saldo_100"]))

    col10, col11, col12 = st.columns(3)

    with col10:
        st.metric("Total para venta 70%", formato_moneda(revision["total_para_venta_70"]))

    with col11:
        st.metric("Saldo al 70%", formato_moneda(revision["saldo_70"]))

    with col12:
        st.metric("Liquidez final", formato_moneda(revision["liquidez_final"]))

    st.info(resultado_liquidez)

    st.markdown(
        f"""
        ### Resumen

        **Ingresos revisión:** {formato_moneda(revision["ingresos"])}  
        **Descuentos:** {formato_moneda(revision["descuentos"])}  
        **Saldo al 100:** {formato_moneda(revision["saldo_100"])}  
        **Saldo al 70:** {formato_moneda(revision["saldo_70"])}  
        **Abono extra / apoyo adicional:** {formato_moneda(revision["abono_extra"])}  
        **Programado:** {formato_moneda(revision["programado"])}  
        **Liquidez final:** {formato_moneda(revision["liquidez_final"])}  
        **Resultado:** {resultado_liquidez}
        """
    )

    st.divider()

    # =========================
    # MENSAJE PARA VENDEDOR
    # =========================

    st.subheader("Mensaje para vendedor")

    st.text_area(
        "Texto formal para copiar y enviar",
        value=mensaje,
        height=220
    )

    st.divider()

    # =========================
    # GENERAR EXCEL
    # =========================

    st.subheader("Generar Excel de revisión")

    if st.button("Generar Excel", type="primary"):
        try:
            ruta_excel = generar_excel_revision(
                datos=datos,
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
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.success(f"Excel generado correctamente: {ruta_excel}")

            try:
                resultado_graph = subir_revision_a_graph(
                    ruta_pdf=str(ruta_pdf),
                    ruta_excel=ruta_excel,
                    anio=int(anio),
                    semana=int(semana),
                    promotor=promotor,
                    nombre_cliente=datos["nombre"],
                    rfc=datos["rfc"]
                )

                st.success("Archivos subidos correctamente a OneDrive/SharePoint.")
                st.write(f"Ruta remota: `{resultado_graph['remote_folder_path']}`")

                if resultado_graph.get("pdf_web_url"):
                    st.link_button("Abrir PDF en OneDrive", resultado_graph["pdf_web_url"])

                if resultado_graph.get("excel_web_url"):
                    st.link_button("Abrir Excel en OneDrive", resultado_graph["excel_web_url"])

            except GraphStorageError as error:
                st.warning("El Excel se generó localmente, pero no se pudo subir a OneDrive/SharePoint.")
                st.error(str(error))

            except Exception as error:
                st.warning("El Excel se generó localmente, pero ocurrió un error inesperado al subir a OneDrive/SharePoint.")
                st.exception(error)

        except FileNotFoundError:
            st.error(
                "No se encontró la plantilla. Verifica que exista el archivo: "
                "templates/plantilla_revision_talon.xlsx"
            )

        except Exception as error:
            st.error("Ocurrió un error al generar el Excel.")
            st.exception(error)

    # =========================
    # DEBUG / REVISIÓN
    # =========================

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
    st.info("Sube un talón en PDF para comenzar.")