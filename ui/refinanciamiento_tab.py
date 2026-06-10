from datetime import datetime

import pandas as pd
import streamlit as st

from services.refinanciamiento_db import RefinanciamientoDatabaseError
from utils.refinanciamiento import (
    OUTPUT_REFINANCIAMIENTO_DIR,
    calcular_facturas_refinanciamiento,
    calcular_resumen_refinanciamiento,
    dataframe_simulacion,
    extraer_fecha_edad_desde_rfc,
    generar_excel_refinanciamiento,
    generar_json_refinanciamiento,
    generar_pdf_refinanciamiento,
    guardar_archivos_refinanciamiento,
    nombre_archivo_refinanciamiento,
    preparar_facturas_desde_bd
)


def render_refinanciamiento(
    promotores,
    formato_moneda,
    limpiar_estado,
    database_url,
    buscar_clientes_cache,
    cargar_facturas_cache
):
    st.markdown("## Calculadora de refinanciamiento")
    st.caption(
        "Busca al cliente en PostgreSQL, selecciona sus facturas y simula "
        "el nuevo saldo por plazo."
    )
    st.button(
        "Limpiar toda la captura",
        key="refi_limpiar_todo",
        on_click=limpiar_estado
    )
    if st.session_state.pop("refi_limpieza_confirmada", False):
        st.success("La captura de refinanciamiento quedó limpia.")

    st.markdown("### A. Datos del cliente")
    with st.container(border=True):
        col_busqueda, col_boton = st.columns([4, 1])
        with col_busqueda:
            busqueda = st.text_input(
                "Buscar por nombre o RFC",
                placeholder="Escribe al menos 2 caracteres",
                key="ref_busqueda_cliente"
            )
        with col_boton:
            st.write("")
            buscar_click = st.button(
                "Buscar",
                use_container_width=True,
                key="ref_buscar_cliente"
            )

        if buscar_click:
            if not database_url:
                st.error(
                    "Falta configurar DATABASE_URL para consultar PostgreSQL."
                )
            elif len(busqueda.strip()) < 2:
                st.warning("Escribe al menos 2 caracteres para buscar.")
            else:
                try:
                    st.session_state["ref_resultados_clientes"] = (
                        buscar_clientes_cache(database_url, busqueda)
                    )
                except RefinanciamientoDatabaseError as error:
                    st.error(str(error))

        resultados = st.session_state.get(
            "ref_resultados_clientes",
            pd.DataFrame()
        )
        if isinstance(resultados, pd.DataFrame) and not resultados.empty:
            st.dataframe(
                resultados[[
                    "cliente",
                    "rfc",
                    "facturas_encontradas",
                    "saldo_total"
                ]].rename(columns={
                    "cliente": "Cliente",
                    "rfc": "RFC",
                    "facturas_encontradas": "Facturas encontradas",
                    "saldo_total": "Saldo total"
                }).style.format({"Saldo total": "${:,.2f}"}),
                use_container_width=True,
                hide_index=True
            )
            clientes = resultados.to_dict(orient="records")
            indice = st.selectbox(
                "Cliente encontrado",
                range(len(clientes)),
                format_func=lambda posicion: (
                    f"{clientes[posicion]['cliente']} | "
                    f"{clientes[posicion]['rfc']} | "
                    f"{formato_moneda(float(clientes[posicion]['saldo_total']))}"
                ),
                key="ref_cliente_seleccion"
            )
            if st.button(
                "Seleccionar cliente y cargar facturas",
                type="primary",
                key="ref_cargar_facturas"
            ):
                cliente = clientes[indice]
                try:
                    facturas_bd = cargar_facturas_cache(
                        database_url,
                        cliente["cliente_id"]
                    )
                    st.session_state["ref_cliente_actual"] = cliente
                    st.session_state["refi_facturas_db"] = (
                        preparar_facturas_desde_bd(facturas_bd)
                    )
                    st.session_state["refi_editor_version"] = (
                        st.session_state.get("refi_editor_version", 0) + 1
                    )
                    st.rerun()
                except RefinanciamientoDatabaseError as error:
                    st.error(str(error))
        elif buscar_click and database_url:
            st.info("No se encontraron clientes con saldo pendiente.")

        cliente_actual = st.session_state.get("ref_cliente_actual")
        cliente_nombre = (
            str(cliente_actual["cliente"]) if cliente_actual else ""
        )
        rfc = str(cliente_actual["rfc"]) if cliente_actual else ""
        if cliente_actual:
            st.success(f"Cliente seleccionado: {cliente_nombre} | {rfc}")

        col_fecha, col_vendedor, col_semana = st.columns(3)
        fecha = col_fecha.date_input(
            "Fecha",
            value=datetime.now().date(),
            format="DD/MM/YYYY",
            key="ref_fecha"
        )
        vendedor = col_vendedor.selectbox(
            "Vendedor",
            promotores,
            key="ref_vendedor"
        )
        semana = col_semana.number_input(
            "Semana",
            min_value=1,
            max_value=53,
            value=datetime.now().isocalendar().week,
            step=1,
            key="ref_semana"
        )
        col_quinquenio, col_aumento = st.columns(2)
        quinquenio = col_quinquenio.number_input(
            "Quinquenio",
            min_value=0,
            value=0,
            step=1,
            key="ref_quinquenio"
        )
        aumento = col_aumento.number_input(
            "Liquidez / aumento en descuento",
            value=0.0,
            step=100.0,
            format="%.2f",
            key="ref_aumento_descuento"
        )
        forzar_1900 = st.checkbox(
            "Forzar siglo 1900 para fecha del RFC",
            value=True,
            key="ref_forzar_1900"
        )
        nacimiento = extraer_fecha_edad_desde_rfc(
            rfc,
            forzar_1900=forzar_1900
        )

    facturas_base = st.session_state.get("refi_facturas_db")
    if not isinstance(facturas_base, pd.DataFrame) or facturas_base.empty:
        st.info("Busca y selecciona un cliente para cargar sus facturas.")
        return

    facturas_calculadas = calcular_facturas_refinanciamiento(facturas_base)
    resumen_inicial = calcular_resumen_refinanciamiento(
        facturas_calculadas,
        aumento
    )
    aptas = facturas_calculadas["ESTATUS"].eq("APTA")
    no_aptas = ~aptas

    st.markdown("### B. Resumen del Cliente")
    with st.container(border=True):
        st.markdown(f"**{cliente_nombre}**  \nRFC: `{rfc or 'Sin RFC'}`")
        col_facturas, col_saldo, col_vta = st.columns(3)
        col_facturas.metric(
            "Facturas encontradas",
            len(facturas_calculadas)
        )
        col_saldo.metric(
            "Saldo total",
            formato_moneda(float(facturas_calculadas["SALDO"].sum()))
        )
        col_vta.metric(
            "Total VTA",
            formato_moneda(float(facturas_calculadas["VTA"].sum()))
        )
        col_aptas, col_no_aptas, col_venta = st.columns(3)
        col_aptas.metric("Facturas aptas", int(aptas.sum()))
        col_no_aptas.metric("Facturas no aptas", int(no_aptas.sum()))
        col_venta.metric(
            "Venta posible estimada (72)",
            formato_moneda(
                resumen_inicial["simulacion"][72]["VENTA POSIBLE"]
            )
        )

    st.markdown("### C. Facturas Encontradas")
    st.caption(
        "FACT, VTA y SALDO vienen de PostgreSQL. Solo puedes cambiar "
        "INCLUIR, quincenas, ABONO y EN COBRO."
    )
    version = st.session_state.get("refi_editor_version", 0)
    facturas_editadas = st.data_editor(
        facturas_calculadas,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_order=[
            "INCLUIR",
            "FACT",
            "VTA",
            "PAGADO",
            "SALDO",
            "QNAS TOMADAS A CUENTA",
            "ABONO",
            "ABONO DE QUINCENAS CONS",
            "SALDO PENDIENTE",
            "EN COBRO",
            "PORCENTAJE PAGADO",
            "ESTATUS"
        ],
        disabled=[
            "FACT", "VTA", "PAGADO", "SALDO",
            "ABONO DE QUINCENAS CONS", "SALDO PENDIENTE",
            "PORCENTAJE PAGADO", "REFINANCIAMIENTO",
            "PUEDE REFINANCIAR", "ESTATUS", "MOTIVO"
        ],
        column_config={
            "INCLUIR": st.column_config.CheckboxColumn("INCLUIR"),
            "VTA": st.column_config.NumberColumn("VTA", format="$ %.2f"),
            "PAGADO": st.column_config.NumberColumn(
                "PAGADO", format="$ %.2f"
            ),
            "SALDO": st.column_config.NumberColumn(
                "SALDO", format="$ %.2f"
            ),
            "ABONO": st.column_config.NumberColumn(
                "ABONO", format="$ %.2f"
            ),
            "QNAS TOMADAS A CUENTA": st.column_config.NumberColumn(
                "QNAS TOMADAS A CUENTA", min_value=0, step=1, format="%d"
            ),
            "PORCENTAJE PAGADO": st.column_config.ProgressColumn(
                "% PAGADO", min_value=0.0, max_value=1.0, format="%.0f%%"
            )
        },
        key=f"refi_facturas_editor_{version}"
    )
    facturas = calcular_facturas_refinanciamiento(facturas_editadas)
    st.session_state["refi_facturas_db"] = facturas.copy()

    saldos_anomalos = facturas["ESTATUS"].eq("REVISAR SALDO")
    if saldos_anomalos.any():
        st.warning(
            f"Hay {int(saldos_anomalos.sum())} factura(s) con SALDO mayor "
            "que VTA. Se marcaron como REVISAR SALDO y no se incluyeron "
            "automáticamente."
        )

    no_aptas = facturas["PUEDE REFINANCIAR"].eq("NO")
    if st.button(
        f"Quitar no aptas ({int(no_aptas.sum())})",
        key="ref_quitar_no_aptas"
    ):
        facturas.loc[no_aptas, "INCLUIR"] = False
        st.session_state["refi_facturas_db"] = facturas
        st.session_state["refi_editor_version"] = version + 1
        st.session_state["refi_no_aptas_quitadas"] = int(no_aptas.sum())
        st.rerun()

    quitadas = st.session_state.pop("refi_no_aptas_quitadas", None)
    if quitadas is not None:
        st.success(
            f"Se excluyeron {quitadas} facturas por tener menos del 40% "
            "pagado o requerir revisión de saldo."
        )

    st.markdown("### D. Facturas No Aptas / Excluidas")
    excluidas = facturas[
        (~facturas["INCLUIR"]) | facturas["PUEDE REFINANCIAR"].eq("NO")
    ]
    if excluidas.empty:
        st.success("No hay facturas excluidas ni por debajo del 40%.")
    else:
        st.caption(
            "Puedes volver a incluirlas activando INCLUIR en la tabla "
            "anterior. El estado APTA/NO APTA no se altera."
        )
        st.dataframe(
            excluidas[[
                "FACT", "VTA", "SALDO", "PORCENTAJE PAGADO",
                "INCLUIR", "MOTIVO"
            ]].style.format({
                "VTA": "${:,.2f}",
                "SALDO": "${:,.2f}",
                "PORCENTAJE PAGADO": "{:.0%}"
            }),
            use_container_width=True,
            hide_index=True
        )

    resumen = calcular_resumen_refinanciamiento(facturas, aumento)
    st.markdown("### E. Simulación de Refinanciamiento")
    simulacion = dataframe_simulacion(resumen)
    filas_destacadas = {
        "VENTA POSIBLE", "DESCUENTO NUEVO", "TOTAL ADEUDO CLIENTE"
    }
    st.dataframe(
        simulacion.style.format("${:,.2f}").apply(
            lambda fila: [
                "font-weight: bold; background-color: #fff2cc"
                if fila.name in filas_destacadas else ""
                for _ in fila
            ],
            axis=1
        ),
        use_container_width=True
    )

    st.markdown("### F. Resultados")
    destacado = resumen["simulacion"][72]
    col_venta, col_descuento, col_adeudo = st.columns(3)
    col_venta.metric(
        "VENTA POSIBLE (72)",
        formato_moneda(destacado["VENTA POSIBLE"])
    )
    col_descuento.metric(
        "DESCUENTO NUEVO",
        formato_moneda(destacado["DESCUENTO NUEVO"])
    )
    col_adeudo.metric(
        "TOTAL ADEUDO CLIENTE (72)",
        formato_moneda(destacado["TOTAL ADEUDO CLIENTE"])
    )
    col_incluidas, col_excluidas = st.columns(2)
    col_incluidas.metric(
        "FACTURAS INCLUIDAS",
        int(facturas["INCLUIR"].sum())
    )
    col_excluidas.metric(
        "FACTURAS EXCLUIDAS",
        int((~facturas["INCLUIR"]).sum())
    )
    st.caption(
        f"Total pagado: {formato_moneda(resumen['total_pagado'])} | "
        f"Saldo pendiente: {formato_moneda(resumen['total_saldo_pendiente'])}"
    )

    st.markdown("### G. Exportación")
    datos_cliente = {
        "fecha": fecha,
        "semana": int(semana),
        "vendedor": vendedor,
        "cliente": cliente_nombre,
        "rfc_nac": rfc,
        "fecha_nacimiento": nacimiento["fecha_nacimiento"],
        "edad": nacimiento["edad"],
        "quinquenio": quinquenio
    }
    excel = generar_excel_refinanciamiento(facturas, resumen, datos_cliente)
    pdf = generar_pdf_refinanciamiento(facturas, resumen, datos_cliente)
    json_data = generar_json_refinanciamiento(facturas, resumen, datos_cliente)

    col_excel, col_pdf, col_json = st.columns(3)
    col_excel.download_button(
        "Descargar Excel",
        excel,
        nombre_archivo_refinanciamiento(extension="xlsx"),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    col_pdf.download_button(
        "Descargar PDF",
        pdf,
        nombre_archivo_refinanciamiento(extension="pdf"),
        "application/pdf",
        use_container_width=True
    )
    col_json.download_button(
        "Descargar JSON",
        json_data,
        nombre_archivo_refinanciamiento(extension="json"),
        "application/json",
        use_container_width=True
    )
    if st.button(
        "Guardar Excel, PDF y JSON en carpeta",
        type="primary",
        key="ref_guardar_carpeta",
        use_container_width=True
    ):
        try:
            rutas = guardar_archivos_refinanciamiento(
                {"xlsx": excel, "pdf": pdf, "json": json_data},
                OUTPUT_REFINANCIAMIENTO_DIR,
                int(semana),
                vendedor,
                cliente_nombre
            )
            st.success(f"Archivos guardados en: {rutas['xlsx'].parent}")
        except Exception as error:
            st.error(f"No se pudieron guardar los archivos: {error}")
