import unittest
from datetime import date
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook

from utils.refinanciamiento import (
    COLUMNAS_MANUALES,
    calcular_facturas_refinanciamiento,
    calcular_resumen_refinanciamiento,
    dataframe_facturas_vacio,
    extraer_fecha_edad_desde_rfc,
    generar_excel_refinanciamiento,
    to_number
)


class RefinanciamientoTest(unittest.TestCase):
    def setUp(self):
        self.facturas = pd.DataFrame([
            {
                "FACT": "12546",
                "VTA": 58564.00,
                "SALDO": 16446.80,
                "QNAS TOMADAS A CUENTA": 0,
                "ABONO": 813.39,
                "EN COBRO": ""
            },
            {
                "FACT": "12547",
                "VTA": 28693.43,
                "SALDO": 8368.91,
                "QNAS TOMADAS A CUENTA": 0,
                "ABONO": 398.52,
                "EN COBRO": ""
            },
            {
                "FACT": "12622",
                "VTA": 29966.96,
                "SALDO": 9988.88,
                "QNAS TOMADAS A CUENTA": 0,
                "ABONO": 416.21,
                "EN COBRO": ""
            },
            {
                "FACT": "12801",
                "VTA": 22251.39,
                "SALDO": 9580.34,
                "QNAS TOMADAS A CUENTA": 0,
                "ABONO": 309.05,
                "EN COBRO": ""
            },
            {
                "FACT": "12719",
                "VTA": 40994.80,
                "SALDO": 14803.78,
                "QNAS TOMADAS A CUENTA": 0,
                "ABONO": 569.37,
                "EN COBRO": ""
            }
        ])

    def test_extrae_fecha_y_edad_desde_rfc(self):
        resultado = extraer_fecha_edad_desde_rfc(
            "J860901UT6",
            fecha_referencia=date(2026, 6, 9)
        )

        self.assertEqual(resultado["fecha_nacimiento"], date(1986, 9, 1))
        self.assertEqual(resultado["edad"], 39)

    def test_rfc_sin_fecha_valida_no_calcula_edad(self):
        resultado = extraer_fecha_edad_desde_rfc(
            "ABC991332XYZ",
            fecha_referencia=date(2026, 6, 9)
        )

        self.assertIsNone(resultado["fecha_nacimiento"])
        self.assertIsNone(resultado["edad"])

    def test_calculos_por_factura_evitan_division_entre_cero(self):
        facturas = calcular_facturas_refinanciamiento(pd.DataFrame([{
            "FACT": "PRUEBA",
            "VTA": 0,
            "SALDO": 500,
            "QNAS TOMADAS A CUENTA": 2,
            "ABONO": 50,
            "EN COBRO": ""
        }]))

        self.assertEqual(facturas.loc[0, "ABONO DE QUINCENAS"], 100)
        self.assertEqual(facturas.loc[0, "PAGADO"], 0)
        self.assertEqual(facturas.loc[0, "SALDO PENDIENTE"], 400)
        self.assertEqual(facturas.loc[0, "PORCENTAJE PAGADO"], 0)
        self.assertEqual(facturas.loc[0, "PUEDE REFINANCIAR"], "NO")

    def test_qnas_vacias_o_invalidas_se_toman_como_cero(self):
        facturas = calcular_facturas_refinanciamiento(pd.DataFrame([
            {
                "FACT": "VACIA",
                "VTA": "$ 1,000.00",
                "SALDO": "600.00",
                "ABONO": "100",
                "EN COBRO": "",
                "QNAS TOMADAS A CUENTA": ""
            },
            {
                "FACT": "TEXTO",
                "VTA": 1000,
                "SALDO": 600,
                "ABONO": 100,
                "EN COBRO": "",
                "QNAS TOMADAS A CUENTA": "dato inválido"
            }
        ]))

        self.assertEqual(facturas.loc[0, "ABONO DE QUINCENAS"], 0)
        self.assertEqual(facturas.loc[1, "ABONO DE QUINCENAS"], 0)
        self.assertEqual(facturas.loc[0, "SALDO PENDIENTE"], 600)
        self.assertEqual(facturas.loc[1, "SALDO PENDIENTE"], 600)

    def test_columnas_manuales_tienen_orden_y_qnas_vacias(self):
        self.assertEqual(
            COLUMNAS_MANUALES,
            [
                "FACT",
                "VTA",
                "SALDO",
                "ABONO",
                "EN COBRO",
                "QNAS TOMADAS A CUENTA"
            ]
        )
        self.assertEqual(
            dataframe_facturas_vacio(1).loc[0, "QNAS TOMADAS A CUENTA"],
            ""
        )

    def test_conversion_numerica_acepta_formatos_comunes(self):
        self.assertEqual(to_number("$ 58,564.00"), 58564)
        self.assertEqual(to_number("72%"), 72)
        self.assertEqual(to_number(""), 0)

    def test_totales_y_simulacion_coinciden_con_ejemplo(self):
        resumen = calcular_resumen_refinanciamiento(self.facturas, 0)

        self.assertEqual(resumen["total_vta"], 180470.58)
        self.assertEqual(resumen["total_pagado"], 121281.87)
        self.assertEqual(resumen["total_saldo"], 59188.71)
        self.assertEqual(resumen["total_abono"], 2506.54)
        self.assertEqual(resumen["abono_ref"], 2506.54)
        self.assertEqual(resumen["total_saldo_pendiente"], 59188.71)
        self.assertEqual(
            resumen["simulacion"][72]["VENTA POSIBLE"],
            180470.88
        )
        self.assertEqual(
            resumen["simulacion"][34]["TOTAL SALDO NUEVO"],
            144411.07
        )

    def test_umbral_de_refinanciamiento_y_aumento(self):
        facturas = pd.DataFrame([
            {
                "FACT": "SI",
                "VTA": 1000,
                "SALDO": 610,
                "QNAS TOMADAS A CUENTA": 0,
                "ABONO": 100,
                "EN COBRO": ""
            },
            {
                "FACT": "NO",
                "VTA": 1000,
                "SALDO": 611,
                "QNAS TOMADAS A CUENTA": 0,
                "ABONO": 200,
                "EN COBRO": ""
            }
        ])
        resultado = calcular_facturas_refinanciamiento(facturas)
        resumen = calcular_resumen_refinanciamiento(resultado, 50)

        self.assertEqual(resultado.loc[0, "PUEDE REFINANCIAR"], "SI")
        self.assertEqual(resultado.loc[1, "PUEDE REFINANCIAR"], "NO")
        self.assertEqual(resumen["abono_ref"], 100)
        self.assertEqual(resumen["abono_antes"], 300)
        self.assertEqual(resumen["total_abono_nuevo"], 150)

    def test_pagado_manual_se_ignora_y_se_recalcula(self):
        resultado = calcular_facturas_refinanciamiento(pd.DataFrame([{
            "FACT": "12546",
            "VTA": 58564,
            "PAGADO": 1,
            "SALDO": 16446.80,
            "QNAS TOMADAS A CUENTA": 0,
            "ABONO": 813.39,
            "EN COBRO": ""
        }]))

        self.assertEqual(resultado.loc[0, "PAGADO"], 42117.20)
        self.assertAlmostEqual(
            resultado.loc[0, "PORCENTAJE PAGADO"],
            0.719165,
            places=6
        )
        self.assertEqual(resultado.loc[0, "REFINANCIAMIENTO"], 42117.20)
        self.assertEqual(resultado.loc[0, "PUEDE REFINANCIAR"], "SI")

    def test_excel_exportado_tiene_dos_hojas(self):
        resumen = calcular_resumen_refinanciamiento(self.facturas, 0)
        contenido = generar_excel_refinanciamiento(
            facturas=self.facturas,
            resumen=resumen,
            datos_cliente={
                "fecha": date(2026, 6, 9),
                "cliente": "CLIENTE PRUEBA",
                "rfc_nac": "J860901UT6",
                "fecha_nacimiento": date(1986, 9, 1),
                "edad": 39,
                "quinquenio": 40
            }
        )
        workbook = load_workbook(BytesIO(contenido), data_only=False)

        self.assertEqual(
            workbook.sheetnames,
            ["Facturas", "Resumen Refinanciamiento"]
        )
        self.assertEqual(workbook["Facturas"]["A2"].value, "12546")
        self.assertEqual(workbook["Facturas"]["C2"].value, 42117.20)
        self.assertEqual(workbook["Facturas"]["J2"].number_format, "0%")
        self.assertEqual(
            workbook["Resumen Refinanciamiento"]["B9"].value,
            2506.54
        )
        self.assertEqual(
            workbook["Resumen Refinanciamiento"]["B22"].value,
            180470.88
        )


if __name__ == "__main__":
    unittest.main()
