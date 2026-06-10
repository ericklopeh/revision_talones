import os

import pandas as pd
from dotenv import load_dotenv
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine
from sqlalchemy.engine.url import make_url
from sshtunnel import SSHTunnelForwarder


load_dotenv()


class RefinanciamientoDatabaseError(RuntimeError):
    pass


def obtener_database_url(secrets=None) -> str:
    if secrets is not None:
        for clave in [
            "SSH_HOST",
            "SSH_PORT",
            "SSH_USERNAME",
            "SSH_PASSWORD",
            "DB_HOST",
            "DB_PORT"
        ]:
            try:
                valor = secrets.get(clave, "")
            except (AttributeError, KeyError, TypeError):
                valor = ""
            if valor:
                os.environ[clave] = str(valor)

        try:
            url = secrets.get("DATABASE_URL", "")
        except (AttributeError, KeyError, TypeError):
            url = ""
        if url:
            return str(url).strip()

    return os.getenv("DATABASE_URL", "").strip()


def crear_engine_refinanciamiento(database_url: str) -> Engine:
    if not database_url:
        raise RefinanciamientoDatabaseError(
            "Falta configurar DATABASE_URL en los secretos de Streamlit "
            "o en las variables de entorno."
        )

    try:
        ssh_host = os.getenv("SSH_HOST", "").strip()
        tunnel = None
        url_engine = database_url

        if ssh_host:
            tunnel = SSHTunnelForwarder(
                (ssh_host, int(os.getenv("SSH_PORT", "22"))),
                ssh_username=os.getenv("SSH_USERNAME", "").strip(),
                ssh_password=os.getenv("SSH_PASSWORD", ""),
                remote_bind_address=(
                    os.getenv("DB_HOST", "localhost").strip(),
                    int(os.getenv("DB_PORT", "5432"))
                ),
                local_bind_address=("127.0.0.1", 0)
            )
            tunnel.start()
            url = make_url(database_url)
            url_engine = url.set(
                host="127.0.0.1",
                port=tunnel.local_bind_port
            )

        engine = create_engine(
            url_engine,
            pool_pre_ping=True,
            pool_recycle=300
        )
        if tunnel is not None:
            engine._refinanciamiento_ssh_tunnel = tunnel
        return engine
    except Exception as error:
        if tunnel is not None and tunnel.is_active:
            tunnel.stop()
        raise RefinanciamientoDatabaseError(
            f"No se pudo preparar la conexión PostgreSQL: {error}"
        ) from error


def buscar_clientes(
    engine: Engine,
    busqueda: str,
    limite: int = 50
) -> pd.DataFrame:
    tokens = [
        token for token in str(busqueda or "").strip().split()
        if len(token) >= 2
    ]
    if not tokens:
        return pd.DataFrame(
            columns=[
                "cliente_id",
                "cliente",
                "rfc",
                "facturas_encontradas",
                "saldo_total"
            ]
        )

    condiciones = []
    puntuaciones = []
    parametros = {"limite": int(limite)}
    for indice, token in enumerate(tokens):
        parametro = f"token_{indice}"
        condiciones.append(
            f"""(
                COALESCE(c.first_name, '') ILIKE :{parametro}
                OR COALESCE(c.last_name, '') ILIKE :{parametro}
                OR COALESCE(c.rfc, '') ILIKE :{parametro}
            )"""
        )
        puntuaciones.append(
            f"""CASE WHEN (
                COALESCE(c.first_name, '') ILIKE :{parametro}
                OR COALESCE(c.last_name, '') ILIKE :{parametro}
                OR COALESCE(c.rfc, '') ILIKE :{parametro}
            ) THEN 1 ELSE 0 END"""
        )
        parametros[parametro] = f"%{token}%"

    consulta = text(
        f"""
        WITH importes_articulos AS (
            SELECT
                si.sale_id,
                COALESCE(SUM(si.real), 0) AS vta_real
            FROM sales_saleitem si
            WHERE NOT COALESCE(si.is_removed, false)
            GROUP BY si.sale_id
        ),
        pagos_por_venta AS (
            SELECT
                p.sale_id,
                COALESCE(SUM(i.amount), 0) AS pagado,
                MODE() WITHIN GROUP (ORDER BY i.amount)
                    FILTER (WHERE i.amount > 0) AS abono_periodico
            FROM payments_payment p
            JOIN payments_installment i ON i.payment_id = p.id
            WHERE NOT COALESCE(p.is_removed, false)
              AND NOT COALESCE(i.is_removed, false)
            GROUP BY p.sale_id
        ),
        facturas AS (
            SELECT
                s.id,
                s.customer_id,
                CASE
                    WHEN ppv.abono_periodico IS NULL
                        OR COALESCE(s.time_limit, 0) <= 0
                        THEN COALESCE(ia.vta_real, 0)
                    WHEN ABS(
                        COALESCE(ia.vta_real, 0)
                        - (ppv.abono_periodico * s.time_limit)
                    ) <= 1
                        THEN COALESCE(ia.vta_real, 0)
                    ELSE ppv.abono_periodico * s.time_limit
                END AS vta,
                COALESCE(ppv.pagado, 0) AS pagado,
                CASE
                    WHEN ppv.abono_periodico IS NULL
                        OR COALESCE(s.time_limit, 0) <= 0
                        THEN COALESCE(ia.vta_real, 0)
                    WHEN ABS(
                        COALESCE(ia.vta_real, 0)
                        - (ppv.abono_periodico * s.time_limit)
                    ) <= 1
                        THEN COALESCE(ia.vta_real, 0)
                    ELSE ppv.abono_periodico * s.time_limit
                END - COALESCE(ppv.pagado, 0) AS saldo
            FROM sales_sale s
            LEFT JOIN importes_articulos ia ON ia.sale_id = s.id
            LEFT JOIN pagos_por_venta ppv ON ppv.sale_id = s.id
            WHERE NOT COALESCE(s.is_removed, false)
        )
        SELECT
            c.id AS cliente_id,
            TRIM(CONCAT_WS(
                ' ',
                NULLIF(c.first_name, ''),
                NULLIF(c.last_name, '')
            )) AS cliente,
            COALESCE(c.rfc, '') AS rfc,
            COUNT(DISTINCT f.id) AS facturas_encontradas,
            COALESCE(SUM(f.saldo), 0) AS saldo_total
        FROM customers_customer c
        JOIN facturas f ON f.customer_id = c.id
        WHERE {" OR ".join(condiciones)}
          AND f.saldo > 0
        GROUP BY c.id, c.first_name, c.last_name, c.rfc
        HAVING COALESCE(SUM(f.saldo), 0) > 0
        ORDER BY ({" + ".join(puntuaciones)}) DESC,
                 c.first_name,
                 c.last_name,
                 c.rfc
        LIMIT :limite
        """
    )

    try:
        with engine.connect() as conexion:
            return pd.read_sql(
                consulta,
                conexion,
                params=parametros
            )
    except Exception as error:
        raise RefinanciamientoDatabaseError(
            f"No se pudieron buscar clientes: {error}"
        ) from error


def cargar_facturas_cliente(
    engine: Engine,
    cliente_id
) -> pd.DataFrame:
    consulta = text(
        """
        WITH importes_articulos AS (
            SELECT
                si.sale_id,
                COALESCE(SUM(si.real), 0) AS vta_real
            FROM sales_saleitem si
            WHERE NOT COALESCE(si.is_removed, false)
            GROUP BY si.sale_id
        ),
        pagos AS (
            SELECT
                p.sale_id,
                COALESCE(SUM(i.amount), 0) AS pagado,
                MODE() WITHIN GROUP (ORDER BY i.amount)
                    FILTER (WHERE i.amount > 0) AS abono_periodico
            FROM payments_payment p
            JOIN payments_installment i ON i.payment_id = p.id
            WHERE NOT COALESCE(p.is_removed, false)
              AND NOT COALESCE(i.is_removed, false)
            GROUP BY p.sale_id
        ),
        facturas AS (
            SELECT
                s.id,
                s.folio,
                s.customer_id,
                s.time_limit AS plazo_venta,
                COALESCE(ia.vta_real, 0) AS vta_real,
                COALESCE(pa.pagado, 0) AS pagado,
                pa.abono_periodico,
                CASE
                    WHEN pa.abono_periodico IS NULL
                        OR COALESCE(s.time_limit, 0) <= 0
                        THEN COALESCE(ia.vta_real, 0)
                    WHEN ABS(
                        COALESCE(ia.vta_real, 0)
                        - (pa.abono_periodico * s.time_limit)
                    ) <= 1
                        THEN COALESCE(ia.vta_real, 0)
                    ELSE pa.abono_periodico * s.time_limit
                END AS vta
            FROM sales_sale s
            LEFT JOIN importes_articulos ia ON ia.sale_id = s.id
            LEFT JOIN pagos pa ON pa.sale_id = s.id
            WHERE NOT COALESCE(s.is_removed, false)
        )
        SELECT
            f.id AS venta_id,
            f.folio AS fact,
            c.id AS cliente_id,
            TRIM(CONCAT_WS(
                ' ',
                NULLIF(c.first_name, ''),
                NULLIF(c.last_name, '')
            )) AS cliente,
            COALESCE(c.rfc, '') AS rfc,
            f.plazo_venta,
            f.vta_real,
            f.abono_periodico,
            f.vta,
            f.pagado AS pagado_db,
            f.vta - f.pagado AS saldo
        FROM facturas f
        JOIN customers_customer c ON c.id = f.customer_id
        WHERE c.id = :cliente_id
          AND f.vta - f.pagado > 0
        ORDER BY f.folio
        """
    )

    try:
        with engine.connect() as conexion:
            return pd.read_sql(
                consulta,
                conexion,
                params={"cliente_id": int(cliente_id)}
            )
    except Exception as error:
        raise RefinanciamientoDatabaseError(
            f"No se pudieron cargar las facturas del cliente: {error}"
        ) from error


def diagnosticar_vta_facturas(
    engine: Engine,
    folios: list
) -> pd.DataFrame:
    folios_limpios = [int(folio) for folio in folios]
    consulta = text(
        """
        SELECT
            s.id AS sale_id,
            s.folio,
            s.time_limit AS plazo_venta,
            si.id AS saleitem_id,
            si.quantity,
            si.real,
            si.price_id,
            si.real * COALESCE(si.quantity, 1) AS real_por_cantidad
        FROM sales_sale s
        JOIN sales_saleitem si ON si.sale_id = s.id
        WHERE s.folio = ANY(:folios)
          AND NOT COALESCE(s.is_removed, false)
          AND NOT COALESCE(si.is_removed, false)
        ORDER BY s.folio, si.id
        """
    )
    with engine.connect() as conexion:
        return pd.read_sql(
            consulta,
            conexion,
            params={"folios": folios_limpios}
        )
