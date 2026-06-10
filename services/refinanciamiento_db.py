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
        parametros[parametro] = f"%{token}%"

    consulta = text(
        f"""
        WITH saldos_por_venta AS (
            SELECT
                p.sale_id,
                COALESCE(SUM(i.amount), 0) AS saldo
            FROM payments_payment p
            JOIN payments_installment i ON i.payment_id = p.id
            GROUP BY p.sale_id
        )
        SELECT
            c.id AS cliente_id,
            TRIM(CONCAT_WS(
                ' ',
                NULLIF(c.first_name, ''),
                NULLIF(c.last_name, '')
            )) AS cliente,
            COALESCE(c.rfc, '') AS rfc,
            COUNT(DISTINCT s.id) AS facturas_encontradas,
            COALESCE(SUM(spv.saldo), 0) AS saldo_total
        FROM customers_customer c
        JOIN sales_sale s ON s.customer_id = c.id
        JOIN saldos_por_venta spv ON spv.sale_id = s.id
        WHERE {" OR ".join(condiciones)}
        GROUP BY c.id, c.first_name, c.last_name, c.rfc
        HAVING COALESCE(SUM(spv.saldo), 0) > 0
        ORDER BY c.first_name, c.last_name, c.rfc
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
        WITH ventas AS (
            SELECT
                si.sale_id,
                COALESCE(SUM(si.real), 0) AS vta
            FROM sales_saleitem si
            GROUP BY si.sale_id
        ),
        saldos AS (
            SELECT
                p.sale_id,
                COALESCE(SUM(i.amount), 0) AS saldo
            FROM payments_payment p
            JOIN payments_installment i ON i.payment_id = p.id
            GROUP BY p.sale_id
        )
        SELECT
            s.id AS venta_id,
            s.folio AS fact,
            c.id AS cliente_id,
            TRIM(CONCAT_WS(
                ' ',
                NULLIF(c.first_name, ''),
                NULLIF(c.last_name, '')
            )) AS cliente,
            COALESCE(c.rfc, '') AS rfc,
            COALESCE(v.vta, 0) AS vta,
            COALESCE(sa.saldo, 0) AS saldo
        FROM sales_sale s
        JOIN customers_customer c ON c.id = s.customer_id
        LEFT JOIN ventas v ON v.sale_id = s.id
        LEFT JOIN saldos sa ON sa.sale_id = s.id
        WHERE c.id = :cliente_id
          AND COALESCE(sa.saldo, 0) > 0
        ORDER BY s.folio
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
