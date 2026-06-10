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
    termino = str(busqueda or "").strip()
    if len(termino) < 2:
        return pd.DataFrame(
            columns=["cliente_id", "cliente", "rfc", "saldo_total"]
        )

    consulta = text(
        """
        SELECT
            c.id AS cliente_id,
            c.name AS cliente,
            COALESCE(c.rfc, '') AS rfc,
            COALESCE(SUM(i.balance), 0) AS saldo_total
        FROM customers_customer c
        JOIN sales_sale s ON s.customer_id = c.id
        JOIN payments_installment i ON i.sale_id = s.id
        WHERE (
            c.name ILIKE :busqueda
            OR COALESCE(c.rfc, '') ILIKE :busqueda
        )
        GROUP BY c.id, c.name, c.rfc
        HAVING COALESCE(SUM(i.balance), 0) > 0
        ORDER BY c.name, c.rfc
        LIMIT :limite
        """
    )

    try:
        with engine.connect() as conexion:
            return pd.read_sql(
                consulta,
                conexion,
                params={
                    "busqueda": f"%{termino}%",
                    "limite": int(limite)
                }
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
        SELECT
            s.id AS venta_id,
            s.folio AS fact,
            c.name AS cliente,
            COALESCE(c.rfc, '') AS rfc,
            COALESCE(s.total, 0) AS vta,
            COALESCE(SUM(i.balance), 0) AS saldo
        FROM sales_sale s
        JOIN customers_customer c ON c.id = s.customer_id
        JOIN payments_installment i ON i.sale_id = s.id
        WHERE c.id = :cliente_id
        GROUP BY s.id, s.folio, c.name, c.rfc, s.total
        HAVING COALESCE(SUM(i.balance), 0) > 0
        ORDER BY s.folio
        """
    )

    try:
        with engine.connect() as conexion:
            return pd.read_sql(
                consulta,
                conexion,
                params={"cliente_id": cliente_id}
            )
    except Exception as error:
        raise RefinanciamientoDatabaseError(
            f"No se pudieron cargar las facturas del cliente: {error}"
        ) from error
