import os
from pathlib import Path
from urllib.parse import quote

import msal
import requests
from dotenv import load_dotenv


load_dotenv()

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


class GraphStorageError(Exception):
    pass


def get_env_required(name: str) -> str:
    value = os.getenv(name)

    if not value:
        raise GraphStorageError(
            f"Falta configurar la variable {name} en el archivo .env"
        )

    return value


def get_access_token() -> str:
    tenant_id = get_env_required("MS_TENANT_ID")
    client_id = get_env_required("MS_CLIENT_ID")
    client_secret = get_env_required("MS_CLIENT_SECRET")

    authority = f"https://login.microsoftonline.com/{tenant_id}"

    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        authority=authority,
        client_credential=client_secret
    )

    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    if "access_token" not in result:
        raise GraphStorageError(
            f"No se pudo obtener token de Graph: {result.get('error_description', result)}"
        )

    return result["access_token"]


def graph_headers() -> dict:
    token = get_access_token()

    return {
        "Authorization": f"Bearer {token}"
    }


def graph_headers_json() -> dict:
    headers = graph_headers()
    headers["Content-Type"] = "application/json"
    return headers


def get_site_id() -> str:
    site_hostname = get_env_required("MS_SITE_HOSTNAME")
    site_path = get_env_required("MS_SITE_PATH")

    url = f"{GRAPH_BASE_URL}/sites/{site_hostname}:{site_path}"

    response = requests.get(url, headers=graph_headers(), timeout=30)

    if response.status_code >= 400:
        raise GraphStorageError(
            f"Error obteniendo site_id: {response.status_code} {response.text}"
        )

    return response.json()["id"]


def get_drive_id(site_id: str) -> str:
    drive_name = os.getenv("MS_DRIVE_NAME", "").strip()

    url = f"{GRAPH_BASE_URL}/sites/{site_id}/drives"

    response = requests.get(url, headers=graph_headers(), timeout=30)

    if response.status_code >= 400:
        raise GraphStorageError(
            f"Error obteniendo drives: {response.status_code} {response.text}"
        )

    drives = response.json().get("value", [])

    if not drives:
        raise GraphStorageError("No se encontraron drives en el sitio.")

    if drive_name:
        for drive in drives:
            if drive.get("name", "").lower() == drive_name.lower():
                return drive["id"]

        disponibles = [drive.get("name", "") for drive in drives]
        raise GraphStorageError(
            f"No se encontró el drive '{drive_name}'. Disponibles: {disponibles}"
        )

    return drives[0]["id"]


def sanitize_remote_name(text: str) -> str:
    invalid_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']

    text = str(text).strip()

    for ch in invalid_chars:
        text = text.replace(ch, "")

    text = " ".join(text.split())

    return text


def build_remote_folder_path(
    anio: int,
    semana: int,
    promotor: str,
    nombre_cliente: str
) -> str:
    root_folder = get_env_required("MS_ROOT_FOLDER").strip("/")

    promotor_clean = sanitize_remote_name(promotor)
    cliente_clean = sanitize_remote_name(nombre_cliente)

    return f"{root_folder}/{anio}/{int(semana):02d}/{promotor_clean}/{cliente_clean}"


def create_folder_if_not_exists(
    site_id: str,
    drive_id: str,
    folder_path: str
):
    parts = [p for p in folder_path.strip("/").split("/") if p]
    current_path = ""

    for part in parts:
        parent_path = current_path
        current_path = f"{current_path}/{part}" if current_path else part

        encoded_current = quote(current_path)
        check_url = (
            f"{GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}"
            f"/root:/{encoded_current}"
        )

        check_response = requests.get(
            check_url,
            headers=graph_headers(),
            timeout=30
        )

        if check_response.status_code == 200:
            continue

        if check_response.status_code != 404:
            raise GraphStorageError(
                f"Error verificando carpeta '{current_path}': "
                f"{check_response.status_code} {check_response.text}"
            )

        if parent_path:
            encoded_parent = quote(parent_path)
            create_url = (
                f"{GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}"
                f"/root:/{encoded_parent}:/children"
            )
        else:
            create_url = (
                f"{GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}"
                f"/root/children"
            )

        payload = {
            "name": part,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "replace"
        }

        create_response = requests.post(
            create_url,
            headers=graph_headers_json(),
            json=payload,
            timeout=30
        )

        if create_response.status_code >= 400:
            raise GraphStorageError(
                f"Error creando carpeta '{current_path}': "
                f"{create_response.status_code} {create_response.text}"
            )


def upload_small_file(
    site_id: str,
    drive_id: str,
    remote_folder_path: str,
    local_file_path: str,
    remote_file_name: str
) -> dict:
    local_path = Path(local_file_path)

    if not local_path.exists():
        raise GraphStorageError(f"No existe el archivo local: {local_file_path}")

    remote_file_name = sanitize_remote_name(remote_file_name)
    remote_path = f"{remote_folder_path.strip('/')}/{remote_file_name}"
    encoded_remote_path = quote(remote_path)

    url = (
        f"{GRAPH_BASE_URL}/sites/{site_id}/drives/{drive_id}"
        f"/root:/{encoded_remote_path}:/content"
    )

    with open(local_path, "rb") as file:
        response = requests.put(
            url,
            headers=graph_headers(),
            data=file,
            timeout=120
        )

    if response.status_code >= 400:
        raise GraphStorageError(
            f"Error subiendo archivo '{remote_file_name}': "
            f"{response.status_code} {response.text}"
        )

    return response.json()


def subir_revision_a_graph(
    ruta_pdf: str,
    ruta_excel: str,
    anio: int,
    semana: int,
    promotor: str,
    nombre_cliente: str,
    rfc: str
) -> dict:
    site_id = get_site_id()
    drive_id = get_drive_id(site_id)

    remote_folder_path = build_remote_folder_path(
        anio=anio,
        semana=semana,
        promotor=promotor,
        nombre_cliente=nombre_cliente
    )

    create_folder_if_not_exists(
        site_id=site_id,
        drive_id=drive_id,
        folder_path=remote_folder_path
    )

    cliente_file = sanitize_remote_name(nombre_cliente).replace(" ", "_")
    promotor_file = sanitize_remote_name(promotor).replace(" ", "_")

    pdf_name = f"TALON_{cliente_file}_{rfc}.pdf"
    excel_name = f"REVISION_{cliente_file}_{promotor_file}.xlsx"

    pdf_result = upload_small_file(
        site_id=site_id,
        drive_id=drive_id,
        remote_folder_path=remote_folder_path,
        local_file_path=ruta_pdf,
        remote_file_name=pdf_name
    )

    excel_result = upload_small_file(
        site_id=site_id,
        drive_id=drive_id,
        remote_folder_path=remote_folder_path,
        local_file_path=ruta_excel,
        remote_file_name=excel_name
    )

    return {
        "remote_folder_path": remote_folder_path,
        "pdf_web_url": pdf_result.get("webUrl"),
        "excel_web_url": excel_result.get("webUrl")
    }
