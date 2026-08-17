from pathlib import Path
import getpass
import json
import os
import shutil
import urllib.error
import urllib.request
from datetime import datetime
from urllib.parse import urlsplit
from uuid import uuid4

import pandas as pd


ENDPOINT_URL = "https://api-refeicao.jfi.com.br/users/licenses"
PASTA_SCRIPT = Path(__file__).resolve().parent
PASTA_PROJETO = PASTA_SCRIPT.parent
OUTPUT_PATH = PASTA_PROJETO / "04 - SAIDAS" / "BASE_SIGO.xlsx"
HISTORICO_PATH = PASTA_PROJETO / "04 - SAIDAS" / "HISTORICO_SIGO"
COLUNAS_OBRIGATORIAS = {
    "externalid", "cpf", "nome", "costcenter", "costcentercode", "isactive",
}


def request_json(url, method="GET", data=None, token=None):
    headers = {"Accept": "application/json"}
    body = None

    if data is not None:
        headers["Content-Type"] = "application/json"
        body = json.dumps(data).encode("utf-8")
    if token:
        headers["Authorization"] = f"Bearer {token}"

    request = urllib.request.Request(url, data=body, headers=headers, method=method)

    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as error:
        detail = error.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"Erro HTTP {error.code}: {detail}") from error
    except urllib.error.URLError as error:
        raise RuntimeError(f"Nao foi possivel conectar ao SIGO: {error.reason}") from error


def main():
    document = os.getenv("SIGO_DOCUMENT", "").strip()
    password = os.getenv("SIGO_PASSWORD", "")
    if not document:
        document = input("Documento de acesso ao SIGO: ").strip()
    if not password:
        password = getpass.getpass("Senha do SIGO: ")
    if not document or not password:
        raise SystemExit("Documento e senha são obrigatórios.")

    endpoint = urlsplit(ENDPOINT_URL)
    auth_url = f"{endpoint.scheme}://{endpoint.netloc}/auth/login"
    login = request_json(
        auth_url,
        method="POST",
        data={"document": document, "password": password},
    )

    token = login.get("accessToken")
    if not token:
        raise RuntimeError("A autenticacao nao retornou accessToken.")

    response = request_json(ENDPOINT_URL, token=token)
    records = response.get("data", []) if isinstance(response, dict) else response
    if not isinstance(records, list):
        raise RuntimeError("O endpoint nao retornou uma lista no campo data.")

    table = pd.json_normalize(records)

    if table.empty:
        raise RuntimeError("O SIGO retornou uma base vazia; a base local não foi alterada.")
    ausentes = COLUNAS_OBRIGATORIAS - set(table.columns)
    if ausentes:
        raise RuntimeError(
            "A resposta do SIGO não contém as colunas obrigatórias: "
            + ", ".join(sorted(ausentes))
        )

    # Mantem todas as colunas retornadas e garante a nova coluna na exportacao.
    if "costcentercode" not in table.columns:
        table["costcentercode"] = pd.NA

    output = Path(OUTPUT_PATH).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    temporario = output.with_name(f".{output.stem}.{uuid4().hex}.tmp.xlsx")
    try:
        table.to_excel(temporario, index=False)
        if output.exists():
            HISTORICO_PATH.mkdir(parents=True, exist_ok=True)
            momento = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            snapshot = HISTORICO_PATH / f"{output.stem}_{momento}{output.suffix}"
            shutil.copy2(output, snapshot)
        temporario.replace(output)
    finally:
        if temporario.exists():
            temporario.unlink()

    print(f"Arquivo gerado: {output}")
    print(f"Registros exportados: {len(table)}")


if __name__ == "__main__":
    main()
