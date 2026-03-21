"""
hermes/tipo_documento.py
Cliente para os endpoints de TipoDocumento do SUPP.

Endpoints cobertos:
  GET  /v1/administrativo/tipo_documento          Lista paginada com filtros
  GET  /v1/administrativo/tipo_documento/count    Contagem com filtros
  GET  /v1/administrativo/tipo_documento/{id}     Busca por ID
"""

from __future__ import annotations

import json
from typing import Any

import httpx

from .config import BASE_URL

_BASE_PATH = "/v1/administrativo/tipo_documento"


class TipoDocumentoError(Exception):
    def __init__(self, status_code: int, body: object) -> None:
        self.status_code = status_code
        self.body = body
        super().__init__(f"HTTP {status_code}: {body}")


def _check(response: httpx.Response) -> Any:
    if not response.is_success:
        try:
            body = response.json()
        except Exception:
            body = response.text
        raise TipoDocumentoError(response.status_code, body)
    try:
        return response.json()
    except Exception:
        return response.text


def _where_str(where: dict | str | None) -> str | None:
    if where is None:
        return None
    return json.dumps(where, ensure_ascii=False) if isinstance(where, dict) else where


def _extract_list(data: Any) -> list[dict]:
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("entities", "data", "results", "items"):
            if key in data and isinstance(data[key], list):
                return data[key]
    return []


class TipoDocumentoClient:
    """
    Cliente síncrono para TipoDocumento do SUPP.

    Uso:
        from hermes.auth import AuthClient
        from hermes.tipo_documento import TipoDocumentoClient

        auth = AuthClient()
        auth.login_ldap("cpf", "senha")
        tdc = TipoDocumentoClient.from_auth(auth)

        tipo = tdc.buscar_por_nome("sentença")
    """

    def __init__(
        self,
        token: str,
        base_url: str = BASE_URL,
        timeout: float = 60.0,
    ) -> None:
        self.token = token
        self.base_url = base_url
        self._http = httpx.Client(
            base_url=base_url,
            timeout=timeout,
            headers={"Authorization": f"Bearer {token}"},
        )

    @classmethod
    def from_auth(cls, auth_client: Any, timeout: float = 60.0) -> "TipoDocumentoClient":
        if not auth_client.token:
            raise RuntimeError("AuthClient sem token. Faça login primeiro.")
        return cls(token=auth_client.token, base_url=auth_client.base_url, timeout=timeout)

    def buscar(self, tipo_id: int | str) -> dict:
        """GET /tipo_documento/{id}"""
        resp = self._http.get(f"{_BASE_PATH}/{tipo_id}")
        return _check(resp)

    def listar(
        self,
        where: dict | str | None = None,
        limit: int = 25,
        offset: int = 0,
    ) -> list[dict]:
        """GET /tipo_documento — Lista paginada com filtros."""
        params: dict[str, Any] = {"limit": limit, "offset": offset}
        if where is not None:
            params["where"] = _where_str(where)
        resp = self._http.get(_BASE_PATH, params=params)
        return _extract_list(_check(resp))

    def contar(self, where: dict | str | None = None) -> int:
        """GET /tipo_documento/count"""
        params: dict[str, Any] = {}
        if where is not None:
            params["where"] = _where_str(where)
        data = _check(self._http.get(f"{_BASE_PATH}/count", params=params))
        if isinstance(data, dict):
            return int(data.get("count", data.get("total", 0)))
        return int(data)

    def buscar_por_nome(self, nome: str) -> list[dict]:
        """
        Busca tipos de documento pelo nome (case-insensitive, parcial).

        Equivale a GET /tipo_documento?where={"nome":"like:%{nome}%"}

        Retorna lista com todos os matches.
        Lança TipoDocumentoError se nenhum tipo for encontrado.
        """
        where = {"nome": f"like:%{nome}%"}
        resultados = self.listar(where=where, limit=50)
        if not resultados:
            raise TipoDocumentoError(404, f"Nenhum TipoDocumento encontrado para nome like '%{nome}%'")
        return resultados

    def buscar_por_sigla(self, sigla: str) -> dict:
        """
        Busca um tipo de documento pela sigla exata (case-sensitive).

        Equivale a GET /tipo_documento?where={"sigla":"eq:{sigla}"}

        Lança TipoDocumentoError se não encontrado.
        """
        where = {"sigla": f"eq:{sigla}"}
        resultados = self.listar(where=where, limit=5)
        if not resultados:
            raise TipoDocumentoError(404, f"Nenhum TipoDocumento encontrado para sigla '{sigla}'")
        return resultados[0]

    def __enter__(self) -> "TipoDocumentoClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()
