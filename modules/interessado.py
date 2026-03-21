"""
hermes/interessado.py
Cliente para os endpoints de Interessado do SUPP.

Endpoints cobertos:
  GET    /v1/administrativo/interessado             Lista paginada com filtros
  GET    /v1/administrativo/interessado/count       Contagem com filtros
  GET    /v1/administrativo/interessado/{id}        Busca por ID

Associações disponíveis no populate:
  processo, modalidadeInteressado, pessoa, criadoPor, atualizadoPor, apagadoPor
"""

from __future__ import annotations

import json
from typing import Any

import httpx

from .config import BASE_URL

_BASE_PATH = "/v1/administrativo/interessado"

_POPULATE_PADRAO = ["modalidadeInteressado", "pessoa"]


class InteressadoError(Exception):
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
        raise InteressadoError(response.status_code, body)
    try:
        return response.json()
    except Exception:
        return response.text


def _where_str(where: dict | str | None) -> str | None:
    if where is None:
        return None
    return json.dumps(where, ensure_ascii=False) if isinstance(where, dict) else where


def _populate_str(populate: list[str] | str | None) -> str | None:
    if populate is None:
        return None
    return json.dumps(populate) if isinstance(populate, list) else populate


def _extract_list(data: Any) -> list[dict]:
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("entities", "data", "results", "items"):
            if key in data and isinstance(data[key], list):
                return data[key]
    return []


class InteressadoClient:
    """
    Cliente síncrono para os endpoints de Interessado do SUPP.

    Uso:
        ic = InteressadoClient.from_auth(auth)
        interessados = ic.listar_por_processo(processo_id=42)
    """

    def __init__(
        self,
        token: str,
        base_url: str = BASE_URL,
        timeout: float = 120.0,
    ) -> None:
        self.token = token
        self.base_url = base_url
        self._http = httpx.Client(
            base_url=base_url,
            timeout=timeout,
            headers={"Authorization": f"Bearer {token}"},
        )

    @classmethod
    def from_auth(cls, auth_client: Any, timeout: float = 120.0) -> "InteressadoClient":
        """Cria InteressadoClient a partir de um AuthClient já autenticado."""
        if not auth_client.token:
            raise RuntimeError("AuthClient sem token. Faça login primeiro.")
        return cls(token=auth_client.token, base_url=auth_client.base_url, timeout=timeout)

    def buscar(
        self,
        interessado_id: int | str,
        populate: list[str] | str | None = None,
    ) -> dict:
        """GET /interessado/{id} — Busca interessado por ID."""
        params: dict[str, Any] = {}
        if populate is not None:
            params["populate"] = _populate_str(populate)
        resp = self._http.get(f"{_BASE_PATH}/{interessado_id}", params=params)
        return _check(resp)

    def listar(
        self,
        where: dict | str | None = None,
        order: dict | str | None = None,
        limit: int = 50,
        offset: int = 0,
        populate: list[str] | str | None = None,
        context: str | None = None,
    ) -> list[dict]:
        """GET /interessado — Lista interessados com paginação e filtros."""
        params: dict[str, Any] = {"limit": limit, "offset": offset}
        if where is not None:
            params["where"] = _where_str(where)
        if order is not None:
            params["order"] = json.dumps(order) if isinstance(order, dict) else order
        if populate is not None:
            params["populate"] = _populate_str(populate)
        if context is not None:
            params["context"] = context
        resp = self._http.get(_BASE_PATH, params=params)
        return _extract_list(_check(resp))

    def contar(self, where: dict | str | None = None) -> int:
        """GET /interessado/count — Contagem com filtro."""
        params: dict[str, Any] = {}
        if where is not None:
            params["where"] = _where_str(where)
        data = _check(self._http.get(f"{_BASE_PATH}/count", params=params))
        if isinstance(data, dict):
            return int(data.get("count", data.get("total", 0)))
        return int(data)

    def listar_por_processo(
        self,
        processo_id: int | str,
        populate: list[str] | str | None = None,
        limit: int = 100,
    ) -> list[dict]:
        """
        Lista todos os interessados de um processo.

        GET /interessado?where={"processo.id":"eq:{id}"}&populate=["modalidadeInteressado","pessoa"]
        """
        where = {"processo.id": f"eq:{processo_id}"}
        pop = populate if populate is not None else _POPULATE_PADRAO
        return self.listar(where=where, populate=pop, limit=limit)

    def __enter__(self) -> "InteressadoClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()
