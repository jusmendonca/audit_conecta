"""
hermes/setor.py
Cliente para os endpoints de Setor do SUPP.

Endpoints cobertos:
  GET  /v1/administrativo/setor         Lista paginada com filtros
  GET  /v1/administrativo/setor/count   Contagem com filtros
  GET  /v1/administrativo/setor/{id}    Busca por ID
"""

from __future__ import annotations

import json
from typing import Any

import httpx

from .config import BASE_URL

_BASE_PATH = "/v1/administrativo/setor"


class SetorError(Exception):
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
        raise SetorError(response.status_code, body)
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


class SetorClient:
    """
    Cliente síncrono para os endpoints de Setor do SUPP.

    Uso:
        from hermes.auth import AuthClient
        from hermes.setor import SetorClient

        auth = AuthClient()
        auth.login_ldap("usuario", "senha")
        sc = SetorClient.from_auth(auth)

        # Busca setores por sigla ou nome
        setores = sc.buscar_por_nome("EFIN")
    """

    def __init__(
        self,
        token: str,
        base_url: str = BASE_URL,
        timeout: float = 30.0,
    ) -> None:
        self.token = token
        self.base_url = base_url
        self._http = httpx.Client(
            base_url=base_url,
            timeout=timeout,
            headers={"Authorization": f"Bearer {token}"},
        )

    @classmethod
    def from_auth(cls, auth_client: Any, timeout: float = 30.0) -> "SetorClient":
        if not auth_client.token:
            raise RuntimeError("AuthClient sem token. Faça login primeiro.")
        return cls(token=auth_client.token, base_url=auth_client.base_url, timeout=timeout)

    def buscar(self, setor_id: int | str, populate: list[str] | None = None) -> dict:
        """GET /setor/{id} — Busca setor por ID."""
        params: dict[str, Any] = {}
        if populate:
            params["populate"] = json.dumps(populate)
        resp = self._http.get(f"{_BASE_PATH}/{setor_id}", params=params)
        return _check(resp)

    def listar(
        self,
        where: dict | str | None = None,
        limit: int = 25,
        offset: int = 0,
    ) -> list[dict]:
        """GET /setor — Lista setores com filtros."""
        params: dict[str, Any] = {"limit": limit, "offset": offset}
        if where is not None:
            params["where"] = _where_str(where)
        resp = self._http.get(_BASE_PATH, params=params)
        return _extract_list(_check(resp))

    def buscar_por_nome(self, termo: str, limit: int = 20) -> list[dict]:
        """
        Busca setores cujo nome ou sigla contenha o termo informado.

        Tenta primeiro filtro server-side com like; se falhar, filtra client-side.
        Retorna lista ordenada: primeiro por sigla exata, depois por correspondência parcial.
        """
        termo_upper = termo.upper().strip()

        # Tentativa 1: like server-side (sigla contém o termo)
        try:
            where = {"sigla": f"like:%{termo_upper}%"}
            resultados = self.listar(where=where, limit=limit)
            if resultados:
                return resultados
        except Exception:
            pass

        # Tentativa 2: like server-side por nome
        try:
            where = {"nome": f"like:%{termo_upper}%"}
            resultados = self.listar(where=where, limit=limit)
            if resultados:
                return resultados
        except Exception:
            pass

        # Tentativa 3: lista ampla e filtra client-side
        try:
            todos = self.listar(limit=200)
            termo_l = termo.lower()
            matches = [
                s for s in todos
                if termo_l in (s.get("sigla") or "").lower()
                or termo_l in (s.get("nome") or "").lower()
            ]
            return matches[:limit]
        except Exception:
            return []

    def __enter__(self) -> "SetorClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()
