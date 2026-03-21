"""
hermes/area_trabalho.py
Cliente para os endpoints de ÁreaTrabalho do SUPP.

Endpoints cobertos:
  GET    /v1/administrativo/area_trabalho              Lista paginada com filtros
  GET    /v1/administrativo/area_trabalho/count        Contagem com filtros
  GET    /v1/administrativo/area_trabalho/{id}         Busca por ID
  POST   /v1/administrativo/area_trabalho              Cria vínculo documento-usuário
  PUT    /v1/administrativo/area_trabalho/{id}         Atualiza vínculo (completo)
  PATCH  /v1/administrativo/area_trabalho/{id}         Atualiza vínculo (parcial)
  DELETE /v1/administrativo/area_trabalho/{id}         Remove vínculo

Schema (campos obrigatórios):
  documento : int  — ID do documento
  usuario   : int  — ID do usuário
  dono      : bool — Indica se o usuário é dono (padrão: true)

Populate (parâmetro `populate`):
  Associações disponíveis: documento, usuario, criadoPor, atualizadoPor, apagadoPor
  Exemplo: populate=["documento", "usuario"]

Filtros (parâmetro `where`):
  A API aceita JSON no formato {"campo": "operador:valor"}.
  Exemplos:
    {"usuario.id": "eq:42"}
    {"dono": "eq:1"}
    {"documento.id": "eq:100"}
"""

from __future__ import annotations

import json
from typing import Any

import httpx

from .config import BASE_URL

_BASE_PATH = "/v1/administrativo/area_trabalho"


class AreaTrabalhoError(Exception):
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
        raise AreaTrabalhoError(response.status_code, body)
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


class AreaTrabalhoClient:
    """
    Cliente síncrono para os endpoints de ÁreaTrabalho do SUPP.

    ÁreaTrabalho representa o vínculo entre um documento e um usuário
    na área de trabalho (pasta pessoal) do sistema.

    Exemplo de uso:

        from hermes.auth import AuthClient
        from hermes.area_trabalho import AreaTrabalhoClient

        auth = AuthClient()
        auth.login_ldap("cpf", "senha")
        area = AreaTrabalhoClient.from_auth(auth)

        # Listar documentos na área de trabalho do usuário 42
        vinculos = area.listar(where={"usuario.id": "eq:42"}, populate=["documento"])
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
    def from_auth(cls, auth_client: Any, timeout: float = 120.0) -> "AreaTrabalhoClient":
        """Cria AreaTrabalhoClient a partir de um AuthClient já autenticado."""
        if not auth_client.token:
            raise RuntimeError("AuthClient sem token. Faça login primeiro.")
        return cls(token=auth_client.token, base_url=auth_client.base_url, timeout=timeout)

    # ── leitura ───────────────────────────────────────────────────────────────

    def listar(
        self,
        where: dict | str | None = None,
        order: dict | str | None = None,
        limit: int = 25,
        offset: int = 0,
        populate: list[str] | str | None = None,
        context: str | None = None,
    ) -> list[dict]:
        """
        GET /v1/administrativo/area_trabalho
        Lista vínculos de área de trabalho com paginação e filtros.

        Args:
            where:    Filtro JSON. Ex: {"usuario.id": "eq:42"}
            order:    Ordenação. Ex: {"criadoEm": "DESC"}
            limit:    Máximo de registros por página (padrão 25).
            offset:   Início da paginação (padrão 0).
            populate: Associações a popular. Ex: ["documento", "usuario"]
            context:  Contexto opcional da API.
        """
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
        data = _check(resp)
        return _extract_list(data)

    def contar(
        self,
        where: dict | str | None = None,
        context: str | None = None,
    ) -> int:
        """
        GET /v1/administrativo/area_trabalho/count
        Retorna o total de vínculos que atendem ao filtro.
        """
        params: dict[str, Any] = {}
        if where is not None:
            params["where"] = _where_str(where)
        if context is not None:
            params["context"] = context

        resp = self._http.get(f"{_BASE_PATH}/count", params=params)
        data = _check(resp)
        if isinstance(data, dict):
            return int(data.get("count", data.get("total", 0)))
        return int(data)

    def buscar(
        self,
        area_id: int | str,
        populate: list[str] | str | None = None,
        context: str | None = None,
    ) -> dict:
        """GET /v1/administrativo/area_trabalho/{id} — Busca vínculo por ID."""
        params: dict[str, Any] = {}
        if populate is not None:
            params["populate"] = _populate_str(populate)
        if context is not None:
            params["context"] = context

        resp = self._http.get(f"{_BASE_PATH}/{area_id}", params=params)
        return _check(resp)

    # ── paginação automática ──────────────────────────────────────────────────

    def listar_todos(
        self,
        where: dict | str | None = None,
        order: dict | str | None = None,
        populate: list[str] | str | None = None,
        context: str | None = None,
        page_size: int = 50,
        verbose: bool = True,
    ) -> list[dict]:
        """
        Busca TODOS os vínculos que atendem ao filtro, paginando automaticamente.

        Args:
            where:     Filtro JSON. Ex: {"usuario.id": "eq:42"}
            order:     Ordenação.
            populate:  Associações a popular.
            context:   Contexto da API.
            page_size: Registros por requisição (padrão 50).
            verbose:   Imprime progresso no terminal.
        """
        total = self.contar(where=where, context=context)
        if verbose:
            print(f"  Total encontrado pelo /count: {total}")

        todos: list[dict] = []
        offset = 0

        while offset < total:
            pagina = self.listar(
                where=where,
                order=order,
                limit=page_size,
                offset=offset,
                populate=populate,
                context=context,
            )
            if not pagina:
                break
            todos.extend(pagina)
            offset += len(pagina)
            if verbose:
                print(f"  Baixados {offset}/{total} vínculos...")

        return todos

    # ── helpers de domínio ────────────────────────────────────────────────────

    def listar_por_usuario(
        self,
        usuario_id: int | str,
        populate: list[str] | str | None = None,
        somente_dono: bool = False,
    ) -> list[dict]:
        """
        Retorna todos os documentos na área de trabalho de um usuário.

        Args:
            usuario_id:   ID do usuário.
            populate:     Associações a popular (padrão: ["documento"]).
            somente_dono: Se True, retorna apenas registros onde dono=True.
        """
        where: dict = {"usuario.id": f"eq:{usuario_id}"}
        if somente_dono:
            where["dono"] = "eq:1"
        if populate is None:
            populate = ["documento"]
        return self.listar_todos(where=where, populate=populate, verbose=False)

    def listar_por_documento(
        self,
        documento_id: int | str,
        populate: list[str] | str | None = None,
    ) -> list[dict]:
        """
        Retorna todos os vínculos de usuários para um documento específico.

        Args:
            documento_id: ID do documento.
            populate:     Associações a popular (padrão: ["usuario"]).
        """
        where = {"documento.id": f"eq:{documento_id}"}
        if populate is None:
            populate = ["usuario"]
        return self.listar_todos(where=where, populate=populate, verbose=False)

    # ── escrita ───────────────────────────────────────────────────────────────

    def criar(self, dados: dict, context: str | None = None) -> dict:
        """
        POST /v1/administrativo/area_trabalho — Cria um novo vínculo.

        Campos obrigatórios em `dados`:
            documento : int  — ID do documento
            usuario   : int  — ID do usuário

        Campo opcional:
            dono      : bool — Indica se o usuário é dono (padrão: true)

        Exemplo:
            area.criar({"documento": 55, "usuario": 42, "dono": True})
        """
        params = {"context": context} if context else {}
        resp = self._http.post(_BASE_PATH, json=dados, params=params)
        return _check(resp)

    def atualizar(self, area_id: int | str, dados: dict, context: str | None = None) -> dict:
        """PUT /v1/administrativo/area_trabalho/{id} — Substitui o vínculo por completo."""
        params = {"context": context} if context else {}
        resp = self._http.put(f"{_BASE_PATH}/{area_id}", json=dados, params=params)
        return _check(resp)

    def atualizar_parcial(self, area_id: int | str, dados: dict, context: str | None = None) -> dict:
        """PATCH /v1/administrativo/area_trabalho/{id} — Atualiza campos específicos."""
        params = {"context": context} if context else {}
        resp = self._http.patch(f"{_BASE_PATH}/{area_id}", json=dados, params=params)
        return _check(resp)

    def deletar(self, area_id: int | str) -> Any:
        """DELETE /v1/administrativo/area_trabalho/{id} — Remove o vínculo."""
        resp = self._http.delete(f"{_BASE_PATH}/{area_id}")
        return _check(resp)

    # ── context manager ───────────────────────────────────────────────────────

    def __enter__(self) -> "AreaTrabalhoClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()
