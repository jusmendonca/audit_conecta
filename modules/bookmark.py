"""
hermes/bookmark.py
Cliente para os endpoints de Bookmark do SUPP.

Endpoints cobertos:
  GET    /v1/administrativo/bookmark              Lista paginada com filtros
  GET    /v1/administrativo/bookmark/count        Contagem com filtros
  GET    /v1/administrativo/bookmark/{id}         Busca por ID
  POST   /v1/administrativo/bookmark              Cria bookmark
  PUT    /v1/administrativo/bookmark/{id}         Atualiza bookmark (completo)
  PATCH  /v1/administrativo/bookmark/{id}         Atualiza bookmark (parcial)
  DELETE /v1/administrativo/bookmark/{id}         Remove bookmark

Schema:
  Obrigatórios:
    nome              : str  — Nome/título do bookmark (máx. 255 chars)
    componenteDigital : int  — ID do componente digital (arquivo)
    processo          : int  — ID do processo
    juntada           : int  — ID da juntada

  Opcionais:
    usuario           : int  — ID do usuário (preenchido automaticamente pela API)
    pagina            : int  — Página do documento (padrão: 0)
    descricao         : str  — Descrição (máx. 512 chars)
    corHexadecimal    : str  — Cor de destaque, ex: "#FF5733"
    geradoPorIa       : bool — Indicador de geração por IA (padrão: false)
    textoReferencia   : str  — Trecho de texto referenciado (máx. 255 chars)

Populate (parâmetro `populate`):
  Associações disponíveis: usuario, componenteDigital, processo, juntada,
                           criadoPor, atualizadoPor, apagadoPor

Filtros (parâmetro `where`):
  A API aceita JSON no formato {"campo": "operador:valor"}.
  Exemplos:
    {"usuario.id": "eq:42"}
    {"processo.id": "eq:100"}
    {"geradoPorIa": "eq:1"}
    {"pagina": "gte:5"}
"""

from __future__ import annotations

import json
from typing import Any

import httpx

from .config import BASE_URL

_BASE_PATH = "/v1/administrativo/bookmark"


class BookmarkError(Exception):
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
        raise BookmarkError(response.status_code, body)
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


class BookmarkClient:
    """
    Cliente síncrono para os endpoints de Bookmark do SUPP.

    Bookmark representa uma marcação em um componente digital (página de um
    documento), vinculada a um processo e juntada específicos. Pode ser criado
    manualmente ou gerado por IA.

    Exemplo de uso:

        from hermes.auth import AuthClient
        from hermes.bookmark import BookmarkClient

        auth = AuthClient()
        auth.login_ldap("cpf", "senha")
        bm = BookmarkClient.from_auth(auth)

        # Listar bookmarks de um processo
        marcacoes = bm.listar_por_processo(processo_id=100, populate=["componenteDigital"])

        # Criar bookmark
        bm.criar({
            "nome": "Cláusula importante",
            "componenteDigital": 55,
            "processo": 100,
            "juntada": 12,
            "pagina": 3,
            "corHexadecimal": "#FF5733",
            "textoReferencia": "O prazo não poderá ser prorrogado.",
        })
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
    def from_auth(cls, auth_client: Any, timeout: float = 120.0) -> "BookmarkClient":
        """Cria BookmarkClient a partir de um AuthClient já autenticado."""
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
        GET /v1/administrativo/bookmark
        Lista bookmarks com paginação e filtros.

        Args:
            where:    Filtro JSON. Ex: {"usuario.id": "eq:42"}
            order:    Ordenação. Ex: {"pagina": "ASC"}
            limit:    Máximo de registros por página (padrão 25).
            offset:   Início da paginação (padrão 0).
            populate: Associações a popular. Ex: ["processo", "componenteDigital"]
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
        GET /v1/administrativo/bookmark/count
        Retorna o total de bookmarks que atendem ao filtro.
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
        bookmark_id: int | str,
        populate: list[str] | str | None = None,
        context: str | None = None,
    ) -> dict:
        """GET /v1/administrativo/bookmark/{id} — Busca bookmark por ID."""
        params: dict[str, Any] = {}
        if populate is not None:
            params["populate"] = _populate_str(populate)
        if context is not None:
            params["context"] = context

        resp = self._http.get(f"{_BASE_PATH}/{bookmark_id}", params=params)
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
        Busca TODOS os bookmarks que atendem ao filtro, paginando automaticamente.

        Args:
            where:     Filtro JSON. Ex: {"processo.id": "eq:100"}
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
                print(f"  Baixados {offset}/{total} bookmarks...")

        return todos

    # ── helpers de domínio ────────────────────────────────────────────────────

    def listar_por_processo(
        self,
        processo_id: int | str,
        populate: list[str] | str | None = None,
        order: dict | str | None = None,
    ) -> list[dict]:
        """
        Retorna todos os bookmarks de um processo, ordenados por página.

        Args:
            processo_id: ID do processo.
            populate:    Associações a popular (padrão: ["componenteDigital", "juntada"]).
            order:       Ordenação (padrão: {"pagina": "ASC"}).
        """
        if populate is None:
            populate = ["componenteDigital", "juntada"]
        if order is None:
            order = {"pagina": "ASC"}
        return self.listar_todos(
            where={"processo.id": f"eq:{processo_id}"},
            order=order,
            populate=populate,
            verbose=False,
        )

    def listar_por_componente(
        self,
        componente_id: int | str,
        populate: list[str] | str | None = None,
    ) -> list[dict]:
        """
        Retorna todos os bookmarks de um componente digital (arquivo), por página.

        Args:
            componente_id: ID do componente digital.
            populate:      Associações a popular (padrão: ["processo"]).
        """
        if populate is None:
            populate = ["processo"]
        return self.listar_todos(
            where={"componenteDigital.id": f"eq:{componente_id}"},
            order={"pagina": "ASC"},
            populate=populate,
            verbose=False,
        )

    def listar_por_usuario(
        self,
        usuario_id: int | str,
        populate: list[str] | str | None = None,
        somente_ia: bool | None = None,
    ) -> list[dict]:
        """
        Retorna todos os bookmarks de um usuário.

        Args:
            usuario_id: ID do usuário.
            populate:   Associações a popular (padrão: ["processo", "componenteDigital"]).
            somente_ia: Se True, retorna apenas bookmarks gerados por IA.
                        Se False, retorna apenas os criados manualmente.
                        Se None (padrão), retorna todos.
        """
        where: dict = {"usuario.id": f"eq:{usuario_id}"}
        if somente_ia is True:
            where["geradoPorIa"] = "eq:1"
        elif somente_ia is False:
            where["geradoPorIa"] = "eq:0"
        if populate is None:
            populate = ["processo", "componenteDigital"]
        return self.listar_todos(where=where, populate=populate, verbose=False)

    # ── escrita ───────────────────────────────────────────────────────────────

    def criar(self, dados: dict, context: str | None = None) -> dict:
        """
        POST /v1/administrativo/bookmark — Cria um novo bookmark.

        Campos obrigatórios em `dados`:
            nome              : str — Nome/título do bookmark
            componenteDigital : int — ID do componente digital
            processo          : int — ID do processo
            juntada           : int — ID da juntada

        Campos opcionais:
            pagina            : int  — Página do documento (padrão: 0)
            descricao         : str  — Descrição livre
            corHexadecimal    : str  — Cor de destaque, ex: "#FF5733"
            textoReferencia   : str  — Trecho de texto referenciado
            geradoPorIa       : bool — Indica geração automática por IA

        Exemplo:
            bm.criar({
                "nome": "Cláusula de rescisão",
                "componenteDigital": 55,
                "processo": 100,
                "juntada": 12,
                "pagina": 7,
                "corHexadecimal": "#FFD700",
            })
        """
        params = {"context": context} if context else {}
        resp = self._http.post(_BASE_PATH, json=dados, params=params)
        return _check(resp)

    def atualizar(self, bookmark_id: int | str, dados: dict, context: str | None = None) -> dict:
        """PUT /v1/administrativo/bookmark/{id} — Substitui o bookmark por completo."""
        params = {"context": context} if context else {}
        resp = self._http.put(f"{_BASE_PATH}/{bookmark_id}", json=dados, params=params)
        return _check(resp)

    def atualizar_parcial(self, bookmark_id: int | str, dados: dict, context: str | None = None) -> dict:
        """PATCH /v1/administrativo/bookmark/{id} — Atualiza campos específicos."""
        params = {"context": context} if context else {}
        resp = self._http.patch(f"{_BASE_PATH}/{bookmark_id}", json=dados, params=params)
        return _check(resp)

    def deletar(self, bookmark_id: int | str) -> Any:
        """DELETE /v1/administrativo/bookmark/{id} — Remove o bookmark."""
        resp = self._http.delete(f"{_BASE_PATH}/{bookmark_id}")
        return _check(resp)

    # ── context manager ───────────────────────────────────────────────────────

    def __enter__(self) -> "BookmarkClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()
