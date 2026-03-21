"""
hermes/folder.py
Cliente para os endpoints de Folder (Pasta) do SUPP.

Endpoints cobertos:
  GET    /v1/administrativo/folder              Lista paginada com filtros
  GET    /v1/administrativo/folder/count        Contagem com filtros
  GET    /v1/administrativo/folder/{id}         Busca por ID
  POST   /v1/administrativo/folder              Cria pasta
  PUT    /v1/administrativo/folder/{id}         Atualiza pasta (completo)
  PATCH  /v1/administrativo/folder/{id}         Atualiza pasta (parcial)
  DELETE /v1/administrativo/folder/{id}         Remove pasta

Schema:
  Obrigatórios:
    usuario           : int  — ID do usuário dono da pasta
    nome              : str  — Nome da pasta (3–255 chars)
    descricao         : str  — Descrição (3–255 chars)

  Opcionais:
    modalidadeFolder  : int  — ID da modalidade da pasta

Regras de negócio (aplicadas pela API):
  - O nome "compartilhadas" é reservado e não pode ser usado.
  - O limite máximo é de 50 pastas por usuário.

Populate (parâmetro `populate`):
  Associações disponíveis: modalidadeFolder, usuario,
                           criadoPor, atualizadoPor, apagadoPor

Filtros (parâmetro `where`):
  A API aceita JSON no formato {"campo": "operador:valor"}.
  Exemplos:
    {"usuario.id": "eq:42"}
    {"nome": "like:Contratos%"}
    {"modalidadeFolder.id": "eq:3"}
"""

from __future__ import annotations

import json
from typing import Any

import httpx

from .config import BASE_URL

_BASE_PATH = "/v1/administrativo/folder"


class FolderError(Exception):
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
        raise FolderError(response.status_code, body)
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


class FolderClient:
    """
    Cliente síncrono para os endpoints de Folder (Pasta) do SUPP.

    Folder representa uma pasta pessoal de organização de documentos de um
    usuário. Cada usuário pode ter no máximo 50 pastas. O nome "compartilhadas"
    é reservado pela API e não pode ser usado.

    Exemplo de uso:

        from hermes.auth import AuthClient
        from hermes.folder import FolderClient

        auth = AuthClient()
        auth.login_ldap("cpf", "senha")
        fc = FolderClient.from_auth(auth)

        # Listar pastas do usuário 42
        pastas = fc.listar_por_usuario(usuario_id=42)

        # Criar pasta
        fc.criar({
            "usuario": 42,
            "nome": "Contratos 2024",
            "descricao": "Pasta de contratos do exercício 2024",
            "modalidadeFolder": 1,
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
    def from_auth(cls, auth_client: Any, timeout: float = 120.0) -> "FolderClient":
        """Cria FolderClient a partir de um AuthClient já autenticado."""
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
        GET /v1/administrativo/folder
        Lista pastas com paginação e filtros.

        Args:
            where:    Filtro JSON. Ex: {"usuario.id": "eq:42"}
            order:    Ordenação. Ex: {"nome": "ASC"}
            limit:    Máximo de registros por página (padrão 25).
            offset:   Início da paginação (padrão 0).
            populate: Associações a popular. Ex: ["modalidadeFolder", "usuario"]
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
        GET /v1/administrativo/folder/count
        Retorna o total de pastas que atendem ao filtro.
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
        folder_id: int | str,
        populate: list[str] | str | None = None,
        context: str | None = None,
    ) -> dict:
        """GET /v1/administrativo/folder/{id} — Busca pasta por ID."""
        params: dict[str, Any] = {}
        if populate is not None:
            params["populate"] = _populate_str(populate)
        if context is not None:
            params["context"] = context

        resp = self._http.get(f"{_BASE_PATH}/{folder_id}", params=params)
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
        Busca TODAS as pastas que atendem ao filtro, paginando automaticamente.

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
                print(f"  Baixadas {offset}/{total} pastas...")

        return todos

    # ── helpers de domínio ────────────────────────────────────────────────────

    def listar_por_usuario(
        self,
        usuario_id: int | str,
        populate: list[str] | str | None = None,
        order: dict | str | None = None,
    ) -> list[dict]:
        """
        Retorna todas as pastas de um usuário, ordenadas por nome.

        Args:
            usuario_id: ID do usuário.
            populate:   Associações a popular (padrão: ["modalidadeFolder"]).
            order:      Ordenação (padrão: {"nome": "ASC"}).
        """
        if populate is None:
            populate = ["modalidadeFolder"]
        if order is None:
            order = {"nome": "ASC"}
        return self.listar_todos(
            where={"usuario.id": f"eq:{usuario_id}"},
            order=order,
            populate=populate,
            verbose=False,
        )

    def buscar_por_nome(
        self,
        nome: str,
        usuario_id: int | str | None = None,
        populate: list[str] | str | None = None,
    ) -> list[dict]:
        """
        Busca pastas pelo nome (busca exata).

        Args:
            nome:       Nome da pasta a buscar.
            usuario_id: Se informado, restringe ao usuário especificado.
            populate:   Associações a popular.
        """
        where: dict = {"nome": f"eq:{nome}"}
        if usuario_id is not None:
            where["usuario.id"] = f"eq:{usuario_id}"
        return self.listar_todos(where=where, populate=populate, verbose=False)

    # ── escrita ───────────────────────────────────────────────────────────────

    def criar(self, dados: dict, context: str | None = None) -> dict:
        """
        POST /v1/administrativo/folder — Cria uma nova pasta.

        Campos obrigatórios em `dados`:
            usuario           : int — ID do usuário dono da pasta
            nome              : str — Nome da pasta (3–255 chars)
            descricao         : str — Descrição (3–255 chars)

        Campos opcionais:
            modalidadeFolder  : int — ID da modalidade da pasta

        Regras de negócio:
            - O nome "compartilhadas" não é permitido.
            - O limite máximo é de 50 pastas por usuário.

        Exemplo:
            fc.criar({
                "usuario": 42,
                "nome": "Contratos 2024",
                "descricao": "Pasta de contratos do exercício 2024",
                "modalidadeFolder": 1,
            })
        """
        params = {"context": context} if context else {}
        resp = self._http.post(_BASE_PATH, json=dados, params=params)
        return _check(resp)

    def atualizar(self, folder_id: int | str, dados: dict, context: str | None = None) -> dict:
        """PUT /v1/administrativo/folder/{id} — Substitui a pasta por completo."""
        params = {"context": context} if context else {}
        resp = self._http.put(f"{_BASE_PATH}/{folder_id}", json=dados, params=params)
        return _check(resp)

    def atualizar_parcial(self, folder_id: int | str, dados: dict, context: str | None = None) -> dict:
        """PATCH /v1/administrativo/folder/{id} — Atualiza campos específicos."""
        params = {"context": context} if context else {}
        resp = self._http.patch(f"{_BASE_PATH}/{folder_id}", json=dados, params=params)
        return _check(resp)

    def deletar(self, folder_id: int | str) -> Any:
        """DELETE /v1/administrativo/folder/{id} — Remove a pasta."""
        resp = self._http.delete(f"{_BASE_PATH}/{folder_id}")
        return _check(resp)

    def renomear(self, folder_id: int | str, novo_nome: str, nova_descricao: str | None = None) -> dict:
        """
        Atalho para renomear uma pasta via PATCH.

        Args:
            folder_id:      ID da pasta.
            novo_nome:      Novo nome (não pode ser "compartilhadas").
            nova_descricao: Nova descrição (opcional; mantém a atual se omitida).
        """
        dados: dict = {"nome": novo_nome}
        if nova_descricao is not None:
            dados["descricao"] = nova_descricao
        return self.atualizar_parcial(folder_id, dados)

    # ── context manager ───────────────────────────────────────────────────────

    def __enter__(self) -> "FolderClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()
