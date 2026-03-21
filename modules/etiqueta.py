"""
hermes/etiqueta.py
Cliente para VinculacaoEtiqueta do SUPP.

VinculacaoEtiqueta é a entidade que vincula uma Etiqueta a um Processo ou Tarefa
(e também a Documento, DocumentoAvulso, Relatório, Setor, Unidade, etc.).

Endpoints cobertos:
  GET    /v1/administrativo/vinculacao_etiqueta                Lista paginada com filtros
  GET    /v1/administrativo/vinculacao_etiqueta/count          Contagem com filtros
  GET    /v1/administrativo/vinculacao_etiqueta/{id}           Busca por ID
  POST   /v1/administrativo/vinculacao_etiqueta                Cria vínculo
  PUT    /v1/administrativo/vinculacao_etiqueta/{id}           Atualiza (completo)
  PATCH  /v1/administrativo/vinculacao_etiqueta/{id}           Atualiza (parcial)
  DELETE /v1/administrativo/vinculacao_etiqueta/{id}           Remove vínculo

Associações disponíveis no populate:
  tarefa, documento, processo, documentoAvulso, relatorio,
  etiqueta, usuario, setor, unidade, modalidadeOrgaoCentral, regraEtiquetaOrigem

Uso típico no sync:
  # Etiquetas do processo
  ec.listar_por_processo(processo_id)

  # Etiquetas da tarefa
  ec.listar_por_tarefa(tarefa_id)
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import httpx

from .config import BASE_URL

_BASE_PATH = "/v1/administrativo/vinculacao_etiqueta"

_POPULATE_PADRAO = ["etiqueta"]


class EtiquetaError(Exception):
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
        raise EtiquetaError(response.status_code, body)
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


# ═══════════════════════════════════════════════════════════════════════════════
# EtiquetaClient
# ═══════════════════════════════════════════════════════════════════════════════

class EtiquetaClient:
    """
    Cliente síncrono para VinculacaoEtiqueta do SUPP.

    Uso:
        from hermes.auth import AuthClient
        from hermes.etiqueta import EtiquetaClient

        auth = AuthClient()
        auth.login_ldap("cpf", "senha")
        ec = EtiquetaClient.from_auth(auth)

        # Etiquetas vinculadas a um processo
        etiquetas = ec.listar_por_processo(processo_id=42)

        # Etiquetas vinculadas a uma tarefa
        etiquetas = ec.listar_por_tarefa(tarefa_id=1001)
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
    def from_auth(cls, auth_client: Any, timeout: float = 120.0) -> "EtiquetaClient":
        """Cria EtiquetaClient a partir de um AuthClient já autenticado."""
        if not auth_client.token:
            raise RuntimeError("AuthClient sem token. Faça login primeiro.")
        return cls(token=auth_client.token, base_url=auth_client.base_url, timeout=timeout)

    # ── consulta ──────────────────────────────────────────────────────────────

    def buscar(
        self,
        vinculacao_id: int | str,
        populate: list[str] | str | None = None,
    ) -> dict:
        """GET /vinculacao_etiqueta/{id} — Busca vínculo por ID."""
        params: dict[str, Any] = {}
        if populate is not None:
            params["populate"] = _populate_str(populate)
        resp = self._http.get(f"{_BASE_PATH}/{vinculacao_id}", params=params)
        return _check(resp)

    def listar(
        self,
        where: dict | str | None = None,
        order: dict | str | None = None,
        limit: int = 25,
        offset: int = 0,
        populate: list[str] | str | None = None,
        context: str | None = None,
    ) -> list[dict]:
        """GET /vinculacao_etiqueta — Lista vínculos com paginação e filtros."""
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

    def contar(
        self,
        where: dict | str | None = None,
        context: str | None = None,
    ) -> int:
        """GET /vinculacao_etiqueta/count — Contagem com filtro."""
        params: dict[str, Any] = {}
        if where is not None:
            params["where"] = _where_str(where)
        if context is not None:
            params["context"] = context
        data = _check(self._http.get(f"{_BASE_PATH}/count", params=params))
        if isinstance(data, dict):
            return int(data.get("count", data.get("total", 0)))
        return int(data)

    def listar_todos(
        self,
        where: dict | str,
        order: dict | str | None = None,
        populate: list[str] | str | None = None,
        page_size: int = 50,
    ) -> list[dict]:
        """
        Busca TODOS os vínculos que atendem ao filtro, paginando automaticamente.

        O parâmetro `where` é obrigatório para evitar timeout em bases grandes.
        """
        total = self.contar(where=where)
        todos: list[dict] = []
        offset = 0
        while offset < total:
            pagina = self.listar(
                where=where,
                order=order,
                limit=page_size,
                offset=offset,
                populate=populate,
            )
            if not pagina:
                break
            todos.extend(pagina)
            offset += len(pagina)
        return todos

    # ── helpers por entidade ──────────────────────────────────────────────────

    def listar_por_processo(
        self,
        processo_id: int | str,
        populate: list[str] | str | None = None,
        page_size: int = 50,
    ) -> list[dict]:
        """
        Lista todas as etiquetas vinculadas a um processo.

        Equivale a:
          GET /vinculacao_etiqueta?where={"processo.id":"eq:{id}"}&populate=["etiqueta"]
        """
        where = {"processo.id": f"eq:{processo_id}"}
        pop = populate if populate is not None else _POPULATE_PADRAO
        return self.listar_todos(where=where, populate=pop, page_size=page_size)

    def listar_por_tarefa(
        self,
        tarefa_id: int | str,
        populate: list[str] | str | None = None,
        page_size: int = 50,
    ) -> list[dict]:
        """
        Lista todas as etiquetas vinculadas a uma tarefa.

        Equivale a:
          GET /vinculacao_etiqueta?where={"tarefa.id":"eq:{id}"}&populate=["etiqueta"]
        """
        where = {"tarefa.id": f"eq:{tarefa_id}"}
        pop = populate if populate is not None else _POPULATE_PADRAO
        return self.listar_todos(where=where, populate=pop, page_size=page_size)

    def listar_por_documento(
        self,
        documento_id: int | str,
        populate: list[str] | str | None = None,
        page_size: int = 50,
    ) -> list[dict]:
        """
        Lista todas as etiquetas vinculadas a um documento.

        Equivale a:
          GET /vinculacao_etiqueta?where={"documento.id":"eq:{id}"}&populate=["etiqueta"]
        """
        where = {"documento.id": f"eq:{documento_id}"}
        pop = populate if populate is not None else _POPULATE_PADRAO
        return self.listar_todos(where=where, populate=pop, page_size=page_size)

    # ── escrita ───────────────────────────────────────────────────────────────

    def criar(self, dados: dict, context: str | None = None) -> dict:
        """
        POST /vinculacao_etiqueta — Cria um vínculo de etiqueta.

        Campos principais do body:
          - etiqueta:           {"id": <etiqueta_id>}  (obrigatório)
          - processo:           {"id": <processo_id>}  (ou tarefa, documento, etc.)
          - tarefa:             {"id": <tarefa_id>}
          - conteudo:           str  (texto livre da etiqueta, se aplicável)
          - privada:            bool
          - dataHoraExpiracao:  ISO 8601 str  (opcional)
        """
        params = {"context": context} if context else {}
        resp = self._http.post(_BASE_PATH, json=dados, params=params)
        return _check(resp)

    def atualizar(
        self,
        vinculacao_id: int | str,
        dados: dict,
        context: str | None = None,
    ) -> dict:
        """PUT /vinculacao_etiqueta/{id} — Substitui o vínculo por completo."""
        params = {"context": context} if context else {}
        resp = self._http.put(f"{_BASE_PATH}/{vinculacao_id}", json=dados, params=params)
        return _check(resp)

    def atualizar_parcial(
        self,
        vinculacao_id: int | str,
        dados: dict,
        context: str | None = None,
    ) -> dict:
        """PATCH /vinculacao_etiqueta/{id} — Atualiza campos específicos."""
        params = {"context": context} if context else {}
        resp = self._http.patch(f"{_BASE_PATH}/{vinculacao_id}", json=dados, params=params)
        return _check(resp)

    def deletar(self, vinculacao_id: int | str) -> Any:
        """DELETE /vinculacao_etiqueta/{id} — Remove o vínculo."""
        resp = self._http.delete(f"{_BASE_PATH}/{vinculacao_id}")
        return _check(resp)

    # ── context manager ───────────────────────────────────────────────────────

    def __enter__(self) -> "EtiquetaClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()
