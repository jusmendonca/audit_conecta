"""
modules/atividade.py
Cliente para os endpoints de Atividade do SUPP.

Endpoints cobertos:
  GET    /v1/administrativo/atividade              Lista paginada com filtros
  GET    /v1/administrativo/atividade/count        Contagem com filtros
  GET    /v1/administrativo/atividade/{id}         Busca por ID
  GET    /v1/judicial/atividade_judicial           Lista atividades judiciais
  GET    /v1/judicial/atividade_judicial/count     Contagem de atividades judiciais
  GET    /v1/consultivo/atividade_consultiva       Lista atividades consultivas
  GET    /v1/consultivo/atividade_consultiva/count Contagem de atividades consultivas

Campos da entidade Atividade (observados nos dados reais):
  tarefa                int/obj  — tarefa vinculada
  dataHoraConclusao     string   — data/hora de conclusão (ISO 8601)
  encerraTarefa         bool     — se a atividade encerra a tarefa
  distribuicaoAutomatica bool    — distribuição automática
  criadoEm              string   — data/hora de criação (ISO 8601)
  atualizadoEm          string   — data/hora de atualização (ISO 8601)

Associações disponíveis no populate:
  tarefa, setor, usuario, especieAtividade, criadoPor, atualizadoPor

Uso típico:
  # Atividades judiciais de uma tarefa (mais comum)
  ac = AtividadeClient(token=token, base_path=BASE_PATH_JUDICIAL)
  ac.listar_por_tarefa(tarefa_id)
"""

from __future__ import annotations

import json
from typing import Any

import httpx

from .config import BASE_URL

# Paths disponíveis — escolha conforme o tipo de tarefa
BASE_PATH_ADMINISTRATIVO = "/v1/administrativo/atividade"
BASE_PATH_JUDICIAL       = "/v1/judicial/atividade_judicial"
BASE_PATH_CONSULTIVO     = "/v1/consultivo/atividade_consultiva"

# Compatibilidade retroativa
_BASE_PATH = BASE_PATH_ADMINISTRATIVO

_POPULATE_PADRAO = ["tarefa", "setor", "usuario", "especieAtividade", "criadoPor"]


class AtividadeError(Exception):
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
        raise AtividadeError(response.status_code, body)
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


class AtividadeClient:
    """
    Cliente síncrono para os endpoints de Atividade do SUPP.

    Suporta os três endpoints de atividade via `base_path`:
      - BASE_PATH_ADMINISTRATIVO (padrão)
      - BASE_PATH_JUDICIAL
      - BASE_PATH_CONSULTIVO

    Uso:
        from modules.atividade import AtividadeClient, BASE_PATH_JUDICIAL

        with AtividadeClient(token=token, base_path=BASE_PATH_JUDICIAL) as ac:
            atividades = ac.listar_por_tarefa(tarefa_id=290361966)
    """

    def __init__(
        self,
        token: str,
        base_url: str = BASE_URL,
        timeout: float = 120.0,
        base_path: str = BASE_PATH_ADMINISTRATIVO,
    ) -> None:
        self.token = token
        self.base_url = base_url
        self._base_path = base_path
        self._http = httpx.Client(
            base_url=base_url,
            timeout=timeout,
            headers={"Authorization": f"Bearer {token}"},
        )

    def __enter__(self) -> "AtividadeClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()

    # ── consulta ──────────────────────────────────────────────────────────────

    def buscar(
        self,
        atividade_id: int | str,
        populate: list[str] | str | None = None,
    ) -> dict:
        """GET {base_path}/{id} — Busca atividade por ID."""
        params: dict[str, Any] = {}
        if populate is not None:
            params["populate"] = _populate_str(populate)
        resp = self._http.get(f"{self._base_path}/{atividade_id}", params=params)
        return _check(resp)

    def listar(
        self,
        where: dict | str | None = None,
        order: dict | str | None = None,
        limit: int = 25,
        offset: int = 0,
        populate: list[str] | str | None = None,
    ) -> list[dict]:
        """GET {base_path} — Lista atividades com filtros."""
        params: dict[str, Any] = {"limit": limit, "offset": offset}
        if where is not None:
            params["where"] = _where_str(where)
        if order is not None:
            params["order"] = json.dumps(order) if isinstance(order, dict) else order
        if populate is not None:
            params["populate"] = _populate_str(populate)
        resp = self._http.get(self._base_path, params=params)
        return _extract_list(_check(resp))

    def contar(self, where: dict | str | None = None) -> int:
        """GET {base_path}/count — Contagem com filtro."""
        params: dict[str, Any] = {}
        if where is not None:
            params["where"] = _where_str(where)
        data = _check(self._http.get(f"{self._base_path}/count", params=params))
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
        Busca TODAS as atividades que atendem ao filtro, paginando automaticamente.
        O parâmetro `where` é obrigatório para evitar timeout.
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

    # ── helper por entidade ───────────────────────────────────────────────────

    def listar_por_tarefa(
        self,
        tarefa_id: int | str,
        populate: list[str] | str | None = None,
        order: dict | str | None = None,
        page_size: int = 50,
    ) -> list[dict]:
        """
        Lista todas as atividades vinculadas a uma tarefa.

        Equivale a:
          GET /atividade?where={"tarefa.id":"eq:{id}"}&populate=[...]
        """
        where = {"tarefa.id": f"eq:{tarefa_id}"}
        pop = populate if populate is not None else _POPULATE_PADRAO
        ord_ = order if order is not None else {"criadoEm": "ASC"}
        return self.listar_todos(where=where, order=ord_, populate=pop, page_size=page_size)
