"""
hermes/formulario_ia.py
Cliente para Formulários de IA do SUPP.

Entidades:
  DadosFormulario  — resposta preenchida de um formulário de IA vinculada a um
                     Documento / Processo (dataValue contém o JSON de análise)
  Formulario       — definição do formulário (schema, nome, sigla)

Endpoints cobertos:
  --- DadosFormulario ---
  GET    /v1/administrativo/dados_formulario                  Lista paginada com filtros
  GET    /v1/administrativo/dados_formulario/count            Contagem com filtros
  GET    /v1/administrativo/dados_formulario/{id}             Busca por ID
  POST   /v1/administrativo/dados_formulario                  Cria
  PUT    /v1/administrativo/dados_formulario/{id}             Atualiza (completo)
  PATCH  /v1/administrativo/dados_formulario/{id}             Atualiza (parcial)
  DELETE /v1/administrativo/dados_formulario/{id}             Remove

  --- Triagem (IA) ---
  GET    /inteligencia_artificial/triagem/{documentoId}/formularios
                                                              Formulários de IA de um documento

  --- Formulario (definição) ---
  GET    /v1/administrativo/formulario                        Lista paginada
  GET    /v1/administrativo/formulario/count                  Contagem
  GET    /v1/administrativo/formulario/{id}                   Busca por ID
  GET    /v1/administrativo/formulario/ia/{id}/fields         Campos do formulário de IA

Associações de DadosFormulario (populate):
  formulario, componenteDigital, documento, processo, criadoPor, atualizadoPor
"""

from __future__ import annotations

import json
from typing import Any

import httpx

from .config import BASE_URL

_DADOS_PATH    = "/v1/administrativo/dados_formulario"
_FORMULARIO_PATH = "/v1/administrativo/formulario"
_TRIAGEM_PATH  = "/inteligencia_artificial/triagem"


class FormularioIaError(Exception):
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
        raise FormularioIaError(response.status_code, body)
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


class FormularioIaClient:
    """
    Cliente síncrono para Formulários de IA do SUPP.

    Uso:
        from hermes.auth import AuthClient
        from hermes.formulario_ia import FormularioIaClient

        auth = AuthClient()
        auth.login_ldap("cpf", "senha")
        fc = FormularioIaClient.from_auth(auth)

        # Formulários de IA de um processo
        formularios = fc.listar_por_processo(proc_id)

        # Formulários de IA de um documento via triagem
        triagem = fc.triagem_por_documento(doc_id)
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
    def from_auth(cls, auth_client: Any, timeout: float = 120.0) -> "FormularioIaClient":
        if not auth_client.token:
            raise RuntimeError("AuthClient sem token. Faça login primeiro.")
        return cls(token=auth_client.token, base_url=auth_client.base_url, timeout=timeout)

    # ── DadosFormulario: consulta ─────────────────────────────────────────────

    def buscar(
        self,
        dados_id: int | str,
        populate: list[str] | str | None = None,
        context: str | None = None,
    ) -> dict:
        """GET /dados_formulario/{id} — Busca por ID."""
        params: dict[str, Any] = {}
        if populate is not None:
            params["populate"] = _populate_str(populate)
        if context is not None:
            params["context"] = context
        resp = self._http.get(f"{_DADOS_PATH}/{dados_id}", params=params)
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
        """GET /dados_formulario — Lista paginada com filtros."""
        params: dict[str, Any] = {"limit": limit, "offset": offset}
        if where is not None:
            params["where"] = _where_str(where)
        if order is not None:
            params["order"] = json.dumps(order) if isinstance(order, dict) else order
        if populate is not None:
            params["populate"] = _populate_str(populate)
        if context is not None:
            params["context"] = context
        resp = self._http.get(_DADOS_PATH, params=params)
        return _extract_list(_check(resp))

    def contar(
        self,
        where: dict | str | None = None,
        context: str | None = None,
    ) -> int:
        """GET /dados_formulario/count — Contagem com filtro."""
        params: dict[str, Any] = {}
        if where is not None:
            params["where"] = _where_str(where)
        if context is not None:
            params["context"] = context
        data = _check(self._http.get(f"{_DADOS_PATH}/count", params=params))
        if isinstance(data, dict):
            return int(data.get("count", data.get("total", 0)))
        return int(data)

    def listar_todos(
        self,
        where: dict | str | None,
        populate: list[str] | str | None = None,
        order: dict | str | None = None,
        context: str | None = None,
        page_size: int = 50,
        verbose: bool = False,
    ) -> list[dict]:
        """Busca todos os DadosFormulario que atendem ao filtro, paginando automaticamente."""
        total = self.contar(where=where, context=context)
        if verbose:
            print(f"    {total} formulário(s) de IA encontrado(s)")

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
        return todos

    def listar_por_processo(
        self,
        processo_id: int | str,
        populate: list[str] | str | None = None,
        verbose: bool = False,
    ) -> list[dict]:
        """
        Lista todos os DadosFormulario vinculados a um processo.

        Equivale a GET /dados_formulario?where={"processo.id":"eq:{id}"}
        com populate=["formulario","documento","componenteDigital"] por padrão.
        """
        pop = populate if populate is not None else ["formulario", "documento", "componenteDigital"]
        where = {"processo.id": f"eq:{processo_id}"}
        return self.listar_todos(where=where, populate=pop, verbose=verbose)

    def listar_por_documento(
        self,
        documento_id: int | str,
        populate: list[str] | str | None = None,
    ) -> list[dict]:
        """
        Lista todos os DadosFormulario vinculados a um documento.
        É aqui que ficam os dados realmente extraídos pela IA (campo dataValue).

        Equivale a GET /dados_formulario?where={"documento.id":"eq:{id}"}
        """
        pop = populate if populate is not None else ["formulario", "componenteDigital"]
        where = {"documento.id": f"eq:{documento_id}"}
        return self.listar_todos(where=where, populate=pop)

    def dados_ia_por_documentos(
        self,
        documento_ids: list[int | str],
        verbose: bool = False,
    ) -> list[dict]:
        """
        Busca os dados extraídos pela IA para cada documento informado.

        Para cada documentoId chama:
          GET /dados_formulario?where={"documento.id":"eq:{id}"}&populate=["formulario","componenteDigital"]

        Retorna lista de entradas no formato:
          {"documentoId": <id>, "dados": [<DadosFormulario>, ...]}

        Apenas documentos com ao menos um registro são incluídos.

        Args:
            documento_ids: IDs dos documentos do processo.
            verbose:       Imprime progresso.
        """
        resultado: list[dict] = []
        for doc_id in documento_ids:
            try:
                dados = self.listar_por_documento(doc_id)
                if dados:
                    resultado.append({"documentoId": doc_id, "dados": dados})
                    if verbose:
                        print(f"      Documento {doc_id} → {len(dados)} registro(s) de IA")
            except FormularioIaError as e:
                if verbose:
                    print(f"      Documento {doc_id} → erro ao buscar dados de IA: {e}")
        return resultado

    # ── Triagem ───────────────────────────────────────────────────────────────

    def triagem_por_documento(self, documento_id: int | str) -> list[dict]:
        """
        GET /inteligencia_artificial/triagem/{documentoId}/formularios
        Retorna os formulários de IA disponíveis para triagem de um documento.
        """
        resp = self._http.get(f"{_TRIAGEM_PATH}/{documento_id}/formularios")
        data = _check(resp)
        return data if isinstance(data, list) else _extract_list(data)

    def triagem_por_processo(
        self,
        documento_ids: list[int | str],
        verbose: bool = False,
    ) -> list[dict]:
        """
        Itera pelos IDs de documentos de um processo e chama
        GET /inteligencia_artificial/triagem/{documentoId}/formularios para cada um.

        Retorna lista de entradas no formato:
          {"documentoId": <id>, "formularios": [...]}

        Apenas documentos que retornarem ao menos um formulário são incluídos.

        Args:
            documento_ids: IDs dos documentos do processo.
            verbose:       Imprime progresso.
        """
        resultado: list[dict] = []
        for doc_id in documento_ids:
            try:
                formularios = self.triagem_por_documento(doc_id)
                if formularios:
                    resultado.append({"documentoId": doc_id, "formularios": formularios})
                    if verbose:
                        print(f"      Documento {doc_id} → {len(formularios)} formulário(s) de IA")
            except FormularioIaError as e:
                if verbose:
                    print(f"      Documento {doc_id} → erro na triagem: {e}")
        return resultado

    # ── Formulario (definição): consulta ──────────────────────────────────────

    def buscar_formulario(
        self,
        formulario_id: int | str,
        populate: list[str] | str | None = None,
    ) -> dict:
        """GET /formulario/{id} — Definição completa do formulário."""
        params: dict[str, Any] = {}
        if populate is not None:
            params["populate"] = _populate_str(populate)
        resp = self._http.get(f"{_FORMULARIO_PATH}/{formulario_id}", params=params)
        return _check(resp)

    def campos_formulario_ia(self, formulario_id: int | str) -> Any:
        """GET /formulario/ia/{id}/fields — Campos do formulário de IA."""
        resp = self._http.get(f"{_FORMULARIO_PATH}/ia/{formulario_id}/fields")
        return _check(resp)

    def listar_formularios(
        self,
        where: dict | str | None = None,
        limit: int = 25,
        offset: int = 0,
        populate: list[str] | str | None = None,
    ) -> list[dict]:
        """GET /formulario — Lista definições de formulários."""
        params: dict[str, Any] = {"limit": limit, "offset": offset}
        if where is not None:
            params["where"] = _where_str(where)
        if populate is not None:
            params["populate"] = _populate_str(populate)
        resp = self._http.get(_FORMULARIO_PATH, params=params)
        return _extract_list(_check(resp))

    # ── DadosFormulario: escrita ──────────────────────────────────────────────

    def criar(self, dados: dict, context: str | None = None) -> dict:
        """POST /dados_formulario"""
        params = {"context": context} if context else {}
        resp = self._http.post(_DADOS_PATH, json=dados, params=params)
        return _check(resp)

    def atualizar(self, dados_id: int | str, dados: dict, context: str | None = None) -> dict:
        """PUT /dados_formulario/{id}"""
        params = {"context": context} if context else {}
        resp = self._http.put(f"{_DADOS_PATH}/{dados_id}", json=dados, params=params)
        return _check(resp)

    def atualizar_parcial(self, dados_id: int | str, dados: dict, context: str | None = None) -> dict:
        """PATCH /dados_formulario/{id}"""
        params = {"context": context} if context else {}
        resp = self._http.patch(f"{_DADOS_PATH}/{dados_id}", json=dados, params=params)
        return _check(resp)

    def deletar(self, dados_id: int | str) -> Any:
        """DELETE /dados_formulario/{id}"""
        resp = self._http.delete(f"{_DADOS_PATH}/{dados_id}")
        return _check(resp)

    # ── context manager ───────────────────────────────────────────────────────

    def __enter__(self) -> "FormularioIaClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()
