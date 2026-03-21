"""
hermes/tarefa.py
Cliente para os endpoints de Tarefa do SUPP.

Endpoints cobertos:
  GET  /v1/administrativo/tarefa/findTarefasPendentesPainel  Tarefas pendentes do painel
  GET  /v1/administrativo/tarefa                             Lista paginada com filtros
  GET  /v1/administrativo/tarefa/count                       Contagem com filtros
  GET  /v1/administrativo/tarefa/{id}                        Busca por ID
  POST /v1/administrativo/tarefa                             Cria tarefa
  PUT  /v1/administrativo/tarefa/{id}                        Atualiza tarefa (completo)
  PATCH /v1/administrativo/tarefa/{id}                       Atualiza tarefa (parcial)

Filtros (parâmetro `where`):
  A API aceita JSON no formato {"campo": "operador:valor"}.
  Exemplos:
    {"urgente": "eq:1"}
    {"dataHoraFinalPrazo": "lt:2026-03-10T23:59:59"}
    {"usuarioResponsavel.id": "eq:42"}

Populate (parâmetro `populate`):
  Associações disponíveis: processo, vinculacaoWorkflow
  Exemplo: populate=["processo", "vinculacaoWorkflow"]
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

import httpx

from .config import BASE_URL

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
TMP_DIR = _PROJECT_ROOT / "tmp"

_BASE_PATH = "/v1/administrativo/tarefa"


class TarefaError(Exception):
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
        raise TarefaError(response.status_code, body)
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


def _extract_list(data: Any, verbose: bool = False) -> list[dict]:
    """
    Normaliza a resposta da API, que pode vir em vários formatos.
    Com verbose=True imprime a estrutura recebida para diagnóstico.
    """
    if verbose:
        if isinstance(data, dict):
            print(f"  [debug] chaves da resposta: {list(data.keys())}")
            for k, v in data.items():
                preview = str(v)[:120]
                print(f"    {k!r}: {preview}")
        elif isinstance(data, list):
            print(f"  [debug] resposta é lista com {len(data)} itens")
        else:
            print(f"  [debug] resposta inesperada: {type(data)} → {str(data)[:200]}")

    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("entities", "data", "results", "items", "tarefas"):
            if key in data and isinstance(data[key], list):
                return data[key]
    return []


class TarefaClient:
    """
    Cliente síncrono para os endpoints de Tarefa do SUPP.

    Requer um token JWT válido:

        from hermes.auth import AuthClient
        from hermes.tarefa import TarefaClient

        auth = AuthClient()
        auth.login_ldap("cpf", "senha")
        tarefas = TarefaClient.from_auth(auth)
    """

    def __init__(
        self,
        token: str,
        base_url: str = BASE_URL,
        timeout: float = 120.0,      # aumentado: queries pesadas precisam de mais tempo
    ) -> None:
        self.token = token
        self.base_url = base_url
        self._http = httpx.Client(
            base_url=base_url,
            timeout=timeout,
            headers={"Authorization": f"Bearer {token}"},
        )

    @classmethod
    def from_auth(cls, auth_client: Any, timeout: float = 120.0) -> "TarefaClient":
        """Cria TarefaClient a partir de um AuthClient já autenticado."""
        if not auth_client.token:
            raise RuntimeError("AuthClient sem token. Faça login primeiro.")
        return cls(token=auth_client.token, base_url=auth_client.base_url, timeout=timeout)

    # ── leitura ───────────────────────────────────────────────────────────────

    def pendentes_painel(self, debug: bool = False) -> list[dict]:
        """
        GET /v1/administrativo/tarefa/findTarefasPendentesPainel
        Tarefas pendentes do painel do usuário logado.

        Use debug=True para imprimir a estrutura bruta da resposta
        e diagnosticar problemas de parsing.
        """
        resp = self._http.get(f"{_BASE_PATH}/findTarefasPendentesPainel")
        data = _check(resp)
        return _extract_list(data, verbose=debug)

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
        GET /v1/administrativo/tarefa
        Lista tarefas com paginação e filtros.

        Args:
            where:    Filtro JSON. Ex: {"urgente": "eq:1"}
            order:    Ordenação. Ex: {"dataHoraFinalPrazo": "ASC"}
            limit:    Máximo de registros por página (padrão 25).
            offset:   Início da paginação (padrão 0).
            populate: Associações a popular. Ex: ["processo"]
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
        GET /v1/administrativo/tarefa/count
        Retorna o total de tarefas que atendem ao filtro.

        Atenção: chamar sem `where` conta TODAS as tarefas do sistema
        e pode ser muito lento. Prefira sempre passar um filtro.
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
        tarefa_id: int | str,
        populate: list[str] | str | None = None,
        context: str | None = None,
    ) -> dict:
        """GET /v1/administrativo/tarefa/{id} — Busca tarefa por ID."""
        params: dict[str, Any] = {}
        if populate is not None:
            params["populate"] = _populate_str(populate)
        if context is not None:
            params["context"] = context

        resp = self._http.get(f"{_BASE_PATH}/{tarefa_id}", params=params)
        return _check(resp)

    # ── paginação automática ──────────────────────────────────────────────────

    def listar_todos(
        self,
        where: dict | str | None,          # obrigatório na prática — sem filtro = timeout
        order: dict | str | None = None,
        populate: list[str] | str | None = None,
        context: str | None = None,
        page_size: int = 50,
        verbose: bool = True,
    ) -> list[dict]:
        """
        Busca TODAS as tarefas que atendem ao filtro, paginando automaticamente.

        O parâmetro `where` é obrigatório na prática: consultas sem filtro
        podem timeout em bases com muitos registros.

        Args:
            where:     Filtro JSON — ex: {"usuarioResponsavel.id": "eq:99"}
            order:     Ordenação.
            populate:  Associações a popular.
            context:   Contexto da API.
            page_size: Registros por requisição (padrão 50).
            verbose:   Imprime progresso no terminal.
        """
        total = self.contar(where=where, context=context)
        if verbose:
            print(f"  Total encontrado pelo /count: {total}")

        todas: list[dict] = []
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
            todas.extend(pagina)
            offset += len(pagina)
            if verbose:
                print(f"  Baixadas {offset}/{total} tarefas...")

        return todas

    # ── download para disco ───────────────────────────────────────────────────

    def baixar_pendentes(
        self,
        dest_dir: Path | str | None = None,
        populate: list[str] | str | None = None,
        verbose: bool = True,
        debug: bool = False,
    ) -> Path:
        """
        Baixa as tarefas pendentes e salva em JSON na pasta temporária.

        Fluxo:
          1. Chama `findTarefasPendentesPainel` (endpoint dedicado, rápido).
          2. Se retornar vazio, exibe orientação — NÃO faz fallback sem filtro
             para evitar timeout em bases grandes. Use `baixar_com_filtro()`.

        Args:
            dest_dir: Pasta de destino. Padrão: <projeto>/tmp/
            populate: Associações a popular (ex: ["processo"]).
            verbose:  Imprime progresso.
            debug:    Imprime estrutura bruta da resposta para diagnóstico.

        Returns:
            Path do arquivo JSON gerado.
        """
        TMP_DIR.mkdir(parents=True, exist_ok=True)
        dest = Path(dest_dir) if dest_dir else TMP_DIR

        if verbose:
            print("Buscando tarefas pendentes via painel...")

        tarefas = self.pendentes_painel(debug=debug)

        if verbose:
            print(f"  Painel retornou {len(tarefas)} tarefas.")

        if not tarefas:
            print(
                "\nAVISO: O painel retornou vazio.\n"
                "Possíveis causas:\n"
                "  1. O endpoint retorna formato não reconhecido — rode com debug=True.\n"
                "  2. Não há tarefas pendentes para este usuário.\n"
                "  3. Use baixar_com_filtro(where={...}) para buscar com filtro explícito.\n"
            )
            # Salva arquivo vazio para não quebrar o chamador
            return self._salvar([], dest, "tarefas_pendentes", verbose)

        return self._salvar(tarefas, dest, "tarefas_pendentes", verbose)

    def baixar_com_filtro(
        self,
        where: dict | str,
        nome_arquivo: str = "tarefas",
        dest_dir: Path | str | None = None,
        populate: list[str] | str | None = None,
        order: dict | str | None = None,
        page_size: int = 50,
        verbose: bool = True,
    ) -> Path:
        """
        Baixa TODAS as tarefas que atendem ao filtro `where`, com paginação
        automática, e salva em JSON.

        Exemplos de filtros:
            {"usuarioResponsavel.id": "eq:42"}
            {"dataHoraFinalPrazo": "lt:2026-03-10T23:59:59"}
            {"urgente": "eq:1"}

        Args:
            where:        Filtro obrigatório.
            nome_arquivo: Prefixo do arquivo gerado (sem extensão).
            dest_dir:     Pasta de destino. Padrão: <projeto>/tmp/
            populate:     Associações a popular.
            order:        Ordenação. Ex: {"dataHoraFinalPrazo": "ASC"}
            page_size:    Registros por requisição.
            verbose:      Imprime progresso.

        Returns:
            Path do arquivo JSON gerado.
        """
        TMP_DIR.mkdir(parents=True, exist_ok=True)
        dest = Path(dest_dir) if dest_dir else TMP_DIR

        if verbose:
            print(f"Buscando tarefas com filtro: {_where_str(where)}")

        tarefas = self.listar_todos(
            where=where,
            order=order,
            populate=populate,
            page_size=page_size,
            verbose=verbose,
        )

        return self._salvar(tarefas, dest, nome_arquivo, verbose)

    # ── escrita ───────────────────────────────────────────────────────────────

    def criar(self, dados: dict, context: str | None = None) -> dict:
        """POST /v1/administrativo/tarefa — Cria uma nova tarefa."""
        params = {"context": context} if context else {}
        resp = self._http.post(_BASE_PATH, json=dados, params=params)
        return _check(resp)

    def atualizar(self, tarefa_id: int | str, dados: dict, context: str | None = None) -> dict:
        """PUT /v1/administrativo/tarefa/{id} — Substitui a tarefa por completo."""
        params = {"context": context} if context else {}
        resp = self._http.put(f"{_BASE_PATH}/{tarefa_id}", json=dados, params=params)
        return _check(resp)

    def atualizar_parcial(self, tarefa_id: int | str, dados: dict, context: str | None = None) -> dict:
        """PATCH /v1/administrativo/tarefa/{id} — Atualiza campos específicos."""
        params = {"context": context} if context else {}
        resp = self._http.patch(f"{_BASE_PATH}/{tarefa_id}", json=dados, params=params)
        return _check(resp)

    # ── helpers internos ──────────────────────────────────────────────────────

    def _salvar(
        self,
        tarefas: list[dict],
        dest: Path,
        prefixo: str,
        verbose: bool,
    ) -> Path:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        arquivo = dest / f"{prefixo}_{timestamp}.json"
        with open(arquivo, "w", encoding="utf-8") as f:
            json.dump(tarefas, f, ensure_ascii=False, indent=2)
        if verbose:
            tamanho_kb = arquivo.stat().st_size / 1024
            print(f"\nSalvo em: {arquivo}")
            print(f"Total: {len(tarefas)} tarefas | Tamanho: {tamanho_kb:.1f} KB")
        return arquivo

    # ── context manager ───────────────────────────────────────────────────────

    def __enter__(self) -> "TarefaClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()
