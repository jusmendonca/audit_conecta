"""
hermes/processo.py
Cliente para os endpoints de Processo do SUPP.

Endpoints cobertos:
  GET    /v1/administrativo/processo                                       Lista paginada com filtros
  GET    /v1/administrativo/processo/count                                 Contagem com filtros
  GET    /v1/administrativo/processo/search                                Busca via Elasticsearch
  GET    /v1/administrativo/processo/{id}                                  Busca por ID
  GET    /v1/administrativo/processo/nup/{nup}                             Busca por NUP
  POST   /v1/administrativo/processo                                       Cria processo
  PUT    /v1/administrativo/processo/{id}                                  Atualiza processo (completo)
  PATCH  /v1/administrativo/processo/{id}                                  Atualiza processo (parcial)
  DELETE /v1/administrativo/processo/{id}                                  Remove processo
  PATCH  /v1/administrativo/processo/{id}/arquivar                         Arquiva processo
  PATCH  /v1/administrativo/processo/{id}/autuar                           Autua processo
  GET    /v1/administrativo/processo/{id}/timeline                         Timeline do processo
  GET    /v1/administrativo/processo/{id}/juntada_index                    Índice de juntadas
  GET    /v1/administrativo/processo/{id}/visibilidade                     Consulta visibilidade
  PUT    /v1/administrativo/processo/{id}/visibilidade                     Cria direito de acesso
  DELETE /v1/administrativo/processo/{processoId}/visibilidade/{id}        Remove direito de acesso
  DELETE /v1/administrativo/processo/{processoId}/deletevisibilidadedocs   Remove acesso aos docs
  GET    /v1/administrativo/processo/{id}/download/{tipo}/{sequencial}     Download de arquivo
  GET    /v1/administrativo/processo/imprime_etiqueta/{processoId}         Imprime etiqueta
  GET    /v1/administrativo/processo/imprime_relatorio/{processoId}        Imprime relatório
  POST   /v1/administrativo/processo/{id}/sendEmail                        Envia por e-mail
  PATCH  /v1/administrativo/processo/{id}/sincronizar_processo_judicial    Sincroniza judicial
  PATCH  /v1/administrativo/processo/{id}/converter_consultivo_em_administrativo
  PATCH  /v1/administrativo/processo/{id}/converter_administrativo_em_consultivo
  PATCH  /v1/administrativo/processo/{id}/converter_disciplinar_em_administrativo
  PATCH  /v1/administrativo/processo/{id}/converter_judicial_em_administrativo

Associações disponíveis no populate:
  processoOrigem, especieProcesso, modalidadeMeio, modalidadeFase,
  documentoAvulsoOrigem, classificacao, procedencia, localizador,
  setorAtual, setorInicial, processoPadLegado, processoJuizoLegado,
  configuracaoNup, criadoPor, atualizadoPor, origemDados
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

_BASE_PATH = "/v1/administrativo/processo"


class ProcessoError(Exception):
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
        raise ProcessoError(response.status_code, body)
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
        for key in ("entities", "data", "results", "items", "processos"):
            if key in data and isinstance(data[key], list):
                return data[key]
    return []


class ProcessoClient:
    """
    Cliente síncrono para os endpoints de Processo do SUPP.

    Uso:
        from hermes.auth import AuthClient
        from hermes.processo import ProcessoClient

        auth = AuthClient()
        auth.login_ldap("cpf", "senha")
        processos = ProcessoClient.from_auth(auth)
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
    def from_auth(cls, auth_client: Any, timeout: float = 120.0) -> "ProcessoClient":
        """Cria ProcessoClient a partir de um AuthClient já autenticado."""
        if not auth_client.token:
            raise RuntimeError("AuthClient sem token. Faça login primeiro.")
        return cls(token=auth_client.token, base_url=auth_client.base_url, timeout=timeout)

    # ── consulta ──────────────────────────────────────────────────────────────

    def buscar(
        self,
        processo_id: int | str,
        populate: list[str] | str | None = None,
        context: str | None = None,
    ) -> dict:
        """
        GET /v1/administrativo/processo/{id}
        Busca processo por ID interno.

        populate sugerido para detalhes completos:
            ["especieProcesso", "setorAtual", "classificacao",
             "localizador", "procedencia", "criadoPor"]
        """
        params: dict[str, Any] = {}
        if populate is not None:
            params["populate"] = _populate_str(populate)
        if context is not None:
            params["context"] = context

        resp = self._http.get(f"{_BASE_PATH}/{processo_id}", params=params)
        return _check(resp)

    def buscar_por_nup(self, nup: str) -> dict:
        """
        GET /v1/administrativo/processo/nup/{nup}
        Busca processo pelo NUP (Número Único de Protocolo).

        O NUP deve ser passado exatamente como registrado, ex:
            "00000.000001/2026-01"
        """
        resp = self._http.get(f"{_BASE_PATH}/nup/{nup}")
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
        """
        GET /v1/administrativo/processo
        Lista processos com paginação e filtros.

        Args:
            where:    Filtro JSON. Ex: {"setorAtual.id": "eq:10"}
            order:    Ordenação. Ex: {"dataHoraAbertura": "DESC"}
            limit:    Máximo de registros por página (padrão 25).
            offset:   Início da paginação (padrão 0).
            populate: Associações a incluir na resposta.
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
        GET /v1/administrativo/processo/count
        Retorna o total de processos que atendem ao filtro.
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

    def pesquisar(
        self,
        where: dict | str | None = None,
        order: dict | str | None = None,
        limit: int = 25,
        offset: int = 0,
    ) -> list[dict]:
        """
        GET /v1/administrativo/processo/search
        Busca processos via Elasticsearch (busca textual avançada).
        Útil para buscas por termos livres no título ou descrição.
        """
        params: dict[str, Any] = {"limit": limit, "offset": offset}
        if where is not None:
            params["where"] = _where_str(where)
        if order is not None:
            params["order"] = json.dumps(order) if isinstance(order, dict) else order

        resp = self._http.get(f"{_BASE_PATH}/search", params=params)
        data = _check(resp)
        return _extract_list(data)

    # ── paginação automática ──────────────────────────────────────────────────

    def listar_todos(
        self,
        where: dict | str,
        order: dict | str | None = None,
        populate: list[str] | str | None = None,
        context: str | None = None,
        page_size: int = 50,
        verbose: bool = True,
    ) -> list[dict]:
        """
        Busca TODOS os processos que atendem ao filtro, paginando automaticamente.

        O parâmetro `where` é obrigatório — queries sem filtro podem
        causar timeout em bases grandes.
        """
        total = self.contar(where=where, context=context)
        if verbose:
            print(f"  Total encontrado: {total} processos")

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
                print(f"  Baixados {offset}/{total} processos...")

        return todos

    # ── informações complementares ────────────────────────────────────────────

    def timeline(self, processo_id: int | str) -> list[dict]:
        """
        GET /v1/administrativo/processo/{id}/timeline
        Retorna a linha do tempo de eventos do processo.
        """
        resp = self._http.get(f"{_BASE_PATH}/{processo_id}/timeline")
        data = _check(resp)
        return data if isinstance(data, list) else _extract_list(data)

    def juntada_index(self, processo_id: int | str) -> Any:
        """
        GET /v1/administrativo/processo/{id}/juntada_index
        Retorna o índice de juntadas do processo.
        """
        resp = self._http.get(f"{_BASE_PATH}/{processo_id}/juntada_index")
        return _check(resp)

    # ── visibilidade / acesso ─────────────────────────────────────────────────

    def visibilidade(self, processo_id: int | str) -> Any:
        """GET /v1/administrativo/processo/{id}/visibilidade — Consulta visibilidade."""
        resp = self._http.get(f"{_BASE_PATH}/{processo_id}/visibilidade")
        return _check(resp)

    def criar_visibilidade(self, processo_id: int | str, dados: dict) -> Any:
        """PUT /v1/administrativo/processo/{id}/visibilidade — Cria direito de acesso."""
        resp = self._http.put(f"{_BASE_PATH}/{processo_id}/visibilidade", json=dados)
        return _check(resp)

    def remover_visibilidade(self, processo_id: int | str, visibilidade_id: int | str) -> Any:
        """DELETE /v1/administrativo/processo/{processoId}/visibilidade/{id}."""
        resp = self._http.delete(f"{_BASE_PATH}/{processo_id}/visibilidade/{visibilidade_id}")
        return _check(resp)

    def remover_visibilidade_docs(self, processo_id: int | str) -> Any:
        """DELETE .../deletevisibilidadedocs — Remove acesso a todos os docs do processo."""
        resp = self._http.delete(f"{_BASE_PATH}/{processo_id}/deletevisibilidadedocs")
        return _check(resp)

    # ── download e impressão ──────────────────────────────────────────────────

    def download(
        self,
        processo_id: int | str,
        tipo: str,
        sequencial: str | int,
        dest_dir: Path | str | None = None,
    ) -> Path:
        """
        GET /v1/administrativo/processo/{id}/download/{tipo}/{sequencial}
        Baixa um arquivo do processo e salva em disco.

        Args:
            tipo:       Tipo do arquivo (ex: "PDF", "ORIGINAL").
            sequencial: Sequencial do componente digital.
            dest_dir:   Pasta de destino. Padrão: <projeto>/tmp/
        """
        TMP_DIR.mkdir(parents=True, exist_ok=True)
        dest = Path(dest_dir) if dest_dir else TMP_DIR

        resp = self._http.get(f"{_BASE_PATH}/{processo_id}/download/{tipo}/{sequencial}")
        if not resp.is_success:
            raise ProcessoError(resp.status_code, resp.text)

        # Tenta inferir nome do arquivo pelo Content-Disposition
        cd = resp.headers.get("content-disposition", "")
        filename = None
        if "filename=" in cd:
            filename = cd.split("filename=")[-1].strip().strip('"')
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"processo_{processo_id}_{tipo}_{sequencial}_{timestamp}.bin"

        arquivo = dest / filename
        arquivo.write_bytes(resp.content)
        return arquivo

    def imprimir_etiqueta(self, processo_id: int | str) -> bytes:
        """GET .../imprime_etiqueta/{processoId} — Retorna o conteúdo binário da etiqueta."""
        resp = self._http.get(f"{_BASE_PATH}/imprime_etiqueta/{processo_id}")
        if not resp.is_success:
            raise ProcessoError(resp.status_code, resp.text)
        return resp.content

    def imprimir_relatorio(self, processo_id: int | str) -> bytes:
        """GET .../imprime_relatorio/{processoId} — Retorna o relatório de documentos."""
        resp = self._http.get(f"{_BASE_PATH}/imprime_relatorio/{processo_id}")
        if not resp.is_success:
            raise ProcessoError(resp.status_code, resp.text)
        return resp.content

    def enviar_email(self, processo_id: int | str, dados: dict | None = None) -> Any:
        """POST /v1/administrativo/processo/{id}/sendEmail — Envia processo por e-mail."""
        resp = self._http.post(f"{_BASE_PATH}/{processo_id}/sendEmail", json=dados or {})
        return _check(resp)

    # ── ações de workflow ─────────────────────────────────────────────────────

    def arquivar(self, processo_id: int | str, dados: dict | None = None) -> Any:
        """PATCH .../arquivar — Arquiva o processo."""
        resp = self._http.patch(f"{_BASE_PATH}/{processo_id}/arquivar", json=dados or {})
        return _check(resp)

    def autuar(self, processo_id: int | str, dados: dict | None = None) -> Any:
        """PATCH .../autuar — Autua o processo."""
        resp = self._http.patch(f"{_BASE_PATH}/{processo_id}/autuar", json=dados or {})
        return _check(resp)

    def sincronizar_judicial(self, processo_id: int | str) -> Any:
        """PATCH .../sincronizar_processo_judicial — Sincroniza com o processo judicial."""
        resp = self._http.patch(f"{_BASE_PATH}/{processo_id}/sincronizar_processo_judicial")
        return _check(resp)

    def converter_consultivo_em_administrativo(self, processo_id: int | str) -> Any:
        """PATCH .../converter_consultivo_em_administrativo"""
        resp = self._http.patch(
            f"{_BASE_PATH}/{processo_id}/converter_consultivo_em_administrativo"
        )
        return _check(resp)

    def converter_administrativo_em_consultivo(self, processo_id: int | str) -> Any:
        """PATCH .../converter_administrativo_em_consultivo"""
        resp = self._http.patch(
            f"{_BASE_PATH}/{processo_id}/converter_administrativo_em_consultivo"
        )
        return _check(resp)

    def converter_disciplinar_em_administrativo(self, processo_id: int | str) -> Any:
        """PATCH .../converter_disciplinar_em_administrativo"""
        resp = self._http.patch(
            f"{_BASE_PATH}/{processo_id}/converter_disciplinar_em_administrativo"
        )
        return _check(resp)

    def converter_judicial_em_administrativo(self, processo_id: int | str) -> Any:
        """PATCH .../converter_judicial_em_administrativo"""
        resp = self._http.patch(
            f"{_BASE_PATH}/{processo_id}/converter_judicial_em_administrativo"
        )
        return _check(resp)

    # ── escrita ───────────────────────────────────────────────────────────────

    def criar(self, dados: dict, context: str | None = None) -> dict:
        """
        POST /v1/administrativo/processo — Abre um novo processo.

        Campos obrigatórios típicos:
            unidadeArquivistica, tipoProtocolo, especieProcesso,
            setorAtual, setorInicial, NUP (ou configuracaoNup)
        """
        params = {"context": context} if context else {}
        resp = self._http.post(_BASE_PATH, json=dados, params=params)
        return _check(resp)

    def atualizar(
        self, processo_id: int | str, dados: dict, context: str | None = None
    ) -> dict:
        """PUT /v1/administrativo/processo/{id} — Substitui o processo por completo."""
        params = {"context": context} if context else {}
        resp = self._http.put(f"{_BASE_PATH}/{processo_id}", json=dados, params=params)
        return _check(resp)

    def atualizar_parcial(
        self, processo_id: int | str, dados: dict, context: str | None = None
    ) -> dict:
        """PATCH /v1/administrativo/processo/{id} — Atualiza campos específicos."""
        params = {"context": context} if context else {}
        resp = self._http.patch(f"{_BASE_PATH}/{processo_id}", json=dados, params=params)
        return _check(resp)

    def deletar(self, processo_id: int | str) -> Any:
        """DELETE /v1/administrativo/processo/{id} — Remove o processo."""
        resp = self._http.delete(f"{_BASE_PATH}/{processo_id}")
        return _check(resp)

    # ── download em lote para disco ───────────────────────────────────────────

    def baixar_com_filtro(
        self,
        where: dict | str,
        nome_arquivo: str = "processos",
        dest_dir: Path | str | None = None,
        populate: list[str] | str | None = None,
        order: dict | str | None = None,
        page_size: int = 50,
        verbose: bool = True,
    ) -> Path:
        """
        Baixa TODOS os processos que atendem ao filtro e salva em JSON.

        Args:
            where:        Filtro obrigatório. Ex: {"setorAtual.id": "eq:10"}
            nome_arquivo: Prefixo do arquivo gerado.
            dest_dir:     Pasta de destino. Padrão: <projeto>/tmp/
            populate:     Associações a popular.
            order:        Ordenação.
            page_size:    Registros por requisição.
            verbose:      Imprime progresso.

        Returns:
            Path do arquivo JSON gerado.
        """
        TMP_DIR.mkdir(parents=True, exist_ok=True)
        dest = Path(dest_dir) if dest_dir else TMP_DIR

        if verbose:
            print(f"Buscando processos com filtro: {_where_str(where)}")

        processos = self.listar_todos(
            where=where,
            order=order,
            populate=populate,
            page_size=page_size,
            verbose=verbose,
        )

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        arquivo = dest / f"{nome_arquivo}_{timestamp}.json"
        with open(arquivo, "w", encoding="utf-8") as f:
            json.dump(processos, f, ensure_ascii=False, indent=2)

        if verbose:
            kb = arquivo.stat().st_size / 1024
            print(f"\nSalvo em: {arquivo}")
            print(f"Total: {len(processos)} processos | Tamanho: {kb:.1f} KB")

        return arquivo

    # ── context manager ───────────────────────────────────────────────────────

    def __enter__(self) -> "ProcessoClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()
