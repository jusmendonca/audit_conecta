"""
hermes/componente_digital.py
Cliente para os endpoints de ComponenteDigital do SUPP.

Endpoints cobertos:
  GET    /v1/administrativo/componente_digital                         Lista paginada
  GET    /v1/administrativo/componente_digital/count                   Contagem
  GET    /v1/administrativo/componente_digital/{id}                    Busca por ID
  POST   /v1/administrativo/componente_digital                         Cria componente
  PUT    /v1/administrativo/componente_digital/{id}                    Atualiza (completo)
  PATCH  /v1/administrativo/componente_digital/{id}                    Atualiza (parcial)
  DELETE /v1/administrativo/componente_digital/{id}                    Remove
  GET    /v1/administrativo/componente_digital/search                  Busca no Elasticsearch
  GET    /v1/administrativo/componente_digital/documento/{uuid}        Busca por UUID do documento
  GET    /v1/administrativo/componente_digital/{id}/download           Download do arquivo binário
  GET    /v1/administrativo/componente_digital/{processoId}/download_latest  Última versão do processo
  GET    /v1/administrativo/componente_digital/{id}/download_p7s       Download com assinatura (p7s)
  GET    /v1/administrativo/componente_digital/{id}/download_vinculado Download vinculado
  GET    /v1/administrativo/componente_digital/{id}/download_html      Download versão HTML
  PATCH  /v1/administrativo/componente_digital/{id}/convertToPdf       Converter para PDF
  PATCH  /v1/administrativo/componente_digital/{id}/convertToHtml      Converter para HTML
  PATCH  /v1/administrativo/componente_digital/{id}/reverter           Reverter para hash anterior
  PATCH  /v1/administrativo/componente_digital/{id}/undelete           Restaurar componente deletado
  POST   /v1/administrativo/componente_digital/aprovar                 Aprovar componente
  POST   /v1/administrativo/componente_digital/render_html_content     Renderizar conteúdo HTML
  POST   /v1/administrativo/componente_digital/{id}/compara_component_digital_com_html
                                                                       Comparar com HTML
  PUT    /v1/administrativo/componente_digital/bulk                    Atualização em lote

Schema:
  Obrigatório:
    fileName : str — Nome do arquivo (3–255 chars)

  Principais campos opcionais:
    hash                    : str  — Hash atual do conteúdo
    conteudo                : str  — Conteúdo em base64 (para upload inline)
    tamanho                 : int  — Tamanho em bytes
    extensao                : str  — Extensão do arquivo (ex: "pdf", "docx")
    mimetype                : str  — MIME type (ex: "application/pdf")
    numeracaoSequencial     : int  — Ordem dentro do documento
    nivelComposicao         : int  — Nível hierárquico de composição
    softwareCriacao         : str  — Software que gerou o arquivo
    versaoSoftwareCriacao   : str  — Versão do software
    dataHoraSoftwareCriacao : str  — Data/hora ISO 8601
    documento               : int  — ID do documento pai
    processoOrigem          : int  — ID do processo de origem
    documentoOrigem         : int  — ID do documento de origem
    tarefaOrigem            : int  — ID da tarefa de origem
    componenteDigitalOrigem : int  — ID do componente de origem (cópia)
    modelo                  : int  — ID do modelo utilizado
    geraModeloEmPdf         : int  — Gerar modelo como PDF (0/1)
    dadosFormulario         : int  — ID dos dados de formulário associados
    statusVerificacaoVirus  : int  — Status de verificação de antivírus

Populate (parâmetro `populate`):
  processoOrigem, tarefaOrigem, documentoAvulsoOrigem,
  modalidadeAlvoInibidor, modalidadeTipoInibidor, modelo,
  documento, criadoPor, atualizadoPor, apagadoPor, origemDados

Filtros (parâmetro `where`):
  Exemplos:
    {"documento.id": "eq:55"}
    {"extensao": "eq:pdf"}
    {"assinado": "eq:1"}
    {"convertidoPdf": "eq:1"}
"""

from __future__ import annotations

import mimetypes
import json
from pathlib import Path
from typing import Any

import httpx

from .config import BASE_URL

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
TMP_DIR = _PROJECT_ROOT / "tmp"

_BASE_PATH = "/v1/administrativo/componente_digital"


class ComponenteDigitalError(Exception):
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
        raise ComponenteDigitalError(response.status_code, body)
    try:
        return response.json()
    except Exception:
        return response.text


def _check_binary(response: httpx.Response) -> bytes:
    if not response.is_success:
        try:
            body = response.json()
        except Exception:
            body = response.text
        raise ComponenteDigitalError(response.status_code, body)
    return response.content


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


def _filename_from_response(resp: httpx.Response, fallback: str) -> str:
    """Infere o nome do arquivo pelo Content-Disposition ou Content-Type."""
    cd = resp.headers.get("content-disposition", "")
    if "filename=" in cd:
        name = cd.split("filename=")[-1].strip().strip('"').strip("'")
        if name:
            return name
    ct = resp.headers.get("content-type", "")
    ext = mimetypes.guess_extension(ct.split(";")[0].strip()) or ""
    return f"{fallback}{ext}"


class ComponenteDigitalClient:
    """
    Cliente síncrono para os endpoints de ComponenteDigital do SUPP.

    ComponenteDigital é o arquivo binário associado a um Documento — pode ser
    um PDF, DOCX, imagem, etc. Suporta versionamento (hash), assinatura digital
    (p7s), conversão de formato e busca full-text via Elasticsearch.

    Exemplo de uso:

        from hermes.auth import AuthClient
        from hermes.componente_digital import ComponenteDigitalClient

        auth = AuthClient()
        auth.login_ldap("cpf", "senha")
        cd = ComponenteDigitalClient.from_auth(auth)

        # Listar componentes de um documento
        componentes = cd.listar_por_documento(documento_id=55)

        # Baixar arquivo para disco
        caminho = cd.download(componente_id=123, dest_dir=Path("tmp/"))
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
    def from_auth(cls, auth_client: Any, timeout: float = 120.0) -> "ComponenteDigitalClient":
        """Cria ComponenteDigitalClient a partir de um AuthClient já autenticado."""
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
        GET /v1/administrativo/componente_digital
        Lista componentes digitais com paginação e filtros.

        Args:
            where:    Filtro JSON. Ex: {"documento.id": "eq:55"}
            order:    Ordenação. Ex: {"numeracaoSequencial": "ASC"}
            limit:    Máximo de registros por página (padrão 25).
            offset:   Início da paginação (padrão 0).
            populate: Associações a popular. Ex: ["documento", "modelo"]
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
        GET /v1/administrativo/componente_digital/count
        Retorna o total de componentes que atendem ao filtro.
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
        componente_id: int | str,
        populate: list[str] | str | None = None,
        context: str | None = None,
    ) -> dict:
        """
        GET /v1/administrativo/componente_digital/{id}
        Busca componente digital por ID.

        Populate disponível: processoOrigem, tarefaOrigem, documentoAvulsoOrigem,
        modalidadeAlvoInibidor, modalidadeTipoInibidor, modelo, documento,
        criadoPor, atualizadoPor, apagadoPor, origemDados
        """
        params: dict[str, Any] = {}
        if populate is not None:
            params["populate"] = _populate_str(populate)
        if context is not None:
            params["context"] = context

        resp = self._http.get(f"{_BASE_PATH}/{componente_id}", params=params)
        return _check(resp)

    def buscar_por_uuid_documento(self, uuid: str) -> dict:
        """
        GET /v1/administrativo/componente_digital/documento/{uuid}
        Busca componente digital pelo UUID do documento pai.

        Args:
            uuid: UUID do documento (formato: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx)
        """
        resp = self._http.get(f"{_BASE_PATH}/documento/{uuid}")
        return _check(resp)

    def pesquisar(
        self,
        where: dict | str | None = None,
        order: dict | str | None = None,
        limit: int = 25,
        offset: int = 0,
        populate: list[str] | str | None = None,
        context: str | None = None,
    ) -> list[dict]:
        """
        GET /v1/administrativo/componente_digital/search
        Busca componentes digitais no índice Elasticsearch (full-text search).

        Indicado para pesquisa por conteúdo de texto dentro dos arquivos.

        Args:
            where:    Filtro de pesquisa.
            order:    Ordenação (ex: {"score": "DESC"} para relevância).
            limit:    Máximo de resultados (padrão 25).
            offset:   Paginação (padrão 0).
            populate: Associações a popular.
            context:  Contexto da API.
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

        resp = self._http.get(f"{_BASE_PATH}/search", params=params)
        data = _check(resp)
        return _extract_list(data)

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
        Busca TODOS os componentes que atendem ao filtro, paginando automaticamente.

        O parâmetro `where` é fortemente recomendado para evitar timeouts.

        Args:
            where:     Filtro JSON. Ex: {"documento.id": "eq:55"}
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
                print(f"  Baixados {offset}/{total} componentes...")

        return todos

    # ── helpers de domínio ────────────────────────────────────────────────────

    def listar_por_documento(
        self,
        documento_id: int | str,
        populate: list[str] | str | None = None,
    ) -> list[dict]:
        """
        Retorna todos os componentes digitais de um documento, em ordem sequencial.

        Args:
            documento_id: ID do documento.
            populate:     Associações a popular (padrão: ["documento"]).
        """
        if populate is None:
            populate = ["documento"]
        return self.listar_todos(
            where={"documento.id": f"eq:{documento_id}"},
            order={"numeracaoSequencial": "ASC"},
            populate=populate,
            verbose=False,
        )

    # ── downloads ─────────────────────────────────────────────────────────────

    def download(
        self,
        componente_id: int | str,
        dest_dir: Path | str | None = None,
        nome_base: str | None = None,
    ) -> Path:
        """
        GET /v1/administrativo/componente_digital/{id}/download
        Baixa o arquivo binário e salva em disco.

        Args:
            componente_id: ID do componente digital.
            dest_dir:      Pasta de destino. Padrão: <projeto>/tmp/
            nome_base:     Nome base do arquivo (sem extensão). Padrão: "cd_{id}".

        Returns:
            Path do arquivo salvo.
        """
        dest = Path(dest_dir) if dest_dir else TMP_DIR
        dest.mkdir(parents=True, exist_ok=True)

        resp = self._http.get(f"{_BASE_PATH}/{componente_id}/download")
        conteudo = _check_binary(resp)

        fallback = nome_base or f"cd_{componente_id}"
        nome_arquivo = _filename_from_response(resp, fallback)
        arquivo = dest / nome_arquivo
        arquivo.write_bytes(conteudo)
        return arquivo

    def download_latest(
        self,
        processo_id: int | str,
        dest_dir: Path | str | None = None,
    ) -> Path:
        """
        GET /v1/administrativo/componente_digital/{processoId}/download_latest
        Baixa a última versão do componente digital de um processo.

        Args:
            processo_id: ID do processo.
            dest_dir:    Pasta de destino. Padrão: <projeto>/tmp/
        """
        dest = Path(dest_dir) if dest_dir else TMP_DIR
        dest.mkdir(parents=True, exist_ok=True)

        resp = self._http.get(f"{_BASE_PATH}/{processo_id}/download_latest")
        conteudo = _check_binary(resp)

        nome_arquivo = _filename_from_response(resp, f"cd_latest_processo_{processo_id}")
        arquivo = dest / nome_arquivo
        arquivo.write_bytes(conteudo)
        return arquivo

    def download_p7s(
        self,
        componente_id: int | str,
        dest_dir: Path | str | None = None,
    ) -> Path:
        """
        GET /v1/administrativo/componente_digital/{id}/download_p7s
        Baixa o arquivo com a assinatura digital embutida (formato p7s/CAdES).

        Args:
            componente_id: ID do componente digital.
            dest_dir:      Pasta de destino. Padrão: <projeto>/tmp/
        """
        dest = Path(dest_dir) if dest_dir else TMP_DIR
        dest.mkdir(parents=True, exist_ok=True)

        resp = self._http.get(f"{_BASE_PATH}/{componente_id}/download_p7s")
        conteudo = _check_binary(resp)

        nome_arquivo = _filename_from_response(resp, f"cd_{componente_id}_assinado.p7s")
        arquivo = dest / nome_arquivo
        arquivo.write_bytes(conteudo)
        return arquivo

    def download_vinculado(
        self,
        componente_id: int | str,
        dest_dir: Path | str | None = None,
    ) -> Path:
        """
        GET /v1/administrativo/componente_digital/{id}/download_vinculado
        Baixa o componente digital vinculado.

        Args:
            componente_id: ID do componente digital.
            dest_dir:      Pasta de destino. Padrão: <projeto>/tmp/
        """
        dest = Path(dest_dir) if dest_dir else TMP_DIR
        dest.mkdir(parents=True, exist_ok=True)

        resp = self._http.get(f"{_BASE_PATH}/{componente_id}/download_vinculado")
        conteudo = _check_binary(resp)

        nome_arquivo = _filename_from_response(resp, f"cd_{componente_id}_vinculado")
        arquivo = dest / nome_arquivo
        arquivo.write_bytes(conteudo)
        return arquivo

    def download_html(
        self,
        componente_id: int | str,
        dest_dir: Path | str | None = None,
    ) -> Path:
        """
        GET /v1/administrativo/componente_digital/{id}/download_html
        Baixa a versão HTML do componente digital.

        Args:
            componente_id: ID do componente digital.
            dest_dir:      Pasta de destino. Padrão: <projeto>/tmp/
        """
        dest = Path(dest_dir) if dest_dir else TMP_DIR
        dest.mkdir(parents=True, exist_ok=True)

        resp = self._http.get(f"{_BASE_PATH}/{componente_id}/download_html")
        conteudo = _check_binary(resp)

        nome_arquivo = _filename_from_response(resp, f"cd_{componente_id}.html")
        arquivo = dest / nome_arquivo
        arquivo.write_bytes(conteudo)
        return arquivo

    # ── escrita ───────────────────────────────────────────────────────────────

    def criar(self, dados: dict, context: str | None = None) -> dict:
        """
        POST /v1/administrativo/componente_digital — Cria um novo componente digital.

        Campo obrigatório em `dados`:
            fileName : str — Nome do arquivo (ex: "contrato.pdf")

        Campos comuns:
            hash        : str — Hash SHA-256 do arquivo (para upload externo)
            conteudo    : str — Conteúdo em base64 (para upload inline)
            tamanho     : int — Tamanho em bytes
            extensao    : str — Extensão (ex: "pdf")
            mimetype    : str — MIME type (ex: "application/pdf")
            documento   : int — ID do documento pai
            modelo      : int — ID do modelo (para gerar minutas)
            geraModeloEmPdf : int — 1 para gerar PDF a partir do modelo

        Exemplo:
            cd.criar({
                "fileName": "parecer.pdf",
                "extensao": "pdf",
                "mimetype": "application/pdf",
                "documento": 55,
                "hash": "abc123...",
                "tamanho": 204800,
            })
        """
        params = {"context": context} if context else {}
        resp = self._http.post(_BASE_PATH, json=dados, params=params)
        return _check(resp)

    def atualizar(self, componente_id: int | str, dados: dict, context: str | None = None) -> dict:
        """PUT /v1/administrativo/componente_digital/{id} — Substitui o componente por completo."""
        params = {"context": context} if context else {}
        resp = self._http.put(f"{_BASE_PATH}/{componente_id}", json=dados, params=params)
        return _check(resp)

    def atualizar_parcial(self, componente_id: int | str, dados: dict, context: str | None = None) -> dict:
        """PATCH /v1/administrativo/componente_digital/{id} — Atualiza campos específicos."""
        params = {"context": context} if context else {}
        resp = self._http.patch(f"{_BASE_PATH}/{componente_id}", json=dados, params=params)
        return _check(resp)

    def deletar(self, componente_id: int | str) -> Any:
        """DELETE /v1/administrativo/componente_digital/{id} — Remove o componente."""
        resp = self._http.delete(f"{_BASE_PATH}/{componente_id}")
        return _check(resp)

    def atualizar_lote(self, dados: list[dict]) -> Any:
        """
        PUT /v1/administrativo/componente_digital/bulk
        Atualiza múltiplos componentes digitais em uma única requisição.

        Args:
            dados: Lista de dicts com os dados de cada componente a atualizar.
                   Cada item deve conter o `id` do componente e os campos a alterar.
        """
        resp = self._http.put(f"{_BASE_PATH}/bulk", json=dados)
        return _check(resp)

    # ── ações ─────────────────────────────────────────────────────────────────

    def reverter(self, componente_id: int | str) -> dict:
        """
        PATCH /v1/administrativo/componente_digital/{id}/reverter
        Reverte o componente digital para o hash anterior (desfaz a última versão).

        Args:
            componente_id: ID do componente a reverter.
        """
        resp = self._http.patch(f"{_BASE_PATH}/{componente_id}/reverter")
        return _check(resp)

    def restaurar(self, componente_id: int | str) -> dict:
        """
        PATCH /v1/administrativo/componente_digital/{id}/undelete
        Restaura um componente digital que foi deletado (soft delete).

        Args:
            componente_id: ID do componente a restaurar.
        """
        resp = self._http.patch(f"{_BASE_PATH}/{componente_id}/undelete")
        return _check(resp)

    def converter_para_pdf(self, componente_id: int | str) -> dict:
        """
        PATCH /v1/administrativo/componente_digital/{id}/convertToPdf
        Solicita a conversão do componente para PDF.

        Funciona com componentes em formato editável (ex: DOCX, ODT).
        Componentes assinados não podem ser convertidos — remova as assinaturas antes.

        Args:
            componente_id: ID do componente a converter.
        """
        resp = self._http.patch(f"{_BASE_PATH}/{componente_id}/convertToPdf")
        return _check(resp)

    def converter_para_html(self, componente_id: int | str) -> dict:
        """
        PATCH /v1/administrativo/componente_digital/{id}/convertToHtml
        Solicita a conversão do componente para HTML (para edição online).

        Args:
            componente_id: ID do componente a converter.
        """
        resp = self._http.patch(f"{_BASE_PATH}/{componente_id}/convertToHtml")
        return _check(resp)

    def aprovar(self, dados: dict) -> dict:
        """
        POST /v1/administrativo/componente_digital/aprovar
        Aprova um pedido de componente digital (fluxo de aprovação de minutas).

        Args:
            dados: Payload de aprovação (campos dependem do fluxo configurado).
        """
        resp = self._http.post(f"{_BASE_PATH}/aprovar", json=dados)
        return _check(resp)

    def renderizar_html(self, dados: dict) -> bytes:
        """
        POST /v1/administrativo/componente_digital/render_html_content
        Renderiza conteúdo HTML e retorna o binário resultante (geralmente PDF).

        Args:
            dados: Payload com o conteúdo HTML a renderizar.

        Returns:
            Bytes do conteúdo renderizado.
        """
        resp = self._http.post(f"{_BASE_PATH}/render_html_content", json=dados)
        return _check_binary(resp)

    def comparar_com_html(self, componente_id: int | str, dados: dict) -> Any:
        """
        POST /v1/administrativo/componente_digital/{id}/compara_component_digital_com_html
        Compara o conteúdo atual do componente com um conteúdo HTML fornecido.
        Útil para detectar diferenças entre a versão salva e a em edição.

        Args:
            componente_id: ID do componente digital de referência.
            dados:         Payload com o conteúdo HTML para comparação.
        """
        resp = self._http.post(
            f"{_BASE_PATH}/{componente_id}/compara_component_digital_com_html",
            json=dados,
        )
        return _check(resp)

    # ── context manager ───────────────────────────────────────────────────────

    def __enter__(self) -> "ComponenteDigitalClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()
