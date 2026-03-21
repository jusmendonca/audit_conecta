"""
hermes/documento.py
Cliente para Documento e ComponenteDigital do SUPP.

Entidades:
  Documento        — metadados do documento (tipo, autor, processo, tarefa…)
  ComponenteDigital — arquivo binário vinculado a um Documento

Endpoints cobertos:
  --- Documento ---
  GET    /v1/administrativo/documento                         Lista paginada com filtros
  GET    /v1/administrativo/documento/count                   Contagem com filtros
  GET    /v1/administrativo/documento/{id}                    Busca por ID
  POST   /v1/administrativo/documento                         Cria documento
  PUT    /v1/administrativo/documento/{id}                    Atualiza (completo)
  PATCH  /v1/administrativo/documento/{id}                    Atualiza (parcial)
  DELETE /v1/administrativo/documento/{id}                    Remove
  PATCH  /v1/administrativo/documento/{id}/undelete           Restaura
  PATCH  /v1/administrativo/documento/convertToPdf/{id}       Converte para PDF
  GET    /v1/administrativo/documento/download/{id}/minuta_anexos/pdf
  POST   /v1/administrativo/documento/{id}/sendEmail          Envia por e-mail
  GET    /v1/administrativo/documento/{id}/visibilidade        Consulta visibilidade
  PUT    /v1/administrativo/documento/{id}/visibilidade        Cria direito de acesso
  DELETE /v1/administrativo/documento/{documentoId}/visibilidade/{id}

  --- ComponenteDigital ---
  GET    /v1/administrativo/componente_digital                 Lista paginada com filtros
  GET    /v1/administrativo/componente_digital/count           Contagem
  GET    /v1/administrativo/componente_digital/{id}            Busca por ID
  GET    /v1/administrativo/componente_digital/{id}/download   Download do arquivo binário
  GET    /v1/administrativo/componente_digital/{processoId}/download_latest
  GET    /v1/administrativo/componente_digital/{id}/download_p7s   Arquivo assinado (p7s)
  GET    /v1/administrativo/componente_digital/{id}/download_html  Versão HTML
  GET    /v1/administrativo/componente_digital/search          Busca Elasticsearch
  PATCH  /v1/administrativo/componente_digital/{id}/convertToPdf

Associações de Documento (populate):
  processoOrigem, tipoDocumento, tarefaOrigem, setorOrigem, juntadaAtual,
  numeroUnicoDocumento, procedencia, modelo, repositorio, criadoPor, atualizadoPor

Associações de ComponenteDigital (populate):
  documento, processoOrigem, tarefaOrigem, modelo, criadoPor, atualizadoPor
"""

from __future__ import annotations

import json
import mimetypes
from datetime import datetime
from pathlib import Path
from typing import Any

import httpx

from .config import BASE_URL

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
TMP_DIR = _PROJECT_ROOT / "tmp"

_DOC_PATH = "/v1/administrativo/documento"
_CD_PATH  = "/v1/administrativo/componente_digital"


class DocumentoError(Exception):
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
        raise DocumentoError(response.status_code, body)
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


# ═══════════════════════════════════════════════════════════════════════════════
# DocumentoClient
# ═══════════════════════════════════════════════════════════════════════════════

class DocumentoClient:
    """
    Cliente síncrono para Documento e ComponenteDigital do SUPP.

    Uso:
        from hermes.auth import AuthClient
        from hermes.documento import DocumentoClient

        auth = AuthClient()
        auth.login_ldap("cpf", "senha")
        dc = DocumentoClient.from_auth(auth)

        # Listar documentos de um processo
        docs = dc.listar_por_processo(processo_id=42)

        # Baixar todos os arquivos de um processo para disco
        dc.baixar_arquivos_processo(processo_id=42, dest_dir=Path("tmp/1001/documentos"))
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
    def from_auth(cls, auth_client: Any, timeout: float = 120.0) -> "DocumentoClient":
        if not auth_client.token:
            raise RuntimeError("AuthClient sem token. Faça login primeiro.")
        return cls(token=auth_client.token, base_url=auth_client.base_url, timeout=timeout)

    # ── Documento: consulta ───────────────────────────────────────────────────

    def buscar(
        self,
        documento_id: int | str,
        populate: list[str] | str | None = None,
        context: str | None = None,
    ) -> dict:
        """GET /documento/{id} — Metadados completos de um documento."""
        params: dict[str, Any] = {}
        if populate is not None:
            params["populate"] = _populate_str(populate)
        if context is not None:
            params["context"] = context
        resp = self._http.get(f"{_DOC_PATH}/{documento_id}", params=params)
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
        """GET /documento — Lista documentos com paginação e filtros."""
        params: dict[str, Any] = {"limit": limit, "offset": offset}
        if where is not None:
            params["where"] = _where_str(where)
        if order is not None:
            params["order"] = json.dumps(order) if isinstance(order, dict) else order
        if populate is not None:
            params["populate"] = _populate_str(populate)
        if context is not None:
            params["context"] = context
        resp = self._http.get(_DOC_PATH, params=params)
        return _extract_list(_check(resp))

    def contar(
        self,
        where: dict | str | None = None,
        context: str | None = None,
    ) -> int:
        """GET /documento/count — Contagem com filtro."""
        params: dict[str, Any] = {}
        if where is not None:
            params["where"] = _where_str(where)
        if context is not None:
            params["context"] = context
        data = _check(self._http.get(f"{_DOC_PATH}/count", params=params))
        if isinstance(data, dict):
            return int(data.get("count", data.get("total", 0)))
        return int(data)

    def listar_por_processo(
        self,
        processo_id: int | str,
        populate: list[str] | str | None = None,
        order: dict | str | None = None,
        page_size: int = 50,
        verbose: bool = False,
    ) -> list[dict]:
        """
        Lista todos os documentos de um processo, paginando automaticamente.

        Equivale a GET /documento?where={"processoOrigem.id":"eq:{id}"}
        """
        where = {"processoOrigem.id": f"eq:{processo_id}"}
        total = self.contar(where=where)
        if verbose:
            print(f"    {total} documento(s) no processo {processo_id}")

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

    def listar_por_tarefa(
        self,
        tarefa_id: int | str,
        populate: list[str] | str | None = None,
        page_size: int = 50,
    ) -> list[dict]:
        """Lista todos os documentos originados de uma tarefa."""
        where = {"tarefaOrigem.id": f"eq:{tarefa_id}"}
        total = self.contar(where=where)
        todos: list[dict] = []
        offset = 0
        while offset < total:
            pagina = self.listar(
                where=where, limit=page_size, offset=offset, populate=populate
            )
            if not pagina:
                break
            todos.extend(pagina)
            offset += len(pagina)
        return todos

    # ── Documento: ações ──────────────────────────────────────────────────────

    def converter_para_pdf(self, documento_id: int | str) -> Any:
        """PATCH /documento/convertToPdf/{id}"""
        resp = self._http.patch(f"{_DOC_PATH}/convertToPdf/{documento_id}")
        return _check(resp)

    def download_minuta_com_anexos(self, documento_id: int | str) -> bytes:
        """GET /documento/download/{id}/minuta_anexos/pdf — PDF com minuta + anexos."""
        resp = self._http.get(f"{_DOC_PATH}/download/{documento_id}/minuta_anexos/pdf")
        if not resp.is_success:
            raise DocumentoError(resp.status_code, resp.text)
        return resp.content

    def enviar_email(self, documento_id: int | str, dados: dict | None = None) -> Any:
        """POST /documento/{id}/sendEmail"""
        resp = self._http.post(f"{_DOC_PATH}/{documento_id}/sendEmail", json=dados or {})
        return _check(resp)

    def visibilidade(self, documento_id: int | str) -> Any:
        """GET /documento/{id}/visibilidade"""
        resp = self._http.get(f"{_DOC_PATH}/{documento_id}/visibilidade")
        return _check(resp)

    def criar_visibilidade(self, documento_id: int | str, dados: dict) -> Any:
        """PUT /documento/{id}/visibilidade"""
        resp = self._http.put(f"{_DOC_PATH}/{documento_id}/visibilidade", json=dados)
        return _check(resp)

    def remover_visibilidade(self, documento_id: int | str, visibilidade_id: int | str) -> Any:
        """DELETE /documento/{documentoId}/visibilidade/{id}"""
        resp = self._http.delete(f"{_DOC_PATH}/{documento_id}/visibilidade/{visibilidade_id}")
        return _check(resp)

    # ── Documento: escrita ────────────────────────────────────────────────────

    def criar(self, dados: dict, context: str | None = None) -> dict:
        """POST /documento"""
        params = {"context": context} if context else {}
        resp = self._http.post(_DOC_PATH, json=dados, params=params)
        return _check(resp)

    def atualizar(self, documento_id: int | str, dados: dict, context: str | None = None) -> dict:
        """PUT /documento/{id}"""
        params = {"context": context} if context else {}
        resp = self._http.put(f"{_DOC_PATH}/{documento_id}", json=dados, params=params)
        return _check(resp)

    def atualizar_parcial(self, documento_id: int | str, dados: dict, context: str | None = None) -> dict:
        """PATCH /documento/{id}"""
        params = {"context": context} if context else {}
        resp = self._http.patch(f"{_DOC_PATH}/{documento_id}", json=dados, params=params)
        return _check(resp)

    def deletar(self, documento_id: int | str) -> Any:
        """DELETE /documento/{id}"""
        resp = self._http.delete(f"{_DOC_PATH}/{documento_id}")
        return _check(resp)

    def restaurar(self, documento_id: int | str) -> Any:
        """PATCH /documento/{id}/undelete"""
        resp = self._http.patch(f"{_DOC_PATH}/{documento_id}/undelete")
        return _check(resp)

    # ── ComponenteDigital: consulta ───────────────────────────────────────────

    def buscar_componente(
        self,
        componente_id: int | str,
        populate: list[str] | str | None = None,
        context: str | None = None,
    ) -> dict:
        """GET /componente_digital/{id} — Metadados do componente digital."""
        params: dict[str, Any] = {}
        if populate is not None:
            params["populate"] = _populate_str(populate)
        if context is not None:
            params["context"] = context
        resp = self._http.get(f"{_CD_PATH}/{componente_id}", params=params)
        return _check(resp)

    def listar_componentes(
        self,
        where: dict | str | None = None,
        order: dict | str | None = None,
        limit: int = 25,
        offset: int = 0,
        populate: list[str] | str | None = None,
        context: str | None = None,
    ) -> list[dict]:
        """GET /componente_digital — Lista componentes com filtros."""
        params: dict[str, Any] = {"limit": limit, "offset": offset}
        if where is not None:
            params["where"] = _where_str(where)
        if order is not None:
            params["order"] = json.dumps(order) if isinstance(order, dict) else order
        if populate is not None:
            params["populate"] = _populate_str(populate)
        if context is not None:
            params["context"] = context
        resp = self._http.get(_CD_PATH, params=params)
        return _extract_list(_check(resp))

    def contar_componentes(
        self,
        where: dict | str | None = None,
        context: str | None = None,
    ) -> int:
        """GET /componente_digital/count"""
        params: dict[str, Any] = {}
        if where is not None:
            params["where"] = _where_str(where)
        if context is not None:
            params["context"] = context
        data = _check(self._http.get(f"{_CD_PATH}/count", params=params))
        if isinstance(data, dict):
            return int(data.get("count", data.get("total", 0)))
        return int(data)

    def pesquisar_componentes(
        self,
        where: dict | str | None = None,
        order: dict | str | None = None,
        limit: int = 25,
        offset: int = 0,
    ) -> list[dict]:
        """GET /componente_digital/search — Elasticsearch."""
        params: dict[str, Any] = {"limit": limit, "offset": offset}
        if where is not None:
            params["where"] = _where_str(where)
        if order is not None:
            params["order"] = json.dumps(order) if isinstance(order, dict) else order
        resp = self._http.get(f"{_CD_PATH}/search", params=params)
        return _extract_list(_check(resp))

    # ── ComponenteDigital: downloads ──────────────────────────────────────────

    def download_componente(
        self,
        componente_id: int | str,
        dest_dir: Path,
        nome_base: str | None = None,
    ) -> Path:
        """
        GET /componente_digital/{id}/download
        Baixa o arquivo binário do componente digital e salva em disco.

        Args:
            componente_id: ID do ComponenteDigital.
            dest_dir:      Pasta de destino (criada automaticamente).
            nome_base:     Nome base do arquivo; se None, usa o retornado pela API.

        Returns:
            Path do arquivo salvo.
        """
        dest_dir.mkdir(parents=True, exist_ok=True)
        resp = self._http.get(f"{_CD_PATH}/{componente_id}/download")
        if not resp.is_success:
            raise DocumentoError(resp.status_code, resp.text)
        fallback = nome_base or f"componente_{componente_id}"
        filename = _filename_from_response(resp, fallback)
        arquivo = dest_dir / filename
        arquivo.write_bytes(resp.content)
        return arquivo

    def download_latest(self, processo_id: int | str, dest_dir: Path) -> Path:
        """
        GET /componente_digital/{processoId}/download_latest
        Baixa a versão mais recente do componente do processo.
        """
        dest_dir.mkdir(parents=True, exist_ok=True)
        resp = self._http.get(f"{_CD_PATH}/{processo_id}/download_latest")
        if not resp.is_success:
            raise DocumentoError(resp.status_code, resp.text)
        filename = _filename_from_response(resp, f"processo_{processo_id}_latest")
        arquivo = dest_dir / filename
        arquivo.write_bytes(resp.content)
        return arquivo

    def download_p7s(self, componente_id: int | str, dest_dir: Path) -> Path:
        """GET /componente_digital/{id}/download_p7s — Arquivo com assinatura digital."""
        dest_dir.mkdir(parents=True, exist_ok=True)
        resp = self._http.get(f"{_CD_PATH}/{componente_id}/download_p7s")
        if not resp.is_success:
            raise DocumentoError(resp.status_code, resp.text)
        filename = _filename_from_response(resp, f"componente_{componente_id}")
        arquivo = dest_dir / filename
        arquivo.write_bytes(resp.content)
        return arquivo

    def download_html(self, componente_id: int | str, dest_dir: Path) -> Path:
        """GET /componente_digital/{id}/download_html — Versão HTML do componente."""
        dest_dir.mkdir(parents=True, exist_ok=True)
        resp = self._http.get(f"{_CD_PATH}/{componente_id}/download_html")
        if not resp.is_success:
            raise DocumentoError(resp.status_code, resp.text)
        filename = _filename_from_response(resp, f"componente_{componente_id}.html")
        arquivo = dest_dir / filename
        arquivo.write_bytes(resp.content)
        return arquivo

    # ── ComponenteDigital: escrita ────────────────────────────────────────────

    def converter_componente_para_pdf(self, componente_id: int | str) -> Any:
        """PATCH /componente_digital/{id}/convertToPdf"""
        resp = self._http.patch(f"{_CD_PATH}/{componente_id}/convertToPdf")
        return _check(resp)

    # ── Download em lote (usado pelo sync) ────────────────────────────────────

    def baixar_arquivos_processo(
        self,
        processo_id: int | str,
        dest_dir: Path,
        verbose: bool = True,
    ) -> list[Path]:
        """
        Baixa todos os documentos (+ seus componentes digitais) de um processo.

        Fluxo:
          1. Lista documentos do processo via GET /documento?where=processoOrigem.id=eq:{id}
          2. Para cada documento, salva metadados em {dest_dir}/{doc_id}_meta.json
          3. Obtém o ComponenteDigital via juntadaAtual ou busca direta
          4. Baixa o binário via GET /componente_digital/{cd_id}/download

        Args:
            processo_id: ID do processo.
            dest_dir:    Pasta de destino (ex: tmp/{tarefa_id}/documentos/).
            verbose:     Imprime progresso.

        Returns:
            Lista de Paths dos arquivos binários baixados.
        """
        dest_dir.mkdir(parents=True, exist_ok=True)
        documentos = self.listar_por_processo(processo_id, verbose=verbose)

        if not documentos:
            if verbose:
                print(f"    Nenhum documento encontrado para o processo {processo_id}.")
            return []

        if verbose:
            print(f"    Baixando {len(documentos)} documento(s)...")

        arquivos_baixados: list[Path] = []

        for doc in documentos:
            doc_id = doc.get("id")

            # Salva metadados do documento
            meta_path = dest_dir / f"{doc_id}_meta.json"
            meta_path.write_text(
                json.dumps(doc, ensure_ascii=False, indent=2), encoding="utf-8"
            )

            # Localiza o ComponenteDigital
            # Pode vir em juntadaAtual.componenteDigital ou componentesDigitais
            cd_id = _extrair_cd_id(doc)

            if cd_id is None:
                # Tenta buscar via listagem de componentes filtrada pelo documento
                componentes = self.listar_componentes(
                    where={"documento.id": f"eq:{doc_id}"},
                    limit=10,
                )
                if componentes:
                    cd_id = componentes[0].get("id")

            if cd_id is None:
                if verbose:
                    print(f"      Documento {doc_id} → sem ComponenteDigital localizado")
                continue

            # Baixa o binário
            try:
                nome_base = doc.get("fileName") or doc.get("nome") or str(doc_id)
                arquivo = self.download_componente(cd_id, dest_dir, nome_base=nome_base)
                arquivos_baixados.append(arquivo)
                if verbose:
                    print(f"      Documento {doc_id} → {arquivo.name}")
            except DocumentoError as e:
                if verbose:
                    print(f"      Documento {doc_id} → erro no download: {e}")

        return arquivos_baixados

    # ── context manager ───────────────────────────────────────────────────────

    def __enter__(self) -> "DocumentoClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()


# ── helpers internos ──────────────────────────────────────────────────────────

def _extrair_cd_id(doc: dict) -> int | str | None:
    """
    Extrai o ID do ComponenteDigital de um documento.
    Tenta vários campos que a API pode devolver dependendo do populate usado.
    """
    # Via juntadaAtual (quando populate=["juntadaAtual"] foi usado)
    juntada = doc.get("juntadaAtual")
    if isinstance(juntada, dict):
        cd = juntada.get("componenteDigital")
        if isinstance(cd, dict):
            return cd.get("id")
        if isinstance(cd, (int, str)):
            return cd

    # Via componenteDigital direto no documento
    cd = doc.get("componenteDigital")
    if isinstance(cd, dict):
        return cd.get("id")
    if isinstance(cd, (int, str)):
        return cd

    # Via lista componentesDigitais
    cds = doc.get("componentesDigitais")
    if isinstance(cds, list) and cds:
        first = cds[0]
        if isinstance(first, dict):
            return first.get("id")
        return first

    return None
