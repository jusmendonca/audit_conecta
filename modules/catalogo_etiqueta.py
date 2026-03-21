"""
hermes/catalogo_etiqueta.py
Cliente para o catálogo de Etiquetas do SUPP.

Diferença em relação a hermes/etiqueta.py:
  - etiqueta.py  → VinculacaoEtiqueta  (vinculação etiqueta↔tarefa/processo/doc)
  - este módulo  → Etiqueta            (catálogo das etiquetas disponíveis)

PROBLEMA de design da API:
  O endpoint GET /etiqueta NÃO permite filtrar por setor nem por usuario porque
  esses campos são armazenados como EntityInterface serializado (sem mapeamento
  Doctrine). Filtrar "setor.id" retorna HTTP 400.

ESTRATÉGIA ADOTADA — via VinculacaoEtiqueta:
  VinculacaoEtiqueta.setor e VinculacaoEtiqueta.usuario são relações Doctrine
  reais (aparecem no populate), portanto SÃO filtráveis.

  Para obter as etiquetas de um setor:
    GET /vinculacao_etiqueta
        ?where={"setor.id":"eq:{setor_id}"}
        &populate=["etiqueta"]
        &order={"criadoEm":"DESC"}
        &limit=500   ← janela deslizante; após deduplicação: ~20-100 etiquetas únicas

  Para obter as etiquetas pessoais de um usuário:
    GET /vinculacao_etiqueta
        ?where={"usuario.id":"eq:{usuario_id}"}
        &populate=["etiqueta"]
        &order={"criadoEm":"DESC"}
        &limit=200

  As etiquetas únicas extraídas dessas consultas representam o catálogo
  efetivamente utilizado pelo setor/usuário — ou seja, as etiquetas "ativas"
  na prática.
"""

from __future__ import annotations

import json
from typing import Any

import httpx

from .config import BASE_URL

_PATH_VE = "/v1/administrativo/vinculacao_etiqueta"

# Quantos registros de VinculacaoEtiqueta buscar para extrair etiquetas únicas
_LIMIT_SETOR   = 500
_LIMIT_PESSOAL = 200


class CatalogoEtiquetaError(Exception):
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
        raise CatalogoEtiquetaError(response.status_code, body)
    try:
        return response.json()
    except Exception:
        return response.text


def _extract_list(data: Any) -> list[dict]:
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("entities", "data", "results", "items"):
            if key in data and isinstance(data[key], list):
                return data[key]
    return []


def _etiquetas_unicas_de_vinculos(vinculos: list[dict]) -> list[dict]:
    """Extrai etiquetas únicas (por id) de uma lista de VinculacaoEtiqueta."""
    vistos: set = set()
    resultado: list[dict] = []
    for v in vinculos:
        etiqueta = v.get("etiqueta")
        if not isinstance(etiqueta, dict):
            continue
        eid = etiqueta.get("id")
        if eid and eid not in vistos:
            vistos.add(eid)
            resultado.append(etiqueta)
    return resultado


class CatalogoEtiquetaClient:
    """
    Cliente síncrono para o catálogo de Etiquetas do SUPP.

    Usa VinculacaoEtiqueta como proxy de catálogo, porque o endpoint /etiqueta
    não suporta filtro por setor ou usuario.

    Uso:
        cec = CatalogoEtiquetaClient.from_auth(auth)
        etiquetas = cec.listar_disponiveis(setor_id=42, usuario_id=99)
    """

    def __init__(
        self,
        token: str,
        base_url: str = BASE_URL,
        timeout: float = 60.0,
    ) -> None:
        self.token = token
        self.base_url = base_url
        self._http = httpx.Client(
            base_url=base_url,
            timeout=timeout,
            headers={"Authorization": f"Bearer {token}"},
        )

    @classmethod
    def from_auth(cls, auth_client: Any, timeout: float = 60.0) -> "CatalogoEtiquetaClient":
        if not auth_client.token:
            raise RuntimeError("AuthClient sem token. Faça login primeiro.")
        return cls(token=auth_client.token, base_url=auth_client.base_url, timeout=timeout)

    # ── consulta via VinculacaoEtiqueta ───────────────────────────────────────

    def _buscar_vinculos(self, where: dict, limit: int) -> list[dict]:
        """Busca VinculacaoEtiqueta com populate=["etiqueta"], mais recentes primeiro."""
        params = {
            "where":    json.dumps(where, ensure_ascii=False),
            "populate": json.dumps(["etiqueta"]),
            "order":    json.dumps({"criadoEm": "DESC"}),
            "limit":    limit,
            "offset":   0,
        }
        resp = self._http.get(_PATH_VE, params=params)
        return _extract_list(_check(resp))

    # ── helpers por escopo ────────────────────────────────────────────────────

    def listar_por_setor(self, setor_id: int | str) -> list[dict]:
        """
        Etiquetas usadas pelo setor — extraídas das VinculacaoEtiqueta mais recentes.

        Busca as {_LIMIT_SETOR} vinculações mais recentes com setor.id == setor_id
        e retorna as etiquetas únicas encontradas.
        """
        vinculos = self._buscar_vinculos(
            where={"setor.id": f"eq:{setor_id}"},
            limit=_LIMIT_SETOR,
        )
        return _etiquetas_unicas_de_vinculos(vinculos)

    def listar_pessoais(self, usuario_id: int | str) -> list[dict]:
        """
        Etiquetas usadas pelo usuário — extraídas das VinculacaoEtiqueta mais recentes.

        Busca as {_LIMIT_PESSOAL} vinculações mais recentes com usuario.id == usuario_id
        e retorna as etiquetas únicas encontradas.
        """
        vinculos = self._buscar_vinculos(
            where={"usuario.id": f"eq:{usuario_id}"},
            limit=_LIMIT_PESSOAL,
        )
        return _etiquetas_unicas_de_vinculos(vinculos)

    def listar_disponiveis(
        self,
        setor_id: int | str | None = None,
        usuario_id: int | str | None = None,
    ) -> list[dict]:
        """
        Combina etiquetas do setor + etiquetas pessoais do usuário, sem duplicatas.

        Ordena: setor primeiro, pessoais depois.
        Marca campo _origem em cada item ("setor" ou "pessoal").
        """
        catalogo: list[dict] = []
        vistos: set = set()

        if setor_id is not None:
            for e in self.listar_por_setor(setor_id):
                eid = e.get("id")
                if eid not in vistos:
                    vistos.add(eid)
                    e["_origem"] = "setor"
                    catalogo.append(e)

        if usuario_id is not None:
            for e in self.listar_pessoais(usuario_id):
                eid = e.get("id")
                if eid not in vistos:
                    vistos.add(eid)
                    e["_origem"] = "pessoal"
                    catalogo.append(e)

        return catalogo

    # ── context manager ───────────────────────────────────────────────────────

    def __enter__(self) -> "CatalogoEtiquetaClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()
