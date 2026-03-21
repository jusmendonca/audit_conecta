"""
hermes/lotacao.py
Cliente para buscar a lotação do usuário no SUPP.

Fluxo típico:
  1. GET /v1/administrativo/colaborador?where={"usuario.id":"eq:{user_id}"}
     → obtém o(s) ID(s) de Colaborador do usuário logado.

  2. GET /v1/administrativo/lotacao?where={"colaborador.id":"in:{ids}"}&populate=["setor"]
     → obtém os Setores onde o usuário está lotado.

Endpoints cobertos:
  GET /v1/administrativo/colaborador        Lista colaboradores com filtros
  GET /v1/administrativo/colaborador/count  Contagem
  GET /v1/administrativo/lotacao            Lista lotações com filtros
  GET /v1/administrativo/lotacao/count      Contagem
"""

from __future__ import annotations

import json
from typing import Any

import httpx

from .config import BASE_URL

_PATH_COLAB  = "/v1/administrativo/colaborador"
_PATH_LOTACA = "/v1/administrativo/lotacao"


class LotacaoError(Exception):
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
        raise LotacaoError(response.status_code, body)
    try:
        return response.json()
    except Exception:
        return response.text


def _where_str(where: dict) -> str:
    return json.dumps(where, ensure_ascii=False)


def _extract_list(data: Any) -> list[dict]:
    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("entities", "data", "results", "items"):
            if key in data and isinstance(data[key], list):
                return data[key]
    return []


class LotacaoClient:
    """
    Cliente síncrono para descobrir a lotação (setores) do usuário.

    Uso:
        from hermes.auth import AuthClient
        from hermes.lotacao import LotacaoClient

        auth = AuthClient()
        auth.login_ldap("usuario", "senha")
        lc = LotacaoClient.from_auth(auth)

        setores = lc.setores_do_usuario(usuario_id=42)
        # Retorna lista de dicts com campos do Setor
    """

    def __init__(
        self,
        token: str,
        base_url: str = BASE_URL,
        timeout: float = 30.0,
    ) -> None:
        self.token = token
        self.base_url = base_url
        self._http = httpx.Client(
            base_url=base_url,
            timeout=timeout,
            headers={"Authorization": f"Bearer {token}"},
        )

    @classmethod
    def from_auth(cls, auth_client: Any, timeout: float = 30.0) -> "LotacaoClient":
        if not auth_client.token:
            raise RuntimeError("AuthClient sem token. Faça login primeiro.")
        return cls(token=auth_client.token, base_url=auth_client.base_url, timeout=timeout)

    # ── colaborador ───────────────────────────────────────────────────────────

    def colaboradores_do_usuario(self, usuario_id: int | str) -> list[dict]:
        """
        Retorna os registros de Colaborador vinculados a um usuário.

        Normalmente um usuário tem apenas um Colaborador, mas pode ter mais.
        """
        params = {
            "where": _where_str({"usuario.id": f"eq:{usuario_id}"}),
            "limit": 10,
        }
        resp = self._http.get(_PATH_COLAB, params=params)
        return _extract_list(_check(resp))

    # ── lotação ───────────────────────────────────────────────────────────────

    def lotacoes_do_colaborador(self, colaborador_id: int | str) -> list[dict]:
        """
        Retorna as Lotações de um Colaborador, com o Setor populado.
        """
        params = {
            "where": _where_str({"colaborador.id": f"eq:{colaborador_id}"}),
            "populate": json.dumps(["setor"]),
            "limit": 50,
        }
        resp = self._http.get(_PATH_LOTACA, params=params)
        return _extract_list(_check(resp))

    def lotacoes_por_usuario_direto(self, usuario_id: int | str) -> list[dict]:
        """
        Tenta filtrar lotação diretamente por colaborador.usuario.id (filtro aninhado).
        Retorna lista vazia se a API não suportar o filtro aninhado.
        """
        params = {
            "where": _where_str({"colaborador.usuario.id": f"eq:{usuario_id}"}),
            "populate": json.dumps(["setor", "colaborador"]),
            "limit": 50,
        }
        try:
            resp = self._http.get(_PATH_LOTACA, params=params)
            return _extract_list(_check(resp))
        except LotacaoError:
            return []

    # ── helper principal ──────────────────────────────────────────────────────

    def setores_do_usuario(self, usuario_id: int | str) -> list[dict]:
        """
        Retorna a lista de Setores onde o usuário está lotado.

        Estratégia:
          1. Tenta filtro aninhado direto (1 chamada).
          2. Se falhar ou retornar vazio: busca colaborador(es) e depois lotações.

        Retorna lista de dicts de Setor com pelo menos: id, sigla, nome.
        Deduplica por setor_id (um colaborador pode ter mais de uma lotação no mesmo setor).
        """
        # Tentativa 1: filtro aninhado colaborador.usuario.id
        lotacoes = self.lotacoes_por_usuario_direto(usuario_id)

        # Tentativa 2: via colaborador
        if not lotacoes:
            colaboradores = self.colaboradores_do_usuario(usuario_id)
            for colab in colaboradores:
                cid = colab.get("id")
                if cid:
                    lotacoes.extend(self.lotacoes_do_colaborador(cid))

        # Extrai setores únicos
        vistos: set = set()
        setores: list[dict] = []
        for lot in lotacoes:
            setor = lot.get("setor")
            if not isinstance(setor, dict):
                continue
            sid = setor.get("id")
            if sid and sid not in vistos:
                vistos.add(sid)
                # Adiciona campo principal para destaque na UI
                setor["_principal"] = lot.get("principal", False)
                setores.append(setor)

        # Ordena: principal primeiro, depois por sigla
        setores.sort(key=lambda s: (not s.get("_principal", False), s.get("sigla", "")))
        return setores

    # ── context manager ───────────────────────────────────────────────────────

    def __enter__(self) -> "LotacaoClient":
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()
