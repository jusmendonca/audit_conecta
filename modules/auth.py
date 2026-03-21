"""
hermes/auth.py
Cliente de autenticação para o SUPP (Sistema Único de Procuradorias Públicas).

Endpoints cobertos:
  POST /auth/get_token              Login usuário/senha → JWT
  POST /auth/ldap_get_token         Login via LDAP
  POST /auth/govbr_get_token        Login via Gov.BR
  GET  /auth/refresh_token          Renova o JWT atual
  GET  /auth/payload_token          Decodifica o payload do JWT
  POST /auth/update_password        Altera senha
  POST /auth/recover_password       Solicita e-mail de recuperação de senha
  GET  /auth/recover_password       Valida token e redefine senha
  POST /auth/totp/setup             Configura 2FA TOTP
  POST /auth/totp/sendMail          Envia challenge TOTP por e-mail
  POST /auth/totp/verify            Verifica código TOTP → JWT final
  GET  /auth/generate_client_secret          Gera secret para robô/automação
  GET  /auth/generate_client_rate_limit_secret  Gera secret de rate-limit
"""

from __future__ import annotations

import httpx

from .config import BASE_URL


class AuthError(Exception):
    """Erro retornado pela API de autenticação."""

    def __init__(self, status_code: int, body: object) -> None:
        self.status_code = status_code
        self.body = body
        super().__init__(f"HTTP {status_code}: {body}")


def _check(response: httpx.Response) -> dict:
    """Levanta AuthError se a resposta não for 2xx, caso contrário retorna o JSON."""
    if not response.is_success:
        try:
            body = response.json()
        except Exception:
            body = response.text
        raise AuthError(response.status_code, body)
    try:
        return response.json()
    except Exception:
        return {"raw": response.text}


class AuthClient:
    """
    Cliente síncrono para os endpoints de autenticação do SUPP.

    Uso básico:
        client = AuthClient()
        token = client.login("12345678900", "senha123")
        # token é uma string JWT

    Após o login o token fica armazenado em client.token e é enviado
    automaticamente nos métodos que requerem Bearer.
    """

    def __init__(self, base_url: str = BASE_URL, timeout: float = 30.0) -> None:
        self.base_url = base_url
        self.token: str | None = None
        self._http = httpx.Client(base_url=base_url, timeout=timeout)

    # ── helpers ───────────────────────────────────────────────────────────────

    def _bearer_headers(self) -> dict:
        if not self.token:
            raise RuntimeError("Nenhum token disponível. Faça login primeiro.")
        return {"Authorization": f"Bearer {self.token}"}

    def _store(self, data: dict) -> str:
        """Extrai e armazena o token JWT retornado pela API."""
        # A API pode retornar {"token": "..."} ou {"access_token": "..."} etc.
        token = (
            data.get("token")
            or data.get("access_token")
            or data.get("jwt")
        )
        if not token:
            raise AuthError(200, f"Token não encontrado na resposta: {data}")
        self.token = token
        return token

    # ── login / obtenção de token ─────────────────────────────────────────────

    def login(self, username: str, password: str) -> str:
        """
        POST /auth/get_token
        Login com usuário (CPF/CNPJ) e senha. Retorna o JWT e o armazena
        internamente para uso automático nos demais métodos.
        """
        resp = self._http.post(
            "/auth/get_token",
            json={"username": username, "password": password},
        )
        return self._store(_check(resp))

    def login_ldap(self, username: str, password: str) -> str:
        """POST /auth/ldap_get_token — Login via LDAP."""
        resp = self._http.post(
            "/auth/ldap_get_token",
            json={"username": username, "password": password},
        )
        return self._store(_check(resp))

    def login_govbr(self, code: str, redirect_uri: str) -> str:
        """POST /auth/govbr_get_token — Login via Gov.BR (OAuth callback)."""
        resp = self._http.post(
            "/auth/govbr_get_token",
            json={"code": code, "redirectUri": redirect_uri},
        )
        return self._store(_check(resp))

    # ── operações com token ───────────────────────────────────────────────────

    def refresh(self) -> str:
        """GET /auth/refresh_token — Renova o JWT atual. Requer login prévio."""
        resp = self._http.get(
            "/auth/refresh_token",
            headers=self._bearer_headers(),
        )
        return self._store(_check(resp))

    def payload(self) -> dict:
        """GET /auth/payload_token — Retorna o payload decodificado do JWT."""
        resp = self._http.get(
            "/auth/payload_token",
            headers=self._bearer_headers(),
        )
        return _check(resp)

    # ── senha ─────────────────────────────────────────────────────────────────

    def update_password(self, old_password: str, new_password: str, token: str = "") -> dict:
        """POST /auth/update_password — Altera a senha do usuário."""
        resp = self._http.post(
            "/auth/update_password",
            json={
                "old_password": old_password,
                "new_password": new_password,
                "token": token,
            },
        )
        return _check(resp)

    def request_password_recovery(self, username: str, email: str) -> dict:
        """POST /auth/recover_password — Solicita e-mail de recuperação de senha."""
        resp = self._http.post(
            "/auth/recover_password",
            json={"username": username, "email": email},
        )
        return _check(resp)

    def reset_password(self, username: str, email: str, reset_token: str) -> dict:
        """GET /auth/recover_password — Valida token e redefine a senha."""
        resp = self._http.get(
            "/auth/recover_password",
            params={"username": username, "email": email, "reset_token": reset_token},
        )
        return _check(resp)

    # ── TOTP (2FA) ────────────────────────────────────────────────────────────

    def totp_setup(self, challenge: str, regenerate: bool = False) -> dict:
        """POST /auth/totp/setup — Configura o segundo fator TOTP."""
        resp = self._http.post(
            "/auth/totp/setup",
            json={"challenge": challenge, "regenerate": regenerate},
        )
        return _check(resp)

    def totp_send_mail(self, challenge: str) -> dict:
        """POST /auth/totp/sendMail — Envia o challenge TOTP por e-mail."""
        resp = self._http.post(
            "/auth/totp/sendMail",
            json={"challenge": challenge},
        )
        return _check(resp)

    def totp_verify(self, challenge: str, code: str) -> str:
        """POST /auth/totp/verify — Verifica o código TOTP e retorna JWT final."""
        resp = self._http.post(
            "/auth/totp/verify",
            json={"challenge": challenge, "code": code},
        )
        return self._store(_check(resp))

    # ── automação / robôs ─────────────────────────────────────────────────────

    def generate_client_secret(self, client_id: str | None = None) -> dict:
        """GET /auth/generate_client_secret — Gera client_secret para robô."""
        params = {"client_id": client_id} if client_id else {}
        resp = self._http.get(
            "/auth/generate_client_secret",
            headers=self._bearer_headers(),
            params=params,
        )
        return _check(resp)

    def generate_client_rate_limit_secret(self, client_id: str | None = None) -> dict:
        """GET /auth/generate_client_rate_limit_secret — Gera secret de rate-limit."""
        params = {"client_id": client_id} if client_id else {}
        resp = self._http.get(
            "/auth/generate_client_rate_limit_secret",
            headers=self._bearer_headers(),
            params=params,
        )
        return _check(resp)

    # ── context manager ───────────────────────────────────────────────────────

    def __enter__(self) -> AuthClient:
        return self

    def __exit__(self, *_) -> None:
        self._http.close()

    def close(self) -> None:
        self._http.close()
