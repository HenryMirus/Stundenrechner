"""
Authentifizierungsmodul für den Stundenrechner.
Verwaltet Microsoft-Anmeldungen via MSAL mit persistentem Token-Cache.

Voraussetzung: Azure App Registration mit folgenden Einstellungen:
  - Kontotyp: Konten in beliebigen Organisationsverzeichnissen + persönliche Konten
  - Umleitungs-URI: Öffentlicher Client/nativ → http://localhost
  - API-Berechtigungen (delegiert): User.Read, Files.ReadWrite
"""

import hashlib
import json
import os
from pathlib import Path

import msal
import requests

# ── Konfiguration ─────────────────────────────────────────────
# Azure App Registration – Client-ID hier eintragen:
MS_CLIENT_ID = "01ef8aca-a8c3-4c98-8c47-4176a77bcb5c"

MS_AUTHORITY = "https://login.microsoftonline.com/common"
MS_SCOPES = ["User.Read", "Files.ReadWrite"]
GRAPH_API = "https://graph.microsoft.com/v1.0"

_APP_DIR = os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")),
    "Stundenrechner",
)
_CACHE_PATH = os.path.join(_APP_DIR, "auth", "token_cache.bin")


class MicrosoftAuth:
    """Verwaltet die Microsoft-Authentifizierung und den Token-Cache."""

    def __init__(self):
        os.makedirs(os.path.dirname(_CACHE_PATH), exist_ok=True)
        self._cache = msal.SerializableTokenCache()
        if os.path.exists(_CACHE_PATH):
            with open(_CACHE_PATH, "r", encoding="utf-8") as f:
                self._cache.deserialize(f.read())

        self._app = msal.PublicClientApplication(
            MS_CLIENT_ID,
            authority=MS_AUTHORITY,
            token_cache=self._cache,
        )
        self._current_account: dict | None = None
        self._current_token: str | None = None

    # ── Cache persistieren ────────────────────────────────────

    def _save_cache(self):
        """Speichert den Token-Cache auf die Festplatte."""
        if self._cache.has_state_changed:
            with open(_CACHE_PATH, "w", encoding="utf-8") as f:
                f.write(self._cache.serialize())

    # ── Konten ────────────────────────────────────────────────

    def get_accounts(self) -> list[dict]:
        """
        Gibt alle gespeicherten Konten aus dem Cache zurück.
        Jedes Konto-Dict enthält mindestens: username, home_account_id, name.
        """
        return self._app.get_accounts()

    def get_current_account(self) -> dict | None:
        """Gibt das aktuell angemeldete Konto zurück."""
        return self._current_account

    @property
    def current_user_id(self) -> str | None:
        """Eindeutige ID des aktuell angemeldeten Benutzers."""
        if self._current_account:
            return self._current_account.get("home_account_id")
        return None

    @property
    def current_user_id_short(self) -> str | None:
        """Kurzform der User-ID als sicherer Dateiname (12-stelliger Hash)."""
        uid = self.current_user_id
        if uid:
            return hashlib.sha256(uid.encode()).hexdigest()[:12]
        return None

    def is_logged_in(self) -> bool:
        """Gibt True zurück, wenn ein Benutzer angemeldet ist und ein Token verfügbar ist."""
        return self._current_account is not None and self._current_token is not None

    # ── Anmeldung ─────────────────────────────────────────────

    def login_interactive(self) -> bool:
        """
        Öffnet den Browser für die Microsoft-Anmeldung.
        Gibt True zurück bei Erfolg, False bei Abbruch oder Fehler.
        """
        result = self._app.acquire_token_interactive(
            scopes=MS_SCOPES,
            prompt="select_account",
        )
        return self._handle_result(result)

    def login_silent(self, account: dict) -> bool:
        """
        Versucht ein Token ohne Benutzerinteraktion zu erneuern.
        Gibt True zurück bei Erfolg, False wenn interaktive Anmeldung nötig.
        """
        result = self._app.acquire_token_silent(
            scopes=MS_SCOPES,
            account=account,
        )
        return self._handle_result(result)

    def _handle_result(self, result: dict | None) -> bool:
        """Verarbeitet ein MSAL-Ergebnis und setzt das aktuelle Konto."""
        if result and "access_token" in result:
            self._current_token = result["access_token"]
            accounts = self._app.get_accounts()
            # Konto anhand der oid/sub aus den id_token_claims identifizieren
            id_claims = result.get("id_token_claims") or {}
            oid = id_claims.get("oid") or id_claims.get("sub")
            matched = None
            if oid:
                for acc in accounts:
                    if oid in acc.get("home_account_id", ""):
                        matched = acc
                        break
            # Fallback: einziges Konto oder zuletzt hinzugefügtes Konto
            if matched is None and accounts:
                matched = accounts[-1]
            self._current_account = matched
            self._save_cache()
            return True
        return False

    def logout(self, account: dict | None = None):
        """
        Entfernt ein Konto aus dem lokalen Cache.
        Das Konto bleibt im Microsoft-System angemeldet, aber die App
        muss sich beim nächsten Mal neu authentifizieren.
        """
        target = account or self._current_account
        if target:
            self._app.remove_account(target)
            self._save_cache()
        if target == self._current_account or account is None:
            self._current_account = None
            self._current_token = None

    def switch_account(self, account: dict) -> bool:
        """Wechselt zu einem anderen gespeicherten Konto (silent-login)."""
        self._current_account = None
        self._current_token = None
        result = self._app.acquire_token_silent(scopes=MS_SCOPES, account=account)
        if result and "access_token" in result:
            # Konto explizit setzen – id_token_claims fehlt bei silent-refresh
            self._current_account = account
            self._current_token = result["access_token"]
            self._save_cache()
            return True
        return False

    # ── Token ─────────────────────────────────────────────────

    def get_token(self) -> str | None:
        """
        Gibt ein gültiges Access-Token zurück.
        Versucht bei Ablauf automatisch ein stilles Token zu erneuern.
        """
        if not self._current_account:
            return None
        # Stilles Erneuern versuchen
        result = self._app.acquire_token_silent(
            scopes=MS_SCOPES,
            account=self._current_account,
        )
        if result and "access_token" in result:
            self._current_token = result["access_token"]
            self._save_cache()
        return self._current_token

    # ── Benutzerdaten abrufen ─────────────────────────────────

    def get_user_info(self) -> dict | None:
        """
        Gibt Benutzername und E-Mail über die Microsoft Graph API zurück.
        Rückgabe: {'name': '...', 'email': '...', 'id': '...'} oder None.
        """
        token = self.get_token()
        if not token:
            return None
        try:
            resp = requests.get(
                f"{GRAPH_API}/me",
                headers={"Authorization": f"Bearer {token}"},
                timeout=10,
            )
            resp.raise_for_status()
            data = resp.json()
            return {
                "name": data.get("displayName") or data.get("userPrincipalName", "Unbekannt"),
                "email": data.get("mail") or data.get("userPrincipalName", ""),
                "id": data.get("id", ""),
            }
        except Exception:
            return None
