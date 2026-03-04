"""
OneDrive-Integrationsmodul für den Stundenrechner.
Ermöglicht das Hochladen von Excel-Dateien in OneDrive via Microsoft Graph API.
"""

import os
from typing import Optional
from urllib.parse import quote

import requests


GRAPH_API = "https://graph.microsoft.com/v1.0"


class OneDriveClient:
    """Verwaltet Dateioperationen in OneDrive via Microsoft Graph API."""

    def __init__(self, auth):
        """
        Args:
            auth: MicrosoftAuth-Instanz für Token-Verwaltung.
        """
        self._auth = auth

    def _headers(self) -> dict:
        """Erstellt Authentifizierungs-Header mit aktuellem Token."""
        token = self._auth.get_token()
        if not token:
            raise RuntimeError("Kein gültiges Zugriffstoken. Bitte erneut anmelden.")
        return {"Authorization": f"Bearer {token}"}

    # ── Ordner-Operationen ────────────────────────────────────

    def list_folder_children(self, folder_id: str = "root") -> list[dict]:
        """
        Listet Unterordner eines OneDrive-Ordners auf.

        Args:
            folder_id: ID des Ordners oder 'root' für das Stammverzeichnis.

        Returns:
            Liste von Dicts mit keys: id, name, parent_id.
        """
        if folder_id == "root":
            url = f"{GRAPH_API}/me/drive/root/children"
        else:
            url = f"{GRAPH_API}/me/drive/items/{folder_id}/children"

        params = {
            "$filter": "folder ne null",
            "$select": "id,name,folder,parentReference",
            "$orderby": "name",
        }

        try:
            resp = requests.get(url, headers=self._headers(), params=params, timeout=15)
            resp.raise_for_status()
            items = resp.json().get("value", [])
            return [
                {
                    "id": item["id"],
                    "name": item["name"],
                    "parent_id": item.get("parentReference", {}).get("id", "root"),
                    "child_count": item.get("folder", {}).get("childCount", 0),
                }
                for item in items
            ]
        except requests.RequestException as e:
            raise RuntimeError(f"Fehler beim Abrufen der OneDrive-Ordner: {e}") from e

    def get_folder_info(self, folder_id: str) -> Optional[dict]:
        """
        Gibt Informationen über einen Ordner zurück.

        Returns:
            Dict mit keys: id, name, path oder None bei Fehler.
        """
        if folder_id == "root":
            url = f"{GRAPH_API}/me/drive/root"
        else:
            url = f"{GRAPH_API}/me/drive/items/{folder_id}"

        try:
            resp = requests.get(
                url,
                headers=self._headers(),
                params={"$select": "id,name,parentReference"},
                timeout=10,
            )
            resp.raise_for_status()
            data = resp.json()
            parent_path = data.get("parentReference", {}).get("path", "")
            # Pfad ohne "/drive/root:" Präfix
            clean_path = parent_path.replace("/drive/root:", "") if parent_path else ""
            full_path = f"{clean_path}/{data['name']}".lstrip("/")
            return {
                "id": data["id"],
                "name": data["name"],
                "path": full_path or data["name"],
            }
        except requests.RequestException:
            return None

    def get_root_info(self) -> Optional[dict]:
        """Gibt Informationen über das OneDrive-Stammverzeichnis zurück."""
        try:
            resp = requests.get(
                f"{GRAPH_API}/me/drive/root",
                headers=self._headers(),
                params={"$select": "id,name"},
                timeout=10,
            )
            resp.raise_for_status()
            data = resp.json()
            return {"id": data["id"], "name": "OneDrive", "path": "OneDrive"}
        except requests.RequestException:
            return None

    # ── Quota-Prüfung ─────────────────────────────────────────

    def get_quota_info(self) -> dict | None:
        """
        Gibt Quota-Informationen des OneDrive zurück.
        Rückgabe: {'total': int, 'used': int, 'remaining': int, 'state': str} oder None.
        """
        try:
            resp = requests.get(
                f"{GRAPH_API}/me/drive",
                headers=self._headers(),
                params={"$select": "quota"},
                timeout=10,
            )
            resp.raise_for_status()
            quota = resp.json().get("quota", {})
            return {
                "total": quota.get("total", 0),
                "used": quota.get("used", 0),
                "remaining": quota.get("remaining", 0),
                "state": quota.get("state", "normal"),
            }
        except requests.RequestException:
            return None

    # ── Datei-Upload ──────────────────────────────────────────

    def upload_file(self, local_path: str, folder_id: str, filename: str) -> bool:
        """
        Lädt eine lokale Datei in einen OneDrive-Ordner hoch (Simple Upload, max ~4 MB).

        Args:
            local_path: Lokaler Pfad der hochzuladenden Datei.
            folder_id: Ziel-Ordner-ID in OneDrive ('root' für Stammverzeichnis).
            filename: Dateiname im OneDrive-Ordner.

        Returns:
            True bei Erfolg.

        Raises:
            RuntimeError: Bei Netzwerk- oder API-Fehlern (mit sprechenden Meldungen).
        """
        # Vor dem Upload Quota prüfen
        try:
            file_size = os.path.getsize(local_path)
        except OSError:
            file_size = 0

        quota = self.get_quota_info()
        if quota and quota["state"] in ("exceeded", "nearing"):
            remaining_mb = quota["remaining"] / 1_048_576
            used_gb = quota["used"] / 1_073_741_824
            total_gb = quota["total"] / 1_073_741_824
            if quota["state"] == "exceeded" or (file_size > 0 and quota["remaining"] < file_size):
                raise RuntimeError(
                    f"OneDrive-Speicher ist voll!\n\n"
                    f"Genutzt: {used_gb:.1f} GB von {total_gb:.1f} GB\n"
                    f"Verbleibend: {remaining_mb:.0f} MB\n\n"
                    f"Bitte Dateien aus OneDrive löschen oder ein größeres Abo wählen:\n"
                    f"https://onedrive.live.com/about/plans/"
                )

        safe_filename = quote(filename, safe="")
        if folder_id == "root":
            url = f"{GRAPH_API}/me/drive/root:/{safe_filename}:/content?@microsoft.graph.conflictBehavior=replace"
        else:
            url = f"{GRAPH_API}/me/drive/items/{folder_id}:/{safe_filename}:/content?@microsoft.graph.conflictBehavior=replace"

        headers = self._headers()
        headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        try:
            with open(local_path, "rb") as f:
                data = f.read()

            resp = requests.put(
                url,
                headers=headers,
                data=data,
                timeout=60,
            )

            # Sprechende Fehlermeldung für bekannte Fehler
            if resp.status_code == 507:
                quota = self.get_quota_info()
                if quota:
                    used_gb = quota["used"] / 1_073_741_824
                    total_gb = quota["total"] / 1_073_741_824
                    raise RuntimeError(
                        f"OneDrive-Speicher ist voll!\n\n"
                        f"Genutzt: {used_gb:.1f} GB von {total_gb:.1f} GB\n\n"
                        f"Bitte Dateien aus OneDrive löschen oder ein größeres Abo wählen:\n"
                        f"https://onedrive.live.com/about/plans/"
                    )
                raise RuntimeError("Upload fehlgeschlagen: OneDrive-Speicher ist voll (507).")

            if resp.status_code == 403:
                raise RuntimeError(
                    "Zugriff verweigert (403).\n"
                    "Bitte prüfen Sie die API-Berechtigungen der Azure App Registration "
                    "(Files.ReadWrite muss genehmigt sein)."
                )

            resp.raise_for_status()
            return True
        except RuntimeError:
            raise
        except requests.RequestException as e:
            raise RuntimeError(f"Netzwerkfehler beim Hochladen: {e}") from e
        except OSError as e:
            raise RuntimeError(f"Fehler beim Lesen der Datei: {e}") from e

    def get_file_web_url(self, folder_id: str, filename: str) -> Optional[str]:
        """
        Gibt die Web-URL einer Datei im OneDrive zurück (zum Öffnen im Browser).

        Returns:
            Web-URL als String oder None bei Fehler.
        """
        safe_filename = quote(filename, safe="")
        if folder_id == "root":
            url = f"{GRAPH_API}/me/drive/root:/{safe_filename}"
        else:
            url = f"{GRAPH_API}/me/drive/items/{folder_id}:/{safe_filename}"

        try:
            resp = requests.get(
                url,
                headers=self._headers(),
                params={"$select": "webUrl"},
                timeout=10,
            )
            resp.raise_for_status()
            return resp.json().get("webUrl")
        except requests.RequestException:
            return None
