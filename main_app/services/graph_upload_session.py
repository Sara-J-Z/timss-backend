import os
import time
from urllib.parse import quote

import requests


GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_INVALID_SEGMENT_CHARS = set('\\/:*?"<>|')


class GraphUploadSessionClient:
    def __init__(self):
        self.tenant_id = os.getenv("AZURE_TENANT_ID")
        self.client_id = os.getenv("AZURE_CLIENT_ID")
        self.client_secret = os.getenv("AZURE_CLIENT_SECRET")
        self.user_email = os.getenv("ONEDRIVE_USER_EMAIL")
        self.root_folder = os.getenv("ONEDRIVE_ROOT_FOLDER", "TIMSS")

        missing = [k for k, v in {
            "AZURE_TENANT_ID": self.tenant_id,
            "AZURE_CLIENT_ID": self.client_id,
            "AZURE_CLIENT_SECRET": self.client_secret,
            "ONEDRIVE_USER_EMAIL": self.user_email,
        }.items() if not v]
        if missing:
            raise RuntimeError(f"Missing env vars: {', '.join(missing)}")

        self._token = None

    def get_app_token(self) -> str:
        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "grant_type": "client_credentials",
            "scope": "https://graph.microsoft.com/.default",
        }
        r = requests.post(token_url, data=data, timeout=30)
        r.raise_for_status()
        self._token = r.json()["access_token"]
        return self._token

    def _headers(self):
        if not self._token:
            self.get_app_token()
        return {"Authorization": f"Bearer {self._token}"}

    @staticmethod
    def _validate_segment(segment: str, label: str) -> str:
        value = (segment or "").strip()
        if not value or value in {".", ".."} or ".." in value:
            raise ValueError(f"Unsafe {label}: {segment!r}")
        if any(ch in value for ch in _INVALID_SEGMENT_CHARS):
            raise ValueError(f"Invalid character in {label}: {segment!r}")
        if value.endswith(".") or value.endswith(" "):
            raise ValueError(f"Unsafe ending in {label}: {segment!r}")
        return value

    def _build_remote_path(self, remote_folder: str, remote_filename: str) -> str:
        # root_folder may itself contain nested folders; validate each segment.
        root_parts = [p for p in self.root_folder.replace("\\", "/").split("/") if p]
        if not root_parts:
            raise ValueError("ONEDRIVE_ROOT_FOLDER is empty or invalid.")

        parts = []
        for idx, part in enumerate(root_parts):
            parts.append(self._validate_segment(part, f"root_folder segment {idx + 1}"))

        parts.append(self._validate_segment(remote_folder, "remote_folder"))
        parts.append(self._validate_segment(remote_filename, "remote_filename"))

        # Encode each path segment separately so spaces/Arabic characters are safe,
        # while preserving the '/' separators between folders.
        return "/".join(quote(part, safe="") for part in parts)

    def download_file(self, remote_folder: str, remote_filename: str, local_path: str) -> bool:
        """
        Download the current OneDrive workbook to local_path.

        Returns:
            True  -> existing workbook downloaded successfully.
            False -> Graph returned 404, so the workbook does not exist.

        Any other failure is raised. Callers must NOT create/replace a workbook
        after an unexpected download error.
        """
        remote_path = self._build_remote_path(remote_folder, remote_filename)
        url = f"{GRAPH_BASE}/users/{quote(self.user_email, safe='@')}/drive/root:/{remote_path}:/content"

        r = requests.get(url, headers=self._headers(), timeout=120)
        if r.status_code == 404:
            return False
        r.raise_for_status()

        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        with open(local_path, "wb") as f:
            f.write(r.content)
        return True

    def create_upload_session(self, remote_path: str) -> str:
        """Create an upload session and replace only the target workbook."""
        url = f"{GRAPH_BASE}/users/{quote(self.user_email, safe='@')}/drive/root:/{remote_path}:/createUploadSession"
        payload = {"item": {"@microsoft.graph.conflictBehavior": "replace"}}

        r = requests.post(
            url,
            headers={**self._headers(), "Content-Type": "application/json"},
            json=payload,
            timeout=30,
        )
        r.raise_for_status()
        return r.json()["uploadUrl"]

    @staticmethod
    def _retry_delay(response, attempt: int) -> int:
        retry_after = None
        if response is not None:
            retry_after = response.headers.get("Retry-After")
        if retry_after:
            try:
                return max(1, min(int(retry_after), 30))
            except (TypeError, ValueError):
                pass
        return min(2 ** attempt, 10)

    def upload_large_file(
        self,
        local_path: str,
        remote_folder: str,
        remote_filename: str,
        chunk_size_mb: int = 10,
        max_retries: int = 3,
    ) -> dict:
        """
        Chunked upload with limited retries for transient Graph/SharePoint errors.
        """
        chunk_size = chunk_size_mb * 1024 * 1024
        remote_path = self._build_remote_path(remote_folder, remote_filename)
        total_size = os.path.getsize(local_path)

        for attempt in range(1, max_retries + 1):
            try:
                upload_url = self.create_upload_session(remote_path)

                start = 0
                with open(local_path, "rb") as f:
                    while start < total_size:
                        end = min(start + chunk_size, total_size) - 1
                        length = (end - start) + 1

                        f.seek(start)
                        chunk = f.read(length)

                        headers = {
                            "Content-Length": str(length),
                            "Content-Range": f"bytes {start}-{end}/{total_size}",
                        }

                        r = requests.put(upload_url, headers=headers, data=chunk, timeout=180)

                        if r.status_code in (200, 201):
                            return r.json()

                        if r.status_code == 202:
                            start = end + 1
                            continue

                        if r.status_code in (409, 423, 429, 503):
                            raise requests.HTTPError(
                                f"{r.status_code} {r.text}", response=r
                            )

                        r.raise_for_status()

                raise RuntimeError("Upload loop finished without completion response.")

            except requests.HTTPError as e:
                status = getattr(e.response, "status_code", None)
                if status in (409, 423, 429, 503) and attempt < max_retries:
                    wait = self._retry_delay(e.response, attempt)
                    print(
                        f"[OneDrive Upload] Transient {status}. "
                        f"Retry {attempt}/{max_retries} after {wait}s"
                    )
                    time.sleep(wait)
                    continue
                raise

            except requests.RequestException as e:
                if attempt < max_retries:
                    wait = min(2 ** attempt, 10)
                    print(
                        f"[OneDrive Upload] Network error. "
                        f"Retry {attempt}/{max_retries} after {wait}s: {e}"
                    )
                    time.sleep(wait)
                    continue
                raise

        raise RuntimeError("Upload failed after max retries.")
