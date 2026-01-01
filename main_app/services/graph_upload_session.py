import os
import time
import requests

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


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

    def create_upload_session(self, remote_path: str) -> str:
        """
        Create an upload session for a file path. We force replace to update the same file.
        """
        url = f"{GRAPH_BASE}/users/{self.user_email}/drive/root:/{remote_path}:/createUploadSession"
        payload = {"item": {"@microsoft.graph.conflictBehavior": "replace"}}

        r = requests.post(
            url,
            headers={**self._headers(), "Content-Type": "application/json"},
            json=payload,
            timeout=30
        )
        r.raise_for_status()
        return r.json()["uploadUrl"]

    def upload_large_file(
        self,
        local_path: str,
        remote_folder: str,
        remote_filename: str,
        chunk_size_mb: int = 10,
        max_retries: int = 5
    ) -> dict:
        """
        Chunked upload with retries for:
        - 423 Locked (file open / temporary lock)
        - 409 Conflict (session conflict / concurrent update)
        - 429/503 throttling
        """
        chunk_size = chunk_size_mb * 1024 * 1024

        remote_path = f"{self.root_folder}/{remote_folder}/{remote_filename}"
        total_size = os.path.getsize(local_path)

        # Outer retry loop: recreate upload session if needed
        for attempt in range(1, max_retries + 1):
            upload_url = None
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

                        # Completed
                        if r.status_code in (200, 201):
                            return r.json()

                        # Continue
                        if r.status_code == 202:
                            start = end + 1
                            continue

                        # Handle lock/throttle/conflict mid-upload
                        if r.status_code in (409, 423, 429, 503):
                            raise requests.HTTPError(f"{r.status_code} {r.text}", response=r)

                        r.raise_for_status()

                raise RuntimeError("Upload loop finished without completion response.")

            except requests.HTTPError as e:
                status = getattr(e.response, "status_code", None)

                # 423 Locked: file open or temporary lock
                if status == 423:
                    wait = min(2 ** attempt, 20)
                    print(f"[OneDrive Upload] Locked (423). Retry {attempt}/{max_retries} after {wait}s")
                    time.sleep(wait)
                    continue

                # 409 Conflict: concurrent update/session conflict
                if status == 409:
                    wait = min(2 ** attempt, 20)
                    print(f"[OneDrive Upload] Conflict (409). Retry {attempt}/{max_retries} after {wait}s")
                    time.sleep(wait)
                    continue

                # 429/503 throttling
                if status in (429, 503):
                    wait = min(2 ** attempt, 20)
                    print(f"[OneDrive Upload] Throttled ({status}). Retry {attempt}/{max_retries} after {wait}s")
                    time.sleep(wait)
                    continue

                # Other errors: stop
                raise

            except Exception:
                # Other exceptions: if more retries, wait a bit
                if attempt < max_retries:
                    wait = min(2 ** attempt, 20)
                    print(f"[OneDrive Upload] Error. Retry {attempt}/{max_retries} after {wait}s")
                    time.sleep(wait)
                    continue
                raise

        raise RuntimeError("Upload failed after max retries.")
