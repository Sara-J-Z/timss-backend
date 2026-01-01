import os
import io
import time
import re
import requests
import openpyxl
from urllib.parse import quote

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def safe_name(name: str) -> str:
    bad = ['/', '\\', ':', '*', '?', '"', '<', '>', '|']
    for ch in bad:
        name = name.replace(ch, '-')
    return name.strip()[:120] or "UnknownSchool"


def safe_table_name(name: str) -> str:
    """
    Excel table names:
    - must start with a letter or underscore
    - cannot contain spaces or most punctuation
    - keep it reasonably short
    """
    name = name.strip()
    name = re.sub(r"\s+", "_", name)
    name = re.sub(r"[^A-Za-z0-9_]", "_", name)
    if not name:
        name = "Sheet1"
    if not (name[0].isalpha() or name[0] == "_"):
        name = f"t_{name}"
    return name[:50]


class GraphExcelAppender:
    """
    Fully automatic:
    - Ensure folder TIMSS/<school> exists
    - Ensure workbook TIMSS/<school>/<school>.xlsx exists (create if missing)
    - Ensure worksheet <subject> exists (create if missing)
    - Ensure a table exists on that worksheet (create if missing)
    - Append a row into the table
    """

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

    # ---------- Auth ----------
    def _get_token(self) -> str:
        if self._token:
            return self._token

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
        return {"Authorization": f"Bearer {self._get_token()}"}

    # ---------- Drive helpers ----------
    def _get_item_by_path(self, path_in_drive: str) -> dict | None:
        url = f"{GRAPH_BASE}/users/{self.user_email}/drive/root:/{path_in_drive}"
        r = requests.get(url, headers=self._headers(), timeout=30)
        if r.status_code == 404:
            return None
        r.raise_for_status()
        return r.json()

    def _create_folder(self, parent_path: str, folder_name: str):
        # parent_path is like "" or "TIMSS" or "TIMSS/School"
        url = f"{GRAPH_BASE}/users/{self.user_email}/drive/root:/{parent_path}:/children"
        payload = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        }
        r = requests.post(
            url,
            headers={**self._headers(), "Content-Type": "application/json"},
            json=payload,
            timeout=30
        )
        if r.status_code == 409:
            return  # already exists
        r.raise_for_status()

    def _ensure_folder_path(self, full_path: str):
        # Ensure folders step-by-step: e.g. TIMSS/School
        parts = [p for p in full_path.split("/") if p.strip()]
        if not parts:
            return

        current = parts[0]
        if not self._get_item_by_path(current):
            self._create_folder("", current)

        for part in parts[1:]:
            next_path = f"{current}/{part}"
            if not self._get_item_by_path(next_path):
                self._create_folder(current, part)
            current = next_path

    def _upload_new_workbook(self, path_in_drive: str, subject: str, headers: list[str]) -> dict:
        """
        Create an XLSX locally in-memory and upload it to OneDrive via:
        PUT /drive/root:/path:/content
        """
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = subject
        ws.append(headers)

        bio = io.BytesIO()
        wb.save(bio)
        bio.seek(0)

        url = f"{GRAPH_BASE}/users/{self.user_email}/drive/root:/{path_in_drive}:/content"
        r = requests.put(url, headers=self._headers(), data=bio.read(), timeout=120)
        r.raise_for_status()
        return r.json()

    # ---------- Workbook helpers ----------
    def _list_worksheets(self, item_id: str) -> list[dict]:
        url = f"{GRAPH_BASE}/users/{self.user_email}/drive/items/{item_id}/workbook/worksheets"
        r = requests.get(url, headers=self._headers(), timeout=30)
        r.raise_for_status()
        return r.json().get("value", [])

    def _ensure_worksheet(self, item_id: str, sheet_name: str):
        sheets = self._list_worksheets(item_id)
        if any(s.get("name") == sheet_name for s in sheets):
            return

        url = f"{GRAPH_BASE}/users/{self.user_email}/drive/items/{item_id}/workbook/worksheets/add"
        payload = {"name": sheet_name}
        r = requests.post(
            url,
            headers={**self._headers(), "Content-Type": "application/json"},
            json=payload,
            timeout=30
        )
        if r.status_code == 409:
            return
        r.raise_for_status()

    def _list_tables(self, item_id: str) -> list[dict]:
        url = f"{GRAPH_BASE}/users/{self.user_email}/drive/items/{item_id}/workbook/tables"
        r = requests.get(url, headers=self._headers(), timeout=30)
        r.raise_for_status()
        return r.json().get("value", [])

    @staticmethod
    def _num_to_excel_col(n: int) -> str:
        result = ""
        while n > 0:
            n, rem = divmod(n - 1, 26)
            result = chr(65 + rem) + result
        return result

    def _ensure_table(self, item_id: str, sheet_name: str, table_name: str, columns_count: int) -> str:
        # If exists, return id
        tables = self._list_tables(item_id)
        for t in tables:
            if t.get("name") == table_name:
                return t.get("id")

        last_col_letter = self._num_to_excel_col(columns_count)

        # ✅ Excel address MUST quote sheet name if it has spaces/special chars
        safe_sheet_for_address = sheet_name.replace("'", "''")
        address = f"'{safe_sheet_for_address}'!A1:{last_col_letter}1"

        # ✅ Worksheet reference in URL must be worksheets('name') and URL-encoded
        sheet_ref = quote(sheet_name)
        url = f"{GRAPH_BASE}/users/{self.user_email}/drive/items/{item_id}/workbook/worksheets('{sheet_ref}')/tables/add"

        payload = {"address": address, "hasHeaders": True}
        r = requests.post(
            url,
            headers={**self._headers(), "Content-Type": "application/json"},
            json=payload,
            timeout=30
        )
        r.raise_for_status()

        return r.json().get("id")

    def _append_row_to_table(self, item_id: str, table_id_or_name: str, values: list):
        url = f"{GRAPH_BASE}/users/{self.user_email}/drive/items/{item_id}/workbook/tables/{table_id_or_name}/rows/add"
        payload = {"values": [values]}

        r = requests.post(
            url,
            headers={**self._headers(), "Content-Type": "application/json"},
            json=payload,
            timeout=30
        )

        # retry on transient errors (workbook busy / throttling)
        if r.status_code in (409, 429, 503):
            time.sleep(1.5)
            r = requests.post(
                url,
                headers={**self._headers(), "Content-Type": "application/json"},
                json=payload,
                timeout=30
            )

        r.raise_for_status()
        return r.json()

    # ---------- Public: Full automatic pipeline ----------
    def ensure_and_append(self, school_name: str, subject: str, headers: list[str], row_values: list):
        safe_school = safe_name(school_name)
        safe_subject = (subject or "Sheet1").strip()[:120] or "Sheet1"

        # 1) Ensure folders
        folder_path = f"{self.root_folder}/{safe_school}"
        self._ensure_folder_path(folder_path)

        # 2) Ensure workbook exists
        workbook_path = f"{folder_path}/{safe_school}.xlsx"
        item = self._get_item_by_path(workbook_path)

        if not item:
            # Create workbook with initial sheet + headers
            item = self._upload_new_workbook(workbook_path, safe_subject, headers)

        item_id = item["id"]

        # 3) Ensure worksheet exists
        self._ensure_worksheet(item_id, safe_subject)

        # 4) Ensure table exists (one per subject)
        table_name = safe_table_name(f"tbl_{safe_subject}")
        table_id = self._ensure_table(item_id, safe_subject, table_name, len(headers))

        # 5) Append row
        self._append_row_to_table(item_id, table_id, row_values)

        return {
            "status": "ok",
            "workbook": workbook_path,
            "sheet": safe_subject,
            "table": table_name
        }
