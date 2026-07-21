from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any


DRIVE_SCOPES = ("https://www.googleapis.com/auth/drive",)
GOOGLE_SHEETS_MIME_TYPE = "application/vnd.google-apps.spreadsheet"
XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@dataclass(frozen=True)
class DriveItem:
    id: str
    name: str
    mime_type: str
    modified_time: str | None = None
    size: str | None = None
    web_view_link: str | None = None


def drive_api_available() -> tuple[bool, str]:
    try:
        import googleapiclient.discovery  # noqa: F401
        import googleapiclient.http  # noqa: F401
        import google.oauth2.service_account  # noqa: F401
    except ModuleNotFoundError as exc:
        return False, exc.name or str(exc)
    return True, ""


def parse_drive_folder_id(value: str) -> str:
    text = value.strip()
    if not text:
        raise ValueError("Paste a Google Drive folder link or folder ID.")

    patterns = (
        r"/folders/([A-Za-z0-9_-]+)",
        r"[?&]id=([A-Za-z0-9_-]+)",
        r"^([A-Za-z0-9_-]{20,})$",
    )
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1)

    raise ValueError("Could not read a folder ID from that Google Drive link.")


def service_account_email(credentials_bytes: bytes) -> str:
    info = json.loads(credentials_bytes.decode("utf-8"))
    email = info.get("client_email")
    if not email:
        raise ValueError("The uploaded credentials file does not include a service-account email.")
    return str(email)


def build_drive_service(credentials_bytes: bytes) -> Any:
    available, missing = drive_api_available()
    if not available:
        raise RuntimeError(f"Google Drive API packages are not installed yet: {missing}")

    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    info = json.loads(credentials_bytes.decode("utf-8"))
    credentials = service_account.Credentials.from_service_account_info(
        info,
        scopes=DRIVE_SCOPES,
    )
    return build("drive", "v3", credentials=credentials)


def list_folder_items(service: Any, folder_id: str) -> list[DriveItem]:
    items: list[DriveItem] = []
    page_token = None
    query = f"'{folder_id}' in parents and trashed = false"

    while True:
        response = (
            service.files()
            .list(
                q=query,
                pageSize=1000,
                pageToken=page_token,
                fields="nextPageToken, files(id, name, mimeType, modifiedTime, size, webViewLink)",
                includeItemsFromAllDrives=True,
                supportsAllDrives=True,
            )
            .execute()
        )
        for file in response.get("files", []):
            items.append(
                DriveItem(
                    id=file["id"],
                    name=file["name"],
                    mime_type=file["mimeType"],
                    modified_time=file.get("modifiedTime"),
                    size=file.get("size"),
                    web_view_link=file.get("webViewLink"),
                )
            )
        page_token = response.get("nextPageToken")
        if not page_token:
            return items


def download_xlsx(service: Any, item: DriveItem, destination: Path) -> None:
    from googleapiclient.http import MediaIoBaseDownload

    if item.mime_type == GOOGLE_SHEETS_MIME_TYPE:
        request = service.files().export_media(fileId=item.id, mimeType=XLSX_MIME_TYPE)
    else:
        request = service.files().get_media(fileId=item.id, supportsAllDrives=True)

    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("wb") as handle:
        downloader = MediaIoBaseDownload(handle, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
