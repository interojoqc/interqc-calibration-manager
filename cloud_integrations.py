from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass
from io import BytesIO
from typing import Any

import pandas as pd


SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/cloud-vision",
]


@dataclass
class CloudStatus:
    google_libs: bool
    credentials: bool
    drive_folder_id: bool
    spreadsheet_id: bool
    vision_ready: bool
    message: str


def get_secret(name: str, default: str = "") -> str:
    value = os.getenv(name)
    if value:
        return value
    try:
        import streamlit as st

        if name in st.secrets:
            return str(st.secrets[name])
    except Exception:
        pass
    return default


def google_libs_available() -> bool:
    try:
        import google.oauth2.service_account  # noqa: F401
        import googleapiclient.discovery  # noqa: F401
        import googleapiclient.http  # noqa: F401

        return True
    except Exception:
        return False


def load_service_account_info() -> dict[str, Any] | None:
    raw = get_secret("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw:
        path = get_secret("GOOGLE_APPLICATION_CREDENTIALS")
        if path and os.path.exists(path):
            with open(path, "r", encoding="utf-8") as file:
                return json.load(file)
        return None
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return None


def cloud_status() -> CloudStatus:
    libs = google_libs_available()
    creds = load_service_account_info() is not None
    drive = bool(get_secret("GOOGLE_DRIVE_FOLDER_ID"))
    sheet = bool(get_secret("GOOGLE_SHEET_ID"))
    vision = libs and creds and get_secret("ENABLE_VISION_OCR", "false").lower() == "true"
    missing = []
    if not libs:
        missing.append("google-api-python-client/google-cloud-vision")
    if not creds:
        missing.append("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not drive:
        missing.append("GOOGLE_DRIVE_FOLDER_ID")
    if not sheet:
        missing.append("GOOGLE_SHEET_ID")
    message = "Google 연동 준비 완료" if not missing else "미설정: " + ", ".join(missing)
    return CloudStatus(libs, creds, drive, sheet, vision, message)


def credentials():
    info = load_service_account_info()
    if not info:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON is not configured.")
    from google.oauth2 import service_account

    return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)


def drive_service():
    from googleapiclient.discovery import build

    return build("drive", "v3", credentials=credentials(), cache_discovery=False)


def sheets_service():
    from googleapiclient.discovery import build

    return build("sheets", "v4", credentials=credentials(), cache_discovery=False)


def upload_file_to_drive(file_name: str, content: bytes, mime_type: str = "application/octet-stream") -> str:
    folder_id = get_secret("GOOGLE_DRIVE_FOLDER_ID")
    if not folder_id:
        raise RuntimeError("GOOGLE_DRIVE_FOLDER_ID is not configured.")
    from googleapiclient.http import MediaIoBaseUpload

    metadata = {"name": file_name, "parents": [folder_id]}
    media = MediaIoBaseUpload(BytesIO(content), mimetype=mime_type, resumable=False)
    created = (
        drive_service()
        .files()
        .create(body=metadata, media_body=media, fields="id, webViewLink, webContentLink")
        .execute()
    )
    return created.get("webViewLink") or f"https://drive.google.com/file/d/{created['id']}/view"


def replace_sheet_with_dataframe(sheet_name: str, df: pd.DataFrame) -> None:
    spreadsheet_id = get_secret("GOOGLE_SHEET_ID")
    if not spreadsheet_id:
        raise RuntimeError("GOOGLE_SHEET_ID is not configured.")
    service = sheets_service()
    values = [list(df.columns)] + df.fillna("").astype(str).values.tolist()
    service.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=f"{sheet_name}!A:ZZ").execute()
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A1",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()


def extract_text_with_vision(file_name: str, content: bytes) -> str:
    try:
        from google.cloud import vision
    except Exception as exc:
        raise RuntimeError("google-cloud-vision is not installed.") from exc

    client = vision.ImageAnnotatorClient(credentials=credentials())
    lower = file_name.lower()
    if lower.endswith(".pdf"):
        raise RuntimeError("PDF OCR은 Cloud Storage 기반 비동기 처리로 다음 단계에서 연결합니다. 우선 이미지 파일 OCR을 사용하세요.")
    image = vision.Image(content=content)
    response = client.document_text_detection(image=image)
    if response.error.message:
        raise RuntimeError(response.error.message)
    return response.full_text_annotation.text or ""


def parse_certificate_candidates(text: str) -> dict[str, str]:
    normalized = re.sub(r"[ \t]+", " ", text or "")
    candidates: dict[str, str] = {}
    cert_patterns = [
        r"(?:성적서\s*(?:번호|No\.?)|Certificate\s*No\.?)\s*[:：]?\s*([A-Za-z0-9_-]{3,})",
        r"\b([A-Z]{1,5}[-_]?\d{3,}[-_]?\d*)\b",
    ]
    for pattern in cert_patterns:
        match = re.search(pattern, normalized, re.IGNORECASE)
        if match:
            candidates["certificate_no"] = match.group(1)
            break
    date_matches = re.findall(r"(\d{4}[./-]\d{1,2}[./-]\d{1,2}|\d{2}[./-]\d{1,2}[./-]\d{1,2})", normalized)
    if date_matches:
        candidates["calibration_date"] = date_matches[0]
    if len(date_matches) > 1:
        candidates["next_due_date"] = date_matches[-1]
    number_matches = re.findall(r"[-+]?\d+(?:\.\d+)?", normalized)
    if number_matches:
        candidates["numeric_values"] = ", ".join(number_matches[:10])
    return candidates
