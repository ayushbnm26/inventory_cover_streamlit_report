from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from inventory_cover.config import GoogleDriveReportConfig, google_drive_report_config_from_values
from inventory_cover.io import google_drive_report_store as store_module
from inventory_cover.io.google_drive_report_store import (
    GoogleDriveFileMetadata,
    GoogleDriveReportStore,
    GoogleDriveReportStoreError,
)


def test_google_drive_config_defaults_disabled(tmp_path: Path) -> None:
    config = google_drive_report_config_from_values({}, project_root=tmp_path)

    assert config.enabled is False
    assert config.folder_id == ""
    assert config.service_account_json_path == tmp_path / "secrets" / "google_drive" / "service_account.json"
    assert config.report_file_name == "Inventory_Cover_Report_latest.xlsx"
    assert config.audit_file_name == "Inventory_Cover_Backend_Audit_latest.xlsx"
    assert config.upload_audit is False
    assert config.fail_on_upload_error is False


def test_google_drive_config_parses_enabled_values(tmp_path: Path) -> None:
    config = google_drive_report_config_from_values(
        {
            "GDRIVE_ENABLED": "true",
            "GDRIVE_FOLDER_ID": "159xM64Uuiat-NEA8sFmkRNdj7nCGlIrr",
            "GDRIVE_SERVICE_ACCOUNT_JSON_PATH": "secrets/google_drive/service_account.json",
            "GDRIVE_SERVICE_ACCOUNT_JSON": '{"type":"service_account"}',
            "GDRIVE_REPORT_FILE_NAME": "Report.xlsx",
            "GDRIVE_AUDIT_FILE_NAME": "Audit.xlsx",
            "GDRIVE_UPLOAD_AUDIT": "yes",
            "GDRIVE_FAIL_ON_UPLOAD_ERROR": "on",
        },
        project_root=tmp_path,
    )

    assert config.enabled is True
    assert config.folder_id == "159xM64Uuiat-NEA8sFmkRNdj7nCGlIrr"
    assert config.service_account_json_path == tmp_path / "secrets" / "google_drive" / "service_account.json"
    assert config.service_account_json == '{"type":"service_account"}'
    assert config.report_file_name == "Report.xlsx"
    assert config.audit_file_name == "Audit.xlsx"
    assert config.upload_audit is True
    assert config.fail_on_upload_error is True


def test_upload_creates_file_when_no_match(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    workbook = tmp_path / "Inventory_Cover_Report_latest.xlsx"
    workbook.write_bytes(b"workbook")
    service = FakeDriveService(files=[])
    store = GoogleDriveReportStore(service)
    monkeypatch.setattr(store_module, "_media_file_upload", lambda path: FakeMedia(path))

    result = store.upload_or_update_file(
        workbook,
        folder_id="folder123",
        drive_file_name="Inventory_Cover_Report_latest.xlsx",
        artifact="inventory_cover_report",
    )

    assert result.action == "created"
    assert result.metadata.file_id == "created-1"
    assert result.duplicate_count == 0
    assert service.resource.created_body == {
        "name": "Inventory_Cover_Report_latest.xlsx",
        "parents": ["folder123"],
    }
    assert service.resource.updated_file_id == ""


def test_upload_updates_newest_duplicate_and_warns(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    workbook = tmp_path / "Inventory_Cover_Report_latest.xlsx"
    workbook.write_bytes(b"workbook")
    service = FakeDriveService(
        files=[
            {
                "id": "older",
                "name": "Inventory_Cover_Report_latest.xlsx",
                "modifiedTime": "2026-06-30T10:00:00Z",
            },
            {
                "id": "newer",
                "name": "Inventory_Cover_Report_latest.xlsx",
                "modifiedTime": "2026-07-01T10:00:00Z",
            },
        ]
    )
    store = GoogleDriveReportStore(service)
    monkeypatch.setattr(store_module, "_media_file_upload", lambda path: FakeMedia(path))

    result = store.upload_or_update_file(
        workbook,
        folder_id="folder123",
        drive_file_name="Inventory_Cover_Report_latest.xlsx",
        artifact="inventory_cover_report",
    )

    assert result.action == "updated"
    assert result.metadata.file_id == "newer"
    assert result.duplicate_count == 2
    assert "2 Google Drive files" in result.warning
    assert service.resource.updated_file_id == "newer"


def test_missing_local_report_raises_clear_error(tmp_path: Path) -> None:
    store = GoogleDriveReportStore(FakeDriveService(files=[]))

    with pytest.raises(GoogleDriveReportStoreError) as exc:
        store.upload_or_update_file(
            tmp_path / "missing.xlsx",
            folder_id="folder123",
            drive_file_name="Inventory_Cover_Report_latest.xlsx",
            artifact="inventory_cover_report",
        )

    assert exc.value.issue_type == "GDRIVE_LOCAL_REPORT_MISSING"


def test_download_by_name_uses_newest_match(monkeypatch: pytest.MonkeyPatch) -> None:
    service = FakeDriveService(
        files=[
            {"id": "old", "name": "Report.xlsx", "modifiedTime": "2026-06-01T00:00:00Z"},
            {"id": "new", "name": "Report.xlsx", "modifiedTime": "2026-07-01T00:00:00Z"},
        ]
    )
    store = GoogleDriveReportStore(service)
    monkeypatch.setattr(store_module, "_download_media_request", lambda request: b"downloaded")

    content, metadata = store.download_file_by_name("folder123", "Report.xlsx")

    assert content == b"downloaded"
    assert metadata.file_id == "new"
    assert service.resource.get_media_file_id == "new"


def test_from_config_uses_json_content_before_file_path(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    called: dict[str, Any] = {}

    def fake_from_info(cls: type[GoogleDriveReportStore], info: str, *, scopes: tuple[str, ...]) -> str:
        called["info"] = info
        called["scopes"] = scopes
        return "store"

    monkeypatch.setattr(GoogleDriveReportStore, "from_service_account_info", classmethod(fake_from_info))
    config = GoogleDriveReportConfig(
        project_root=tmp_path,
        enabled=True,
        service_account_json='{"type":"service_account"}',
    )

    assert GoogleDriveReportStore.from_config(config) == "store"
    assert called["info"] == '{"type":"service_account"}'
    assert called["scopes"] == config.scopes


class FakeRequest:
    def __init__(self, payload: dict[str, Any]):
        self.payload = payload

    def execute(self) -> dict[str, Any]:
        return self.payload


class FakeMedia:
    def __init__(self, path: Path):
        self.path = path


class FakeDriveService:
    def __init__(self, *, files: list[dict[str, Any]]):
        self.resource = FakeFilesResource(files)

    def files(self) -> "FakeFilesResource":
        return self.resource


class FakeFilesResource:
    def __init__(self, files: list[dict[str, Any]]):
        self.files_payload = files
        self.created_body: dict[str, Any] = {}
        self.updated_file_id = ""
        self.get_media_file_id = ""

    def list(self, **kwargs: Any) -> FakeRequest:
        self.list_kwargs = kwargs
        return FakeRequest({"files": self.files_payload})

    def create(self, **kwargs: Any) -> FakeRequest:
        self.created_body = dict(kwargs["body"])
        return FakeRequest(
            {
                "id": "created-1",
                "name": self.created_body["name"],
                "modifiedTime": "2026-07-01T12:00:00Z",
                "size": "8",
                "webViewLink": "https://drive.google.com/file/d/created-1/view",
            }
        )

    def update(self, **kwargs: Any) -> FakeRequest:
        self.updated_file_id = str(kwargs["fileId"])
        return FakeRequest(
            {
                "id": self.updated_file_id,
                "name": kwargs["body"]["name"],
                "modifiedTime": "2026-07-01T12:00:00Z",
                "size": "8",
                "webViewLink": f"https://drive.google.com/file/d/{self.updated_file_id}/view",
            }
        )

    def get(self, **kwargs: Any) -> FakeRequest:
        file_id = str(kwargs["fileId"])
        return FakeRequest(
            {
                "id": file_id,
                "name": "Report.xlsx",
                "modifiedTime": "2026-07-01T12:00:00Z",
                "size": "10",
            }
        )

    def get_media(self, **kwargs: Any) -> FakeRequest:
        self.get_media_file_id = str(kwargs["fileId"])
        return FakeRequest({})
