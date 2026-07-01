"""Google Drive storage for latest Inventory Cover report workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import io
import json
import logging
from pathlib import Path
import time
from typing import Any, Mapping

from inventory_cover.config import GoogleDriveReportConfig
from inventory_cover.exceptions import PipelineError
from inventory_cover.inventory_cover_schemas import InventoryCoverPipelineRunResult
from inventory_cover.logging_utils import setup_run_logger, write_json_file


EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
DRIVE_FILE_FIELDS = "id,name,mimeType,modifiedTime,size,webViewLink,parents"


class GoogleDriveReportStoreError(PipelineError):
    """Actionable Google Drive latest-report storage failure."""

    def __init__(
        self,
        message: str,
        *,
        issue_type: str = "GDRIVE_ERROR",
        classification: str = "api",
    ):
        super().__init__(message)
        self.issue_type = issue_type
        self.classification = classification


@dataclass(frozen=True)
class GoogleDriveFileMetadata:
    """Small, audit-safe subset of Google Drive file metadata."""

    file_id: str
    name: str
    modified_time: str = ""
    size: str = ""
    web_view_link: str = ""
    mime_type: str = ""

    @classmethod
    def from_api(cls, payload: Mapping[str, Any]) -> "GoogleDriveFileMetadata":
        return cls(
            file_id=str(payload.get("id") or ""),
            name=str(payload.get("name") or ""),
            modified_time=str(payload.get("modifiedTime") or ""),
            size=str(payload.get("size") or ""),
            web_view_link=str(payload.get("webViewLink") or ""),
            mime_type=str(payload.get("mimeType") or ""),
        )

    def as_json(self) -> dict[str, str]:
        return {
            "file_id": self.file_id,
            "name": self.name,
            "modified_time": self.modified_time,
            "size": self.size,
            "web_view_link": self.web_view_link,
            "mime_type": self.mime_type,
        }


@dataclass(frozen=True)
class GoogleDriveUploadResult:
    """Result for one uploaded or updated Drive workbook."""

    artifact: str
    local_path: Path
    drive_file_name: str
    action: str
    metadata: GoogleDriveFileMetadata
    duplicate_count: int = 0
    warning: str = ""

    def as_json(self) -> dict[str, Any]:
        return {
            "artifact": self.artifact,
            "local_path": str(self.local_path),
            "drive_file_name": self.drive_file_name,
            "action": self.action,
            "duplicate_count": self.duplicate_count,
            "warning": self.warning,
            "metadata": self.metadata.as_json(),
        }


@dataclass(frozen=True)
class GoogleDriveUploadSummary:
    """Traceable outcome of the optional Drive upload layer."""

    run_id: str
    enabled: bool
    status: str
    audit_file: Path
    log_file: Path
    started_at: str
    completed_at: str
    duration_seconds: float
    uploads: tuple[GoogleDriveUploadResult, ...] = ()
    error_type: str = ""
    error_message_sanitized: str = ""
    warnings: tuple[str, ...] = ()

    def as_json(self) -> dict[str, Any]:
        return {
            "run_id": self.run_id,
            "enabled": self.enabled,
            "status": self.status,
            "started_at": self.started_at,
            "completed_at": self.completed_at,
            "duration_seconds": self.duration_seconds,
            "uploads": [upload.as_json() for upload in self.uploads],
            "error_type": self.error_type,
            "error_message_sanitized": self.error_message_sanitized,
            "warnings": list(self.warnings),
        }


class GoogleDriveReportStore:
    """Upload/update/download stable latest report files in one Drive folder."""

    def __init__(self, service: Any):
        self._service = service

    @classmethod
    def from_service_account_file(
        cls,
        path: Path,
        *,
        scopes: tuple[str, ...],
    ) -> "GoogleDriveReportStore":
        """Build a Drive store from a local service-account JSON key file."""

        if not Path(path).exists():
            raise GoogleDriveReportStoreError(
                f"Google Drive service account JSON file is missing: {path}",
                issue_type="GDRIVE_SERVICE_ACCOUNT_FILE_MISSING",
                classification="configuration",
            )
        try:
            from google.oauth2 import service_account
            from googleapiclient.discovery import build
        except ImportError as exc:
            raise GoogleDriveReportStoreError(
                "Google Drive dependencies are not installed. Run `python -m pip install -e .` first.",
                issue_type="GDRIVE_DEPENDENCY_MISSING",
                classification="dependency",
            ) from exc

        try:
            credentials = service_account.Credentials.from_service_account_file(
                str(path),
                scopes=list(scopes),
            )
            service = build("drive", "v3", credentials=credentials, cache_discovery=False)
            return cls(service)
        except Exception as exc:
            raise GoogleDriveReportStoreError(
                f"Google Drive service account authentication failed: {_classify_exception_message(exc)}",
                issue_type="GDRIVE_AUTHENTICATION_FAILED",
                classification="authentication",
            ) from exc

    @classmethod
    def from_service_account_info(
        cls,
        info: Mapping[str, Any] | str,
        *,
        scopes: tuple[str, ...],
    ) -> "GoogleDriveReportStore":
        """Build a Drive store from service-account JSON content or a TOML dict."""

        try:
            from google.oauth2 import service_account
            from googleapiclient.discovery import build
        except ImportError as exc:
            raise GoogleDriveReportStoreError(
                "Google Drive dependencies are not installed. Run `python -m pip install -e .` first.",
                issue_type="GDRIVE_DEPENDENCY_MISSING",
                classification="dependency",
            ) from exc

        try:
            credentials_info = _normalize_service_account_info(info)
            credentials = service_account.Credentials.from_service_account_info(
                credentials_info,
                scopes=list(scopes),
            )
            service = build("drive", "v3", credentials=credentials, cache_discovery=False)
            return cls(service)
        except json.JSONDecodeError as exc:
            raise GoogleDriveReportStoreError(
                "Google Drive service account JSON content is not valid JSON.",
                issue_type="GDRIVE_SERVICE_ACCOUNT_JSON_INVALID",
                classification="configuration",
            ) from exc
        except Exception as exc:
            raise GoogleDriveReportStoreError(
                f"Google Drive service account authentication failed: {_classify_exception_message(exc)}",
                issue_type="GDRIVE_AUTHENTICATION_FAILED",
                classification="authentication",
            ) from exc

    @classmethod
    def from_oauth_user(
        cls,
        *,
        credentials_path: Path,
        token_path: Path,
        scopes: tuple[str, ...],
    ) -> "GoogleDriveReportStore":
        """Build a Drive store from a local OAuth Desktop client and user token."""

        credentials_path = Path(credentials_path)
        token_path = Path(token_path)
        if not credentials_path.exists():
            raise GoogleDriveReportStoreError(
                f"Google Drive OAuth credentials file is missing: {credentials_path}",
                issue_type="GDRIVE_OAUTH_CREDENTIALS_FILE_MISSING",
                classification="configuration",
            )
        try:
            from google.auth.exceptions import RefreshError, TransportError
            from google.auth.transport.requests import Request
            from google.oauth2.credentials import Credentials
            from google_auth_oauthlib.flow import InstalledAppFlow
            from googleapiclient.discovery import build
        except ImportError as exc:
            raise GoogleDriveReportStoreError(
                "Google Drive OAuth dependencies are not installed. Run `python -m pip install -e .` first.",
                issue_type="GDRIVE_DEPENDENCY_MISSING",
                classification="dependency",
            ) from exc

        requested_scopes = list(scopes)
        credentials = None
        try:
            if token_path.exists():
                credentials = Credentials.from_authorized_user_file(str(token_path), requested_scopes)
                if credentials and not credentials.has_scopes(requested_scopes):
                    credentials = None
            if credentials and credentials.expired and credentials.refresh_token:
                credentials.refresh(Request())
            if not credentials or not credentials.valid:
                flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), requested_scopes)
                credentials = flow.run_local_server(port=0)
            token_path.parent.mkdir(parents=True, exist_ok=True)
            token_path.write_text(credentials.to_json(), encoding="utf-8")
            service = build("drive", "v3", credentials=credentials, cache_discovery=False)
            return cls(service)
        except RefreshError as exc:
            raise GoogleDriveReportStoreError(
                "Google Drive OAuth token refresh failed. Delete the Drive token file and re-authorize.",
                issue_type="GDRIVE_OAUTH_TOKEN_ERROR",
                classification="oauth",
            ) from exc
        except TransportError as exc:
            raise GoogleDriveReportStoreError(
                f"Google Drive OAuth transport failed: {_classify_exception_message(exc)}",
                issue_type="GDRIVE_TRANSPORT_FAILURE",
                classification="transport",
            ) from exc
        except Exception as exc:
            raise GoogleDriveReportStoreError(
                f"Google Drive OAuth authentication failed: {_classify_exception_message(exc)}",
                issue_type="GDRIVE_AUTHENTICATION_FAILED",
                classification="authentication",
            ) from exc

    @classmethod
    def from_config(cls, config: GoogleDriveReportConfig) -> "GoogleDriveReportStore":
        """Build a Drive store from the configured upload authentication mode."""

        resolved = config.resolved()
        if resolved.auth_mode == "oauth_user":
            return cls.from_oauth_user(
                credentials_path=resolved.oauth_credentials_path,
                token_path=resolved.oauth_token_path,
                scopes=resolved.scopes,
            )
        if resolved.service_account_json.strip():
            return cls.from_service_account_info(resolved.service_account_json, scopes=resolved.scopes)
        return cls.from_service_account_file(resolved.service_account_json_path, scopes=resolved.scopes)

    def find_files_by_name(self, folder_id: str, file_name: str) -> list[GoogleDriveFileMetadata]:
        """Return non-trashed exact filename matches inside one Drive folder."""

        _require_value(folder_id, "GDRIVE_FOLDER_ID")
        _require_value(file_name, "Drive file name")
        query = (
            f"'{_escape_drive_query_literal(folder_id)}' in parents and "
            f"name = '{_escape_drive_query_literal(file_name)}' and trashed = false"
        )
        files_payload: list[dict[str, Any]] = []
        page_token = None
        try:
            while True:
                response = (
                    self._service.files()
                    .list(
                        q=query,
                        spaces="drive",
                        fields=f"nextPageToken,files({DRIVE_FILE_FIELDS})",
                        orderBy="modifiedTime desc",
                        pageSize=100,
                        pageToken=page_token,
                        supportsAllDrives=True,
                        includeItemsFromAllDrives=True,
                    )
                    .execute()
                )
                files_payload.extend(response.get("files", []))
                page_token = response.get("nextPageToken")
                if not page_token:
                    break
        except Exception as exc:
            raise _drive_api_error(exc, f"Could not search Google Drive folder for {file_name!r}.") from exc
        files = [GoogleDriveFileMetadata.from_api(item) for item in files_payload]
        return sorted(files, key=lambda item: item.modified_time, reverse=True)

    def upload_or_update_file(
        self,
        local_path: Path,
        *,
        folder_id: str,
        drive_file_name: str,
        artifact: str,
    ) -> GoogleDriveUploadResult:
        """Create the Drive file if missing, otherwise update the newest exact match."""

        if not Path(local_path).is_file():
            raise GoogleDriveReportStoreError(
                f"Local report file does not exist for Google Drive upload: {local_path}",
                issue_type="GDRIVE_LOCAL_REPORT_MISSING",
                classification="configuration",
            )
        matches = self.find_files_by_name(folder_id, drive_file_name)
        duplicate_warning = ""
        if len(matches) > 1:
            duplicate_warning = (
                f"{len(matches)} Google Drive files named {drive_file_name!r} exist in the folder; "
                "updated the most recently modified file and left duplicates untouched."
            )
        media = _media_file_upload(local_path)
        try:
            if matches:
                selected = matches[0]
                response = (
                    self._service.files()
                    .update(
                        fileId=selected.file_id,
                        body={"name": drive_file_name},
                        media_body=media,
                        fields=DRIVE_FILE_FIELDS,
                        supportsAllDrives=True,
                    )
                    .execute()
                )
                action = "updated"
            else:
                response = (
                    self._service.files()
                    .create(
                        body={"name": drive_file_name, "parents": [folder_id]},
                        media_body=media,
                        fields=DRIVE_FILE_FIELDS,
                        supportsAllDrives=True,
                    )
                    .execute()
                )
                action = "created"
        except Exception as exc:
            raise _drive_api_error(exc, f"Could not upload {drive_file_name!r} to Google Drive.") from exc
        return GoogleDriveUploadResult(
            artifact=artifact,
            local_path=Path(local_path),
            drive_file_name=drive_file_name,
            action=action,
            metadata=GoogleDriveFileMetadata.from_api(response),
            duplicate_count=len(matches),
            warning=duplicate_warning,
        )

    def download_file_by_id(self, file_id: str) -> tuple[bytes, GoogleDriveFileMetadata]:
        """Download a Drive binary file by ID."""

        _require_value(file_id, "Google Drive file ID")
        try:
            metadata_payload = (
                self._service.files()
                .get(fileId=file_id, fields=DRIVE_FILE_FIELDS, supportsAllDrives=True)
                .execute()
            )
            request = self._service.files().get_media(fileId=file_id, supportsAllDrives=True)
            content = _download_media_request(request)
        except Exception as exc:
            raise _drive_api_error(exc, f"Could not download Google Drive file {file_id!r}.") from exc
        return content, GoogleDriveFileMetadata.from_api(metadata_payload)

    def download_file_by_name(self, folder_id: str, file_name: str) -> tuple[bytes, GoogleDriveFileMetadata]:
        """Download the newest exact filename match inside one Drive folder."""

        matches = self.find_files_by_name(folder_id, file_name)
        if not matches:
            raise GoogleDriveReportStoreError(
                f"Google Drive report file was not found in the configured folder: {file_name}",
                issue_type="GDRIVE_REPORT_FILE_MISSING",
                classification="not_found",
            )
        return self.download_file_by_id(matches[0].file_id)


def upload_inventory_cover_reports_to_drive(
    result: InventoryCoverPipelineRunResult,
    config: GoogleDriveReportConfig,
    *,
    store: GoogleDriveReportStore | None = None,
) -> GoogleDriveUploadSummary:
    """Upload the latest Inventory Cover report and optional audit workbook."""

    resolved = config.resolved()
    notifications_dir = result.run_dir / "notifications"
    audit_file = notifications_dir / "google_drive_upload.json"
    log_file = result.run_dir / "logs" / "google_drive_upload.log"
    logger = setup_run_logger(log_file, logger_name=f"inventory_cover.google_drive.{result.run_id}")
    started_monotonic = time.monotonic()
    started_at = _now_iso()
    uploads: list[GoogleDriveUploadResult] = []
    warnings: list[str] = []

    try:
        if not resolved.enabled:
            summary = GoogleDriveUploadSummary(
                run_id=result.run_id,
                enabled=False,
                status="DISABLED",
                audit_file=audit_file,
                log_file=log_file,
                started_at=started_at,
                completed_at=_now_iso(),
                duration_seconds=round(time.monotonic() - started_monotonic, 3),
            )
            _write_upload_audit(summary, logger=logger)
            return summary

        _require_value(resolved.folder_id, "GDRIVE_FOLDER_ID")
        active_store = store or GoogleDriveReportStore.from_config(resolved)
        logger.info(
            "Google Drive upload requested. folder_id=%s report_file=%s upload_audit=%s fail_on_error=%s",
            resolved.folder_id,
            resolved.report_file_name,
            resolved.upload_audit,
            resolved.fail_on_upload_error,
        )
        uploads.append(
            active_store.upload_or_update_file(
                result.team_latest_file,
                folder_id=resolved.folder_id,
                drive_file_name=resolved.report_file_name,
                artifact="inventory_cover_report",
            )
        )
        if resolved.upload_audit:
            uploads.append(
                active_store.upload_or_update_file(
                    result.backend_latest_file,
                    folder_id=resolved.folder_id,
                    drive_file_name=resolved.audit_file_name,
                    artifact="backend_audit",
                )
            )
        for upload in uploads:
            if upload.warning:
                warnings.append(upload.warning)
                logger.warning(upload.warning)
            logger.info(
                "Google Drive %s %s. file_id=%s name=%s modified_time=%s size=%s",
                upload.artifact,
                upload.action,
                upload.metadata.file_id,
                upload.metadata.name,
                upload.metadata.modified_time,
                upload.metadata.size,
            )
        summary = GoogleDriveUploadSummary(
            run_id=result.run_id,
            enabled=True,
            status="SUCCESS",
            audit_file=audit_file,
            log_file=log_file,
            started_at=started_at,
            completed_at=_now_iso(),
            duration_seconds=round(time.monotonic() - started_monotonic, 3),
            uploads=tuple(uploads),
            warnings=tuple(warnings),
        )
        _write_upload_audit(summary, logger=logger)
        return summary
    except Exception as exc:
        if isinstance(exc, GoogleDriveReportStoreError):
            error_type = exc.issue_type
            classification = exc.classification
        else:
            error_type = exc.__class__.__name__
            classification = "api"
        sanitized = _classify_exception_message(exc)
        logger.warning("Google Drive upload failed. error_type=%s error=%s", error_type, sanitized)
        summary = GoogleDriveUploadSummary(
            run_id=result.run_id,
            enabled=resolved.enabled,
            status="FAILED",
            audit_file=audit_file,
            log_file=log_file,
            started_at=started_at,
            completed_at=_now_iso(),
            duration_seconds=round(time.monotonic() - started_monotonic, 3),
            uploads=tuple(uploads),
            error_type=error_type,
            error_message_sanitized=sanitized,
            warnings=(f"Google Drive upload failed with classification {classification}.",),
        )
        _write_upload_audit(summary, logger=logger)
        if resolved.fail_on_upload_error:
            raise GoogleDriveReportStoreError(
                f"{error_type}: {sanitized}",
                issue_type=error_type,
                classification=classification,
            ) from exc
        return summary


def _write_upload_audit(summary: GoogleDriveUploadSummary, *, logger: logging.Logger) -> None:
    write_json_file(summary.audit_file, summary.as_json())
    logger.info("Google Drive upload audit written: %s", summary.audit_file)


def _normalize_service_account_info(info: Mapping[str, Any] | str) -> dict[str, Any]:
    if isinstance(info, str):
        parsed = json.loads(info)
    else:
        parsed = dict(info)
    if "private_key" in parsed and isinstance(parsed["private_key"], str):
        parsed["private_key"] = parsed["private_key"].replace("\\n", "\n")
    return parsed


def _media_file_upload(path: Path) -> Any:
    try:
        from googleapiclient.http import MediaFileUpload
    except ImportError as exc:
        raise GoogleDriveReportStoreError(
            "Google Drive dependencies are not installed. Run `python -m pip install -e .` first.",
            issue_type="GDRIVE_DEPENDENCY_MISSING",
            classification="dependency",
        ) from exc
    return MediaFileUpload(str(path), mimetype=EXCEL_MIME_TYPE, resumable=False)


def _download_media_request(request: Any) -> bytes:
    try:
        from googleapiclient.http import MediaIoBaseDownload
    except ImportError as exc:
        raise GoogleDriveReportStoreError(
            "Google Drive dependencies are not installed. Run `python -m pip install -e .` first.",
            issue_type="GDRIVE_DEPENDENCY_MISSING",
            classification="dependency",
        ) from exc
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buffer.getvalue()


def _drive_api_error(exc: Exception, prefix: str) -> GoogleDriveReportStoreError:
    status = getattr(getattr(exc, "resp", None), "status", None)
    reason = _google_error_reason(exc)
    classification = "storage_quota" if reason == "storageQuotaExceeded" else _classify_google_status(status)
    issue_type = {
        "authentication": "GDRIVE_AUTHENTICATION_FAILED",
        "access": "GDRIVE_ACCESS_DENIED",
        "not_found": "GDRIVE_FOLDER_OR_FILE_NOT_FOUND",
        "quota": "GDRIVE_API_QUOTA_FAILURE",
        "storage_quota": "GDRIVE_STORAGE_QUOTA_EXCEEDED",
        "transport": "GDRIVE_TRANSPORT_FAILURE",
    }.get(classification, "GDRIVE_API_FAILURE")
    return GoogleDriveReportStoreError(
        f"{prefix} {classification.upper()}: {_classify_exception_message(exc)}",
        issue_type=issue_type,
        classification=classification,
    )


def _classify_google_status(status: Any) -> str:
    if status in {401}:
        return "authentication"
    if status in {403}:
        return "access"
    if status in {404}:
        return "not_found"
    if status in {429}:
        return "quota"
    if isinstance(status, int) and status >= 500:
        return "transport"
    return "api"


def _google_error_reason(exc: Exception) -> str:
    content = getattr(exc, "content", b"")
    if isinstance(content, bytes):
        content = content.decode("utf-8", errors="replace")
    if not isinstance(content, str) or not content.strip():
        return ""
    try:
        payload = json.loads(content)
    except json.JSONDecodeError:
        return ""
    errors = payload.get("error", {}).get("errors", [])
    if errors and isinstance(errors[0], Mapping):
        return str(errors[0].get("reason") or "")
    return str(payload.get("error", {}).get("reason") or "")


def _require_value(value: str, name: str) -> None:
    if not str(value or "").strip():
        raise GoogleDriveReportStoreError(
            f"{name} is required for Google Drive report storage.",
            issue_type="GDRIVE_CONFIG_MISSING",
            classification="configuration",
        )


def _escape_drive_query_literal(value: str) -> str:
    return str(value).replace("\\", "\\\\").replace("'", "\\'")


def _classify_exception_message(exc: Exception) -> str:
    text = str(exc)
    if "private_key" in text:
        return "Google Drive authentication failed with a private key related error."
    return text


def _now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")
