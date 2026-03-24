"""Flask web entrypoint for excel_webdatamerger."""
from concurrent.futures import ThreadPoolExecutor
import json
import re
import os
import shutil
import threading
import traceback
from datetime import datetime, timezone
from functools import wraps
from pathlib import Path
from urllib.parse import urlparse
from uuid import uuid4
import math

import pandas as pd
from flask import (
    Flask,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.middleware.proxy_fix import ProxyFix

from excelmerger.config_manager import ConfigManager
from excelmerger.io_utils import read_file, save_file
from excelmerger.logger import setup_logger
from excelmerger.merger import ExcelMergerCore
from .config import WebConfig


def create_app() -> Flask:
    """Create and configure the Flask application."""
    app = Flask(
        __name__,
        template_folder="templates",
        static_folder="static",
    )
    app.config.from_object(WebConfig)
    app.secret_key = app.config["SECRET_KEY"]
    app.config["UPLOAD_ROOT"].mkdir(parents=True, exist_ok=True)

    # 关键：让 Flask 正确识别 Nginx 反向代理 + HTTPS
    app.wsgi_app = ProxyFix(
        app.wsgi_app,
        x_for=1,
        x_proto=1,
        x_host=1,
        x_port=1,
        x_prefix=1,
    )

    logger = setup_logger("ExcelMergerWeb")
    upload_root: Path = app.config["UPLOAD_ROOT"]
    metadata_lock = threading.Lock()
    merge_executor = ThreadPoolExecutor(max_workers=2)

    if app.config["USERNAME"] == "admin" or app.config["PASSWORD"] == "admin123":
        logger.warning("Using default web credentials is unsafe in production")
    if app.config["SECRET_KEY"] == "replace-this-secret":
        logger.warning("Using default Flask SECRET_KEY is unsafe in production")

    def is_safe_next_url(target: str) -> bool:
        if not target:
            return False
        parsed = urlparse(target)
        return not parsed.netloc and target.startswith("/")

    def purge_expired_tasks() -> None:
        now = datetime.now(timezone.utc)
        for job_dir in upload_root.iterdir():
            if not job_dir.is_dir():
                continue
            metadata = load_task_metadata(job_dir.name)
            created_at = get_task_expiry_reference(job_dir, metadata)
            if created_at is None:
                continue
            age_seconds = (now - created_at).total_seconds()
            if age_seconds > app.config["CLEANUP_MINUTES"] * 60:
                cleanup_job_dir(job_dir)

    def login_required(func):
        """Simple login-required decorator."""

        @wraps(func)
        def wrapper(*args, **kwargs):
            if not session.get("user"):
                return redirect(url_for("login", next=request.path))
            return func(*args, **kwargs)

        return wrapper

    def allowed_file(filename: str) -> bool:
        return Path(filename).suffix.lower() in app.config["ALLOWED_EXTENSIONS"]

    def cleanup_job_dir(path: Path) -> None:
        """Remove temporary job directory safely."""
        try:
            if path.exists():
                shutil.rmtree(path)
        except Exception as cleanup_err:
            logger.warning("Failed to cleanup job dir %s: %s", path, cleanup_err)

    def task_metadata_path(task_id: str) -> Path:
        return upload_root / task_id / "metadata.json"

    def parse_utc_datetime(value: str | None) -> datetime | None:
        if not value:
            return None
        try:
            parsed = datetime.fromisoformat(value)
            if parsed.tzinfo is None:
                return parsed.replace(tzinfo=timezone.utc)
            return parsed.astimezone(timezone.utc)
        except ValueError:
            return None

    def load_task_metadata(task_id: str) -> dict | None:
        metadata_path = task_metadata_path(task_id)
        if not metadata_path.exists():
            return None
        try:
            with metadata_lock:
                with metadata_path.open("r", encoding="utf-8") as fh:
                    loaded = json.load(fh)
                    if isinstance(loaded, dict):
                        return loaded
        except (OSError, json.JSONDecodeError) as exc:
            logger.warning("Failed to load task metadata for %s: %s", task_id, exc)
        return None

    def save_task_metadata(task_id: str, payload: dict) -> None:
        metadata_path = task_metadata_path(task_id)
        with metadata_lock:
            with metadata_path.open("w", encoding="utf-8") as fh:
                json.dump(payload, fh, ensure_ascii=False, indent=2)

    def update_task_metadata(task_id: str, **updates) -> dict | None:
        metadata = load_task_metadata(task_id)
        if metadata is None:
            return None
        metadata.update(updates)
        save_task_metadata(task_id, metadata)
        return metadata

    def get_task_expiry_reference(job_dir: Path, metadata: dict | None) -> datetime | None:
        """Return the best available UTC timestamp for task cleanup decisions."""
        created_at = parse_utc_datetime((metadata or {}).get("created_at"))
        if created_at is not None:
            return created_at
        try:
            return datetime.fromtimestamp(job_dir.stat().st_mtime, tz=timezone.utc)
        except OSError as exc:
            logger.warning("Failed to stat job dir %s: %s", job_dir, exc)
            return None

    def build_default_download_stem() -> str:
        return f"merged_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

    def sanitize_download_name(raw_name: str, fmt: str) -> str:
        cleaned = re.sub(r'[\x00-\x1f\x7f<>:"/\\|?*]+', "_", raw_name or "")
        cleaned = cleaned.strip().strip(".")
        if not cleaned:
            return f"{build_default_download_stem()}.{fmt}"
        normalized = cleaned.lower()
        known_suffixes = {".csv", ".xlsx", ".xls", ".txt"}
        matched_suffix = next(
            (suffix for suffix in known_suffixes if normalized.endswith(suffix)),
            "",
        )
        if matched_suffix:
            stem = cleaned[: -len(matched_suffix)]
        else:
            stem = cleaned
        stem = stem.strip().strip("._- ")
        if not stem:
            return f"{build_default_download_stem()}.{fmt}"
        return f"{stem}.{fmt}"

    def build_task_status_payload(task_id: str) -> tuple[dict, int]:
        metadata = load_task_metadata(task_id)
        if not metadata:
            return {"ok": False, "error": "任务不存在或已过期"}, 404

        status = metadata.get("status", "unknown")
        payload = {
            "ok": status != "failed",
            "task_id": task_id,
            "status": status,
            "suggested_filename": metadata.get("suggested_filename", ""),
            "format": metadata.get("format", "xlsx"),
        }
        if status == "completed":
            payload["download_url"] = url_for("download_result", task_id=task_id)
        if status == "failed":
            payload["error"] = metadata.get("error", "合并失败")
        return payload, 200

    def process_merge_task(
        task_id: str,
        saved_paths: list[Path],
        *,
        normalize_columns: bool,
        enable_fuzzy: bool,
        remove_duplicates: bool,
        smart_dedup: bool,
        dedup_keys: list[str],
        exclude_columns: set[str],
        output_format: str,
    ) -> None:
        job_dir = upload_root / task_id
        update_task_metadata(task_id, status="running", started_at=datetime.now(timezone.utc).isoformat())

        try:
            config_manager = ConfigManager()
            merger = ExcelMergerCore(config_manager)

            all_dfs = []
            mapping_report = {}

            for file_path in saved_paths:
                sheets = read_file(str(file_path))
                for sheet_name, df in sheets.items():
                    if df.empty:
                        logger.info(
                            "Skip empty sheet %s - %s", file_path.name, sheet_name
                        )
                        continue

                    if normalize_columns:
                        df = merger.normalize_columns(
                            df, enable_fuzzy=enable_fuzzy
                        )
                        current_mapping = merger.get_mapping_report()
                        if current_mapping:
                            mapping_report[
                                f"{file_path.name}-{sheet_name}"
                            ] = current_mapping

                    filename_without_ext = file_path.stem
                    df.insert(0, "来源文件", filename_without_ext)
                    df.insert(1, "工作表", sheet_name)

                    if exclude_columns:
                        cols_to_keep = [
                            c
                            for c in df.columns
                            if str(c) not in exclude_columns
                            or str(c) in {"来源文件", "工作表"}
                        ]
                        if len(cols_to_keep) < len(df.columns):
                            df = df[cols_to_keep]

                    all_dfs.append(df)

            if not all_dfs:
                raise ValueError("没有可合并的数据")

            merged = pd.concat(all_dfs, join="outer", ignore_index=True, sort=False)
            logger.info(
                "Merged %s files into %s rows x %s cols",
                len(saved_paths),
                len(merged),
                len(merged.columns),
            )

            original_count = len(merged)
            if smart_dedup and dedup_keys:
                merged = merger.deduplicate_smart(
                    merged, key_columns=dedup_keys
                )
                removed = original_count - len(merged)
                if removed > 0:
                    logger.info("Smart dedup removed %s rows", removed)
            elif remove_duplicates:
                merged = merger.deduplicate_smart(merged)
                removed = original_count - len(merged)
                if removed > 0:
                    logger.info("Full-row dedup removed %s rows", removed)

            quality_report = merger.validate_data(merged)
            logger.info("Quality report: %s", quality_report)
            if mapping_report:
                logger.info("Column mapping: %s", mapping_report)

            output_path = job_dir / f"merged.{output_format}"
            save_file(merged, output_path, file_format=output_format)

            update_task_metadata(
                task_id,
                status="completed",
                path=output_path.name,
                completed_at=datetime.now(timezone.utc).isoformat(),
                error="",
            )
        except Exception as exc:  # noqa: BLE001
            logger.error("Merge failed: %s\n%s", exc, traceback.format_exc())
            update_task_metadata(
                task_id,
                status="failed",
                error=str(exc),
                completed_at=datetime.now(timezone.utc).isoformat(),
            )

    @app.route("/login", methods=["GET", "POST"])
    def login():
        error = None
        next_url = request.args.get("next", "")
        if request.method == "POST":
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "").strip()
            next_url = request.form.get("next", "")
            if (
                username == app.config["USERNAME"]
                and password == app.config["PASSWORD"]
            ):
                session["user"] = username
                if is_safe_next_url(next_url):
                    return redirect(next_url)
                return redirect(url_for("merge_page"))
            error = "用户名或密码错误"
            logger.warning("Login failed for user: %s", username)
        return render_template("login.html", error=error, next_url=next_url)

    @app.route("/logout")
    def logout():
        session.clear()
        return redirect(url_for("login"))

    @app.before_request
    def enforce_login():
        purge_expired_tasks()
        # Allow login page and static files without auth
        if request.endpoint in {"login", "static"}:
            return None
        if not session.get("user"):
            return redirect(url_for("login", next=request.path))
        return None

    @app.route("/")
    @login_required
    def merge_page():
        return render_template("merge.html", user=session.get("user"))

    @app.route("/merge", methods=["POST"])
    @login_required
    def merge_endpoint():
        # Basic payload size guard
        if (
            request.content_length
            and request.content_length > app.config["MAX_CONTENT_LENGTH"]
        ):
            return (
                jsonify({"ok": False, "error": "上传大小超出限制"}),
                413,
            )

        files = request.files.getlist("files")
        if not files:
            return jsonify({"ok": False, "error": "请至少上传一个文件"}), 400

        normalize_columns = request.form.get("normalize_columns") == "on"
        enable_fuzzy = request.form.get("enable_fuzzy") == "on"
        remove_duplicates = request.form.get("remove_duplicates") == "on"
        smart_dedup = request.form.get("smart_dedup") == "on"
        dedup_keys_raw = request.form.get("dedup_keys", "")
        dedup_keys = [k.strip() for k in dedup_keys_raw.split(",") if k.strip()]
        exclude_raw = request.form.get("exclude_columns", "")
        exclude_columns = {c.strip() for c in exclude_raw.split(",") if c.strip()}
        output_format = request.form.get("output_format", "xlsx").lower()
        if output_format not in {"xlsx", "csv"}:
            output_format = "xlsx"

        task_id = str(uuid4())
        job_dir = upload_root / task_id
        job_dir.mkdir(parents=True, exist_ok=True)
        saved_paths = []
        suggested_filename = build_default_download_stem()

        try:
            total_size = 0
            for f in files:
                if not f or not f.filename:
                    continue
                if not allowed_file(f.filename):
                    cleanup_job_dir(job_dir)
                    return (
                        jsonify(
                            {"ok": False, "error": f"不支持的文件类型: {f.filename}"}
                        ),
                        400,
                    )
                dest = job_dir / Path(f.filename).name
                f.save(dest)
                file_size = dest.stat().st_size
                total_size += file_size
                if total_size > app.config["MAX_CONTENT_LENGTH"]:
                    cleanup_job_dir(job_dir)
                    return (
                        jsonify({"ok": False, "error": "上传文件总大小超出限制"}),
                        413,
                    )
                saved_paths.append(dest)

            save_task_metadata(
                task_id,
                {
                    "created_at": datetime.now(timezone.utc).isoformat(),
                    "format": output_format,
                    "status": "queued",
                    "suggested_filename": suggested_filename,
                    "error": "",
                },
            )

            task_kwargs = {
                "normalize_columns": normalize_columns,
                "enable_fuzzy": enable_fuzzy,
                "remove_duplicates": remove_duplicates,
                "smart_dedup": smart_dedup,
                "dedup_keys": dedup_keys,
                "exclude_columns": exclude_columns,
                "output_format": output_format,
            }

            if app.config.get("MERGE_ASYNC", True):
                merge_executor.submit(
                    process_merge_task,
                    task_id,
                    list(saved_paths),
                    **task_kwargs,
                )
            else:
                process_merge_task(
                    task_id,
                    list(saved_paths),
                    **task_kwargs,
                )

            return jsonify(
                {
                    "ok": True,
                    "task_id": task_id,
                    "status": load_task_metadata(task_id).get("status", "queued"),
                    "suggested_filename": suggested_filename,
                    "status_url": url_for(
                        "task_status", task_id=task_id
                    ),
                }
            ), 202
        except Exception as exc:  # noqa: BLE001
            cleanup_job_dir(job_dir)
            logger.error("Merge failed: %s\n%s", exc, traceback.format_exc())
            status_code = 400 if isinstance(exc, ValueError) else 500
            return jsonify({"ok": False, "error": str(exc)}), status_code

    @app.route("/task/<task_id>")
    @login_required
    def task_status(task_id: str):
        payload, status_code = build_task_status_payload(task_id)
        return jsonify(payload), status_code

    @app.route("/download/<task_id>")
    @login_required
    def download_result(task_id: str):
        metadata = load_task_metadata(task_id)
        if not metadata:
            return "任务不存在或已过期", 404
        if metadata.get("status") != "completed":
            return "任务尚未完成", 409

        output_path = upload_root / task_id / metadata.get("path", "")
        if not output_path.exists():
            cleanup_job_dir(upload_root / task_id)
            return "文件不存在", 404

        fmt = metadata.get("format", "xlsx")
        requested_name = request.args.get("filename", "")
        filename = sanitize_download_name(requested_name, fmt)
        if fmt == "csv":
            mimetype = "text/csv"
        else:
            mimetype = (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        return send_file(
            output_path,
            as_attachment=True,
            download_name=filename,
            mimetype=mimetype,
        )

    @app.after_request
    def disable_cache(response):
        """Disable caching to avoid stale HTML/JS during rapid iterations."""
        response.headers["Cache-Control"] = "no-store"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response

    def sanitize_json(obj):
        """Recursively replace NaN/inf with None for strict JSON."""
        if isinstance(obj, float):
            if math.isnan(obj) or math.isinf(obj):
                return None
            return obj
        if isinstance(obj, dict):
            return {k: sanitize_json(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [sanitize_json(v) for v in obj]
        return obj

    def remove_path_contents(path: Path, files_only: bool = False):
        """Delete files (and optionally subdirectories) under a path."""
        removed = 0
        errors = []
        if not path.exists():
            return removed, errors
        for item in path.iterdir():
            try:
                if item.is_dir():
                    if files_only:
                        continue
                    shutil.rmtree(item)
                else:
                    item.unlink()
                removed += 1
            except Exception as exc:  # noqa: BLE001
                errors.append(f"{item}: {exc}")
        return removed, errors

    def cleanup_expired_job_dirs():
        """Delete only expired task directories to avoid racing active downloads."""
        removed = 0
        skipped = 0
        errors = []
        now = datetime.now(timezone.utc)
        for item in upload_root.iterdir():
            if not item.is_dir():
                continue
            metadata = load_task_metadata(item.name)
            created_at = get_task_expiry_reference(item, metadata)
            if created_at is None:
                skipped += 1
                continue
            age_seconds = (now - created_at).total_seconds()
            if age_seconds <= app.config["CLEANUP_MINUTES"] * 60:
                skipped += 1
                continue
            try:
                shutil.rmtree(item)
                removed += 1
            except Exception as exc:  # noqa: BLE001
                errors.append(f"{item}: {exc}")
        return removed, skipped, errors

    @app.route("/mapping", methods=["GET", "POST"])
    @login_required
    def mapping_endpoint():
        """Fetch or update column mapping configuration."""
        cm = ConfigManager()
        if request.method == "GET":
            return jsonify({"ok": True, "mappings": cm.get_mappings()})

        try:
            payload = request.get_json(force=True)
            mappings = payload.get("mappings") if isinstance(payload, dict) else None
            if not isinstance(mappings, dict):
                return jsonify({"ok": False, "error": "映射格式应为对象（键为标准列名，值为别名列表）"}), 400
            # 确保值为列表
            cleaned = {}
            for k, v in mappings.items():
                if not isinstance(v, list):
                    return jsonify({"ok": False, "error": f"映射 {k} 的值必须是列表"}), 400
                cleaned[str(k)] = [str(alias) for alias in v if str(alias).strip()]
            if not cm.save_mappings(cleaned):
                return jsonify({"ok": False, "error": "映射保存失败，请检查文件权限"}), 500
            return jsonify({"ok": True})
        except Exception as exc:  # noqa: BLE001
            logger.error("Save mapping failed: %s\n%s", exc, traceback.format_exc())
            return jsonify({"ok": False, "error": str(exc)}), 500

    @app.route("/cleanup", methods=["POST"])
    @login_required
    def cleanup_endpoint():
        """Clean logs or temporary directories."""
        target = request.form.get("target")
        if target == "logs":
            target_path = Path(__file__).resolve().parent.parent / "logs"
            removed, errors = remove_path_contents(target_path, files_only=True)
            skipped = 0
            logger.info("Cleanup logs removed %s files", removed)
        elif target == "temp":
            removed, skipped, errors = cleanup_expired_job_dirs()
            logger.info("Cleanup temp removed %s entries and skipped %s active entries", removed, skipped)
        else:
            return jsonify({"ok": False, "error": "未知清理目标"}), 400

        return jsonify({"ok": True, "removed": removed, "skipped": skipped, "errors": errors})

    @app.route("/inspect", methods=["POST"])
    @login_required
    def inspect_files():
        """Generate column list and preview rows for selected files."""
        if (
            request.content_length
            and request.content_length > app.config["MAX_CONTENT_LENGTH"]
        ):
            return jsonify({"ok": False, "error": "上传大小超出限制"}), 413

        files = request.files.getlist("files")
        if not files:
            return jsonify({"ok": False, "error": "请上传文件"}), 400

        normalize_columns = request.form.get("normalize_columns") == "on"
        enable_fuzzy = request.form.get("enable_fuzzy") == "on"

        job_dir = app.config["UPLOAD_ROOT"] / Path(str(uuid4()))
        job_dir.mkdir(parents=True, exist_ok=True)

        previews = []
        column_info = {}
        mapping_report = {}

        try:
            for f in files:
                if not f or not f.filename:
                    continue
                if not allowed_file(f.filename):
                    cleanup_job_dir(job_dir)
                    return (
                        jsonify({"ok": False, "error": f"不支持的文件类型: {f.filename}"}),
                        400,
                    )
                dest = job_dir / Path(f.filename).name
                f.save(dest)

            merger = ExcelMergerCore(ConfigManager())

            for file_path in job_dir.iterdir():
                sheets = read_file(str(file_path))
                for sheet_name, df in sheets.items():
                    if normalize_columns:
                        df = merger.normalize_columns(df, enable_fuzzy=enable_fuzzy)
                        current_mapping = merger.get_mapping_report()
                        if current_mapping:
                            mapping_report[
                                f"{file_path.name}-{sheet_name}"
                            ] = current_mapping

                    filename_without_ext = file_path.stem
                    df.insert(0, "来源文件", filename_without_ext)
                    df.insert(1, "工作表", sheet_name)

                    # 记录列信息（以映射后列名为准）
                    for col in df.columns:
                        col_key = str(col)
                        if col_key not in column_info:
                            column_info[col_key] = {
                                "sources": set(),
                                "is_meta": col_key in {"来源文件", "工作表"},
                            }
                        column_info[col_key]["sources"].add(
                            f"{file_path.name}-{sheet_name}"
                        )

                    preview_df = df.head(5).copy()
                    preview_df = preview_df.where(pd.notnull(preview_df), None)
                    previews.append(
                        {
                            "file": file_path.name,
                            "sheet": sheet_name,
                            "columns": list(df.columns),
                            "rows": preview_df.to_dict(orient="records"),
                        }
                    )

            # 转换 set -> list
            columns_payload = [
                {
                    "name": name,
                    "sources": sorted(info["sources"]),
                    "is_meta": info["is_meta"],
                }
                for name, info in sorted(column_info.items())
            ]

            return jsonify(
                {
                    "ok": True,
                    "columns": sanitize_json(columns_payload),
                    "previews": sanitize_json(previews),
                    "mapping": sanitize_json(mapping_report),
                }
            )
        except Exception as exc:  # noqa: BLE001
            logger.error("Inspect failed: %s\n%s", exc, traceback.format_exc())
            return jsonify({"ok": False, "error": str(exc)}), 500
        finally:
            cleanup_job_dir(job_dir)

    return app


# gunicorn 入口：web_app.app:app
app = create_app()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8000")))
