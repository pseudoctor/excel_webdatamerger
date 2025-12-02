"""Flask web entrypoint for excel_datamerger."""
import os
import shutil
import traceback
from datetime import datetime
from functools import wraps
from pathlib import Path
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

    logger = setup_logger("ExcelMergerWeb")

    # In-memory task store: {task_id: {"path": Path, "created_at": datetime}}
    tasks = {}

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

    @app.route("/login", methods=["GET", "POST"])
    def login():
        error = None
        if request.method == "POST":
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "").strip()
            if (
                username == app.config["USERNAME"]
                and password == app.config["PASSWORD"]
            ):
                session["user"] = username
                return redirect(url_for("merge_page"))
            error = "用户名或密码错误"
            logger.warning("Login failed for user: %s", username)
        return render_template("login.html", error=error)

    @app.route("/logout")
    def logout():
        session.clear()
        return redirect(url_for("login"))

    @app.before_request
    def enforce_login():
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

        job_dir = app.config["UPLOAD_ROOT"] / Path(str(uuid4()))
        job_dir.mkdir(parents=True, exist_ok=True)
        saved_paths = []

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

                    # 列删除过滤
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
                cleanup_job_dir(job_dir)
                return jsonify({"ok": False, "error": "没有可合并的数据"}), 400

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

            task_id = str(uuid4())
            tasks[task_id] = {
                "path": output_path,
                "created_at": datetime.utcnow(),
                "format": output_format,
            }

            return jsonify(
                {
                    "ok": True,
                    "task_id": task_id,
                    "download_url": url_for(
                        "download_result", task_id=task_id
                    ),
                }
            )
        except Exception as exc:  # noqa: BLE001
            cleanup_job_dir(job_dir)
            logger.error("Merge failed: %s\n%s", exc, traceback.format_exc())
            return jsonify({"ok": False, "error": str(exc)}), 500

    @app.route("/download/<task_id>")
    @login_required
    def download_result(task_id: str):
        task = tasks.get(task_id)
        if not task:
            return "任务不存在或已过期", 404
        output_path: Path = task["path"]
        if not output_path.exists():
            return "文件不存在", 404

        fmt = task.get("format", "xlsx")
        filename = f"merged_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{fmt}"
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
                cleaned[k] = v
            cm.save_mappings(cleaned)
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
            logger.info("Cleanup logs removed %s files", removed)
        elif target == "temp":
            target_path = Path(app.config["UPLOAD_ROOT"])
            removed, errors = remove_path_contents(target_path, files_only=False)
            logger.info("Cleanup temp removed %s entries", removed)
        else:
            return jsonify({"ok": False, "error": "未知清理目标"}), 400

        return jsonify({"ok": True, "removed": removed, "errors": errors})

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


app = create_app()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8000")))
