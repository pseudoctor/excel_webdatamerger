import logging
import os
from datetime import datetime
from pathlib import Path


def setup_logger(name="ExcelMerger"):
    """设置日志记录。日志文件不可写时退回到标准输出。"""
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    logger.propagate = False

    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    log_dir = Path(
        os.getenv(
            "MERGER_LOG_DIR",
            Path(__file__).resolve().parent.parent / "logs",
        )
    )
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
        log_path = log_dir / f"{datetime.now().strftime('%Y%m%d')}.log"
        file_handler = logging.FileHandler(log_path, encoding="utf-8")
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except OSError as exc:
        logger.warning("日志文件不可写，已退回标准输出: %s", exc)

    return logger
