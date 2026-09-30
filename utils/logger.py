"""日志工具模块"""
import logging
import os
from datetime import datetime


def setup_logger(log_dir: str = "logs", level: str = "INFO") -> logging.Logger:
    os.makedirs(log_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join(log_dir, f"run_{timestamp}.log")

    logger = logging.getLogger("mail_audit")
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    # 连续运行时关闭旧文件句柄并移除旧 GUI/控制台 handler，避免日志重复和
    # Windows 上残留句柄导致日志文件无法移动或删除。
    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        try:
            handler.close()
        except Exception:
            pass

    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    logger.info(f"日志初始化完成, 日志文件: {log_file}")
    return logger


class GuiLogHandler(logging.Handler):
    """GUI 日志处理器，将日志发送到回调函数"""

    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def emit(self, record):
        try:
            msg = self.format(record)
            self.callback(msg)
        except Exception:
            pass
