#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PDF批量转Word - 命令行入口

- RotatingFileHandler 限制日志体积(2MB x 5)
- 按转换结果返回进程退出码(0=全部成功)
"""

import logging
import os
import sys
from logging.handlers import RotatingFileHandler

from pdf_to_word_converter import PDFToWordConverter

APP_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(APP_DIR, "logs")
LOG_FILE = os.path.join(LOG_DIR, "converter.log")
LOG_MAX_BYTES = 2 * 1024 * 1024
LOG_BACKUP_COUNT = 5


def setup_logging():
    """配置日志: 控制台 + 轮转文件"""
    os.makedirs(LOG_DIR, exist_ok=True)

    logger = logging.getLogger("converter")
    logger.setLevel(logging.INFO)
    if logger.handlers:
        return logger, LOG_FILE

    fmt = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler = RotatingFileHandler(
        LOG_FILE, maxBytes=LOG_MAX_BYTES, backupCount=LOG_BACKUP_COUNT, encoding='utf-8'
    )
    file_handler.setFormatter(fmt)
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(fmt)
    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)

    # pdf2docx 日志降噪: 只保留警告以上
    logging.getLogger("pdf2docx").setLevel(logging.WARNING)
    return logger, LOG_FILE


def main():
    try:
        if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

    logger, log_file = setup_logging()
    logger.info("开始执行PDF转Word转换程序...")
    logger.info(f"日志文件: {log_file}")

    input_dir = os.path.join(APP_DIR, "data")
    output_dir = os.path.join(APP_DIR, "output")

    if not os.path.exists(input_dir):
        os.makedirs(input_dir)
        logger.warning(f"输入目录不存在, 已创建: {input_dir}, 请将PDF文件放入后重新运行")
        return 0

    converter = PDFToWordConverter(input_dir, output_dir, logger=logger)
    results = converter.convert_all()

    if not results:
        logger.warning(f"{input_dir} 目录中没有PDF文件, 未执行转换")
        return 0

    ok = [r for r in results if r['status'] == 'ok']
    failed = [r for r in results if r['status'] != 'ok']
    total_pages = sum(r.get('pages') or 0 for r in ok)
    warned = [r for r in ok if r.get('warning')]

    logger.info("-" * 60)
    logger.info(f"汇总: {len(ok)}/{len(results)} 个文件转换成功, 共 {total_pages} 页")
    if warned:
        logger.warning(f"其中 {len(warned)} 个文件有质量警告(疑似扫描件)")
    if failed:
        for r in failed:
            logger.error(f"失败: {r['file']} ({r['status']}: {r['error']})")
    logger.info(f"输出目录: {output_dir}, 日志文件: {log_file}")
    logger.info("转换程序执行完成!")

    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
