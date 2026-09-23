#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PDF 转 Word 批量转换器

特性:
- 文本层 PDF: pdf2docx 保留版式(段落/表格/图片/字体样式), 文字可编辑
- 扫描件/图片型 PDF: 逐页渲染 + RapidOCR 离线识别, 生成纯文本可编辑 Word
- 混合 PDF(文字页+图片页): 逐页判断, 文字页保留版式, 图片页 OCR, 合并为单文档
- 单次打开完成页数统计、逐页文本层检测与转换
- 输出文件名批内防同名覆盖
- 加密/损坏/空文档给出明确错误
- 支持 on_progress 回调(供 GUI 逐文件进度更新)
- 日志只记录文件级信息(文件名/页数/耗时/错误), 不记录文档内容
"""

import logging
import os
import time

import fitz  # PyMuPDF
from pdf2docx import Converter

# OCR 渲染分辨率(DPI)与识别置信度下限
OCR_DPI = 200
OCR_MIN_SCORE = 0.45

_ocr_engine = None


def _get_ocr():
    """懒加载全局 OCR 引擎(进程内复用, 避免重复加载模型)"""
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
    return _ocr_engine


class PDFToWordConverter:
    """PDF转Word转换器"""

    def __init__(self, input_dir="data", output_dir="output", logger=None):
        """
        初始化转换器

        Args:
            input_dir: 输入PDF文件的目录
            output_dir: 输出Word文件的目录
            logger: 外部注入的logger(可选)
        """
        self.input_dir = input_dir
        self.output_dir = output_dir
        self.logger = logger or logging.getLogger(__name__)
        os.makedirs(output_dir, exist_ok=True)

    # ------------------------------------------------------------------
    # 对外接口
    # ------------------------------------------------------------------

    def convert_all(self, on_progress=None):
        """转换指定目录中的所有PDF文件, 返回每个文件的结果统计列表

        Args:
            on_progress: 可选回调, 签名 (file_result: dict, current: int, total: int)
                         每个PDF处理完成后调用, 用于GUI逐文件进度更新
        """
        pdf_files = self._get_pdf_files()

        if not pdf_files:
            self.logger.warning(f"在 {self.input_dir} 目录中未找到PDF文件")
            return []

        self.logger.info(f"找到 {len(pdf_files)} 个PDF文件")
        results = []
        used_names = set()
        for i, pdf_file in enumerate(pdf_files, 1):
            result = self._convert_single_pdf(pdf_file, used_names)
            results.append(result)
            if on_progress:
                on_progress(result, i, len(pdf_files))
        return results

    def convert_files(self, pdf_paths, on_progress=None, on_page_progress=None):
        """转换指定的PDF文件列表(供GUI按勾选范围转换)

        Args:
            on_progress: 文件级回调 (result, current, total)
            on_page_progress: 页级回调 (page_current, page_total), 仅OCR模式触发
        """
        pdf_paths = [p for p in pdf_paths if os.path.isfile(p)]
        if not pdf_paths:
            return []

        self.logger.info(f"开始转换 {len(pdf_paths)} 个PDF文件")
        results = []
        used_names = set()
        for i, pdf_file in enumerate(pdf_paths, 1):
            result = self._convert_single_pdf(
                pdf_file, used_names, on_page_progress=on_page_progress)
            results.append(result)
            if on_progress:
                on_progress(result, i, len(pdf_paths))
        return results

    # ------------------------------------------------------------------
    # 主流程
    # ------------------------------------------------------------------

    def _get_pdf_files(self):
        """获取输入目录中的所有PDF文件"""
        if not os.path.exists(self.input_dir):
            return []
        pdf_files = [
            os.path.join(self.input_dir, f)
            for f in os.listdir(self.input_dir)
            if f.lower().endswith('.pdf')
        ]
        return sorted(pdf_files)

    def _convert_single_pdf(self, pdf_path, used_names=None, on_page_progress=None):
        """
        转换单个PDF文件

        Returns:
            结果统计 dict: file/status/pages/output/warning/error/elapsed/method
        """
        filename = os.path.basename(pdf_path)
        self.logger.info(f"正在处理: {filename}")
        started = time.time()

        # 打开阶段: 页数统计 + 文本层检测 + 加密/损坏检查
        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            self.logger.error(f"无法打开PDF {filename}: {type(e).__name__}: {e}")
            return self._result(filename, 'error', error=f"{type(e).__name__}: {e}")

        try:
            with doc:
                if doc.needs_pass:
                    self.logger.error(f"PDF已加密, 需要密码: {filename}")
                    return self._result(filename, 'error', error='加密PDF, 请先解除密码')
                if doc.page_count == 0:
                    self.logger.error(f"PDF无页面: {filename}")
                    return self._result(filename, 'error', error='空文档')
                pages = doc.page_count
                # 逐页文本层检测: 区分 纯文本(layout) / 纯扫描(ocr) / 混合(hybrid)
                text_flags = [
                    bool(doc.load_page(i).get_text().strip())
                    for i in range(pages)
                ]
                n_text = sum(text_flags)

                # 输出路径(批内防同名覆盖)
                output_filename = self._generate_output_filename(filename)
                if used_names is not None:
                    base, ext = os.path.splitext(output_filename)
                    n = 2
                    while output_filename in used_names:
                        output_filename = f"{base}({n}){ext}"
                        n += 1
                    used_names.add(output_filename)
                output_path = os.path.join(self.output_dir, output_filename)

                warning = None
                if n_text == pages:
                    # 全部有文本层: pdf2docx 版式转换(doc 仅用于预检)
                    method = 'layout'
                elif n_text == 0:
                    # 全部无文本层(扫描件): 必须在 doc 关闭前完成逐页 OCR
                    method = 'ocr'
                    warning = '扫描件已OCR识别为可编辑文字, 个别字符可能识别偏差'
                    self.logger.warning(
                        f"{filename}: 无文本层(扫描件), 启用OCR离线识别生成可编辑文字")

                    try:
                        ok = self._convert_with_ocr(
                            doc, output_path, filename,
                            on_page=on_page_progress)
                    except SystemExit:
                        raise
                    except Exception as e:
                        self.logger.error(f"{filename}: 转换失败: {type(e).__name__}: {e}")
                        return self._result(filename, 'error', pages=pages,
                                            warning=warning,
                                            error=f"{type(e).__name__}: {e}")
                    if not ok:
                        self.logger.error(f"{filename}: OCR未识别到任何文字")
                        self._cleanup(output_path)
                        return self._result(filename, 'error', pages=pages,
                                            warning=warning,
                                            error='OCR未识别到文字, 无法生成可编辑文档')
                else:
                    # 混合文档: 文字页保留版式, 图片页逐页 OCR 成可编辑文字
                    method = 'hybrid'
                    warning = ('混合文档(文字页+图片页): 图片页已OCR识别为可编辑文字, '
                               '个别字符可能识别偏差')
                    self.logger.warning(
                        f"{filename}: 混合文档({n_text}/{pages}页含文本层), "
                        f"文字页走版式, 图片页走OCR")

                    try:
                        ok = self._convert_hybrid(
                            doc, pdf_path, output_path, filename,
                            on_page=on_page_progress)
                    except SystemExit:
                        raise
                    except Exception as e:
                        self.logger.error(f"{filename}: 转换失败: {type(e).__name__}: {e}")
                        return self._result(filename, 'error', pages=pages,
                                            warning=warning,
                                            error=f"{type(e).__name__}: {e}")
                    if not ok:
                        self.logger.error(f"{filename}: 未能生成有效内容")
                        self._cleanup(output_path)
                        return self._result(filename, 'error', pages=pages,
                                            warning=warning,
                                            error='未能生成有效内容')
        except Exception as e:
            if isinstance(e, SystemExit):
                raise
            self.logger.error(f"读取PDF {filename}: {type(e).__name__}: {e}")
            return self._result(filename, 'error', error=f"{type(e).__name__}: {e}")

        # 文本层 PDF 在 doc 关闭后再调 pdf2docx(按路径独立打开)
        if method == 'layout':
            try:
                ok = self._convert_with_pdf2docx(pdf_path, output_path)
                if not ok:
                    self.logger.error(f"{filename}: 未生成有效的Word文件")
                    self._cleanup(output_path)
                    return self._result(filename, 'error', pages=pages,
                                        error='未生成有效的Word文件')
            except SystemExit:
                raise
            except Exception as e:
                self.logger.error(f"{filename}: 转换失败: {type(e).__name__}: {e}")
                return self._result(filename, 'error', pages=pages,
                                    error=f"{type(e).__name__}: {e}")

        if not os.path.isfile(output_path) or os.path.getsize(output_path) == 0:
            self.logger.error(f"{filename}: 未生成有效的Word文件")
            self._cleanup(output_path)
            return self._result(filename, 'error', pages=pages, warning=warning,
                                error='未生成有效的Word文件')

        elapsed = time.time() - started
        size_kb = os.path.getsize(output_path) / 1024
        method_note = {'ocr': 'OCR', 'layout': '版式', 'hybrid': '混合'}[method]
        self.logger.info(
            f"成功转换: {output_filename} ({pages}页, {size_kb:.0f}KB, "
            f"{method_note}模式, 耗时{elapsed:.1f}s)"
            + (f" [警告: {warning}]" if warning else "")
        )
        return self._result(filename, 'ok', pages=pages, output=output_filename,
                            warning=warning, elapsed=round(elapsed, 1), method=method)

    # ------------------------------------------------------------------
    # 文本层 PDF: pdf2docx 版式转换(文字可编辑)
    # ------------------------------------------------------------------

    @staticmethod
    def _convert_with_pdf2docx(pdf_path, output_path):
        cv = Converter(pdf_path)
        try:
            cv.convert(output_path, start=0, end=None)
        finally:
            try:
                cv.close()
            except Exception:
                pass
        return os.path.isfile(output_path) and os.path.getsize(output_path) > 0

    # ------------------------------------------------------------------
    # 扫描件 PDF: 渲染 + OCR -> 纯文本可编辑 Word
    # ------------------------------------------------------------------

    def _convert_with_ocr(self, doc, output_path, filename, on_page=None):
        from docx import Document
        from docx.enum.text import WD_BREAK
        from docx.oxml.ns import qn
        from docx.shared import Pt

        ocr = _get_ocr()
        document = Document()

        # 默认字体: 中英文均可正常显示编辑
        style = document.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)
        style.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

        total_pages = doc.page_count
        total_lines = 0
        for i in range(total_pages):
            page = doc.load_page(i)
            pix = page.get_pixmap(dpi=OCR_DPI)
            raw = ocr(pix.tobytes('png'))
            result = raw[0] if isinstance(raw, tuple) else raw
            lines = self._normalize_ocr_lines(result)
            paragraphs = self._lines_to_paragraphs(lines)
            self.logger.info(f"{filename}: 第 {i + 1}/{total_pages} 页 OCR "
                             f"{len(paragraphs)} 段")

            if i > 0:
                p = document.add_paragraph()
                p.add_run().add_break(WD_BREAK.PAGE)

            if paragraphs:
                for text in paragraphs:
                    document.add_paragraph(text)
                total_lines += len(paragraphs)
            else:
                document.add_paragraph('')

            if on_page:
                on_page(i + 1, total_pages)

        document.save(output_path)
        return total_lines > 0

    # ------------------------------------------------------------------
    # 混合文档: 文字页 pdf2docx 保留版式, 图片页 OCR -> 合并为单文档
    # ------------------------------------------------------------------

    def _convert_hybrid(self, doc, pdf_path, output_path, filename, on_page=None):
        """逐页判断: 有文本层的页走 pdf2docx 版式, 纯图片页走 OCR

        输出一个 docx, 保证每一页都是可编辑内容; 纯图片页 OCR 无结果时
        退回嵌入该页原图, 避免照片页丢失。
        """
        from docx import Document
        from docx.enum.text import WD_BREAK
        from docx.oxml.ns import qn
        from docx.shared import Pt

        ocr = _get_ocr()
        document = Document()

        # 默认字体: 中英文均可正常显示编辑
        style = document.styles['Normal']
        style.font.name = 'Calibri'
        style.font.size = Pt(11)
        style.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

        total_pages = doc.page_count
        text_flags = [
            bool(doc.load_page(i).get_text().strip())
            for i in range(total_pages)
        ]

        # 连续同类型页聚成段, 便于 pdf2docx 整段转换
        segments = []  # [is_text, [page_idx...]]
        for i, f in enumerate(text_flags):
            if segments and segments[-1][0] == f:
                segments[-1][1].append(i)
            else:
                segments.append([f, [i]])

        total_lines = 0
        first_content = True
        for is_text, indices in segments:
            if is_text:
                if not first_content:
                    p = document.add_paragraph()
                    p.add_run().add_break(WD_BREAK.PAGE)
                self._append_pdf2docx_range(
                    pdf_path, document, indices[0], indices[-1])
                first_content = False
                if on_page:
                    on_page(indices[-1] + 1, total_pages)
            else:
                for pi in indices:
                    if not first_content:
                        p = document.add_paragraph()
                        p.add_run().add_break(WD_BREAK.PAGE)
                    paragraphs = self._ocr_page_to_paragraphs(
                        doc.load_page(pi), ocr)
                    if paragraphs:
                        for text in paragraphs:
                            document.add_paragraph(text)
                        total_lines += len(paragraphs)
                    else:
                        self._embed_page_image(document, doc.load_page(pi))
                    first_content = False
                    if on_page:
                        on_page(pi + 1, total_pages)

        document.save(output_path)
        return total_lines > 0

    def _append_pdf2docx_range(self, pdf_path, target_doc, start, end):
        """把 PDF 的 [start, end] 页(含)用 pdf2docx 转换后追加进 target_doc

        处理: 跳过源文档 sectPr、重映射内嵌图片与超链接关系。
        """
        import copy
        import io as _io
        import os
        import tempfile

        from docx import Document
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.oxml.ns import qn
        from pdf2docx import Converter

        fd, tmp = tempfile.mkstemp(suffix='.docx')
        os.close(fd)
        try:
            cv = Converter(pdf_path)
            try:
                cv.convert(tmp, start=start, end=end + 1)
            finally:
                cv.close()

            src = Document(tmp)
            src_body = src.element.body
            dst_body = target_doc.element.body

            # 图片关系重映射: 源 r:embed -> 目标文档中重新注册的 rId
            rel_map = {}
            blips = src_body.findall('.//' + qn('a:blip'))
            for blip in blips:
                rid = blip.get(qn('r:embed'))
                if not rid or rid in rel_map:
                    continue
                part = src.part.related_parts.get(rid)
                if part is None:
                    continue
                new_rid, _ = target_doc.part.get_or_add_image(
                    _io.BytesIO(part.blob))
                rel_map[rid] = new_rid
            for blip in blips:
                old = blip.get(qn('r:embed'))
                if old and old in rel_map:
                    blip.set(qn('r:embed'), rel_map[old])

            # 超链接关系重映射
            for hyperlink in src_body.findall('.//' + qn('w:hyperlink')):
                rid = hyperlink.get(qn('r:id'))
                if not rid:
                    continue
                rel = src.part.rels.get(rid)
                if rel is not None and rel.is_external:
                    new_rid = target_doc.part.relate_to(
                        rel.target_ref, RT.HYPERLINK, is_external=True)
                    hyperlink.set(qn('r:id'), new_rid)

            # 追加正文(跳过源 sectPr, 目标文档保持单一节)
            sect_pr = dst_body.find(qn('w:sectPr'))
            for child in list(src_body):
                if child.tag == qn('w:sectPr'):
                    continue
                node = copy.deepcopy(child)
                if sect_pr is not None:
                    sect_pr.addprevious(node)
                else:
                    dst_body.append(node)
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass

    @staticmethod
    def _ocr_page_to_paragraphs(page, ocr):
        """对单页渲染 + OCR, 返回重建后的段落列表"""
        pix = page.get_pixmap(dpi=OCR_DPI)
        raw = ocr(pix.tobytes('png'))
        result = raw[0] if isinstance(raw, tuple) else raw
        lines = PDFToWordConverter._normalize_ocr_lines(result)
        return PDFToWordConverter._lines_to_paragraphs(lines)

    @staticmethod
    def _embed_page_image(document, page):
        """OCR 无结果时把整页渲染图嵌入 docx, 避免照片页丢失"""
        import io as _io

        from docx.shared import Inches

        pix = page.get_pixmap(dpi=OCR_DPI)
        buf = _io.BytesIO(pix.tobytes('png'))
        usable_in = max((page.rect.width / 72.0) - 2.0, 1.0)
        p = document.add_paragraph()
        p.add_run().add_picture(buf, width=Inches(usable_in))

    @staticmethod
    def _normalize_ocr_lines(result):
        """OCR 结果 -> [(x0, y0, x1, y1, text)], 按 y/x 排序, 过滤低置信度"""
        if not result:
            return []
        lines = []
        for item in result:
            try:
                box, text, score = item[0], str(item[1]), float(item[2])
            except (IndexError, TypeError, ValueError):
                continue
            text = text.strip()
            if not text or score < OCR_MIN_SCORE:
                continue
            xs = [float(p[0]) for p in box]
            ys = [float(p[1]) for p in box]
            lines.append((min(xs), min(ys), max(xs), max(ys), text))
        lines.sort(key=lambda t: (t[1], t[0]))
        return lines

    @staticmethod
    def _lines_to_paragraphs(lines):
        """将 OCR 文本框重建为段落列表(同行合并、按行距分段)"""
        if not lines:
            return []

        # 1) 聚类为视觉行: y 区间重叠视为同一行
        visual_rows = []  # 每项: [y0, y1, [(x0, x1, text), ...]]
        for x0, y0, x1, y1, text in lines:
            h = max(y1 - y0, 1.0)
            placed = False
            for row in visual_rows:
                ry0, ry1 = row[0], row[1]
                overlap = min(y1, ry1) - max(y0, ry0)
                if overlap > 0.4 * min(h, ry1 - ry0):
                    row[0] = min(row[0], y0)
                    row[1] = max(row[1], y1)
                    row[2].append((x0, x1, text))
                    placed = True
                    break
            if not placed:
                visual_rows.append([y0, y1, [(x0, x1, text)]])
        visual_rows.sort(key=lambda r: r[0])

        # 2) 行内按 x 排序, 间距大则空格分隔
        joined_rows = []  # [(y0, y1, text)]
        for ry0, ry1, parts in visual_rows:
            parts.sort(key=lambda p: p[0])
            heights = [max(ry1 - ry0, 1.0)]
            buf = []
            for idx, (x0, x1, text) in enumerate(parts):
                if idx == 0:
                    buf.append(text)
                    continue
                prev_x1 = parts[idx - 1][1]
                gap = x0 - prev_x1
                if gap > 0.35 * heights[0]:
                    buf.append(' ')
                # 直接拼接(过小间隙视为连字符/词内)
                # 中英文混排时统一加空格更易编辑
                if gap > 0.1 * heights[0] and not buf[-1].endswith(' ') \
                        and not text.startswith(' '):
                    # 已在上面处理大间隙; 中等间隙补一个空格
                    if gap <= 0.35 * heights[0]:
                        pass
                buf.append(text)
            joined_rows.append((ry0, ry1, ''.join(buf).strip()))

        # 3) 按行距分段: 行距 > 0.85 * 行高 视为段落间隔
        paragraphs = []
        cur = []
        prev_y1 = None
        for y0, y1, text in joined_rows:
            if not text:
                continue
            h = max(y1 - y0, 1.0)
            if prev_y1 is not None and (y0 - prev_y1) > 0.85 * h and cur:
                paragraphs.append(''.join(cur).strip())
                cur = []
            # 行间无大间隙时拼接(英文单词跨行需空格; 中文直接连)
            if cur:
                last = cur[-1]
                if last and text:
                    if last[-1].isascii() and last[-1].isalnum() \
                            and text[0].isascii() and text[0].isalnum():
                        cur.append(' ')
                    cur.append(text)
                else:
                    cur.append(text)
            else:
                cur.append(text)
            prev_y1 = y1
        if cur:
            paragraphs.append(''.join(cur).strip())

        return [p for p in paragraphs if p]

    @staticmethod
    def _cleanup(path):
        try:
            if os.path.isfile(path):
                os.remove(path)
        except OSError:
            pass

    # ------------------------------------------------------------------
    # 输出命名
    # ------------------------------------------------------------------

    def _generate_output_filename(self, input_filename):
        """生成输出文件名(.docx)"""
        base_name = os.path.splitext(input_filename)[0]
        return f"{base_name}.docx"

    @staticmethod
    def _result(filename, status, pages=None, output=None, warning=None,
                error=None, elapsed=None, method=None):
        return {
            'file': filename, 'status': status, 'pages': pages,
            'output': output, 'warning': warning, 'error': error,
            'elapsed': elapsed, 'method': method,
        }
