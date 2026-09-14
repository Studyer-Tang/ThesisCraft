"""Recognition operations for the shared document engine."""

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph
from .constants import RE_TITLE_H1, RE_TITLE_H2


class RecognitionMixin:
    def _find_title_and_subtitle_paragraphs(self, doc, is_from_txt, start_index=0):
        """
        查找题目和副标题段落的索引范围
        返回: (title_indices, subtitle_indices)
        title_indices: 题目行的索引列表
        subtitle_indices: 副标题行的索引列表
        """
        all_blocks = list(self._iter_block_items(doc))

        # 查找首个标题行
        first_title_idx = -1

        if is_from_txt:
            self._log("文档源自 TXT，采用智能规则查找题目...")
            for idx in range(start_index, len(all_blocks)):
                block = all_blocks[idx]
                if isinstance(block, Paragraph) and block.text.strip():
                    text_to_check = block.text.strip()
                    if RE_TITLE_H1.match(text_to_check) or RE_TITLE_H2.match(
                        text_to_check
                    ):
                        self._log(
                            f"  > 首个非空行 (块 {idx + 1}) 符合标题格式，认定本文档无独立题目。"
                        )
                        return [], []
                    else:
                        self._log(
                            f"  > 在块 {idx + 1} 发现首个非空段落，认定为题目首行。"
                        )
                        first_title_idx = idx
                        break
        else:
            self._log("正在预扫描以确定居中题目位置...")
            for idx in range(start_index, len(all_blocks)):
                block = all_blocks[idx]
                if not isinstance(block, Paragraph) or not block.text.strip():
                    continue
                para = block
                text_to_check = para.text.lstrip()
                if RE_TITLE_H1.match(text_to_check) or RE_TITLE_H2.match(text_to_check):
                    self._log("  > 发现一级/二级标题，在此之前未找到居中题目。")
                    return [], []
                if self._get_paragraph_alignment(para) == WD_ALIGN_PARAGRAPH.CENTER:
                    self._log(f"  > 在块 {idx + 1} 发现潜在题目首行。")
                    first_title_idx = idx
                    break

        if first_title_idx == -1:
            self._log("  > 扫描结束，未能找到题目。")
            return [], []

        # 获取首个标题行的字体字号信息
        first_title_para = all_blocks[first_title_idx]
        title_font, title_size = self._get_paragraph_font_info(first_title_para)

        # 向下查找连续的标题行
        title_indices = [first_title_idx]
        idx = first_title_idx + 1

        while idx < len(all_blocks):
            block = all_blocks[idx]
            if not isinstance(block, Paragraph):
                break

            para = block
            text = para.text.strip()

            # 遇到空行，停止标题识别
            if not text:
                self._log(f"  > 在块 {idx + 1} 遇到空行，标题识别结束。")
                break

            # 检查是否居中
            if self._get_paragraph_alignment(para) != WD_ALIGN_PARAGRAPH.CENTER:
                break

            # 检查字体字号是否与首行相同
            para_font, para_size = self._get_paragraph_font_info(para)
            if para_font == title_font and para_size == title_size:
                self._log(f"  > 块 {idx + 1} 也是标题行（居中且字体字号相同）。")
                title_indices.append(idx)
                idx += 1
            else:
                # 字体字号不同，可能是副标题的开始
                break

        self._log(f"  > 共识别到 {len(title_indices)} 行标题。")

        # 查找副标题
        subtitle_indices = []
        subtitle_start_idx = idx

        # 跳过空行
        while subtitle_start_idx < len(all_blocks):
            block = all_blocks[subtitle_start_idx]
            if isinstance(block, Paragraph) and block.text.strip():
                break
            if isinstance(block, Paragraph):
                subtitle_start_idx += 1
            else:
                # 遇到非段落（如表格），停止
                break

        # 检查是否有副标题
        if subtitle_start_idx < len(all_blocks):
            block = all_blocks[subtitle_start_idx]
            if isinstance(block, Paragraph):
                para = block
                text = para.text.strip()

                # 副标题必须居中
                if (
                    text
                    and self._get_paragraph_alignment(para) == WD_ALIGN_PARAGRAPH.CENTER
                ):
                    # 检查字体字号是否与标题不同
                    para_font, para_size = self._get_paragraph_font_info(para)
                    if para_font != title_font or para_size != title_size:
                        self._log(
                            f"  > 在块 {subtitle_start_idx + 1} 发现副标题首行（居中且字体字号与标题不同）。"
                        )
                        subtitle_indices.append(subtitle_start_idx)

                        # 查找连续的副标题行
                        subtitle_font, subtitle_size = para_font, para_size
                        idx = subtitle_start_idx + 1

                        while idx < len(all_blocks):
                            block = all_blocks[idx]
                            if not isinstance(block, Paragraph):
                                break

                            para = block
                            text = para.text.strip()

                            # 遇到空行，停止副标题识别
                            if not text:
                                self._log(
                                    f"  > 在块 {idx + 1} 遇到空行，副标题识别结束。"
                                )
                                break

                            # 检查是否居中
                            if (
                                self._get_paragraph_alignment(para)
                                != WD_ALIGN_PARAGRAPH.CENTER
                            ):
                                break

                            # 检查字体字号是否与副标题首行相同
                            para_font, para_size = self._get_paragraph_font_info(para)
                            if (
                                para_font == subtitle_font
                                and para_size == subtitle_size
                            ):
                                self._log(
                                    f"  > 块 {idx + 1} 也是副标题行（居中且字体字号相同）。"
                                )
                                subtitle_indices.append(idx)
                                idx += 1
                            else:
                                break

                        self._log(f"  > 共识别到 {len(subtitle_indices)} 行副标题。")

        return title_indices, subtitle_indices
