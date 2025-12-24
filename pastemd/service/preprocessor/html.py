"""HTML content preprocessor."""

from bs4 import BeautifulSoup
from .base import BasePreprocessor
from ...utils.html_formatter import clean_html_content, convert_strikethrough_to_del, remove_empty_paragraphs, unwrap_li_paragraphs
from ...utils.logging import log


class HtmlPreprocessor(BasePreprocessor):
    """HTML 内容预处理器（无状态）"""

    def process(self, html: str, config: dict) -> str:
        """
        预处理 HTML 内容

        处理步骤:
        1. 清理无效元素（SVG等）
        2. 转换删除线标记
        3. 清理 LaTeX 公式块中的 br 标签
        4. 其他自定义处理...

        Args:
            html: 原始 HTML 内容
            config: 配置字典

        Returns:
            预处理后的 HTML 内容
        """
        log("Preprocessing HTML content")

        # 使用 html_formatter 进行清理
        soup = BeautifulSoup(html, "html.parser")
        clean_html_content(soup, config)

        if config.get("convert_strikethrough", True):
            convert_strikethrough_to_del(soup)

        unwrap_li_paragraphs(soup)
        remove_empty_paragraphs(soup)

        html_output = str(soup)
        
        # 仅在 HTML 不包含 DOCTYPE 时才添加
        if "<!DOCTYPE" not in html_output.upper():
            html_output = f"<!DOCTYPE html>\n<meta charset='utf-8'>\n{html_output}"
        
        return html_output
