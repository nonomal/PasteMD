"""WPS document workflow."""
import sys

from pastemd.app.workflows.word.word_base import WordBaseWorkflow
from pastemd.service.document import WPSPlacer
from pastemd.i18n import t

from pastemd.core.errors import ClipboardError, PandocError
from pastemd.utils.system_detect import is_windows
from pastemd.utils.html_formatter import postprocess_pandoc_html


class WPSWorkflow(WordBaseWorkflow):
    """WPS 文档工作流"""

    def __init__(self):
        super().__init__()
        self._placer = WPSPlacer()

    @property
    def app_name(self) -> str:
        return "WPS 文字"

    @property
    def placer(self):
        return self._placer

    def execute(self) -> None:
        """
        macOS WPS：
        - 使用 HTML+RTF 富文本粘贴（WPS 偏好 HTML，RTF 作为兜底）
        - 公式统一保留为 `$...$` / `$$...$$` 文本（keep-latex-math.lua）
        """
        if is_windows():
            return super().execute()

        content_type: str | None = None
        from_md_file = False
        md_file_count = 0

        try:
            content_type, content, from_md_file, md_file_count = self._read_clipboard()
            self._log(f"Clipboard content type: {content_type}")
            config = self.config.copy()
            config["Keep_original_formula"] = True  # 保留公式为 LaTeX 文本
            if content_type == "html":
                content = self.html_preprocessor.process(content, config)
                md_text = self.doc_generator.convert_html_to_markdown_text(
                    content, config
                )
            else:
                # markdown
                md_text = self.markdown_preprocessor.process(content, config)
            
            html_text = self.doc_generator.convert_markdown_to_html_text(
                md_text, config
            )
            
            # 后处理 Pandoc 输出的 HTML，修复代码块格式等问题
            html_text = postprocess_pandoc_html(html_text)

            # 内容落地由 placer 负责（写剪贴板 + Cmd+V）
            result = self.placer.place(
                None,
                self.config, _plain_text=md_text, _rtf_bytes=None, _html_text=html_text
            )

            if not result.success:
                self._notify_error(result.error or t("workflow.generic.failure"))
                return

            if from_md_file:
                if md_file_count > 1:
                    msg = t(
                        "workflow.md_file.insert_success_multi",
                        count=md_file_count,
                        app=self.app_name,
                    )
                else:
                    msg = t("workflow.md_file.insert_success", app=self.app_name)
            elif content_type == "html":
                msg = t("workflow.html.insert_success", app=self.app_name)
            else:
                msg = t("workflow.word.insert_success", app=self.app_name)

            self._notify_success(msg)

        except ClipboardError as e:
            self._log(f"Clipboard error: {e}")
            self._notify_error(t("workflow.clipboard.read_failed"))
        except PandocError as e:
            self._log(f"Pandoc error: {e}")
            self._notify_error(t("workflow.markdown.convert_failed"))
        except Exception as e:
            self._log(f"{self.app_name} workflow failed: {e}")
            import traceback

            traceback.print_exc()
            self._notify_error(t("workflow.generic.failure"))
