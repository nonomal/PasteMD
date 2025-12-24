"""Shared workflow logic for Word-family document apps."""

from __future__ import annotations

from abc import ABC, abstractmethod

from pastemd.app.workflows.base import BaseWorkflow
from pastemd.core.errors import ClipboardError, PandocError
from pastemd.i18n import t
from pastemd.utils.clipboard import (
    get_clipboard_html,
    get_clipboard_text,
    is_clipboard_empty,
    read_markdown_files_from_clipboard,
)
from pastemd.utils.fs import generate_output_path
from pastemd.utils.html_analyzer import is_plain_html_fragment
from pastemd.utils.markdown_utils import merge_markdown_contents


class WordBaseWorkflow(BaseWorkflow, ABC):
    """Word/WPS 文字共用工作流逻辑。"""

    @property
    @abstractmethod
    def app_name(self) -> str: ...

    @property
    @abstractmethod
    def placer(self): ...

    def execute(self) -> None:
        content_type: str | None = None
        from_md_file = False
        md_file_count = 0

        try:
            content_type, content, from_md_file, md_file_count = self._read_clipboard()
            self._log(f"Clipboard content type: {content_type}")

            if content_type == "markdown":
                content = self.markdown_preprocessor.process(content, self.config)
            elif content_type == "html":
                # 预处理 HTML，清理 LaTeX 公式块中的 br 标签等
                content = self.html_preprocessor.process(content, self.config)

            if content_type == "html":
                docx_bytes = self.doc_generator.convert_html_to_docx_bytes(
                    content, self.config
                )
            else:
                docx_bytes = self.doc_generator.convert_markdown_to_docx_bytes(
                    content, self.config
                )

            result = self.placer.place(docx_bytes, self.config)

            if result.success:
                if result.method:
                    self._log(f"Insert method: {result.method}")

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
            else:
                self._notify_error(result.error or t("workflow.generic.failure"))

            if result.success and self.config.get("keep_file", False):
                self._save_docx(docx_bytes)

        except ClipboardError as e:
            self._log(f"Clipboard error: {e}")
            self._notify_error(t("workflow.clipboard.read_failed"))
        except PandocError as e:
            self._log(f"Pandoc error: {e}")
            if content_type == "html":
                self._notify_error(t("workflow.html.convert_failed_generic"))
            else:
                self._notify_error(t("workflow.markdown.convert_failed"))
        except Exception as e:
            self._log(f"{self.app_name} workflow failed: {e}")
            import traceback

            traceback.print_exc()
            self._notify_error(t("workflow.generic.failure"))

    def _read_clipboard(self) -> tuple[str, str, bool, int]:
        """
        读取剪贴板,返回 (类型, 内容, 是否来自 MD 文件, MD 文件数量)
        """
        try:
            html = get_clipboard_html(self.config)
            if not is_plain_html_fragment(html):
                return ("html", html, False, 0)
        except ClipboardError:
            pass

        if not is_clipboard_empty():
            return ("markdown", get_clipboard_text(), False, 0)

        found, files_data, _ = read_markdown_files_from_clipboard()
        if found:
            merged = merge_markdown_contents(files_data)
            return ("markdown", merged, True, len(files_data))

        raise ClipboardError("剪贴板为空或无有效内容")

    def _save_docx(self, docx_bytes: bytes) -> None:
        try:
            output_path = generate_output_path(
                keep_file=True,
                save_dir=self.config.get("save_dir", ""),
                md_text="",
            )
            with open(output_path, "wb") as f:
                f.write(docx_bytes)
            self._log(f"Saved DOCX to: {output_path}")
        except Exception as e:
            self._log(f"Failed to save DOCX: {e}")

