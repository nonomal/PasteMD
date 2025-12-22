"""Output executor - unified handler for document and spreadsheet output actions."""

import os
from typing import List, Tuple, Optional

from ...utils.clipboard import copy_files_to_clipboard
from ...utils.logging import log
from ...domains.awakener import AppLauncher
from ...domains.spreadsheet.generator import SpreadsheetGenerator
from ...utils.fs import generate_unique_path
from ...i18n import t


class OutputExecutor:
    """
    统一输出执行器
    
    负责处理 DOCX 和 XLSX 两类产物的 open/save/clipboard 三种输出动作。
    封装了写入文件、执行动作、显示通知的共同逻辑。
    """
    
    def __init__(self, notification_manager):
        """
        初始化执行器
        
        Args:
            notification_manager: 通知管理器实例
        """
        self.notification_manager = notification_manager
    
    def execute_docx(
        self,
        action: str,
        docx_bytes: bytes,
        output_path: str,
        *,
        from_md_file: bool = False,
        from_html: bool = False
    ) -> bool:
        """
        执行 DOCX 文档输出
        
        Args:
            action: 输出动作 ("open" | "save" | "clipboard")
            docx_bytes: DOCX 字节流
            output_path: 输出文件路径
            from_md_file: 是否来源于 MD 文件（影响通知文案）
            from_html: 是否来源于 HTML（影响通知文案）
            
        Returns:
            True 如果操作成功
        """
        try:
            # 1. 写入文件
            with open(output_path, "wb") as f:
                f.write(docx_bytes)
            log(f"Generated DOCX: {output_path}")
            
            # 2. 根据动作执行并通知
            if action == "open":
                return self._docx_open(output_path, from_md_file, from_html)
            elif action == "save":
                return self._docx_save(output_path, from_md_file)
            elif action == "clipboard":
                return self._docx_clipboard(output_path, from_md_file)
            else:
                log(f"Unknown DOCX action: {action}")
                return False
                
        except Exception as e:
            log(f"DOCX output failed: {e}")
            # 根据旧行为选择错误通知
            if action == "clipboard":
                # 复制到剪贴板失败（Markdown/HTML 统一）
                self.notification_manager.notify(
                    "PasteMD", t("workflow.action.clipboard_failed"), ok=False
                )
            elif action == "save":
                # 保存失败（Markdown/HTML 统一）
                self.notification_manager.notify(
                    "PasteMD", t("workflow.document.save_failed"), ok=False
                )
            elif from_html:
                # HTML open 失败时旧逻辑使用 html.generate_failed
                self.notification_manager.notify(
                    "PasteMD", t("workflow.html.generate_failed"), ok=False
                )
            else:
                # Markdown open 失败时旧逻辑使用 document.generate_failed
                self.notification_manager.notify(
                    "PasteMD", t("workflow.document.generate_failed"), ok=False
                )
            return False
    
    def execute_docx_batch(
        self,
        action: str,
        items: List[Tuple[bytes, str, str]],
        *,
        from_md_file: bool = False,
        from_html: bool = False,
        pre_failures: Optional[List[Tuple[str, str]]] = None,
    ) -> dict:
        """
        批量执行 DOCX 输出（仅用于无应用多文件分支）

        Args:
            action: 输出动作 ("open" | "save" | "clipboard")
            items: [(docx_bytes, output_path, source_filename), ...]
            from_md_file: 是否来源于 MD 文件（影响通知文案）
            from_html: 是否来源于 HTML（影响通知文案）
            pre_failures: 生成阶段已失败的 [(filename, error), ...]

        Returns:
            {"success_paths": [...], "failures": [(filename, error), ...]}
        """
        if not items:
            return {"success_paths": [], "failures": list(pre_failures or [])}

        # 确保批内输出路径唯一，避免同名覆盖
        seen_paths: set[str] = set()
        normalized_items: List[Tuple[bytes, str, str]] = []
        for docx_bytes, output_path, source_filename in items:
            unique_path = output_path

            if unique_path in seen_paths:
                base_dir = os.path.dirname(unique_path)
                stem, ext = os.path.splitext(os.path.basename(unique_path))
                idx = 1
                candidate = unique_path
                while candidate in seen_paths or os.path.exists(candidate):
                    candidate = os.path.join(base_dir, f"{stem}_batch{idx}{ext}")
                    idx += 1
                unique_path = candidate
            else:
                unique_path = generate_unique_path(unique_path)
                while unique_path in seen_paths or os.path.exists(unique_path):
                    unique_path = generate_unique_path(unique_path)

            seen_paths.add(unique_path)
            normalized_items.append((docx_bytes, unique_path, source_filename))

        success_paths: List[str] = []
        failures: List[Tuple[str, str]] = list(pre_failures or [])

        for docx_bytes, output_path, source_filename in normalized_items:
            try:
                with open(output_path, "wb") as f:
                    f.write(docx_bytes)
                log(f"Generated DOCX (batch): {output_path}")

                if action == "open":
                    ok = self._docx_open(
                        output_path, from_md_file, from_html, notify_success=False
                    )
                elif action == "save":
                    ok = self._docx_save(output_path, from_md_file, notify_success=False)
                elif action == "clipboard":
                    ok = True
                else:
                    log(f"Unknown DOCX action in batch: {action}")
                    ok = False

                if ok:
                    success_paths.append(output_path)
                else:
                    failures.append((source_filename, f"action_failed:{action}"))
            except Exception as e:
                log(f"DOCX batch item failed ({source_filename}): {e}")
                # 逐项失败通知保持旧语义
                if action == "clipboard":
                    self.notification_manager.notify(
                        "PasteMD", t("workflow.action.clipboard_failed"), ok=False
                    )
                elif action == "save":
                    self.notification_manager.notify(
                        "PasteMD", t("workflow.document.save_failed"), ok=False
                    )
                elif from_html:
                    self.notification_manager.notify(
                        "PasteMD", t("workflow.html.generate_failed"), ok=False
                    )
                else:
                    self.notification_manager.notify(
                        "PasteMD", t("workflow.document.generate_failed"), ok=False
                    )
                failures.append((source_filename, str(e)))

        # clipboard 动作：末尾一次性写入剪贴板（CF_HDROP 多路径）
        if action == "clipboard" and success_paths:
            try:
                copy_files_to_clipboard(success_paths)
            except Exception as e:
                log(f"DOCX batch clipboard failed: {e}")
                self.notification_manager.notify(
                    "PasteMD", t("workflow.action.clipboard_failed"), ok=False
                )
                failures.append(("_batch_clipboard", str(e)))
                return {"success_paths": success_paths, "failures": failures}

        # 多文件批量成功通知收敛为 1 条
        total_attempted = len(items) + len(pre_failures or [])
        if total_attempted > 1 and success_paths:
            action_name = (
                t(f"action.{action}")
                if action in ("open", "save", "clipboard", "none")
                else action
            )
            msg = t("workflow.md_file.batch_success", count=len(success_paths))
            msg += "\n" + t("workflow.md_file.batch_action_line", action=action_name)

            failed_items = [
                name
                for name, _ in failures
                if name and not name.startswith("_batch_")
            ]
            if failed_items:
                msg += "\n" + t(
                    "workflow.md_file.batch_failure_line",
                    failed_count=len(failed_items),
                    failed_files=", ".join(failed_items),
                )

            self.notification_manager.notify("PasteMD", msg, ok=True)

        return {"success_paths": success_paths, "failures": failures}

    def execute_xlsx(
        self,
        action: str,
        table_data: List[List[str]],
        output_path: str,
        keep_format: bool = True
    ) -> bool:
        """
        执行 XLSX 表格输出
        
        Args:
            action: 输出动作 ("open" | "save" | "clipboard")
            table_data: 表格数据（二维数组）
            output_path: 输出文件路径
            keep_format: 是否保留格式
            
        Returns:
            True 如果操作成功
        """
        try:
            if action == "open":
                return self._xlsx_open(table_data, output_path, keep_format)
            elif action == "save":
                return self._xlsx_save(table_data, output_path, keep_format)
            elif action == "clipboard":
                return self._xlsx_clipboard(table_data, output_path, keep_format)
            else:
                log(f"Unknown XLSX action: {action}")
                return False
                
        except Exception as e:
            log(f"XLSX output failed: {e}")
            if action == "clipboard":
                self.notification_manager.notify(
                    "PasteMD", t("workflow.action.clipboard_failed"), ok=False
                )
            else:
                self.notification_manager.notify(
                    "PasteMD", t("workflow.table.export_failed"), ok=False
                )
            return False
    
    # ==================== DOCX 私有方法 ====================
    
    def _docx_open(
        self,
        output_path: str,
        from_md_file: bool,
        from_html: bool,
        *,
        notify_success: bool = True,
    ) -> bool:
        """打开 DOCX 文件"""
        if AppLauncher.awaken_and_open_document(output_path):
            if notify_success:
                if from_html:
                    msg = t("workflow.html.generated_and_opened", path=output_path)
                elif from_md_file:
                    msg = t("workflow.md_file.generated_and_opened", path=output_path)
                else:
                    msg = t("workflow.document.generated_and_opened", path=output_path)
                self.notification_manager.notify("PasteMD", msg, ok=True)
            return True
        else:
            self.notification_manager.notify(
                "PasteMD",
                t("workflow.document.open_failed", path=output_path),
                ok=False
            )
            return False
    
    def _docx_save(
        self,
        output_path: str,
        from_md_file: bool,
        *,
        notify_success: bool = True,
    ) -> bool:
        """保存 DOCX 文件（文件已写入，仅通知）"""
        if notify_success:
            if from_md_file:
                msg = t("workflow.md_file.saved", path=output_path)
            else:
                msg = t("workflow.action.saved", path=output_path)
            self.notification_manager.notify("PasteMD", msg, ok=True)
        return True
    
    def _docx_clipboard(self, output_path: str, from_md_file: bool) -> bool:
        """复制 DOCX 文件到剪贴板"""
        copy_files_to_clipboard([output_path])
        if from_md_file:
            msg = t("workflow.md_file.clipboard_copied")
        else:
            msg = t("workflow.action.clipboard_copied")
        self.notification_manager.notify("PasteMD", msg, ok=True)
        return True
    
    # ==================== XLSX 私有方法 ====================
    
    def _xlsx_open(self, table_data: List[List[str]], output_path: str, keep_format: bool) -> bool:
        """生成并打开 XLSX 文件"""
        try:
            xlsx_bytes = SpreadsheetGenerator.generate_xlsx_bytes(table_data, keep_format)
            if not xlsx_bytes:
                raise Exception("Generated XLSX bytes are empty")
            with open(output_path, "wb") as f:
                f.write(xlsx_bytes)
            log(f"Successfully generated spreadsheet: {output_path}")
            if AppLauncher.awaken_and_open_spreadsheet(output_path):
                self.notification_manager.notify(
                    "PasteMD",
                    t("workflow.table.export_success", rows=len(table_data), path=output_path),
                    ok=True
                )
                return True
            else:
                self.notification_manager.notify(
                    "PasteMD",
                    t("workflow.table.export_open_failed", path=output_path),
                    ok=False
                )
                return False
        except Exception as e:
            log(f"Failed to generate spreadsheet: {e}")
            self.notification_manager.notify(
                "PasteMD",
                t("workflow.table.export_failed"),
                ok=False
            )
            return False
    
    def _xlsx_save(self, table_data: List[List[str]], output_path: str, keep_format: bool) -> bool:
        """生成 XLSX 文件（仅保存）"""
        try:
            xlsx_bytes = SpreadsheetGenerator.generate_xlsx_bytes(table_data, keep_format)
            if not xlsx_bytes:
                raise Exception("Generated XLSX bytes are empty")
            with open(output_path, "wb") as f:
                f.write(xlsx_bytes)
            log(f"Successfully generated spreadsheet: {output_path}")
            self.notification_manager.notify(
                "PasteMD",
                t("workflow.action.saved", path=output_path),
                ok=True
            )
            return True
        except Exception as e:
            log(f"Failed to generate spreadsheet: {e}")
            self.notification_manager.notify(
                "PasteMD",
                t("workflow.table.export_failed"),
                ok=False
            )
            return False
    
    def _xlsx_clipboard(self, table_data: List[List[str]], output_path: str, keep_format: bool) -> bool:
        """生成 XLSX 文件并复制到剪贴板"""
        try:
            xlsx_bytes = SpreadsheetGenerator.generate_xlsx_bytes(table_data, keep_format)
            if not xlsx_bytes:
                raise Exception("Generated XLSX bytes are empty")
            with open(output_path, "wb") as f:
                f.write(xlsx_bytes)
            log(f"Successfully generated spreadsheet: {output_path}")
            copy_files_to_clipboard([output_path])
            self.notification_manager.notify(
                "PasteMD",
                t("workflow.action.clipboard_copied"),
                ok=True
            )
            return True
        except Exception as e:
            log(f"Failed to generate spreadsheet: {e}")
            self.notification_manager.notify(
                "PasteMD",
                t("workflow.table.export_failed"),
                ok=False
            )
            return False