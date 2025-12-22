"""Excel spreadsheet workflow."""

from .base import BaseWorkflow
from ...domains.spreadsheet import ExcelPlacer
from ...utils.clipboard import get_clipboard_text, is_clipboard_empty
from ...domains.spreadsheet.parser import parse_markdown_table
from ...core.errors import ClipboardError


class ExcelWorkflow(BaseWorkflow):
    """Excel 表格工作流"""
    
    def __init__(self):
        super().__init__()
        self.placer = ExcelPlacer()  # 无需工厂,直接实例化
    
    def execute(self) -> None:
        """执行 Excel 工作流"""
        if not self.config.get("enable_excel", True):
            self. _log("Excel workflow is disabled in config.")
            self._notify_error("Excel 表格功能未启用")
            return  # 未启用则跳过
        try:
            # 1. 读取剪贴板
            table_data = self._read_clipboard_table()
            self._log(f"Parsed table with {len(table_data)} rows")
            
            # 2. 落地内容(不做降级,失败即报错)
            result = self.placer.place(table_data, self.config)
            
            # 3. 通知结果
            if result.success:
                method_str = result.method or "unknown"
                self._notify_success(f"成功插入到 Excel (方式: {method_str})")
            else:
                self._notify_error(result.error or "Excel 插入失败")
            
            # 4. 可选保存
            if result.success and self.config.get("keep_file", False):
                self._save_xlsx(table_data)
        
        except ClipboardError as e:
            self._log(f"Clipboard error: {e}")
            self._notify_error("剪贴板读取失败或无有效表格")
        except Exception as e:
            self._log(f"Excel workflow failed: {e}")
            import traceback
            traceback.print_exc()
            self._notify_error("操作失败")
    
    def _read_clipboard_table(self) -> list:
        """
        读取剪贴板中的 Markdown 表格
        
        Returns:
            二维数组表格数据
            
        Raises:
            ClipboardError: 剪贴板为空或无表格
        """
        if is_clipboard_empty():
            raise ClipboardError("剪贴板为空")
        
        markdown_text = get_clipboard_text()
        table_data = parse_markdown_table(markdown_text)
        
        if not table_data:
            raise ClipboardError("剪贴板中无有效 Markdown 表格")
        
        return table_data
    
    def _save_xlsx(self, table_data: list):
        """保存 XLSX 到磁盘"""
        try:
            from ...utils.fs import generate_output_path
            from ...domains.spreadsheet import SpreadsheetGenerator
            
            # 生成 XLSX 字节流
            xlsx_bytes = SpreadsheetGenerator.generate_xlsx_bytes(
                table_data,
                keep_format=self.config.get("excel_keep_format", self.config.get("keep_format", True))
            )
            
            # 保存到文件
            output_path = generate_output_path(
                keep_file=True,
                save_dir=self.config.get("save_dir", ""),
                md_text="",
                extension=".xlsx"
            )
            with open(output_path, "wb") as f:
                f.write(xlsx_bytes)
            self._log(f"Saved XLSX to: {output_path}")
        except Exception as e:
            self._log(f"Failed to save XLSX: {e}")
