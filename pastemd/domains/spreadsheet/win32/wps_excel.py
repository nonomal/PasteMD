"""Windows WPS Excel spreadsheet placer."""

from typing import List
from ..base import BaseSpreadsheetPlacer
from ....core.types import PlacementResult
from ....utils.logging import log
from ....i18n import t

# 复用现有 COM 插入器
from .wps_excel_inserter import WPSExcelInserter


class WPSExcelPlacer(BaseSpreadsheetPlacer):
    """Windows WPS Excel 内容落地器"""
    
    def __init__(self):
        self.com_inserter = WPSExcelInserter()
    
    def place(self, table_data: List[List[str]], config: dict) -> PlacementResult:
        """通过 COM 插入表格数据,失败不降级"""
        try:
            keep_format = config.get("excel_keep_format", config.get("keep_format", True))
            success = self.com_inserter.insert(table_data, keep_format=keep_format)
            
            if success:
                return PlacementResult(success=True, method="com")
            else:
                raise Exception(t("placer.win32_wps_excel.insert_failed_unknown"))
        
        except Exception as e:
            log(f"WPS Excel COM 插入失败: {e}")
            return PlacementResult(
                success=False,
                method="com",
                error=t("placer.win32_wps_excel.insert_failed", error=str(e))
            )
