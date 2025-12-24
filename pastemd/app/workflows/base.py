"""Base workflow class."""

from abc import ABC, abstractmethod
from ...core.state import app_state
from ...service.notification.manager import NotificationManager
from ...service.document import DocumentGenerator
from ...service.spreadsheet import SpreadsheetGenerator
from ...service.preprocessor import HtmlPreprocessor, MarkdownPreprocessor
from ...utils.logging import log


class BaseWorkflow(ABC):
    """工作流基类"""
    
    def __init__(self):
        # 轻量级初始化
        self.notification_manager = NotificationManager()  # 复用单例
        
        # 延迟初始化(避免 Pandoc 开销)
        self._doc_generator = None
        self._sheet_generator = None
        
        # 无状态预处理器（可复用）
        self._markdown_preprocessor = MarkdownPreprocessor()
        self._html_preprocessor = HtmlPreprocessor()
    
    @property
    def config(self):
        """实时获取最新配置"""
        return app_state.config
    
    @property
    def doc_generator(self):
        """懒加载 DocumentGenerator"""
        if self._doc_generator is None:
            self._doc_generator = DocumentGenerator()
        return self._doc_generator
    
    @property
    def sheet_generator(self):
        """懒加载 SpreadsheetGenerator"""
        if self._sheet_generator is None:
            self._sheet_generator = SpreadsheetGenerator()
        return self._sheet_generator
    
    @property
    def markdown_preprocessor(self):
        """获取无状态的 MarkdownPreprocessor"""
        return self._markdown_preprocessor
    
    @property
    def html_preprocessor(self):
        """获取无状态的 HtmlPreprocessor"""
        return self._html_preprocessor
    
    @abstractmethod
    def execute(self) -> None:
        """执行工作流(子类实现)"""
        pass
    
    # 公共辅助方法
    def _notify_success(self, msg: str):
        """通知成功"""
        self.notification_manager.notify("PasteMD", msg, ok=True)
    
    def _notify_error(self, msg: str):
        """通知错误"""
        self.notification_manager.notify("PasteMD", msg, ok=False)
    
    def _log(self, msg: str):
        """记录日志"""
        log(msg)
