"""Base preprocessor class."""

from abc import ABC, abstractmethod


class BasePreprocessor(ABC):
    """内容预处理器基类（无状态）"""

    @abstractmethod
    def process(self, content: any, config: dict) -> any:
        """
        预处理内容

        Args:
            content: 原始内容
            config: 配置字典

        Returns:
            处理后的内容
        """
        pass
