"""
文档处理器基础类
定义所有文档处理器必须实现的接口
"""
from abc import ABC, abstractmethod
from pathlib import Path
from typing import List, Optional, Set
from ..core.models import FileType, PrintSettings


class BaseDocumentHandler(ABC):
    """文档处理器基础抽象类"""
    
    @abstractmethod
    def get_supported_file_types(self) -> Set[FileType]:
        """
        获取支持的文件类型
        
        Returns:
            支持的FileType集合
        """
        pass
    
    @abstractmethod
    def get_supported_extensions(self) -> Set[str]:
        """
        获取支持的文件扩展名
        
        Returns:
            支持的文件扩展名集合（小写，包含点号，如：{'.pdf', '.docx'}）
        """
        pass
    
    @abstractmethod
    def can_handle_file(self, file_path: Path) -> bool:
        """
        检查是否能处理指定文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否能处理该文件
        """
        pass
    
    @abstractmethod
    def print_document(self, file_path: Path, settings: PrintSettings) -> bool:
        """
        打印文档
        
        Args:
            file_path: 文件路径
            settings: 打印设置
            
        Returns:
            是否打印成功
        """
        pass
    
    @abstractmethod
    def count_pages(self, file_path: Path) -> int:
        """
        统计文档页数
        
        Args:
            file_path: 文件路径
            
        Returns:
            页数（如果计算失败抛出异常）
        """
        pass
    
    def get_handler_name(self) -> str:
        """
        获取处理器名称
        
        Returns:
            处理器名称
        """
        return self.__class__.__name__
    
    def validate_file_exists(self, file_path: Path) -> bool:
        """
        验证文件是否存在
        
        Args:
            file_path: 文件路径
            
        Returns:
            文件是否存在
        """
        return file_path.exists() and file_path.is_file()
    
    def get_file_size_mb(self, file_path: Path) -> float:
        """
        获取文件大小（MB）
        
        Args:
            file_path: 文件路径
            
        Returns:
            文件大小（MB）
        """
        try:
            return file_path.stat().st_size / (1024 * 1024)
        except Exception:
            return 0.0 