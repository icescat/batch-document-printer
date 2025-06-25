"""
文档处理器注册中心
负责管理和分发各种文档格式的处理器
"""
from pathlib import Path
from typing import Dict, List, Optional, Set
from ..core.models import FileType
from .base_handler import BaseDocumentHandler


class HandlerRegistry:
    """处理器注册中心"""
    
    def __init__(self):
        """初始化注册中心"""
        self._handlers: List[BaseDocumentHandler] = []
        self._file_type_handlers: Dict[FileType, BaseDocumentHandler] = {}
        self._extension_handlers: Dict[str, BaseDocumentHandler] = {}
    
    def register_handler(self, handler: BaseDocumentHandler):
        """
        注册处理器
        
        Args:
            handler: 要注册的处理器
        """
        if handler in self._handlers:
            print(f"处理器 {handler.get_handler_name()} 已经注册")
            return
        
        # 添加到处理器列表
        self._handlers.append(handler)
        
        # 按文件类型注册
        for file_type in handler.get_supported_file_types():
            if file_type in self._file_type_handlers:
                print(f"警告: 文件类型 {file_type} 已有处理器 {self._file_type_handlers[file_type].get_handler_name()}, "
                      f"将被 {handler.get_handler_name()} 替换")
            self._file_type_handlers[file_type] = handler
        
        # 按文件扩展名注册
        for extension in handler.get_supported_extensions():
            ext_lower = extension.lower()
            if ext_lower in self._extension_handlers:
                print(f"警告: 文件扩展名 {ext_lower} 已有处理器 {self._extension_handlers[ext_lower].get_handler_name()}, "
                      f"将被 {handler.get_handler_name()} 替换")
            self._extension_handlers[ext_lower] = handler
        
        print(f"已注册处理器: {handler.get_handler_name()}")
        print(f"  - 支持文件类型: {[ft.value for ft in handler.get_supported_file_types()]}")
        print(f"  - 支持扩展名: {sorted(handler.get_supported_extensions())}")
    
    def unregister_handler(self, handler: BaseDocumentHandler):
        """
        注销处理器
        
        Args:
            handler: 要注销的处理器
        """
        if handler not in self._handlers:
            print(f"处理器 {handler.get_handler_name()} 未注册")
            return
        
        # 从处理器列表移除
        self._handlers.remove(handler)
        
        # 从文件类型映射移除
        for file_type in handler.get_supported_file_types():
            if file_type in self._file_type_handlers and self._file_type_handlers[file_type] == handler:
                del self._file_type_handlers[file_type]
        
        # 从扩展名映射移除
        for extension in handler.get_supported_extensions():
            ext_lower = extension.lower()
            if ext_lower in self._extension_handlers and self._extension_handlers[ext_lower] == handler:
                del self._extension_handlers[ext_lower]
        
        print(f"已注销处理器: {handler.get_handler_name()}")
    
    def get_handler_by_file_type(self, file_type: FileType) -> Optional[BaseDocumentHandler]:
        """
        根据文件类型获取处理器
        
        Args:
            file_type: 文件类型
            
        Returns:
            处理器实例或None
        """
        return self._file_type_handlers.get(file_type)
    
    def get_handler_by_extension(self, extension: str) -> Optional[BaseDocumentHandler]:
        """
        根据文件扩展名获取处理器
        
        Args:
            extension: 文件扩展名（可以包含或不包含点号）
            
        Returns:
            处理器实例或None
        """
        # 标准化扩展名
        if not extension.startswith('.'):
            extension = '.' + extension
        return self._extension_handlers.get(extension.lower())
    
    def get_handler_by_file_path(self, file_path: Path) -> Optional[BaseDocumentHandler]:
        """
        根据文件路径获取处理器
        
        Args:
            file_path: 文件路径
            
        Returns:
            处理器实例或None
        """
        extension = file_path.suffix.lower()
        return self.get_handler_by_extension(extension)
    
    def can_handle_file(self, file_path: Path) -> bool:
        """
        检查是否有处理器可以处理指定文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否可以处理
        """
        handler = self.get_handler_by_file_path(file_path)
        if handler:
            return handler.can_handle_file(file_path)
        return False
    
    def get_all_supported_extensions(self) -> Set[str]:
        """
        获取所有支持的文件扩展名
        
        Returns:
            所有支持的扩展名集合
        """
        return set(self._extension_handlers.keys())
    
    def get_all_supported_file_types(self) -> Set[FileType]:
        """
        获取所有支持的文件类型
        
        Returns:
            所有支持的文件类型集合
        """
        return set(self._file_type_handlers.keys())
    
    def get_registered_handlers(self) -> List[BaseDocumentHandler]:
        """
        获取所有已注册的处理器
        
        Returns:
            处理器列表
        """
        return self._handlers.copy()
    
    def print_registry_info(self):
        """打印注册中心信息"""
        print("\n=== 处理器注册中心信息 ===")
        print(f"已注册处理器数量: {len(self._handlers)}")
        
        for i, handler in enumerate(self._handlers, 1):
            print(f"\n{i}. {handler.get_handler_name()}")
            print(f"   文件类型: {[ft.value for ft in handler.get_supported_file_types()]}")
            print(f"   扩展名: {sorted(handler.get_supported_extensions())}")
        
        print(f"\n支持的所有扩展名: {sorted(self.get_all_supported_extensions())}")
        print("========================\n") 