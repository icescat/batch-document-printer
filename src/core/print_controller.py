"""
批量打印控制器 v5.0
基于新的处理器架构的重构版本
"""
import time
from pathlib import Path
from typing import List, Optional, Callable
from concurrent.futures import ThreadPoolExecutor, Future

from .models import Document, PrintSettings, PrintStatus, FileType
from .printer_config_manager import PrinterConfigManager
from ..handlers import HandlerRegistry, PDFDocumentHandler, WordDocumentHandler, PowerPointDocumentHandler, ExcelDocumentHandler, ImageDocumentHandler, TextDocumentHandler


class PrintController:
    """批量打印控制器 - 基于处理器架构"""
    
    def __init__(self):
        """初始化打印控制器"""
        self._print_queue: List[Document] = []
        self._current_settings: Optional[PrintSettings] = None
        self._is_printing = False
        self._print_progress_callback: Optional[Callable] = None
        self._executor = ThreadPoolExecutor(max_workers=2)
        
        # 初始化打印机配置管理器
        self._printer_config_manager = PrinterConfigManager()
        
        # 初始化处理器注册中心
        self._setup_handlers()
    
    def _setup_handlers(self):
        """设置文档处理器"""
        self._handler_registry = HandlerRegistry()
        
        # 注册所有处理器
        handlers = [
            PDFDocumentHandler(),
            WordDocumentHandler(),
            PowerPointDocumentHandler(),
            ExcelDocumentHandler(),
            ImageDocumentHandler(),
            TextDocumentHandler()
        ]
        
        for handler in handlers:
            self._handler_registry.register_handler(handler)
        
        print("📋 打印控制器已初始化处理器架构")
        self._handler_registry.print_registry_info()
    
    def set_print_settings(self, settings: PrintSettings):
        """
        设置打印参数
        
        Args:
            settings: 打印设置对象
        """
        self._current_settings = settings
        print(f"打印设置已更新: 打印机={settings.printer_name}, 纸张={settings.paper_size}")
    
    def set_progress_callback(self, callback: Callable[[int, int, str], None]):
        """
        设置打印进度回调函数
        
        Args:
            callback: 回调函数，参数为(当前进度, 总数, 状态消息)
        """
        self._print_progress_callback = callback
    
    def add_documents_to_queue(self, documents: List[Document]):
        """
        添加文档到打印队列
        
        Args:
            documents: 文档列表
        """
        self._print_queue.extend(documents)
        print(f"已添加 {len(documents)} 个文档到打印队列，队列总数: {len(self._print_queue)}")
    
    def clear_queue(self):
        """清空打印队列"""
        self._print_queue.clear()
        print("打印队列已清空")
    
    @property
    def queue_size(self) -> int:
        """获取队列大小"""
        return len(self._print_queue)
    
    @property
    def is_printing(self) -> bool:
        """是否正在打印"""
        return self._is_printing
    
    def start_batch_print(self) -> Future:
        """
        开始批量打印
        
        Returns:
            Future对象，可用于跟踪打印任务状态
        """
        if self._is_printing:
            raise RuntimeError("打印任务已在进行中")
        
        if not self._print_queue:
            raise ValueError("打印队列为空")
        
        if not self._current_settings:
            raise ValueError("未设置打印参数")
        
        print(f"🖨️ 开始批量打印，共 {len(self._print_queue)} 个文档")
        
        # 提交打印任务到线程池
        future = self._executor.submit(self._execute_batch_print)
        return future
    
    def _execute_batch_print(self):
        """执行批量打印（在后台线程中运行）"""
        try:
            self._is_printing = True
            total_docs = len(self._print_queue)
            success_count = 0
            error_count = 0
            
            # 在批量打印开始前，应用系统级打印机配置
            printer_config_applied = False
            if self._current_settings and self._current_settings.printer_name:
                print(f"🔧 正在应用系统级打印机配置...")
                printer_config_applied = self._printer_config_manager.apply_batch_print_settings(
                    self._current_settings.printer_name, 
                    self._current_settings
                )
                if printer_config_applied:
                    print(f"✅ 系统级打印机配置已应用，所有文档将使用统一设置")
                else:
                    print(f"⚠️ 系统级配置应用失败，将使用各处理器的独立设置")
            
            for i, document in enumerate(self._print_queue):
                try:
                    # 更新进度
                    if self._print_progress_callback:
                        self._print_progress_callback(
                            i + 1, total_docs, 
                            f"正在打印: {document.file_name}"
                        )
                    
                    # 更新文档状态为打印中
                    document.print_status = PrintStatus.PRINTING
                    
                    # 执行打印
                    success = self._print_single_document(document)
                    
                    if success:
                        document.print_status = PrintStatus.COMPLETED
                        success_count += 1
                        print(f"✓ 打印成功: {document.file_name}")
                    else:
                        document.print_status = PrintStatus.ERROR
                        error_count += 1
                        print(f"✗ 打印失败: {document.file_name}")
                    
                    # 短暂延迟，避免打印队列拥堵
                    time.sleep(0.5)
                    
                except Exception as e:
                    document.print_status = PrintStatus.ERROR
                    error_count += 1
                    print(f"✗ 打印异常 {document.file_name}: {e}")
            
            # 打印完成
            if self._print_progress_callback:
                self._print_progress_callback(
                    total_docs, total_docs,
                    f"批量打印完成！成功: {success_count}, 失败: {error_count}"
                )
            
            print(f"🎉 批量打印任务完成: 成功 {success_count}/{total_docs} 个文档")
            
            # 恢复打印机原始配置
            if printer_config_applied:
                print(f"🔄 正在恢复打印机原始配置...")
                self._printer_config_manager.restore_printer_config()
                print(f"✅ 打印机配置已恢复")
            
        finally:
            self._is_printing = False
    
    def _print_single_document(self, document: Document) -> bool:
        """
        打印单个文档（使用新的处理器架构）
        
        Args:
            document: 要打印的文档
            
        Returns:
            是否打印成功
        """
        try:
            file_path = document.file_path
            
            if not file_path.exists():
                print(f"文件不存在: {file_path}")
                return False
            
            # 使用处理器注册中心获取相应的处理器
            handler = self._handler_registry.get_handler_by_file_type(document.file_type)
            
            if handler is None:
                print(f"❌ 不支持的文件类型: {document.file_type}")
                return False
            
            if not handler.can_handle_file(file_path):
                print(f"❌ 处理器无法处理文件: {file_path}")
                return False
            
            # 使用处理器打印文档
            print(f"🔧 使用 {handler.get_handler_name()} 打印文档")
            if self._current_settings is None:
                print("❌ 打印设置为空")
                return False
            success = handler.print_document(file_path, self._current_settings)
            
            return success
                
        except Exception as e:
            print(f"打印文档失败 {document.file_name}: {e}")
            return False
    
    def cancel_current_print(self):
        """取消当前打印任务"""
        if self._is_printing:
            print("请求取消当前打印任务...")
    
    def get_print_queue_status(self) -> dict:
        """
        获取打印队列状态统计
        
        Returns:
            状态统计字典
        """
        status_count = {
            'pending': 0,
            'printing': 0,
            'completed': 0,
            'error': 0
        }
        
        for doc in self._print_queue:
            status_count[doc.print_status.value] += 1
        
        return {
            'total': len(self._print_queue),
            'status_count': status_count,
            'is_printing': self._is_printing,
            'supported_types': [ft.value for ft in self._handler_registry.get_all_supported_file_types()],
            'supported_extensions': sorted(self._handler_registry.get_all_supported_extensions())
        }
    
    def get_supported_file_types(self) -> List[FileType]:
        """获取支持的文件类型列表"""
        return list(self._handler_registry.get_all_supported_file_types())
    
    def get_supported_extensions(self) -> List[str]:
        """获取支持的文件扩展名列表"""
        return sorted(self._handler_registry.get_all_supported_extensions())
    
    def __del__(self):
        """析构函数，清理资源"""
        if hasattr(self, '_executor'):
            self._executor.shutdown(wait=True) 