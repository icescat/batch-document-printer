"""
批量打印控制器
负责调用Windows打印API执行批量打印任务
"""
import os
import subprocess
import time
from pathlib import Path
from typing import List, Optional, Callable
import win32api
import win32print
from concurrent.futures import ThreadPoolExecutor, Future

from .models import Document, PrintSettings, PrintStatus, FileType


class PrintController:
    """批量打印控制器类"""
    
    def __init__(self):
        """初始化打印控制器"""
        self._print_queue: List[Document] = []
        self._current_settings: Optional[PrintSettings] = None
        self._is_printing = False
        self._print_progress_callback: Optional[Callable] = None
        self._executor = ThreadPoolExecutor(max_workers=2)
    
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
        
        print(f"开始批量打印，共 {len(self._print_queue)} 个文档")
        
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
            
            print(f"批量打印任务完成: 成功 {success_count}/{total_docs} 个文档")
            
        finally:
            self._is_printing = False
    
    def _print_single_document(self, document: Document) -> bool:
        """
        打印单个文档
        
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
            
            # 根据文件类型选择打印方法
            if document.file_type == FileType.PDF:
                return self._print_pdf(file_path)
            elif document.file_type == FileType.WORD:
                return self._print_word_document(file_path)
            elif document.file_type == FileType.PPT:
                return self._print_powerpoint(file_path)
            else:
                print(f"不支持的文件类型: {document.file_type}")
                return False
                
        except Exception as e:
            print(f"打印文档失败 {document.file_name}: {e}")
            return False
    
    def _print_pdf(self, file_path: Path) -> bool:
        """
        打印PDF文件
        
        Args:
            file_path: PDF文件路径
            
        Returns:
            是否打印成功
        """
        try:
            # 使用Adobe Reader或默认PDF查看器进行打印
            # 构建打印命令
            cmd = [
                "powershell", "-Command",
                f"Start-Process -FilePath '{file_path}' -Verb Print -WindowStyle Hidden"
            ]
            
            # 执行打印命令
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0:
                time.sleep(2)  # 等待打印开始
                return True
            else:
                print(f"PDF打印命令失败: {result.stderr}")
                return False
                
        except subprocess.TimeoutExpired:
            print(f"PDF打印超时: {file_path.name}")
            return False
        except Exception as e:
            print(f"PDF打印异常: {e}")
            return False
    
    def _print_word_document(self, file_path: Path) -> bool:
        """
        打印Word文档
        
        Args:
            file_path: Word文件路径
            
        Returns:
            是否打印成功
        """
        try:
            import comtypes.client
            
            # 创建Word应用程序对象
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False
            
            try:
                # 打开文档
                doc = word.Documents.Open(str(file_path))
                
                # 设置打印参数
                if self._current_settings:
                    # 设置打印机
                    word.ActivePrinter = self._current_settings.printer_name
                    
                    # 打印文档
                    doc.PrintOut(
                        Copies=self._current_settings.copies,
                        Collate=True,
                        PrintToFile=False
                    )
                else:
                    doc.PrintOut()
                
                # 关闭文档
                doc.Close(SaveChanges=False)
                
                return True
                
            finally:
                # 退出Word应用程序
                word.Quit()
                
        except Exception as e:
            print(f"Word文档打印失败: {e}")
            return False
    
    def _print_powerpoint(self, file_path: Path) -> bool:
        """
        打印PowerPoint文档
        
        Args:
            file_path: PPT文件路径
            
        Returns:
            是否打印成功
        """
        try:
            import comtypes.client
            
            # 创建PowerPoint应用程序对象
            ppt = comtypes.client.CreateObject("PowerPoint.Application")
            ppt.Visible = False
            
            try:
                # 打开演示文稿
                presentation = ppt.Presentations.Open(str(file_path))
                
                # 打印演示文稿
                if self._current_settings:
                    presentation.PrintOptions.ActivePrinter = self._current_settings.printer_name
                    presentation.PrintOptions.NumberOfCopies = self._current_settings.copies
                
                presentation.PrintOut()
                
                # 关闭演示文稿
                presentation.Close()
                
                return True
                
            finally:
                # 退出PowerPoint应用程序
                ppt.Quit()
                
        except Exception as e:
            print(f"PowerPoint文档打印失败: {e}")
            return False
    
    def cancel_current_print(self):
        """取消当前打印任务（注意：可能无法立即停止正在打印的文档）"""
        if self._is_printing:
            print("请求取消当前打印任务...")
            # 在实际应用中，这里可以设置一个标志来停止打印循环
            # 但已经提交到打印队列的文档可能仍会继续打印
    
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
            'is_printing': self._is_printing
        }
    
    def __del__(self):
        """析构函数，清理资源"""
        if hasattr(self, '_executor'):
            self._executor.shutdown(wait=True) 