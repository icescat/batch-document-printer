"""
页数统计管理器 v5.0
基于新的处理器架构的重构版本
"""
import time
from pathlib import Path
from typing import List, Optional, Callable, Tuple
from dataclasses import dataclass, field
from enum import Enum
from concurrent.futures import ThreadPoolExecutor, Future, as_completed

from .models import Document, FileType
from ..handlers import HandlerRegistry, PDFDocumentHandler, WordDocumentHandler, PowerPointDocumentHandler, ExcelDocumentHandler, ImageDocumentHandler, TextDocumentHandler


class PageCountStatus(Enum):
    """页数统计状态"""
    UNKNOWN = "unknown"
    CALCULATING = "calculating"
    SUCCESS = "success"
    SKIPPED_LARGE = "skipped_large"
    SKIPPED_ENCRYPTED = "skipped_encrypted"
    SKIPPED_DAMAGED = "skipped_damaged"
    SKIPPED_NO_ACCESS = "skipped_no_access"
    SKIPPED_NO_OFFICE = "skipped_no_office"
    ERROR = "error"


@dataclass
class PageCountResult:
    """页数统计结果"""
    document: Document
    page_count: Optional[int] = None
    status: PageCountStatus = PageCountStatus.UNKNOWN
    error_message: str = ""
    calculation_time: float = 0.0


@dataclass
class PageCountSummary:
    """页数统计汇总"""
    total_files: int = 0
    total_pages: int = 0
    success_count: int = 0
    skipped_count: int = 0
    error_count: int = 0
    
    # 按文件类型统计
    word_files: int = 0
    word_pages: int = 0
    ppt_files: int = 0
    ppt_pages: int = 0
    excel_files: int = 0
    excel_pages: int = 0
    pdf_files: int = 0
    pdf_pages: int = 0
    image_files: int = 0
    image_pages: int = 0
    
    # 问题文件列表
    skipped_files: List[PageCountResult] = field(default_factory=list)
    error_files: List[PageCountResult] = field(default_factory=list)


class PageCountManager:
    """页数统计管理器 - 基于处理器架构"""
    
    def __init__(self):
        """初始化页数统计管理器"""
        self._cancel_flag = False
        self._progress_callback: Optional[Callable] = None
        self._executor = ThreadPoolExecutor(max_workers=4)
        
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
        
        print("📊 页数统计管理器已初始化处理器架构")
        self._handler_registry.print_registry_info()
    
    def set_progress_callback(self, callback: Callable[[int, int, str], None]):
        """
        设置进度回调函数
        
        Args:
            callback: 回调函数，参数为(当前进度, 总数, 状态消息)
        """
        self._progress_callback = callback
    
    def cancel_calculation(self):
        """取消页数统计"""
        self._cancel_flag = True
    
    def calculate_all_pages(self, documents: List[Document]) -> PageCountSummary:
        """
        批量计算所有文档的页数（使用新的处理器架构）
        
        Args:
            documents: 文档列表
            
        Returns:
            统计汇总结果
        """
        if not documents:
            return PageCountSummary()
        
        print(f"📊 开始批量页数统计，共 {len(documents)} 个文档")
        self._cancel_flag = False
        
        results = []
        total_docs = len(documents)
        
        try:
            # 使用线程池并行计算页数
            futures = []
            for i, document in enumerate(documents):
                if self._cancel_flag:
                    break
                
                future = self._executor.submit(self._calculate_single_document, document)
                futures.append((i, future))
            
            # 收集结果
            for i, future in futures:
                try:
                    if self._cancel_flag:
                        break
                    
                    # 更新进度
                    if self._progress_callback:
                        self._progress_callback(
                            len(results) + 1, total_docs,
                            f"正在统计: {documents[i].file_name}"
                        )
                    
                    result = future.result(timeout=60)  # 60秒超时
                    results.append(result)
                    
                except Exception as e:
                    # 创建错误结果
                    error_result = PageCountResult(
                        document=documents[i],
                        status=PageCountStatus.ERROR,
                        error_message=f"计算超时或异常: {e}"
                    )
                    results.append(error_result)
            
            # 生成汇总
            summary = self._generate_summary(results)
            
            if self._progress_callback:
                self._progress_callback(
                    len(results), total_docs,
                    f"页数统计完成！总页数: {summary.total_pages}"
                )
            
            print(f"📊 页数统计完成: {summary.success_count}/{len(results)} 成功")
            return summary
            
        except Exception as e:
            print(f"批量页数统计失败: {e}")
            return PageCountSummary()
    
    def _calculate_single_document(self, document: Document) -> PageCountResult:
        """
        计算单个文档的页数（使用新的处理器架构）
        
        Args:
            document: 要计算的文档
            
        Returns:
            计算结果
        """
        start_time = time.time()
        result = PageCountResult(document=document)
        
        try:
            # 检查是否应该跳过此文档
            should_skip, skip_status = self._should_skip_document(document)
            if should_skip:
                result.status = skip_status
                result.error_message = self._get_skip_message(skip_status)
                return result
            
            # 使用处理器注册中心获取相应的处理器
            handler = self._handler_registry.get_handler_by_file_type(document.file_type)
            
            if handler is None:
                result.status = PageCountStatus.ERROR
                result.error_message = f"不支持的文件类型: {document.file_type}"
                return result
            
            if not handler.can_handle_file(document.file_path):
                result.status = PageCountStatus.ERROR
                result.error_message = "处理器无法处理此文件"
                return result
            
            # 更新状态为计算中
            result.status = PageCountStatus.CALCULATING
            
            # 使用处理器计算页数
            print(f"🔧 使用 {handler.get_handler_name()} 统计页数: {document.file_name}")
            page_count = handler.count_pages(document.file_path)
            
            # 验证结果
            if page_count < 0:
                result.status = PageCountStatus.ERROR
                result.error_message = "获取到无效的页数"
            else:
                result.page_count = page_count
                result.status = PageCountStatus.SUCCESS
                print(f"✅ 页数统计成功: {document.file_name} = {page_count} 页")
            
        except Exception as e:
            error_message = self._get_user_friendly_error(e, document.file_type)
            
            if "文件被加密" in error_message:
                result.status = PageCountStatus.SKIPPED_ENCRYPTED
            elif "文件已损坏" in error_message:
                result.status = PageCountStatus.SKIPPED_DAMAGED
            elif "需要安装" in error_message:
                result.status = PageCountStatus.SKIPPED_NO_OFFICE
            elif "无法访问" in error_message:
                result.status = PageCountStatus.SKIPPED_NO_ACCESS
            else:
                result.status = PageCountStatus.ERROR
            
            result.error_message = error_message
            print(f"❌ 页数统计失败: {document.file_name} - {error_message}")
        
        finally:
            result.calculation_time = time.time() - start_time
        
        return result
    
    def _should_skip_document(self, document: Document) -> Tuple[bool, PageCountStatus]:
        """
        检查是否应该跳过文档统计
        
        Args:
            document: 文档对象
            
        Returns:
            (是否跳过, 跳过状态)
        """
        # 检查文件是否存在
        if not document.file_path.exists():
            return True, PageCountStatus.SKIPPED_NO_ACCESS
        
        # 检查文件大小（超过100MB跳过）
        try:
            file_size_mb = document.file_path.stat().st_size / (1024 * 1024)
            if file_size_mb > 100:
                return True, PageCountStatus.SKIPPED_LARGE
        except:
            pass
        
        return False, PageCountStatus.UNKNOWN
    
    def _get_user_friendly_error(self, error: Exception, file_type: FileType) -> str:
        """获取用户友好的错误信息"""
        error_str = str(error)
        
        if "文件被加密" in error_str:
            return "文件被加密，无法统计页数"
        elif "文件已损坏" in error_str:
            return "文件已损坏，无法读取"
        elif "需要安装" in error_str:
            return f"需要安装相应的Office组件来处理{file_type.value}文件"
        else:
            return f"统计页数失败: {error_str}"
    
    def _get_skip_message(self, status: PageCountStatus) -> str:
        """获取跳过原因的用户友好描述"""
        messages = {
            PageCountStatus.SKIPPED_LARGE: "文件过大（>100MB），已跳过",
            PageCountStatus.SKIPPED_ENCRYPTED: "文件被加密，无法访问",
            PageCountStatus.SKIPPED_DAMAGED: "文件已损坏，无法读取",
            PageCountStatus.SKIPPED_NO_ACCESS: "文件不存在或无法访问",
            PageCountStatus.SKIPPED_NO_OFFICE: "缺少必要的Office组件"
        }
        return messages.get(status, "未知原因跳过")
    
    def _generate_summary(self, results: List[PageCountResult]) -> PageCountSummary:
        """生成统计汇总"""
        summary = PageCountSummary()
        summary.total_files = len(results)
        
        for result in results:
            # 统计成功、跳过、错误数量
            if result.status == PageCountStatus.SUCCESS:
                summary.success_count += 1
                summary.total_pages += result.page_count or 0
                
                # 按文件类型统计
                if result.document.file_type == FileType.WORD:
                    summary.word_files += 1
                    summary.word_pages += result.page_count or 0
                elif result.document.file_type == FileType.PPT:
                    summary.ppt_files += 1
                    summary.ppt_pages += result.page_count or 0
                elif result.document.file_type == FileType.EXCEL:
                    summary.excel_files += 1
                    summary.excel_pages += result.page_count or 0
                elif result.document.file_type == FileType.PDF:
                    summary.pdf_files += 1
                    summary.pdf_pages += result.page_count or 0
                elif result.document.file_type == FileType.IMAGE:
                    summary.image_files += 1
                    summary.image_pages += result.page_count or 0
                    
            elif result.status in [
                PageCountStatus.SKIPPED_LARGE,
                PageCountStatus.SKIPPED_ENCRYPTED,
                PageCountStatus.SKIPPED_DAMAGED,
                PageCountStatus.SKIPPED_NO_ACCESS,
                PageCountStatus.SKIPPED_NO_OFFICE
            ]:
                summary.skipped_count += 1
                summary.skipped_files.append(result)
            else:
                summary.error_count += 1
                summary.error_files.append(result)
        
        return summary
    
    def get_supported_file_types(self) -> List[FileType]:
        """获取支持的文件类型列表"""
        return list(self._handler_registry.get_all_supported_file_types())
    
    def get_supported_extensions(self) -> List[str]:
        """获取支持的文件扩展名列表"""
        return sorted(self._handler_registry.get_all_supported_extensions())
    
    def __del__(self):
        """析构函数，清理资源"""
        try:
            if hasattr(self, '_executor'):
                self._executor.shutdown(wait=True)
        except Exception:
            pass 