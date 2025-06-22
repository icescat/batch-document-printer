"""
页数统计管理器
负责计算不同类型文档的页数，处理各种异常情况
"""
import time
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Callable, Any, Union
from dataclasses import dataclass, field
from enum import Enum
import threading
import signal
import functools

from .models import Document, FileType


def timeout(seconds):
    """超时装饰器"""
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            # 在Windows上使用线程实现超时
            result: List[Any] = [None]
            exception: List[Optional[Exception]] = [None]
            
            def target():
                try:
                    result[0] = func(*args, **kwargs)
                except Exception as e:
                    exception[0] = e
            
            thread = threading.Thread(target=target)
            thread.daemon = True
            thread.start()
            thread.join(timeout=seconds)
            
            if thread.is_alive():
                # 超时了，但我们无法强制终止线程
                # 只能抛出超时异常
                raise Exception(f"文档处理超时（超过{seconds}秒）")
            
            if exception[0] is not None:
                raise exception[0]
            
            return result[0]
        return wrapper
    return decorator


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
    
    # 问题文件列表
    skipped_files: List[PageCountResult] = field(default_factory=list)
    error_files: List[PageCountResult] = field(default_factory=list)
    



class PageCountManager:
    """页数统计管理器"""
    
    def __init__(self):
        """初始化页数统计管理器"""
        self.max_file_size_mb = 100  # 最大文件大小限制（MB）
        self.calculation_timeout = 30  # 单个文件计算超时时间（秒）
        self._progress_callback: Optional[Callable] = None
        self._cancel_flag = False
    
    def set_progress_callback(self, callback: Callable[[int, int, str], None]):
        """
        设置进度回调函数
        
        Args:
            callback: 回调函数，参数为(当前进度, 总数, 状态消息)
        """
        self._progress_callback = callback
    
    def cancel_calculation(self):
        """取消计算"""
        self._cancel_flag = True
    
    def calculate_all_pages(self, documents: List[Document]) -> PageCountSummary:
        """
        计算所有文档的页数
        
        Args:
            documents: 文档列表
            
        Returns:
            页数统计汇总
        """
        self._cancel_flag = False
        results = []
        
        for i, document in enumerate(documents):
            if self._cancel_flag:
                break
                
            # 更新进度
            if self._progress_callback:
                self._progress_callback(
                    i + 1, len(documents),
                    f"{document.file_name}"
                )
            
            # 计算单个文档页数
            result = self._calculate_single_document(document)
            results.append(result)
            
            # 短暂延迟，避免系统过载
            time.sleep(0.1)
        
        # 生成汇总报告
        summary = self._generate_summary(results)
        return summary
    
    def _calculate_single_document(self, document: Document) -> PageCountResult:
        """
        计算单个文档的页数
        
        Args:
            document: 文档对象
            
        Returns:
            页数统计结果
        """
        result = PageCountResult(document=document)
        start_time = time.time()
        
        try:
            # 检查是否应该跳过
            should_skip, skip_reason = self._should_skip_document(document)
            if should_skip:
                result.status = skip_reason
                result.error_message = self._get_skip_message(skip_reason)
                return result
            
            result.status = PageCountStatus.CALCULATING
            
            # 根据文件类型计算页数
            if document.file_type == FileType.PDF:
                page_count = self._calculate_pdf_pages(document.file_path)
            elif document.file_type == FileType.WORD:
                page_count = self._calculate_word_pages(document.file_path)
            elif document.file_type == FileType.PPT:
                page_count = self._calculate_ppt_pages(document.file_path)
            elif document.file_type == FileType.EXCEL:
                page_count = self._calculate_excel_pages(document.file_path)
            else:
                raise ValueError(f"不支持的文件类型: {document.file_type}")
            
            result.page_count = page_count
            result.status = PageCountStatus.SUCCESS
            
        except Exception as e:
            result.status = PageCountStatus.ERROR
            result.error_message = self._get_user_friendly_error(e, document.file_type)[:150]  # 限制错误消息长度
        
        finally:
            result.calculation_time = time.time() - start_time
        
        return result
    
    def _should_skip_document(self, document: Document) -> Tuple[bool, PageCountStatus]:
        """
        检查是否应该跳过文档
        
        Args:
            document: 文档对象
            
        Returns:
            (是否跳过, 跳过原因)
        """
        # 检查文件是否存在
        if not document.file_path.exists():
            return True, PageCountStatus.SKIPPED_NO_ACCESS
        
        # 检查文件大小
        file_size_mb = document.file_path.stat().st_size / (1024 * 1024)
        if file_size_mb > self.max_file_size_mb:
            return True, PageCountStatus.SKIPPED_LARGE
        
        # 检查文件访问权限
        try:
            with open(document.file_path, 'rb') as f:
                f.read(1)  # 尝试读取1字节
        except (PermissionError, OSError):
            return True, PageCountStatus.SKIPPED_NO_ACCESS
        
        return False, PageCountStatus.UNKNOWN
    
    def _get_user_friendly_error(self, error: Exception, file_type: FileType) -> str:
        """将技术错误转换为用户友好的错误信息（简化版）"""
        error_str = str(error).lower()
        
        # 简化的错误分类：只区分加密和损坏两种情况
        if ("password" in error_str or "protected" in error_str or "encrypted" in error_str or 
            "access" in error_str or "permission" in error_str):
            return "文件被加密或权限受限"
        else:
            return "文件已损坏或格式不正确"
    
    def _get_skip_message(self, status: PageCountStatus) -> str:
        """获取跳过原因的描述"""
        messages = {
            PageCountStatus.SKIPPED_LARGE: f"文件过大 (>{self.max_file_size_mb}MB)",
            PageCountStatus.SKIPPED_ENCRYPTED: "文件已加密或需要密码",
            PageCountStatus.SKIPPED_DAMAGED: "文件损坏或格式错误",
            PageCountStatus.SKIPPED_NO_ACCESS: "文件无法访问或权限不足",
            PageCountStatus.SKIPPED_NO_OFFICE: "Office应用程序不可用"
        }
        return messages.get(status, "未知原因")
    
    @timeout(30)  # 30秒超时
    def _calculate_pdf_pages(self, file_path: Path) -> int:
        """计算PDF文件页数（改进版本）"""
        try:
            # 首先尝试使用PyPDF2
            import PyPDF2
            
            with open(file_path, 'rb') as file:
                try:
                    reader = PyPDF2.PdfReader(file)
                    
                    # 检查是否加密
                    if reader.is_encrypted:
                        raise Exception("PDF文件已加密，需要密码")
                    
                    page_count = len(reader.pages)
                    
                    # 验证结果的合理性
                    if page_count < 0:
                        raise Exception("获取到无效的页数")
                    
                    return page_count
                    
                except Exception as pdf_error:
                    error_str = str(pdf_error).lower()
                    if "encrypted" in error_str or "password" in error_str:
                        raise Exception("文件被加密")
                    else:
                        raise Exception("文件已损坏")
                        
        except ImportError:
            raise Exception("需要安装PyPDF2库来处理PDF文件")
        except Exception as e:
            # 简化的错误处理
            error_str = str(e).lower()
            if "文件被加密" in str(e) or "文件已损坏" in str(e):
                # 重新抛出我们的自定义错误
                raise e
            elif "encrypted" in error_str or "password" in error_str or "permission" in error_str or "access" in error_str:
                raise Exception("文件被加密")
            else:
                raise Exception("文件已损坏")
    
    @timeout(30)  # 30秒超时
    def _calculate_word_pages(self, file_path: Path) -> int:
        """计算Word文档页数（强化密码保护处理）"""
        import win32com.client
        import pythoncom
        import time
        
        word_app = None
        document = None
        
        try:
            # 初始化COM
            pythoncom.CoInitialize()
            
            # 创建Word应用程序实例
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False  # 禁用所有警告对话框，包括密码对话框
            
            try:
                # 使用更强化的密码检测方式
                document = word_app.Documents.Open(
                    FileName=str(file_path),
                    ReadOnly=True,              # 只读模式
                    AddToRecentFiles=False,     # 不添加到最近文件
                    Visible=False,              # 不显示文档
                    OpenAndRepair=False,        # 不尝试修复
                    NoEncodingDialog=True,      # 不显示编码对话框
                    PasswordDocument="",        # 空密码，快速检测是否需要密码
                    PasswordTemplate="",        # 空模板密码
                    ConfirmConversions=False,   # 不确认转换
                    Revert=False               # 不恢复
                )
                
                # 验证文档是否成功打开
                if document is None:
                    raise Exception("文件被加密")
                
                # 尝试访问文档内容，进一步验证是否需要密码
                try:
                    # 尝试获取页数统计，如果需要密码这里会立即失败
                    pages = document.ComputeStatistics(2)  # 2 = wdStatisticPages
                except Exception:
                    # 如果无法获取统计信息，通常是密码保护
                    raise Exception("文件被加密")
                
                # 验证结果的合理性
                if pages < 0:
                    raise Exception("获取到无效的页数")
                
                return pages
                
            except Exception as open_error:
                error_str = str(open_error).lower()
                
                # 密码相关错误的快速识别
                if ("password" in error_str or "protected" in error_str or 
                    "access" in error_str or "permission" in error_str or
                    "encrypted" in error_str or "locked" in error_str):
                    raise Exception("文件被加密")
                else:
                    raise Exception("文件已损坏")
                
        except ImportError:
            raise Exception("需要安装pywin32库来处理Word文档")
        except Exception as e:
            if "文件被加密" in str(e) or "文件已损坏" in str(e):
                # 重新抛出我们的自定义错误
                raise e
            else:
                # 其他错误统一处理为文件已损坏
                raise Exception("文件已损坏")
        finally:
            # 强制清理所有COM对象
            try:
                if document is not None:
                    document.Close(SaveChanges=False)
                    document = None
            except:
                pass
            
            try:
                if word_app is not None:
                    # 强制关闭所有打开的文档
                    try:
                        for doc in word_app.Documents:
                            try:
                                doc.Close(SaveChanges=False)
                            except: pass
                    except: pass
                    word_app.Quit()
                    word_app = None
            except:
                pass
            
            # 强制垃圾回收，清理COM对象
            try:
                import gc
                gc.collect()
                time.sleep(0.1)  # 给系统一点时间清理
            except:
                pass
            
            # 清理COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    

    
    @timeout(30)  # 30秒超时
    def _calculate_ppt_pages(self, file_path: Path) -> int:
        """计算PowerPoint幻灯片数（根据文件格式选择最佳方法）"""
        file_extension = file_path.suffix.lower()
        
        # 对于.ppt文件，优先使用COM方式（兼容性更好）
        if file_extension == '.ppt':
            try:
                return self._calculate_ppt_pages_com(file_path)
            except Exception as com_error:
                # 如果COM失败，尝试其他方法
                try:
                    # 先尝试python-pptx（虽然兼容性有限）
                    return self._calculate_ppt_pages_pptx(file_path)
                except Exception:
                    # 所有方法都失败
                    raise Exception(f"无法统计.ppt文件页数。COM方式失败: {com_error}")
        
        # 对于.pptx文件，优先使用python-pptx（更稳定）
        elif file_extension == '.pptx':
            try:
                return self._calculate_ppt_pages_pptx(file_path)
            except Exception as pptx_error:
                # 如果python-pptx失败，尝试COM方式作为备用
                try:
                    return self._calculate_ppt_pages_com(file_path)
                except Exception:
                    # 如果两种方式都失败，抛出python-pptx的错误
                    raise pptx_error
        
        # 未知扩展名，尝试两种方式
        else:
            try:
                # 先尝试python-pptx
                return self._calculate_ppt_pages_pptx(file_path)
            except Exception:
                # 再尝试COM方式
                return self._calculate_ppt_pages_com(file_path)
    
    def _calculate_ppt_pages_pptx(self, file_path: Path) -> int:
        """使用python-pptx库计算PowerPoint幻灯片数"""
        try:
            from pptx import Presentation
            
            # 使用python-pptx库
            prs = Presentation(file_path)
            slides_count = len(prs.slides)
            
            # 验证结果
            if slides_count < 0:
                raise Exception("获取到无效的幻灯片数量")
            
            return slides_count
            
        except ImportError:
            raise Exception("需要安装python-pptx库来处理PowerPoint文档")
        except Exception as e:
            error_str = str(e).lower()
            if "password" in error_str or "protected" in error_str or "permission" in error_str or "access" in error_str:
                raise Exception("文件被加密")
            else:
                raise Exception("文件已损坏")
    
    def _calculate_ppt_pages_com(self, file_path: Path) -> int:
        """PowerPoint页数统计的COM方法（优化版）"""
        import win32com.client
        import pythoncom
        import time
        import os
        
        # 验证文件路径
        if not file_path.exists():
            raise Exception("PowerPoint文档不存在")
        
        # 转换为绝对路径，避免路径问题
        abs_path = file_path.resolve()
        file_extension = file_path.suffix.lower()
        
        # 对于.ppt文件，使用可见模式（经验证最有效）
        if file_extension == '.ppt':
            return self._try_com_method_visible(abs_path)
        else:
            # 对于.pptx文件，使用标准隐藏模式
            return self._try_com_method_standard(abs_path)
    
    def _try_com_method_visible(self, abs_path: Path) -> int:
        """可见模式COM方式（专为.ppt文件设计，基于用户成功经验）"""
        import win32com.client
        import pythoncom
        import time
        
        ppt_app = None
        presentation = None
        
        try:
            pythoncom.CoInitialize()
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            
            # 关键：让PowerPoint可见，这对老版本.ppt文件的兼容性至关重要
            ppt_app.Visible = True
            ppt_app.DisplayAlerts = False  # 禁用密码对话框等警告
            
            # 给PowerPoint启动一点时间
            time.sleep(0.3)
            
            # 尝试打开文件，如果遇到密码保护会快速失败
            try:
                presentation = ppt_app.Presentations.Open(
                    FileName=str(abs_path),
                    ReadOnly=True,
                    Untitled=False,
                    WithWindow=True,  # 允许创建窗口
                    OpenAndRepair=False  # 不尝试修复，避免额外对话框
                )
            except Exception as open_error:
                # 如果打开失败（可能是密码保护），立即抛出异常
                error_str = str(open_error).lower()
                if "password" in error_str or "protected" in error_str:
                    raise Exception("文件被加密")
                else:
                    raise Exception("文件已损坏")
            
            if presentation is None:
                raise Exception("文档打开失败")
            
            # 等待文件完全加载，但限制时间
            time.sleep(0.5)
            
            # 尝试获取幻灯片数量，如果失败说明文件有问题
            try:
                slides_count = presentation.Slides.Count
            except Exception:
                # 如果无法获取幻灯片数量，可能是密码保护或损坏
                raise Exception("文件被加密")
            
            # 验证结果有效性
            if slides_count <= 0:
                raise Exception("获取到的幻灯片数量无效")
            
            return slides_count
            
        finally:
            # 强制关闭所有相关对象，防止卡住
            try:
                if presentation: 
                    presentation.Close()
                    time.sleep(0.2)
            except: pass
            try:
                if ppt_app: 
                    # 强制关闭所有打开的演示文稿
                    try:
                        for pres in ppt_app.Presentations:
                            try:
                                pres.Close()
                            except: pass
                    except: pass
                    ppt_app.Quit()
                    time.sleep(0.3)
            except: pass
            try:
                import gc
                gc.collect()
                time.sleep(0.2)
                pythoncom.CoUninitialize()
            except: pass
    
    def _try_com_method_standard(self, abs_path: Path) -> int:
        """标准COM方式（隐藏模式）"""
        import win32com.client
        import pythoncom
        import time
        
        ppt_app = None
        presentation = None
        
        try:
            pythoncom.CoInitialize()
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            
            # 标准隐藏模式
            ppt_app.Visible = False
            ppt_app.DisplayAlerts = False  # 禁用所有对话框
            
            # 尝试打开文件，添加错误处理
            try:
                presentation = ppt_app.Presentations.Open(
                    FileName=str(abs_path),
                    ReadOnly=True,
                    Untitled=False,
                    WithWindow=False,
                    OpenAndRepair=False
                )
            except Exception as open_error:
                error_str = str(open_error).lower()
                if "password" in error_str or "protected" in error_str:
                    raise Exception("文件被加密")
                else:
                    raise Exception("文件已损坏")
            
            if presentation is None:
                raise Exception("文档打开失败")
            
            # 尝试获取幻灯片数量
            try:
                slides_count = presentation.Slides.Count
            except Exception:
                raise Exception("文件被加密")
            
            return slides_count
            
        finally:
            # 强制关闭所有相关对象
            try:
                if presentation: 
                    presentation.Close()
                    time.sleep(0.1)
            except: pass
            try:
                if ppt_app: 
                    # 强制关闭所有打开的演示文稿
                    try:
                        for pres in ppt_app.Presentations:
                            try:
                                pres.Close()
                            except: pass
                    except: pass
                    ppt_app.Quit()
                    time.sleep(0.2)
            except: pass
            try:
                import gc
                gc.collect()
                time.sleep(0.1)
                pythoncom.CoUninitialize()
            except: pass
    

    
    @timeout(60)  # Excel文件可能更复杂，给60秒超时
    def _calculate_excel_pages(self, file_path: Path) -> int:
        """计算Excel表格打印页数（强化密码保护处理）"""
        excel_app = None
        workbook = None
        
        try:
            import win32com.client
            
            # 创建Excel应用程序实例
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False  # 禁用所有警告对话框，包括密码对话框
            
            try:
                # 使用强化的密码检测方式打开工作簿
                workbook = excel_app.Workbooks.Open(
                    Filename=str(file_path), 
                    ReadOnly=True,
                    Password="",  # 空密码，快速检测是否需要密码
                    WriteResPassword="",  # 空写入密码
                    IgnoreReadOnlyRecommended=True,  # 忽略只读建议
                    Origin=None,  # 不指定来源
                    Delimiter=None,  # 不指定分隔符
                    Editable=False,  # 不可编辑
                    Notify=False,  # 不通知
                    Converter=None,  # 不转换
                    AddToMru=False,  # 不添加到最近使用
                    Local=False,  # 不使用本地设置
                    CorruptLoad=0  # 不尝试修复损坏的文件
                )
                
                # 验证工作簿是否成功打开
                if workbook is None:
                    raise Exception("文件被加密")
                    
            except Exception as open_error:
                error_str = str(open_error).lower()
                
                # 密码相关错误的快速识别
                if ("password" in error_str or "protected" in error_str or 
                    "access" in error_str or "permission" in error_str or
                    "encrypted" in error_str or "locked" in error_str):
                    raise Exception("文件被加密")
                else:
                    raise Exception("文件已损坏")
            
            total_pages = 0
            
            # 遍历所有工作表
            for worksheet in workbook.Worksheets:
                try:
                    # 检查工作表是否有内容
                    used_range = worksheet.UsedRange
                    if used_range is None or used_range.Rows.Count == 0:
                        continue
                    
                    # 尝试获取实际的打印页数（最准确的方法）
                    try:
                        # 设置页面设置为默认值（A4纸张）
                        page_setup = worksheet.PageSetup
                        
                        # 获取水平和垂直分页数
                        h_page_breaks = worksheet.HPageBreaks.Count + 1  # 水平分页数
                        v_page_breaks = worksheet.VPageBreaks.Count + 1  # 垂直分页数
                        
                        # 计算总页数
                        worksheet_pages = h_page_breaks * v_page_breaks
                        
                        # 如果没有分页符，使用改进的估算方法
                        if worksheet_pages <= 1:
                            rows = used_range.Rows.Count
                            cols = used_range.Columns.Count
                            
                            # 根据页面设置估算
                            # A4纸张大约能容纳45行和8列（取决于字体大小）
                            estimated_row_pages = max(1, (rows + 44) // 45)
                            estimated_col_pages = max(1, (cols + 7) // 8)
                            worksheet_pages = estimated_row_pages * estimated_col_pages
                        
                        total_pages += worksheet_pages
                        
                    except Exception:
                        # 如果获取分页信息失败，使用简化估算
                        rows = used_range.Rows.Count
                        cols = used_range.Columns.Count
                        
                        # 改进的估算：考虑行数和列数
                        if rows <= 45 and cols <= 8:
                            # 单页内容
                            worksheet_pages = 1
                        else:
                            # 多页内容
                            row_pages = max(1, (rows + 44) // 45)
                            col_pages = max(1, (cols + 7) // 8)
                            worksheet_pages = row_pages * col_pages
                        
                        total_pages += worksheet_pages
                        
                except Exception:
                    # 如果某个工作表完全无法访问，估算为1页
                    total_pages += 1
            
            return max(1, total_pages)  # 至少1页
                
        except ImportError:
            raise Exception("需要安装pywin32库来处理Excel文档")
        except Exception as e:
            error_str = str(e).lower()
            if "password" in error_str or "protected" in error_str or "access" in error_str or "permission" in error_str:
                raise Exception("文件被加密")
            else:
                raise Exception("文件已损坏")
        finally:
            # 清理资源
            try:
                if workbook:
                    workbook.Close(SaveChanges=False)
            except:
                pass
            
            try:
                if excel_app:
                    excel_app.Quit()
            except:
                pass
    
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

 