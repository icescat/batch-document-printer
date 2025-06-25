"""
Word文档处理器
负责Word文件的打印和页数统计
"""
import time
from pathlib import Path
from typing import Set, Optional
from ..core.models import FileType, PrintSettings
from .base_handler import BaseDocumentHandler


class WordDocumentHandler(BaseDocumentHandler):
    """Word文档处理器"""
    
    def __init__(self):
        """初始化Word处理器"""
        self._timeout_seconds = 30
    
    def get_supported_file_types(self) -> Set[FileType]:
        """获取支持的文件类型"""
        return {FileType.WORD}
    
    def get_supported_extensions(self) -> Set[str]:
        """获取支持的文件扩展名"""
        return {'.doc', '.docx', '.wps'}
    
    def can_handle_file(self, file_path: Path) -> bool:
        """检查是否能处理指定文件"""
        if not self.validate_file_exists(file_path):
            return False
        
        extension = file_path.suffix.lower()
        return extension in self.get_supported_extensions()
    
    def print_document(self, file_path: Path, settings: PrintSettings) -> bool:
        """
        打印Word文档
        
        Args:
            file_path: Word文件路径
            settings: 打印设置
            
        Returns:
            是否打印成功
        """
        try:
            import comtypes.client
            
            print(f"开始打印Word文档: {file_path}")
            
            # 创建Word应用程序对象
            word = comtypes.client.CreateObject("Word.Application")
            # 保持Word可见性设置，Word通常可以隐藏
            word.Visible = False
            
            try:
                # 打开文档
                doc = word.Documents.Open(str(file_path))
                
                # 设置打印参数
                try:
                    # 设置打印机
                    word.ActivePrinter = settings.printer_name
                    
                    # 尝试强制应用双面打印设置到Word
                    if settings.duplex:
                        try:
                            # 通过Word的Application对象设置打印机属性
                            import win32print
                            import win32con
                            
                            # 获取当前打印机的DEVMODE并强制设置双面打印
                            printer_name = settings.printer_name
                            printer_handle = win32print.OpenPrinter(printer_name)
                            try:
                                printer_info = win32print.GetPrinter(printer_handle, 2)
                                devmode = printer_info.get('pDevMode')
                                if devmode and hasattr(devmode, 'Duplex'):
                                    print(f"🔍 Word打印前验证: 打印机双面设置为 {devmode.Duplex}")
                            finally:
                                win32print.ClosePrinter(printer_handle)
                        except Exception as duplex_check:
                            print(f"⚠️ Word双面打印验证失败: {duplex_check}")
                    
                    # 使用简化的打印调用，依赖系统级打印机配置
                    # (双面打印等设置已通过PrinterConfigManager在系统级配置)
                    doc.PrintOut(
                        Copies=settings.copies,
                        Collate=True,
                        PrintToFile=False
                    )
                    print(f"✅ Word文档打印调用完成，使用系统级配置")
                except Exception as print_error:
                    print(f"设置打印参数失败，使用默认设置: {print_error}")
                    # 如果设置打印选项失败，使用默认设置
                    doc.PrintOut()
                
                # 关闭文档
                doc.Close(SaveChanges=False)
                
                print(f"✅ Word文档打印成功: {file_path.name}")
                return True
                
            finally:
                # 退出Word应用程序
                try:
                    word.Quit()
                except Exception as quit_error:
                    print(f"退出Word应用程序时出错: {quit_error}")
                
        except ImportError:
            print("❌ 缺少comtypes库，无法打印Word文档")
            return False
        except Exception as e:
            print(f"❌ Word文档打印失败: {e}")
            return False
    
    def count_pages(self, file_path: Path) -> int:
        """
        统计Word文档页数
        
        Args:
            file_path: Word文件路径
            
        Returns:
            页数
            
        Raises:
            Exception: 统计失败时抛出异常
        """
        import win32com.client
        import pythoncom
        
        word_app = None
        document = None
        
        try:
            pythoncom.CoInitialize()
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False
            
            try:
                document = word_app.Documents.Open(
                    FileName=str(file_path),
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    Visible=False,
                    OpenAndRepair=False,
                    NoEncodingDialog=True,
                    PasswordDocument="",
                    PasswordTemplate="",
                    ConfirmConversions=False,
                    Revert=False
                )
                
                if document is None:
                    raise Exception("文件被加密")
                
                try:
                    pages = document.ComputeStatistics(2)  # wdStatisticPages
                except Exception:
                    raise Exception("文件被加密")
                
                if pages < 0:
                    raise Exception("获取到无效的页数")
                
                return pages
                
            except Exception as open_error:
                error_str = str(open_error).lower()
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
                raise e
            else:
                raise Exception("文件已损坏")
        finally:
            # 清理资源
            try:
                if document is not None:
                    document.Close(SaveChanges=False)
            except:
                pass
            
            try:
                if word_app is not None:
                    word_app.Quit()
            except:
                pass
            
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    

 