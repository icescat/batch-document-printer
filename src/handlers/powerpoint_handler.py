"""
PowerPoint文档处理器
负责PowerPoint文件的打印和页数统计
"""
import time
from pathlib import Path
from typing import Set, Optional
from ..core.models import FileType, PrintSettings
from .base_handler import BaseDocumentHandler


class PowerPointDocumentHandler(BaseDocumentHandler):
    """PowerPoint文档处理器"""
    
    def __init__(self):
        """初始化PowerPoint处理器"""
        self._timeout_seconds = 30
    
    def get_supported_file_types(self) -> Set[FileType]:
        """获取支持的文件类型"""
        return {FileType.PPT}
    
    def get_supported_extensions(self) -> Set[str]:
        """获取支持的文件扩展名"""
        return {'.ppt', '.pptx', '.dps'}
    
    def can_handle_file(self, file_path: Path) -> bool:
        """检查是否能处理指定文件"""
        if not self.validate_file_exists(file_path):
            return False
        
        extension = file_path.suffix.lower()
        return extension in self.get_supported_extensions()
    
    def print_document(self, file_path: Path, settings: PrintSettings) -> bool:
        """
        打印PowerPoint文档
        
        Args:
            file_path: PowerPoint文件路径
            settings: 打印设置
            
        Returns:
            是否打印成功
        """
        import time
        
        try:
            import comtypes.client
            
            print(f"开始打印PowerPoint文档: {file_path}")
            
            # 创建PowerPoint应用程序对象
            ppt = comtypes.client.CreateObject("PowerPoint.Application")
            
            # 设置为不可见模式（静默打印）
            try:
                ppt.Visible = False
                ppt.WindowState = 2  # ppWindowMinimized
            except Exception:
                # 如果设置不可见失败，继续执行但会显示界面
                print("⚠️ 无法设置PowerPoint为隐藏模式，将显示界面")
            
            presentation = None
            
            try:
                # 打开演示文稿
                presentation = ppt.Presentations.Open(
                    str(file_path),
                    ReadOnly=True,
                    Untitled=False,
                    WithWindow=False  # 不显示窗口
                )
                
                # 设置打印参数
                try:
                    # 使用更明确的打印机设置方式
                    if settings.printer_name:
                        presentation.PrintOptions.ActivePrinter = settings.printer_name
                        print(f"🖨️ 设置PowerPoint打印机: {settings.printer_name}")
                    
                    presentation.PrintOptions.NumberOfCopies = settings.copies
                    
                    # 验证系统级双面打印设置
                    if settings.duplex:
                        try:
                            import win32print
                            printer_handle = win32print.OpenPrinter(settings.printer_name)
                            try:
                                printer_info = win32print.GetPrinter(printer_handle, 2)
                                devmode = printer_info.get('pDevMode')
                                if devmode and hasattr(devmode, 'Duplex'):
                                    print(f"🔍 PowerPoint打印前验证: 打印机双面设置为 {devmode.Duplex}")
                            finally:
                                win32print.ClosePrinter(printer_handle)
                        except Exception as duplex_check:
                            print(f"⚠️ PowerPoint双面打印验证失败: {duplex_check}")
                    
                    # 强制前台打印，确保打印作业真正发送
                    presentation.PrintOptions.PrintInBackground = False
                    print(f"✅ PowerPoint打印设置完成（强制前台打印）")
                    
                except Exception as print_error:
                    print(f"⚠️ 设置打印参数失败，使用默认设置: {print_error}")
                
                # 执行打印 - 使用简单可靠的方式
                print("📤 正在发送PowerPoint打印作业...")
                presentation.PrintOut()
                print("✅ PowerPoint打印作业已发送到打印机")
                
                # 等待打印队列处理
                time.sleep(3)
                
                print(f"✅ PowerPoint文档打印成功: {file_path.name}")
                return True
                
            finally:
                # 关闭演示文稿
                try:
                    if presentation is not None:
                        presentation.Close()
                        presentation = None
                        print("📁 PowerPoint文档已关闭")
                except Exception as close_error:
                    print(f"⚠️ 关闭PowerPoint文档时出错: {close_error}")
                
                # 退出PowerPoint应用程序
                try:
                    ppt.Quit()
                    print("🔚 PowerPoint应用程序已退出")
                    
                    # 等待PowerPoint完全关闭
                    time.sleep(1)
                    
                except Exception as quit_error:
                    print(f"⚠️ 退出PowerPoint应用程序时出错: {quit_error}")
                
        except ImportError:
            print("❌ 缺少comtypes库，无法打印PowerPoint文档")
            return False
        except Exception as e:
            print(f"❌ PowerPoint文档打印失败: {e}")
            return False
    
    def count_pages(self, file_path: Path) -> int:
        """
        统计PowerPoint文档页数（幻灯片数）
        
        Args:
            file_path: PowerPoint文件路径
            
        Returns:
            幻灯片数量
            
        Raises:
            Exception: 统计失败时抛出异常
        """
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
        """可见模式COM方式（专为.ppt文件设计）"""
        import win32com.client
        import pythoncom
        import time
        
        ppt_app = None
        presentation = None
        
        try:
            # 初始化COM
            pythoncom.CoInitialize()
            
            # 创建PowerPoint应用程序实例（可见模式）
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            ppt_app.Visible = True  # 对于.ppt文件，使用可见模式
            
            try:
                # 打开演示文稿
                presentation = ppt_app.Presentations.Open(
                    str(abs_path),
                    ReadOnly=True,
                    Untitled=False,
                    WithWindow=True
                )
                
                # 获取幻灯片数量
                slides_count = presentation.Slides.Count
                
                # 验证结果的合理性
                if slides_count < 0:
                    raise Exception("获取到无效的幻灯片数量")
                
                return slides_count
                
            except Exception as open_error:
                error_str = str(open_error).lower()
                if ("password" in error_str or "protected" in error_str or 
                    "access" in error_str or "permission" in error_str or
                    "encrypted" in error_str or "locked" in error_str):
                    raise Exception("文件被加密")
                else:
                    raise Exception("文件已损坏")
                
        except ImportError:
            raise Exception("需要安装pywin32库来处理PowerPoint文档")
        except Exception as e:
            if "文件被加密" in str(e) or "文件已损坏" in str(e):
                raise e
            else:
                raise Exception("文件已损坏")
        finally:
            # 强制清理所有COM对象
            try:
                if presentation is not None:
                    presentation.Close()
                    presentation = None
            except:
                pass
            
            try:
                if ppt_app is not None:
                    ppt_app.Quit()
                    ppt_app = None
            except:
                pass
            
            # 清理COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def _try_com_method_standard(self, abs_path: Path) -> int:
        """标准隐藏模式COM方式（专为.pptx文件设计）"""
        import win32com.client
        import pythoncom
        import time
        
        ppt_app = None
        presentation = None
        
        try:
            # 初始化COM
            pythoncom.CoInitialize()
            
            # 创建PowerPoint应用程序实例（隐藏模式）
            ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            ppt_app.Visible = False  # 对于.pptx文件，使用隐藏模式
            
            try:
                # 打开演示文稿
                presentation = ppt_app.Presentations.Open(
                    str(abs_path),
                    ReadOnly=True,
                    Untitled=False,
                    WithWindow=False
                )
                
                # 获取幻灯片数量
                slides_count = presentation.Slides.Count
                
                # 验证结果的合理性
                if slides_count < 0:
                    raise Exception("获取到无效的幻灯片数量")
                
                return slides_count
                
            except Exception as open_error:
                error_str = str(open_error).lower()
                if ("password" in error_str or "protected" in error_str or 
                    "access" in error_str or "permission" in error_str or
                    "encrypted" in error_str or "locked" in error_str):
                    raise Exception("文件被加密")
                else:
                    raise Exception("文件已损坏")
                
        except ImportError:
            raise Exception("需要安装pywin32库来处理PowerPoint文档")
        except Exception as e:
            if "文件被加密" in str(e) or "文件已损坏" in str(e):
                raise e
            else:
                raise Exception("文件已损坏")
        finally:
            # 强制清理所有COM对象
            try:
                if presentation is not None:
                    presentation.Close()
                    presentation = None
            except:
                pass
            
            try:
                if ppt_app is not None:
                    ppt_app.Quit()
                    ppt_app = None
            except:
                pass
            
            # 清理COM
            try:
                pythoncom.CoUninitialize()
            except:
                pass 