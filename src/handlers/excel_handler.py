"""
Excel文档处理器
负责Excel文件的打印和页数统计
使用xlwings库获得精确的打印页数
"""
import time
from pathlib import Path
from typing import Set, Optional
from ..core.models import FileType, PrintSettings
from .base_handler import BaseDocumentHandler

try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False


class ExcelDocumentHandler(BaseDocumentHandler):
    """Excel文档处理器"""
    
    def __init__(self):
        """初始化Excel处理器"""
        self._timeout_seconds = 60  # Excel文件可能更复杂，给60秒超时
        if not XLWINGS_AVAILABLE:
            raise ImportError("xlwings库未安装，无法处理Excel文件")
    
    def get_supported_file_types(self) -> Set[FileType]:
        """获取支持的文件类型"""
        return {FileType.EXCEL}
    
    def get_supported_extensions(self) -> Set[str]:
        """获取支持的文件扩展名"""
        return {'.xls', '.xlsx', '.et'}
    
    def can_handle_file(self, file_path: Path) -> bool:
        """检查是否能处理指定文件"""
        if not self.validate_file_exists(file_path):
            return False
        
        extension = file_path.suffix.lower()
        return extension in self.get_supported_extensions()
    
    def count_pages(self, file_path: Path) -> int:
        """
        使用xlwings获取Excel精确打印页数
        通过Excel原生API访问分页符信息
        """
        if not self.can_handle_file(file_path):
            return 0
            
        try:
            return self._count_pages_with_xlwings(file_path)
        except Exception as e:
            print(f"Excel页数统计失败 {file_path}: {e}")
            # 如果xlwings失败，至少返回1页避免返回0
            return 1
    
    def _count_pages_with_xlwings(self, file_path: Path) -> int:
        """使用xlwings获取Excel精确打印页数"""
        app = None
        wb = None
        total_pages = 0
        
        try:
            # 创建Excel应用实例，设置为不可见且不显示警告
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False
            
            # 打开工作簿
            wb = app.books.open(str(file_path))
            
            # 遍历所有工作表
            for sheet in wb.sheets:
                try:
                    # 获取水平和垂直分页符数量
                    h_page_breaks = len(sheet.api.HPageBreaks)
                    v_page_breaks = len(sheet.api.VPageBreaks)
                    
                    # 计算页数：(水平分页符+1) × (垂直分页符+1)
                    sheet_pages = (h_page_breaks + 1) * (v_page_breaks + 1)
                    
                    # 确保每个工作表至少有1页
                    if sheet_pages < 1:
                        sheet_pages = 1
                        
                    total_pages += sheet_pages
                    
                except Exception as e:
                    print(f"工作表 {sheet.name} 页数统计失败: {e}")
                    # 如果单个工作表失败，至少计为1页
                    total_pages += 1
            
            return max(total_pages, 1)  # 确保至少返回1页
            
        except Exception as e:
            print(f"xlwings处理Excel文件失败 {file_path}: {e}")
            return 1
            
        finally:
            # 清理资源
            try:
                if wb:
                    wb.close()
                if app:
                    app.quit()
            except:
                pass
    
    def print_document(self, file_path: Path, settings: PrintSettings) -> bool:
        """使用Excel COM接口执行打印"""
        if not self.can_handle_file(file_path):
            return False
            
        app = None
        wb = None
        
        try:
            # 创建Excel应用实例
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False
            
            # 打开工作簿
            wb = app.books.open(str(file_path))
            
            # 设置打印选项
            if settings.printer_name:
                try:
                    # Excel的打印机设置方式 - 使用更可靠的方法
                    current_printer = app.api.ActivePrinter
                    print(f"📋 Excel当前打印机: {current_printer}")
                    
                    app.api.ActivePrinter = settings.printer_name
                    print(f"🖨️ 设置Excel打印机: {settings.printer_name}")
                    
                    # 验证设置是否成功
                    new_printer = app.api.ActivePrinter
                    print(f"✅ Excel打印机设置后: {new_printer}")
                    
                except Exception as printer_error:
                    print(f"⚠️ 设置Excel打印机失败，使用默认打印机: {printer_error}")
                    # 继续使用默认打印机
            
            # 验证系统级双面打印设置
            if settings.duplex and settings.printer_name:
                try:
                    import win32print
                    printer_handle = win32print.OpenPrinter(settings.printer_name)
                    try:
                        printer_info = win32print.GetPrinter(printer_handle, 2)
                        devmode = printer_info.get('pDevMode')
                        if devmode and hasattr(devmode, 'Duplex'):
                            print(f"🔍 Excel打印前验证: 打印机双面设置为 {devmode.Duplex}")
                    finally:
                        win32print.ClosePrinter(printer_handle)
                except Exception as duplex_check:
                    print(f"⚠️ Excel双面打印验证失败: {duplex_check}")
            
            # 执行打印 - 使用更明确的参数
            try:
                print("📤 正在发送Excel打印作业...")
                wb.api.PrintOut(
                    Copies=settings.copies,
                    Collate=True,
                    Preview=False,
                    PrintToFile=False,
                    ActivePrinter=settings.printer_name if settings.printer_name else None
                )
                print(f"✅ Excel打印作业已发送到打印机")
                
                # 等待打印作业处理
                time.sleep(2)
                
            except Exception as print_error:
                print(f"❌ Excel详细打印失败: {print_error}")
                # 尝试简单的打印调用
                try:
                    wb.api.PrintOut(Copies=settings.copies)
                    print("🔄 使用简单方式重试Excel打印")
                except Exception as simple_error:
                    print(f"❌ Excel简单打印也失败: {simple_error}")
                    raise simple_error
            
            return True
            
        except Exception as e:
            print(f"Excel打印失败 {file_path}: {e}")
            return False
            
        finally:
            # 清理资源
            try:
                if wb:
                    wb.close()
                if app:
                    app.quit()
            except:
                pass 