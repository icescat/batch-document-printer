"""
打印设置管理器
负责检测打印机、管理打印设置配置
"""
import win32print
import win32api
import win32gui
from typing import List, Dict, Optional
from .models import PrintSettings, ColorMode, Orientation


class PrinterSettingsManager:
    """打印机设置管理器类"""
    
    # 标准纸张尺寸
    PAPER_SIZES = {
        'A4': (210, 297),
        'A3': (297, 420), 
        'A5': (148, 210),
        'Letter': (216, 279),
        'Legal': (216, 356),
        'Tabloid': (279, 432)
    }
    
    def __init__(self):
        """初始化设置管理器"""
        self._available_printers: List[str] = []
        self._default_printer: Optional[str] = None
        self._refresh_printers()
    
    def _refresh_printers(self):
        """刷新可用打印机列表"""
        try:
            # 获取所有打印机
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
            self._available_printers = [printer[2] for printer in printers]
            
            # 获取默认打印机
            try:
                self._default_printer = win32print.GetDefaultPrinter()
            except:
                self._default_printer = self._available_printers[0] if self._available_printers else None
                
            print(f"发现 {len(self._available_printers)} 台打印机")
            if self._default_printer:
                print(f"默认打印机: {self._default_printer}")
                
        except Exception as e:
            print(f"获取打印机列表失败: {e}")
            self._available_printers = []
            self._default_printer = None
    
    @property
    def available_printers(self) -> List[str]:
        """获取可用打印机列表"""
        return self._available_printers.copy()
    
    @property
    def default_printer(self) -> Optional[str]:
        """获取默认打印机"""
        return self._default_printer
    
    @property
    def paper_sizes(self) -> List[str]:
        """获取支持的纸张尺寸列表"""
        return list(self.PAPER_SIZES.keys())
    
    def validate_printer(self, printer_name: str) -> bool:
        """
        验证打印机是否可用
        
        Args:
            printer_name: 打印机名称
            
        Returns:
            是否可用
        """
        return printer_name in self._available_printers
    
    def get_printer_info(self, printer_name: str) -> Optional[Dict]:
        """
        获取打印机详细信息
        
        Args:
            printer_name: 打印机名称
            
        Returns:
            打印机信息字典
        """
        try:
            if not self.validate_printer(printer_name):
                return None
                
            # 打开打印机
            handle = win32print.OpenPrinter(printer_name)
            
            # 获取打印机信息
            printer_info = win32print.GetPrinter(handle, 2)
            
            # 关闭打印机句柄
            win32print.ClosePrinter(handle)
            
            return {
                'name': printer_info['pPrinterName'],
                'driver': printer_info['pDriverName'],
                'port': printer_info['pPortName'],
                'location': printer_info.get('pLocation', ''),
                'comment': printer_info.get('pComment', ''),
                'status': printer_info['Status']
            }
            
        except Exception as e:
            print(f"获取打印机信息失败 {printer_name}: {e}")
            return None
    
    def test_printer_connection(self, printer_name: str) -> bool:
        """
        测试打印机连接
        
        Args:
            printer_name: 打印机名称
            
        Returns:
            连接是否正常
        """
        try:
            handle = win32print.OpenPrinter(printer_name)
            win32print.ClosePrinter(handle)
            return True
        except Exception as e:
            print(f"打印机连接测试失败 {printer_name}: {e}")
            return False
    
    def create_default_settings(self) -> PrintSettings:
        """
        创建默认打印设置
        
        Returns:
            默认PrintSettings对象
        """
        printer_name = self._default_printer or ""
        return PrintSettings(
            printer_name=printer_name,
            paper_size="A4",
            copies=1,
            duplex=False,
            color_mode=ColorMode.COLOR,
            orientation=Orientation.PORTRAIT
        )
    
    def validate_settings(self, settings: PrintSettings) -> tuple[bool, List[str]]:
        """
        验证打印设置
        
        Args:
            settings: 打印设置对象
            
        Returns:
            (是否有效, 错误信息列表)
        """
        errors = []
        
        # 验证打印机
        if not settings.printer_name:
            errors.append("未选择打印机")
        elif not self.validate_printer(settings.printer_name):
            errors.append(f"打印机不可用: {settings.printer_name}")
        
        # 验证纸张尺寸
        if settings.paper_size not in self.PAPER_SIZES:
            errors.append(f"不支持的纸张尺寸: {settings.paper_size}")
        
        # 验证打印数量
        if settings.copies < 1 or settings.copies > 999:
            errors.append("打印数量必须在1-999之间")
        
        return len(errors) == 0, errors
    
    def get_printer_capabilities(self, printer_name: str) -> Dict:
        """
        获取打印机功能特性
        
        Args:
            printer_name: 打印机名称
            
        Returns:
            功能特性字典
        """
        try:
            if not self.validate_printer(printer_name):
                return {}
            
            # 获取设备能力
            hdc = win32gui.CreateDC("WINSPOOL", printer_name, None)
            
            capabilities = {
                'supports_color': win32gui.GetDeviceCaps(hdc, 12) > 1,  # 12 是 BITSPIXEL 常量
                'supports_duplex': True,  # 大部分现代打印机都支持
                'paper_sizes': self.paper_sizes,
                'max_copies': 999
            }
            
            win32gui.DeleteDC(hdc)
            return capabilities
            
        except Exception as e:
            print(f"获取打印机功能失败 {printer_name}: {e}")
            return {
                'supports_color': True,
                'supports_duplex': True,
                'paper_sizes': self.paper_sizes,
                'max_copies': 999
            }
    
    def refresh_printer_list(self):
        """手动刷新打印机列表"""
        self._refresh_printers()
    
    def set_default_printer(self, printer_name: str) -> bool:
        """
        设置默认打印机
        
        Args:
            printer_name: 打印机名称
            
        Returns:
            是否设置成功
        """
        try:
            if not self.validate_printer(printer_name):
                return False
            
            win32print.SetDefaultPrinter(printer_name)
            self._default_printer = printer_name
            print(f"已设置默认打印机: {printer_name}")
            return True
            
        except Exception as e:
            print(f"设置默认打印机失败: {e}")
            return False 