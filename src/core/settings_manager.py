"""
打印设置管理器
负责检测打印机、管理打印设置配置
"""
import win32print
import win32api
from typing import List, Dict, Optional, Tuple
from .models import PrintSettings, ColorMode, Orientation

# WIN32 API 常量定义
# DeviceCapabilities功能常量
DC_FIELDS = 1
DC_PAPERS = 2
DC_PAPERSIZE = 3
DC_MINEXTENT = 4
DC_MAXEXTENT = 5
DC_BINS = 6
DC_DUPLEX = 7
DC_SIZE = 8
DC_EXTRA = 9
DC_VERSION = 10
DC_DRIVER = 11
DC_BINNAMES = 12
DC_ENUMRESOLUTIONS = 13
DC_FILEDEPENDENCIES = 14
DC_TRUETYPE = 15
DC_PAPERNAMES = 16
DC_ORIENTATION = 17
DC_COPIES = 18
DC_BINADJUST = 19
DC_EMF_COMPLIANT = 20
DC_DATATYPE_PRODUCED = 21
DC_COLLATE = 22
DC_MANUFACTURER = 23
DC_MODEL = 24
DC_PERSONALITY = 25
DC_PRINTRATE = 26
DC_PRINTRATEUNIT = 27
DC_PRINTERMEM = 28
DC_MEDIAREADY = 29
DC_STAPLE = 30
DC_PRINTRATEPPM = 31
DC_COLORDEVICE = 32
DC_NUP = 33
DC_MEDIATYPES = 34
DC_MEDIATYPENAMES = 35
DC_MAXCOPIES = 36


class PrinterSettingsManager:
    """打印机设置管理器类"""
    
    # 标准纸张尺寸映射表（作为备用）
    STANDARD_PAPER_SIZES = {
        'A4': (210, 297),
        'A3': (297, 420), 
        'A5': (148, 210),
        'B5': (176, 250)
    }
    
    # 常见纸张尺寸代码映射（DMPAPER_* 常量）
    PAPER_CODE_MAPPING = {
        1: 'Letter',
        2: 'Lettersmall',
        3: 'Tabloid',
        4: 'Ledger',
        5: 'Legal',
        6: 'Statement',
        7: 'Executive',
        8: 'A3',
        9: 'A4',
        10: 'A4small',
        11: 'A5',
        12: 'B4',
        13: 'B5',
        14: 'Folio',
        15: 'Quarto',
        16: '10x14',
        17: '11x17',
        18: 'Note',
        19: 'Envelope #9',
        20: 'Envelope #10',
        21: 'Envelope #11',
        22: 'Envelope #12',
        23: 'Envelope #14',
        24: 'C size sheet',
        25: 'D size sheet',
        26: 'E size sheet',
        27: 'Envelope DL',
        28: 'Envelope C5',
        29: 'Envelope C3',
        30: 'Envelope C4',
        31: 'Envelope C6',
        32: 'Envelope C65',
        33: 'Envelope B4',
        34: 'Envelope B5',
        35: 'Envelope B6',
        36: 'Envelope',
        37: 'Envelope Monarch',
        38: '6 3/4 Envelope',
        39: 'US Std Fanfold',
        40: 'German Std Fanfold',
        41: 'German Legal Fanfold',
        42: 'B4 (ISO)',
        43: 'Japanese Postcard',
        44: '9x11',
        45: '10x11',
        46: '15x11',
        47: 'Envelope Invite',
        50: 'Letter Extra',
        51: 'Legal Extra',
        52: 'Tabloid Extra',
        53: 'A4 Extra',
        54: 'Letter Transverse',
        55: 'A4 Transverse',
        56: 'Letter Extra Transverse',
        57: 'A Plus',
        58: 'B Plus',
        59: 'Letter Plus',
        60: 'A4 Plus',
        61: 'A5 Transverse',
        62: 'B5 (JIS) Transverse',
        63: 'A3 Extra',
        64: 'A5 Extra',
        65: 'B5 (ISO) Extra',
        66: 'A2',
        67: 'A3 Transverse',
        68: 'A3 Extra Transverse',
        69: 'Japanese Double Postcard',
        70: 'A6',
        71: 'Japanese Envelope Kaku #2',
        72: 'Japanese Envelope Kaku #3',
        73: 'Japanese Envelope Chou #3',
        74: 'Japanese Envelope Chou #4',
        75: 'Letter Rotated',
        76: 'A3 Rotated',
        77: 'A4 Rotated',
        78: 'A5 Rotated',
        79: 'B4 (JIS) Rotated',
        80: 'B5 (JIS) Rotated',
        81: 'Japanese Postcard Rotated',
        82: 'Double Japanese Postcard Rotated',
        83: 'A6 Rotated',
        84: 'Japanese Envelope Kaku #2 Rotated',
        85: 'Japanese Envelope Kaku #3 Rotated',
        86: 'Japanese Envelope Chou #3 Rotated',
        87: 'Japanese Envelope Chou #4 Rotated',
        88: 'B6 (JIS)',
        89: 'B6 (JIS) Rotated',
        90: '12x11',
        91: 'Japanese Envelope You #4',
        92: 'Japanese Envelope You #4 Rotated',
        93: 'PRC 16K',
        94: 'PRC 32K',
        95: 'PRC 32K(Big)',
        96: 'PRC Envelope #1',
        97: 'PRC Envelope #2',
        98: 'PRC Envelope #3',
        99: 'PRC Envelope #4',
        100: 'PRC Envelope #5',
        101: 'PRC Envelope #6',
        102: 'PRC Envelope #7',
        103: 'PRC Envelope #8',
        104: 'PRC Envelope #9',
        105: 'PRC Envelope #10',
        106: 'PRC 16K Rotated',
        107: 'PRC 32K Rotated',
        108: 'PRC 32K(Big) Rotated',
        109: 'PRC Envelope #1 Rotated',
        110: 'PRC Envelope #2 Rotated',
        111: 'PRC Envelope #3 Rotated',
        112: 'PRC Envelope #4 Rotated',
        113: 'PRC Envelope #5 Rotated',
        114: 'PRC Envelope #6 Rotated',
        115: 'PRC Envelope #7 Rotated',
        116: 'PRC Envelope #8 Rotated',
        117: 'PRC Envelope #9 Rotated',
        118: 'PRC Envelope #10 Rotated'
    }

    def __init__(self):
        """初始化设置管理器"""
        self._available_printers: List[str] = []
        self._default_printer: Optional[str] = None
        self._printer_paper_sizes: Dict[str, List[str]] = {}  # 缓存各打印机的纸张大小
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
                
            # 清空纸张尺寸缓存，强制重新读取
            self._printer_paper_sizes.clear()
                
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
        """获取默认打印机支持的纸张尺寸列表"""
        if self._default_printer:
            return self.get_printer_paper_sizes(self._default_printer)
        return list(self.STANDARD_PAPER_SIZES.keys())
    
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
            duplex=False,  # 默认关闭双面打印
            color_mode=ColorMode.GRAYSCALE,  # 黑白打印
            orientation=Orientation.PORTRAIT  # 竖向（纵向）
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
        
        # 验证纸张尺寸（针对具体打印机）
        if settings.printer_name:
            supported_papers = self.get_printer_paper_sizes(settings.printer_name)
            if settings.paper_size not in supported_papers:
                errors.append(f"打印机 {settings.printer_name} 不支持纸张尺寸: {settings.paper_size}")
        elif settings.paper_size not in self.STANDARD_PAPER_SIZES:
            errors.append(f"不支持的纸张尺寸: {settings.paper_size}")
        
        # 验证打印数量
        if settings.copies < 1 or settings.copies > 999:
            errors.append("打印数量必须在1-999之间")
        
        return len(errors) == 0, errors
    
    def get_printer_paper_sizes(self, printer_name: str) -> List[str]:
        """
        获取指定打印机支持的纸张大小列表
        
        Args:
            printer_name: 打印机名称
            
        Returns:
            纸张大小列表
        """
        if not printer_name or not self.validate_printer(printer_name):
            return list(self.STANDARD_PAPER_SIZES.keys())
        
        # 检查缓存
        if printer_name in self._printer_paper_sizes:
            return self._printer_paper_sizes[printer_name]
        
        try:
            # 获取打印机信息以获取端口名称
            printer_info = self.get_printer_info(printer_name)
            port_name = printer_info.get('port', '') if printer_info else ''
            
            # 使用 DeviceCapabilities 获取纸张尺寸代码和名称
            paper_codes = win32print.DeviceCapabilities(printer_name, port_name, DC_PAPERS)
            paper_names = win32print.DeviceCapabilities(printer_name, port_name, DC_PAPERNAMES)
            
            # 组合代码和名称，优先使用系统返回的名称
            supported_papers = []
            
            if paper_codes and paper_names:
                for i, code in enumerate(paper_codes):
                    if i < len(paper_names):
                        # 使用系统返回的纸张名称，去除空白字符
                        paper_name = paper_names[i].strip()
                        if paper_name:
                            supported_papers.append(paper_name)
                    elif code in self.PAPER_CODE_MAPPING:
                        # 如果没有系统名称，使用映射表
                        supported_papers.append(self.PAPER_CODE_MAPPING[code])
            
            # 去重并排序
            if supported_papers:
                supported_papers = sorted(list(set(supported_papers)))
                print(f"打印机 {printer_name} 支持 {len(supported_papers)} 种纸张格式")
            else:
                # 如果无法获取，使用标准尺寸作为备用
                supported_papers = list(self.STANDARD_PAPER_SIZES.keys())
                print(f"无法获取打印机 {printer_name} 的纸张信息，使用标准格式")
            
            # 缓存结果
            self._printer_paper_sizes[printer_name] = supported_papers
            return supported_papers
            
        except Exception as e:
            print(f"获取打印机 {printer_name} 纸张信息失败: {e}")
            # 返回标准纸张尺寸作为备用
            fallback_papers = list(self.STANDARD_PAPER_SIZES.keys())
            self._printer_paper_sizes[printer_name] = fallback_papers
            return fallback_papers

    def get_printer_paper_details(self, printer_name: str) -> Dict[str, Tuple[float, float]]:
        """
        获取指定打印机的纸张尺寸详细信息（尺寸以毫米为单位）
        
        Args:
            printer_name: 打印机名称
            
        Returns:
            纸张名称到尺寸(宽, 高)的映射字典
        """
        if not printer_name or not self.validate_printer(printer_name):
            return self.STANDARD_PAPER_SIZES
        
        try:
            # 获取打印机信息以获取端口名称
            printer_info = self.get_printer_info(printer_name)
            port_name = printer_info.get('port', '') if printer_info else ''
            
            # 获取纸张尺寸详情
            paper_sizes = win32print.DeviceCapabilities(printer_name, port_name, DC_PAPERSIZE)
            paper_names = win32print.DeviceCapabilities(printer_name, port_name, DC_PAPERNAMES)
            
            paper_details = {}
            
            if paper_sizes and paper_names:
                for i, size_dict in enumerate(paper_sizes):
                    if i < len(paper_names):
                        paper_name = paper_names[i].strip()
                        if paper_name and 'cx' in size_dict and 'cy' in size_dict:
                            # 将1/10毫米转换为毫米
                            width_mm = size_dict['cx'] / 10.0
                            height_mm = size_dict['cy'] / 10.0
                            paper_details[paper_name] = (width_mm, height_mm)
            
            if paper_details:
                print(f"获取到打印机 {printer_name} 的 {len(paper_details)} 种纸张详细信息")
                return paper_details
            else:
                print(f"无法获取打印机 {printer_name} 的纸张详细信息，使用标准信息")
                return self.STANDARD_PAPER_SIZES
                
        except Exception as e:
            print(f"获取打印机 {printer_name} 纸张详细信息失败: {e}")
            return self.STANDARD_PAPER_SIZES

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
                return {'paper_sizes': list(self.STANDARD_PAPER_SIZES.keys())}
            
            capabilities = {
                'supports_color': True,  # 假设支持彩色打印
                'supports_duplex': True,  # 假设支持双面打印
                'paper_sizes': self.get_printer_paper_sizes(printer_name),
                'paper_details': self.get_printer_paper_details(printer_name),
                'max_copies': 999
            }
            
            # 尝试获取更多打印机功能
            try:
                # 获取打印机信息以获取端口名称
                printer_info = self.get_printer_info(printer_name)
                port_name = printer_info.get('port', '') if printer_info else ''
                
                # 检查是否支持双面打印
                duplex_caps = win32print.DeviceCapabilities(printer_name, port_name, DC_DUPLEX)
                capabilities['supports_duplex'] = bool(duplex_caps)
                
                # 检查颜色支持
                color_caps = win32print.DeviceCapabilities(printer_name, port_name, DC_COLORDEVICE)
                capabilities['supports_color'] = bool(color_caps)
                
            except Exception as e:
                print(f"获取打印机 {printer_name} 详细功能时出错: {e}")
            
            return capabilities
            
        except Exception as e:
            print(f"获取打印机功能失败 {printer_name}: {e}")
            return {'paper_sizes': list(self.STANDARD_PAPER_SIZES.keys())}
    
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