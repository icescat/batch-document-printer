"""
打印机配置管理器
负责在系统级别管理打印机设置，支持批量打印的统一配置
"""
import win32print
import win32con
from typing import Optional, Dict, Any
from .models import PrintSettings


class PrinterConfigManager:
    """打印机配置管理器 - 系统级设置"""
    
    def __init__(self):
        """初始化配置管理器"""
        self._original_configs: Dict[str, Any] = {}
        self._current_printer: Optional[str] = None
    
    def backup_printer_config(self, printer_name: str) -> bool:
        """
        备份打印机当前配置
        
        Args:
            printer_name: 打印机名称
            
        Returns:
            是否备份成功
        """
        try:
            # 打开打印机
            printer_handle = win32print.OpenPrinter(printer_name)
            
            try:
                # 获取当前打印机信息
                printer_info = win32print.GetPrinter(printer_handle, 2)
                devmode = printer_info.get('pDevMode')
                
                if devmode:
                    # 备份关键设置
                    self._original_configs[printer_name] = {
                        'duplex': getattr(devmode, 'Duplex', 1),
                        'copies': getattr(devmode, 'Copies', 1),
                        'color': getattr(devmode, 'Color', 1),
                        'orientation': getattr(devmode, 'Orientation', 1),
                        'paper_size': getattr(devmode, 'PaperSize', 9),  # A4 = 9
                        'fields': getattr(devmode, 'Fields', 0)
                    }
                    
                    print(f"✅ 已备份打印机配置: {printer_name}")
                    return True
                else:
                    print(f"⚠️ 无法获取打印机配置: {printer_name}")
                    return False
                    
            finally:
                win32print.ClosePrinter(printer_handle)
                
        except Exception as e:
            print(f"❌ 备份打印机配置失败 {printer_name}: {e}")
            return False
    
    def apply_batch_print_settings(self, printer_name: str, settings: PrintSettings) -> bool:
        """
        应用批量打印设置到系统级打印机配置
        
        Args:
            printer_name: 打印机名称
            settings: 打印设置
            
        Returns:
            是否应用成功
        """
        try:
            # 先备份当前配置
            if not self.backup_printer_config(printer_name):
                return False
            
            # 打开打印机
            printer_handle = win32print.OpenPrinter(printer_name)
            
            try:
                # 获取当前设备模式
                printer_info = win32print.GetPrinter(printer_handle, 2)
                devmode = printer_info.get('pDevMode')
                
                if not devmode:
                    # 如果没有devmode，尝试获取默认的
                    devmode = win32print.DocumentProperties(
                        0, printer_handle, printer_name, None, None, 0
                    )
                
                if devmode:
                    # 设置双面打印
                    if settings.duplex:
                        devmode.Duplex = 2  # 2 = 双面长边翻转, 3 = 双面短边翻转
                        devmode.Fields |= win32con.DM_DUPLEX
                        print(f"🔄 设置双面打印: 长边翻转")
                    else:
                        devmode.Duplex = 1  # 1 = 单面打印
                        devmode.Fields |= win32con.DM_DUPLEX
                        print(f"📄 设置单面打印")
                    
                    # 设置打印份数（默认为1，实际份数由应用程序控制）
                    devmode.Copies = 1
                    devmode.Fields |= win32con.DM_COPIES
                    
                    # 设置颜色模式
                    if settings.color_mode.value == "color":
                        devmode.Color = 2  # 2 = 彩色
                    else:
                        devmode.Color = 1  # 1 = 黑白
                    devmode.Fields |= win32con.DM_COLOR
                    
                    # 设置页面方向
                    if settings.orientation.value == "landscape":
                        devmode.Orientation = 2  # 2 = 横向
                    else:
                        devmode.Orientation = 1  # 1 = 纵向
                    devmode.Fields |= win32con.DM_ORIENTATION
                    
                    # 设置纸张大小
                    paper_size_map = {
                        'A4': 9,
                        'A3': 8,
                        'A5': 11,
                        'Letter': 1,
                        'Legal': 5,
                        'Tabloid': 3
                    }
                    devmode.PaperSize = paper_size_map.get(settings.paper_size, 9)
                    devmode.Fields |= win32con.DM_PAPERSIZE
                    
                    # 应用配置 - 使用SetPrinter方式
                    try:
                        # 方法1: 使用DocumentProperties设置
                        win32print.DocumentProperties(
                            0, printer_handle, printer_name, 
                            devmode, devmode, 
                            win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER
                        )
                        
                        # 方法2: 直接通过SetPrinter应用
                        printer_info['pDevMode'] = devmode
                        win32print.SetPrinter(printer_handle, 2, printer_info, 0)
                        print("✅ 使用SetPrinter方式应用配置")
                        
                    except Exception as set_error:
                        print(f"⚠️ SetPrinter方式失败: {set_error}")
                        # 回退到DocumentProperties方式
                        win32print.DocumentProperties(
                            0, printer_handle, printer_name, 
                            devmode, devmode, 
                            win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER
                        )
                    
                    self._current_printer = printer_name
                    print(f"✅ 已应用批量打印配置到系统级打印机: {printer_name}")
                    print(f"   - 双面打印: {'开启' if settings.duplex else '关闭'}")
                    print(f"   - 颜色模式: {settings.color_mode.value}")
                    print(f"   - 页面方向: {settings.orientation.value}")
                    print(f"   - 纸张大小: {settings.paper_size}")
                    
                    # 验证设置是否生效
                    self._verify_printer_settings(printer_handle, printer_name)
                    
                    return True
                else:
                    print(f"❌ 无法获取打印机设备模式: {printer_name}")
                    return False
                    
            finally:
                win32print.ClosePrinter(printer_handle)
                
        except Exception as e:
            print(f"❌ 应用打印机配置失败 {printer_name}: {e}")
            return False
    
    def restore_printer_config(self, printer_name: Optional[str] = None) -> bool:
        """
        恢复打印机原始配置
        
        Args:
            printer_name: 打印机名称，如果为None则恢复当前打印机
            
        Returns:
            是否恢复成功
        """
        target_printer = printer_name or self._current_printer
        
        if not target_printer:
            print("⚠️ 没有需要恢复的打印机配置")
            return True
        
        if target_printer not in self._original_configs:
            print(f"⚠️ 没有找到打印机的备份配置: {target_printer}")
            return False
        
        try:
            # 打开打印机
            printer_handle = win32print.OpenPrinter(target_printer)
            
            try:
                # 获取当前设备模式
                printer_info = win32print.GetPrinter(printer_handle, 2)
                devmode = printer_info.get('pDevMode')
                
                if devmode:
                    # 恢复原始配置
                    original = self._original_configs[target_printer]
                    
                    devmode.Duplex = original['duplex']
                    devmode.Copies = original['copies']
                    devmode.Color = original['color']
                    devmode.Orientation = original['orientation']
                    devmode.PaperSize = original['paper_size']
                    devmode.Fields = original['fields']
                    
                    # 应用恢复的配置
                    win32print.DocumentProperties(
                        0, printer_handle, target_printer,
                        devmode, devmode,
                        win32con.DM_IN_BUFFER | win32con.DM_OUT_BUFFER
                    )
                    
                    print(f"✅ 已恢复打印机原始配置: {target_printer}")
                    
                    # 清理备份
                    del self._original_configs[target_printer]
                    if self._current_printer == target_printer:
                        self._current_printer = None
                    
                    return True
                else:
                    print(f"❌ 无法获取打印机设备模式进行恢复: {target_printer}")
                    return False
                    
            finally:
                win32print.ClosePrinter(printer_handle)
                
        except Exception as e:
            print(f"❌ 恢复打印机配置失败 {target_printer}: {e}")
            return False
    
    def restore_all_configs(self):
        """恢复所有备份的打印机配置"""
        printers_to_restore = list(self._original_configs.keys())
        
        for printer_name in printers_to_restore:
            self.restore_printer_config(printer_name)
        
        print(f"✅ 已恢复所有打印机配置")
    
    def _verify_printer_settings(self, printer_handle, printer_name: str):
        """验证打印机设置是否正确应用"""
        try:
            # 重新读取打印机配置
            printer_info = win32print.GetPrinter(printer_handle, 2)
            devmode = printer_info.get('pDevMode')
            
            if devmode:
                current_duplex = getattr(devmode, 'Duplex', 1)
                current_color = getattr(devmode, 'Color', 1)
                current_orientation = getattr(devmode, 'Orientation', 1)
                
                print(f"🔍 验证打印机设置:")
                print(f"   - 当前双面设置: {current_duplex} (1=单面, 2=长边, 3=短边)")
                print(f"   - 当前颜色设置: {current_color} (1=黑白, 2=彩色)")
                print(f"   - 当前方向设置: {current_orientation} (1=纵向, 2=横向)")
            else:
                print("⚠️ 无法验证打印机设置")
                
        except Exception as e:
            print(f"⚠️ 验证打印机设置时出错: {e}")
    
    def __del__(self):
        """析构函数，确保配置被恢复"""
        try:
            self.restore_all_configs()
        except:
            pass 