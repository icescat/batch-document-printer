"""
打印工具模块
提供通用的打印机验证和配置函数
"""
import win32print
from typing import Optional
from ..core.models import PrintSettings


def verify_printer_duplex_setting(printer_name: str, handler_name: str = "") -> Optional[int]:
    """
    验证打印机的双面打印设置
    
    Args:
        printer_name: 打印机名称
        handler_name: 处理器名称（用于日志）
        
    Returns:
        双面打印设置值（1=单面, 2=长边, 3=短边），如果失败返回None
    """
    try:
        printer_handle = win32print.OpenPrinter(printer_name)
        try:
            printer_info = win32print.GetPrinter(printer_handle, 2)
            devmode = printer_info.get('pDevMode')
            if devmode and hasattr(devmode, 'Duplex'):
                duplex_value = devmode.Duplex
                duplex_name = {1: "单面", 2: "双面长边", 3: "双面短边"}.get(duplex_value, f"未知({duplex_value})")
                print(f"🔍 {handler_name}打印前验证: 打印机双面设置为 {duplex_value} ({duplex_name})")
                return duplex_value
            else:
                print(f"⚠️ {handler_name}无法获取打印机双面设置")
                return None
        finally:
            win32print.ClosePrinter(printer_handle)
    except Exception as e:
        print(f"⚠️ {handler_name}双面打印验证失败: {e}")
        return None


def log_print_start(file_path, handler_name: str, settings: PrintSettings):
    """
    记录打印开始信息
    
    Args:
        file_path: 文件路径
        handler_name: 处理器名称
        settings: 打印设置
    """
    print(f"📋 使用 {handler_name} 打印文档: {file_path.name}")
    
    # 如果启用了双面打印，验证系统设置
    if settings.duplex and settings.printer_name:
        verify_printer_duplex_setting(settings.printer_name, handler_name)


def log_print_success(file_path, handler_name: str, method: str = ""):
    """
    记录打印成功信息
    
    Args:
        file_path: 文件路径  
        handler_name: 处理器名称
        method: 打印方法（可选）
    """
    method_info = f"（{method}）" if method else ""
    print(f"✅ {handler_name}打印成功{method_info}: {file_path.name}")


def log_print_error(file_path, handler_name: str, error: Exception):
    """
    记录打印错误信息
    
    Args:
        file_path: 文件路径
        handler_name: 处理器名称  
        error: 错误信息
    """
    print(f"❌ {handler_name}打印失败 {file_path.name}: {error}") 