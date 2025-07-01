"""
路径工具模块
处理PyInstaller打包后的资源文件路径问题
"""
import os
import sys
from pathlib import Path


def get_resource_path(relative_path: str) -> Path:
    """
    获取资源文件的绝对路径
    
    在开发环境中，使用相对于项目根目录的路径
    在PyInstaller打包的exe中，使用临时解压目录的路径
    
    Args:
        relative_path: 相对于项目根目录的路径
        
    Returns:
        资源文件的绝对路径
    """
    try:
        # PyInstaller创建临时文件夹，并将路径存储在_MEIPASS中
        base_path = sys._MEIPASS
        print(f"🔧 检测到PyInstaller环境，使用临时路径: {base_path}")
    except AttributeError:
        # 开发环境，使用脚本所在目录的上级目录（项目根目录）
        base_path = Path(__file__).parent.parent.parent
        print(f"🔧 开发环境，使用项目根目录: {base_path}")
    
    resource_path = Path(base_path) / relative_path
    print(f"🔧 资源路径解析: {relative_path} -> {resource_path}")
    
    return resource_path


def get_sumatra_pdf_path() -> Path:
    """
    获取SumatraPDF可执行文件的路径
    
    Returns:
        SumatraPDF.exe的绝对路径
    """
    return get_resource_path("external/SumatraPDF/SumatraPDF.exe")


def get_app_icon_path() -> Path:
    """
    获取应用程序图标的路径
    
    Returns:
        应用程序图标的绝对路径
    """
    return get_resource_path("resources/app_icon.ico")


def ensure_resource_exists(resource_path: Path) -> bool:
    """
    检查资源文件是否存在
    
    Args:
        resource_path: 资源文件路径
        
    Returns:
        资源文件是否存在
    """
    exists = resource_path.exists()
    if exists:
        print(f"✅ 资源文件存在: {resource_path}")
    else:
        print(f"❌ 资源文件不存在: {resource_path}")
        
        # 尝试列出父目录内容以帮助调试
        parent_dir = resource_path.parent
        if parent_dir.exists():
            print(f"📁 父目录内容: {list(parent_dir.iterdir())}")
        else:
            print(f"❌ 父目录也不存在: {parent_dir}")
    
    return exists


def get_executable_dir() -> Path:
    """
    获取可执行文件所在目录
    
    Returns:
        可执行文件所在目录的路径
    """
    if getattr(sys, 'frozen', False):
        # PyInstaller打包的exe
        return Path(sys.executable).parent
    else:
        # 开发环境
        return Path(__file__).parent.parent.parent


def debug_paths():
    """调试路径信息"""
    print("=" * 50)
    print("路径调试信息")
    print("=" * 50)
    
    print(f"sys.executable: {sys.executable}")
    print(f"sys.argv[0]: {sys.argv[0]}")
    print(f"__file__: {__file__}")
    print(f"frozen: {getattr(sys, 'frozen', False)}")
    
    if hasattr(sys, '_MEIPASS'):
        print(f"sys._MEIPASS: {sys._MEIPASS}")
    
    print(f"当前工作目录: {os.getcwd()}")
    print(f"可执行文件目录: {get_executable_dir()}")
    
    # 测试关键资源路径
    sumatra_path = get_sumatra_pdf_path()
    print(f"SumatraPDF路径: {sumatra_path}")
    print(f"SumatraPDF存在: {ensure_resource_exists(sumatra_path)}")
    
    print("=" * 50) 