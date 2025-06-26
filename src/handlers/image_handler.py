"""
图片文档处理器
使用SumatraPDF实现图片打印，支持各种图片格式的高效打印和页数统计
"""
import os
import sys
import subprocess
from pathlib import Path
from typing import List, Set, Optional, Dict, Any

# 添加项目根目录到Python路径
project_root = Path(__file__).parents[2]
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.core.models import FileType, Document
from src.handlers.base_handler import BaseDocumentHandler

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False


class ImageDocumentHandler(BaseDocumentHandler):
    """基于SumatraPDF的图片文档处理器"""
    
    def __init__(self):
        """初始化图片处理器"""
        super().__init__()
        
        # 支持的图片格式（SumatraPDF支持的格式）
        self._supported_extensions = {
            '.jpg', '.jpeg', '.png', '.gif', '.webp', 
            '.tiff', '.tif', '.tga', '.bmp', '.dib'
        }
        self._supported_file_types = {FileType.IMAGE}
        
        # SumatraPDF路径
        self._sumatra_path = project_root / "external" / "SumatraPDF" / "SumatraPDF.exe"
        self._sumatra_available = self._sumatra_path.exists()
        
        if not self._sumatra_available:
            print("⚠️ SumatraPDF不可用，图片打印将使用Windows系统方案")
            # 回退到Windows系统支持的格式
            self._supported_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif'}
        
        # 如果PIL不可用，进一步限制格式
        if not PIL_AVAILABLE:
            print("⚠️ PIL/Pillow 未安装，多页TIFF支持受限")
    
    def get_handler_name(self) -> str:
        """获取处理器名称"""
        if self._sumatra_available:
            return "图片文档处理器 (SumatraPDF)"
        return "图片文档处理器 (Windows系统)"
    
    def get_supported_file_types(self) -> Set[FileType]:
        """获取支持的文件类型"""
        return self._supported_file_types.copy()
    
    def get_supported_extensions(self) -> Set[str]:
        """获取支持的文件扩展名"""
        return self._supported_extensions.copy()
    
    def can_handle_file(self, file_path: Path) -> bool:
        """
        检查是否可以处理指定文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否可以处理
        """
        if not file_path.exists() or not file_path.is_file():
            return False
        
        extension = file_path.suffix.lower()
        if extension not in self._supported_extensions:
            return False
        
        # 基本的图片文件验证
        try:
            # 检查文件大小（过滤掉异常小的文件）
            if file_path.stat().st_size < 10:
                return False
            
            # 如果有PIL，进行更严格的验证
            if PIL_AVAILABLE:
                try:
                    with Image.open(file_path) as img:
                        img.verify()
                    return True
                except Exception:
                    return False
            else:
                # 没有PIL时，仅基于扩展名判断
                return True
                
        except Exception as e:
            print(f"验证图片文件失败 {file_path}: {e}")
            return False
    
    def count_pages(self, file_path: Path) -> int:
        """
        统计图片文件页数
        对于多数图片文件返回1页，但TIFF可能有多页
        
        Args:
            file_path: 文件路径
            
        Returns:
            页数
        """
        if not self.can_handle_file(file_path):
            raise ValueError(f"无法处理的图片文件: {file_path}")
        
        # 检查是否为TIFF格式
        extension = file_path.suffix.lower()
        if extension in ['.tiff', '.tif']:
            return self._count_tiff_pages(file_path)
        
        # 其他图片格式固定为1页
        return 1
    
    def _count_tiff_pages(self, file_path: Path) -> int:
        """
        统计TIFF文件的页数
        
        Args:
            file_path: TIFF文件路径
            
        Returns:
            页数
        """
        try:
            if PIL_AVAILABLE:
                # 使用PIL计算TIFF页数
                with Image.open(file_path) as img:
                    page_count = 0
                    try:
                        while True:
                            img.seek(page_count)
                            page_count += 1
                    except EOFError:
                        # 到达文件末尾，正常结束
                        pass
                    
                    return max(page_count, 1)  # 至少返回1页
            else:
                # 没有PIL时，TIFF文件假设为1页
                print(f"⚠️ 缺少PIL库，无法准确统计TIFF页数，假设为1页: {file_path.name}")
                return 1
                
        except Exception as e:
            print(f"统计TIFF页数失败 {file_path}: {e}")
            return 1  # 失败时返回1页
    
    def print_document(self, file_path: Path, settings: Any) -> bool:
        """
        打印图片文档 - 智能打印策略：SumatraPDF优先，Windows系统备用
        
        Args:
            file_path: 文件路径
            settings: 打印设置
            
        Returns:
            打印是否成功
        """
        try:
            if not self.can_handle_file(file_path):
                raise ValueError(f"无法处理的图片文件: {file_path}")
            
            # 主策略：SumatraPDF打印（优先）
            if self._sumatra_available:
                print(f"🎯 使用SumatraPDF打印图片: {file_path.name}")
                success = self._print_with_sumatra(file_path, settings)
                if success:
                    return True
                print("⚠️ SumatraPDF打印失败，启用Windows系统备用方案...")
            
            # 备用策略：Windows系统方案
            print(f"🔄 使用Windows系统方案打印图片: {file_path.name}")
            return self._print_with_windows_system(file_path, settings)
                
        except Exception as e:
            print(f"✗ 打印图片文档失败 {file_path.name}: {e}")
            return False
    
    def _print_with_sumatra(self, file_path: Path, settings: Any) -> bool:
        """使用SumatraPDF打印图片"""
        try:
            # 构建SumatraPDF命令
            cmd = [
                str(self._sumatra_path),
                "-print-to", settings.printer_name,
                "-silent"
            ]
            
            # 图片打印设置
            print_settings = []
            print_settings.append(settings.duplex_mode_str)
            print_settings.append(settings.orientation_str)
            
            if print_settings:
                cmd.extend(["-print-settings", ",".join(print_settings)])
            
            cmd.append(str(file_path))
            
            # 执行打印
            result = subprocess.run(
                cmd, 
                capture_output=True, 
                text=True, 
                timeout=30
            )
            
            if result.returncode == 0:
                print(f"✓ SumatraPDF图片打印成功: {file_path.name}")
                return True
            else:
                print(f"✗ SumatraPDF图片打印失败: {file_path.name}")
                if result.stderr:
                    print(f"   错误信息: {result.stderr}")
                return False
                
        except Exception as e:
            print(f"✗ SumatraPDF图片打印异常: {e}")
            return False
    
    def _print_with_windows_system(self, file_path: Path, settings: Any) -> bool:
        """使用Windows系统方案打印图片（备用）"""
        try:
            # 检查是否为多页TIFF
            extension = file_path.suffix.lower()
            if extension in ['.tiff', '.tif']:
                page_count = self._count_tiff_pages(file_path)
                if page_count > 1:
                    print(f"⚠️ 检测到多页TIFF文件: {file_path.name} ({page_count}页)")
                    print("   注意: Windows系统方案可能只能打印第一页")
            
            if os.name == 'nt':  # Windows系统
                try:
                    # 使用Windows的默认打印命令
                    result = subprocess.run([
                        'rundll32.exe',
                        'shimgvw.dll,ImageView_PrintTo',
                        str(file_path),
                        settings.printer_name
                    ], capture_output=True, text=True, timeout=30)
                    
                    if result.returncode == 0:
                        print(f"✓ Windows系统图片打印成功: {file_path.name}")
                        return True
                    else:
                        print(f"✗ Windows系统打印命令失败，使用默认程序打开")
                        # 备用方案：使用系统默认程序打开
                        os.startfile(str(file_path), 'print')
                        print(f"✓ 已使用默认程序打开图片: {file_path.name} (请手动打印)")
                        return True
                        
                except Exception as e:
                    print(f"✗ Windows图片打印失败: {e}")
                    # 最后的备用方案：直接打开文件
                    try:
                        os.startfile(str(file_path))
                        print(f"✓ 已打开图片文件: {file_path.name} (请手动打印)")
                        return True
                    except Exception as e2:
                        print(f"✗ 打开图片文件失败: {e2}")
                        return False
            else:
                # 非Windows系统的处理
                print(f"✗ 非Windows系统，图片打印功能需要手动实现")
                return False
                
        except Exception as e:
            print(f"✗ Windows系统图片打印异常: {e}")
            return False
    
    def get_file_info(self, file_path: Path) -> Dict[str, Any]:
        """
        获取图片文件信息
        
        Args:
            file_path: 文件路径
            
        Returns:
            文件信息字典
        """
        info = {
            'file_path': str(file_path),
            'file_name': file_path.name,
            'file_size': file_path.stat().st_size,
            'pages': self.count_pages(file_path),
            'format': file_path.suffix.upper().lstrip('.'),
            'dimensions': None,
            'color_mode': None,
            'handler': self.get_handler_name()
        }
        
        # 如果有PIL，获取更详细的信息
        if PIL_AVAILABLE and self.can_handle_file(file_path):
            try:
                with Image.open(file_path) as img:
                    info['dimensions'] = f"{img.width} x {img.height}"
                    info['color_mode'] = img.mode
                    
                    # 获取图片格式信息
                    if hasattr(img, 'format'):
                        info['format'] = img.format
                        
            except Exception as e:
                print(f"获取图片详细信息失败 {file_path}: {e}")
        
        return info
    
    def validate_print_settings(self, print_settings: Any) -> tuple[bool, List[str]]:
        """
        验证图片打印设置
        
        Args:
            print_settings: 打印设置
            
        Returns:
            (是否有效, 错误信息列表)
        """
        errors = []
        
        # 检查打印机名称
        if not hasattr(print_settings, 'printer_name') or not print_settings.printer_name:
            errors.append("未指定打印机")
        
        # 图片打印建议设置
        if hasattr(print_settings, 'color_mode') and print_settings.color_mode.value == 'grayscale':
            print("💡 提示: 建议图片使用彩色打印以获得最佳效果")
        
        return len(errors) == 0, errors 