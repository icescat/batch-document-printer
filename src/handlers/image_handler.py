"""
图片文档处理器
处理各种图片格式的打印和页数统计功能
"""
import os
import sys
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
    """图片文档处理器"""
    
    def __init__(self):
        """初始化图片处理器"""
        super().__init__()
        
        # 支持的图片格式
        self._supported_extensions = {
            '.jpg', '.jpeg', '.png', '.bmp', 
            '.tiff', '.tif', '.webp'  # 移除了 .gif, .ico 和 .pcx
        }
        self._supported_file_types = {FileType.IMAGE}
        
        # 如果PIL不可用，减少支持的格式
        if not PIL_AVAILABLE:
            print("警告: PIL/Pillow 未安装，图片支持功能受限")
            # 只保留Windows系统原生支持的格式
            self._supported_extensions = {'.jpg', '.jpeg', '.png', '.bmp'}
    
    def get_handler_name(self) -> str:
        """获取处理器名称"""
        return "图片文档处理器"
    
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
            if file_path.stat().st_size < 10:  # 小于10字节的文件很可能不是有效图片
                return False
            
            # 如果有PIL，进行更严格的验证
            if PIL_AVAILABLE:
                try:
                    with Image.open(file_path) as img:
                        img.verify()  # 验证图片文件完整性
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
                print(f"警告: 缺少PIL库，无法准确统计TIFF页数，假设为1页: {file_path.name}")
                return 1
                
        except Exception as e:
            print(f"统计TIFF页数失败 {file_path}: {e}")
            return 1  # 失败时返回1页
    
    def print_document(self, file_path: Path, settings: Any) -> bool:
        """
        打印图片文档
        
        Args:
            file_path: 文件路径
            settings: 打印设置
            
        Returns:
            打印是否成功
        """
        try:
            if not self.can_handle_file(file_path):
                raise ValueError(f"无法处理的图片文件: {file_path}")
            
            # 检查是否为多页TIFF
            extension = file_path.suffix.lower()
            if extension in ['.tiff', '.tif']:
                page_count = self._count_tiff_pages(file_path)
                if page_count > 1:
                    print(f"⚠️ 检测到多页TIFF文件: {file_path.name} ({page_count}页)")
                    print("   注意: 使用默认查看器可能只能打印第一页")
                    print("   建议: 使用专业TIFF查看器打印所有页面")
            
            print(f"开始打印图片: {file_path.name}")
            
            # Windows系统下使用默认图片查看器打印
            # 这里使用系统的打印命令
            if os.name == 'nt':  # Windows系统
                try:
                    # 使用Windows的默认打印命令
                    # 注意：这会打开默认的图片查看器，用户需要手动点击打印
                    import subprocess
                    result = subprocess.run([
                        'rundll32.exe',
                        'shimgvw.dll,ImageView_PrintTo',
                        str(file_path),
                        settings.printer_name
                    ], capture_output=True, text=True, timeout=30)
                    
                    if result.returncode == 0:
                        print(f"✓ 图片打印命令执行成功: {file_path.name}")
                        return True
                    else:
                        print(f"✗ 图片打印命令执行失败: {file_path.name}, 错误: {result.stderr}")
                        
                        # 备用方案：使用系统默认程序打开（需要用户手动打印）
                        os.startfile(str(file_path), 'print')
                        print(f"✓ 已使用默认程序打开图片: {file_path.name} (请手动打印)")
                        return True
                        
                except Exception as e:
                    print(f"✗ Windows图片打印失败: {e}")
                    # 最后的备用方案
                    try:
                        os.startfile(str(file_path))
                        print(f"✓ 已打开图片文件: {file_path.name} (请手动打印)")
                        return True
                    except Exception as e2:
                        print(f"✗ 打开图片文件失败: {e2}")
                        return False
            else:
                # 非Windows系统的处理（Linux/Mac）
                print(f"非Windows系统，图片打印功能需要手动实现")
                return False
                
        except Exception as e:
            print(f"✗ 打印图片文档失败 {file_path.name}: {e}")
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
            'pages': 1,  # 图片固定1页
            'format': file_path.suffix.upper().lstrip('.'),
            'dimensions': None,
            'color_mode': None
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
            print("提示: 建议图片使用彩色打印以获得最佳效果")
        
        return len(errors) == 0, errors 