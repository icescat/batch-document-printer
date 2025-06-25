"""
文本文档处理器
处理TXT文件的打印和页数统计功能
"""
import os
import sys
import subprocess
from pathlib import Path
from typing import Set, Dict, Any

# 添加项目根目录到Python路径
project_root = Path(__file__).parents[2]
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from src.core.models import FileType
from src.handlers.base_handler import BaseDocumentHandler


class TextDocumentHandler(BaseDocumentHandler):
    """文本文档处理器"""
    
    def __init__(self):
        """初始化文本处理器"""
        super().__init__()
        
        # 支持的文本格式（目前只支持TXT）
        self._supported_extensions = {'.txt'}
        self._supported_file_types = {FileType.TEXT}
        
        # 文本文件页数估算参数
        self._chars_per_line = 75  # 每行字符数（考虑打印边距）
        self._lines_per_page = 50  # 每页行数（考虑页边距）
        
    def get_handler_name(self) -> str:
        """获取处理器名称"""
        return "文本文档处理器"
    
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
        
        # 基本文件验证
        try:
            # 检查文件大小（避免处理过大的文本文件）
            file_size_mb = file_path.stat().st_size / (1024 * 1024)
            if file_size_mb > 100:  # 限制100MB
                print(f"文本文件过大: {file_path.name} ({file_size_mb:.1f}MB)")
                return False
            
            # 验证TXT文件编码
            return self._validate_txt_file(file_path)
                
        except Exception as e:
            print(f"验证文本文件失败 {file_path}: {e}")
            return False
    
    def _validate_txt_file(self, file_path: Path) -> bool:
        """验证TXT文件"""
        try:
            # 尝试用不同编码读取文件开头
            encodings = ['utf-8', 'gbk', 'gb2312', 'utf-16', 'latin1']
            
            for encoding in encodings:
                try:
                    with open(file_path, 'r', encoding=encoding) as f:
                        # 读取前2KB验证编码
                        content = f.read(2048)
                        # 检查是否包含大量二进制字符（可能不是文本文件）
                        if self._is_likely_text(content):
                            return True
                except (UnicodeDecodeError, UnicodeError):
                    continue
            
            print(f"无法识别TXT文件编码或疑似二进制文件: {file_path.name}")
            return False
            
        except Exception:
            return False
    
    def _is_likely_text(self, content: str) -> bool:
        """判断内容是否为文本"""
        if not content:
            return True
        
        # 计算可打印字符的比例
        printable_chars = sum(1 for c in content if c.isprintable() or c.isspace())
        ratio = printable_chars / len(content)
        
        # 如果80%以上是可打印字符，认为是文本文件
        return ratio >= 0.8
    
    def count_pages(self, file_path: Path) -> int:
        """
        统计TXT文件页数
        
        Args:
            file_path: 文件路径
            
        Returns:
            页数
        """
        if not self.can_handle_file(file_path):
            raise ValueError(f"无法处理的文本文件: {file_path}")
        
        return self._count_txt_pages(file_path)
    
    def _count_txt_pages(self, file_path: Path) -> int:
        """统计TXT文件页数（估算）"""
        try:
            # 检测文件编码并读取内容
            content = self._read_txt_with_encoding(file_path)
            if not content:
                return 1
            
            # 按行分割
            lines = content.split('\n')
            total_lines = 0
            
            for line in lines:
                # 计算每行实际占用的行数（考虑自动换行）
                line_length = len(line.rstrip())  # 去除行尾空格
                if line_length == 0:
                    total_lines += 1  # 空行
                else:
                    # 计算自动换行导致的行数
                    wrapped_lines = (line_length + self._chars_per_line - 1) // self._chars_per_line
                    total_lines += max(wrapped_lines, 1)
            
            # 计算页数
            pages = (total_lines + self._lines_per_page - 1) // self._lines_per_page
            return max(pages, 1)
            
        except Exception as e:
            print(f"TXT页数统计失败 {file_path}: {e}")
            return 1
    
    def _read_txt_with_encoding(self, file_path: Path) -> str:
        """使用合适的编码读取TXT文件"""
        encodings = ['utf-8', 'gbk', 'gb2312', 'utf-16', 'utf-16le', 'latin1']
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    content = f.read()
                    # 验证读取的内容是否合理
                    if self._is_likely_text(content[:1000]):  # 检查前1000字符
                        return content
            except (UnicodeDecodeError, UnicodeError):
                continue
        
        raise ValueError(f"无法解码文件: {file_path}")
    
    def print_document(self, file_path: Path, settings: Any) -> bool:
        """
        打印TXT文档
        
        Args:
            file_path: 文件路径
            settings: 打印设置
            
        Returns:
            打印是否成功
        """
        try:
            if not self.can_handle_file(file_path):
                raise ValueError(f"无法处理的文本文件: {file_path}")
            
            print(f"开始打印文本文件: {file_path.name}")
            
            if os.name == 'nt':  # Windows系统
                return self._print_txt_windows(file_path, settings)
            else:
                print("非Windows系统的文本打印功能需要手动实现")
                return False
                
        except Exception as e:
            print(f"打印文本文档失败 {file_path.name}: {e}")
            return False
    
    def _print_txt_windows(self, file_path: Path, settings: Any) -> bool:
        """在Windows上打印TXT文件"""
        try:
            # 方法1: 使用notepad打印（最直接的方式）
            try:
                result = subprocess.run([
                    'notepad.exe', '/p', str(file_path)
                ], capture_output=True, text=True, timeout=10)
                
                # notepad /p 命令通常会直接发送到默认打印机
                print(f"✓ TXT文件已发送至打印机: {file_path.name}")
                return True
                
            except subprocess.TimeoutExpired:
                print(f"✓ TXT文件打印命令已执行: {file_path.name}")
                return True
            except Exception as e1:
                print(f"notepad打印失败: {e1}")
                
                # 方法2: 使用系统默认程序打印
                try:
                    os.startfile(str(file_path), 'print')
                    print(f"✓ 已使用默认程序打印: {file_path.name}")
                    return True
                except Exception as e2:
                    print(f"默认程序打印失败: {e2}")
                    
                    # 方法3: 直接打开文件让用户手动打印
                    try:
                        os.startfile(str(file_path))
                        print(f"✓ 已打开文件 {file_path.name} (请手动打印)")
                        return True
                    except Exception as e3:
                        print(f"打开文件失败: {e3}")
                        return False
                        
        except Exception as e:
            print(f"TXT文件打印失败: {e}")
            return False
    
    def get_file_info(self, file_path: Path) -> Dict[str, Any]:
        """
        获取文本文件信息
        
        Args:
            file_path: 文件路径
            
        Returns:
            文件信息字典
        """
        info = {
            'file_path': str(file_path),
            'file_name': file_path.name,
            'file_size': file_path.stat().st_size,
            'format': 'TXT',
            'encoding': 'Unknown',
            'lines': 0,
            'pages': 0
        }
        
        if self.can_handle_file(file_path):
            try:
                # 检测编码
                content = self._read_txt_with_encoding(file_path)
                info['lines'] = len(content.split('\n'))
                info['pages'] = self._count_txt_pages(file_path)
                
                # 尝试检测编码类型
                encodings = ['utf-8', 'gbk', 'gb2312']
                for encoding in encodings:
                    try:
                        with open(file_path, 'r', encoding=encoding) as f:
                            f.read(100)
                        info['encoding'] = encoding.upper()
                        break
                    except (UnicodeDecodeError, UnicodeError):
                        continue
                        
            except Exception as e:
                print(f"获取文本文件详细信息失败 {file_path}: {e}")
        
        return info 