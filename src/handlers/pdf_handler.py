"""
PDF文档处理器 - 基于SumatraPDF绿色版的可靠打印方案
集成SumatraPDF绿色版，提供高兼容性的PDF打印支持
支持系统关联和SumatraPDF双重方案
"""
import subprocess
import time
from pathlib import Path
from typing import Set, Optional
from ..core.models import FileType, PrintSettings
from .base_handler import BaseDocumentHandler


class PDFDocumentHandler(BaseDocumentHandler):
    """PDF文档处理器 - 集成SumatraPDF绿色版"""
    
    def __init__(self):
        """初始化PDF处理器"""
        self._timeout_seconds = 30
        self._sumatra_path = None
        self._setup_sumatra_pdf()
    
    def _setup_sumatra_pdf(self):
        """设置SumatraPDF环境"""
        # SumatraPDF存放路径
        self._sumatra_dir = Path("external/SumatraPDF")
        self._sumatra_exe = self._sumatra_dir / "SumatraPDF.exe"
        
        # 检查是否已存在
        if self._sumatra_exe.exists():
            self._sumatra_path = str(self._sumatra_exe)
            print(f"✅ SumatraPDF已就绪: {self._sumatra_exe}")
        else:
            self._sumatra_path = None
            print(f"❌ SumatraPDF未找到: {self._sumatra_exe}")
    
    def ensure_sumatra_available(self) -> bool:
        """确保SumatraPDF可用"""
        return self._sumatra_path and Path(self._sumatra_path).exists()
    

    
    def get_supported_file_types(self) -> Set[FileType]:
        """获取支持的文件类型"""
        return {FileType.PDF}
    
    def get_supported_extensions(self) -> Set[str]:
        """获取支持的文件扩展名"""
        return {'.pdf'}
    
    def can_handle_file(self, file_path: Path) -> bool:
        """检查是否能处理指定文件"""
        if not self.validate_file_exists(file_path):
            return False
        
        extension = file_path.suffix.lower()
        return extension in self.get_supported_extensions()
    
    def print_document(self, file_path: Path, settings: PrintSettings) -> bool:
        """
        打印PDF文档 - 使用SumatraPDF可靠打印
        
        Args:
            file_path: PDF文件路径
            settings: 打印设置
            
        Returns:
            是否打印成功
        """
        print(f"🖨️ 开始打印PDF文件: {file_path.name}")
        
        # 获取打印机名称
        printer_name = self._get_printer_name(settings)
        if not printer_name:
            print("❌ 无法获取打印机")
            return False
        
        print(f"📋 使用打印机: {printer_name}")
        
        # 使用SumatraPDF打印
        if self.ensure_sumatra_available():
            return self._print_with_sumatra_pdf(file_path, printer_name, settings)
        else:
            print("❌ SumatraPDF不可用，请检查安装")
            return False
    
    def _print_with_sumatra_pdf(self, file_path: Path, printer_name: str, settings: PrintSettings) -> bool:
        """使用SumatraPDF打印 - 支持完整的打印设置"""
        try:
            if not self._sumatra_path:
                print("❌ SumatraPDF路径未设置")
                return False
            
            print(f"🖨️ 使用SumatraPDF打印: {file_path.name}")
            
            # 构建SumatraPDF命令行参数
            cmd = [
                self._sumatra_path,
                "-print-to", printer_name,
                "-silent"  # 静默模式
            ]
            
            # 构建打印设置
            print_settings = []
            
            # 双面打印设置
            print_settings.append(settings.duplex_mode_str)
            
            # 纸张方向设置
            print_settings.append(settings.orientation_str)
            
            # 如果有打印设置，添加到命令中
            if print_settings:
                cmd.extend(["-print-settings", ",".join(print_settings)])
            
            # 添加文件路径
            cmd.append(str(file_path))
            
            # 显示打印设置信息
            print(f"📄 双面设置: {settings.duplex_mode_str}")
            print(f"📐 纸张方向: {settings.orientation_str}")
            print(f"🔢 打印份数: {settings.copies}")
            print(f"🎨 色彩模式: {'彩色' if settings.color else '黑白'}")
            
            print(f"🔧 SumatraPDF命令: {' '.join(cmd)}")
            
            # 执行打印命令
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=self._timeout_seconds
            )
            
            if result.returncode == 0:
                print("✅ SumatraPDF打印命令执行成功")
                # 等待打印作业完成
                time.sleep(2)
                return True
            else:
                print(f"❌ SumatraPDF打印失败:")
                if result.stdout:
                    print(f"   输出: {result.stdout}")
                if result.stderr:
                    print(f"   错误: {result.stderr}")
                return False
                
        except subprocess.TimeoutExpired:
            print(f"❌ SumatraPDF打印超时（{self._timeout_seconds}秒）")
            return False
        except Exception as e:
            print(f"❌ SumatraPDF打印异常: {e}")
            return False
    

    
    def _get_printer_name(self, settings: PrintSettings) -> Optional[str]:
        """获取打印机名称"""
        if settings.printer_name:
            return settings.printer_name
            
        try:
            import win32print
            return win32print.GetDefaultPrinter()
        except Exception as e:
            print(f"⚠️ 获取默认打印机失败: {e}")
            
            # 备用方案：使用PowerShell获取默认打印机
            try:
                result = subprocess.run(
                    ['powershell', '-Command', 'Get-WmiObject -Class Win32_Printer | Where-Object {$_.Default -eq $true} | Select-Object -ExpandProperty Name'],
                    capture_output=True,
                    text=True,
                    timeout=10
                )
                if result.returncode == 0 and result.stdout.strip():
                    return result.stdout.strip()
            except Exception:
                pass
                
            return None
    
    def count_pages(self, file_path: Path) -> int:
        """
        统计PDF页数 - 使用PyPDF2高效解析
        
        Args:
            file_path: PDF文件路径
            
        Returns:
            页数，失败返回-1
        """
        return self._count_pages_with_pypdf2(file_path)
    
    def _count_pages_with_pypdf2(self, file_path: Path) -> int:
        """使用PyPDF2统计页数"""
        try:
            from PyPDF2 import PdfReader
            
            with open(file_path, 'rb') as file:
                pdf_reader = PdfReader(file)
                page_count = len(pdf_reader.pages)
                print(f"📄 PDF页数统计成功: {page_count}")
                return page_count
                
        except ImportError:
            print("⚠️ 缺少PyPDF2库，无法统计PDF页数")
            return -1
        except Exception as e:
            print(f"⚠️ PDF页数统计失败: {e}")
            return -1 