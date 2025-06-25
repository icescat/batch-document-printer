"""
PDF文档处理器
负责PDF文件的打印和页数统计
"""
import os
import subprocess
import time
from pathlib import Path
from typing import Set, Optional
from ..core.models import FileType, PrintSettings
from .base_handler import BaseDocumentHandler


class PDFDocumentHandler(BaseDocumentHandler):
    """PDF文档处理器"""
    
    def __init__(self):
        """初始化PDF处理器"""
        self._timeout_seconds = 30
    
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
        打印PDF文档
        
        Args:
            file_path: PDF文件路径
            settings: 打印设置
            
        Returns:
            是否打印成功
        """
        try:
            print(f"开始打印PDF文件: {file_path}")
            
            # 验证系统级双面打印设置
            if settings.duplex and settings.printer_name:
                try:
                    import win32print
                    printer_handle = win32print.OpenPrinter(settings.printer_name)
                    try:
                        printer_info = win32print.GetPrinter(printer_handle, 2)
                        devmode = printer_info.get('pDevMode')
                        if devmode and hasattr(devmode, 'Duplex'):
                            print(f"🔍 PDF打印前验证: 打印机双面设置为 {devmode.Duplex}")
                    finally:
                        win32print.ClosePrinter(printer_handle)
                except Exception as duplex_check:
                    print(f"⚠️ PDF双面打印验证失败: {duplex_check}")
            
            # 优先使用系统API方式，利用系统级打印机配置
            success = self._print_with_system_api(file_path, settings)
            
            if success:
                print(f"✅ PDF打印成功（系统API）: {file_path.name}")
                return True
            else:
                # 如果系统API失败，回退到Edge方式
                print(f"⚠️ 系统API打印失败，尝试Edge方式")
                success = self._print_with_edge(file_path, settings.printer_name)
                if success:
                    print(f"✅ PDF打印成功（Edge）: {file_path.name}")
                    return True
                else:
                    print(f"❌ PDF打印失败: {file_path.name}")
                    return False
                
        except Exception as e:
            print(f"PDF打印过程中发生错误: {e}")
            return False
    
    def count_pages(self, file_path: Path) -> int:
        """
        统计PDF文档页数
        
        Args:
            file_path: PDF文件路径
            
        Returns:
            页数
            
        Raises:
            Exception: 统计失败时抛出异常
        """
        try:
            # 首先尝试使用PyPDF2
            import PyPDF2
            
            with open(file_path, 'rb') as file:
                try:
                    reader = PyPDF2.PdfReader(file)
                    
                    # 检查是否加密
                    if reader.is_encrypted:
                        raise Exception("PDF文件已加密，需要密码")
                    
                    page_count = len(reader.pages)
                    
                    # 验证结果的合理性
                    if page_count < 0:
                        raise Exception("获取到无效的页数")
                    
                    return page_count
                    
                except Exception as pdf_error:
                    error_str = str(pdf_error).lower()
                    if "encrypted" in error_str or "password" in error_str:
                        raise Exception("文件被加密")
                    else:
                        raise Exception("文件已损坏")
                        
        except ImportError:
            raise Exception("需要安装PyPDF2库来处理PDF文件")
        except Exception as e:
            # 简化的错误处理
            error_str = str(e).lower()
            if "文件被加密" in str(e) or "文件已损坏" in str(e):
                # 重新抛出我们的自定义错误
                raise e
            elif "encrypted" in error_str or "password" in error_str or "permission" in error_str or "access" in error_str:
                raise Exception("文件被加密")
            else:
                raise Exception("文件已损坏")
    
    def _print_with_edge(self, file_path: Path, printer_name: Optional[str] = None) -> bool:
        """
        使用Microsoft Edge打印PDF文件
        
        Args:
            file_path: PDF文件路径
            printer_name: 打印机名称
            
        Returns:
            是否打印成功
        """
        try:
            # 构建Edge启动命令
            edge_paths = [
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
            ]
            
            edge_path = None
            for path in edge_paths:
                if os.path.exists(path):
                    edge_path = path
                    break
            
            if not edge_path:
                print("未找到Microsoft Edge")
                return False
            
            # 构建打印命令
            cmd = [
                edge_path,
                "--headless",
                "--disable-gpu",
                "--run-all-compositor-stages-before-draw",
                "--print-to-pdf-no-header",
                f"--print-to-pdf={file_path}",
                str(file_path)
            ]
            
            if printer_name:
                cmd.extend([f"--printer-name={printer_name}"])
            
            # 执行打印命令
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=30,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            if result.returncode == 0:
                print(f"Edge打印命令执行成功")
                return True
            else:
                print(f"Edge打印命令失败: {result.stderr}")
                return False
                
        except subprocess.TimeoutExpired:
            print("Edge打印命令超时")
            return False
        except Exception as e:
            print(f"使用Edge打印时出错: {e}")
            return False
    
    def _print_with_system_api(self, file_path: Path, settings: PrintSettings) -> bool:
        """
        使用Windows API打印PDF文件（支持双面打印）
        
        Args:
            file_path: PDF文件路径
            settings: 打印设置
            
        Returns:
            是否打印成功
        """
        try:
            import win32api
            import win32print
            import win32con
            import subprocess
            
            # 获取打印机名称
            printer_name = settings.printer_name or win32print.GetDefaultPrinter()
            print(f"🖨️ PDF使用打印机: {printer_name}")
            
            # 方法1: 尝试使用Adobe Reader或其他PDF阅读器的命令行打印
            success = self._try_pdf_reader_print(file_path, printer_name, settings)
            if success:
                return True
            
            # 方法2: 使用Windows默认的打印命令
            try:
                print("📤 使用Windows默认打印命令...")
                
                # 使用系统默认程序打印
                result = subprocess.run([
                    "rundll32.exe", 
                    "mshtml.dll,PrintHTML", 
                    str(file_path)
                ], capture_output=True, text=True, timeout=10)
                
                if result.returncode == 0:
                    print("✅ Windows默认打印命令执行成功")
                    time.sleep(3)
                    return True
                else:
                    print(f"⚠️ Windows默认打印命令失败: {result.stderr}")
                    
            except Exception as default_error:
                print(f"⚠️ Windows默认打印失败: {default_error}")
            
            # 方法3: 使用ShellExecute打印
            try:
                print("📤 使用ShellExecute打印...")
                win32api.ShellExecute(
                    0,
                    "print",
                    str(file_path),
                    f'/d:"{printer_name}"',
                    "",
                    0  # SW_HIDE
                )
                
                print("✅ ShellExecute打印命令已发送")
                time.sleep(3)
                return True
                
            except Exception as shell_error:
                print(f"❌ ShellExecute打印失败: {shell_error}")
                return False
                
        except Exception as e:
            print(f"❌ PDF系统API打印失败: {e}")
            return False
    
    def _try_pdf_reader_print(self, file_path: Path, printer_name: str, settings: PrintSettings) -> bool:
        """尝试使用PDF阅读器的命令行打印功能"""
        try:
            import subprocess
            import os
            
            # 常见PDF阅读器的可能路径
            pdf_readers = [
                r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
                r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
                r"C:\Program Files\Foxit Software\Foxit Reader\FoxitReader.exe"
            ]
            
            for reader_path in pdf_readers:
                if os.path.exists(reader_path):
                    try:
                        print(f"📖 尝试使用PDF阅读器打印: {os.path.basename(reader_path)}")
                        
                        if "Acrobat" in reader_path or "AcroRd32" in reader_path:
                            # Adobe Reader/Acrobat命令行参数
                            cmd = [
                                reader_path,
                                "/t",  # 打印
                                str(file_path),
                                printer_name
                            ]
                        elif "Foxit" in reader_path:
                            # Foxit Reader命令行参数
                            cmd = [
                                reader_path,
                                "/t",
                                str(file_path),
                                printer_name
                            ]
                        else:
                            continue
                        
                        result = subprocess.run(
                            cmd,
                            capture_output=True,
                            text=True,
                            timeout=15,
                            creationflags=subprocess.CREATE_NO_WINDOW
                        )
                        
                        if result.returncode == 0:
                            print(f"✅ PDF阅读器打印成功")
                            time.sleep(2)
                            return True
                        else:
                            print(f"⚠️ PDF阅读器打印失败: {result.stderr}")
                            
                    except Exception as reader_error:
                        print(f"⚠️ PDF阅读器打印异常: {reader_error}")
                        continue
            
            return False
            
        except Exception as e:
            print(f"⚠️ PDF阅读器打印检查失败: {e}")
            return False
    
    def _print_with_default_program(self, file_path: Path) -> bool:
        """
        使用系统默认程序打印PDF文件
        
        Args:
            file_path: PDF文件路径
            
        Returns:
            是否打印成功
        """
        try:
            # 使用Windows的默认打印命令
            import win32api
            import win32print
            
            # 获取默认打印机
            default_printer = win32print.GetDefaultPrinter()
            
            # 使用ShellExecute打印
            win32api.ShellExecute(
                0,
                "print",
                str(file_path),
                f'/d:"{default_printer}"',
                "",
                0
            )
            
            # 等待一段时间让打印任务开始
            time.sleep(2)
            
            return True
            
        except Exception as e:
            print(f"使用默认程序打印时出错: {e}")
            return False 