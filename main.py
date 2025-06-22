"""
办公文档批量打印器 - 主程序入口
支持批量打印Word、PowerPoint、PDF文档的桌面应用程序
"""
import sys
import os
from pathlib import Path

# 添加源代码路径到系统路径
src_path = Path(__file__).parent / "src"
sys.path.insert(0, str(src_path))

try:
    # 导入主窗口
    from src.gui.main_window import MainWindow
    
    def main():
        """主函数"""
        print("=" * 50)
        print("        办公文档批量打印器 v1.0")
        print("    支持 Word、PowerPoint、PDF 文档")
        print("=" * 50)
        
        try:
            # 检查系统要求
            check_system_requirements()
            
            # 创建并运行主窗口
            app = MainWindow()
            app.run()
            
        except Exception as e:
            print(f"应用程序运行失败: {e}")
            import tkinter.messagebox as messagebox
            messagebox.showerror("错误", f"应用程序运行失败:\n{e}")
            sys.exit(1)
    
    def check_system_requirements():
        """检查系统要求"""
        import platform
        
        # 检查操作系统
        if platform.system() != "Windows":
            raise RuntimeError("此应用程序仅支持Windows操作系统")
        
        # 检查Python版本
        if sys.version_info < (3, 8):
            raise RuntimeError("需要Python 3.8或更高版本")
        
        # 检查必要的依赖
        required_modules = [
            'tkinter',
            'pathlib',
            'win32print',
            'win32api'
        ]
        
        missing_modules = []
        for module in required_modules:
            try:
                __import__(module)
            except ImportError:
                missing_modules.append(module)
        
        if missing_modules:
            raise RuntimeError(
                f"缺少必要的依赖模块: {', '.join(missing_modules)}\n"
                "请运行: pip install -r requirements.txt"
            )
        
        print("✓ 系统要求检查通过")

    if __name__ == "__main__":
        main()

except ImportError as e:
    print(f"导入模块失败: {e}")
    print("请确保已安装所有依赖: pip install -r requirements.txt")
    sys.exit(1) 