#!/usr/bin/env python3
"""
办公文档批量打印器 - 自动化打包脚本
用于生成exe可执行文件
"""
import os
import sys
import subprocess
import shutil
from pathlib import Path
import time


def print_header(title):
    """打印标题"""
    print("\n" + "="*60)
    print(f"  {title}")
    print("="*60)


def check_requirements():
    """检查打包要求"""
    print_header("检查打包要求")
    
    # 检查Python版本
    python_version = sys.version_info
    print(f"✓ Python版本: {python_version.major}.{python_version.minor}.{python_version.micro}")
    
    if python_version < (3, 8):
        print("❌ 需要Python 3.8或更高版本")
        return False
    
    # 检查PyInstaller
    try:
        import PyInstaller
        print(f"✓ PyInstaller版本: {PyInstaller.__version__}")
    except ImportError:
        print("❌ PyInstaller未安装，请运行: pip install PyInstaller")
        return False
    
    # 检查必要文件
    required_files = [
        "main.py",
        "办公文档批量打印器v5.0.spec",
        "requirements.txt",
        "resources/app_icon.ico",
        "external/SumatraPDF/SumatraPDF.exe"
    ]
    
    for file_path in required_files:
        if Path(file_path).exists():
            print(f"✓ {file_path}")
        else:
            print(f"❌ 缺少文件: {file_path}")
            return False
    
    return True


def clean_build_dirs():
    """清理构建目录"""
    print_header("清理构建目录")
    
    dirs_to_clean = ["build", "dist", "__pycache__"]
    
    for dir_name in dirs_to_clean:
        if Path(dir_name).exists():
            print(f"删除目录: {dir_name}")
            shutil.rmtree(dir_name)
        else:
            print(f"目录不存在: {dir_name}")


def run_pyinstaller():
    """运行PyInstaller"""
    print_header("开始打包")
    
    spec_file = "办公文档批量打印器v5.0.spec"
    
    # 构建命令
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--clean",  # 清理缓存
        "--noconfirm",  # 不询问覆盖
        spec_file
    ]
    
    print(f"执行命令: {' '.join(cmd)}")
    print("\n开始打包，请耐心等待...")
    
    start_time = time.time()
    
    try:
        # 运行PyInstaller
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        
        end_time = time.time()
        duration = end_time - start_time
        
        print(f"\n打包完成，耗时: {duration:.1f}秒")
        
        if result.returncode == 0:
            print("✓ 打包成功!")
            return True
        else:
            print("❌ 打包失败!")
            print("\n错误输出:")
            print(result.stderr)
            return False
            
    except Exception as e:
        print(f"❌ 打包过程中发生错误: {e}")
        return False


def check_output():
    """检查输出文件"""
    print_header("检查输出文件")
    
    exe_path = Path("dist/办公文档批量打印器v5.0.exe")
    
    if exe_path.exists():
        file_size = exe_path.stat().st_size / (1024 * 1024)  # MB
        print(f"✓ EXE文件已生成: {exe_path}")
        print(f"✓ 文件大小: {file_size:.1f} MB")
        
        # 检查关键资源文件是否包含
        dist_dir = Path("dist")
        if dist_dir.exists():
            print(f"✓ 输出目录: {dist_dir}")
            
        return True
    else:
        print("❌ EXE文件未生成")
        return False


def create_version_info():
    """创建版本信息文件"""
    print_header("创建版本信息")
    
    version_info = '''# UTF-8
#
# For more details about fixed file info 'ffi' see:
# http://msdn.microsoft.com/en-us/library/ms646997.aspx
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=(5,0,0,0),
    prodvers=(5,0,0,0),
    mask=0x3f,
    flags=0x0,
    OS=0x40004L,
    fileType=0x1L,
    subtype=0x0L,
    date=(0, 0)
    ),
  kids=[
    StringFileInfo(
      [
      StringTable(
        u'080404B0',
        [StringStruct(u'CompanyName', u'喵言喵语'),
        StringStruct(u'FileDescription', u'办公文档批量打印器'),
        StringStruct(u'FileVersion', u'5.0.0.0'),
        StringStruct(u'InternalName', u'办公文档批量打印器'),
        StringStruct(u'LegalCopyright', u'Copyright © 2024 喵言喵语'),
        StringStruct(u'OriginalFilename', u'办公文档批量打印器v5.0.exe'),
        StringStruct(u'ProductName', u'办公文档批量打印器'),
        StringStruct(u'ProductVersion', u'5.0.0.0')])
      ]), 
    VarFileInfo([VarStruct(u'Translation', [2052, 1200])])
  ]
)'''
    
    with open("version_info.txt", "w", encoding="utf-8") as f:
        f.write(version_info)
    
    print("✓ 版本信息文件已创建: version_info.txt")


def main():
    """主函数"""
    print_header("办公文档批量打印器 - 自动化打包工具")
    
    # 检查要求
    if not check_requirements():
        print("\n❌ 检查失败，无法继续打包")
        return False
    
    # 创建版本信息
    create_version_info()
    
    # 清理构建目录
    clean_build_dirs()
    
    # 运行打包
    if not run_pyinstaller():
        print("\n❌ 打包失败")
        return False
    
    # 检查输出
    if not check_output():
        print("\n❌ 输出检查失败")
        return False
    
    print_header("打包完成")
    print("✓ 所有步骤完成!")
    print(f"✓ 可执行文件位置: dist/办公文档批量打印器v5.0.exe")
    print("\n您现在可以运行生成的exe文件进行测试。")
    
    return True


if __name__ == "__main__":
    success = main()
    if not success:
        sys.exit(1) 