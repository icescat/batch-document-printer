"""
文件导入功能处理器
处理所有与文件导入相关的功能：拖拽、文件选择、文件夹选择、文件过滤等
"""
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from typing import List, Set, Callable, Optional
import re

# 拖拽支持
try:
    from tkinterdnd2 import DND_FILES
    DRAG_DROP_AVAILABLE = True
except ImportError:
    DRAG_DROP_AVAILABLE = False


class FileImportHandler:
    """文件导入功能处理器"""
    
    def __init__(self, document_manager, get_enabled_file_types_callback: Callable, refresh_callback: Optional[Callable] = None):
        """
        初始化文件导入处理器
        
        Args:
            document_manager: 文档管理器实例
            get_enabled_file_types_callback: 获取启用文件类型的回调函数
            refresh_callback: 界面刷新回调函数
        """
        self.document_manager = document_manager
        self.get_enabled_file_types = get_enabled_file_types_callback
        self.refresh_callback = refresh_callback
        
    def setup_drag_drop(self, target_widget):
        """设置拖拽功能"""
        if not DRAG_DROP_AVAILABLE:
            print("拖拽功能不可用：tkinterdnd2 未安装")
            return False
            
        try:
            # 为目标控件注册拖拽
            target_widget.drop_target_register(DND_FILES)  # type: ignore
            target_widget.dnd_bind('<<Drop>>', self._on_drop_files)  # type: ignore
            
            print("✓ 拖拽功能已启用")
            return True
        except Exception as e:
            print(f"✗ 拖拽功能初始化失败: {e}")
            return False
    
    def _on_drop_files(self, event):
        """处理拖拽文件事件"""
        try:
            # 获取拖拽的文件路径
            raw_data = event.data.strip()
            print(f"拖拽数据: {raw_data}")
            print(f"拖拽数据长度: {len(raw_data)} 字符")
            
            # 解析文件路径
            files = self.parse_drag_data(raw_data)
            
            if not files:
                print("解析结果为空，显示警告")
                messagebox.showwarning("拖拽导入", "无法识别拖拽的文件路径，请确保文件或文件夹存在")
                return "copy"
            
            print(f"解析到 {len(files)} 个路径: {files}")
            
            # 处理拖拽的路径
            added_count = self.process_dropped_paths(files)
            
            if added_count > 0:
                print(f"拖拽导入成功: {added_count} 个文档")
                # 去掉成功提示窗口
                # 调用界面刷新回调
                if self.refresh_callback:
                    self.refresh_callback()
            else:
                print("拖拽导入: 未找到支持的文档格式、文件类型被过滤或文件已存在")
                # 去掉提示窗口，只保留控制台输出
                
            # 返回 "copy" 表示拖拽操作成功
            return "copy"
                
        except Exception as e:
            error_msg = f"处理拖拽文件时出错：{str(e)}"
            print(error_msg)
            import traceback
            traceback.print_exc()
            messagebox.showerror("拖拽导入失败", error_msg)
            return "copy"
    
    def parse_drag_data(self, raw_data: str) -> List[str]:
        """
        统一的拖拽数据解析器
        支持多种拖拽格式，优先使用最可靠的解析方法
        """
        # 去除整体外层大括号（如果有）
        if raw_data.startswith('{') and raw_data.endswith('}') and raw_data.count('{') == 1:
            raw_data = raw_data[1:-1]
        
        files = []
        
        # 检查单个路径
        if Path(raw_data).exists():
            return [raw_data]
        
        # 方法1: 换行符分割 (最常见的标准格式)
        if '\n' in raw_data:
            print("   使用换行符分割")
            for line in raw_data.split('\n'):
                path = line.strip('\r\n "\'{}')
                if path and Path(path).exists():
                    files.append(path)
        
        # 方法2: null字符分割 (某些系统使用)
        elif '\0' in raw_data:
            print("   使用null字符分割")
            for part in raw_data.split('\0'):
                path = part.strip('\0 "\'{}')
                if path and Path(path).exists():
                    files.append(path)
        
        # 方法3: 大括号混合格式 (Windows含空格文件的特殊格式)
        elif '{' in raw_data:
            print("   使用大括号混合格式")
            
            # 提取大括号包围的路径
            braced_pattern = r'\{([^}]+)\}'
            braced_files = re.findall(braced_pattern, raw_data)
            for path in braced_files:
                path = path.strip()
                if path and Path(path).exists():
                    files.append(path)
            
            # 提取不含大括号的路径
            remaining = re.sub(braced_pattern, ' ', raw_data)
            remaining = re.sub(r'\s+', ' ', remaining).strip()
            if remaining:
                for path in remaining.split():
                    path = path.strip(' "\'')
                    if path and Path(path).exists():
                        files.append(path)
        
        # 方法4: 引号分割 (带引号的路径列表)
        elif '"' in raw_data or "'" in raw_data:
            print("   使用引号分割")
            quoted_paths = re.findall(r'["\']([^"\']+)["\']', raw_data)
            for path in quoted_paths:
                if path and Path(path).exists():
                    files.append(path)
        
        # 方法5: 智能重建 (兜底方案，处理被错误分割的路径)
        if not files:
            print("   使用智能重建")
            files = self._smart_rebuild_paths(raw_data.split())
        
        return files
    
    def _smart_rebuild_paths(self, parts: List[str]) -> List[str]:
        """智能重建被空格分割的文件路径"""
        files = []
        i = 0
        
        while i < len(parts):
            current = parts[i].strip('"\'{}')
            
            # 检查当前片段是否为有效路径
            if Path(current).exists():
                files.append(current)
                i += 1
                continue
            
            # 尝试与后续片段组合重建路径
            rebuilt = current
            j = i + 1
            while j < len(parts):
                rebuilt += " " + parts[j].strip('"\'{}')
                if Path(rebuilt).exists():
                    files.append(rebuilt)
                    i = j + 1
                    break
                j += 1
            else:
                i += 1  # 无法重建，跳过当前片段
        
        return files
    
    def process_dropped_paths(self, paths: List[str]) -> int:
        """处理拖拽的路径列表，应用过滤器并添加到文档管理器"""
        # 获取文件类型过滤设置
        enabled_types = self.get_enabled_file_types()
        allowed_extensions = self.get_allowed_extensions(enabled_types)
        
        file_paths = []
        folder_paths = []
        
        # 分类和过滤路径
        for path_str in paths:
            path = Path(path_str)
            
            if path.is_file():
                if path.suffix.lower() in allowed_extensions:
                    file_paths.append(path)
                    print(f"   ✅ {path.name}")
                else:
                    print(f"   ❌ {path.name} (文件类型未启用)")
            elif path.is_dir():
                folder_paths.append(path)
                print(f"   📁 {path.name}")
        
        # 添加文件和文件夹
        added_count = 0
        
        if file_paths:
            try:
                added_docs = self.document_manager.add_files(file_paths)
                added_count += len(added_docs)
            except Exception as e:
                print(f"添加文件时出错: {e}")
        
        if folder_paths:
            for folder in folder_paths:
                try:
                    added_docs = self.document_manager.add_folder(folder, True, enabled_types)
                    added_count += len(added_docs)
                except Exception as e:
                    print(f"处理文件夹 {folder.name} 时出错: {e}")
        
        return added_count
    
    def add_files_dialog(self) -> int:
        """显示文件选择对话框"""
        # 根据当前启用的文件类型动态生成文件类型过滤器
        enabled_types = self.get_enabled_file_types()
        allowed_extensions = self.get_allowed_extensions(enabled_types)
        
        # 构建文件类型过滤器
        file_types = []
        supported_patterns = []
        
        if enabled_types.get('word', False):
            file_types.append(("Word文档", "*.doc;*.docx;*.wps"))
            supported_patterns.extend(["*.doc", "*.docx", "*.wps"])
        if enabled_types.get('ppt', False):
            file_types.append(("PowerPoint", "*.ppt;*.pptx;*.dps"))
            supported_patterns.extend(["*.ppt", "*.pptx", "*.dps"])
        if enabled_types.get('excel', False):
            file_types.append(("Excel表格", "*.xls;*.xlsx;*.et"))
            supported_patterns.extend(["*.xls", "*.xlsx", "*.et"])
        if enabled_types.get('pdf', False):
            file_types.append(("PDF文件", "*.pdf"))
            supported_patterns.append("*.pdf")
        if enabled_types.get('image', False):
            file_types.append(("图片文件", "*.jpg;*.jpeg;*.png;*.bmp;*.tiff;*.tif;*.webp"))
            supported_patterns.extend(["*.jpg", "*.jpeg", "*.png", "*.bmp", "*.tiff", "*.tif", "*.webp"])
        if enabled_types.get('text', False):
            file_types.append(("文本文件", "*.txt"))
            supported_patterns.append("*.txt")
        
        # 添加支持的文档总览
        if supported_patterns:
            file_types.insert(0, ("支持的文档", ";".join(supported_patterns)))
        
        # 添加所有文件选项
        file_types.append(("所有文件", "*.*"))
        
        if not file_types or len(file_types) <= 1:  # 只有"所有文件"选项
            messagebox.showwarning("提示", "请先在文件类型过滤器中启用至少一种文档类型")
            return 0
        
        files = filedialog.askopenfilenames(
            title="选择要打印的文档",
            filetypes=file_types
        )
        
        if files:
            # 应用文件类型过滤
            file_paths = [Path(f) for f in files]
            filtered_paths = []
            filtered_out_count = 0
            
            for file_path in file_paths:
                if file_path.suffix.lower() in allowed_extensions:
                    filtered_paths.append(file_path)
                else:
                    filtered_out_count += 1
                    print(f"   ❌ {file_path.name} (文件类型未启用)")
            
            if filtered_out_count > 0:
                messagebox.showwarning("文件过滤", 
                    f"有 {filtered_out_count} 个文件因文件类型过滤被跳过。\n"
                    f"请检查文件类型过滤器设置。")
            
            if filtered_paths:
                added_docs = self.document_manager.add_files(filtered_paths)
                
                if added_docs:
                    # 去掉成功提示窗口
                    print(f"成功添加 {len(added_docs)} 个文档")
                    # 调用界面刷新回调
                    if self.refresh_callback:
                        self.refresh_callback()
                    return len(added_docs)
        
        return 0
    
    def add_folder_dialog(self) -> int:
        """显示文件夹选择对话框"""
        folder = filedialog.askdirectory(title="选择包含文档的文件夹")
        
        if folder:
            folder_path = Path(folder)
            
            # 询问是否递归搜索
            recursive = messagebox.askyesno(
                "搜索选项", 
                "是否搜索子文件夹中的文档？"
            )
            
            # 获取当前的文件类型过滤设置
            enabled_file_types = self.get_enabled_file_types()
            
            added_docs = self.document_manager.add_folder(folder_path, recursive, enabled_file_types)
            
            if added_docs:
                # 去掉成功提示窗口
                print(f"从文件夹中成功添加 {len(added_docs)} 个文档")
                # 调用界面刷新回调
                if self.refresh_callback:
                    self.refresh_callback()
                return len(added_docs)
            else:
                # 去掉警告窗口，只保留控制台输出
                print("在指定文件夹中未找到支持的文档")
        
        return 0
    
    def get_allowed_extensions(self, enabled_types: dict) -> Set[str]:
        """根据启用的文件类型获取允许的扩展名集合"""
        extensions = set()
        
        if enabled_types.get('word', False):
            extensions.update(['.doc', '.docx', '.wps'])
        if enabled_types.get('ppt', False):
            extensions.update(['.ppt', '.pptx', '.dps'])
        if enabled_types.get('excel', False):
            extensions.update(['.xls', '.xlsx', '.et'])
        if enabled_types.get('pdf', False):
            extensions.add('.pdf')
        if enabled_types.get('image', False):
            extensions.update(['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'])
        if enabled_types.get('text', False):
            extensions.add('.txt')
        
        return extensions 