"""
主窗口界面
办公文档批量打印应用的主界面 (重构版本)
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import List, Optional
import threading

# 拖拽支持
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DRAG_DROP_AVAILABLE = True
except ImportError:
    print("警告: tkinterdnd2 未安装，拖拽功能将不可用")
    DRAG_DROP_AVAILABLE = False

from src.core.document_manager import DocumentManager
from src.core.settings_manager import PrinterSettingsManager
from src.core.print_controller import PrintController
from src.core.models import Document, PrintSettings, PrintStatus
from src.utils.config_utils import ConfigManager
from src.gui.print_settings_dialog import PrintSettingsDialog
from src.gui.page_count_dialog import show_page_count_dialog

# 导入功能处理器
from src.gui.components import FileImportHandler, ListOperationHandler, WindowManager, create_button_tooltip


class MainWindow:
    """主窗口类 (重构版本)"""
    
    def __init__(self):
        """初始化主窗口"""
        # 创建支持拖拽的主窗口
        if DRAG_DROP_AVAILABLE:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        
        # 初始化核心管理器
        self.document_manager = DocumentManager()
        self.printer_manager = PrinterSettingsManager()
        self.print_controller = PrintController()
        self.config_manager = ConfigManager()
        
        # 加载配置
        self.app_config = self.config_manager.load_app_config()
        self.current_print_settings = self.config_manager.load_print_settings()
        
        # 初始化功能处理器
        self._setup_handlers()
        
        # 设置打印控制器
        self.print_controller.set_print_settings(self.current_print_settings)
        self.print_controller.set_progress_callback(self._on_print_progress)
        
        # 创建界面
        self._create_widgets()
        self._setup_layout()
        self._bind_events()
        
        # 设置窗口属性
        self._setup_window()
        
        print("批量文档打印器已启动 (重构版本)")
    
    def _setup_handlers(self):
        """初始化功能处理器"""
        # 窗口管理器
        self.window_manager = WindowManager(self.root, self.config_manager)
        
        # 文件导入处理器 (需要在创建界面后初始化拖拽功能)
        self.file_import_handler = FileImportHandler(
            self.document_manager, 
            self._get_enabled_file_types,
            self._on_files_imported
        )
        
        # 列表操作处理器 (需要在创建树形控件后初始化)
        self.list_operation_handler = None  # 稍后初始化
    
    def _setup_window(self):
        """设置窗口属性"""
        # 设置窗口标题
        self.window_manager.set_window_title("办公文档批量打印器", "v5.0 by.喵言喵语")
        
        # 设置窗口最小尺寸
        self.window_manager.set_window_minimum_size(800, 500)
        
        # 恢复窗口几何属性
        self.window_manager.restore_window_geometry(self.app_config)
        
        # 设置窗口关闭处理
        self.window_manager.setup_window_close_handler(self._on_window_closing)
    
    def _create_widgets(self):
        """创建界面组件"""
        # 工具栏
        self._create_toolbar()
        
        # 主要内容区域
        self.main_frame = ttk.Frame(self.root)
        
        # 文档列表区域
        self._create_document_list()
        
        # 状态区域
        self._create_status_area()
        
        # 进度区域
        self._create_progress_area()
        
        # 创建列表操作处理器 (现在树形控件已创建)
        self.list_operation_handler = ListOperationHandler(
            self.document_manager, 
            self.doc_tree,
            self._on_list_operation_completed
        )
        
        # 设置界面提示功能
        self._setup_tooltips()
    
    def _create_toolbar(self):
        """创建工具栏"""
        self.toolbar = ttk.Frame(self.root)
        
        # 文件操作按钮
        self.btn_add_files = ttk.Button(
            self.toolbar, text="添加文件", 
            command=self._add_files
        )
        
        self.btn_add_folder = ttk.Button(
            self.toolbar, text="添加文件夹", 
            command=self._add_folder
        )
        
        self.btn_remove_selected = ttk.Button(
            self.toolbar, text="移除选中", 
            command=self._remove_selected_documents
        )
        
        self.btn_clear = ttk.Button(
            self.toolbar, text="文件过滤", 
            command=self._filter_documents
        )
        
        self.btn_filter = ttk.Button(
            self.toolbar, text="清空列表", 
            command=self._clear_documents
        )
        
        # 功能按钮
        self.btn_print_settings = ttk.Button(
            self.toolbar, text="打印设置", 
            command=self._show_print_settings
        )
        
        self.btn_help = ttk.Button(
            self.toolbar, text="使用说明", 
            command=self._show_help
        )
        
        self.btn_calculate_pages = ttk.Button(
            self.toolbar, text="计算页数", 
            command=self._calculate_pages
        )
        
        self.btn_start_print = ttk.Button(
            self.toolbar, text="开始打印", 
            command=self._start_printing,
            style="Accent.TButton"
        )
    
    def _create_document_list(self):
        """创建文档列表组件"""
        # 创建标题框架
        title_frame = ttk.Frame(self.main_frame)
        
        # 文档列表标题
        title_label = ttk.Label(title_frame, text="文档列表")
        
        # 文件类型过滤器
        self._create_file_type_filters(title_frame)
        
        # 文档列表框架
        list_frame = ttk.LabelFrame(self.main_frame, labelwidget=title_frame, padding="5")
        self.list_frame = list_frame
        
        # 创建树形视图
        self._create_tree_view(list_frame)
    
    def _create_file_type_filters(self, parent):
        """创建文件类型过滤器"""
        # 文件类型勾选框变量
        self.var_word = tk.BooleanVar(value=self.app_config.enabled_file_types.get('word', True))
        self.var_ppt = tk.BooleanVar(value=self.app_config.enabled_file_types.get('ppt', True))
        self.var_excel = tk.BooleanVar(value=self.app_config.enabled_file_types.get('excel', False))
        self.var_pdf = tk.BooleanVar(value=self.app_config.enabled_file_types.get('pdf', True))
        self.var_image = tk.BooleanVar(value=self.app_config.enabled_file_types.get('image', True))
        self.var_text = tk.BooleanVar(value=self.app_config.enabled_file_types.get('text', True))
        
        # 标题
        title_label = ttk.Label(parent, text="文档列表")
        title_label.pack(side="left")
        
        # 文件类型勾选框
        self.chk_word = ttk.Checkbutton(
            parent, text="Word", variable=self.var_word,
            command=self._on_filter_changed
        )
        self.chk_ppt = ttk.Checkbutton(
            parent, text="PPT", variable=self.var_ppt,
            command=self._on_filter_changed
        )
        self.chk_excel = ttk.Checkbutton(
            parent, text="Excel", variable=self.var_excel,
            command=self._on_filter_changed
        )
        self.chk_pdf = ttk.Checkbutton(
            parent, text="PDF", variable=self.var_pdf,
            command=self._on_filter_changed
        )
        self.chk_image = ttk.Checkbutton(
            parent, text="图片", variable=self.var_image,
            command=self._on_filter_changed
        )
        self.chk_text = ttk.Checkbutton(
            parent, text="文本", variable=self.var_text,
            command=self._on_filter_changed
        )
        
        # 布局
        self.chk_word.pack(side="left", padx=(10, 2))
        self.chk_ppt.pack(side="left", padx=2)
        self.chk_excel.pack(side="left", padx=2)
        self.chk_pdf.pack(side="left", padx=2)
        self.chk_image.pack(side="left", padx=2)
        self.chk_text.pack(side="left", padx=2)
    
    def _create_tree_view(self, parent):
        """创建树形视图"""
        # 创建树形视图容器
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill="both", expand=True)
        
        # 创建Treeview
        columns = ("文件名", "类型", "大小", "状态", "路径")
        self.doc_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=15)
        
        # 设置列标题和宽度
        self.doc_tree.heading("文件名", text="文件名")
        self.doc_tree.heading("类型", text="类型")
        self.doc_tree.heading("大小", text="大小(MB)")
        self.doc_tree.heading("状态", text="状态")
        self.doc_tree.heading("路径", text="文件路径")
        
        self.doc_tree.column("文件名", width=240)
        self.doc_tree.column("类型", width=80)
        self.doc_tree.column("大小", width=80)
        self.doc_tree.column("状态", width=50)
        self.doc_tree.column("路径", width=310)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.doc_tree.yview)
        self.doc_tree.configure(yscrollcommand=scrollbar.set)
        
        # 布局
        self.doc_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def _create_status_area(self):
        """创建状态显示区域"""
        status_frame = ttk.LabelFrame(self.main_frame, text="状态信息", padding="5")
        self.status_frame = status_frame
        
        # 状态标签
        self.lbl_doc_count = ttk.Label(status_frame, text="文档总数: 0")
        self.lbl_total_size = ttk.Label(status_frame, text="总大小: 0 MB")
        self.lbl_printer = ttk.Label(status_frame, text="打印机: 未设置")
        self.lbl_print_status = ttk.Label(status_frame, text="就绪")
        
        # 布局
        self.lbl_doc_count.grid(row=0, column=0, sticky="w", padx=5)
        self.lbl_total_size.grid(row=0, column=1, sticky="w", padx=5)
        self.lbl_printer.grid(row=1, column=0, sticky="w", padx=5)
        self.lbl_print_status.grid(row=1, column=1, sticky="w", padx=5)
    
    def _create_progress_area(self):
        """创建进度显示区域"""
        progress_frame = ttk.LabelFrame(self.main_frame, text="打印进度", padding="5")
        self.progress_frame = progress_frame
        
        # 进度条
        self.progress_bar = ttk.Progressbar(
            progress_frame, mode="determinate", length=400
        )
        
        # 进度标签
        self.lbl_progress = ttk.Label(progress_frame, text="等待开始...")
        
        # 布局
        self.progress_bar.grid(row=0, column=0, sticky="ew", padx=5)
        self.lbl_progress.grid(row=1, column=0, sticky="w", padx=5)
        
        progress_frame.grid_columnconfigure(0, weight=1)
    
    def _setup_layout(self):
        """设置布局"""
        # 工具栏布局 - 统一风格和间距
        self.btn_add_files.pack(side="left", padx=3)
        self.btn_add_folder.pack(side="left", padx=3)
        self.btn_remove_selected.pack(side="left", padx=3)
        self.btn_clear.pack(side="left", padx=3)
        self.btn_filter.pack(side="left", padx=3)
        self.btn_print_settings.pack(side="left", padx=3)
        self.btn_help.pack(side="left", padx=3)
        
        # 右侧按钮 - 统一间距
        self.btn_start_print.pack(side="right", padx=3)
        self.btn_calculate_pages.pack(side="right", padx=3)
        
        # 主框架布局
        self.toolbar.pack(fill="x", padx=10, pady=5)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 主内容布局
        self.list_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0, 5))
        self.status_frame.grid(row=1, column=0, sticky="ew", padx=(0, 5))
        self.progress_frame.grid(row=1, column=1, sticky="ew")
        
        # 设置权重
        self.main_frame.grid_rowconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=1)
    
    def _bind_events(self):
        """绑定事件"""
        # 确保列表操作处理器已初始化
        if self.list_operation_handler:
            # 设置列排序功能
            self.list_operation_handler.setup_column_sorting()
        
        # 设置拖拽功能
        self.file_import_handler.setup_drag_drop(self.doc_tree)
        
        # 创建右键菜单
        self._create_context_menu()
        
        # 文档列表事件
        self.doc_tree.bind("<Button-3>", self._show_context_menu)
        self.doc_tree.bind("<Double-1>", self._on_double_click)
        self.doc_tree.bind("<Delete>", self._on_delete_key)
    
    def _setup_tooltips(self):
        """设置界面提示功能"""
        # 只为文件类型过滤器添加tooltip
        create_button_tooltip(self.chk_word, 
            "Word文档过滤器\n支持格式: .doc, .docx, .wps")
        
        create_button_tooltip(self.chk_ppt, 
            "PowerPoint演示文稿过滤器\n支持格式: .ppt, .pptx, .dps")
        
        create_button_tooltip(self.chk_excel, 
            "Excel表格过滤器\n支持格式: .xls, .xlsx, .et\n注意:此格式只能模糊统计页数，请先手动再文件内排版好再打印")
        
        create_button_tooltip(self.chk_pdf, 
            "PDF文档过滤器\n支持标准PDF文件")
        
        create_button_tooltip(self.chk_image, 
            "图片文件过滤器\n支持格式: .jpg, .jpeg, .png, .bmp, .tiff, .tif, .webp\n注意: TIFF可能包含多页，其他图片按1页计算")
        
        create_button_tooltip(self.chk_text, 
            "文本文件过滤器\n支持格式: .txt\n注意:此格式只能模糊统计页数")
        
        # 为工具栏按钮添加tooltip
        create_button_tooltip(self.btn_filter, 
            "文件过滤\n根据上方勾选的文件类型过滤列表\n只保留勾选类型的文档，移除未勾选类型的文档")
        

    

    
    # === 文件操作相关方法 ===
    def _add_files(self):
        """添加文件"""
        added_count = self.file_import_handler.add_files_dialog()
        if added_count > 0:
            self._refresh_document_list()
            self._update_status()
    
    def _add_folder(self):
        """添加文件夹"""
        added_count = self.file_import_handler.add_folder_dialog()
        if added_count > 0:
            self._refresh_document_list()
            self._update_status()
    
    def _remove_selected_documents(self):
        """删除选中的文档"""
        if not self.list_operation_handler:
            return
        removed_count = self.list_operation_handler.remove_selected_documents()
        if removed_count > 0:
            self._refresh_document_list()
            self._update_status()
    
    def _clear_documents(self):
        """清空文档列表"""
        if not self.list_operation_handler:
            return
        if self.list_operation_handler.clear_all_documents():
            self._refresh_document_list()
            self._update_status()
    
    def _filter_documents(self):
        """根据勾选的文件类型过滤文档列表"""
        if not self.list_operation_handler:
            return
        
        # 获取当前启用的文件类型
        enabled_types = self._get_enabled_file_types()
        
        # 执行过滤操作
        filtered_count = self.list_operation_handler.filter_documents_by_enabled_types(enabled_types)
        
        if filtered_count > 0:
            self._refresh_document_list()
            self._update_status()
    
    # === 文件类型过滤器相关 ===
    def _get_enabled_file_types(self) -> dict:
        """获取当前启用的文件类型"""
        return {
            'word': self.var_word.get(),
            'ppt': self.var_ppt.get(),
            'excel': self.var_excel.get(),
            'pdf': self.var_pdf.get(),
            'image': self.var_image.get(),
            'text': self.var_text.get()
        }
    
    def _on_filter_changed(self):
        """文件类型过滤器变更事件"""
        enabled_types = self._get_enabled_file_types()
        self.window_manager.save_user_preferences(self.app_config, enabled_types)
        print(f"文件类型过滤器已更新: {enabled_types}")
    
    def _on_files_imported(self):
        """文件导入完成后的回调函数"""
        self._refresh_document_list()
        self._update_status()
    
    def _on_list_operation_completed(self):
        """列表操作完成后的回调函数"""
        self._refresh_document_list()
        self._update_status()
    
    # === 打印相关方法 ===
    def _show_print_settings(self):
        """显示打印设置对话框"""
        dialog = PrintSettingsDialog(
            self.root, 
            self.printer_manager, 
            self.current_print_settings
        )
        
        if dialog.result:
            self.current_print_settings = dialog.result
            self.print_controller.set_print_settings(self.current_print_settings)
            self.config_manager.save_print_settings(self.current_print_settings)
            self._update_status()
            print("打印设置已更新")
    
    def _start_printing(self):
        """开始批量打印"""
        if self.print_controller.is_printing:
            messagebox.showwarning("提示", "打印任务正在进行中")
            return
        
        if self.document_manager.document_count == 0:
            messagebox.showwarning("提示", "请先添加要打印的文档")
            return
        
        if not self.current_print_settings.printer_name:
            messagebox.showerror("错误", "请先设置打印机")
            return
        
        # 确认开始打印
        if not messagebox.askyesno("确认", f"确定要打印 {self.document_manager.document_count} 个文档吗？"):
            return
        
        try:
            # 添加文档到打印队列
            documents = self.document_manager.documents
            self.print_controller.clear_queue()
            self.print_controller.add_documents_to_queue(documents)
            
            # 开始打印
            future = self.print_controller.start_batch_print()
            
            # 禁用开始打印按钮
            self.btn_start_print.config(state="disabled")
            
            print("批量打印任务已启动")
            
        except Exception as e:
            messagebox.showerror("错误", f"启动打印失败: {e}")
    
    def _on_print_progress(self, current: int, total: int, message: str):
        """打印进度回调"""
        # 更新进度条
        progress = (current / total) * 100
        self.progress_bar['value'] = progress
        
        # 更新状态标签
        self.lbl_progress.config(text=f"{current}/{total} - {message}")
        self.lbl_print_status.config(text=f"打印中 ({current}/{total})")
        
        # 如果打印完成，重新启用按钮
        if current >= total:
            self.btn_start_print.config(state="normal")
            self.lbl_print_status.config(text="打印完成")
            self._refresh_document_list()  # 刷新状态显示
    
    # === 页数统计相关 ===
    def _calculate_pages(self):
        """计算页数"""
        if self.document_manager.document_count == 0:
            messagebox.showwarning("提示", "请先添加要统计的文档")
            return
        
        # 显示页数统计对话框
        show_page_count_dialog(self.root, self.document_manager.documents)
    
    def _calculate_selected_pages(self):
        """计算选中文档的页数"""
        if not self.list_operation_handler:
            return
        selected_documents = self.list_operation_handler.get_selected_document_objects()
        if not selected_documents:
            messagebox.showwarning("提示", "请先选择要计算页数的文档")
            return
        
        # 调用页数统计功能
        try:
            show_page_count_dialog(self.root, selected_documents)
        except Exception as e:
            messagebox.showerror("错误", f"页数统计功能出错: {e}")
    
    # === 界面更新相关方法 ===
    def _refresh_document_list(self):
        """刷新文档列表显示"""
        # 清空现有项目
        for item in self.doc_tree.get_children():
            self.doc_tree.delete(item)
        
        # 添加文档
        for doc in self.document_manager.documents:
            status_text = {
                PrintStatus.PENDING: "待打印",
                PrintStatus.PRINTING: "打印中",
                PrintStatus.COMPLETED: "已完成",
                PrintStatus.ERROR: "失败"
            }.get(doc.print_status, "未知")
            
            self.doc_tree.insert("", "end", values=(
                doc.file_name,
                doc.type_display,
                doc.size_mb,
                status_text,
                str(doc.file_path)
            ))
        
        # 保持排序指示器显示
        if self.list_operation_handler:
            self.list_operation_handler.maintain_sort_indicators()
    
    def _update_status(self):
        """更新状态显示"""
        summary = self.document_manager.get_summary()
        
        # 更新文档统计
        self.lbl_doc_count.config(text=f"文档总数: {summary['total']}")
        self.lbl_total_size.config(text=f"总大小: {summary['total_size_mb']} MB")
        
        # 更新打印机信息
        printer_name = self.current_print_settings.printer_name or "未设置"
        self.lbl_printer.config(text=f"打印机: {printer_name}")
        
        # 更新按钮状态
        if self.document_manager.document_count == 0:
            self.btn_calculate_pages.config(state="disabled")
        else:
            self.btn_calculate_pages.config(state="normal")
    
    # === 右键菜单相关 ===
    def _create_context_menu(self):
        """创建右键菜单"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
    
    def _show_context_menu(self, event):
        """显示右键菜单"""
        selection = self.doc_tree.selection()
        if not selection:
            return
        
        # 清空现有菜单项
        self.context_menu.delete(0, "end")
        
        # 根据选中项目数量构建菜单
        selected_count = len(selection)
        
        if selected_count == 1:
            # 单文件菜单
            self.context_menu.add_command(
                label="📁  打开所在文件夹",
                command=lambda: self.list_operation_handler.open_file_location() if self.list_operation_handler else None
            )
            self.context_menu.add_command(
                label="📄  仅打印选中文档",
                command=self._print_selected_documents
            )
            self.context_menu.add_command(
                label="❌  从列表中移除",
                command=self._remove_selected_documents
            )
            self.context_menu.add_separator()
            self.context_menu.add_command(
                label="🔄  重置排序",
                command=self._reset_sort
            )
            self.context_menu.add_command(
                label="🔍  文件过滤",
                command=self._filter_documents
            )
            self.context_menu.add_separator()
            self.context_menu.add_command(
                label="📊  计算选中文档页数",
                command=self._calculate_selected_pages
            )
        else:
            # 多文件菜单
            self.context_menu.add_command(
                label=f"📄  仅打印选中文档 ({selected_count}个文件)",
                command=self._print_selected_documents
            )
            self.context_menu.add_command(
                label=f"❌  从列表中移除 ({selected_count}个文件)",
                command=self._remove_selected_documents
            )
            self.context_menu.add_separator()
            self.context_menu.add_command(
                label="🔄  重置排序",
                command=self._reset_sort
            )
            self.context_menu.add_command(
                label="🔍  文件过滤",
                command=self._filter_documents
            )
            self.context_menu.add_separator()
            self.context_menu.add_command(
                label=f"📊  计算选中文档页数 ({selected_count}个文件)",
                command=self._calculate_selected_pages
            )
            self.context_menu.add_command(
                label=f"💾  导出选中文档列表 ({selected_count}个文件)",
                command=lambda: self.list_operation_handler.export_document_list("selected") if self.list_operation_handler else None
            )
        
        # 显示菜单
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
    
    def _print_selected_documents(self):
        """仅打印选中的文档"""
        if not self.list_operation_handler:
            return
        selected_documents = self.list_operation_handler.get_selected_document_objects()
        if not selected_documents:
            messagebox.showwarning("提示", "请先选择要打印的文档")
            return
        
        if self.print_controller.is_printing:
            messagebox.showwarning("提示", "打印任务正在进行中")
            return
        
        if not self.current_print_settings.printer_name:
            messagebox.showerror("错误", "请先设置打印机")
            return
        
        # 确认打印
        count = len(selected_documents)
        if not messagebox.askyesno("确认", f"确定要打印选中的 {count} 个文档吗？"):
            return
        
        try:
            # 添加选中文档到打印队列
            self.print_controller.clear_queue()
            self.print_controller.add_documents_to_queue(selected_documents)
            
            # 开始打印
            future = self.print_controller.start_batch_print()
            
            # 禁用开始打印按钮
            self.btn_start_print.config(state="disabled")
            
            print(f"已开始打印选中的 {count} 个文档")
            
        except Exception as e:
            messagebox.showerror("错误", f"启动打印失败: {e}")
    
    def _reset_sort(self):
        """重置排序"""
        if not self.list_operation_handler:
            return
        self.list_operation_handler.reset_sort()
        self._refresh_document_list()
    
    # === 事件处理 ===
    def _on_double_click(self, event):
        """双击文档列表项处理"""
        # 检查是否点击在列标题区域
        region = self.doc_tree.identify_region(event.x, event.y)
        if region == "heading":
            # 点击在列标题区域，不处理双击事件
            return
        
        # 检查是否有选中的项目
        item = self.doc_tree.identify_row(event.y)
        if item and self.list_operation_handler:
            self.list_operation_handler.open_file_with_default_app()
    
    def _on_delete_key(self, event):
        """删除键处理"""
        self._remove_selected_documents()
    
    def _on_window_closing(self):
        """窗口关闭事件处理"""
        # 检查是否有打印任务正在进行
        if self.print_controller.is_printing:
            if not messagebox.askyesno("确认", "打印任务正在进行中，确定要退出吗？"):
                return False  # 取消关闭
        
        # 保存窗口几何属性
        self.window_manager.save_window_geometry(self.app_config)
        
        return True  # 允许关闭
    
    # === 使用说明 ===
    def _show_help(self):
        """显示使用说明"""
        help_text = """
📖 办公文档批量打印器使用说明 V5.0

═══════════════════════════════════════

🎯 软件功能
• 批量添加和打印多种格式文档
• 方便的过滤各种文档
• 灵活的打印设置配置
• 便捷的文档管理
• 页数统计功能

═══════════════════════════════════════

📂 支持的文件格式

📝 Office文档：
   • Word文档：.doc, .docx, .wps (WPS文字)
   • PowerPoint：.ppt, .pptx, .dps (WPS演示)  
   • Excel表格：.xls, .xlsx, .et (WPS表格)
   请慎重选择excel格式，使用前先在文件内把打印排版调校好

📄 通用文档：
   • PDF文件：.pdf
   • 文本文件：.txt

🖼️ 图片文件：
   • 常见格式：.jpg, .jpeg, .png, .bmp
   • 高级格式：.tiff, .tif, .webp
   • 注意：TIFF可能包含多页，其他按1页计算

═══════════════════════════════════════

📋 使用步骤

1 文件类型过滤
   • 勾选/取消勾选各种文件类型过滤器
   • Word、PPT、Excel、PDF、图片、文本独立控制
   • 鼠标悬停过滤器可查看支持的扩展名

2 添加文档
   • 点击按钮添加选择单个或多个文档以及文件夹
   • 直接拖拽文件或文件夹到程序窗口进行快速添加
   • 支持递归搜索子文件夹

3 管理文档
   • 可单、多选文档后进行操作
   • 双击文档可用默认程序打开预览

4 页数统计
   • 点击"计算页数"统计所有文档的页数
   • 可多选后用右键统计选中文档
   • 可导出统计结果到Excel

5 配置打印
   • 点击"打印设置"配置打印参数
   • 选择打印机、纸张尺寸、页面方向
   • 设置打印份数、双面打印、颜色模式

6 开始打印
   • 确认文档列表和打印设置
   • 点击"开始打印"执行批量打印
   • 观察进度条了解打印状态
   • 支持右键菜单打印选中文档

═══════════════════════════════════════

🚀 v5.0 新特性

🏗️ 架构升级：
   • 重新编写构架使用模块化文件处理器设计

📈 功能增强：
   • 新增图片文件支持（.jpg, .png, .bmp, .tiff, .webp等）
   • 新增文本文件支持（.txt）
   • 新增WPS格式完整支持（.wps, .dps, .et）
   • 新增文件过滤功能
   • 新增右键增强功能
   • 新增文件排序功能
   • 新增浮窗提示功能
🔧 技术改进：
   • 优化拖拽导入功能
   • 增强双面打印成功率
   • 修复PDF偶尔无法打印问题
   • 修复纸张规格无法同步问题

═══════════════════════════════════════

💡 使用提示

• Excel和文本文件页数为估算值，打印前建议预览确认
• 大文件处理可能需要较长时间，请耐心等待
• 确保打印机驱动程序正确安装
• 建议定期检查打印机连接状态
• 支持的最大文本文件大小：100MB

═══════════════════════════════════════

💝 感谢使用办公文档批量打印器！
开发者：喵言喵语 by.52pojie
        """
        
        # 创建帮助窗口
        help_window = tk.Toplevel(self.root)
        help_window.title("使用说明")
        help_window.geometry("650x700")
        help_window.resizable(True, True)
        help_window.transient(self.root)
        help_window.grab_set()
        
        # 居中显示
        self.window_manager.center_window(help_window)
        
        # 创建滚动文本框
        main_frame = ttk.Frame(help_window, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # 文本显示区域
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill="both", expand=True)
        
        # 创建文本控件和滚动条
        text_widget = tk.Text(
            text_frame,
            wrap=tk.WORD,
            font=("Microsoft YaHei", 10),
            bg="white",
            fg="black",
            relief="flat",
            borderwidth=0,
            state="normal"
        )
        
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        # 布局
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 插入帮助文本
        text_widget.insert("1.0", help_text)
        text_widget.config(state="disabled")  # 设为只读
        
        # 关闭按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(10, 0))
        
        close_btn = ttk.Button(
            button_frame,
            text="关闭",
            command=help_window.destroy,
            width=10
        )
        close_btn.pack(side="right")
    
    def run(self):
        """运行主循环"""
        # 初始更新
        self._update_status()
        
        # 启动主循环
        self.root.mainloop()


def main():
    """主函数"""
    try:
        app = MainWindow()
        app.run()
    except Exception as e:
        print(f"应用程序启动失败: {e}")
        messagebox.showerror("错误", f"应用程序启动失败: {e}")


if __name__ == "__main__":
    main() 