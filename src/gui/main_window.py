"""
主窗口界面
办公文档批量打印应用的主界面
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


class MainWindow:
    """主窗口类"""
    
    def __init__(self):
        """初始化主窗口"""
        # 创建支持拖拽的主窗口
        if DRAG_DROP_AVAILABLE:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        
        self.root.title("办公文档批量打印器 v4.0 by.喵言喵语")
        self.root.geometry("900x600")
        self.root.minsize(800, 500)
        
        # 初始化管理器
        self.document_manager = DocumentManager()
        self.printer_manager = PrinterSettingsManager()
        self.print_controller = PrintController()
        self.config_manager = ConfigManager()
        
        # 加载配置
        self.app_config = self.config_manager.load_app_config()
        self.current_print_settings = self.config_manager.load_print_settings()
        
        # 设置打印控制器
        self.print_controller.set_print_settings(self.current_print_settings)
        self.print_controller.set_progress_callback(self._on_print_progress)
        
        # 创建界面
        self._create_widgets()
        self._setup_layout()
        self._bind_events()
        
        # 恢复窗口几何属性
        self._restore_window_geometry()
        
        print("批量文档打印器已启动")
    
    def _create_widgets(self):
        """创建界面组件"""
        # 工具栏
        self.toolbar = ttk.Frame(self.root)
        
        # 添加文件按钮
        self.btn_add_files = ttk.Button(
            self.toolbar, text="添加文件", 
            command=self._add_files
        )
        
        # 添加文件夹按钮
        self.btn_add_folder = ttk.Button(
            self.toolbar, text="添加文件夹", 
            command=self._add_folder
        )
        
        # 删除选中按钮
        self.btn_remove_selected = ttk.Button(
            self.toolbar, text="删除选中", 
            command=self._remove_selected_documents
        )
        
        # 清空列表按钮
        self.btn_clear = ttk.Button(
            self.toolbar, text="清空列表", 
            command=self._clear_documents
        )
        
        # 打印设置按钮
        self.btn_print_settings = ttk.Button(
            self.toolbar, text="打印设置", 
            command=self._show_print_settings
        )
        
        # 使用说明按钮
        self.btn_help = ttk.Button(
            self.toolbar, text="使用说明", 
            command=self._show_help
        )
        
        # 计算页数按钮
        self.btn_calculate_pages = ttk.Button(
            self.toolbar, text="计算页数", 
            command=self._calculate_pages
        )
        
        # 开始打印按钮
        self.btn_start_print = ttk.Button(
            self.toolbar, text="开始打印", 
            command=self._start_printing,
            style="Accent.TButton"
        )
        
        # 主要内容区域
        self.main_frame = ttk.Frame(self.root)
        
        # 文档列表区域
        self._create_document_list()
        
        # 状态区域
        self._create_status_area()
        
        # 进度区域
        self._create_progress_area()
    
    def _create_document_list(self):
        """创建文档列表组件"""
        # 创建标题框架用于放置在LabelFrame的标题位置
        title_frame = ttk.Frame(self.main_frame)
        
        # 文档列表标题
        title_label = ttk.Label(title_frame, text="文档列表")
        
        # 文件类型勾选框变量
        self.var_word = tk.BooleanVar(value=self.app_config.enabled_file_types.get('word', True))
        self.var_ppt = tk.BooleanVar(value=self.app_config.enabled_file_types.get('ppt', True))
        self.var_excel = tk.BooleanVar(value=self.app_config.enabled_file_types.get('excel', False))
        self.var_pdf = tk.BooleanVar(value=self.app_config.enabled_file_types.get('pdf', True))
        
        # 文件类型勾选框
        self.chk_word = ttk.Checkbutton(
            title_frame, text="Word", variable=self.var_word,
            command=self._on_filter_changed
        )
        self.chk_ppt = ttk.Checkbutton(
            title_frame, text="PPT", variable=self.var_ppt,
            command=self._on_filter_changed
        )
        self.chk_excel = ttk.Checkbutton(
            title_frame, text="Excel", variable=self.var_excel,
            command=self._on_filter_changed
        )
        self.chk_pdf = ttk.Checkbutton(
            title_frame, text="PDF", variable=self.var_pdf,
            command=self._on_filter_changed
        )
        
        # 布局标题和过滤器
        title_label.pack(side="left")
        self.chk_word.pack(side="left", padx=(10, 2))
        self.chk_ppt.pack(side="left", padx=2)
        self.chk_excel.pack(side="left", padx=2)
        self.chk_pdf.pack(side="left", padx=2)
        
        # 文档列表框架
        list_frame = ttk.LabelFrame(self.main_frame, labelwidget=title_frame, padding="5")
        self.list_frame = list_frame
        
        # 创建树形视图容器
        tree_frame = ttk.Frame(list_frame)
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
        
        self.doc_tree.column("文件名", width=200)
        self.doc_tree.column("类型", width=100)
        self.doc_tree.column("大小", width=80)
        self.doc_tree.column("状态", width=80)
        self.doc_tree.column("路径", width=300)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.doc_tree.yview)
        self.doc_tree.configure(yscrollcommand=scrollbar.set)
        
        # 布局
        self.doc_tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
    
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
        # 工具栏布局
        self.btn_add_files.pack(side="left", padx=2)
        self.btn_add_folder.pack(side="left", padx=2)
        self.btn_remove_selected.pack(side="left", padx=2)
        self.btn_clear.pack(side="left", padx=2)
        self.btn_print_settings.pack(side="left", padx=10)
        self.btn_help.pack(side="left", padx=2)
        
        # 右侧按钮（从右到左的顺序）
        self.btn_start_print.pack(side="right", padx=5)
        self.btn_calculate_pages.pack(side="right", padx=2)
        
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
        # 窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # 文档列表右键菜单
        self.doc_tree.bind("<Button-3>", self._show_context_menu)
        
        # 双击文档列表项打开文件
        self.doc_tree.bind("<Double-1>", self._on_double_click)
        
        # 删除键删除选中文档
        self.doc_tree.bind("<Delete>", self._on_delete_key)
        
        # 拖拽支持
        if DRAG_DROP_AVAILABLE:
            self._setup_drag_drop()
    
    def _setup_drag_drop(self):
        """设置拖拽功能"""
        try:
            # 为文档列表区域注册拖拽
            self.doc_tree.drop_target_register(DND_FILES)
            self.doc_tree.dnd_bind('<<Drop>>', self._on_drop_files)
            
            # 为主窗口注册拖拽（备选）
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self._on_drop_files)
            
            print("✓ 拖拽功能已启用")
        except Exception as e:
            print(f"✗ 拖拽功能初始化失败: {e}")
    
    def _on_drop_files(self, event):
        """处理拖拽文件事件"""
        try:
            # 获取拖拽的文件路径
            files = event.data.split()
            if not files:
                return
            
            file_paths = []
            folder_paths = []
            
            # 分离文件和文件夹
            for file_str in files:
                # 去除可能的引号
                file_str = file_str.strip('"\'')
                file_path = Path(file_str)
                
                if file_path.exists():
                    if file_path.is_file():
                        file_paths.append(file_path)
                    elif file_path.is_dir():
                        folder_paths.append(file_path)
            
            # 处理文件
            added_count = 0
            if file_paths:
                added_docs = self.document_manager.add_files(file_paths)
                added_count += len(added_docs)
                print(f"拖拽添加文件: {len(added_docs)} 个")
            
            # 处理文件夹
            if folder_paths:
                enabled_file_types = self._get_enabled_file_types()
                for folder_path in folder_paths:
                    # 默认递归搜索
                    added_docs = self.document_manager.add_folder(folder_path, True, enabled_file_types)
                    added_count += len(added_docs)
                    print(f"拖拽添加文件夹 {folder_path.name}: {len(added_docs)} 个文档")
            
            # 更新界面
            if added_count > 0:
                self._refresh_document_list()
                self._update_status()
                messagebox.showinfo("拖拽导入成功", f"成功导入 {added_count} 个文档")
            else:
                messagebox.showwarning("拖拽导入", "未找到支持的文档格式或文件已存在")
                
        except Exception as e:
            print(f"拖拽处理错误: {e}")
            messagebox.showerror("拖拽导入失败", f"处理拖拽文件时出错：{str(e)}")
    
    def _add_files(self):
        """添加文件"""
        file_types = [
            ("支持的文档", "*.pdf;*.doc;*.docx;*.ppt;*.pptx;*.xls;*.xlsx"),
            ("PDF文件", "*.pdf"),
            ("Word文档", "*.doc;*.docx"),
            ("PowerPoint", "*.ppt;*.pptx"),
            ("Excel表格", "*.xls;*.xlsx"),
            ("所有文件", "*.*")
        ]
        
        files = filedialog.askopenfilenames(
            title="选择要打印的文档",
            filetypes=file_types
        )
        
        if files:
            file_paths = [Path(f) for f in files]
            added_docs = self.document_manager.add_files(file_paths)
            
            if added_docs:
                self._refresh_document_list()
                self._update_status()
                messagebox.showinfo("成功", f"成功添加 {len(added_docs)} 个文档")
    
    def _add_folder(self):
        """添加文件夹"""
        folder = filedialog.askdirectory(title="选择包含文档的文件夹")
        
        if folder:
            folder_path = Path(folder)
            
            # 询问是否递归搜索
            recursive = messagebox.askyesno(
                "搜索选项", 
                "是否搜索子文件夹中的文档？"
            )
            
            # 获取当前的文件类型过滤设置
            enabled_file_types = self._get_enabled_file_types()
            
            added_docs = self.document_manager.add_folder(folder_path, recursive, enabled_file_types)
            
            if added_docs:
                self._refresh_document_list()
                self._update_status()
                messagebox.showinfo("成功", f"从文件夹中成功添加 {len(added_docs)} 个文档")
            else:
                messagebox.showwarning("提示", "在指定文件夹中未找到支持的文档")
    
    def _get_enabled_file_types(self) -> dict:
        """获取当前启用的文件类型"""
        return {
            'word': self.var_word.get(),
            'ppt': self.var_ppt.get(),
            'excel': self.var_excel.get(),
            'pdf': self.var_pdf.get()
        }
    
    def _on_filter_changed(self):
        """文件类型过滤器变更事件"""
        # 更新应用配置
        self.app_config.enabled_file_types = self._get_enabled_file_types()
        
        # 保存配置
        self.config_manager.save_app_config(self.app_config)
        
        print(f"文件类型过滤器已更新: {self.app_config.enabled_file_types}")
    
    def _remove_selected_documents(self):
        """删除选中的文档"""
        selection = self.doc_tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先选择要删除的文档")
            return
        
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selection)} 个文档吗？"):
            # 获取选中文档的文件名并删除
            removed_count = 0
            for item in selection:
                values = self.doc_tree.item(item, 'values')
                if values:
                    file_name = values[0]
                    # 从文档管理器中找到并移除对应文档
                    for doc in self.document_manager.documents:
                        if doc.file_name == file_name:
                            if self.document_manager.remove_document(doc.id):
                                removed_count += 1
                            break
            
            # 刷新显示
            self._refresh_document_list()
            self._update_status()
            
            if removed_count > 0:
                messagebox.showinfo("完成", f"已删除 {removed_count} 个文档")
    
    def _clear_documents(self):
        """清空文档列表"""
        if self.document_manager.document_count == 0:
            return
        
        if messagebox.askyesno("确认", "确定要清空所有文档吗？"):
            self.document_manager.clear_all()
            self._refresh_document_list()
            self._update_status()
    
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
    
    def _calculate_pages(self):
        """计算页数"""
        if self.document_manager.document_count == 0:
            messagebox.showwarning("提示", "请先添加要统计的文档")
            return
        
        # 显示页数统计对话框
        show_page_count_dialog(self.root, self.document_manager.documents)
    
    def _update_calculate_button_state(self):
        """更新计算页数按钮状态"""
        if self.document_manager.document_count == 0:
            self.btn_calculate_pages.config(state="disabled")
        else:
            self.btn_calculate_pages.config(state="normal")
    
    def _show_help(self):
        """显示使用说明"""
        help_text = """
            📖 办公文档批量打印器使用说明 V4.0

═══════════════════════════════════════

🎯 软件功能
• 批量添加和打印Word、PowerPoint、Excel、PDF文档
• 文件类型过滤器：选择要扫描的文档类型
• 灵活的打印设置配置
• 实时打印进度显示
• 便捷的文档管理
• 页数统计功能

═══════════════════════════════════════

📋 使用步骤

1️⃣ 添加文档
   • 点击"添加文件"选择单个或多个文档
   • 点击"添加文件夹"批量添加整个文件夹中的文档
   • 🆕 直接拖拽文件或文件夹到程序窗口进行快速添加
   • 支持递归搜索子文件夹
   • 使用文件类型过滤器选择要扫描的文档类型（Word、PPT、Excel、PDF）
   • 默认不扫码excel，表格打印容易排版错位，请先手动调整好排版


2️⃣ 管理文档
   • 选中文档后点击"删除选中"可移除特定文档
   • 点击"清空列表"可移除所有文档
   • 双击文档可用默认程序打开预览
   • 按Delete键可快速删除选中文档

3️⃣ 页数统计
   • 点击"计算页数"可统计所有文档的页数
   • 支持PDF、Word、PowerPoint、Excel文档页数计算
   • 显示详细统计报告和问题文件
   • 可导出完整报告或错误报告

4️⃣ 配置打印
   • 点击"打印设置"配置打印参数：
     - 选择打印机
     - 设置纸张尺寸（A4、A3、Letter等）
     - 选择页面方向（纵向/横向）
     - 设置打印数量（1-999份）
     - 选择颜色模式（彩色/黑白）
     - 启用双面打印（如果打印机支持）

5️⃣ 开始打印
   • 确认文档列表和打印设置
   • 点击"开始打印"执行批量打印
   • 观察进度条了解打印状态

═══════════════════════════════════════

📁 支持的文件格式
• Word文档：.doc、.docx
• PowerPoint演示文稿：.ppt、.pptx
• Excel表格：.xls、.xlsx （慎重选择，打印前先手动调整好排版）
• PDF文件：.pdf

═══════════════════════════════════════

⚠️ 注意事项

🔹 系统要求
• Windows 10/11 操作系统
• 已安装Microsoft Office（Word、PowerPoint和Excel打印需要）
• 至少一台可用的打印机

🔹 使用提示
• 打印前请确保打印机正常连接
• 大批量打印时请确保纸张充足
• PDF文件使用系统默认程序打印
• Word、PowerPoint和Excel通过Office应用打印
• 使用文件类型过滤器可控制扫描文件夹时包含的文档类型
• 打印过程中请勿关闭应用程序

🔹 故障排除
• 如打印失败，请检查文件是否被其他程序占用
• 确认打印机驱动程序已正确安装
• 对于PDF文件，确保系统已安装PDF阅读器
• 页面统计大文件时间会较长请耐心等待
• 页面统计遇到加密文件会卡主需手动关闭打开的文档

═══════════════════════════════════════

💝 感谢使用办公文档批量打印器！
开发者：喵言喵语 2025.6.22 by.52pojie
        """
        
        # 创建帮助窗口
        help_window = tk.Toplevel(self.root)
        help_window.title("使用说明")
        help_window.geometry("650x700")
        help_window.resizable(True, True)
        help_window.transient(self.root)
        help_window.grab_set()
        
        # 居中显示
        self._center_window(help_window)
        
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
    
    def _center_window(self, window):
        """将窗口居中显示"""
        window.update_idletasks()
        
        # 获取主窗口位置和大小
        main_x = self.root.winfo_rootx()
        main_y = self.root.winfo_rooty()
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()
        
        # 计算子窗口位置
        window_width = window.winfo_reqwidth()
        window_height = window.winfo_reqheight()
        
        x = main_x + (main_width - window_width) // 2
        y = main_y + (main_height - window_height) // 2
        
        window.geometry(f"+{x}+{y}")
    
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
        self._update_calculate_button_state()
    
    def _show_context_menu(self, event):
        """显示右键菜单"""
        # 简化实现：选中项目时显示删除选项
        selection = self.doc_tree.selection()
        if selection:
            # 这里可以添加右键菜单
            pass
    
    def _on_double_click(self, event):
        """双击文档列表项处理"""
        selection = self.doc_tree.selection()
        if selection:
            # 获取选中的文档路径
            item = selection[0]
            values = self.doc_tree.item(item, 'values')
            if values and len(values) > 4:
                file_path = values[4]  # 文件路径在第5列
                try:
                    # 使用系统默认程序打开文件
                    import os
                    os.startfile(file_path)
                except Exception as e:
                    messagebox.showerror("错误", f"无法打开文件: {e}")
    
    def _on_delete_key(self, event):
        """删除键处理"""
        selection = self.doc_tree.selection()
        if selection:
            if messagebox.askyesno("确认", f"确定要从列表中移除 {len(selection)} 个文档吗？"):
                # 获取选中文档的ID并移除
                for item in selection:
                    values = self.doc_tree.item(item, 'values')
                    if values:
                        file_name = values[0]
                        # 从文档管理器中找到并移除对应文档
                        for doc in self.document_manager.documents:
                            if doc.file_name == file_name:
                                self.document_manager.remove_document(doc.id)
                                break
                
                # 刷新显示
                self._refresh_document_list()
                self._update_status()
    
    def _restore_window_geometry(self):
        """恢复窗口几何属性"""
        geometry = self.app_config.window_geometry
        if geometry:
            try:
                self.root.geometry(f"{geometry.get('width', 900)}x{geometry.get('height', 600)}")
                if 'x' in geometry and 'y' in geometry:
                    self.root.geometry(f"+{geometry['x']}+{geometry['y']}")
            except:
                pass  # 使用默认几何属性
    
    def _save_window_geometry(self):
        """保存窗口几何属性"""
        try:
            geometry = self.root.geometry()
            # 解析几何字符串 "widthxheight+x+y"
            parts = geometry.replace('+', ' +').replace('-', ' -').split()
            if len(parts) >= 3:
                size_part = parts[0]
                width, height = map(int, size_part.split('x'))
                x = int(parts[1])
                y = int(parts[2])
                
                self.app_config.window_geometry = {
                    'width': width,
                    'height': height,
                    'x': x,
                    'y': y
                }
                
                self.config_manager.save_app_config(self.app_config)
        except:
            pass  # 忽略保存错误
    
    def _on_closing(self):
        """窗口关闭事件处理"""
        # 检查是否有打印任务正在进行
        if self.print_controller.is_printing:
            if not messagebox.askyesno("确认", "打印任务正在进行中，确定要退出吗？"):
                return
        
        # 保存窗口几何属性
        self._save_window_geometry()
        
        # 关闭窗口
        self.root.destroy()
    
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