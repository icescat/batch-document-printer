"""
列表操作功能处理器
处理文档列表的各种操作：排序、文件操作、导出、右键菜单功能等
"""
import tkinter as tk
from tkinter import messagebox, filedialog
from pathlib import Path
from typing import List, Dict, Any, Optional
import subprocess
import csv


class ListOperationHandler:
    """列表操作功能处理器"""
    
    def __init__(self, document_manager, tree_widget, refresh_callback=None):
        """
        初始化列表操作处理器
        
        Args:
            document_manager: 文档管理器实例
            tree_widget: 文档列表树形控件
            refresh_callback: 界面刷新回调函数
        """
        self.document_manager = document_manager
        self.tree_widget = tree_widget
        self.refresh_callback = refresh_callback
        
        # 列名到排序键的映射
        self.column_sort_map = {
            "文件名": "name",
            "类型": "type", 
            "大小": "size",
            "状态": "status",
            "路径": "path"
        }
        
        # 列标题的基础文本
        self.base_column_texts = {
            "文件名": "文件名",
            "类型": "类型",
            "大小": "大小(MB)",
            "状态": "状态",
            "路径": "文件路径"
        }
    
    def setup_column_sorting(self):
        """设置列标题排序功能"""
        for column in self.column_sort_map.keys():
            self.tree_widget.heading(column, command=lambda col=column: self.handle_column_click(col))
    
    def handle_column_click(self, column_name: str):
        """处理列标题点击事件"""
        if column_name not in self.column_sort_map:
            return
        
        sort_key = self.column_sort_map[column_name]
        
        # 切换排序状态
        is_reverse = self.document_manager.toggle_sort(sort_key)
        
        # 更新列标题显示排序指示器
        self.update_sort_indicators(column_name, is_reverse)
        
        print(f"按 {column_name} {'降序' if is_reverse else '升序'} 排序")
        
        # 触发界面刷新
        if hasattr(self, 'refresh_callback') and self.refresh_callback:
            self.refresh_callback()
        
        return is_reverse
    
    def update_sort_indicators(self, active_column: str, is_reverse: bool):
        """更新列标题显示排序指示器"""
        # 更新所有列标题
        for column in self.column_sort_map.keys():
            base_text = self.base_column_texts.get(column, column)
            if column == active_column:
                # 活跃列显示排序指示器
                indicator = " ▼" if is_reverse else " ▲"
                text = base_text + indicator
            else:
                # 非活跃列不显示指示器
                text = base_text
            
            self.tree_widget.heading(column, text=text)
    
    def maintain_sort_indicators(self):
        """维护排序指示器显示（用于刷新列表时保持排序状态）"""
        sort_key, is_reverse = self.document_manager.current_sort_info
        if sort_key:
            # 找到对应的列名
            column_name = None
            for col, key in self.column_sort_map.items():
                if key == sort_key:
                    column_name = col
                    break
            
            if column_name:
                self.update_sort_indicators(column_name, is_reverse)
    
    def reset_sort(self):
        """重置排序到原始顺序"""
        # 清除DocumentManager中的排序状态
        self.document_manager._current_sort_key = None
        self.document_manager._current_sort_reverse = False
        
        # 按添加时间重新排序（恢复原始顺序）
        self.document_manager.sort_documents('added_time', False)
        
        # 清除所有列标题的排序指示器
        for column, text in self.base_column_texts.items():
            self.tree_widget.heading(column, text=text)
        
        print("排序已重置到原始顺序")
    
    def get_selected_documents(self) -> List[Dict[str, Any]]:
        """获取选中的文档信息"""
        selection = self.tree_widget.selection()
        selected_documents = []
        
        for item in selection:
            values = self.tree_widget.item(item, 'values')
            if values and len(values) >= 5:
                doc_info = {
                    'file_name': values[0],
                    'type': values[1],
                    'size': values[2],
                    'status': values[3],
                    'path': values[4]
                }
                selected_documents.append(doc_info)
        
        return selected_documents
    
    def get_selected_document_objects(self) -> List:
        """获取选中的文档对象"""
        selected_info = self.get_selected_documents()
        selected_file_names = [info['file_name'] for info in selected_info]
        
        # 从文档管理器中找到对应的文档对象
        selected_documents = []
        for doc in self.document_manager.documents:
            if doc.file_name in selected_file_names:
                selected_documents.append(doc)
        
        return selected_documents
    
    def remove_selected_documents(self) -> int:
        """删除选中的文档"""
        selected_docs = self.get_selected_documents()
        if not selected_docs:
            messagebox.showinfo("提示", "请先选择要删除的文档")
            return 0
        
        if messagebox.askyesno("确认", f"确定要删除选中的 {len(selected_docs)} 个文档吗？"):
            # 根据文件名删除文档
            removed_count = 0
            for doc_info in selected_docs:
                file_name = doc_info['file_name']
                # 从文档管理器中找到并移除对应文档
                for doc in self.document_manager.documents:
                    if doc.file_name == file_name:
                        if self.document_manager.remove_document(doc.id):
                            removed_count += 1
                        break
            
            if removed_count > 0:
                # 去掉删除完成提示窗口
                print(f"已删除 {removed_count} 个文档")
            
            return removed_count
        
        return 0
    
    def open_file_location(self, file_path: Optional[str] = None):
        """打开文件所在文件夹"""
        if not file_path:
            selected_docs = self.get_selected_documents()
            if not selected_docs:
                messagebox.showwarning("提示", "请先选择要定位的文档")
                return
            file_path = selected_docs[0]['path']
        
        if not file_path:
            return
            
        try:
            path = Path(file_path)
            # 在Windows文件管理器中打开并选中文件
            result = subprocess.run(['explorer', '/select,', str(path)], 
                                  capture_output=True, text=True)
            # 只有在真正失败时才显示错误（例如文件不存在）
            if result.returncode != 0 and not path.exists():
                messagebox.showerror("错误", f"文件不存在: {path}")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹: {e}")
    
    def open_file_with_default_app(self, file_path: Optional[str] = None):
        """用默认程序打开文件"""
        if not file_path:
            selected_docs = self.get_selected_documents()
            if not selected_docs:
                messagebox.showwarning("提示", "请先选择要打开的文档")
                return
            file_path = selected_docs[0]['path']
        
        if not file_path:
            return
            
        try:
            # 使用系统默认程序打开文件
            import os
            os.startfile(file_path)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件: {e}")
    
    def export_document_list(self, export_type: str = "selected", file_format: str = "csv"):
        """
        导出文档列表
        
        Args:
            export_type: "selected" 或 "all"
            file_format: "csv", "xlsx", "txt"
        """
        if export_type == "selected":
            documents_data = self.get_selected_documents()
            if not documents_data:
                messagebox.showwarning("提示", "请先选择要导出的文档")
                return
        else:
            # 导出所有文档
            documents_data = []
            for item in self.tree_widget.get_children():
                values = self.tree_widget.item(item, 'values')
                if values and len(values) >= 5:
                    doc_info = {
                        'file_name': values[0],
                        'type': values[1],
                        'size': values[2],
                        'status': values[3],
                        'path': values[4]
                    }
                    documents_data.append(doc_info)
        
        if not documents_data:
            messagebox.showwarning("提示", "没有可导出的数据")
            return
        
        # 选择保存位置
        if file_format == "csv":
            default_ext = ".csv"
            file_types = [("CSV文件", "*.csv")]
        elif file_format == "xlsx":
            default_ext = ".xlsx"
            file_types = [("Excel文件", "*.xlsx")]
        else:
            default_ext = ".txt"
            file_types = [("文本文件", "*.txt")]
        
        file_types.append(("所有文件", "*.*"))
        
        file_path = filedialog.asksaveasfilename(
            title="保存文档列表",
            defaultextension=default_ext,
            filetypes=file_types
        )
        
        if not file_path:
            return
        
        try:
            save_path = Path(file_path)
            
            if save_path.suffix.lower() == '.csv' or file_format == "csv":
                self._export_to_csv(save_path, documents_data)
            elif save_path.suffix.lower() == '.txt' or file_format == "txt":
                self._export_to_text(save_path, documents_data)
            else:
                messagebox.showwarning("提示", "当前只支持导出为CSV和文本格式")
                return
            
            # 去掉导出成功提示窗口
            print(f"已导出 {len(documents_data)} 个文档信息到: {save_path}")
                
        except Exception as e:
            messagebox.showerror("导出失败", f"导出文档列表时出错: {e}")
    
    def _export_to_csv(self, file_path: Path, documents_data: List[Dict[str, Any]]):
        """导出为CSV格式"""
        with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
            if documents_data:
                # 重新映射字段名为中文
                fieldnames = ['文件名', '类型', '大小(MB)', '状态', '文件路径']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                
                # 重新映射数据
                for doc in documents_data:
                    mapped_doc = {
                        '文件名': doc['file_name'],
                        '类型': doc['type'],
                        '大小(MB)': doc['size'],
                        '状态': doc['status'],
                        '文件路径': doc['path']
                    }
                    writer.writerow(mapped_doc)
    
    def _export_to_text(self, file_path: Path, documents_data: List[Dict[str, Any]]):
        """导出为文本格式"""
        with open(file_path, 'w', encoding='utf-8') as txtfile:
            txtfile.write("文档列表导出报告\n")
            txtfile.write("=" * 50 + "\n\n")
            
            for i, doc in enumerate(documents_data, 1):
                txtfile.write(f"{i}. {doc['file_name']}\n")
                txtfile.write(f"   类型: {doc['type']}\n")
                txtfile.write(f"   大小: {doc['size']} MB\n")
                txtfile.write(f"   状态: {doc['status']}\n")
                txtfile.write(f"   路径: {doc['path']}\n")
                txtfile.write("\n")
    
    def clear_all_documents(self) -> bool:
        """清空所有文档"""
        if self.document_manager.document_count == 0:
            return False
        
        if messagebox.askyesno("确认", "确定要清空所有文档吗？"):
            self.document_manager.clear_all()
            return True
        
        return False 
    
    def filter_documents_by_enabled_types(self, enabled_file_types: dict) -> int:
        """
        根据启用的文件类型过滤文档列表
        
        Args:
            enabled_file_types: 启用的文件类型字典，如 {'word': True, 'pdf': False, ...}
            
        Returns:
            int: 被过滤掉的文档数量
        """
        if self.document_manager.document_count == 0:
            messagebox.showinfo("提示", "文档列表为空，无需过滤")
            return 0
        
        # 获取启用的文件类型列表
        enabled_types = [file_type for file_type, enabled in enabled_file_types.items() if enabled]
        
        if not enabled_types:
            messagebox.showwarning("提示", "请至少勾选一种文件类型")
            return 0
        
        # 统计过滤前的文档数量
        original_count = self.document_manager.document_count
        
        # 找出需要移除的文档
        documents_to_remove = []
        for doc in self.document_manager.documents:
            # 根据文档的文件类型判断是否需要移除
            doc_type = self._get_document_type_key(doc.type_display)
            if doc_type not in enabled_types:
                documents_to_remove.append(doc)
        
        if not documents_to_remove:
            messagebox.showinfo("提示", "没有需要过滤的文档")
            return 0
        
        # 确认过滤操作
        filter_count = len(documents_to_remove)
        if not messagebox.askyesno("确认过滤", 
                                 f"将过滤掉 {filter_count} 个未勾选类型的文档，确定继续吗？"):
            return 0
        
        # 执行过滤操作
        removed_count = 0
        for doc in documents_to_remove:
            if self.document_manager.remove_document(doc.id):
                removed_count += 1
        
        if removed_count > 0:
            enabled_type_names = [self._get_type_display_name(t) for t in enabled_types]
            print(f"文件过滤完成：保留了 {', '.join(enabled_type_names)} 类型文档，过滤掉 {removed_count} 个文档")
        
        return removed_count
    
    def _get_document_type_key(self, type_display: str) -> str:
        """
        根据文档显示类型获取对应的类型键
        
        Args:
            type_display: 文档显示类型，如 "Word文档"、"PDF文件" 等
            
        Returns:
            str: 对应的类型键，如 "word"、"pdf" 等
        """
        type_mapping = {
            "Word文档": "word",
            "PowerPoint": "ppt", 
            "Excel表格": "excel",
            "PDF文件": "pdf",
            "图片文件": "image",
            "文本文件": "text"
        }
        return type_mapping.get(type_display, "unknown")
    
    def _get_type_display_name(self, type_key: str) -> str:
        """
        根据类型键获取显示名称
        
        Args:
            type_key: 类型键，如 "word"、"pdf" 等
            
        Returns:
            str: 显示名称
        """
        display_mapping = {
            "word": "Word",
            "ppt": "PPT", 
            "excel": "Excel",
            "pdf": "PDF",
            "image": "图片",
            "text": "文本"
        }
        return display_mapping.get(type_key, type_key)