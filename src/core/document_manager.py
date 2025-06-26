"""
文档管理器
负责文档的添加、删除、管理和验证功能
"""
import os
from pathlib import Path
from typing import List, Optional, Set, Dict, Callable
from .models import Document, FileType, PrintStatus


class DocumentManager:
    """文档管理器类"""
    
    # 支持的文件扩展名
    SUPPORTED_EXTENSIONS = {
        '.doc', '.docx', '.wps',    # Word文档（包含WPS文字）
        '.ppt', '.pptx', '.dps',    # PowerPoint（包含WPS演示）
        '.xls', '.xlsx', '.et',     # Excel表格（包含WPS表格）
        '.pdf',                     # PDF文件
        '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp',  # 图片文件（已移除.gif, .ico和.pcx）
        '.txt'                      # 文本文件
    }
    
    def __init__(self):
        """初始化文档管理器"""
        self._documents: List[Document] = []
        self._document_paths: Set[str] = set()  # 用于快速查重
        self._current_sort_key: Optional[str] = None  # 当前排序键
        self._current_sort_reverse: bool = False  # 当前排序方向
    
    @property
    def documents(self) -> List[Document]:
        """获取所有文档列表"""
        return self._documents.copy()
    
    @property
    def document_count(self) -> int:
        """获取文档总数"""
        return len(self._documents)
    
    @property
    def current_sort_info(self) -> tuple:
        """获取当前排序信息 (sort_key, reverse)"""
        return (self._current_sort_key, self._current_sort_reverse)
    
    def add_file(self, file_path: Path) -> Optional[Document]:
        """
        添加单个文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            成功添加的Document对象，失败返回None
        """
        try:
            # 验证文件
            if not self._validate_file(file_path):
                return None
            
            # 检查是否已存在
            file_path_str = str(file_path.resolve())
            if file_path_str in self._document_paths:
                print(f"文档已存在: {file_path.name}")
                return None
            
            # 创建文档对象
            document = Document(file_path=file_path)
            
            # 添加到列表
            self._documents.append(document)
            self._document_paths.add(file_path_str)
            
            # 如果当前有排序，重新应用排序
            if self._current_sort_key:
                self._apply_current_sort()
            
            print(f"成功添加文档: {document.file_name}")
            return document
            
        except Exception as e:
            print(f"添加文件失败 {file_path}: {e}")
            return None
    
    def add_files(self, file_paths: List[Path]) -> List[Document]:
        """
        批量添加文件
        
        Args:
            file_paths: 文件路径列表
            
        Returns:
            成功添加的Document对象列表
        """
        added_documents = []
        for file_path in file_paths:
            doc = self.add_file(file_path)
            if doc:
                added_documents.append(doc)
        
        print(f"批量添加完成: 成功 {len(added_documents)} 个，总计 {len(file_paths)} 个")
        return added_documents
    
    def add_folder(self, folder_path: Path, recursive: bool = True, enabled_file_types: Optional[Dict[str, bool]] = None) -> List[Document]:
        """
        添加文件夹中的所有支持的文档
        
        Args:
            folder_path: 文件夹路径
            recursive: 是否递归搜索子文件夹
            enabled_file_types: 启用的文件类型字典，如 {'word': True, 'ppt': True, 'excel': False, 'pdf': True}
            
        Returns:
            成功添加的Document对象列表
        """
        if not folder_path.exists() or not folder_path.is_dir():
            print(f"无效的文件夹路径: {folder_path}")
            return []
        
        # 如果没有提供文件类型过滤器，默认启用所有类型
        if enabled_file_types is None:
            enabled_file_types = {'word': True, 'ppt': True, 'excel': True, 'pdf': True, 'image': True, 'text': True}
        
        # 根据启用的文件类型构建允许的扩展名集合
        allowed_extensions = set()
        if enabled_file_types.get('word', False):
            allowed_extensions.update(['.doc', '.docx', '.wps'])
        if enabled_file_types.get('ppt', False):
            allowed_extensions.update(['.ppt', '.pptx', '.dps'])
        if enabled_file_types.get('excel', False):
            allowed_extensions.update(['.xls', '.xlsx', '.et'])
        if enabled_file_types.get('pdf', False):
            allowed_extensions.add('.pdf')
        if enabled_file_types.get('image', False):
            allowed_extensions.update(['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'])
        if enabled_file_types.get('text', False):
            allowed_extensions.add('.txt')
        
        # 获取文件列表
        pattern = "**/*" if recursive else "*"
        all_files = list(folder_path.glob(pattern))
        
        # 过滤支持的文件（根据启用的文件类型）
        supported_files = [
            f for f in all_files 
            if f.is_file() and f.suffix.lower() in allowed_extensions
        ]
        
        print(f"在文件夹 {folder_path.name} 中找到 {len(supported_files)} 个支持的文档")
        return self.add_files(supported_files)
    
    def remove_document(self, document_id: str) -> bool:
        """
        移除指定文档
        
        Args:
            document_id: 文档ID
            
        Returns:
            是否成功移除
        """
        for i, doc in enumerate(self._documents):
            if doc.id == document_id:
                # 从路径集合中移除
                file_path_str = str(doc.file_path.resolve())
                self._document_paths.discard(file_path_str)
                
                # 从文档列表中移除
                removed_doc = self._documents.pop(i)
                print(f"已移除文档: {removed_doc.file_name}")
                return True
        
        print(f"未找到文档ID: {document_id}")
        return False
    
    def clear_all(self):
        """清空所有文档"""
        count = len(self._documents)
        self._documents.clear()
        self._document_paths.clear()
        self._current_sort_key = None
        self._current_sort_reverse = False
        print(f"已清空所有文档 ({count} 个)")
    
    def get_document_by_id(self, document_id: str) -> Optional[Document]:
        """
        根据ID获取文档
        
        Args:
            document_id: 文档ID
            
        Returns:
            Document对象或None
        """
        for doc in self._documents:
            if doc.id == document_id:
                return doc
        return None
    
    def get_documents_by_status(self, status: PrintStatus) -> List[Document]:
        """
        根据打印状态获取文档列表
        
        Args:
            status: 打印状态
            
        Returns:
            符合条件的文档列表
        """
        return [doc for doc in self._documents if doc.print_status == status]
    
    def get_documents_by_type(self, file_type: FileType) -> List[Document]:
        """
        根据文件类型获取文档列表
        
        Args:
            file_type: 文件类型
            
        Returns:
            符合条件的文档列表
        """
        return [doc for doc in self._documents if doc.file_type == file_type]
    
    def update_document_status(self, document_id: str, status: PrintStatus) -> bool:
        """
        更新文档打印状态
        
        Args:
            document_id: 文档ID
            status: 新的打印状态
            
        Returns:
            是否更新成功
        """
        doc = self.get_document_by_id(document_id)
        if doc:
            old_status = doc.print_status
            doc.print_status = status
            print(f"文档 {doc.file_name} 状态更新: {old_status.value} -> {status.value}")
            return True
        return False
    
    def sort_documents(self, sort_key: str, reverse: bool = False) -> None:
        """
        对文档列表进行排序
        
        Args:
            sort_key: 排序键 ('name', 'type', 'size', 'status', 'path', 'added_time')
            reverse: 是否倒序排列
        """
        if not self._documents:
            return
        
        # 定义排序键函数
        sort_functions = {
            'name': lambda doc: doc.file_name.lower(),  # 文件名不区分大小写排序
            'type': lambda doc: (doc.file_type.value, doc.file_name.lower()),  # 先按类型，再按文件名
            'size': lambda doc: (doc.file_size, doc.file_name.lower()),  # 先按大小，再按文件名
            'status': lambda doc: (doc.print_status.value, doc.file_name.lower()),  # 先按状态，再按文件名
            'path': lambda doc: str(doc.file_path).lower(),  # 路径不区分大小写排序
            'added_time': lambda doc: (doc.added_time, doc.file_name.lower())  # 先按添加时间，再按文件名
        }
        
        if sort_key not in sort_functions:
            print(f"不支持的排序键: {sort_key}")
            return
        
        try:
            # 执行排序
            self._documents.sort(key=sort_functions[sort_key], reverse=reverse)
            
            # 保存当前排序状态
            self._current_sort_key = sort_key
            self._current_sort_reverse = reverse
            
            print(f"文档已按 {sort_key} {'降序' if reverse else '升序'} 排列")
            
        except Exception as e:
            print(f"排序失败: {e}")
    
    def toggle_sort(self, sort_key: str) -> bool:
        """
        切换排序状态：如果当前是该字段升序，则改为降序；否则设为升序
        
        Args:
            sort_key: 排序键
            
        Returns:
            返回当前是否为降序
        """
        if self._current_sort_key == sort_key:
            # 如果是同一个字段，切换排序方向
            reverse = not self._current_sort_reverse
        else:
            # 如果是不同字段，默认从升序开始
            reverse = False
        
        self.sort_documents(sort_key, reverse)
        return reverse
    
    def _apply_current_sort(self):
        """应用当前的排序设置"""
        if self._current_sort_key:
            self.sort_documents(self._current_sort_key, self._current_sort_reverse)
    
    def get_summary(self) -> dict:
        """
        获取文档统计摘要
        
        Returns:
            包含各种统计信息的字典
        """
        if not self._documents:
            return {
                'total': 0,
                'by_type': {},
                'by_status': {},
                'total_size_mb': 0
            }
        
        # 按类型统计
        by_type = {}
        for file_type in FileType:
            count = len(self.get_documents_by_type(file_type))
            if count > 0:
                by_type[file_type.value] = count
        
        # 按状态统计
        by_status = {}
        for status in PrintStatus:
            count = len(self.get_documents_by_status(status))
            if count > 0:
                by_status[status.value] = count
        
        # 总大小
        total_size = sum(doc.file_size for doc in self._documents)
        total_size_mb = round(total_size / (1024 * 1024), 2)
        
        return {
            'total': self.document_count,
            'by_type': by_type,
            'by_status': by_status,
            'total_size_mb': total_size_mb
        }
    
    def _validate_file(self, file_path: Path) -> bool:
        """
        验证文件是否有效
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否有效
        """
        # 检查文件是否存在
        if not file_path.exists():
            print(f"文件不存在: {file_path}")
            return False
        
        # 检查是否为文件
        if not file_path.is_file():
            print(f"不是文件: {file_path}")
            return False
        
        # 过滤临时文件和隐藏文件
        if self._is_temp_or_hidden_file(file_path):
            print(f"跳过临时/隐藏文件: {file_path.name}")
            return False
        
        # 检查文件扩展名
        if file_path.suffix.lower() not in self.SUPPORTED_EXTENSIONS:
            print(f"不支持的文件类型: {file_path.suffix}")
            return False
        
        # 检查文件大小（避免过大的文件）
        try:
            file_size = file_path.stat().st_size
            max_size = 100 * 1024 * 1024  # 100MB限制
            if file_size > max_size:
                print(f"文件过大 (>{max_size/(1024*1024):.1f}MB): {file_path.name}")
                return False
        except OSError as e:
            print(f"无法读取文件信息: {file_path} - {e}")
            return False
        
        return True
    
    def _is_temp_or_hidden_file(self, file_path: Path) -> bool:
        """
        检查是否为临时文件或隐藏文件
        
        Args:
            file_path: 文件路径
            
        Returns:
            是否为临时文件或隐藏文件
        """
        file_name = file_path.name
        
        # 1. 检查隐藏文件（以点开头）
        if file_name.startswith('.'):
            return True
        
        # 2. 检查Office临时文件
        # Word临时文件：~$开头
        if file_name.startswith('~$'):
            return True
        
        # Excel临时文件：~$开头或者.tmp结尾
        if file_name.startswith('~$') or file_name.lower().endswith('.tmp'):
            return True
        
        # PowerPoint特殊临时文件：pptE或pptF开头的临时文件
        if file_name.startswith('pptE') or file_name.startswith('pptF'):
            return True
        
        # 3. 检查Windows系统临时文件
        temp_patterns = [
            '~',          # 以~开头的文件
            'Thumbs.db',  # Windows缩略图文件
            'desktop.ini', # Windows桌面配置文件
            '.DS_Store',  # macOS系统文件
        ]
        
        for pattern in temp_patterns:
            if file_name.startswith(pattern) or file_name == pattern:
                return True
        
        # 4. 检查Office自动恢复文件
        recovery_patterns = [
            'AutoRecovery save of',  # Word自动恢复
            '自动恢复的',              # 中文自动恢复
        ]
        
        for pattern in recovery_patterns:
            if pattern in file_name:
                return True
        
        # 5. 检查备份文件
        backup_patterns = [
            '.bak',    # 备份文件
            '.backup', # 备份文件
            '副本',    # 中文副本文件
            ' - 副本', # 中文副本文件
        ]
        
        for pattern in backup_patterns:
            if file_name.lower().endswith(pattern.lower()) or pattern in file_name:
                return True
        
        # 6. 使用Windows API检查隐藏属性（如果可用）
        try:
            import os
            import stat
            
            # 检查文件属性
            file_stat = file_path.stat()
            
            # 在Windows上检查隐藏属性
            if os.name == 'nt':
                try:
                    import win32api
                    import win32con
                    
                    attrs = win32api.GetFileAttributes(str(file_path))
                    if attrs & win32con.FILE_ATTRIBUTE_HIDDEN:
                        return True
                    if attrs & win32con.FILE_ATTRIBUTE_SYSTEM:
                        return True
                    if attrs & win32con.FILE_ATTRIBUTE_TEMPORARY:
                        return True
                except ImportError:
                    # win32api不可用，跳过Windows特定检查
                    pass
                except Exception:
                    # 其他错误，跳过
                    pass
            
        except Exception:
            # 如果无法获取文件属性，不影响判断
            pass
        
        return False 