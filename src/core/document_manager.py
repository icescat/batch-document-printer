"""
文档管理器
负责文档的添加、删除、管理和验证功能
"""
import os
from pathlib import Path
from typing import List, Optional, Set
from .models import Document, FileType, PrintStatus


class DocumentManager:
    """文档管理器类"""
    
    # 支持的文件扩展名
    SUPPORTED_EXTENSIONS = {
        '.doc', '.docx',    # Word文档
        '.ppt', '.pptx',    # PowerPoint
        '.pdf'              # PDF文件
    }
    
    def __init__(self):
        """初始化文档管理器"""
        self._documents: List[Document] = []
        self._document_paths: Set[str] = set()  # 用于快速查重
    
    @property
    def documents(self) -> List[Document]:
        """获取所有文档列表"""
        return self._documents.copy()
    
    @property
    def document_count(self) -> int:
        """获取文档总数"""
        return len(self._documents)
    
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
    
    def add_folder(self, folder_path: Path, recursive: bool = True) -> List[Document]:
        """
        添加文件夹中的所有支持的文档
        
        Args:
            folder_path: 文件夹路径
            recursive: 是否递归搜索子文件夹
            
        Returns:
            成功添加的Document对象列表
        """
        if not folder_path.exists() or not folder_path.is_dir():
            print(f"无效的文件夹路径: {folder_path}")
            return []
        
        # 获取文件列表
        pattern = "**/*" if recursive else "*"
        all_files = list(folder_path.glob(pattern))
        
        # 过滤支持的文件
        supported_files = [
            f for f in all_files 
            if f.is_file() and f.suffix.lower() in self.SUPPORTED_EXTENSIONS
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