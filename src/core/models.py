"""
数据模型定义
定义了应用中使用的所有数据结构
"""
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Optional, List


class FileType(Enum):
    """文件类型枚举"""
    WORD = "word"
    PPT = "ppt"
    EXCEL = "excel"
    PDF = "pdf"
    IMAGE = "image"
    TEXT = "text"


class PrintStatus(Enum):
    """打印状态枚举"""
    PENDING = "pending"      # 待打印
    PRINTING = "printing"    # 打印中
    COMPLETED = "completed"  # 已完成
    ERROR = "error"         # 错误


class ColorMode(Enum):
    """颜色模式枚举"""
    COLOR = "color"         # 彩色
    GRAYSCALE = "grayscale" # 黑白


class Orientation(Enum):
    """页面方向枚举"""
    PORTRAIT = "portrait"   # 纵向
    LANDSCAPE = "landscape" # 横向


@dataclass
class Document:
    """文档数据模型"""
    file_path: Path
    file_name: str = field(init=False)
    file_type: FileType = field(init=False)
    file_size: int = field(init=False)
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    added_time: datetime = field(default_factory=datetime.now)
    print_status: PrintStatus = PrintStatus.PENDING
    
    def __post_init__(self):
        """初始化后处理"""
        self.file_name = self.file_path.name
        self.file_size = self.file_path.stat().st_size if self.file_path.exists() else 0
        
        # 根据文件扩展名确定文件类型
        suffix = self.file_path.suffix.lower()
        if suffix in ['.doc', '.docx', '.wps']:  # 添加WPS文字格式
            self.file_type = FileType.WORD
        elif suffix in ['.ppt', '.pptx', '.dps']:  # 添加WPS演示格式
            self.file_type = FileType.PPT
        elif suffix in ['.xls', '.xlsx', '.et']:  # 添加WPS表格格式
            self.file_type = FileType.EXCEL
        elif suffix == '.pdf':
            self.file_type = FileType.PDF
        elif suffix in ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp']:
            self.file_type = FileType.IMAGE
        elif suffix == '.txt':
            self.file_type = FileType.TEXT
        else:
            raise ValueError(f"不支持的文件类型: {suffix}")
    
    @property
    def size_mb(self) -> float:
        """获取文件大小（MB）"""
        return round(self.file_size / (1024 * 1024), 2)
    
    @property
    def type_display(self) -> str:
        """获取文件类型显示名称"""
        type_map = {
            FileType.WORD: "Word文档",
            FileType.PPT: "PowerPoint",
            FileType.EXCEL: "Excel表格",
            FileType.PDF: "PDF文件",
            FileType.IMAGE: "图片文件",
            FileType.TEXT: "文本文件"
        }
        return type_map.get(self.file_type, "未知")


@dataclass
class PrintSettings:
    """打印设置数据模型"""
    printer_name: str = ""
    paper_size: str = "A4"
    copies: int = 1
    duplex: bool = True  # 默认启用双面打印
    color_mode: ColorMode = ColorMode.GRAYSCALE  # 默认黑白打印
    orientation: Orientation = Orientation.PORTRAIT  # 默认竖向
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'printer_name': self.printer_name,
            'paper_size': self.paper_size,
            'copies': self.copies,
            'duplex': self.duplex,
            'color_mode': self.color_mode.value,
            'orientation': self.orientation.value
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'PrintSettings':
        """从字典创建实例"""
        return cls(
            printer_name=data.get('printer_name', ''),
            paper_size=data.get('paper_size', 'A4'),
            copies=data.get('copies', 1),
            duplex=data.get('duplex', True),  # 默认启用双面打印
            color_mode=ColorMode(data.get('color_mode', 'grayscale')),  # 默认黑白
            orientation=Orientation(data.get('orientation', 'portrait'))
        )


@dataclass
class AppConfig:
    """应用配置数据模型"""
    last_printer: str = ""
    default_settings: PrintSettings = field(default_factory=PrintSettings)
    window_geometry: dict = field(default_factory=dict)
    recent_folders: List[str] = field(default_factory=list)
    # 文件类型过滤器设置 - 默认启用Word、PPT、PDF、图片、文本，不启用Excel
    enabled_file_types: dict = field(default_factory=lambda: {
        'word': True,
        'ppt': True,
        'excel': False,
        'pdf': True,
        'image': True,
        'text': True
    })
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'last_printer': self.last_printer,
            'default_settings': self.default_settings.to_dict(),
            'window_geometry': self.window_geometry,
            'recent_folders': self.recent_folders,
            'enabled_file_types': self.enabled_file_types
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'AppConfig':
        """从字典创建实例"""
        # 默认的文件类型启用状态
        default_enabled_types = {
            'word': True,
            'ppt': True,
            'excel': False,
            'pdf': True,
            'image': True,
            'text': True
        }
        enabled_types = data.get('enabled_file_types', default_enabled_types)
        # 确保所有文件类型都有设置
        for file_type in default_enabled_types:
            if file_type not in enabled_types:
                enabled_types[file_type] = default_enabled_types[file_type]
        
        return cls(
            last_printer=data.get('last_printer', ''),
            default_settings=PrintSettings.from_dict(data.get('default_settings', {})),
            window_geometry=data.get('window_geometry', {}),
            recent_folders=data.get('recent_folders', []),
            enabled_file_types=enabled_types
        ) 