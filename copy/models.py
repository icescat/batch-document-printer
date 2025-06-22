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
    PDF = "pdf"


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
        if suffix in ['.doc', '.docx']:
            self.file_type = FileType.WORD
        elif suffix in ['.ppt', '.pptx']:
            self.file_type = FileType.PPT
        elif suffix == '.pdf':
            self.file_type = FileType.PDF
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
            FileType.PDF: "PDF文件"
        }
        return type_map.get(self.file_type, "未知")


@dataclass
class PrintSettings:
    """打印设置数据模型"""
    printer_name: str = ""
    paper_size: str = "A4"
    copies: int = 1
    duplex: bool = False
    color_mode: ColorMode = ColorMode.COLOR
    orientation: Orientation = Orientation.PORTRAIT
    
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
            duplex=data.get('duplex', False),
            color_mode=ColorMode(data.get('color_mode', 'color')),
            orientation=Orientation(data.get('orientation', 'portrait'))
        )


@dataclass
class AppConfig:
    """应用配置数据模型"""
    last_printer: str = ""
    default_settings: PrintSettings = field(default_factory=PrintSettings)
    window_geometry: dict = field(default_factory=dict)
    recent_folders: List[str] = field(default_factory=list)
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'last_printer': self.last_printer,
            'default_settings': self.default_settings.to_dict(),
            'window_geometry': self.window_geometry,
            'recent_folders': self.recent_folders
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'AppConfig':
        """从字典创建实例"""
        return cls(
            last_printer=data.get('last_printer', ''),
            default_settings=PrintSettings.from_dict(data.get('default_settings', {})),
            window_geometry=data.get('window_geometry', {}),
            recent_folders=data.get('recent_folders', [])
        ) 