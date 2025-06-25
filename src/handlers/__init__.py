"""
文档处理器模块
提供基于文件格式的模块化文档处理能力
"""

try:
    from .base_handler import BaseDocumentHandler
    from .handler_registry import HandlerRegistry
    from .pdf_handler import PDFDocumentHandler
    from .word_handler import WordDocumentHandler
    from .powerpoint_handler import PowerPointDocumentHandler
    from .excel_handler import ExcelDocumentHandler
    from .image_handler import ImageDocumentHandler
    from .text_handler import TextDocumentHandler
    
    __all__ = [
        'BaseDocumentHandler',
        'HandlerRegistry',
        'PDFDocumentHandler',
        'WordDocumentHandler',
        'PowerPointDocumentHandler',
        'ExcelDocumentHandler',
        'ImageDocumentHandler',
        'TextDocumentHandler'
    ]
except ImportError as e:
    print(f"警告: 导入处理器模块时出错: {e}")
    __all__ = [] 