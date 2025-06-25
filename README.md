# 办公文档批量打印器

> 🚀 **最新版本**: v5.0.0 | 📥 **[立即下载](https://github.com/icescat/batch-document-printer/releases/latest)** | 🌟 **[GitHub 仓库](https://github.com/icescat/batch-document-printer)**

## 项目概括

本项目基于Python 3.12，采用**策略模式+注册器模式的模块化处理器架构**，支持批量打印多种文档格式：Word、PowerPoint、Excel、PDF、图片和文本文件。提供直观的图形用户界面，支持拖拽导入、智能文件类型过滤、精确页数统计、WPS格式兼容、灵活打印设置等功能，能显著提升办公文档打印效率。

## 🏗️ 架构设计

### 核心架构 - 模块化处理器模式

### 🔧 技术选型

- **主要编程语言**: Python 3.12+
- **GUI框架**: tkinter, tkinterdnd2 (拖拽支持)
- **架构模式**: 策略模式 + 注册器模式 + 插件化架构
- **文档处理库**:
  - python-docx (Word文档处理)
  - python-pptx (PowerPoint处理)
  - xlwings (Excel处理，精确页数统计)
  - PyPDF2/pypdf (PDF处理)
  - comtypes/pywin32 (Windows COM接口，用于调用Office应用)
- **打印控制**: win32print, win32api (Windows打印API)
- **应用打包**: PyInstaller (生成独立exe文件)
- **文件操作**: pathlib, os (文件系统操作)
- **版本控制**: Git
- **其他工具**:
  - pytest (单元测试)
  - black (代码格式化)
  - cx_Freeze 或 auto-py-to-exe (备选打包方案)

## 📁 项目结构 / 模块划分

### v5.0 完整架构目录结构

```text
printer/
├── src/
│   ├── handlers/                    # 🆕 文档处理器模块 (核心架构)
│   │   ├── __init__.py              # 处理器导出
│   │   ├── base_handler.py          # 基础处理器接口
│   │   ├── handler_registry.py      # 处理器注册中心
│   │   ├── pdf_handler.py           # PDF文档处理器
│   │   ├── word_handler.py          # Word文档处理器
│   │   ├── powerpoint_handler.py    # PowerPoint文档处理器
│   │   ├── excel_handler.py         # Excel文档处理器
│   │   ├── image_handler.py         # 图片文档处理器
│   │   ├── text_handler.py          # 文本文档处理器
│   │   └── print_utils.py           # 打印工具类
│   ├── gui/                         # 图形用户界面模块
│   │   ├── components/              # GUI功能组件
│   │   │   ├── __init__.py          # 组件导出
│   │   │   ├── file_import_handler.py    # 文件导入功能处理器
│   │   │   ├── list_operation_handler.py # 列表操作功能处理器
│   │   │   ├── window_manager.py         # 窗口管理器
│   │   │   └── tooltip.py                # 浮窗提示组件
│   │   ├── main_window.py           # 主窗口界面 (纯界面层)
│   │   ├── print_settings_dialog.py # 打印设置对话框
│   │   └── page_count_dialog.py     # 页数统计对话框
│   ├── core/                        # 核心业务逻辑模块 (重构)
│   │   ├── document_manager.py      # 文档管理器
│   │   ├── print_controller.py      # 基于处理器架构
│   │   ├── page_count_manager.py    # 基于处理器架构
│   │   ├── settings_manager.py      # 设置管理器
│   │   ├── printer_config_manager.py # 打印机配置管理器
│   │   └── models.py                # 数据模型定义
│   └── utils/                       # 工具类和辅助函数
│       └── config_utils.py          # 配置文件操作
├── data/                            # 应用数据存储目录
├── resources/                       # 资源文件目录
│   └── app_icon.ico                 # 应用程序图标
├── main.py                         # 程序入口点
├── requirements.txt                # Python依赖管理
├── build_v5.0.bat                  # v5.0构建脚本
├── 办公文档批量打印器v5.0.spec      # PyInstaller打包配置
└── README.md                       # 项目文档
```

### 架构分层说明

```text
📋 v5.0 完整分层架构
┌─────────────────────────────────────────────────────────┐
│                    应用入口层                            │
│  main.py - 程序启动入口和系统初始化                     │
├─────────────────────────────────────────────────────────┤
│                    GUI界面层                            │
│  主界面   │  设置对话框  │  页数统计对话框  │  帮助界面    │
├─────────────────────────────────────────────────────────┤
│                 GUI功能处理器层 🆕                       │
│ FileImportHandler │ ListOperationHandler │ WindowManager │
├─────────────────────────────────────────────────────────┤
│                   核心业务逻辑层                         │
│ DocumentManager │ PrintController │ SettingsManager     │
├─────────────────────────────────────────────────────────┤
│                 文档处理器层 🆕                          │
│ WordHandler │ PDFHandler │ PPTHandler │ ExcelHandler │ ImageHandler │ TextHandler │
├─────────────────────────────────────────────────────────┤
│                   基础工具层                            │
│  ConfigUtils  │  HandlerRegistry  │  Models  │  Types  │
└─────────────────────────────────────────────────────────┘
```

## 🚀 核心功能 / 模块详解

### 🔧 文档处理器架构 (v5.0 新特性)

#### 基础处理器接口 (`base_handler.py`)

定义所有文档处理器必须实现的统一接口：

```python
class BaseDocumentHandler(ABC):
    @abstractmethod
    def get_supported_file_types(self) -> Set[FileType]
    def get_supported_extensions(self) -> Set[str]
    def can_handle_file(self, file_path: Path) -> bool
    def print_document(self, file_path: Path, settings: PrintSettings) -> bool
    def count_pages(self, file_path: Path) -> int
```

#### 处理器注册中心 (`handler_registry.py`)

负责管理和分发各种文档格式的处理器：

- 动态处理器注册和发现
- 基于文件类型和扩展名的智能路由
- 支持运行时添加新的处理器
- 统一的错误处理和日志记录

#### 具体文档处理器

| 处理器 | 支持格式 | 功能 |
|--------|----------|------|
| **PDFDocumentHandler** | `.pdf` | PDF打印 + 页数统计 |
| **WordDocumentHandler** | `.doc`, `.docx`, `.wps` | Word打印 + 页数统计，支持WPS文字 |
| **PowerPointDocumentHandler** | `.ppt`, `.pptx`, `.dps` | PPT打印 + 幻灯片统计，支持WPS演示 |
| **ExcelDocumentHandler** | `.xls`, `.xlsx`, `.et` | Excel打印 + 工作表统计，支持WPS表格 |
| **ImageDocumentHandler** | `.jpg`, `.png`, `.bmp`, `.tiff`, `.webp` | 图片打印 + 页数统计（TIFF多页支持） |
| **TextDocumentHandler** | `.txt` | 文本打印 + 页数估算，智能编码检测 |

### 🎨 GUI功能处理器架构 (v5.0 新特性)

#### 功能处理器详解

#### 1. FileImportHandler (文件导入处理器)

- **拖拽导入**: 支持多种格式的拖拽数据解析
- **文件选择**: 文件/文件夹选择对话框
- **格式过滤**: 根据用户设置过滤文件类型
- **路径解析**: 智能处理复杂文件路径

#### 2. ListOperationHandler (列表操作处理器)

- **智能排序**: 列表头点击排序，支持多字段排序
- **文件操作**: 删除、清空、打开文件位置
- **右键菜单**: 上下文相关的操作菜单
- **列表导出**: CSV/文本格式的文档列表导出

#### 3. WindowManager (窗口管理器)

- **几何状态**: 窗口位置和大小的保存/恢复
- **用户偏好**: 文件类型过滤器等设置持久化
- **窗口行为**: 居中显示、最小尺寸、关闭处理
- **配置管理**: 与ConfigManager的集成

### 4. 重构后的打印控制器 (`print_controller.py`)

- **基于处理器架构**: 不再包含具体的格式处理逻辑
- **智能路由**: 根据文件类型自动选择合适的处理器
- **统一接口**: 所有文档类型使用相同的打印接口
- **增强错误处理**: 处理器级别的错误隔离
- **并行优化**: 支持多个处理器并行工作

### 5. 重构后的页数统计管理器 (`page_count_manager.py`)

- **模块化统计**: 每种格式由专门的处理器负责
- **并行计算**: 多线程并发处理大批量文档
- **统一结果格式**: 所有处理器返回标准化的统计结果
- **缓存机制**: 避免重复计算相同文档的页数
- **进度追踪**: 大批量统计时提供实时进度更新

### 6. 核心业务模块

- **文档管理器** (`document_manager.py`): 支持多种添加方式、文件格式过滤验证、文档列表管理与预览
- **设置管理器** (`settings_manager.py`): 打印机选择、纸张设置、打印参数配置、用户偏好持久化
- **配置工具** (`config_utils.py`): 应用程序配置文件操作、用户偏好设置、窗口状态保存

## 🔄 架构演进历程

### v5.0新特性亮点总结

#### 1. **完全模块化的文档处理器架构**

- 支持6种文档格式：PDF、Word、PowerPoint、Excel、图片、文本文件
- 统一的BaseDocumentHandler接口，确保所有处理器的一致性
- 智能注册中心HandlerRegistry，支持运行时动态加载处理器
- **WPS兼容性**：完整支持.wps、.dps、.et格式

#### 2. **GUI架构的彻底重构**

- 实现了界面与业务逻辑的完全分离

#### 3. **增强的扩展性**

- 添加新文档格式只需实现新的处理器类，无需修改现有代码
- 插件式架构支持第三方扩展
- 配置驱动的功能开关

#### 4. **更好的错误处理和性能**

- 处理器级别的错误隔离，单个格式的问题不影响其他格式
- 并行处理能力，支持多文档同时处理
- 详细的错误信息和建议

#### 5. **新增文件格式支持** 🆕

- **图片文件支持**：.jpg, .jpeg, .png, .bmp, .tiff, .tif, .webp
- **文本文件支持**：.txt，智能编码检测（UTF-8, GBK, GB2312等）
- **WPS格式支持**：.wps, .dps, .et，通过复用Office处理器实现

#### 6. **用户体验优化**

- 智能的文件类型识别和过滤
- 流畅的拖拽导入体验
- 详细的处理进度显示
- **提示系统**：仅在文件类型过滤器显示tooltip提示

#### 7. **错误隔离**

- 单个处理器的错误不会影响其他处理器
- 更好的错误定位和调试能力
- 提供详细的错误信息和建议

#### 8. **GUI模块化重构优势** 🆕

- **职责分离**: 界面与功能逻辑彻底分离
- **代码可读性**: 从1357行巨大文件拆分为4个清晰模块
- **维护便利**: 修改功能无需接触界面代码
- **测试友好**: 功能处理器可独立进行单元测试
- **复用性强**: 功能处理器可在其他界面中复用

## 🛠️ 安装与使用

### 环境要求

- **操作系统**: Windows 10/11 (64位)
- **Python版本**: Python 3.8+ (推荐3.12+)
- **Office软件**: Microsoft Office 2016+ (用于.docx/.pptx/.xlsx文件处理)
- **系统权限**: 需要打印机访问权限

### 快速安装

#### 方式一: 直接下载可执行文件 (推荐)

```bash
# 1. 下载最新版本
https://github.com/icescat/batch-document-printer/releases/latest

# 2. 解压并运行
办公文档批量打印器v5.0.exe
```

#### 方式二: 从源码安装

```bash
# 1. 克隆仓库
git clone https://github.com/icescat/batch-document-printer.git
cd batch-document-printer

# 2. 创建虚拟环境
python -m venv .venv
.venv\Scripts\activate

# 3. 安装依赖
pip install -r requirements.txt

# 4. 运行程序
python main.py
```

#### 方式三: 使用pip安装

```bash
pip install batch-document-printer
batch-printer
```

## 📊 技术实现细节

### 文档处理器实现原理

#### PDF处理器 (`pdf_handler.py`)

```python
class PDFDocumentHandler(BaseDocumentHandler):
    def count_pages(self, file_path: Path) -> int:
        """使用PyPDF2库解析PDF页数"""
        
    def print_document(self, file_path: Path, settings: PrintSettings) -> bool:
        """通过Adobe Reader或Windows默认程序打印"""
```

#### Word处理器 (`word_handler.py`)

```python
class WordDocumentHandler(BaseDocumentHandler):
    def count_pages(self, file_path: Path) -> int:
        """通过python-docx或COM接口获取页数"""
        
    def print_document(self, file_path: Path, settings: PrintSettings) -> bool:
        """使用Microsoft Word COM接口执行打印"""
```

#### PowerPoint处理器 (`powerpoint_handler.py`)

```python
class PowerPointDocumentHandler(BaseDocumentHandler):
    def count_pages(self, file_path: Path) -> int:
        """通过python-pptx或COM接口获取幻灯片数"""
        
    def print_document(self, file_path: Path, settings: PrintSettings) -> bool:
        """使用Microsoft PowerPoint COM接口执行打印"""
```

#### Excel处理器 (`excel_handler.py`)

```python
class ExcelDocumentHandler(BaseDocumentHandler):
    def count_pages(self, file_path: Path) -> int:
        """
        使用xlwings获取Excel精确打印页数:
        1. 通过Excel原生API访问HPageBreaks和VPageBreaks
        2. 计算实际打印页数: (水平分页符+1) × (垂直分页符+1)
        3. 遍历所有工作表累加总页数
        4. 确保每个工作表至少统计1页
        """
        
    def print_document(self, file_path: Path, settings: PrintSettings) -> bool:
        """使用Microsoft Excel COM接口执行打印"""
```

#### 图片处理器 (`image_handler.py`) 🆕

```python
class ImageDocumentHandler(BaseDocumentHandler):
    def count_pages(self, file_path: Path) -> int:
        """
        智能图片页数统计:
        - 一般图片格式(jpg, png, bmp, webp): 1页
        - TIFF格式: 使用PIL检测实际页数(支持多页TIFF)
        """
        
    def print_document(self, file_path: Path, settings: PrintSettings) -> bool:
        """通过Windows关联程序或系统打印服务打印图片"""
```

#### 文本处理器 (`text_handler.py`) 🆕

```python
class TextDocumentHandler(BaseDocumentHandler):
    def count_pages(self, file_path: Path) -> int:
        """
        智能文本页数估算:
        1. 多编码检测: UTF-8, GBK, GB2312, ANSI等
        2. 行数统计: 分析换行符和自动换行
        3. 页数估算: 基于标准A4纸张尺寸和字体大小
        4. 文件大小限制: 最大支持100MB文本文件
        """
        
    def print_document(self, file_path: Path, settings: PrintSettings) -> bool:
        """使用notepad /p命令打印文本文件"""
```

**Excel页数统计技术细节:**

- **xlwings集成**: 利用xlwings库直接访问Excel的原生分页符API，获得真实的打印页数
- **分页符分析**: 统计水平分页符(HPageBreaks)和垂直分页符(VPageBreaks)，精确计算页数
- **工作表遍历**: 逐个处理每个工作表，累加总页数
- **错误处理**: 单个工作表失败时自动补偿，确保统计结果的可靠性

### 核心算法

#### 1. 智能文件类型识别

```python
def detect_file_type(file_path: Path) -> FileType:
    """
    多层次文件类型识别:
    1. 文件扩展名检查
    2. MIME类型检测  
    3. 文件头魔数验证
    """
```

#### 2. 并行处理调度

```python
class ParallelProcessor:
    """
    基于线程池的并行文档处理:
    - 动态调整工作线程数量
    - 智能任务分配和负载均衡
    - 实时进度监控和异常处理
    """
```

#### 3. 错误恢复机制

```python
class ErrorRecoveryManager:
    """
    多级错误恢复策略:
    - 自动重试机制
    - 降级处理方案
    - 详细错误日志记录
    """
```

## 🧪 开发与测试

### 开发环境搭建

```bash
# 1. 克隆开发分支
git clone -b develop https://github.com/icescat/batch-document-printer.git

# 2. 安装开发依赖
pip install -r requirements-dev.txt

# 3. 安装pre-commit钩子
pre-commit install
```

### 运行测试

```bash
# 运行所有测试
pytest

# 运行特定模块测试
pytest tests/test_handlers/

# 生成测试覆盖率报告
pytest --cov=src --cov-report=html
```

### 代码质量检查

```bash
# 代码格式化
black src/

# 静态类型检查  
mypy src/

# 代码风格检查
flake8 src/
```

## 📦 构建与分发

### 构建可执行文件

```bash
# 使用提供的构建脚本
.\build_v5.0.bat

# 或手动使用PyInstaller
pyinstaller 办公文档批量打印器v5.0.spec
```

### 项目打包

```bash
# 创建分发包
python setup.py sdist bdist_wheel

# 上传到PyPI
twine upload dist/*
```

## 📄 许可证

本项目采用 [MIT许可证](LICENSE) - 查看LICENSE文件了解详情。

## 📞 支持与反馈

- **GitHub**: [项目仓库](https://github.com/icescat/batch-document-printer)
- **Issues**: [问题追踪](https://github.com/icescat/batch-document-printer/issues)
- **文档**: [在线文档](https://github.com/icescat/batch-document-printer/wiki)

## 📅 版本更新历史

### 🚀 v5.0 (2025-06-25) - 架构重构版本

#### 新增功能
- ✅ **多格式支持扩展**: 新增图片文件(.jpg, .jpeg, .png, .bmp, .tiff, .tif, .webp)和文本文件(.txt)支持
- ✅ **WPS兼容性**: 完整支持WPS格式(.wps, .dps, .et)，复用Office处理器
- ✅ **智能提示系统**: 优化的tooltip提示，仅在必要位置显示帮助信息
- ✅ **增强编码支持**: 文本文件智能编码检测(UTF-8, GBK, GB2312等)

#### 架构重构
- 🔧 **策略模式+注册器**: 全新的模块化处理器架构，每种文件格式独立处理器
- 🔧 **GUI组件分离**: GUI功能模块化，提升代码可维护性
- 🔧 **并发优化**: 改进的多线程页数统计，提升处理效率

#### 技术升级
- 📈 支持Python 3.12+，向下兼容3.8+
- 📈 改进内存管理，减少大文件处理时的内存占用
- 📈 更快的启动速度和更稳定的COM接口调用

---

### 🔄 v4.1 (2025-06-23) - 拖拽修复版本

#### 问题修复
- 🐛 **拖拽导入修复**: 解决文件名包含空格时的导入失败问题
- 🐛 **路径解析优化**: 改进含特殊字符路径的处理逻辑

---

### 📁 v4.0 (2025-06-23) - 拖拽导入版本

#### 新增功能  
- ✅ **拖拽导入功能**: 支持直接拖拽文件和文件夹到程序窗口
- ✅ **递归文件夹扫描**: 自动搜索子文件夹中的支持文档
- ✅ **智能路径解析**: 自动处理文件路径格式和特殊字符

#### 用户体验改进
- 🎨 更直观的文件添加方式
- 🎨 实时的文件类型过滤器应用
- 🎨 改进的操作反馈提示

---

### 📊 v3.0 (2025-06-22) - 页数统计版本

#### 重大功能新增
- ✅ **页数统计功能**: 全新的文档页数/张数统计功能
- ✅ **统计结果展示**: 专用的页数统计对话框，提供详细结果
- ✅ **批量统计**: 支持多文档同时统计，大幅提升效率
- ✅ **成本预估**: 为用户提供打印页数参考，便于成本评估

#### 技术实现
- 🔧 优化数据模型，增加页数相关字段
- 🔧 完善界面交互，增强用户体验

---

### 📋 v2.0 (2025-06-22) - Excel支持版本

#### 核心功能扩展
- ✅ **Excel文件支持**: 新增.xls和.xlsx格式支持
- ✅ **文件类型过滤器**: 独立的Word、PPT、Excel、PDF过滤开关
- ✅ **界面优化**: 重新设计过滤器布局，提升使用便利性

#### 功能改进
- 🎨 文件类型可视化选择
- 🎨 改进的文件添加逻辑
- 🎨 优化的界面布局和交互

---

### 🎯 v1.0 (2025-06-22) - 首发版本

#### 基础功能实现
- ✅ **核心打印功能**: 支持Word(.doc/.docx)、PowerPoint(.ppt/.pptx)、PDF文件的批量打印
- ✅ **图形化界面**: 基于tkinter的直观用户界面
- ✅ **打印设置**: 完整的打印参数配置(打印机、纸张、份数等)
- ✅ **文档管理**: 支持添加、删除、清空文档列表

#### 技术基础
- 🔧 基于Python + tkinter的跨平台桌面应用
- 🔧 COM接口调用Microsoft Office组件
- 🔧 PyPDF2处理PDF文档
- 🔧 模块化的代码架构设计

---
⭐ 如果这个项目对您有帮助，请给个Star支持一下！ | 💬 欢迎提出建议和反馈
