# 办公文档批量打印器

> 🚀 **最新版本**: v4.0 | 📥 **[立即下载](https://github.com/icescat/batch-document-printer/releases/latest)** | 🌟 **[GitHub 仓库](https://github.com/icescat/batch-document-printer)**

## 项目概括
本项目基于Python 3.12，用于批量打印Word文档、PowerPoint演示文稿、Excel表格和PDF文件。提供图形用户界面，支持灵活的文档添加方式（包括拖拽导入）、文件类型过滤、文档页数统计、便捷打印设置以及一键批量打印功能，能显著提升办公文档打印效率。

## 技术选型
- **主要编程语言**: Python 3.12+
- **GUI框架**: tkinter, tkinterdnd2 (拖拽支持)
- **文档处理库**: 
  - python-docx (Word文档处理)
  - python-pptx (PowerPoint处理)
  - openpyxl (Excel处理，通过COM接口)
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

## 项目结构 / 模块划分
- `/src/`: 核心源代码目录
  - `/gui/`: 图形用户界面模块
    - `main_window.py`: 主窗口界面
    - `print_settings_dialog.py`: 打印设置对话框
    - `page_count_dialog.py`: 页数统计对话框
  - `/core/`: 核心业务逻辑模块
    - `document_manager.py`: 文档管理器
    - `print_controller.py`: 打印控制器
    - `settings_manager.py`: 设置管理器
    - `page_count_manager.py`: 页数统计管理器
    - `models.py`: 数据模型定义
  - `/utils/`: 工具类和辅助函数
    - `config_utils.py`: 配置文件操作
- `/data/`: 应用数据存储目录
  - `config.json`: 应用配置文件
  - `print_settings.json`: 打印设置配置
- `/resources/`: 资源文件目录
  - `app_icon.ico`: 应用程序图标
- `/tests/`: 测试代码目录
- `main.py`: 程序入口点
- `requirements.txt`: Python依赖管理
- `.gitignore`: Git忽略配置

## 核心功能 / 模块详解
- **文档添加管理** (`document_manager.py`): 支持多种添加方式（按钮选择、拖拽导入）、单个文件和文件夹批量添加、文件格式过滤验证（.docx, .doc, .pptx, .ppt, .xlsx, .xls, .pdf）、文档列表管理与预览。
- **文件类型过滤器**: 界面顶部提供Word、PPT、Excel、PDF四个勾选框，控制扫描文件夹时包含的文档类型，默认启用Word、PPT、PDF。
- **页数统计功能** (`page_count_manager.py`): 智能识别并统计Word、PowerPoint、Excel、PDF文档的页数/张数，支持批量统计、实时显示总页数、按文件类型分类统计，为打印成本评估提供重要参考。
- **打印设置配置** (`settings_manager.py`): 检测和选择可用打印机、纸张尺寸设置（A4、A3、Letter等）、打印数量控制、双面打印选项、彩色/黑白模式选择。
- **批量打印控制** (`print_controller.py`): 调用Windows系统打印API、支持Word/PPT/Excel(COM接口)和PDF(系统默认程序)的打印调度、打印队列管理、打印进度显示、错误处理和重试机制。
- **图形用户界面** (`main_window.py`): 直观的主界面设计、文档列表展示、拖拽导入支持、实时状态更新、打印进度条显示、页数统计结果展示。
- **应用配置管理** (`config_utils.py`): 用户偏好设置持久化、最近使用的打印设置、应用程序状态保存与恢复。

## 数据模型
- **Document**: { id (UUID), file_path (Path), file_name (str), file_type (enum: WORD|PPT|EXCEL|PDF), file_size (int), page_count (int), added_time (datetime), print_status (enum: PENDING|PRINTING|COMPLETED|ERROR) }
- **PrintSettings**: { printer_name (str), paper_size (str), copies (int), duplex (bool), color_mode (enum: COLOR|GRAYSCALE), page_range (str), orientation (enum: PORTRAIT|LANDSCAPE) }
- **AppConfig**: { last_printer (str), default_settings (PrintSettings), window_geometry (dict), recent_folders (List[str]), enabled_file_types (dict) }
- **PageCountResult**: { total_pages (int), word_pages (int), ppt_slides (int), excel_sheets (int), pdf_pages (int), error_files (List[str]) }

## 技术实现细节

### 核心架构设计
- **数据模型层**: 使用dataclass定义Document、PrintSettings、AppConfig等核心数据结构
- **业务逻辑层**: DocumentManager处理文档管理，PrinterSettingsManager管理打印设置，PrintController执行打印任务，PageCountManager处理页数统计
- **界面交互层**: 基于tkinter的MainWindow主界面、PrintSettingsDialog设置对话框和PageCountDialog页数统计对话框
- **配置管理**: 通过ConfigManager实现JSON格式的配置持久化

### 文档管理器实现 (#document_manager)
- 支持.doc/.docx/.ppt/.pptx/.xls/.xlsx/.pdf格式的文件验证和添加
- 实现文件去重机制，防止重复添加相同文档
- 提供批量文件夹扫描功能，支持递归搜索
- 文件类型过滤：根据用户选择的文件类型过滤器扫描指定类型的文档
- 文档状态管理：PENDING → PRINTING → COMPLETED/ERROR

### 页数统计管理器实现 (#page_count_manager)
- **Word文档统计**: 通过python-docx库读取.docx文件的段落和分页符，计算实际页数
- **PowerPoint统计**: 使用python-pptx库获取演示文稿的幻灯片数量
- **Excel文档统计**: 通过COM接口调用Excel应用程序，统计工作簿中所有工作表的页数
- **PDF文档统计**: 使用PyPDF2/pypdf库读取PDF文件的页面数量
- **批量统计**: 支持多线程并发处理，提高大批量文档的统计效率
- **错误处理**: 对损坏或受保护的文档进行异常处理，记录错误信息
- **实时反馈**: 统计过程中提供进度条和状态更新

### 打印设置管理 (#settings_manager)
- 通过win32print API检测系统可用打印机
- 支持A4/A3/Letter等标准纸张尺寸设置
- 彩色/黑白模式选择，双面打印控制
- 打印机连接测试和状态验证

### 批量打印控制器 (#print_controller)
- 多线程打印执行，避免界面阻塞
- 支持PDF(通过系统关联)、Word(COM接口)、PowerPoint(COM接口)、Excel(COM接口)四种打印方式
- Excel打印支持：通过COM接口调用Excel应用程序，支持工作簿的所有工作表打印
- 实时进度回调和状态更新
- 错误处理和重试机制

### 配置持久化 (#config_utils)
- JSON格式配置文件存储在data目录
- 支持应用设置和打印设置的分离存储
- 配置备份和恢复功能
- 窗口几何属性保存
- 文件类型过滤器设置持久化

### 文件类型过滤器实现 (#file_type_filter)
- 界面顶部四个勾选框：Word、PPT、Excel、PDF
- 默认配置：Word(✓)、PPT(✓)、Excel(✗)、PDF(✓)
- 实时保存用户选择到配置文件
- 扫描文件夹时根据过滤器设置筛选文档类型

### 页数统计对话框实现 (#page_count_dialog)
- 独立的页数统计窗口，显示详细的统计结果
- 按文件类型分类显示：Word页数、PPT幻灯片数、Excel工作表数、PDF页数
- 总页数汇总和打印成本估算参考
- 统计进度条和状态显示
- 支持统计结果导出和保存

### 拖拽导入功能实现 (#drag_drop_import)
- **拖拽库集成**: 使用tkinterdnd2库提供跨平台拖拽支持
- **多目标支持**: 支持拖拽到文档列表区域或主窗口
- **智能识别**: 自动区分拖拽的文件和文件夹
- **批量处理**: 支持同时拖拽多个文件和文件夹
- **类型过滤**: 遵循用户设置的文件类型过滤器
- **错误处理**: 完善的异常处理和用户反馈机制
- **递归搜索**: 拖拽文件夹时自动递归搜索子目录

## 代码检查与问题记录

### 已修复问题
| 问题描述 | 发现时间 | 解决方案 | 修复状态 |
|----------|----------|----------|----------|
| 拖拽导入文件名含空格失败 | 2024-12-19 | 修改_on_drop_files方法，正确解析包含空格的文件路径 | ✅已修复 |

**问题详情**:
- **问题现象**: 拖拽导入的文件或文件夹名称如果含有空格会导入失败，提示"未找到支持的文档格式或文件已存在"
- **根本原因**: 原代码使用`event.data.split()`简单分割文件路径，导致包含空格的路径被错误分割成多个部分
- **解决方案**: 改进路径解析逻辑，支持换行符、null字符分隔的多文件路径，正确处理包含空格的文件名和文件夹名
- **影响范围**: 所有包含空格的文件路径（文件名、文件夹名、完整路径中的任何部分）
- **测试建议**: 测试拖拽含空格的文件名、文件夹名，以及同时拖拽多个文件的场景

## 开发状态跟踪
| 模块/功能        | 状态     | 备注与链接 |
|------------------|----------|----------|
| 项目架构搭建     | ✅已完成  | [技术实现](#core-architecture) |
| 文档管理器       | ✅已完成  | [文档管理](#document_manager) |
| 打印设置管理     | ✅已完成  | [设置管理](#settings_manager) |
| 主界面设计       | ✅已完成  | [主界面](#main_window) |
| 打印控制器       | ✅已完成  | [打印控制](#print_controller) |
| 工具函数模块     | ✅已完成  | [配置管理](#config_utils) |
| 应用配置管理     | ✅已完成  | [配置持久化](#config_utils) |
| GitHub发布部署   | ✅已完成  | [v1.0.0 Release](https://github.com/icescat/batch-document-printer/releases) |
| Excel文档支持    | ✅已完成  | [Excel打印](#print_controller) |
| 文件类型过滤器   | ✅已完成  | [过滤器实现](#file_type_filter) |
| v2.0版本发布     | ✅已完成  | v2.0功能开发完成 |
| 项目结构清理     | ✅已完成  | 删除冗余文件，优化项目结构 |
| 页数统计功能     | ✅已完成  | [页数统计](#page_count_manager) |
| 页数统计对话框   | ✅已完成  | [统计界面](#page_count_dialog) |
| v3.0版本发布     | ✅已完成  | 页数统计功能实现 |
| 拖拽导入功能     | ✅已完成  | [拖拽导入](#drag_drop_import) |
| v4.0版本发布     | ✅已完成  | 拖拽导入功能实现 |
| 拖拽导入Bug修复  | ✅已完成  | 修复文件名含空格的拖拽导入问题 |
| v4.1版本发布     | ✅已完成  | Bug修复版本 |

## 环境设置与运行指南
### 开发环境要求
- Python 3.12+
- Windows 10/11 操作系统
- Microsoft Office (用于Word和PPT文档处理)
- 至少一台可用的打印机

### 依赖安装
```bash
pip install -r requirements.txt
```

### 开发运行
```bash
# 直接运行主程序
python main.py

# 或者在src目录下运行
cd src
python -m gui.main_window
```

### 应用构建
```bash
# 使用构建脚本（推荐）
./build_v4.1.bat

# 手动构建
pyinstaller --clean "办公文档批量打印器v4.1.spec"

# 清理构建文件
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue
```

#### 图标配置
- **应用图标位置**: `resources/app_icon.ico`
- **图标格式**: ICO格式，支持多尺寸
- **推荐尺寸**: 16x16, 32x32, 48x48, 64x64, 128x128, 256x256像素
- **如果没有图标**: 将使用PyInstaller默认图标

#### 构建输出
- **可执行文件**: `dist/办公文档批量打印器v4.1.exe`
- **文件大小**: 约30-60MB（包含所有依赖）
- **注意**: 构建后的临时文件已配置在.gitignore中，不会提交到版本控制


## 下载与安装

### 📥 直接下载（推荐）
**[点击这里下载最新版本](https://github.com/icescat/batch-document-printer/releases/latest)**

1. 在 Release 页面下载 `办公文档批量打印器.exe`
2. 双击运行即可使用，无需安装
3. 首次启动可能需要10-30秒初始化时间

### 💻 系统要求
- **操作系统**: Windows 10/11 (64位)
- **必需软件**: Microsoft Office (用于Word、PowerPoint和Excel文档)
- **可选软件**: PDF阅读器 (用于PDF文档打印)
- **硬件要求**: 至少一台可用的打印机

### 🔧 从源码运行
如果您是开发者或想要自定义功能：
```bash
git clone https://github.com/icescat/batch-document-printer.git
cd batch-document-printer
pip install -r requirements.txt
python main.py
```

## 部署与发布
- **发布渠道**: GitHub Releases
- **更新方式**: 手动下载新版本覆盖
- **版本管理**: 语义化版本控制 (Semantic Versioning)
- **发布频率**: 根据功能更新和Bug修复情况 