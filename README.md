# 批量文档打印桌面应用

## 项目概括
本项目旨在开发一个基于Python 3.12的Windows桌面应用，用于批量打印Word文档、PowerPoint演示文稿和PDF文件。该应用提供图形用户界面，支持灵活的文档添加方式、完善的打印设置功能，以及一键批量打印，显著提升办公文档打印效率。

## 技术选型
- **主要编程语言**: Python 3.12+
- **GUI框架**: tkinter (内置，无需额外依赖) 或 PyQt5/PySide6 (更现代的界面)
- **文档处理库**: 
  - python-docx (Word文档处理)
  - python-pptx (PowerPoint处理)
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
    - `document_list_widget.py`: 文档列表组件
  - `/core/`: 核心业务逻辑模块
    - `document_manager.py`: 文档管理器
    - `print_controller.py`: 打印控制器
    - `settings_manager.py`: 设置管理器
  - `/utils/`: 工具类和辅助函数
    - `file_utils.py`: 文件操作工具
    - `print_utils.py`: 打印相关工具函数
    - `config_utils.py`: 配置文件操作
- `/data/`: 应用数据存储目录
  - `config.json`: 应用配置文件
  - `print_settings.json`: 打印设置配置
- `/tests/`: 测试代码目录
- `/build/`: 构建输出目录
- `/dist/`: 最终发布目录
- `main.py`: 程序入口点
- `requirements.txt`: Python依赖管理
- `.gitignore`: Git忽略配置
- `build.bat`: Windows构建脚本

## 核心功能 / 模块详解
- **文档添加管理** (`document_manager.py`): 支持拖拽添加单个文件、按文件夹批量添加、文件格式过滤验证（.docx, .doc, .pptx, .ppt, .pdf）、文档列表管理与预览。
- **打印设置配置** (`settings_manager.py`): 检测和选择可用打印机、纸张尺寸设置（A4、A3、Letter等）、打印数量控制、双面打印选项、彩色/黑白模式选择。
- **批量打印控制** (`print_controller.py`): 调用Windows系统打印API、支持多种文档格式的打印调度、打印队列管理、打印进度显示、错误处理和重试机制。
- **图形用户界面** (`main_window.py`): 直观的主界面设计、文档列表展示、拖拽支持、实时状态更新、打印进度条显示。
- **应用配置管理** (`config_utils.py`): 用户偏好设置持久化、最近使用的打印设置、应用程序状态保存与恢复。

## 数据模型
- **Document**: { id (UUID), file_path (Path), file_name (str), file_type (enum: WORD|PPT|PDF), file_size (int), added_time (datetime), print_status (enum: PENDING|PRINTING|COMPLETED|ERROR) }
- **PrintSettings**: { printer_name (str), paper_size (str), copies (int), duplex (bool), color_mode (enum: COLOR|GRAYSCALE), page_range (str), orientation (enum: PORTRAIT|LANDSCAPE) }
- **AppConfig**: { last_printer (str), default_settings (PrintSettings), window_geometry (dict), recent_folders (List[str]) }

## 技术实现细节

### 核心架构设计
- **数据模型层**: 使用dataclass定义Document、PrintSettings、AppConfig等核心数据结构
- **业务逻辑层**: DocumentManager处理文档管理，PrinterSettingsManager管理打印设置，PrintController执行打印任务
- **界面交互层**: 基于tkinter的MainWindow主界面和PrintSettingsDialog设置对话框
- **配置管理**: 通过ConfigManager实现JSON格式的配置持久化

### 文档管理器实现 (#document_manager)
- 支持.doc/.docx/.ppt/.pptx/.pdf格式的文件验证和添加
- 实现文件去重机制，防止重复添加相同文档
- 提供批量文件夹扫描功能，支持递归搜索
- 文档状态管理：PENDING → PRINTING → COMPLETED/ERROR

### 打印设置管理 (#settings_manager)
- 通过win32print API检测系统可用打印机
- 支持A4/A3/Letter等标准纸张尺寸设置
- 彩色/黑白模式选择，双面打印控制
- 打印机连接测试和状态验证

### 批量打印控制器 (#print_controller)
- 多线程打印执行，避免界面阻塞
- 支持PDF(通过系统关联)、Word(COM接口)、PowerPoint(COM接口)三种打印方式
- 实时进度回调和状态更新
- 错误处理和重试机制

### 配置持久化 (#config_utils)
- JSON格式配置文件存储在data目录
- 支持应用设置和打印设置的分离存储
- 配置备份和恢复功能
- 窗口几何属性保存

## 开发状态跟踪
| 模块/功能        | 状态     | 负责人 | 计划完成日期 | 实际完成日期 | 备注与链接 |
|------------------|----------|--------|--------------|--------------|------------|
| 项目架构搭建     | ✅已完成  | AI     | 2024-12-19   | 2024-12-19   | [技术实现](#core-architecture) |
| 文档管理器       | ✅已完成  | AI     | 2024-12-19   | 2024-12-19   | [文档管理](#document_manager) |
| 打印设置管理     | ✅已完成  | AI     | 2024-12-19   | 2024-12-19   | [设置管理](#settings_manager) |
| 主界面设计       | ✅已完成  | AI     | 2024-12-19   | 2024-12-19   | [主界面](#main_window) |
| 打印控制器       | ✅已完成  | AI     | 2024-12-19   | 2024-12-19   | [打印控制](#print_controller) |
| 工具函数模块     | ✅已完成  | AI     | 2024-12-19   | 2024-12-19   | [配置管理](#config_utils) |
| 应用配置管理     | ✅已完成  | AI     | 2024-12-19   | 2024-12-19   | [配置持久化](#config_utils) |
| 单元测试编写     | 待完成   | AI     | 2024-12-20   |              | 需要后续添加 |
| 应用打包构建     | ✅已完成  | AI     | 2024-12-20   | 2024-12-19   | build.bat脚本 |

## 代码检查与问题记录
[本部分用于记录代码检查结果和开发过程中遇到的问题及其解决方案。]

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
build.bat

# 手动构建
pyinstaller --onefile --windowed --name="批量文档打印器" --icon=resources/app_icon.ico main.py
```

#### 图标配置
- **应用图标位置**: `resources/app_icon.ico`
- **图标格式**: ICO格式，支持多尺寸
- **推荐尺寸**: 16x16, 32x32, 48x48, 64x64, 128x128, 256x256像素
- **如果没有图标**: 将使用PyInstaller默认图标

#### 构建脚本功能
`build.bat` 脚本包含以下功能：
- ✅ 自动检查Python环境
- ✅ 安装项目依赖
- ✅ 检测应用图标文件
- ✅ 生成版本信息
- ✅ 清理旧构建文件
- ✅ 优化打包参数
- ✅ 包含资源文件

#### 构建输出
- **可执行文件**: `dist/批量文档打印器.exe`
- **版本信息**: 包含开发者署名"喵言喵语"
- **文件大小**: 约30-60MB（包含所有依赖）

## 部署与发布
- **目标平台**: Windows 10/11 (64位)
- **发布方式**: 独立exe可执行文件
- **运行要求**: 
  - Windows系统
  - Microsoft Office (用于Word/PPT打印)
  - PDF阅读器 (用于PDF打印)
  - 可用打印机
- **安装说明**: 无需安装，下载exe文件直接运行
- **首次运行**: 可能需要较长时间初始化 