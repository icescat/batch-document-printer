"""
窗口管理器
处理窗口状态管理、配置持久化、用户偏好设置等功能
"""
import tkinter as tk
from typing import Dict, Any, Optional


class WindowManager:
    """窗口管理器"""
    
    def __init__(self, root_window, config_manager):
        """
        初始化窗口管理器
        
        Args:
            root_window: 主窗口对象
            config_manager: 配置管理器实例
        """
        self.root = root_window
        self.config_manager = config_manager
        
    def restore_window_geometry(self, app_config):
        """恢复窗口几何属性"""
        geometry = app_config.window_geometry
        if geometry:
            try:
                width = geometry.get('width', 900)
                height = geometry.get('height', 600)
                self.root.geometry(f"{width}x{height}")
                
                if 'x' in geometry and 'y' in geometry:
                    x = geometry['x']
                    y = geometry['y']
                    self.root.geometry(f"+{x}+{y}")
                    
                print(f"窗口几何属性已恢复: {width}x{height}+{x}+{y}")
            except Exception as e:
                print(f"恢复窗口几何属性失败: {e}")
                # 使用默认几何属性
                self.root.geometry("900x600")
        else:
            # 首次启动，使用默认几何属性
            self.root.geometry("900x600")
    
    def save_window_geometry(self, app_config):
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
                
                app_config.window_geometry = {
                    'width': width,
                    'height': height,
                    'x': x,
                    'y': y
                }
                
                self.config_manager.save_app_config(app_config)
                print(f"窗口几何属性已保存: {width}x{height}+{x}+{y}")
        except Exception as e:
            print(f"保存窗口几何属性失败: {e}")
    
    def center_window(self, window, width: Optional[int] = None, height: Optional[int] = None):
        """将窗口居中显示"""
        window.update_idletasks()
        
        # 获取主窗口位置和大小
        main_x = self.root.winfo_rootx()
        main_y = self.root.winfo_rooty()
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()
        
        # 计算子窗口位置
        if width is None or height is None:
            window_width = window.winfo_reqwidth()
            window_height = window.winfo_reqheight()
        else:
            window_width = width
            window_height = height
        
        x = main_x + (main_width - window_width) // 2
        y = main_y + (main_height - window_height) // 2
        
        window.geometry(f"+{x}+{y}")
    
    def setup_window_close_handler(self, close_callback):
        """设置窗口关闭事件处理"""
        def on_closing():
            try:
                # 执行传入的关闭回调
                if close_callback:
                    result = close_callback()
                    if result is False:  # 如果回调返回False，取消关闭
                        return
                
                # 关闭窗口
                self.root.destroy()
            except Exception as e:
                print(f"窗口关闭处理出错: {e}")
                self.root.destroy()
        
        self.root.protocol("WM_DELETE_WINDOW", on_closing)
    
    def save_user_preferences(self, app_config, enabled_file_types: Dict[str, bool]):
        """保存用户偏好设置"""
        try:
            # 更新文件类型过滤设置
            app_config.enabled_file_types = enabled_file_types
            
            # 保存配置
            self.config_manager.save_app_config(app_config)
            print(f"用户偏好已保存: {enabled_file_types}")
        except Exception as e:
            print(f"保存用户偏好失败: {e}")
    
    def load_user_preferences(self) -> Dict[str, Any]:
        """加载用户偏好设置"""
        try:
            app_config = self.config_manager.load_app_config()
            return {
                'enabled_file_types': app_config.enabled_file_types,
                'window_geometry': app_config.window_geometry
            }
        except Exception as e:
            print(f"加载用户偏好失败: {e}")
            # 返回默认设置
            return {
                'enabled_file_types': {
                    'word': True,
                    'ppt': True,
                    'excel': False,
                    'pdf': True
                },
                'window_geometry': None
            }
    
    def set_window_title(self, title: str, version: Optional[str] = None):
        """设置窗口标题"""
        if version:
            full_title = f"{title} {version}"
        else:
            full_title = title
            
        self.root.title(full_title)
    
    def set_window_icon(self, icon_path: str):
        """设置窗口图标"""
        try:
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"设置窗口图标失败: {e}")
    
    def set_window_minimum_size(self, min_width: int, min_height: int):
        """设置窗口最小尺寸"""
        self.root.minsize(min_width, min_height)
    
    def set_window_resizable(self, width_resizable: bool = True, height_resizable: bool = True):
        """设置窗口是否可调整大小"""
        self.root.resizable(width_resizable, height_resizable) 