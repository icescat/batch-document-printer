"""
鼠标悬停提示框组件
为GUI元素添加悬停时显示的提示信息
"""
import tkinter as tk
from tkinter import ttk
from typing import Optional, Union

# ========== TOOLTIP 配置设置 (可在此处修改) ==========
# 延迟显示时间 (毫秒) - 鼠标悬停多久后显示tooltip
FILTER_DELAY = 50         # 文件类型过滤器tooltip延迟

# 文本换行设置 (像素) - tooltip文本的最大宽度
FILTER_WRAPLENGTH = 250    # 过滤器tooltip文本宽度

# 外观设置
BACKGROUND_COLOR = '#ffffe0'  # 背景颜色 (浅黄色)
FOREGROUND_COLOR = 'black'    # 文字颜色
ALPHA = 0.8                   # 透明度 (0.0-1.0)
FONT_FAMILY = '微软雅黑'       # 字体
FONT_SIZE = 9                 # 字体大小
BORDER_WIDTH = 1              # 边框宽度
# ================================================


def update_tooltip_timing(filter_delay=None):
    """
    快速更新tooltip时间设置的便捷函数
    
    Args:
        filter_delay: 过滤器tooltip延迟时间
    
    使用示例:
        # 让tooltip更快显示
        update_tooltip_timing(filter_delay=200)
    """
    global FILTER_DELAY
    
    if filter_delay is not None:
        FILTER_DELAY = filter_delay
    
    print(f"Tooltip时间已更新: 过滤器={FILTER_DELAY}ms")


class ToolTip:
    """
    鼠标悬停提示框组件
    
    使用方法:
    tooltip = ToolTip(widget, "这是提示文本")
    """
    
    def __init__(self, widget: tk.Widget, text: str, delay: int = 500, wraplength: int = 300):
        """
        初始化提示框
        
        Args:
            widget: 要添加提示的控件
            text: 提示文本
            delay: 延迟显示时间(毫秒)
            wraplength: 文本换行长度(像素)
        """
        self.widget = widget
        self.text = text
        self.delay = delay
        self.wraplength = wraplength
        
        self.tooltip_window: Optional[tk.Toplevel] = None
        self.hover_timer: Optional[str] = None
        
        # 绑定鼠标事件
        self.widget.bind('<Enter>', self._on_enter)
        self.widget.bind('<Leave>', self._on_leave)
        self.widget.bind('<Motion>', self._on_motion)
        self.widget.bind('<ButtonPress>', self._on_click)
    
    def _on_enter(self, event=None):
        """鼠标进入控件时"""
        self._schedule_tooltip()
    
    def _on_leave(self, event=None):
        """鼠标离开控件时"""
        self._cancel_tooltip()
        self._hide_tooltip()
    
    def _on_motion(self, event=None):
        """鼠标移动时"""
        # 如果已显示tooltip，更新位置
        if self.tooltip_window:
            self._update_tooltip_position(event)
    
    def _on_click(self, event=None):
        """鼠标点击时隐藏tooltip"""
        self._hide_tooltip()
    
    def _schedule_tooltip(self):
        """安排显示tooltip"""
        self._cancel_tooltip()
        self.hover_timer = self.widget.after(self.delay, self._show_tooltip)
    
    def _cancel_tooltip(self):
        """取消显示tooltip"""
        if self.hover_timer:
            self.widget.after_cancel(self.hover_timer)
            self.hover_timer = None
    
    def _show_tooltip(self):
        """显示tooltip"""
        if self.tooltip_window or not self.text:
            return
        
        # 创建顶级窗口
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)  # 无边框
        self.tooltip_window.wm_attributes('-topmost', True)  # 置顶
        
        # 在某些系统上设置透明度
        try:
            self.tooltip_window.wm_attributes('-alpha', ALPHA)
        except tk.TclError:
            pass  # 某些系统不支持透明度
        
        # 创建标签
        label = tk.Label(
            self.tooltip_window,
            text=self.text,
            justify='left',
            background=BACKGROUND_COLOR,
            foreground=FOREGROUND_COLOR,
            relief='solid',
            borderwidth=BORDER_WIDTH,
            wraplength=self.wraplength,
            font=(FONT_FAMILY, FONT_SIZE)
        )
        label.pack()
        
        # 设置位置
        self._position_tooltip()
    
    def _position_tooltip(self):
        """设置tooltip位置"""
        if not self.tooltip_window:
            return
        
        # 获取鼠标位置
        x = self.widget.winfo_pointerx()
        y = self.widget.winfo_pointery()
        
        # 更新窗口以获取准确尺寸
        self.tooltip_window.update_idletasks()
        
        # 获取tooltip尺寸
        tooltip_width = self.tooltip_window.winfo_reqwidth()
        tooltip_height = self.tooltip_window.winfo_reqheight()
        
        # 获取屏幕尺寸
        screen_width = self.tooltip_window.winfo_screenwidth()
        screen_height = self.tooltip_window.winfo_screenheight()
        
        # 调整位置避免超出屏幕
        if x + tooltip_width + 10 > screen_width:
            x = x - tooltip_width - 10
        else:
            x = x + 10
        
        if y + tooltip_height + 10 > screen_height:
            y = y - tooltip_height - 10
        else:
            y = y + 10
        
        # 设置窗口位置
        self.tooltip_window.wm_geometry(f'+{x}+{y}')
    
    def _update_tooltip_position(self, event):
        """更新tooltip位置"""
        if self.tooltip_window:
            self._position_tooltip()
    
    def _hide_tooltip(self):
        """隐藏tooltip"""
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None
    
    def update_text(self, new_text: str):
        """更新提示文本"""
        self.text = new_text
        # 如果tooltip正在显示，重新创建
        if self.tooltip_window:
            self._hide_tooltip()
            self._show_tooltip()
    
    def set_enabled(self, enabled: bool):
        """启用或禁用tooltip"""
        if enabled:
            self.widget.bind('<Enter>', self._on_enter)
            self.widget.bind('<Leave>', self._on_leave)
            self.widget.bind('<Motion>', self._on_motion)
        else:
            self.widget.unbind('<Enter>')
            self.widget.unbind('<Leave>')
            self.widget.unbind('<Motion>')
            self._hide_tooltip()



def create_button_tooltip(button: tk.Widget, text: str, delay: Optional[int] = None) -> ToolTip:
    """
    为文件类型过滤器按钮创建tooltip的便捷函数
    
    Args:
        button: 按钮控件
        text: 提示文本
        delay: 延迟时间，如果为None则使用配置默认值
    
    Returns:
        ToolTip对象
    """
    if delay is None:
        delay = FILTER_DELAY
    return ToolTip(button, text, delay=delay, wraplength=FILTER_WRAPLENGTH) 