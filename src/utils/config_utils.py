"""
配置文件操作工具
负责应用配置的保存和加载
"""
import json
from pathlib import Path
from typing import Optional
from src.core.models import AppConfig, PrintSettings


class ConfigManager:
    """配置管理器类"""
    
    def __init__(self, config_dir: Optional[Path] = None):
        """
        初始化配置管理器
        
        Args:
            config_dir: 配置文件目录，默认为data目录
        """
        if config_dir is None:
            # 默认使用项目根目录下的data文件夹
            self.config_dir = Path(__file__).parent.parent.parent / "data"
        else:
            self.config_dir = config_dir
        
        # 确保配置目录存在
        self.config_dir.mkdir(parents=True, exist_ok=True)
        
        # 配置文件路径
        self.app_config_file = self.config_dir / "config.json"
        self.print_settings_file = self.config_dir / "print_settings.json"
    
    def load_app_config(self) -> AppConfig:
        """
        加载应用配置
        
        Returns:
            AppConfig对象
        """
        try:
            if self.app_config_file.exists():
                with open(self.app_config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return AppConfig.from_dict(data)
            else:
                print("配置文件不存在，使用默认配置")
                return AppConfig()
                
        except Exception as e:
            print(f"加载应用配置失败: {e}")
            return AppConfig()
    
    def save_app_config(self, config: AppConfig) -> bool:
        """
        保存应用配置
        
        Args:
            config: AppConfig对象
            
        Returns:
            是否保存成功
        """
        try:
            with open(self.app_config_file, 'w', encoding='utf-8') as f:
                json.dump(config.to_dict(), f, indent=2, ensure_ascii=False)
            print(f"应用配置已保存到: {self.app_config_file}")
            return True
            
        except Exception as e:
            print(f"保存应用配置失败: {e}")
            return False
    
    def load_print_settings(self) -> PrintSettings:
        """
        加载打印设置
        
        Returns:
            PrintSettings对象
        """
        try:
            if self.print_settings_file.exists():
                with open(self.print_settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return PrintSettings.from_dict(data)
            else:
                print("打印设置文件不存在，使用默认设置")
                return PrintSettings()
                
        except Exception as e:
            print(f"加载打印设置失败: {e}")
            return PrintSettings()
    
    def save_print_settings(self, settings: PrintSettings) -> bool:
        """
        保存打印设置
        
        Args:
            settings: PrintSettings对象
            
        Returns:
            是否保存成功
        """
        try:
            with open(self.print_settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings.to_dict(), f, indent=2, ensure_ascii=False)
            print(f"打印设置已保存到: {self.print_settings_file}")
            return True
            
        except Exception as e:
            print(f"保存打印设置失败: {e}")
            return False
    
    def backup_config(self, backup_name: Optional[str] = None) -> bool:
        """
        备份配置文件
        
        Args:
            backup_name: 备份文件名后缀，默认使用时间戳
            
        Returns:
            是否备份成功
        """
        try:
            from datetime import datetime
            
            if backup_name is None:
                backup_name = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            backup_dir = self.config_dir / "backups"
            backup_dir.mkdir(exist_ok=True)
            
            # 备份应用配置
            if self.app_config_file.exists():
                backup_app_file = backup_dir / f"config_{backup_name}.json"
                backup_app_file.write_text(
                    self.app_config_file.read_text(encoding='utf-8'),
                    encoding='utf-8'
                )
            
            # 备份打印设置
            if self.print_settings_file.exists():
                backup_print_file = backup_dir / f"print_settings_{backup_name}.json"
                backup_print_file.write_text(
                    self.print_settings_file.read_text(encoding='utf-8'),
                    encoding='utf-8'
                )
            
            print(f"配置文件已备份到: {backup_dir}")
            return True
            
        except Exception as e:
            print(f"备份配置文件失败: {e}")
            return False
    
    def restore_config(self, backup_name: str) -> bool:
        """
        恢复配置文件
        
        Args:
            backup_name: 备份文件名后缀
            
        Returns:
            是否恢复成功
        """
        try:
            backup_dir = self.config_dir / "backups"
            
            # 恢复应用配置
            backup_app_file = backup_dir / f"config_{backup_name}.json"
            if backup_app_file.exists():
                self.app_config_file.write_text(
                    backup_app_file.read_text(encoding='utf-8'),
                    encoding='utf-8'
                )
            
            # 恢复打印设置
            backup_print_file = backup_dir / f"print_settings_{backup_name}.json"
            if backup_print_file.exists():
                self.print_settings_file.write_text(
                    backup_print_file.read_text(encoding='utf-8'),
                    encoding='utf-8'
                )
            
            print(f"配置文件已从备份恢复: {backup_name}")
            return True
            
        except Exception as e:
            print(f"恢复配置文件失败: {e}")
            return False
    
    def reset_to_defaults(self) -> bool:
        """
        重置为默认配置
        
        Returns:
            是否重置成功
        """
        try:
            # 先备份当前配置
            self.backup_config("before_reset")
            
            # 创建默认配置
            default_app_config = AppConfig()
            default_print_settings = PrintSettings()
            
            # 保存默认配置
            success = (
                self.save_app_config(default_app_config) and 
                self.save_print_settings(default_print_settings)
            )
            
            if success:
                print("配置已重置为默认值")
            
            return success
            
        except Exception as e:
            print(f"重置配置失败: {e}")
            return False
    
    def get_config_info(self) -> dict:
        """
        获取配置文件信息
        
        Returns:
            配置信息字典
        """
        info = {
            'config_dir': str(self.config_dir),
            'app_config_exists': self.app_config_file.exists(),
            'print_settings_exists': self.print_settings_file.exists(),
            'app_config_size': 0,
            'print_settings_size': 0,
            'last_modified': {}
        }
        
        try:
            if self.app_config_file.exists():
                stat = self.app_config_file.stat()
                info['app_config_size'] = stat.st_size
                info['last_modified']['app_config'] = stat.st_mtime
            
            if self.print_settings_file.exists():
                stat = self.print_settings_file.stat()
                info['print_settings_size'] = stat.st_size
                info['last_modified']['print_settings'] = stat.st_mtime
                
        except Exception as e:
            print(f"获取配置信息失败: {e}")
        
        return info 