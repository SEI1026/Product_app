"""
設定ファイル復旧システム - 破損した設定の自動修復
"""

import os
import json
import logging
import shutil
from pathlib import Path
from typing import Dict, Any, Optional
from PyQt5.QtCore import QSettings

class ConfigRecoveryManager:
    """設定ファイルの破損検出と復旧"""
    
    def __init__(self, app_name: str = "商品登録入力ツール"):
        self.app_name = app_name
        self.config_backup_dir = Path.home() / ".config_backup" / app_name
        self.config_backup_dir.mkdir(parents=True, exist_ok=True)
    
    def create_config_backup(self):
        """現在の設定をバックアップ"""
        try:
            settings = QSettings("株式会社大宝家具", self.app_name)
            
            # QSettingsから全設定を取得
            config_data = {}
            for key in settings.allKeys():
                config_data[key] = settings.value(key)
            
            # バックアップファイルに保存
            backup_file = self.config_backup_dir / "settings_backup.json"
            with open(backup_file, 'w', encoding='utf-8') as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2, default=str)
            
            logging.info(f"設定バックアップを作成: {backup_file}")
            
        except Exception as e:
            logging.error(f"設定バックアップ作成エラー: {e}")
    
    def detect_config_corruption(self) -> bool:
        """設定ファイルの破損をチェック"""
        try:
            settings = QSettings("株式会社大宝家具", self.app_name)
            
            # 初回起動チェック - 設定キーが1つも存在しない場合は初回起動
            if len(settings.allKeys()) == 0:
                logging.info("初回起動を検出しました - デフォルト設定を作成します")
                return False  # 破損ではなく初回起動
            
            # 基本的な設定キーが存在するかチェック
            # geometry は optional として扱う（初回起動時は存在しない）
            required_keys = ["update/auto_check_enabled"]
            
            corruption_detected = False
            for key in required_keys:
                if not settings.contains(key):
                    logging.debug(f"必須設定キー '{key}' が見つかりません")
                    corruption_detected = True
            
            # 設定値の整合性チェック
            geometry = settings.value("geometry")
            if geometry is not None and hasattr(geometry, '__len__') and len(geometry) == 0:
                logging.debug("geometry設定が空です")
                corruption_detected = True
            
            return corruption_detected
            
        except Exception as e:
            logging.error(f"設定破損チェックエラー: {e}")
            return True
    
    def restore_config_from_backup(self) -> bool:
        """バックアップから設定を復元"""
        try:
            backup_file = self.config_backup_dir / "settings_backup.json"
            if not backup_file.exists():
                logging.warning("設定バックアップファイルが見つかりません")
                return False
            
            with open(backup_file, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            
            settings = QSettings("株式会社大宝家具", self.app_name)
            settings.clear()  # 既存設定をクリア
            
            for key, value in config_data.items():
                settings.setValue(key, value)
            
            settings.sync()
            logging.info("設定をバックアップから復元しました")
            return True
            
        except Exception as e:
            logging.error(f"設定復元エラー: {e}")
            return False
    
    def reset_to_defaults(self) -> Dict[str, Any]:
        """デフォルト設定にリセット"""
        default_settings = {
            "update/auto_check_enabled": True,
            "ui/theme": "light",
            "autosave/interval_ms": 30000,
            "validation/strict_mode": False
        }
        
        try:
            settings = QSettings("株式会社大宝家具", self.app_name)
            settings.clear()
            
            for key, value in default_settings.items():
                settings.setValue(key, value)
            
            settings.sync()
            logging.info("設定をデフォルトにリセットしました")
            return default_settings
            
        except Exception as e:
            logging.error(f"デフォルトリセットエラー: {e}")
            return default_settings


def check_and_recover_config(app_name: str) -> bool:
    """設定ファイルをチェックして必要に応じて復旧"""
    recovery_manager = ConfigRecoveryManager(app_name)
    
    # 初回起動の場合はデフォルト設定を作成
    settings = QSettings("株式会社大宝家具", app_name)
    if len(settings.allKeys()) == 0:
        logging.info("初回起動 - デフォルト設定を作成します")
        recovery_manager.reset_to_defaults()
        return False
    
    if recovery_manager.detect_config_corruption():
        logging.warning("設定ファイルの破損を検出しました")
        
        # まずバックアップからの復元を試行
        if recovery_manager.restore_config_from_backup():
            return True
        
        # バックアップからの復元に失敗した場合はデフォルトに戻す
        logging.warning("設定をデフォルトにリセットします")
        recovery_manager.reset_to_defaults()
        return True
    
    # 正常な場合はバックアップを作成
    recovery_manager.create_config_backup()
    return False
