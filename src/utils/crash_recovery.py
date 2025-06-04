"""
クラッシュ復旧システム - 予期しない終了からの自動復旧
"""

import os
import sys
import json
import logging
import tempfile
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional

class CrashRecoveryManager:
    """アプリケーションのクラッシュ復旧を管理"""
    
    def __init__(self, app_name: str = "商品登録入力ツール"):
        self.app_name = app_name
        self.temp_dir = Path(tempfile.gettempdir()) / "ProductAppRecovery"
        self.temp_dir.mkdir(exist_ok=True)
        
        # 復旧ファイルのパス
        self.crash_log_file = self.temp_dir / "crash_log.json"
        self.session_file = self.temp_dir / "current_session.json"
        self.backup_data_file = self.temp_dir / "emergency_backup.json"
    
    def start_session(self, session_data: Dict[str, Any]):
        """セッション開始時の情報を記録"""
        try:
            session_info = {
                "start_time": datetime.now().isoformat(),
                "pid": os.getpid(),
                "version": session_data.get("version", "unknown"),
                "user_data_dir": session_data.get("user_data_dir", ""),
                "manage_file_path": session_data.get("manage_file_path", ""),
                "last_heartbeat": datetime.now().isoformat()
            }
            
            with open(self.session_file, 'w', encoding='utf-8') as f:
                json.dump(session_info, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            logging.error(f"セッション開始記録エラー: {e}")
    
    def update_heartbeat(self, current_data: Optional[Dict[str, Any]] = None):
        """生存証明の更新（定期的に呼び出す）"""
        try:
            if self.session_file.exists():
                with open(self.session_file, 'r', encoding='utf-8') as f:
                    session_info = json.load(f)
                
                session_info["last_heartbeat"] = datetime.now().isoformat()
                if current_data:
                    session_info["current_data"] = current_data
                
                with open(self.session_file, 'w', encoding='utf-8') as f:
                    json.dump(session_info, f, ensure_ascii=False, indent=2)
                    
        except Exception as e:
            logging.error(f"ハートビート更新エラー: {e}")
    
    def create_emergency_backup(self, data: Dict[str, Any]):
        """緊急バックアップの作成"""
        try:
            backup_data = {
                "timestamp": datetime.now().isoformat(),
                "data": data,
                "source": "emergency_backup"
            }
            
            with open(self.backup_data_file, 'w', encoding='utf-8') as f:
                json.dump(backup_data, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            logging.error(f"緊急バックアップ作成エラー: {e}")
    
    def log_crash(self, error_info: str):
        """クラッシュ情報をログに記録"""
        try:
            crash_info = {
                "timestamp": datetime.now().isoformat(),
                "error": error_info,
                "traceback": traceback.format_exc(),
                "pid": os.getpid()
            }
            
            # 既存のクラッシュログを読み込み
            crash_history = []
            if self.crash_log_file.exists():
                with open(self.crash_log_file, 'r', encoding='utf-8') as f:
                    crash_history = json.load(f)
            
            crash_history.append(crash_info)
            
            # 最新10件のみ保持
            crash_history = crash_history[-10:]
            
            with open(self.crash_log_file, 'w', encoding='utf-8') as f:
                json.dump(crash_history, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            logging.error(f"クラッシュログ記録エラー: {e}")
    
    def check_for_crash(self) -> Optional[Dict[str, Any]]:
        """前回のセッションがクラッシュしたかチェック"""
        try:
            if not self.session_file.exists():
                return None
            
            with open(self.session_file, 'r', encoding='utf-8') as f:
                session_info = json.load(f)
            
            # ハートビートが5分以上前の場合はクラッシュと判定
            last_heartbeat = datetime.fromisoformat(session_info.get("last_heartbeat", ""))
            if (datetime.now() - last_heartbeat).total_seconds() > 300:  # 5分
                return session_info
            
            return None
            
        except Exception as e:
            logging.error(f"クラッシュチェックエラー: {e}")
            return None
    
    def get_emergency_backup(self) -> Optional[Dict[str, Any]]:
        """緊急バックアップデータを取得"""
        try:
            if self.backup_data_file.exists():
                with open(self.backup_data_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return None
            
        except Exception as e:
            logging.error(f"緊急バックアップ取得エラー: {e}")
            return None
    
    def clean_session(self):
        """正常終了時のクリーンアップ"""
        try:
            if self.session_file.exists():
                self.session_file.unlink()
                
        except Exception as e:
            logging.error(f"セッションクリーンアップエラー: {e}")


def setup_crash_handler(recovery_manager: CrashRecoveryManager):
    """グローバル例外ハンドラーを設定"""
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        
        error_msg = f"Uncaught exception: {exc_type.__name__}: {exc_value}"
        logging.critical(error_msg, exc_info=(exc_type, exc_value, exc_traceback))
        recovery_manager.log_crash(error_msg)
    
    sys.excepthook = handle_exception


# PyQt5用の例外ハンドラー
def setup_qt_exception_handler(recovery_manager: CrashRecoveryManager):
    """PyQt5例外ハンドラーを設定"""
    try:
        from PyQt5.QtCore import qInstallMessageHandler, QtMsgType
        
        def qt_message_handler(mode, context, message):
            if mode == QtMsgType.QtCriticalMsg or mode == QtMsgType.QtFatalMsg:
                error_msg = f"Qt {mode}: {message}"
                logging.error(error_msg)
                recovery_manager.log_crash(error_msg)
            
        qInstallMessageHandler(qt_message_handler)
        
    except ImportError:
        pass  # PyQt5が利用できない場合はスキップ