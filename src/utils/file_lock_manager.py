"""
ファイルロック管理システム - 重複起動や排他制御
"""

import os
import time
import psutil
import logging
from pathlib import Path
from typing import Optional, List, Dict
from PyQt5.QtWidgets import QMessageBox

class FileLockManager:
    """ファイルロックと重複起動の管理"""
    
    def __init__(self, app_name: str = "商品登録入力ツール"):
        self.app_name = app_name
        self.lock_dir = Path.home() / ".app_locks"
        self.lock_dir.mkdir(exist_ok=True)
        self.lock_file = self.lock_dir / f"{app_name}.lock"
        self.current_pid = os.getpid()
    
    def acquire_app_lock(self) -> bool:
        """アプリケーションロックを取得"""
        try:
            if self.lock_file.exists():
                # 既存のロックファイルをチェック
                with open(self.lock_file, 'r') as f:
                    existing_pid = int(f.read().strip())
                
                # プロセスが実際に動いているかチェック
                if self._is_process_running(existing_pid):
                    return False  # 既に起動中
                else:
                    # ゾンビロックファイルを削除
                    self.lock_file.unlink()
                    logging.info(f"ゾンビロックファイルを削除: PID {existing_pid}")
            
            # 新しいロックファイルを作成
            with open(self.lock_file, 'w') as f:
                f.write(str(self.current_pid))
            
            return True
            
        except Exception as e:
            logging.error(f"アプリロック取得エラー: {e}")
            return True  # エラー時は起動を許可
    
    def release_app_lock(self):
        """アプリケーションロックを解放"""
        try:
            if self.lock_file.exists():
                with open(self.lock_file, 'r') as f:
                    lock_pid = int(f.read().strip())
                
                if lock_pid == self.current_pid:
                    self.lock_file.unlink()
                    logging.info("アプリロックを解放しました")
                    
        except Exception as e:
            logging.error(f"アプリロック解放エラー: {e}")
    
    def _is_process_running(self, pid: int) -> bool:
        """指定されたPIDのプロセスが動いているかチェック"""
        try:
            process = psutil.Process(pid)
            return process.is_running() and self.app_name in process.name()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            return False
    
    def check_file_conflicts(self, file_paths: List[str]) -> List[str]:
        """ファイルが他のプロセスで開かれていないかチェック"""
        conflicted_files = []
        
        for file_path in file_paths:
            if self._is_file_locked(file_path):
                conflicted_files.append(file_path)
        
        return conflicted_files
    
    def _is_file_locked(self, file_path: str) -> bool:
        """ファイルが他のプロセスで開かれているかチェック"""
        try:
            # Windowsでのファイルロックチェック
            if os.name == 'nt':
                try:
                    # ファイルを排他モードで開いてみる
                    with open(file_path, 'r+b') as f:
                        pass
                    return False
                except (IOError, OSError):
                    return True
            else:
                # Unix系での簡易チェック
                import fcntl
                try:
                    with open(file_path, 'r') as f:
                        fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                        fcntl.flock(f.fileno(), fcntl.LOCK_UN)
                    return False
                except (IOError, OSError):
                    return True
                    
        except Exception:
            return False  # エラー時は非ロック状態と判定
    
    def wait_for_file_release(self, file_path: str, timeout: int = 30) -> bool:
        """ファイルのロック解除を待機"""
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            if not self._is_file_locked(file_path):
                return True
            time.sleep(1)
        
        return False
    
    def find_processes_using_file(self, file_path: str) -> List[Dict]:
        """ファイルを使用しているプロセスを特定"""
        using_processes = []
        
        try:
            for proc in psutil.process_iter(['pid', 'name', 'open_files']):
                try:
                    if proc.info['open_files']:
                        for f in proc.info['open_files']:
                            if os.path.samefile(f.path, file_path):
                                using_processes.append({
                                    'pid': proc.info['pid'],
                                    'name': proc.info['name']
                                })
                except (psutil.NoSuchProcess, psutil.AccessDenied, OSError):
                    continue
                    
        except Exception as e:
            logging.error(f"プロセス検索エラー: {e}")
        
        return using_processes


def handle_duplicate_launch(parent=None) -> bool:
    """重複起動の処理"""
    lock_manager = FileLockManager()
    
    if not lock_manager.acquire_app_lock():
        reply = QMessageBox.question(
            parent,
            "重複起動の検出",
            f"""{lock_manager.app_name} は既に起動しています。

既存のアプリケーションをアクティブにしますか？

「はい」: 既存アプリを前面に表示
「いいえ」: 強制的に新しいインスタンスを起動""",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        
        if reply == QMessageBox.Yes:
            # 既存アプリをアクティブにする処理
            # （実装は環境依存のため省略）
            return False  # 起動を中止
        else:
            # 強制起動の場合はロックファイルを無視
            return True
    
    return True  # 正常起動


def handle_file_conflicts(file_paths: List[str], parent=None) -> bool:
    """ファイル競合の処理"""
    lock_manager = FileLockManager()
    conflicted_files = lock_manager.check_file_conflicts(file_paths)
    
    if conflicted_files:
        file_list = "\n".join([f"• {f}" for f in conflicted_files])
        
        reply = QMessageBox.question(
            parent,
            "ファイル使用中",
            f"""以下のファイルが他のアプリケーションで開かれています：

{file_list}

これらのファイルを閉じてから続行してください。

「再試行」: ファイルのロック解除を待機
「強制実行」: 警告を無視して続行
「キャンセル」: 処理を中止""",
            QMessageBox.Retry | QMessageBox.Ignore | QMessageBox.Cancel,
            QMessageBox.Retry
        )
        
        if reply == QMessageBox.Retry:
            # ファイルロック解除を待機
            for file_path in conflicted_files:
                if not lock_manager.wait_for_file_release(file_path):
                    QMessageBox.warning(
                        parent,
                        "タイムアウト",
                        f"ファイル '{file_path}' のロック解除を待機中にタイムアウトしました。"
                    )
                    return False
            return True
            
        elif reply == QMessageBox.Ignore:
            return True  # 強制実行
        else:
            return False  # キャンセル
    
    return True  # 競合なし