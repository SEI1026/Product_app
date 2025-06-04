"""
ネットワーク監視システム - 接続障害の検出と対策
"""

import socket
import requests
import threading
import time
import logging
from typing import Optional, Dict, Callable
from urllib.parse import urlparse
from PyQt5.QtCore import QTimer, QObject, pyqtSignal
from PyQt5.QtWidgets import QMessageBox

class NetworkMonitor(QObject):
    """ネットワーク接続の監視"""
    
    # シグナル定義
    connection_lost = pyqtSignal()
    connection_restored = pyqtSignal()
    connection_degraded = pyqtSignal(float)  # レスポンス時間を送信
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.is_online = True
        self.last_check_time = 0
        self.response_times = []
        self.max_response_time_samples = 10
        self.timeout_seconds = 10
        
        # 監視タイマー
        self.monitor_timer = QTimer()
        self.monitor_timer.timeout.connect(self._check_connectivity)
        
        # テスト用URL（複数設定で冗長性確保）
        self.test_urls = [
            "https://www.google.com",
            "https://www.microsoft.com", 
            "https://www.github.com"
        ]
    
    def start_monitoring(self, interval_ms: int = 30000):
        """ネットワーク監視を開始"""
        self.monitor_timer.start(interval_ms)
        logging.info("ネットワーク監視を開始しました")
    
    def stop_monitoring(self):
        """ネットワーク監視を停止"""
        self.monitor_timer.stop()
        logging.info("ネットワーク監視を停止しました")
    
    def _check_connectivity(self):
        """接続性をチェック（非同期）"""
        def check_in_thread():
            current_status = self._test_internet_connection()
            
            # 状態変化を検出
            if current_status != self.is_online:
                if current_status:
                    logging.info("ネットワーク接続が復旧しました")
                    self.connection_restored.emit()
                else:
                    logging.warning("ネットワーク接続が失われました")
                    self.connection_lost.emit()
                
                self.is_online = current_status
        
        # バックグラウンドスレッドで実行
        thread = threading.Thread(target=check_in_thread, daemon=True)
        thread.start()
    
    def _test_internet_connection(self) -> bool:
        """インターネット接続をテスト"""
        for url in self.test_urls:
            try:
                start_time = time.time()
                response = requests.get(url, timeout=self.timeout_seconds)
                end_time = time.time()
                
                if response.status_code == 200:
                    response_time = end_time - start_time
                    self._record_response_time(response_time)
                    
                    # レスポンス時間が遅い場合は警告
                    if response_time > 5.0:
                        self.connection_degraded.emit(response_time)
                    
                    return True
                    
            except (requests.RequestException, socket.timeout):
                continue
        
        return False
    
    def _record_response_time(self, response_time: float):
        """レスポンス時間を記録"""
        self.response_times.append(response_time)
        if len(self.response_times) > self.max_response_time_samples:
            self.response_times.pop(0)
    
    def get_average_response_time(self) -> float:
        """平均レスポンス時間を取得"""
        if not self.response_times:
            return 0.0
        return sum(self.response_times) / len(self.response_times)
    
    def check_specific_url(self, url: str) -> Dict[str, any]:
        """特定のURLへの接続をテスト"""
        try:
            start_time = time.time()
            response = requests.get(url, timeout=self.timeout_seconds)
            end_time = time.time()
            
            return {
                "success": True,
                "status_code": response.status_code,
                "response_time": end_time - start_time,
                "url": url
            }
            
        except requests.RequestException as e:
            return {
                "success": False,
                "error": str(e),
                "url": url
            }


class OfflineManager:
    """オフライン時の機能管理"""
    
    def __init__(self):
        self.pending_operations = []
        self.offline_mode = False
    
    def enable_offline_mode(self):
        """オフラインモードを有効化"""
        self.offline_mode = True
        logging.info("オフラインモードを有効化しました")
    
    def disable_offline_mode(self):
        """オフラインモードを無効化"""
        self.offline_mode = False
        logging.info("オフラインモードを無効化しました")
    
    def queue_operation(self, operation_type: str, operation_data: Dict):
        """オフライン時の操作をキューに追加"""
        operation = {
            "type": operation_type,
            "data": operation_data,
            "timestamp": time.time()
        }
        self.pending_operations.append(operation)
        logging.info(f"オフライン操作をキューに追加: {operation_type}")
    
    def process_pending_operations(self, processor_func: Callable) -> int:
        """接続復旧時に保留中の操作を処理"""
        processed_count = 0
        failed_operations = []
        
        for operation in self.pending_operations:
            try:
                success = processor_func(operation)
                if success:
                    processed_count += 1
                else:
                    failed_operations.append(operation)
            except Exception as e:
                logging.error(f"保留操作処理エラー: {e}")
                failed_operations.append(operation)
        
        # 失敗した操作のみを保持
        self.pending_operations = failed_operations
        
        logging.info(f"保留操作処理完了: {processed_count}件成功, {len(failed_operations)}件失敗")
        return processed_count
    
    def get_pending_count(self) -> int:
        """保留中の操作数を取得"""
        return len(self.pending_operations)


class NetworkAwareUpdateChecker:
    """ネットワーク状況を考慮した更新チェッカー"""
    
    def __init__(self, parent=None):
        self.parent = parent
        self.network_monitor = NetworkMonitor(parent)
        self.offline_manager = OfflineManager()
        
        # シグナル接続
        self.network_monitor.connection_lost.connect(self._on_connection_lost)
        self.network_monitor.connection_restored.connect(self._on_connection_restored)
        self.network_monitor.connection_degraded.connect(self._on_connection_degraded)
    
    def _on_connection_lost(self):
        """接続断での処理"""
        self.offline_manager.enable_offline_mode()
        
        QMessageBox.information(
            self.parent,
            "ネットワーク接続断",
            """インターネット接続が失われました。

オフラインモードに切り替えました。
以下の機能が制限されます：
• 自動更新チェック
• オンライン機能の利用

接続が復旧次第、自動的に通常モードに戻ります。"""
        )
    
    def _on_connection_restored(self):
        """接続復旧での処理"""
        self.offline_manager.disable_offline_mode()
        
        # 保留中の操作を処理
        pending_count = self.offline_manager.get_pending_count()
        if pending_count > 0:
            QMessageBox.information(
                self.parent,
                "接続復旧",
                f"""インターネット接続が復旧しました。

保留中の操作 {pending_count} 件を処理します。"""
            )
        else:
            QMessageBox.information(
                self.parent,
                "接続復旧",
                "インターネット接続が復旧しました。\n通常モードに戻りました。"
            )
    
    def _on_connection_degraded(self, response_time: float):
        """接続品質低下での処理"""
        logging.warning(f"ネットワーク接続が低下しています: {response_time:.2f}秒")
        
        # 重い処理（更新チェックなど）を延期
        QMessageBox.information(
            self.parent,
            "接続品質低下",
            f"""ネットワーク接続が遅くなっています。
（応答時間: {response_time:.1f}秒）

大きなファイルのダウンロードや更新チェックを
一時的に延期します。"""
        )
    
    def check_for_updates_with_retry(self, check_function: Callable, max_retries: int = 3):
        """リトライ機能付きの更新チェック"""
        for attempt in range(max_retries):
            try:
                # ネットワーク接続を確認
                if not self.network_monitor.is_online:
                    self.offline_manager.queue_operation("update_check", {})
                    return False
                
                # 更新チェックを実行
                result = check_function()
                return result
                
            except Exception as e:
                logging.error(f"更新チェック試行 {attempt + 1} 失敗: {e}")
                if attempt < max_retries - 1:
                    time.sleep(2 ** attempt)  # 指数バックオフ
                else:
                    # 最終試行失敗時はオフラインキューに追加
                    self.offline_manager.queue_operation("update_check", {})
                    return False
        
        return False


def setup_network_monitoring(app_instance):
    """アプリケーションにネットワーク監視を設定"""
    if not hasattr(app_instance, 'network_checker'):
        app_instance.network_checker = NetworkAwareUpdateChecker(app_instance)
        app_instance.network_checker.network_monitor.start_monitoring()
        logging.info("ネットワーク監視システムを初期化しました")


def check_network_before_operation(operation_name: str, parent=None) -> bool:
    """ネットワークを必要とする操作の前にチェック"""
    monitor = NetworkMonitor()
    
    # 即座に接続テスト
    is_connected = monitor._test_internet_connection()
    
    if not is_connected:
        reply = QMessageBox.question(
            parent,
            "ネットワーク未接続",
            f"""'{operation_name}' を実行するにはインターネット接続が必要です。

現在ネットワークに接続されていません。

「再試行」: 接続を再確認
「オフライン」: オフラインで続行
「キャンセル」: 操作を中止""",
            QMessageBox.Retry | QMessageBox.Ignore | QMessageBox.Cancel,
            QMessageBox.Retry
        )
        
        if reply == QMessageBox.Retry:
            return check_network_before_operation(operation_name, parent)
        elif reply == QMessageBox.Ignore:
            return True  # オフラインで続行
        else:
            return False  # キャンセル
    
    return True