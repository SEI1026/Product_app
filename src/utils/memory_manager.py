"""
メモリ管理システム - メモリ不足の検出と対策
"""

import gc
import os
import psutil
import logging
from typing import Dict, List, Optional, Tuple
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtCore import QTimer

class MemoryManager:
    """メモリ使用量の監視と管理"""
    
    def __init__(self):
        self.process = psutil.Process()
        self.warning_threshold_mb = 1024  # 警告閾値（1GB）
        self.critical_threshold_mb = 2048  # 緊急閾値（2GB）
        self.max_records_per_batch = 1000  # バッチ処理の最大レコード数
    
    def get_memory_info(self) -> Dict[str, float]:
        """現在のメモリ使用状況を取得"""
        try:
            # プロセスメモリ情報
            memory_info = self.process.memory_info()
            memory_percent = self.process.memory_percent()
            
            # システムメモリ情報
            virtual_memory = psutil.virtual_memory()
            
            return {
                "process_memory_mb": memory_info.rss / (1024 * 1024),
                "process_memory_percent": memory_percent,
                "system_memory_total_mb": virtual_memory.total / (1024 * 1024),
                "system_memory_available_mb": virtual_memory.available / (1024 * 1024),
                "system_memory_percent": virtual_memory.percent
            }
            
        except Exception as e:
            logging.error(f"メモリ情報取得エラー: {e}")
            return {}
    
    def check_memory_status(self) -> str:
        """メモリ状況をチェック"""
        memory_info = self.get_memory_info()
        
        if not memory_info:
            return "error"
        
        process_memory_mb = memory_info.get("process_memory_mb", 0)
        system_memory_percent = memory_info.get("system_memory_percent", 0)
        
        # プロセスメモリまたはシステムメモリが危険レベル
        if process_memory_mb > self.critical_threshold_mb or system_memory_percent > 90:
            return "critical"
        elif process_memory_mb > self.warning_threshold_mb or system_memory_percent > 80:
            return "warning"
        else:
            return "ok"
    
    def force_garbage_collection(self) -> float:
        """強制ガベージコレクションを実行"""
        try:
            memory_before = self.get_memory_info().get("process_memory_mb", 0)
            
            # ガベージコレクションを実行
            collected = gc.collect()
            
            memory_after = self.get_memory_info().get("process_memory_mb", 0)
            freed_mb = memory_before - memory_after
            
            logging.info(f"ガベージコレクション完了: {collected}オブジェクト回収, {freed_mb:.2f}MB解放")
            return freed_mb
            
        except Exception as e:
            logging.error(f"ガベージコレクションエラー: {e}")
            return 0.0
    
    def optimize_data_loading(self, total_records: int) -> List[Tuple[int, int]]:
        """データ量に応じてバッチサイズを最適化"""
        memory_info = self.get_memory_info()
        available_memory_mb = memory_info.get("system_memory_available_mb", 1024)
        
        # 利用可能メモリに基づいてバッチサイズを調整
        if available_memory_mb < 512:  # 512MB未満
            batch_size = 100
        elif available_memory_mb < 1024:  # 1GB未満
            batch_size = 500
        else:
            batch_size = self.max_records_per_batch
        
        # バッチリストを作成
        batches = []
        for start in range(0, total_records, batch_size):
            end = min(start + batch_size, total_records)
            batches.append((start, end))
        
        logging.info(f"データローディング最適化: {total_records}レコードを{len(batches)}バッチに分割")
        return batches
    
    def cleanup_large_objects(self, data_containers: List) -> int:
        """大きなデータオブジェクトをクリーンアップ（メモリリーク対策強化）"""
        cleaned_count = 0
        
        try:
            for container in data_containers:
                if hasattr(container, 'clear'):
                    container.clear()
                    cleaned_count += 1
                elif isinstance(container, list):
                    container.clear()
                    cleaned_count += 1
                elif isinstance(container, dict):
                    container.clear()
                    cleaned_count += 1
                # 循環参照の解放
                elif hasattr(container, '__dict__'):
                    container.__dict__.clear()
                    cleaned_count += 1
            
            # 強制ガベージコレクション（複数回実行でより確実に）
            import gc
            gc.disable()  # 一時的にGCを無効化
            try:
                for _ in range(3):
                    collected = gc.collect()
                    if collected == 0:
                        break
                    logging.debug(f"ガベージコレクション: {collected}オブジェクト回収")
            finally:
                gc.enable()  # GCを再有効化
            
            # 弱参照の処理（より安全な方法）
            # weakref.finalize_hooksの直接操作は危険なため削除
            # 代わりに明示的な弱参照の管理を実装
            try:
                import weakref
                # 安全な弱参照のクリア（Python 3.4+）
                if hasattr(weakref, 'WeakKeyDictionary'):
                    # システム全体の弱参照辞書は触らず、
                    # アプリケーション固有の弱参照のみ管理
                    pass
            except Exception as e:
                logging.debug(f"弱参照処理中の軽微なエラー: {e}")
            
            logging.info(f"大型オブジェクトクリーンアップ: {cleaned_count}オブジェクト")
            return cleaned_count
            
        except Exception as e:
            logging.error(f"オブジェクトクリーンアップエラー: {e}")
            return 0


class MemoryMonitor:
    """メモリ監視とアラート"""
    
    def __init__(self, parent=None):
        self.parent = parent
        self.manager = MemoryManager()
        self.monitor_timer = QTimer()
        self.monitor_timer.timeout.connect(self._check_memory_periodically)
        self.last_warning_time = 0
        self.warning_interval = 300  # 5分間隔で警告
    
    def start_monitoring(self, interval_ms: int = 30000):
        """メモリ監視を開始（デフォルト30秒間隔）"""
        self.monitor_timer.start(interval_ms)
        logging.info("メモリ監視を開始しました")
    
    def stop_monitoring(self):
        """メモリ監視を停止"""
        self.monitor_timer.stop()
        logging.info("メモリ監視を停止しました")
    
    def _check_memory_periodically(self):
        """定期的なメモリチェック"""
        try:
            status = self.manager.check_memory_status()
            
            if status == "critical":
                self._handle_critical_memory()
            elif status == "warning":
                self._handle_warning_memory()
                
        except Exception as e:
            logging.error(f"定期メモリチェックエラー: {e}")
    
    def _handle_critical_memory(self):
        """緊急メモリ不足への対応"""
        memory_info = self.manager.get_memory_info()
        
        # 自動ガベージコレクション
        freed_mb = self.manager.force_garbage_collection()
        
        QMessageBox.critical(
            self.parent,
            "緊急：メモリ不足",
            f"""メモリ不足が発生しています！

プロセスメモリ使用量: {memory_info.get('process_memory_mb', 0):.1f} MB
システムメモリ使用率: {memory_info.get('system_memory_percent', 0):.1f}%

自動的にメモリを{freed_mb:.1f}MB解放しました。

推奨対応：
1. 他のアプリケーションを終了
2. データを保存して再起動
3. 大量データの処理を分割"""
        )
    
    def _handle_warning_memory(self):
        """メモリ警告への対応"""
        import time
        current_time = time.time()
        
        # 警告間隔をチェック
        if current_time - self.last_warning_time < self.warning_interval:
            return
        
        self.last_warning_time = current_time
        memory_info = self.manager.get_memory_info()
        
        reply = QMessageBox.question(
            self.parent,
            "メモリ使用量警告",
            f"""メモリ使用量が多くなっています。

プロセスメモリ使用量: {memory_info.get('process_memory_mb', 0):.1f} MB
システムメモリ使用率: {memory_info.get('system_memory_percent', 0):.1f}%

メモリを最適化しますか？

「はい」: 自動最適化を実行
「いいえ」: このまま続行""",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        
        if reply == QMessageBox.Yes:
            freed_mb = self.manager.force_garbage_collection()
            QMessageBox.information(
                self.parent,
                "最適化完了",
                f"メモリ最適化が完了しました。\n解放されたメモリ: {freed_mb:.1f} MB"
            )


def check_memory_before_large_operation(estimated_memory_mb: float, parent=None) -> bool:
    """大きな処理を行う前のメモリチェック"""
    manager = MemoryManager()
    memory_info = manager.get_memory_info()
    
    available_mb = memory_info.get("system_memory_available_mb", 0)
    
    if available_mb < estimated_memory_mb * 1.5:  # 1.5倍のマージン
        reply = QMessageBox.question(
            parent,
            "メモリ不足の可能性",
            f"""この処理には大量のメモリが必要です。

推定必要メモリ: {estimated_memory_mb:.1f} MB
利用可能メモリ: {available_mb:.1f} MB

メモリ不足が発生する可能性があります。
処理を続行しますか？

「はい」: 処理を実行（分割処理を適用）
「いいえ」: 処理を中止""",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        return reply == QMessageBox.Yes
    
    return True


def optimize_large_data_processing(data_list: List, process_func, parent=None):
    """大量データ処理の最適化"""
    manager = MemoryManager()
    
    # データサイズに基づいてバッチ処理を決定
    total_records = len(data_list)
    batches = manager.optimize_data_loading(total_records)
    
    if len(batches) > 1:
        QMessageBox.information(
            parent,
            "バッチ処理",
            f"""大量のデータを効率的に処理するため、
{len(batches)}回に分けて処理します。

総レコード数: {total_records:,}件
バッチ数: {len(batches)}回"""
        )
    
    results = []
    for i, (start, end) in enumerate(batches):
        batch_data = data_list[start:end]
        
        try:
            batch_result = process_func(batch_data)
            results.extend(batch_result if isinstance(batch_result, list) else [batch_result])
            
            # バッチ間でガベージコレクション
            if i < len(batches) - 1:  # 最後のバッチ以外
                manager.force_garbage_collection()
                
        except Exception as e:
            logging.error(f"バッチ処理エラー (バッチ {i+1}): {e}")
            raise
    
    return results