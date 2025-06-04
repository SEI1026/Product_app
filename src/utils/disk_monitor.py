"""
ディスク容量監視システム - 容量不足の検出と対策
"""

import os
import shutil
import psutil
import logging
from pathlib import Path
from typing import Dict, List, Tuple
from PyQt5.QtWidgets import QMessageBox

class DiskSpaceMonitor:
    """ディスク容量の監視と管理"""
    
    def __init__(self):
        self.min_free_space_mb = 100  # 最小必要容量（MB）
        self.warning_threshold_mb = 500  # 警告閾値（MB）
        self.critical_threshold_mb = 200  # 緊急閾値（MB）
    
    def check_disk_space(self, path: str) -> Dict[str, any]:
        """指定パスのディスク容量をチェック"""
        try:
            disk_usage = shutil.disk_usage(path)
            total_bytes = disk_usage.total
            used_bytes = disk_usage.used
            free_bytes = disk_usage.free
            
            total_mb = total_bytes / (1024 * 1024)
            used_mb = used_bytes / (1024 * 1024)
            free_mb = free_bytes / (1024 * 1024)
            
            usage_percent = (used_bytes / total_bytes) * 100
            
            # 状況判定
            if free_mb < self.critical_threshold_mb:
                status = "critical"
            elif free_mb < self.warning_threshold_mb:
                status = "warning"
            else:
                status = "ok"
            
            return {
                "status": status,
                "total_mb": total_mb,
                "used_mb": used_mb,
                "free_mb": free_mb,
                "usage_percent": usage_percent,
                "path": path
            }
            
        except Exception as e:
            logging.error(f"ディスク容量チェックエラー: {e}")
            return {"status": "error", "error": str(e)}
    
    def check_required_space(self, file_path: str, estimated_size_mb: float) -> bool:
        """必要な容量があるかチェック"""
        try:
            disk_info = self.check_disk_space(os.path.dirname(file_path))
            if disk_info["status"] == "error":
                return False
            
            return disk_info["free_mb"] >= (estimated_size_mb + self.min_free_space_mb)
            
        except Exception as e:
            logging.error(f"必要容量チェックエラー: {e}")
            return False
    
    def estimate_file_size(self, data_count: int, avg_record_size_kb: float = 2.0) -> float:
        """データ量から推定ファイルサイズを計算（MB）"""
        estimated_kb = data_count * avg_record_size_kb
        return estimated_kb / 1024  # MB変換
    
    def find_cleanup_candidates(self, directory: str) -> List[Dict]:
        """クリーンアップ対象ファイルを検出"""
        cleanup_candidates = []
        
        try:
            dir_path = Path(directory)
            if not dir_path.exists():
                return cleanup_candidates
            
            # 一時ファイル
            temp_patterns = ["*.tmp", "*.temp", "*~", "*.bak"]
            for pattern in temp_patterns:
                for file_path in dir_path.rglob(pattern):
                    if file_path.is_file():
                        size_mb = file_path.stat().st_size / (1024 * 1024)
                        cleanup_candidates.append({
                            "path": str(file_path),
                            "size_mb": size_mb,
                            "type": "temp",
                            "description": "一時ファイル"
                        })
            
            # 古いログファイル
            log_files = list(dir_path.rglob("*.log"))
            for log_file in log_files:
                if log_file.is_file():
                    size_mb = log_file.stat().st_size / (1024 * 1024)
                    if size_mb > 10:  # 10MB以上のログファイル
                        cleanup_candidates.append({
                            "path": str(log_file),
                            "size_mb": size_mb,
                            "type": "log",
                            "description": "大きなログファイル"
                        })
            
            # 古いバックアップファイル
            backup_patterns = ["*.backup_*", "*.bak"]
            for pattern in backup_patterns:
                for backup_file in dir_path.rglob(pattern):
                    if backup_file.is_file():
                        size_mb = backup_file.stat().st_size / (1024 * 1024)
                        cleanup_candidates.append({
                            "path": str(backup_file),
                            "size_mb": size_mb,
                            "type": "backup",
                            "description": "古いバックアップファイル"
                        })
            
        except Exception as e:
            logging.error(f"クリーンアップ候補検索エラー: {e}")
        
        # サイズの大きい順にソート
        cleanup_candidates.sort(key=lambda x: x["size_mb"], reverse=True)
        return cleanup_candidates
    
    def perform_cleanup(self, file_paths: List[str]) -> Tuple[bool, float]:
        """指定されたファイルをクリーンアップ"""
        total_freed_mb = 0.0
        success = True
        
        for file_path in file_paths:
            try:
                if os.path.exists(file_path):
                    size_mb = os.path.getsize(file_path) / (1024 * 1024)
                    os.remove(file_path)
                    total_freed_mb += size_mb
                    logging.info(f"クリーンアップ: {file_path} ({size_mb:.2f}MB)")
                    
            except Exception as e:
                logging.error(f"ファイル削除エラー {file_path}: {e}")
                success = False
        
        return success, total_freed_mb


def check_disk_space_before_save(file_path: str, estimated_data_count: int, parent=None) -> bool:
    """保存前のディスク容量チェック"""
    monitor = DiskSpaceMonitor()
    
    # ファイルサイズを推定
    estimated_size_mb = monitor.estimate_file_size(estimated_data_count)
    
    # 容量チェック
    if not monitor.check_required_space(file_path, estimated_size_mb):
        disk_info = monitor.check_disk_space(os.path.dirname(file_path))
        
        if disk_info["status"] == "error":
            QMessageBox.critical(
                parent,
                "ディスク容量エラー",
                "ディスク容量の確認中にエラーが発生しました。"
            )
            return False
        
        # 容量不足の警告
        reply = QMessageBox.question(
            parent,
            "ディスク容量不足",
            f"""ディスク容量が不足しています。

現在の空き容量: {disk_info['free_mb']:.1f} MB
必要な容量: {estimated_size_mb + monitor.min_free_space_mb:.1f} MB
不足分: {(estimated_size_mb + monitor.min_free_space_mb) - disk_info['free_mb']:.1f} MB

クリーンアップを実行しますか？

「はい」: 不要ファイルを自動削除
「いいえ」: 保存を中止""",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        
        if reply == QMessageBox.Yes:
            return perform_disk_cleanup(os.path.dirname(file_path), parent)
        else:
            return False
    
    return True


def perform_disk_cleanup(directory: str, parent=None) -> bool:
    """ディスククリーンアップの実行"""
    monitor = DiskSpaceMonitor()
    
    # クリーンアップ候補を検索
    candidates = monitor.find_cleanup_candidates(directory)
    
    if not candidates:
        QMessageBox.information(
            parent,
            "クリーンアップ",
            "削除可能な不要ファイルが見つかりませんでした。\n手動でディスク容量を確保してください。"
        )
        return False
    
    # ユーザーに確認
    file_list = "\n".join([
        f"• {os.path.basename(c['path'])} ({c['size_mb']:.1f}MB) - {c['description']}"
        for c in candidates[:10]  # 上位10件を表示
    ])
    
    total_size = sum(c["size_mb"] for c in candidates)
    
    reply = QMessageBox.question(
        parent,
        "ファイル削除の確認",
        f"""以下のファイルを削除してディスク容量を確保します：

{file_list}

合計削除サイズ: {total_size:.1f} MB

削除を実行しますか？

注意: 削除したファイルは復元できません。""",
        QMessageBox.Yes | QMessageBox.No,
        QMessageBox.Yes
    )
    
    if reply == QMessageBox.Yes:
        file_paths = [c["path"] for c in candidates]
        success, freed_mb = monitor.perform_cleanup(file_paths)
        
        if success:
            QMessageBox.information(
                parent,
                "クリーンアップ完了",
                f"クリーンアップが完了しました。\n解放された容量: {freed_mb:.1f} MB"
            )
            return True
        else:
            QMessageBox.warning(
                parent,
                "クリーンアップ警告",
                f"一部のファイルの削除に失敗しました。\n解放された容量: {freed_mb:.1f} MB"
            )
            return freed_mb > 0
    
    return False


def monitor_disk_space_continuously(paths: List[str], parent=None):
    """継続的なディスク容量監視"""
    monitor = DiskSpaceMonitor()
    
    for path in paths:
        disk_info = monitor.check_disk_space(path)
        
        if disk_info["status"] == "critical":
            QMessageBox.critical(
                parent,
                "緊急：ディスク容量不足",
                f"""ディスク容量が緊急レベルまで減少しています！

パス: {path}
空き容量: {disk_info['free_mb']:.1f} MB
使用率: {disk_info['usage_percent']:.1f}%

即座にファイルを削除するか、別のドライブに移動してください。
アプリケーションが正常に動作しない可能性があります。"""
            )
        elif disk_info["status"] == "warning":
            QMessageBox.warning(
                parent,
                "警告：ディスク容量少",
                f"""ディスク容量が少なくなっています。

パス: {path}
空き容量: {disk_info['free_mb']:.1f} MB
使用率: {disk_info['usage_percent']:.1f}%

不要なファイルを削除することをお勧めします。"""
            )