"""
バージョンチェッカー - GitHub上の最新バージョンを確認し、自動更新を管理
"""

import json
import logging
import os
import sys
import subprocess
import tempfile
import zipfile
import shutil
import urllib.error
import webbrowser
import re
import time
import hashlib
import requests
from typing import Optional, Dict, Any, Tuple
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError
from PyQt5.QtCore import QThread, pyqtSignal, QObject, QTimer
from PyQt5.QtWidgets import QMessageBox, QProgressDialog, QPushButton, QApplication


def safe_extract_zip(zip_file_path: str, extract_to: str) -> str:
    """
    安全なZIPファイル抽出 - ディレクトリトラバーサル攻撃を防止
    
    Args:
        zip_file_path: 抽出するZIPファイルのパス
        extract_to: 抽出先ディレクトリ
    
    Returns:
        str: 抽出先ディレクトリのパス
        
    Raises:
        ValueError: 危険なファイルパスが検出された場合
        Exception: その他のエラー
    """
    # 抽出先ディレクトリを正規化
    extract_to = os.path.abspath(extract_to)
    
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        # ZIPファイルの整合性をチェック
        bad_file = zip_ref.testzip()
        if bad_file:
            raise Exception(f"ZIPファイルが破損しています: {bad_file}")
        
        # 各ファイルのパスを検証
        for member in zip_ref.infolist():
            # パスを正規化
            normalized_path = os.path.normpath(member.filename)
            
            # 危険なパスをチェック
            if (normalized_path.startswith('/') or 
                normalized_path.startswith('\\') or 
                '..' in normalized_path or
                ':' in normalized_path):
                raise ValueError(f"危険なファイルパス検出: {member.filename}")
            
            # 抽出先の完全パスを計算
            full_extract_path = os.path.abspath(os.path.join(extract_to, normalized_path))
            
            # 抽出先が指定ディレクトリ内にあることを確認
            if not full_extract_path.startswith(extract_to):
                raise ValueError(f"ディレクトリトラバーサル攻撃を検出: {member.filename}")
            
            # ファイル名の長さ制限
            if len(normalized_path) > 255:
                raise ValueError(f"ファイル名が長すぎます: {member.filename}")
        
        # 全て安全であれば抽出を実行
        zip_ref.extractall(extract_to)
        logging.info(f"ZIPファイルを安全に抽出: {extract_to}")
        
    return extract_to

# 現在のアプリケーションバージョン
CURRENT_VERSION = "2.7.1"

# GitHub上のversion.jsonのURL
# 株式会社大宝家具の商品登録入力ツール
# ※実際のGitHubリポジトリURLに合わせて変更してください
VERSION_CHECK_URL = "https://raw.githubusercontent.com/SEI1026/Product_app/main/version.json"


class VersionInfo:
    """バージョン情報を格納するクラス"""
    
    def __init__(self, version_data: Dict[str, Any]):
        self.version = self._sanitize_string(version_data.get("version", "0.0.0"))
        self.release_date = self._sanitize_string(version_data.get("release_date", ""))
        self.download_url = self._sanitize_string(version_data.get("download_url", ""))
        self.changelog = self._sanitize_changelog(version_data.get("changelog", {}))
        self.minimum_required_version = self._sanitize_string(version_data.get("minimum_required_version", "0.0.0"))
    
    def _sanitize_string(self, text: str) -> str:
        """文字列から不正な文字を除去"""
        if not isinstance(text, str):
            return str(text)
        # 文字化けや不正な文字を置換
        import re
        # 制御文字を除去
        text = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', text)
        # 不正なUnicode文字を?に置換
        text = text.encode('utf-8', errors='replace').decode('utf-8')
        return text
    
    def _sanitize_changelog(self, changelog: Dict[str, Any]) -> Dict[str, Any]:
        """変更履歴の文字列を全てサニタイズ"""
        sanitized = {}
        for version, changes in changelog.items():
            sanitized[version] = {}
            for category, items in changes.items():
                if isinstance(items, list):
                    sanitized[version][category] = [self._sanitize_string(item) for item in items]
                else:
                    sanitized[version][category] = self._sanitize_string(items)
        return sanitized
        
    def get_latest_changes(self) -> str:
        """最新バージョンの変更点を取得"""
        if self.version in self.changelog:
            changes = self.changelog[self.version]
            features = changes.get("features", [])
            improvements = changes.get("improvements", [])
            bug_fixes = changes.get("bug_fixes", [])
            
            result = []
            if features:
                result.append("【新機能】")
                result.extend([f"• {f}" for f in features])
            if improvements:
                result.append("\n【改善点】")
                result.extend([f"• {i}" for i in improvements])
            if bug_fixes:
                result.append("\n【バグ修正】")
                result.extend([f"• {b}" for b in bug_fixes])
                
            return "\n".join(result)
        return "変更点の情報がありません"


class UpdateDownloader(QThread):
    """バックグラウンドで更新をダウンロードするスレッド"""
    
    progress = pyqtSignal(int)  # ダウンロード進捗
    status = pyqtSignal(str)    # ステータスメッセージ
    finished = pyqtSignal(bool, str)  # 完了シグナル（成功/失敗, メッセージ）
    
    def __init__(self, download_url: str, target_dir: str):
        super().__init__()
        self.download_url = download_url
        self.target_dir = target_dir
        self.temp_file = None
        self._cancelled = False
        self.extract_dir = None
        
        logging.info(f"UpdateDownloader初期化: URL={download_url}, ターゲット={target_dir}")
        
    def run(self):
        """更新ファイルをダウンロードして展開"""
        step = "初期化"
        crash_log_file = None
        
        try:
            # クラッシュログファイルを作成
            import tempfile
            crash_log_file = os.path.join(tempfile.gettempdir(), f"update_crash_log_{os.getpid()}.txt")
            
            logging.info("=== 自動更新プロセス開始 ===")
            self._write_crash_log(crash_log_file, f"=== 自動更新プロセス開始 ===\n開始時刻: {self._get_timestamp()}\n")
            
            step = "キャンセルチェック"
            if self._cancelled:
                logging.info("ダウンロード開始前にキャンセルされました")
                self._write_crash_log(crash_log_file, f"ステップ: {step} - キャンセルされました\n")
                return
                
            # 一時ファイル作成
            step = "一時ファイル作成"
            temp_dir = tempfile.gettempdir()
            self.temp_file = os.path.join(temp_dir, f'update_{os.getpid()}.zip')
            logging.info(f"一時ファイル: {self.temp_file}")
            self._write_crash_log(crash_log_file, f"ステップ: {step} - 一時ファイル: {self.temp_file}\n")
            
            step = "URL検証"
            logging.info(f"ダウンロード開始: {self.download_url}")
            self._write_crash_log(crash_log_file, f"ステップ: {step} - URL: {self.download_url}\n")
            
            # 🚨 URL検証 - 詳細確認
            logging.info(f"🔍 ダウンロードURL詳細確認:")
            logging.info(f"  📋 URL: {self.download_url}")
            logging.info(f"  ✅ URL存在: {bool(self.download_url)}")
            logging.info(f"  ✅ HTTPS: {self.download_url.startswith('https://') if self.download_url else False}")
            logging.info(f"  📏 URL長: {len(self.download_url) if self.download_url else 0}")
            
            if not self.download_url or not self.download_url.startswith('https://'):
                error_msg = f"無効なダウンロードURL: {self.download_url}"
                logging.error(error_msg)
                self._write_crash_log(crash_log_file, f"エラー: {error_msg}\n")
                self.finished.emit(False, "無効なダウンロードURLです")
                return
                
            # ダウンロード
            step = "ダウンロード"
            self.status.emit("更新ファイルをダウンロード中...")
            logging.info("ダウンロード処理開始")
            self._write_crash_log(crash_log_file, f"ステップ: {step} - ダウンロード開始\n")
            success = self._download_file()
            
            if not success or self._cancelled:
                error_msg = "ダウンロードがキャンセルまたは失敗しました"
                logging.warning(error_msg)
                self._write_crash_log(crash_log_file, f"エラー: {error_msg}\n")
                return
            
            # ハッシュ検証
            step = "ハッシュ検証"
            self.status.emit("ダウンロードファイルを検証中...")
            logging.info("ハッシュ検証処理開始")
            self._write_crash_log(crash_log_file, f"ステップ: {step} - ハッシュ検証開始\n")
            
            hash_success = self._verify_file_hash()
            if not hash_success or self._cancelled:
                error_msg = "ファイルのハッシュ検証に失敗しました"
                logging.error(error_msg)
                self._write_crash_log(crash_log_file, f"エラー: {error_msg}\n")
                self.finished.emit(False, "ダウンロードファイルの整合性検証に失敗しました")
                return
            
            logging.info("ハッシュ検証完了")
            self._write_crash_log(crash_log_file, f"ハッシュ検証完了\n")
                
            # 展開
            step = "ZIP展開"
            self.status.emit("更新ファイルを展開中...")
            logging.info("ZIP展開処理開始")
            self._write_crash_log(crash_log_file, f"ステップ: {step} - ZIP展開開始\n")
            self.extract_dir = self._extract_zip()
            logging.info(f"展開完了: {self.extract_dir}")
            self._write_crash_log(crash_log_file, f"展開完了: {self.extract_dir}\n")
            
            if self._cancelled:
                logging.info("展開後にキャンセルされました") 
                self._write_crash_log(crash_log_file, f"ステップ: {step} - キャンセルされました\n")
                return
                
            # ファイル更新
            step = "ファイル更新"
            self.status.emit("ファイルを更新中...")
            logging.info(f"ファイル更新開始: extract_dir={self.extract_dir}, target_dir={self.target_dir}")
            self._write_crash_log(crash_log_file, f"ステップ: {step} - ファイル更新開始\n")
            self._write_crash_log(crash_log_file, f"展開ディレクトリ: {self.extract_dir}\n")
            self._write_crash_log(crash_log_file, f"ターゲットディレクトリ: {self.target_dir}\n")
            
            # === 🚨 包括的デバッグ情報収集 ===
            logging.info("=" * 80)
            logging.info("🚨🚨🚨 包括的デバッグ情報収集開始 🚨🚨🚨")
            logging.info("=" * 80)
            
            # sys.frozen状態を確認
            is_frozen = getattr(sys, 'frozen', False)
            logging.info(f"📊 sys.frozen: {is_frozen}")
            logging.info(f"📊 sys.executable: {sys.executable}")
            logging.info(f"📊 sys.argv[0]: {sys.argv[0] if sys.argv else 'None'}")
            logging.info(f"📊 現在の作業ディレクトリ: {os.getcwd()}")
            
            # 展開ディレクトリの完全構造表示
            logging.info("🔍 展開ディレクトリの完全構造:")
            try:
                for root, dirs, files in os.walk(self.extract_dir):
                    level = root.replace(self.extract_dir, '').count(os.sep)
                    indent = '  ' * level
                    rel_path = os.path.relpath(root, self.extract_dir)
                    logging.info(f"{indent}📁 {rel_path}/ ({len(files)} files, {len(dirs)} dirs)")
                    
                    for file in files:
                        file_path = os.path.join(root, file)
                        file_size = os.path.getsize(file_path)
                        if file.endswith('.exe'):
                            logging.info(f"{indent}  🔥🔥 EXE: {file} ({file_size:,} bytes) 🔥🔥")
                        else:
                            logging.info(f"{indent}  📄 {file} ({file_size:,} bytes)")
            except Exception as e:
                logging.error(f"❌ 展開ディレクトリ構造表示エラー: {e}")
            
            # ターゲットディレクトリの確認
            logging.info(f"🎯 ターゲットディレクトリ確認: {self.target_dir}")
            logging.info(f"🎯 ターゲットディレクトリ存在: {os.path.exists(self.target_dir)}")
            if os.path.exists(self.target_dir):
                try:
                    target_files = os.listdir(self.target_dir)
                    logging.info(f"🎯 ターゲットディレクトリ内容: {len(target_files)} items")
                    for item in target_files[:10]:  # 最初の10個
                        item_path = os.path.join(self.target_dir, item)
                        if os.path.isfile(item_path):
                            size = os.path.getsize(item_path)
                            logging.info(f"  📄 {item} ({size:,} bytes)")
                        else:
                            logging.info(f"  📁 {item}/")
                except Exception as e:
                    logging.error(f"❌ ターゲットディレクトリ内容確認エラー: {e}")
            
            # .newファイルの事前確認
            expected_new_file = os.path.join(self.target_dir, '商品登録入力ツール.exe.new')
            logging.info(f"🔍 事前.newファイル確認: {expected_new_file}")
            logging.info(f"🔍 事前.newファイル存在: {os.path.exists(expected_new_file)}")
            
            logging.info("=" * 80)
            
            try:
                # 環境情報をログに記録
                self._log_system_info()
                
                # 展開されたディレクトリ内から実際の更新ファイルがあるディレクトリを特定
                actual_source_dir = self._find_actual_source_directory(self.extract_dir)
                self._write_crash_log(crash_log_file, f"実際のソースディレクトリ: {actual_source_dir}\n")
                logging.info(f"実際のソースディレクトリ: {actual_source_dir}")
                
                self._update_files(actual_source_dir, self.target_dir)
                logging.info("ファイル更新が正常に完了しました")
                self._write_crash_log(crash_log_file, f"ファイル更新正常完了\n")
            except Exception as update_error:
                error_msg = f"ファイル更新中にエラーが発生: {update_error}"
                logging.error(error_msg, exc_info=True)
                self._write_crash_log(crash_log_file, f"重大エラー: {error_msg}\n")
                self._write_crash_log(crash_log_file, f"エラータイプ: {type(update_error).__name__}\n")
                self._write_crash_log(crash_log_file, f"エラー詳細: {str(update_error)}\n")
                
                # 詳細な環境情報も含めてエラー報告
                error_details = self._collect_error_context(update_error)
                self._write_crash_log(crash_log_file, f"コンテキスト情報:\n{error_details}\n")
                
                # クラッシュログの場所をエラーメッセージに含める
                self.finished.emit(False, f"ファイル更新エラー（{step}）: {update_error}\n\n{error_details}\n\nクラッシュログ: {crash_log_file}")
                return
            
            step = "完了処理"
            if not self._cancelled:
                logging.info("更新処理が完全に完了 - 成功シグナル送信")
                self._write_crash_log(crash_log_file, f"ステップ: {step} - 成功完了\n")
                self.finished.emit(True, "更新が正常に完了しました")
                logging.info("更新処理が完全に完了しました")
            
        except urllib.error.URLError as e:
            error_msg = f"ネットワークエラー（{step}）: {e}"
            logging.error(error_msg)
            if crash_log_file:
                self._write_crash_log(crash_log_file, f"ネットワークエラー: {error_msg}\n")
            self.finished.emit(False, f"{error_msg}\n\nクラッシュログ: {crash_log_file}")
        except zipfile.BadZipFile as e:
            error_msg = f"ZIPファイルエラー（{step}）: {e}"
            logging.error(error_msg)
            if crash_log_file:
                self._write_crash_log(crash_log_file, f"ZIPエラー: {error_msg}\n")
            self.finished.emit(False, f"{error_msg}\n\nクラッシュログ: {crash_log_file}")
        except PermissionError as e:
            error_msg = f"ファイルアクセスエラー（{step}）: {e}"
            logging.error(error_msg)
            if crash_log_file:
                self._write_crash_log(crash_log_file, f"権限エラー: {error_msg}\n")
            self.finished.emit(False, f"{error_msg}\n\nクラッシュログ: {crash_log_file}")
        except Exception as e:
            error_msg = f"予期しないエラー（{step}）: {e}"
            logging.error(f"更新エラー: {error_msg}", exc_info=True)
            if crash_log_file:
                self._write_crash_log(crash_log_file, f"予期しないエラー: {error_msg}\nエラータイプ: {type(e).__name__}\n")
                import traceback
                self._write_crash_log(crash_log_file, f"トレースバック:\n{traceback.format_exc()}\n")
            self.finished.emit(False, f"{error_msg}\n\nクラッシュログ: {crash_log_file}")
            
        finally:
            self._cleanup()
    
    def _download_file(self):
        """ファイルダウンロード処理"""
        response = None
        file_handle = None
        
        try:
            req = Request(self.download_url, headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            })
            
            response = urlopen(req, timeout=30)
            
            if response.getcode() != 200:
                raise Exception(f"HTTPエラー: {response.getcode()}")
            
            total_size = int(response.headers.get('Content-Length', 0))
            downloaded = 0
            
            file_handle = open(self.temp_file, 'wb')
            
            while not self._cancelled:
                try:
                    chunk = response.read(8192)
                    if not chunk:
                        break
                        
                    file_handle.write(chunk)
                    downloaded += len(chunk)
                    
                    if total_size > 0:
                        progress = min(int((downloaded / total_size) * 100), 100)
                        self.progress.emit(progress)
                        
                except Exception as e:
                    logging.error(f"チャンク読み込みエラー: {e}")
                    raise
                    
            # キャンセルされた場合
            if self._cancelled:
                logging.info("ダウンロードがキャンセルされました")
                return False
                
            # ダウンロード完了確認
            if total_size > 0 and downloaded < total_size:
                raise Exception(f"ダウンロード不完全: {downloaded}/{total_size} bytes")
                
            logging.info(f"ダウンロード完了: {downloaded} bytes")
            return True
            
        except Exception as e:
            logging.error(f"ダウンロードエラー: {e}")
            raise
            
        finally:
            # リソースの確実なクリーンアップ
            if file_handle:
                try:
                    file_handle.close()
                except Exception as e:
                    logging.warning(f"ファイルクローズエラー: {e}")
                    
            if response:
                try:
                    response.close()
                except Exception as e:
                    logging.warning(f"レスポンスクローズエラー: {e}")
    
    def _verify_file_hash(self):
        """ダウンロードファイルのSHA-256ハッシュを検証"""
        if not os.path.exists(self.temp_file):
            logging.error("ハッシュ検証: ファイルが存在しません")
            return False
            
        try:
            # 期待されるハッシュ値を取得
            expected_hash = None
            if hasattr(self, 'version_info') and self.version_info:
                expected_hash = self.version_info.get('file_hash_sha256')
            
            if not expected_hash or expected_hash == "example_hash_placeholder_to_be_replaced_with_actual_hash":
                logging.warning("ハッシュ検証: 期待ハッシュが設定されていません - スキップします")
                return True  # 後方互換性のため、ハッシュがない場合は成功とする
            
            # ファイルのハッシュを計算
            sha256_hash = hashlib.sha256()
            with open(self.temp_file, 'rb') as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    sha256_hash.update(chunk)
            
            calculated_hash = sha256_hash.hexdigest()
            
            # ハッシュを比較
            if calculated_hash.lower() == expected_hash.lower():
                logging.info("ハッシュ検証: 成功")
                return True
            else:
                logging.error(f"ハッシュ検証: 失敗 - 期待値: {expected_hash}, 実際値: {calculated_hash}")
                return False
                
        except Exception as e:
            logging.error(f"ハッシュ検証エラー: {e}")
            return False
    
    def _extract_zip(self):
        """ZIPファイルの展開"""
        if not os.path.exists(self.temp_file):
            raise Exception("ダウンロードファイルが見つかりません")
            
        file_size = os.path.getsize(self.temp_file)
        if file_size < 1000:
            raise Exception(f"ダウンロードファイルが不完全です（{file_size} bytes）")
        
        extract_dir = tempfile.mkdtemp(prefix='update_extract_')
        safe_extract_zip(self.temp_file, extract_dir)
        return extract_dir
    
    def _cleanup(self):
        """クリーンアップ処理"""
        logging.info("アップデートクリーンアップ開始")
        
        # 一時ファイルを削除
        if self.temp_file and os.path.exists(self.temp_file):
            try:
                file_size = os.path.getsize(self.temp_file)
                os.unlink(self.temp_file)
                logging.info(f"一時ファイル削除完了: {self.temp_file} ({file_size} bytes)")
            except Exception as e:
                logging.warning(f"一時ファイル削除エラー: {e}")
        
        # 展開ディレクトリを削除
        if self.extract_dir and os.path.exists(self.extract_dir):
            try:
                shutil.rmtree(self.extract_dir)
                logging.info(f"展開ディレクトリ削除完了: {self.extract_dir}")
            except Exception as e:
                logging.warning(f"展開ディレクトリ削除エラー: {e}")
                
        # キャンセル時の中途半端なファイルをクリーンアップ
        if self._cancelled and hasattr(self, 'target_dir'):
            try:
                self._cleanup_partial_files()
            except Exception as e:
                logging.warning(f"部分ファイルクリーンアップエラー: {e}")
                
        logging.info("アップデートクリーンアップ完了")
    
    def _cleanup_partial_files(self):
        """中途半端なファイル（.newファイルなど）をクリーンアップ"""
        if not hasattr(self, 'target_dir') or not self.target_dir:
            return
            
        try:
            # .newファイルを検索して削除
            import glob
            new_files = glob.glob(os.path.join(self.target_dir, "*.new"))
            for new_file in new_files:
                try:
                    file_size = os.path.getsize(new_file)
                    os.remove(new_file)
                    logging.info(f"中途半端な.newファイルを削除: {new_file} ({file_size} bytes)")
                except Exception as e:
                    logging.warning(f".newファイル削除エラー {new_file}: {e}")
                    
        except Exception as e:
            logging.error(f"部分ファイルクリーンアップエラー: {e}")
    
    def _copy_large_file(self, source_file: str, target_file: str, file_size: int):
        """大きなファイルを安全にコピー（チャンク方式）"""
        try:
            logging.info(f"大きなファイルのチャンクコピー開始: {source_file} -> {target_file} ({file_size} bytes)")
            
            chunk_size = 1024 * 1024  # 1MBずつコピー
            copied = 0
            
            with open(source_file, 'rb') as src, open(target_file, 'wb') as dst:
                while copied < file_size:
                    if self._cancelled:
                        logging.info("大きなファイルコピー中にキャンセルされました")
                        return False
                    
                    # チャンクサイズを調整（残りサイズが小さい場合）
                    current_chunk_size = min(chunk_size, file_size - copied)
                    chunk = src.read(current_chunk_size)
                    
                    if not chunk:
                        break
                    
                    dst.write(chunk)
                    copied += len(chunk)
                    
                    # 進捗更新（5%刻み）
                    progress = int((copied / file_size) * 100)
                    if progress % 5 == 0:
                        self.status.emit(f"大きなファイルをコピー中: {progress}% ({copied/1024/1024:.1f}MB/{file_size/1024/1024:.1f}MB)")
                    
                    # 少し待機してCPU負荷を軽減
                    import time
                    time.sleep(0.001)  # 1ms待機
            
            # コピー完了確認
            if copied != file_size:
                raise Exception(f"ファイルコピーが不完全: {copied}/{file_size} bytes")
            
            logging.info(f"大きなファイルのチャンクコピー完了: {copied} bytes")
            return True
            
        except Exception as e:
            logging.error(f"大きなファイルコピーエラー: {e}")
            # 失敗した場合は部分ファイルを削除
            if os.path.exists(target_file):
                try:
                    os.remove(target_file)
                    logging.info(f"失敗した大きなファイルを削除: {target_file}")
                except:
                    pass
            raise
    
    def _write_crash_log(self, crash_log_file, message):
        """クラッシュログファイルに情報を書き込み"""
        try:
            if crash_log_file:
                with open(crash_log_file, 'a', encoding='utf-8') as f:
                    f.write(message)
                    f.flush()  # 即座にディスクに書き込み
        except Exception as e:
            # ログファイル書き込み失敗はサイレントに処理
            pass
    
    def _get_timestamp(self):
        """現在時刻の文字列を取得"""
        import datetime
        return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    def _log_system_info(self):
        """システム情報をログに記録"""
        try:
            import psutil
            import platform
            
            # メモリ情報
            memory = psutil.virtual_memory()
            logging.info(f"メモリ情報: 総容量={memory.total/1024/1024/1024:.1f}GB, "
                        f"使用量={memory.percent}%, 利用可能={memory.available/1024/1024/1024:.1f}GB")
            
            # ディスク容量
            disk = psutil.disk_usage(self.target_dir)
            logging.info(f"ディスク容量: 総容量={disk.total/1024/1024/1024:.1f}GB, "
                        f"使用量={disk.used/1024/1024/1024:.1f}GB, 空き容量={disk.free/1024/1024/1024:.1f}GB")
            
            # OS情報
            logging.info(f"OS情報: {platform.system()} {platform.release()} {platform.version()}")
            
            # プロセス情報
            current_process = psutil.Process()
            logging.info(f"プロセス情報: メモリ使用量={current_process.memory_info().rss/1024/1024:.1f}MB, "
                        f"CPU使用率={current_process.cpu_percent()}%")
            
        except ImportError:
            logging.info("psutilが利用できません - 基本的なシステム情報のみ記録")
            try:
                # psutilなしでも取得できる情報
                disk_free = shutil.disk_usage(self.target_dir).free
                logging.info(f"ディスク空き容量: {disk_free/1024/1024/1024:.1f}GB")
            except:
                pass
        except Exception as e:
            logging.warning(f"システム情報取得エラー: {e}")
    
    def _collect_error_context(self, error):
        """エラー発生時のコンテキスト情報を収集"""
        try:
            context_info = []
            
            # エラーの種類
            context_info.append(f"エラータイプ: {type(error).__name__}")
            
            # ディスク容量確認
            try:
                disk = shutil.disk_usage(self.target_dir)
                free_gb = disk.free / 1024 / 1024 / 1024
                context_info.append(f"ディスク空き容量: {free_gb:.1f}GB")
                if free_gb < 0.5:  # 500MB未満
                    context_info.append("⚠️ ディスク容量不足の可能性")
            except:
                context_info.append("ディスク容量取得失敗")
            
            # 権限確認
            try:
                if not os.access(self.target_dir, os.W_OK):
                    context_info.append("⚠️ ターゲットディレクトリへの書き込み権限なし")
                else:
                    context_info.append("✓ ターゲットディレクトリへの書き込み権限あり")
            except:
                context_info.append("権限確認失敗")
            
            # 一時ファイル確認
            try:
                if self.temp_file and os.path.exists(self.temp_file):
                    temp_size = os.path.getsize(self.temp_file)
                    context_info.append(f"一時ファイル: {temp_size/1024/1024:.1f}MB")
                else:
                    context_info.append("⚠️ 一時ファイルが見つからない")
            except:
                context_info.append("一時ファイル確認失敗")
            
            # 展開ディレクトリ確認
            try:
                if self.extract_dir and os.path.exists(self.extract_dir):
                    extracted_files = []
                    for root, dirs, files in os.walk(self.extract_dir):
                        extracted_files.extend(files)
                    context_info.append(f"展開済みファイル数: {len(extracted_files)}個")
                else:
                    context_info.append("⚠️ 展開ディレクトリが見つからない")
            except:
                context_info.append("展開ディレクトリ確認失敗")
            
            return "\n".join(context_info)
            
        except Exception as e:
            return f"コンテキスト情報収集エラー: {e}"
    
    def terminate(self):
        """ダウンロードをキャンセル"""
        logging.info("アップデートダウンロードのキャンセルが要求されました")
        self._cancelled = True
        
        # スレッドの安全な終了を待つ
        if self.isRunning():
            # 最大5秒待つ
            if not self.wait(5000):
                logging.warning("ダウンロードスレッドが5秒以内に終了しませんでした")
                # 強制終了は避ける（危険なため）
    
    def _update_files(self, source_dir: str, target_dir: str):
        """ファイルを更新（実行中のファイルは.newとして保存、ユーザーデータは保護）"""
        try:
            # 実際の実行ファイルパスを取得
            if getattr(sys, 'frozen', False):
                # PyInstallerでビルドされた場合
                current_exe = os.path.join(target_dir, '商品登録入力ツール.exe')
            else:
                current_exe = os.path.abspath(sys.executable)
                
            current_exe_name = os.path.basename(current_exe)
            updated_exe = False
            
            logging.info(f"ファイル更新開始: {source_dir} -> {target_dir}")
            logging.info(f"現在の実行ファイル: {current_exe}")
            logging.info(f"実行ファイル名: {current_exe_name}")
            
            # ソースディレクトリの構造をログ出力
            logging.info("=== ファイル更新詳細ログ ===")
            logging.info(f"ソースディレクトリ: {source_dir}")
            logging.info(f"ターゲットディレクトリ: {target_dir}")
            logging.info(f"sys.frozen: {getattr(sys, 'frozen', False)}")
            
            # ファイル更新前の状態確認
            self._log_directory_structure(source_dir, "更新前ソース構造")
            self._log_directory_structure(target_dir, "更新前ターゲット構造")
            
            # ユーザーデータファイル（保護対象）のパターン
            protected_patterns = [
                'item_manage.xlsm',  # ユーザーの商品管理ファイル
                '*_user_*',          # ユーザー作成ファイル
                '*.backup',          # バックアップファイル  
                'user_settings.json', # ユーザー設定
                'config.ini',        # 設定ファイル
            ]
            
            # ユーザーデータのバックアップを作成
            try:
                backup_created = self._create_user_data_backup(target_dir)
                if backup_created:
                    self.status.emit("ユーザーデータのバックアップを作成しました")
                    logging.info("ユーザーデータバックアップ作成完了")
            except Exception as e:
                logging.error(f"ユーザーデータバックアップ作成エラー: {e}")
                # バックアップ失敗は続行可能
        
            # 展開されたファイルを探す
            file_count = 0
            for root, dirs, files in os.walk(source_dir):
                if self._cancelled:
                    logging.info("ファイル更新中にキャンセルされました")
                    return
                    
                rel_path = os.path.relpath(root, source_dir)
                target_root = os.path.join(target_dir, rel_path) if rel_path != '.' else target_dir
                
                # ディレクトリマッピングをログ出力
                if rel_path != '.':
                    logging.info(f"ディレクトリマッピング: {rel_path} -> {target_root}")
                
                # ディレクトリを作成
                try:
                    if not os.path.exists(target_root) and rel_path != '.':
                        os.makedirs(target_root, exist_ok=True)
                        logging.info(f"ディレクトリ作成: {target_root}")
                except Exception as e:
                    logging.error(f"ディレクトリ作成エラー {target_root}: {e}")
                    raise
                
                for file in files:
                    if self._cancelled:
                        logging.info("ファイルコピー中にキャンセルされました")
                        return
                        
                    try:
                        source_file = os.path.join(root, file)
                        target_file = os.path.join(target_root, file)
                        file_count += 1
                        
                        logging.info(f"🔍 処理中のファイル[{file_count}]: {file}")
                        logging.info(f"   📁 rel_path: '{rel_path}'")
                        logging.info(f"   📂 source_file: {source_file}")
                        logging.info(f"   📂 target_file: {target_file}")
                        logging.info(f"   🖥️ sys.frozen: {getattr(sys, 'frozen', False)}")
                        
                        # ユーザーデータファイルの保護チェック
                        if self._is_user_data_file(file, rel_path, protected_patterns):
                            # 既存のユーザーデータファイルがある場合は保護
                            if os.path.exists(target_file):
                                self.status.emit(f"ユーザーデータを保護: {file}")
                                logging.info(f"ユーザーデータファイルを保護: {target_file}")
                                continue  # このファイルはスキップ
                        
                        # PyInstallerでビルドされたexeファイルの更新
                        if getattr(sys, 'frozen', False):
                            # 商品登録入力ツール.exe の場合（どの階層にあっても）
                            if file == '商品登録入力ツール.exe':
                                # 実行中のexeファイルは.newとして保存
                                original_target = target_file
                                target_file = os.path.join(target_dir, '商品登録入力ツール.exe.new')
                                updated_exe = True
                                self.status.emit(f"実行ファイルを更新中: {file}")
                                logging.info(f"🎯 EXEファイル検出: {file}")
                                logging.info(f"   📂 rel_path: '{rel_path}'")
                                logging.info(f"   📂 元のターゲット: {original_target}")
                                logging.info(f"   🆕 新しいターゲット: {target_file}")
                                logging.info(f"⚡ 重要: .newファイルを作成します: {target_file}")
                                logging.info(f"🔥🔥🔥 ZIP階層構造でEXEファイル発見！ 🔥🔥🔥")
                            else:
                                logging.info(f"📄 通常ファイル: {file}")
                        else:
                            # 開発環境の場合、実行中のスクリプトと同じ場合
                            if os.path.abspath(target_file) == current_exe:
                                target_file = target_file + '.new'
                                updated_exe = True
                                logging.info(f"開発環境ファイル更新: {file}")
                            else:
                                logging.info(f"📄 開発環境通常ファイル: {file}")
                        
                        # ファイルをコピー
                        retry_count = 0
                        max_retries = 3
                        
                        while retry_count < max_retries:
                            try:
                                # ファイルサイズ確認
                                source_size = os.path.getsize(source_file)
                                if source_size == 0:
                                    logging.warning(f"ソースファイルのサイズが0: {source_file}")
                                
                                # 大きなファイルの場合は進捗表示
                                if source_size > 1024 * 1024:  # 1MB以上
                                    self.status.emit(f"大きなファイルをコピー中: {file} ({source_size/1024/1024:.1f}MB)")
                                    logging.info(f"大きなファイルのコピー開始: {file} ({source_size} bytes)")
                                
                                # キャンセルチェック
                                if self._cancelled:
                                    logging.info(f"ファイルコピー前にキャンセル: {file}")
                                    return
                                
                                # リトライ表示
                                if retry_count > 0:
                                    self.status.emit(f"ファイルコピー再試行中 ({retry_count+1}/{max_retries}): {file}")
                                    logging.info(f"ファイルコピー再試行 {retry_count+1}/{max_retries}: {file}")
                                
                                # 🚨 ファイルコピー前の最終確認
                                logging.info(f"🚨 ファイルコピー実行前:")
                                logging.info(f"   📂 source_file: {source_file}")
                                logging.info(f"   📂 target_file: {target_file}")
                                logging.info(f"   📏 source_size: {source_size:,} bytes")
                                logging.info(f"   ✅ source exists: {os.path.exists(source_file)}")
                                logging.info(f"   📁 target dir exists: {os.path.exists(os.path.dirname(target_file))}")
                                logging.info(f"   🖥️ is EXE file: {file == '商品登録入力ツール.exe'}")
                                logging.info(f"   🎯 is .new target: {target_file.endswith('.new')}")
                                
                                # ファイルコピー実行（チャンク方式で安全にコピー）
                                if source_size > 10 * 1024 * 1024:  # 10MB以上の大きなファイル
                                    logging.info(f"🔄 大きなファイルのチャンクコピー開始")
                                    self._copy_large_file(source_file, target_file, source_size)
                                else:
                                    logging.info(f"🔄 通常のファイルコピー開始")
                                    shutil.copy2(source_file, target_file)
                                
                                # 🚨 ファイルコピー後の確認
                                logging.info(f"🚨 ファイルコピー実行後:")
                                logging.info(f"   ✅ target exists: {os.path.exists(target_file)}")
                                if os.path.exists(target_file):
                                    actual_size = os.path.getsize(target_file)
                                    logging.info(f"   📏 actual size: {actual_size:,} bytes")
                                    logging.info(f"   ✅ size match: {actual_size == source_size}")
                                    
                                    # .newファイルの場合は特別な確認
                                    if target_file.endswith('.new'):
                                        logging.info(f"🔥🔥🔥 .NEWファイル作成成功: {target_file} 🔥🔥🔥")
                                        logging.info(f"🔥 .NEWファイルサイズ: {actual_size:,} bytes")
                                else:
                                    logging.error(f"❌❌❌ ファイルコピー後にファイルが存在しません: {target_file}")
                                
                                # コピー成功した場合はループを抜ける
                                break
                                
                            except (PermissionError, OSError) as copy_error:
                                retry_count += 1
                                if retry_count < max_retries:
                                    # リトライ前に少し待機
                                    import time
                                    wait_time = retry_count * 2  # 2秒、4秒と増加
                                    logging.warning(f"ファイルコピーエラー（{retry_count}/{max_retries}）: {copy_error}")
                                    logging.info(f"{wait_time}秒待機してリトライします...")
                                    time.sleep(wait_time)
                                else:
                                    # 最大リトライ回数に達した場合
                                    raise copy_error
                            except Exception as copy_error:
                                # その他のエラーは即座に失敗
                                raise copy_error
                        
                        # キャンセルチェック（コピー後）
                        if self._cancelled:
                            logging.info(f"ファイルコピー後にキャンセル: {file}")
                            # 中途半端なファイルを削除
                            if os.path.exists(target_file):
                                try:
                                    os.remove(target_file)
                                    logging.info(f"中途半端なファイルを削除: {target_file}")
                                except Exception as e:
                                    logging.warning(f"中途半端ファイル削除エラー: {e}")
                            return
                        
                        # コピー後のサイズ確認
                        if os.path.exists(target_file):
                            target_size = os.path.getsize(target_file)
                            if source_size != target_size:
                                logging.error(f"ファイルサイズ不一致: {file} - 期待値:{source_size} 実際:{target_size}")
                                raise Exception(f"ファイルサイズ不一致: {source_size} != {target_size}")
                                
                            logging.info(f"ファイルコピー完了: {file} ({source_size} bytes)")
                            logging.info(f"  コピー元: {source_file}")
                            logging.info(f"  コピー先: {target_file}")
                        else:
                            raise Exception(f"コピー後にファイルが存在しません: {target_file}")
                    
                    except Exception as file_error:
                        logging.error(f"ファイル処理エラー {file}: {file_error}")
                        # 個別ファイルエラーは続行しない（重要なファイルの可能性があるため）
                        raise file_error
        
            logging.info(f"ファイル更新完了: {file_count}個のファイルを処理")
            
            # 🚨🚨🚨 最終結果検証 - 絶対確認モード 🚨🚨🚨
            logging.info("=" * 80)
            logging.info("🚨🚨🚨 最終結果検証 - 絶対確認モード 🚨🚨🚨")
            logging.info("=" * 80)
            
            # .newファイルの徹底確認
            new_exe_path = os.path.join(target_dir, '商品登録入力ツール.exe.new')
            logging.info(f"🔍 期待される.newファイルパス: {new_exe_path}")
            
            if os.path.exists(new_exe_path):
                new_size = os.path.getsize(new_exe_path)
                logging.info(f"🔥🔥🔥 .NEWファイル確認成功! 🔥🔥🔥")
                logging.info(f"✅ パス: {new_exe_path}")
                logging.info(f"✅ サイズ: {new_size:,} bytes")
                
                # 元のexeファイルと比較
                original_exe = os.path.join(target_dir, '商品登録入力ツール.exe')
                if os.path.exists(original_exe):
                    original_size = os.path.getsize(original_exe)
                    logging.info(f"📊 サイズ比較: 旧={original_size:,} bytes, 新={new_size:,} bytes")
                    size_diff = new_size - original_size
                    logging.info(f"📊 サイズ差: {size_diff:,} bytes")
                else:
                    logging.warning(f"⚠️ 元のEXEファイルが見つかりません: {original_exe}")
                    
            else:
                logging.error(f"❌❌❌ .NEWファイルが見つかりません! ❌❌❌")
                logging.error(f"❌ 期待パス: {new_exe_path}")
                
                # 徹底的なディレクトリ確認
                logging.info("🔍 ターゲットディレクトリの徹底確認:")
                try:
                    all_files = os.listdir(target_dir)
                    logging.info(f"📁 ディレクトリ内アイテム数: {len(all_files)}")
                    
                    # .newファイル検索
                    new_files = [f for f in all_files if f.endswith('.new')]
                    logging.info(f"🔍 .newファイル数: {len(new_files)}")
                    for new_file in new_files:
                        full_path = os.path.join(target_dir, new_file)
                        size = os.path.getsize(full_path)
                        logging.info(f"  🔥 発見: {new_file} ({size:,} bytes)")
                    
                    # EXEファイル検索
                    exe_files = [f for f in all_files if f.endswith('.exe')]
                    logging.info(f"🔍 EXEファイル数: {len(exe_files)}")
                    for exe_file in exe_files:
                        full_path = os.path.join(target_dir, exe_file)
                        size = os.path.getsize(full_path)
                        logging.info(f"  ⚡ EXE: {exe_file} ({size:,} bytes)")
                    
                    # 全ファイル表示（最大20個）
                    logging.info("📋 全ファイル/フォルダ一覧（最大20個）:")
                    for i, item in enumerate(all_files[:20]):
                        item_path = os.path.join(target_dir, item)
                        if os.path.isfile(item_path):
                            size = os.path.getsize(item_path)
                            logging.info(f"  {i+1:2d}. 📄 {item} ({size:,} bytes)")
                        else:
                            logging.info(f"  {i+1:2d}. 📁 {item}/")
                    
                    if len(all_files) > 20:
                        logging.info(f"  ... 他 {len(all_files)-20} 個のアイテム")
                        
                except Exception as e:
                    logging.error(f"❌ ディレクトリ確認エラー: {e}")
                
                # 絶対パス確認
                abs_target_dir = os.path.abspath(target_dir)
                abs_new_path = os.path.abspath(new_exe_path)
                logging.info(f"🔍 絶対パス確認:")
                logging.info(f"  📁 ターゲットdir: {abs_target_dir}")
                logging.info(f"  📄 .newファイル: {abs_new_path}")
                logging.info(f"  ✅ ターゲットdir存在: {os.path.exists(abs_target_dir)}")
                logging.info(f"  ✅ .newファイル存在: {os.path.exists(abs_new_path)}")
            
            logging.info("=" * 80)
            
            # C#フォルダの確認
            cs_dir = os.path.join(target_dir, 'C#')
            if os.path.exists(cs_dir):
                cs_files = os.listdir(cs_dir)
                logging.info(f"✅ C#フォルダ確認: {len(cs_files)}個のファイル/フォルダ")
                
                # 重要ファイルの確認
                important_cs_files = ['csharp.exe', 'ec_csv_tool']
                for important_file in important_cs_files:
                    path = os.path.join(cs_dir, important_file)
                    exists = os.path.exists(path)
                    logging.info(f"  {'✅' if exists else '❌'} {important_file}")
            else:
                logging.error(f"❌ C#フォルダが見つかりません: {cs_dir}")
            
            # 更新後のディレクトリ構造をログ
            self._log_directory_structure(target_dir, "更新後ターゲット構造")
            
            if not updated_exe and getattr(sys, 'frozen', False):
                # exeファイルが見つからなかった場合の警告
                logging.error("❌ 重大エラー: 更新パッケージ内に実行ファイルが見つかりませんでした")
            else:
                logging.info(f"✅ 実行ファイル更新フラグ: {updated_exe}")
                
        except Exception as e:
            logging.error(f"ファイル更新エラー: {e}")
            raise
    
    def _is_user_data_file(self, filename: str, rel_path: str, protected_patterns: list) -> bool:
        """ファイルがユーザーデータかどうかを判定"""
        try:
            import fnmatch
            
            # ファイル名パターンマッチング
            for pattern in protected_patterns:
                if fnmatch.fnmatch(filename.lower(), pattern.lower()):
                    logging.debug(f"ユーザーデータファイル検出（パターン）: {filename}")
                    return True
            
            # 特定のディレクトリ内のファイル（C#ツール内など）
            if 'C#' in rel_path and filename.endswith('.xlsm'):
                logging.debug(f"ユーザーデータファイル検出（C#ディレクトリ）: {filename}")
                return True
                
            # ファイルサイズ・更新日時による判定（テンプレートより大きい場合はユーザーデータの可能性）
            # item_manage.xlsmがitem_template.xlsmより大きい場合など
            
            return False
            
        except Exception as e:
            logging.error(f"ユーザーデータファイル判定エラー {filename}: {e}")
            # エラーの場合は保護対象として扱う（安全側に倒す）
            return True
    
    def _create_user_data_backup(self, target_dir: str) -> bool:
        """重要なユーザーデータのバックアップを作成"""
        try:
            import datetime
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_dir = os.path.join(target_dir, f"backup_before_update_{timestamp}")
            
            # バックアップ対象ファイル
            important_files = [
                'item_manage.xlsm',
                'config.ini', 
                'user_settings.json'
            ]
            
            backup_created = False
            for filename in important_files:
                if self._cancelled:
                    logging.info("バックアップ作成中にキャンセルされました")
                    return backup_created
                    
                try:
                    source_file = os.path.join(target_dir, filename)
                    if os.path.exists(source_file):
                        if not os.path.exists(backup_dir):
                            os.makedirs(backup_dir, exist_ok=True)
                            logging.info(f"バックアップディレクトリ作成: {backup_dir}")
                        
                        backup_file = os.path.join(backup_dir, filename)
                        shutil.copy2(source_file, backup_file)
                        backup_created = True
                        logging.info(f"バックアップ作成: {source_file} -> {backup_file}")
                except Exception as e:
                    logging.error(f"個別ファイルバックアップエラー {filename}: {e}")
                    # 個別ファイルのエラーは続行
            
            if backup_created:
                logging.info(f"ユーザーデータバックアップ完了: {backup_dir}")
            else:
                logging.info("バックアップ対象ファイルが見つかりませんでした")
                
            return backup_created
            
        except Exception as e:
            logging.error(f"バックアップ作成エラー: {e}")
            return False
    
    def _find_actual_source_directory(self, extract_dir: str) -> str:
        """展開されたZIPファイル内から実際の更新ファイルがあるディレクトリを特定"""
        try:
            logging.info(f"=== ソースディレクトリ検索開始 ===")
            logging.info(f"展開ディレクトリ: {extract_dir}")
            
            # 展開されたディレクトリの構造を詳細に表示
            self._log_directory_structure(extract_dir, "展開構造")
            
            # 重要なファイルのリスト（優先順位付き）
            critical_files = ['商品登録入力ツール.exe', 'product_app.py']
            important_files = ['constants.py', 'version.json']
            directory_markers = ['C#']  # 重要ディレクトリ
            
            # === 詳細な全ファイル探索 ===
            logging.info("=== 全ファイル詳細探索 ===")
            all_files_found = []
            exe_files_found = []
            for root, dirs, files in os.walk(extract_dir):
                rel_path = os.path.relpath(root, extract_dir)
                logging.info(f"📁 ディレクトリ: {rel_path}")
                for file in files:
                    all_files_found.append(f"{rel_path}/{file}")
                    if file.endswith('.exe'):
                        exe_files_found.append(f"{rel_path}/{file}")
                        logging.info(f"  🔥 EXEファイル発見: {file}")
                    logging.info(f"  📄 {file}")
            
            logging.info(f"=== 検索結果サマリー ===")
            logging.info(f"総ファイル数: {len(all_files_found)}")
            logging.info(f"EXEファイル数: {len(exe_files_found)}")
            logging.info(f"EXEファイルリスト: {exe_files_found}")
            
            # 商品登録入力ツール.exe を特別検索
            target_exe_paths = []
            for root, dirs, files in os.walk(extract_dir):
                for file in files:
                    if file == '商品登録入力ツール.exe':
                        full_path = os.path.join(root, file)
                        target_exe_paths.append(full_path)
                        logging.info(f"🎯 ターゲットEXE発見: {full_path}")
            
            if target_exe_paths:
                logging.info(f"✅ 商品登録入力ツール.exe が {len(target_exe_paths)} 箇所で発見されました")
                # 最初に見つかったものを使用
                target_dir = os.path.dirname(target_exe_paths[0])
                logging.info(f"🚀 強制的にターゲットディレクトリを設定: {target_dir}")
                return target_dir
            else:
                logging.error(f"❌ 商品登録入力ツール.exe が見つかりません！")
            
            best_candidate = None
            best_score = 0
            
            # 全ディレクトリをチェック
            for root, dirs, files in os.walk(extract_dir):
                rel_path = os.path.relpath(root, extract_dir)
                score = 0
                
                # クリティカルファイルのチェック（高得点）
                for critical_file in critical_files:
                    if critical_file in files:
                        score += 10
                        logging.info(f"  ✓ クリティカルファイル発見: {critical_file} in {rel_path}")
                
                # 重要ファイルのチェック（中得点）
                for important_file in important_files:
                    if important_file in files:
                        score += 5
                        logging.info(f"  ✓ 重要ファイル発見: {important_file} in {rel_path}")
                
                # 重要ディレクトリのチェック（中得点）
                for directory_marker in directory_markers:
                    if directory_marker in dirs:
                        score += 5
                        logging.info(f"  ✓ 重要ディレクトリ発見: {directory_marker} in {rel_path}")
                
                logging.info(f"ディレクトリスコア: {rel_path} -> {score}点")
                
                if score > best_score:
                    best_score = score
                    best_candidate = root
                    logging.info(f"  → 新しいベスト候補: {root} ({score}点)")
            
            # 結果の決定
            if best_candidate and best_score >= 10:  # クリティカルファイルが最低1つ必要
                logging.info(f"✅ 最適ソースディレクトリ確定: {best_candidate} ({best_score}点)")
                return best_candidate
            
            # フォールバック: ProductRegisterTool で始まるディレクトリ
            subdirs = [d for d in os.listdir(extract_dir) if os.path.isdir(os.path.join(extract_dir, d))]
            for subdir in subdirs:
                if subdir.startswith('ProductRegisterTool'):
                    subdir_path = os.path.join(extract_dir, subdir)
                    logging.info(f"⚠️ フォールバック1: ProductRegisterToolディレクトリ使用: {subdir_path}")
                    return subdir_path
            
            # 最終フォールバック
            if subdirs:
                first_subdir = os.path.join(extract_dir, subdirs[0])
                logging.info(f"⚠️ フォールバック2: 最初のサブディレクトリ使用: {first_subdir}")
                return first_subdir
            
            # 最後の手段
            logging.warning(f"⚠️ フォールバック3: extract_dir自体を使用: {extract_dir}")
            return extract_dir
            
        except Exception as e:
            logging.error(f"ソースディレクトリ検索エラー: {e}")
            return extract_dir
    
    def _log_directory_structure(self, directory: str, title: str, max_depth: int = 3):
        """ディレクトリ構造を詳細にログ出力"""
        try:
            logging.info(f"=== {title} ===")
            for root, dirs, files in os.walk(directory):
                level = root.replace(directory, '').count(os.sep)
                if level >= max_depth:
                    continue
                    
                indent = '  ' * level
                rel_path = os.path.relpath(root, directory)
                if rel_path == '.':
                    rel_path = '(ルート)'
                    
                logging.info(f"{indent}{rel_path}/ ({len(files)}ファイル)")
                
                # 重要なファイルをハイライト
                important_extensions = ['.exe', '.py', '.json']
                for file in files[:10]:  # 最大10ファイルまで表示
                    file_ext = os.path.splitext(file)[1].lower()
                    marker = "⭐" if file_ext in important_extensions else "  "
                    logging.info(f"{indent}  {marker} {file}")
                    
                if len(files) > 10:
                    logging.info(f"{indent}  ... 他{len(files)-10}ファイル")
                    
        except Exception as e:
            logging.error(f"ディレクトリ構造ログエラー: {e}")


class VersionChecker:
    """バージョンチェックと更新管理を行うクラス"""
    
    def __init__(self, parent=None):
        self.parent = parent
        self.logger = logging.getLogger(__name__)
        
    def check_for_updates(self, silent: bool = False) -> Optional[VersionInfo]:
        """
        GitHub上の最新バージョンをチェック
        
        Args:
            silent: Trueの場合、最新版の場合にメッセージを表示しない
            
        Returns:
            新しいバージョンがある場合はVersionInfo、それ以外はNone
        """
        try:
            self.logger.info(f"バージョンチェック開始: 現在={CURRENT_VERSION}, URL={VERSION_CHECK_URL}")
            
            # GitHub APIからversion.jsonを取得
            req = Request(VERSION_CHECK_URL, headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                'Cache-Control': 'no-cache'
            })
            
            with urlopen(req, timeout=15) as response:
                if response.getcode() != 200:
                    raise Exception(f"HTTP {response.getcode()}: バージョン情報の取得に失敗")
                    
                raw_data = response.read()
                self.logger.info(f"バージョンデータ取得成功: {len(raw_data)} bytes")
                
                # エンコーディングエラーを回避するため、複数のエンコーディングを試す
                encodings = ['utf-8', 'utf-8-sig', 'shift_jis', 'cp932']
                decoded_data = None
                
                for encoding in encodings:
                    try:
                        decoded_data = raw_data.decode(encoding)
                        self.logger.info(f"エンコーディング成功: {encoding}")
                        break
                    except UnicodeDecodeError:
                        self.logger.warning(f"エンコーディング失敗: {encoding}")
                        continue
                
                if decoded_data is None:
                    # どのエンコーディングでも失敗した場合、エラーを無視してデコード
                    decoded_data = raw_data.decode('utf-8', errors='replace')
                    self.logger.warning("エンコーディングエラーを無視してデコードしました")
                
                # JSONパース時にエラーが発生した場合の対策
                try:
                    version_data = json.loads(decoded_data)
                except json.JSONDecodeError as e:
                    # 文字化けを含む可能性があるため、正規表現で修正を試みる
                    import re
                    # 文字化けしたテキストを除去または置換
                    cleaned_data = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', decoded_data)
                    # 不正なUnicode文字を置換
                    cleaned_data = re.sub(r'[^\x00-\x7F\u3000-\u303F\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF\uFF00-\uFFEF]+', '?', cleaned_data)
                    self.logger.warning(f"JSONパースエラー。クリーニング後に再試行: {e}")
                    version_data = json.loads(cleaned_data)
            
            version_info = VersionInfo(version_data)
            remote_version = version_info.version
            
            self.logger.info(f"バージョン比較: 現在={CURRENT_VERSION}, リモート={remote_version}")
            
            # バージョン比較
            if self._is_newer_version(remote_version, CURRENT_VERSION):
                self.logger.info(f"新しいバージョンを検出: {remote_version}")
                return version_info
            else:
                self.logger.info(f"最新バージョンを使用中: {CURRENT_VERSION}")
                if not silent:
                    QMessageBox.information(
                        self.parent,
                        "更新確認",
                        f"お使いのバージョン ({CURRENT_VERSION}) は最新です。\n\n"
                        f"リモートバージョン: {remote_version}\n"
                        f"チェック日時: {version_info.release_date}"
                    )
                return None
            
        except (URLError, HTTPError) as e:
            error_msg = f"ネットワークエラー: {e}"
            self.logger.error(f"バージョンチェック中の{error_msg}")
            if not silent:
                QMessageBox.warning(
                    self.parent,
                    "更新確認エラー",
                    f"更新の確認中にエラーが発生しました。\n\n"
                    f"エラー詳細: {error_msg}\n\n"
                    f"現在のバージョン: {CURRENT_VERSION}\n"
                    f"チェックURL: {VERSION_CHECK_URL}\n\n"
                    f"インターネット接続を確認してください。"
                )
            return None
            
        except json.JSONDecodeError as e:
            error_msg = f"JSON解析エラー: {e}"
            self.logger.error(f"バージョンチェック中の{error_msg}")
            if not silent:
                QMessageBox.warning(
                    self.parent,
                    "更新確認エラー",
                    f"バージョン情報の解析中にエラーが発生しました。\n\n"
                    f"エラー詳細: {error_msg}\n\n"
                    f"GitHub上のversion.jsonファイルを確認してください。"
                )
            return None
            
        except UnicodeDecodeError as e:
            error_msg = f"エンコーディングエラー: {e}"
            self.logger.error(f"バージョンチェック中の{error_msg}")
            if not silent:
                QMessageBox.warning(
                    self.parent,
                    "更新確認エラー",
                    f"文字エンコーディングエラーが発生しました。\n\n"
                    f"エラー詳細: {error_msg}"
                )
            return None
            
        except Exception as e:
            error_msg = f"予期しないエラー: {e}"
            self.logger.error(f"バージョンチェック中の{error_msg}", exc_info=True)
            if not silent:
                QMessageBox.warning(
                    self.parent,
                    "更新確認エラー",
                    f"更新の確認中に予期しないエラーが発生しました。\n\n"
                    f"エラー詳細: {error_msg}\n\n"
                    f"現在のバージョン: {CURRENT_VERSION}\n"
                    f"ネットワーク接続と設定を確認してください。"
                )
            return None
    
    def prompt_for_update(self, version_info: VersionInfo) -> bool:
        """
        更新するかユーザーに確認
        
        Returns:
            更新する場合True
        """
        message = f"""新しいバージョン {version_info.version} が利用可能です。
（現在のバージョン: {CURRENT_VERSION}）

リリース日: {version_info.release_date}

{version_info.get_latest_changes()}

📋 データ保護機能:
• ユーザーの商品データ (item_manage.xlsm) は自動保護
• 設定ファイルとバックアップは自動作成
• 更新前にバックアップフォルダを生成

今すぐ更新しますか？"""
        
        reply = QMessageBox.question(
            self.parent,
            "更新の確認",
            message,
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        
        return reply == QMessageBox.Yes
    
    def download_and_install_update(self, version_info: VersionInfo):
        """更新をダウンロードしてインストール"""
        try:
            # プログレスダイアログを作成
            progress = QProgressDialog("更新ファイルをダウンロード中...", "キャンセル", 0, 100, self.parent)
            progress.setWindowTitle(f"商品登録入力ツール v{version_info.version} へのアップデート")
            progress.setModal(True)
            progress.setAutoClose(False)
            progress.setMinimumDuration(0)  # すぐに表示
            progress.setMinimumWidth(400)  # 幅を広げる
            progress.show()
            
            # アプリケーションディレクトリを自動検出
            app_dir = self._detect_application_directory()
            logging.info(f"自動検出されたアプリケーションディレクトリ: {app_dir}")
            
            # 自動ダウンロード機能：ユーザーの選択によって自動または手動
            msg_box = QMessageBox(self.parent)
            msg_box.setIcon(QMessageBox.Question)
            msg_box.setWindowTitle("更新方法の選択")
            msg_box.setText(f"新しいバージョン {version_info.version} をダウンロードします。")
            msg_box.setInformativeText(
                "どちらの方法で更新しますか？\n\n"
                "🔄 自動ダウンロード: アプリが自動でダウンロード・インストール\n"
                "🌐 手動ダウンロード: ブラウザでダウンロードページを開く"
            )
            
            auto_btn = msg_box.addButton("自動ダウンロード（推奨）", QMessageBox.AcceptRole)
            manual_btn = msg_box.addButton("手動ダウンロード", QMessageBox.ActionRole) 
            cancel_btn = msg_box.addButton("キャンセル", QMessageBox.RejectRole)
            msg_box.setDefaultButton(auto_btn)
            
            msg_box.exec_()
            clicked_button = msg_box.clickedButton()
            
            if clicked_button == cancel_btn:
                progress.close()
                return
            elif clicked_button == manual_btn:
                # 手動ダウンロード（直接ダウンロードURL）
                progress.close()
                import webbrowser
                
                # 直接ダウンロードURLを開く
                if version_info.download_url and version_info.download_url.startswith('https://'):
                    # 直接ダウンロードURLをブラウザで開く
                    webbrowser.open(version_info.download_url)
                    
                    QMessageBox.information(
                        self.parent,
                        "ダウンロード開始",
                        f"ブラウザで直接ダウンロードを開始しました。\n\n"
                        f"ダウンロードファイル: ProductRegisterTool-v{version_info.version}.zip\n\n"
                        f"ダウンロード完了後:\n"
                        f"1. このアプリを終了\n"
                        f"2. ZIPファイルを適当なフォルダに解凍\n"
                        f"3. 新しいバージョンを起動"
                    )
                else:
                    # フォールバック: リリースページを開く
                    download_url_parts = version_info.download_url.split('/')
                    if len(download_url_parts) >= 8 and download_url_parts[5] == 'releases':
                        tag_name = download_url_parts[7]  # v2.2.6
                        repo_path = '/'.join(download_url_parts[:5])  # https://github.com/SEI1026/Product_app
                        release_url = f"{repo_path}/releases/tag/{tag_name}"
                    else:
                        # 最終フォールバック: リリース一覧ページ
                        release_url = "https://github.com/SEI1026/Product_app/releases"
                    
                    webbrowser.open(release_url)
                    
                    QMessageBox.information(
                        self.parent,
                        "ダウンロードページ",
                        "ブラウザでリリースページを開きました。\n"
                        "手動でZIPファイルをダウンロードしてください。"
                    )
                return
            else:
                # 自動ダウンロード（新機能）
                # ダウンロードURL検証
                if not version_info.download_url or not version_info.download_url.startswith('https://'):
                    progress.close()
                    QMessageBox.critical(
                        self.parent,
                        "更新エラー",
                        "無効なダウンロードURLです。手動ダウンロードをお試しください。"
                    )
                    return
                
                # ダウンロード用スレッドを作成
                downloader = UpdateDownloader(version_info.download_url, app_dir)
                
                # プログレス状態を保存する変数
                self._update_result = None
                self._update_message = None
                
                # 完了時の処理（UIスレッドで実行するためにタイマーを使用）
                def on_finished(success: bool, message: str):
                    try:
                        logging.info(f"更新完了コールバック: success={success}")
                        
                        # 結果を保存
                        self._update_result = success
                        self._update_message = message
                        
                        # プログレスダイアログを閉じる
                        if progress and not progress.wasCanceled():
                            progress.close()
                            
                        # UIスレッドで結果処理を実行（100ms後）
                        QTimer.singleShot(100, self._handle_update_result)
                        
                    except Exception as e:
                        logging.error(f"完了コールバック処理エラー: {e}", exc_info=True)
                
                def on_cancel():
                    """キャンセル処理"""
                    try:
                        downloader.terminate()
                        if progress:
                            progress.close()
                        logging.info("ユーザーがアップデートをキャンセルしました")
                    except Exception as e:
                        logging.error(f"キャンセル処理エラー: {e}")
                
                # シグナル接続（エラーハンドリング付き）
                try:
                    downloader.progress.connect(progress.setValue)
                    downloader.status.connect(progress.setLabelText)
                    downloader.finished.connect(on_finished)
                    progress.canceled.connect(on_cancel)
                except Exception as e:
                    logging.error(f"シグナル接続エラー: {e}")
                    progress.close()
                    QMessageBox.critical(self.parent, "更新エラー", "更新処理の初期化に失敗しました")
                    return
                
                # ダウンロード開始前の事前チェック
                try:
                    # ディスク容量チェック（簡易）
                    import shutil
                    total, used, free = shutil.disk_usage(app_dir)
                    free_mb = free / (1024 * 1024)
                    if free_mb < 100:  # 100MB未満の場合は警告
                        logging.warning(f"ディスク容量不足の可能性: {free_mb:.1f}MB")
                        reply = QMessageBox.question(
                            self.parent,
                            "容量警告",
                            f"ディスクの空き容量が少ないです（{free_mb:.1f}MB）。\n更新を続行しますか？",
                            QMessageBox.Yes | QMessageBox.No,
                            QMessageBox.No
                        )
                        if reply == QMessageBox.No:
                            progress.close()
                            return
                            
                    # ターゲットディレクトリの書き込み権限チェック
                    if not os.access(app_dir, os.W_OK):
                        progress.close()
                        QMessageBox.critical(
                            self.parent,
                            "権限エラー", 
                            f"アプリケーションディレクトリに書き込み権限がありません：\n{app_dir}"
                        )
                        return
                        
                except Exception as e:
                    logging.warning(f"事前チェックエラー: {e}")
                    # 事前チェックエラーは続行
                
                # ダウンロード開始
                logging.info("アップデートダウンロード開始")
                downloader.start()
                return
        except Exception as e:
            logging.error(f"更新ダイアログ作成中にエラー: {e}")
            QMessageBox.critical(
                self.parent,
                "更新エラー",
                f"更新の準備中にエラーが発生しました:\n{str(e)}"
            )
            return
    
    def _is_newer_version(self, version1: str, version2: str) -> bool:
        """
        バージョン比較（version1 > version2 の場合True）
        """
        try:
            v1_parts = [int(x) for x in version1.split('.')]
            v2_parts = [int(x) for x in version2.split('.')]
            
            # バージョン番号の長さを揃える
            max_len = max(len(v1_parts), len(v2_parts))
            v1_parts.extend([0] * (max_len - len(v1_parts)))
            v2_parts.extend([0] * (max_len - len(v2_parts)))
            
            return v1_parts > v2_parts
            
        except ValueError:
            # バージョン番号のパースに失敗した場合は文字列比較
            return version1 > version2
    
    def _get_log_file_info(self):
        """ログファイルの情報を取得"""
        try:
            # 一般的なログファイルのパスを確認
            possible_log_paths = [
                "application.log",
                "app.log", 
                "product_app.log",
                os.path.join(os.path.expanduser("~"), "AppData", "Local", "ProductApp", "logs", "app.log"),
                os.path.join(os.path.dirname(sys.executable), "logs", "app.log"),
                os.path.join(tempfile.gettempdir(), "product_app.log")
            ]
            
            log_info = []
            
            # 現在のロガーの設定を確認
            current_logger = logging.getLogger()
            if current_logger.handlers:
                for handler in current_logger.handlers:
                    if hasattr(handler, 'baseFilename'):
                        log_path = handler.baseFilename
                        if os.path.exists(log_path):
                            log_size = os.path.getsize(log_path)
                            log_info.append(f"ログファイル: {log_path}")
                            log_info.append(f"ファイルサイズ: {log_size/1024:.1f}KB")
                            return "\n".join(log_info)
            
            # 既知のパスから検索
            for log_path in possible_log_paths:
                if os.path.exists(log_path):
                    log_size = os.path.getsize(log_path)
                    log_info.append(f"ログファイル: {log_path}")
                    log_info.append(f"ファイルサイズ: {log_size/1024:.1f}KB")
                    return "\n".join(log_info)
            
            return "ログファイルが見つかりませんでした"
            
        except Exception as e:
            return f"ログ情報取得エラー: {e}"
    
    def _create_restart_script(self):
        """再起動用のスクリプトを作成（安全なプロセス終了）"""
        # Windows以外でもテスト用にバッチファイルを生成
        if sys.platform == 'win32' or getattr(sys, 'frozen', False):
            # 実際のアプリケーションパスを取得
            app_dir = self._detect_application_directory()
            
            # 実行可能ファイルのパスを正しく取得
            if getattr(sys, 'frozen', False):
                # PyInstallerでビルドされた場合
                exe_path = os.path.join(app_dir, '商品登録入力ツール.exe')
            else:
                # 開発環境の場合
                exe_path = sys.executable
                
            exe_dir = os.path.dirname(exe_path)
            exe_name = os.path.basename(exe_path)
            script_path = os.path.join(app_dir, 'update_restart.ps1')
            
            # パスの存在確認 - 重要: _update_filesメソッドと同じロジックで.newファイルパスを構築
            if getattr(sys, 'frozen', False):
                # PyInstallerでビルドされた場合は、target_dirベースで.newファイルパスを構築
                # _update_filesメソッドの line 614 と同じロジック
                new_file_path = os.path.join(app_dir, '商品登録入力ツール.exe.new')
            else:
                # 開発環境の場合
                new_file_path = exe_path + '.new'
            
            logging.info(f"=== 再起動スクリプト情報 ===")
            logging.info(f"アプリディレクトリ: {app_dir}")
            logging.info(f"実行ファイルパス: {exe_path}")
            logging.info(f".newファイルパス: {new_file_path}")
            logging.info(f"スクリプトパス: {script_path}")
            logging.info(f".newファイル存在: {os.path.exists(new_file_path)}")
            
            # 追加のデバッグ情報
            if not os.path.exists(new_file_path):
                logging.warning(f"⚠️ .newファイルが見つかりません！")
                # app_dir内の.newファイルを検索
                try:
                    import glob
                    new_files = glob.glob(os.path.join(app_dir, "*.new"))
                    if new_files:
                        logging.info(f"発見された.newファイル: {new_files}")
                    else:
                        logging.error(f"❌ {app_dir}内に.newファイルが存在しません")
                except Exception as e:
                    logging.error(f"glob検索エラー: {e}")
            
            # 現在のプロセスIDを取得
            current_pid = os.getpid()
            
            # PowerShellスクリプトを作成（より確実な実行のため）
            with open(script_path, 'w', encoding='utf-8-sig') as f:
                f.write(f'''# Product Registration Tool Auto Update Script
Write-Host "=======================================" -ForegroundColor Green
Write-Host "   Product Registration Tool Auto Update" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Green
Write-Host ""
Write-Host "Start Time: $(Get-Date)" -ForegroundColor Yellow
Write-Host "Process ID: {current_pid}" -ForegroundColor Yellow
Write-Host "App Directory: {app_dir}" -ForegroundColor Yellow
Write-Host "Executable: {exe_path}" -ForegroundColor Yellow
Write-Host "New File: {new_file_path}" -ForegroundColor Yellow
Write-Host ""

# Check for .new file
Write-Host "[1/4] Checking update file..." -ForegroundColor Cyan
if (Test-Path "{new_file_path}") {{
    Write-Host "  ✓ New file exists" -ForegroundColor Green
    $newFileSize = (Get-Item "{new_file_path}").Length
    Write-Host "  File size: $($newFileSize / 1MB) MB" -ForegroundColor Green
}} else {{
    Write-Host "  ✗ Error: New file not found" -ForegroundColor Red
    Write-Host "  Expected path: {new_file_path}" -ForegroundColor Red
    Write-Host "  Directory contents:" -ForegroundColor Red
    Get-ChildItem "{app_dir}" -Filter "*.new" | ForEach-Object {{ Write-Host "    $($_.Name)" }}
    if (-not (Get-ChildItem "{app_dir}" -Filter "*.new")) {{
        Write-Host "    (No .new files found)" -ForegroundColor Red
    }}
    Write-Host ""
    Write-Host "  Update cancelled. Starting original application..." -ForegroundColor Yellow
    Start-Process "{exe_path}" -ErrorAction SilentlyContinue
    if ($?) {{
        Write-Host "  ✓ Application started" -ForegroundColor Green
    }} else {{
        Write-Host "  ✗ Failed to start application" -ForegroundColor Red
        Write-Host "  Please start manually: {exe_path}" -ForegroundColor Red
    }}
    Read-Host "Press Enter to continue"
    exit 1
}}

# Wait for process to end
Write-Host "[2/4] Waiting for process to end..." -ForegroundColor Cyan
$maxWait = 30
$count = 0
while ($count -lt $maxWait) {{
    $process = Get-Process -Id {current_pid} -ErrorAction SilentlyContinue
    if (-not $process) {{
        Write-Host "  ✓ Process termination confirmed" -ForegroundColor Green
        break
    }}
    if ($count -ge $maxWait) {{
        Write-Host "  ⚠ Timeout: Forcing process termination" -ForegroundColor Yellow
        Stop-Process -Id {current_pid} -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        break
    }}
    Start-Sleep -Seconds 1
    $count++
    Write-Host "  Waiting... ($count/$maxWait seconds)" -ForegroundColor Gray
}}

# Replace file
Write-Host "[3/4] Replacing file..." -ForegroundColor Cyan
$maxRetries = 3
$attempt = 0
$success = $false

while ($attempt -lt $maxRetries -and -not $success) {{
    try {{
        Move-Item "{new_file_path}" "{exe_path}" -Force -ErrorAction Stop
        Write-Host "  ✓ File replacement completed" -ForegroundColor Green
        $success = $true
    }} catch {{
        $attempt++
        Write-Host "  ✗ File replacement failed (attempt $attempt/$maxRetries)" -ForegroundColor Red
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
        if ($attempt -lt $maxRetries) {{
            Write-Host "  Retrying in 2 seconds..." -ForegroundColor Yellow
            Start-Sleep -Seconds 2
        }}
    }}
}}

if (-not $success) {{
    Write-Host "  ✗ Failed to replace file after $maxRetries attempts" -ForegroundColor Red
    Read-Host "Press Enter to continue"
    exit 1
}}

# Start application
Write-Host "[4/4] Starting application..." -ForegroundColor Cyan
Write-Host "  Start path: {exe_path}" -ForegroundColor Gray
try {{
    Start-Process "{exe_path}" -ErrorAction Stop
    Write-Host "  ✓ Application started successfully" -ForegroundColor Green
}} catch {{
    Write-Host "  ✗ Failed to start application" -ForegroundColor Red
    Write-Host "  Please start manually: {exe_path}" -ForegroundColor Red
    Read-Host "Press Enter to continue"
    exit 1
}}

Write-Host ""
Write-Host "Update process completed successfully!" -ForegroundColor Green
Write-Host "This window will close automatically in 3 seconds..." -ForegroundColor Yellow
Start-Sleep -Seconds 3

# Clean up script file
Remove-Item $MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue
''')
            
            # PowerShellスクリプトを実行（デバッグのためコンソール表示）
            # Windows環境でのみ実行
            try:
                # PowerShellを管理者権限で実行
                # CREATE_NEW_CONSOLE = 0x00000010
                powershell_cmd = ['powershell.exe', '-ExecutionPolicy', 'Bypass', '-File', script_path]
                subprocess.Popen(powershell_cmd, creationflags=0x00000010)
                logging.info("✅ PowerShell再起動スクリプト実行開始")
            except (ValueError, OSError) as e:
                # Linux等での実行時はスクリプト生成のみ
                logging.info(f"ℹ️ PowerShellスクリプト生成完了（実行は次回Windows環境で）: {e}")
                if not sys.platform.startswith('win'):
                    logging.info("  → Linuxテスト環境のため、スクリプト実行をスキップしました")
            
            # 現在のアプリケーションを優雅に終了
            QApplication.quit()
        else:
            # Unix系の場合
            script_path = os.path.join(os.path.dirname(sys.executable), 'restart.sh')
            with open(script_path, 'w') as f:
                f.write(f'''#!/bin/bash
sleep 2
if [ -f "{sys.executable}.new" ]; then
    mv -f "{sys.executable}.new" "{sys.executable}"
fi
"{sys.executable}" &
rm -f "$0"
''')
            os.chmod(script_path, 0o755)
            subprocess.Popen(['/bin/bash', script_path])
    
    def _handle_update_result(self):
        """更新結果をUIスレッドで処理（ダイアログが消えないように）"""
        try:
            logging.info(f"更新結果処理開始: success={self._update_result}")
            
            if self._update_result and hasattr(self, '_update_message'):
                # 更新成功 - 再起動確認
                logging.info("更新成功 - 再起動確認ダイアログ表示")
                
                # メインウィンドウを親として確実に指定
                parent = self.parent
                if parent is None:
                    parent = QApplication.activeWindow()
                
                reply = QMessageBox.question(
                    parent,
                    "更新完了",
                    f"{self._update_message}\n\n"
                    f"📋 重要: 更新を適用するにはアプリケーションの再起動が必要です\n\n"
                    f"💾 現在の作業内容は自動保存されています\n"
                    f"🔄 再起動中は一時的にアプリが終了します（数秒程度）\n"
                    f"✅ 更新後は最新バージョンで再開されます\n\n"
                    f"今すぐアプリケーションを再起動しますか？",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )
                
                logging.info(f"ユーザー選択: {reply}")
                
                if reply == QMessageBox.Yes:
                    try:
                        logging.info("再起動スクリプト実行")
                        self._create_restart_script()
                        logging.info("再起動スクリプト実行完了")
                    except Exception as e:
                        logging.error(f"再起動エラー: {e}")
                        QMessageBox.warning(
                            parent,
                            "再起動エラー",
                            "手動でアプリケーションを再起動してください。"
                        )
                else:
                    QMessageBox.information(
                        parent,
                        "更新予定",
                        "更新は次回起動時に適用されます。"
                    )
                    
            elif self._update_result is False and hasattr(self, '_update_message'):
                # 更新失敗
                logging.error(f"更新失敗: {self._update_message}")
                
                parent = self.parent if self.parent else QApplication.activeWindow()
                
                QMessageBox.critical(
                    parent, 
                    "更新エラー", 
                    f"更新中にエラーが発生しました:\n\n{self._update_message}\n\n"
                    f"問題が継続する場合は開発者にご報告ください。"
                )
            else:
                logging.info("更新がキャンセルされました")
                
        except Exception as e:
            logging.error(f"更新結果処理エラー: {e}", exc_info=True)
            # 最後の手段として基本的なエラーダイアログを表示
            try:
                QMessageBox.critical(
                    None,
                    "更新処理エラー",
                    f"更新の完了処理中にエラーが発生しました:\n{e}"
                )
            except:
                pass
        finally:
            # 状態をクリア
            self._update_result = None
            self._update_message = None
    
    def _detect_application_directory(self) -> str:
        """実行中のアプリケーションディレクトリを確実に検出"""
        try:
            candidate_dirs = []
            
            # 方法1: 現在の作業ディレクトリをチェック（最も確実）
            cwd = os.getcwd()
            logging.info(f"現在の作業ディレクトリ: {cwd}")
            candidate_dirs.append(cwd)
            
            # 方法2: sys.argv[0]から取得
            if sys.argv and sys.argv[0]:
                script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
                logging.info(f"sys.argv[0]ディレクトリ: {script_dir}")
                candidate_dirs.append(script_dir)
            
            # 方法3: sys.executableから取得
            if getattr(sys, 'frozen', False):
                exe_dir = os.path.dirname(sys.executable)
                logging.info(f"sys.executableディレクトリ: {exe_dir}")
                candidate_dirs.append(exe_dir)
            
            # 方法4: __file__から取得（開発環境用）
            try:
                this_file_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
                logging.info(f"__file__ベースディレクトリ: {this_file_dir}")
                candidate_dirs.append(this_file_dir)
            except:
                pass
            
            # 各候補をチェックして最適なものを選択
            for candidate in candidate_dirs:
                if self._validate_app_directory(candidate):
                    logging.info(f"✅ 確定されたアプリケーションディレクトリ: {candidate}")
                    return candidate
                    
            # 全て失敗した場合のフォールバック
            logging.warning("⚠️ 適切なディレクトリが見つからず、作業ディレクトリを使用")
            return cwd
            
        except Exception as e:
            logging.error(f"アプリケーションディレクトリ検出エラー: {e}")
            return os.getcwd()
    
    def _validate_app_directory(self, directory: str) -> bool:
        """ディレクトリがアプリケーションディレクトリかどうかを検証"""
        try:
            # 重要ファイルの存在を確認
            important_files = [
                'product_app.py',      # Pythonスクリプト版
                '商品登録入力ツール.exe',   # ビルド版
                'constants.py',        # 設定ファイル
                'version.json'         # バージョン情報
            ]
            
            found_files = 0
            for file in important_files:
                if os.path.exists(os.path.join(directory, file)):
                    found_files += 1
                    logging.debug(f"  ✓ {file}")
                    
            # 2つ以上の重要ファイルがあれば有効
            is_valid = found_files >= 2
            logging.info(f"ディレクトリ検証: {directory} -> {found_files}/4ファイル → {'有効' if is_valid else '無効'}")
            return is_valid
            
        except Exception as e:
            logging.error(f"ディレクトリ検証エラー {directory}: {e}")
            return False
    


def check_for_updates_simple(parent=None, silent=False):
    """
    シンプルな更新チェック（通知のみ）
    """
    try:
        
        # GitHub APIから最新バージョンを取得
        response = requests.get("https://api.github.com/repos/SEI1026/Product_app/releases/latest", timeout=5)
        latest = response.json()
        latest_version = latest['tag_name'].replace('v', '')
        
        logging.info(f"バージョンチェック: 現在={CURRENT_VERSION}, 最新={latest_version}")
        
        if latest_version > CURRENT_VERSION:
            # 新しいバージョンがある場合
            result = QMessageBox.question(
                parent, 
                "更新のお知らせ", 
                f"新しいバージョン v{latest_version} が利用可能です。\n\n"
                f"現在のバージョン: v{CURRENT_VERSION}\n"
                f"最新のバージョン: v{latest_version}\n\n"
                f"自動更新を実行しますか？",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if result == QMessageBox.Yes:
                # GitHubリリースからダウンロードURLを構築
                download_url = f"https://github.com/SEI1026/Product_app/releases/download/v{latest_version}/ProductRegisterTool-v{latest_version}.zip"
                simple_auto_update(parent, download_url, latest_version)
                
        elif not silent:
            QMessageBox.information(
                parent,
                "更新確認",
                f"お使いのバージョン v{CURRENT_VERSION} は最新です。"
            )
            
    except Exception as e:
        logging.error(f"更新チェックエラー: {e}")
        if not silent:
            QMessageBox.warning(
                parent,
                "更新確認エラー", 
                "更新の確認中にエラーが発生しました。\n"
                "インターネット接続を確認してください。"
            )


def simple_auto_update(parent, download_url, new_version):
    """
    シンプルな自動更新（データ移行付き）
    """
    
    try:
        # プログレスダイアログ
        progress = QProgressDialog("更新をダウンロード中...", "キャンセル", 0, 100, parent)
        progress.setWindowModality(2)  # Qt.WindowModal
        progress.show()
        QApplication.processEvents()
        
        # 1. 新バージョンをダウンロード
        logging.info(f"更新ダウンロード開始: {download_url}")
        response = requests.get(download_url, stream=True)
        total_size = int(response.headers.get('content-length', 0))
        
        # 元のツールの親ディレクトリにダウンロード（権限チェック付き）
        current_dir = os.path.dirname(os.path.abspath(sys.argv[0])) if getattr(sys, 'frozen', False) else os.getcwd()
        parent_dir = os.path.dirname(current_dir)
        
        # 親ディレクトリへの書き込み権限をチェック
        try:
            # テスト用一時ファイルで書き込み権限を確認
            test_file = os.path.join(parent_dir, f"test_write_{os.getpid()}.tmp")
            with open(test_file, 'w') as f:
                f.write('test')
            os.remove(test_file)
            download_dir = parent_dir
            logging.info(f"親ディレクトリへの書き込み権限OK: {parent_dir}")
        except (PermissionError, OSError) as e:
            logging.warning(f"親ディレクトリ権限なし({parent_dir}): {e}")
            
            # UAC昇格を試行
            if _try_uac_elevation(parent_dir, new_version, download_url):
                logging.info("UAC昇格による更新が完了しました")
                return  # UAC昇格で成功した場合はここで終了
            
            # UAC昇格が失敗または拒否された場合はフォールバック
            download_dir = tempfile.gettempdir()
            logging.info(f"フォールバック: 一時ディレクトリを使用: {download_dir}")
            QMessageBox.information(
                parent,
                "保存場所の変更",
                f"管理者権限が取得できないため、更新ファイルを一時フォルダに保存します:\n{download_dir}\n\n"
                f"※完了後、手動で適切な場所に移動してください。"
            )
        
        temp_zip = os.path.join(download_dir, f"ProductRegisterTool-v{new_version}.zip")
        
        with open(temp_zip, 'wb') as f:
            downloaded = 0
            for chunk in response.iter_content(chunk_size=8192):
                if progress.wasCanceled():
                    return
                    
                f.write(chunk)
                downloaded += len(chunk)
                if total_size > 0:
                    progress.setValue(int(downloaded * 50 / total_size))  # 50%まで
                QApplication.processEvents()
        
        progress.setLabelText("更新ファイルを展開中...")
        progress.setValue(60)
        QApplication.processEvents()
        
        # 2. ZIPを展開（権限に応じたディレクトリに）
        extract_dir = os.path.join(download_dir, f"ProductRegisterTool-v{new_version}")
        if os.path.exists(extract_dir):
            shutil.rmtree(extract_dir)
            
        safe_extract_zip(temp_zip, extract_dir)
        
        # 展開されたフォルダ内のメインフォルダを見つける
        logging.info(f"🔍 展開ディレクトリ内容確認: {extract_dir}")
        try:
            all_items = os.listdir(extract_dir)
            logging.info(f"📁 展開されたアイテム: {all_items}")
            
            subdirs = [d for d in all_items if os.path.isdir(os.path.join(extract_dir, d))]
            files = [f for f in all_items if os.path.isfile(os.path.join(extract_dir, f))]
            
            logging.info(f"📁 サブディレクトリ: {subdirs}")
            logging.info(f"📄 ファイル: {files}")
            
            # 実行ファイルを直接探す
            exe_found = False
            main_folder = extract_dir
            
            # まず、直接実行ファイルがあるかチェック
            direct_exe = os.path.join(extract_dir, "商品登録入力ツール.exe")
            if os.path.exists(direct_exe):
                logging.info(f"🔥 実行ファイルを直接発見: {direct_exe}")
                main_folder = extract_dir
                exe_found = True
            else:
                # サブディレクトリを検索
                for subdir in subdirs:
                    subdir_path = os.path.join(extract_dir, subdir)
                    sub_exe = os.path.join(subdir_path, "商品登録入力ツール.exe")
                    if os.path.exists(sub_exe):
                        logging.info(f"🔥 実行ファイルをサブディレクトリで発見: {sub_exe}")
                        main_folder = subdir_path
                        exe_found = True
                        break
                
                if not exe_found and subdirs:
                    # フォールバック: 最初のサブディレクトリを使用
                    main_folder = os.path.join(extract_dir, subdirs[0])
                    logging.info(f"⚠️ フォールバック: 最初のサブディレクトリを使用: {main_folder}")
            
            logging.info(f"🎯 選択されたメインフォルダ: {main_folder}")
            
        except Exception as e:
            logging.error(f"❌ 展開ディレクトリ確認エラー: {e}")
            main_folder = extract_dir
            
        progress.setValue(80)
        QApplication.processEvents()
        
        # 3. ユーザーデータを新フォルダにコピー
        progress.setLabelText("ユーザーデータを移行中...")
        user_files = ["item_manage.xlsm"]
        current_dir = os.path.dirname(os.path.abspath(sys.argv[0])) if getattr(sys, 'frozen', False) else os.getcwd()
        
        for file in user_files:
            source_file = os.path.join(current_dir, file)
            if os.path.exists(source_file):
                dest_file = os.path.join(main_folder, file)
                shutil.copy2(source_file, dest_file)
                logging.info(f"ユーザーデータをコピー: {source_file} → {dest_file}")
        
        progress.setValue(100)
        progress.close()
        
        # 4. 完了通知と新バージョン起動
        QMessageBox.information(
            parent,
            "更新完了",
            f"更新が完了しました！\n\n"
            f"新しいバージョン v{new_version} を起動します。\n"
            f"新しいフォルダ: {main_folder}\n\n"
            f"※起動に失敗した場合は、上記フォルダの実行ファイルを手動で起動してください。"
        )
        
        # 5. 新バージョン起動 & 現在終了
        new_exe = os.path.join(main_folder, "商品登録入力ツール.exe")
        logging.info(f"🔍 実行ファイル最終確認: {new_exe}")
        logging.info(f"🔍 実行ファイル存在: {os.path.exists(new_exe)}")
        
        if os.path.exists(new_exe):
            logging.info(f"🚀 新バージョン起動: {new_exe}")
            
            try:
                
                # Windowsの場合、絶対パスで起動
                abs_exe = os.path.abspath(new_exe)
                abs_cwd = os.path.abspath(main_folder)
                
                logging.info(f"🚀 起動コマンド: {abs_exe}")
                logging.info(f"🚀 作業ディレクトリ: {abs_cwd}")
                
                # アプリケーションロックを解放
                from ..utils.file_lock_manager import FileLockManager
                lock_manager = FileLockManager()
                lock_manager.release_app_lock()
                logging.info("🔓 アプリケーションロックを解放しました")
                
                # 少し待ってロックファイルが確実に削除されるようにする
                time.sleep(0.5)
                
                # 新しいプロセスをバックグラウンドで起動
                subprocess.Popen([abs_exe], 
                               cwd=abs_cwd,
                               creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if hasattr(subprocess, 'CREATE_NEW_PROCESS_GROUP') else 0)
                
                # 少し待ってから終了
                time.sleep(1)
                
                logging.info("🚀 新バージョン起動完了、現在のアプリを終了")
                QApplication.quit()
                
            except Exception as start_error:
                logging.error(f"❌ 新バージョン起動エラー: {start_error}")
                error_msg = f"新しいバージョンの起動に失敗しました。\n\n"
                error_msg += f"エラー: {start_error}\n\n"
                error_msg += f"手動で以下のファイルを実行してください:\n{new_exe}"
                QMessageBox.warning(parent, "起動エラー", error_msg)
        else:
            # 詳細なエラー情報
            error_msg = f"新しい実行ファイルが見つかりません。\n\n"
            error_msg += f"期待されるパス: {new_exe}\n"
            error_msg += f"メインフォルダ: {main_folder}\n"
            error_msg += f"一時展開先: {extract_dir}\n\n"
            error_msg += f"手動で以下のフォルダを確認してください:\n{extract_dir}"
            
            logging.error(f"❌ 実行ファイルが見つかりません: {new_exe}")
            logging.error(f"❌ メインフォルダ内容:")
            try:
                if os.path.exists(main_folder):
                    items = os.listdir(main_folder)
                    for item in items:
                        item_path = os.path.join(main_folder, item)
                        if os.path.isfile(item_path):
                            logging.error(f"  📄 {item}")
                        else:
                            logging.error(f"  📁 {item}/")
                else:
                    logging.error(f"❌ メインフォルダが存在しません: {main_folder}")
            except Exception as e:
                logging.error(f"❌ フォルダ内容確認エラー: {e}")
            
            QMessageBox.warning(parent, "更新エラー", error_msg)
            
        # デバッグのため一時ファイルを保持
        logging.info(f"📁 ダウンロードファイル保存先: {temp_zip}")
        logging.info(f"📁 展開フォルダ保存先: {extract_dir}")
        # os.remove(temp_zip)  # デバッグ用にコメントアウト
        
    except Exception as e:
        logging.error(f"自動更新エラー: {e}")
        QMessageBox.critical(
            parent,
            "更新エラー",
            f"更新中にエラーが発生しました:\n{e}\n\n"
            f"手動でGitHubから最新版をダウンロードしてください。"
        )


def _try_uac_elevation(parent_dir: str, new_version: str, download_url: str) -> bool:
    """
    UAC昇格を使用して管理者権限で更新を実行
    """
    try:
        import platform
        import sys
        
        # Windows以外では対応しない
        if platform.system() != 'Windows':
            logging.info("非Windows環境のため、UAC昇格をスキップ")
            return False
        
        logging.info("UAC昇格を試行中...")
        
        # 確認ダイアログ
        reply = QMessageBox.question(
            None,
            "管理者権限が必要",
            f"更新ファイルを以下の場所に保存するには管理者権限が必要です:\n{parent_dir}\n\n"
            f"管理者として実行しますか？\n\n"
            f"※「いいえ」を選択すると一時フォルダに保存されます。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )
        
        if reply != QMessageBox.Yes:
            logging.info("ユーザーがUAC昇格を拒否しました")
            return False
        
        # 現在のスクリプトファイルパスを取得
        if getattr(sys, 'frozen', False):
            # PyInstallerでビルドされた場合
            current_exe = sys.executable
        else:
            # 開発環境の場合
            current_exe = sys.executable
            script_args = [__file__] if '__file__' in globals() else []
        
        # UAC昇格用のコマンドライン引数を作成
        elevation_args = [
            current_exe,
            '-c',
            f'''
import sys
import os
import tempfile
import zipfile
import shutil
import requests
import logging

# ログ設定
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

try:
    logging.info("UAC昇格プロセス開始")
    logging.info("Download URL: {download_url}")
    logging.info("Target dir: {parent_dir}")
    logging.info("Version: {new_version}")
    
    # ダウンロード
    response = requests.get("{download_url}", stream=True)
    temp_zip = os.path.join("{parent_dir}", "ProductRegisterTool-v{new_version}.zip")
    
    with open(temp_zip, "wb") as f:
        for chunk in response.iter_content(chunk_size=8192):
            f.write(chunk)
    
    # 展開
    extract_dir = os.path.join("{parent_dir}", "ProductRegisterTool-v{new_version}")
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    
    safe_extract_zip(temp_zip, extract_dir)
    
    logging.info("UAC昇格による更新完了")
    
except Exception as e:
    logging.error(f"UAC昇格エラー: {e}")
    sys.exit(1)
'''
        ]
        
        # PowerShellを使用してUAC昇格実行
        import subprocess
        
        # PowerShellコマンドを作成
        ps_command = f'''
Start-Process -FilePath "python" -ArgumentList @("-c", @"
import sys
import os
import tempfile
import zipfile
import shutil
import requests
import logging

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

try:
    logging.info('UAC昇格プロセス開始')
    logging.info('Download URL: {download_url}')
    logging.info('Target dir: {parent_dir}')
    logging.info('Version: {new_version}')
    
    # ダウンロード
    response = requests.get('{download_url}', stream=True)
    temp_zip = os.path.join(r'{parent_dir}', 'ProductRegisterTool-v{new_version}.zip')
    
    with open(temp_zip, 'wb') as f:
        for chunk in response.iter_content(chunk_size=8192):
            f.write(chunk)
    
    # 展開
    extract_dir = os.path.join(r'{parent_dir}', 'ProductRegisterTool-v{new_version}')
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    
    safe_extract_zip(temp_zip, extract_dir)
    
    logging.info('UAC昇格による更新完了')
    
except Exception as e:
    logging.error(f'UAC昇格エラー: {e}')
    import sys
    sys.exit(1)
"@) -Verb RunAs -Wait
'''
        
        # PowerShellスクリプトを一時ファイルに保存
        ps_script_path = os.path.join(tempfile.gettempdir(), f"uac_elevation_{os.getpid()}.ps1")
        with open(ps_script_path, 'w', encoding='utf-8') as f:
            f.write(ps_command)
        
        # PowerShellを実行
        result = subprocess.run([
            'powershell.exe', 
            '-ExecutionPolicy', 'Bypass',
            '-File', ps_script_path
        ], capture_output=True, text=True, timeout=300)
        
        # 一時ファイルを削除
        try:
            os.remove(ps_script_path)
        except:
            pass
        
        if result.returncode == 0 and "SUCCESS" in result.stdout:
            logging.info("UAC昇格による更新が成功しました")
            
            # 成功メッセージ
            QMessageBox.information(
                None,
                "更新完了",
                f"管理者権限で更新が完了しました！\n\n"
                f"新しいバージョン v{new_version} が以下に保存されました:\n"
                f"{parent_dir}\\ProductRegisterTool-v{new_version}\\"
            )
            return True
        else:
            logging.error(f"UAC昇格が失敗しました: {result.stderr}")
            return False
            
    except subprocess.TimeoutExpired:
        logging.error("UAC昇格がタイムアウトしました")
        return False
    except Exception as e:
        logging.error(f"UAC昇格エラー: {e}")
        return False


def check_for_updates_on_startup(parent=None):
    """
    起動時の自動更新チェック（新しいシンプル版）
    """
    check_for_updates_simple(parent, silent=True)
