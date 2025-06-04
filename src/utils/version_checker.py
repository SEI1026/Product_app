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
from typing import Optional, Dict, Any, Tuple
from urllib.request import urlopen, Request
from urllib.error import URLError, HTTPError
from PyQt5.QtCore import QThread, pyqtSignal, QObject
from PyQt5.QtWidgets import QMessageBox, QProgressDialog, QPushButton, QApplication

# 現在のアプリケーションバージョン
CURRENT_VERSION = "2.2.0"

# GitHub上のversion.jsonのURL
# 株式会社大宝家具の商品登録入力ツール
# ※実際のGitHubリポジトリURLに合わせて変更してください
VERSION_CHECK_URL = "https://raw.githubusercontent.com/SEI1026/Product_app/main/version.json"


class VersionInfo:
    """バージョン情報を格納するクラス"""
    
    def __init__(self, version_data: Dict[str, Any]):
        self.version = version_data.get("version", "0.0.0")
        self.release_date = version_data.get("release_date", "")
        self.download_url = version_data.get("download_url", "")
        self.changelog = version_data.get("changelog", {})
        self.minimum_required_version = version_data.get("minimum_required_version", "0.0.0")
        
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
        
    def run(self):
        """更新ファイルをダウンロードして展開"""
        extract_dir = None
        try:
            # キャンセルチェック
            if self._cancelled:
                return
                
            # 一時ディレクトリを使用（より安全）
            temp_dir = tempfile.gettempdir()
            self.temp_file = os.path.join(temp_dir, f'update_{os.getpid()}.zip')
            
            logging.info(f"一時ファイルパス: {self.temp_file}")
            logging.info(f"ダウンロードURL: {self.download_url}")
            
            # URL検証
            if not self.download_url or not self.download_url.startswith('https://'):
                raise Exception("無効なダウンロードURLです")
                
            self.status.emit("更新ファイルをダウンロード中...")
            
            # ダウンロード処理を安全にラップ
            try:
                self._download_file()
                
                if self._cancelled:
                    return
                    
                self.status.emit("更新ファイルを展開中...")
                
                # ZIPファイルの検証と展開
                extract_dir = self._extract_and_validate_zip()
                
                if self._cancelled:
                    return
                
                # ファイル更新処理
                self._update_files(extract_dir, self.target_dir)
                
                self.finished.emit(True, "更新が正常に完了しました")
                
            except Exception as process_e:
                error_msg = str(process_e)
                logging.error(f"更新処理エラー: {error_msg}")
                self.finished.emit(False, f"更新のダウンロードに失敗しました: {error_msg}")
            
        except Exception as e:
            error_msg = str(e)
            logging.error(f"更新ダウンロード中の予期しないエラー: {error_msg}")
            if not self._cancelled:
                self.finished.emit(False, f"更新中に予期しないエラーが発生しました: {error_msg}")
            
        finally:
            # クリーンアップ処理
            self._cleanup(extract_dir)
    
    def _download_file(self):
        """ファイルダウンロード処理"""
        req = Request(self.download_url, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        
        try:
            with urlopen(req, timeout=30) as response:
                # HTTPステータスコードをチェック
                if response.getcode() != 200:
                    raise Exception(f"HTTPエラー: {response.getcode()} - ダウンロードファイルが見つかりません")
                
                # Content-Lengthを安全に取得
                try:
                    total_size = int(response.headers.get('Content-Length', 0))
                except (ValueError, TypeError):
                    total_size = 0
                
                downloaded = 0
                logging.info(f"ダウンロード開始 - サイズ: {total_size} bytes")
                
                with open(self.temp_file, 'wb') as f:
                    while True:
                        if self._cancelled:
                            break
                            
                        chunk = response.read(8192)
                        if not chunk:
                            break
                            
                        f.write(chunk)
                        downloaded += len(chunk)
                        
                        # プログレス更新
                        if total_size > 0:
                            progress = min(int((downloaded / total_size) * 100), 100)
                            self.progress.emit(progress)
                            
        except HTTPError as e:
            raise Exception(f"HTTPエラー {e.code}: ダウンロードファイルが見つかりません。\nURL: {self.download_url}")
        except URLError as e:
            raise Exception(f"ネットワークエラー: {e.reason}")
        except OSError as e:
            raise Exception(f"ファイル書き込みエラー: {str(e)}")
        except Exception as e:
            raise Exception(f"ダウンロードエラー: {str(e)}")
    
    def _extract_and_validate_zip(self):
        """ZIPファイルの検証と展開"""
        # ZIPファイルの存在確認
        if not os.path.exists(self.temp_file):
            raise Exception("ダウンロードしたファイルが見つかりません")
            
        # ZIPファイルのサイズ確認
        file_size = os.path.getsize(self.temp_file)
        if file_size < 1000:  # 1KB未満の場合は無効なファイル
            raise Exception(f"ダウンロードしたファイルが不完全です（サイズ: {file_size} bytes）")
        
        # ZIPファイルの整合性確認と展開
        try:
            with zipfile.ZipFile(self.temp_file, 'r') as zip_ref:
                # ZIPファイルの整合性テスト
                bad_file = zip_ref.testzip()
                if bad_file:
                    raise Exception(f"ZIPファイルが破損しています: {bad_file}")
                
                # 一時ディレクトリに展開
                extract_dir = tempfile.mkdtemp(prefix='update_extract_')
                logging.info(f"展開先ディレクトリ: {extract_dir}")
                
                zip_ref.extractall(extract_dir)
                return extract_dir
                
        except zipfile.BadZipFile:
            raise Exception("ダウンロードしたファイルが有効なZIPファイルではありません")
        except PermissionError:
            raise Exception("ファイル展開の権限がありません")
        except OSError as e:
            raise Exception(f"ファイル展開中のシステムエラー: {str(e)}")
        except Exception as e:
            raise Exception(f"ファイル展開エラー: {str(e)}")
    
    def _cleanup(self, extract_dir=None):
        """クリーンアップ処理"""
        # 一時ファイルを削除
        if self.temp_file and os.path.exists(self.temp_file):
            try:
                os.unlink(self.temp_file)
                logging.info(f"一時ファイルを削除: {self.temp_file}")
            except Exception as e:
                logging.warning(f"一時ファイル削除エラー: {e}")
        
        # 展開ディレクトリを削除
        if extract_dir and os.path.exists(extract_dir):
            try:
                shutil.rmtree(extract_dir)
                logging.info(f"展開ディレクトリを削除: {extract_dir}")
            except Exception as e:
                logging.warning(f"展開ディレクトリ削除エラー: {e}")
    
    def terminate(self):
        """ダウンロードをキャンセル"""
        self._cancelled = True
        super().terminate()
    
    def _update_files(self, source_dir: str, target_dir: str):
        """ファイルを更新（実行中のファイルは.newとして保存、ユーザーデータは保護）"""
        current_exe = os.path.abspath(sys.executable)
        current_exe_name = os.path.basename(current_exe)
        updated_exe = False
        
        # ユーザーデータファイル（保護対象）のパターン
        protected_patterns = [
            'item_manage.xlsm',  # ユーザーの商品管理ファイル
            '*_user_*',          # ユーザー作成ファイル
            '*.backup',          # バックアップファイル  
            'user_settings.json', # ユーザー設定
            'config.ini',        # 設定ファイル
        ]
        
        # ユーザーデータのバックアップを作成
        backup_created = self._create_user_data_backup(target_dir)
        if backup_created:
            self.status.emit("ユーザーデータのバックアップを作成しました")
        
        # 展開されたファイルを探す
        for root, dirs, files in os.walk(source_dir):
            rel_path = os.path.relpath(root, source_dir)
            target_root = os.path.join(target_dir, rel_path) if rel_path != '.' else target_dir
            
            # ディレクトリを作成
            if not os.path.exists(target_root):
                os.makedirs(target_root, exist_ok=True)
            
            for file in files:
                source_file = os.path.join(root, file)
                target_file = os.path.join(target_root, file)
                
                # ユーザーデータファイルの保護チェック
                if self._is_user_data_file(file, rel_path, protected_patterns):
                    # 既存のユーザーデータファイルがある場合は保護
                    if os.path.exists(target_file):
                        self.status.emit(f"ユーザーデータを保護: {file}")
                        logging.info(f"ユーザーデータファイルを保護: {target_file}")
                        continue  # このファイルはスキップ
                
                # PyInstallerでビルドされたexeファイルの更新
                if getattr(sys, 'frozen', False):
                    # 実行ファイル名と一致する場合（商品登録ツール.exe など）
                    if file.lower() == current_exe_name.lower() or file.lower().endswith('.exe'):
                        # 実行中のexeファイルは.newとして保存
                        target_file = current_exe + '.new'
                        updated_exe = True
                        self.status.emit(f"実行ファイルを更新中: {file}")
                else:
                    # 開発環境の場合、実行中のスクリプトと同じ場合
                    if os.path.abspath(target_file) == current_exe:
                        target_file = target_file + '.new'
                        updated_exe = True
                
                # ファイルをコピー
                try:
                    shutil.copy2(source_file, target_file)
                    logging.info(f"更新ファイルをコピー: {source_file} -> {target_file}")
                except Exception as e:
                    logging.error(f"ファイルコピーエラー: {e}")
                    raise
        
        if not updated_exe and getattr(sys, 'frozen', False):
            # exeファイルが見つからなかった場合の警告
            logging.warning("更新パッケージ内に実行ファイルが見つかりませんでした")
    
    def _is_user_data_file(self, filename: str, rel_path: str, protected_patterns: list) -> bool:
        """ファイルがユーザーデータかどうかを判定"""
        import fnmatch
        
        # ファイル名パターンマッチング
        for pattern in protected_patterns:
            if fnmatch.fnmatch(filename.lower(), pattern.lower()):
                return True
        
        # 特定のディレクトリ内のファイル（C#ツール内など）
        if 'C#' in rel_path and filename.endswith('.xlsm'):
            return True
            
        # ファイルサイズ・更新日時による判定（テンプレートより大きい場合はユーザーデータの可能性）
        # item_manage.xlsmがitem_template.xlsmより大きい場合など
        
        return False
    
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
                source_file = os.path.join(target_dir, filename)
                if os.path.exists(source_file):
                    if not os.path.exists(backup_dir):
                        os.makedirs(backup_dir, exist_ok=True)
                    
                    backup_file = os.path.join(backup_dir, filename)
                    shutil.copy2(source_file, backup_file)
                    backup_created = True
                    logging.info(f"バックアップ作成: {source_file} -> {backup_file}")
            
            return backup_created
            
        except Exception as e:
            logging.error(f"バックアップ作成エラー: {e}")
            return False


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
            # GitHub APIからversion.jsonを取得
            req = Request(VERSION_CHECK_URL, headers={'User-Agent': 'Mozilla/5.0'})
            with urlopen(req, timeout=10) as response:
                version_data = json.loads(response.read().decode('utf-8'))
            
            version_info = VersionInfo(version_data)
            
            # バージョン比較
            if self._is_newer_version(version_info.version, CURRENT_VERSION):
                return version_info
            elif not silent:
                QMessageBox.information(
                    self.parent,
                    "更新確認",
                    f"お使いのバージョン ({CURRENT_VERSION}) は最新です。"
                )
            return None
            
        except (URLError, HTTPError) as e:
            self.logger.error(f"バージョンチェック中のネットワークエラー: {e}")
            if not silent:
                QMessageBox.warning(
                    self.parent,
                    "更新確認エラー",
                    "更新の確認中にエラーが発生しました。\nインターネット接続を確認してください。"
                )
            return None
            
        except UnicodeEncodeError as e:
            self.logger.error(f"バージョンチェック中のエンコーディングエラー: {e}")
            if not silent:
                QMessageBox.warning(
                    self.parent,
                    "更新確認エラー",
                    "更新の確認中にエンコーディングエラーが発生しました。\n設定を確認してください。"
                )
            return None
            
        except Exception as e:
            self.logger.error(f"バージョンチェック中の予期しないエラー: {e}")
            if not silent:
                QMessageBox.warning(
                    self.parent,
                    "更新確認エラー",
                    "更新の確認中にエラーが発生しました。\nネットワーク接続を確認してください。"
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
            
            # アプリケーションディレクトリを取得
            if getattr(sys, 'frozen', False):
                # PyInstallerでビルドされた場合
                app_dir = os.path.dirname(sys.executable)
            else:
                # 開発環境の場合
                app_dir = os.path.dirname(os.path.abspath(__file__))
            
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
                # 手動ダウンロード（従来の方法）
                progress.close()
                import webbrowser
                
                # GitHubリリースページを開く
                download_url_parts = version_info.download_url.split('/')
                if len(download_url_parts) >= 8 and download_url_parts[5] == 'releases':
                    tag_name = download_url_parts[7]  # v2.1.7
                    repo_path = '/'.join(download_url_parts[:5])  # https://github.com/SEI1026/Product_app
                    release_url = f"{repo_path}/releases/tag/{tag_name}"
                else:
                    # フォールバック: リリース一覧ページ
                    release_url = version_info.download_url.rsplit('/releases/', 1)[0] + '/releases'
                
                webbrowser.open(release_url)
                
                QMessageBox.information(
                    self.parent,
                    "ダウンロード開始",
                    "ブラウザでダウンロードページを開きました。\n"
                    "ダウンロード完了後、このアプリを終了してから\n"
                    "新しいバージョンをインストールしてください。"
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
                
                # プログレスバーの更新を安全に接続
                try:
                    downloader.progress.connect(progress.setValue)
                    downloader.status.connect(progress.setLabelText)
                except Exception as e:
                    logging.error(f"シグナル接続エラー: {e}")
                    progress.close()
                    QMessageBox.critical(
                        self.parent,
                        "更新エラー",
                        f"ダウンロード準備中にエラーが発生しました: {str(e)}"
                    )
                    return
                
                # 完了時の処理を安全にラップ
                def on_finished(success: bool, message: str):
                    try:
                        progress.close()
                        
                        if success:
                            # 更新成功
                            msg_box = QMessageBox(self.parent)
                            msg_box.setIcon(QMessageBox.Information)
                            msg_box.setWindowTitle("更新完了")
                            msg_box.setText(f"{message}")
                            msg_box.setInformativeText("今すぐアプリケーションを再起動して更新を適用しますか？")
                            
                            restart_btn = msg_box.addButton("今すぐ再起動", QMessageBox.AcceptRole)
                            later_btn = msg_box.addButton("後で再起動", QMessageBox.RejectRole)
                            msg_box.setDefaultButton(restart_btn)
                            
                            msg_box.exec_()
                            
                            if msg_box.clickedButton() == restart_btn:
                                # 自動再起動スクリプトを実行
                                try:
                                    self._create_restart_script()
                                except Exception as restart_e:
                                    logging.error(f"再起動スクリプト作成エラー: {restart_e}")
                                    QMessageBox.warning(
                                        self.parent,
                                        "再起動エラー",
                                        "自動再起動に失敗しました。手動でアプリケーションを再起動してください。"
                                    )
                            else:
                                # 次回起動時に更新が適用されることを通知
                                QMessageBox.information(
                                    self.parent,
                                    "更新予定",
                                    "更新は次回アプリケーション起動時に適用されます。"
                                )
                        else:
                            # 更新失敗
                            QMessageBox.critical(
                                self.parent,
                                "更新エラー",
                                message
                            )
                    except Exception as e:
                        logging.error(f"完了処理エラー: {e}")
                        try:
                            QMessageBox.critical(
                                self.parent,
                                "更新エラー",
                                f"更新完了処理中にエラーが発生しました: {str(e)}"
                            )
                        except:
                            pass  # メッセージボックス表示も失敗した場合
                
                try:
                    downloader.finished.connect(on_finished)
                    
                    # キャンセルボタンの処理
                    def on_cancel():
                        try:
                            downloader.terminate()
                            downloader.wait(3000)  # 3秒待機
                        except Exception as e:
                            logging.error(f"ダウンロードキャンセルエラー: {e}")
                    
                    progress.canceled.connect(on_cancel)
                    
                    # ダウンロード開始
                    downloader.start()
                    
                    # メインスレッドでイベントループを維持
                    while downloader.isRunning():
                        QApplication.processEvents()
                        downloader.msleep(50)
                    
                except Exception as e:
                    logging.error(f"ダウンロード開始エラー: {e}")
                    progress.close()
                    QMessageBox.critical(
                        self.parent,
                        "更新エラー",
                        f"ダウンロード開始中にエラーが発生しました: {str(e)}"
                    )
                    try:
                        downloader.terminate()
                    except:
                        pass
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
    
    def _create_restart_script(self):
        """再起動用のスクリプトを作成"""
        if sys.platform == 'win32':
            # Windowsの場合
            exe_path = sys.executable
            exe_dir = os.path.dirname(exe_path)
            exe_name = os.path.basename(exe_path)
            script_path = os.path.join(exe_dir, 'update_restart.bat')
            
            # バッチファイルを作成（日本語対応）
            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(f'''@echo off
chcp 65001 >nul
echo 更新を適用しています...
timeout /t 3 /nobreak > nul
:retry
if exist "{exe_path}.new" (
    taskkill /f /im "{exe_name}" >nul 2>&1
    timeout /t 1 /nobreak > nul
    move /y "{exe_path}.new" "{exe_path}" >nul 2>&1
    if errorlevel 1 (
        echo ファイルの置換に失敗しました。再試行します...
        timeout /t 2 /nobreak > nul
        goto retry
    )
)
echo 更新が完了しました。アプリケーションを起動します...
start "" "{exe_path}"
del "%~f0"
''')
            # バッチファイルを実行して即座に終了
            subprocess.Popen(['cmd', '/c', script_path], 
                           creationflags=subprocess.CREATE_NEW_CONSOLE)
            # アプリケーションを終了
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


def check_for_updates_on_startup(parent=None):
    """
    起動時の自動更新チェック（非同期）
    """
    checker = VersionChecker(parent)
    version_info = checker.check_for_updates(silent=True)
    
    if version_info:
        # 新しいバージョンがある場合
        if checker.prompt_for_update(version_info):
            checker.download_and_install_update(version_info)
