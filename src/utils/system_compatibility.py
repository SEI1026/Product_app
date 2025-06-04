"""
システム互換性チェック - OS、ランタイム、依存関係の確認
"""

import os
import sys
import platform
import subprocess
import logging
from typing import Dict, List, Tuple, Optional
from packaging import version
from PyQt5.QtWidgets import QMessageBox

# Try to import pkg_resources with fallback for newer Python versions
try:
    from importlib.metadata import version as get_version, PackageNotFoundError
    _use_importlib = True
except ImportError:
    # Fall back to pkg_resources for older Python versions
    import pkg_resources
    _use_importlib = False

class SystemCompatibilityChecker:
    """システム互換性のチェック"""
    
    def __init__(self):
        self.requirements = {
            "python_min_version": "3.8.0",
            "python_max_version": "3.13.99",  # 現在サポートされている最新
            "required_packages": {
                "PyQt5": "5.15.0",
                "openpyxl": "3.0.0",
                "requests": "2.25.0"
            },
            "system_requirements": {
                "windows": {
                    "min_version": "10.0.0",
                    "required_features": [".NET Framework"]
                },
                "darwin": {  # macOS
                    "min_version": "10.15.0",
                    "required_features": []
                },
                "linux": {
                    "min_version": "18.04",  # Ubuntu version
                    "required_features": []
                }
            }
        }
    
    def check_python_version(self) -> Dict[str, any]:
        """Pythonバージョンの互換性チェック"""
        current_version = platform.python_version()
        min_version = self.requirements["python_min_version"]
        max_version = self.requirements["python_max_version"]
        
        try:
            is_compatible = (
                version.parse(current_version) >= version.parse(min_version) and
                version.parse(current_version) <= version.parse(max_version)
            )
            
            return {
                "compatible": is_compatible,
                "current_version": current_version,
                "min_required": min_version,
                "max_supported": max_version,
                "recommendation": "Python 3.9-3.11を推奨" if not is_compatible else "OK"
            }
            
        except Exception as e:
            logging.error(f"Pythonバージョンチェックエラー: {e}")
            return {"compatible": False, "error": str(e)}
    
    def check_required_packages(self) -> Dict[str, Dict]:
        """必須パッケージの確認"""
        results = {}
        
        for package_name, min_version in self.requirements["required_packages"].items():
            try:
                # インストール済みバージョンを確認
                if _use_importlib:
                    try:
                        installed_version = get_version(package_name)
                        is_compatible = version.parse(installed_version) >= version.parse(min_version)
                        
                        results[package_name] = {
                            "installed": True,
                            "version": installed_version,
                            "min_required": min_version,
                            "compatible": is_compatible
                        }
                    except PackageNotFoundError:
                        results[package_name] = {
                            "installed": False,
                            "version": None,
                            "min_required": min_version,
                            "compatible": False
                        }
                else:
                    installed_version = pkg_resources.get_distribution(package_name).version
                    is_compatible = version.parse(installed_version) >= version.parse(min_version)
                    
                    results[package_name] = {
                        "installed": True,
                        "version": installed_version,
                        "min_required": min_version,
                        "compatible": is_compatible
                    }
                
            except Exception as e:
                if not _use_importlib and hasattr(e, '__class__') and e.__class__.__name__ == 'DistributionNotFound':
                    results[package_name] = {
                        "installed": False,
                        "version": None,
                        "min_required": min_version,
                        "compatible": False
                    }
                else:
                    results[package_name] = {
                        "installed": False,
                        "error": str(e),
                        "compatible": False
                    }
        
        return results
    
    def check_system_requirements(self) -> Dict[str, any]:
        """システム要件の確認"""
        system_name = platform.system().lower()
        
        if system_name == "windows":
            return self._check_windows_requirements()
        elif system_name == "darwin":
            return self._check_macos_requirements()
        elif system_name == "linux":
            return self._check_linux_requirements()
        else:
            return {
                "compatible": False,
                "error": f"未サポートのOS: {system_name}"
            }
    
    def _check_windows_requirements(self) -> Dict[str, any]:
        """Windows固有の要件チェック"""
        try:
            # Windowsバージョン確認
            win_version = platform.version()
            win_release = platform.release()
            
            # .NET Framework確認
            dotnet_installed = self._check_dotnet_framework()
            
            # Windows 10以降かチェック
            is_win10_or_later = float(platform.release()) >= 10.0
            
            return {
                "compatible": is_win10_or_later and dotnet_installed,
                "os_version": f"Windows {win_release} ({win_version})",
                "dotnet_framework": dotnet_installed,
                "requirements_met": is_win10_or_later,
                "recommendations": [] if is_win10_or_later else ["Windows 10以降にアップグレードしてください"]
            }
            
        except Exception as e:
            return {"compatible": False, "error": str(e)}
    
    def _check_macos_requirements(self) -> Dict[str, any]:
        """macOS固有の要件チェック"""
        try:
            mac_version = platform.mac_ver()[0]
            min_version = self.requirements["system_requirements"]["darwin"]["min_version"]
            
            is_compatible = version.parse(mac_version) >= version.parse(min_version)
            
            return {
                "compatible": is_compatible,
                "os_version": f"macOS {mac_version}",
                "min_required": min_version,
                "requirements_met": is_compatible
            }
            
        except Exception as e:
            return {"compatible": False, "error": str(e)}
    
    def _check_linux_requirements(self) -> Dict[str, any]:
        """Linux固有の要件チェック"""
        try:
            # ディストリビューション情報を取得
            distro_info = platform.freedesktop_os_release()
            
            return {
                "compatible": True,  # Linuxは基本的に互換性あり
                "distribution": distro_info.get("NAME", "Unknown"),
                "version": distro_info.get("VERSION", "Unknown"),
                "requirements_met": True
            }
            
        except Exception as e:
            return {"compatible": False, "error": str(e)}
    
    def _check_dotnet_framework(self) -> bool:
        """.NET Frameworkの確認"""
        try:
            # レジストリから.NET Frameworkの情報を確認
            import winreg
            
            reg_path = r"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
                    version_value = winreg.QueryValueEx(key, "Version")[0]
                    return True
            except FileNotFoundError:
                return False
                
        except ImportError:
            # winregが利用できない場合（非Windows環境）
            return True
        except Exception:
            return False
    
    def check_available_memory(self) -> Dict[str, any]:
        """利用可能メモリの確認"""
        try:
            import psutil
            
            memory = psutil.virtual_memory()
            total_gb = memory.total / (1024**3)
            available_gb = memory.available / (1024**3)
            
            # 最小要件は2GB、推奨は4GB
            min_required_gb = 2
            recommended_gb = 4
            
            return {
                "total_memory_gb": total_gb,
                "available_memory_gb": available_gb,
                "meets_minimum": total_gb >= min_required_gb,
                "meets_recommended": total_gb >= recommended_gb,
                "recommendation": "4GB以上のRAMを推奨" if total_gb < recommended_gb else "OK"
            }
            
        except ImportError:
            return {"error": "psutilパッケージが必要です"}
        except Exception as e:
            return {"error": str(e)}
    
    def check_disk_space(self, path: str = ".") -> Dict[str, any]:
        """ディスク容量の確認"""
        try:
            import shutil
            
            total, used, free = shutil.disk_usage(path)
            
            total_gb = total / (1024**3)
            free_gb = free / (1024**3)
            
            # 最小要件は1GB、推奨は5GB
            min_required_gb = 1
            recommended_gb = 5
            
            return {
                "total_space_gb": total_gb,
                "free_space_gb": free_gb,
                "meets_minimum": free_gb >= min_required_gb,
                "meets_recommended": free_gb >= recommended_gb,
                "recommendation": "5GB以上の空き容量を推奨" if free_gb < recommended_gb else "OK"
            }
            
        except Exception as e:
            return {"error": str(e)}
    
    def generate_compatibility_report(self) -> Dict[str, any]:
        """包括的な互換性レポートを生成"""
        report = {
            "python": self.check_python_version(),
            "packages": self.check_required_packages(),
            "system": self.check_system_requirements(),
            "memory": self.check_available_memory(),
            "disk": self.check_disk_space(),
            "overall_compatible": True,
            "warnings": [],
            "errors": []
        }
        
        # 全体的な互換性を判定
        if not report["python"].get("compatible", False):
            report["overall_compatible"] = False
            report["errors"].append("Pythonバージョンが互換性要件を満たしていません")
        
        # パッケージチェック
        for pkg_name, pkg_info in report["packages"].items():
            if not pkg_info.get("compatible", False):
                if not pkg_info.get("installed", False):
                    report["errors"].append(f"必須パッケージ '{pkg_name}' がインストールされていません")
                else:
                    report["warnings"].append(f"パッケージ '{pkg_name}' のバージョンが古い可能性があります")
        
        # システム要件チェック
        if not report["system"].get("compatible", False):
            report["overall_compatible"] = False
            report["errors"].append("システム要件を満たしていません")
        
        # メモリ・ディスク容量の警告
        if not report["memory"].get("meets_recommended", True):
            report["warnings"].append("推奨メモリ容量を下回っています")
        
        if not report["disk"].get("meets_recommended", True):
            report["warnings"].append("推奨ディスク容量を下回っています")
        
        return report


def check_system_compatibility(parent=None) -> bool:
    """システム互換性をチェックしてユーザーに報告"""
    checker = SystemCompatibilityChecker()
    report = checker.generate_compatibility_report()
    
    if report["overall_compatible"] and not report["warnings"]:
        # 完全に互換性がある場合は通知しない
        return True
    
    # 警告またはエラーがある場合は表示
    message_parts = []
    
    if report["errors"]:
        message_parts.append("【エラー】")
        message_parts.extend([f"• {error}" for error in report["errors"]])
        message_parts.append("")
    
    if report["warnings"]:
        message_parts.append("【警告】")
        message_parts.extend([f"• {warning}" for warning in report["warnings"]])
        message_parts.append("")
    
    if report["overall_compatible"]:
        message_parts.append("アプリケーションは動作しますが、最適な性能のため上記の改善をお勧めします。")
        message_type = QMessageBox.Warning
        title = "システム互換性警告"
    else:
        message_parts.append("アプリケーションが正常に動作しない可能性があります。")
        message_type = QMessageBox.Critical
        title = "システム互換性エラー"
    
    QMessageBox(
        message_type,
        title,
        "\n".join(message_parts),
        QMessageBox.Ok,
        parent
    ).exec_()
    
    return report["overall_compatible"]


def get_system_info() -> str:
    """システム情報の詳細を取得（トラブルシューティング用）"""
    try:
        info_parts = [
            f"Python: {platform.python_version()}",
            f"OS: {platform.system()} {platform.release()} {platform.version()}",
            f"Architecture: {platform.machine()}",
            f"Processor: {platform.processor()}",
        ]
        
        # メモリ情報
        try:
            import psutil
            memory = psutil.virtual_memory()
            info_parts.append(f"Memory: {memory.total / (1024**3):.1f}GB total, {memory.available / (1024**3):.1f}GB available")
        except ImportError:
            pass
        
        # インストール済みパッケージ
        try:
            import PyQt5
            info_parts.append(f"PyQt5: {PyQt5.QtCore.PYQT_VERSION_STR}")
        except ImportError:
            info_parts.append("PyQt5: Not installed")
        
        try:
            import openpyxl
            info_parts.append(f"openpyxl: {openpyxl.__version__}")
        except ImportError:
            info_parts.append("openpyxl: Not installed")
        
        return "\n".join(info_parts)
        
    except Exception as e:
        return f"システム情報取得エラー: {e}"