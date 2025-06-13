"""
セキュリティ検証システム - 入力値の検証とセキュリティチェック
"""

import os
import re
import html
import logging
from typing import Any, Dict, List, Union, Optional
from urllib.parse import urlparse


class SecurityValidator:
    """セキュリティ検証クラス"""
    
    def __init__(self):
        self.max_input_length = 1000
        self.dangerous_patterns = [
            r'<script[^>]*>.*?</script>',
            r'javascript:',
            r'vbscript:',
            r'on\w+\s*=',
            r'expression\s*\(',
            r'eval\s*\(',
            r'setTimeout\s*\(',
            r'setInterval\s*\(',
        ]
        self.sql_injection_patterns = [
            r'(\bOR\b|\bAND\b)\s+\d+\s*=\s*\d+',
            r'UNION\s+SELECT',
            r'DROP\s+TABLE',
            r'DELETE\s+FROM',
            r'INSERT\s+INTO',
            r'UPDATE\s+SET',
            r';\s*--',
            r'/\*.*?\*/',
        ]
    
    def validate_input(self, value: Any) -> str:
        """入力値の総合的な検証とサニタイゼーション"""
        if value is None:
            return ""
        
        # 文字列変換
        text_value = str(value)
        
        # 長さ制限
        if len(text_value) > self.max_input_length:
            logging.warning(f"入力値が最大長を超えています: {len(text_value)} > {self.max_input_length}")
            text_value = text_value[:self.max_input_length]
        
        # HTMLエスケープ
        sanitized = html.escape(text_value)
        
        # 制御文字の除去（タブ、改行、復帰文字以外）
        sanitized = ''.join(char for char in sanitized if ord(char) >= 32 or char in '\t\n\r')
        
        # 危険なパターンの検出と除去
        for pattern in self.dangerous_patterns:
            if re.search(pattern, sanitized, re.IGNORECASE):
                logging.error(f"セキュリティ警告: 危険なパターンが検出されました: {pattern}")
                sanitized = re.sub(pattern, '', sanitized, flags=re.IGNORECASE)
        
        # SQLインジェクションパターンの検出
        for pattern in self.sql_injection_patterns:
            if re.search(pattern, sanitized, re.IGNORECASE):
                logging.error(f"セキュリティ警告: SQLインジェクション可能性のあるパターン: {pattern}")
                # SQLパターンは除去せず、ログのみ記録（データ損失を避けるため）
        
        return sanitized.strip()
    
    def validate_file_path(self, filepath: str, allowed_dirs: List[str] = None) -> str:
        """ファイルパスの検証"""
        if not filepath:
            raise ValueError("ファイルパスが空です")
        
        # パス正規化
        normalized_path = os.path.normpath(filepath)
        
        # ディレクトリトラバーサル攻撃の検出
        if '..' in normalized_path.split(os.sep):
            raise ValueError(f"不正なファイルパス: ディレクトリトラバーサルが検出されました: {filepath}")
        
        # 絶対パスの場合の検証
        if os.path.isabs(normalized_path):
            abs_path = normalized_path
        else:
            abs_path = os.path.abspath(normalized_path)
        
        # 許可されたディレクトリ内かチェック
        if allowed_dirs:
            is_allowed = any(abs_path.startswith(os.path.abspath(allowed_dir)) for allowed_dir in allowed_dirs)
            if not is_allowed:
                raise ValueError(f"許可されていないディレクトリへのアクセス: {filepath}")
        
        return abs_path
    
    def validate_url(self, url: str) -> bool:
        """URL の検証"""
        if not url:
            return False
        
        try:
            parsed = urlparse(url)
            # HTTPSまたはHTTPのみ許可
            if parsed.scheme not in ['http', 'https']:
                logging.warning(f"許可されていないURLスキーム: {parsed.scheme}")
                return False
            
            # ローカルネットワークへのアクセスを制限
            hostname = parsed.hostname
            if hostname:
                # プライベートIPアドレスの検出
                import ipaddress
                try:
                    ip = ipaddress.ip_address(hostname)
                    if ip.is_private or ip.is_loopback:
                        logging.warning(f"プライベートIPアドレスへのアクセス試行: {hostname}")
                        return False
                except ValueError:
                    # ホスト名の場合
                    if hostname.lower() in ['localhost', '127.0.0.1', '::1']:
                        logging.warning(f"ローカルホストへのアクセス試行: {hostname}")
                        return False
            
            return True
        except Exception as e:
            logging.error(f"URL検証エラー: {e}")
            return False
    
    def validate_numeric_input(self, value: Any, min_val: Optional[float] = None, max_val: Optional[float] = None) -> Union[int, float, None]:
        """数値入力の検証"""
        if value is None or value == "":
            return None
        
        try:
            # 文字列から数値への変換を試行
            if isinstance(value, str):
                value = value.strip()
                if '.' in value:
                    num_value = float(value)
                else:
                    num_value = int(value)
            else:
                num_value = float(value)
            
            # 範囲チェック
            if min_val is not None and num_value < min_val:
                raise ValueError(f"値が最小値を下回っています: {num_value} < {min_val}")
            if max_val is not None and num_value > max_val:
                raise ValueError(f"値が最大値を上回っています: {num_value} > {max_val}")
            
            return num_value
        except (ValueError, TypeError) as e:
            logging.error(f"数値変換エラー: {e}")
            raise ValueError(f"無効な数値形式: {value}")
    
    def check_data_integrity(self, data: Dict) -> Dict[str, Any]:
        """データ整合性の検証"""
        issues = []
        
        # 必須フィールドのチェック（例）
        required_fields = ['mycode', '商品名_正式表記']
        for field in required_fields:
            if field not in data or not data[field]:
                issues.append(f"必須フィールドが不足: {field}")
        
        # データ型チェック
        numeric_fields = ['当店通常価格_税込み', 'ソート']
        for field in numeric_fields:
            if field in data and data[field]:
                try:
                    self.validate_numeric_input(data[field], min_val=0)
                except ValueError as e:
                    issues.append(f"数値フィールドエラー ({field}): {e}")
        
        # 文字数制限チェック
        length_limits = {
            'mycode': 20,
            '商品名_正式表記': 255,
            'Y_キャッチコピー': 60,
        }
        for field, limit in length_limits.items():
            if field in data and data[field] and len(str(data[field])) > limit:
                issues.append(f"文字数制限超過 ({field}): {len(str(data[field]))} > {limit}")
        
        return {
            'is_valid': len(issues) == 0,
            'issues': issues,
            'data': data
        }


# グローバルインスタンス
security_validator = SecurityValidator()


def validate_input(value: Any) -> str:
    """入力値の検証（関数形式）"""
    return security_validator.validate_input(value)


def validate_file_path(filepath: str, allowed_dirs: List[str] = None) -> str:
    """ファイルパスの検証（関数形式）"""
    return security_validator.validate_file_path(filepath, allowed_dirs)


def validate_url(url: str) -> bool:
    """URLの検証（関数形式）"""
    return security_validator.validate_url(url)
