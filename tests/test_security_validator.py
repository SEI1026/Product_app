#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
セキュリティバリデーターのテスト
"""

import pytest
import tempfile
import os
import sys
from unittest.mock import patch, Mock

# テスト対象のモジュールをインポートできるようにパスを追加
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'src', 'utils'))

from security_validator import SecurityValidator


class TestSecurityValidator:
    """SecurityValidator クラスのテスト"""
    
    def setup_method(self):
        """各テストメソッドの前に実行"""
        self.validator = SecurityValidator()
    
    def test_validate_input_basic(self):
        """基本的な入力検証"""
        # 通常の入力
        assert self.validator.validate_input("テスト") == "テスト"
        assert self.validator.validate_input("123") == "123"
        assert self.validator.validate_input("") == ""
        assert self.validator.validate_input(None) == ""
    
    def test_validate_input_html_escape(self):
        """HTMLエスケープのテスト"""
        # HTMLタグの検証
        result = self.validator.validate_input("<div>テスト</div>")
        assert "&lt;div&gt;テスト&lt;/div&gt;" == result
        
        # 特殊文字のエスケープ
        result = self.validator.validate_input("&<>\"'")
        assert "&amp;&lt;&gt;&quot;&#x27;" == result
    
    def test_validate_input_xss_protection(self):
        """XSS攻撃の検証"""
        # スクリプトタグ
        result = self.validator.validate_input("<script>alert('xss')</script>")
        assert "<script>" not in result.lower()
        
        # javascriptスキーム
        result = self.validator.validate_input("javascript:alert('xss')")
        assert "javascript:" not in result.lower()
        
        # データURL
        result = self.validator.validate_input("data:text/html,<script>alert('xss')</script>")
        assert "data:" not in result.lower()
        
        # SVG XSS
        result = self.validator.validate_input("<svg onload=alert('xss')></svg>")
        assert "<svg" not in result.lower()
    
    def test_validate_input_length_limit(self):
        """入力長制限のテスト"""
        # 長い入力
        long_input = "A" * 1500
        result = self.validator.validate_input(long_input)
        assert len(result) <= self.validator.max_input_length
    
    def test_validate_csv_input_basic(self):
        """CSV入力検証の基本テスト"""
        # 通常の入力
        assert self.validator.validate_csv_input("テスト") == "テスト"
        assert self.validator.validate_csv_input("123") == "123"
        assert self.validator.validate_csv_input("") == ""
        assert self.validator.validate_csv_input(None) == ""
    
    def test_validate_csv_input_formula_protection(self):
        """CSVフォーミュラ攻撃の検証"""
        # 等号で始まる入力
        result = self.validator.validate_csv_input("=SUM(A1:A10)")
        assert result.startswith("'=")
        
        # プラス記号で始まる入力
        result = self.validator.validate_csv_input("+1+1")
        assert result.startswith("'+")
        
        # マイナス記号で始まる入力
        result = self.validator.validate_csv_input("-1")
        assert result.startswith("'-")
        
        # アットマークで始まる入力
        result = self.validator.validate_csv_input("@SUM(1,2)")
        assert result.startswith("'@")
    
    def test_validate_file_path_basic(self):
        """ファイルパス検証の基本テスト"""
        # 正常なパス
        with tempfile.NamedTemporaryFile() as tmp:
            result = self.validator.validate_file_path(tmp.name)
            assert os.path.isabs(result)
    
    def test_validate_file_path_traversal_protection(self):
        """ディレクトリトラバーサル攻撃の検証"""
        # ディレクトリトラバーサル攻撃
        with pytest.raises(ValueError, match="ディレクトリトラバーサル"):
            self.validator.validate_file_path("../../../etc/passwd")
        
        with pytest.raises(ValueError, match="ディレクトリトラバーサル"):
            self.validator.validate_file_path("test/../../../etc/passwd")
    
    def test_validate_file_path_allowed_dirs(self):
        """許可ディレクトリの検証"""
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_file = os.path.join(temp_dir, "test.txt")
            with open(temp_file, 'w') as f:
                f.write("test")
            
            # 許可されたディレクトリ内
            result = self.validator.validate_file_path(temp_file, [temp_dir])
            assert result == os.path.abspath(temp_file)
            
            # 許可されていないディレクトリ
            with pytest.raises(ValueError, match="許可されていないディレクトリ"):
                self.validator.validate_file_path(temp_file, ["/other/dir"])
    
    def test_validate_file_type_extension(self):
        """ファイル拡張子の検証"""
        with tempfile.NamedTemporaryFile(suffix=".csv") as tmp:
            # 許可された拡張子
            assert self.validator.validate_file_type(tmp.name, [".csv", ".txt"])
            
            # 許可されていない拡張子
            assert not self.validator.validate_file_type(tmp.name, [".jpg", ".png"])
    
    @patch('magic.from_file')
    def test_validate_file_type_mime(self, mock_magic):
        """MIMEタイプの検証"""
        mock_magic.return_value = "text/csv"
        
        with tempfile.NamedTemporaryFile(suffix=".csv") as tmp:
            # 許可されたMIMEタイプ
            assert self.validator.validate_file_type(
                tmp.name, 
                allowed_mime_types=["text/csv", "text/plain"]
            )
            
            # 許可されていないMIMEタイプ
            assert not self.validator.validate_file_type(
                tmp.name,
                allowed_mime_types=["image/jpeg", "image/png"]
            )
    
    def test_validate_url_basic(self):
        """URL検証の基本テスト"""
        # 正常なURL
        assert self.validator.validate_url("https://example.com")
        assert self.validator.validate_url("http://example.com")
        
        # 無効なURL
        assert not self.validator.validate_url("")
        assert not self.validator.validate_url("ftp://example.com")
        assert not self.validator.validate_url("file:///etc/passwd")
    
    def test_validate_url_private_ip_protection(self):
        """プライベートIP保護の検証"""
        # プライベートIP
        assert not self.validator.validate_url("http://192.168.1.1")
        assert not self.validator.validate_url("http://10.0.0.1")
        assert not self.validator.validate_url("http://172.16.0.1")
        
        # ローカルホスト
        assert not self.validator.validate_url("http://localhost")
        assert not self.validator.validate_url("http://127.0.0.1")
    
    def test_validate_numeric_input(self):
        """数値入力の検証"""
        # 正常な数値
        assert self.validator.validate_numeric_input("123") == 123
        assert self.validator.validate_numeric_input("123.45") == 123.45
        assert self.validator.validate_numeric_input(456) == 456
        
        # 無効な入力
        assert self.validator.validate_numeric_input("") is None
        assert self.validator.validate_numeric_input(None) is None
        assert self.validator.validate_numeric_input("abc") is None
        
        # 範囲チェック
        assert self.validator.validate_numeric_input("50", min_val=0, max_val=100) == 50
        assert self.validator.validate_numeric_input("-10", min_val=0, max_val=100) is None
        assert self.validator.validate_numeric_input("150", min_val=0, max_val=100) is None


if __name__ == "__main__":
    pytest.main([__file__])