# -*- coding: utf-8 -*-
"""
utils.py モジュールのテスト

主要な関数のテストケースを実装:
- normalize_text
- normalize_wave_dash
- get_byte_count_excel_lenb
"""
import pytest
import sys
import os

# テスト対象のモジュールをインポートできるようにパスを追加
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils import normalize_text, normalize_wave_dash, get_byte_count_excel_lenb


class TestNormalizeText:
    """normalize_text 関数のテスト"""
    
    def test_normalize_basic_text(self):
        """基本的なテキスト正規化"""
        # 半角カナを全角に変換
        assert normalize_text("ｱｲｳｴｵ") == "アイウエオ"
        
        # 全角英数字を半角に変換
        assert normalize_text("１２３ＡＢＣ") == "123ABC"
        
        # 全角スペースを半角に変換
        assert normalize_text("テスト　データ") == "テスト データ"
    
    def test_normalize_mixed_text(self):
        """混在したテキストの正規化"""
        input_text = "商品名：テーブル１２３ＣＭ　高サ７５ＣＭ"
        expected = "商品名:テーブル123CM 高サ75CM"
        assert normalize_text(input_text) == expected
    
    def test_normalize_empty_text(self):
        """空文字列の処理"""
        assert normalize_text("") == ""
        assert normalize_text(None) == ""
    
    def test_normalize_already_normalized(self):
        """既に正規化済みのテキスト"""
        text = "テストデータ 123 ABC"
        assert normalize_text(text) == text


class TestNormalizeWaveDash:
    """normalize_wave_dash 関数のテスト"""
    
    def test_wave_dash_conversion(self):
        """波ダッシュの変換"""
        # 全角チルダを波ダッシュに変換
        assert normalize_wave_dash("テスト〜データ") == "テスト～データ"
        
        # 複数の波ダッシュ
        assert normalize_wave_dash("範囲〜〜〜終了") == "範囲～～～終了"
    
    def test_no_wave_dash(self):
        """波ダッシュがないテキスト"""
        text = "普通のテキストです"
        assert normalize_wave_dash(text) == text
    
    def test_empty_wave_dash(self):
        """空文字列の処理"""
        assert normalize_wave_dash("") == ""
        assert normalize_wave_dash(None) == ""


class TestGetByteCountExcelLenb:
    """get_byte_count_excel_lenb 関数のテスト"""
    
    def test_ascii_characters(self):
        """ASCII文字のバイト数"""
        assert get_byte_count_excel_lenb("ABC") == 3
        assert get_byte_count_excel_lenb("123") == 3
    
    def test_japanese_characters(self):
        """日本語文字のバイト数"""
        # ひらがな、カタカナ、漢字は2バイト
        assert get_byte_count_excel_lenb("あいう") == 6
        assert get_byte_count_excel_lenb("アイウ") == 6
        assert get_byte_count_excel_lenb("商品名") == 6
    
    def test_mixed_characters(self):
        """混在した文字のバイト数"""
        # "商品A123" = 商品(4) + A(1) + 123(3) = 8バイト
        assert get_byte_count_excel_lenb("商品A123") == 8
    
    def test_empty_string(self):
        """空文字列のバイト数"""
        assert get_byte_count_excel_lenb("") == 0
        assert get_byte_count_excel_lenb(None) == 0
    
    def test_special_characters(self):
        """特殊文字のバイト数"""
        # 記号は1バイト
        assert get_byte_count_excel_lenb("!@#$%") == 5
        
        # 全角記号は2バイト
        assert get_byte_count_excel_lenb("！＠＃") == 6


if __name__ == "__main__":
    pytest.main([__file__])
