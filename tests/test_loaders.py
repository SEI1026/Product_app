# -*- coding: utf-8 -*-
"""
loaders.py モジュールのテスト

主要な関数のテストケースを実装:
- load_categories_from_csv
- load_explanation_mark_icons
- load_material_spec_master
"""
import pytest
import sys
import os
import tempfile
import csv
from unittest.mock import Mock, patch

# テスト対象のモジュールをインポートできるようにパスを追加
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from loaders import load_categories_from_csv, load_explanation_mark_icons, load_material_spec_master


class TestLoadCategoriesFromCsv:
    """load_categories_from_csv 関数のテスト"""
    
    def test_load_valid_csv(self):
        """正常なCSVファイルの読み込み"""
        # テスト用のCSVデータを作成
        csv_data = [
            ["レベル", "カテゴリ名", "親カテゴリ名"],
            ["1", "カテゴリA", "親A"],
            ["2", "カテゴリB", "親B"],
            ["3", "カテゴリC", "親C"]
        ]
        
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerows(csv_data)
            temp_path = f.name
        
        try:
            result = load_categories_from_csv(temp_path)
            
            assert len(result) == 3
            # 結果は (level, name, parent) のタプルのリスト
            assert result[0] == (1, "カテゴリA", "親A")
            assert result[1] == (2, "カテゴリB", "親B")
            assert result[2] == (3, "カテゴリC", "親C")
            
        finally:
            os.unlink(temp_path)
    
    def test_load_nonexistent_file(self):
        """存在しないファイルの処理"""
        with pytest.raises(FileNotFoundError):
            load_categories_from_csv("nonexistent.csv")
    
    def test_load_empty_csv(self):
        """空のCSVファイルの処理"""
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv', encoding='utf-8-sig') as f:
            temp_path = f.name
        
        try:
            result = load_categories_from_csv(temp_path)
            assert result == []
        finally:
            os.unlink(temp_path)


class TestLoadExplanationMarkIcons:
    """load_explanation_mark_icons 関数のテスト"""
    
    @patch('os.path.isdir')
    @patch('os.listdir')
    def test_load_explanation_icons(self, mock_listdir, mock_isdir):
        """説明マークアイコンの読み込み"""
        # モックの設定
        mock_isdir.return_value = True
        mock_listdir.return_value = [
            "1_中国製.jpg",
            "2_タイ製.jpg", 
            "3_ベトナム製.jpg",
            "readme.txt"  # 画像ファイル以外
        ]
        
        result = load_explanation_mark_icons("/fake/path")
        
        # 結果は辞書のリストで、必要なキーが含まれている
        assert len(result) == 3
        # 各要素が辞書であり、必要なキーを持つことを確認
        for icon in result:
            assert isinstance(icon, dict)
            assert "id" in icon
            assert "description" in icon
            assert "path" in icon
            assert "filename" in icon
        
        # 特定のアイコンの内容を確認
        icon_1 = next(icon for icon in result if icon["id"] == "1")
        assert icon_1["description"] == "中国製"
        assert icon_1["filename"] == "1_中国製.jpg"
    
    @patch('os.path.isdir')
    def test_load_nonexistent_directory(self, mock_isdir):
        """存在しないディレクトリの処理"""
        mock_isdir.return_value = False
        
        result = load_explanation_mark_icons("/nonexistent/path")
        assert result == []


class TestLoadMaterialSpecMaster:
    """load_material_spec_master 関数のテスト"""
    
    def test_load_valid_material_spec(self):
        """正常な材質・仕様マスターの読み込み"""
        # テスト用のCSVデータを作成
        csv_data = [
            ["名称", "説明"],
            ["無垢材", "天然木を使用した材質"],
            ["合板", "複数の板を接着した材質"],
            ["MDF", "中密度繊維板"]
        ]
        
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerows(csv_data)
            temp_path = f.name
        
        try:
            result = load_material_spec_master(temp_path)
            
            assert len(result) == 3
            assert result["無垢材"] == "天然木を使用した材質"
            assert result["合板"] == "複数の板を接着した材質"
            assert result["MDF"] == "中密度繊維板"
            
        finally:
            os.unlink(temp_path)
    
    def test_load_invalid_material_spec(self):
        """不正な材質・仕様マスターの処理"""
        # カラムが不足しているCSVデータ
        csv_data = [
            ["名称"],  # 説明カラムがない
            ["無垢材"],
            ["合板"]
        ]
        
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerows(csv_data)
            temp_path = f.name
        
        try:
            result = load_material_spec_master(temp_path)
            # エラーの場合は空の辞書が返される
            assert result == {}
            
        finally:
            os.unlink(temp_path)


# 統合テスト
class TestLoadersIntegration:
    """ローダー機能の統合テスト"""
    
    def test_all_loaders_handle_encoding(self):
        """全てのローダーがエンコーディングを適切に処理"""
        # 日本語を含むテストデータ
        japanese_data = [
            ["1", "テーブル", "家具"],
            ["2", "椅子", "家具"]
        ]
        
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerows(japanese_data)
            temp_path = f.name
        
        try:
            result = load_categories_from_csv(temp_path)
            
            assert len(result) == 2
            # 結果は (level, name, parent) のタプル
            assert result[0][1] == "テーブル"  # name部分
            assert result[1][1] == "椅子"     # name部分
            
        finally:
            os.unlink(temp_path)


if __name__ == "__main__":
    pytest.main([__file__])