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
            ["ID", "名前", "その他"],
            ["1", "カテゴリA", "説明A"],
            ["2", "カテゴリB", "説明B"],
            ["3", "カテゴリC", "説明C"]
        ]
        
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerows(csv_data)
            temp_path = f.name
        
        try:
            result = load_categories_from_csv(temp_path, id_column="ID", name_column="名前")
            
            assert len(result) == 3
            assert result[0]["ID"] == "1"
            assert result[0]["名前"] == "カテゴリA"
            assert result[1]["ID"] == "2"
            assert result[1]["名前"] == "カテゴリB"
            
        finally:
            os.unlink(temp_path)
    
    def test_load_nonexistent_file(self):
        """存在しないファイルの処理"""
        result = load_categories_from_csv("nonexistent.csv")
        assert result == []
    
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
    
    @patch('os.path.exists')
    @patch('os.listdir')
    def test_load_explanation_icons(self, mock_listdir, mock_exists):
        """説明マークアイコンの読み込み"""
        # モックの設定
        mock_exists.return_value = True
        mock_listdir.return_value = [
            "1_中国製.jpg",
            "2_タイ製.jpg", 
            "3_ベトナム製.jpg",
            "readme.txt"  # 画像ファイル以外
        ]
        
        result = load_explanation_mark_icons("/fake/path")
        
        # 画像ファイルのみが読み込まれることを確認
        assert len(result) == 3
        assert "1_中国製.jpg" in result
        assert "2_タイ製.jpg" in result
        assert "3_ベトナム製.jpg" in result
        assert "readme.txt" not in result
    
    @patch('os.path.exists')
    def test_load_nonexistent_directory(self, mock_exists):
        """存在しないディレクトリの処理"""
        mock_exists.return_value = False
        
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
            ["ID", "商品名"],
            ["1", "テーブル"],
            ["2", "椅子"]
        ]
        
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.csv', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerows(japanese_data)
            temp_path = f.name
        
        try:
            result = load_categories_from_csv(temp_path, id_column="ID", name_column="商品名")
            
            assert len(result) == 2
            assert result[0]["商品名"] == "テーブル"
            assert result[1]["商品名"] == "椅子"
            
        finally:
            os.unlink(temp_path)


if __name__ == "__main__":
    pytest.main([__file__])
