#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
データモデルのテスト
"""

import pytest
import sys
import os
from unittest.mock import Mock, patch
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import Qt, QModelIndex

# テスト対象のモジュールをインポートできるようにパスを追加
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models import SkuTableModel


class TestSkuTableModel:
    """SkuTableModel のテスト"""
    
    @classmethod
    def setup_class(cls):
        """テストクラス全体の前に実行"""
        if not QApplication.instance():
            cls.app = QApplication([])
        else:
            cls.app = QApplication.instance()
    
    def setup_method(self):
        """各テストメソッドの前に実行"""
        self.headers = ["SKU", "商品名", "価格", "在庫"]
        self.model = SkuTableModel(self.headers)
    
    def test_model_creation(self):
        """モデル作成のテスト"""
        assert self.model is not None
        assert self.model.columnCount() == len(self.headers)
        assert self.model.rowCount() == 0
    
    def test_header_data(self):
        """ヘッダーデータのテスト"""
        for i, header in enumerate(self.headers):
            header_data = self.model.headerData(i, Qt.Horizontal, Qt.DisplayRole)
            assert header_data == header
    
    def test_add_row(self):
        """行追加のテスト"""
        initial_count = self.model.rowCount()
        test_data = ["SKU001", "テスト商品", "1000", "10"]
        
        self.model.add_row(test_data)
        
        assert self.model.rowCount() == initial_count + 1
        
        # データの確認
        for i, value in enumerate(test_data):
            index = self.model.index(0, i)
            assert self.model.data(index, Qt.DisplayRole) == value
    
    def test_set_data(self):
        """データ設定のテスト"""
        # まず行を追加
        test_data = ["SKU001", "テスト商品", "1000", "10"]
        self.model.add_row(test_data)
        
        # データを変更
        new_value = "更新された商品名"
        index = self.model.index(0, 1)  # 商品名の列
        
        result = self.model.setData(index, new_value, Qt.EditRole)
        assert result is True
        
        # 変更されたデータの確認
        updated_data = self.model.data(index, Qt.DisplayRole)
        assert updated_data == new_value
    
    def test_remove_rows(self):
        """行削除のテスト"""
        # テストデータを追加
        test_data = [
            ["SKU001", "商品1", "1000", "10"],
            ["SKU002", "商品2", "2000", "20"],
            ["SKU003", "商品3", "3000", "30"]
        ]
        
        for data in test_data:
            self.model.add_row(data)
        
        initial_count = self.model.rowCount()
        
        # 中間の行を削除
        result = self.model.removeRows(1, 1)
        assert result is True
        assert self.model.rowCount() == initial_count - 1
        
        # 残ったデータの確認
        first_row_data = self.model.data(self.model.index(0, 0), Qt.DisplayRole)
        second_row_data = self.model.data(self.model.index(1, 0), Qt.DisplayRole)
        
        assert first_row_data == "SKU001"
        assert second_row_data == "SKU003"
    
    def test_flags(self):
        """アイテムフラグのテスト"""
        # 行を追加
        test_data = ["SKU001", "テスト商品", "1000", "10"]
        self.model.add_row(test_data)
        
        index = self.model.index(0, 0)
        flags = self.model.flags(index)
        
        # 編集可能であることを確認
        assert flags & Qt.ItemIsEditable
        assert flags & Qt.ItemIsEnabled
        assert flags & Qt.ItemIsSelectable
    
    def test_get_row_data(self):
        """行データ取得のテスト"""
        test_data = ["SKU001", "テスト商品", "1000", "10"]
        self.model.add_row(test_data)
        
        row_data = self.model.get_row_data(0)
        assert row_data == test_data
    
    def test_clear_data(self):
        """データクリアのテスト"""
        # テストデータを追加
        test_data = ["SKU001", "テスト商品", "1000", "10"]
        self.model.add_row(test_data)
        
        assert self.model.rowCount() > 0
        
        self.model.clear_data()
        
        assert self.model.rowCount() == 0
    
    def test_invalid_index_handling(self):
        """無効なインデックスの処理テスト"""
        # 無効なインデックス
        invalid_index = self.model.index(-1, 0)
        assert not invalid_index.isValid()
        
        # 範囲外のインデックス
        out_of_range_index = self.model.index(100, 0)
        data = self.model.data(out_of_range_index, Qt.DisplayRole)
        assert data is None or data == ""
    
    def test_data_validation(self):
        """データ検証のテスト"""
        test_data = ["SKU001", "テスト商品", "1000", "10"]
        self.model.add_row(test_data)
        
        index = self.model.index(0, 2)  # 価格の列
        
        # 数値データの設定
        assert self.model.setData(index, "2000", Qt.EditRole)
        assert self.model.data(index, Qt.DisplayRole) == "2000"
        
        # 無効なデータの処理（実装に依存）
        # 注意: 実際の実装によってはバリデーションが含まれる場合があります
    
    def test_bulk_operations(self):
        """一括操作のテスト"""
        # 複数行の追加
        test_data_list = [
            ["SKU001", "商品1", "1000", "10"],
            ["SKU002", "商品2", "2000", "20"],
            ["SKU003", "商品3", "3000", "30"],
            ["SKU004", "商品4", "4000", "40"],
            ["SKU005", "商品5", "5000", "50"]
        ]
        
        for data in test_data_list:
            self.model.add_row(data)
        
        assert self.model.rowCount() == len(test_data_list)
        
        # 複数行の削除
        result = self.model.removeRows(1, 3)  # 1行目から3行削除
        assert result is True
        assert self.model.rowCount() == len(test_data_list) - 3


if __name__ == "__main__":
    pytest.main([__file__])