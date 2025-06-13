# -*- coding: utf-8 -*-
"""
SearchPanel クラスのテスト

検索・置換機能のテストケースを実装
"""
import pytest
import sys
import os
from unittest.mock import Mock, patch, MagicMock

# PyQt5のテスト用インポート
try:
    from PyQt5.QtWidgets import QApplication, QWidget, QLineEdit, QTextEdit
    from PyQt5.QtCore import Qt
    from PyQt5.QtTest import QTest
    import pytest_qt
    PYQT_AVAILABLE = True
except ImportError:
    PYQT_AVAILABLE = False

# テスト対象のモジュールをインポートできるようにパスを追加
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

if PYQT_AVAILABLE:
    from product_app import SearchPanel

pytestmark = pytest.mark.skipif(not PYQT_AVAILABLE, reason="PyQt5 not available")


class TestSearchPanel:
    """SearchPanel クラスのテスト"""
    
    @pytest.fixture
    def app(self, qapp):
        """テスト用のQApplicationインスタンス"""
        return qapp
    
    @pytest.fixture
    def mock_parent(self):
        """SearchPanelの親アプリケーションのモック"""
        parent = Mock()
        parent.main_fields = {
            'field1': QLineEdit(),
            'field2': QTextEdit(),
            'field3': QLineEdit()
        }
        parent.main_fields['field1'].setText("テストデータ1")
        parent.main_fields['field2'].setPlainText("テストデータ2 検索対象")
        parent.main_fields['field3'].setText("別のデータ")
        return parent
    
    @pytest.fixture
    def search_panel(self, qtbot, mock_parent):
        """テスト用のSearchPanelインスタンス"""
        panel = SearchPanel(parent=mock_parent)
        qtbot.addWidget(panel)
        return panel
    
    def test_search_panel_initialization(self, search_panel):
        """SearchPanelの初期化テスト"""
        assert search_panel.windowTitle() == "検索と置換"
        assert search_panel.width() == 350
        assert search_panel.current_results == []
        assert search_panel.current_index == -1
    
    def test_auto_search_checkbox(self, search_panel):
        """自動検索チェックボックスのテスト"""
        # デフォルトで有効になっている
        assert search_panel.auto_search.isChecked() == True
        
        # チェックボックスの状態変更
        search_panel.auto_search.setChecked(False)
        assert search_panel.auto_search.isChecked() == False
    
    def test_search_text_input(self, qtbot, search_panel):
        """検索テキスト入力のテスト"""
        # 自動検索が有効な場合
        search_panel.auto_search.setChecked(True)
        
        # テキスト入力をシミュレート
        qtbot.keyClicks(search_panel.search_input, "テスト")
        
        # ボタンが有効になることを確認
        assert search_panel.find_next_btn.isEnabled() == True
        assert search_panel.find_prev_btn.isEnabled() == True
        assert search_panel.find_all_btn.isEnabled() == True
    
    def test_auto_search_disabled(self, qtbot, search_panel):
        """自動検索無効時のテスト"""
        # 自動検索を無効に設定
        search_panel.auto_search.setChecked(False)
        
        # 検索メソッドをモック
        search_panel.perform_search = Mock()
        
        # テキスト入力をシミュレート
        qtbot.keyClicks(search_panel.search_input, "テスト")
        
        # 自動検索が実行されないことを確認
        search_panel.perform_search.assert_not_called()
    
    def test_manual_search_on_enter(self, qtbot, search_panel):
        """Enterキーでの手動検索テスト"""
        # 自動検索を無効に設定
        search_panel.auto_search.setChecked(False)
        
        # 検索メソッドをモック
        search_panel.perform_search = Mock()
        
        # テキスト入力
        search_panel.search_input.setText("テスト")
        
        # Enterキーを押下
        qtbot.keyPress(search_panel.search_input, Qt.Key_Return)
        
        # 手動検索が実行されることを確認
        search_panel.perform_search.assert_called_once_with(auto_jump=True)
    
    def test_search_options(self, search_panel):
        """検索オプションのテスト"""
        # 大文字小文字を区別する
        search_panel.case_sensitive.setChecked(True)
        assert search_panel.case_sensitive.isChecked() == True
    
    def test_search_scope_selection(self, search_panel):
        """検索対象の選択テスト"""
        # 検索対象のコンボボックス
        assert search_panel.scope_combo.count() == 3
        
        # デフォルトは「商品一覧」
        assert search_panel.scope_combo.currentIndex() == 0
        
        # 「現在の商品のフィールド」に変更
        search_panel.scope_combo.setCurrentIndex(1)
        assert search_panel.scope_combo.currentIndex() == 1
    
    def test_replace_input(self, qtbot, search_panel):
        """置換入力のテスト"""
        # 置換テキスト入力
        qtbot.keyClicks(search_panel.replace_input, "置換後")
        assert search_panel.replace_input.text() == "置換後"
    
    def test_esc_key_handling(self, qtbot, search_panel):
        """ESCキーでのパネル閉じるテスト"""
        # パネルを表示
        search_panel.show()
        assert search_panel.isVisible() == True
        
        # close_panel メソッドをモック
        search_panel.close_panel = Mock()
        
        # ESCキーを押下
        qtbot.keyPress(search_panel, Qt.Key_Escape)
        
        # パネルが閉じられることを確認
        search_panel.close_panel.assert_called_once()


class TestSearchFunctionality:
    """検索機能のテスト"""
    
    @pytest.fixture
    def search_panel_with_data(self, qtbot, mock_parent):
        """テストデータ付きのSearchPanel"""
        # より詳細なテストデータを設定
        mock_parent.main_fields['商品名'] = QLineEdit()
        mock_parent.main_fields['商品名'].setText("テーブル 木製 120cm")
        
        mock_parent.main_fields['説明'] = QTextEdit()
        mock_parent.main_fields['説明'].setPlainText("美しい木製テーブルです。サイズは120cmです。")
        
        panel = SearchPanel(parent=mock_parent)
        qtbot.addWidget(panel)
        return panel
    
    def test_search_current_product_fields(self, search_panel_with_data):
        """現在の商品フィールド検索のテスト"""
        panel = search_panel_with_data
        
        # 検索対象を「現在の商品のフィールド」に設定
        panel.scope_combo.setCurrentIndex(1)
        
        # 検索実行をモック
        with patch.object(panel, 'search_current_product') as mock_search:
            panel.search_input.setText("テーブル")
            panel.perform_search()
            
            # search_current_product が呼ばれることを確認
            mock_search.assert_called_once()
    
    def test_field_monitoring_setup(self, search_panel_with_data):
        """フィールド監視の設定テスト"""
        panel = search_panel_with_data
        
        # フィールド監視が設定されることを確認
        # (実際の実装では textChanged シグナルが接続される)
        assert hasattr(panel, '_update_timer')


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
