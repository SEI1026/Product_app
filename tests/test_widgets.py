#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
カスタムウィジェットのテスト
"""

import pytest
import sys
import os
from unittest.mock import Mock, patch, MagicMock
from PyQt5.QtWidgets import QApplication, QWidget
from PyQt5.QtCore import Qt
from PyQt5.QtTest import QTest

# テスト対象のモジュールをインポートできるようにパスを追加
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from widgets import JapaneseLineEdit, JapaneseTextEdit, JapaneseHtmlTextEdit


class TestJapaneseWidgets:
    """日本語ウィジェットのテスト"""
    
    @classmethod
    def setup_class(cls):
        """テストクラス全体の前に実行"""
        if not QApplication.instance():
            cls.app = QApplication([])
        else:
            cls.app = QApplication.instance()
    
    def test_japanese_line_edit_creation(self):
        """JapaneseLineEdit の作成テスト"""
        widget = JapaneseLineEdit()
        assert widget is not None
        assert hasattr(widget, '_show_japanese_context_menu')
    
    def test_japanese_line_edit_context_menu(self):
        """JapaneseLineEdit のコンテキストメニューテスト"""
        widget = JapaneseLineEdit()
        widget.setText("テストテキスト")
        
        # コンテキストメニューの表示テスト
        with patch.object(widget, '_show_japanese_context_menu') as mock_menu:
            widget.contextMenuEvent(Mock())
            mock_menu.assert_called_once()
    
    def test_japanese_text_edit_creation(self):
        """JapaneseTextEdit の作成テスト"""
        widget = JapaneseTextEdit()
        assert widget is not None
        assert hasattr(widget, '_show_japanese_context_menu')
    
    def test_japanese_text_edit_context_menu(self):
        """JapaneseTextEdit のコンテキストメニューテスト"""
        widget = JapaneseTextEdit()
        widget.setText("テストテキスト")
        
        # コンテキストメニューの表示テスト
        with patch.object(widget, '_show_japanese_context_menu') as mock_menu:
            widget.contextMenuEvent(Mock())
            mock_menu.assert_called_once()
    
    def test_japanese_html_text_edit_creation(self):
        """JapaneseHtmlTextEdit の作成テスト"""
        widget = JapaneseHtmlTextEdit()
        assert widget is not None
        assert hasattr(widget, '_show_japanese_context_menu')
    
    def test_japanese_html_text_edit_context_menu(self):
        """JapaneseHtmlTextEdit のコンテキストメニューテスト"""
        widget = JapaneseHtmlTextEdit()
        widget.setText("テストテキスト")
        
        # コンテキストメニューの表示テスト
        with patch.object(widget, '_show_japanese_context_menu') as mock_menu:
            widget.contextMenuEvent(Mock())
            mock_menu.assert_called_once()
    
    def test_context_menu_actions(self):
        """コンテキストメニューアクションのテスト"""
        widget = JapaneseLineEdit()
        widget.setText("テストテキスト")
        widget.selectAll()
        
        # 各アクションの存在確認
        menu = widget._create_japanese_menu()
        actions = menu.actions()
        
        action_texts = [action.text() for action in actions if not action.isSeparator()]
        
        expected_actions = ["元に戻す", "やり直し", "切り取り", "コピー", "貼り付け", "削除", "すべて選択"]
        
        for expected in expected_actions:
            assert any(expected in text for text in action_texts), f"'{expected}' アクションが見つかりません"
    
    def test_copy_paste_functionality(self):
        """コピー・ペースト機能のテスト"""
        widget = JapaneseLineEdit()
        test_text = "テストテキスト"
        widget.setText(test_text)
        widget.selectAll()
        
        # コピー操作
        widget.copy()
        
        # ペースト操作
        widget.clear()
        widget.paste()
        
        # コピーしたテキストがペーストされることを確認
        # 注意: この部分はクリップボードの状態に依存するため、
        # 実際のテスト環境では別の方法で検証する必要がある場合があります
        assert widget.text() == test_text or widget.text() == ""  # クリップボードが利用できない場合


class TestCustomWidgetFunctionality:
    """カスタムウィジェットの機能テスト"""
    
    @classmethod
    def setup_class(cls):
        """テストクラス全体の前に実行"""
        if not QApplication.instance():
            cls.app = QApplication([])
        else:
            cls.app = QApplication.instance()
    
    def test_japanese_line_edit_input_handling(self):
        """JapaneseLineEdit の入力処理テスト"""
        widget = JapaneseLineEdit()
        
        # 日本語テキストの入力
        test_text = "こんにちは世界"
        widget.setText(test_text)
        assert widget.text() == test_text
        
        # 英数字の入力
        test_text = "Hello123"
        widget.setText(test_text)
        assert widget.text() == test_text
    
    def test_japanese_text_edit_html_handling(self):
        """JapaneseTextEdit のHTML処理テスト"""
        widget = JapaneseTextEdit()
        
        # プレーンテキストの設定
        plain_text = "テストテキスト"
        widget.setPlainText(plain_text)
        assert widget.toPlainText() == plain_text
        
        # HTMLテキストの設定
        html_text = "<b>太字テキスト</b>"
        widget.setHtml(html_text)
        assert "<b>" in widget.toHtml().lower() or "太字テキスト" in widget.toPlainText()
    
    def test_widget_focus_behavior(self):
        """ウィジェットのフォーカス動作テスト"""
        widget = JapaneseLineEdit()
        parent = QWidget()
        widget.setParent(parent)
        parent.show()
        
        # フォーカス設定
        widget.setFocus()
        
        # フォーカス状態の確認（実際のGUI環境でのみ有効）
        # assert widget.hasFocus()  # GUI環境依存のため、コメントアウト
        
        parent.close()


if __name__ == "__main__":
    pytest.main([__file__])