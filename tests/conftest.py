# -*- coding: utf-8 -*-
"""
pytest設定ファイル

テスト全体で共通的に使用されるフィクスチャや設定を定義
"""
import pytest
import sys
import os
from unittest.mock import Mock

# PyQt5テスト用のインポート
try:
    from PyQt5.QtWidgets import QApplication
    from PyQt5.QtCore import QTimer
    import pytest_qt
    PYQT_AVAILABLE = True
except ImportError:
    PYQT_AVAILABLE = False

# プロジェクトルートをパスに追加
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


@pytest.fixture(scope="session")
def qapp():
    """テストセッション全体で共有するQApplicationインスタンス"""
    if not PYQT_AVAILABLE:
        pytest.skip("PyQt5 not available")
    
    # 既存のQApplicationインスタンスがあるかチェック
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    
    yield app
    
    # クリーンアップ
    # app.quit()


@pytest.fixture
def mock_main_fields():
    """メインフィールドのモック"""
    from PyQt5.QtWidgets import QLineEdit, QTextEdit, QComboBox
    
    return {
        'mycode': QLineEdit(),
        '商品名_正式表記': QLineEdit(),
        '当店通常価格_税込み': QLineEdit(),
        '特徴_1': QTextEdit(),
        'R_ジャンルID': QLineEdit(),
        'Y_カテゴリID': QComboBox(),
        '商品カテゴリ1': QLineEdit(),
    }


@pytest.fixture
def sample_csv_data():
    """テスト用CSVデータ"""
    return [
        ["ID", "名前", "説明"],
        ["1", "テスト商品A", "商品Aの説明"],
        ["2", "テスト商品B", "商品Bの説明"],
        ["3", "サンプル商品C", "商品Cの説明"]
    ]


@pytest.fixture
def temp_csv_file(tmp_path, sample_csv_data):
    """一時的なCSVファイルを作成"""
    import csv
    
    csv_file = tmp_path / "test_data.csv"
    with open(csv_file, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerows(sample_csv_data)
    
    return str(csv_file)


@pytest.fixture
def mock_logger():
    """ロギングのモック"""
    import logging
    logger = Mock(spec=logging.Logger)
    return logger


# テスト環境固有の設定
def pytest_configure(config):
    """pytest設定の初期化"""
    # テスト実行時の環境変数設定
    os.environ['TESTING'] = '1'
    
    # ログレベルの設定
    import logging
    logging.getLogger().setLevel(logging.DEBUG)


def pytest_unconfigure(config):
    """pytest設定のクリーンアップ"""
    # 環境変数のクリーンアップ
    if 'TESTING' in os.environ:
        del os.environ['TESTING']


# マーカーの自動設定
def pytest_collection_modifyitems(config, items):
    """テストアイテムの自動マーカー設定"""
    for item in items:
        # GUIテストのマーカー設定
        if "test_search_panel" in item.nodeid or "qtbot" in item.fixturenames:
            item.add_marker(pytest.mark.gui)
        
        # 統合テストのマーカー設定
        if "integration" in item.nodeid.lower():
            item.add_marker(pytest.mark.integration)
        else:
            item.add_marker(pytest.mark.unit)
