"""
商品登録入力ツール - ユーティリティ関数モジュール
"""
import os
import sys
import csv
import functools
import unicodedata
import logging
from contextlib import contextmanager
from typing import Optional, List, Dict, Union, Tuple
from PyQt5.QtCore import QStandardPaths
from PyQt5.QtWidgets import QApplication

from constants import (
    DEFAULT_ENCODING, FALLBACK_ENCODING, PROGRESS_UPDATE_ROW_INTERVAL,
    APP_DATA_SUBDIR
)


@contextmanager
def open_csv_file_with_fallback(filepath, mode='r', progress_dialog=None, file_label="CSVファイル"):
    """
    CSVファイルをUTF-8 (BOM付き)で開き、失敗したらShift_JISで再試行するコンテキストマネージャ。
    ファイルオブジェクト、デリミタ、使用されたエンコーディングをyieldする。
    ファイルが見つからない場合やデコードに失敗した場合は例外を発生させる。
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"{file_label} '{filepath}' が見つかりません。")

    encodings_to_try = [DEFAULT_ENCODING, FALLBACK_ENCODING]
    file_obj = None
    base_filename = os.path.basename(filepath)

    for encoding_idx, encoding in enumerate(encodings_to_try):
        if progress_dialog:
            label_suffix = f" ({encoding})" if encoding != DEFAULT_ENCODING else ""
            progress_dialog.setLabelText(f"{file_label} ({base_filename}{label_suffix}) を読み込み中...")
            QApplication.processEvents()
        try:
            file_obj = open(filepath, mode, encoding=encoding, newline='')
            delimiter = None
            if 'r' in mode:  # 読み込みモードの場合のみデリミタ検出
                first_line = file_obj.readline()
                file_obj.seek(0)
                delimiter = '\t' if '\t' in first_line and ',' not in first_line else ','
            
            try:
                yield file_obj, delimiter, encoding
            finally:
                if file_obj:
                    file_obj.close()
            return  # 成功したら終了
        except UnicodeDecodeError:
            if file_obj: 
                file_obj.close()
            if encoding == FALLBACK_ENCODING:  # 最後のフォールバックでも失敗
                raise  # 元のUnicodeDecodeErrorを再発生させる
            logging.info(f"{file_label} '{filepath}' を{encoding}で読み込み失敗。{FALLBACK_ENCODING}で再試行します。")
        except Exception as e:  # FileNotFoundErrorもここでキャッチされる可能性あり
            if file_obj: 
                file_obj.close()
            logging.warning(f"{file_label} '{filepath}' ({encoding}) のオープン/読み込み中に予期せぬエラー。", exc_info=True)
            if encoding == FALLBACK_ENCODING or isinstance(e, FileNotFoundError):
                raise  # 最後のフォールバックでのエラー、またはFileNotFoundErrorは再発生


@functools.lru_cache(maxsize=1000)
def normalize_text(text: Optional[Union[str, int, float]]) -> str:
    """全角英数字、記号、カタカナを半角に、ひらがなをカタカナに変換し、大文字にする"""
    if text is None: 
        return ""
    text_str = str(text)  # 数値なども文字列として扱えるように
    text_str = unicodedata.normalize('NFKC', text_str).upper()
    # ひらがなをカタカナに変換（効率的な変換）
    hiragana_to_katakana = str.maketrans(
        'あいうえおかきくけこさしすせそたちつてとなにぬねのはひふへほまみむめもやゆよらりるれろわをん'
        'ぁぃぅぇぉゃゅょっ',
        'アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン'
        'ァィゥェォャュョッ'
    )
    return text_str.translate(hiragana_to_katakana)


@functools.lru_cache(maxsize=500)
def normalize_wave_dash(text: Optional[Union[str, int, float]]) -> str:
    """波ダッシュ(〜)とチルダ(~)を全角チルダ(～)に変換"""
    if text is None: 
        return ""
    return str(text).replace('\u301c', '\uff5e').replace('~', '\uff5e')


def get_byte_count_excel_lenb(text: Optional[str]) -> int:
    """ExcelのLENB関数と同様に、CP932エンコーディングでのバイト数を返す"""
    if text is None:
        return 0
    try:
        return len(str(text).encode('cp932'))
    except Exception: 
        return -1


def get_user_data_dir(preferred_dir: Optional[str] = None) -> str:
    """
    ユーザーデータを保存するディレクトリパスを取得・作成する。
    preferred_dir が指定され、書き込み可能であればそこを使用する。
    それ以外の場合はドキュメントフォルダ内のサブディレクトリを使用する。
    """
    if preferred_dir:
        try:
            # preferred_dir が存在しない場合は作成を試みる
            if not os.path.exists(preferred_dir):
                os.makedirs(preferred_dir, exist_ok=True)

            # 書き込み可能か簡易チェック
            if os.access(preferred_dir, os.W_OK):
                logging.info(f"ユーザーデータは優先ディレクトリ '{preferred_dir}' に保存します。")
                return preferred_dir
            else:
                # preferred_dir は存在するが書き込めない場合
                raise OSError(f"優先ディレクトリ '{preferred_dir}' は存在しますが、書き込み権限がありません。")
        except Exception as e:
            logging.warning(f"優先ディレクトリ '{preferred_dir}' への書き込み/作成に失敗しました: {e}")
            logging.info(f"ドキュメントフォルダ内の '{APP_DATA_SUBDIR}' にフォールバックします。")

    # フォールバック: ドキュメントフォルダ
    try:
        # QStandardPathsを使用して標準的なドキュメントディレクトリを取得
        docs_path = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
        if not docs_path or not os.path.exists(docs_path):  # パスが取得できないか、存在しない場合
            logging.warning(f"標準ドキュメントディレクトリが見つかりません。ホームディレクトリを試みます。")
            docs_path = os.path.expanduser('~')
            if not os.path.exists(docs_path):  # ホームディレクトリも存在しない場合 (非常に稀)
                logging.warning(f"ホームディレクトリも見つかりません。カレントディレクトリにフォールバックします。")
                return os.getcwd()

        app_data_dir = os.path.join(docs_path, APP_DATA_SUBDIR)
        os.makedirs(app_data_dir, exist_ok=True)
        logging.info(f"ユーザーデータはドキュメントフォルダ内の '{app_data_dir}' に保存します。")
        return app_data_dir
    except Exception as e:
        logging.warning(f"ユーザーデータディレクトリの取得/作成に失敗しました: {e}")
        return os.getcwd()  # フォールバックとしてカレントディレクトリ
