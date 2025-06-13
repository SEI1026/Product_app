"""
商品登録入力ツール - データローダーモジュール
"""
import os
import sys
import csv
import re
import logging
from typing import Optional, List, Dict, Tuple
from PyQt5.QtWidgets import QApplication, QDialog


from constants import (
    YSPEC_CSV_FILE, YSPEC_COL_CATEGORY_ID, YSPEC_COL_SPEC_ID, YSPEC_COL_SPEC_NAME,
    YSPEC_COL_SPEC_VALUE_NAME, YSPEC_COL_SPEC_VALUE_ID, YSPEC_COL_SELECTION_TYPE,
    YSPEC_COL_DATA_TYPE, DEFAULT_ENCODING, PROGRESS_UPDATE_ROW_INTERVAL,
    COL_GENRE_ID, COL_ITEM_NAME_JP, COL_ORDER, COL_UNIT_EXISTS,
    COL_RECOMMENDED_UNIT_SOURCE, COL_INPUT_METHOD, COL_DEFINITION_GROUP,
    COL_MULTIPLE_SELECT_ENABLED, COL_REQUIRED_OPTIONAL, REC_COL_DEFINITION_GROUP,
    REC_COL_ITEM_NAME_JP, REC_COL_RECOMMENDED_VALUE, EXCEPTIONALLY_MULTIPLE_FIELDS_COMMA_DELIMITED,
    DEFINITION_CSV_FILE, RECOMMENDED_LIST_CSV_FILE, MASTER_MATERIAL_SPEC_NAME_COL,
    MASTER_MATERIAL_SPEC_DESC_COL, MATERIAL_SPEC_MASTER_FILE_NAME,
    EXPLANATION_MARK_ICONS_SUBDIR, MASTER_ID_COLUMN_DEFAULT, MASTER_HIERARCHY_COLUMN_DEFAULT
)
from utils import open_csv_file_with_fallback, normalize_wave_dash


class YSpecDefinitionLoader:
    """Yahoo!スペック定義を読み込むクラス"""
    
    def __init__(self, base_path, progress_dialog=None):
        self.base_path = base_path
        self.progress_dialog = progress_dialog
        self.spec_definitions = {}  # {category_id: [{spec_id: ..., spec_name: ..., options: [...], selection_type: ..., data_type: ...}, ...]}
        self._load_spec_data()

    def _load_spec_data(self):
        filepath = os.path.join(self.base_path, YSPEC_CSV_FILE)
        try:
            with open_csv_file_with_fallback(filepath, 'r', self.progress_dialog, "Yahoo!スペック定義") as (f, delimiter, encoding_name):
                reader = csv.DictReader(f, delimiter=delimiter)
                required_cols = [
                    YSPEC_COL_CATEGORY_ID, YSPEC_COL_SPEC_ID, YSPEC_COL_SPEC_NAME,
                    YSPEC_COL_SELECTION_TYPE, YSPEC_COL_DATA_TYPE
                ]

                if not reader.fieldnames or not all(col in reader.fieldnames for col in required_cols):
                    encoding_label = f" ({encoding_name})" if encoding_name != DEFAULT_ENCODING else ""
                    logging.warning(f"Yahoo!スペック定義書 '{filepath}'{encoding_label} に必須ヘッダーが見つかりません。不足: {[h for h in required_cols if h not in (reader.fieldnames or [])]}")
                    return

                temp_specs_by_cat_and_spec_id = {}  # {(category_id, spec_id): spec_detail}

                for row_num, row_dict_raw in enumerate(reader, start=2):
                    row_data = {str(k).strip(): str(v).strip() if v is not None else "" for k, v in row_dict_raw.items()}

                    category_id = row_data.get(YSPEC_COL_CATEGORY_ID)
                    spec_id = row_data.get(YSPEC_COL_SPEC_ID)
                    spec_name = row_data.get(YSPEC_COL_SPEC_NAME)
                    selection_type_str = row_data.get(YSPEC_COL_SELECTION_TYPE)
                    data_type_str = row_data.get(YSPEC_COL_DATA_TYPE)
                    spec_value_name = row_data.get(YSPEC_COL_SPEC_VALUE_NAME)
                    spec_value_id = row_data.get(YSPEC_COL_SPEC_VALUE_ID)

                    if not category_id or not spec_id or not spec_name or not selection_type_str or not data_type_str:
                        logging.warning(f"Yahoo!スペック定義書 行 {row_num}: 必須情報 (カテゴリID, specID, spec名, selection_type, data_type) が不足しています。スキップします。")
                        continue

                    try:
                        selection_type = int(selection_type_str)
                        data_type = int(data_type_str)
                    except ValueError:
                        logging.warning(f"Yahoo!スペック定義書 行 {row_num}: selection_typeまたはdata_typeが数値ではありません。スキップします。")
                        continue

                    current_spec_key = (category_id, spec_id)

                    if current_spec_key not in temp_specs_by_cat_and_spec_id:
                        temp_specs_by_cat_and_spec_id[current_spec_key] = {
                            "spec_id": spec_id,
                            "spec_name": spec_name,
                            "selection_type": selection_type,
                            "data_type": data_type,
                            "options": []  # { "value_id": ..., "value_name": ... }
                        }
                    
                    # data_type が 1 (テキスト選択) の場合のみ、選択肢を追加
                    if data_type == 1 and spec_value_name and spec_value_id:
                        # 既に同じspec_value_idの選択肢がないか確認
                        if not any(opt["value_id"] == spec_value_id for opt in temp_specs_by_cat_and_spec_id[current_spec_key]["options"]):
                            temp_specs_by_cat_and_spec_id[current_spec_key]["options"].append({
                                "value_id": spec_value_id,
                                "value_name": spec_value_name
                            })
                    
                    if reader.line_num % (PROGRESS_UPDATE_ROW_INTERVAL * 4) == 0 and self.progress_dialog:
                        QApplication.processEvents()
                
                # temp_specs_by_cat_and_spec_id から self.spec_definitions に再構成
                for (cat_id, _), spec_detail in temp_specs_by_cat_and_spec_id.items():
                    if cat_id not in self.spec_definitions:
                        self.spec_definitions[cat_id] = []
                    # spec_id の重複を避けるため (既に同じspec_idの項目がなければ追加)
                    if not any(s["spec_id"] == spec_detail["spec_id"] for s in self.spec_definitions[cat_id]):
                        self.spec_definitions[cat_id].append(spec_detail)
                
                # 各カテゴリのスペックリストを spec_id の昇順でソート
                for cat_id in self.spec_definitions:
                    self.spec_definitions[cat_id].sort(key=lambda x: int(x["spec_id"]) if x["spec_id"].isdigit() else float('inf'))

        except FileNotFoundError:
            return
        except UnicodeDecodeError as e:
            logging.error(f"エンコーディングエラー '{filepath}': {e}", exc_info=True)
            raise UnicodeDecodeError(
                e.encoding, e.object, e.start, e.end,
                f"ファイル '{filepath}' の文字エンコーディングが正しくありません。UTF-8またはShift_JISで保存し直してください。"
            )
        except csv.Error as e:
            logging.error(f"Yahoo!スペック定義書 '{filepath}' のCSVパースエラーが発生しました: {e}", exc_info=True)
        except MemoryError as e:
            logging.error(f"Yahoo!スペック定義書 '{filepath}' の処理中にメモリ不足が発生しました。", exc_info=True)
        except Exception as e:
            logging.error(f"Yahoo!スペック定義書 '{filepath}' の処理中に予期せぬエラーが発生しました。", exc_info=True)
        
        if self.spec_definitions:
            total_specs_count = sum(len(specs) for specs in self.spec_definitions.values())
            logging.info(f"{len(self.spec_definitions)}カテゴリ、合計{total_specs_count}件のYahoo!スペック項目定義を読み込みました。")
        else:
            logging.warning(f"Yahoo!スペック定義書から有効なデータが読み込まれませんでした。")

    def get_specs_for_category(self, category_id):
        """指定されたYカテゴリIDに対応するスペック定義のリストを返す"""
        return self.spec_definitions.get(str(category_id).strip(), [])


class RakutenAttributeDefinitionLoader:
    """楽天商品属性定義を読み込むクラス"""
    
    def __init__(self, base_path, progress_dialog=None):
        self.base_path = base_path
        self.genre_definitions = {}
        self.recommended_values_map = {}
        self.progress_dialog = progress_dialog
        self._load_definition_data()

    def _load_definition_data(self):
        definition_file_path = os.path.join(self.base_path, DEFINITION_CSV_FILE)
        recommended_list_file_path = os.path.join(self.base_path, RECOMMENDED_LIST_CSV_FILE)
        self._parse_definition_csv(definition_file_path)
        self._parse_recommended_list_csv(recommended_list_file_path)
        
        if not self.genre_definitions:
            logging.warning(f"'{definition_file_path}' から楽天属性定義の読み込みに失敗しました。")
        else:
            logging.info(f"{sum(len(attrs) for attrs in self.genre_definitions.values())}件の楽天属性定義({len(self.genre_definitions)}ジャンルID)を読み込みました。")
        
        if self.recommended_values_map:
            logging.info(f"{len(self.recommended_values_map)}件の楽天推奨値キーを'{recommended_list_file_path}'から読み込みました。")

    def _parse_definition_csv(self, filepath):
        try:
            with open_csv_file_with_fallback(filepath, 'r', self.progress_dialog, "属性定義書") as (f, delimiter, encoding_name):
                reader = csv.DictReader(f, delimiter=delimiter)
                required_cols = [COL_GENRE_ID, COL_ITEM_NAME_JP, COL_ORDER]
                if not reader.fieldnames or not all(col in reader.fieldnames for col in required_cols):
                    encoding_label = f" ({encoding_name})" if encoding_name != DEFAULT_ENCODING else ""
                    logging.warning(f"楽天属性定義書 '{filepath}'{encoding_label} に必須ヘッダーが見つかりません。不足: {[h for h in required_cols if h not in (reader.fieldnames or [])]}")
                    return

                source_file_label = f"{os.path.basename(filepath)}"
                if encoding_name != DEFAULT_ENCODING:
                    source_file_label += f" ({encoding_name})"

                for row_num, row_dict_raw in enumerate(reader, start=2):
                    row_data = {str(k).strip(): str(v).strip() if v is not None else "" for k, v in row_dict_raw.items()}
                    self._process_definition_row(row_data, f"{source_file_label} (行 {row_num})")
                    if reader.line_num % PROGRESS_UPDATE_ROW_INTERVAL == 0:
                        QApplication.processEvents()
                        
        except FileNotFoundError:
            logging.warning(f"楽天属性定義書ファイル '{filepath}' が見つかりません。")
            return
        except UnicodeDecodeError:
            logging.warning(f"楽天属性定義書 '{filepath}' のデコードに全てのエンコーディングで失敗しました。", exc_info=True)
        except csv.Error as e:
            logging.error(f"楽天属性定義書 '{filepath}' のCSVパースエラーが発生しました: {e}", exc_info=True)
        except MemoryError as e:
            logging.error(f"楽天属性定義書 '{filepath}' の処理中にメモリ不足が発生しました。", exc_info=True)
        except Exception as e:
            logging.error(f"楽天属性定義書 '{filepath}' の処理中に予期せぬエラーが発生しました。", exc_info=True)

    def _process_definition_row(self, row_data, source_info):
        genre_id = row_data.get(COL_GENRE_ID, "")
        item_name = row_data.get(COL_ITEM_NAME_JP, "")
        order_str = row_data.get(COL_ORDER, "")
        if not genre_id or not item_name or not order_str:
            return

        is_exceptionally_multiple = item_name in EXCEPTIONALLY_MULTIPLE_FIELDS_COMMA_DELIMITED
        exception_delimiter = ',' if is_exceptionally_multiple else '|'

        try:
            order = int(order_str)
        except ValueError:
            logging.warning(f"楽天属性定義書 ({source_info}): 並び順 '{order_str}' が数値ではありません。")
            return
            
        unit_options_str = row_data.get(COL_RECOMMENDED_UNIT_SOURCE, "")
        unit_options_list = [opt.strip() for opt in unit_options_str.split('|') if opt.strip()] if unit_options_str else []
        
        attribute_detail = {
            "name": item_name,
            "order": order,
            "unit_exists_raw": row_data.get(COL_UNIT_EXISTS, ""),
            "unit_options_list": unit_options_list,
            "input_method": row_data.get(COL_INPUT_METHOD, ""),
            "definition_group": row_data.get(COL_DEFINITION_GROUP, ""),
            "options": [],
            "is_multiple_select": row_data.get(COL_MULTIPLE_SELECT_ENABLED, "不可").strip() == "可",
            "is_required": row_data.get(COL_REQUIRED_OPTIONAL, "任意").strip() == "必須",
            "is_exceptionally_multiple": is_exceptionally_multiple,
            "exception_delimiter": exception_delimiter
        }
        
        if genre_id not in self.genre_definitions:
            self.genre_definitions[genre_id] = []
        self.genre_definitions[genre_id].append(attribute_detail)
        self.genre_definitions[genre_id].sort(key=lambda x: x.get("order", float('inf')))

    def _parse_recommended_list_csv(self, filepath):
        try:
            with open_csv_file_with_fallback(filepath, 'r', self.progress_dialog, "推奨値リスト") as (f, delimiter, encoding_name):
                reader = csv.DictReader(f, delimiter=delimiter)
                required_rec_cols = [REC_COL_DEFINITION_GROUP, REC_COL_ITEM_NAME_JP, REC_COL_RECOMMENDED_VALUE]
                if not reader.fieldnames or not all(col in reader.fieldnames for col in required_rec_cols):
                    encoding_label = f" ({encoding_name})" if encoding_name != DEFAULT_ENCODING else ""
                    logging.warning(f"楽天推奨値リスト '{filepath}'{encoding_label} に必須ヘッダーが見つかりません。不足: {[h for h in required_rec_cols if h not in (reader.fieldnames or [])]}")
                    return

                for row_dict_raw in reader:
                    row_data = {str(k).strip(): str(v).strip() if v is not None else "" for k, v in row_dict_raw.items()}
                    def_group = row_data.get(REC_COL_DEFINITION_GROUP)
                    item_name = row_data.get(REC_COL_ITEM_NAME_JP)
                    rec_value = row_data.get(REC_COL_RECOMMENDED_VALUE)
                    if def_group and item_name and rec_value:
                        key = (def_group, item_name)
                        if key not in self.recommended_values_map:
                            self.recommended_values_map[key] = []
                        if rec_value not in self.recommended_values_map[key]:
                            self.recommended_values_map[key].append(rec_value)
                    if reader.line_num % (PROGRESS_UPDATE_ROW_INTERVAL * 4) == 0:
                        QApplication.processEvents()
                        
        except FileNotFoundError:
            logging.info(f"楽天推奨値リストファイル '{filepath}' が見つかりません。")
            return
        except UnicodeDecodeError:
            logging.warning(f"楽天推奨値リスト '{filepath}' のデコードに全てのエンコーディングで失敗しました。", exc_info=True)
        except csv.Error as e:
            logging.error(f"楽天推奨値リスト '{filepath}' のCSVパースエラーが発生しました: {e}", exc_info=True)
        except MemoryError as e:
            logging.error(f"楽天推奨値リスト '{filepath}' の処理中にメモリ不足が発生しました。", exc_info=True)
        except Exception as e:
            logging.error(f"楽天推奨値リスト '{filepath}' の処理中に予期せぬエラーが発生しました。", exc_info=True)

    def get_attribute_details_for_genre(self, genre_id):
        genre_id_str = str(genre_id).strip()
        details_list_original = self.genre_definitions.get(genre_id_str, [])
        details_list_with_options = []
        for detail_orig in details_list_original:
            detail_copy = detail_orig.copy()
            key_for_rec = (detail_copy.get("definition_group"), detail_copy.get("name"))
            detail_copy["options"] = self.recommended_values_map.get(key_for_rec, [])
            details_list_with_options.append(detail_copy)
        return details_list_with_options


def load_categories_from_csv(filepath: str, progress_dialog: Optional[QDialog] = None) -> List[Tuple[int, str, str]]:
    """カテゴリCSVファイルを読み込む"""
    categories: List[Tuple[int, str, str]] = []
    try:
        with open_csv_file_with_fallback(filepath, 'r', progress_dialog, "カテゴリ") as (f, delimiter, encoding_name):
            reader = csv.reader(f, delimiter=delimiter)
            next(reader, None)  # ヘッダー行をスキップ
            for row in reader:
                if len(row) >= 3:
                    try:
                        level = int(row[0])
                        raw_category_name = row[1]  # 元のカテゴリ名
                        raw_parent_name = row[2]    # 元の親カテゴリ名

                        name = normalize_wave_dash(raw_category_name)
                        parent = normalize_wave_dash(raw_parent_name)

                        categories.append((level, name, parent))
                        if reader.line_num % PROGRESS_UPDATE_ROW_INTERVAL == 0:
                            QApplication.processEvents()
                    except ValueError:
                        logging.warning(f"カテゴリファイル '{filepath}' の不正な行 (レベルが数値ではありません): {row}")
                        continue
    except FileNotFoundError:
        logging.error(f"カテゴリファイル '{filepath}' が見つかりません。")
        raise
    except UnicodeDecodeError:
        logging.error(f"カテゴリファイル '{filepath}' のデコードに失敗しました。", exc_info=True)
        raise
    except Exception as e:
        logging.error(f"カテゴリファイル '{filepath}' の読み込み中に予期せぬエラーが発生しました。", exc_info=True)
        raise
    return categories


def load_explanation_mark_icons(base_path: str, progress_dialog: Optional[QDialog] = None) -> List[Dict[str, str]]:
    """説明マークアイコンファイルを読み込む"""
    icons_data: List[Dict[str, str]] = []
    icons_dir = os.path.join(base_path, EXPLANATION_MARK_ICONS_SUBDIR)
    file_label = "説明マークアイコン"

    if progress_dialog:
        progress_dialog.setLabelText(f"{file_label} ({EXPLANATION_MARK_ICONS_SUBDIR}) を検索中...")
        QApplication.processEvents()

    if not os.path.isdir(icons_dir):
        logging.info(f"{file_label}ディレクトリ '{icons_dir}' が見つかりません。説明マークアイコン機能は利用できません。")
        return []

    try:
        for filename in os.listdir(icons_dir):
            # jpg と png をサポート
            if filename.lower().endswith((".jpg", ".jpeg", ".png")):
                match = re.match(r"(\d+)_(.+)\.(jpg|jpeg|png)", filename, re.IGNORECASE)
                if match:
                    icon_id = match.group(1)
                    description = match.group(2)
                    filepath = os.path.join(icons_dir, filename)
                    icons_data.append({
                        "id": icon_id,
                        "description": description,
                        "path": filepath,
                        "filename": filename
                    })
                else:
                    logging.warning(f"{file_label}ファイル名 '{filename}' の形式が不正です (例: 1_説明.jpg)。スキップします。")
            if progress_dialog and len(icons_data) % 20 == 0:
                QApplication.processEvents()
        
        icons_data.sort(key=lambda x: int(x["id"]))  # IDの昇順でソート
        logging.info(f"{file_label}ディレクトリ '{icons_dir}' から {len(icons_data)} 件のアイコン情報を読み込みました。")
    except Exception as e:
        logging.warning(f"{file_label}ディレクトリ '{icons_dir}' の読み込み中にエラーが発生しました。", exc_info=True)
    return icons_data


def load_material_spec_master(filepath: str, progress_dialog: Optional[QDialog] = None) -> Dict[str, str]:
    """材質・仕様マスターCSVファイルを読み込む"""
    master_data: Dict[str, str] = {}  # {"名称": "説明"}
    file_label = "材質・仕様マスター"
    
    try:
        with open_csv_file_with_fallback(filepath, 'r', progress_dialog, file_label) as (f, delimiter, encoding_name):
            reader = csv.DictReader(f, delimiter=delimiter)
            
            name_col = MASTER_MATERIAL_SPEC_NAME_COL
            desc_col = MASTER_MATERIAL_SPEC_DESC_COL

            if not reader.fieldnames or name_col not in reader.fieldnames or desc_col not in reader.fieldnames:
                missing_cols = []
                if not reader.fieldnames:
                    missing_cols = [name_col, desc_col]
                else:
                    if name_col not in reader.fieldnames:
                        missing_cols.append(name_col)
                    if desc_col not in reader.fieldnames:
                        missing_cols.append(desc_col)
                
                encoding_label = f" ({encoding_name})" if encoding_name != DEFAULT_ENCODING else ""
                logging.warning(f"{file_label} '{filepath}'{encoding_label} に必須ヘッダー ({name_col}, {desc_col}) が見つかりません。不足: {missing_cols}")
                return {}

            for row_num, row_dict_raw in enumerate(reader, start=2):
                row_data = {str(k).strip(): str(v).strip() if v is not None else "" for k, v in row_dict_raw.items()}
                name = row_data.get(name_col)
                description = row_data.get(desc_col)
                if name:
                    if name in master_data:
                        logging.warning(f"{file_label} '{filepath}' 行 {row_num}: 名称 '{name}' が重複しています。最初の定義を使用します。")
                    else:
                        master_data[name] = description
                else:
                    logging.warning(f"{file_label} '{filepath}' 行 {row_num}: 名称が空です。スキップします。")
                
                if reader.line_num % PROGRESS_UPDATE_ROW_INTERVAL == 0 and progress_dialog:
                    QApplication.processEvents()
            logging.info(f"{file_label} '{filepath}' から {len(master_data)} 件のデータを読み込みました。")
    except FileNotFoundError:
        logging.info(f"{file_label}ファイル '{filepath}' が見つかりません。材質・仕様マスター機能は利用できません。")
        return {}
    except Exception as e:
        logging.warning(f"{file_label} '{filepath}' の読み込み中にエラーが発生しました。", exc_info=True)
        return {}
    return master_data


def load_id_master_data(filepath, id_col_header, name_col_header, hierarchy_col_header,
                       progress_dialog=None, file_label="IDマスター"):
    """IDマスターデータを読み込む"""
    all_searchable_data_list = []

    # セキュリティ強化: ファイルパス検証
    if not os.path.isabs(filepath):
        base_dir = sys._MEIPASS if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        # パス正規化でディレクトリトラバーサルを防ぐ
        normalized_filepath = os.path.normpath(filepath)
        if '..' in normalized_filepath or normalized_filepath.startswith(('/') if os.name != 'nt' else ('/', '\\')):
            logging.error(f"セキュリティ警告: 不正なファイルパスが指定されました: {filepath}")
            raise ValueError("不正なファイルパスです")
        effective_filepath = os.path.join(base_dir, normalized_filepath)
    else:
        # 絶対パスの場合も検証
        effective_filepath = os.path.abspath(os.path.normpath(filepath))
        
    # ファイルパスが許可された範囲内にあることを確認
    allowed_base = os.path.abspath(".")
    if not effective_filepath.startswith(allowed_base):
        logging.error(f"セキュリティ警告: 許可されていないディレクトリへのアクセスが試行されました: {effective_filepath}")
        raise ValueError("許可されていないファイルパスです")

    try:
        with open_csv_file_with_fallback(effective_filepath, 'r', progress_dialog, file_label) as (f, delimiter, encoding_name):
            reader = csv.DictReader(f, delimiter=delimiter)
            required_headers = [id_col_header, hierarchy_col_header]
            if name_col_header:
                required_headers.append(name_col_header)

            if not reader.fieldnames or not all(h in reader.fieldnames for h in required_headers):
                missing_h = [h for h in required_headers if h not in (reader.fieldnames or [])]
                encoding_label = f" ({encoding_name})" if encoding_name != DEFAULT_ENCODING else ""
                logging.warning(f"IDマスターファイル '{effective_filepath}'{encoding_label} に必須ヘッダーが見つかりません。不足: {missing_h}")
                return []

            for row_num, row in enumerate(reader):
                item_id = row.get(id_col_header, "").strip()
                item_hierarchy = normalize_wave_dash(row.get(hierarchy_col_header, "")).strip()
                item_name = normalize_wave_dash(row.get(name_col_header, "")).strip() if name_col_header else ""

                if item_id and item_hierarchy:
                    data_entry = {'id': item_id, 'name': item_name, 'hierarchy': item_hierarchy}
                    all_searchable_data_list.append(data_entry)
                if reader.line_num % PROGRESS_UPDATE_ROW_INTERVAL == 0:
                    QApplication.processEvents()
    except FileNotFoundError:
        logging.info(f"IDマスターファイル '{effective_filepath}' が見つかりません。")
        return []
    except UnicodeDecodeError:
        logging.error(f"IDマスターファイル '{effective_filepath}' のデコードに失敗しました。", exc_info=True)
        return []
    except Exception as e:
        logging.warning(f"IDマスターファイル '{effective_filepath}' の読み込み中にエラーが発生しました。", exc_info=True)
        return []

    if not all_searchable_data_list:
        logging.info(f"IDマスターファイル '{effective_filepath}' は空か、有効なデータを含んでいません。")
    return all_searchable_data_list
