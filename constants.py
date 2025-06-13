"""
商品登録入力ツール - 定数定義モジュール
"""

# アプリケーション設定
APP_NAME = "商品登録入力ツール"
APP_DATA_SUBDIR = "ProductAppUserData"

# ファイル名
TEMPLATE_FILE_NAME = "item_template.xlsm"
CATEGORY_FILE_NAME = "カテゴリ.csv"
MANAGE_FILE_NAME = "item_manage.xlsm"
OUTPUT_FILE_NAME = "item.xlsm"
MATERIAL_SPEC_MASTER_FILE_NAME = "材質・仕様マスタ.csv"

# エンコーディング
DEFAULT_ENCODING = "utf-8-sig"
FALLBACK_ENCODING = "shift_jis"

# UI設定
AUTO_SAVE_INTERVAL_MS = 30000
FROZEN_TABLE_COLUMN_COUNT = 2
PROGRESS_UPDATE_ROW_INTERVAL = 100

# 処理設定
MAX_WORKER_THREADS = 5

# バリデーション設定
DIGIT_COUNT_MYCODE_MAX = 10

# Static constants that don't need configuration (these remain as constants)

# IDマスターファイル
R_GENRE_MASTER_FILE = "r_genre_master.csv"
Y_CATEGORY_MASTER_FILE = "y_category_master.csv"
YA_CATEGORY_MASTER_FILE = "ya_category_master.csv"

# IDマスターファイル内のカラム名
MASTER_ID_COLUMN_DEFAULT = "ID"
MASTER_HIERARCHY_COLUMN_DEFAULT = "階層"
MASTER_NAME_COLUMN_R_GENRE = "RGenreName"
MASTER_NAME_COLUMN_Y_CATEGORY = "YCategoryName"
MASTER_NAME_COLUMN_YA_CATEGORY = None

# ユーザーデータディレクトリに保存されるファイル名（上記で定義済み）

# シート名
MAIN_SHEET_NAME = "Main"
SKU_SHEET_NAME = "SKU"

# 主要なヘッダー名
HEADER_CONTROL_COLUMN = "コントロールカラム"
HEADER_MYCODE = "mycode"
HEADER_PRODUCT_NAME = "商品名_正式表記"
HEADER_PRICE_TAX_INCLUDED = "当店通常価格_税込み"
HEADER_SORT_FIELD = "ソート"
HEADER_SKU_CODE = "SKUコード"
HEADER_CHOICE_NAME = "選択肢名"
HEADER_MEMO = "メモ"
HEADER_GROUP = "グループ"
HEADER_PRODUCT_CODE_SKU = "商品コード"
HEADER_ATTR_ITEM_PREFIX = "商品属性（項目）"
HEADER_ATTR_VALUE_PREFIX = "商品属性（値）"
HEADER_ATTR_UNIT_PREFIX = "商品属性（単位）"
HEADER_R_GENRE_ID = "R_ジャンルID"
HEADER_Y_CATEGORY_ID = "Y_カテゴリID"
HEADER_YA_CATEGORY_ID = "YA_カテゴリID"
HEADER_IMAGE_PATH_RAKUTEN = "画像パス_楽天"
HEADER_IMAGE_DESCRIPTION = "画像説明"
HEADER_YAHOO_ABSTRACT = "Y_abstract"

# --- 材質・仕様マスター関連 ---
MASTER_MATERIAL_SPEC_NAME_COL = "名称"  # CSVのA列ヘッダー
MASTER_MATERIAL_SPEC_DESC_COL = "説明"  # CSVのB列ヘッダー

# --- 説明マークアイコン関連 ---
EXPLANATION_MARK_ICONS_SUBDIR = "explanation_icons"  # バンドルされるアイコンのサブディレクトリ名
EXPLANATION_MARK_FIELD_NAME = "説明マーク_1"  # フォームレイアウトでこの名前のフィールドを特別扱い

# SKU関連
MAX_SKU_ATTRIBUTES = 41
SKU_CODE_SUFFIX_INITIAL = "010"
SKU_CODE_SUFFIX_INCREMENT = 10
SKU_CODE_SUFFIX_MAX = 1990

# UI表示用
UI_HEADER_UNIT = "単位"

# バイト数制限
BYTE_LIMITS = {
    "R_商品名": 255,
    "Y_商品名": 150,
    "R_キャッチコピー": 174,
    "Y_metadesc": 160,
    "Y_キャッチコピー": 60,
}

# HTML入力が想定される複数行フィールド（サニタイゼーション強化対象）
HTML_TEXTEDIT_FIELDS = ["特徴_1", "材質_1", "仕様_1"]

# HTMLサニタイゼーション設定
HTML_SANITIZE_CONFIG = {
    "max_length": 5000,
    "allowed_tags": ["br", "p", "strong", "em", "ul", "ol", "li"],
    "strip_dangerous": True
}

# よく使われる商品色
COMMON_PRODUCT_COLORS = [
    "ナチュラル", "ブラウン", "ホワイト", "グレー", "アイボリー", "ブラック", 
    "レッド", "ブルー", "グリーン", "イエロー", "ピンク", "オレンジ", 
    "パープル", "シルバー", "ゴールド"
]

# タイマー設定（ミリ秒）
LOADING_ANIMATION_INTERVAL_MS = 200  # 200ミリ秒

# UI関連の定数
EXPANDABLE_GROUP_TOGGLE_BUTTON_SIZE = 20  # 展開ボタンのサイズ
TABLE_PADDING = 2  # テーブルビューのパディング

# --- RakutenAttributeDefinitionLoader と関連定数 ---
DEFINITION_CSV_FILE = "mystore_attribute_list_20250520.csv"
RECOMMENDED_LIST_CSV_FILE = "ichiba_recommended_list_20250520.csv"

COL_GENRE_ID = "ジャンルID"
COL_ITEM_NAME_JP = "項目名（日本語）"
COL_ORDER = "並び順"
COL_UNIT_EXISTS = "単位有無"
COL_RECOMMENDED_UNIT_SOURCE = "楽天推奨単位"
COL_INPUT_METHOD = "入力方式"
COL_DEFINITION_GROUP = "商品属性定義詳細グループ"
COL_MULTIPLE_SELECT_ENABLED = "複数値可不可"
COL_REQUIRED_OPTIONAL = "必須/任意"

REC_COL_DEFINITION_GROUP = "商品属性定義詳細グループ"
REC_COL_ITEM_NAME_JP = "項目名（日本語）"
REC_COL_RECOMMENDED_VALUE = "推奨値"

# 例外的に複数選択（カンマ区切り）を許可するSKU属性項目名
EXCEPTIONALLY_MULTIPLE_FIELDS_COMMA_DELIMITED = ["素材", "金属の種類", "樹種"]

# 楽天SKU属性のサイズ情報を格納する項目名
RAKUTEN_SKU_ATTR_NAME_SIZE_INFO = "サイズ情報"

# --- YSpecDefinitionLoader のための定数 ---
YSPEC_CSV_FILE = "Y_spec_data.csv"
YSPEC_COL_CATEGORY_ID = "id"
YSPEC_COL_SPEC_ID = "spec_id"
YSPEC_COL_SPEC_NAME = "spec_name"
YSPEC_COL_SPEC_VALUE_NAME = "spec_value_name"
YSPEC_COL_SPEC_VALUE_ID = "spec_value_id"
YSPEC_COL_SELECTION_TYPE = "selection_type"  # 0:単数, 1:複数
YSPEC_COL_DATA_TYPE = "data_type"            # 1:選択式, 2:整数, 4:整数or小数

# --- Y_specの同期対象項目名 (完全一致) ---
YSPEC_NAME_WIDTH_CM = "幅（cm）"
YSPEC_NAME_DEPTH_CM = "奥行き（cm）"
YSPEC_NAME_HEIGHT_CM = "高さ（cm）"
YSPEC_NAME_WEIGHT = "重量（kg）"