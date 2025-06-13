"""
商品登録入力ツール - テーブルモデルモジュール
"""
from typing import Optional, List, Dict, Any, Union
from PyQt5.QtCore import Qt, QAbstractTableModel, QModelIndex, QVariant
from PyQt5.QtGui import QColor

from constants import (
    HEADER_ATTR_ITEM_PREFIX, HEADER_ATTR_VALUE_PREFIX, HEADER_ATTR_UNIT_PREFIX,
    UI_HEADER_UNIT
)


class SkuTableModel(QAbstractTableModel):
    """SKUテーブル用のモデルクラス"""
    
    HIGHLIGHT_COLOR = QColor(255, 255, 180)
    
    def __init__(self, data=None, headers=None, defined_attr_details=None, parent=None):
        super().__init__(parent)
        self._data = data if data is not None else []
        self._headers = headers if headers is not None else []
        self._defined_attr_details = defined_attr_details if defined_attr_details is not None else []

    def rowCount(self, parent=QModelIndex()):
        return len(self._data)

    def columnCount(self, parent=QModelIndex()):
        return len(self._headers)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Union[str, QColor, None]:
        if not index.isValid() or not (0 <= index.row() < len(self._data) and 0 <= index.column() < len(self._headers)):
            return None
        row = index.row()
        col = index.column()
        header_key = self._headers[col]
        
        if role == Qt.BackgroundRole:
            if self._data[row].get(f"_highlight_{header_key}", False):
                return self.HIGHLIGHT_COLOR
            return None
        
        if role in (Qt.DisplayRole, Qt.EditRole):
            return str(self._data[row].get(header_key, ""))
        
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                if 0 <= section < len(self._headers):
                    original_header = self._headers[section]
                    if HEADER_ATTR_VALUE_PREFIX in original_header:
                        try:
                            attr_num = int(original_header.replace(HEADER_ATTR_VALUE_PREFIX, "").strip())
                            if 1 <= attr_num <= len(self._defined_attr_details):
                                attr_detail = self._defined_attr_details[attr_num - 1]
                                header_text = attr_detail.get("name", original_header)
                                if attr_detail.get("is_multiple_select", False):
                                    header_text += " (複数可)"
                                return header_text
                        except ValueError:
                            pass
                        return original_header
                    elif HEADER_ATTR_UNIT_PREFIX in original_header:
                        return UI_HEADER_UNIT
                    return original_header
            elif orientation == Qt.Vertical:
                return str(section + 1)
        return None

    def setData(self, index: QModelIndex, value: Any, role: int = Qt.EditRole) -> bool:
        if not index.isValid() or role != Qt.EditRole or \
           not (0 <= index.row() < len(self._data) and 0 <= index.column() < len(self._headers)):
            return False
        
        row = index.row()
        col = index.column()
        header_key = self._headers[col]
        
        # 実際のデータ変更をチェック
        old_value = self._data[row].get(header_key, "")
        new_value = str(value)
        
        # 値が実際に変更された場合のみ処理
        if old_value != new_value:
            self._data[row][header_key] = new_value
            
            highlight_flag_key = f"_highlight_{header_key}"
            if highlight_flag_key in self._data[row]:
                self._data[row][highlight_flag_key] = False
            
            self.dataChanged.emit(index, index, [role])
            
            # ProductAppの参照を持っている場合はmark_dirtyを呼び出す
            try:
                parent_obj = self.parent()
                if parent_obj is not None and hasattr(parent_obj, 'mark_dirty') and callable(getattr(parent_obj, 'mark_dirty', None)):
                    parent_obj.mark_dirty()
            except (AttributeError, TypeError) as e:
                import logging
                logging.debug(f"mark_dirty呼び出しでAttributeError/TypeError: {e}")
            except Exception as e:
                import logging
                logging.warning(f"mark_dirty呼び出しで予期しないエラー: {e}", exc_info=True)
        
        return True

    def flags(self, index):
        if not index.isValid():
            return Qt.NoItemFlags
        
        flags = super().flags(index)
        if index.column() < len(self._headers):
            header_key = self._headers[index.column()]
            if HEADER_ATTR_ITEM_PREFIX in header_key:
                return flags & ~Qt.ItemIsEditable
            if not (HEADER_ATTR_ITEM_PREFIX in header_key):
                flags |= Qt.ItemIsEditable
        return flags

    def update_data(self, new_data, new_headers, new_defined_attr_details=None):
        """データの更新"""
        self.beginResetModel()
        self._data = new_data if new_data is not None else []
        self._headers = new_headers if new_headers is not None else []
        self._defined_attr_details = new_defined_attr_details if new_defined_attr_details is not None else []
        self.endResetModel()