"""
商品登録入力ツール - カスタムウィジェットモジュール
"""
from typing import Optional, List, Any
from PyQt5.QtCore import Qt, QTimer, QSize
from PyQt5.QtWidgets import (
    QTextEdit, QTableView, QWidget, QHBoxLayout, QLineEdit, QPushButton,
    QSizePolicy, QDialog, QListWidget, QListWidgetItem, QDialogButtonBox,
    QVBoxLayout, QStyledItemDelegate, QComboBox, QCompleter, QMessageBox,
    QLabel, QProgressBar
)

from constants import (
    HEADER_ATTR_VALUE_PREFIX, HEADER_ATTR_UNIT_PREFIX,
    HEADER_ATTR_ITEM_PREFIX
)


class CustomHtmlTextEdit(QTextEdit):
    """HTMLタグの手動入力をサポートするカスタムテキストエディット"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptRichText(False)  # HTMLタグは手動入力なのでリッチテキストは無効化

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Tab and event.modifiers() == Qt.ControlModifier:
            # Ctrl+Tab でタブ文字を挿入
            self.insertPlainText("\t")
            event.accept()
        elif event.key() == Qt.Key_Tab and not event.modifiers():
            # Tabキーのみの場合は、次のウィジェットにフォーカスを移す
            event.ignore()  # デフォルトのフォーカス処理に任せる
        elif event.key() == Qt.Key_Backtab:  # Shift+Tab
            # Shift+Tabキーの場合は、前のウィジェットにフォーカスを移す
            event.ignore()  # デフォルトのフォーカス処理に任せる
        else:
            super().keyPressEvent(event)  # その他のキーは通常通り処理


class FocusControllingTableView(QTableView):
    """フォーカス制御機能を持つテーブルビュー（固定列用）"""
    
    def __init__(self, product_app_instance, parent=None):
        super().__init__(parent)
        self.other_table_view = None  # ペアとなるテーブルビューを保持
        self.product_app_ref = product_app_instance  # ProductAppの参照を保持

    def setOtherTableView(self, other_table):
        self.other_table_view = other_table

    def keyPressEvent(self, event):
        if not self.other_table_view or not self.model() or not self.currentIndex().isValid():
            super().keyPressEvent(event)
            return

        current_index = self.currentIndex()

        if event.key() == Qt.Key_Tab and not event.modifiers():  # Tabキーのみ
            is_last_visible_column = True
            # 現在の列より右に表示されている列があるかチェック
            for col in range(current_index.column() + 1, self.model().columnCount()):
                if not self.isColumnHidden(col):
                    is_last_visible_column = False
                    break
            
            if is_last_visible_column:
                target_row = current_index.row()
                target_col = 0
                while target_col < self.other_table_view.model().columnCount() and self.other_table_view.isColumnHidden(target_col):
                    target_col += 1
                
                if target_col < self.other_table_view.model().columnCount():
                    self.other_table_view.setCurrentIndex(self.other_table_view.model().index(target_row, target_col))
                    self.other_table_view.setFocus()
                    event.accept()
                    return
        elif event.key() == Qt.Key_Delete:  # Deleteキー処理
            if self.product_app_ref and hasattr(self.product_app_ref, 'delete_selected_skus'):
                self.product_app_ref.delete_selected_skus()
                event.accept()
                return
        super().keyPressEvent(event)


class ScrollableFocusControllingTableView(QTableView):
    """フォーカス制御機能を持つテーブルビュー（スクロール可能列用）"""
    
    def __init__(self, product_app_instance, parent=None):
        super().__init__(parent)
        self.other_table_view = None  # ペアとなるテーブルビューを保持
        self.product_app_ref = product_app_instance  # ProductAppの参照を保持

    def setOtherTableView(self, other_table):
        self.other_table_view = other_table

    def keyPressEvent(self, event):
        if not self.other_table_view or not self.model() or not self.currentIndex().isValid():
            super().keyPressEvent(event)
            return
        
        # SKU削除のDeleteキー処理
        if event.key() == Qt.Key_Delete:
            if self.product_app_ref and hasattr(self.product_app_ref, 'delete_selected_skus'):
                self.product_app_ref.delete_selected_skus()
                event.accept()
                return
        super().keyPressEvent(event)


class MultipleSelectDialog(QDialog):
    """複数選択ダイアログ"""
    
    def __init__(self, options, current_selected_values=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("複数選択")
        self.resize(400, 500)
        self.options = options if options else []
        self.selected_values = list(current_selected_values) if current_selected_values else []

        layout = QVBoxLayout(self)
        self.list_widget = QListWidget()
        for option_text in self.options:
            item = QListWidgetItem(option_text)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            if option_text in self.selected_values:
                item.setCheckState(Qt.Checked)
            else:
                item.setCheckState(Qt.Unchecked)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget)
        self.list_widget.itemClicked.connect(self._toggle_item_check_state_on_click)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def get_selected_values(self):
        """選択されている項目のテキストリストを返します"""
        return sorted(list(set(self.selected_values)))

    def _toggle_item_check_state_on_click(self, item):
        """リストウィジェットのアイテムがクリックされたときに、チェック状態をトグル"""
        option_text = item.text()
        if option_text in self.selected_values:
            # 既に選択されている場合は、選択解除する
            item.setCheckState(Qt.Unchecked)
            self.selected_values.remove(option_text)
        else:
            # まだ選択されていない場合は、選択する
            item.setCheckState(Qt.Checked)
            self.selected_values.append(option_text)


class SkuMultipleAttributeEditor(QWidget):
    """SKU属性の複数選択エディター"""
    
    def __init__(self, options_param, current_value_str="", parent=None, editable_line_edit=False, delimiter_char='|'):
        super().__init__(parent)

        self.delimiter_char = delimiter_char

        # Ensure self.options is a Python list
        if isinstance(options_param, list):
            self.options = options_param
        elif options_param is not None:
            self.options = list(options_param) if hasattr(options_param, '__iter__') else []
        else:
            self.options = []
            
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.line_edit = QLineEdit(self)
        self.line_edit.setReadOnly(not editable_line_edit)
        self.line_edit.setText(current_value_str)
        
        if editable_line_edit and not self.options:
            self.line_edit.setPlaceholderText(f"{self.delimiter_char}区切りで複数入力")
        elif editable_line_edit and self.options:
            self.line_edit.setPlaceholderText(f"{self.delimiter_char}区切りで入力、または選択...")
        
        layout.addWidget(self.line_edit)
        layout.setStretchFactor(self.line_edit, 1)

        self.select_button = QPushButton("選択...", self)
        self.select_button.clicked.connect(self.open_dialog)
        self.select_button.setMinimumWidth(30)
        self.select_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        self.select_button.setVisible(bool(self.options))
        layout.addWidget(self.select_button)
        self.setLayout(layout)

        if bool(self.options):
            QTimer.singleShot(0, self._check_button_visibility_later)

    def _check_button_visibility_later(self):
        """This method will be called after the current event loop cycle"""
        self.select_button.setVisible(bool(self.options))
        if self.select_button.isVisible():
            self.select_button.show()
        self.update()

    def open_dialog(self):
        current_text_in_line_edit = self.line_edit.text()
        current_values_for_dialog = [s.strip() for s in current_text_in_line_edit.split(self.delimiter_char) if s.strip()]

        dialog = MultipleSelectDialog(self.options, current_values_for_dialog, self)
        if dialog.exec_() == QDialog.Accepted:
            selected_values = dialog.get_selected_values()
            self.line_edit.setText(self.delimiter_char.join(selected_values))

    def text(self):
        return self.line_edit.text()

    def setText(self, text_value_str):
        self.line_edit.setText(str(text_value_str))


class SkuAttributeDelegate(QStyledItemDelegate):
    """SKU属性テーブル用のカスタムデリゲート"""
    
    def __init__(self, parent=None):
        super().__init__(parent)

    def createEditor(self, parent, option, index):
        model = index.model()
        if not hasattr(model, '_headers') or not hasattr(model, '_defined_attr_details'):
            return super().createEditor(parent, option, index)

        header_key = model._headers[index.column()]
        attr_num = -1
        if HEADER_ATTR_VALUE_PREFIX in header_key:
            try:
                attr_num = int(header_key.replace(HEADER_ATTR_VALUE_PREFIX, "").strip())
            except ValueError:
                pass
        elif HEADER_ATTR_UNIT_PREFIX in header_key:
            try:
                attr_num = int(header_key.replace(HEADER_ATTR_UNIT_PREFIX, "").strip())
            except ValueError:
                pass

        attr_detail = None
        if 1 <= attr_num <= len(model._defined_attr_details):
            attr_detail = model._defined_attr_details[attr_num - 1]

        if attr_detail:
            options_list = attr_detail.get("options", [])
            unit_options_list_val = attr_detail.get("unit_options_list", [])
            is_multiple_from_def = attr_detail.get("is_multiple_select", False)
            input_method = attr_detail.get("input_method", "").strip()

            # 例外処理のためのフラグと区切り文字を取得
            is_exceptionally_multiple = attr_detail.get("is_exceptionally_multiple", False)
            exception_delimiter = attr_detail.get("exception_delimiter", '|')

            if HEADER_ATTR_VALUE_PREFIX in header_key:
                # --- 例外処理 ---
                if is_exceptionally_multiple:
                    # システム上は単数・記述式だが、UI上は複数選択・指定区切り文字にしたい場合
                    return SkuMultipleAttributeEditor(options_list, "", parent, editable_line_edit=True, delimiter_char=exception_delimiter)

                if input_method == "選択式":
                    if not is_multiple_from_def:  # 単一選択・選択式
                        editor_combo = QComboBox(parent)
                        editor_combo.setEditable(True)
                        editor_combo.addItem("")
                        if options_list:
                            editor_combo.addItems(options_list)
                        # ドロップダウンリストの幅を調整
                        fm = editor_combo.fontMetrics()
                        max_w = 0
                        for i in range(editor_combo.count()):
                            max_w = max(max_w, fm.horizontalAdvance(editor_combo.itemText(i)))
                        # パディングやスクロールバーを考慮して少し余裕を持たせる
                        editor_combo.view().setMinimumWidth(max_w + 70)
                        return editor_combo
                    else:  # 複数選択・選択式
                        if options_list:
                            return SkuMultipleAttributeEditor(options_list, "", parent, editable_line_edit=False, delimiter_char='|')
                        else:
                            editor_line_edit = QLineEdit(parent)
                            editor_line_edit.setReadOnly(True)
                            editor_line_edit.setPlaceholderText("選択肢がありません")
                            return editor_line_edit
                elif input_method == "記述式":
                    if not is_multiple_from_def:  # 単一選択・記述式
                        if options_list:
                            editor_combo = QComboBox(parent)
                            editor_combo.setEditable(True)
                            editor_combo.addItem("")
                            editor_combo.addItems(options_list)
                            completer = QCompleter(options_list, editor_combo)
                            completer.setCaseSensitivity(Qt.CaseInsensitive)
                            completer.setFilterMode(Qt.MatchContains)
                            editor_combo.setCompleter(completer)

                            # ドロップダウンリストの幅を調整
                            fm = editor_combo.fontMetrics()
                            max_w = 0
                            for i in range(editor_combo.count()):
                                max_w = max(max_w, fm.horizontalAdvance(editor_combo.itemText(i)))
                            editor_combo.view().setMinimumWidth(max_w + 70)
                            return editor_combo
                        else:
                            return QLineEdit(parent)
                    else:  # 複数選択・記述式
                        if options_list:
                            return SkuMultipleAttributeEditor(options_list, "", parent, editable_line_edit=True, delimiter_char='|')
                        else:
                            editor_line_edit = QLineEdit(parent)
                            editor_line_edit.setPlaceholderText("|区切りで複数入力")
                            return editor_line_edit
                else:
                    return super().createEditor(parent, option, index)
            elif HEADER_ATTR_UNIT_PREFIX in header_key:
                if unit_options_list_val and \
                   (len(unit_options_list_val) > 1 or \
                    (len(unit_options_list_val) == 1 and unit_options_list_val[0] != '-')):
                    combo = QComboBox(parent)
                    combo.addItem("")
                    combo.addItems(unit_options_list_val)
                    fm = combo.fontMetrics()
                    max_w = 0
                    for i in range(combo.count()):
                        max_w = max(max_w, fm.horizontalAdvance(combo.itemText(i)))
                    combo.view().setMinimumWidth(max_w + 70)
                    return combo
                else:
                    pass  # Fall back to default editor if no specific unit options
        return super().createEditor(parent, option, index)

    def sizeHint(self, option, index):
        model = index.model()
        if not hasattr(model, '_headers') or not hasattr(model, '_defined_attr_details'):
            return super().sizeHint(option, index)

        header_key = model._headers[index.column()]
        attr_num = -1
        if HEADER_ATTR_VALUE_PREFIX in header_key:
            try:
                attr_num = int(header_key.replace(HEADER_ATTR_VALUE_PREFIX, "").strip())
            except ValueError:
                pass

        attr_detail = None
        if 1 <= attr_num <= len(model._defined_attr_details):
            attr_detail = model._defined_attr_details[attr_num - 1]

        if attr_detail and attr_detail.get("input_method", "").strip() in ["選択式", "記述式"]:
            options_list = attr_detail.get("options", [])
            if options_list:
                longest_text = max(options_list, key=len)
                # 描画に必要な幅を計算
                fontMetrics = option.fontMetrics
                text_width = fontMetrics.horizontalAdvance(longest_text)
                # セルのパディングとコンボボックスの矢印部分の幅を考慮
                padding = 20  # 適切なパディング量に調整
                arrow_width = 20  # 矢印部分の幅
                return QSize(text_width + padding + arrow_width, option.rect.height())

        return super().sizeHint(option, index)

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.EditRole)
        if isinstance(editor, SkuMultipleAttributeEditor):
            editor.setText(str(value))
        elif isinstance(editor, QComboBox):
            if editor.count() > 0:
                current_model_value = str(value)
                idx = editor.findText(current_model_value)
                if idx != -1:
                    editor.setCurrentIndex(idx)
                elif not current_model_value:
                    empty_idx = editor.findText("")
                    if empty_idx != -1:
                        editor.setCurrentIndex(empty_idx)
        elif isinstance(editor, QLineEdit) and not isinstance(editor.parent(), SkuMultipleAttributeEditor):
            editor.setText(str(value))
        else:
            super().setEditorData(editor, index)

    def setModelData(self, editor, model, index):
        current_editor_value = ""
        is_value_column_editor = False
        
        if isinstance(editor, SkuMultipleAttributeEditor):
            current_editor_value = editor.text()
            is_value_column_editor = True
        elif isinstance(editor, QComboBox) and editor.isEditable():
            current_editor_value = editor.currentText()
            is_value_column_editor = True
        elif isinstance(editor, QLineEdit) and not isinstance(editor.parent(), SkuMultipleAttributeEditor):
            current_editor_value = editor.text()
            is_value_column_editor = True
        elif isinstance(editor, QComboBox) and not editor.isEditable():  # Unit column
            if editor.count() > 0:
                current_editor_value = editor.currentText()
            else:
                current_editor_value = ""
        else:
            super().setModelData(editor, model, index)
            return

        if is_value_column_editor and hasattr(model, '_defined_attr_details') and hasattr(model, '_headers'):
            header_key = model._headers[index.column()]
            attr_num = -1
            if HEADER_ATTR_VALUE_PREFIX in header_key:
                try:
                    attr_num = int(header_key.replace(HEADER_ATTR_VALUE_PREFIX, "").strip())
                except ValueError:
                    pass

            if 1 <= attr_num <= len(model._defined_attr_details):
                attr_detail = model._defined_attr_details[attr_num - 1]
                input_method = attr_detail.get("input_method", "").strip()
                options = attr_detail.get("options", [])
                is_required_attr = attr_detail.get("is_required", False)

                if is_required_attr and not current_editor_value:
                    current_editor_value = "-"

                if input_method == "選択式":
                    if current_editor_value != "" and current_editor_value != "-" and current_editor_value not in options:
                        QMessageBox.warning(
                            editor.parentWidget(),
                            "入力値エラー",
                            f"属性「{attr_detail.get('name', '不明')}」の値「{current_editor_value}」は、\n"
                            f"定義された選択肢にありません。\n\n"
                            f"選択可能な値: {', '.join(options[:10])}{'...' if len(options) > 10 else ''}\n\n"
                            "正しい値を入力または選択してください。"
                        )
                        if is_required_attr:
                            current_editor_value = "-"
                        else:
                            return
        model.setData(index, current_editor_value, Qt.EditRole)
    
class LoadingDialog(QDialog):
    """起動時に表示される進捗ダイアログ"""

    def __init__(self, message: str, total_steps: int, parent=None):
        super().__init__(parent)
        self.setWindowTitle("読み込み中")
        layout = QVBoxLayout(self)
        self.label = QLabel(message)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, total_steps)
        layout.addWidget(self.label)
        layout.addWidget(self.progress_bar)
        self.setLayout(layout)
        self.setModal(True)
        self.resize(400, 120)

    def update_progress(self, step: int):
        self.progress_bar.setValue(step)

    def setValue(self, value: int):
        self.progress_bar.setValue(value)

    def setLabelText(self, text: str):
        self.label.setText(text)

    def stop_animation(self):
        """アニメーション停止用のダミーメソッド（現時点では何もしない）"""
        pass
