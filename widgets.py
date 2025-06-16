"""
商品登録入力ツール - カスタムウィジェットモジュール
"""
import logging
from typing import Optional, List, Any
from PyQt5.QtCore import Qt, QTimer, QSize
from PyQt5.QtWidgets import (
    QTextEdit, QTableView, QWidget, QHBoxLayout, QLineEdit, QPushButton,
    QSizePolicy, QDialog, QListWidget, QListWidgetItem, QDialogButtonBox,
    QVBoxLayout, QStyledItemDelegate, QComboBox, QCompleter, QMessageBox,
    QLabel, QProgressBar, QPlainTextEdit, QInputDialog, QAbstractItemView,
    QApplication, QAbstractItemDelegate
)
from PyQt5.QtGui import QInputMethodEvent

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


class SimpleIMELineEdit(QLineEdit):
    """シンプルなIME対応QLineEdit"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        # 最小限の設定でIME問題を回避
        self.setAttribute(Qt.WA_InputMethodEnabled, True)
        self.setInputMethodHints(Qt.ImhNone)
        self.setFocusPolicy(Qt.StrongFocus)
        
        # 日本語フォントを明示的に指定
        from PyQt5.QtGui import QFont
        font = QFont()
        font.setFamily("Yu Gothic UI")  # Windows標準の日本語フォント
        font.setPointSize(9)
        self.setFont(font)
        
        self.setStyleSheet("""
            QLineEdit { 
                background-color: white;
                border: 1px solid #ccc;
                padding: 2px;
                font-family: "Yu Gothic UI", "Meiryo UI", "MS UI Gothic";
                font-size: 9pt;
            }
        """)
    
    def focusInEvent(self, event):
        """フォーカス取得時の処理"""
        super().focusInEvent(event)
        # フォーカス取得時に全選択しない（IME問題回避）
        self.deselect()
        # カーソルを末尾に移動
        self.setCursorPosition(len(self.text()))
    
    def setText(self, text):
        """テキスト設定時にUTF-8処理を確実に行う"""
        if isinstance(text, bytes):
            try:
                text = text.decode('utf-8')
            except UnicodeDecodeError as e:
                # デコードエラーの場合、エラー位置を記録して安全な文字列に変換
                import logging
                logging.warning(f"Unicode decode error in setText: {e}")
                # 問題のある部分を?に置き換えて処理を継続
                text = text.decode('utf-8', errors='replace')
        elif text is None:
            text = ""
        else:
            text = str(text)
        super().setText(text)


class FocusControllingTableView(QTableView):
    """フォーカス制御機能を持つテーブルビュー（固定列用）"""
    
    def __init__(self, product_app_instance, parent=None):
        super().__init__(parent)
        self.other_table_view = None  # ペアとなるテーブルビューを保持
        self.product_app_ref = product_app_instance  # ProductAppの参照を保持
        self._row_header_clicked = False  # 行ヘッダークリックフラグ
        # IME入力対応の強化
        self.setAttribute(Qt.WA_InputMethodEnabled, True)
        self.setAttribute(Qt.WA_KeyCompression, False)
        self.setFocusPolicy(Qt.StrongFocus)
        # シングルクリックで編集開始
        self.setEditTriggers(QAbstractItemView.CurrentChanged | QAbstractItemView.SelectedClicked)
        # 選択モードを行全体に設定
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        
        # TabキーナビゲーションをOFFにして、手動で制御
        self.setTabKeyNavigation(False)
        
        # 行ヘッダーのクリック処理を設定
        self.verticalHeader().sectionPressed.connect(self._on_row_header_clicked)

    def setOtherTableView(self, other_table):
        self.other_table_view = other_table
    
    def _on_row_header_clicked(self, logical_index):
        """行ヘッダーがクリックされたときの処理"""
        # 行ヘッダークリックフラグを設定
        self._row_header_clicked = True
        
        # 編集状態を強制終了（現在編集中のセルがあれば）
        current_index = self.currentIndex()
        if self.state() == QAbstractItemView.EditingState and current_index.isValid():
            self.closePersistentEditor(current_index)
            
        # 行を選択（編集状態にはしない）
        self.selectRow(logical_index)
        
        # フラグをリセット（遅延実行）
        QTimer.singleShot(100, lambda: setattr(self, '_row_header_clicked', False))
    
    def edit(self, index, trigger=None, event=None):
        """編集開始をコントロール - 行ヘッダークリック時は編集しない"""
        if self._row_header_clicked:
            return False
        # 引数の数に応じて適切に呼び出す
        if trigger is None and event is None:
            return super().edit(index)
        return super().edit(index, trigger, event)
    
    def mousePressEvent(self, event):
        """マウスクリックイベントの処理 - 行ヘッダークリック時の編集防止"""
        index = self.indexAt(event.pos())
        vheader_width = self.verticalHeader().width()
        
        # 行ヘッダー部分がクリックされた場合
        if event.pos().x() < vheader_width:
            # 行を選択するが編集状態にはしない
            if index.isValid():
                self.selectRow(index.row())
            event.accept()
            return
        
        # 通常のセルクリックは元の処理
        super().mousePressEvent(event)
    
    def focusNextPrevChild(self, next):
        """Tabキーナビゲーションを完全に制御"""
        if (hasattr(self.product_app_ref, 'smart_navigation_enabled') and 
            self.product_app_ref.smart_navigation_enabled):
            
            # 疑似的なキーイベントを作成
            from PyQt5.QtGui import QKeyEvent
            from PyQt5.QtCore import QEvent
            
            if next:  # Tab
                fake_event = QKeyEvent(QEvent.KeyPress, Qt.Key_Tab, Qt.NoModifier)
                self.product_app_ref._handle_sku_enter_navigation(self, fake_event)
            else:  # Shift+Tab
                fake_event = QKeyEvent(QEvent.KeyPress, Qt.Key_Backtab, Qt.ShiftModifier)
                self.product_app_ref._handle_sku_backtab_navigation(self, fake_event)
            
            return True  # フォーカス移動を処理したと報告
        
        return super().focusNextPrevChild(next)

    def keyPressEvent(self, event):
        # スマートナビゲーション有効時のTab/Enterキー処理
        if (hasattr(self.product_app_ref, 'smart_navigation_enabled') and 
            self.product_app_ref.smart_navigation_enabled):
            
            if event.key() == Qt.Key_Tab and not event.modifiers():
                self.product_app_ref._handle_sku_enter_navigation(self, event)
                event.accept()
                return
            elif event.key() == Qt.Key_Return:
                self.product_app_ref._handle_sku_enter_navigation(self, event)
                event.accept()
                return
            elif event.key() == Qt.Key_Backtab:
                self.product_app_ref._handle_sku_backtab_navigation(self, event)
                event.accept()
                return
        
        # SKU削除のDeleteキー処理
        if event.key() == Qt.Key_Delete:
            # フォーカスがあるウィジェットがエディタでない場合のみ行削除実行
            focused_widget = QApplication.focusWidget()
            if (not focused_widget or 
                not isinstance(focused_widget, (QLineEdit, QTextEdit, QPlainTextEdit))):
                
                if self.product_app_ref and hasattr(self.product_app_ref, 'delete_selected_skus'):
                    self.product_app_ref.delete_selected_skus()
                    event.accept()
                    return
        
        # 従来のTab処理（スマートナビゲーション無効時）
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
        self._row_header_clicked = False  # 行ヘッダークリックフラグ
        # IME入力対応の強化
        self.setAttribute(Qt.WA_InputMethodEnabled, True)
        self.setAttribute(Qt.WA_KeyCompression, False)
        self.setFocusPolicy(Qt.StrongFocus)
        # シングルクリックで編集開始
        self.setEditTriggers(QAbstractItemView.CurrentChanged | QAbstractItemView.SelectedClicked)
        # 選択モードを行全体に設定
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        
        # 行ヘッダーのクリック処理を設定
        self.verticalHeader().sectionPressed.connect(self._on_row_header_clicked)

    def setOtherTableView(self, other_table):
        self.other_table_view = other_table
    
    def _on_row_header_clicked(self, logical_index):
        """行ヘッダーがクリックされたときの処理"""
        # 行ヘッダークリックフラグを設定
        self._row_header_clicked = True
        
        # 編集状態を強制終了（現在編集中のセルがあれば）
        current_index = self.currentIndex()
        if self.state() == QAbstractItemView.EditingState and current_index.isValid():
            self.closePersistentEditor(current_index)
            
        # 行を選択（編集状態にはしない）
        self.selectRow(logical_index)
        
        # フラグをリセット（遅延実行）
        QTimer.singleShot(100, lambda: setattr(self, '_row_header_clicked', False))
    
    def edit(self, index, trigger=None, event=None):
        """編集開始をコントロール - 行ヘッダークリック時は編集しない"""
        if self._row_header_clicked:
            return False
        # 引数の数に応じて適切に呼び出す
        if trigger is None and event is None:
            return super().edit(index)
        return super().edit(index, trigger, event)
    
    def mousePressEvent(self, event):
        """マウスクリックイベントの処理 - 行ヘッダークリック時の編集防止"""
        index = self.indexAt(event.pos())
        vheader_width = self.verticalHeader().width()
        
        # 行ヘッダー部分がクリックされた場合
        if event.pos().x() < vheader_width:
            # 行を選択するが編集状態にはしない
            if index.isValid():
                self.selectRow(index.row())
            event.accept()
            return
        
        # 通常のセルクリックは元の処理
        super().mousePressEvent(event)

    def keyPressEvent(self, event):
        # スマートナビゲーション有効時のEnterキー処理
        if (hasattr(self.product_app_ref, 'smart_navigation_enabled') and 
            self.product_app_ref.smart_navigation_enabled):
            
            if event.key() == Qt.Key_Return:
                self.product_app_ref._handle_sku_enter_navigation(self, event)
                event.accept()
                return
            elif event.key() == Qt.Key_Backtab:
                self.product_app_ref._handle_sku_backtab_navigation(self, event)
                event.accept()
                return
        
        if not self.other_table_view or not self.model() or not self.currentIndex().isValid():
            super().keyPressEvent(event)
            return
        
        # SKU削除のDeleteキー処理
        if event.key() == Qt.Key_Delete:
            # フォーカスがあるウィジェットがエディタでない場合のみ行削除実行
            focused_widget = QApplication.focusWidget()
            if (not focused_widget or 
                not isinstance(focused_widget, (QLineEdit, QTextEdit, QPlainTextEdit))):
                
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
        layout.setSpacing(4)  # ボタンとLineEditの間隔を狭く
        self.line_edit = SimpleIMELineEdit(self)
        self.line_edit.setReadOnly(not editable_line_edit)
        self.line_edit.setText(current_value_str)
        
        if editable_line_edit and not self.options:
            self.line_edit.setPlaceholderText(f"{self.delimiter_char}区切りで複数入力")
        elif editable_line_edit and self.options:
            self.line_edit.setPlaceholderText(f"{self.delimiter_char}区切りで入力、または選択...")
        
        layout.addWidget(self.line_edit)
        layout.setStretchFactor(self.line_edit, 1)

        self.select_button = QPushButton("選択", self)
        self.select_button.clicked.connect(self.open_dialog)
        self.select_button.setMaximumWidth(50)
        # Tabナビゲーションから除外
        self.select_button.setFocusPolicy(Qt.NoFocus)
        # 高さ制限を緩めて自然なサイズに
        self.select_button.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Preferred)
        self.select_button.setStyleSheet("""
            QPushButton {
                padding: 1px 4px;
                font-size: 9px;
                border: 1px solid #cbd5e1;
                border-radius: 2px;
                background-color: #f8fafc;
                min-height: 18px;
                max-height: 22px;
            }
            QPushButton:hover {
                background-color: #e2e8f0;
                border-color: #94a3b8;
            }
            QPushButton:pressed {
                background-color: #cbd5e1;
            }
        """)

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
                            return JapaneseSimpleIMELineEdit(parent)
                    else:  # 複数選択・記述式
                        if options_list:
                            return SkuMultipleAttributeEditor(options_list, "", parent, editable_line_edit=True, delimiter_char='|')
                        else:
                            editor_line_edit = JapaneseSimpleIMELineEdit(parent)
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
        
        # デフォルトエディタを作成（IME対応強化）
        editor = super().createEditor(parent, option, index)
        if isinstance(editor, QLineEdit):
            # 標準のQLineEditをカスタムIME対応版に置き換え
            ime_editor = JapaneseSimpleIMELineEdit(parent)
            ime_editor.setText(editor.text())
            return ime_editor
        return editor

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
        elif isinstance(editor, (QLineEdit, SimpleIMELineEdit)) and not isinstance(editor.parent(), SkuMultipleAttributeEditor):
            text_value = str(value) if value is not None else ""
            editor.setText(text_value)
            # IME問題回避：全選択状態を解除し、カーソルを末尾に移動
            if hasattr(editor, 'deselect'):
                editor.deselect()
                editor.setCursorPosition(len(text_value))
        else:
            super().setEditorData(editor, index)

    def setModelData(self, editor, model, index):
        current_editor_value = ""
        is_value_column_editor = False
        
        # 入力値の取得と基本的なサニタイゼーション
        if isinstance(editor, SkuMultipleAttributeEditor):
            current_editor_value = self._sanitize_input(editor.text())
            is_value_column_editor = True
        elif isinstance(editor, QComboBox) and editor.isEditable():
            current_editor_value = self._sanitize_input(editor.currentText())
            is_value_column_editor = True
        elif isinstance(editor, (QLineEdit, SimpleIMELineEdit)) and not isinstance(editor.parent(), SkuMultipleAttributeEditor):
            current_editor_value = self._sanitize_input(editor.text())
            is_value_column_editor = True
        elif isinstance(editor, QComboBox) and not editor.isEditable():  # Unit column
            if editor.count() > 0:
                current_editor_value = self._sanitize_input(editor.currentText())
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
    
    def _sanitize_input(self, text_value: str) -> str:
        """入力値のサニタイゼーションとセキュリティ検証"""
        if not isinstance(text_value, str):
            text_value = str(text_value)
        
        # HTMLエスケープ処理
        import html
        sanitized = html.escape(text_value)
        
        # 最大長制限（XSSやバッファオーバーフロー対策）
        MAX_INPUT_LENGTH = 1000
        if len(sanitized) > MAX_INPUT_LENGTH:
            logging.warning(f"入力値が最大長を超えています: {len(sanitized)} > {MAX_INPUT_LENGTH}")
            sanitized = sanitized[:MAX_INPUT_LENGTH]
        
        # NULL文字やその他の制御文字を除去
        sanitized = ''.join(char for char in sanitized if ord(char) >= 32 or char in '\t\n\r')
        
        # 危険な文字パターンの検出
        dangerous_patterns = [
            r'<script',
            r'javascript:',
            r'vbscript:',
            r'on\w+\s*=',
            r'expression\s*\(',
        ]
        
        import re
        for pattern in dangerous_patterns:
            if re.search(pattern, sanitized, re.IGNORECASE):
                logging.error(f"セキュリティ警告: 危険なパターンが検出されました: {pattern}")
                # 危険なパターンを除去
                sanitized = re.sub(pattern, '', sanitized, flags=re.IGNORECASE)
        
        return sanitized.strip()
    
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


class JapaneseLineEdit(QLineEdit):
    """日本語コンテキストメニューを持つQLineEdit"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_japanese_context_menu)
    
    def _show_japanese_context_menu(self, position):
        """日本語化されたコンテキストメニューを表示"""
        from PyQt5.QtWidgets import QMenu, QAction
        
        menu = QMenu(self)
        
        # 元に戻す
        undo_action = QAction("元に戻す", self)
        undo_action.setShortcut("Ctrl+Z")
        undo_action.setEnabled(self.isUndoAvailable())
        undo_action.triggered.connect(self.undo)
        menu.addAction(undo_action)
        
        # やり直し
        redo_action = QAction("やり直し", self)
        redo_action.setShortcut("Ctrl+Y")
        redo_action.setEnabled(self.isRedoAvailable())
        redo_action.triggered.connect(self.redo)
        menu.addAction(redo_action)
        
        menu.addSeparator()
        
        # 切り取り
        cut_action = QAction("切り取り", self)
        cut_action.setShortcut("Ctrl+X")
        cut_action.setEnabled(self.hasSelectedText())
        cut_action.triggered.connect(self.cut)
        menu.addAction(cut_action)
        
        # コピー
        copy_action = QAction("コピー", self)
        copy_action.setShortcut("Ctrl+C")
        copy_action.setEnabled(self.hasSelectedText())
        copy_action.triggered.connect(self.copy)
        menu.addAction(copy_action)
        
        # 貼り付け
        paste_action = QAction("貼り付け", self)
        paste_action.setShortcut("Ctrl+V")
        from PyQt5.QtWidgets import QApplication
        clipboard = QApplication.clipboard()
        paste_action.setEnabled(bool(clipboard.text()))
        paste_action.triggered.connect(self.paste)
        menu.addAction(paste_action)
        
        # 削除
        delete_action = QAction("削除", self)
        delete_action.setShortcut("Delete")
        delete_action.setEnabled(self.hasSelectedText())
        delete_action.triggered.connect(self._delete_selected)
        menu.addAction(delete_action)
        
        menu.addSeparator()
        
        # すべて選択
        select_all_action = QAction("すべて選択", self)
        select_all_action.setShortcut("Ctrl+A")
        select_all_action.setEnabled(bool(self.text()))
        select_all_action.triggered.connect(self.selectAll)
        menu.addAction(select_all_action)
        
        # メニュー表示
        menu.exec_(self.mapToGlobal(position))
    
    def _delete_selected(self):
        """選択されたテキストを削除"""
        if self.hasSelectedText():
            self.del_()


class JapaneseTextEdit(QTextEdit):
    """日本語コンテキストメニューを持つQTextEdit"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_japanese_context_menu)
    
    def _show_japanese_context_menu(self, position):
        """日本語化されたコンテキストメニューを表示"""
        from PyQt5.QtWidgets import QMenu, QAction
        from PyQt5.QtGui import QTextCursor
        
        menu = QMenu(self)
        
        # 元に戻す
        undo_action = QAction("元に戻す", self)
        undo_action.setShortcut("Ctrl+Z")
        undo_action.setEnabled(self.document().isUndoAvailable())
        undo_action.triggered.connect(self.undo)
        menu.addAction(undo_action)
        
        # やり直し
        redo_action = QAction("やり直し", self)
        redo_action.setShortcut("Ctrl+Y")
        redo_action.setEnabled(self.document().isRedoAvailable())
        redo_action.triggered.connect(self.redo)
        menu.addAction(redo_action)
        
        menu.addSeparator()
        
        # 切り取り
        cut_action = QAction("切り取り", self)
        cut_action.setShortcut("Ctrl+X")
        cursor = self.textCursor()
        cut_action.setEnabled(cursor.hasSelection())
        cut_action.triggered.connect(self.cut)
        menu.addAction(cut_action)
        
        # コピー
        copy_action = QAction("コピー", self)
        copy_action.setShortcut("Ctrl+C")
        copy_action.setEnabled(cursor.hasSelection())
        copy_action.triggered.connect(self.copy)
        menu.addAction(copy_action)
        
        # 貼り付け
        paste_action = QAction("貼り付け", self)
        paste_action.setShortcut("Ctrl+V")
        from PyQt5.QtWidgets import QApplication
        clipboard = QApplication.clipboard()
        paste_action.setEnabled(bool(clipboard.text()))
        paste_action.triggered.connect(self.paste)
        menu.addAction(paste_action)
        
        # 削除
        delete_action = QAction("削除", self)
        delete_action.setShortcut("Delete")
        delete_action.setEnabled(cursor.hasSelection())
        delete_action.triggered.connect(self._delete_selected)
        menu.addAction(delete_action)
        
        menu.addSeparator()
        
        # すべて選択
        select_all_action = QAction("すべて選択", self)
        select_all_action.setShortcut("Ctrl+A")
        select_all_action.setEnabled(bool(self.toPlainText()))
        select_all_action.triggered.connect(self.selectAll)
        menu.addAction(select_all_action)
        
        # メニュー表示
        menu.exec_(self.mapToGlobal(position))
    
    def _delete_selected(self):
        """選択されたテキストを削除"""
        cursor = self.textCursor()
        if cursor.hasSelection():
            cursor.removeSelectedText()


class JapaneseHtmlTextEdit(CustomHtmlTextEdit):
    """日本語コンテキストメニューを持つHTMLテキストエディット"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self._show_japanese_context_menu)
    
    def _show_japanese_context_menu(self, position):
        """日本語化されたコンテキストメニューを表示（HTML用の追加機能付き）"""
        from PyQt5.QtWidgets import QMenu, QAction
        from PyQt5.QtGui import QTextCursor
        
        menu = QMenu(self)
        
        # 元に戻す
        undo_action = QAction("元に戻す", self)
        undo_action.setShortcut("Ctrl+Z")
        undo_action.setEnabled(self.document().isUndoAvailable())
        undo_action.triggered.connect(self.undo)
        menu.addAction(undo_action)
        
        # やり直し
        redo_action = QAction("やり直し", self)
        redo_action.setShortcut("Ctrl+Y")
        redo_action.setEnabled(self.document().isRedoAvailable())
        redo_action.triggered.connect(self.redo)
        menu.addAction(redo_action)
        
        menu.addSeparator()
        
        # 切り取り
        cut_action = QAction("切り取り", self)
        cut_action.setShortcut("Ctrl+X")
        cursor = self.textCursor()
        cut_action.setEnabled(cursor.hasSelection())
        cut_action.triggered.connect(self.cut)
        menu.addAction(cut_action)
        
        # コピー
        copy_action = QAction("コピー", self)
        copy_action.setShortcut("Ctrl+C")
        copy_action.setEnabled(cursor.hasSelection())
        copy_action.triggered.connect(self.copy)
        menu.addAction(copy_action)
        
        # 貼り付け
        paste_action = QAction("貼り付け", self)
        paste_action.setShortcut("Ctrl+V")
        from PyQt5.QtWidgets import QApplication
        clipboard = QApplication.clipboard()
        paste_action.setEnabled(bool(clipboard.text()))
        paste_action.triggered.connect(self.paste)
        menu.addAction(paste_action)
        
        # 削除
        delete_action = QAction("削除", self)
        delete_action.setShortcut("Delete")
        delete_action.setEnabled(cursor.hasSelection())
        delete_action.triggered.connect(self._delete_selected)
        menu.addAction(delete_action)
        
        menu.addSeparator()
        
        # すべて選択
        select_all_action = QAction("すべて選択", self)
        select_all_action.setShortcut("Ctrl+A")
        select_all_action.setEnabled(bool(self.toPlainText()))
        select_all_action.triggered.connect(self.selectAll)
        menu.addAction(select_all_action)
        
        menu.addSeparator()
        
        # HTML用の特別な機能
        insert_br_action = QAction("改行タグ挿入 (<br>)", self)
        insert_br_action.triggered.connect(lambda: self.insertPlainText("<br>"))
        menu.addAction(insert_br_action)
        
        # メニュー表示
        menu.exec_(self.mapToGlobal(position))
    
    def _delete_selected(self):
        """選択されたテキストを削除"""
        cursor = self.textCursor()
        if cursor.hasSelection():
            cursor.removeSelectedText()


class JapaneseSimpleIMELineEdit(JapaneseLineEdit):
    """日本語コンテキストメニュー付きのIME対応QLineEdit"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        # IME問題を回避する設定
        self.setAttribute(Qt.WA_InputMethodEnabled, True)
        self.setInputMethodHints(Qt.ImhNone)
        self.setFocusPolicy(Qt.StrongFocus)
        
        # 日本語フォントを明示的に指定
        from PyQt5.QtGui import QFont
        font = QFont()
        font.setFamily("Yu Gothic UI")  # Windows標準の日本語フォント
        font.setPointSize(9)
        self.setFont(font)
        
        self.setStyleSheet("""
            QLineEdit { 
                background-color: white;
                border: 1px solid #ccc;
                padding: 2px;
                font-family: "Yu Gothic UI", "Meiryo UI", "MS UI Gothic";
                font-size: 9pt;
            }
        """)
    
    def focusInEvent(self, event):
        """フォーカス取得時の処理"""
        super().focusInEvent(event)
        # IMEを確実に有効化
        self.setAttribute(Qt.WA_InputMethodEnabled, True)
    
    def inputMethodEvent(self, event):
        """IME入力イベントの処理"""
        # 通常の処理を実行
        super().inputMethodEvent(event)
        # 変換中のテキストがある場合は、その長さを保持
        preedit_str = event.preeditString()
        if preedit_str:
            self.setProperty("ime_preedit_length", len(preedit_str))
        else:
            self.setProperty("ime_preedit_length", 0)


class SearchLineEdit(JapaneseLineEdit):
    """検索用のラインエディット（ESCキー対応）"""
    
    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            # 親のSearchPanelを直接閉じる
            search_panel = self.parent()
            if hasattr(search_panel, 'close_panel'):
                search_panel.close_panel()
            event.accept()
            return
        super().keyPressEvent(event)
