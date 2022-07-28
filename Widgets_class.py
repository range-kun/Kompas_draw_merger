import os

from PyQt5 import QtCore, QtGui, QtWidgets


class MakeWidgets(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        QtWidgets.QMainWindow.__init__(self, parent=None)

    def make_button(self, *, text, command=None, size=None, font=None, size_policy=None, parent=None, enabled=True):
        button = QtWidgets.QPushButton(parent or self.centralwidget)
        button.setEnabled(enabled)
        if size:
            button.setMaximumSize(QtCore.QSize(*size))
        if command:
            button.clicked.connect(command)
        if font:
            button.setFont(font)
        if size_policy:
            button.setSizePolicy(size_policy)
        button.setText(text)
        return button

    def make_line_edit(self, *, text=None, status=None, parent=None, font=None):
        line = QtWidgets.QLineEdit(parent or self.centralwidget)
        line.setStatusTip(status)
        if font:
            line.setFont(font)
        line.setPlaceholderText(text)
        return line

    @staticmethod
    def make_plain_text(*, text=None, status=None, parent=None, size_policy=None, font=None):
        plain_text = QtWidgets.QPlainTextEdit(parent)
        plain_text.setStatusTip(status)
        plain_text.setPlainText(text)
        plain_text.setFont(font)
        if size_policy:
            plain_text.setSizePolicy(size_policy)
        return plain_text

    def make_text_edit(
            self, *,
            text=None,
            status=None,
            parent=None,
            placeholder=None,
            size_policy=None,
            font=None
    ):
        text_edit = QtWidgets.QTextEdit(parent or self.centralwidget)
        text_edit.setStatusTip(status)
        text_edit.setPlainText(text)
        text_edit.setFont(font)
        text_edit.setPlaceholderText(placeholder)
        if size_policy:
            text_edit.setSizePolicy(size_policy)
        return text_edit

    def make_line(self, orientation, width=None):
        line = QtWidgets.QFrame(self.centralwidget)
        if orientation == 'horizontal':
            line.setFrameShape(QtWidgets.QFrame.HLine)
        elif orientation == 'vertical':
            line.setFrameShape(QtWidgets.QFrame.VLine)
        if width:
            line.setLineWidth(width)
            line.setMidLineWidth(width)
        return line

    def make_label(self, *, text, font=None, parent=None):
        label = QtWidgets.QLabel(text, parent or self.centralwidget)
        if not font:
            font = QtGui.QFont()
            font.setFamily("Times New Roman")
            font.setPointSize(12)
            font.setBold(True)
        label.setFont(font)
        label.setAlignment(QtCore.Qt.AlignCenter)
        return label

    def make_date(self, *, date_par=None, status=None, font=None, parent=None):
        if parent:
            widget = QtWidgets.QDateEdit(parent)
        else:
            widget = QtWidgets.QDateEdit(self.centralwidget)
        if font:
            widget.setFont(font)
        if date_par:
            self.set_date(date_par, widget)
        widget.setStatusTip(status)
        return widget

    @staticmethod
    def set_date(date_par, widget):
        widget.setDate(QtCore.QDate(*date_par))

    def make_combobox(self, status=None, size=None, array=None, parent=None, font=None):
        combobox = QtWidgets.QComboBox(parent or self.centralwidget)
        if size:
            combobox.setMinimumSize(QtCore.QSize(*size))
        if status:
            combobox.setStatusTip(status)
        if array:
            self.fill_combo_box(array, combobox)
        if font:
            combobox.setFont(font)
        return combobox

    @staticmethod
    def fill_combo_box(array, widget):
        if array is None:
            return
        if type(widget) == CheckableComboBox:
            widget.fill_combo_box(array)
            return
        for i in array:
            widget.addItem(i)

    def create_checkable_combobox(self, parent=None, font=None):
        checkable_combobox = CheckableComboBox(parent or self.centralwidget)
        if font:
            checkable_combobox.setFont(font)
        return checkable_combobox

    def add_item(self, array, name, widget):
        if all(name.lower() != i.lower() for i in array):
            name = name.lower().capitalize()
            array.append(name)
            widget.addItem(name)
        else:
            self.send_error('Имя уже добавлено в список')

    def send_error(self, message):
        error_dialog = QtWidgets.QErrorMessage(self)
        error_dialog.setWindowModality(QtCore.Qt.WindowModal)
        error_dialog.showMessage(message)
        error_dialog.exec_()

    def make_checkbox(self, *, font=None, text=None, activate=False, command=None, parent=None):
        if parent:
            widget = QtWidgets.QCheckBox(parent)
        else:
            widget = QtWidgets.QCheckBox(self.gridLayoutWidget)
        if activate:
            widget.setChecked(True)
        if font:
            widget.setFont(font)
        if command:
            widget.clicked.connect(command)
        widget.setText(text)
        return widget

    def make_radio_button(self, *, text, command=None, parent=None, font=None):
        radio_button = QtWidgets.QRadioButton(text, parent or self.centralwidget)
        if font:
            radio_button.setFont(font)
        if command:
            radio_button.clicked.connect(command)
        return radio_button

    @staticmethod
    def get_items_in_list(list_widget):
        return [str(list_widget.item(i).text()) for i in range(list_widget.count())
                if list_widget.item(i).checkState()]


class CheckableComboBox(QtWidgets.QComboBox):
    def __init__(self, parent):
        super().__init__(parent)
        self.view().pressed.connect(self.handle_item_pressed)

    def handle_item_pressed(self, index):
        item = self.model().itemFromIndex(index)
        if item.checkState() == QtCore.Qt.Checked:
            item.setCheckState(QtCore.Qt.Unchecked)
        else:
            item.setCheckState(QtCore.Qt.Checked)

    def fill_combo_box(self, array):
        for index, element in enumerate(array):
            self.addItem(element)
            item = self.model().item(index, 0)
            item.setCheckState(QtCore.Qt.Unchecked)

    def item_checked(self, index):
        item = self.model().item(index, 0)
        return item.checkState() == QtCore.Qt.Checked

    def collect_checked_items(self):
        checked_items = []
        for index in range(self.count()):
            if self.item_checked(index):
                checked_items.append(self.model().item(index, 0).text())
        return checked_items


class ListWidget(QtWidgets.QListWidget):
    def __init__(self, parent):
        super().__init__(parent)

    def fill_list(self, *, draw_list,):
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(9)

        for file_path in draw_list:
            item = QtWidgets.QListWidgetItem()
            item.setText(file_path)
            item.setCheckState(QtCore.Qt.Checked)
            item.setFont(font)
            self.addItem(item)

    def get_items_text_data(self):
        items = self.get_selected_items()
        return [str(item.text()) for item in items]

    def get_selected_items(self):
        items = (self.item(index) for index in range(self.count()) if self.item(index).checkState())
        return items

    def get_not_selected_items(self):
        items = (self.item(index) for index in range(self.count()) if not self.item(index).checkState())
        return items


class MainListWidget(ListWidget):
    def __init__(self, parent):
        super().__init__(parent)
        self.itemDoubleClicked.connect(self.open_item)

    @staticmethod
    def open_item(item):
        path = item.text()
        os.system(fr'explorer "{os.path.normpath(os.path.dirname(path))}"')
        os.startfile(path)

    def select_all(self):
        items = self.get_not_selected_items()
        if items:
            for item in items:
                item.setCheckState(QtCore.Qt.Checked)

    def unselect_all(self):
        items = self.get_selected_items()
        if items:
            for item in items:
                item.setCheckState(QtCore.Qt.Unchecked)

    def move_item_down(self):
        current_row = self.currentRow()
        current_item = self.takeItem(current_row)
        self.insertItem(current_row + 1, current_item)
        self.setCurrentRow(current_row + 1)

    def move_item_up(self):
        current_row = self.currentRow()
        current_item = self.takeItem(current_row)
        self.insertItem(current_row - 1, current_item)
        self.setCurrentRow(current_row - 1)


class ExcludeFolderListWidget(ListWidget):
    def __init__(self, parent):
        super().__init__(parent)

    def add_folder(self):
        folder_name, ok = QtWidgets.QInputDialog.getText(self, "Дилог ввода текста", "Введите название папки")
        if ok:
            self.fill_list(draw_list=[folder_name])

    def remove_item(self):
        list_items = self.selectedItems()
        if not list_items:
            return
        for item in list_items:
            self.takeItem(self.row(item))
