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

    def make_text_edit(self, *, text=None, status=None, parent=None,
                        placeholder=None, size_policy=None, font=None):
        text_edit = QtWidgets.QTextEdit(parent or self.centralwidget)
        text_edit.setStatusTip(status)
        text_edit.setPlainText(text)
        font.setPointSize(11)
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

    def make_date(self, *, date_par, status=None, font=None, parent=None):
        if parent:
            widget = QtWidgets.QDateEdit(parent)
        else:
            widget = QtWidgets.QDateEdit(self.centralwidget)
        if font:
            widget.setFont(font)
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
        for i in array:
            widget.addItem(i)

    def add_item(self, array, name, widget):
        if all(name.lower() != i.lower() for i in array):
            name = name.lower().capitalize()
            array.append(name)
            widget.addItem(name)
        else:
            self.send_error('Имя уже добавлено в список')

    def make_menu(self):
        menu = self.menuBar()
        for name, items in self.menu_list:
            pulldown=menu.addMenu(name)
            self.addMenuItems(pulldown, items)

    def addMenuItems(self, pulldown, items):
        for item in items:
            command = QtWidgets.QAction(item[0],self)
            command.triggered.connect(item[1])
            pulldown.addAction(command)

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

    def fill_list(self, *, draw_list, widget_list=None):
        if widget_list != None:
            widget = widget_list
        else:
            widget = self.listWidget
        for file in draw_list:
            item = QtWidgets.QListWidgetItem()
            item.setText(file)
            item.setCheckState(QtCore.Qt.Checked)
            widget.addItem(item)

    def remove_item(self, *, widget_list=None):
        if widget_list.baseSize:
            widget = widget_list
        else:
            widget = self.listWidget
        list_items = widget.selectedItems()
        if not list_items:
            return
        for item in list_items:
            widget.takeItem(widget.row(item))

    def get_items_in_list(self, list_widget):
        return [str(list_widget.item(i).text()) for i in range(list_widget.count())
                if list_widget.item(i).checkState()]

