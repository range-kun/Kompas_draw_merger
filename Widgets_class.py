from PyQt5 import QtCore, QtGui, QtWidgets

class MakeWidgets(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        QtWidgets.QMainWindow.__init__(self, parent=None)

    def make_button(self, *, text, command=None, size=None, font=None):
        button = QtWidgets.QPushButton(self.centralwidget)
        if size:
            button.setMaximumSize(QtCore.QSize(*size))
        if command:
            button.clicked.connect(command)
        if font:
            button.setFont(font)
        button.setText(text)
        return button

    def make_line_edit(self, text=None, status=None, parent=None):
        line = QtWidgets.QLineEdit(parent)
        line.setStatusTip(status)
        line.setPlaceholderText(text)
        return line

    @staticmethod
    def make_plain_text(*, text=None, status=None, parent=None):
        plain_text = QtWidgets.QPlainTextEdit(parent)
        plain_text.setStatusTip(status)
        plain_text.setPlainText(text)
        return plain_text

    def make_line(self, orientation,width=None):
        line = QtWidgets.QFrame(self.centralwidget)
        if orientation == 'horizontal':
            line.setFrameShape(QtWidgets.QFrame.HLine)
        elif orientation == 'vertical':
            line.setFrameShape(QtWidgets.QFrame.VLine)
        if width:
            line.setLineWidth(width)
            line.setMidLineWidth(width)
        return line

    def make_label(self, *, text, font=None):
        label = QtWidgets.QLabel(text, self.centralwidget)
        if not font:
            font = QtGui.QFont()
            font.setFamily("Times New Roman")
            font.setPointSize(12)
            font.setBold(True)
        label.setFont(font)
        label.setAlignment(QtCore.Qt.AlignCenter)
        return label

    def make_date(self, *, date_par, status=None, font=None):
        widget = QtWidgets.QDateEdit(self.centralwidget)
        if font:
            widget.setFont(font)
        self.set_date(date_par, widget)
        widget.setStatusTip(status)
        return widget

    @staticmethod
    def set_date(date_par, widget):
        widget.setDate(QtCore.QDate(*date_par))

    def make_combobox(self,status=None,size=None,list=None):
        combobox = QtWidgets.QComboBox(self.centralwidget)
        if size:
            combobox.setMinimumSize(QtCore.QSize(*size))
        if status:
            combobox.setStatusTip(status)
        if list:
            self.fill_combo_box(list, combobox)
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
            self.error('Имя уже добавлено в список')

    def make_menu(self):
        menu=self.menuBar()
        for name, items in self.menu_list:
            pulldown=menu.addMenu(name)
            self.addMenuItems(pulldown, items)

    def addMenuItems(self, pulldown, items):
        for item in items:
            command=QtWidgets.QAction(item[0],self)
            command.triggered.connect(item[1])
            pulldown.addAction(command)

    def error(self, message):
        self.error_dialog = QtWidgets.QErrorMessage()
        self.error_dialog.showMessage(message)

    def make_checkbox(self, *, font=None, text=None, activate=False, command=None):
        widget = QtWidgets.QCheckBox(self.gridLayoutWidget)
        if activate:
            widget.setChecked(True)
        if font:
            widget.setFont(font)
        if command:
            widget.clicked.connect(command)
        widget.setText(text)
        return widget

    def fill_list(self, draw_list):
        for file in draw_list:
            item = QtWidgets.QListWidgetItem()
            item.setText(file)
            item.setCheckState(QtCore.Qt.Checked)
            self.listWidget.addItem(item)
        self.pushButton_5.setEnabled(True)

