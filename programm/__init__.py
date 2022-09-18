import sys

from PyQt5 import QtWidgets

from programm.main import UiMerger, except_hook
from programm.kompas_api import CoreKompass


if __name__ == "__main__":
    _core_kompas_api = CoreKompass()
    app = QtWidgets.QApplication(sys.argv)
    app.aboutToQuit.connect(_core_kompas_api.exit_kompas)
    app.setStyle('Fusion')

    merger = UiMerger(_core_kompas_api)
    merger.show()
    sys.excepthook = except_hook
    sys.exit(app.exec_())
