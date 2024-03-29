from __future__ import annotations

import sys

from PyQt5 import QtWidgets

from programm.kompas_api import CoreKompass
from programm.main import except_hook
from programm.main import KompasMerger


if __name__ == "__main__":
    _core_kompas_api = CoreKompass()
    app = QtWidgets.QApplication(sys.argv)
    app.aboutToQuit.connect(_core_kompas_api.exit_kompas)
    app.setStyle("Fusion")

    merger = KompasMerger(_core_kompas_api)
    merger.mw.show()
    sys.excepthook = except_hook
    sys.exit(app.exec_())
