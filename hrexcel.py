# -*- coding: utf-8 -*-

import sys
from PyQt5 import QtWidgets, QtCore, QtGui
from excelwidget import ExcelWidget


def main():
    app = QtWidgets.QApplication(sys.argv)
    excelWidget = ExcelWidget()
    excelWidget.exec()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()


