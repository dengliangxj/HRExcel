# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'excel.ui'
#
# Created by: PyQt5 UI code generator 5.14.1
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(1003, 768)
        self.groupBox = QtWidgets.QGroupBox(Dialog)
        self.groupBox.setGeometry(QtCore.QRect(30, 40, 511, 251))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(14)
        self.groupBox.setFont(font)
        self.groupBox.setObjectName("groupBox")
        self.btnSplit = QtWidgets.QPushButton(self.groupBox)
        self.btnSplit.setGeometry(QtCore.QRect(290, 60, 200, 70))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(14)
        self.btnSplit.setFont(font)
        self.btnSplit.setObjectName("btnSplit")
        self.pBarSplit = QtWidgets.QProgressBar(self.groupBox)
        self.pBarSplit.setEnabled(True)
        self.pBarSplit.setGeometry(QtCore.QRect(20, 170, 471, 28))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(12)
        self.pBarSplit.setFont(font)
        self.pBarSplit.setProperty("value", 24)
        self.pBarSplit.setAlignment(QtCore.Qt.AlignCenter)
        self.pBarSplit.setTextDirection(QtWidgets.QProgressBar.TopToBottom)
        self.pBarSplit.setObjectName("pBarSplit")
        self.btnSplitImport = QtWidgets.QPushButton(self.groupBox)
        self.btnSplitImport.setGeometry(QtCore.QRect(20, 60, 200, 70))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(14)
        self.btnSplitImport.setFont(font)
        self.btnSplitImport.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))
        self.btnSplitImport.setObjectName("btnSplitImport")
        self.groupBox_2 = QtWidgets.QGroupBox(Dialog)
        self.groupBox_2.setGeometry(QtCore.QRect(30, 340, 511, 391))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(14)
        self.groupBox_2.setFont(font)
        self.groupBox_2.setObjectName("groupBox_2")
        self.label_3 = QtWidgets.QLabel(self.groupBox_2)
        self.label_3.setGeometry(QtCore.QRect(20, 50, 441, 61))
        font = QtGui.QFont()
        font.setFamily("黑体")
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.btnMerge = QtWidgets.QPushButton(self.groupBox_2)
        self.btnMerge.setGeometry(QtCore.QRect(290, 130, 200, 70))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(14)
        self.btnMerge.setFont(font)
        self.btnMerge.setObjectName("btnMerge")
        self.btnMergeImport = QtWidgets.QPushButton(self.groupBox_2)
        self.btnMergeImport.setGeometry(QtCore.QRect(20, 130, 200, 70))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(14)
        self.btnMergeImport.setFont(font)
        self.btnMergeImport.setCursor(QtGui.QCursor(QtCore.Qt.OpenHandCursor))
        self.btnMergeImport.setObjectName("btnMergeImport")
        self.pBarMerge = QtWidgets.QProgressBar(self.groupBox_2)
        self.pBarMerge.setEnabled(True)
        self.pBarMerge.setGeometry(QtCore.QRect(20, 340, 471, 28))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(12)
        self.pBarMerge.setFont(font)
        self.pBarMerge.setProperty("value", 24)
        self.pBarMerge.setAlignment(QtCore.Qt.AlignCenter)
        self.pBarMerge.setTextDirection(QtWidgets.QProgressBar.TopToBottom)
        self.pBarMerge.setObjectName("pBarMerge")
        self.labelMergeDirPath = QtWidgets.QLabel(self.groupBox_2)
        self.labelMergeDirPath.setGeometry(QtCore.QRect(120, 220, 371, 101))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(12)
        self.labelMergeDirPath.setFont(font)
        self.labelMergeDirPath.setText("")
        self.labelMergeDirPath.setWordWrap(True)
        self.labelMergeDirPath.setObjectName("labelMergeDirPath")
        self.label = QtWidgets.QLabel(self.groupBox_2)
        self.label.setGeometry(QtCore.QRect(20, 220, 91, 101))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.groupBox_3 = QtWidgets.QGroupBox(Dialog)
        self.groupBox_3.setGeometry(QtCore.QRect(570, 40, 401, 691))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(14)
        self.groupBox_3.setFont(font)
        self.groupBox_3.setObjectName("groupBox_3")
        self.label_2 = QtWidgets.QLabel(self.groupBox_3)
        self.label_2.setGeometry(QtCore.QRect(30, 30, 351, 111))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(6)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.listView = QtWidgets.QListView(self.groupBox_3)
        self.listView.setGeometry(QtCore.QRect(30, 150, 341, 521))
        font = QtGui.QFont()
        font.setFamily("黑体")
        font.setPointSize(14)
        self.listView.setFont(font)
        self.listView.setFocusPolicy(QtCore.Qt.NoFocus)
        self.listView.setContextMenuPolicy(QtCore.Qt.NoContextMenu)
        self.listView.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.listView.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.listView.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.listView.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.listView.setDragDropMode(QtWidgets.QAbstractItemView.DragOnly)
        self.listView.setDefaultDropAction(QtCore.Qt.IgnoreAction)
        self.listView.setAlternatingRowColors(False)
        self.listView.setSelectionMode(QtWidgets.QAbstractItemView.MultiSelection)
        self.listView.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.listView.setTextElideMode(QtCore.Qt.ElideRight)
        self.listView.setMovement(QtWidgets.QListView.Free)
        self.listView.setProperty("isWrapping", False)
        self.listView.setResizeMode(QtWidgets.QListView.Fixed)
        self.listView.setUniformItemSizes(False)
        self.listView.setWordWrap(True)
        self.listView.setSelectionRectVisible(True)
        self.listView.setItemAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.listView.setObjectName("listView")

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Mindray表格处理平台"))
        self.groupBox.setTitle(_translate("Dialog", "表格拆分"))
        self.btnSplit.setText(_translate("Dialog", "拆分表格"))
        self.pBarSplit.setFormat(_translate("Dialog", "%v/%m"))
        self.btnSplitImport.setText(_translate("Dialog", "选择拆分表格文件"))
        self.groupBox_2.setTitle(_translate("Dialog", "表格合并"))
        self.label_3.setText(_translate("Dialog", "<html><head/><body><p align=\"justify\"><span style=\" font-size:12pt; color:#ff0000;\">请注意！！！</span></p><p align=\"justify\"><span style=\" font-size:12pt; color:#ff0000;\">必须确保合并表格的表头完全一致。</span></p></body></html>"))
        self.btnMerge.setText(_translate("Dialog", "合并表格"))
        self.btnMergeImport.setText(_translate("Dialog", "选择合并表格路径"))
        self.pBarMerge.setFormat(_translate("Dialog", "%v/%m"))
        self.label.setText(_translate("Dialog", "合并路径："))
        self.groupBox_3.setTitle(_translate("Dialog", "拆分条件选择"))
        self.label_2.setText(_translate("Dialog", "<html><head/><body><p align=\"justify\"><span style=\" font-size:12pt;\">选择拆分表格文件完成后自动展示所有</span></p><p align=\"justify\"><span style=\" font-size:12pt;\">表头项，选择拆分的条件后便按照选项</span></p><p align=\"justify\"><span style=\" font-size:12pt;\">进行拆分，支持多选。</span></p></body></html>"))
