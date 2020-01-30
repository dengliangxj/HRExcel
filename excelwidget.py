import os
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import QObject, pyqtSlot, pyqtSignal
from excel import Ui_Dialog
from excelloader import ExcelLoader
from excelsplit import ExcelSplit
from excelmerge import ExcelMerge


class ExcelWidget(QtWidgets.QDialog):

    def __init__(self):
        QtWidgets.QWidget.__init__(self)
        self.ui = Ui_Dialog()

        # 构造拆分表格的数据字典
        self.__splitList__ = []
        self.__mergeList__ = []

        # 初始化当前路径
        self.__lastFilePath__ = QtCore.QDir.currentPath()
        self.__mergeFileDir__ = self.__lastFilePath__

        # 构造Excel加载器
        self.__excelLoader__ = ExcelLoader(self)

        # 构造Excel拆分器
        self.__excelSplit__ = ExcelSplit(self)

        # 构造拆分的ListView列表的model
        self.__splitModel__ = QtCore.QStringListModel()

        # 构造拆分表格的文件名称
        self.__splitXlsFile__ = ''
        self.__splitSheetName__ = ''

        # 构造Excel合并器
        self.__excelMerge__ = ExcelMerge(self)

        # 初始化UI界面以及信号连接
        self.initUi()

    # 初始化UI界面
    def initUi(self):

        # 构造UI界面
        self.ui.setupUi(self)

        # 初始化ListView控件
        self.ui.listView.setModel(self.__splitModel__)
        self.ui.listView.setSpacing(10)
        self.ui.listView.setSelectionMode(QtWidgets.QAbstractItemView.MultiSelection)

        # 隐藏进度条的展示
        self.ui.pBarSplit.hide()
        self.ui.pBarMerge.hide()

        # 连接拆分和合并按钮的响应函数
        self.ui.btnSplitImport.clicked.connect(self.onSplitImport)
        self.ui.btnMergeImport.clicked.connect(self.onMergeImport)
        self.ui.btnSplit.clicked.connect(self.onSplitExecute)
        self.ui.btnMerge.clicked.connect(self.onMergeExecute)

    # 声明响应拆分导入槽函数
    @pyqtSlot()
    def onSplitImport(self):
        # 弹出选择对话框
        dlgTitle = '选择一个文件'
        fileFilter = 'Excel文件(*.xls *.xlsx)'
        chooseFilePath = QtWidgets.QFileDialog.getOpenFileName(None,dlgTitle,self.__lastFilePath__,fileFilter)

        print('chooseFilePath:' + chooseFilePath[0])

        # 如果路径为空，则直接返回
        if chooseFilePath[0] == '':
            print('chooseFilePath is empty!')
            return

        # 灰显所有的按钮
        self.enableBtns(False)

        # 连接完成信号
        self.__excelLoader__.signal_LoadFinished.connect(self.onSplitImportFinished)

        # 记录待拆分的XLS文件名称
        self.__splitXlsFile__ = chooseFilePath[0]

        # 更新最近的文件目录
        self.__lastFilePath__ = os.path.dirname(self.__splitXlsFile__)

        # 启动Excel文件的加载
        self.__excelLoader__.load(self.__splitXlsFile__)

    # 声明响应Excel文件加载完成的槽函数
    @pyqtSlot(list, str)
    def onSplitImportFinished(self, listData, sheetName):

        print('onSplitImportFinished sheetName: ', sheetName)

        # 断链信号
        self.__excelLoader__.signal_LoadFinished.disconnect()

        # 亮显按钮
        self.enableBtns(True)

        # 更新数据字典
        self.__splitList__ = listData
        self.__splitSheetName__ = sheetName

        # 提示加载完成或者加载失败
        tipMessage = '加载Excel文件失败'

        if listData:
            # 在界面上展示对应的拆分表头
            self.__splitModel__.setStringList(listData[0])
            tipMessage = '加载待拆分表格完成！'

        else:
            # 清空待记录的拆分表格文件
            self.__splitXlsFile__ = ''

        QtWidgets.QMessageBox.information(None, '提示', tipMessage)

    # 声明响应拆分命令槽函数
    @pyqtSlot()
    def onSplitExecute(self):
        print('onSplitExecute')

        # 有效响应：A. splitDict有数据  B. listView中有选中的条目
        if (self.__splitXlsFile__ == '') or (not self.__splitList__):
            # 提示没有加载Excel文件
            QtWidgets.QMessageBox.information(None, '提示', '请先加载待拆分的Excel文件！')
            return

        # 如果用户没有选择表头，则提示需要选择表头数据
        selectedIndexList = self.ui.listView.selectedIndexes()

        # 获取当前用户选择的字段
        if 0 == len(selectedIndexList):
            # 提示未选择数据
            QtWidgets.QMessageBox.information(None, '提示', '请先在右侧窗口选择拆分字段！')
            return

        # 根据选择的条目，获得当前的拆分字段
        headerStringList = self.__splitModel__.stringList()
        headerIndexes = []

        # 获取选择的表头Index集合
        for selectedIndex in selectedIndexList:
            headerIndexes.append(selectedIndex.row())

        # 灰显按钮
        self.enableBtns(False)

        # 执行表格的拆分
        self.__excelSplit__.signal_Progressed.connect(self.onSplitProgressed)
        self.__excelSplit__.finished.connect(self.onSplitFinished)

        self.__excelSplit__.split(self.__splitXlsFile__, self.__splitSheetName__, self.__splitList__, headerIndexes)

    # 声明响应拆分表格进度通知槽函数
    @pyqtSlot(int, int)
    def onSplitProgressed(self, curProgress, maxProgress):
        # 首先展示进度条
        if not self.ui.pBarSplit.isVisible():
            self.ui.pBarSplit.setVisible(True)

        # 设置进度条的最大和当前刻度
        self.ui.pBarSplit.setRange(0, maxProgress)
        self.ui.pBarSplit.setValue(curProgress)

    # 声明响应拆分表格完成槽函数
    @pyqtSlot()
    def onSplitFinished(self):
        # 断连信号
        self.__excelSplit__.signal_Progressed.disconnect()
        self.__excelSplit__.finished.disconnect()

        # 隐藏进度条
        self.ui.pBarSplit.setVisible(False)

        # 清理进度条
        self.ui.pBarSplit.reset()

        # 亮显按钮
        self.enableBtns(True)

        # 提示拆分完成
        QtWidgets.QMessageBox.information(None, '提示', '表格拆分完成！')

    # 声明响应合并导入槽函数
    @pyqtSlot()
    def onMergeImport(self):
        # 弹出目录选择框，供用户选择需要合并的目录
        dlgTitle = '选择合并目录'
        chooseFilePath = QtWidgets.QFileDialog.getExistingDirectory(None, dlgTitle, self.__lastFilePath__)

        if chooseFilePath == '':
            return

        # 更新最近的文件路径
        self.__lastFilePath__ = chooseFilePath
        self.__mergeFileDir__ = chooseFilePath

        # 在界面中显示文件路径
        self.ui.labelMergeDirPath.setText(chooseFilePath)

    # 声明响应合并导入槽函数
    @pyqtSlot()
    def onMergeExecute(self):
        # 首先判定是否有选择路径
        if self.__mergeFileDir__ == '':
            # 提示先选择待合并的路径
            QtWidgets.QMessageBox.information(None, '提示', '请先选择合并文件路径！')
            return

        # 获取目录下的所有Excel文件
        dirFiles = os.listdir(self.__mergeFileDir__)
        allExcelFileList = []

        # 设置为绝对路径
        for dirFileName in dirFiles:
            walkFileName = os.path.join(self.__mergeFileDir__, dirFileName)

            # 过滤非excel后缀文件以及合并的Excel文件
            if ('.xls' in walkFileName) and ('Merge' not in walkFileName):
                allExcelFileList.append(walkFileName)

        # 如果所有文件为空的话，提示用户，没有发现XLS文件
        if not allExcelFileList:
            QtWidgets.QMessageBox.information(None, '提示', '合并文件目录未发现Excel文件！')
            return

        # 灰显按钮
        self.enableBtns(False)

        # 执行所有的Excel的合并
        self.__excelMerge__.signal_Progressed.connect(self.onMergeProgressed)
        self.__excelMerge__.signal_Finished.connect(self.onMergeFinished)

        # 执行Excel的合并操作
        self.__excelMerge__.merge(allExcelFileList)

    # 声明响应合并进度的槽函数
    @pyqtSlot(int, int)
    def onMergeProgressed(self, curProgress, maxProgress):
        # 首先展示进度条
        if not self.ui.pBarMerge.isVisible():
            self.ui.pBarMerge.setVisible(True)

        # 设置进度条的最大和当前刻度
        self.ui.pBarMerge.setRange(0, maxProgress)
        self.ui.pBarMerge.setValue(curProgress)

    # 声明响应合并表格完成槽函数
    @pyqtSlot(bool, str)
    def onMergeFinished(self, success, errStrings):
        # 断连信号
        self.__excelMerge__.signal_Progressed.disconnect()
        self.__excelMerge__.signal_Finished.disconnect()

        # 隐藏进度条
        self.ui.pBarMerge.setVisible(False)

        # 清理进度条
        self.ui.pBarMerge.reset()

        # 亮显按钮
        self.enableBtns(True)

        if success:
            # 提示拆分完成
            QtWidgets.QMessageBox.information(None, '提示', '表格合并完成！')
        else:
            # 提示故障字符串
            QtWidgets.QMessageBox.information(None, '提示', errStrings)

    # 灰显所有的按钮
    def enableBtns(self, enable):
        self.ui.btnSplit.setEnabled(enable)
        self.ui.btnSplitImport.setEnabled(enable)
        self.ui.btnMerge.setEnabled(enable)
        self.ui.btnMergeImport.setEnabled(enable)
        self.ui.listView.setEnabled(enable)
