from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import QObject, pyqtSlot, pyqtSignal
from excel import Ui_Dialog
from excelloader import ExcelLoader
from excelsplit import ExcelSplit


class ExcelWidget(QtWidgets.QDialog):

    def __init__(self):
        QtWidgets.QWidget.__init__(self)
        self.ui = Ui_Dialog()

        # 构造拆分表格的数据字典
        self.__splitDict__ = {}
        self.__mergeDict__ = {}

        # 初始化当前路径
        self.__lastFilePath__ = QtCore.QDir.currentPath()

        # 构造Excel加载器
        self.__excelLoader__ = ExcelLoader(self)

        # 构造Excel拆分器
        self.__excelSplit__ = ExcelSplit(self)

        # 构造拆分的ListView列表的model
        self.__splitModel__ = QtCore.QStringListModel()

        # 构造拆分表格的文件名称
        self.__splitXlsFile__ = ''
        self.__splitSheetName__ = ''

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

        # 更新最近的文件目录 TBD

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

        # 启动Excel文件的加载
        self.__excelLoader__.load(self.__splitXlsFile__)

    # 声明响应Excel文件加载完成的槽函数
    @pyqtSlot(dict, str)
    def onSplitImportFinished(self, dataDict, sheetName):

        print('onSplitImportFinished sheetName: ', sheetName)

        # 断链信号
        self.__excelLoader__.signal_LoadFinished.disconnect()

        # 亮显按钮
        self.enableBtns(True)

        # 更新数据字典
        self.__splitDict__ = dataDict
        self.__splitSheetName__ = sheetName

        # 在界面上展示对应的拆分表头
        self.__splitModel__.setStringList(dataDict.keys())

        # 提示加载完成或者加载失败
        tipMessage = '加载Excel文件失败'

        if dataDict:
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
        if (self.__splitXlsFile__ == '') or (not self.__splitDict__):
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
        dataFilter = []

        for selectedIndex in selectedIndexList:
            rowIndex = selectedIndex.row()
            dataFilter.append(headerStringList[rowIndex])

        # 灰显按钮
        self.enableBtns(False)

        # 执行表格的拆分
        self.__excelSplit__.signal_Progressed.connect(self.onSplitProgressed)
        self.__excelSplit__.finished.connect(self.onSplitFinished)

        self.__excelSplit__.split(self.__splitXlsFile__, self.__splitSheetName__, self.__splitDict__, dataFilter)

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
        pass

    # 声明响应合并导入槽函数
    @pyqtSlot()
    def onMergeExecute(self):
        pass

    # 灰显所有的按钮
    def enableBtns(self, enable):
        self.ui.btnSplit.setEnabled(enable)
        self.ui.btnSplitImport.setEnabled(enable)
        self.ui.btnMerge.setEnabled(enable)
        self.ui.btnMergeImport.setEnabled(enable)
        self.ui.listView.setEnabled(enable)
