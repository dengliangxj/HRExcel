import sys
import time
import xlrd
from PyQt5 import QtCore
from PyQt5.QtCore import QObject, pyqtSlot, pyqtSignal


class ExcelLoader(QtCore.QThread):

    # 定义启动信号
    signal_LoadStarted = pyqtSignal(int)

    # 定义结束信号
    signal_LoadFinished = pyqtSignal(dict, str)

    def __init__(self, parent):
        QtCore.QThread.__init__(self, parent)
        self.__loadPath__ = ''

    # 异步加载数据库
    def load(self, excelPath):
        # 缓存当前的加载路径
        self.__loadPath__ = excelPath

        # 直接启动线程
        self.start()

    # 执行数据所有字段的加载
    def loadFieldsFromSheet(self, sheet):
        # 获取sheet中的所有字段大小
        print(sheet.name, sheet.nrows, sheet.ncols)

        # 定义一个字典
        dict = {}
        headerRow = 0
        headerStrings = []

        # 搜索第一个字段开头为“工号”，即为表头项目
        for theRowNo in range(sheet.nrows):
            # 读取每一行的数据
            valueStrings = sheet.row_values(theRowNo)

            # 判定是否含有“工号”
            if '工号' in valueStrings:
                # 记录表头行的位置
                print('found header rowNo: ', theRowNo)
                headerRow = theRowNo

                # 记录表头(非空字段)
                for headerKeyString in valueStrings:
                    if headerKeyString.strip() != '':
                        headerStrings.append(headerKeyString)

                break

        # 开始遍历表格，记录所有的信息
        for theColIndex in range(len(headerStrings)):
            # 如果表头非空的，则读取所有的信息到dict中
            if headerStrings[theColIndex] != '':
                # 读取表格中所有的对应行的数据
                cellDataList = []
                for theRowNo in range(headerRow+1, sheet.nrows):
                    cell_string = sheet.cell_value(theRowNo, theColIndex)
                    cellDataList.append(cell_string)

                # 保存到字典中
                dict[headerStrings[theColIndex]] = cellDataList

        return dict

    # 重载线程的run函数
    def run(self):
        # 通知外部开始启动
        self.signal_LoadStarted.emit(5)

        print('start load excel file : ' + self.__loadPath__)

        # 执行Excel的加载
        dict = {}

        # 读取Excel文件
        excelFile = xlrd.open_workbook(self.__loadPath__)

        # 获取所有的sheet名字
        sheetNameList = excelFile.sheet_names()

        # 如果一个sheet都没有则返回空字典
        if not sheetNameList:
            # 直接发送空字典
            self.signal_LoadFinished.emit(dict)

        # 记录第一个sheet的名称
        sheetName = sheetNameList[0]

        # 读取第一个worksheet
        print('start process first worksheet: ' + sheetName)

        # 加载所有的表格内容
        sheet = excelFile.sheet_by_name(sheetName)

        if not sheet:
            print('load sheet failed: ' + sheetName)
            self.signal_LoadFinished.emit(dict)

        # 从sheet中读取所有的字段
        dict = self.loadFieldsFromSheet(sheet)

        print('finish load excel file : ' + self.__loadPath__)

        # 通过信号将字典发送到UI界面
        self.signal_LoadFinished.emit(dict, sheetName)
