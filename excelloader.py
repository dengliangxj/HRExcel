import sys
import time
import xlrd
from PyQt5 import QtCore
from PyQt5.QtCore import QObject, pyqtSlot, pyqtSignal


class ExcelLoader(QtCore.QThread):

    # 定义启动信号
    signal_LoadStarted = pyqtSignal(int)

    # 定义结束信号
    signal_LoadFinished = pyqtSignal(list, str)

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
    @staticmethod
    def loadFieldsFromSheet(sheetFileName):
        # 读取Excel文件
        excelFile = xlrd.open_workbook(sheetFileName)

        # 获取所有的sheet名字
        sheetNameList = excelFile.sheet_names()

        # 定义Excel的sheetName
        sheetName = ''

        # 定义一个列表集合
        listData = []

        # 如果一个sheet都没有则返回空字典
        if not sheetNameList:
            # 直接发送空字典
            return listData, sheetName

        # 记录第一个sheet的名称
        sheetName = sheetNameList[0]

        # 加载所有的表格内容
        sheet = excelFile.sheet_by_name(sheetName)

        # 打印数据表中的信息
        print('start load excel file : ', sheetFileName, sheet.name, sheet.nrows, sheet.ncols)

        headerRow = 0
        headerStrings = []
        dataColumnIndexes = []

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
                for i in range(len(valueStrings)):
                    if valueStrings[i].strip() != '':
                        headerStrings.append(valueStrings[i])
                        dataColumnIndexes.append(i)
                break

        # 第一行记录表头
        listData.append(headerStrings)

        # 开始遍历表格，记录所有的信息
        for rowIndex in range(headerRow+1, sheet.nrows):

            # 构造一行的数据
            cellDataList = []

            # 遍历每一行保存数据
            for colIndex in dataColumnIndexes:
                cellString = sheet.cell_value(rowIndex, colIndex)
                cellDataList.append(cellString)

            # 保存列表集合
            listData.append(cellDataList)

        print('finish load excel file : ', sheetFileName)

        return listData, sheetName

    # 重载线程的run函数
    def run(self):
        # 通知外部开始启动
        self.signal_LoadStarted.emit(5)

        # 从sheet中读取所有的字段
        loadResultList = self.loadFieldsFromSheet(self.__loadPath__)

        # 通过信号将字典发送到UI界面
        self.signal_LoadFinished.emit(loadResultList[0], loadResultList[1])
