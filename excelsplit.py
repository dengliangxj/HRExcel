import sys
import time
import xlsxwriter
from PyQt5 import QtCore
from PyQt5.QtCore import QObject, pyqtSlot, pyqtSignal

#########################################################
# Excel文件拆分类：
# 用于将字典转化为excel进行导出,通知外部拆分进度
#########################################################


class ExcelSplit(QtCore.QThread):

    # 定义拆分进度通知的信号
    signal_Progressed = pyqtSignal(int, int)

    def __init__(self, parent):
        QtCore.QThread.__init__(self, parent)
        self.__xlsPath__ = ''
        self.__sheetName__ = ''
        self.__dataDict__ = {}
        self.__dataFilter__ = []

    # 定义Excel拆分方法
    def split(self, excelFilePath, sheetName, dataDict, dataFilter):
        # 保存当前的文件路径、数据字典、数据拆分字段列表
        self.__xlsPath__ = excelFilePath
        self.__sheetName__ = sheetName
        self.__dataDict__ = dataDict
        self.__dataFilter__ = dataFilter

        # 启动线程
        self.start()

    # 重载线程的run函数
    def run(self):
        # 输出当前需要拆分的Keys
        print('start split by:', self.__dataFilter__)

        # 获取字典的行数
        dictRows = len(self.__dataDict__[self.__dataFilter__[0]])

        # 遍历所有的数据行，获取按照dataFilter对应列的键值对
        dataValueKeyList = []

        for rowNo in range(dictRows):
            # 构造一行数据的键值对
            keyValuePair = []

            # 添加键值
            for dictKey in self.__dataFilter__:
                keyValuePair.append(self.__dataDict__[dictKey][rowNo])

            # 保存非重复的键值对
            if keyValuePair not in dataValueKeyList:
                dataValueKeyList.append(keyValuePair)

        # 计算需要拆分的Excel总数
        totalSplitNum = len(dataValueKeyList)
        print('dataValueKeyList size: ', totalSplitNum)

        # 执行所有Excel的拆分
        for i in range(totalSplitNum):
            print('start split excelNo: ', i+1)
            # 通知进度
            self.signal_Progressed.emit(i+1, totalSplitNum)

            # 执行对应键值的Excel拆分
            self.__executeOneExcelSplit__(dataValueKeyList[i], self.__dataDict__)

    # 根据键值和原Excel的文件名称，构造新的Excel文件名称
    @staticmethod
    def __buildSplitExcelPath__(keyPairs, excelPath):
        # 去掉文件后缀
        splitXlsName = excelPath.strip('.xlsx').strip('.xls')

        # 拼接当前的Excel文件名称
        for keyString in keyPairs:
            splitXlsName += '_'
            splitXlsName += keyString

        # 添加后缀
        splitXlsName += '.xlsx'

        # 重新保存文件名称
        return splitXlsName

    # 执行所有表格的拆分
    def __executeOneExcelSplit__(self, keyPairs, dataDict):
        # 构造新的Excel文件的名称
        splitXlsName = self.__buildSplitExcelPath__(keyPairs, self.__xlsPath__)

        print('begin split xlsx file: ', splitXlsName)

        # 创建Excel文档
        workBook = xlsxwriter.Workbook(splitXlsName)
        workSheet = workBook.add_worksheet(self.__sheetName__)

        # 写入表头
        headerDataList = dataDict.keys()
        self.__writeXlsHeader__(workBook, workSheet, headerDataList)

        # 从dict中筛选出满足Key要求的字段，并写入Excel
        self.__writeXlsDataArea__(workBook, workSheet, keyPairs, headerDataList, dataDict)

        # 关闭Excel文件
        workBook.close()

        print('finish split xlsx file: ', splitXlsName)

    # 写入表头
    @staticmethod
    def __writeXlsHeader__(workBook, workSheet, headerDataList):
        # 构造表头的格式
        headerStyle = workBook.add_format({
            'font_size': 12,  # 字体大小
            'bold': True,  # 是否粗体
            'bg_color': '#92CDDC',  # 表格背景颜色
            'font_color': '#000000',  # 字体颜色
            'align': 'center',  # 居中对齐

            # 后面参数是线条宽度
            'top': 2,    # 上边框
            'left': 2,   # 左边框
            'right': 2,  # 右边框
            'bottom': 2  # 底边框
        })

        # 设置每列宽度
        workSheet.set_column(0, len(headerDataList) - 1, 15)

        # 设置表头高度30
        workSheet.set_row(0, 30)

        # 写入表头数据
        workSheet.write_row(0, 0, headerDataList, headerStyle)

    # 判定数据行是否包含键值
    @staticmethod
    def __containKeysByRowData__(keyPairs, rowDataList):
        for keyItem in keyPairs:
            # 判定该项是否在rowDataList中
            if keyItem not in rowDataList:
                return False

        return True

    # 写入数据区
    def __writeXlsDataArea__(self, workBook, workSheet, keyPairs, headerDataList, dataDict):
        # 构造数据区的格式
        dataAreaStyle = workBook.add_format({
            'font_size': 10,  # 字体大小
            'font_color': '#000000',  # 字体颜色
            'align': 'center',  # 居中对齐

            # 后面参数是线条宽度
            'top': 1,    # 上边框
            'left': 1,   # 左边框
            'right': 1,  # 右边框
            'bottom': 1  # 底边框
        })

        # 起始行号为1
        rows = 1

        # 开始写满足键值的行
        dictRows = len(self.__dataDict__[self.__dataFilter__[0]])

        # 只要满足该行的数据包含所有的键值，则写入Excel中
        for i in range(dictRows):

            # 构造每行的数据组
            rowDataList = []

            # 添加每行的数据
            for headerKey in headerDataList:
                rowDataList.append(dataDict[headerKey][i])

            # 判定该行是否含有键值
            if self.__containKeysByRowData__(keyPairs, rowDataList):
                # 设置行高
                workSheet.set_row(rows, 20)

                # 将该行写入到Excel
                workSheet.write_row(rows, 0, rowDataList)
                rows += 1
