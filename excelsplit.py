import sys
import os
import time
import xlsxwriter
from PyQt5 import QtCore
from PyQt5.QtCore import QObject, pyqtSlot, pyqtSignal
from exceldatawriter import ExcelDataWriter

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
        self.__dataList__ = []
        self.__headerIndexes__ = []

    # 定义Excel拆分方法
    def split(self, excelFilePath, sheetName, dataList, headerIndexes):
        # 保存当前的文件路径、数据字典、数据拆分字段列表
        self.__xlsPath__ = excelFilePath
        self.__sheetName__ = sheetName
        self.__dataList__ = dataList
        self.__headerIndexes__ = headerIndexes

        # 启动线程
        self.start()

    # 重载线程的run函数
    def run(self):
        # 输出当前需要拆分的Keys
        print('start split:', self.__xlsPath__)

        # 获取字段的行数
        dictRows = len(self.__dataList__)

        # 遍历所有的数据行，获取按照dataFilter对应列的键值对
        dataValueKeyList = []

        # 第一行固定为表头
        for rowNo in range(1, dictRows):
            # 构造一行数据的键值对
            keyValuePair = []

            # 添加键值
            for filterIndex in self.__headerIndexes__:
                keyValuePair.append(self.__dataList__[rowNo][filterIndex])

            # 保存非重复的键值对
            if keyValuePair not in dataValueKeyList:
                dataValueKeyList.append(keyValuePair)

        # 计算需要拆分的Excel总数
        totalSplitNum = len(dataValueKeyList)
        print('dataValueKeyList size: ', totalSplitNum)

        # 获取当前系统时间
        dateTimePrefix = time.strftime("Split_%Y-%m-%d_%H-%M-%S", time.localtime())

        # 执行所有Excel的拆分
        for i in range(totalSplitNum):
            print('start split excelNo: ', i+1)
            # 通知进度
            self.signal_Progressed.emit(i+1, totalSplitNum)

            # 执行对应键值的Excel拆分
            self.__executeOneExcelSplit__(dataValueKeyList[i], self.__dataList__, dateTimePrefix)

    # 根据键值和原Excel的文件名称，构造新的Excel文件名称
    @staticmethod
    def __buildSplitExcelPath__(keyPairs, excelPath, dateTimePrefix):

        # 获取单纯的文件名称
        baseFileName = os.path.basename(excelPath).strip('.xlsx').strip('.xls')

        # 拼接当前的Excel文件名称
        for keyString in keyPairs:
            baseFileName += '_'
            baseFileName += keyString

        # 添加后缀
        baseFileName += '.xlsx'

        # 获取当前文件的路径
        filePathDir = os.path.dirname(excelPath)

        # 添加日期字样作为前缀
        filePathDir += '/'
        filePathDir += dateTimePrefix
        filePathDir += '/'

        # 如果不存在该目录的话，则创建该目录
        if not os.path.exists(filePathDir):
            os.makedirs(filePathDir)

        splitFilePath = os.path.join(filePathDir + baseFileName)

        # 重新保存文件名称
        return splitFilePath

    # 判定数据行是否包含键值
    @staticmethod
    def __containKeysByRowData__(keyPairs, rowDataList):
        for keyItem in keyPairs:
            # 判定该项是否在rowDataList中
            if keyItem not in rowDataList:
                return False

        return True

    # 执行所有表格的拆分
    def __executeOneExcelSplit__(self, keyPairs, dataList, dateTimePrefix):
        # 构造新的Excel文件的名称
        splitXlsName = self.__buildSplitExcelPath__(keyPairs, self.__xlsPath__, dateTimePrefix)

        # 定义拆分后的数据表
        splitDataList = []

        # 根据Filter从全集的Dict提取出对应的部分Dict
        splitDataList.append(dataList[0])

        # 添加对应含有键值对的行
        for rowDataList in dataList:
            if self.__containKeysByRowData__(keyPairs, rowDataList):
                splitDataList.append(rowDataList)

        # 写入到Excel中
        excelDataWriter = ExcelDataWriter()
        excelDataWriter.doWrite(splitXlsName, self.__sheetName__, splitDataList)
