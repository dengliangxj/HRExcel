import os
import sys
import time
from PyQt5 import QtCore
from PyQt5.QtCore import QObject, pyqtSlot, pyqtSignal
from excelloader import ExcelLoader
from exceldatawriter import ExcelDataWriter


#########################################################
# Excel文件合并类：
# 用于将读取所有的Excel文件，并将其合并为一个文件
#########################################################

class ExcelMerge(QtCore.QThread):

    # 定义合并进度通知的信号
    signal_Progressed = pyqtSignal(int, int)

    # 定义完成的信号
    signal_Finished = pyqtSignal(bool, str)

    def __init__(self, parent):
        QtCore.QThread.__init__(self, parent)
        self.__mergeFileList__ = []

    # 执行Excel文件的合并
    def merge(self, excelFileList):
        # 记录需要解析的Excel文件列表
        self.__mergeFileList__ = excelFileList

        # 直接启动线程
        self.start()

    # 重载线程的run函数
    def run(self):
        # 构造内容集合
        allDataContentList = []

        # 表头集合
        headerContents = []

        # 合并表格的sheetName
        mergeSheetName = ''

        # 构造Excel加载类
        excelLoader = ExcelLoader(self)

        # 构造Excel写入类
        excelWriter = ExcelDataWriter()

        # 依次加载所有的文件
        totalFileNum = len(self.__mergeFileList__)

        for i in range(totalFileNum):
            # 获取Excel文件名
            excelFileName = self.__mergeFileList__[i]

            # 从Excel文件中，加载信息
            loadResultList = excelLoader.loadFieldsFromSheet(excelFileName)

            # 进行结果和Sheet的赋值
            dataContentList = loadResultList[0]
            mergeSheetName = loadResultList[1]

            # 如果加载失败，则直接通知结束失败
            if not dataContentList:
                errString = 'Excel文件解析失败： ' + excelFileName
                self.signal_Finished.emit(False, errString)
                return

            # 记录表头
            if not headerContents:
                headerContents = dataContentList[0]
                allDataContentList.append(headerContents)

            # 判定当前的Excel文件头是否与表头一致
            if headerContents != dataContentList[0]:
                errString = 'Excel文件表头信息不统一： ' + excelFileName
                self.signal_Finished.emit(False, errString)
                return

            # 合并数据区域到统一的列表集合中(数据区从第2行开始)
            for dataIndex in range(1, len(dataContentList)):
                allDataContentList.append(dataContentList[dataIndex])

            # 通知进度信息
            self.signal_Progressed.emit(i+1, totalFileNum)

        if self.__mergeFileList__:
            # 构造当前目录的合并文件名称
            mergeFileName = self.__buildMergeExcelPath__(self.__mergeFileList__[0])

            # 将数据写入到Excel文件中
            excelWriter.doWrite(mergeFileName, mergeSheetName, allDataContentList)

        # 通知正常完成
        self.signal_Finished.emit(True, '')

    # 构造合并文件的名称
    @staticmethod
    def __buildMergeExcelPath__(excelFilePath):
        # 获取当前文件的目录
        pathDir = os.path.dirname(excelFilePath)
        orgExcelBaseNameList = os.path.basename(excelFilePath).strip('.xlsx').strip('.xls').split('_')

        # 固定合并文件的Excel名称
        mergeBaseName = orgExcelBaseNameList[0] + '_Merge.xlsx'

        # 合并为完成文件名称
        mergeFileName = os.path.join(pathDir, mergeBaseName)

        return mergeFileName
