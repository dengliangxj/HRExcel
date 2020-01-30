import sys
import os
import xlsxwriter

#####################################################
# 将数据内容写入到Excel文件中
#####################################################


class ExcelDataWriter:

    def __init__(self):
        pass

    # 将listData写入到Excel文件中
    def doWrite(self, excelFilePath, sheetName, dataList):

        # 入参判定
        if (not excelFilePath) or (not dataList):
            print('ExcelDataWriter invalid input param')
            return

        print('write excel file: ', excelFilePath, ' sheetName: ', sheetName, ' begin')

        # 创建Excel文档
        workBook = xlsxwriter.Workbook(excelFilePath)
        workSheet = workBook.add_worksheet(sheetName)

        # 写入表头
        self.__writeXlsHeader__(workBook, workSheet, dataList[0])

        # 写入内容
        self.__writeXlsDataArea__(workBook, workSheet, dataList)

        # 关闭Excel文件
        workBook.close()

        print('write excel file: ', excelFilePath, ' sheetName: ', sheetName, ' end')

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

    # 写入数据区
    @staticmethod
    def __writeXlsDataArea__(workBook, workSheet, dataList):
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

        # 只要满足该行的数据包含所有的键值，则写入Excel中
        for i in range(1, len(dataList)):
            # 设置行高
            workSheet.set_row(i, 20)

            # 将该行写入到Excel
            workSheet.write_row(i, 0, dataList[i])
