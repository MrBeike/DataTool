# -*-coding:utf-8-*-

import os
import openpyxl
import configparser

class DataTool:
    def __init__(self):
         # 获取当前工作目录
        self.path = os.getcwd()
        self.sep = os.path.sep
        self.config = configparser.ConfigParser()
        self.config.read("config.ini", encoding="utf-8")

    def getFileList(self):
        '''
        读取当前目录下所有J文件，后缀名为Dat
        '''
        # 列出文件夹下所有文件
        fileList = os.listdir(self.path)
        datFileList = []
        for file in fileList:
            # 筛选出DAT后缀文件
            if file.endswith('.DAT'):
                datFileList.append(file)
        return datFileList
        
    def readDataFile(self,fileName):
        '''
        文件名解析：BJ 6d03 3417000 20191130 421.DAT
        BJ:报数用户代码+文件类型  0:1  此处不涉及，不处理
        6do3: 机构代码  2:5
        3417000: 地区代码  6:12
        20191130: 报表数据日期  13:20
        421: 报表批次代码 21:23  
        '''
        fileInfo = {}
        fileInfo['organCode'] = fileName[2:6]
        fileInfo['localCode'] = fileName[6:13]
        fileInfo['date'] = fileName[13:21]
        fileInfo['formCode'] = fileName[21:24]
        fileData= {}
        fileInfo['data'] = fileData
        filePath = self.path+self.sep+fileName
        with open(filePath) as dataFile:
            '''
            J文件解析：I00002|34R01|1600
            I00002：报表指标Index 此处不涉及，不处理
            34R01：指标名称
            1600：指标数值
            '''
            for line in dataFile:
                # 去除每行末尾 \n
                line = line.strip()
                lineList = line.split('|')
                fileData[lineList[1]]=lineList[2]
        return fileInfo    

    def dataWriter(self,fileInfo):
        # 根据ini配置文件读取机构报表等相关信息
        organName = self.config.get("organCode", fileInfo['organCode'])
        sheetName = self.config.get("formCode",fileInfo['formCode']).split(',')
        print(sheetName)
        # 打开对应机构xlsx表
        excelFile = openpyxl.load_workbook(organName+'.xlsx')
        # 打开对应sheet,存在多个sheet的情况
        for sheet in sheetName:
            columnIndex = self.config.get('columnIndex',sheet).split(',')
            print(columnIndex)
            sheetFile = excelFile[sheet]
            tagLine = sheetFile[1]
            tags = [(x.value,x.column_letter) for x in tagLine]
            


            

k =DataTool()
datFileList = k.getFileList()
for filename in datFileList:
    fileInfo = k.readDataFile(filename)
    k.dataWriter(fileInfo)