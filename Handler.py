# -*- coding: utf-8 -*-

import xlrd
import xlwt
import types
import json
import chardet
import sys
import os
import os.path
reload(sys)
sys.setdefaultencoding('utf-8')
def openExcel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print "error" + str(e)



def readExcel(file='file.xls',colnameindex=0,by_index=0):
    data = openExcel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    colnames = table.row_values(0) #某一行数据
    list = []
    header = []
    for i in range(ncols):
        tData = colnames[i].encode('utf-8')
        header.append(tData)
    for cols in range(0,ncols):
        numbers = {}
        colNum = table.col_values(cols)
        temp = []
        for i in range(1,len(colNum)):
            temp.append(colNum[i])
        numbers[header[cols]] = temp
        list.append(numbers)
    return list


def set_style(name,bold=False,color = 0,height = 200):
    style = xlwt.XFStyle()
    fnt = xlwt.Font()                        # 创建一个文本格式，包括字体、字号和颜色样式特性
    fnt.name = name                # 设置其字体
    fnt.height = height
    fnt.colour_index = color               # 设置其字体颜色
    fnt.bold = bold
    style.font = fnt
    return  style

def executeExcel(data = []):
    calculateData = []
    dateKey = data[0].keys()[0]
    dataInfo = data[0][dateKey]
    lenDate = len(dataInfo)
    calculateData.append({dateKey:dataInfo[lenDate/2:lenDate]})
    for i in range(1,len(data)):
        key = data[i].keys()[0]
        templist = []
        tempMap ={}
        numbers = data[i][key]
        skip = len(numbers)/2
        for j in range(0,len(numbers)/2):
            t1 = numbers[j]
            t2 = numbers[j+skip ]
            ratio = (t2-t1)/t1
            templist.append(ratio)
        tempMap[key] = templist
        calculateData.append(tempMap)
    return calculateData


def writeExcel(sheet1,parent,shortName,start,pre = [],trend = []):
    posFix = unicode('涨跌', "utf-8")
    length = 0
    style_red = set_style(u'Times New Roman',False,2) ##红色字体
    style_blue = set_style('Times New Roman',False,4) ##蓝色字体
    try:

        ##写入详细数据
        sheet1.write_merge(start, start, 0, 10, '广告位' + shortName + "详细数据")
        start = start + 1
        for i in range(len(pre)):
            key = pre[i].keys()[0]
            number = pre[i][key]
            sheet1.write(start,i,key)
            length = len(number) + 1
            for j in range(len(number)):
                info = (number[j])
                sheet1.write(start+j+1,i,info)

        start = start + length + 2


        ##写入本周期涨跌趋势数据
        sheet1.write_merge(start, start, 0, 10, '广告位' + shortName + "涨跌")
        start = start + 1
        for i in range(len(trend)):
            key = trend[i].keys()[0]
            number = trend[i][key]
            length = len(number) + 1
            sheet1.write(start,i,key)
            for j in range(len(number)):
                info =  number[j]
                sheet1.write(start+j+1,i,info)
                if(i == 0):
                    sheet1.write(start+j+1,i,info)
                elif(info < 0):
                    info = "{:.3%}".format(info)
                    sheet1.write(start+j+1,i,info,style_red)

                else :
                    info = "{:.3%}".format(info)
                    sheet1.write(start+j+1,i,info,style_blue)
        start = start + length + 10

    except Exception,e:
        print e

    return start


def main():

    #存放待处理excel的文件夹
    rootdir = r'/Users/dainping/Desktop/人肉监控/'
    dirName = r"1221-1227"
    rootdir = rootdir + dirName
    start = 0
    f = xlwt.Workbook(encoding = 'utf-8') #创建工作簿
    sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True) #创建sheet
    for parent,dirnames,filenames in os.walk(rootdir):
        for filename in filenames:
            path = os.path.join(parent,filename)
            (shotname,extension) = os.path.splitext(filename)
            #下载下来的是.xlsx格式
            if(extension == ".xlsx"):
                ##读取excel内容
                ans = readExcel(path)
                ##处理excel数据，计算增长率
                cal = executeExcel(ans)
                ##写入excel文件，暂时只支持写成.xls
                start = writeExcel(sheet1,parent,shotname,start,ans,cal)
    path = os.path.join(parent,unicode('商详页'+ dirName +'周期数据分析.xls', "utf-8"))
    f.save(path)

if __name__=="__main__":
    main()