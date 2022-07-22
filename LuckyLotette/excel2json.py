 
import os
import sys
import codecs
import xlrd  # http://pypi.python.org/pypi/xlrd
 
 
def FloatToString(aFloat):
    if type(aFloat) != float:
        return ""
    strTemp = str(aFloat)
    strList = strTemp.split(".")
    if len(strList) == 1:
        return strTemp
    else:
        if strList[1] == "0":
            return strList[0]
        else:
            return strTemp
 
 
def table2json(table, jsonfilename, fileDir):
    hang = table.nrows
    lie = table.ncols
    f = codecs.open(fileDir+jsonfilename, "w", "utf-8")
    # f.write(u"[\n")
    f.write(u"{\n")
    # print(f)
    # 这里我们这的表是从第2行开始读
    for r in range(2, hang):
        # 每一行一组数据 一个对象的开始用{
        print("参数：%s，总计%d,%a"%(r, hang, lie))
        

        # 取第一列的参数作为主键，不需要的话可以修改成数组---------start----------
        strKeyValue = u""
        keyObj = table.cell_value(r, 0)
        # print("获得第一个参数：%s"%(keyObj))
        if type(keyObj) == float:
                strKeyValue = FloatToString(keyObj)
        else:
            strKeyValue = str(keyObj)
        strTmp1 = u"    \"" + strKeyValue + u"\": "
        f.write(strTmp1 + u"{")
        # 取第一列的参数作为主键，不需要的话可以修改成数组---------end----------
        
        # f.write(u"  {")

        for c in range(0, lie):
            # 定义一个空的字符串
            strCellValue = u""
            # 获取一个单元格的值
            CellObj = table.cell_value(r, c)
            # print("读取：%s"%(curPath+'\\'+a))
            print("获得：%s，%s"%(c, CellObj))
            # 判断数据类型如果是float类型要转成字符串
            if type(CellObj) == float:
                # print("1获得float")
                strCellValue = FloatToString(CellObj)
            else:
                # print("2获得不是float")
                # 转成字符串
                strCellValue = str(CellObj)
                # 值里面写"在里面防止转义报错要去掉 做过滤
                strCellValue = strCellValue.replace(u"\"", u"")
                # 变成Json的值字符串要加"
                strCellValue = u'\"'+strCellValue+u'\"'
            strTmp = u"\"" + table.cell_value(0, c) + u"\":" + strCellValue
            # print("获得strTmp：%s"%(strTmp))
            # 如果不是最后一个需要加,
            if c < lie-1:
                strTmp += u", "
            # 写字符串到{}中
            f.write(strTmp)
        f.write(u"}")
        # 每一个对象后面要加,
        if r < hang-1:
            f.write(u",")
        # 换行
        f.write(u"\n")
    # 最后所有的数据要用]反中括号包起来
    # f.write(u"]")
    f.write(u"}")
    # 关闭文件
    f.close()
    print("转换完成表 ", jsonfilename)
    return
 
 
# 取当前目录
curPath = os.path.dirname(__file__)
# 在当前目录下创建一个文件夹JSON
jsonDir = curPath+'\\JSON\\'
# 判断文件夹是否存在决定建不建文件夹
isExists = os.path.exists(jsonDir)
if not isExists:
    os.makedirs(jsonDir)
 
# 遍历当前目录查询出所有的excel表
fileNameList = os.listdir(curPath)
print(fileNameList)
for a in fileNameList:
    print(a)
    # 这里只能读取xlsx的表 如果是其他的表请加入判断
    extName = os.path.splitext(a)
    # 剔除缓存的表
    if(extName[0].find("~") >= 0 or extName[0].find("$") >= 0):
        continue
    # 只有这三种格式的才转 其他的不管
    if(extName[1] == '.xlsx' or extName[1] == ".csv" or extName[1] == ".xls"):
        print("读取sheet名字：%s"%(curPath+'\\'+a))
        # data = xlrd.open_workbook(curPath+'\\'+'item.xlsx')
        data = xlrd.open_workbook(curPath+'\\'+a)
        for sheetName in data.sheet_names():
            table = data.sheet_by_name(sheetName)
            print("读取sheet获得table：%s"%(table))
            table2json(table, sheetName+'.json', jsonDir)
 
        # table = data.sheet_by_index(0)
        # print("读取sheet获得table：%s"%(table))
        # table2json(table, a.replace(".xlsx", "")+'.json', jsonDir)
print("所有的表转换完成")