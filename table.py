# JM导表工具
import os
import xlrd
import json
import os.path

# 表格地址
file_src = 'D:\\表格\\数值表格'
# 类型模板    
classTmplate = "/** \n   {notes} \n*/\ninterface C{name} {\n{prop}\n}\n\n"
propTemplate = "    /** {notes} */\n    {name}: {type};\n"

# 提取文件下所有xlsx
def getFiles():
    for root, dirs, files in os.walk(file_src):
        return files

# 跳过空行
def IsEmptyLine(arr):
    if len(arr) > 0:
        if arr[0] != "" and arr[0] == "是":
            return False
    return True

# 字段类型检查(置空和填写错误)
def checkType(sheet): 
    colTitle = sheet.row_values(0)
    colName = sheet.row_values(1)
    colType = sheet.row_values(2)
    for s in colTitle:
        if s == '':
            print('Error: 字段名没有填写')
            return False
    for a in colName:
        if a == '':
            print('Error: 字段名没有填写')
            return False
    for b in colType:
        if b != 'int' and b != 'str' and b != 'bool' and b != 'array' and b != 'float' and b != 'auto' and b != 'json':
            print('Error: 字段类型书写错误：' + b)
            return False
    return True

# 获取字段类型
def getColNames(sheet):
    column = {}
    colTitle = sheet.row_values(0)
    colName = sheet.row_values(1)
    colType = sheet.row_values(2)
    for index in range(len(colName)):
        if colName[index] != 'export':
            title = colTitle[index]
            type = colType[index]
            cname = colName[index].replace('*', '').replace('#', '')
            if (colType[index] == 'int') or (colType[index] == 'float'):
                type = 'number'
            elif (colType[index] == 'json') or (colType[index] == 'auto'):
                type = 'any'
            elif colType[index] == 'str':
                type = 'string'
            elif colType[index] == 'bool':
                type = 'boolean'
            elif colType[index] == 'array':
                type = 'any[]'
            if cname == colName[index]:
                column[cname + '?'] = [type, title]
            else:
                column[cname] = [type, title]
    return column

# 获取表格每行数据
def getRowData(sheet):
    data = []
    colType = sheet.row_values(2)
    del colType[0]
    for i in range(3, sheet.nrows):
        if IsEmptyLine(sheet.row_values(i)):  #跳过空行
            continue
        row_list = sheet.row_values(i)
        del row_list[0]
        for j in range(len(row_list)):
            #布尔类型转换
            if row_list[j] == '是':
                row_list[j] = True
            elif row_list[j] == '否':
                row_list[j] = False

            #excel浮点型问题
            if (colType[j] == 'int'):
                row_list[j] = int(row_list[j])
        data.append(row_list)
    return data

# 表格类型转换为数据类型
def xlsxTotype():
    isPass = True
    interface = ""
    fieldDict = []
    nameDict = []
    for name in getFiles():
        workbook = xlrd.open_workbook(file_src + '\\' + name)
        for str in workbook.sheet_names():
            name1 = str.split('|')[0].upper()
            name2 = str.split('|')[1]
            nameDict.append([name1, name2]);
        for idx in range(0, workbook.nsheets):
            sheet = workbook.sheet_by_index(idx)
            if checkType(sheet):
                fieldDict.append(getColNames(sheet))
            else:
                isPass = False
                break
    if isPass:
        for index in range(len(nameDict)):
            prop = ""
            for type in fieldDict[index]:
                prop += propTemplate.replace("{notes}", fieldDict[index][type][1]).replace("{name}", type).replace("{type}", fieldDict[index][type][0])
            interface += classTmplate.replace("{notes}", nameDict[index][1]).replace("{name}", nameDict[index][0]).replace("{prop}", prop)

        output = \
        open('D:\\表格\\table.d.ts', 'w', encoding = "utf-8")
        output.write(interface)
        output.close()
        print('=====================Success: 数据类型导出成功=======================')        
    else:
        print('=====================Error: 数据类型导出失败=========================')

# 表格数据转换为Json
def xlsxTojson():
    isPass = True
    nameDict = []   #文件名
    fieldDict = []  #字段名
    dataDict = []   #数据

    for name in getFiles():
        workbook = xlrd.open_workbook(file_src + '\\' + name)
        for str in workbook.sheet_names():
            name = str.split('|')[0] + '.json'
            nameDict.append(name);
        for idx in range(0, workbook.nsheets):
            sheet = workbook.sheet_by_index(idx)
            if checkType(sheet):
                arr = []
                colName = sheet.row_values(1)
                for index in range(len(colName)):
                    if colName[index] != 'export':
                        cname = colName[index].replace('*', '').replace('#', '')
                        arr.append(cname)
                fieldDict.append(arr)
            else:
                isPass = False
                break
            dataDict.append(getRowData(sheet))

    if isPass:
        for i in range(len(nameDict)):
            dict = {}
            for j in range(len(dataDict[i])):
                data = {}
                arr1 = fieldDict[i]
                arr2 = dataDict[i][j]
                for s in range(len(arr2)):
                    data[arr1[s]] = arr2[s]                 
                dict[arr2[0]] = data
            output = \
            open('D:\\表格\\table\\' + nameDict[i], "w", encoding = "utf-8")
            output.write(json.dumps(dict, indent = 4))
            output.close()
        print('=====================Success: 数据JSON导出成功=======================')
    else:
        print('=====================Error: 数据JSON导出失败=========================')

def main():
    xlsxTotype()
    xlsxTojson()
main()
