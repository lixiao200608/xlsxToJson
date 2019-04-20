# JM导表工具
import os
import xlrd
import json
import os.path

# 字段类型检查(置空和填写错误)
def checkType(sheet): 
    colName = sheet.row_values(1)
    colType = sheet.row_values(2)
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
    
# 获取文件下所有excel
def getFileName(file_dir):
    interface = ""
    isPass = True
    fieldDict = []
    nameDict = []
    fileList = []  
    classTmplate = "/** \n   {notes} \n*/\n{name} {\n{prop}\n}\n\n"
    propTemplate = "    /** {notes} */\n    {name}: {type};\n"
    for root, dirs, files in os.walk(file_dir):
        fileList = files
    for name in fileList:
        workbook = xlrd.open_workbook(file_dir + '\\' + name)
        for str in workbook.sheet_names():
            name1 = 'interface C' + str.split('|')[0].upper()
            name2 = str.split('|')[1].upper()
            nameDict.append([name1, name2]);
        for idx in range(0, workbook.nsheets):
            sheet = workbook.sheet_by_index(idx)
            if checkType(sheet):
                fieldDict.append(getColNames(sheet))
            else:
                isPass = False
    
    for index in range(len(nameDict)):
        prop = ""
        for type in fieldDict[index]:
            prop += propTemplate.replace("{notes}", fieldDict[index][type][1]).replace("{name}", type).replace("{type}", fieldDict[index][type][0])
        interface += classTmplate.replace("{notes}", nameDict[index][1]).replace("{name}", nameDict[index][0]).replace("{prop}", prop)

    return [isPass, interface]

def main():
    src = 'D:\\表格\\数值表格'
    dataType = getFileName(src)
    if dataType[0]:
        output = \
        open('D:\\表格\\table.d.ts', 'w', encoding = "utf-8")
        output.write(dataType[1])
        output.close()
        print('=====================Success: 表格导出成功=======================')
    else:
        print('=====================Error: 表格导出失败=======================')
        return
main()