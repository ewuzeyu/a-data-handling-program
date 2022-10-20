from msilib.schema import Error
from typing import Dict
import xlrd
import xlwt
import re
import sys
import os

# global variance
# local变量为1时为指定路径测试用，为0时为拖拽文件应用
local = 0
# local = 1

# repeat = 3
def mean(l:list) -> float:
    sum = 0
    for i in l:
        # if i == '':
        #     pass
        # else:
        #     sum += i
        sum += refloat(i)
    return sum/len(l)


def dic2list(d:dict) -> list:
    l = []
    for key in d:
        l.append(mean(d[key]))
    return l

def refloat(i):
    if isinstance(i,float):
        return i
    else:
        return 0

def getname(sample:str) -> str:
    res = re.match(r'(.*)-[0-9]+$',sample)
    if hasattr(res,'group'):
        return res.group(1)
    else:
        return sample

def average(l:list) -> float:
    if len(l) == 0:
        return 0
    s = 0
    for i in l:
        s = s + refloat(i)
    return s/len(l)

def stdev(l:list) -> float:
    if len(l) == 0 or len(l) == 1:
        return 0
    s2 = 0
    ave = average(l)
    for i in l:
        s2 = s2 + (refloat(i) - ave)**2
    return (s2/(len(l)-1))**0.5

def calcv(l:list) -> float:
    if average(l) == 0:
        return 0
    else:
        return stdev(l)/average(l)

# 四舍五入函数
def round(n) -> int:
    i = int(n)
    if n - i > 0.5:
        return i+1
    else: return i

class reporter:
    '''报告基因类
    
    Attributes:
        name: name of reporter gene
        virus: 
    '''
    def __init__(self, name:str) -> None:
        # name为报告基因名称
        self.name = name
        # virus为实际作用的病毒,初始仍为报告基因名称
        self.virus = name
        self.con = {}


# file = 'C:\\Users\\admin\\Desktop\\qPCR\\导出数据\\20220922 TIQUYOUHUA.xls'

# wk = xlrd.open_workbook(filename=file)
# sheet = wk.sheet_by_name('结果分析')
# row48 = sheet.row_values(48)

# patt = r'(.*)-[0-9]+$'
# res = re.match(patt,row48[3])

# print(getname(row48[3]))

# print(sheet.row_values(49)[8])

# 读入模块
def main():
    try:
        if local == 0:
            file = sys.argv[1].replace('\\','/').strip()
        else:
            file = r'C:\Users\admin\Desktop\qPCR\导出数据\20220927 zhilikuozeng.xls'

        wk = xlrd.open_workbook(filename=file)

        sheet = wk.sheet_by_name('结果分析')

        # reporter_gene = ['FAM','VIC','ROX','CY5']
        # dre = []
        # dre.append(reporter('FAM'))
        # dre.append(reporter('VIC'))
        # dre.append(reporter)
        dFAM = {} # FAM
        dVIC = {} # VIC
        dROX = {}
        dCY5 = {}

        for i in range(47,sheet.nrows):
            row = sheet.row_values(i)
            if row[6] == 'FAM':
                name = getname(row[3])
                if name == '':
                    continue
                if name not in dFAM:# 若samplename没有在字典中
                    dFAM[name] = {}# 则加入samplename
                    dFAM[name][row[3]] = [row[8]] # 添加sample
                else:
                    if row[3] not in dFAM[name]:# 若没有sample
                        dFAM[name][row[3]] = [row[8]]# 添加sample
                    else:
                        dFAM[name][row[3]].append(row[8])
            elif row[6] == 'VIC':
                name = getname(row[3])
                if name == '':
                    continue
                if name not in dVIC:# 若samplename没有在字典中
                    dVIC[name] = {}# 则加入samplename
                    dVIC[name][row[3]] = [row[8]] # 添加sample
                else:
                    if row[3] not in dVIC[name]:# 若没有sample
                        dVIC[name][row[3]] = [row[8]]# 添加sample
                    else:
                        dVIC[name][row[3]].append(row[8])
            elif row[6] == 'ROX':
                name = getname(row[3])
                if name == '':
                    continue
                if name not in dROX:# 若samplename没有在字典中
                    dROX[name] = {}# 则加入samplename
                    dROX[name][row[3]] = [row[8]] # 添加sample
                else:
                    if row[3] not in dROX[name]:# 若没有sample
                        dROX[name][row[3]] = [row[8]]# 添加sample
                    else:
                        dROX[name][row[3]].append(row[8])
            elif row[6] == 'CY5':
                name = getname(row[3])
                if name == '':
                    continue
                if name not in dCY5:# 若samplename没有在字典中
                    dCY5[name] = {}# 则加入samplename
                    dCY5[name][row[3]] = [row[8]] # 添加sample
                else:
                    if row[3] not in dCY5[name]:# 若没有sample
                        dCY5[name][row[3]] = [row[8]]# 添加sample
                    else:
                        dCY5[name][row[3]].append(row[8])
            else:
                pass


        sum_l = 0
        max_l = 0
        for k2 in dFAM:
            # 平均repeat
            # print(len(dFAM[k2].keys()))
            sum_l += len(dFAM[k2].keys())
            # 最大repeat
            if max_l < len(dFAM[k2].keys()):
                max_l = len(dFAM[k2].keys())
            else:
                pass

        sum_d = 0
        max_d =0
        for k2 in dFAM:
            for k3 in dFAM[k2]:
                if len(dFAM[k2][k3]) > max_d:
                    max_d = len(dFAM[k2][k3])
                else:
                    pass

        # 使用最大或者平均的重复数
        repeat = max_l
        # repeat = round(sum_l/len(dFAM.keys()))
        dupli = max_d
        # 测试区

        # for key in dFAM:
        #     print(dFAM[key])

        print(repeat)
        # print(dFAM)
        print(dupli)


        # 数据输出

        # 样式定义

        # 居中
        alignment = xlwt.Alignment()
        alignment.horz = 0x02
        alignment.vert = 0x01
        # styCT 供CT值单元格,两位小数数字
        styCT = xlwt.XFStyle()
        styCT.num_format_str = '0.00'
        styCT.alignment = alignment
        # styCV 供CV值单元格,百分比+两位小数
        styCV = xlwt.XFStyle()
        styCV.num_format_str = '0.00%'
        styCV.alignment = alignment
        # 居中样式
        stycent = xlwt.XFStyle()
        stycent.alignment = alignment

        # 初始化表头
        tablehead = ['样本','sample_name']+[str(i) for i in range(1,repeat+1)]+['CT','std','CV','备注']

        f = xlwt.Workbook()
        sheet1 = f.add_sheet('数据输出',cell_overwrite_ok=True)

        # 记录操作行
        ptr = 0

        sheet1.write(ptr,0,'FAM')
        ptr += 1


        for i in range(0,len(tablehead)):
            sheet1.write(ptr,i,tablehead[i],stycent)

        ptr += 1

        for key in dFAM:
            sheet1.write(ptr,1,key,stycent)
            for (i,value) in zip(range(2,2+repeat),dic2list(dFAM[key])):
                sheet1.write(ptr,i,value,styCT)
            sheet1.write(ptr,2+repeat,average(dic2list(dFAM[key])),styCT)
            sheet1.write(ptr,3+repeat,stdev(dic2list(dFAM[key])),styCT)
            sheet1.write(ptr,4+repeat,calcv(dic2list(dFAM[key])),styCV)
            ptr += 1

        ptr += 1

        sheet1.write(ptr,0,'VIC')
        ptr += 1

        for i in range(0,len(tablehead)):
            sheet1.write(ptr,i,tablehead[i],stycent)

        ptr += 1

        for key in dVIC:
            sheet1.write(ptr,1,key,stycent)
            for (i,value) in zip(range(2,2+repeat),dic2list(dVIC[key])):
                sheet1.write(ptr,i,value,styCT)
            sheet1.write(ptr,2+repeat,average(dic2list(dVIC[key])),styCT)
            sheet1.write(ptr,3+repeat,stdev(dic2list(dVIC[key])),styCT)
            sheet1.write(ptr,4+repeat,calcv(dic2list(dVIC[key])),styCV)
            ptr += 1

        ptr += 1

        sheet1.write(ptr,0,'ROX')
        ptr += 1

        for i in range(0,len(tablehead)):
            sheet1.write(ptr,i,tablehead[i],stycent)

        ptr += 1

        for key in dROX:
            sheet1.write(ptr,1,key,stycent)
            for (i,value) in zip(range(2,2+repeat),dic2list(dROX[key])):
                sheet1.write(ptr,i,value,styCT)
            sheet1.write(ptr,2+repeat,average(dic2list(dROX[key])),styCT)
            sheet1.write(ptr,3+repeat,stdev(dic2list(dROX[key])),styCT)
            sheet1.write(ptr,4+repeat,calcv(dic2list(dROX[key])),styCV)
            ptr += 1

        ptr += 1

        sheet1.write(ptr,0,'CY5')
        ptr += 1

        for i in range(0,len(tablehead)):
            sheet1.write(ptr,i,tablehead[i],stycent)

        ptr += 1

        for key in dCY5:
            sheet1.write(ptr,1,key,stycent)
            for (i,value) in zip(range(2,2+repeat),dic2list(dCY5[key])):
                sheet1.write(ptr,i,value,styCT)
            sheet1.write(ptr,2+repeat,average(dic2list(dCY5[key])),styCT)
            sheet1.write(ptr,3+repeat,stdev(dic2list(dCY5[key])),styCT)
            sheet1.write(ptr,4+repeat,calcv(dic2list(dCY5[key])),styCV)
            ptr += 1




        # duplicates

        tablehead2 = ['样本','sample_name']+[str(i) for i in range(1,dupli+1)]+['CT','std','CV','备注']
        sheet2 = f.add_sheet('数据输出2',cell_overwrite_ok=True)
        ptr2 = 0

        # 输出单个

        sheet2.write(ptr2,0,'FAM')
        ptr2 += 1

        for i in range(0,len(tablehead2)):
            sheet2.write(ptr2,i,tablehead2[i],stycent)

        ptr2 += 1

        for key in dFAM:
            for key2 in dFAM[key]:
                sheet2.write(ptr2,1,key2,stycent)
                for (i,value) in zip(range(2,2+dupli),dFAM[key][key2]):
                    sheet2.write(ptr2,i,value,styCT)
                sheet2.write(ptr2,2+dupli,average(dFAM[key][key2]),styCT)
                sheet2.write(ptr2,3+dupli,stdev(dFAM[key][key2]),styCT)
                sheet2.write(ptr2,4+dupli,calcv(dFAM[key][key2]),styCV)
                ptr2 += 1

        ptr2 += 1

        sheet2.write(ptr2,0,'VIC')
        ptr2 += 1

        for i in range(0,len(tablehead2)):
            sheet2.write(ptr2,i,tablehead2[i],stycent)

        ptr2 += 1

        for key in dVIC:
            for key2 in dVIC[key]:
                sheet2.write(ptr2,1,key2,stycent)
                for (i,value) in zip(range(2,2+dupli),dVIC[key][key2]):
                    sheet2.write(ptr2,i,value,styCT)
                sheet2.write(ptr2,2+dupli,average(dVIC[key][key2]),styCT)
                sheet2.write(ptr2,3+dupli,stdev(dVIC[key][key2]),styCT)
                sheet2.write(ptr2,4+dupli,calcv(dVIC[key][key2]),styCV)
                ptr2 += 1

        ptr2 += 1

        sheet2.write(ptr2,0,'ROX')
        ptr2 += 1

        for i in range(0,len(tablehead2)):
            sheet2.write(ptr2,i,tablehead2[i],stycent)

        ptr2 += 1

        for key in dROX:
            for key2 in dROX[key]:
                sheet2.write(ptr2,1,key2,stycent)
                for (i,value) in zip(range(2,2+dupli),dROX[key][key2]):
                    sheet2.write(ptr2,i,value,styCT)
                sheet2.write(ptr2,2+dupli,average(dROX[key][key2]),styCT)
                sheet2.write(ptr2,3+dupli,stdev(dROX[key][key2]),styCT)
                sheet2.write(ptr2,4+dupli,calcv(dROX[key][key2]),styCV)
                ptr2 += 1

        ptr2 += 1

        sheet2.write(ptr2,0,'CY5')
        ptr2 += 1

        for i in range(0,len(tablehead2)):
            sheet2.write(ptr2,i,tablehead2[i],stycent)

        ptr2 += 1

        for key in dCY5:
            for key2 in dCY5[key]:
                sheet2.write(ptr2,1,key2,stycent)
                for (i,value) in zip(range(2,2+dupli),dCY5[key][key2]):
                    sheet2.write(ptr2,i,value,styCT)
                sheet2.write(ptr2,2+dupli,average(dCY5[key][key2]),styCT)
                sheet2.write(ptr2,3+dupli,stdev(dCY5[key][key2]),styCT)
                sheet2.write(ptr2,4+dupli,calcv(dCY5[key][key2]),styCV)
                ptr2 += 1


        # 输出文件

        if local == 0:
            a = re.match(r'.+\/(.*)?.((?:xlsx)|(?:xls))$',file)
            newfilename = a.group(1)
        else:
            newfilename = ''

        f.save(newfilename + '数据处理' + '.xls')
    except:
        print(sys.exc_info())
    finally:
        pass

if __name__ == '__main__':
    main()