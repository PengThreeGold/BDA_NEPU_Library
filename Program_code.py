# -*- coding: utf-8 -*-
"""
Created on Sun May 30 14:37:43 2021

@author: Three_Gold
"""

"""
1、对提供的图书大数据做初步探索性分析。分析每一张表的数据结构。
2、对给定数据做数据预处理（数据清洗），去重，去空。
3、按照分配的院系保留自己需要的数据（教师的和学生的都要），这部分数据为处理好的数据。
4、将处理好的数据保存，文件夹名称为预处理后数据。
"""

import pandas as pd
import time
import os
import jieba
import wordcloud
path = os.getcwd()#获取项目根目录
path = path.replace('\\', '/')#设置项目根目录路径
Path = path + '/代码/12暖暖小组01号彭鑫项目/预处理后数据/'#设置预处理后数据根目录路径

#数据清洗，去重，去空

A1_UserID = pd.read_excel(path + '/代码/12暖暖小组01号彭鑫项目/原始数据/读者信息.xlsx')
A1_UserID = A1_UserID.dropna()#去除空值
A1_UserID = A1_UserID.drop_duplicates()#去重
A1_UserID = A1_UserID.reset_index(drop=True)#重设索引
A1_UserID.to_excel(path + '/代码/12暖暖小组01号彭鑫项目/清洗后数据/读者信息.xlsx')#保存数据

book_list = pd.read_excel(path + '/代码/12暖暖小组01号彭鑫项目/原始数据/图书目录.xlsx')#读取图书目录预处理后数据
book_list = book_list.dropna()#去除空值
book_list = book_list.drop_duplicates()#去重
book_list = book_list.reset_index(drop=True)#重设索引
book_list.to_excel(path + '/代码/12暖暖小组01号彭鑫项目/清洗后数据/图书目录.xlsx')#保存数据

def bookyearsdata():#对借还信息进行去重去空再保存
    Year_All = ['2014','2015','2016','2017']
    for year in Year_All:
        address = path + '/代码/12暖暖小组01号彭鑫项目/原始数据/图书借还' + year +'.xlsx'#获得预处理后数据路径
        address_last = path + '/代码/12暖暖小组01号彭鑫项目/清洗后数据/图书借还' + year +'.xlsx'#获得清洗后数据保存路径
        book = pd.read_excel(address)#读取预处理后数据
        book = book.dropna()#去除空值
        book = book.drop_duplicates()#去重
        book = book.reset_index(drop=True)#重设索引
        book.to_excel(address_last)#保存清洗后数据至新路径
        pass
    pass

bookyearsdata()#调用上述方法
print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))#获取打印本地当时时间
print('数据清洗，去重，去空完毕')

#利用原有数据重新制作图书分类表

#制作大类表
n=0#截取数据得到基本的大类对应表
with open(path + '/代码/12暖暖小组01号彭鑫项目/原始数据/《中国图书馆图书分类法》简表.txt',"r", encoding='UTF-8') as f:#用文件管理器打开预处理后数据文档
    with open(path + '/代码/12暖暖小组01号彭鑫项目/预处理后数据/图书大类表.txt', "w", encoding='UTF-8') as f1:#用文件管理器打开（创建）新的分类表
     for line in f.readlines():#读取预处理后数据文档的每一行
      n=n+1#行数累加
      if n <22:#截取行数小于22的行，对应所有图书大类的分类详情
          f1.writelines(line)#将行数小于22的行信息写入新的分类表
All_Book_List = pd.read_csv(path + '/代码/12暖暖小组01号彭鑫项目/预处理后数据/图书大类表.txt', names= ['图书分类号及类别'], sep='\n', encoding='utf8')#读取新的分类表
All_Book_List = All_Book_List.drop_duplicates()#去重
All_Book_List = All_Book_List.applymap(lambda x:x.strip() if type(x)==str else x)#去除左右空格
All_Book_List = All_Book_List.replace('K　历史、地理 N　自然科学总论','K　历史、地理')#清洗数据
All_Book_List = All_Book_List.replace('V    航空、航天','V　航空、航天')#针对性修改数据
All_Book_List = All_Book_List.replace('X    环境科学、劳动保护科学（安全科学）','X　环境科学、劳动保护科学（安全科学）')#针对性修改数据
All_Book_List.loc[21] = ("N　自然科学总论")#针对性添加数据
All_Book_List['图书分类号'] = All_Book_List['图书分类号及类别'].map(lambda x:x.split()[0])#截取中间空格前字符
All_Book_List['类别'] = All_Book_List['图书分类号及类别'].map(lambda x:x.split()[-1])#截取中间空格后字符
All_Book_List = All_Book_List[['图书分类号','类别']]#保留所需列
All_Book_List.to_excel(path + '/代码/12暖暖小组01号彭鑫项目/预处理后数据/图书大类表.xlsx')#写入保存
os.remove(path + '/代码/12暖暖小组01号彭鑫项目/预处理后数据/图书大类表.txt')#删除过渡数据
print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))#获取打印本地当时时间
print("制作大类表完毕")

#制作小说类别表，用于求小说类等
Novel_List = pd.read_csv(path + '/代码/12暖暖小组01号彭鑫项目/原始数据/《中国图书馆图书分类法》简表.txt',"r" ,names= ['图书分类号及类别'],encoding='UTF-8')#读取总表信息
Novel_List =Novel_List.loc[Novel_List['图书分类号及类别'].str.contains('I24')]#模糊搜索小说类别截取保存
Novel_List = Novel_List.drop_duplicates()#去重
Novel_List = Novel_List.applymap(lambda x:x.strip() if type(x)==str else x)#去除左右空格
Novel_List['图书分类号'] = Novel_List['图书分类号及类别'].map(lambda x:x.split(' ',1)[0])#截取中间空格前字符
Novel_List['类别'] = Novel_List['图书分类号及类别'].map(lambda x:x.split(' ',1)[-1])#截取中间空格后字符
Novel_List = Novel_List[['图书分类号','类别']]#保留所需列
for index,row in Novel_List['图书分类号'].iteritems():#添加未在分类法里有而数据里有的细化分类号
    if len(row) == 6:#判定为细化分类号的行
        I_row = pd.DataFrame()#创建临时空值表
        I_row_ = Novel_List.loc[Novel_List['图书分类号'] == row]#截取判定为细化分类号的行
        for i in range(10):#为截取判定为细化分类号的行细化数据添加尾号0-9
            i = str(i)#转型为str
            I_row_i = I_row_.replace(row,row+i)#再截取数据上进行修改加上尾号
            Novel_List = Novel_List.append(I_row_i)#写入添加架上尾号的数据
        pass
    pass
Novel_List.to_excel(path + '/代码/12暖暖小组01号彭鑫项目/预处理后数据/小说类别表.xlsx')#写入保存
print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))#获取打印本地当时时间
print("制作小说类别表完毕")


#制作图书目录字典
Book_xlsx = pd.DataFrame(pd.read_excel(path + '/代码/12暖暖小组01号彭鑫项目/清洗后数据/图书目录.xlsx',index_col=0))#导入数据
Book_Dic_Numble = Book_xlsx.set_index("图书ID").to_dict()['图书分类号']#创建图书id：图书分类号字典
Book_Dic_Name = Book_xlsx.set_index("图书ID").to_dict()['书名']#创建图书id:书名字典


#选择对应院系信息
A3_UserID = pd.read_excel(path + '/代码/12暖暖小组01号彭鑫项目/清洗后数据/读者信息.xlsx')#导入数据
A3_TeacherUserID = A3_UserID[A3_UserID['单位'] == '计算机与科学技术学院']#得到教师信息
A3_TeacherUserID = A3_TeacherUserID.append(A3_UserID[A3_UserID['单位'] == '计算机与信息技术学院'])#得到教师信息
A3_UserID = pd.DataFrame(A3_UserID[A3_UserID['读者类型'] == '本科生'])#得到所有本科生信息
A3_UserID_a = A3_UserID[A3_UserID['单位'] == '计2011-2']#筛选特殊专业
A3_UserID_b = A3_UserID[A3_UserID['单位'] == '计2012-08']#筛选特殊专业
A3_UserID_c = A3_UserID[A3_UserID['单位'] == '计2012-2']#筛选特殊专业
A3_UserID_All = pd.DataFrame()#学生数据总表

A3_UserID_All = A3_UserID_All.append(A3_UserID_a)#添加
A3_UserID_All = A3_UserID_All.append(A3_UserID_b)#添加
A3_UserID_All = A3_UserID_All.append(A3_UserID_c)#添加
glass = ['计2013', '计2014', '计2015', '计2016', '计2017',
         '软工2011','软工2013',
         '软件2014','软件2015','软件2016','软件2017',
         '物联2013','物联2014','物联2015','物联2016','物联2017',
         '信息2013','信息2014','信息2015','信息2016','信息2017']#制作院系专业表
for j in glass:#循环遍历得到所有专业表学生数据
    for i in range(1,7):
        glass_test = str(j)+ '-' + str(i)
        A3_UserID_test = A3_UserID[A3_UserID['单位']==glass_test]
        A3_UserID_All = A3_UserID_All.append(A3_UserID_test)
A3_StudentUserID_All = A3_UserID_All
A3_AllUserID = A3_StudentUserID_All.append(A3_TeacherUserID)#制成学院总表
A3_TeacherUserID = A3_TeacherUserID[['读者ID','读者类型','单位']]
A3_TeacherUserID.to_excel(path + '/代码/12暖暖小组01号彭鑫项目/预处理后数据/计算机院教师读者信息表.xlsx')#写入保存
A3_StudentUserID_All = A3_StudentUserID_All[['读者ID','读者类型','单位']]
A3_StudentUserID_All.to_excel(path + '/代码/12暖暖小组01号彭鑫项目/预处理后数据/计算机院学生读者信息表.xlsx')#写入保存
A3_AllUserID = A3_AllUserID[['读者ID','读者类型','单位']]
A3_AllUserID.to_excel(path + '/代码/12暖暖小组01号彭鑫项目/预处理后数据/计算机院读者信息表.xlsx')#写入保存

print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
print("选择对应院系信息完毕")

#得到图书借还的院系信息
def ChooseInf():#初步清洗各个年份数据，得到计算机院数据并将其保存
    Year_All = ['2014','2015','2016','2017']
    for Year in Year_All:
        AfterPath = path + '/代码/12暖暖小组01号彭鑫项目/清洗后数据/'
        A3_AllUser_bb = pd.DataFrame(pd.read_excel(AfterPath + '图书借还' + Year + '.xlsx',index_col=0))
        A3_TeacherUser = pd.DataFrame(pd.read_excel(path + '/代码/12暖暖小组01号彭鑫项目/预处理后数据/计算机院教师读者信息表.xlsx',index_col=0))
        A3_StudentUser = pd.DataFrame(pd.read_excel(path + '/代码/12暖暖小组01号彭鑫项目/预处理后数据/计算机院学生读者信息表.xlsx',index_col=0))
        A3_AllUser_bb['读者ID'] = A3_AllUser_bb['读者ID'].apply(str)#转换为统一数据类型
        A3_TeacherUser['读者ID'] = A3_TeacherUser['读者ID'].apply(str)#转换为统一数据类型
        A3_StudentUser['读者ID'] = A3_StudentUser['读者ID'].apply(str)#转换为统一数据类型
        A3_Result_Teacher = pd.merge(A3_AllUser_bb,A3_TeacherUser,how='outer',left_on = A3_AllUser_bb['读者ID'],right_on = A3_TeacherUser['读者ID'])#拼接两个数据表，得到有效数据和无效空值数据
        A3_Result_Teacher = A3_Result_Teacher.dropna()#去除缺失值行
        A3_Result_Teacher = A3_Result_Teacher.drop_duplicates()#去除重复值行
        A3_Result_Teacher = A3_Result_Teacher[['操作时间','操作类型','图书ID']]#保留有效信息
        A3_Result_Teacher.to_excel(path + '/代码/12暖暖小组01号彭鑫项目/预处理后数据/计算机院教师图书借还' + Year + '.xlsx')#保存该年份清洗后数据
        print('保存' + Year + '教师借阅初始信息成功')
        A3_Result_Student = pd.merge(A3_AllUser_bb,A3_StudentUser,how='outer',left_on = A3_AllUser_bb['读者ID'],right_on = A3_StudentUser['读者ID'])#拼接两个数据表，得到有效数据和无效空值数据
        A3_Result_Student = A3_Result_Student.dropna()#去除缺失值行
        A3_Result_Student = A3_Result_Student.drop_duplicates()#去除重复值行
        A3_Result_Student = A3_Result_Student[['操作时间','操作类型','图书ID']]#保留有效信息
        A3_Result_Student.to_excel(path + '/代码/12暖暖小组01号彭鑫项目/预处理后数据/计算机院学生图书借还' + Year + '.xlsx')#拼接两个数据表，得到有效数据和无效空值数据
        print('保存' + Year + '学生借阅初始信息成功')
    pass

ChooseInf()
print(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()))
print("得到图书借还的院系信息完毕")

def AddInf():#将各个年份数据添加到总表，去除还书数据，得到完整的借书数据
    Year_All = ['2014','2015','2016','2017']
    Da_Teacher_All = pd.DataFrame()
    Da_Student_All = pd.DataFrame()
    for Year in Year_All:
        Da_Teacher = pd.DataFrame(pd.read_excel(Path + '计算机院教师图书借还' + Year +'.xlsx',index_col=0))
        Da_Teacher_All = Da_Teacher_All.append(Da_Teacher)
        print('已添加%s教师信息至总表'%(Year))
        Da_Student = pd.DataFrame(pd.read_excel(Path + '计算机院学生图书借还' + Year +'.xlsx',index_col=0))
        Da_Student_All = Da_Student_All.append(Da_Student)
        print('已添加%s学生信息至总表'%(Year))
        pass
    Da_Teacher_All = Da_Teacher_All.reset_index(drop=True)
    Da_Student_All = Da_Student_All.reset_index(drop=True)
    #去除所有同一本书的"还"数据
    Da_Teacher_All_Back = Da_Teacher_All['操作类型'].isin(['还'])
    Da_Student_All_Back = Da_Student_All['操作类型'].isin(['还'])
    Da_Teacher_All_Borrow = Da_Teacher_All[~Da_Teacher_All_Back]
    Da_Student_All_Borrow = Da_Student_All[~Da_Student_All_Back]
    #去除同一本书同时被两个读者借的情况（即一个读者两个ID,数据默认保留第一条）
    Da_Teacher_All_Borrow = Da_Teacher_All_Borrow.drop_duplicates(subset=['操作时间','图书ID'],keep ='first')
    Da_Student_All_Borrow = Da_Student_All_Borrow.drop_duplicates(subset=['操作时间','图书ID'],keep ='first')
    Da_Student_All_Borrow = Da_Student_All_Borrow.reset_index()
    Da_Student_All_Borrow.to_excel(Path + '计算机院学生图书借总表.xlsx')
    print('生成计算机院学生图书借总表成功')
    Da_Teacher_All_Borrow = Da_Teacher_All_Borrow.reset_index()
    Da_Teacher_All_Borrow.to_excel(Path + '计算机院教师图书借总表.xlsx')
    print('生成计算机院教师图书借总表成功')
    pass
AddInf()


'''以上为任务基础代码'''
'''==========================================================================================================================================='''
'''以下为各个任务代码'''

#加载必要的库
import pandas as pd
import os
import time
import jieba
import wordcloud
#设置路径
path = os.getcwd()#获取项目根目录
path = path.replace('\\', '/')#设置项目根目录路径
Path = path + '/代码/12暖暖小组01号彭鑫项目/预处理后数据/'#设置预处理后数据根目录路径

#为下面程序提供数据,避免重复加载
Da_Teacher_All_Borrow = pd.DataFrame(pd.read_excel(Path + '计算机院教师图书借总表.xlsx',index_col=0))#导入数据
Da_Teacher_All_Borrow['操作时间']=Da_Teacher_All_Borrow['操作时间'].apply(str)#将数据转化为str类型
Da_Student_All_Borrow = pd.DataFrame(pd.read_excel(Path + '计算机院学生图书借总表.xlsx',index_col=0))#导入数据
Da_Student_All_Borrow['操作时间']=Da_Student_All_Borrow['操作时间'].apply(str)#将数据转化为str类型
All_Book_List = pd.DataFrame(pd.read_excel(Path + '图书大类表.xlsx',index_col=0))#导入数据
Book_xlsx = pd.DataFrame(pd.read_excel(path + '/代码/12暖暖小组01号彭鑫项目/清洗后数据/图书目录.xlsx',index_col=0))#导入数据
Novel_List = pd.DataFrame(pd.read_excel(Path + '小说类别表.xlsx',index_col=0))#导入数据

# 5、统计你分配的那个学院的教师2014年，2015年，2016年，2017年所借书籍类别占前10的都是什么类别。
def Task5():
    Year_All = ['2014','2015','2016','2017']
    for Year in Year_All:
        Book_xlsx['图书ID'] = Book_xlsx['图书ID'].apply(str)#转型以便拼接
        Da_Teacher_All_Borrow_Year = Da_Teacher_All_Borrow[Da_Teacher_All_Borrow['操作时间'].str.contains(Year)]#截取该年份信息
        Da_Teacher_All_Borrow_Year = pd.merge(Da_Teacher_All_Borrow_Year,Book_xlsx,how='left',left_on=Da_Teacher_All_Borrow_Year['图书ID'],right_on=Book_xlsx['图书ID'])#将借信息与图书目录信息按照图书id进行拼接
        Da_Teacher_All_Borrow_Year_BookNumble = Da_Teacher_All_Borrow_Year[['图书分类号']]#保留所需信息
        Da_Teacher_All_Borrow_Year_BookNumble = Da_Teacher_All_Borrow_Year_BookNumble['图书分类号'].str.extract('([A-Z])', expand=False)#利用正则表达式获取图书分类号首字母
        Da_Teacher_All_Borrow_Year_mes = pd.merge(Da_Teacher_All_Borrow_Year_BookNumble,All_Book_List)#将获取的首字母与图书大类表进行拼接
        Da_Teacher_All_Borrow_Year_10 = Da_Teacher_All_Borrow_Year_mes['类别'].value_counts()[:10]#保留排名前10的信息
        Da_Teacher_All_Borrow_Year_10 = Da_Teacher_All_Borrow_Year_10.reset_index()#重设索引
        Da_Teacher_All_Borrow_Year_10.columns = ['类别','数量']#设置列名
        Da_Teacher_All_Borrow_Year_10.index = Da_Teacher_All_Borrow_Year_10.index + 1#将索引加1得到排名
        Da_Teacher_All_Borrow_Year_10.index.name = '排名'#将索引名称设置为排名
        print('---------%s---------'%(Year) + '\n' + '%s教师所借书籍前十'%(Year))
        print(Da_Teacher_All_Borrow_Year_10)#打印信息
        pass
    pass
pass

# 6、统计你分配的那个学院的教师2014年，2015年，2016年，2017年所最喜欢看的小说分别是那一本。
def Task6():
    Year_All = ['2014','2015','2016','2017']
    for Year in Year_All:
        Book_xlsx['图书ID'] = Book_xlsx['图书ID'].apply(str)#转型以便拼接
        Da_Teacher_All_Borrow_Year = Da_Teacher_All_Borrow[Da_Teacher_All_Borrow['操作时间'].str.contains(Year)]#截取该年份信息
        Da_Teacher_All_Borrow_Year = pd.merge(Da_Teacher_All_Borrow_Year,Book_xlsx,how='left',left_on=Da_Teacher_All_Borrow_Year['图书ID'],right_on=Book_xlsx['图书ID'])#将借信息与图书目录信息按照图书id进行拼接
        Da_Teacher_All_Borrow_Year_BookNumble = Da_Teacher_All_Borrow_Year[['图书分类号','书名']]#保留所需信息
        Da_Teacher_All_Borrow_Year_BookNumble['图书分类号'] = Da_Teacher_All_Borrow_Year_BookNumble['图书分类号'].apply(str)#转型以便拼接
        Da_Teacher_All_Borrow_Year_novel = Da_Teacher_All_Borrow_Year_BookNumble.loc[Da_Teacher_All_Borrow_Year_BookNumble['图书分类号'].str.contains('I24')]#模糊搜索小说信息截取保存
        Da_Teacher_All_Borrow_Year_novel_1 =Da_Teacher_All_Borrow_Year_novel['书名'].value_counts()[:1]#对书名进行数量排序，并得到数量最多的书名
        Da_Teacher_All_Borrow_Year_novel_1.index.name = '书名'#将索引名设置为'书名'
        print('---------%s---------'%(Year) + '\n' + '%s教师最喜欢看的小说'%(Year))
        print(Da_Teacher_All_Borrow_Year_novel_1)#打印信息
        pass
    pass


# 7、统计你分配的那个学院的教师2014年，2015年，2016年，2017年一共借了多少书，专业书籍多少本。
def Task7():
    Year_All = ['2014','2015','2016','2017']
    All_Book = 0#初始化数据
    All_MBook = 0#初始化数据
    for Year in Year_All:
        Book_xlsx['图书ID'] = Book_xlsx['图书ID'].apply(str)#转型以便拼接
        Da_Teacher_All_Borrow_Year = Da_Teacher_All_Borrow[Da_Teacher_All_Borrow['操作时间'].str.contains(Year)]#截取该年份信息
        Da_Teacher_All_Borrow_Year = pd.merge(Da_Teacher_All_Borrow_Year,Book_xlsx,how='left',left_on=Da_Teacher_All_Borrow_Year['图书ID'],right_on=Book_xlsx['图书ID'])#将借信息与图书目录信息按照图书id进行拼接
        Da_Teacher_All_Borrow_Year_BookNumble = Da_Teacher_All_Borrow_Year[['图书分类号','书名']]#保留所需信息
        Da_Teacher_All_Borrow_Year_major = Da_Teacher_All_Borrow_Year_BookNumble#数据另存为
        Da_Teacher_All_Borrow_Year_major['图书分类号'] = Da_Teacher_All_Borrow_Year_major['图书分类号'].apply(str)#转型以便拼接
        Da_Teacher_All_Borrow_Year_major = Da_Teacher_All_Borrow_Year_major.loc[Da_Teacher_All_Borrow_Year_major['图书分类号'].str.contains('TP3')]#模糊搜索专业书籍信息截取保存
        All_Book = All_Book + Da_Teacher_All_Borrow_Year_BookNumble.shape[0]#加上每一年借书本数
        All_MBook = All_MBook + Da_Teacher_All_Borrow_Year_major.shape[0]#加上每一年借专业书本书
        print('---------%s---------'%(Year))
        print('%s教师总共借书'%(Year) + str(Da_Teacher_All_Borrow_Year_BookNumble.shape[0]) + '本')#打印信息
        print('%s教师借了专业书籍'%(Year) + str(Da_Teacher_All_Borrow_Year_major.shape[0]) + '本')#打印信息
        pass
    print('-------------------------------')
    print('计算机与信息技术学院教师（2014年-2017年）一共借'+str(All_Book)+'书')#打印信息
    print('计算机与信息技术学院教师（2014年-2017年）一共借了专业书'+str(All_MBook)+'书')#打印信息
    pass


# 8、统计你分配的那个学院的教师2014年，2015年，2016年，2017年一共有多少本书没有归还。没有归还的书籍哪类书籍最多。
def Task8():
     Da_Teacher_All = pd.DataFrame(pd.read_excel(Path + '计算机院教师图书借总表.xlsx',index_col=0))#读取院系借还信息
     Da_Teacher_All = Da_Teacher_All.sort_values(by = '操作时间')#根据时间降序排列
     Da_Teacher_All_Unback = Da_Teacher_All.drop_duplicates(subset=['图书ID'],keep ='last')#根据图书ID去重，保留最后一条信息
     Da_Teacher_All_Unback_Num = pd.merge(Da_Teacher_All_Unback,Book_xlsx)#将信息与图书目录按照图书id拼接
     Da_Teacher_All_Unback_Num = Da_Teacher_All_Unback_Num[['图书分类号']]#保留所需信息
     Da_Teacher_All_Unback_Num['图书分类号'] = Da_Teacher_All_Unback_Num['图书分类号'].apply(str)#转型以便拼接
     Da_Teacher_All_Unback_Num = Da_Teacher_All_Unback_Num['图书分类号'].str.extract('([A-Z])', expand=False)#利用正则表达式获取图书分类号首字母
     Da_Teacher_All_Unback_Num = pd.merge(Da_Teacher_All_Unback_Num,All_Book_List)#将获取的首字母与图书大类表进行拼接
     Da_Teacher_All_Unback_Num = Da_Teacher_All_Unback_Num['类别'].value_counts()[:1]#对类别进行数量排序，并得到数量最多的类别
     print('------------------' + '\n' + '2014~2017教师总共未归还书籍' + str(Da_Teacher_All_Unback.shape[0]) + '本')#打印信息
     print('教师未归还书籍中最多的类别是')
     print(Da_Teacher_All_Unback_Num)
     print('\n')
     pass


'''下列9-12代码结构与上面基本相似省略注释'''
# 9、统计你分配的那个学院的学生2014年，2015年，2016年，2017年所借书籍类别占前10的都是什么类别。
def Task9():
    Year_All = ['2014','2015','2016','2017']
    for Year in Year_All:
        Book_xlsx['图书ID'] = Book_xlsx['图书ID'].apply(str)
        Da_Student_All_Borrow_Year = Da_Student_All_Borrow[Da_Student_All_Borrow['操作时间'].str.contains(Year)]
        Da_Student_All_Borrow_Year = pd.merge(Da_Student_All_Borrow_Year,Book_xlsx,how='left',left_on=Da_Student_All_Borrow_Year['图书ID'],right_on=Book_xlsx['图书ID'])
        Da_Student_All_Borrow_Year_BookNumble = Da_Student_All_Borrow_Year[['图书分类号']]
        Da_Student_All_Borrow_Year_BookNumble['图书分类号'] = Da_Student_All_Borrow_Year_BookNumble['图书分类号'].apply(str)
        Da_Student_All_Borrow_Year_BookNumble = Da_Student_All_Borrow_Year_BookNumble['图书分类号'].str.extract('([A-Z])', expand=False)
        Da_Student_All_Borrow_Year_mes = pd.merge(Da_Student_All_Borrow_Year_BookNumble,All_Book_List)
        Da_Student_All_Borrow_Year_10 = Da_Student_All_Borrow_Year_mes['类别'].value_counts()[:10]
        Da_Student_All_Borrow_Year_10 = Da_Student_All_Borrow_Year_10.reset_index()
        Da_Student_All_Borrow_Year_10.columns = ['类别','数量']
        Da_Student_All_Borrow_Year_10.index = Da_Student_All_Borrow_Year_10.index + 1
        Da_Student_All_Borrow_Year_10.index.name = '排名'
        print('---------%s---------'%(Year) + '\n' + '%s学生所借书籍前十'%(Year))
        print(Da_Student_All_Borrow_Year_10)
        pass
    pass


# 10、统计你分配的那个学院的学生2014年，2015年，2016年，2017年所最喜欢看的小说分别是那一本。
def Task10():
    Year_All = ['2014','2015','2016','2017']
    for Year in Year_All:
        Book_xlsx['图书ID'] = Book_xlsx['图书ID'].apply(str)
        Da_Student_All_Borrow_Year = Da_Student_All_Borrow[Da_Student_All_Borrow['操作时间'].str.contains(Year)]
        Da_Student_All_Borrow_Year = pd.merge(Da_Student_All_Borrow_Year,Book_xlsx,how='left',left_on=Da_Student_All_Borrow_Year['图书ID'],right_on=Book_xlsx['图书ID'])
        Da_Student_All_Borrow_Year_BookNumble = Da_Student_All_Borrow_Year[['图书分类号','书名']]
        Da_Student_All_Borrow_Year_BookNumble['图书分类号'] = Da_Student_All_Borrow_Year_BookNumble['图书分类号'].apply(str)
        Da_Student_All_Borrow_Year_novel = Da_Student_All_Borrow_Year_BookNumble.loc[Da_Student_All_Borrow_Year_BookNumble['图书分类号'].str.contains('I24')]
        Da_Student_All_Borrow_Year_novel_1 =Da_Student_All_Borrow_Year_novel['书名'].value_counts()[:1]
        Da_Student_All_Borrow_Year_novel_1.index.name = '书名'
        print('---------%s---------'%(Year) + '\n' + '%s学生最喜欢看的小说'%(Year))
        print(Da_Student_All_Borrow_Year_novel_1)
        pass
    pass


# 11、统计你分配的那个学院的学生2014年，2015年，2016年，2017年一共借了多少书，专业书籍多少本。
def Task11():
    Year_All = ['2014','2015','2016','2017']
    All_Book = 0
    All_MBook = 0
    for Year in Year_All:
        Book_xlsx['图书ID'] = Book_xlsx['图书ID'].apply(str)
        Da_Student_All_Borrow_Year = Da_Student_All_Borrow[Da_Student_All_Borrow['操作时间'].str.contains(Year)]
        Da_Student_All_Borrow_Year = pd.merge(Da_Student_All_Borrow_Year,Book_xlsx,how='left',left_on=Da_Student_All_Borrow_Year['图书ID'],right_on=Book_xlsx['图书ID'])
        Da_Student_All_Borrow_Year_BookNumble = Da_Student_All_Borrow_Year[['图书分类号','书名']]
        Da_Student_All_Borrow_Year_major = Da_Student_All_Borrow_Year_BookNumble
        Da_Student_All_Borrow_Year_major['图书分类号'] = Da_Student_All_Borrow_Year_major['图书分类号'].apply(str)
        Da_Student_All_Borrow_Year_major = Da_Student_All_Borrow_Year_major.loc[Da_Student_All_Borrow_Year_major['图书分类号'].str.contains('TP3')]
        All_Book = All_Book + Da_Student_All_Borrow_Year_BookNumble.shape[0]
        All_MBook = All_MBook + Da_Student_All_Borrow_Year_major.shape[0]
        print('---------%s---------'%(Year))
        print('%s学生总共借书'%(Year) + str(Da_Student_All_Borrow_Year_BookNumble.shape[0]) + '本')
        print('%s学生借了专业书籍'%(Year) + str(Da_Student_All_Borrow_Year_major.shape[0]) + '本')
        pass
    print('-------------------------------')
    print('计算机与信息技术学院学生（2014年-2017年）一共借'+str(All_Book)+'书')
    print('计算机与信息技术学院学生（2014年-2017年）一共借了专业书'+str(All_MBook)+'书')
    pass


# 12、统计你分配的那个学院的学生2014年，2015年，2016年，2017年一共有多少本书没有归还。没有归还的书籍哪类书籍最多。
def Task12():
     Da_Student_All = pd.DataFrame(pd.read_excel(Path + '计算机院学生图书借总表.xlsx',index_col=0))
     Da_Student_All = Da_Student_All.sort_values(by = '操作时间')
     Da_Student_All_Unback = Da_Student_All.drop_duplicates(subset=['图书ID'],keep ='last')
     Da_Student_All_Unback_Num = pd.merge(Da_Student_All_Unback,Book_xlsx)
     Da_Student_All_Unback_Num = Da_Student_All_Unback_Num[['图书分类号']]
     Da_Student_All_Unback_Num['图书分类号'] = Da_Student_All_Unback_Num['图书分类号'].apply(str)
     Da_Student_All_Unback_Num = Da_Student_All_Unback_Num['图书分类号'].str.extract('([A-Z])', expand=False)
     Da_Student_All_Unback_Num = pd.merge(Da_Student_All_Unback_Num,All_Book_List)
     Da_Student_All_Unback_Num = Da_Student_All_Unback_Num['类别'].value_counts()[:1]
     print('------------------' + '\n' + '2014~2017学生未归还书籍' + str(Da_Student_All_Unback.shape[0]) + '本')
     print('学生未归还书籍中最多的类别是')
     print(Da_Student_All_Unback_Num)
     print('\n')
     pass


# 13、用折线图画出你分配的那个学院的教师2014年，2015年，2016年，2017年所借书籍类别占前10的书籍的走向图。
def Task13():
    import pandas as pd
    import matplotlib as mpl
    import matplotlib.pyplot as plt
    Year_All = ['2014','2015','2016','2017']
    All_Nnmble_10 = pd.DataFrame()#初始化数据
    for Year in Year_All:
        Book_xlsx['图书ID'] = Book_xlsx['图书ID'].apply(str)#转型以便拼接
        Da_Teacher_All_Borrow_Year = Da_Teacher_All_Borrow[Da_Teacher_All_Borrow['操作时间'].str.contains(Year)]#截取该年份信息
        Da_Teacher_All_Borrow_Year = pd.merge(Da_Teacher_All_Borrow_Year,Book_xlsx,how='left',left_on=Da_Teacher_All_Borrow_Year['图书ID'],right_on=Book_xlsx['图书ID'])#将借信息与图书目录信息按照图书id进行拼接
        Da_Teacher_All_Borrow_Year_BookNumble = Da_Teacher_All_Borrow_Year[['图书分类号']]#保留所需信息
        Da_Teacher_All_Borrow_Year_BookNumble = Da_Teacher_All_Borrow_Year_BookNumble['图书分类号'].str.extract('([A-Z])', expand=False)#利用正则表达式获取图书分类号首字母
        Da_Teacher_All_Borrow_Year_mes = pd.merge(Da_Teacher_All_Borrow_Year_BookNumble,All_Book_List)#将获取的首字母与图书大类表进行拼接
        Da_Teacher_All_Borrow_Year_Numble = Da_Teacher_All_Borrow_Year_mes['类别'].value_counts()[:20]#对类别进行数量排序，并得到数量前20的类别
        Da_Teacher_All_Borrow_Year_Numble = Da_Teacher_All_Borrow_Year_Numble.reset_index()#重置索引，将原先索引变为数值列
        Da_Teacher_All_Borrow_Year_Numble.index = Da_Teacher_All_Borrow_Year_Numble.index + 1#索引加1，得到排名
        Da_Teacher_All_Borrow_Year_Numble = Da_Teacher_All_Borrow_Year_Numble.reset_index()#再次重置索引，将原先索引变为数值列
        Da_Teacher_All_Borrow_Year_Numble = Da_Teacher_All_Borrow_Year_Numble[['level_0','index']]#保留所需信息
        if Year == '2014' :
            All_Nnmble_10 = Da_Teacher_All_Borrow_Year_Numble#将2014年数据作为基础数据
        else:
            All_Nnmble_10 = pd.merge( All_Nnmble_10,Da_Teacher_All_Borrow_Year_Numble,how = 'left',left_on=All_Nnmble_10['index'],right_on=Da_Teacher_All_Borrow_Year_Numble['index'])#将其他年份信息和基础信息拼接
            All_Nnmble_10 = All_Nnmble_10.rename(columns={'key_0':'index'})#保留所需信息，以便下次拼接使用
            pass
        pass
    All_Nnmble_10 = All_Nnmble_10[['index','level_0_x','level_0_y']]#保留所需信息（图书目录，2014~2017年排名信息）
    All_Nnmble_10.columns = ['书类','2014','2015','2016','2017']#设置列名
    All_Nnmble_10 = All_Nnmble_10[:10]#截取前十
    All_Nnmble_10_last = pd.DataFrame()#初始化数据
    for Year in Year_All:
        Sort = All_Nnmble_10[Year].sort_values(na_position='last')
        Sort = Sort.reset_index()
        Sort.index = Sort.index +1
        Sort = Sort.reset_index()
        Sort = Sort[['level_0','index']]
        if Year == '2014':          
            All_Nnmble_10_last = Sort
        else:
            All_Nnmble_10_last = pd.merge(All_Nnmble_10_last,Sort,how = 'left',left_on = All_Nnmble_10_last['index'],right_on = Sort['index'])
            All_Nnmble_10_last = All_Nnmble_10_last.rename(columns={'key_0':'index'})
    All_Nnmble_10_last = All_Nnmble_10_last[['level_0_x','level_0_y']]#保留所需信息（图书目录，2014~2017年排名信息）
    All_Nnmble_10_last.columns = ['2014','2015','2016','2017']
    All_Nnmble_10 = All_Nnmble_10.set_index('书类')#将书类列设置为索引
    All_Nnmble_10_last.index = All_Nnmble_10.index
    All_Nnmble_10_last = All_Nnmble_10_last.T#将表格转置
    mpl.rcParams['font.sans-serif'] = ['KaiTi']#设置字体
    fig, ax = plt.subplots(figsize=(5,8))
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    for name, value in All_Nnmble_10_last[3:4].max().items():
        plt.text(3.25,value, name, size=10, ha='left', va='center')
    plt.plot(All_Nnmble_10_last,'-o')
    plt.ylim([11,0])
    plt.yticks([i for i in range(1,11)])
    plt.title('2014-2017计算机院教师借阅书籍类别前十',fontsize=20)
    plt.savefig(path+'/代码/12暖暖小组01号彭鑫项目/项目可视化图/任务13.png',dpi=300,bbox_inches='tight')
    plt.show()
    pass


# 14、用折线图画出你分配的那个学院的学生2014年，2015年，2016年，2017年所借书籍类别占前10的书籍的走向图。（注释同Task13）
def Task14():
    import pandas as pd
    import matplotlib as mpl
    import matplotlib.pyplot as plt
    Year_All = ['2014','2015','2016','2017']
    All_numble_10 = pd.DataFrame()
    for Year in Year_All:
        Book_xlsx['图书ID'] = Book_xlsx['图书ID'].apply(str)
        Da_Student_All_Borrow_Year = Da_Student_All_Borrow[Da_Student_All_Borrow['操作时间'].str.contains(Year)]
        Da_Student_All_Borrow_Year = pd.merge(Da_Student_All_Borrow_Year,Book_xlsx,how='left',left_on=Da_Student_All_Borrow_Year['图书ID'],right_on=Book_xlsx['图书ID'])
        Da_Student_All_Borrow_Year_BookNumble = Da_Student_All_Borrow_Year[['图书分类号']]
        Da_Student_All_Borrow_Year_BookNumble = Da_Student_All_Borrow_Year_BookNumble['图书分类号'].str.extract('([A-Z])', expand=False)
        Da_Student_All_Borrow_Year_mes = pd.merge(Da_Student_All_Borrow_Year_BookNumble,All_Book_List)
        Da_Student_All_Borrow_Year_Numble = Da_Student_All_Borrow_Year_mes['类别'].value_counts()[:20]
        Da_Student_All_Borrow_Year_Numble = Da_Student_All_Borrow_Year_Numble.reset_index()
        Da_Student_All_Borrow_Year_Numble.index = Da_Student_All_Borrow_Year_Numble.index + 1
        Da_Student_All_Borrow_Year_Numble = Da_Student_All_Borrow_Year_Numble.reset_index()
        Da_Student_All_Borrow_Year_Numble = Da_Student_All_Borrow_Year_Numble[['level_0','index']]
        if Year == '2014' :
            All_numble_10 = Da_Student_All_Borrow_Year_Numble
        else:
            All_numble_10 = pd.merge( All_numble_10,Da_Student_All_Borrow_Year_Numble,how = 'left',left_on=All_numble_10['index'],right_on=Da_Student_All_Borrow_Year_Numble['index'])
            All_numble_10 = All_numble_10.rename(columns={'key_0':'index'})
            pass
        pass
    All_numble_10 = All_numble_10[['index','level_0_x','level_0_y']]#保留所需信息
    All_numble_10.columns = ['书类','2014','2015','2016','2017']#设置列名
    All_numble_10 = All_numble_10[:10]#截取前十
    All_numble_10_last = pd.DataFrame()#初始化数据
    for Year in Year_All:
        Sort = All_numble_10[Year].sort_values(na_position='last')
        Sort = Sort.reset_index()
        Sort.index = Sort.index +1
        Sort = Sort.reset_index()
        Sort = Sort[['level_0','index']]
        if Year == '2014':
            All_numble_10_last = Sort
        else:
            All_numble_10_last = pd.merge(All_numble_10_last,Sort,how = 'left',left_on = All_numble_10_last['index'],right_on = Sort['index'])
            All_numble_10_last = All_numble_10_last.rename(columns={'key_0':'index'})
    All_numble_10_last = All_numble_10_last[['level_0_x','level_0_y']]
    All_numble_10_last.columns = ['2014','2015','2016','2017']
    All_numble_10 = All_numble_10.set_index('书类')#将书类列设置为索引
    All_numble_10_last.index = All_numble_10.index
    All_numble_10_last = All_numble_10_last.T#将表格转置
    mpl.rcParams['font.sans-serif'] = ['KaiTi']#设置字体
    fig, ax = plt.subplots(figsize=(5,8))#设置画布，设置画布大小，宽5高8
    ax.spines['right'].set_visible(False)#取消右边框
    ax.spines['top'].set_visible(False)#取消左边框
    for name, value in All_numble_10_last[3:4].max().items():#传入最后一次值（即2017年排名信息）作为图例位置信息
        plt.text(3.25,value, name, size=10, ha='left', va='center')
    plt.plot(All_numble_10_last,'-o')#传入画图数据，设置图像节点符号为小圆点
    plt.ylim([11,0])#设置y轴限度
    plt.yticks([i for i in range(1,11)])#设置y轴刻度
    plt.title('2014-2017计算机院学生借阅书籍类别前十',fontsize=20)#设置标题
    plt.savefig(path+'/代码/12暖暖小组01号彭鑫项目/项目可视化图/任务14.png',dpi=300,bbox_inches='tight')#保存图像在指定位置。设置像素大小。保存完整图像
    plt.show()
    pass


    
# 15、用折线图画出你分配的那个学院的教师2014年，2015年，2016年，2017年所喜欢的排名前10的小说的走向图。
def Task15():
    import pandas as pd
    import matplotlib as mpl
    import matplotlib.pyplot as plt
    Year_All = ['2014','2015','2016','2017']
    All_novel_10 = pd.DataFrame()#初始化数据
    for Year in Year_All:
        Book_xlsx['图书ID'] = Book_xlsx['图书ID'].apply(str)#转型以便拼接
        Da_Teacher_All_Borrow_Year = Da_Teacher_All_Borrow[Da_Teacher_All_Borrow['操作时间'].str.contains(Year)]#截取该年份信息
        Da_Teacher_All_Borrow_Year = pd.merge(Da_Teacher_All_Borrow_Year,Book_xlsx,how='left',left_on=Da_Teacher_All_Borrow_Year['图书ID'],right_on=Book_xlsx['图书ID'])#将借信息与图书目录信息按照图书id进行拼接
        Da_Teacher_All_Borrow_Year_BookNumble = Da_Teacher_All_Borrow_Year[['图书分类号','书名']]#保留所需信息
        Da_Teacher_All_Borrow_Year_BookNumble['图书分类号'] = Da_Teacher_All_Borrow_Year_BookNumble['图书分类号'].apply(str)#转型以便拼接
        Da_Teacher_All_Borrow_Year_novel_Numble = Da_Teacher_All_Borrow_Year_BookNumble.loc[Da_Teacher_All_Borrow_Year_BookNumble['图书分类号'].str.contains('I24')]#模糊搜索小说信息截取保存
        Da_Teacher_All_Borrow_Year_novel_Numble['图书分类号'] = Da_Teacher_All_Borrow_Year_novel_Numble['图书分类号'].map(lambda x:x.split('/')[0])#截取中间斜杠前字符
        Da_Teacher_All_Borrow_Year_novel_Numble = pd.merge(Da_Teacher_All_Borrow_Year_novel_Numble,Novel_List,how='left',left_on=Da_Teacher_All_Borrow_Year_novel_Numble['图书分类号'],right_on = Novel_List['图书分类号'])#将借信息与小说分类表信息按照图书分类号进行拼接
        Da_Teacher_All_Borrow_Year_novel_Numble = Da_Teacher_All_Borrow_Year_novel_Numble['类别'].value_counts()[:50]#对类别进行数量排序，并得到数量前50（大于10即可）的类别
        Da_Teacher_All_Borrow_Year_novel_Numble = Da_Teacher_All_Borrow_Year_novel_Numble.reset_index()#重置索引，将原先索引变为数值列
        Da_Teacher_All_Borrow_Year_novel_Numble.index = Da_Teacher_All_Borrow_Year_novel_Numble.index +1#索引加1，得到排名
        Da_Teacher_All_Borrow_Year_novel_Numble = Da_Teacher_All_Borrow_Year_novel_Numble.reset_index()#再次重置索引，将原先索引变为数值列
        Da_Teacher_All_Borrow_Year_novel_Numble = Da_Teacher_All_Borrow_Year_novel_Numble[['level_0','index']]#保留所需信息
        if Year == '2014' :
            All_novel_10 = Da_Teacher_All_Borrow_Year_novel_Numble#将2014年数据作为基础数据
        else:
            All_novel_10 = pd.merge( All_novel_10,Da_Teacher_All_Borrow_Year_novel_Numble,how = 'outer',left_on=All_novel_10['index'],right_on=Da_Teacher_All_Borrow_Year_novel_Numble['index'])#将其他年份信息和基础信息拼接
            All_novel_10 = All_novel_10.rename(columns={'key_0':'index'})#保留所需信息，以便下次拼接使用
            pass
        pass
    All_novel_10 = All_novel_10[['index','level_0_x','level_0_y']]#保留所需信息
    All_novel_10.columns = ['书类','2014','2015','2016','2017']#设置列名
    All_novel_10 = All_novel_10[:10]#截取前十
    All_novel_10_last = pd.DataFrame()#初始化数据
    for Year in Year_All:
        Sort = All_novel_10[Year].sort_values(na_position='last')
        Sort = Sort.reset_index()
        Sort.index = Sort.index +1
        Sort = Sort.reset_index()
        Sort = Sort[['level_0','index']]
        if Year == '2014':
            All_novel_10_last = Sort
        else:
            All_novel_10_last = pd.merge(All_novel_10_last,Sort,how = 'left',left_on = All_novel_10_last['index'],right_on = Sort['index'])
            All_novel_10_last = All_novel_10_last.rename(columns={'key_0':'index'})
    All_novel_10_last = All_novel_10_last[['level_0_x','level_0_y']]
    All_novel_10_last.columns = ['2014','2015','2016','2017']
    All_novel_10 = All_novel_10.set_index('书类')#将书类列设置为索引
    All_novel_10_last.index = All_novel_10.index
    All_novel_10_last = All_novel_10_last.T#将表格转置
    mpl.rcParams['font.sans-serif'] = ['KaiTi']#设置字体
    fig, ax = plt.subplots(figsize=(5,8))
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    for name, value in All_novel_10_last[3:4].max().items():
        plt.text(3.25,value, name, size=10, ha='left', va='center')
    plt.plot(All_novel_10_last,'-o')
    plt.ylim([11,0])
    plt.yticks([i for i in range(1,11)])
    plt.title('2014-2017计算机院教师借阅书籍小说类别前十',fontsize=20)
    plt.savefig(path+'/代码/12暖暖小组01号彭鑫项目/项目可视化图/任务15.png',dpi=300,bbox_inches='tight')
    plt.show()
    pass

    
# 16、用折线图画出你分配的那个学院的学生2014年，2015年，2016年，2017年所喜欢的排名前10的小说的走向图。(注释同Task15)
def Task16():
    import pandas as pd
    import matplotlib as mpl
    import matplotlib.pyplot as plt
    Year_All = ['2014','2015','2016','2017']
    All_novel_10 = pd.DataFrame()
    for Year in Year_All:
        Book_xlsx['图书ID'] = Book_xlsx['图书ID'].apply(str)
        Da_Student_All_Borrow_Year = Da_Student_All_Borrow[Da_Student_All_Borrow['操作时间'].str.contains(Year)]
        Da_Student_All_Borrow_Year = pd.merge(Da_Student_All_Borrow_Year,Book_xlsx,how='left',left_on=Da_Student_All_Borrow_Year['图书ID'],right_on=Book_xlsx['图书ID'])
        Da_Student_All_Borrow_Year_BookNumble = Da_Student_All_Borrow_Year[['图书分类号','书名']]
        Da_Student_All_Borrow_Year_BookNumble['图书分类号'] = Da_Student_All_Borrow_Year_BookNumble['图书分类号'].apply(str)
        Da_Student_All_Borrow_Year_novel_Numble = Da_Student_All_Borrow_Year_BookNumble.loc[Da_Student_All_Borrow_Year_BookNumble['图书分类号'].str.contains('I24')]
        Da_Student_All_Borrow_Year_novel_Numble['图书分类号'] = Da_Student_All_Borrow_Year_novel_Numble['图书分类号'].map(lambda x:x.split('/')[0])#截取中间斜杠前字符
        Da_Student_All_Borrow_Year_novel_Numble = pd.merge(Da_Student_All_Borrow_Year_novel_Numble,Novel_List,how='left',left_on=Da_Student_All_Borrow_Year_novel_Numble['图书分类号'],right_on = Novel_List['图书分类号'])
        Da_Student_All_Borrow_Year_novel_Numble = Da_Student_All_Borrow_Year_novel_Numble['类别'].value_counts()[:len(Da_Student_All_Borrow_Year_novel_Numble)]
        Da_Student_All_Borrow_Year_novel_Numble = Da_Student_All_Borrow_Year_novel_Numble.reset_index()
        Da_Student_All_Borrow_Year_novel_Numble.index = Da_Student_All_Borrow_Year_novel_Numble.index +1
        Da_Student_All_Borrow_Year_novel_Numble = Da_Student_All_Borrow_Year_novel_Numble.reset_index()
        Da_Student_All_Borrow_Year_novel_Numble = Da_Student_All_Borrow_Year_novel_Numble[['level_0','index']]
        if Year == '2014' :
            All_novel_10 = Da_Student_All_Borrow_Year_novel_Numble
        else:
            All_novel_10 = pd.merge( All_novel_10,Da_Student_All_Borrow_Year_novel_Numble,how = 'outer',left_on=All_novel_10['index'],right_on=Da_Student_All_Borrow_Year_novel_Numble['index'])
            All_novel_10 = All_novel_10.rename(columns={'key_0':'index'})
            pass
        pass
    All_novel_10 = All_novel_10[['index','level_0_x','level_0_y']]#保留所需信息
    All_novel_10.columns = ['书类','2014','2015','2016','2017']#设置列名
    All_novel_10 = All_novel_10[:10]#截取前十
    All_novel_10_last = pd.DataFrame()#初始化数据
    for Year in Year_All:
        Sort = All_novel_10[Year].sort_values(na_position='last')
        Sort = Sort.reset_index()
        Sort.index = Sort.index +1
        Sort = Sort.reset_index()
        Sort = Sort[['level_0','index']]
        if Year == '2014':
            All_novel_10_last = Sort
        else:
            All_novel_10_last = pd.merge(All_novel_10_last,Sort,how = 'left',left_on = All_novel_10_last['index'],right_on = Sort['index'])
            All_novel_10_last = All_novel_10_last.rename(columns={'key_0':'index'})
    All_novel_10_last = All_novel_10_last[['level_0_x','level_0_y']]
    All_novel_10_last.columns = ['2014','2015','2016','2017']
    All_novel_10 = All_novel_10.set_index('书类')#将书类列设置为索引
    All_novel_10_last.index = All_novel_10.index
    All_novel_10_last = All_novel_10_last.T#将表格转置
    mpl.rcParams['font.sans-serif'] = ['KaiTi']#设置字体
    fig, ax = plt.subplots(figsize=(5,8))
    ax.spines['right'].set_visible(False)
    ax.spines['top'].set_visible(False)
    for name, value in All_novel_10_last[3:4].max().items():
        plt.text(3.25,value, name, size=10, ha='left', va='center')
    plt.plot(All_novel_10_last,'-o')
    plt.ylim([11,0])
    plt.yticks([i for i in range(1,11)])
    plt.title('2014-2017计算机院学生借阅书籍小说类别前十',fontsize=20)
    plt.savefig(path+'/代码/12暖暖小组01号彭鑫项目/项目可视化图/任务16.png',dpi=300,bbox_inches='tight')
    plt.show()
    pass


#利用stopwords停词库，用jieba分词，并制作词云
import csv
from wordcloud import WordCloud
import numpy as np
from PIL import Image
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
#自定义停词库
stopwords = pd.read_csv(path + '/代码/12暖暖小组01号彭鑫项目/原始数据/stopwords.txt',header= None,names=['stopwords'] ,encoding='UTF-8',quoting=csv.QUOTE_NONE)#读取停词库信息
stopwords = stopwords['stopwords']#将DataFrame表格转变成Series类型
stopwords = stopwords.tolist()#将Series类型转变成list类型
for word in iter(stopwords):#遍历停词库
    jieba.add_word(word)#加入各个停止词
Book_xlsx['图书ID'] = Book_xlsx['图书ID'].apply(str)#转型以便拼接

#教师数据
def jieba_Ter_Pic():
    Da_Teacher_All_Borrow_jieba = pd.merge(Da_Teacher_All_Borrow,Book_xlsx,how='left',left_on=Da_Teacher_All_Borrow['图书ID'],right_on=Book_xlsx['图书ID'])#将借信息与图书目录信息按照图书ID进行拼接
    Da_Teacher_All_Borrow_jieba = Da_Teacher_All_Borrow_jieba['书名']#保留所需信息
    Da_Teacher_All_Borrow_jieba = Da_Teacher_All_Borrow_jieba.apply(str)#转型以便拼接
    Da_Teacher_All_Borrow_jieba = Da_Teacher_All_Borrow_jieba.tolist()#将Series类型转变成list类型
    Da_Teacher_All_Borrow_jieba = ''.join(Da_Teacher_All_Borrow_jieba)#将list拼接成一个字符串
    words = list(jieba.cut(Da_Teacher_All_Borrow_jieba))#利用jieba停词，并将所有结果制成list
    words = ' '.join(words)#将list所有元素用空格拼接成一个字符串
    image = np.array(Image.open(path + '/代码/12暖暖小组01号彭鑫项目/原始数据/Teacher.jpg'))#导入蒙版图
    wordcloud = WordCloud(font_path=r'C:\Windows\Fonts\msyh.ttc',  # 调用系统自带字体(微软雅黑)    
                          background_color='white',  # 背景色
                          width = 20000, #输出的画布宽度
                          height = 10000, #输出的画布高度
                          max_words=200,  # 最大显示单词数
                          max_font_size=60,  # 频率最大单词字体大小
                          mask = image,#设置图片遮罩
                          scale=1#按照比例进行放大画布
                          ).generate(words)
    plt.imshow(wordcloud, interpolation='bilinear')#传入数值，选择图像插值方式（双线性插值）
    plt.axis('off')#关闭标尺
    plt.savefig(path+'/代码/12暖暖小组01号彭鑫项目/项目可视化图/wordcloud_Teacher.png',dpi=300,bbox_inches='tight')
    plt.show()#展示图例
    pass

#学生数据
def jieba_Stu_Pic():
    Da_Student_All_Borrow_jieba = pd.merge(Da_Student_All_Borrow,Book_xlsx,how='left',left_on=Da_Student_All_Borrow['图书ID'],right_on=Book_xlsx['图书ID'])#将借信息与图书目录信息按照图书ID进行拼接
    Da_Student_All_Borrow_jieba = Da_Student_All_Borrow_jieba['书名']#保留所需信息
    Da_Student_All_Borrow_jieba = Da_Student_All_Borrow_jieba.apply(str)#转型以便拼接
    Da_Student_All_Borrow_jieba = Da_Student_All_Borrow_jieba.tolist()#将Series类型转变成list类型
    Da_Student_All_Borrow_jieba = ''.join(Da_Student_All_Borrow_jieba)#将list拼接成一个字符串
    words = list(jieba.cut(Da_Student_All_Borrow_jieba))#利用jieba停词，并将所有结果制成list
    words = ' '.join(words)#将list所有元素用空格拼接成一个字符串
    image = np.array(Image.open(path + '/代码/12暖暖小组01号彭鑫项目/原始数据/Student.jpg'))#导入蒙版图
    wordcloud = WordCloud(font_path=r'C:\Windows\Fonts\msyh.ttc',  # 调用系统自带字体(微软雅黑)
                          background_color='white',  # 背景色
                          width = 20000, #输出的画布宽度
                          height = 10000, #输出的画布高度
                          max_words=200,  # 最大显示单词数
                          max_font_size=60,  # 频率最大单词字体大小
                          mask = image,#设置图片遮罩
                          scale=1#按照比例进行放大画布
                          ).generate(words)
    plt.imshow(wordcloud, interpolation='bilinear')#传入数值，选择图像插值方式（双线性插值）
    plt.axis('off')#关闭标尺
    plt.savefig(path+'/代码/12暖暖小组01号彭鑫项目/项目可视化图/wordcloud_Student.png',dpi=300,bbox_inches='tight')
    plt.show()#展示图例
    pass

#运行所需任务
Task5()
Task6()
Task7()
Task8()
Task9()
Task10()
Task11()
Task12()
Task13()
Task14()
Task15()
Task16()
jieba_Ter_Pic()
jieba_Stu_Pic()


