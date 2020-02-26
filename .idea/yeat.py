import  xlrd
import  numpy as np
import pandas as pd
import xlwt
import warnings
warnings.filterwarnings("ignore")

import os
import traceback

#判断文件夹是否存在
def judge():
     if(os.path.exists("../new")==False): #判断一个目录是否存在
        os.makedirs("../new")
        print("正在创建new文件夹")
     if(os.path.exists("../file")==False): #判断一个目录是否存在
        os.makedirs("../file")
        print("正在创建file文件夹")
     if(os.path.exists("../result")==False): #判断一个目录是否存在
        os.makedirs("../result")
        print("正在创建file文件夹")
     print("初始化成功")

#初始化构建新表 传过来的参数就是哪一年那个表    如果没有及跳过直接要return
def read_初始化(i,j):
    if j<10:
        temp = '0%d'%(j)
        # i 代表年份 temp代表月份   生成201808
        filename = "%d%s"%(i,temp)
    else:
        filename="%d%d"%(i,j)
    #print(temp)
    if os.path.exists("../file/%s.xls"%(filename)):
        #print("%s存在"%(filename))

        #文件存在 执行下边操作
        data = pd.read_excel("../file/%s.xls"%(filename), sheet_name = 0, skiprows= 5,skipfooter = 2)#调过末尾两行
        s = "%d.%s"%(i,temp) #备注改
        data['备注(被引用年份)']=s

        #写入数据
        data.to_excel('../new/%s_new.xls'%(filename),sheet_name=0 ,index=None)  #sheng成的一个新表
        return data


    else:
        #print("文件不存在")
        #print("%s"%(filename))
        return #文件不存在返回






#传过去i代表读第i个年信息 j代表月  information
def read(i,j):  #读取新建的more个表
    if j<10:
        temp = '0%d'%(j)
        # i 代表年份 temp代表月份   生成201808
        filename = "%d%s"%(i,temp)
    else:
         filename = "%d%d"%(i,j)
    #print(temp)
    if os.path.exists("../new/%s_new.xls"%(filename)):
        #如果信息存在
        data = pd.read_excel("../new/%s_new.xls"% (filename))
        #print(data)
        return data

    else:
        #表不存在

        return 1



#创建一个新表 从第i个开始那么 就从第i个读取
def Creat_newExcel(start_year,end_year,start_month,end_month,die_start,die_end):
    print("本次开始年份")
    print(start_year)
    print("本次结束年份")
    print(end_year)
    print("本次开始月份")
    print(start_month)
    print("本次结束月份")
    print(end_month)
    print("固定开始时间")
    print(die_start)
    print("固定结束时间")
    print(die_end)
    print("---------------")
#调整开始年月份输入
    if start_month<10:
        temp = '0%d'%(start_month)
        # i 代表年份 temp代表月份   生成201808
        #print(temp)
        start = "%d%s"%(start_year,temp)
        #die_start="%d%s"%(start_year,temp)
    else:
        start = "%d%d"%(start_year,start_month)
        #die_start="%d%s"%(start_year,start_month)
   # print("我是开始年月份"+start)
    #print("我是固定死开始年月份"+die_start)


#调整结束年月份输入
    if end_month<10:
        temp_1 = '0%d'%(end_month)
        #print(temp)
        end="%d%s"% (end_year,temp_1)
        #die_end="%d%s"%(end_year,temp) #死的固定死了
    else:
        end="%d%d"% (end_year,end_month)
       # die_end="%d%s"%(end_year,end_month)

   # print("我是结束年月份"+end)
    #print("我是固定死结束年月份"+die_end)

    #如果初始表存在     读取数据
    if os.path.exists("../new/%s_new.xls"%(start)):
        data = read(start_year,start_month)
        #新生成的表 在result里边
        data.to_excel('../result/%s-%s_data.xls'%(die_start,die_end),sheet_name= 'data',index=None)  #sheng成的一个新表


        data= pd.read_excel('../result/%s-%s_data.xls'%(die_start,die_end), sheet_name = 'data', skiprows= 0)  #从新生成的第一行开始读取
        print(data)
        Operate(start_year,end_year,start_month,end_month,die_start,die_end)
        return data



    #如果表不存在   不断递归寻找
    else:
        if start_year==end_year and start_month == end_month:
            print("为什么一个文件没有呢")
            return

        else:
            if start_month !=12: #如果当前月份不是12月 就递归下一个
                Creat_newExcel(start_year,end_year,start_month+1,end_month,die_start,die_end)
            else: #如果当前是12月 那么开始月份就归 1
                 Creat_newExcel(start_year+1,end_year,1,end_month,die_start,die_end)  #12月的话那么就月份归1 年加 1







#这个跳转是由Creat_newExcel 传递过来的参数
                #参数说明 start_year 是查找的（开始日期和结束日期之间）第一个存在的表的名字年数  end_year 不变 start_month是查找的存在的第一个表的月数 end_month

 #从第start_year表到第end个表每个表的每个信息进行比对
def Operate(start_year,end_year,start_month,end_month,die_start,die_end):
    #首先对传过来的数据进行加工


    #先对年份开始遍历
    for i in range(start_year,end_year+1):
     #对这一年的月份进行遍历

        #如果开始年等于结束年
        if start_year == end_year:
            #直接遍历所有月份
            for j in range(start_month,end_month+1):
                add(i,j,start_year,end_year,start_month,end_month,die_start,die_end)


        #如果开始年不等于结束年
        else:
            #对待第一年效果不同·如果第一年
            if i==start_year:
                #第一年遍历 开始月份到12月
                for j in range(start_month,12):
                    add(i,j,start_year,end_year,start_month,end_month,die_start,die_end)


            #如果当前年份是最后一年
            elif i==end_year:
                #直接遍历1月到结束那一月
                for j in range(1,end_month+1):
                    add(i,j,start_year,end_year,start_month,end_month,die_start,die_end)


            #如果当前年份是卡在中间的年份
            else:
                #遍历中间所有月
                for j in range(1,12):
                    add(i,j,start_year,end_year,start_month,end_month,die_start,die_end)


















#对具体的表操作函数
def add(i,j,start_year,end_year,start_month,end_month,die_start,die_end):
#参数说明 i:当前年份    j：当前月份      start_year：具体有数据的年数 start_month:具体有数据的月数  die_start,die_end作用是用来不断读取新的表

#重点读取数据之前 要先看看有没有这个表
    data_1=read(i,j) #读取i年j月的数据
    #print("我是%d年%d月的"%(i,j))
    #print(data_1)

    #i 代表年 j代表月份
        #输出表的行数

    #新生成的表
    data= pd.read_excel('../result/%s-%s_data.xls'%(die_start,die_end), sheet_name = 'data', skiprows= 0)  #从新生成的第一行开始读取
    print(data)

    if isinstance(data_1,int)== True : #如果表不存在
        return

        #print("我是%d年%d月的表 我不存在"% (i,j))



    else:
        #print(type(data_1))

        for a in range(0,data_1.shape[0]):#对应
            temp = data_1["Accession Number"][a]   #i年j月 每行
            #开始遍历老表 一一比对
            num = 1 #每次换个月表就置一 这样记录新表遍历所在的行数
            for k in range(0,data.shape[0]):
                #print(data_1.shape[0])
                if temp == data_1["Accession Number"][j]: #如果比对成功     记住这时候是第k行
                    s_1=""
                    #处理信息
                    if j<10:#如果小于10月
                        s_1 = "%d.0%d"%(i,j)
                    else:
                        s_1="%d%d"%(i,j)

                    data_temp="" #置空
                    data_temp=data["备注(被引用年份)"][k]
                    print(data_temp)
                    #data["备注(被引用年份)"][k]=data_temp+s_1
                    break

                else: #如果没有比对成功

                    #如果还没有遍历到新表最下那行
                    if num!=data.shape[0]:
                        num = num+1
                    #不然的话 就是已经到最后了表 追加一个数据
                    else:
                        d = pd.DataFrame(data_1, index=[a])#读取的老表的一行    老表的a行
                        data = data.append(d, ignore_index=True)  #追加一个数据


        #这时候就要注入数据了 新的数据  每当遍历完一张 i j 表后注入
        data.to_excel('../result/%s-%s_data.xls'%(die_start,die_end),sheet_name='data' , index=None)  #sheng成的一个新表








                #temp = data["Accession Number"][k]

               # print("%d年%d月"% (i,j)+temp)








if __name__=="__main__":
    start_year=int(2000)
    end_year=int(2019)
    start_month=int(1)
    end_month=int(12)
    judge()
    print("初始化完成")
    start_year = int(input("请输入开始年份"))
    end_year = int(input("请输入结束年份"))
    start_month = int(input("请输入开始月份"))
    end_month=  int(input("请输入结束月份"))
    if start_year>end_year:
        print("输出错误")
    #if start_month>end_month:
        #print("输出错误")
    for i in range(start_year,end_year+1): #初始化年份
        for j in range(start_month,end_month+1):#初始化月份
            read_初始化(i,j)
            read(i,j)

#g固定开始结束年月份
    if start_month<10:
        temp = '0%d'%(start_month)
        # i 代表年份 temp代表月份   生成201808
        #print(temp)
        die_start="%d%s"%(start_year,temp)
    else:
        die_start="%d%s"%(start_year,start_month)
    #print("我是固定死开始年月份"+die_start)

    if end_month<10:
        temp = '0%d'%(end_month)
        #print(temp)
        die_end="%d%s"%(end_year,temp) #死的固定死了
    else:
        die_end="%d%s"%(end_year,end_month)

  #  print("我是结束年月份"+end)
    #print("我是固定死结束年月份"+die_end)
    print("本次开始年份")
    print(start_year)
    print("本次结束年份")
    print(end_year)
    print("本次开始月份")
    print(start_month)
    print("本次结束月份")
    print(end_month)
    print("固定开始时间")
    print(die_start)
    print("固定结束时间")
    print(die_end)
    print("---------------")




    Creat_newExcel(start_year,end_year,start_month,end_month,die_start,die_end)







