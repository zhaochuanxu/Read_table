import  xlrd
import  numpy as np
import pandas as pd
import xlwt
import warnings
warnings.filterwarnings("ignore")

import os
import traceback


export_start = 1
export_end=2

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
     #print(filename)
    if os.path.exists("../file/%s.xls"%(filename)):
        #print("%s存在"%(filename))

        #文件存在 执行下边操作
        data = pd.read_excel("../file/%s.xls"%(filename), sheet_name = 0, skiprows= 5,skipfooter = 2)#调过末尾两行
        s = "%d-%s"%(i,temp) #备注改
        data['备注(被引用年份)']=s

        #写入数据
        data.to_excel('../new/%s_new.xls'%(filename),sheet_name='data' ,index=None)  #sheng成的一个新表
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
    # 名称转换接口print(filename)
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
        export_start=start_year
        export_end=start_month
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

    #如果部存在文件就返回1 存在就返回一个data
    #print("我是%d年%d月数据"%(i,j))
    #print(data_1)
    #print(data_1)
    data= pd.read_excel('../result/%s-%s_data.xls'%(die_start,die_end), sheet_name = 'data', skiprows= 0)  #从新生成的第一行开始读取
    if isinstance(data_1,int)== True : #如果表不存在
        return


    else: #表存在:

        #遍历这个表
        #print(data_1) data_1代表new里边的表      检查数据接口
        if i == start_year and j ==start_month:
            t = 3
        else:
            for m in range(0,data_1.shape[0]):   #遍历老 ij表的m行
                #print(m)    测试输出行
                #遍历新表
                data_1_temp = data_1["Accession Number"][m]

                #print("-----")
                #print(temp)
                #print("------")
                for k in range(0,data.shape[0]):
                    #print(data["Accession Number"][k]) # 测试输特殊字段
                    if data_1_temp == data["Accession Number"][k]:
                       # print("找到个一样的了")
                        s=""
                        #print(type(data["Accession Number"][k]))
                        #s=data["Accession Number"][k]+"test"
                        # print(s)

                        if j<10:#如果小于10月
                            s = "  %d-0%d  "%(i,j)
                        else:
                            s="  %d-%d  "%(i,j)

                        #print(type(data['备注(被引用年份)'][k]))
                        s=data['备注(被引用年份)'][k]+s
                        #print(type(s))
                       # print(s)
                        # print(data_temp)
                        data['备注(被引用年份)'][k]=s
                        #print(data['备注(被引用年份)'][k])
                        #print(k)
                        break





                    else: #没有比对成功的话:
                        if k == data.shape[0]-1:
                            d = pd.DataFrame(data_1, index=[m])#读取的老表的一行    老表的a行
                            # print(d)   检测出不一样的
                            data = data.append(d, ignore_index=True)  #追加一个数据
                        # print(k)



                        else:
                           # print("向下可能还会有哦")
                            continue
                        #写入


                data.to_excel('../result/%s-%s_data.xls'%(die_start,die_end),sheet_name='data',index=None)  #sheng成的一个新表






















if __name__=="__main__":
    while 1:
        try:
            print("提示：")
            print("1：请将需要操作的文件统一命名为m_()   括号内内容为月份，如3月就是m_3"
                      "并存放在file目录:\n")
            print("2:提取的结果文件放置在result文件夹下边，如果提取的为1-3月的，文件名字就为1_3_data.xls")
            print("\n\n")
            print("请输入：你要统计的月份（默认是2000.1—2019.12）")
            start_year=int(2000)
            end_year=int(2019)
            start_month=int(1)
            end_month=int(12)
            judge()
            print("初始化完成")
            start_year = int(input("请输入开始年份"))
            start_month = int(input("请输入开始月份"))
            end_year = int(input("请输入结束年份"))
            end_month=  int(input("请输入结束月份"))
            if start_year>end_year:
                print("输出错误")
            #if start_month>end_month:
                #print("输出错误")
            #print(type(start_month))
            else:

                for i in range(start_year,end_year+1):
                    #如果开始年等于结束年
                    if start_year == end_year:
                        #直接遍历所有月份
                        for j in range(start_month,end_month+1):
                            read_初始化(i,j)



                    #如果开始年不等于结束年
                    else:
                        #对待第一年效果不同·如果第一年
                        if i==start_year:
                            #第一年遍历 开始月份到12月
                            for j in range(start_month,12):
                                read_初始化(i,j)



                        #如果当前年份是最后一年
                        elif i==end_year:
                            #直接遍历1月到结束那一月
                            for j in range(1,end_month+1):
                                read_初始化(i,j)



                        #如果当前年份是卡在中间的年份
                        else:
                            #遍历中间所有月
                            for j in range(1,12):
                                read_初始化(i,j)

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

                print("完成")
            #print(a)
                print("任意键退出")
                i=input()

                if i != 99999:
                    break



        except:
            print ("内容写入文件成功")
            #traceback.print_exc()
            #
            print("操作错误")













