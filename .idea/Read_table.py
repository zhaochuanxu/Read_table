import  xlrd
import  numpy as np
import pandas as pd
import xlwt
import warnings
warnings.filterwarnings("ignore")

import os
import traceback

a = ['0','0','0','0','0','0','0','0','0','0','0','0','0']
#判断文件是否存在
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



#初始化所有表
def read_初始化(i):
    #print(type(os.getcwd())) #获取当前目录
   # os.makedirs("../new")
   # if(print(os.path.exists("../new"))==false): #判断一个目录是否存在
        #os.makedirs("../new")
    if os.path.exists("../file/m_%d.xls"%(i)):
        a[i]='1'
        data = pd.read_excel("../file/m_%d.xls"%(i), sheet_name = 0, skiprows= 5,skipfooter = 2)#调过末尾两行
        s='2018-%d'%(i)
        #print(type(s))
        #print(s)
        data['备注(被引用年份)']=s
        # print(data['备注(被引用年份)'])

        #将数据读取出来 写入新表
        data.to_excel('../new/m_%d_new.xls'%(i),sheet_name='data',index=None)  #sheng成的一个新表
        return data
        #print("未创建名为new文件夹")
    else:
        a[i]=0
        return



#传过去i代表读第i个月信息
def read(i):  #读取新建的12个表
    if os.path.exists("../new/m_%d_new.xls"%(i)):
        data = pd.read_excel("../new/m_%d_new.xls"% (i))
        return data

    else:

        return



#创建一个新表 从第i个开始那么 就从第i个读取
def Creat_newExcel(start,end):
     #举个例子 只读取第11表

    #data = pd.read_excel("../m_%d_new.xls"%(start))  #跳过末尾两行


    #print(data)
    #Operate(data,start,end)
      #data['备注'][3]='已经被引用'  测试
      #for i in range(0,)

    #data['备注']='  2018.%d '%start #向 DataFrame 添加一列，该列为同一值   初始化  先不需要初始化
    if a[start]== '1':
        data = read(start)

        #新生成的表 在result里边
        data.to_excel('../result/%d-%d_data.xls'%(start,end),sheet_name='data,',index=None)  #sheng成的一个新表


        data= pd.read_excel('../result/%d-%d_data.xls'%(start,end), sheet_name = 0, skiprows= 0)  #从新生成的第一行开始读取

        Operate(data,start,end)

        return data
    else:
        Creat_newExcel(start+1,end)






#对新生成的表进行操作

#从第start个表到第end个表每个表的每个信息进行比对
def Operate(data,start,end):

    try:
        #对每个老表i个 表进行循环比较
        for i in range(start+1,end+1):
            if a[i] =='1': #如果年份
                 #acd=read(i)
                 #print(acd)
                data_1=read(i)
                 #   print(data_1) 测试
                #对每行（每个那老表）进行比对
                for j in range(0,data_1.shape[0]):
                                     #print(data_1["Accession Number"][j])  #测试
                    temp = data_1["Accession Number"][j]
                    #对新表每行遍历查找
                    num = 1 #计时 对新表每行始计数
                    flage=0
                    for k in  range(0,data.shape[0]):
                        if temp  == data["Accession Number"][k]:

                                #flage=flage + 1   #第一次查询到了就为1
                            s="  2018-%d  "%(i)
                                #print(type(data['备注(被引用年份)'][k]))
                            data_temp=""
                            # print(data_temp)
                            data_temp=data["备注(被引用年份)"][k]
                            # print(data_temp)
                            data["备注(被引用年份)"][k]=""
                            data["备注(被引用年份)"][k]=data_temp+s
                            #print(type(data["备注(被引用年份)"][k]))
                                #print(data["备注(被引用年份)"][k]) 追加
                            #print("3")
                            break

                        else:# 就是没有 检查是否遍历完一遍
                            if num!=data.shape[0]:
                                num=num+1
                            else:   #一遍了

                                d = pd.DataFrame(data_1, index=[j])#读取的老表的一行
                                data = data.append(d, ignore_index=True)  #追加一个数据


                data.to_excel('../result/%d-%d_data.xls'%(start,end),sheet_name='data',index=None)  #sheng成的一个新表







    except:
        traceback.print_exc()
        print ("内容写入文件成功")





















#主函数
if __name__=="__main__":
    #测试read函数
    #data=read(11)
    #print(data)
    #read(1)
    # Creat_newExcel(1,2)
    #data = pd.read_excel("../m_4_new.xls")
        #for i in range(1,13):#初始化一年中所有的表
         #   read_初始化(i)
      #  print("初始化完成")
    i = 788
    while(1):
        try:
            judge()

            print("初始化完成")
            start_year=2000
            end_year=2019

            start_mon = 1
            end = 12
            print("提示：")
            print("1：请将需要操作的文件统一命名为m_()   括号内内容为月份，如3月就是m_3"
                  "并存放在file目录:\n")
            print("2:提取的结果文件放置在result文件夹下边，如果提取的为1-3月的，文件名字就为1_3_data.xls")
            print("\n\n")
            print("请输入：你要统计的月份（默认是1—12月）")
            start=int(input("开始月份： "))
            end = int(input("结束月份： "))
            if start>end:
                print("操作错误：开始月份应该小于结束月份")
            for i in range(start,end+1):#初始化一年中所有的表
                read_初始化(i)
            Creat_newExcel(start,end)
            print("完成")
            #print(a)
            print("任意键退出")
            i=input()

            if i != 99999:
                break

        except:
            traceback.print_exc()
            print("操作错误：请重试")














