import  xlrd
import  numpy as np
import pandas as pd
import xlwt


#初始化所有表
def read_初始化(i):
    data = pd.read_excel("../m_%d.xls"%(i), sheet_name = 0, skiprows= 5,skipfooter = 2)#调过末尾两行
    s='2019-%d'%(i)
    #print(type(s))
    #print(s)
    data['备注(被引用年份)']=s
   # print(data['备注(被引用年份)'])

     #将数据读取出来 写入新表
    data.to_excel('../new/m_%d_new.xls'%(i),sheet_name='data,',index=None)  #sheng成的一个新表
    return data


#传过去i代表读第i个月信息
def read(i):  #读取新建的12个表
    data = pd.read_excel("../new/m_%d_new.xls"% (i))
    return data



#创建一个新表 从第i个开始那么 就从第i个读取
def Creat_newExcel(start,end):
     #举个例子 只读取第11表

    #data = pd.read_excel("../m_%d_new.xls"%(start))  #跳过末尾两行


    #print(data)
    #Operate(data,start,end)
      #data['备注'][3]='已经被引用'  测试
      #for i in range(0,)

    #data['备注']='  2018.%d '%start #向 DataFrame 添加一列，该列为同一值   初始化  先不需要初始化

    data = read(start)

    #新生成的表 在result里边
    data.to_excel('../result/%d-%d_data.xls'%(start,end),sheet_name='data,',index=None)  #sheng成的一个新表


    data= pd.read_excel('../result/%d-%d_data.xls'%(start,end), sheet_name = 0, skiprows= 0)  #从新生成的第一行开始读取

    Operate(data,start,end)

    return data





#对新生成的表进行操作

#从第start个表到第end个表每个表的每个信息进行比对
def Operate(data,start,end):

    #对每个老表i个 表进行循环比较
    for i in range(start+1,end+1):
        data_1=  pd.read_excel("../new/m_%d_new.xls"%(i))
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










        data.to_excel('../result/%d-%d_data.xls'%(start,end),sheet_name='data,',index=None)  #sheng成的一个新表






















#主函数
if __name__=="__main__":
    #测试read函数
    #data=read(11)
    #print(data)
    #read(1)
    # Creat_newExcel(1,2)
    #data = pd.read_excel("../m_4_new.xls")
    for i in range(1,13):#初始化一年中所有的表
       read_初始化(i)

    start = 1
    end = 12
    print("请输入：你要统计的月份（默认是1—12月）")
    start=int(input("开始月份"))
    end = int(input("结束月份"))
    Creat_newExcel(start,end)
    print("完成")














