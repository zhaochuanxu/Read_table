import  xlrd
import  numpy as np
import pandas as pd
import xlwt


#传过去i代表读第i个月信息
def read(i):
    data = pd.read_excel("../m_%d.xls"%(i), sheet_name = 0, skiprows= 5,skipfooter = 2)  #跳过末尾两行
    return data





#创建一个新表 从第i个开始那么 就从第i个读取
def Creat_newExcel(start,end):
     #举个例子 只读取第11表
    data = pd.read_excel("../m_%d.xls"%(start), sheet_name = 0, skiprows= 5,skipfooter = 2)  #跳过末尾两行
    data['备注']=''#向 DataFrame 添加一列，该列为同一值
    #data['备注'][3]='已经被引用'  测试
    data.to_excel('../%d_%d_data.xls'%(start,end),sheet_name='data,',index=None)  #sheng成的一个新表
    data= pd.read_excel('../%d_%d_data.xls'%(start,end), sheet_name = 0, skiprows= 0)  #从新生成的第一行开始读取
    print(data)
    Operate(data,start,end)





#对新生成的表进行操作

#从第start个表到第end个表每个表的每个信息进行比对
def Operate(data,start,end):
    #data= pd.read_excel('../%d_%d_data.xls'%(start,end), sheet_name = 0, skiprows= 0)  #从新生成的第一行开始读取
    data['备注'][2]='我已修改'  #data表示新生成的表
    temp = data.iloc[-1]#读取最后一行
    print(temp)

    # 重点比对措施    用i个表一行和新表所有j行比对
    #i个表
    for i in range(start+1,end+1):
        data_1=read(i)     #第i个表

            #遍历i表所有行
        #data_1.shape[0]行  选一行
        for j in range (0,data_1.shape[0]): #行数

            old_table = data_1["Accession Number'"][j]

          #遍历新表所有行
            for k  in range(0,data_shape[0]):
                #如果老表中找的这一行 出现在新表中
                if old_table == data["Accession Number'"][k]:
                    s = "2018.%d "% i
                    data['备注'][k]=  data['备注'][k]+s
                    break



















#主函数
if __name__=="__main__":
    #测试read函数
    #data=read(11)
    #print(data)
    start = 1
    end = 12
    print("请输入：你要统计的月份（默认是1—12月）")
    start=int(input("开始月份"))
    end = int(input("结束月份"))
    Creat_newExcel(start,end)





 #遍历新表所有行
            for k  in range(0,data_shape[0]):
                #如果老表中找的这一行 出现在新表中
                if old_table == data["Accession Number'"][k]:
                    s = "2018.%d "% i
                    data['备注'][k]=  data['备注'][k]+s
                    break
                else:
                    data = data.append(old_table, ignore_index=True)  #追加一个数据
                    print(data)







       Accession Number  ...     备注
0   WOS:000275117900034  ...  被引用月份
1   WOS:000320613800004  ...  被引用月份
2   WOS:000301990600062  ...  被引用月份
3   WOS:000288227800007  ...  被引用月份
4   WOS:000333551800037  ...  被引用月份
5   WOS:000277463600009  ...  被引用月份
6   WOS:000328699400004  ...  被引用月份
7   WOS:000331941300008  ...  被引用月份
8   WOS:000334337500008  ...  被引用月份
9   WOS:000355036800014  ...  被引用月份
10  WOS:000338943900043  ...  被引用月份
11  WOS:000351439200002  ...  被引用月份
12  WOS:000309785000017  ...  被引用月份
13  WOS:000370962900011  ...  被引用月份
14  WOS:000348366100001  ...  被引用月份
15  WOS:000375520700009  ...  被引用月份
16  WOS:000368783600087  ...  被引用月份
17  WOS:000401982100002  ...  被引用月份
18  WOS:000396043900003  ...  被引用月份
19  WOS:000371241100010  ...  被引用月份
20  WOS:000396186600058  ...  被引用月份
21  WOS:000390505200015  ...  被引用月份
22  WOS:000430729100030  ...  被引用月份
23  WOS:000418971900017  ...  被引用月份
24  WOS:000427537400001  ...  被引用月份


Accession Number  ...     备注
0   WOS:000275117900034  ...  被引用月份
1   WOS:000320613800004  ...  被引用月份
2   WOS:000301990600062  ...  被引用月份
3   WOS:000288227800007  ...  被引用月份
4   WOS:000333551800037  ...  被引用月份
5   WOS:000277463600009  ...  被引用月份
6   WOS:000328699400004  ...  被引用月份
7   WOS:000331941300008  ...  被引用月份
8   WOS:000334337500008  ...  被引用月份
9   WOS:000355036800014  ...  被引用月份
10  WOS:000338943900043  ...  被引用月份
11  WOS:000351439200002  ...  被引用月份
12  WOS:000309785000017  ...  被引用月份
13  WOS:000370962900011  ...  被引用月份
14  WOS:000348366100001  ...  被引用月份
15  WOS:000375520700009  ...  被引用月份
16  WOS:000368783600087  ...  被引用月份
17  WOS:000401982100002  ...  被引用月份
18  WOS:000396043900003  ...  被引用月份
19  WOS:000371241100010  ...  被引用月份
20  WOS:000396186600058  ...  被引用月份
21  WOS:000390505200015  ...  被引用月份
22  WOS:000430729100030  ...  被引用月份
23  WOS:000418971900017  ...  被引用月份
24  WOS:000427537400001  ...  被引用月份

[25 rows x 13 columns]



 if a[i]=='1':  #如果文件存在
                acd=read(i)
                print(acd)
                print(i)



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
            else:
                continue






                D:\python\pythonw.exe E:/新建软件/Read_table/.idea/Year_table.py
初始化成功
初始化完成
请输入开始年份2017
请输入结束年份2018
请输入开始月份11
请输入结束月份5
本次开始年份
2017
本次结束年份
2018
本次开始月份
11
本次结束月份
5
固定开始时间
201711
固定结束时间
201805
---------------
本次开始年份
2017
本次结束年份
2018
本次开始月份
12
本次结束月份
2018
固定开始时间
201711
固定结束时间
201805
---------------
本次开始年份
2018
本次结束年份
2018
本次开始月份
1
本次结束月份
2018
固定开始时间
201711
固定结束时间
201805
---------------
本次开始年份
2018
本次结束年份
2018
本次开始月份
2
本次结束月份
2018
固定开始时间
201711
固定结束时间
201805
---------------
本次开始年份
2018
本次结束年份
2018
本次开始月份
3
本次结束月份
2018
固定开始时间
201711
固定结束时间
201805
---------------
                             Accession Number  ... 备注(被引用年份)
0                         WOS:000275117900034  ...   2018.03
1                         WOS:000320613800004  ...   2018.03
2                         WOS:000301990600062  ...   2018.03
3                         WOS:000288227800007  ...   2018.03
4                         WOS:000333551800037  ...   2018.03
5                         WOS:000277463600009  ...   2018.03
6                         WOS:000328699400004  ...   2018.03
7                         WOS:000331941300008  ...   2018.03
8                         WOS:000334337500008  ...   2018.03
9                         WOS:000355036800014  ...   2018.03
10                        WOS:000338943900043  ...   2018.03
11                        WOS:000351439200002  ...   2018.03
12                        WOS:000309785000017  ...   2018.03
13                        WOS:000370962900011  ...   2018.03
14                        WOS:000348366100001  ...   2018.03
15                        WOS:000375520700009  ...   2018.03
16                        WOS:000368783600087  ...   2018.03
17                        WOS:000401982100002  ...   2018.03
18                        WOS:000396043900003  ...   2018.03
19                        WOS:000371241100010  ...   2018.03
20                        WOS:000396186600058  ...   2018.03
21                        WOS:000390505200015  ...   2018.03
22                        WOS:000430729100030  ...   2018.03
23                        WOS:000418971900017  ...   2018.03
24  WOS:0004189719000170000000000000000000000  ...   2018.03
25                        WOS:000427537400001  ...   2018.03

[26 rows x 13 columns]

进程已结束,退出代码0








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







#print("我是%d年%d月的"%(i,j))
    #print(data_1)

    #i 代表年 j代表月份
        #输出表的行数

    #新生成的表
    data= pd.read_excel('../result/%s-%s_data.xls'%(die_start,die_end), sheet_name = 'data', skiprows= 0)  #从新生成的第一行开始读取  原始数据

    #print(data)  #
    #data数据是新生成的表
    if isinstance(data_1,int)== True : #如果表不存在
        return

        #print("我是%d年%d月的表 我不存在"% (i,j))



    else:  #如果表存在
        #print(type(data_1))

        for a in range(0,data_1.shape[0]):#对应
            temp = data_1["Accession Number"][a]   #i年j月 每行
            #开始遍历老表 一一比对
            #num = 1 #每次换个月表就置一 这样记录新表遍历所在的行数
            for k in range(0,data.shape[0]):   #d对新表进行循环
                #print(data_1.shape[0])
                if temp == data["Accession Number"][k]: #如果比对成功     记住这时候是第k行
                    s1=""
                    #处理信息
                    if j<10:#如果小于10月
                        s1 = "%d.0%d"%(i,j)
                    else:
                        s1="%d%d"%(i,j)

                    data_temp="" #置空
                    data_temp=data["备注(被引用年份)"][k]
                    print(data_temp)
                    #data["备注(被引用年份)"][k]=""
                    #data["备注(被引用年份)"][k]=data_temp+s1
                    break

                else: #如果没有比对成功

                    #如果还没有遍历到新表最下那行
                    if k != data.shape[0]-1:
                        #num = num+1
                        continue
                    #不然的话 就是已经到最后了表 追加一个数据
                    else:
                        d = pd.DataFrame(data_1, index=[a])#读取的老表的一行    老表的a行
                        data = data.append(d, ignore_index=True)  #追加一个数据


        #这时候就要注入数据了 新的数据  每当遍历完一张 i j 表后注入
        data.to_excel('../result/%s-%s_data.xls'%(die_start,die_end),sheet_name='data' , index=None)  #sheng成的一个新表








                #temp = data["Accession Number"][k]

               # print("%d年%d月"% (i,j)+temp)







