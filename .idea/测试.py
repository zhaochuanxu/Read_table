import  xlrd
import  numpy as np
import pandas as pd
import xlwt
import warnings
warnings.filterwarnings("ignore")

def test(i):
    data = pd.read_excel("../m_%s.xls"%(i), sheet_name = 0, skiprows= 5,skipfooter = 2)  #跳过末尾两行
        # enan = ['组长','组员','组员','组员','组员','组员','组员','组员']
        # data['备注'] = enan
    data['备注']='该未被初始化'#向 DataFrame 添加一列，该列为同一值
        #data['备注'][3]='已经被引用'  测试
    data.to_excel('../data%s,%s.xls'%('star','teh'),sheet_name='data,',index=None)

if __name__=="__main__":
    test(11)

     data_1 = data_1.append(temp, ignore_index=True)  #追加一个数据
        print(data_1)



start = 1
    end = 12
    print("请输入：你要统计的月份（默认是1—12月）")
    start=int(input("开始月份"))
    end = int(input("结束月份"))
    Creat_newExcel(start,end)

       start = 1
    end = 12
    print("请输入：你要统计的月份（默认是1—12月）")
    start=int(input("开始月份"))
    end = int(input("结束月份"))
    Creat_newExcel(start,end)

     start = 1
    end = 12
    print("请输入：你要统计的月份（默认是1—12月）")
    start=int(input("开始月份"))
    end = int(input("结束月份"))
    Creat_newExcel(start,end)




  #data= pd.read_excel('../%d_%d_data.xls'%(start,end), sheet_name = 0, skiprows= 0)  #从新生成的第一行开始读取
    #temp = data.iloc[-1]#读取最后一行
    # print(temp)
    #print("最后————————一行")

    # 重点比对措施    用i个表一行和新表所有j行比对
    #i个表
    for i in range(start+1,end+1):
        data_1=read(i)     #第i个表

            #遍历i表所有行
        #data_1.shape[0]行  选一行
        for j in range (0,data_1.shape[0]): #行数

            old_table = data_1["Accession Number"][j]
            #print(old_table)  #输出每行
            num = 1 #每个表固定的查询次数
            for k  in range(0,data.shape[0]):
                #print(data["Accession Number"][data.shape[0]-1])


                #一共25个 下表是0 - 24  实际是2  - 26
                #如果老表中找的这一行 出现在新表中
                if old_table == data["Accession Number"][k]:
                    s = "   2018.%d "% i
                    data['备注'][k]=  data['备注'][k]+s
                    #print(type( data['备注'][k]))

                   # print(data['备注'][k])
                    #print(data)
                    break
                else: #追加功能
                    #d=data_1(index=[j])
                    #print("")
                    #print(data.shape[0])
                    if num!=data.shape[0]:
                        num=num+1
                    else:  #老表中j行


                        h = "我是第%d行"% j

                        s = "   2018.%d "% i
                        data_1['备注'][j]=data['备注'][j]+s

                        d = pd.DataFrame(data_1, index=[j])#读取的老表的一行


                        print('我就要添加了')
                        #print(h)
                      # print(d)

                        data = data.append(d, ignore_index=True)  #追加一个数据
                        data.to_excel('../%d-%d_data.xls'%(start,end),sheet_name='data,',index=None)  #sheng成的一个新表

                        print('添加成功')
                        print(data)
                        #print(type(old_table))





data_2=""   #先记住原来的
                    data_2=data["备注(被引用年份)"][k]
                    data["备注(被引用年份)"][k]=data_2+s

