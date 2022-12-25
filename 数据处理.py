#bug数据不准确请核对编号是否正确，一定是ID、No、T的数据模式，而且要均等，追加了样本编号自查，应该问题不大
import numpy
import xlwt
import xlrd
import time
date=xlrd.open_workbook('D:\\TF\\华西\\课题\\黄老师\\生物学变异\\00成熟运算流程\\数据格式.xls')#读取表格
table=date.sheets()[0]#转存数据
nonr2=numpy.array(table.col_values(0))#获取第一列
n=numpy.sum(nonr2 == '')#获取第一列中不为空的数量
nonr1=table.nrows#表格行数
nonc1=table.ncols#表格列数
r1=nonr1-n-1#表格行数，去掉抬头，即具体测试结果数量
c1=nonc1-1-1-1-1#表格列数-ID-No-T-sex，即项目数
#print(r1,c1)#获取测试与样本数
#print(table.col_values(1)[0])#列，行
IDmax=int(table.col_values(0)[r1])#患者数
IDg=int(r1/IDmax)#同一患者测试数据个数
Nomax=int(table.col_values(1)[r1])#样本数
Nog=int(r1/Nomax)#每个样本测定次数
Tg=int(table.col_values(2)[r1])
Sampg=Nomax/IDmax
#print(Tg)
#print(r1)
if Tg==r1:
    a000=1
else:
    print("测试编号错误")
    exit() 
print('患者'+str(int(IDmax))+'例')
print('每例患者'+str(int(Sampg))+'个样本')
print('总计样本'+str(int(Nomax))+'个')
print('每个样本重复测定'+str(int(Nog))+'次')
print('总计测定'+str(int(r1))+'次')
def datef(ax):#定义一个函数遍历对应列的所有非空格数据
    datemean=[]
    for fi in range(0,r1):
        Qdata=xlrd.open_workbook('D:\\TF\\华西\\课题\\黄老师\\生物学变异\\00成熟运算流程\\re0.xls')#读取转存表格
        Qtable=Qdata.sheets()[0]
        zdatemean1=Qtable.col_values(ax)[fi]
        if zdatemean1=="异常值":
            a000=1
        else:
            datemean.append(zdatemean1)
    return datemean

def cvf(ax):#根据列的编码计算出生物学变异数据
    file=xlwt.Workbook()#转存数据
    tablen=file.add_sheet('info',cell_overwrite_ok=True)
    datemean=[]
    for i in range(1,r1+1):
        #print(i)
        zdatemean1=table.col_values(ax)[i]
        if zdatemean1=="":
            #print('异常值')
            tablen.write(i-1,0,"异常值")
        elif int(zdatemean1)<0:
            #print('异常值')
            tablen.write(i-1,0,"异常值")
        else:
            datemean.append(zdatemean1)
            tablen.write(i-1,0,zdatemean1)
    #print(datemean)
    file.save('re0.xls')
    mean=numpy.mean(datemean)
    sd=numpy.std(datemean)
    #print(mean,sd)

    #去离群值
    t=0
    tn=0
    while t in range(0,r1+1):
        #print (t)
        tn=tn+1
        Qdata=xlrd.open_workbook('D:\\TF\\华西\\课题\\黄老师\\生物学变异\\00成熟运算流程\\re0.xls')#读取转存表格
        Qtable=Qdata.sheets()[0]
        datemean=datef(0)
        mean=numpy.mean(datemean)
        sd=numpy.std(datemean)
        #print(mean,sd)
        Qn=0
        for j in range(0,r1):
            if Qtable.col_values(0)[j]=="异常值":
                a00=0
            elif abs(round(Qtable.col_values(0)[j],6)-mean)/sd > 3:
                #print(round(Qtable.col_values(0)[j],6))
                Qn=Qn+1
                tablen.write(j,0,"异常值")
            else:
                a0=0
        file.save('re0.xls')
        #print(Qn)
        if Qn==0:
            break
    print('去离群值次数：'+str(tn-1)+'次')
    #去离群值之后的均数与标准差
    mean=numpy.mean(datemean)
    sd=numpy.std(datemean)    
    Qdata=xlrd.open_workbook('D:\\TF\\华西\\课题\\黄老师\\生物学变异\\00成熟运算流程\\re0.xls')#读取转存表格
    Qtable=Qdata.sheets()[0]
    zdate1=[]#以ID分组数据
    i=0
    j=0
    gammano1=0#开始计算γ
    SST1=0
    SSE1=0
    nt1=0
    for i in range(0,IDmax):
        #print(i)
        zdaten1=[]
        for j in range(i*IDg,(i+1)*IDg):
            #print(j)
            zdated=Qtable.col_values(0)[j]
            if zdated=="异常值":
                a000=0
            else:
                zdaten1.append(zdated)
                SST10=(zdated-mean)**2
                #print(SST10)
                SST1=SST10+SST1#总的
                nt1=nt1+1
        #print(zdaten1)
        
        if zdaten1==[]:
            a000=0
        else:
            gammano0=(len(zdaten1)/Nog)**2
            #print(gammano0)
            gammano1=gammano1+gammano0
            zdate1.append(zdaten1)
            cc=[aa-numpy.mean(zdaten1) for aa in zdaten1]
            SSE10=sum(c*c for c in cc)
            #print(SSE0)
            SSE1=SSE1+SSE10#组内
    #print(zdate1)
    SSI1=SST1-SSE1#组间
    dfni1=len(zdate1)-1#组间
    dfnt1=nt1-1#总的
    dfne1=dfnt1-dfni1#组内
    msi1=SSI1/dfni1
    mse1=SSE1/dfne1
    #print(gammano1)
    #print(SSI1,SSE1,SST1,dfni1,dfne1,dfnt1,msi1,mse1)

    zdate2=[]#以No分组数据
    i=0
    j=0
    SST2=0
    SSE2=0
    nt2=0
    for i in range(0,Nomax):
        #print(i)
        zdaten2=[]
        for j in range(i*Nog,(i+1)*Nog):
            #print(j)
            zdated0=Qtable.col_values(0)[j]
            if zdated0=="异常值":
                a000=0
            else:
                zdaten2.append(zdated0)
                SST20=(zdated0-mean)**2
                #print(SST20)
                SST2=SST20+SST2#总的
                nt2=nt2+1
        #print(zdaten1)
        
        if zdaten2==[]:
            a000=0
        else:
            zdate2.append(zdaten2)
            cc=[aa-numpy.mean(zdaten2) for aa in zdaten2]
            SSE20=sum(c*c for c in cc)
            #print(SSE0)
            SSE2=SSE2+SSE20#组内
    #print(zdate1)
    SSI2=SST2-SSE2#组间
    dfni2=len(zdate2)-1#组间
    dfnt2=nt2-1#总的
    dfne2=dfnt2-dfni2#组内
    msi2=SSI2/dfni2
    mse2=SSE2/dfne2
    #print(SSI2,dfni2)
    #print(SSI2,SSE2,SST2,dfni2,dfne2,dfnt2,msi2,mse2)

    zdate2=[]#以No分组数据
    i=0
    j=0
    for i in range(0,Nomax):
        zdaten2=[]
        for j in range(i*Nog,(i+1)*Nog):
            zdateo=Qtable.col_values(0)[j]
            if zdateo=="异常值":
                a000=0
            else:
                zdaten2.append(zdateo)
        if zdaten2==[]:
            a000=0
        else:
            zdate2.append(zdaten2)
    #开始计算
    #print(mean,sd)

    sa=mse2
    si=(SSE1/(((dfni2+1)/(dfni1+1)-1)*(dfni1+1))-mse2)/2
    sg=(msi1-SSE1/(((dfni2+1)/(dfni1+1)-1)*(dfni1+1)))/(1/dfni1*((dfnt2+1)-1/(dfnt2+1)*gammano1*(((dfnt2+1)/(dfni1+1))/((dfni2+1)/(dfni1+1)))**2))
    cva=sa**0.5/mean
    cvi=si**0.5/mean
    cvg=sg**0.5/mean
    #print(cva,cvi,cvg)
    tablerd.write(0,0,"项目名称")
    tablerd.write(0,1,"均数")#行，列
    tablerd.write(0,2,"CVa")
    tablerd.write(0,3,"CVi")
    tablerd.write(0,4,"CVg")
    tablerd.write(ax-3,0,table.col_values(ax)[0])
    tablerd.write(ax-3,1,mean)
    tablerd.write(ax-3,2,cva)
    tablerd.write(ax-3,3,cvi)
    tablerd.write(ax-3,4,cvg)
    result="OK"
    return result
filerd=xlwt.Workbook()#数据结果写入excel
tablerd=filerd.add_sheet('sheet1',cell_overwrite_ok=True)
i=0
for i in range(4,nonc1):
    print(cvf(i),'完成：'+str(round(((i-3)/(nonc1-4)*100),2))+'%')
    #print('完成：'+str(round((i/(nonc1-1)*100),2))+'%')
#filerd.save('result'+str(time.time())+'.xls')
filerd.save('result.xls')