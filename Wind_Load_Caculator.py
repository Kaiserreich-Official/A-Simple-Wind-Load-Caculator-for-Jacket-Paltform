from easygui import msgbox,enterbox,multenterbox,buttonbox,textbox,filesavebox
from sys import exit
from numpy import array,dot,log,cos,sin,pi,sqrt
from xlwt import Workbook
workbook = Workbook()

sheet1 = workbook.add_sheet('源数据')
sheet2 = workbook.add_sheet('计算结果')

msgbox('加载时请根据结果的XYZ坐标，使用nsel命令选择构建中心Node加载','使用须知')
msgbox('善用Tab键:它可以在表格间快速切换而不用鼠标点击','你知道吗')

velocity = float(enterbox('请输入风速(m/s)','简易风载荷计算器','34.5'))
theta = float(enterbox('来流方向(度)','简易风载荷计算器','45'))/180*pi
dens = 1.29

Diam_lst = []
x1_lst = []
y1_lst = []
z1_lst = []
x2_lst = []
y2_lst = []
z2_lst = []
x_lst = []
y_lst = []
z_lst = []
fx_lst = []
fy_lst = []

def calc(Diam,x1,y1,z1,x2,y2,z2):
    Diam = float(Diam)
    loc1 = array([float(x1),float(y1),float(z1)])
    loc2 = array([float(x2),float(y2),float(z2)])
    #方向矢量旋转
    Rz = array([[cos(theta),-sin(theta),0],[sin(theta),cos(theta),0],[0,0,1]])
    loc1 = dot(Rz,loc1)
    loc2 = dot(Rz,loc2)
    x1 = float(loc1[0])
    z1 = float(loc1[2])
    x2 = float(loc2[0])
    z2 = float(loc2[2])
    length = sqrt((x1-x2)**2+(z1-z2)**2)
    F = dens/2*wind(z1,z2)**2*0.5*length*Diam
    return F*sin(theta),F*cos(theta)

def wind(z1,z2):
    return velocity*(1+0.137*log((z1+z2)/2/10)-0.047*log(3/600))

def align(text,lst,n):
    for elem in lst:
        if len(str(elem)) < n:
            para = str(elem)
            for i in range(n -len(str(elem))):
                para += ' '
        else:
            para = str(elem)
        para += '|'
        text += para
    text += '\n'
    return text

def menu():
    choice = buttonbox('欢迎使用简易风载荷计算器','简易风载荷计算器',('添加一个构件','查看构件表','计算结果','结果查看','结果保存为Excel','关于作者','退出'))
    if choice == '添加一个构件':
        try:
            Diam,x1,y1,z1,x2,y2,z2 = multenterbox('请输入参数','简易风载荷计算器',['外径D(m)','x1','y1','z1','x2','y2','z2'])
            Diam_lst.append(float(Diam))
            x1_lst.append(float(x1))
            y1_lst.append(float(y1))
            z1_lst.append(float(z1))
            x2_lst.append(float(x2))
            y2_lst.append(float(y2))
            z2_lst.append(float(z2))
        except:
            msgbox('未输入数据或者格式有误!','Warning')
    if choice == '查看构件表':
        try:
            text = '外径D(m):'
            text = align(text,Diam_lst,5)
            text += 'x1      :'
            text = align(text,x1_lst,5)
            text += 'y1      :'
            text = align(text,y1_lst,5)
            text += 'z1      :'
            text = align(text,z1_lst,5)
            text += 'x2      :'
            text = align(text,x2_lst,5)
            text += 'y2      :'
            text = align(text,y2_lst,5)
            text += 'z2      :'
            text = align(text,z2_lst,5)
            textbox('构件汇总','简易风载荷计算器',text)
        except:
            msgbox('尚未输入数据!','Warning')
    if choice == '计算结果':
        if Diam_lst:
            for i in range(len(Diam_lst)):
                x_lst.append((x1_lst[i]+x2_lst[i])/2)
                y_lst.append((y1_lst[i]+y2_lst[i])/2)
                z_lst.append((z1_lst[i]+z2_lst[i])/2)
                fx,fy = calc(Diam_lst[i],x1_lst[i],y1_lst[i],z1_lst[i],x2_lst[i],y2_lst[i],z2_lst[i])
                fx_lst.append(format(fx,'.3f'))
                fy_lst.append(format(fy,'.3f'))
            msgbox('计算完成！','Success')
        else:
            msgbox('尚未输入数据!','Warning')
    if choice == '结果查看':
        if x_lst:
            text = 'x :'
            text = align(text,x_lst,8)
            text += 'y :'
            text = align(text,y_lst,8)
            text += 'z :'
            text = align(text,z_lst,8)
            text += 'FX:'
            text = align(text,fx_lst,8)
            text += 'FY:'
            text = align(text,fy_lst,8)
            textbox('计算结果汇总','简易风载荷计算器',text)
        else:
            msgbox('尚未计算结果或尚未输入数据!','Warning')
    if choice == '结果保存为Excel':
        try:
            path = filesavebox(title='保存计算表格结果',default='./计算表格.xls',filetypes='.xls')
            sheet1.write(0,0,"外径D(m)")
            sheet1.write(1,0,"x1")
            sheet1.write(2,0,"y1")
            sheet1.write(3,0,"z1")
            sheet1.write(4,0,"x2")
            sheet1.write(5,0,"y2")
            sheet1.write(6,0,"z2")
            sheet2.write(0,0,"x")
            sheet2.write(1,0,"y")
            sheet2.write(2,0,"z")
            sheet2.write(3,0,"Fx(N)")
            sheet2.write(4,0,"Fy(N)")
            for i in range(len(Diam_lst)):
                sheet1.write(0,i+1,Diam_lst[i])
                sheet1.write(1,i+1,x1_lst[i])
                sheet1.write(2,i+1,y1_lst[i])
                sheet1.write(3,i+1,z1_lst[i])
                sheet1.write(4,i+1,x2_lst[i])
                sheet1.write(5,i+1,y2_lst[i])
                sheet1.write(6,i+1,z2_lst[i])
                sheet2.write(0,i+1,x_lst[i])
                sheet2.write(1,i+1,y_lst[i])
                sheet2.write(2,i+1,z_lst[i])
                sheet2.write(3,i+1,format(float(fx_lst[i]),'.3f'))
                sheet2.write(4,i+1,format(float(fy_lst[i]),'.3f'))
            workbook.save(path)
        except:
            msgbox('保存错误,请检查保存路径是否可读写!','Warning')
    if choice == '关于作者':
        msgbox('作者:Kaiserreich-Official\nQQ:2313594637\nMail:bibizidan@hotmail.com','关于作者')
    if choice == '退出':
        exit(0)

while True:
    menu()