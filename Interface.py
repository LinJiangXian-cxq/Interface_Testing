import http.client
import xlrd

conn = http.client.HTTPConnection('localhost')
conn.request('GET','/agileone/')
content = conn.getresponse()
cookies = content.getheader('Set-Cookie')
cookie = cookies.split(';')[0]

param = 'username=admin&password=admin&savelogin=true'
header = {'Content-type':'application/x-www-form-urlencoded','Cookie':cookie}
conn = http.client.HTTPConnection('localhost')
conn.request('POST','/agileone/index.php/common/login HTTP/1.1',param,header)
response = conn.getresponse().read()
print('--开始登陆啦：')
print('返回的是：',response)
if response.decode() == 'successful':
    print('[输入正确信息测试]-[成功登录]')
else:
    print('[输入正确信息测试]-[登录失败]')

biaoge=xlrd.open_workbook('d:3.xlsx')
diyigebiao=biaoge.sheets()[0]

print('')
print('--新增数据模块接口测试开始啦：')
l1=[]

l3=[]
for i in range(1,6):
    list=diyigebiao.row_values(i)
    list1=list[0:1]
    l1.append(list1)
    list3=list[4:7]
    l3.append(list3)
for j in range(0,5):
    l3[j][0]=str(l3[j][0])
    paramj='content='+l3[j][2]+'&type='+l3[j][1]+'&userstoryid='+l3[j][0]
    headerj={'Content-type':'application/x-www-form-urlencoded','Cookie':cookie}
    conn = http.client.HTTPConnection('localhost')
    conn.request('POST','/agileone/index.php/spec/add HTTP/1.1',paramj,headerj)
    responsej=conn.getresponse().read()
    print('模块',l1[j][0],'第',j+1,'条测试用例:','返回的是：',responsej.decode())
    if type(responsej.decode())=='<class \'bytes\'>':
        print('新增成功')
    else:
        print('新增失败')



print('')
print('--搜索数据模块接口测试开始啦：')
l4=[]
l6=[]
for ii in range(6,11):
    listii=diyigebiao.row_values(ii)
    list4=listii[0:1]
    l4.append(list4)
    list6=listii[3:9]
    l6.append(list6)
for jj in range(0,5):
    l6[jj][0]=str(l6[jj][0])
    l6[jj][1]=str(l6[jj][1])
    l6[jj][4]=str(l6[jj][4])

    paramjj='creator='+l6[jj][5]+'&currentpage='+l6[jj][4]+'&specid='+l6[jj][0]+'&content='+l6[jj][3]+'&type='+l6[jj][2]+'&userstoryid='+l6[jj][1]
    headerjj={'Content-type':'application/x-www-form-urlencoded','Cookie':cookie}
    conn = http.client.HTTPConnection('localhost')
    conn.request('POST','/agileone/index.php/spec/query HTTP/1.1',paramjj,headerjj)
    responsejj=conn.getresponse().read()
    print('模块',l4[jj][0],'第',jj+1,'条测试用例:','返回的是：',responsejj.decode())
    if responsejj.decode()=='[{"totalRecord":0}]':
        print('搜索失败')
    else:
        print('搜索成功')





print('')
print('--删除数据模块接口测试开始啦：')
l7=[]
l8=[]
for iii in range(11,15):
    listiii=diyigebiao.row_values(iii)
    list7=listiii[0:1]
    l7.append(list7)
    list8=listiii[3:4]
    l8.append(list8)
for jjj in range(0,4):
    l8[jjj][0]=str(l8[jjj][0])


    paramjjj='specid='+l8[jjj][0]
    headerjjj={'Content-type':'application/x-www-form-urlencoded','Cookie':cookie}
    conn = http.client.HTTPConnection('localhost')
    conn.request('POST','/agileone/index.php/spec/delete HTTP/1.1',paramjjj,headerjjj)
    responsejjj=conn.getresponse().read()
    print('模块',l7[jjj][0],'第',jjj+1,'条测试用例:','返回的是：',responsejjj.decode())
    if responsejjj.decode()=='1':
        print('删除成功')
    else:
        print('删除失败')
    print('')


