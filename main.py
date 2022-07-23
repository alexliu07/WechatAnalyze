from xml.dom.minidom import parse
from pandas import read_csv
import hashlib,os,time
from openpyxl.styles import Border,Side,colors
import openpyxl
if not os.path.exists('output'):
    os.mkdir('output')
#输入
authfilepath = input('auth_info_key_prefs.xml路径：').replace('"','').replace('\\','/')
dbpath = input('EnMicroMsg.db路径：')
#获取uin
print("获取UIN中......",end='')
xmlf = parse(authfilepath)
data = xmlf.documentElement
mains = data.getElementsByTagName('int')
for i in mains:
    name = i.getAttribute('name')
    if name == '_auth_uin':
        uin = i.getAttribute('value')
print(uin)
#获取密码
print('计算数据库密码中......',end='')
keytext = '1234567890ABCDEF'+str(uin)
input_name = hashlib.md5()
input_name.update(keytext.encode("utf-8"))
key = (input_name.hexdigest()).lower()[0:7]
print(key)
#写入解密指令
print('解密数据库中......',end='')
decryptcmd = "PRAGMA key = '"+key+"';\nPRAGMA cipher_use_hmac = off;\nPRAGMA kdf_iter = 4000;\nATTACH DATABASE 'wechat.db' AS wechat KEY '';\nSELECT sqlcipher_export('wechat');\nDETACH DATABASE wechat;\n.e"
decryptfile = open('decrypt.txt','w+',encoding='utf-8')
decryptfile.write(decryptcmd)
decryptfile.close()
#解密数据库
os.system('getdb.bat '+dbpath)
#删除解密指令
os.remove('decrypt.txt')
print('完成')
#解析数据库
print('正在导出数据表......',end='')
#写入csv
os.system('sqlite3\\sqlite3.exe wechat.db < exportcsv.txt')
#删除解密后的数据库
os.remove('wechat.db')
print('完成')
#导入及分析数据
print('正在分析数据......',end='')
chatfile = read_csv('message.csv')
listin = chatfile.values.tolist()
listout = []
for i in listin:
    if i[2] == 1:
        listout.append([i[6],i[7],i[8],i[4]])
    elif i[2] == 3:
        listout.append([i[6],i[7],'[图片]',i[4]])
#删除csv
print('完成')
#导入excel
print('正在导入Excel......',end='')
book = openpyxl.Workbook()
sheet = book.active
sheet.title = 'WechatMsg'
#设置表头
sheet['A1'] = '发送时间(时间戳)'
sheet['B1'] = '发送时间(实际时间)'
sheet['C1'] = '聊天对象'
sheet['D1'] = '是否为本人发出'
sheet['E1'] = '消息'
#将时间戳转换为实际时间并存入其他数值
for i in range(len(listout)):
    sheet['A'+str(i+2)] = listout[i][0]
    timearray = time.localtime(listout[i][0]/1000)
    currentTime = time.strftime("%Y-%m-%d %H:%M:%S", timearray)
    sheet['B'+str(i+2)] = currentTime
    sheet['C'+str(i+2)] = listout[i][1]
    if listout[i][3] == 0:
        sheet['D'+str(i+2)] = '否'
    else:
        sheet['D'+str(i+2)] = '是'
    sheet['E'+str(i+2)] = listout[i][2]
#居中及框线
alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center", text_rotation=0)
border_set = Border(left=Side(style='thin', color=colors.BLACK),right=Side(style='thin', color=colors.BLACK),top=Side(style='thin', color=colors.BLACK),bottom=Side(style='thin', color=colors.BLACK))
for i in sheet['A']:
        i.alignment = alignment
        i.border = border_set
for i in sheet['B']:
    i.alignment = alignment
    i.border = border_set
for i in sheet['C']:
    i.alignment = alignment
    i.border = border_set
for i in sheet['D']:
    i.alignment = alignment
    i.border = border_set
for i in sheet['E']:
    i.alignment = alignment
    i.border = border_set
#调整列宽
sheet.column_dimensions['A'].width = 17
sheet.column_dimensions['B'].width = 21
sheet.column_dimensions['C'].width = 23
sheet.column_dimensions['D'].width = 15
sheet.column_dimensions['E'].width = 72
#储存
book.save('output/message.xlsx')
print('完成')
#聊天记录写入txt
print('正在写入聊天记录......',end='')
for i in listout:
    #检测是否为群聊
    if '@chatroom' in i[1]:
        filename = 'chatroom'+(i[1].replace('@chatroom',''))
        sender = ''
    else:
        filename = i[1]
        if i[3] == 0:
            sender = i[1]+': '
        else:
            sender = '我: '
    if not os.path.exists('output/chats'):
        os.mkdir('output/chats')
    chatfile = open('output/chats/'+filename+'.txt','a',encoding='utf-8')
    timearray = time.localtime(i[0]/1000)
    currentTime = time.strftime("%Y-%m-%d %H:%M:%S", timearray)
    chatfile.write(currentTime+'：\n'+sender+i[2].replace('\n',' ')+'\n')
    chatfile.close()
print('完成')