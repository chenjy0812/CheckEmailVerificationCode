#-*- encoding:utf-8 -*-

import email, getpass, poplib, sys
import pyperclip
import win32api,win32con
import re, time

hostname = 'pop.gmail.com'
user = 'xxx@gmail.com'
passwd = '******'

VK_CODE={
    'enter':0x0D,
    'ctrl':0x11,
    'v':0x56}

def keyDown(keyName):
    #按下按键
    win32api.keybd_event(VK_CODE[keyName],0,0,0)

def keyUp(keyName):
    #释放按键
    win32api.keybd_event(VK_CODE[keyName],0,win32con.KEYEVENTF_KEYUP,0)
    
def oneKey(key):#对前两个方法的调用
    #模拟单个按键
    keyDown(key)
    keyUp(key)

def pressTwoKeys(key1,key2):#对前面函数的调用
    #模拟两个组合键
    keyDown(key1)
    keyDown(key2)
    keyUp(key2)
    keyUp(key1)

# 循环接收新邮件
for count in range(10000):
    time.sleep(1)
    p = poplib.POP3_SSL('pop.gmail.com') #与SMTP一样，登录gmail需要使用POP3_SSL() 方法，返回class POP3实例
    try:
        # 使用POP3.user(), POP3.pass_()方法来登录个人账户
        p.user(user) 
        p.pass_(passwd)
        print('log on')
    except poplib.error_proto: #可能出现的异常
        print('login failed')
    else:
        response, listings, octets = p.list()
        for listing in listings:
            number, size = listing.split() #取出message-id
            number = bytes.decode(number) 
            """
            size = bytes.decode(size) 
            print('Message', number, '( size is ', size, 'bytes)')
            response, lines, octets = p.top(number , 0)
            # 继续把Byte类型转化成普通字符串
            for i in range(0, len(lines)):
                lines[i] = bytes.decode(lines[i])
            #利用email库函数转化成Message类型邮件
            message = email.message_from_string('\n'.join(lines))
            # 输出From, To, Subject, Date头部及其信息
            for header in 'From', 'To', 'Subject', 'Date':
                if header in message:
                    print(header + ':' , message[header])
            """#
            #与用户交互是否想查看邮件内容
            response, lines, octets = p.retr(number) #检索message并返回
            for i in range(0, len(lines)):
                lines[i] = bytes.decode(lines[i])
            message = email.message_from_string('\n'.join(lines)) 
            print('-' * 72)
            maintype = message.get_content_maintype()
            if maintype == 'multipart':
                for part in message.get_payload():
                    if part.get_content_maintype() == 'text':
                        mail_content = part.get_payload(decode=True).strip()
            elif maintype == 'text':
                mail_content = message.get_payload(decode=True).strip()
            try:
                mail_content = mail_content.decode('utf8')
            except UnicodeDecodeError:
                mail_content = mail_content.decode('gbk')
            
            mat = re.search('[^(\d|#)](\d{6}?)([^\d]|$)', mail_content)
            if mat != None:
                code = mat.group(1)
                pyperclip.copy(code)
                print('receive mail code:' + pyperclip.paste())
                pressTwoKeys('ctrl','v')
    finally:
        print('log out')
        p.quit()
