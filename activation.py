#coding=utf-8
'''
  Title: QQ Mail Sender with Selenium Activator
  Version: V1.0
  Author: qq956302200
  Date: Last update 2017-11-03
  Email: 956302200@qq.com
'''

import hashlib
import win32clipboard
import win32con

if __name__=='__main__':
      print('\n软件名称：QQ邮箱模拟浏览器发送邮件软件 注册机 1.4')
      print('----------------------------------------------------------------------------------------------------------------')
      print('\n1.激活码生成后，会自动复制到剪贴板中，在记事本或者聊天窗口中按下 ctrl+v ,即可将激活码粘贴到窗口中，发送给客户。')
      print('\n2.您也可以将本软件目录下生成的certificate.key文件发给客户。')
      print('\n3.默认只对比前10位，因此在手动输入的情况下，发送前十位即可。')
      while True:
            machinecode=input('\n请输入机器码:').replace('\n','')
            hash=hashlib.sha1()
            hash.update(bytes(machinecode+'000', encoding='utf-8'))
            activationcode=hash.hexdigest()
            f=open('certificate.key','w',encoding='utf-8')
            f.write(activationcode)
            f.close()            
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, activationcode)
            win32clipboard.CloseClipboard()
            print('\n激活码为:%s'%activationcode)

            print('\n------------------------继续输入机器码，以生成激活码-----------------------------------------------')