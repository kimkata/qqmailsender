#coding=utf-8
'''
  Title: QQ Mail Sender with Selenium
  Version: V1.2
  Author: qq956302200
  Date: Last update 2017-11-03
  Email: 956302200@qq.com
'''
from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from threading import Thread
import threading
from PIL import Image #pip install pillow
import os, time, sys,datetime,winsound,csv,re,socket,hashlib,json,requests,math
import win32clipboard
import win32con
import pythoncom

class QQMailSender(Thread):
    #类变量
    CHROMEDRIVE = r'chromedriver.exe'
    MAIL_SENDER_LIST_FILE_PATH=os.path.join('settings',u'发件箱.csv')
    MAIL_TO_LIST_FILE_PATH=os.path.join('settings',u'收件箱.csv')
    MAIL_CONTENT_FILE_PATH=os.path.join('settings',u'邮件内容.txt')
    MAIL_SEND_RESULT_FILE_PATH=os.path.join('settings',u'发送结果.csv')
    MAIL_ATTACHMENTS_PATH=u'附件文件夹'
    MAIL_SETTINGS_FILE_PATH=os.path.join('settings','settings.json')
    #默认设置
    DEFAULT_SETTINGS={'threads_max':1,'mails_count_in_each_thread':50,'address_num_in_each_mail':5,
    'settings_of_lianzhong':{
    'username':'','password':''
    },
    'settings_of_chrome':{
                    'profile.default_content_settings.popups': 0,
                    'profile.default_content_setting_values.images':0,
                    'profile.default_content_setting_values.background_sync':0,
                    'profile.default_content_setting_values.cookies': 0,
                    'profile.default_content_setting_values.javascript':0,
                    'profile.default_content_setting_values.notifications': 0,
                    'profile.default_content_setting_values.plugins':0
    }}


    def __init__(self,message,PROXY=None,settings_dir=None,create_browser=True):
        threading.Thread.__init__(self)
        self.msg=message
        self.PROXY=PROXY
        self.settings=self.read_settings()
        self.chrome_init(settings_dir)
        self.desired_capabilities=None
        if self.PROXY:
            self.desired_capabilities = self.options.to_capabilities()
            desired_capabilities['proxy'] = {
                "httpProxy":self.PROXY,
                "ftpProxy":self.PROXY,
                "sslProxy":self.PROXY,
                "noProxy":None,
                "proxyType":"MANUAL",
                "class":"org.openqa.selenium.Proxy",
                "autodetect":False}
        if create_browser:
            self.browser = webdriver.Chrome(self.CHROMEDRIVE, chrome_options=self.options,desired_capabilities=self.desired_capabilities)
        #如果不存在settings文件夹，则创建它
        if not os.path.exists('./settings'):
            os.mkdir('./settings')
            print('\n成功创建settings文件夹...')

        #如果不存在附件文件夹，则创建它
        if not os.path.exists(self.MAIL_ATTACHMENTS_PATH):
            os.mkdir(self.MAIL_ATTACHMENTS_PATH)
            print('\n成功创建文件夹=>%s'%(self.MAIL_ATTACHMENTS_PATH))

        if not os.path.exists(self.MAIL_SENDER_LIST_FILE_PATH):
            f=open(self.MAIL_SENDER_LIST_FILE_PATH,'w',encoding='utf-8-sig',newline='')
            writer=csv.writer(f)
            writer.writerows([(u'QQ账号',u'QQ密码')])
            f.close()

        if not os.path.exists(self.MAIL_TO_LIST_FILE_PATH):
            f=open(self.MAIL_TO_LIST_FILE_PATH,'w',encoding='utf-8-sig',newline='')
            writer=csv.writer(f)
            writer.writerows([[u'收件箱']])
            f.close()

        if not os.path.exists(self.MAIL_CONTENT_FILE_PATH):
            f=open(self.MAIL_CONTENT_FILE_PATH,'w',encoding='utf-8')
            text=u'\n第1行=>填写邮件主题\n从第2行开始=>填写邮件正文'
            f.write(text)
            f.close()

    def __del__(self):
        pass       

    def chrome_init(self, settings_dir):
        if not os.path.exists(self.CHROMEDRIVE):
            print(self.CHROMEDRIVE + '\n=>驱动程序不存在,请下载chromedriver.exe文件(即:谷歌浏览器驱动程序)')
            sys.exit()
        self.settings_dir = self.check_settings_dir(settings_dir)
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--start-maximized")
        self.options.add_argument("disable-extensions")
        prefs=self.settings['settings_of_chrome']
        self.options.add_experimental_option('prefs', prefs)

        
    @staticmethod
    def check_settings_dir(settings_dir=None):
        if settings_dir == None:
            settings_dir = os.getcwd() + os.sep + r'settings'
        if not os.path.exists(settings_dir):
            os.mkdir(settings_dir)
        return settings_dir


    def get_attachments_list(self):
        attachments_list=[]
        if os.path.exists(self.MAIL_ATTACHMENTS_PATH):
            for root, dirs, files in os.walk(os.getcwd()+os.sep+self.MAIL_ATTACHMENTS_PATH):
                for file in files:
                    file_full_name=os.path.join(root,file).replace('\\','/')
                    attachments_list.append(file_full_name)
        return attachments_list

    def get_sender_list(self):
        sender_list=[]
        if os.path.exists(self.MAIL_SENDER_LIST_FILE_PATH):
            try:
                f=open(self.MAIL_SENDER_LIST_FILE_PATH,'r',encoding='utf-8')
                lines=f.readlines()
                for line in lines[1:]:
                    item=line.replace('\t',',').split(',')
                    try:
                        sender_list.append((item[0],item[1]))
                    except Exception as e:
                        self.write_error_into_file(str(e))
                        print('\n本行发件箱信息添加失败，请仔细查看是否填写有误=>%s'%(line))
            except Exception as e:
                   self.write_error_into_file(str(e))
                   print('\n读取文件失败！=>%s'%(self.MAIL_SENDER_LIST_FILE_PATH))

        else:
            print(u'\n没有找到邮件发件箱列表文件，请检查!')
        return sender_list


    def get_to_list(self):
        to_list=[]
        if os.path.exists(self.MAIL_TO_LIST_FILE_PATH):
            try:
                f=open(self.MAIL_TO_LIST_FILE_PATH,'r',encoding='utf-8')
                lines=f.readlines()
                for line in lines:
                    #注意：这是对单个邮件地址的验证，如果是多个邮件地址写在同一行，就会错误。
                    to_list+=re.findall('^[a-zA-Z0-9._%+-]+@(?!.*\.\..*)[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$',self.fix_mailaddress(line.replace('\n','').replace('\r','').replace('\t','')))
                return to_list
            except Exception as e:
                   self.write_error_into_file(str(e))
                   print('读取文件失败！=>%s'%(self.MAIL_SENDER_LIST_FILE_PATH))
        else:
            print(u'\n没有找到邮件收件人列表文件，请检查!')           
        return to_list

    #以下函数用于处理只填写了号码的qq邮箱。
    @staticmethod
    def fix_mailaddress(mailaddress):
        #注意要对mailaddress进行去空白字符处理。
        if re.findall('^\d{5,10}$',re.sub('\s+','',mailaddress)):
            mailaddress+='@qq.com'
        return mailaddress


    def get_content_text(self):
        if os.path.exists(self.MAIL_CONTENT_FILE_PATH):
            f=open(self.MAIL_CONTENT_FILE_PATH,'r',encoding='utf-8')
            lines=f.readlines()
            #读出第一行作为发件的主题。
            if len(lines)>1:
                subject=lines[0].replace('\n','').replace('\r','').strip() #主题不可能包含换行符，如果有换行符，应该将其去掉，包括前后的空格。
                #从第二行的内容作为文件的文本内容。
                content_text=''
                for line in lines[1:]:
                    content_text+=line
            else:
                subject='None'
                content_text='None'
        else:
            subject='None'
            content_text='None'

        return {'subject':subject,'content_text':content_text}

    def run(self):
        try:
            if self.login(self.msg['username'],self.msg['password']):
                message_item=msg
                while self.msg['to']:
                    to_str=''
                    for num in range(self.settings['address_num_in_each_mail']):
                        if self.msg['to']:
                            to_str+=self.msg['to'].pop()+';'
                    message_item['to']=to_str
                    self.send_mails(message_item)
            else:
                print('\n邮箱登陆未成功，即将退出浏览器。')
        except Exception as e:
            self.write_error_into_file(str(e))
            print(str(e))

        winsound.Beep(3000,500)
        self.browser.close()
        self.browser.quit()
        print('线程ID：%s =>已经关闭!'%(threading.current_thread()))


    def check_element_existed_by_id(self,element_id):
        try:
            self.browser.find_element_by_id(element_id)
            return True
        except :
            return False

    def check_element_existed_by_xpath(self,element_xpath):
        try:
            self.browser.find_element_by_xpath(element_xpath)
            return True
        except :
            return False



    def login(self,username,password):
        print('\n线程ID:%s, 将使用发送账号:%s, 密码:%s'%(threading.current_thread(),username,password))
        self.browser.get('https://mail.qq.com/')
        self.browser.switch_to.frame('login_frame')
        WebDriverWait(self.browser,5,0.5).until(EC.presence_of_element_located((By.XPATH,"//div/a[@id='switcher_plogin']"))).click()
        self.browser.find_element_by_xpath("//input[@id='u']").send_keys(username)
        self.browser.find_element_by_xpath("//input[@id='p']").send_keys(password)
        self.browser.find_element_by_xpath("//input[@id='login_button']").click()
        time.sleep(3)
        #检测是否有验证提示窗口，如果有，则发出声音,并且休眠20秒。
        if self.check_element_existed_by_id("composebtn"):
            return True
        else:
            winsound.Beep(2000,500)
            print('\n线程ID:%s,请检测是否需要验证,本软件将此进程休眠20秒，请在20秒内在浏览器页面完成验证操作,如果未完成验证，本浏览器窗口将退出。'%(threading.current_thread()))
            time.sleep(20)
            #在此检测是否进入了邮箱发件页面，如果没有，那么将浏览器退出，并且不再执行后面的发送操作。
            if self.check_element_existed_by_id('composebtn'):
                return True
            else:
                return False

    def get_lianzhong_check_result(self,api_username, api_password, file_name, api_post_url, yzm_min, yzm_max, yzm_type, tools_token):
        '''
                main() 参数介绍
                api_username    （API账号）             --必须提供
                api_password    （API账号密码）         --必须提供
                file_name       （需要打码的图片路径）   --必须提供
                api_post_url    （API接口地址）         --必须提供
                yzm_min         （验证码最小值）        --可空提供
                yzm_max         （验证码最大值）        --可空提供
                yzm_type        （验证码类型）          --可空提供
                tools_token     （工具或软件token）     --可空提供
        '''
        # api_username =
        # api_password = 
        # file_name = 'c:/temp/lianzhong_vcode.png'
        # api_post_url = "http://v1-http-api.jsdama.com/api.php?mod=php&act=upload"
        # yzm_min = '1'
        # yzm_max = '8'
        # yzm_type = '1303'
        # tools_token = api_username

        # proxies = {'http': 'http://127.0.0.1:8888'}
        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3',
            'Accept-Encoding': 'gzip, deflate',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64; rv:53.0) Gecko/20100101 Firefox/53.0',
            # 'Content-Type': 'multipart/form-data; boundary=---------------------------227973204131376',
            'Connection': 'keep-alive',
            'Host': 'v1-http-api.jsdama.com',
            'Upgrade-Insecure-Requests': '1'
        }

        files = {
            'upload': (file_name, open(file_name, 'rb'), 'image/png')
        }

        data = {
            'user_name': api_username,
            'user_pw': api_password,
            'yzm_minlen': yzm_min,
            'yzm_maxlen': yzm_max,
            'yzmtype_mark': yzm_type,
            'zztool_token': tools_token
        }
        s = requests.session()
        # r = s.post(api_post_url, headers=headers, data=data, files=files, verify=False, proxies=proxies)
        r = s.post(api_post_url, headers=headers, data=data, files=files, verify=False)
        print(r.text)
        response_list=json.loads(r.text)
        #{"result":false,"data":"\u672a\u4e0a\u4f20\u9a8c\u8bc1\u7801\u56fe\u7247"}
        if response_list['result']:
            print('\n联众服务器成功返回验证码！')
            return response_list['data']['val']
        else:
            winsound.Beep(2000,500)
            print('\n返回的验证码错误，错误信息:%s'%response_list['data'])
            return 'None'

    def send_mails(self,msg):
        # time.sleep(3)
        self.browser.switch_to_default_content()         
        WebDriverWait(self.browser,5,0.5).until(EC.presence_of_element_located((By.ID,"composebtn"))).click()
        self.__switch_to_iframe('mainFrame') #注意这里包含有等待，因此后面不用sleep
        self.browser.find_element_by_xpath('''//*[@id="toAreaCtrl"]/div[2]/input''').send_keys(msg['to'])
        if msg['subject']!='None':
            self.browser.find_element_by_xpath('''//*[@id="subject"]''').send_keys(msg['subject'])
        #添加附件
        for attachment in msg['attachments_list']:
            self.browser.find_element_by_xpath('//*[@id="AttachFrame"]/span/input').send_keys(attachment)
        #定位‘正文’iframe 位置
        if msg['body']!='None':
            #定位‘正文’iframe 位置
            main_body= self.browser.find_element_by_xpath("//*[@id='QMEditorArea']/table/tbody/tr[2]/td/iframe")
            self.browser.switch_to.frame(main_body)
            self.browser.find_element_by_xpath("/html/body").send_keys(msg['body'])
            time.sleep(2)
            #注意:frame只能一层一层往里面定位，如果要退出iframe,使用switch_to_default_content
        self.browser.switch_to_default_content()
        self.__switch_to_iframe('mainFrame') 
        self.browser.find_element_by_xpath('''//*[@id="toolbar"]/div/a[1]''').click()
        # 以下检测验证码窗口的方式使用了智能显示等等，因此这里不用sleep
        #实际使用中发现，主题，收件人，内容的输入都是异步的
        try:
            print('\n正在检测是否有验证码窗口!')
            WebDriverWait(self.browser,5,0.5).until(
                EC.presence_of_element_located((By.XPATH,'//*[@id="QMVerify_QMDialog_verify_img_code"]'))
                )
        except :
            print('\n未检测到输入验证码提示!')
        self.browser.switch_to_default_content()
        #检测是否有验证码出现
        while True:
            if self.check_element_existed_by_xpath('//*[@id="QMVerify_QMDialog_verify_img_code"]'):
                print('\n检测到点击发送后有验证码窗口弹出。')             
                #如果设置了联众的账号和密码，则调用联众打码，如果没有，直接休眠20秒。
                if self.settings['settings_of_lianzhong']['username']: 
                    try:
                        element=self.browser.find_element_by_xpath('//*[@id="QMVerify_QMDialog_verify_img_code"]')
                        picture_url=element.get_attribute('src') 
                        #获取element的位置和大小
                        location = element.location  #获取验证码x,y轴坐标
                        size=element.size  #获取验证码的长宽
                        print('\n验证码图片的坐标位置:%s'%str(location))
                        print('\n验证码图片的大小为:%s'%str(size))                             
                        print('\n验证码图片的链接为:%s'%picture_url)
                        rangle=(int(location['x']),int(location['y']),int(location['x']+size['width']),int(location['y']+size['height'])) 
                        #写成我们需要截取的位置坐标
                        key=time.time()
                        filename_png='img_%s.png'%(str(key).replace('.','_'))
                        self.browser.save_screenshot(filename_png)  #截取当前网页，该网页有我们需要的验证码
                        i=Image.open(filename_png) #打开截图
                        verify_image=i.crop(rangle)  #使用Image的crop函数，从截图中再次截取我们需要的区域
                        verify_image.save(filename_png)
                        verify_image.close()
                        #接入联众打码平台，获取验证结果。
                        file_name_upload=str(os.getcwd()+os.sep+filename_png).replace('\\','/')
                        username=self.settings['settings_of_lianzhong']['username']
                        password=self.settings['settings_of_lianzhong']['password']
                        print(file_name_upload)
                        print('\n软件设置中包含有联众平台的打码账号，下面将接入联众打码平台...')
                        img_text=self.get_lianzhong_check_result(username,
                                                                 password,
                                                                 file_name_upload,
                                                                "http://v1-http-api.jsdama.com/api.php?mod=php&act=upload", 
                                                                 '1',
                                                                 '8',
                                                                 '1001',
                                                                 '')                        
                        folder_of_image='verify_images'
                        if os.path.exists(folder_of_image):
                        	os.mkdir(folder_of_image)
                        os.rename(filename_png,os.path.join(folder_of_image,(img_text+'.png'))) #打码后，将图片移动到文件夹中。
                        #以下的输入出验证码值和点击确认的方法有效，经过了测试。
                        self.browser.find_element_by_xpath('//*[@id="QMVerify_QMDialog_verifycodeinput"]').send_keys(img_text)
                        self.browser.find_element_by_xpath('//*[@id="QMVerify_QMDialog_btnConfirm"]').click()
                        time.sleep(2)
                    except Exception as e:
                        print(str(e))
                else:
                    print('\n点击写邮件按钮出现异常，可能是需要验证码。下面休眠20秒，请完成验证。如果未及时验证，浏览器将退出。')
                    time.sleep(20)
                    break
            else:
                print('\n=>未检测到输入验证码提示,跳出验证码处理...')
                break

    def write_error_into_file(self,errorstr):
        time_str=time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
        file = open('错误信息.txt', 'a+', encoding='utf-8-sig')
        text=('时间:%s\n错误信息:%s\n'+'-'*50+'\n')%(time_str,errorstr)
        file.write(text)
        file.close()

    def __switch_to_iframe(self, iframe_name, timeout=10):
        WebDriverWait(self.browser, timeout).until(lambda driver: driver.find_element_by_id(iframe_name).is_displayed())
        self.browser.switch_to.frame(iframe_name)

    def read_settings(self):
        self.check_settings_dir()
        if os.path.exists(self.MAIL_SETTINGS_FILE_PATH):
            with open(self.MAIL_SETTINGS_FILE_PATH) as jsonfile:
                data=json.load(jsonfile)
                return data
        else:
            with open(self.MAIL_SETTINGS_FILE_PATH,'w') as jsonfile:
                jsonfile.write(json.dumps(self.DEFAULT_SETTINGS,indent=1))
            return self.DEFAULT_SETTINGS

if __name__ == '__main__': 
    print('\n软件名称：QQ邮箱模拟浏览器发送邮件客户端 2.0')
    print('\n本软件可以接入联众打码平台，联众网址：https://www.jsdama.com,使用前请在settings.json中设置联众打码用户名和密码。')  
    print('----------------------------------------------------------------------------------------------------------------')    
    time_now=1510482629.0
    left_days=365
    r=requests.get('https://www.baidu.com')
    date_str=r.headers['Date']
    baidu_date=time.strptime(date_str,'%a, %d %b %Y %H:%M:%S GMT')
    #注意百度和腾讯等公司返回的都是英国格林威治时间，与中国的时间落后8小时。
    baidu_time=time.mktime((baidu_date.tm_year,baidu_date.tm_mon,baidu_date.tm_mday,baidu_date.tm_hour+8, baidu_date.tm_min,baidu_date.tm_sec, baidu_date.tm_wday, baidu_date.tm_yday, baidu_date.tm_isdst))
    time_str=time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(baidu_time))
    print('\n当前网络时间: %s'%(time_str))
    if baidu_time>time_now+left_days*24*60*60:
        print('\n提示:此软件过期已经过期，请与软件商联系获取新版,邮箱:956302200@qq.com！')
        os.system('pause')
        sys.exit()
    p=os.popen('wmic CPU get ProcessorID')
    res=p.read()
    machine_code=re.findall('[a-z0-9]{16}',res,re.I)[0]
    hash=hashlib.sha1()
    hash.update(bytes(machine_code+'000', encoding='utf-8'))
    activation_code=hash.hexdigest()
    # print(activation_code)
    #检测certificate.key文件是否存在，如果存在，则读取这个文件，验证里面的值和正确的激活码是否相同。
    is_activated=False
    if os.path.exists('certificate.key'):
        f=open('certificate.key','r',encoding='utf-8')
        key=f.read().replace('\n','').replace('\t','').replace(' ','')
        f.close()
        is_activated=True if key==activation_code else False
        
    if is_activated==False:
        pythoncom.CoInitialize()
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, machine_code)
        win32clipboard.CloseClipboard()
        pythoncom.CoUninitialize()

    time_start=time.time()
    is_now_tested=False
    if not is_activated:
        while True:
            print('\n您的机器码是:%s\n\n本软件已经将机器码存入剪贴板，在键盘上按下 ctrl+v 即可将机器码粘贴到记事本或聊天窗口，请将其发给商家以获取授权！'%machine_code)
            print('\n您可以在本窗口中输入激活码，也可以将激活码放入certificate.key文件中，然后重新运行软件。')
            print('\n输入123456，可以试用10分钟，如果要购买授权，请联系商家!')
            activation_code_input=input('\n请在本行输入激活码:')
            #试用模式，可以试用10分钟。
            if str.lower(activation_code_input)==u'123456':
                print('\n检测到您启用了试用模式，在此模式下，您可以试用10分钟，如果您喜欢此软件，请联系商家购买正版授权！')
                time_start=time.time()
                is_now_tested=True
                break
            if activation_code[:10]==str.lower(activation_code_input[:10]):
                is_activated=True
                f=open('certificate.key','w',encoding='utf-8')
                f.write(activation_code)
                f.close()
                #将些激活码写入到certificate.key中
                print('\n激活成功，激活密钥已经写入到certificate.key中，请不要删除此文件。')
                break
            else:
                print('\n激活失败，请重新输入!')
                continue

    #在此处批量处理，不使用多线程操作。每个邮箱创建一个浏览器窗口。
    mail_sender = QQMailSender(message=None,create_browser=False)
    #将反复调用同一个实例对象的run方法进行发送，这样确保窗口不关闭。    
    msg=dict()
    #以下三项:主题，正文，附件是不会变化的。
    msg['subject']=mail_sender.get_content_text()['subject']
    msg['body']=mail_sender.get_content_text()['content_text']
    msg['attachments_list']=mail_sender.get_attachments_list()
    sender_list=mail_sender.get_sender_list()
    sender_list.reverse()

    if not sender_list:
        print('\n发件人列表为空，请打开 发件人.csv,然后输入发件QQ号和密码!')
        os.system('pause')

    to_list=mail_sender.get_to_list()
    to_list.reverse()
    if not to_list:
        print('\n收件人列表为空，请打开 收件人.csv,然后输入收件人邮箱。提示:如果是QQ邮箱，可以只填写QQ号。')
        os.system('pause')

    settings=mail_sender.settings
    #将发送的数据放在msg中，一个msg包括1个发件账号，10个收件箱，主题，正文，附件地址。    
    THREADS_MAX=settings['threads_max']
    TO_MAIL_BOX_NUM_OF_EACH_SENDER=settings['mails_count_in_each_thread']
    ADDRESS_NUM_IN_EACH_MAIL=settings['address_num_in_each_mail']
    print('\n基本设置:最大打开的浏览器窗口数量=>%d 个，每一个窗口最大发送邮件数量=>%d 封，每一个发件箱同时发送的收件人数量=>%d 个。\n'%(THREADS_MAX,TO_MAIL_BOX_NUM_OF_EACH_SENDER,ADDRESS_NUM_IN_EACH_MAIL))
    mail_sender=None
    #打印出发件箱数量，收件箱数量等。
    print('\n本次发件箱数量: %d 个，收件箱数量: %d 个，预计开启 %d 个浏览器窗口进行发送。'%(len(sender_list),len(to_list),math.ceil(len(to_list)/TO_MAIL_BOX_NUM_OF_EACH_SENDER)))
    total_of_to_list=len(to_list)
    while to_list:
        if time.time()-time_start>10*60 and is_now_tested:
            print('\n试用结束，如果您喜欢此软件，请联系商家购买正版！')
            break
        if threading.active_count()<=THREADS_MAX:
            sender=sender_list.pop()
            msg['username']=sender[0]
            msg['password']=sender[1]
            sender_list.insert(0,sender) 
            #在msg['to']中装入10个发件地址。
            msg['to']=[]
            for y in range(TO_MAIL_BOX_NUM_OF_EACH_SENDER):
                if to_list:
                    msg['to'].append(to_list.pop())
            if msg['to']:
                message=msg.copy()
                QQMailSender(message=message).start()
            print('\n开始新浏览器窗口，已经发送的收件人数量为: %d个，剩余收件人数量为: %d 个，完成率： %.2f%%'%(total_of_to_list-len(to_list),len(to_list),100*((total_of_to_list-len(to_list))/total_of_to_list)))
            #将未发送的收件人放入到一个表中。
            f=open(r'settings/收件人_剩余.csv','w',encoding='utf-8-sig',newline='')
            writer=csv.writer(f)
            data=[[item] for item in to_list]
            writer.writerows(data)
            f.close()
    if not to_list:
        print(u'\n发件箱分配完毕!')
    os.system('pause')
