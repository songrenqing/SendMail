import os
import sys
import csv
import time
import datetime
import smtplib
from email.mime.base import MIMEBase 
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr
from email.header import Header
from email import encoders

# ==============发送服务器配置==========
sender_host = 'smtp.163.com'  # 默认服务器地址
sender_port = '25' #服务器端口
sender_user = '18516001681@163.com' #用户名
sender_pwd = 'wySR890'
sender_name = '上海百分百公司'

# ===============获取通讯录中的人名和相应的邮箱地址======
def getAddrBook(addrBook):
    '''
        @作用：根据输入的CSV文件，形成相应的通讯录字典
        @返回：字典类型，name为人名，value为对应的邮件地址
    '''
    with open(addrBook,'r') as addrFile: 
        reader = csv.reader(addrFile)
        name = []
        value = []
        for row in reader:
            name.append(row[0])
            value.append(row[1])
    
    addrs = dict(zip(name, value))
    return addrs

# addrs = {name : value}

# ==========根据附件名称中获得的人名，查找通讯录，找到对应的邮件地址=========

def getRecvAddr(addrs,person_name):
    if not person_name in addrs:
        print("没有<"+person_name+">的邮箱地址! 请在联系人中添加此人邮箱后重试。")
        anykey = input("请按任意数字键【0-9】退出程序：")
        if anykey != '':
            time.sleep(1)
            sys.exit(0)
    return addrs[person_name]

# =============读取Html中的内容==============

def getHtmlContent(html):
    content_html = ''

    if not os.path.exists(html):
        print("html文件不存在")
        exit(0)

    with open(html,'rb') as f:
        content_html = f.read()

    return content_html

# =============加载邮件正文中的图片=========

def getMailImage(image):
    img = ''

    if not os.path.exists(image):
        print("图片不存在")
        exit(0)

    with open(image,'rb') as f:
        img = MIMEImage(f.read())
        

    return img

# =============添加附件==================

def addAttch(attach_file):
    att = MIMEBase('application','octet-stream')  # 这两个参数不知道啥意思，二进制流文件
    att.set_payload(open(attach_file,'rb').read())
    # 此时的附件名称为****.xlsx，截取文件名
    att.add_header('Content-Disposition', 'attachment', filename=('gbk','',attach_file.split("//")[-1]))
    encoders.encode_base64(att)
    return att


# =============发送邮件==================
def send(attach_path,addrBook,html,image):
    smtp = smtplib.SMTP()   # 新建smtp对象
    smtp.connect(sender_host)
    smtp.login(sender_user, sender_pwd)
    for root,dirs,files in os.walk(attach_path):
        for attach_file in [file for file in files if file.endswith('.xlsx')]:      # attach_file : ***_2_***.xlsx
            msg = MIMEMultipart('related') #('alternative') 创建纯文本与超文本实例 ('mixed')  #创建带附件的实例 ('related')  #创建内嵌资源的实例
            att_name = attach_file.split(".")[0]
            subject = str(datetime.datetime.now())[0:10]+att_name
            msg['Subject'] = Header(subject,'utf-8')  # 设置邮件主题
            person_name = subject.split("_")[-1]

            addrs = getAddrBook(addrBook)
            recv_addr = getRecvAddr(addrs,person_name)
            
            msg['From'] = formataddr([sender_name,sender_user]) # 设置发件人名称
            msg['To'] = formataddr([person_name,recv_addr]) # 设置收件人名称  
            # mail_content = getMailContent(content_path)
            # msg.attach(MIMEText(mail_content,'plain', 'utf-8'))  # 正文  MIMEText(content,'plain','utf-8')
            content_html = getHtmlContent(html)
            msg.attach(MIMEText(content_html,'html','utf-8'))
            mail_img = getMailImage(image)
            mail_img.add_header('Content-ID', 'image1')
            msg.attach(mail_img)
            attach_file = root+"//"+attach_file
            att = addAttch(attach_file)
            msg.attach(att)  # 附件
            try:
                smtp.sendmail(sender_user, [recv_addr,], msg.as_string())  # smtp.sendmail(from_addr, to_addrs, msg)
                print("已发送： "+person_name+" <"+recv_addr+">")
            except smtplib.SMTPConnectError as e:
                print('邮件发送失败，连接失败:', e.smtp_code, e.smtp_error)
            except smtplib.SMTPAuthenticationError as e:
                print('邮件发送失败，认证错误:', e.smtp_code, e.smtp_error)
            except smtplib.SMTPSenderRefused as e:
                print('邮件发送失败，发件人被拒绝:', e.smtp_code, e.smtp_error)
            except smtplib.SMTPRecipientsRefused as e:
                print('邮件发送失败，收件人被拒绝:', e.smtp_code, e.smtp_error)
            except smtplib.SMTPDataError as e:
                print('邮件发送失败，数据接收拒绝:', e.smtp_code, e.smtp_error)
            except smtplib.SMTPException as e:
                print('邮件发送失败, ', e.message)
            except Exception as e:
                print('邮件发送异常, ', str(e))
        smtp.quit()

    