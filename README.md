一、实现的功能
1.根据不同的人发送不同的附件
2.直接读取html文件，正文中插入图片。如果邮件内容是纯文本，使用txt文件就可以；但是如果需要在正文中添加图片，就必须使用html格式。而且html可以使用超链接，可以用css美化，建议不管是不是纯文本，直接使用html文件就好。
3.GUI图形界面，选择附件、通讯录、图片等

二、工具
1.smtplib库，主要用到SMTP()类，用来连接服务器，发送邮件，常用方法：
![Image text](https://github.com/songrenqing/SendMail/blob/master/image/1.jpeg)
2.email库，用来编辑邮件内容，包括标题，发件人、接收人、正文等
![Image text](https://github.com/songrenqing/SendMail/blob/master/image/2.jpeg)
3.PyQt5库，图形工具，主要是提供一个操作界面，用来选择附件、图片等

三、准备工作
SMTP中的密码不是指邮箱的登录密码，而是客户端的授权密码，需要先去邮箱里面开启SMTP服务，然后获取授权密码
1.163邮箱，先开启POP3/SMTP服务，然后设置授权密码
![Image text](https://github.com/songrenqing/SendMail/blob/master/image/3.jpeg）
2.QQ邮箱也是一样，在设置中找到开启SMTP服务，然后设置密码
![Image text](https://github.com/songrenqing/SendMail/blob/master/image/4.jpeg）

四、运行界面与结果
1.GUI界面
![Image text](https://github.com/songrenqing/SendMail/blob/master/image/5.jpeg）
2.运行结果
![Image text](https://github.com/songrenqing/SendMail/blob/master/image/6.jpeg）
3.收到的邮件内容
（1）QQ邮箱
![Image text](https://github.com/songrenqing/SendMail/blob/master/image/7.jpeg）
（2）163邮箱
![Image text](https://github.com/songrenqing/SendMail/blob/master/image/8.jpeg）

五、可能遇到的异常
异常：554 b'DT:SPM 163 smtp10,DsCowADnLCYUXxleA5NlLg--.58715S3 1578721044
描述：这个问题是163邮箱发送到qq邮箱遇到的，QQ邮箱拒绝接受此邮件
解决办法：关掉电脑连接的wifi，手机使用移动网络，然后电脑连接手机的热点，再次发送即可。
